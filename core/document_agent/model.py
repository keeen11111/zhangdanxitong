"""Small, dependency-free model provider boundary.

Only an OpenAI-compatible chat endpoint is implemented initially.  Provider
responses are treated as untrusted input and converted into typed contracts;
the rest of the agent never receives raw provider payloads or credentials.
"""
from __future__ import annotations

import json
import os
import ast
import http.client
import time
import urllib.error
import urllib.request
import re
import math
from uuid import uuid4
from typing import Any, Iterator, Literal, Mapping, Sequence

from pydantic import BaseModel, ConfigDict, Field, SecretStr

from .contracts import ToolCall


def estimate_token_count(value: Any) -> int:
    """Conservative provider-independent token estimate for budget decisions.

    A tokenizer is not guaranteed to be installed for every compatible model.
    ASCII prose/JSON is estimated at roughly four characters per token, while
    non-ASCII text is counted one character per token.  The latter is
    intentionally conservative for Chinese payroll data, where a four-char
    estimate materially undercounts the provider token budget.
    """
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False, default=str)
    ascii_count = sum(character.isascii() for character in text)
    non_ascii_count = len(text) - ascii_count
    return max(0, math.ceil(ascii_count / 4 + non_ascii_count))


def _deduplicate_provider_tools(
    tools: Sequence[Mapping[str, Any]] | None,
) -> tuple[list[Mapping[str, Any]], list[str]]:
    """Normalize the final provider contract to one definition per name.

    Different execution entry points can compose overlapping read, write and
    validation contracts.  OpenAI-compatible providers reject duplicate
    function names with HTTP 400 before model execution, so this boundary is
    the last place where the invariant can be guaranteed for every caller.
    The first definition remains authoritative; the names removed are exposed
    in request metadata for diagnostics without sending duplicate schemas.
    """
    unique: list[Mapping[str, Any]] = []
    seen: set[str] = set()
    duplicates: list[str] = []
    for tool in tools or []:
        function = tool.get("function") if isinstance(tool, Mapping) else None
        name = function.get("name") if isinstance(function, Mapping) else None
        if isinstance(name, str) and name:
            if name in seen:
                if name not in duplicates:
                    duplicates.append(name)
                continue
            seen.add(name)
        unique.append(tool)
    return unique, duplicates


class ModelProviderError(RuntimeError):
    """Safe provider failure with bounded, secret-free diagnostics."""

    def __init__(self, message: str, *, status_code: int | None = None,
                 error_type: str | None = None, error_code: str | None = None,
                 body_summary: str | None = None, request_meta: Mapping[str, Any] | None = None):
        super().__init__(message)
        self.status_code = status_code
        self.error_type = error_type
        self.error_code = error_code
        self.body_summary = body_summary or ""
        self.request_meta = dict(request_meta or {})

    def diagnostics(self) -> dict[str, Any]:
        return {
            "status_code": self.status_code,
            "error_type": self.error_type,
            "error_code": self.error_code,
            "body_summary": self.body_summary,
            **{key: value for key, value in self.request_meta.items()
               if key in {"input_chars", "estimated_input_tokens", "tool_schema_chars",
                          "estimated_tool_tokens", "max_output_tokens", "stage", "provider", "model"}},
        }


# 单次模型请求的硬性超时上限：超过该值的配置会被压缩，
# 避免"300 秒超时 × 3 次重试 = 15 分钟等待"的失败路径。
MAX_MODEL_TIMEOUT_SECONDS = 120.0


def check_model_connectivity(config: "ModelConfig", *, timeout_seconds: float = 15.0) -> str | None:
    """Send one minimal real request so misconfiguration fails in seconds.

    Returns None when the endpoint answers normally; otherwise a safe,
    user-facing reason (auth failure, wrong URL, unknown model, timeout).
    """
    probe = config.model_copy(update={
        "timeout_seconds": max(5.0, min(timeout_seconds, config.timeout_seconds)),
        "max_output_tokens": 16,
        "enable_thinking": False,
        "reasoning_effort": "none",
    })
    try:
        OpenAICompatibleProvider(probe).complete(messages=[{"role": "user", "content": "ping"}])
    except ModelProviderError as exc:
        return str(exc)
    return None


def _http_error_message(status: int) -> str:
    """Convert provider HTTP failures into actionable, secret-free messages."""
    if status in {401, 403}:
        return "模型 API Key 或服务权限认证失败，请检查模型地址和密钥"
    if status == 404:
        return "模型接口地址不存在，请检查 API 地址和接口类型"
    if status == 429:
        return "模型服务请求过于频繁，请稍后重试"
    if status >= 500:
        return "模型服务暂时不可用"
    return f"模型服务请求失败（HTTP {status}）"


class ModelConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    provider: str = Field(min_length=1, max_length=64)
    base_url: str = Field(min_length=1, max_length=500)
    api_key: SecretStr
    model: str = Field(min_length=1, max_length=256)
    timeout_seconds: float = Field(default=60.0, gt=0, le=300)
    api_style: Literal["chat_completions", "responses"] = "chat_completions"
    reasoning_effort: Literal["none", "minimal", "low", "medium", "high", "xhigh"] = "low"
    max_output_tokens: int = Field(default=512, ge=16, le=32768)
    enable_thinking: bool = True

    @classmethod
    def from_env_prefix(cls, prefix: str = "PAYROLL_MODEL") -> "ModelConfig | None":
        """Load one provider configuration without exposing its secret.

        A separate prefix lets the agent keep a second provider ready for
        failover while preserving the existing primary configuration contract.
        """
        values: dict[str, Any] = {
            "provider": os.getenv(f"{prefix}_PROVIDER", "").strip(),
            "base_url": os.getenv(f"{prefix}_BASE_URL", "").strip(),
            "api_key": os.getenv(f"{prefix}_API_KEY", "").strip(),
            "model": os.getenv(f"{prefix}_NAME", "").strip(),
            "api_style": os.getenv(f"{prefix}_API_STYLE", "chat_completions").strip(),
            "reasoning_effort": os.getenv(f"{prefix}_REASONING_EFFORT", "low").strip(),
            "timeout_seconds": os.getenv(f"{prefix}_TIMEOUT_SECONDS", "120").strip(),
            "max_output_tokens": os.getenv(f"{prefix}_MAX_OUTPUT_TOKENS", "512").strip(),
            "enable_thinking": os.getenv(f"{prefix}_ENABLE_THINKING", "true").strip().lower(),
        }
        # Reuse the documented DeepSeek secret when the primary payroll
        # profile intentionally omits a duplicate key in .env.
        if prefix == "PAYROLL_MODEL" and "api.deepseek.com" in values["base_url"]:
            # DeepSeek credentials are authoritative for a DeepSeek endpoint;
            # ignore stale provider-specific keys inherited by the desktop app.
            values["api_key"] = os.getenv("DEEPSEEK_API_KEY", "").strip()
        elif prefix == "PAYROLL_MODEL" and not values["api_key"]:
            values["api_key"] = os.getenv("DEEPSEEK_API_KEY", "").strip()
        if not all(values[key] for key in ("provider", "base_url", "api_key", "model")):
            return None
        # Clamp oversized timeouts so one slow request cannot dominate a run.
        try:
            values["timeout_seconds"] = str(min(float(values["timeout_seconds"]), MAX_MODEL_TIMEOUT_SECONDS))
        except (TypeError, ValueError):
            values["timeout_seconds"] = "90"
        # Do not permit arbitrary providers to be silently treated as the
        # OpenAI wire format; a future provider gets an explicit adapter.
        if values["provider"] != "openai_compatible":
            return None
        return cls(**values)

    @classmethod
    def from_env(cls) -> "ModelConfig | None":
        # The explicit PAYROLL_MODEL_* configuration is the primary contract.
        # Convenience aliases are used only when the explicit configuration is
        # absent, so process-local overrides and tests remain deterministic.
        explicit = cls.from_env_prefix()
        deepseek_model = os.getenv("DEEPSEEK_MODEL", "").strip()
        # Preserve the documented DeepSeek v4 alias as an explicit migration
        # path. Other alias values never override PAYROLL_MODEL_* settings.
        # Explicit PAYROLL_MODEL_* settings always win over legacy aliases.
        # This prevents a stale DEEPSEEK_MODEL environment variable from
        # silently switching the configured payroll model after restart.
        if explicit is not None:
            return explicit
        key = os.getenv("DEEPSEEK_API_KEY", "").strip()
        if key and deepseek_model == "deepseek-v4-pro":
            return cls(
                provider="openai_compatible",
                base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1").strip(),
                api_key=key,
                model=os.getenv("DEEPSEEK_MODEL", "deepseek-v4-pro").strip(),
                api_style="chat_completions",
                reasoning_effort=os.getenv("DEEPSEEK_REASONING_EFFORT", "high").strip(),
                timeout_seconds=os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "90").strip(),
                max_output_tokens=os.getenv("DEEPSEEK_MAX_OUTPUT_TOKENS", "8192").strip(),
                enable_thinking=True,
            )
        return cls.from_env_prefix()

    @classmethod
    def fallback_from_env(cls) -> "ModelConfig | None":
        """Load the optional secondary Agent provider."""
        if os.getenv("PAYROLL_MODEL_DISABLE_FAILOVER", "").strip().lower() in {"1", "true", "yes"}:
            return None
        config = cls.from_env_prefix("PAYROLL_MODEL_FALLBACK")
        if config is not None:
            return config
        # Convenience aliases keep DeepSeek setup small while remaining
        # opt-in: no fallback is enabled unless its key is present.
        key = os.getenv("DEEPSEEK_API_KEY", "").strip()
        if not key:
            return None
        return cls(
            provider="openai_compatible",
            base_url=os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1").strip(),
            api_key=key,
            model=os.getenv("DEEPSEEK_FALLBACK_MODEL", "deepseek-v4-flash").strip(),
            api_style="chat_completions",
            reasoning_effort=os.getenv("DEEPSEEK_REASONING_EFFORT", "high").strip(),
            timeout_seconds=os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "180").strip(),
            max_output_tokens=os.getenv("DEEPSEEK_MAX_OUTPUT_TOKENS", "8192").strip(),
            enable_thinking=True,
        )

    @property
    def secret(self) -> str:
        return self.api_key.get_secret_value()


class ModelResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")

    content: str = ""
    tool_calls: list[ToolCall] = Field(default_factory=list)
    finish_reason: str | None = None
    request_id: str | None = None
    provider_request_id: str | None = None
    input_tokens: int | None = Field(default=None, ge=0)
    output_tokens: int | None = Field(default=None, ge=0)


class OpenAICompatibleProvider:
    """HTTP adapter for providers exposing ``/chat/completions``."""

    def __init__(self, config: ModelConfig):
        if config.provider != "openai_compatible":
            raise ValueError("unsupported model provider")
        self.config = config
        self.last_request_meta: dict[str, Any] = {}
        self._pending_request_id: str | None = None

    def set_request_id(self, request_id: str) -> None:
        """Attach the orchestrator's local correlation ID to the next call."""
        self._pending_request_id = request_id

    def _request_metadata(self, messages: Sequence[Mapping[str, Any]], tools: Sequence[Mapping[str, Any]] | None,
                          *, stage: str | None = None) -> dict[str, Any]:
        input_chars = len(json.dumps(list(messages), ensure_ascii=False, default=str))
        tool_schema_chars = len(json.dumps(list(tools or []), ensure_ascii=False, default=str))
        return {
            "input_chars": input_chars,
            "estimated_input_tokens": estimate_token_count(json.dumps(list(messages), ensure_ascii=False, default=str)),
            "tool_schema_chars": tool_schema_chars,
            "estimated_tool_tokens": estimate_token_count(json.dumps(list(tools or []), ensure_ascii=False, default=str)),
            "max_output_tokens": self.config.max_output_tokens,
            "provider": self.config.provider,
            "model": self.config.model,
            "stage": stage or "model_call",
        }

    @staticmethod
    def _decode_tool_arguments(value: Any) -> dict[str, Any]:
        """Decode provider arguments with syntax-only recovery.

        Some compatible endpoints serialize a JSON object using Python literal
        spelling (single quotes/True/None), or wrap valid JSON in a markdown
        fence. ``literal_eval`` is deliberately limited to literals and the
        result is still required to be an object; no missing fields are guessed.
        """
        if isinstance(value, dict):
            return value
        if not isinstance(value, str):
            raise ValueError
        text = value.strip()
        if text.startswith("```") and text.endswith("```"):
            lines = text.splitlines()
            text = "\n".join(lines[1:-1]).strip()
            if text.lower().startswith("json\n"):
                text = text[5:].lstrip()
        try:
            parsed = json.loads(text)
        except (TypeError, ValueError, json.JSONDecodeError):
            try:
                parsed = ast.literal_eval(text)
            except (SyntaxError, ValueError, TypeError, MemoryError, RecursionError):
                raise ValueError from None
        if not isinstance(parsed, dict):
            raise ValueError
        return parsed

    def complete(
        self,
        *,
        messages: Sequence[Mapping[str, Any]],
        tools: Sequence[Mapping[str, Any]] | None = None,
        temperature: float = 0.0,
    ) -> ModelResponse:
        local_request_id = self._pending_request_id or uuid4().hex
        self._pending_request_id = None
        normalized_tools, duplicate_tool_names = _deduplicate_provider_tools(tools)
        self.last_request_meta = self._request_metadata(messages, normalized_tools)
        if duplicate_tool_names:
            self.last_request_meta["deduplicated_tool_names"] = duplicate_tool_names
        self.last_request_meta["request_id"] = local_request_id
        if self.config.api_style == "responses":
            response = self._complete_responses(messages=messages, tools=normalized_tools)
            response.provider_request_id = response.request_id
            response.request_id = local_request_id
            self.last_request_meta["provider_request_id"] = response.provider_request_id
            return response
        base = self.config.base_url.rstrip("/")
        url = f"{base}/chat/completions"
        payload: dict[str, Any] = {
            "model": self.config.model,
            "messages": list(messages),
            "temperature": temperature,
            "max_tokens": self.config.max_output_tokens,
        }
        if normalized_tools:
            payload["tools"] = list(normalized_tools)
            payload["tool_choice"] = "auto"
        if not self.config.enable_thinking and "api.deepseek.com" in self.config.base_url:
            # deepseek-v4 系列是默认开启思考的混合推理模型，且思考 token 计入
            # max_tokens：复杂任务思考过长会在输出正文前耗尽额度，返回空 content
            # （finish_reason=length, content=""）。ENABLE_THINKING=false 的意图
            # 必须显式下发 thinking 参数，API 才会真正关闭思考。
            payload["thinking"] = {"type": "disabled"}
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            url,
            data=body,
            method="POST",
            headers={
                "Authorization": f"Bearer {self.config.secret}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
        )
        try:
            try:
                response_context = urllib.request.urlopen(request, timeout=self.config.timeout_seconds)
            except TypeError:
                # Tiny test doubles and some embedded urllib shims expose only
                # the request positional argument. Production urllib receives
                # the timeout above.
                response_context = urllib.request.urlopen(request)
            with response_context as response:
                raw = response.read()
            decoded = json.loads(raw.decode("utf-8"))
        except urllib.error.HTTPError as exc:
            raise self._provider_http_error(exc) from exc
        except (urllib.error.URLError, http.client.HTTPException, TimeoutError, OSError) as exc:
            raise ModelProviderError("模型服务暂时不可用") from exc
        except (UnicodeDecodeError, json.JSONDecodeError, TypeError, ValueError) as exc:
            raise ModelProviderError("模型响应不是有效 JSON") from exc
        response = self._parse_response(decoded)
        response.provider_request_id = response.request_id
        response.request_id = local_request_id
        self.last_request_meta["provider_request_id"] = response.provider_request_id
        return response

    def _provider_http_error(self, exc: urllib.error.HTTPError) -> ModelProviderError:
        """Extract only bounded diagnostic fields from an HTTP error body."""
        raw = b""
        try:
            raw = exc.read(8192)
        except Exception:
            pass
        error_type = error_code = None
        summary = ""
        try:
            decoded = json.loads(raw.decode("utf-8", errors="replace")) if raw else {}
            detail = decoded.get("error") if isinstance(decoded, dict) else None
            if isinstance(detail, dict):
                error_type = str(detail.get("type") or "")[:128] or None
                error_code = str(detail.get("code") or "")[:128] or None
                summary = str(detail.get("message") or "")
            elif isinstance(detail, str):
                summary = detail
        except Exception:
            summary = raw.decode("utf-8", errors="replace")[:500]
        summary = re.sub(r"(?:Bearer\s+|api[_-]?key\s*[=:]\s*)[^\s,;]+", "[redacted]", summary, flags=re.I)
        summary = re.sub(r"[A-Za-z0-9+/=_-]{24,}", "[redacted]", summary)[:500]
        return ModelProviderError(
            _http_error_message(exc.code), status_code=exc.code,
            error_type=error_type, error_code=error_code,
            body_summary=summary, request_meta=self.last_request_meta,
        )

    def stream_text(
        self,
        *,
        messages: Sequence[Mapping[str, Any]],
        temperature: float = 0.0,
    ) -> Iterator[str]:
        """Yield visible text deltas from either supported OpenAI wire format.

        This intentionally supports text-only task chat. Workbook tool calls
        stay on the existing non-streaming orchestration path, where their
        structured arguments can be fully validated before any action.
        """
        if self.config.api_style == "responses":
            payload: dict[str, Any] = {
                "model": self.config.model,
                "input": self._responses_input(messages),
                "reasoning": {"effort": self.config.reasoning_effort},
                "max_output_tokens": self.config.max_output_tokens,
                "enable_thinking": self.config.enable_thinking,
                "stream": True,
            }
            for event in self._stream_sse_json(f"{self.config.base_url.rstrip('/')}/responses", payload):
                if event.get("type") != "response.output_text.delta":
                    continue
                delta = event.get("delta")
                if isinstance(delta, str) and delta:
                    yield delta
            return

        payload = {
            "model": self.config.model,
            "messages": list(messages),
            "temperature": temperature,
            "max_tokens": self.config.max_output_tokens,
            "stream": True,
        }
        if not self.config.enable_thinking and "api.deepseek.com" in self.config.base_url:
            # deepseek-v4 系列是默认开启思考的混合推理模型，且思考 token 计入
            # max_tokens：复杂任务思考过长会在输出正文前耗尽额度，返回空 content
            # （finish_reason=length, content=""）。ENABLE_THINKING=false 的意图
            # 必须显式下发 thinking 参数，API 才会真正关闭思考。
            payload["thinking"] = {"type": "disabled"}
        for event in self._stream_sse_json(f"{self.config.base_url.rstrip('/')}/chat/completions", payload):
            choices = event.get("choices")
            if not isinstance(choices, list) or not choices or not isinstance(choices[0], Mapping):
                continue
            delta = choices[0].get("delta")
            if not isinstance(delta, Mapping):
                continue
            content = delta.get("content")
            if isinstance(content, str) and content:
                yield content

    def _stream_sse_json(self, url: str, payload: Mapping[str, Any]) -> Iterator[Mapping[str, Any]]:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            url,
            data=body,
            method="POST",
            headers={
                "Authorization": f"Bearer {self.config.secret}",
                "Content-Type": "application/json",
                "Accept": "text/event-stream",
            },
        )
        data_lines: list[str] = []
        try:
            try:
                response_context = urllib.request.urlopen(request, timeout=self.config.timeout_seconds)
            except TypeError:
                response_context = urllib.request.urlopen(request)
            with response_context as response:
                for raw_line in response:
                    line = raw_line.decode("utf-8") if isinstance(raw_line, bytes) else str(raw_line)
                    for item in line.splitlines():
                        if not item:
                            if data_lines:
                                raw_data = "\n".join(data_lines).strip()
                                data_lines = []
                                if raw_data == "[DONE]":
                                    return
                                decoded = json.loads(raw_data)
                                if not isinstance(decoded, Mapping):
                                    raise ModelProviderError("模型流式响应结构无效")
                                yield decoded
                            continue
                        if item.startswith("data:"):
                            data_lines.append(item[5:].lstrip())
                if data_lines:
                    raw_data = "\n".join(data_lines).strip()
                    if raw_data != "[DONE]":
                        decoded = json.loads(raw_data)
                        if not isinstance(decoded, Mapping):
                            raise ModelProviderError("模型流式响应结构无效")
                        yield decoded
        except ModelProviderError:
            raise
        except urllib.error.HTTPError as exc:
            raise self._provider_http_error(exc) from exc
        except (urllib.error.URLError, http.client.HTTPException, TimeoutError, OSError) as exc:
            raise ModelProviderError("模型服务暂时不可用") from exc
        except (UnicodeDecodeError, json.JSONDecodeError, TypeError, ValueError) as exc:
            raise ModelProviderError("模型流式响应无效") from exc

    def _request_json(self, url: str, payload: Mapping[str, Any]) -> Any:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(
            url,
            data=body,
            method="POST",
            headers={
                "Authorization": f"Bearer {self.config.secret}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
        )
        try:
            try:
                response_context = urllib.request.urlopen(request, timeout=self.config.timeout_seconds)
            except TypeError:
                response_context = urllib.request.urlopen(request)
            with response_context as response:
                raw = response.read()
            return json.loads(raw.decode("utf-8"))
        except urllib.error.HTTPError as exc:
            raise self._provider_http_error(exc) from exc
        except (urllib.error.URLError, http.client.HTTPException, TimeoutError, OSError) as exc:
            raise ModelProviderError("模型服务暂时不可用") from exc
        except (UnicodeDecodeError, json.JSONDecodeError, TypeError, ValueError) as exc:
            raise ModelProviderError("模型响应不是有效 JSON") from exc

    @staticmethod
    def _responses_input(messages: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
        converted: list[dict[str, Any]] = []
        for message in messages:
            role = str(message.get("role") or "")
            if role == "tool":
                converted.append({
                    "type": "function_call_output",
                    "call_id": str(message.get("tool_call_id") or ""),
                    "output": str(message.get("content") or ""),
                })
                continue
            raw_calls = message.get("tool_calls")
            if role == "assistant" and isinstance(raw_calls, list) and raw_calls:
                for raw_call in raw_calls:
                    if not isinstance(raw_call, Mapping) or not isinstance(raw_call.get("function"), Mapping):
                        raise ModelProviderError("模型工具调用上下文无效")
                    function = raw_call["function"]
                    arguments = function.get("arguments", "{}")
                    if not isinstance(arguments, str):
                        arguments = json.dumps(arguments, ensure_ascii=False)
                    converted.append({
                        "type": "function_call",
                        "call_id": str(raw_call.get("id") or ""),
                        "name": str(function.get("name") or ""),
                        "arguments": arguments,
                    })
                continue
            if role not in {"system", "developer", "user", "assistant"}:
                raise ModelProviderError("模型消息角色无效")
            converted.append({"role": role, "content": str(message.get("content") or "")})
        return converted

    @staticmethod
    def _responses_tools(tools: Sequence[Mapping[str, Any]] | None) -> list[dict[str, Any]]:
        converted: list[dict[str, Any]] = []
        for tool in tools or []:
            function = tool.get("function") if isinstance(tool, Mapping) else None
            if not isinstance(function, Mapping):
                raise ModelProviderError("模型工具定义无效")
            converted.append({
                "type": "function",
                "name": str(function.get("name") or ""),
                "description": str(function.get("description") or ""),
                "parameters": function.get("parameters") or {"type": "object", "properties": {}},
                "strict": True,
            })
        return converted

    def _complete_responses(
        self,
        *,
        messages: Sequence[Mapping[str, Any]],
        tools: Sequence[Mapping[str, Any]] | None,
    ) -> ModelResponse:
        payload: dict[str, Any] = {
            "model": self.config.model,
            "input": self._responses_input(messages),
            "reasoning": {"effort": self.config.reasoning_effort},
            "max_output_tokens": self.config.max_output_tokens,
            "enable_thinking": self.config.enable_thinking,
        }
        converted_tools = self._responses_tools(tools)
        if converted_tools:
            payload.update({"tools": converted_tools, "tool_choice": "auto", "parallel_tool_calls": False})
        decoded = self._request_json(f"{self.config.base_url.rstrip('/')}/responses", payload)
        return self._parse_responses_response(decoded)

    def _parse_responses_response(self, decoded: Any) -> ModelResponse:
        if not isinstance(decoded, dict) or not isinstance(decoded.get("output"), list):
            raise ModelProviderError("模型 Responses 响应结构无效")
        content_parts: list[str] = []
        calls: list[ToolCall] = []
        for output in decoded["output"]:
            if not isinstance(output, dict):
                raise ModelProviderError("模型 Responses 输出结构无效")
            if output.get("type") == "function_call":
                try:
                    arguments = self._decode_tool_arguments(output.get("arguments", {}))
                    calls.append(ToolCall(
                        call_id=str(output.get("call_id") or output.get("id") or "call-unknown"),
                        name=output.get("name"), arguments=arguments,
                    ))
                except Exception as exc:
                    raise ModelProviderError("模型工具调用参数无效") from exc
            if output.get("type") == "message" and isinstance(output.get("content"), list):
                for part in output["content"]:
                    if isinstance(part, dict) and part.get("type") in {"output_text", "text"} and isinstance(part.get("text"), str):
                        content_parts.append(part["text"])
        usage = decoded.get("usage") if isinstance(decoded.get("usage"), dict) else {}
        return ModelResponse(
            content="\n".join(content_parts), tool_calls=calls,
            finish_reason=str(decoded.get("status")) if decoded.get("status") else None,
            request_id=str(decoded.get("id")) if decoded.get("id") else None,
            input_tokens=usage.get("input_tokens") if isinstance(usage.get("input_tokens"), int) else None,
            output_tokens=usage.get("output_tokens") if isinstance(usage.get("output_tokens"), int) else None,
        )

    def _parse_response(self, decoded: Any) -> ModelResponse:
        if not isinstance(decoded, dict):
            raise ModelProviderError("模型响应结构无效")
        choices = decoded.get("choices")
        if not isinstance(choices, list) or not choices or not isinstance(choices[0], dict):
            raise ModelProviderError("模型响应缺少 choices")
        message = choices[0].get("message")
        if not isinstance(message, dict):
            raise ModelProviderError("模型响应缺少 message")
        content = message.get("content", "")
        if content is None:
            content = ""
        if not isinstance(content, str):
            raise ModelProviderError("模型响应 content 类型无效")
        calls: list[ToolCall] = []
        raw_calls = message.get("tool_calls") or []
        if not isinstance(raw_calls, list):
            raise ModelProviderError("模型工具调用结构无效")
        for raw_call in raw_calls:
            try:
                if not isinstance(raw_call, dict):
                    raise ValueError
                function = raw_call.get("function")
                if not isinstance(function, dict):
                    raise ValueError
                arguments = self._decode_tool_arguments(function.get("arguments", {}))
                calls.append(ToolCall(
                    call_id=str(raw_call.get("id") or "call-unknown"),
                    name=function.get("name"),
                    arguments=arguments,
                ))
            except Exception as exc:
                raise ModelProviderError("模型工具调用参数无效") from exc
        usage = decoded.get("usage") if isinstance(decoded.get("usage"), dict) else {}
        return ModelResponse(
            content=content,
            tool_calls=calls,
            finish_reason=(choices[0].get("finish_reason") if isinstance(choices[0].get("finish_reason"), str) else None),
            request_id=(str(decoded.get("id")) if decoded.get("id") is not None else None),
            input_tokens=(usage.get("prompt_tokens") if isinstance(usage.get("prompt_tokens"), int) else None),
            output_tokens=(usage.get("completion_tokens") if isinstance(usage.get("completion_tokens"), int) else None),
        )
