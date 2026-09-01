from __future__ import annotations

import json
import urllib.error
from typing import Any

import pytest
from pydantic import ValidationError

from core.document_agent.contracts import AgentEvent, ModelTrace, ToolCall, ToolResult
from core.document_agent.model import (
    ModelConfig,
    ModelProviderError,
    ModelResponse,
    OpenAICompatibleProvider,
)


def test_model_config_reads_only_complete_openai_compatible_settings(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    for name in (
        "PAYROLL_MODEL_PROVIDER",
        "PAYROLL_MODEL_BASE_URL",
        "PAYROLL_MODEL_API_KEY",
        "PAYROLL_MODEL_NAME",
    ):
        monkeypatch.delenv(name, raising=False)
    assert ModelConfig.from_env() is None

    monkeypatch.setenv("PAYROLL_MODEL_PROVIDER", "openai_compatible")
    monkeypatch.setenv("PAYROLL_MODEL_BASE_URL", "https://model.example/v1")
    monkeypatch.setenv("PAYROLL_MODEL_API_KEY", "test-secret")
    monkeypatch.setenv("PAYROLL_MODEL_NAME", "finance-model")

    config = ModelConfig.from_env()
    assert config is not None
    assert config.provider == "openai_compatible"
    assert config.base_url == "https://model.example/v1"
    assert config.model == "finance-model"
    assert config.enable_thinking is True
    assert "test-secret" not in repr(config)


def test_model_config_reads_deepseek_fallback_from_alias_without_leaking_key(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("DEEPSEEK_API_KEY", "deepseek-secret")
    monkeypatch.delenv("PAYROLL_MODEL_FALLBACK_PROVIDER", raising=False)
    monkeypatch.delenv("PAYROLL_MODEL_FALLBACK_API_KEY", raising=False)

    config = ModelConfig.fallback_from_env()

    assert config is not None
    assert config.base_url == "https://api.deepseek.com/v1"
    assert config.model == "deepseek-v4-flash"
    assert config.api_style == "chat_completions"
    assert "deepseek-secret" not in repr(config)


def test_model_config_prefers_explicit_primary_over_legacy_deepseek_alias(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("PAYROLL_MODEL_PROVIDER", "openai_compatible")
    monkeypatch.setenv("PAYROLL_MODEL_BASE_URL", "https://qwen.example/v1")
    monkeypatch.setenv("PAYROLL_MODEL_API_KEY", "qwen-secret")
    monkeypatch.setenv("PAYROLL_MODEL_NAME", "qwen3.8-max")
    monkeypatch.setenv("DEEPSEEK_API_KEY", "deepseek-secret")
    monkeypatch.setenv("DEEPSEEK_MODEL", "deepseek-v4-pro")
    monkeypatch.setenv("DEEPSEEK_FALLBACK_MODEL", "deepseek-v4-flash")

    primary = ModelConfig.from_env()
    fallback = ModelConfig.fallback_from_env()

    assert primary is not None and primary.model == "qwen3.8-max"
    assert primary.base_url == "https://qwen.example/v1"
    assert fallback is not None and fallback.model == "deepseek-v4-flash"


def test_openai_provider_parses_text_and_validated_tool_calls(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    config = ModelConfig(
        provider="openai_compatible",
        base_url="https://model.example/v1",
        api_key="test-secret",
        model="finance-model", max_output_tokens=2048,
    )
    provider = OpenAICompatibleProvider(config)
    captured: dict[str, Any] = {}

    def fake_request(request: Any) -> Any:
        captured["url"] = request.full_url
        captured["authorization"] = request.get_header("Authorization")
        captured["payload"] = json.loads(request.data)
        return _Response({
            "choices": [{"message": {
                "content": "正在核对张三。",
                "tool_calls": [{
                    "id": "call-1",
                    "type": "function",
                    "function": {
                        "name": "read_range",
                        "arguments": '{"sheet":"工资核算","range":"J12:K12"}',
                    },
                }],
            }}],
        })

    monkeypatch.setattr("urllib.request.urlopen", fake_request)
    response = provider.complete(
        messages=[{"role": "user", "content": "核对张三"}],
        tools=[{"type": "function", "function": {"name": "read_range", "parameters": {}}}],
    )

    assert isinstance(response, ModelResponse)
    assert response.content == "正在核对张三。"
    assert response.tool_calls[0].name == "read_range"
    assert response.tool_calls[0].arguments == {"sheet": "工资核算", "range": "J12:K12"}
    assert captured["url"] == "https://model.example/v1/chat/completions"
    assert captured["authorization"] == "Bearer test-secret"
    assert captured["payload"]["model"] == "finance-model"
    assert captured["payload"]["max_tokens"] == 2048


def test_openai_provider_rejects_malformed_tool_arguments(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    provider = OpenAICompatibleProvider(ModelConfig(
        provider="openai_compatible", base_url="https://model.example/v1/",
        api_key="test-secret", model="finance-model",
    ))

    monkeypatch.setattr("urllib.request.urlopen", lambda _request: _Response({
        "choices": [{"message": {"tool_calls": [{
            "id": "call-1", "type": "function",
            "function": {"name": "read_range", "arguments": "not-json"},
        }]}}],
    }))

    with pytest.raises(ModelProviderError, match="工具调用"):
        provider.complete(messages=[{"role": "user", "content": "继续"}], tools=[])


def test_openai_provider_reports_authentication_failures_without_provider_body(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    provider = OpenAICompatibleProvider(ModelConfig(
        provider="openai_compatible", base_url="https://model.example/v1",
        api_key="test-secret", model="finance-model",
    ))

    def unauthorized(_request: Any, **_kwargs: Any) -> Any:
        raise urllib.error.HTTPError("https://model.example/v1/chat/completions", 401, "Unauthorized", {}, None)

    monkeypatch.setattr("urllib.request.urlopen", unauthorized)

    with pytest.raises(ModelProviderError, match="API Key|认证失败"):
        provider.complete(messages=[{"role": "user", "content": "继续"}])


def test_openai_provider_streams_chat_completion_text_deltas(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    provider = OpenAICompatibleProvider(ModelConfig(
        provider="openai_compatible", base_url="https://model.example/v1",
        api_key="test-secret", model="finance-model",
    ))
    captured: dict[str, Any] = {}

    def fake_request(request: Any, **_kwargs: Any) -> Any:
        captured["url"] = request.full_url
        captured["accept"] = request.get_header("Accept")
        captured["payload"] = json.loads(request.data)
        return _StreamingResponse([
            'data: {"choices":[{"delta":{"content":"正在"}}]}\n\n'.encode("utf-8"),
            'data: {"choices":[{"delta":{"content":"核对。"}}]}\n\n'.encode("utf-8"),
            b'data: [DONE]\n\n',
        ])

    monkeypatch.setattr("urllib.request.urlopen", fake_request)

    assert list(provider.stream_text(messages=[{"role": "user", "content": "继续"}])) == ["正在", "核对。"]
    assert captured["url"] == "https://model.example/v1/chat/completions"
    assert captured["accept"] == "text/event-stream"
    assert captured["payload"]["stream"] is True


def test_responses_provider_streams_output_text_deltas(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    provider = OpenAICompatibleProvider(ModelConfig(
        provider="openai_compatible", base_url="https://model.example/v1",
        api_key="test-secret", model="finance-model", api_style="responses",
    ))

    monkeypatch.setattr("urllib.request.urlopen", lambda _request, **_kwargs: _StreamingResponse([
        b"event: response.output_text.delta\n",
        'data: {"type":"response.output_text.delta","delta":"回复"}\n\n'.encode("utf-8"),
        'data: {"type":"response.output_text.delta","delta":"完成"}\n\n'.encode("utf-8"),
        b'data: [DONE]\n\n',
    ]))

    assert list(provider.stream_text(messages=[{"role": "user", "content": "继续"}])) == ["回复", "完成"]


def test_responses_provider_converts_chat_tools_and_parses_function_calls(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    provider = OpenAICompatibleProvider(ModelConfig(
        provider="openai_compatible",
        base_url="https://model.example/v1",
        api_key="test-secret",
        model="finance-model",
        api_style="responses",
        reasoning_effort="low",
    ))
    captured: dict[str, Any] = {}

    def fake_request(request: Any, **_kwargs: Any) -> Any:
        captured["url"] = request.full_url
        captured["payload"] = json.loads(request.data)
        return _Response({
            "id": "resp-1",
            "status": "completed",
            "output": [{
                "type": "function_call",
                "call_id": "call-1",
                "name": "read_range",
                "arguments": '{"sheet":"工资核算","range":"J2:K2"}',
            }],
            "usage": {"input_tokens": 12, "output_tokens": 8},
        })

    monkeypatch.setattr("urllib.request.urlopen", fake_request)
    response = provider.complete(
        messages=[{"role": "user", "content": "读取奖金"}],
        tools=[{"type": "function", "function": {
            "name": "read_range", "description": "Read cells",
            "parameters": {"type": "object", "properties": {}},
        }}],
    )

    assert captured["url"] == "https://model.example/v1/responses"
    assert captured["payload"]["reasoning"] == {"effort": "low"}
    assert captured["payload"]["enable_thinking"] is True
    assert captured["payload"]["tools"][0]["name"] == "read_range"
    assert response.tool_calls[0].arguments == {"sheet": "工资核算", "range": "J2:K2"}
    assert response.input_tokens == 12


def test_responses_provider_converts_tool_outputs_for_follow_up(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    provider = OpenAICompatibleProvider(ModelConfig(
        provider="openai_compatible", base_url="https://model.example/v1",
        api_key="test-secret", model="finance-model", api_style="responses",
    ))
    captured: dict[str, Any] = {}

    def fake_request(request: Any, **_kwargs: Any) -> Any:
        captured["payload"] = json.loads(request.data)
        return _Response({"id": "resp-2", "status": "completed", "output": []})

    monkeypatch.setattr("urllib.request.urlopen", fake_request)
    provider.complete(messages=[
        {"role": "assistant", "content": "", "tool_calls": [{
            "id": "call-1", "type": "function",
            "function": {"name": "read_range", "arguments": '{"sheet":"工资核算"}'},
        }]},
        {"role": "tool", "tool_call_id": "call-1", "name": "read_range", "content": '{"values":[[2000,0]]}'},
    ])

    assert captured["payload"]["input"][0]["type"] == "function_call"
    assert captured["payload"]["input"][1]["type"] == "function_call_output"


def test_typed_audit_contracts_reject_untyped_or_unsafe_payloads() -> None:
    call = ToolCall(call_id="call-1", name="read_range", arguments={"sheet": "工资核算"})
    result = ToolResult(call_id="call-1", name="read_range", ok=True, observation={"values": [[2000]]})
    trace = ModelTrace(provider="openai_compatible", model="finance-model", response=ModelResponse(
        content="读取完成", tool_calls=[call],
    ))
    event = AgentEvent(kind="tool_result", run_id="run-1", tool_result=result, model_trace=trace)

    assert event.tool_result == result
    assert event.model_trace.response.tool_calls[0].call_id == "call-1"
    with pytest.raises(ValidationError):
        ToolCall(call_id="call-2", name="arbitrary_python", arguments={})
    with pytest.raises(ValidationError):
        AgentEvent(kind="tool_result", run_id="run-1", tool_result={"ok": True})


class _Response:
    def __init__(self, payload: dict[str, Any]):
        self._body = json.dumps(payload).encode("utf-8")

    def read(self) -> bytes:
        return self._body

    def __enter__(self) -> "_Response":
        return self

    def __exit__(self, *_args: object) -> None:
        return None


class _StreamingResponse:
    def __init__(self, chunks: list[bytes]):
        self._chunks = chunks

    def __iter__(self):
        return iter(self._chunks)

    def __enter__(self) -> "_StreamingResponse":
        return self

    def __exit__(self, *_args: object) -> None:
        return None
