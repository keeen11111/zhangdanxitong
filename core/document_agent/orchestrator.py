"""Bounded model/tool orchestration for document-driven workbook runs."""
from __future__ import annotations

from datetime import datetime, timezone
import json
import os
from typing import Any, Callable, Mapping, Sequence
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field

from .contracts import AgentEvent, ToolCall, ToolResult
from .model import ModelConfig, ModelProviderError, ModelResponse, OpenAICompatibleProvider

REPEATED_TOOL_CALL_LIMIT = 3
# Keep enough context for a multi-step workbook change without replaying many
# large range reads on every request.
RECENT_CONVERSATION_TURN_LIMIT = 6
TOOL_RESULT_MAX_CHARS = 12000
MODEL_RESPONSE_RETRY_LIMIT = 3
WRITE_TOOL_NAMES = {
    "prepare_workbook_copy",
    "apply_source_cells",
    "apply_formula_divisors",
    "insert_and_copy_row",
    "apply_cell_changes",
    "copy_formula_from_reference",
    "rollback_work_item",
    "publish_workbook",
}


class ToolExecutionError(RuntimeError):
    """A deliberately safe business error that may be shown to the model."""


class OrchestrationResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    status: str
    code: str | None = None
    content: str = ""
    events: list[AgentEvent] = Field(default_factory=list)


class ToolRegistry:
    """Allow-list of deterministic workbook operations exposed to the model."""

    def __init__(self) -> None:
        self._handlers: dict[str, Callable[..., Any]] = {}

    def register(self, name: str, handler: Callable[..., Any]) -> None:
        if name not in {
            "prepare_workbook_copy", "run_basic_payroll_processor", "apply_source_cells", "search_cells", "apply_formula_divisors", "insert_and_copy_row",
            "inspect_workbook", "classify_file", "find_table", "read_range",
            "inspect_source_file", "read_source_range",
            "match_person", "propose_changes", "apply_cell_changes",
            "select_sheet_mapping",
            "copy_formula_from_reference", "validate_workbook", "rollback_work_item",
            "publish_workbook",
        }:
            raise ValueError(f"unsupported tool: {name}")
        self._handlers[name] = handler

    def execute(self, call: ToolCall) -> ToolResult:
        handler = self._handlers.get(call.name)
        if handler is None:
            return ToolResult(call_id=call.call_id, name=call.name, status="failed", ok=False, error="工具未注册")
        try:
            output = handler(**call.arguments)
            return ToolResult(call_id=call.call_id, name=call.name, status="succeeded", ok=True, output=output, observation=output)
        except ToolExecutionError as exc:
            return ToolResult(call_id=call.call_id, name=call.name, status="failed", ok=False, error=str(exc))
        except Exception:
            # Never leak stack traces, workbook paths or provider content to the model.
            return ToolResult(call_id=call.call_id, name=call.name, status="failed", ok=False, error="工具执行失败")

    def restrict_to(self, names: set[str]) -> None:
        """Remove handlers that this model phase is not authorized to call."""
        self._handlers = {name: handler for name, handler in self._handlers.items() if name in names}


class ModelOrchestrator:
    """Drive model turns while keeping all side effects behind ToolRegistry."""

    def __init__(
        self,
        provider: Any | None = None,
        registry: ToolRegistry | None = None,
        max_turns: int | None = None,
        fallback_providers: Sequence[Any] | None = None,
    ):
        self.provider = provider
        self.fallback_providers = list(fallback_providers or [])
        self.registry = registry or ToolRegistry()
        # Normal project processing has no artificial turn ceiling. Tests and
        # controlled callers can still request an explicit bounded run.
        self.max_turns = None if max_turns is None else max(1, max_turns)

    def run(
        self,
        *,
        run_id: str,
        messages: Sequence[Mapping[str, Any]],
        tools: Sequence[Mapping[str, Any]] | None = None,
        on_event: Callable[[AgentEvent], None] | None = None,
    ) -> OrchestrationResult:
        provider = self.provider
        if provider is None:
            config = ModelConfig.from_env()
            if config is None:
                return OrchestrationResult(status="blocked", code="MODEL_CONFIGURATION_REQUIRED")
            provider = OpenAICompatibleProvider(config)
        # Failover is opt-in for payroll runs. A model change can alter
        # business decisions, so deployments may require the configured
        # primary model to be retried instead of silently switching models.
        disable_failover = os.getenv("PAYROLL_MODEL_DISABLE_FAILOVER", "").strip().lower() in {"1", "true", "yes"}
        providers = [provider] if disable_failover else [provider, *self.fallback_providers]
        provider_index = 0
        initial_messages = list(messages)
        conversation = list(initial_messages)
        recent_turns: list[list[dict[str, Any]]] = []
        events: list[AgentEvent] = []
        revision = 0
        content = ""
        last_provider_error = ""
        turn = 0
        repeated_call_count = 0
        previous_call_fingerprint: str | None = None
        read_fingerprints_since_write: set[str] = set()

        def record(event: AgentEvent) -> None:
            """Expose an event before the next potentially slow model turn."""
            events.append(event)
            if on_event is not None:
                on_event(event)

        while self.max_turns is None or turn < self.max_turns:
            turn += 1
            request_messages = conversation
            switched_provider = False
            for attempt in range(1, MODEL_RESPONSE_RETRY_LIMIT + 1):
                revision += 1
                record(AgentEvent(
                    event_id=uuid4().hex,
                    run_id=run_id,
                    revision=revision,
                    type="model_request",
                    payload={
                        "stage": "model_call",
                        "turn": turn,
                        "attempt": attempt,
                        "label": "正在请求 Agent 分析并等待受控操作指令",
                    },
                ))
                try:
                    response: ModelResponse = provider.complete(messages=request_messages, tools=tools or [])
                except ModelProviderError as exc:
                    last_provider_error = str(exc).strip() or "模型服务请求失败"
                    if "HTTP 400" in last_provider_error and attempt < MODEL_RESPONSE_RETRY_LIMIT:
                        # A compatible endpoint may reject an oversized or
                        # malformed accumulated tool transcript. Persisted
                        # workbook writes are authoritative, so retry from the
                        # original task context with a compact continuation
                        # message instead of replaying the whole transcript.
                        request_messages = [
                            *initial_messages,
                            {
                                "role": "user",
                                "content": "上一轮请求被模型服务拒绝（HTTP 400）。请从工作簿当前已保存进度继续，仅调用一个符合工具 schema 的受控工具；不要重复已完成写入。",
                            },
                        ]
                    if "工具调用参数无效" in last_provider_error and attempt < MODEL_RESPONSE_RETRY_LIMIT:
                        # DeepSeek occasionally emits a malformed tool call.
                        # Keep the same conversation but explicitly request a
                        # fresh JSON call; never infer missing arguments locally.
                        request_messages = [
                            *conversation,
                            {
                                "role": "user",
                                "content": "上一轮工具参数不是合法 JSON。请重新输出一次完整、严格符合工具 schema 的 JSON 工具调用，不要省略必填字段。",
                            },
                        ]
                    if attempt < MODEL_RESPONSE_RETRY_LIMIT:
                        continue
                    if provider_index + 1 < len(providers):
                        provider_index += 1
                        provider = providers[provider_index]
                        revision += 1
                        record(AgentEvent(
                            event_id=uuid4().hex, run_id=run_id, revision=revision,
                            type="model_request", payload={
                                "stage": "model_failover", "provider_index": provider_index,
                                "label": "主模型连续请求失败，已切换备用 Agent；当前上下文已保留",
                            },
                        ))
                        switched_provider = True
                        break
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex,
                        run_id=run_id,
                        revision=revision,
                        type="needs_user_input",
                        payload={
                            "code": "MODEL_PROVIDER_ERROR",
                            "detail": last_provider_error,
                        },
                    ))
                    failure_content = (
                        f"{content}\n\n模型服务错误：{last_provider_error}"
                        if content.strip()
                        else last_provider_error
                    )
                    return OrchestrationResult(
                        status="execution_incomplete",
                        code="MODEL_PROVIDER_ERROR",
                        content=failure_content,
                        events=events,
                    )
                revision += 1
                record(AgentEvent(
                    event_id=uuid4().hex,
                    run_id=run_id,
                    revision=revision,
                    type="model_response",
                    payload={"content": response.content, "tool_call_count": len(response.tool_calls)},
                ))
                if response.content.strip() or response.tool_calls:
                    break
                # An empty response is commonly caused by the provider
                # exhausting context/reasoning budget.  Retrying the full
                # tool transcript only reproduces that condition and makes a
                # run appear stuck.  Workbook writes are already durable, so
                # retry from the original task with a compact continuation.
                if attempt < MODEL_RESPONSE_RETRY_LIMIT:
                    request_messages = [
                        *initial_messages,
                        {
                            "role": "user",
                            "content": (
                                "上一轮模型返回空响应。请从工作簿当前已保存进度继续，"
                                "只调用一个符合工具 schema 的受控工具；不要重复已完成写入。"
                            ),
                        },
                    ]
                if attempt == MODEL_RESPONSE_RETRY_LIMIT:
                    if provider_index + 1 < len(providers):
                        provider_index += 1
                        provider = providers[provider_index]
                        revision += 1
                        record(AgentEvent(
                            event_id=uuid4().hex, run_id=run_id, revision=revision,
                            type="model_request", payload={
                                "stage": "model_failover", "provider_index": provider_index,
                                "label": "主模型返回空响应，已切换备用 Agent；当前上下文已保留",
                            },
                        ))
                        switched_provider = True
                        break
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex,
                        run_id=run_id,
                        revision=revision,
                        type="needs_user_input",
                        payload={
                            "code": "EMPTY_MODEL_RESPONSE",
                            "detail": "模型连续返回空响应，已保留当前处理进度",
                        },
                    ))
                    return OrchestrationResult(
                        status="execution_incomplete",
                        code="EMPTY_MODEL_RESPONSE",
                        content=content,
                        events=events,
                    )
            if switched_provider:
                continue
            content = response.content
            if not response.tool_calls:
                return OrchestrationResult(status="completed", content=content, events=events)
            turn_messages = [{
                "role": "assistant", "content": response.content or "",
                "tool_calls": [
                    {"id": call.call_id, "type": "function", "function": {
                        "name": call.name,
                        "arguments": json.dumps(call.arguments, ensure_ascii=False),
                    }}
                    for call in response.tool_calls
                ],
            }]
            for call in response.tool_calls:
                fingerprint = json.dumps(
                    {"name": call.name, "arguments": call.arguments},
                    ensure_ascii=False,
                    sort_keys=True,
                    default=str,
                )
                repeated_call_count = (
                    repeated_call_count + 1
                    if fingerprint == previous_call_fingerprint
                    else 1
                )
                previous_call_fingerprint = fingerprint
                repeat_limit = 2 if call.name in WRITE_TOOL_NAMES else REPEATED_TOOL_CALL_LIMIT
                if call.name not in WRITE_TOOL_NAMES and fingerprint in read_fingerprints_since_write:
                    # Re-reading is harmless but does not advance the task.
                    # Feed a deterministic tool error back to the same model
                    # so it can choose a write, a different range, or finish;
                    # reserve hard blocking for repeated write calls below.
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="tool_call", payload={"call_id": call.call_id, "name": call.name}, item_id=None,
                    ))
                    revision += 1
                    detail = "该读取范围已读取过，请改读尚未核对的范围，执行写入/校验，或明确说明未完成事项"
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="tool_result", payload={"call_id": call.call_id, "name": call.name, "status": "failed", "error": detail},
                    ))
                    turn_messages.append({
                        "role": "tool", "tool_call_id": call.call_id, "name": call.name,
                        "content": json.dumps({"error": detail}, ensure_ascii=False),
                    })
                    read_fingerprints_since_write.discard(fingerprint)
                    repeated_call_count = 0
                    previous_call_fingerprint = None
                    continue
                if repeated_call_count >= repeat_limit:
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex,
                        run_id=run_id,
                        revision=revision,
                        type="run_blocked",
                        payload={
                            "code": "NO_PROGRESS_DETECTED",
                            "tool": call.name,
                            "detail": "模型连续重复同一工具请求，已停止以避免重复写入",
                        },
                    ))
                    return OrchestrationResult(
                        status="blocked",
                        code="NO_PROGRESS_DETECTED",
                        content=content,
                        events=events,
                    )
                revision += 1
                record(AgentEvent(event_id=uuid4().hex, run_id=run_id, revision=revision, type="tool_call", payload={"call_id": call.call_id, "name": call.name}, item_id=None))
                result = self.registry.execute(call)
                revision += 1
                record(AgentEvent(event_id=uuid4().hex, run_id=run_id, revision=revision, type="tool_result", payload={"call_id": result.call_id, "name": result.name, "status": result.status, "error": result.error}))
                tool_content = json.dumps(
                    result.output if result.status == "succeeded" else {"error": result.error},
                    ensure_ascii=False,
                    default=str,
                )
                if len(tool_content) > TOOL_RESULT_MAX_CHARS:
                    tool_content = (
                        tool_content[:TOOL_RESULT_MAX_CHARS]
                        + "\n[工具结果过长，已截断；请缩小读取范围后继续]"
                    )
                turn_messages.append({
                    "role": "tool",
                    "tool_call_id": result.call_id,
                    "name": result.name,
                    "content": tool_content,
                })
                if call.name in WRITE_TOOL_NAMES and result.status == "succeeded":
                    read_fingerprints_since_write.clear()
                elif call.name not in WRITE_TOOL_NAMES and result.status == "succeeded":
                    read_fingerprints_since_write.add(fingerprint)
            recent_turns.append(turn_messages)
            conversation = [
                *initial_messages,
                *(message for prior_turn in recent_turns[-RECENT_CONVERSATION_TURN_LIMIT:] for message in prior_turn),
            ]
        revision += 1
        record(AgentEvent(event_id=uuid4().hex, run_id=run_id, revision=revision, type="needs_user_input", payload={"code": "MAX_TURNS_EXCEEDED"}))
        return OrchestrationResult(status="execution_incomplete", code="MAX_TURNS_EXCEEDED", content=content, events=events)
