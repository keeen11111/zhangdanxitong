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
# 会话总字符预算（初始消息 + 读取存档 + 回放窗口）。没有总量控制时，
# “6 轮回放 × 每轮多个 40K 工具结果 + 100K 存档”可逼近 300K 字符，
# 远超模型 128K 上下文窗口——超限请求会得到空 content（非报错），
# 表现为“读完数据后模型连续空响应、run 卡死 0 写入”。
# 组装时从最新轮往前保留，超出预算的旧轮整体丢弃（其读取数据已在存档）。
CONVERSATION_TOTAL_CHAR_BUDGET = 96_000
# 工具结果回放上限：必须容得下一次“读全整个目标区域”的结果
# （读取工具单次最多1000格，JSON 约 20-40KB）。截断到低于这个量，
# 模型永远看不到完整数据，会陷入“读不全→再读→遗忘→再读”的循环。
TOOL_RESULT_MAX_CHARS = 40_000
MODEL_RESPONSE_RETRY_LIMIT = 3
# 网络超时/服务不可用类错误只重试一次：连续等待两次超时后应快速失败并保留进度，
# 而不是把同一问题拖到分钟级。格式类错误（HTTP 400、工具参数无效）仍走 3 次。
TRANSIENT_MODEL_RETRY_LIMIT = 2
# 模型“只输出权衡文字、不调用任何工具”时的催促上限：连续 N 次催促后仍
# 无工具调用才接受其为最终答复，防止把模型的中间思考误判为任务完成。
NO_TOOL_CALL_NUDGE_LIMIT = 2
# A write can be only one step of a larger workbook plan (for example,
# deleting unmatched rows before copying wage fields).  Some providers emit a
# long chain-of-thought-like continuation with no tool call at that point.  It
# must not be accepted as a completed run merely because an earlier write
# happened.  These markers are intentionally limited to explicit future-work
# language; ordinary completion summaries are unaffected.
INCOMPLETE_PROSE_MARKERS = (
    "now i need",
    "i need to ",
    "next i will",
    "let me reconsider",
    "let me carefully",
    "let me read",
    "let me map",
    "let me build",
    "接下来需要",
    "下一步将",
    "还需要",
    "继续调用",
    "继续写入",
)
# 写入型任务的“只读不写”防护阈值：连续 N 轮模型只调用读取类工具、
# 一次写入都没有时，注入催促要求立即基于已读数据开始写入。
# 没有这个防护，模型会在“读取→超出回放窗口遗忘→再读新范围”里
# 打转直到轮数耗尽（表现为 0 写入 + 连续 MAX_TURNS_EXCEEDED）。
READ_ONLY_TURN_NUDGE_LIMIT = 8
# 真正搬运业务数据的读取工具：其结果必须在写入阶段仍然可见。
# “先读全、再大批写”的执行模式下，写入发生在读取之后的若干轮，
# 若读取结果随回放窗口淘汰，模型到写入时会“失忆”且被读取上限
# 拦住无法重读，形成“读完了却写不了”的死局。
READ_DATA_TOOL_NAMES = {"read_range", "read_source_range", "search_cells"}
# 读取数据存档的总字符上限：超出时丢弃最早的存档。读取次数本身有
# 硬上限，正常任务达不到此值；这是上下文长度的最后防线。
READ_ARCHIVE_MAX_CHARS = 72_000
CONVERSATION_RESERVE_CHARS = 8_000


def _default_max_turns() -> int:
    """Total model-turn ceiling for one orchestration run (env-tunable)."""
    try:
        return max(1, int(os.getenv("AGENT_MAX_TURNS", "16")))
    except ValueError:
        return 16


def _prose_declares_pending_work(content: str) -> bool:
    """Detect explicit future-work language in a no-tool model response."""
    normalized = " ".join(str(content or "").lower().split())
    return any(marker in normalized for marker in INCOMPLETE_PROSE_MARKERS)
WRITE_TOOL_NAMES = {
    "prepare_workbook_copy",
    "run_basic_payroll_processor",
    "run_keyuan_workflow",
    "apply_source_cells",
    "apply_formula_divisors",
    "insert_and_copy_row",
    "delete_rows",
    "apply_cell_changes",
    "copy_formula_from_reference",
    "rollback_work_item",
    "publish_workbook",
}

# 真正写入业务数据的工具。prepare_workbook_copy 只是创建草稿副本，
# 没有写入任何数据：把它算作“已写入”会让 require_writes 保护失效
# （模型建完草稿就输出纯文字总结也会被判定为完成，0 写入收场）。
DATA_WRITE_TOOL_NAMES = WRITE_TOOL_NAMES - {"prepare_workbook_copy"}


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
            "prepare_workbook_copy", "run_basic_payroll_processor", "run_keyuan_workflow", "apply_source_cells", "search_cells", "apply_formula_divisors", "insert_and_copy_row", "delete_rows",
            "inspect_workbook", "classify_file", "find_table", "read_range",
            "inspect_source_file", "read_source_range",
            "match_person", "propose_changes", "apply_cell_changes",
            "select_sheet_mapping",
            "copy_formula_from_reference", "validate_workbook", "rollback_work_item",
            "validate_with_officecli", "publish_workbook",
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
        require_writes: bool = False,
        should_stop: Callable[[], bool] | None = None,
        read_call_limit: int | None = None,
    ):
        self.provider = provider
        self.fallback_providers = list(fallback_providers or [])
        self.registry = registry or ToolRegistry()
        # 一次编排会话有总轮数上限，防止畸形模型响应无限循环；
        # 可用环境变量 AGENT_MAX_TURNS 调整（默认 24）。
        self.max_turns = _default_max_turns() if max_turns is None else max(1, max_turns)
        # require_writes=True 时（如薪资写入任务），模型在没有任何写入的
        # 情况下输出纯文字会被视为中间思考而催促继续执行，防止把权衡
        # 文本误判为任务完成。只读核对类任务保持默认 False。
        self.require_writes = require_writes
        # 协作式停止：每轮模型调用前检查；触发后以 USER_STOPPED 结束，
        # 已完成的写入与事件保持不变，由上层落盘为可续跑状态。
        self.should_stop = should_stop
        # 读取次数硬上限：成功的读取类工具调用累计到该次数后，后续读取
        # 直接被拒绝（返回引导性错误），强制模型转入写入阶段。None 表示
        # 不限制（只读核对任务、测试均保持默认）。用于把“一次读全、
        # 大批写入”从提示词约定升级为确定性保障，杜绝反复读取烧钱。
        self.read_call_limit = read_call_limit

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
        read_fingerprints_since_write: dict[str, int] = {}
        has_written = False
        no_tool_call_nudges = 0
        turns_since_write = 0
        successful_reads = 0
        # 读取数据存档：成功的数据读取结果以 user 消息形式跨轮保留，
        # 不随 RECENT_CONVERSATION_TURN_LIMIT 淘汰。写入阶段（含空响应
        # /HTTP 400 后的紧凑重试）据此仍能看到全部已读数据。
        read_archive: list[dict[str, Any]] = []

        def message_size(message: Mapping[str, Any]) -> int:
            size = len(str(message.get("content") or ""))
            for tool_call in message.get("tool_calls") or []:
                size += len(str((tool_call.get("function") or {}).get("arguments") or ""))
            return size

        initial_message_chars = sum(message_size(message) for message in initial_messages)
        archive_char_budget = min(
            READ_ARCHIVE_MAX_CHARS,
            max(0, CONVERSATION_TOTAL_CHAR_BUDGET - initial_message_chars - CONVERSATION_RESERVE_CHARS),
        )

        def archive_read_result(
            fingerprint: str, call_id: str, name: str, content: str,
        ) -> bool:
            """Archive an exact read result, or reject it before context can overflow."""
            retained = [e for e in read_archive if e["fingerprint"] != fingerprint]
            message = {
                "role": "user",
                "content": f"[已读取数据存档 · {name} · 请求 {fingerprint}]\n{content}",
            }
            if sum(len(e["message"]["content"]) for e in retained) + len(message["content"]) > archive_char_budget:
                return False
            read_archive[:] = retained
            read_archive.append({
                "fingerprint": fingerprint,
                "call_id": call_id,
                # 当前轮即将 append 进 recent_turns 的下标；存档条目只在
                # 其所属轮中用短引用代替原始工具结果，避免同一数据重复回放。
                "turn": len(recent_turns),
                "message": message,
            })
            return True

        def assemble_conversation() -> list[dict[str, Any]]:
            window_start = max(0, len(recent_turns) - RECENT_CONVERSATION_TURN_LIMIT)
            # 总字符预算：防止“多轮大读取 + 存档”把请求撑爆模型上下文窗口。
            # 先为全部读取存档预留空间，再从最新轮往前保留完整工具轮。这样
            # 小型近期读取仍保持 assistant -> tool 的原始 JSON 协议；滑出窗口
            # 或因预算未回放的读取，最后用独立存档消息补入，不重复占上下文。
            archive_reserve = sum(message_size(entry["message"]) for entry in read_archive)
            remaining = (
                CONVERSATION_TOTAL_CHAR_BUDGET
                - sum(message_size(message) for message in initial_messages)
                - archive_reserve
            )
            kept_turns: list[tuple[int, list[dict[str, Any]]]] = []
            for turn_index in range(len(recent_turns) - 1, window_start - 1, -1):
                prior_turn = recent_turns[turn_index]
                size = sum(message_size(message) for message in prior_turn)
                if remaining < size:
                    continue
                kept_turns.append((turn_index, prior_turn))
                remaining -= size
            kept_turns.reverse()
            replayed_call_ids = {
                str(message.get("tool_call_id") or "")
                for _, prior_turn in kept_turns
                for message in prior_turn
                if message.get("role") == "tool"
            }
            detached_archives = [
                entry["message"] for entry in read_archive
                if entry["call_id"] not in replayed_call_ids
            ]
            return [
                *initial_messages,
                *(message for _, prior_turn in kept_turns for message in prior_turn),
                *detached_archives,
            ]

        def compact_request_messages(note: str) -> list[dict[str, Any]]:
            """空响应/HTTP 400 后的紧凑重试：丢弃轮次回放但保留读取存档。"""
            return [
                *initial_messages,
                *(entry["message"] for entry in read_archive),
                {"role": "user", "content": note},
            ]

        def record(event: AgentEvent) -> None:
            """Expose an event before the next potentially slow model turn."""
            events.append(event)
            if on_event is not None:
                on_event(event)

        while self.max_turns is None or turn < self.max_turns:
            if self.should_stop is not None and self.should_stop():
                revision += 1
                record(AgentEvent(
                    event_id=uuid4().hex, run_id=run_id, revision=revision,
                    type="progress", payload={
                        "stage": "user_stop",
                        "label": "已按用户要求停止，当前进度已保留",
                    },
                ))
                return OrchestrationResult(
                    status="execution_incomplete",
                    code="USER_STOPPED",
                    content=content,
                    events=events,
                )
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
                        request_messages = compact_request_messages(
                            "上一轮请求被模型服务拒绝（HTTP 400）。请从工作簿当前已保存进度继续，仅调用一个符合工具 schema 的受控工具；不要重复已完成写入。"
                        )
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
                    # 超时/服务不可用/限流属于瞬时错误：只重试一次即快速失败，
                    # 由上层保留进度并标记待确认，避免一个问题等待数分钟。
                    transient_failure = (
                        "暂时不可用" in last_provider_error
                        or "过于频繁" in last_provider_error
                    )
                    attempt_limit = TRANSIENT_MODEL_RETRY_LIMIT if transient_failure else MODEL_RESPONSE_RETRY_LIMIT
                    if attempt < attempt_limit:
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
                    request_messages = compact_request_messages(
                        "上一轮模型返回空响应。请从工作簿当前已保存进度继续："
                        "此前已成功读取的全部数据在上下文的“已读取数据存档”中，直接使用即可，不要重读。"
                        "若草稿尚未创建先调用 prepare_workbook_copy；否则立即基于存档数据调用"
                        " apply_source_cells 分批写入（每批最多500格，可多批）；不要重复已完成写入。"
                    )
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
                # 写入型任务中，模型只输出文字（常是权衡/自言自语）却尚未
                # 写过任何数据时，这大概率是中间思考而非最终答复：注入催促
                # 消息要求它继续调用工具，最多催促 NO_TOOL_CALL_NUDGE_LIMIT
                # 次之后才接受为最终答复。只读任务（require_writes=False）
                # 保持原行为，纯文字答复即视为完成。
                if (
                    self.require_writes
                    and not has_written
                    and no_tool_call_nudges < NO_TOOL_CALL_NUDGE_LIMIT
                ):
                    no_tool_call_nudges += 1
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="progress", payload={
                            "stage": "nudge_no_tool_call",
                            "label": "模型仅输出权衡文字且尚未写入，已要求其继续执行",
                            "nudge": no_tool_call_nudges,
                        },
                    ))
                    recent_turns.append([{
                        "role": "assistant", "content": response.content or "",
                    }, {
                        "role": "user",
                        "content": (
                            "以上内容是中间思考，不能作为任务结束。请立即继续执行："
                            "如需核对数据就调用读取工具；核对充分后必须调用写入工具完成计划中的写入步骤"
                            "（apply_source_cells 每批最多500格，可分批多次调用）。"
                            "只有当计划的所有写入步骤均已实际完成后，才允许输出不含工具调用的最终总结。"
                        ),
                    }])
                    conversation = assemble_conversation()
                    continue
                if self.require_writes and not has_written:
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex,
                        run_id=run_id,
                        revision=revision,
                        type="needs_user_input",
                        payload={
                            "code": "NO_WRITES_PERFORMED",
                            "detail": "模型未调用任何数据写入工具，已保留当前进度供自动续跑",
                        },
                    ))
                    return OrchestrationResult(
                        status="execution_incomplete",
                        code="NO_WRITES_PERFORMED",
                        content=content,
                        events=events,
                    )
                if (
                    self.require_writes
                    and has_written
                    and _prose_declares_pending_work(response.content)
                ):
                    # A previous write is not proof that the whole plan is
                    # complete. Keep the durable workbook checkpoint and ask
                    # the model to continue with the remaining steps.
                    if no_tool_call_nudges < NO_TOOL_CALL_NUDGE_LIMIT:
                        no_tool_call_nudges += 1
                        revision += 1
                        record(AgentEvent(
                            event_id=uuid4().hex, run_id=run_id,
                            revision=revision, type="progress", payload={
                                "stage": "nudge_incomplete_summary",
                                "label": "模型输出显示仍有后续步骤，已要求继续执行",
                                "nudge": no_tool_call_nudges,
                            },
                        ))
                        recent_turns.append([{
                            "role": "assistant", "content": response.content or "",
                        }, {
                            "role": "user", "content": (
                                "以上内容显示任务仍未完成。请立即继续调用受控工具，"
                                "完成计划中尚未执行的写入和校验步骤；只有全部步骤实际完成后，"
                                "才可以输出最终总结。"
                            ),
                        }])
                        conversation = assemble_conversation()
                        continue
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="needs_user_input", payload={
                            "code": "INCOMPLETE_MODEL_RESPONSE",
                            "detail": "模型输出仍显示有未完成步骤，已保留当前写入进度供自动续跑",
                        },
                    ))
                    return OrchestrationResult(
                        status="execution_incomplete",
                        code="INCOMPLETE_MODEL_RESPONSE",
                        content=content,
                        events=events,
                    )
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
            wrote_this_turn = False
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
                # 数据读取工具的重复调用拦截：同一数据范围在回放窗口内重复出现时
                # 提示模型换范围或开始写入；但回放窗口（最近 N 轮）之外的
                # 旧读取对模型已经“失忆”，此时允许重读——拦住它只会逼模型
                # 去读更多新范围，形成“读了又忘、忘了再读新”的死循环。
                # inspect/find_table 等结构探查结果很小，并且是空响应恢复后确认
                # 草稿状态的必要动作，永远不进入该去重拦截。
                recorded_turn = read_fingerprints_since_write.get(fingerprint)
                stale_read = (
                    call.name in READ_DATA_TOOL_NAMES
                    and recorded_turn is not None
                    and turn - recorded_turn < RECENT_CONVERSATION_TURN_LIMIT
                )
                if stale_read:
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
                    if self.read_call_limit is not None and successful_reads >= self.read_call_limit:
                        # 读取阶段已结束：不要再引导“改读新范围”，
                        # 直接把模型推向写入，否则它会去读更多新范围耗尽轮数。
                        detail = (
                            f"该范围已读取过且读取阶段已结束（已成功读取 {successful_reads} 次）。"
                            "全部已读数据在上下文“已读取数据存档”中：立即调用写入工具"
                            "（apply_source_cells 每批最多500格，可多批）完成计划写入；"
                            "确实缺少必要数据时，直接输出缺口清单结束。"
                        )
                    else:
                        detail = "该读取范围已读取过，请改读尚未核对的范围，执行写入/校验，或明确说明未完成事项"
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="tool_result", payload={"call_id": call.call_id, "name": call.name, "status": "failed", "error": detail},
                    ))
                    turn_messages.append({
                        "role": "tool", "tool_call_id": call.call_id, "name": call.name,
                        "content": json.dumps({"error": detail}, ensure_ascii=False),
                    })
                    read_fingerprints_since_write.pop(fingerprint, None)
                    repeated_call_count = 0
                    previous_call_fingerprint = None
                    continue
                if repeated_call_count >= repeat_limit:
                    revision += 1
                    read_only_repeat = call.name not in WRITE_TOOL_NAMES
                    if read_only_repeat:
                        detail = (
                            "该读取或结构检查请求已连续重复，未再次执行；"
                            "请立即基于已读取数据调用写入工具，或明确列出缺少的必要数据后结束"
                        )
                        record(AgentEvent(
                            event_id=uuid4().hex,
                            run_id=run_id,
                            revision=revision,
                            type="tool_call",
                            payload={"call_id": call.call_id, "name": call.name},
                        ))
                        revision += 1
                        record(AgentEvent(
                            event_id=uuid4().hex,
                            run_id=run_id,
                            revision=revision,
                            type="tool_result",
                            payload={
                                "call_id": call.call_id,
                                "name": call.name,
                                "status": "failed",
                                "error": detail,
                            },
                        ))
                        turn_messages.append({
                            "role": "tool",
                            "tool_call_id": call.call_id,
                            "name": call.name,
                            "content": json.dumps({"error": detail}, ensure_ascii=False),
                        })
                        repeated_call_count = 0
                        previous_call_fingerprint = None
                        continue
                    detail = (
                        "模型连续重复同一写入请求，已停止以避免重复写入"
                    )
                    record(AgentEvent(
                        event_id=uuid4().hex,
                        run_id=run_id,
                        revision=revision,
                        type="run_blocked",
                        payload={
                            "code": "NO_PROGRESS_DETECTED",
                            "tool": call.name,
                            "detail": detail,
                        },
                    ))
                    return OrchestrationResult(
                        status="execution_incomplete" if read_only_repeat else "blocked",
                        code="NO_PROGRESS_DETECTED",
                        content=content,
                        events=events,
                    )
                revision += 1
                record(AgentEvent(event_id=uuid4().hex, run_id=run_id, revision=revision, type="tool_call", payload={"call_id": call.call_id, "name": call.name}, item_id=None))
                # 读取次数硬上限：读取阶段结束（读够了）后拒绝一切新的读取，
                # 引导模型立即转入写入。只统计真正搬运数据的读取
                # （read_range/read_source_range/search_cells）；
                # inspect/find_table 等结构探查结果小、不烧 token，
                # 不占读取额度，否则结构核对就会把额度提前耗尽。
                counts_toward_read_cap = call.name in READ_DATA_TOOL_NAMES
                if (
                    counts_toward_read_cap
                    and self.read_call_limit is not None
                    and successful_reads >= self.read_call_limit
                ):
                    revision += 1
                    read_cap_detail = (
                        f"读取阶段已结束（已成功读取 {successful_reads} 次，达到上限）。"
                        "不要再读取任何新范围：立即基于已读取核对的数据调用写入工具"
                        "（apply_source_cells 每批最多500格，可多批）完成计划写入；"
                        "确实缺少必要数据时，直接输出缺口清单结束，不要请求读取。"
                    )
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="tool_result", payload={"call_id": call.call_id, "name": call.name, "status": "failed", "error": read_cap_detail},
                    ))
                    turn_messages.append({
                        "role": "tool", "tool_call_id": call.call_id, "name": call.name,
                        "content": json.dumps({"error": read_cap_detail}, ensure_ascii=False),
                    })
                    repeated_call_count = 0
                    previous_call_fingerprint = None
                    continue
                result = self.registry.execute(call)
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
                if (
                    counts_toward_read_cap
                    and result.status == "succeeded"
                    and not archive_read_result(fingerprint, call.call_id, call.name, tool_content)
                ):
                    archive_error = (
                        "本次读取结果会使模型上下文超过安全预算，未把该结果交给模型。"
                        "请缩小行列范围后重读，或先基于已读取数据执行写入。"
                    )
                    result = result.model_copy(update={
                        "status": "failed", "ok": False, "output": None,
                        "observation": None, "error": archive_error,
                    })
                    tool_content = json.dumps({"error": archive_error}, ensure_ascii=False)
                revision += 1
                record(AgentEvent(event_id=uuid4().hex, run_id=run_id, revision=revision, type="tool_result", payload={"call_id": result.call_id, "name": result.name, "status": result.status, "error": result.error}))
                turn_messages.append({
                    "role": "tool",
                    "tool_call_id": result.call_id,
                    "name": result.name,
                    "content": tool_content,
                })
                if call.name in DATA_WRITE_TOOL_NAMES and result.status == "succeeded":
                    read_fingerprints_since_write.clear()
                    read_archive.clear()
                    has_written = True
                    wrote_this_turn = True
                elif call.name in READ_DATA_TOOL_NAMES and result.status == "succeeded":
                    # 只记录真正搬运数据的读取；结构探查允许按需重复。
                    read_fingerprints_since_write[fingerprint] = turn
                    if counts_toward_read_cap:
                        successful_reads += 1
            recent_turns.append(turn_messages)
            # 只读不写防护：写入型任务中连续多轮没有任何成功写入时，
            # 注入催促把模型从“读取-遗忘-再读”循环里拉出来，
            # 要求其立即基于已读数据分批写入或明确报告缺口。
            if self.require_writes:
                if wrote_this_turn:
                    turns_since_write = 0
                else:
                    turns_since_write += 1
                if turns_since_write >= READ_ONLY_TURN_NUDGE_LIMIT:
                    turns_since_write = 0
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="progress", payload={
                            "stage": "nudge_read_only_loop",
                            "label": "模型连续多轮只读取未写入，已要求其立即开始写入",
                        },
                    ))
                    recent_turns.append([{
                        "role": "user",
                        "content": (
                            "你已连续多轮只调用读取工具、没有执行任何写入。"
                            "不要再读取新范围：立即基于已读取核对的数据分批调用写入工具"
                            "（apply_source_cells 每批最多500格，可多批），完成计划中的写入步骤；"
                            "确实缺少必要数据时，停止读取并明确列出缺口，不要继续读表。"
                        ),
                    }])
            conversation = assemble_conversation()
        revision += 1
        record(AgentEvent(event_id=uuid4().hex, run_id=run_id, revision=revision, type="needs_user_input", payload={"code": "MAX_TURNS_EXCEEDED"}))
        return OrchestrationResult(status="execution_incomplete", code="MAX_TURNS_EXCEEDED", content=content, events=events)
