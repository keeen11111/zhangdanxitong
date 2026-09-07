"""Bounded model/tool orchestration for document-driven workbook runs."""
from __future__ import annotations

from datetime import datetime, timezone
import json
import os
import hashlib
from typing import Any, Callable, Mapping, Sequence
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field

from .contracts import AgentEvent, ToolCall, ToolResult
from .model import ModelConfig, ModelProviderError, ModelResponse, OpenAICompatibleProvider, estimate_token_count

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
# Token budget is the primary guard. The character budget remains as a legacy
# upper bound for clients that only expose character metrics.
CONTEXT_INPUT_TOKEN_BUDGET = 24_000
# 工具结果回放上限：必须容得下一次“读全整个目标区域”的结果
# （读取工具单次最多1000格，JSON 约 20-40KB）。截断到低于这个量，
# 模型永远看不到完整数据，会陷入“读不全→再读→遗忘→再读”的循环。
TOOL_RESULT_MAX_CHARS = 40_000
# Small reads remain in the immediate assistant/tool transcript for protocol
# compatibility. Larger reads are represented by their bounded checkpoint even
# before the next write, otherwise the very next request replays the raw range.
INLINE_READ_RESULT_MAX_CHARS = 8_000
MODEL_RESPONSE_RETRY_LIMIT = 3
# Empty responses are usually caused by context or reasoning exhaustion. One
# compact retry is useful; repeating the same large transcript three times
# only reproduces the failure and burns a run's turn budget.
EMPTY_RESPONSE_RETRY_LIMIT = 2
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
    # Explicit failure/incompleteness statements seen in real runs.  These
    # are intentionally phrase-level markers rather than a bare ``未`` or
    # ``缺少`` so that successful reports such as ``未发现公式错误`` do not
    # become false positives.
    "cannot reliably",
    "cannot safely",
    "cannot complete",
    "could not complete",
    "unable to",
    "not complete",
    "not yet complete",
    "incomplete",
    "blocked",
    "must report",
    "未能完成",
    "无法完成",
    "尚未完成",
    "任务未完成",
    "无法可靠",
    "无法安全",
    "无法写入",
    "不能可靠",
    "不能完成",
    "未能写入",
    "尚未写入",
    "尚未处理",
    "仍有未",
    "执行阻塞",
    "阻塞",
    "待处理",
    "待确认",
    "需要人工",
    "执行缺口",
    "写入缺口",
    "不完整",
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


def _provider_error_detail(error: ModelProviderError) -> str:
    """Render bounded provider diagnostics without exposing secrets."""
    message = str(error).strip() or "模型服务请求失败"
    diagnostics = error.diagnostics() if hasattr(error, "diagnostics") else {}
    code = diagnostics.get("error_code") or diagnostics.get("error_type")
    summary = str(diagnostics.get("body_summary") or "").strip()
    if code:
        message = f"{message}（{code}）"
    if summary:
        message = f"{message}：{summary[:500]}"
    return message[:1200]
WRITE_TOOL_NAMES = {
    "prepare_workbook_copy",
    "run_basic_payroll_processor",
    "run_keyuan_workflow",
    "apply_source_cells",
    "apply_formula_divisors",
    "insert_and_copy_row",
    "delete_rows",
    "roll_forward_month",
    "apply_cell_changes",
    "copy_formula_from_reference",
    "rollback_work_item",
    "publish_workbook",
}

# 真正写入业务数据的工具。prepare_workbook_copy 只是创建草稿副本，
# 没有写入任何数据：把它算作“已写入”会让 require_writes 保护失效
# （模型建完草稿就输出纯文字总结也会被判定为完成，0 写入收场）。
DATA_WRITE_TOOL_NAMES = WRITE_TOOL_NAMES - {"prepare_workbook_copy"}


def _successful_mutation_count(name: str, output: Any) -> int:
    """Count physical workbook mutations from tool evidence, fail-closed."""
    if name not in DATA_WRITE_TOOL_NAMES:
        return 0
    if name == "roll_forward_month":
        if not isinstance(output, Mapping) or not output:
            return 0
        deferred = output.get("deferred_rule")
        deferred_updates = deferred.get("updates") if isinstance(deferred, Mapping) else None
        nested_count = len(deferred_updates) if isinstance(deferred_updates, list) else 0
        return 1 + nested_count
    if name in {"run_basic_payroll_processor", "run_keyuan_workflow"}:
        if not isinstance(output, Mapping) or output.get("status") not in {"passed", "needs_review"}:
            return 0
        return max(0, int(output.get("change_count") or 0))
    if isinstance(output, list):
        return len(output)
    if not isinstance(output, Mapping):
        return 0
    for key in ("updates", "changes", "deleted_rows"):
        values = output.get(key)
        if isinstance(values, list):
            return len(values)
    if name == "insert_and_copy_row" and output.get("inserted_row"):
        return 1
    if name == "delete_rows" and output.get("deleted") is True:
        return 1
    if name in {"apply_cell_changes", "copy_formula_from_reference", "rollback_work_item"}:
        return 1 if output.get("target_cell") or output.get("item_id") else 0
    return 0


def _is_business_write(name: str, mutation_count: int, output: Any = None) -> bool:
    """A month-row insertion or rollback is not evidence of current data completion."""
    if name == "roll_forward_month" and isinstance(output, Mapping):
        deferred = output.get("deferred_rule")
        deferred_updates = deferred.get("updates") if isinstance(deferred, Mapping) else None
        return isinstance(deferred_updates, list) and len(deferred_updates) > 0
    return mutation_count > 0 and name not in {
        "roll_forward_month", "rollback_work_item", "publish_workbook",
    }


class ToolExecutionError(RuntimeError):
    """A deliberately safe business error that may be shown to the model."""


class OrchestrationResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    status: str
    code: str | None = None
    content: str = ""
    events: list[AgentEvent] = Field(default_factory=list)
    checkpoints: list[dict[str, Any]] = Field(default_factory=list)
    context_state: dict[str, Any] = Field(default_factory=dict)


def _input_token_budget() -> int:
    try:
        return max(1, int(os.getenv("AGENT_INPUT_TOKEN_BUDGET", str(CONTEXT_INPUT_TOKEN_BUDGET))))
    except ValueError:
        return CONTEXT_INPUT_TOKEN_BUDGET


def _safe_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, default=str, separators=(",", ":"))


def _semantic_checkpoint_content(value: Any) -> str:
    """Return only checkpoint evidence, never a complete range payload.

    Older persisted checkpoints may contain ``values``/``rows`` alongside the
    summary.  Those fields are deliberately removed when a checkpoint is
    reloaded so a resumed run cannot re-inflate its prompt with an old raw
    workbook read.  The bounded ``current_cells`` and ``write_evidence`` are
    the only cell-level data the executor needs to continue safely.
    """
    parsed: Any = value
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except (TypeError, ValueError, json.JSONDecodeError):
            # A malformed/legacy archive is not safe to replay.  Returning a
            # bounded error marker makes the missing evidence explicit and
            # forces the normal source-evidence guard to request a fresh,
            # targeted read instead of leaking an arbitrary raw payload.
            return _safe_json({"checkpoint_unavailable": True})
    if not isinstance(parsed, Mapping):
        return _safe_json(parsed)
    allowed = {
        "checkpoint_id", "version", "tool", "sheet", "filename", "range",
        "row_count", "identity_fields", "current_cells", "data_hash",
        "complete", "omitted", "required_fields_missing", "write_evidence",
        "call_id", "read_fingerprint",
    }
    return _safe_json({key: parsed.get(key) for key in allowed if key in parsed})


def _checkpoint_summary(
    name: str,
    arguments: Mapping[str, Any],
    output: Any,
    *,
    sequence: int,
    requirements: Sequence[Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Create a bounded semantic checkpoint; never retain the raw range payload."""
    raw = output if isinstance(output, Mapping) else {"value": output}
    sheet = str(raw.get("sheet") or arguments.get("sheet") or "")
    area = str(raw.get("range") or arguments.get("range") or "")
    rows = raw.get("values") or raw.get("rows") or raw.get("cells") or raw.get("cached_rows") or []
    if not isinstance(rows, list):
        rows = []
    row_count = len(rows)
    headers = list(rows[0]) if rows and isinstance(rows[0], list) else []
    identity_indexes = [
        index for index, value in enumerate(headers)
        if any(term in str(value).lower() for term in ("姓名", "人员", "工号", "编号", "id"))
    ]
    requirements = list(requirements or [])
    try:
        from openpyxl.utils.cell import range_boundaries
        left, top, _right, _bottom = range_boundaries(area)
    except (TypeError, ValueError):
        left, top = 1, 1
    required_row_indexes: set[int] = set()
    required_column_indexes: set[int] = set()
    import re
    for requirement in requirements:
        for coordinate in requirement.get("source_cells") or []:
            match = re.fullmatch(r"([A-Z]+)([1-9][0-9]{0,6})", str(coordinate))
            if not match:
                continue
            required_row_indexes.add(int(match.group(2)) - top)
            column_number = 0
            for letter in match.group(1):
                column_number = column_number * 26 + ord(letter) - 64
            required_column_indexes.add(column_number - left)
    # Keep a bounded, explicit batch: headers, identity columns, and requested
    # source cells. Arbitrary target rows must survive semantic cropping.
    current_cells: list[dict[str, Any]] = []
    selected_rows = set(range(min(len(rows), 16))) | {
        index for index in required_row_indexes if 0 <= index < len(rows)
    }
    selected_columns = (set(range(len(headers))) if len(rows) <= 64 else set(range(8))) | set(identity_indexes) | {
        index for index in required_column_indexes if index >= 0
    }
    for row_index in sorted(selected_rows):
        row = rows[row_index]
        if not isinstance(row, list):
            continue
        indexes = [index for index in sorted(selected_columns) if index < len(row)]
        for column_index in indexes:
            if len(current_cells) >= 512:
                break
            value = row[column_index] if column_index < len(row) else None
            if value is None:
                continue
            current_cells.append({"row": row_index, "column": column_index, "value": value})
    digest = hashlib.sha256(_safe_json(raw).encode("utf-8")).hexdigest()
    checkpoint_id = f"cp-{sequence:04d}-{digest[:12]}"
    write_evidence: list[dict[str, Any]] = []
    required_fields_missing: list[dict[str, Any]] = []
    for requirement in requirements:
        source_coordinates: list[dict[str, Any]] = []
        for coordinate in requirement.get("source_cells") or []:
            match = re.fullmatch(r"([A-Z]+)([1-9][0-9]{0,6})", str(coordinate))
            value_found = False
            value = None
            if match:
                row_index = int(match.group(2)) - top
                column_number = 0
                for letter in match.group(1):
                    column_number = column_number * 26 + ord(letter) - 64
                column_index = column_number - left
                if 0 <= row_index < len(rows) and isinstance(rows[row_index], list) and 0 <= column_index < len(rows[row_index]):
                    value = rows[row_index][column_index]
                    # ``None`` is a valid, explicit blank source value. A
                    # field is missing only when its coordinate is outside
                    # the returned rectangle.
                    value_found = True
            if not value_found:
                required_fields_missing.append({"field": "source_cell", "coordinate": coordinate})
            source_coordinates.append({"cell": coordinate, "value": value})
        write_evidence.append({
            "source_file": requirement.get("source_file"),
            "source_sheet": requirement.get("source_sheet"),
            "source_coordinates": source_coordinates,
            "target_sheet": requirement.get("target_sheet"),
            "target_coordinate": requirement.get("target_cell"),
            "expected_value": requirement.get("expected_value"),
        })
    omitted = []
    if row_count > 16 or len(_safe_json(raw)) > 8000:
        omitted.append("unrelated_rows_or_columns")
    return {
        "checkpoint_id": checkpoint_id,
        "version": sequence,
        "tool": name,
        "sheet": sheet,
        "filename": str(raw.get("filename") or arguments.get("filename") or ""),
        "range": area,
        "row_count": row_count,
        "identity_fields": [headers[index] for index in identity_indexes if index < len(headers)],
        "current_cells": current_cells,
        "data_hash": digest,
        "complete": not omitted,
        "omitted": omitted,
        "required_fields_missing": required_fields_missing,
        "write_evidence": write_evidence,
    }


def _compact_tool_schemas(
    tools: Sequence[Mapping[str, Any]], *, stage: str | None = None,
) -> list[dict[str, Any]]:
    """Keep only stage-relevant tool contracts and strip schema prose.

    A retry must not carry the full read/write/validation registry.  Tool
    descriptions and JSON-schema annotations are explanatory, not execution
    state, so they are removed while names, properties and required fields are
    retained for strict provider validation.
    """
    read_names = {
        "inspect_workbook", "classify_file", "find_table", "read_range",
        "inspect_source_file", "read_source_range", "search_cells",
        "match_person", "select_sheet_mapping", "prepare_workbook_copy",
    }
    write_names = {
        "apply_source_cells", "apply_formula_divisors", "insert_and_copy_row",
        "delete_rows", "roll_forward_month", "run_basic_payroll_processor",
        "run_keyuan_workflow", "validate_workbook", "validate_with_officecli",
    }
    allowed = read_names if stage in {"execution", "reading"} else write_names if stage in {"writing", "validating"} else None

    def compact_schema(value: Any) -> Any:
        if isinstance(value, Mapping):
            return {
                str(key): compact_schema(item)
                for key, item in value.items()
                if key not in {"description", "title", "default", "examples", "example", "$comment"}
            }
        if isinstance(value, list):
            return [compact_schema(item) for item in value]
        return value

    compacted: list[dict[str, Any]] = []
    for tool in tools:
        if not isinstance(tool, Mapping):
            continue
        function = tool.get("function")
        if not isinstance(function, Mapping):
            continue
        name = str(function.get("name") or "")
        if allowed is not None and name not in allowed:
            continue
        compacted.append({"type": "function", "function": {
            "name": name,
            "parameters": compact_schema(function.get("parameters") or {"type": "object", "properties": {}}),
        }})
    return compacted


def _missing_write_evidence(
    checkpoints: Sequence[Mapping[str, Any]], change: Mapping[str, Any],
) -> list[str]:
    """Return source coordinates absent from the latest semantic read proof."""
    source_file = str(change.get("source_file") or "")
    source_sheet = str(change.get("source_sheet") or "")
    checkpoint = next(
        (
            item for item in reversed(checkpoints)
            if item.get("filename") == source_file and item.get("sheet") == source_sheet
        ),
        None,
    )
    if checkpoint is None:
        return [str(cell) for cell in change.get("source_cells") or []]
    try:
        from openpyxl.utils.cell import range_boundaries
        left, top, _right, _bottom = range_boundaries(str(checkpoint.get("range") or ""))
    except (TypeError, ValueError):
        return [str(cell) for cell in change.get("source_cells") or []]
    indexed = {
        (int(cell.get("row", -1)), int(cell.get("column", -1)))
        for cell in checkpoint.get("current_cells") or []
        if isinstance(cell, Mapping)
    }
    import re
    missing: list[str] = []
    for coordinate in change.get("source_cells") or []:
        match = re.fullmatch(r"([A-Z]+)([1-9][0-9]{0,6})", str(coordinate))
        if not match:
            missing.append(str(coordinate))
            continue
        column = 0
        for letter in match.group(1):
            column = column * 26 + ord(letter) - 64
        key = (int(match.group(2)) - top, column - left)
        if key not in indexed:
            missing.append(str(coordinate))
    return missing


class ToolRegistry:
    """Allow-list of deterministic workbook operations exposed to the model."""

    def __init__(self) -> None:
        self._handlers: dict[str, Callable[..., Any]] = {}

    def register(self, name: str, handler: Callable[..., Any]) -> None:
        if name not in {
            "prepare_workbook_copy", "run_basic_payroll_processor", "run_keyuan_workflow", "apply_source_cells", "search_cells", "apply_formula_divisors", "insert_and_copy_row", "delete_rows",
            "roll_forward_month",
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
        stage: str = "execution",
        checkpoint_sink: Callable[[dict[str, Any]], None] | None = None,
        initial_context_state: Mapping[str, Any] | None = None,
        initial_checkpoints: Sequence[Mapping[str, Any]] | None = None,
        required_validation_tools: Sequence[str] | None = None,
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
        self.stage = stage
        self.checkpoint_sink = checkpoint_sink
        self.initial_context_state = dict(initial_context_state or {})
        self.initial_checkpoints = [
            dict(checkpoint) for checkpoint in (initial_checkpoints or [])
            if isinstance(checkpoint, Mapping)
        ][-20:]
        self.required_validation_tools = {
            str(name) for name in (required_validation_tools or [])
            if str(name) in {"validate_workbook", "validate_with_officecli"}
        }

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
        request_tools = list(tools or [])
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
        validation_nudges = 0
        turns_since_write = 0
        successful_reads = 0
        # 读取数据存档：成功的数据读取结果以 user 消息形式跨轮保留，
        # 不随 RECENT_CONVERSATION_TURN_LIMIT 淘汰。写入阶段（含空响应
        # /HTTP 400 后的紧凑重试）据此仍能看到全部已读数据。
        read_archive: list[dict[str, Any]] = []
        checkpoints: list[dict[str, Any]] = list(self.initial_checkpoints)
        checkpoint_sequence = max(
            (int(checkpoint.get("version") or 0) for checkpoint in checkpoints),
            default=0,
        )
        # When a provider fails after context compaction, failover must start
        # from that compact continuation instead of replaying the stale raw
        # transcript.  Otherwise the backup provider sees the same oversized
        # request and the run falls back into the empty-response loop.
        failover_conversation: list[dict[str, Any]] | None = None
        context_state: dict[str, Any] = {
            "current_stage": self.stage,
            "completed_steps": [], "completed_writes": [], "pending_writes": [],
            "business_write_count": 0,
            "unresolved_items": [], "source_target_mappings": [],
            "validation_failures": [], "last_read_checkpoint": None,
            "last_write_checkpoint": None,
            "required_validation_tools": sorted(self.required_validation_tools),
        }
        for key in context_state:
            if key in self.initial_context_state:
                value = self.initial_context_state[key]
                context_state[key] = list(value) if isinstance(value, list) else value
        has_written = int(context_state.get("business_write_count") or 0) > 0 or any(
            (
                str(entry.get("tool") or "") not in {
                    "roll_forward_month", "rollback_work_item", "publish_workbook",
                    "run_basic_payroll_processor", "run_keyuan_workflow",
                }
                and int(entry.get("count") or 0) > 0
            )
            or bool(entry.get("target_cell") or entry.get("target_coordinate"))
            for entry in context_state["completed_writes"]
            if isinstance(entry, Mapping)
        )
        compact_attempted = False

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
            content = _semantic_checkpoint_content(content)
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

        # A resumable segment must restore the bounded values behind its read
        # checkpoints, not only the range/hash summary. Otherwise the model is
        # forced to read the workbook again before it can write accurately.
        for checkpoint in checkpoints:
            archive_content = checkpoint.get("archive_content")
            fingerprint = str(checkpoint.get("read_fingerprint") or "")
            if not archive_content or not fingerprint:
                continue
            if archive_read_result(
                fingerprint,
                str(checkpoint.get("call_id") or checkpoint.get("checkpoint_id") or ""),
                str(checkpoint.get("tool") or "read_range"),
                str(archive_content),
            ):
                read_fingerprints_since_write[fingerprint] = 0
                successful_reads += 1

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

        conversation = assemble_conversation()

        def compact_request_messages(note: str) -> list[dict[str, Any]]:
            """Build a token-fitting continuation without replaying transcript.

            Character slicing is unsafe for Chinese and for JSON-heavy tool
            schemas.  We reserve space for the current tool definitions,
            retain only the latest checkpoints, then progressively shorten the
            initial messages.  The final candidate is measured with the same
            estimator used by the request guard; if it still cannot fit, the
            caller fails closed instead of sending another oversized request.
            """
            budget = _input_token_budget()
            tool_tokens = estimate_token_count(_safe_json(request_tools))
            available = max(0, budget - tool_tokens - 64)

            def build(checkpoint_count: int, initial_limit: int) -> list[dict[str, Any]]:
                bounded_initial: list[dict[str, Any]] = []
                remaining = initial_limit
                for message in initial_messages:
                    if remaining <= 0:
                        break
                    role = str(message.get("role") or "user")
                    original = str(message.get("content") or "")
                    if estimate_token_count(original) <= remaining:
                        content = original
                    else:
                        # Binary search the largest prefix that fits the
                        # remaining token allowance, preserving role and a
                        # machine-readable truncation marker.
                        lo, hi = 0, len(original)
                        marker = "\n[initial context compacted]"
                        while lo < hi:
                            mid = (lo + hi + 1) // 2
                            if estimate_token_count(original[:mid] + marker) <= remaining:
                                lo = mid
                            else:
                                hi = mid - 1
                        content = original[:lo] + marker if lo else marker
                    bounded_initial.append({"role": role, "content": content})
                    remaining -= estimate_token_count(content)
                def bound_items(values: Any, limit: int) -> list[Any]:
                    items = list(values or []) if isinstance(values, list) else []
                    bounded: list[Any] = []
                    for item in items[-limit:]:
                        if isinstance(item, Mapping):
                            bounded.append({
                                str(key): str(value)[:500]
                                for key, value in item.items()
                                if key in {
                                    "tool", "count", "source_file", "source_sheet",
                                    "target_sheet", "target_cell", "target_coordinate",
                                    "source_cells", "new_value", "expected_value",
                                    "id", "person_key", "status", "error",
                                }
                            })
                        else:
                            bounded.append(str(item)[:500])
                    return bounded

                state = {
                    "current_stage": context_state.get("current_stage"),
                    "completed_steps": bound_items(context_state.get("completed_steps"), 30),
                    "completed_writes": bound_items(context_state.get("completed_writes"), 20),
                    "pending_writes": bound_items(context_state.get("pending_writes"), 40),
                    "unresolved_items": bound_items(context_state.get("unresolved_items"), 30),
                    "source_target_mappings": bound_items(context_state.get("source_target_mappings"), 30),
                    "validation_failures": bound_items(context_state.get("validation_failures"), 15),
                    "last_read_checkpoint": str(context_state.get("last_read_checkpoint") or "")[:128],
                    "last_write_checkpoint": str(context_state.get("last_write_checkpoint") or "")[:128],
                }
                checkpoint_payload = []
                if checkpoint_count:
                    for entry in checkpoints[-checkpoint_count:]:
                        # ``archive_content`` is a server-side resume aid and
                        # may contain a large legacy payload.  Never embed it
                        # in a compact model request; the semantic fields are
                        # sufficient for the executor and source-evidence gate.
                        compact_entry = {
                            key: value for key, value in entry.items()
                            if key != "archive_content"
                        }
                        checkpoint_payload.append(compact_entry)
                payload = {"event": "compact_context", "note": note, "state": state,
                           "checkpoints": checkpoint_payload}
                return [*bounded_initial, {"role": "user", "content": _safe_json(payload)}]

            # Checkpoint evidence is more valuable than old prose.  Try several
            # progressively smaller evidence sets before dropping initial text.
            candidates: list[list[dict[str, Any]]] = []
            for checkpoint_count in (8, 4, 2, 1, 0):
                # Reserve the state/checkpoint message first.  The previous
                # implementation gave the initial messages the entire
                # allowance and appended state afterwards, so a compact
                # request could still exceed the budget.
                base = build(checkpoint_count, 0)
                base_tokens = estimate_token_count(_safe_json(base))
                initial_allowance = max(0, available - base_tokens)
                candidate = build(checkpoint_count, initial_allowance)
                # If token-estimation rounding still leaves it over budget,
                # shrink the initial prefix with a binary search.
                if estimate_token_count(_safe_json(candidate)) + tool_tokens > budget:
                    lo, hi = 0, initial_allowance
                    while lo < hi:
                        mid = (lo + hi) // 2
                        trial = build(checkpoint_count, mid)
                        if estimate_token_count(_safe_json(trial)) + tool_tokens <= budget:
                            lo = mid + 1
                        else:
                            hi = mid
                    candidate = build(checkpoint_count, max(0, lo - 1))
                candidates.append(candidate)
            for candidate in candidates:
                if estimate_token_count(_safe_json(candidate)) + tool_tokens <= budget:
                    return candidate
            # Last-resort minimal continuation: latest checkpoint/state only.
            # The state-only message is guaranteed to be smaller than any
            # candidate carrying initial prose.  The caller will fail closed
            # only when even this machine-readable continuation cannot fit.
            return build(0, 0)

        def reset_transcript_after_write() -> None:
            """Drop stale read transcripts after a durable write checkpoint.

            The assistant/tool transcript becomes obsolete after an atomic
            write, but bounded read archives remain available for later write
            batches. This avoids replaying protocol noise without discarding
            the source values that subsequent writes still require.
            """
            recent_turns.clear()

        def context_metrics(request_messages: Sequence[Mapping[str, Any]], provider_obj: Any) -> dict[str, Any]:
            input_chars = len(_safe_json(list(request_messages)))
            tool_chars = len(_safe_json(request_tools))
            config = getattr(provider_obj, "config", None)
            return {
                "input_chars": input_chars,
                "estimated_input_tokens": estimate_token_count(_safe_json(list(request_messages))),
                "tool_schema_chars": tool_chars,
                "estimated_tool_tokens": estimate_token_count(_safe_json(request_tools)),
                "max_output_tokens": getattr(config, "max_output_tokens", None),
                "provider": getattr(config, "provider", provider_obj.__class__.__name__),
                "model": getattr(config, "model", "unknown"),
                "stage": context_state.get("current_stage", self.stage),
                "recent_tool_result_chars": len(str(read_archive[-1]["message"].get("content") or "")) if read_archive else 0,
                "input_token_budget": _input_token_budget(),
            }

        def record(event: AgentEvent) -> None:
            """Expose an event before the next potentially slow model turn."""
            events.append(event)
            if on_event is not None:
                on_event(event)

        def validation_passed(name: str, output: Any) -> bool:
            if not isinstance(output, Mapping):
                return False
            if name == "validate_with_officecli":
                return (
                    output.get("valid") is True
                    and output.get("recalculated") is True
                    and not output.get("formula_errors")
                    and not output.get("unevaluated_formulas")
                )
            if name == "validate_workbook":
                return (
                    output.get("readable") is True
                    and int(output.get("unresolved_item_count") or 0) == 0
                    and not output.get("formula_errors")
                    and not output.get("summary_range_errors")
                    and not output.get("duplicate_identities")
                    and not output.get("identity_errors")
                    and (output.get("can_publish") is True or output.get("passed") is True)
                )
            return False

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
            request_messages = failover_conversation or conversation
            failover_conversation = None
            compact_attempted = False
            switched_provider = False
            for attempt in range(1, MODEL_RESPONSE_RETRY_LIMIT + 1):
                metrics = context_metrics(request_messages, provider)
                request_correlation_id = uuid4().hex
                setter = getattr(provider, "set_request_id", None)
                if callable(setter):
                    setter(request_correlation_id)
                metrics["request_id"] = request_correlation_id
                revision += 1
                record(AgentEvent(
                    event_id=uuid4().hex,
                    run_id=run_id,
                    revision=revision,
                    type="model_request",
                    payload={**metrics, **{
                        # Legacy execution consumers use ``model_call`` as
                        # the event stage; explicit planner/executor stages
                        # remain visible for stage-aware runs.
                        "stage": "model_call" if self.stage == "execution" and context_state.get("current_stage") == "execution" else metrics["stage"],
                        "operation": "model_call",
                        "turn": turn,
                        "attempt": attempt,
                        "request_id": request_correlation_id,
                        "label": "正在请求 Agent 分析并等待受控操作指令",
                    }},
                ))
                if metrics["estimated_input_tokens"] + metrics["estimated_tool_tokens"] > _input_token_budget():
                    if compact_attempted:
                        revision += 1
                        record(AgentEvent(
                            event_id=uuid4().hex, run_id=run_id, revision=revision,
                            type="needs_user_input", payload={
                                "code": "CONTEXT_BUDGET_EXCEEDED",
                                "detail": "上下文压缩后仍超过输入预算，已停止读取并保留 checkpoint",
                                **metrics,
                            },
                        ))
                        return OrchestrationResult(
                            status="execution_incomplete", code="CONTEXT_BUDGET_EXCEEDED",
                            content=content, events=events, checkpoints=checkpoints,
                            context_state=context_state,
                        )
                    compact_attempted = True
                    request_tools = _compact_tool_schemas(
                        request_tools, stage=context_state.get("current_stage"),
                    )
                    request_messages = compact_request_messages("请求超过输入 token 预算，已压缩上下文")
                    continue
                try:
                    response: ModelResponse = provider.complete(messages=request_messages, tools=request_tools)
                except ModelProviderError as exc:
                    last_provider_error = (
                        f"{_provider_error_detail(exc)}"
                        if isinstance(exc, ModelProviderError)
                        else str(exc).strip() or "模型服务请求失败"
                    )
                    diagnostics = exc.diagnostics() if hasattr(exc, "diagnostics") else {}
                    if getattr(exc, "status_code", None) == 400 and not compact_attempted:
                        compact_attempted = True
                        request_tools = _compact_tool_schemas(
                            request_tools, stage=context_state.get("current_stage"),
                        )
                        request_messages = compact_request_messages("模型服务返回 HTTP 400，已压缩上下文后重试")
                        continue
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
                        failover_conversation = list(request_messages)
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
                            "diagnostics": diagnostics,
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
                        checkpoints=checkpoints, context_state=context_state,
                    )
                revision += 1
                response_meta = context_metrics(request_messages, provider)
                record(AgentEvent(
                    event_id=uuid4().hex,
                    run_id=run_id,
                    revision=revision,
                    type="model_response",
                    payload={
                        **response_meta,
                        "request_id": request_correlation_id,
                        "provider_request_id": getattr(provider, "last_request_meta", {}).get("provider_request_id") or getattr(response, "provider_request_id", None) or response.request_id,
                        "finish_reason": response.finish_reason,
                        "input_tokens": response.input_tokens,
                        "output_tokens": response.output_tokens,
                        "tool_call_count": len(response.tool_calls),
                        "has_tool_calls": bool(response.tool_calls),
                        "content_empty": not bool(response.content.strip()),
                        "reached_output_limit": response.finish_reason in {"length", "max_tokens", "incomplete"},
                        "content": response.content[:4000],
                        "checkpoint_id": context_state.get("last_read_checkpoint"),
                    },
                ))
                if response.finish_reason in {"length", "max_tokens", "incomplete"} and not compact_attempted:
                    compact_attempted = True
                    request_tools = _compact_tool_schemas(
                        request_tools, stage=context_state.get("current_stage"),
                    )
                    request_messages = compact_request_messages("模型输出达到上限，已压缩上下文并继续当前阶段")
                    continue
                if response.finish_reason in {"length", "max_tokens", "incomplete"}:
                    revision += 1
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="needs_user_input", payload={
                            "code": "CONTEXT_BUDGET_EXCEEDED",
                            "detail": "压缩后模型仍达到输出上限，已停止并保留当前 checkpoint",
                        },
                    ))
                    return OrchestrationResult(
                        status="execution_incomplete", code="CONTEXT_BUDGET_EXCEEDED",
                        content=content, events=events, checkpoints=checkpoints,
                        context_state=context_state,
                    )
                if response.content.strip() or response.tool_calls:
                    break
                # An empty response is commonly caused by the provider
                # exhausting context/reasoning budget.  Retrying the full
                # tool transcript only reproduces that condition and makes a
                # run appear stuck.  Workbook writes are already durable, so
                # retry from the original task with a compact continuation.
                if not compact_attempted:
                    compact_attempted = True
                    request_tools = _compact_tool_schemas(
                        request_tools, stage=context_state.get("current_stage"),
                    )
                    request_messages = compact_request_messages(
                        "上一轮模型返回空响应。请从工作簿当前已保存进度继续："
                        "此前已成功读取的全部数据在上下文的“已读取数据存档”中，直接使用即可，不要重读。"
                        "若草稿尚未创建先调用 prepare_workbook_copy；否则立即基于存档数据调用"
                        " apply_source_cells 分批写入（每批最多500格，可多批）；不要重复已完成写入。"
                    )
                    continue
                else:
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
                        failover_conversation = list(request_messages)
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
                        checkpoints=checkpoints,
                        context_state=context_state,
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
                        checkpoints=checkpoints, context_state=context_state,
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
                        checkpoints=checkpoints, context_state=context_state,
                    )
                missing_validations = self.required_validation_tools - set(context_state["completed_steps"])
                if self.require_writes and has_written and missing_validations:
                    if context_state["validation_failures"]:
                        revision += 1
                        detail = "工作簿机器校验未通过，不能标记为完成"
                        record(AgentEvent(
                            event_id=uuid4().hex, run_id=run_id, revision=revision,
                            type="needs_user_input", payload={
                                "code": "VALIDATION_FAILED", "detail": detail,
                                "failures": context_state["validation_failures"][-10:],
                            },
                        ))
                        return OrchestrationResult(
                            status="execution_incomplete", code="VALIDATION_FAILED",
                            content=content, events=events, checkpoints=checkpoints,
                            context_state=context_state,
                        )
                    if validation_nudges < NO_TOOL_CALL_NUDGE_LIMIT:
                        validation_nudges += 1
                        revision += 1
                        record(AgentEvent(
                            event_id=uuid4().hex, run_id=run_id, revision=revision,
                            type="progress", payload={
                                "stage": "nudge_validation",
                                "label": "实际写入已完成，正在要求执行机器校验",
                                "missing_tools": sorted(missing_validations),
                            },
                        ))
                        recent_turns.append([{
                            "role": "assistant", "content": response.content or "",
                        }, {
                            "role": "user", "content": (
                                "写入记录不是完成凭据。请依次调用以下机器校验工具并确认返回通过："
                                + "、".join(sorted(missing_validations))
                                + "。禁止用自然语言代替校验。"
                            ),
                        }])
                        conversation = assemble_conversation()
                        continue
                    revision += 1
                    detail = "缺少 Skill 要求的实际校验工具记录，不能标记为完成"
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="needs_user_input", payload={
                            "code": "VALIDATION_NOT_PERFORMED", "detail": detail,
                            "missing_tools": sorted(missing_validations),
                        },
                    ))
                    return OrchestrationResult(
                        status="execution_incomplete", code="VALIDATION_NOT_PERFORMED",
                        content=content, events=events, checkpoints=checkpoints,
                        context_state=context_state,
                    )
                context_state["current_stage"] = "completed"
                return OrchestrationResult(status="completed", content=content, events=events,
                                            checkpoints=checkpoints, context_state=context_state)
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
                # 数据读取工具的重复调用拦截：Run 的输入由哈希固定，成功
                # 读取的同一范围在发生写入前不会自行变化，因此指纹不能按
                # 回放窗口过期。完整读取值已进入受限存档，重复请求只返回
                # 结构化反馈，避免“拒绝一次、执行一次”的交替死循环。
                # inspect/find_table 等结构探查结果很小，并且是空响应恢复后确认
                # 草稿状态的必要动作，永远不进入该去重拦截。
                recorded_turn = read_fingerprints_since_write.get(fingerprint)
                stale_read = (
                    call.name in READ_DATA_TOOL_NAMES
                    and recorded_turn is not None
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
                    and context_state.get("current_stage") in {"writing", "validating"}
                ):
                    revision += 1
                    detail = (
                        f"当前已进入 {context_state['current_stage']} 阶段，禁止重新读取数据范围。"
                        "请使用已保存的读取 checkpoint 完成剩余写入，或调用工作簿校验工具；"
                        "若证据不足，明确列出缺失坐标并以 execution_incomplete 结束。"
                    )
                    record(AgentEvent(
                        event_id=uuid4().hex, run_id=run_id, revision=revision,
                        type="tool_result", payload={
                            "call_id": call.call_id, "name": call.name,
                            "status": "failed", "code": "READ_FORBIDDEN_AFTER_WRITING",
                            "error": detail,
                        },
                    ))
                    turn_messages.append({
                        "role": "tool", "tool_call_id": call.call_id, "name": call.name,
                        "content": json.dumps({
                            "code": "READ_FORBIDDEN_AFTER_WRITING", "error": detail,
                            "last_read_checkpoint": context_state.get("last_read_checkpoint"),
                            "last_write_checkpoint": context_state.get("last_write_checkpoint"),
                        }, ensure_ascii=False),
                    })
                    repeated_call_count = 0
                    previous_call_fingerprint = None
                    continue
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
                if call.name == "apply_source_cells":
                    missing_coordinates: list[str] = []
                    for change in call.arguments.get("changes") or []:
                        if isinstance(change, Mapping):
                            missing_coordinates.extend(_missing_write_evidence(checkpoints, change))
                    if missing_coordinates:
                        detail = (
                            "写入前缺少来源读取证据，已拒绝本批写入；请先读取并核对坐标："
                            + ", ".join(dict.fromkeys(missing_coordinates))
                        )
                        result = ToolResult(
                            call_id=call.call_id, name=call.name, status="failed", ok=False,
                            error=detail,
                        )
                    else:
                        result = self.registry.execute(call)
                else:
                    result = self.registry.execute(call)
                if call.name in {"validate_workbook", "validate_with_officecli"}:
                    passed = result.status == "succeeded" and validation_passed(call.name, result.output)
                    if passed:
                        if call.name not in context_state["completed_steps"]:
                            context_state["completed_steps"].append(call.name)
                    else:
                        failure = {
                            "tool": call.name,
                            "error": result.error or "校验工具返回未通过",
                        }
                        context_state["validation_failures"].append(failure)
                        if result.status == "succeeded":
                            result = result.model_copy(update={
                                "status": "failed", "ok": False,
                                "error": "校验工具返回未通过",
                            })
                raw_tool_content = _safe_json(
                    result.output if result.status == "succeeded" else {"error": result.error}
                )
                archive_content = raw_tool_content
                if counts_toward_read_cap and result.status == "succeeded":
                    checkpoint_sequence += 1
                    checkpoint = _checkpoint_summary(call.name, call.arguments, result.output,
                                                      sequence=checkpoint_sequence,
                                                      requirements=context_state.get("pending_writes") or [])
                    # Keep the complete read result on the server only long
                    # enough to derive the bounded checkpoint.  The model is
                    # never sent the original values/rows payload; it receives
                    # headers, identity fields, selected cells, hashes and
                    # explicit write evidence instead.
                    archive_content = _semantic_checkpoint_content(checkpoint)
                    checkpoints.append(checkpoint)
                    context_state["current_stage"] = "reading"
                    context_state["last_read_checkpoint"] = checkpoint["checkpoint_id"]
                    checkpoint["call_id"] = call.call_id
                    checkpoint["read_fingerprint"] = fingerprint
                    checkpoint["archive_content"] = archive_content
                    # Reads larger than the tool-result boundary cannot be
                    # shown faithfully. Reject them and require a smaller
                    # range instead of silently dropping tail cells.
                    tool_content = archive_content
                else:
                    tool_content = raw_tool_content
                    if len(tool_content) > TOOL_RESULT_MAX_CHARS:
                        tool_content = tool_content[:TOOL_RESULT_MAX_CHARS] + "\n[结果已裁剪]"
                if len(tool_content) > TOOL_RESULT_MAX_CHARS:
                    tool_content = tool_content[:TOOL_RESULT_MAX_CHARS] + "\n[结果已裁剪]"
                if (
                    counts_toward_read_cap
                    and result.status == "succeeded"
                    and (
                        len(archive_content) > TOOL_RESULT_MAX_CHARS
                        or not archive_read_result(fingerprint, call.call_id, call.name, archive_content)
                    )
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
                    failed_checkpoint = checkpoints.pop()
                    prior_read = next(
                        (
                            item for item in reversed(checkpoints)
                            if item.get("stage") != "writing"
                        ),
                        None,
                    )
                    context_state["last_read_checkpoint"] = (
                        prior_read.get("checkpoint_id") if prior_read else None
                    )
                    if failed_checkpoint.get("checkpoint_id") == context_state.get("last_write_checkpoint"):
                        context_state["last_write_checkpoint"] = None
                elif counts_toward_read_cap and result.status == "succeeded":
                    if self.checkpoint_sink is not None:
                        self.checkpoint_sink(checkpoints[-1])
                mutation_count = (
                    _successful_mutation_count(call.name, result.output)
                    if result.status == "succeeded" else 0
                )
                business_write = (
                    _is_business_write(call.name, mutation_count, result.output)
                    if mutation_count > 0 else False
                )
                revision += 1
                result_payload = {"call_id": result.call_id, "name": result.name,
                                  "status": result.status, "error": result.error}
                if mutation_count > 0:
                    result_payload.update({
                        "write_count": mutation_count,
                        "business_write": business_write,
                    })
                if counts_toward_read_cap and result.status == "succeeded" and checkpoints:
                    result_payload.update({
                        "checkpoint_id": checkpoints[-1]["checkpoint_id"],
                        "checkpoint_version": checkpoints[-1]["version"],
                        "sheet": checkpoints[-1]["sheet"],
                        "range": checkpoints[-1]["range"],
                        "row_count": checkpoints[-1]["row_count"],
                        "result_chars": len(raw_tool_content),
                    })
                record(AgentEvent(event_id=uuid4().hex, run_id=run_id, revision=revision,
                                  type="tool_result", payload=result_payload))
                turn_messages.append({
                    "role": "tool",
                    "tool_call_id": result.call_id,
                    "name": result.name,
                    "content": tool_content,
                })
                if mutation_count > 0:
                    read_fingerprints_since_write.clear()
                    context_state["completed_steps"] = [
                        step for step in context_state["completed_steps"]
                        if step not in self.required_validation_tools
                    ]
                    context_state["validation_failures"] = []
                    has_written = has_written or business_write
                    if business_write:
                        context_state["business_write_count"] = (
                            int(context_state.get("business_write_count") or 0) + mutation_count
                        )
                    wrote_this_turn = True
                    context_state["current_stage"] = "writing"
                    context_state["completed_writes"].append({
                        "tool": call.name,
                        "count": mutation_count,
                    })
                    checkpoint_sequence += 1
                    write_checkpoint = {
                        "checkpoint_id": f"cp-{checkpoint_sequence:04d}-write-{turn}",
                        "version": checkpoint_sequence,
                        "tool": call.name,
                        "stage": "writing",
                        "status": "succeeded",
                        "write_count": mutation_count,
                        "business_write": business_write,
                    }
                    checkpoints.append(write_checkpoint)
                    context_state["last_write_checkpoint"] = write_checkpoint["checkpoint_id"]
                    if self.checkpoint_sink is not None:
                        self.checkpoint_sink(write_checkpoint)
                elif call.name in {"validate_workbook", "validate_with_officecli"} and result.status == "succeeded":
                    context_state["current_stage"] = "validating"
                elif call.name in READ_DATA_TOOL_NAMES and result.status == "succeeded":
                    # 只记录真正搬运数据的读取；结构探查允许按需重复。
                    read_fingerprints_since_write[fingerprint] = turn
                    if counts_toward_read_cap:
                        successful_reads += 1
            if wrote_this_turn:
                # Do not replay the read transcript after a write. A compact
                # checkpoint is sufficient and preserves the assistant/tool
                # protocol for the just-completed write call.
                reset_transcript_after_write()
                recent_turns.append([{
                    "role": "user",
                    "content": _safe_json({
                        "event": "write_checkpoint",
                        "checkpoint_id": context_state.get("last_write_checkpoint"),
                        "completed_writes": context_state.get("completed_writes", [])[-1:],
                        "instruction": "草稿已原子保存。继续执行尚未完成的写入或验证，不要重读已完成范围。",
                    }),
                }])
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
        return OrchestrationResult(status="execution_incomplete", code="MAX_TURNS_EXCEEDED", content=content,
                                   events=events, checkpoints=checkpoints, context_state=context_state)
