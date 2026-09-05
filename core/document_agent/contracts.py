"""Pydantic contracts and deterministic rules used by the Excel Agent."""
from __future__ import annotations

import math
from decimal import Decimal
from datetime import datetime, timezone
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


DecisionAction = Literal["apply_proposed", "keep_current", "skip"]
WorkItemStatus = Literal["pending", "needs_review", "resolved", "skipped", "failed", "auto_applied"]

ToolName = Literal[
    "prepare_workbook_copy", "apply_source_cells", "search_cells", "apply_formula_divisors",
    "inspect_workbook", "classify_file", "find_table", "read_range",
    "inspect_source_file", "read_source_range",
    "match_person", "propose_changes", "apply_cell_changes",
    "select_sheet_mapping",
    "copy_formula_from_reference", "validate_workbook", "validate_with_officecli", "rollback_work_item",
    "publish_workbook",
    "insert_and_copy_row", "delete_rows",
    "run_basic_payroll_processor", "run_keyuan_workflow",
]
ToolCallStatus = Literal["pending", "running", "succeeded", "failed"]
AgentEventType = Literal[
    "run_started", "model_request", "model_response", "tool_call", "tool_result",
    "work_item_updated", "user_message", "assistant_message", "validation",
    "needs_user_input", "run_blocked", "run_completed", "run_failed", "progress",
]


class ToolCall(BaseModel):
    """A model-requested invocation of one allow-listed workbook tool."""

    model_config = ConfigDict(extra="forbid")

    call_id: str = Field(min_length=1, max_length=128)
    name: ToolName
    arguments: dict[str, Any] = Field(default_factory=dict)


class ToolResult(BaseModel):
    """The sanitized observation returned to the model after a tool call."""

    model_config = ConfigDict(extra="forbid")

    call_id: str = Field(min_length=1, max_length=128)
    name: ToolName
    status: ToolCallStatus = "succeeded"
    output: Any = None
    # Compatibility aliases used by the event API.  ``ok``/``observation``
    # keep the wire contract readable while status/output remain canonical.
    ok: bool | None = None
    observation: Any = None

    @model_validator(mode="after")
    def _normalise_aliases(self) -> "ToolResult":
        if self.ok is False and self.status == "succeeded":
            self.status = "failed"
        if self.ok is True and self.status == "failed":
            self.status = "succeeded"
        if self.output is None and self.observation is not None:
            self.output = self.observation
        if self.observation is None and self.output is not None:
            self.observation = self.output
        if self.status == "failed" and not self.error:
            self.error = "工具执行失败"
        return self
    error: str | None = Field(default=None, max_length=2000)


class ModelTrace(BaseModel):
    """Auditable model exchange metadata. Prompt and credentials are never stored."""

    model_config = ConfigDict(extra="forbid")

    provider: str = Field(min_length=1, max_length=128)
    model: str = Field(min_length=1, max_length=256)
    request_id: str | None = Field(default=None, max_length=256)
    latency_ms: int | None = Field(default=None, ge=0)
    input_tokens: int | None = Field(default=None, ge=0)
    output_tokens: int | None = Field(default=None, ge=0)
    finish_reason: str | None = Field(default=None, max_length=128)
    response: Any = None


class AgentEvent(BaseModel):
    """Append-only event used by the UI to recover and render a run."""

    model_config = ConfigDict(extra="forbid")

    event_id: str = Field(default="event", min_length=1, max_length=128)
    run_id: str = Field(min_length=1, max_length=128)
    revision: int = Field(default=1, ge=1)
    type: AgentEventType | None = None
    kind: AgentEventType | None = None
    at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))
    item_id: str | None = Field(default=None, max_length=128)
    payload: dict[str, Any] = Field(default_factory=dict)
    trace: ModelTrace | None = None
    tool_result: ToolResult | None = None
    model_trace: ModelTrace | None = None

    @model_validator(mode="after")
    def _normalise_aliases(self) -> "AgentEvent":
        # Accept either ``type``/``trace`` (new stream API) or
        # ``kind``/``model_trace`` (legacy clients), normalising in memory.
        if self.type is None and self.kind is not None:
            self.type = self.kind
        if self.kind is None and self.type is not None:
            self.kind = self.type
        if self.trace is None and self.model_trace is not None:
            self.trace = self.model_trace
        if self.model_trace is None and self.trace is not None:
            self.model_trace = self.trace
        if self.type is None:
            raise ValueError("事件必须包含 type 或 kind")
        return self


class BonusSourceDecision(BaseModel):
    model_config = ConfigDict(extra="forbid")

    status: Literal["auto", "no_update", "needs_conversation"]
    source: Literal["J", "K"] | None = None
    value: int | float | None = None
    reason: Literal["exactly_one_nonzero", "both_zero", "ambiguous_or_invalid"]


class Decision(BaseModel):
    """A user-confirmed action.  It is data, never executable code."""

    model_config = ConfigDict(extra="forbid")

    item_id: str = Field(min_length=1, max_length=128)
    action: DecisionAction
    value: str | int | float | bool | None = None
    reason: str = Field(default="", max_length=1000)
    confidence: float | None = Field(default=None, ge=0, le=1)
    source: str | None = Field(default=None, max_length=128)
    remember: bool = False


class WorkItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    item_id: str = Field(min_length=1, max_length=128)
    person_key: str = Field(min_length=1, max_length=256)
    person_category: str | None = Field(default=None, max_length=128)
    person_name: str | None = Field(default=None, max_length=256)
    status: WorkItemStatus = "pending"
    target_sheet: str | None = None
    target_cell: str | None = None
    current_value: Any = None
    proposed_value: Any = None
    candidate_values: list[Any] = Field(default_factory=list)
    reason: str = ""
    risk_level: Literal["low", "medium", "high"] = "low"
    decision: Decision | None = None


class ExecutionPlan(BaseModel):
    model_config = ConfigDict(extra="forbid")

    run_id: str = Field(min_length=1, max_length=128)
    company_id: str = Field(min_length=1, max_length=128)
    salary_month: str = Field(min_length=1, max_length=32)
    rule_version: str = Field(min_length=1, max_length=128)
    items: list[WorkItem] = Field(default_factory=list)
    auto_publish: bool = False
    month_confirmation_required: bool = False
    month_confirmed: bool = False


def _number(value: Any) -> float | int | None:
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        return None
    number = float(value)
    if not math.isfinite(number) or number < 0:
        return None
    return value


def choose_bonus_source(j_value: Any, k_value: Any) -> BonusSourceDecision:
    """Select the only non-zero bonus source without guessing ambiguous data."""
    j_number, k_number = _number(j_value), _number(k_value)
    if j_number is None or k_number is None:
        return BonusSourceDecision(status="needs_conversation", reason="ambiguous_or_invalid")
    j_nonzero, k_nonzero = j_number > 0, k_number > 0
    if j_nonzero and not k_nonzero:
        return BonusSourceDecision(status="auto", source="J", value=j_value, reason="exactly_one_nonzero")
    if k_nonzero and not j_nonzero:
        return BonusSourceDecision(status="auto", source="K", value=k_value, reason="exactly_one_nonzero")
    if not j_nonzero and not k_nonzero:
        return BonusSourceDecision(status="no_update", reason="both_zero")
    return BonusSourceDecision(status="needs_conversation", reason="ambiguous_or_invalid")
