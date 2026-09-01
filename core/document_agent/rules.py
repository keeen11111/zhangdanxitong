"""Versioned company rule-package contracts.

The compiler that turns DOCX/transcripts into these objects can be swapped in
later.  Execution only consumes validated packages and never executes text
from a document as Python or an Excel formula.
"""
from __future__ import annotations

import re
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


RuleSourceKind = Literal[
    "explicit_instruction",
    "active_rule_package",
    "rule_package",
    "manual",
    "recording",
]
RuleActionType = Literal[
    "set_value_from_source",
    "keep_current",
    "skip",
    "choose_unique_nonzero",
]

_DANGEROUS_KEYS = frozenset(
    {"code", "python", "formula", "formulas", "script", "expression", "eval", "exec"}
)
_UNSAFE_TEXT = re.compile(r"(?:^|[\s_])(python|formula|script|eval|exec)(?:$|[\s_])", re.I)


def _validate_safe_payload(value: Any, *, path: str = "") -> None:
    """Reject executable-looking values while allowing ordinary rule facts."""
    if isinstance(value, dict):
        for key, nested in value.items():
            normalized = str(key).strip().lower()
            if normalized in _DANGEROUS_KEYS:
                raise ValueError(f"不允许可执行字段: {path + '.' if path else ''}{key}")
            _validate_safe_payload(nested, path=f"{path}.{key}" if path else str(key))
        return
    if isinstance(value, (list, tuple)):
        for index, nested in enumerate(value):
            _validate_safe_payload(nested, path=f"{path}[{index}]")
        return
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.startswith("=") or _UNSAFE_TEXT.search(stripped):
            raise ValueError(f"不允许可执行字段: {path or '值'}")


def _validate_field_name(value: str, *, field: str) -> str:
    value = value.strip()
    if not value or any(char in value for char in "=;(){}[]\\\"'\n\r"):
        raise ValueError(f"字段 {field} 包含不安全字符")
    return value


class RuleAction(BaseModel):
    """Closed, declarative action vocabulary consumed by the workbook tools."""

    model_config = ConfigDict(extra="forbid")

    type: RuleActionType
    target_field: str | None = Field(default=None, max_length=128)
    source_field: str | None = Field(default=None, max_length=128)
    source_fields: list[str] = Field(default_factory=list, max_length=16)

    @model_validator(mode="after")
    def _validate_shape(self) -> "RuleAction":
        if self.target_field is not None:
            _validate_field_name(self.target_field, field="target_field")
        if self.source_field is not None:
            _validate_field_name(self.source_field, field="source_field")
        fields = [_validate_field_name(item, field="source_fields") for item in self.source_fields]
        self.source_fields = fields
        if self.type == "set_value_from_source":
            if not self.target_field or not self.source_field:
                raise ValueError("set_value_from_source 必须同时提供 target_field 和 source_field")
        elif self.type == "choose_unique_nonzero":
            if not self.target_field or len(fields) < 2 or len(set(fields)) != len(fields):
                raise ValueError("choose_unique_nonzero 必须提供至少两个不重复的 source_fields 和 target_field")
        elif self.type in {"keep_current", "skip"} and self.source_field is not None:
            raise ValueError(f"{self.type} 不接受 source_field")
        return self


class RuleSource(BaseModel):
    model_config = ConfigDict(extra="forbid")

    kind: RuleSourceKind
    ref: str = Field(min_length=1, max_length=500)
    excerpt: str = Field(min_length=1, max_length=2000)
    # Higher revisions win when two manuals have the same source priority.
    revision: int = Field(default=0, ge=0, le=1_000_000)


class Rule(BaseModel):
    model_config = ConfigDict(extra="forbid")

    rule_id: str = Field(min_length=1, max_length=128)
    condition: dict[str, Any] = Field(default_factory=dict)
    action: RuleAction
    exceptions: list[dict[str, Any]] = Field(default_factory=list)
    source_refs: list[str] = Field(min_length=1, max_length=64)

    @model_validator(mode="after")
    def _validate_payload(self) -> "Rule":
        _validate_safe_payload(self.condition, path="condition")
        _validate_safe_payload(self.exceptions, path="exceptions")
        return self


class RulePackage(BaseModel):
    model_config = ConfigDict(extra="forbid")

    package_id: str = Field(min_length=1, max_length=128)
    company_id: str = Field(min_length=1, max_length=128)
    version: str = Field(min_length=1, max_length=128)
    effective_period: str = Field(min_length=1, max_length=32)
    status: Literal["candidate", "active", "retired"] = "candidate"
    rules: list[Rule] = Field(default_factory=list)
    sources: list[RuleSource] = Field(default_factory=list)


_PRIORITY = {
    "explicit_instruction": 0,
    "active_rule_package": 1,
    "rule_package": 1,
    "manual": 2,
    "recording": 3,
}


def resolve_rule_priority(source_kind: str) -> int:
    """Return the lower-is-stronger priority for one supported source."""
    try:
        return _PRIORITY[source_kind]
    except KeyError as exc:
        raise ValueError(f"不支持的规则来源：{source_kind}") from exc
