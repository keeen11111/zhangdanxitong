"""Deterministic compilation of material evidence into a safe rule package.

An LLM may propose :class:`RuleCandidate` objects, but this module is the
trust boundary: it validates evidence, resolves source precedence, rejects
same-priority disagreement, and emits only the closed action vocabulary from
``rules.py``. No document text is ever evaluated as Python, SQL, or Excel
formula.
"""
from __future__ import annotations

import json
from typing import Any, Iterable

from pydantic import BaseModel, ConfigDict, Field, model_validator

from .rules import (
    Rule,
    RuleAction,
    RulePackage,
    RuleSource,
    _validate_safe_payload,
    resolve_rule_priority,
)


class RuleCompilationError(ValueError):
    """Raised when material evidence cannot produce one unambiguous rule."""


class RuleCandidate(BaseModel):
    """One model-extracted rule assertion plus its locatable evidence."""

    model_config = ConfigDict(extra="forbid")

    rule_id: str = Field(min_length=1, max_length=128)
    condition: dict[str, Any] = Field(default_factory=dict)
    action: RuleAction
    exceptions: list[dict[str, Any]] = Field(default_factory=list)
    source: RuleSource

    @model_validator(mode="after")
    def _validate_candidate(self) -> "RuleCandidate":
        _validate_safe_payload(self.condition, path="condition")
        _validate_safe_payload(self.exceptions, path="exceptions")
        return self


def _canonical(value: Any) -> str:
    """Stable comparison key for JSON-like declarative values."""
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), default=str)


def _candidate_key(candidate: RuleCandidate) -> tuple[str, str]:
    return candidate.rule_id, _canonical(candidate.condition)


def _action_key(candidate: RuleCandidate) -> str:
    return _canonical(candidate.action.model_dump(mode="json", exclude_none=True))


def _exceptions_key(candidate: RuleCandidate) -> str:
    return _canonical(candidate.exceptions)


def _select_strongest(candidates: list[RuleCandidate]) -> list[RuleCandidate]:
    priorities = [resolve_rule_priority(candidate.source.kind) for candidate in candidates]
    strongest_priority = min(priorities)
    selected = [
        candidate
        for candidate in candidates
        if resolve_rule_priority(candidate.source.kind) == strongest_priority
    ]

    # Newer manual/rule-package revisions supersede older assertions at the
    # same priority. Explicit run instructions remain eligible for conflicts.
    if selected and selected[0].source.kind in {"manual", "rule_package", "active_rule_package"}:
        latest_revision = max(candidate.source.revision for candidate in selected)
        selected = [candidate for candidate in selected if candidate.source.revision == latest_revision]
    return selected


def _merge_group(candidates: list[RuleCandidate]) -> tuple[Rule, list[RuleSource]]:
    selected = _select_strongest(candidates)
    action_keys = {_action_key(candidate) for candidate in selected}
    exception_keys = {_exceptions_key(candidate) for candidate in selected}
    if len(action_keys) > 1 or len(exception_keys) > 1:
        refs = ", ".join(sorted(candidate.source.ref for candidate in selected))
        raise RuleCompilationError(f"同优先级规则冲突: {candidates[0].rule_id} ({refs})")

    evidence = sorted(
        {candidate.source.ref: candidate.source for candidate in selected}.values(),
        key=lambda source: source.ref,
    )
    chosen = selected[0]
    rule = Rule(
        rule_id=chosen.rule_id,
        condition=chosen.condition,
        action=chosen.action,
        exceptions=chosen.exceptions,
        source_refs=[source.ref for source in evidence],
    )
    return rule, evidence


def validate_rule_package(package: RulePackage) -> RulePackage:
    """Validate cross-references and safety invariants on a compiled package."""
    sources_by_ref: dict[str, RuleSource] = {}
    for source in package.sources:
        if source.ref in sources_by_ref:
            raise ValueError(f"重复规则来源: {source.ref}")
        if not source.excerpt.strip():
            raise ValueError(f"规则来源缺少摘录: {source.ref}")
        sources_by_ref[source.ref] = source

    rule_ids: set[str] = set()
    for rule in package.rules:
        if rule.rule_id in rule_ids:
            raise ValueError(f"重复规则 ID: {rule.rule_id}")
        rule_ids.add(rule.rule_id)
        if not rule.source_refs:
            raise ValueError(f"规则缺少来源: {rule.rule_id}")
        for ref in rule.source_refs:
            if ref not in sources_by_ref:
                raise ValueError(f"未知来源: {ref}")
    return package


def compile_rule_package(
    *,
    package_id: str,
    company_id: str,
    version: str,
    effective_period: str,
    candidates: Iterable[RuleCandidate],
    status: str = "candidate",
) -> RulePackage:
    """Compile candidates using explicit > package > manual > recording precedence."""
    normalized = [
        candidate if isinstance(candidate, RuleCandidate) else RuleCandidate.model_validate(candidate)
        for candidate in candidates
    ]
    groups: dict[tuple[str, str], list[RuleCandidate]] = {}
    for candidate in normalized:
        groups.setdefault(_candidate_key(candidate), []).append(candidate)

    rules: list[Rule] = []
    sources: list[RuleSource] = []
    for key in sorted(groups):
        rule, evidence = _merge_group(groups[key])
        rules.append(rule)
        sources.extend(evidence)

    unique_sources = {source.ref: source for source in sources}
    package = RulePackage(
        package_id=package_id,
        company_id=company_id,
        version=version,
        effective_period=effective_period,
        status=status,
        rules=rules,
        sources=sorted(unique_sources.values(), key=lambda source: source.ref),
    )
    return validate_rule_package(package)


__all__ = [
    "RuleCandidate",
    "RuleCompilationError",
    "compile_rule_package",
    "validate_rule_package",
]
