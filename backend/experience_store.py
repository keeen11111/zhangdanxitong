"""Tenant-scoped memory for reviewed workbook decisions.

The store keeps reusable, non-PII conditions and the decision made by a
reviewer. It does not execute workbook changes. Rules are kept in tenant
files for backwards compatibility with the existing review workflow.
"""
from __future__ import annotations

from collections.abc import Mapping, Sequence
from datetime import UTC, date, datetime
from hashlib import sha256
import json
import math
from pathlib import Path
from typing import Any
from uuid import uuid4


SENSITIVE_FIELD_TOKENS = (
    "工号",
    "姓名",
    "身份证",
    "证件",
    "手机",
    "电话",
    "邮箱",
    "地址",
    "银行卡",
    "银行账号",
    "账户",
)
MEMORY_STATUSES = frozenset({"candidate", "active", "suspended"})
DEFAULT_SCOPE = "company"
DEFAULT_RULE_VERSION = "1"


class ExperienceStore:
    """Persist reviewed memory rules separated by tenant."""

    def __init__(self, root: Path | str) -> None:
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)

    def record(
        self,
        *,
        tenant_id: str,
        project_id: str,
        created_by: str,
        diff_type: str,
        item: dict[str, Any],
        match_fields: list[str],
        decision: str,
        updates: dict[str, Any],
        note: str,
        status: str = "candidate",
        scope: str = DEFAULT_SCOPE,
        period_key: str | None = None,
        evidence: Any = None,
        rule_version: str | None = None,
        metadata: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Record a decision using the legacy item/match-fields contract."""
        return self.record_memory(
            tenant_id=tenant_id,
            project_id=project_id,
            created_by=created_by,
            diff_type=diff_type,
            item=item,
            match_fields=match_fields,
            decision=decision,
            updates=updates,
            note=note,
            status=status,
            scope=scope,
            period_key=period_key,
            evidence=evidence,
            rule_version=rule_version,
            metadata=metadata,
            _deduplicate=False,
        )

    def record_memory(
        self,
        *,
        tenant_id: str,
        project_id: str = "",
        created_by: str = "",
        diff_type: str,
        item: Mapping[str, Any] | None = None,
        match_fields: Sequence[str] | None = None,
        conditions: Mapping[str, Any] | None = None,
        decision: str,
        updates: Mapping[str, Any] | None = None,
        note: str = "",
        status: str = "candidate",
        scope: str = DEFAULT_SCOPE,
        period_key: str | None = None,
        evidence: Any = None,
        rule_version: str | None = None,
        metadata: Mapping[str, Any] | None = None,
        _deduplicate: bool = True,
    ) -> dict[str, Any]:
        """Record or update a company memory rule.

        When ``period_key`` is present, recording the same decision twice for
        the same period is idempotent. Distinct periods increase consistency
        counters once each; three consistent periods promote a candidate to
        ``active``. A conflicting decision suspends the previous rule and is
        stored as a new candidate for auditability.
        """
        if not isinstance(diff_type, str) or not diff_type.strip():
            raise ValueError("经验规则的差异类型不能为空")
        if not isinstance(decision, str) or not decision.strip():
            raise ValueError("经验规则决定不能为空")
        decision = decision.strip()
        if status not in MEMORY_STATUSES:
            raise ValueError("经验规则状态无效")
        scope = _normalise_text(scope, DEFAULT_SCOPE)
        period_key = _normalise_period_key(period_key)
        rule_version = _normalise_text(rule_version, DEFAULT_RULE_VERSION)

        if conditions is not None:
            normalised_conditions = self._normalise_conditions(conditions)
        else:
            normalised_conditions = self._conditions_from_item(
                item or {}, list(match_fields or [])
            )
        if not normalised_conditions:
            raise ValueError("经验规则至少需要一个非个人标识的匹配字段")

        now = _utc_now()
        rule = {
            "id": uuid4().hex,
            "project_id": str(project_id or ""),
            "created_by": str(created_by or ""),
            "diff_type": diff_type.strip(),
            "conditions": normalised_conditions,
            "decision": decision,
            "updates": dict(updates or {}),
            "note": str(note or "").strip(),
            "created_at": now,
            "updated_at": now,
            "status": status,
            "scope": scope,
            "period_key": period_key,
            "period_keys": [period_key] if period_key else [],
            "evidence": evidence,
            "use_count": 1,
            "consistent_count": 1,
            "conflict_count": 0,
            "rule_version": rule_version,
            "metadata": dict(metadata or {}),
        }
        _ensure_json_serialisable(rule)

        rules = self._load(tenant_id)
        if not _deduplicate:
            rules.append(rule)
            self._save(tenant_id, rules)
            return rule

        identity = self._rule_identity(rule)
        matching = [existing for existing in rules if self._rule_identity(existing) == identity]
        previous = next(reversed(matching), None)
        if previous is not None:
            periods = _periods_for(previous)
            same_period = bool(period_key and period_key in periods)
            if previous.get("decision") == decision and not same_period:
                if period_key:
                    periods.append(period_key)
                    previous["period_keys"] = periods
                    previous["use_count"] = _counter(previous, "use_count") + 1
                    previous["consistent_count"] = _counter(previous, "consistent_count") + 1
                    if previous.get("status") != "suspended" and previous["consistent_count"] >= 3:
                        previous["status"] = "active"
                previous["updated_at"] = _utc_now()
                if evidence is not None:
                    previous["evidence"] = evidence
                if metadata:
                    previous.setdefault("metadata", {}).update(dict(metadata))
                _ensure_json_serialisable(previous)
                self._save(tenant_id, rules)
                return previous
            if previous.get("decision") == decision and same_period:
                return previous
            previous["conflict_count"] = _counter(previous, "conflict_count") + 1
            previous["status"] = "suspended"
            previous["updated_at"] = _utc_now()

        rules.append(rule)
        self._save(tenant_id, rules)
        return rule

    def observe(
        self,
        tenant_id: str,
        rule_id: str,
        period_key: str,
        *,
        consistent: bool = True,
        evidence: Any = None,
        rule_version: str | None = None,
        metadata: Mapping[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Record one deduplicated observation for an existing memory rule.

        ``period_key`` is the caller's accounting period (for example,
        ``2026-07``), not a wall-clock timestamp. This keeps retries within a
        monthly run from inflating the confidence counters.
        """
        normalized_period = _normalise_period_key(period_key)
        if not normalized_period:
            raise ValueError("观察记录需要有效的周期")
        rules = self._load(tenant_id)
        for rule in rules:
            if rule.get("id") != rule_id:
                continue
            periods = _periods_for(rule)
            if normalized_period in periods:
                return self._with_defaults(rule)

            periods.append(normalized_period)
            rule["period_keys"] = periods
            rule["observed_periods"] = list(periods)
            rule["use_count"] = _counter(rule, "use_count") + 1
            if consistent:
                rule["consistent_count"] = _counter(rule, "consistent_count") + 1
                if (
                    rule.get("status") == "candidate"
                    and rule["consistent_count"] >= 3
                ):
                    rule["status"] = "active"
            else:
                rule["conflict_count"] = _counter(rule, "conflict_count") + 1
                rule["status"] = "suspended"
                rule["suspended_at"] = _utc_now()
            rule["updated_at"] = _utc_now()
            if evidence is not None:
                rule["evidence"] = evidence
            if rule_version is not None:
                rule["rule_version"] = _normalise_text(rule_version, DEFAULT_RULE_VERSION)
            if metadata:
                rule.setdefault("metadata", {}).update(dict(metadata))
            _ensure_json_serialisable(rule)
            self._save(tenant_id, rules)
            return self._with_defaults(rule)
        raise ValueError("经验规则不存在")

    def record_agent_decision(
        self,
        *,
        tenant_id: str,
        project_id: str = "",
        created_by: str = "",
        diff_type: str | None = None,
        item: Mapping[str, Any] | None = None,
        match_fields: Sequence[str] | None = None,
        decision: str | Mapping[str, Any] | None = None,
        updates: Mapping[str, Any] | None = None,
        note: str = "",
        scope: str = DEFAULT_SCOPE,
        period_key: str | None = None,
        evidence: Any = None,
        rule_version: str | None = None,
        confidence: float | None = None,
        metadata: Mapping[str, Any] | None = None,
        agent_decision: Mapping[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Persist a structured Agent decision as a company-scoped memory.

        The documented ``decision/value/reason/confidence`` shape is accepted
        alongside the explicit diff arguments used by the review workflow.
        Agent-only fields are retained as metadata and never become matching
        conditions.
        """
        structured: dict[str, Any] = dict(agent_decision or {})
        if isinstance(decision, Mapping):
            structured = {**structured, **dict(decision)}
            resolved_decision = structured.get("decision") or structured.get("action")
        else:
            resolved_decision = decision or structured.get("decision") or structured.get("action")
        resolved_diff_type = diff_type or structured.get("diff_type")
        if item is None:
            item = structured.get("item")
        resolved_scope = scope if scope != DEFAULT_SCOPE else str(structured.get("scope", scope))
        resolved_note = note or str(structured.get("reason") or structured.get("note") or "")
        resolved_updates = dict(updates or {})
        if not resolved_updates and "value" in structured:
            resolved_updates = {"value": structured["value"]}
        resolved_metadata = dict(metadata or {})
        if confidence is not None:
            resolved_metadata["confidence"] = confidence
        for key in ("item_id", "work_item_id"):
            if structured.get(key) is not None:
                resolved_metadata[key] = structured[key]
        if isinstance(structured.get("metadata"), Mapping):
            resolved_metadata.update(structured["metadata"])
        if not resolved_diff_type or not resolved_decision or item is None:
            raise ValueError("Agent 决策需要 diff_type、decision 和 item")
        return self.record_memory(
            tenant_id=tenant_id,
            project_id=project_id,
            created_by=created_by,
            diff_type=str(resolved_diff_type),
            item=item,
            match_fields=list(match_fields or []),
            decision=str(resolved_decision),
            updates=resolved_updates,
            note=resolved_note,
            scope=resolved_scope,
            period_key=period_key,
            evidence=evidence,
            rule_version=rule_version,
            metadata=resolved_metadata,
        )

    def suggest(
        self,
        tenant_id: str,
        diff_type: str,
        item: Mapping[str, Any],
        metadata: Mapping[str, Any] | None = None,
    ) -> list[dict[str, Any]]:
        """Return exact condition matches, newest rule first.

        The no-metadata form intentionally preserves the old response shape.
        Callers that pass ``metadata`` receive explainability fields alongside
        the legacy fields.
        """
        suggestions: list[dict[str, Any]] = []
        for raw_rule in reversed(self._load(tenant_id)):
            rule = self._with_defaults(raw_rule)
            if rule.get("status") == "suspended" or rule.get("diff_type") != diff_type:
                continue
            conditions = rule.get("conditions")
            if not isinstance(conditions, dict) or not conditions:
                continue
            if not all(item.get(field) == value for field, value in conditions.items()):
                continue
            suggestion = {
                "rule_id": rule["id"],
                "decision": rule["decision"],
                "updates": rule.get("updates", {}),
                "note": rule.get("note", ""),
                "matched_fields": sorted(conditions),
            }
            if metadata is not None:
                suggestion.update({
                    "status": rule["status"],
                    "scope": rule["scope"],
                    "period_key": rule.get("period_key"),
                    "evidence": rule.get("evidence"),
                    "use_count": rule["use_count"],
                    "consistent_count": rule["consistent_count"],
                    "conflict_count": rule["conflict_count"],
                    "rule_version": rule["rule_version"],
                    "metadata": rule.get("metadata", {}),
                })
            suggestions.append(suggestion)
        return suggestions

    def suggest_memory(
        self,
        tenant_id: str,
        diff_type: str,
        item: Mapping[str, Any],
        metadata: Mapping[str, Any] | None = None,
    ) -> list[dict[str, Any]]:
        """Agent-facing alias for :meth:`suggest`."""
        return self.suggest(tenant_id, diff_type, item, metadata=metadata)

    def list_rules(self, tenant_id: str, offset: int, limit: int) -> tuple[list[dict[str, Any]], int]:
        """Return paginated rules without exposing tenant storage internals."""
        rules = [self._with_defaults(rule) for rule in reversed(self._load(tenant_id))]
        return rules[offset:offset + limit], len(rules)

    def set_status(self, tenant_id: str, rule_id: str, status: str) -> dict[str, Any]:
        """Explicitly activate or suspend one tenant rule after validation."""
        if status not in MEMORY_STATUSES:
            raise ValueError("经验规则状态无效")
        rules = self._load(tenant_id)
        for rule in rules:
            if str(rule.get("id")) != str(rule_id):
                continue
            rule["status"] = status
            rule["updated_at"] = _utc_now()
            if status == "suspended":
                rule["suspended_at"] = rule["updated_at"]
            _ensure_json_serialisable(rule)
            self._save(tenant_id, rules)
            return self._with_defaults(rule)
        raise ValueError("经验规则不存在")

    @staticmethod
    def _conditions_from_item(item: Mapping[str, Any], match_fields: Sequence[str]) -> dict[str, Any]:
        conditions: dict[str, Any] = {}
        for field in match_fields:
            if not isinstance(field, str):
                continue
            field = field.strip()
            if _is_sensitive_field(field) or field not in item:
                continue
            value = item[field]
            if isinstance(value, (dict, list, tuple, set)) or value in (None, ""):
                continue
            if isinstance(value, (datetime, date)):
                value = value.isoformat()
            if isinstance(value, float) and math.isnan(value):
                continue
            if not isinstance(value, (str, int, float, bool)):
                continue
            conditions[field] = value.strip() if isinstance(value, str) else value
        return conditions

    @classmethod
    def _normalise_conditions(cls, conditions: Mapping[str, Any]) -> dict[str, Any]:
        return cls._conditions_from_item(conditions, list(conditions))

    @staticmethod
    def _rule_identity(rule: Mapping[str, Any]) -> tuple[Any, ...]:
        conditions = rule.get("conditions") or {}
        return (
            rule.get("diff_type"),
            json.dumps(conditions, ensure_ascii=False, sort_keys=True, separators=(",", ":")),
            rule.get("scope", DEFAULT_SCOPE),
        )

    @staticmethod
    def _with_defaults(rule: Mapping[str, Any]) -> dict[str, Any]:
        result = dict(rule)
        result.setdefault("status", "candidate")
        result.setdefault("scope", DEFAULT_SCOPE)
        result.setdefault("period_key", None)
        result.setdefault("period_keys", [result["period_key"]] if result.get("period_key") else [])
        result.setdefault("observed_periods", list(result["period_keys"]))
        result.setdefault("evidence", None)
        result.setdefault("use_count", 1)
        result.setdefault("consistent_count", 1)
        result.setdefault("conflict_count", 0)
        result.setdefault("rule_version", DEFAULT_RULE_VERSION)
        result.setdefault("metadata", {})
        result.setdefault("updates", {})
        result.setdefault("note", "")
        return result

    def _path_for(self, tenant_id: str) -> Path:
        tenant_key = sha256(str(tenant_id).encode("utf-8")).hexdigest()
        return self.root / f"{tenant_key}.json"

    def _load(self, tenant_id: str) -> list[dict[str, Any]]:
        path = self._path_for(tenant_id)
        if not path.exists():
            return []
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise ValueError("经验库文件损坏，无法读取") from exc
        if not isinstance(raw, list) or not all(isinstance(rule, dict) for rule in raw):
            raise ValueError("经验库文件格式无效")
        return raw

    def _save(self, tenant_id: str, rules: list[dict[str, Any]]) -> None:
        path = self._path_for(tenant_id)
        temp_path = path.with_suffix(".tmp")
        temp_path.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
        temp_path.replace(path)


def _utc_now() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")


def _normalise_text(value: Any, default: str) -> str:
    if value is None:
        return default
    if not isinstance(value, str):
        raise ValueError("经验规则文本字段必须是字符串")
    value = value.strip()
    return value or default


def _normalise_period_key(value: Any) -> str | None:
    if value is None or value == "":
        return None
    if not isinstance(value, str):
        raise ValueError("经验规则周期必须是字符串")
    value = value.strip()
    if not value or len(value) > 64:
        raise ValueError("经验规则周期长度无效")
    return value


def _periods_for(rule: Mapping[str, Any]) -> list[str]:
    periods = rule.get("period_keys")
    if not isinstance(periods, list):
        periods = rule.get("observed_periods")
    if isinstance(periods, list):
        return list(dict.fromkeys(period for period in periods if isinstance(period, str)))
    period = rule.get("period_key")
    return [period] if isinstance(period, str) and period else []


def _counter(rule: Mapping[str, Any], key: str) -> int:
    value = rule.get(key, 0)
    return value if isinstance(value, int) and value >= 0 else 0


def _ensure_json_serialisable(value: Any) -> None:
    try:
        json.dumps(value, ensure_ascii=False)
    except (TypeError, ValueError) as exc:
        raise ValueError("经验规则包含无法保存的字段") from exc


def select_unique_nonzero_source(
    values: Mapping[str, Any],
    columns: Sequence[str] = ("J", "K"),
) -> tuple[str, Any] | None:
    """Return the only non-zero source column/value, otherwise ``None``."""
    candidates: list[tuple[str, Any]] = []
    for column in columns:
        value = values.get(column)
        if _is_nonzero(value):
            candidates.append((column, value))
    return candidates[0] if len(candidates) == 1 else None


def _is_nonzero(value: Any) -> bool:
    if value is None or value is False:
        return False
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return False
        try:
            return float(stripped.replace(",", "")) != 0
        except ValueError:
            return True
    if isinstance(value, (int, float)):
        return value != 0 and not (isinstance(value, float) and math.isnan(value))
    return bool(value)


def _is_sensitive_field(field: str) -> bool:
    return any(token in field for token in SENSITIVE_FIELD_TOKENS)
