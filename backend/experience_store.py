"""Tenant-scoped, human-confirmed review experience rules.

The store intentionally contains only reusable conditions and a human decision.
It never applies a decision to payroll data automatically.
"""
from __future__ import annotations

from datetime import date, datetime
from hashlib import sha256
import json
from pathlib import Path
from typing import Any
from uuid import uuid4


IDENTITY_FIELDS = frozenset({"工号", "姓名", "身份证号", "身份证号码", "证件号码"})


class ExperienceStore:
    """Persist reviewed cases locally, separated by tenant."""

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
    ) -> dict[str, Any]:
        """Record one human-reviewed decision as a reusable, non-PII rule."""
        conditions = self._conditions_from_item(item, match_fields)
        if not conditions:
            raise ValueError("经验规则至少需要一个非个人标识的匹配字段")

        rule = {
            "id": uuid4().hex,
            "project_id": project_id,
            "created_by": created_by,
            "diff_type": diff_type,
            "conditions": conditions,
            "decision": decision,
            "updates": updates,
            "note": note,
            "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        }
        rules = self._load(tenant_id)
        rules.append(rule)
        self._save(tenant_id, rules)
        return rule

    def suggest(
        self,
        tenant_id: str,
        diff_type: str,
        item: dict[str, Any],
    ) -> list[dict[str, Any]]:
        """Return exact condition matches as suggestions, newest rule first."""
        suggestions: list[dict[str, Any]] = []
        for rule in reversed(self._load(tenant_id)):
            if rule.get("diff_type") != diff_type:
                continue
            conditions = rule.get("conditions")
            if not isinstance(conditions, dict) or not conditions:
                continue
            if not all(item.get(field) == value for field, value in conditions.items()):
                continue
            suggestions.append({
                "rule_id": rule["id"],
                "decision": rule["decision"],
                "updates": rule.get("updates", {}),
                "note": rule.get("note", ""),
                "matched_fields": sorted(conditions),
            })
        return suggestions

    def list_rules(self, tenant_id: str, offset: int, limit: int) -> tuple[list[dict[str, Any]], int]:
        """Return paginated rules without exposing tenant storage internals."""
        rules = list(reversed(self._load(tenant_id)))
        return rules[offset:offset + limit], len(rules)

    @staticmethod
    def _conditions_from_item(item: dict[str, Any], match_fields: list[str]) -> dict[str, Any]:
        conditions: dict[str, Any] = {}
        for field in match_fields:
            if field in IDENTITY_FIELDS or field not in item:
                continue
            value = item[field]
            if isinstance(value, (dict, list, tuple, set)) or value in (None, ""):
                continue
            if isinstance(value, (datetime, date)):
                value = value.isoformat()
            if not isinstance(value, (str, int, float, bool)):
                continue
            conditions[field] = value.strip() if isinstance(value, str) else value
        return conditions

    def _path_for(self, tenant_id: str) -> Path:
        tenant_key = sha256(tenant_id.encode("utf-8")).hexdigest()
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
