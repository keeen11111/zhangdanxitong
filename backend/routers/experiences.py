"""Endpoints for human-confirmed, tenant-scoped review experience rules."""
from __future__ import annotations

import os
from typing import Any, Literal, Mapping

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, ConfigDict, Field, model_validator
from sqlalchemy.orm import Session

from backend.auth import get_current_user
from backend.database import DATA_DIR, get_db
from backend.experience_store import ExperienceStore
from backend.models import User
from backend.routers.pipeline import _load_json, _load_project_or_404, _save_json


router = APIRouter(tags=["experiences"])
experience_store = ExperienceStore(os.path.join(DATA_DIR, "experience_rules"))


class ExperienceSuggestion(BaseModel):
    model_config = ConfigDict(extra="ignore")

    rule_id: str
    decision: str
    updates: dict[str, Any]
    note: str
    matched_fields: list[str]
    status: Literal["candidate", "active", "suspended"] = "candidate"
    scope: str = "company"
    period_key: str | None = None
    evidence: Any = None
    use_count: int = 1
    consistent_count: int = 1
    conflict_count: int = 0
    rule_version: str = "1"
    metadata: dict[str, Any] = Field(default_factory=dict)


class LegacyExperienceSuggestion(BaseModel):
    """Original pipeline response shape; kept stable for existing clients."""

    rule_id: str
    decision: str
    updates: dict[str, Any]
    note: str
    matched_fields: list[str]


class ExperienceSuggestionItem(BaseModel):
    diff_type: str
    item_index: int
    suggestions: list[LegacyExperienceSuggestion]


class ExperienceSuggestionResponse(BaseModel):
    items: list[ExperienceSuggestionItem]
    total: int


class ExperienceRecordIn(BaseModel):
    model_config = ConfigDict(extra="forbid")

    diff_type: str = Field(min_length=1, max_length=64)
    item_index: int = Field(ge=0)
    match_fields: list[str] = Field(min_length=1, max_length=5)
    decision: Literal["confirmed", "ignored"]
    note: str = Field(min_length=1, max_length=1000)
    status: Literal["candidate", "active", "suspended"] = "candidate"
    scope: str = Field(default="company", min_length=1, max_length=64)
    period_key: str | None = Field(default=None, max_length=64)
    evidence: Any = None
    rule_version: str | None = Field(default=None, max_length=128)
    metadata: dict[str, Any] = Field(default_factory=dict)
    updates: dict[str, Any] = Field(default_factory=dict)


class ExperienceRuleOut(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: str
    diff_type: str
    conditions: dict[str, Any]
    decision: str
    updates: dict[str, Any] = Field(default_factory=dict)
    note: str
    created_at: str
    updated_at: str | None = None
    status: Literal["candidate", "active", "suspended"] = "candidate"
    scope: str = "company"
    period_key: str | None = None
    evidence: Any = None
    use_count: int = 1
    consistent_count: int = 1
    conflict_count: int = 0
    rule_version: str = "1"
    metadata: dict[str, Any] = Field(default_factory=dict)


class ExperienceRulePage(BaseModel):
    items: list[ExperienceRuleOut]
    page: int
    page_size: int
    total: int


class AgentMemoryRecordIn(BaseModel):
    """Validated, tenant-scoped memory input for Agent conversations."""

    model_config = ConfigDict(extra="forbid")

    project_id: str = Field(default="", max_length=128)
    diff_type: str = Field(min_length=1, max_length=64)
    item: dict[str, Any] | None = None
    match_fields: list[str] | None = Field(default=None, max_length=10)
    conditions: dict[str, Any] | None = None
    decision: str = Field(min_length=1, max_length=64)
    updates: dict[str, Any] = Field(default_factory=dict)
    note: str = Field(default="", max_length=1000)
    status: Literal["candidate", "active", "suspended"] = "candidate"
    scope: str = Field(default="company", min_length=1, max_length=64)
    period_key: str | None = Field(default=None, max_length=64)
    evidence: Any = None
    rule_version: str | None = Field(default=None, max_length=128)
    metadata: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="after")
    def require_memory_conditions(self) -> "AgentMemoryRecordIn":
        if self.conditions is None and (self.item is None or not self.match_fields):
            raise ValueError("记忆规则需要 conditions，或 item 与 match_fields")
        if self.match_fields is not None and len(set(self.match_fields)) != len(self.match_fields):
            raise ValueError("匹配字段不能重复")
        return self


class AgentMemorySuggestIn(BaseModel):
    model_config = ConfigDict(extra="forbid")

    diff_type: str = Field(min_length=1, max_length=64)
    item: dict[str, Any]
    metadata: dict[str, Any] | None = None


class AgentMemorySuggestionResponse(BaseModel):
    suggestions: list[ExperienceSuggestion]
    total: int


def _reviewable_diff_items(diff: dict[str, Any]) -> list[tuple[str, list[dict[str, Any]]]]:
    return [
        (diff_type, items)
        for diff_type, items in diff.items()
        if diff_type != "summary" and isinstance(items, list)
    ]


@router.get("/api/pipeline/{project_id}/experience-suggestions", response_model=ExperienceSuggestionResponse)
def get_experience_suggestions(
    project_id: str,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> ExperienceSuggestionResponse:
    """Return past human decisions that exactly match current pending differences."""
    project = _load_project_or_404(project_id, user, db)
    data = _load_json(project.id, "diff")
    if not data:
        raise HTTPException(status_code=400, detail="尚未计算差异")

    suggested_items: list[ExperienceSuggestionItem] = []
    for diff_type, items in _reviewable_diff_items(data.get("diff", {})):
        for item_index, item in enumerate(items):
            if item.get("status", "pending") != "pending":
                continue
            suggestions = experience_store.suggest(user.tenant_id, diff_type, item)
            if suggestions:
                suggested_items.append(ExperienceSuggestionItem(
                    diff_type=diff_type,
                    item_index=item_index,
                    suggestions=[LegacyExperienceSuggestion(**suggestion) for suggestion in suggestions],
                ))
    return ExperienceSuggestionResponse(items=suggested_items, total=len(suggested_items))


@router.post("/api/pipeline/{project_id}/experiences", response_model=ExperienceRuleOut, status_code=201)
def record_review_experience(
    project_id: str,
    payload: ExperienceRecordIn,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> ExperienceRuleOut:
    """Confirm one decision and save non-PII matching conditions for later review."""
    project = _load_project_or_404(project_id, user, db)
    data = _load_json(project.id, "diff")
    if not data:
        raise HTTPException(status_code=400, detail="尚未计算差异")

    diff = data.get("diff", {})
    items = diff.get(payload.diff_type)
    if not isinstance(items, list) or payload.item_index >= len(items):
        raise HTTPException(status_code=400, detail="差异类型或索引无效")
    item = items[payload.item_index]
    if item.get("status", "pending") != "pending":
        raise HTTPException(status_code=409, detail="该差异项已审核，不能覆盖历史决定")
    match_fields = [field.strip() for field in payload.match_fields if field.strip()]
    if len(match_fields) != len(set(match_fields)):
        raise HTTPException(status_code=422, detail="匹配字段不能重复")
    if any(field not in item for field in match_fields):
        raise HTTPException(status_code=422, detail="匹配字段必须存在于当前差异项")

    try:
        rule = experience_store.record(
            tenant_id=user.tenant_id,
            project_id=project.id,
            created_by=user.id,
            diff_type=payload.diff_type,
            item=item,
            match_fields=match_fields,
            decision=payload.decision,
            updates=payload.updates,
            note=payload.note.strip(),
            status=payload.status,
            scope=payload.scope,
            period_key=payload.period_key,
            evidence=payload.evidence,
            rule_version=payload.rule_version,
            metadata=payload.metadata,
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc

    item["status"] = payload.decision
    _save_json(project.id, "diff", data)
    return ExperienceRuleOut(**rule)


@router.get("/api/pipeline/{project_id}/experiences", response_model=ExperienceRulePage)
def list_review_experiences(
    project_id: str,
    page: int = 1,
    page_size: int = 20,
    user: User = Depends(get_current_user),
    db: Session = Depends(get_db),
) -> ExperienceRulePage:
    """List this tenant's reusable review rules without exposing other tenants' data."""
    _load_project_or_404(project_id, user, db)
    if page < 1 or page_size < 1 or page_size > 100:
        raise HTTPException(status_code=422, detail="分页参数无效")
    rules, total = experience_store.list_rules(
        user.tenant_id,
        offset=(page - 1) * page_size,
        limit=page_size,
    )
    return ExperienceRulePage(
        items=[ExperienceRuleOut(**rule) for rule in rules],
        page=page,
        page_size=page_size,
        total=total,
    )


@router.post("/api/agent/memory", response_model=ExperienceRuleOut, status_code=201)
@router.post("/api/agent/memory/record", response_model=ExperienceRuleOut, status_code=201)
def record_agent_memory(
    payload: AgentMemoryRecordIn,
    user: User = Depends(get_current_user),
) -> ExperienceRuleOut:
    """Save one Agent decision in the current tenant's memory."""
    try:
        rule = experience_store.record_memory(
            tenant_id=user.tenant_id,
            project_id=payload.project_id,
            created_by=user.id,
            diff_type=payload.diff_type,
            item=payload.item,
            match_fields=payload.match_fields,
            conditions=payload.conditions,
            decision=payload.decision,
            updates=payload.updates,
            note=payload.note,
            status=payload.status,
            scope=payload.scope,
            period_key=payload.period_key,
            evidence=payload.evidence,
            rule_version=payload.rule_version,
            metadata=payload.metadata,
        )
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    return ExperienceRuleOut(**rule)


@router.post("/api/agent/memory/suggest", response_model=AgentMemorySuggestionResponse)
def suggest_agent_memory(
    payload: AgentMemorySuggestIn,
    user: User = Depends(get_current_user),
) -> AgentMemorySuggestionResponse:
    """Suggest reusable memory rules without crossing tenant boundaries."""
    suggestions = experience_store.suggest_memory(
        user.tenant_id,
        payload.diff_type,
        payload.item,
        metadata=payload.metadata,
    )
    return AgentMemorySuggestionResponse(
        suggestions=[ExperienceSuggestion(**suggestion) for suggestion in suggestions],
        total=len(suggestions),
    )


@router.get("/api/agent/memory", response_model=ExperienceRulePage)
def list_agent_memory(
    page: int = 1,
    page_size: int = 20,
    user: User = Depends(get_current_user),
) -> ExperienceRulePage:
    """List the current tenant's Agent memory for the company workbench."""
    if page < 1 or page_size < 1 or page_size > 100:
        raise HTTPException(status_code=422, detail="分页参数无效")
    rules, total = experience_store.list_rules(
        user.tenant_id,
        offset=(page - 1) * page_size,
        limit=page_size,
    )
    return ExperienceRulePage(
        items=[ExperienceRuleOut(**rule) for rule in rules],
        page=page,
        page_size=page_size,
        total=total,
    )


def _set_agent_memory_status(rule_id: str, status: str, user: User) -> ExperienceRuleOut:
    try:
        rule = experience_store.set_status(user.tenant_id, rule_id, status)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return ExperienceRuleOut(**rule)


@router.post("/api/agent/memory/{rule_id}/suspend", response_model=ExperienceRuleOut)
def suspend_agent_memory(rule_id: str, user: User = Depends(get_current_user)) -> ExperienceRuleOut:
    return _set_agent_memory_status(rule_id, "suspended", user)


@router.post("/api/agent/memory/{rule_id}/activate", response_model=ExperienceRuleOut)
def activate_agent_memory(rule_id: str, user: User = Depends(get_current_user)) -> ExperienceRuleOut:
    return _set_agent_memory_status(rule_id, "active", user)
