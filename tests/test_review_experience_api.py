from types import SimpleNamespace

import pytest
from fastapi import HTTPException

from backend.experience_store import ExperienceStore
from backend.routers import experiences


def test_recording_experience_confirms_current_item_and_keeps_only_reusable_fields(
    tmp_path,
    monkeypatch,
) -> None:
    data = {
        "diff": {
            "salary_change": [{
                "姓名": "张三",
                "工号": "E001",
                "变动项": "月基本薪资",
                "source": "薪资数据",
                "status": "pending",
            }],
            "summary": {},
        },
    }
    saved: list[dict] = []
    monkeypatch.setattr(experiences, "experience_store", ExperienceStore(tmp_path))
    monkeypatch.setattr(experiences, "_load_project_or_404", lambda *_: SimpleNamespace(id="project-1"))
    monkeypatch.setattr(experiences, "_load_json", lambda *_: data)
    monkeypatch.setattr(experiences, "_save_json", lambda *_args: saved.append(_args[2]))

    result = experiences.record_review_experience(
        "project-1",
        experiences.ExperienceRecordIn(
            diff_type="salary_change",
            item_index=0,
            match_fields=["工号", "变动项", "source"],
            decision="confirmed",
            note="已核验调薪通知。",
        ),
        user=SimpleNamespace(id="user-1", tenant_id="tenant-a"),
        db=None,
    )

    assert result.decision == "confirmed"
    assert result.conditions == {"变动项": "月基本薪资", "source": "薪资数据"}
    assert data["diff"]["salary_change"][0]["status"] == "confirmed"
    assert saved == [data]


def test_recording_experience_rejects_identity_only_match_fields(tmp_path, monkeypatch) -> None:
    data = {
        "diff": {
            "salary_change": [{"工号": "E001", "姓名": "张三", "status": "pending"}],
            "summary": {},
        },
    }
    monkeypatch.setattr(experiences, "experience_store", ExperienceStore(tmp_path))
    monkeypatch.setattr(experiences, "_load_project_or_404", lambda *_: SimpleNamespace(id="project-1"))
    monkeypatch.setattr(experiences, "_load_json", lambda *_: data)

    with pytest.raises(HTTPException) as error:
        experiences.record_review_experience(
            "project-1",
            experiences.ExperienceRecordIn(
                diff_type="salary_change",
                item_index=0,
                match_fields=["工号", "姓名"],
                decision="ignored",
                note="仅作测试。",
            ),
            user=SimpleNamespace(id="user-1", tenant_id="tenant-a"),
            db=None,
        )

    assert error.value.status_code == 422


def test_recording_experience_does_not_overwrite_a_previous_review(tmp_path, monkeypatch) -> None:
    data = {
        "diff": {
            "salary_change": [{
                "变动项": "月基本薪资",
                "source": "薪资数据",
                "status": "confirmed",
            }],
            "summary": {},
        },
    }
    monkeypatch.setattr(experiences, "experience_store", ExperienceStore(tmp_path))
    monkeypatch.setattr(experiences, "_load_project_or_404", lambda *_: SimpleNamespace(id="project-1"))
    monkeypatch.setattr(experiences, "_load_json", lambda *_: data)

    with pytest.raises(HTTPException) as error:
        experiences.record_review_experience(
            "project-1",
            experiences.ExperienceRecordIn(
                diff_type="salary_change",
                item_index=0,
                match_fields=["变动项"],
                decision="ignored",
                note="不应覆盖历史审核。",
            ),
            user=SimpleNamespace(id="user-1", tenant_id="tenant-a"),
            db=None,
        )

    assert error.value.status_code == 409


def test_agent_memory_endpoint_records_a_tenant_scoped_rule_with_metadata(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(experiences, "experience_store", ExperienceStore(tmp_path))

    result = experiences.record_agent_memory(
        experiences.AgentMemoryRecordIn(
            diff_type="bonus_source",
            conditions={"人员类别": "店员", "姓名": "张三"},
            decision="confirmed",
            updates={"奖金来源": "J"},
            note="店员优先使用唯一非零来源。",
            period_key="2026-07",
            evidence={"source": "对话"},
        ),
        user=SimpleNamespace(id="user-1", tenant_id="tenant-a"),
    )

    assert result.status == "candidate"
    assert result.conditions == {"人员类别": "店员"}
    assert result.period_key == "2026-07"
    assert result.evidence == {"source": "对话"}


def test_agent_memory_suggestions_are_tenant_isolated_and_include_metadata(tmp_path, monkeypatch) -> None:
    store = ExperienceStore(tmp_path)
    monkeypatch.setattr(experiences, "experience_store", store)
    store.record_memory(
        tenant_id="tenant-a",
        diff_type="bonus_source",
        item={"人员类别": "店员"},
        match_fields=["人员类别"],
        decision="confirmed",
        updates={"奖金来源": "J"},
        note="唯一非零来源。",
        period_key="2026-07",
    )

    result = experiences.suggest_agent_memory(
        experiences.AgentMemorySuggestIn(
            diff_type="bonus_source",
            item={"人员类别": "店员"},
            metadata={"include_rule_metadata": True},
        ),
        user=SimpleNamespace(id="user-2", tenant_id="tenant-a"),
    )
    other_tenant = experiences.suggest_agent_memory(
        experiences.AgentMemorySuggestIn(
            diff_type="bonus_source",
            item={"人员类别": "店员"},
        ),
        user=SimpleNamespace(id="user-3", tenant_id="tenant-b"),
    )

    assert result.total == 1
    assert result.suggestions[0].status == "candidate"
    assert result.suggestions[0].period_key == "2026-07"
    assert other_tenant.total == 0


def test_agent_memory_input_rejects_missing_conditions_and_unknown_fields() -> None:
    with pytest.raises(Exception):
        experiences.AgentMemoryRecordIn(
            diff_type="bonus_source",
            decision="confirmed",
            note="缺少条件。",
        )

    with pytest.raises(Exception):
        experiences.AgentMemoryRecordIn(
            diff_type="bonus_source",
            conditions={"人员类别": "店员"},
            decision="confirmed",
            note="不允许额外字段。",
            unexpected="value",
        )
