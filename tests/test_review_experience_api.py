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
