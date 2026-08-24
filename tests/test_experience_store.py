from pathlib import Path

from backend.experience_store import ExperienceStore


def test_records_a_review_decision_and_suggests_it_for_matching_diff(tmp_path: Path) -> None:
    store = ExperienceStore(tmp_path)
    source_item = {
        "姓名": "张三",
        "工号": "E001",
        "变动项": "月基本薪资",
        "source": "薪资数据",
        "status": "pending",
    }

    rule = store.record(
        tenant_id="tenant-a",
        project_id="project-1",
        created_by="user-1",
        diff_type="salary_change",
        item=source_item,
        match_fields=["变动项", "source"],
        decision="confirmed",
        updates={"审核备注": "本月调薪通知已核验"},
        note="有正式调薪通知时确认。",
    )

    suggestions = store.suggest(
        "tenant-a",
        "salary_change",
        {
            "姓名": "李四",
            "工号": "E002",
            "变动项": "月基本薪资",
            "source": "薪资数据",
            "status": "pending",
        },
    )

    assert rule["decision"] == "confirmed"
    assert rule["conditions"] == {"变动项": "月基本薪资", "source": "薪资数据"}
    assert suggestions == [{
        "rule_id": rule["id"],
        "decision": "confirmed",
        "updates": {"审核备注": "本月调薪通知已核验"},
        "note": "有正式调薪通知时确认。",
        "matched_fields": ["source", "变动项"],
    }]


def test_does_not_match_experience_rules_across_tenants(tmp_path: Path) -> None:
    store = ExperienceStore(tmp_path)
    item = {"变动项": "月基本薪资", "source": "薪资数据"}
    store.record(
        tenant_id="tenant-a",
        project_id="project-1",
        created_by="user-1",
        diff_type="salary_change",
        item=item,
        match_fields=["变动项"],
        decision="ignored",
        updates={},
        note="重复导入，不处理。",
    )

    assert store.suggest("tenant-b", "salary_change", item) == []


def test_ignores_identity_fields_when_recording_conditions(tmp_path: Path) -> None:
    store = ExperienceStore(tmp_path)
    item = {"工号": "E001", "姓名": "张三", "变动项": "月基本薪资"}

    rule = store.record(
        tenant_id="tenant-a",
        project_id="project-1",
        created_by="user-1",
        diff_type="salary_change",
        item=item,
        match_fields=["工号", "姓名", "变动项"],
        decision="confirmed",
        updates={},
        note="按调薪规则处理。",
    )

    assert rule["conditions"] == {"变动项": "月基本薪资"}


def test_ignores_identity_field_aliases_when_recording_conditions(tmp_path: Path) -> None:
    store = ExperienceStore(tmp_path)
    item = {"员工工号": "E001", "人员姓名": "张三", "变动项": "月基本薪资"}

    rule = store.record(
        tenant_id="tenant-a",
        project_id="project-1",
        created_by="user-1",
        diff_type="salary_change",
        item=item,
        match_fields=["员工工号", "人员姓名", "变动项"],
        decision="confirmed",
        updates={},
        note="按调薪规则处理。",
    )

    assert rule["conditions"] == {"变动项": "月基本薪资"}
