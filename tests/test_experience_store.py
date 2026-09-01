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


def test_records_company_memory_metadata_and_starts_as_candidate(tmp_path: Path) -> None:
    store = ExperienceStore(tmp_path)

    rule = store.record_memory(
        tenant_id="tenant-a",
        project_id="project-1",
        created_by="user-1",
        diff_type="bonus_source",
        item={"人员类别": "店员", "姓名": "张三", "工号": "E001"},
        match_fields=["人员类别"],
        decision="confirmed",
        updates={"奖金来源": "J"},
        note="店员唯一非零列优先。",
        scope="company",
        period_key="2026-07",
        evidence={"source": "用户对话", "message_id": "m-1"},
        rule_version="keyuan-2026.08.1",
    )

    assert rule["status"] == "candidate"
    assert rule["scope"] == "company"
    assert rule["period_key"] == "2026-07"
    assert rule["evidence"] == {"source": "用户对话", "message_id": "m-1"}
    assert rule["use_count"] == 1
    assert rule["consistent_count"] == 1
    assert rule["conflict_count"] == 0
    assert rule["rule_version"] == "keyuan-2026.08.1"
    assert "姓名" not in rule["conditions"]
    assert "工号" not in rule["conditions"]


def test_same_company_memory_is_idempotent_per_period_and_activates_after_three_periods(
    tmp_path: Path,
) -> None:
    store = ExperienceStore(tmp_path)
    kwargs = {
        "tenant_id": "tenant-a",
        "project_id": "project-1",
        "created_by": "user-1",
        "diff_type": "bonus_source",
        "item": {"人员类别": "店员"},
        "match_fields": ["人员类别"],
        "decision": "confirmed",
        "updates": {"奖金来源": "J"},
        "note": "唯一非零列。",
        "scope": "company",
        "rule_version": "keyuan-1",
    }

    first = store.record_memory(**kwargs, period_key="2026-05")
    duplicate = store.record_memory(**kwargs, period_key="2026-05")
    second = store.record_memory(**kwargs, period_key="2026-06")
    third = store.record_memory(**kwargs, period_key="2026-07")

    assert duplicate["id"] == first["id"]
    assert duplicate["use_count"] == 1
    assert second["use_count"] == 2
    assert third["use_count"] == 3
    assert third["consistent_count"] == 3
    assert third["status"] == "active"


def test_consistency_accumulates_when_rule_package_version_changes(tmp_path: Path) -> None:
    store = ExperienceStore(tmp_path)
    for period, version in (("2026-05", "keyuan-1"), ("2026-06", "keyuan-2"), ("2026-07", "keyuan-3")):
        result = store.record_memory(
            tenant_id="tenant-a",
            diff_type="bonus_source",
            item={"人员类别": "店员"},
            match_fields=["人员类别"],
            decision="confirmed",
            updates={"奖金来源": "J"},
            note="唯一非零列。",
            scope="company",
            period_key=period,
            rule_version=version,
        )

    assert result["use_count"] == 3
    assert result["consistent_count"] == 3
    assert result["status"] == "active"


def test_conflicting_memory_is_suspended_and_recorded_as_a_new_candidate(tmp_path: Path) -> None:
    store = ExperienceStore(tmp_path)
    base = {
        "tenant_id": "tenant-a",
        "project_id": "project-1",
        "created_by": "user-1",
        "diff_type": "bonus_source",
        "item": {"人员类别": "店员"},
        "match_fields": ["人员类别"],
        "updates": {},
        "note": "用户确认。",
        "scope": "company",
    }
    original = store.record_memory(**base, decision="confirmed", period_key="2026-06")
    replacement = store.record_memory(**base, decision="ignored", period_key="2026-07")

    rules, total = store.list_rules("tenant-a", offset=0, limit=20)

    assert total == 2
    assert replacement["id"] != original["id"]
    assert replacement["status"] == "candidate"
    old = next(rule for rule in rules if rule["id"] == original["id"])
    assert old["status"] == "suspended"
    assert old["conflict_count"] == 1


def test_suggest_can_include_memory_metadata_without_changing_legacy_default(tmp_path: Path) -> None:
    store = ExperienceStore(tmp_path)
    rule = store.record_memory(
        tenant_id="tenant-a",
        project_id="project-1",
        created_by="user-1",
        diff_type="bonus_source",
        item={"人员类别": "店员"},
        match_fields=["人员类别"],
        decision="confirmed",
        updates={"奖金来源": "J"},
        note="唯一非零列。",
        period_key="2026-07",
        evidence=[{"sheet": "奖金-7月", "cell": "J12"}],
    )

    legacy = store.suggest("tenant-a", "bonus_source", {"人员类别": "店员"})
    enriched = store.suggest(
        "tenant-a",
        "bonus_source",
        {"人员类别": "店员"},
        metadata={"include_rule_metadata": True},
    )

    assert legacy == [{
        "rule_id": rule["id"],
        "decision": "confirmed",
        "updates": {"奖金来源": "J"},
        "note": "唯一非零列。",
        "matched_fields": ["人员类别"],
    }]
    assert enriched[0]["status"] == "candidate"
    assert enriched[0]["period_key"] == "2026-07"
    assert enriched[0]["evidence"] == [{"sheet": "奖金-7月", "cell": "J12"}]


def test_bonus_selector_returns_only_unique_nonzero_source() -> None:
    from backend.experience_store import select_unique_nonzero_source

    assert select_unique_nonzero_source({"J": 2000, "K": 0}) == ("J", 2000)
    assert select_unique_nonzero_source({"J": 0, "K": 2000}) == ("K", 2000)
    assert select_unique_nonzero_source({"J": 2000, "K": 3000}) is None
    assert select_unique_nonzero_source({"J": 0, "K": 0}) is None
