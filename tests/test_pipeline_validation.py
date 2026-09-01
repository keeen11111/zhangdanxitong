"""Pipeline 最终库校验的特征测试，不读取真实薪资数据。"""

import os

import pytest
from fastapi import HTTPException

from backend.models import UploadFile
from backend.routers import pipeline
from backend.routers.pipeline import _ensure_exportable, _reuse_cached_base_db, _validate_final_entities


class _FakeQuery:
    def __init__(self, rows):
        self.rows = rows

    def filter(self, *args, **kwargs):
        return self

    def order_by(self, *args, **kwargs):
        return self

    def all(self):
        return self.rows


class _FakeDb:
    def __init__(self, rows):
        self.rows = rows

    def query(self, model):
        return _FakeQuery(self.rows)


def test_confirming_department_transfer_updates_only_the_selected_employee() -> None:
    final_data = {
        "entities": {
            "employee_profile": [
                {"工号": "E001", "姓名": "张三", "部门": "原部门"},
                {"工号": "E002", "姓名": "李四", "部门": "原部门二"},
            ],
        },
        "issues": [
            {
                "issue_id": "department-e001",
                "issue_type": "department_transfer",
                "employee_id": "E001",
                "proposed_value": "新部门",
                "status": "pending",
            },
            {
                "issue_id": "formula-review",
                "issue_type": "manual_formula_review",
                "status": "pending",
            },
        ],
    }

    updated = pipeline._confirm_department_transfers(final_data, ["department-e001"])

    assert updated == 1
    assert final_data["entities"]["employee_profile"][0]["部门"] == "新部门"
    assert final_data["entities"]["employee_profile"][1]["部门"] == "原部门二"
    assert final_data["issues"][0]["status"] == "confirmed"
    assert final_data["issues"][1]["status"] == "pending"


def test_legacy_department_transfer_is_hydrated_when_employee_and_target_are_unambiguous() -> None:
    final_data = {
        "entities": {
            "employee_profile": [
                {"工号": "E001", "姓名": "张三", "部门": "原部门"},
            ],
        },
        "issues": [
            {
                "issue_type": "department_transfer",
                "person_name": "张三",
                "message": "部门从总表值“原部门”变更为“新部门”，未自动覆盖，请人工确认。",
            },
        ],
    }

    pipeline._hydrate_legacy_department_transfer_issues(final_data)

    assert final_data["issues"][0]["employee_id"] == "E001"
    assert final_data["issues"][0]["current_value"] == "原部门"
    assert final_data["issues"][0]["proposed_value"] == "新部门"


def test_legacy_department_transfer_stays_unselectable_when_target_has_multiple_values() -> None:
    final_data = {
        "entities": {
            "employee_profile": [
                {"工号": "E001", "姓名": "张三", "部门": "原部门"},
            ],
        },
        "issues": [
            {
                "issue_type": "department_transfer",
                "person_name": "张三",
                "message": "部门从总表值“原部门”变更为“新部门一、新部门二”，未自动覆盖，请人工确认。",
            },
        ],
    }

    pipeline._hydrate_legacy_department_transfer_issues(final_data)

    assert final_data["issues"][0].get("proposed_value") is None


def test_template_selection_rejects_source_only_project_instead_of_using_builtin(tmp_path, monkeypatch) -> None:
    source_path = tmp_path / "source.xlsx"
    builtin_path = tmp_path / "builtin.xlsx"
    source_path.write_bytes(b"source")
    builtin_path.write_bytes(b"builtin")
    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir()
    (upload_dir / "source.xlsx").write_bytes(source_path.read_bytes())

    source_file = UploadFile(
        id="source-id",
        project_id="project-id",
        original_name="工资汇总表.xlsx",
        stored_path="source.xlsx",
        file_type="source",
    )
    monkeypatch.setattr(pipeline, "UPLOAD_DIR", str(upload_dir))
    monkeypatch.setattr(pipeline, "BUILTIN_TEMPLATE_PATH", str(builtin_path))
    monkeypatch.setattr(
        pipeline,
        "_supports_unified_export",
        lambda path: os.path.abspath(path) == os.path.abspath(str(builtin_path)),
    )

    selected = pipeline._select_template_file("project-id", _FakeDb([source_file]))

    assert selected is None


def test_template_selection_only_uses_explicit_master_role(tmp_path, monkeypatch) -> None:
    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir()
    (upload_dir / "master.xlsx").write_bytes(b"master")
    (upload_dir / "change.xlsx").write_bytes(b"change")
    master = UploadFile(
        id="master-id",
        project_id="project-id",
        original_name="本月工资总表.xlsx",
        stored_path="master.xlsx",
        file_type="template",
    )
    misleading_change = UploadFile(
        id="change-id",
        project_id="project-id",
        original_name="工资核算变更.xlsx",
        stored_path="change.xlsx",
        file_type="source",
    )
    monkeypatch.setattr(pipeline, "UPLOAD_DIR", str(upload_dir))
    monkeypatch.setattr(pipeline, "_supports_unified_export", lambda path: True)

    selected = pipeline._select_template_file(
        "project-id",
        _FakeDb([misleading_change, master]),
    )

    assert selected == (master, str(upload_dir / "master.xlsx"))


def test_template_selection_accepts_existing_financial_master_for_payroll_export(tmp_path, monkeypatch) -> None:
    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir()
    (upload_dir / "master.xlsx").write_bytes(b"master")
    master = UploadFile(
        id="master-id",
        project_id="project-id",
        original_name="本月工资总表.xlsx",
        stored_path="master.xlsx",
        file_type="financial_master",
    )
    monkeypatch.setattr(pipeline, "UPLOAD_DIR", str(upload_dir))
    monkeypatch.setattr(pipeline, "_supports_unified_export", lambda path: True)

    selected = pipeline._select_template_file("project-id", _FakeDb([master]))

    assert selected == (master, str(upload_dir / "master.xlsx"))


def test_template_selection_blocks_multiple_explicit_masters_even_when_one_is_a_change(tmp_path, monkeypatch) -> None:
    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir()
    for name in ("master.xlsx", "misclassified-change.xlsx"):
        (upload_dir / name).write_bytes(b"workbook")
    first = UploadFile(
        id="master-id",
        project_id="project-id",
        original_name="总表.xlsx",
        stored_path="master.xlsx",
        file_type="template",
    )
    second = UploadFile(
        id="change-id",
        project_id="project-id",
        original_name="人员异动.xlsx",
        stored_path="misclassified-change.xlsx",
        file_type="template",
    )
    monkeypatch.setattr(pipeline, "UPLOAD_DIR", str(upload_dir))
    monkeypatch.setattr(pipeline, "_supports_unified_export", lambda path: True)

    assert pipeline._select_template_file("project-id", _FakeDb([first, second])) is None


def test_source_signature_counts_only_master_and_change_roles() -> None:
    rows = [
        UploadFile(id="master", original_name="总表.xlsx", stored_path="m", file_type="template"),
        UploadFile(id="change", original_name="考勤.xlsx", stored_path="c", file_type="source"),
        UploadFile(id="legacy", original_name="旧分类.xlsx", stored_path="l", file_type="current"),
    ]

    signature = pipeline._project_source_signature(rows)

    assert signature == pipeline._project_source_signature(rows[:2])


def test_source_signature_includes_existing_financial_upload_roles() -> None:
    rows = [
        UploadFile(id="master", original_name="总表.xlsx", stored_path="m", file_type="financial_master"),
        UploadFile(id="change", original_name="考勤.xlsx", stored_path="c", file_type="financial_source"),
    ]

    assert pipeline._project_source_signature(rows) != pipeline._project_source_signature([])


def test_semantic_issue_key_is_stable_for_the_same_source_task() -> None:
    issue = {
        "issue_type": "empty_source_sheet",
        "source_files": ["来源.xlsx"],
        "source_sheets": ["补发补扣"],
        "message": "来源 Sheet 没有可处理的数据记录，未自动写入。",
    }

    assert pipeline._semantic_issue_key(issue) == pipeline._semantic_issue_key(dict(issue))


def test_pipeline_validation_accepts_complete_employee() -> None:
    entities = {
        "employee_profile": [
            {
                "工号": "E001",
                "姓名": "测试员工",
                "身份证号": "110101199001011234",
                "转正日期": "2026-01-01",
                "是否试用期": "否",
            }
        ],
        "salary_detail": [
            {
                "工号": "E001",
                "月基本薪资": 8000,
                "交通通讯补贴标准": 300,
            }
        ],
    }

    result = _validate_final_entities(entities)

    assert result["metrics"]["high_risk_count"] == 0
    assert result["blockers"] == []


def test_pipeline_validation_blocks_missing_identity_and_salary() -> None:
    entities = {
        "employee_profile": [{"工号": "E002", "姓名": "缺项员工"}],
        "salary_detail": [],
    }

    result = _validate_final_entities(entities)

    assert result["metrics"]["high_risk_count"] == 1
    assert result["blockers"][0]["employee_id"] == "E002"
    assert "身份证号" in result["blockers"][0]["missing_fields"]
    assert "基本工资" in result["blockers"][0]["missing_fields"]


def test_pipeline_export_is_rejected_when_blockers_exist() -> None:
    entities = {
        "employee_profile": [{"工号": "E003", "姓名": "待补齐员工"}],
        "salary_detail": [],
    }

    with pytest.raises(HTTPException) as exc_info:
        _ensure_exportable(entities)

    assert exc_info.value.status_code == 400
    assert "阻断项" in str(exc_info.value.detail)


def test_base_db_refresh_bypasses_cached_session() -> None:
    cached = {"entities": {"employee_profile": [{"工号": "OLD"}]}}

    assert _reuse_cached_base_db(cached, refresh=False) is True
    assert _reuse_cached_base_db(cached, refresh=True) is False
