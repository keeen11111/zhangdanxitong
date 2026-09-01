from pathlib import Path

import openpyxl

from backend.routers.agent import _build_run_tool_registry
from core.document_agent import ToolCall


def _workbook(path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "工资核算"
    sheet["J2"] = 2000
    sheet["K2"] = 0
    workbook.save(path)
    workbook.close()


def test_run_registry_reads_only_the_scoped_workbook(tmp_path: Path) -> None:
    path = tmp_path / "draft.xlsx"
    _workbook(path)
    run = {
        "run_id": "run-1",
        "project_id": "project-1",
        "salary_month": "2026.07",
        "draft_filename": path.name,
        "items": [{"id": "item-1", "person_key": "E001", "person_name": "李某", "status": "needs_review"}],
    }
    registry = _build_run_tool_registry(run, workbook_path=path)
    result = registry.execute(ToolCall(
        call_id="call-1", name="read_range",
        arguments={"sheet": "工资核算", "range": "J2:K2"},
    ))
    assert result.ok is True
    assert result.output["values"] == [[2000, 0]]


def test_run_registry_never_applies_an_ambiguous_item(tmp_path: Path) -> None:
    path = tmp_path / "draft.xlsx"
    _workbook(path)
    run = {
        "run_id": "run-1", "project_id": "project-1", "salary_month": "2026.07",
        "draft_filename": path.name,
        "items": [{
            "id": "item-1", "person_key": "E001", "status": "needs_review",
            "issue_type": "conflicting_value", "target_sheet": "工资核算",
            "target_cell": "J2", "candidate_values": [2000, 3000],
        }],
    }
    registry = _build_run_tool_registry(run, workbook_path=path)
    result = registry.execute(ToolCall(
        call_id="call-2", name="apply_cell_changes",
        arguments={"item_id": "item-1", "value": 3000},
    ))
    assert result.ok is False
    assert "人工确认" in str(result.error)
    workbook = openpyxl.load_workbook(path, data_only=False)
    assert workbook["工资核算"]["J2"].value == 2000
    workbook.close()


def test_run_registry_selects_only_a_candidate_sheet_and_calls_reintegration(tmp_path: Path) -> None:
    path = tmp_path / "draft.xlsx"
    _workbook(path)
    calls: list[tuple[str, str, str]] = []
    run = {
        "run_id": "run-1",
        "project_id": "project-1",
        "salary_month": "2026.07",
        "draft_filename": path.name,
        "items": [{
            "id": "item-1",
            "person_key": "E001",
            "issue_type": "ambiguous_sheet",
            "source_files": ["更新表.xlsx"],
            "source_sheets": ["奖金-7月"],
            "candidate_target_sheets": ["工资核算", "台账"],
            "status": "needs_review",
        }],
    }
    registry = _build_run_tool_registry(
        run,
        workbook_path=path,
        sheet_mapping_handler=lambda item, target, reason: calls.append((item["id"], target, reason)) or {"applied": True},
    )
    result = registry.execute(ToolCall(
        call_id="call-map",
        name="select_sheet_mapping",
        arguments={"item_id": "item-1", "target_sheet": "工资核算", "reason": "奖金明细应进入工资核算"},
    ))
    assert result.ok is True
    assert run["sheet_mappings"] == [{
        "source_file": "更新表.xlsx", "source_sheet": "奖金-7月", "target_sheet": "工资核算",
        "reason": "奖金明细应进入工资核算",
    }]
    assert calls == [("item-1", "工资核算", "奖金明细应进入工资核算")]

    rejected = registry.execute(ToolCall(
        call_id="call-map-invalid",
        name="select_sheet_mapping",
        arguments={"item_id": "item-1", "target_sheet": "不存在", "reason": "猜测"},
    ))
    assert rejected.ok is False
    assert "候选" in str(rejected.error)
