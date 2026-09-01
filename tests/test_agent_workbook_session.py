from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import openpyxl
from openpyxl.styles import Font
import pytest

from core.document_agent.orchestrator import ToolExecutionError
from core.document_agent.workbook_session import WorkbookSession


@pytest.fixture
def workbooks(tmp_path: Path) -> SimpleNamespace:
    master = tmp_path / "原始总表.xlsx"
    source = tmp_path / "奖金来源.xlsx"
    draft = tmp_path / "独立草稿.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "工资核算"
    sheet["J2"] = 0
    sheet["K2"] = 100
    sheet["J3"] = "=SUM(J2:K2)"
    sheet["J2"].number_format = "#,##0.00"
    sheet["J2"].font = Font(bold=True, color="FF0000")
    workbook.create_sheet("历史记录")["A1"] = "保留历史数据"
    workbook.save(master)
    workbook.close()
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "奖金明细"
    sheet["B2"] = 2000
    sheet["B3"] = 500
    sheet["B4"] = "=B2+B3"
    sheet["B5"] = "待核对"
    workbook.save(source)
    workbook.close()
    return SimpleNamespace(
        master=master, source=source, draft=draft,
        originals={master: master.read_bytes(), source: source.read_bytes()},
    )


def _session(workbooks: SimpleNamespace) -> WorkbookSession:
    return WorkbookSession(
        master=workbooks.master, draft=workbooks.draft,
        sources={workbooks.source.name: workbooks.source},
    )


def _change(**overrides: Any) -> dict[str, Any]:
    return {
        "source_file": "奖金来源.xlsx", "source_sheet": "奖金明细", "source_cells": ["B2"],
        "target_sheet": "工资核算", "target_cell": "J2", "expected_value": 0,
        "aggregation": "copy", **overrides,
    }


def test_prepare_copy_creates_independent_draft_without_touching_inputs(workbooks: SimpleNamespace) -> None:
    session = _session(workbooks)

    session.prepare_copy()

    assert workbooks.draft.is_file()
    assert workbooks.draft.read_bytes() == workbooks.originals[workbooks.master]
    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


@pytest.mark.parametrize("input_name", ["master", "source"])
def test_prepare_copy_rejects_draft_that_is_any_input(workbooks: SimpleNamespace, input_name: str) -> None:
    with pytest.raises(ToolExecutionError):
        session = WorkbookSession(
            master=workbooks.master, draft=getattr(workbooks, input_name),
            sources={workbooks.source.name: workbooks.source},
        )
        session.prepare_copy()

    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


def test_apply_source_cells_copies_and_sums_with_auditable_updates(workbooks: SimpleNamespace) -> None:
    session = _session(workbooks)
    session.prepare_copy()

    updates = session.apply_source_cells([
        _change(),
        _change(source_cells=["B2", "B3"], target_cell="K2", expected_value=100, aggregation="sum"),
    ])

    assert isinstance(updates, list)
    assert [{key: update[key] for key in (
        "old_value", "new_value", "target_sheet", "target_cell", "source_file", "source_sheet",
    )} for update in updates] == [
        {"old_value": 0, "new_value": 2000, "target_sheet": "工资核算", "target_cell": "J2",
         "source_file": "奖金来源.xlsx", "source_sheet": "奖金明细"},
        {"old_value": 100, "new_value": 2500, "target_sheet": "工资核算", "target_cell": "K2",
         "source_file": "奖金来源.xlsx", "source_sheet": "奖金明细"},
    ]
    workbook = openpyxl.load_workbook(workbooks.draft, data_only=False)
    try:
        sheet = workbook["工资核算"]
        assert sheet["J2"].value == 2000
        assert sheet["K2"].value == 2500
        assert sheet["J3"].value == "=SUM(J2:K2)"
        assert sheet["J2"].number_format == "#,##0.00"
        assert sheet["J2"].font.bold is True
        assert workbook["历史记录"]["A1"].value == "保留历史数据"
    finally:
        workbook.close()
    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


@pytest.mark.parametrize("change", [
    pytest.param(_change(target_cell="J3", expected_value="=SUM(J2:K2)"), id="target-formula"),
    pytest.param(_change(source_cells=["B4"]), id="source-formula-without-cache"),
    pytest.param(_change(source_file="未授权来源.xlsx"), id="unregistered-source"),
    pytest.param(_change(expected_value=999), id="stale-target-value"),
    pytest.param(_change(source_sheet="不存在的来源页"), id="unknown-source-sheet"),
    pytest.param(_change(target_sheet="不存在的目标页"), id="unknown-target-sheet"),
    pytest.param(_change(source_cells=["B0"]), id="invalid-source-cell"),
    pytest.param(_change(target_cell="J0"), id="invalid-target-cell"),
    pytest.param(_change(aggregation="execute"), id="unsupported-aggregation"),
    pytest.param(_change(source_cells=["B2", "B5"], aggregation="sum"), id="nonnumeric-sum"),
])
def test_invalid_change_is_sanitized_and_leaves_draft_and_inputs_unchanged(
    workbooks: SimpleNamespace, change: dict[str, Any],
) -> None:
    session = _session(workbooks)
    session.prepare_copy()
    draft_before = workbooks.draft.read_bytes()

    with pytest.raises(ToolExecutionError):
        session.apply_source_cells([change])

    assert workbooks.draft.read_bytes() == draft_before
    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


@pytest.mark.parametrize("invalid_change", [
    _change(source_cells=["B4"], target_cell="K2", expected_value=100),
    _change(target_cell="K2", expected_value=999),
])
def test_batch_failure_never_persists_an_earlier_valid_change(
    workbooks: SimpleNamespace, invalid_change: dict[str, Any],
) -> None:
    session = _session(workbooks)
    session.prepare_copy()
    draft_before = workbooks.draft.read_bytes()

    with pytest.raises(ToolExecutionError):
        session.apply_source_cells([_change(), invalid_change])

    assert workbooks.draft.read_bytes() == draft_before
    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


def test_insert_and_copy_row_preserves_style_and_translates_formulas(workbooks: SimpleNamespace) -> None:
    workbook = openpyxl.load_workbook(workbooks.master)
    try:
        sheet = workbook["工资核算"]
        sheet["A2"] = "2026年6月"
        sheet["B2"] = 100
        sheet["C2"] = "=B2*2"
        sheet.row_dimensions[2].height = 24
        sheet["B2"].font = Font(bold=True, color="00FF00")
        workbook.save(workbooks.master)
    finally:
        workbook.close()
    workbooks.originals[workbooks.master] = workbooks.master.read_bytes()
    session = _session(workbooks)
    session.prepare_copy()

    update = session.insert_and_copy_row({
        "sheet": "工资核算",
        "source_row": 2,
        "insert_before_row": 3,
        "expected_source_cells": {"A2": "2026年6月", "C2": "=B2*2"},
        "copy_max_column": 3,
    })

    assert update == {
        "sheet": "工资核算", "source_row": 2, "inserted_row": 3,
        "copied_columns": 3,
    }
    workbook = openpyxl.load_workbook(workbooks.draft, data_only=False)
    try:
        sheet = workbook["工资核算"]
        assert [sheet.cell(3, column).value for column in range(1, 4)] == ["2026年6月", 100, "=B3*2"]
        assert sheet.row_dimensions[3].height == 24
        assert sheet["B3"].font.bold is True
        assert sheet["B3"].font.color.rgb == "0000FF00"
    finally:
        workbook.close()
    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


@pytest.mark.parametrize("change", [
    {"sheet": "工资核算", "source_row": 2, "insert_before_row": 2,
     "expected_source_cells": {"A2": "2026年6月"}, "copy_max_column": 3},
    {"sheet": "工资核算", "source_row": 2, "insert_before_row": 3,
     "expected_source_cells": {"A2": "stale"}, "copy_max_column": 3},
    {"sheet": "不存在", "source_row": 2, "insert_before_row": 3,
     "expected_source_cells": {"A2": "2026年6月"}, "copy_max_column": 3},
])
def test_insert_and_copy_row_rejects_invalid_or_stale_requests_atomically(
    workbooks: SimpleNamespace, change: dict[str, Any],
) -> None:
    workbook = openpyxl.load_workbook(workbooks.master)
    try:
        workbook["工资核算"]["A2"] = "2026年6月"
        workbook.save(workbooks.master)
    finally:
        workbook.close()
    workbooks.originals[workbooks.master] = workbooks.master.read_bytes()
    session = _session(workbooks)
    session.prepare_copy()
    draft_before = workbooks.draft.read_bytes()

    with pytest.raises(ToolExecutionError):
        session.insert_and_copy_row(change)

    assert workbooks.draft.read_bytes() == draft_before
    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


def test_apply_verified_values_updates_non_formula_cells_without_changing_styles(workbooks: SimpleNamespace) -> None:
    session = _session(workbooks)
    session.prepare_copy()

    updates = session.apply_verified_values([
        {"target_sheet": "工资核算", "target_cell": "J2", "expected_value": 0, "new_value": None},
        {"target_sheet": "工资核算", "target_cell": "K2", "expected_value": 100, "new_value": 23},
    ])

    assert updates == [
        {"target_sheet": "工资核算", "target_cell": "J2", "old_value": 0, "new_value": None},
        {"target_sheet": "工资核算", "target_cell": "K2", "old_value": 100, "new_value": 23},
    ]
    workbook = openpyxl.load_workbook(workbooks.draft, data_only=False)
    try:
        sheet = workbook["工资核算"]
        assert sheet["J2"].value is None
        assert sheet["J2"].number_format == "#,##0.00"
        assert sheet["J2"].font.bold is True
        assert sheet["K2"].value == 23
        assert sheet["J3"].value == "=SUM(J2:K2)"
    finally:
        workbook.close()
    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


@pytest.mark.parametrize("changes", [
    [{"target_sheet": "工资核算", "target_cell": "J3", "expected_value": "=SUM(J2:K2)", "new_value": 1}],
    [{"target_sheet": "工资核算", "target_cell": "J2", "expected_value": 999, "new_value": 1}],
    [{"target_sheet": "工资核算", "target_cell": "J2", "expected_value": 0, "new_value": "=1+1"}],
    [
        {"target_sheet": "工资核算", "target_cell": "J2", "expected_value": 0, "new_value": 1},
        {"target_sheet": "工资核算", "target_cell": "J2", "expected_value": 0, "new_value": 2},
    ],
])
def test_apply_verified_values_rejects_unsafe_requests_atomically(
    workbooks: SimpleNamespace, changes: list[dict[str, Any]],
) -> None:
    session = _session(workbooks)
    session.prepare_copy()
    draft_before = workbooks.draft.read_bytes()

    with pytest.raises(ToolExecutionError):
        session.apply_verified_values(changes)

    assert workbooks.draft.read_bytes() == draft_before
    assert {path: path.read_bytes() for path in workbooks.originals} == workbooks.originals


def test_roll_forward_summary_row_keeps_current_formulas_and_freezes_history(workbooks: SimpleNamespace) -> None:
    workbook = openpyxl.load_workbook(workbooks.master)
    try:
        sheet = workbook["工资核算"]
        sheet["A2"] = "6月"
        sheet["B2"] = 40
        sheet["C2"] = "=B2*2"
        workbook.save(workbooks.master)
    finally:
        workbook.close()
    workbooks.originals[workbooks.master] = workbooks.master.read_bytes()
    session = _session(workbooks)
    session.prepare_copy()

    update = session.roll_forward_summary_row({
        "sheet": "工资核算", "current_row": 2, "expected_current_month": "6月", "new_month": "8月",
        "start_column": 1, "end_column": 3,
        "history_values": {"A2": "6月", "B2": 40, "C2": 80},
    })

    assert update == {"sheet": "工资核算", "current_row": 2, "history_row": 3, "new_month": "8月"}
    workbook = openpyxl.load_workbook(workbooks.draft, data_only=False)
    try:
        sheet = workbook["工资核算"]
        assert [sheet.cell(2, column).value for column in range(1, 4)] == ["8月", 40, "=B2*2"]
        assert [sheet.cell(3, column).value for column in range(1, 4)] == ["6月", 40, 80]
    finally:
        workbook.close()


def test_roll_forward_summary_row_rejects_missing_history_values_atomically(workbooks: SimpleNamespace) -> None:
    session = _session(workbooks)
    session.prepare_copy()
    draft_before = workbooks.draft.read_bytes()

    with pytest.raises(ToolExecutionError):
        session.roll_forward_summary_row({
            "sheet": "工资核算", "current_row": 2, "expected_current_month": None, "new_month": "8月",
            "start_column": 1, "end_column": 3,
            "history_values": {"A2": None},
        })

    assert workbooks.draft.read_bytes() == draft_before
