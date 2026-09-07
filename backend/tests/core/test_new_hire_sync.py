"""Characterize the existing mapper and exercise standalone hire notices.

Synthetic data deliberately contains no customer identifiers or pay amounts.
"""
from pathlib import Path
from datetime import datetime
from types import SimpleNamespace

from openpyxl import Workbook, load_workbook
import pytest

from core.sheet_mapper import apply_semantic_sheet_updates


def save_book(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    book = Workbook()
    first = True
    for title, rows in sheets.items():
        sheet = book.active if first else book.create_sheet()
        first = False
        sheet.title = title
        for row in rows:
            sheet.append(row)
    book.save(path)
    book.close()


def test_existing_non_hire_ambiguity_stays_unresolved(tmp_path: Path) -> None:
    master, source, output = [tmp_path / name for name in ("master.xlsx", "source.xlsx", "output.xlsx")]
    save_book(master, {
        "人员档案": [["工号", "姓名", "部门"], ["001", "测试甲", "原部门"]],
        "人员异动": [["工号", "姓名", "部门"], ["001", "测试甲", "原部门"]],
    })
    save_book(source, {"变更": [["工号", "姓名", "部门"], ["001", "测试甲", "新部门"]]})
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["auto_update_count"] == 0
    assert result["issues"][0]["issue_type"] == "ambiguous_sheet"


def test_existing_mapper_preserves_formula_when_updating_known_person(tmp_path: Path) -> None:
    master, source, output = [tmp_path / name for name in ("master.xlsx", "source.xlsx", "output.xlsx")]
    save_book(master, {"档案": [["工号", "姓名", "部门", "计算"], ["001", "测试甲", "原部门", "=1+2"]]})
    save_book(source, {"变更": [["工号", "姓名", "部门"], ["001", "测试甲", "新部门"]]})
    result = apply_semantic_sheet_updates(master, [source], output)
    book = load_workbook(output)
    try:
        assert result["issues"] == []
        assert book.active["C2"].value == "新部门"
        assert book.active["D2"].value == "=1+2"
    finally:
        book.close()


def hire_files(tmp_path: Path) -> tuple[Path, Path, Path]:
    master, source, output = [tmp_path / name for name in ("master.xlsx", "source.xlsx", "output.xlsx")]
    save_book(master, {
        "员工档案": [["工号", "姓名", "公司", "部门", "入职日期"], ["001", "测试甲", "公司甲", "部门甲"]],
        "工资核算": [
            ["工号", "姓名", "月基本薪资", "月绩效标准", "合计"],
            ["001", "测试甲", 100, 20, "=C2+D2"],
            ["合计", None, "=SUM(C2:C2)", "=SUM(D2:D2)", "=SUM(E2:E2)"],
        ],
        "社保登记": [["工号", "姓名", "社保基数", "公积金基数"], ["001", "测试甲", 100, 100]],
        "奖金明细": [["工号", "姓名", "业绩奖金"], ["001", "测试甲", 10]],
        "汇总": [["工资合计"], ["='工资核算'!E3"]],
    })
    save_book(source, {"Sheet1": [
        ["入职"],
        ["姓名", "新工号", "新公司", "新成本中心/部门", "入职时间", "基本工资", "绩效工资", "社保基数", "公积金基数"],
        ["测试乙", "002", "公司甲", "部门乙", datetime(2026, 5, 12), 200, 40, 200, 200],
    ]})
    return master, source, output


def test_hire_notice_fans_out_to_related_sheets_and_extends_totals(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["issues"] == []
    book = load_workbook(output)
    try:
        assert book["员工档案"]["A3"].value == "002"
        assert book["员工档案"]["E3"].value == datetime(2026, 5, 12)
        assert book["工资核算"]["C3"].value == 200
        assert book["工资核算"]["E3"].value == "=C3+D3"
        assert book["工资核算"]["E4"].value == "=SUM(E2:E3)"
        assert book["汇总"]["A2"].value == "='工资核算'!E4"
        assert book["社保登记"]["D3"].value == 200
        assert book["奖金明细"].max_row == 2
        assert len(book.sheetnames) == 5
    finally:
        book.close()
    assert {m["target_sheet"] for m in result["matches"]} == {"员工档案", "工资核算", "社保登记"}
    assert result["personnel_coverage"]["processed_person_count"] == 1
    assert result["personnel_coverage"]["match_rate"] == 1.0
    assert all(u["source_file"] == source.name for u in result["updates"])


def test_hire_repeated_run_is_idempotent(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    apply_semantic_sheet_updates(master, [source], output)
    second = tmp_path / "second.xlsx"
    result = apply_semantic_sheet_updates(output, [source], second)
    assert result["issues"] == []
    assert result["auto_update_count"] == 0
    book = load_workbook(second)
    try:
        assert book["工资核算"].max_row == 4
        assert book["员工档案"].max_row == 3
    finally:
        book.close()


def test_hire_identity_conflict_prevents_partial_cross_sheet_writes(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(master)
    book["社保登记"].append(["002", "另一人", 90, 90])
    book.save(master)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["auto_update_count"] == 0
    assert any(i["issue_type"] == "hire_identity_conflict" for i in result["issues"])
    assert result["personnel_coverage"]["match_rate"] == 0
    book = load_workbook(output)
    try:
        assert book["员工档案"].max_row == 2
    finally:
        book.close()


def test_hire_keeps_allowance_standard_and_payment_note_distinct(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(master)
    sheet = book.create_sheet("薪资标准")
    sheet.append(["工号", "姓名", "高温费", "高温费备注", "本月高温费"])
    sheet.append(["001", "测试甲", 800, "6-9月发放，每月200元", "=0"])
    book.save(master)
    book.close()
    book = load_workbook(source)
    book.active["J2"] = "高温费"
    book.active["K2"] = "高温费别再"
    book.active["J3"] = 800
    book.active["K3"] = "6-9月发放，每月200元"
    book.save(source)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output, salary_month="2026-05")
    assert result["issues"] == []
    book = load_workbook(output)
    try:
        assert book["薪资标准"]["C3"].value == 800
        assert book["薪资标准"]["D3"].value == "6-9月发放，每月200元"
        assert book["薪资标准"]["E3"].value == "=0"
    finally:
        book.close()


def test_insert_does_not_expand_previous_employees_horizontal_sum(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(master)
    book["工资核算"]["E2"] = "=SUM(C2:D2)"
    book.save(master)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["issues"] == []
    book = load_workbook(output)
    try:
        assert book["工资核算"]["E2"].value == "=SUM(C2:D2)"
        assert book["工资核算"]["E3"].value == "=SUM(C3:D3)"
    finally:
        book.close()


def test_hire_does_not_claim_full_match_when_source_field_has_no_target(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(source)
    book.active["J2"] = "银行卡号"
    book.active["J3"] = "00001234"
    book.save(source)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["auto_update_count"] == 0
    assert result["personnel_coverage"]["match_rate"] == 0
    assert any(issue["issue_type"] == "hire_unmapped_fields" for issue in result["issues"])


def test_hire_source_conflicts_are_checked_before_writing(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    second = tmp_path / "second_notice.xlsx"
    book = load_workbook(source)
    book.active["F3"] = 999
    book.save(second)
    book.close()
    result = apply_semantic_sheet_updates(master, [source, second], output)
    assert result["auto_update_count"] == 0
    assert result["issues"][0]["issue_type"] == "hire_conflicting_source"


def test_homonymous_new_hires_are_blocked_independent_of_source_order(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    second = tmp_path / "second_notice.xlsx"
    book = load_workbook(source)
    book.active["B3"] = "003"
    book.save(second)
    book.close()
    for sources in ([source, second], [second, source]):
        result = apply_semantic_sheet_updates(master, sources, output)
        assert result["auto_update_count"] == 0
        assert result["personnel_coverage"]["input_person_count"] == 2
        assert result["personnel_coverage"]["processed_person_count"] == 0


def test_input_workbooks_and_unrelated_employee_values_are_preserved(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    master_bytes, source_bytes = master.read_bytes(), source.read_bytes()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["issues"] == []
    assert master.read_bytes() == master_bytes
    assert source.read_bytes() == source_bytes
    book = load_workbook(output)
    try:
        assert list(book["员工档案"].values)[1] == ("001", "测试甲", "公司甲", "部门甲", None)
        assert list(book["奖金明细"].values)[1] == ("001", "测试甲", 10)
    finally:
        book.close()


def test_insertion_preserves_styles_print_area_and_absolute_total_references(tmp_path: Path) -> None:
    from openpyxl.styles import PatternFill
    from openpyxl.workbook.defined_name import DefinedName

    master, source, output = hire_files(tmp_path)
    book = load_workbook(master)
    sheet = book["工资核算"]
    sheet["C2"].fill = PatternFill("solid", fgColor="112233")
    sheet["C2"].number_format = "0.00"
    sheet.row_dimensions[2].height = 28
    sheet.print_area = "A1:E3"
    sheet["E3"] = "=SUM($E$2:$E$2)"
    book.defined_names.add(DefinedName("PayTotal", attr_text="'工资核算'!$E$3"))
    book.save(master)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["issues"] == []
    book = load_workbook(output)
    try:
        sheet = book["工资核算"]
        assert sheet["C3"]._style == sheet["C2"]._style
        assert sheet.row_dimensions[3].height == 28
        assert sheet["E4"].value == "=SUM($E$2:$E$3)"
        assert "$E$4" in sheet.print_area
        assert book.defined_names["PayTotal"].attr_text == "'工资核算'!$E$4"
    finally:
        book.close()


def test_missing_hire_id_does_not_create_an_employee_or_report_success(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(source)
    book.active["B3"] = None
    book.save(source)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["auto_update_count"] == 0
    assert result["personnel_coverage"]["input_person_count"] == 1
    assert result["personnel_coverage"]["match_rate"] == 0


def test_numeric_formatted_employee_id_keeps_leading_zeros(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(source)
    book.active["B3"] = 2
    book.active["B3"].number_format = "000"
    book.save(source)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["issues"] == []
    book = load_workbook(output)
    try:
        assert book["员工档案"]["A3"].value == "002"
    finally:
        book.close()


def test_existing_update_master_endpoint_handles_notice_without_payroll_reconstruction(tmp_path: Path, monkeypatch) -> None:
    import backend.routers.pipeline as pipeline

    master, source, _ = hire_files(tmp_path)
    project = SimpleNamespace(id="project", salary_month="2026-05")
    source_file = SimpleNamespace(id="source", stored_path=source.name, original_name="客户入职通知.xlsx")
    master_file = SimpleNamespace(id="master", original_name=master.name)

    class Query:
        def filter(self, *args):
            return self

        def order_by(self, *args):
            return self

        def all(self):
            return [source_file]

    monkeypatch.setattr(pipeline, "UPLOAD_DIR", str(tmp_path))
    monkeypatch.setattr(pipeline, "EXPORT_DIR", str(tmp_path / "exports"))
    monkeypatch.setattr(pipeline, "SESSION_DIR", str(tmp_path / "sessions"))
    (tmp_path / "sessions").mkdir()
    monkeypatch.setattr(pipeline, "_load_project_or_404", lambda *args: project)
    monkeypatch.setattr(pipeline, "_select_template_file", lambda *args: (master_file, str(master)))
    monkeypatch.setattr(pipeline, "_load_template_base_entities", lambda *args: ({}, []))
    monkeypatch.setattr(pipeline, "_current_source_signature", lambda *args: ("input-signature", 2))
    monkeypatch.setattr(pipeline, "export_unified_workbook", lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("Must not rebuild the payroll")))

    result = pipeline.integrate_pipeline(project.id, user=SimpleNamespace(), db=SimpleNamespace(query=lambda *args: Query()))

    assert result.status == "completed"
    assert result.issue_count == 0
    assert result.issues_filename is None
    assert result.validation["scope"] == "personnel_updates"
    assert result.validation["personnel_coverage"]["match_rate"] == 1
    meta = pipeline._load_json(project.id, "export_meta")
    assert meta["source_signature"] == "input-signature"
    assert (tmp_path / "exports" / meta["path"]).is_file()
    assert all(update["source_file"] == source_file.original_name for update in meta["updates"])


@pytest.mark.parametrize("cell,value", [("F3", "未确定"), ("F3", "NaN"), ("F3", True), ("E3", "2026-02-30")])
def test_invalid_hire_input_never_reaches_any_master_sheet(tmp_path: Path, cell: str, value: object) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(source)
    book.active[cell] = value
    book.save(source)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["auto_update_count"] == 0
    assert any(issue["issue_type"] == "hire_invalid_value" for issue in result["issues"])
    assert result["personnel_coverage"]["match_rate"] == 0


def test_text_amounts_and_dates_are_normalized_before_merging(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(source)
    book.active["F3"] = "￥1,200.50"
    book.active["E3"] = "2026/05/12"
    book.save(source)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["issues"] == []
    book = load_workbook(output)
    try:
        assert book["工资核算"]["C3"].value == 1200.50
        assert book["员工档案"]["E3"].value == datetime(2026, 5, 12)
    finally:
        book.close()


@pytest.mark.parametrize("missing_cell", ["A2", "B2"])
def test_existing_partial_identity_is_completed_without_duplicate_employee(tmp_path: Path, missing_cell: str) -> None:
    master, source, output = hire_files(tmp_path)
    book = load_workbook(source)
    book.active["A3"] = "测试甲"
    book.active["B3"] = "001"
    book.save(source)
    book.close()
    book = load_workbook(master)
    book["员工档案"][missing_cell] = None
    book["工资核算"]["A2"].number_format = "000000"
    book.save(master)
    book.close()
    result = apply_semantic_sheet_updates(master, [source], output)
    assert result["issues"] == []
    book = load_workbook(output)
    try:
        assert book["员工档案"].max_row == 2
        assert book["员工档案"]["A2"].value == "001"
        assert book["员工档案"]["B2"].value == "测试甲"
        assert book["工资核算"]["A2"].number_format == "000000"
    finally:
        book.close()


def test_invalid_duplicate_notice_blocks_its_person_in_every_source_order(tmp_path: Path) -> None:
    master, source, output = hire_files(tmp_path)
    second = tmp_path / "invalid-notice.xlsx"
    book = load_workbook(source)
    book.active["F3"] = "尚未确定"
    book.save(second)
    book.close()
    for sources in ([source, second], [second, source]):
        result = apply_semantic_sheet_updates(master, sources, output)
        assert result["auto_update_count"] == 0
        assert result["personnel_coverage"]["input_person_count"] == 1
        assert result["personnel_coverage"]["processed_person_count"] == 0
