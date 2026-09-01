from __future__ import annotations

from pathlib import Path
import json

import openpyxl
import pytest


def test_basic_processor_updates_explicit_payroll_sources_and_returns_a_change_log(
    tmp_path: Path, capsys: pytest.CaptureFixture[str],
) -> None:
    from scripts.keyuan_basic_processor import process_keyuan_basics

    master_path = tmp_path / "master.xlsx"
    master = openpyxl.Workbook()
    payroll = master.active
    payroll.title = "工资核算"
    payroll.cell(4, 2).value = "张三"
    attendance = master.create_sheet("考勤")
    attendance.cell(2, 2).value = "张三"
    duty = master.create_sheet("配送员值班费")
    duty.cell(2, 2).value = "张三"
    adjustment = master.create_sheet("其他调差累计")
    master.save(master_path)
    master.close()

    source_path = tmp_path / "薪资数据-科园-7月薪资.xlsx"
    source = openpyxl.Workbook()
    bonus = source.active
    bonus.title = "奖金-7月"
    bonus.cell(4, 4).value = "张三"
    bonus.cell(4, 10).value = 100
    bonus.cell(4, 11).value = 20
    bonus.cell(4, 12).value = 30
    source_attendance = source.create_sheet("考勤-7月")
    source_attendance.cell(9, 1).value = "张三"
    source_attendance.cell(9, 7).value = 21
    source_attendance.cell(9, 17).value = 2
    source_attendance.cell(9, 30).value = 3
    source_attendance.cell(9, 32).value = 4
    source_attendance.cell(9, 33).value = 5
    source_attendance.cell(9, 34).value = 6
    source_duty = source.create_sheet("值班")
    source_duty.cell(3, 3).value = "张三"
    source_duty.cell(4, 3).value = "张三"
    source_adjustment = source.create_sheet("补发补扣")
    source_adjustment.cell(1, 1).value = "张三"
    source_adjustment.cell(1, 2).value = "加班补发 775.98元"
    source.save(source_path)
    source.close()

    output_path = tmp_path / "processed.xlsx"
    result = process_keyuan_basics(master_path, source_path, output_path)

    assert output_path.is_file()
    assert result["change_count"] >= 10
    assert {change["rule"] for change in result["changes"]} >= {"奖金 J+K/L", "考勤字段", "值班次数", "补发补扣"}
    output = openpyxl.load_workbook(output_path, data_only=False)
    assert output["工资核算"].cell(4, 30).value == 120
    assert output["工资核算"].cell(4, 31).value == 30
    assert output["考勤"].cell(2, 7).value == 21
    assert output["考勤"].cell(2, 9).value == 7
    assert output["考勤"].cell(2, 16).value == 6
    assert output["配送员值班费"].cell(2, 4).value == 2
    assert [output["其他调差累计"].cell(2, column).value for column in range(1, 4)] == ["张三", 775.98, "加班补发 775.98元"]
    output.close()

    from scripts.keyuan_basic_processor import main

    cli_output = tmp_path / "cli-processed.xlsx"
    assert main([str(master_path), str(source_path), str(cli_output)]) == 0
    assert json.loads(capsys.readouterr().out)["change_count"] >= 10


def test_basic_processor_never_overwrites_the_salary_source(tmp_path: Path) -> None:
    from scripts.keyuan_basic_processor import process_keyuan_basics

    master_path = tmp_path / "master.xlsx"
    master = openpyxl.Workbook()
    master.active.title = "工资核算"
    master.create_sheet("考勤")
    master.create_sheet("配送员值班费")
    master.create_sheet("其他调差累计")
    master.save(master_path)
    master.close()

    source_path = tmp_path / "salary.xlsx"
    source = openpyxl.Workbook()
    source.active.title = "奖金-7月"
    for sheet in ("考勤-7月", "值班", "补发补扣"):
        source.create_sheet(sheet)
    source.save(source_path)
    original = source_path.read_bytes()
    source.close()

    with pytest.raises(ValueError, match="独立草稿"):
        process_keyuan_basics(master_path, source_path, source_path)

    assert source_path.read_bytes() == original


def test_basic_processor_accepts_a_different_month_sheet_suffix(tmp_path: Path) -> None:
    from scripts.keyuan_basic_processor import process_keyuan_basics

    master_path = tmp_path / "master.xlsx"
    master = openpyxl.Workbook()
    payroll = master.active
    payroll.title = "工资核算"
    payroll.cell(4, 2).value = "张三"
    attendance = master.create_sheet("考勤")
    attendance.cell(2, 2).value = "张三"
    duty = master.create_sheet("配送员值班费")
    duty.cell(2, 2).value = "张三"
    master.create_sheet("其他调差累计")
    master.save(master_path)
    master.close()

    source_path = tmp_path / "202608薪资更新.xlsx"
    source = openpyxl.Workbook()
    bonus = source.active
    bonus.title = "奖金-8月"
    bonus.cell(4, 4).value = "张三"
    bonus.cell(4, 10).value = 100
    source_attendance = source.create_sheet("考勤-8月")
    source_attendance.cell(9, 1).value = "张三"
    source_attendance.cell(9, 7).value = 22
    source.create_sheet("值班")
    source.create_sheet("补发补扣")
    source.save(source_path)
    source.close()

    result = process_keyuan_basics(master_path, source_path, tmp_path / "processed.xlsx")

    assert result["status"] in {"passed", "needs_review"}
    output = openpyxl.load_workbook(tmp_path / "processed.xlsx", data_only=False)
    assert output["工资核算"].cell(4, 30).value == 100
    assert output["考勤"].cell(2, 7).value == 22
    output.close()
