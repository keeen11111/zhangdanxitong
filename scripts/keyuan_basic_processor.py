"""Deterministic, auditable basics for the 科园 payroll workbook.

This module deliberately covers only source-to-target operations that can be
verified without a model: bonus values, attendance fields, duty counts and
explicit adjustment amounts.  It always writes a separate output workbook;
the input master and source workbook are opened read-only.
"""
from __future__ import annotations

import re
import shutil
import json
import sys
from pathlib import Path
from typing import Any

import openpyxl


def _text(value: Any) -> str:
    return str(value or "").strip()


def _number(value: Any) -> float:
    if value is None or isinstance(value, bool):
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _write(
    changes: list[dict[str, Any]],
    sheet: str,
    cell: openpyxl.cell.cell.Cell,
    value: Any,
    *,
    source: str,
    rule: str,
    issues: list[dict[str, str]] | None = None,
) -> None:
    """Write one non-formula target and retain a physical-change audit row."""
    if cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("=")):
        if issues is not None:
            issues.append({"item": f"{sheet}!{cell.coordinate}", "detail": "目标单元格是公式，基础处理器跳过写入"})
            return
        raise ValueError(f"{sheet}!{cell.coordinate} 是公式，基础处理器不能覆盖")
    if cell.value == value:
        return
    before = cell.value
    cell.value = value
    changes.append({
        "sheet": sheet,
        "cell": cell.coordinate,
        "before": before,
        "after": value,
        "source": source,
        "rule": rule,
    })


def _rows_by_name(sheet: openpyxl.worksheet.worksheet.Worksheet, column: int, start_row: int) -> dict[str, int]:
    return {
        name: row
        for row in range(start_row, sheet.max_row + 1)
        if (name := _text(sheet.cell(row, column).value))
    }


def _sheet_with_prefix(workbook: openpyxl.Workbook, prefixes: tuple[str, ...]) -> openpyxl.worksheet.worksheet.Worksheet | None:
    """Find a monthly source sheet without hard-coding a particular month."""
    for sheet_name in workbook.sheetnames:
        if any(sheet_name == prefix or sheet_name.startswith(prefix) for prefix in prefixes):
            return workbook[sheet_name]
    return None


def process_keyuan_basics(master_path: Path, salary_source_path: Path, output_path: Path) -> dict[str, Any]:
    """Create an output draft with the safely deterministic 科园 source updates.

    The return value is a portable audit payload suitable for a future Agent
    tool result.  Unsupported or unparseable records remain in ``issues``;
    they are never converted into guessed values.
    """
    master_path = Path(master_path)
    salary_source_path = Path(salary_source_path)
    output_path = Path(output_path)
    if not master_path.is_file() or not salary_source_path.is_file():
        raise FileNotFoundError("总表或薪资来源文件不存在")
    if output_path.resolve() in {master_path.resolve(), salary_source_path.resolve()}:
        raise ValueError("基础处理器只能写入独立草稿，不能覆盖原始总表或来源文件")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not output_path.exists():
        shutil.copy2(master_path, output_path)
    master = openpyxl.load_workbook(output_path, data_only=False)
    source = openpyxl.load_workbook(salary_source_path, data_only=True, read_only=True)
    changes: list[dict[str, Any]] = []
    issues: list[dict[str, str]] = []
    source_name = salary_source_path.name
    try:
        payroll = master["工资核算"]
        attendance = master["考勤"]
        duty = master["配送员值班费"]
        adjustment = master["其他调差累计"]
        bonus = _sheet_with_prefix(source, ("奖金-", "奖金"))
        source_attendance = _sheet_with_prefix(source, ("考勤-", "考勤"))
        source_duty = _sheet_with_prefix(source, ("值班",))
        source_adjustment = _sheet_with_prefix(source, ("补发补扣", "调差"))
        for label, sheet in (("奖金", bonus), ("考勤", source_attendance), ("值班", source_duty), ("补发补扣", source_adjustment)):
            if sheet is None:
                issues.append({"item": label, "detail": f"来源文件未找到{label} Sheet，本项交由 Agent 继续核对"})

        bonus_values = {
            _text(bonus.cell(row, 4).value): (
                _number(bonus.cell(row, 10).value) + _number(bonus.cell(row, 11).value),
                _number(bonus.cell(row, 12).value),
            )
            for row in range(4, bonus.max_row + 1)
            if _text(bonus.cell(row, 4).value)
        } if bonus is not None else {}
        for name, row in _rows_by_name(payroll, 2, 4).items():
            if name not in bonus_values:
                continue
            first, second = bonus_values[name]
            _write(changes, "工资核算", payroll.cell(row, 30), first, source=source_name, rule="奖金 J+K/L", issues=issues)
            _write(changes, "工资核算", payroll.cell(row, 31), second, source=source_name, rule="奖金 J+K/L", issues=issues)

        attendance_values = {
            _text(source_attendance.cell(row, 1).value): (
                _number(source_attendance.cell(row, 7).value),
                _number(source_attendance.cell(row, 17).value),
                _number(source_attendance.cell(row, 30).value) + _number(source_attendance.cell(row, 32).value),
                _number(source_attendance.cell(row, 33).value),
                _number(source_attendance.cell(row, 34).value),
            )
            for row in range(9, source_attendance.max_row + 1)
            if _text(source_attendance.cell(row, 1).value)
        } if source_attendance is not None else {}
        for name, row in _rows_by_name(attendance, 2, 2).items():
            if name not in attendance_values:
                continue
            values = attendance_values[name]
            for column, value in zip((7, 13, 9, 10, 16), values, strict=True):
                _write(changes, "考勤", attendance.cell(row, column), value, source=source_name, rule="考勤字段", issues=issues)

        duty_counts: dict[str, int] = {}
        if source_duty is not None:
            for row in range(3, source_duty.max_row + 1):
                name = _text(source_duty.cell(row, 3).value)
                if name:
                    duty_counts[name] = duty_counts.get(name, 0) + 1
        for name, row in _rows_by_name(duty, 2, 2).items():
            _write(changes, "配送员值班费", duty.cell(row, 4), duty_counts.get(name, 0), source=source_name, rule="值班次数", issues=issues)

        if source_adjustment is not None:
            for row in range(2, adjustment.max_row + 1):
                for column in range(1, 4):
                    _write(changes, "其他调差累计", adjustment.cell(row, column), None, source="本次处理", rule="清空当月调差区", issues=issues)
        output_row = 2
        amount_pattern = re.compile(r"(?<!\d)(\d+(?:\.\d{1,2})?)(?=元)")
        if source_adjustment is not None:
            for row in range(1, source_adjustment.max_row + 1):
                name = _text(source_adjustment.cell(row, 1).value)
                note = _text(source_adjustment.cell(row, 2).value)
                if not name or not note:
                    continue
                matches = amount_pattern.findall(note)
                if not matches:
                    issues.append({"item": "补发补扣", "detail": f"{name} 的金额无法从说明中识别"})
                    continue
                amount = float(matches[-1])
                _write(changes, "其他调差累计", adjustment.cell(output_row, 1), name, source=source_name, rule="补发补扣", issues=issues)
                _write(changes, "其他调差累计", adjustment.cell(output_row, 2), amount, source=source_name, rule="补发补扣", issues=issues)
                _write(changes, "其他调差累计", adjustment.cell(output_row, 3), note, source=source_name, rule="补发补扣", issues=issues)
                output_row += 1

        master.save(output_path)
    finally:
        source.close()
        master.close()

    return {
        "status": "passed" if not issues else "needs_review",
        "output_path": str(output_path),
        "change_count": len(changes),
        "changes": changes,
        "issues": issues,
    }


def main(arguments: list[str] | None = None) -> int:
    """Run the processor from PowerShell and print only its audit payload."""
    values = list(arguments if arguments is not None else sys.argv[1:])
    if len(values) != 3:
        raise SystemExit("usage: keyuan_basic_processor.py <master.xlsx> <salary-source.xlsx> <output.xlsx>")
    result = process_keyuan_basics(Path(values[0]), Path(values[1]), Path(values[2]))
    print(json.dumps(result, ensure_ascii=False, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
