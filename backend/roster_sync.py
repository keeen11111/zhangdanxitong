"""Deterministic roster-scoped payroll workbook synchronization."""
from __future__ import annotations

import os
import re
import tempfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter


class RosterSyncError(ValueError):
    """Raised when a roster-scoped update cannot be performed safely."""


SOURCE_TO_TARGET_HEADERS = {
    "基本工资": "基本工资",
    "岗位津贴": "岗位津贴",
    "绩效基数": "绩效基数",
    "实际绩效": "实际绩效",
    "绩效调整": "绩效调整",
    "排班绩效": "排班绩效",
    "加班工资": "加班费",
    "翻班津贴": "翻班津贴",
    "餐饮津贴": "餐费补贴",
    "奖金": "奖金",
    "节日福利": "节日福利",
    "高温费": "高温费",
    "税前调整": "税前调整",
    "病假扣款": "病假扣款",
    "事假扣款": "事假扣款",
    "缺勤扣款": "缺勤扣款",
    "税后调整": "税后调整",
}

EARNINGS_HEADERS = (
    "基本工资", "岗位津贴", "企龄津贴", "差旅津贴", "实际绩效", "绩效调整",
    "排班绩效", "加班费", "翻班津贴", "餐费补贴", "奖金", "节日福利", "高温费",
    "税前调整", "最低薪资补偿", "独生子女",
)
DEDUCTION_HEADERS = ("病假扣款", "事假扣款", "旷工扣款", "其他扣款", "缺勤扣款")
FINANCIAL_HEADERS = {
    "社保基数", "公积金基数", "养老企业", "医疗企业", "失业企业", "工伤企业",
    "医疗企业（地方）", "单位社保合计", "残保金企业", "养老个人", "医疗个人",
    "失业个人", "个人社保合计", "基本公积金企业", "基本公积金个人", "保险总费用合计",
    *EARNINGS_HEADERS, *DEDUCTION_HEADERS, "应发工资", "税后调整", "经济补偿金", "应税工资",
    "个调税", "实发工资", "管理费", "代收代付", "合计", "累计收入额", "累计减除费用",
    "累计专项扣除", "累计已扣个调税", "累计应缴个调税",
}
DERIVED_HEADERS = {"保险总费用合计", "应发工资", "应税工资", "实发工资", "代收代付", "合计"}


def _normalize_text(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).strip()


def _find_header_row(sheet: Any, required: set[str], *, limit: int = 20) -> int:
    for row in range(1, min(sheet.max_row, limit) + 1):
        values = {_normalize_text(cell.value) for cell in sheet[row] if cell.value is not None}
        if required.issubset(values):
            return row
    raise RosterSyncError(f"工作表“{sheet.title}”未找到必需表头：{'、'.join(sorted(required))}")


def _header_map(sheet: Any, row: int) -> dict[str, int]:
    result: dict[str, int] = {}
    for cell in sheet[row]:
        header = _normalize_text(cell.value)
        if not header:
            continue
        if header in result:
            raise RosterSyncError(f"工作表“{sheet.title}”存在重复表头：{header}")
        result[header] = cell.column
    return result


def _number(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, bool):
        return float(value)
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise RosterSyncError(f"需要数值的单元格包含非数值内容：{value}") from exc


def _money(value: float) -> float | int:
    rounded = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if rounded == rounded.to_integral():
        return int(rounded)
    return float(rounded)


def _rmb_upper(value: Any) -> str:
    amount = Decimal(str(_number(value))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if amount < 0 or amount >= Decimal("100000000000000"):
        raise RosterSyncError("人民币大写仅支持 0 至 100 万亿元以内的金额")
    digits = "零壹贰叁肆伍陆柒捌玖"
    small_units = ("", "拾", "佰", "仟")
    section_units = ("", "万", "亿", "万亿")

    def section_text(section: int) -> str:
        output = ""
        pending_zero = False
        for position in range(3, -1, -1):
            divisor = 10 ** position
            digit = section // divisor
            section %= divisor
            if digit:
                if pending_zero and output:
                    output += "零"
                output += digits[digit] + small_units[position]
                pending_zero = False
            elif output and section:
                pending_zero = True
        return output

    integer = int(amount)
    if integer == 0:
        integer_text = "零"
    else:
        sections: list[int] = []
        while integer:
            sections.append(integer % 10000)
            integer //= 10000
        pieces: list[str] = []
        zero_between = False
        for index in range(len(sections) - 1, -1, -1):
            section = sections[index]
            if section == 0:
                if pieces:
                    zero_between = True
                continue
            if pieces and (zero_between or section < 1000):
                pieces.append("零")
            pieces.append(section_text(section) + section_units[index])
            zero_between = False
        integer_text = "".join(pieces)

    cents = int((amount * 100) % 100)
    jiao, fen = divmod(cents, 10)
    if not jiao and not fen:
        fraction_text = "整"
    elif jiao and fen:
        fraction_text = f"{digits[jiao]}角{digits[fen]}分"
    elif jiao:
        fraction_text = f"{digits[jiao]}角"
    else:
        fraction_text = f"零{digits[fen]}分"
    return f"{integer_text}圆{fraction_text}"


def _set_value(sheet: Any, row: int, column: int, value: Any, changes: list[dict[str, Any]]) -> None:
    cell = sheet.cell(row, column)
    before = cell.value
    if before == value:
        return
    cell.value = value
    changes.append({
        "sheet": sheet.title,
        "cell": f"{get_column_letter(column)}{row}",
        "before": before,
        "after": value,
    })


def _reject_formula_cells(sheet: Any, coordinates: set[tuple[int, int]]) -> None:
    for row, column in sorted(coordinates):
        cell = sheet.cell(row, column)
        if cell.data_type == "f":
            raise RosterSyncError(f"工作表“{sheet.title}”的 {cell.coordinate} 是公式，已拒绝覆盖")


def _set_period_label(sheet: Any, label: str, period: str, changes: list[dict[str, Any]]) -> None:
    normalized_label = _normalize_text(label)
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 10)):
        for cell in row:
            if _normalize_text(cell.value) == normalized_label:
                _set_value(sheet, cell.row, cell.column + 1, period, changes)
                return


def _update_derived_values(sheet: Any, row: int, headers: dict[str, int], changes: list[dict[str, Any]]) -> None:
    def value(header: str) -> float:
        column = headers.get(header)
        return _number(sheet.cell(row, column).value) if column else 0.0

    def update(header: str, amount: float) -> None:
        column = headers.get(header)
        if column:
            _set_value(sheet, row, column, _money(amount), changes)

    unit_social = value("单位社保合计")
    personal_social = value("个人社保合计")
    residual = value("残保金企业")
    company_fund = value("基本公积金企业")
    personal_fund = value("基本公积金个人")
    gross = sum(value(header) for header in EARNINGS_HEADERS) - sum(value(header) for header in DEDUCTION_HEADERS)
    taxable = gross - personal_social - personal_fund
    net = taxable - value("个调税") + value("税后调整") + value("经济补偿金")
    pass_through = gross + unit_social + residual + company_fund
    total = pass_through + value("管理费")
    update("保险总费用合计", unit_social + personal_social + residual + company_fund + personal_fund)
    update("应发工资", gross)
    update("应税工资", taxable)
    update("实发工资", net)
    update("代收代付", pass_through)
    update("合计", total)


def _update_total_row(sheet: Any, header_row: int, total_row: int, headers: dict[str, int], changes: list[dict[str, Any]]) -> None:
    first_employee_row = header_row + 1
    _set_value(sheet, total_row, 1, "*合计*", changes)
    for marker_header in ("姓名", "证件类型", "状态"):
        column = headers.get(marker_header)
        if column:
            _set_value(sheet, total_row, column, "79", changes)
    for header, column in headers.items():
        if header not in FINANCIAL_HEADERS:
            continue
        total = sum(_number(sheet.cell(row, column).value) for row in range(first_employee_row, total_row))
        _set_value(sheet, total_row, column, _money(total), changes)


def _update_notice(
    notice: Any,
    period: str,
    employee_count: int,
    detail: Any,
    detail_headers: dict[str, int],
    total_row: int,
    changes: list[dict[str, Any]],
) -> None:
    _set_period_label(notice, "费用所属期间：", period, changes)

    def total(header: str) -> float:
        column = detail_headers.get(header)
        return _number(detail.cell(total_row, column).value) if column else 0.0

    special_counts = {"社保小计", "残保金企业", "公积金小计"}
    for row in range(1, notice.max_row + 1):
        raw_label = _normalize_text(notice.cell(row, 1).value)
        label = raw_label.rstrip("：:")
        if label == "管理费":
            _set_value(notice, row, 5, _money(total("管理费")), changes)
        elif label == "本期总计应付款":
            _set_value(notice, row, 5, _money(total("合计")), changes)
        elif label == "人民币大写":
            _set_value(notice, row, 5, _rmb_upper(total("合计")), changes)
        elif label in detail_headers:
            _set_value(notice, row, 5, "79" if label in special_counts else employee_count, changes)
            _set_value(notice, row, 6, _money(total(label)), changes)
        elif label == "社保小计":
            _set_value(notice, row, 5, "79", changes)
            _set_value(notice, row, 6, _money(total("单位社保合计") + total("个人社保合计")), changes)
        elif label == "公积金小计":
            _set_value(notice, row, 5, "79", changes)
            _set_value(notice, row, 6, _money(total("基本公积金企业") + total("基本公积金个人")), changes)


def sync_master_to_source_roster(
    *,
    master_path: Path,
    source_path: Path,
    output_path: Path,
    source_sheet: str,
    target_sheet: str,
    target_period: str,
) -> dict[str, Any]:
    """Synchronize one target payroll table to the exact roster of one source sheet.

    Existing target rows are retained in their original order. A source person absent
    from the target is rejected because identity, bank and social-insurance fields
    cannot be safely fabricated from the salary source alone.
    """
    period = str(target_period or "").strip().replace(".", "-").replace("/", "-")
    period_match = re.fullmatch(r"(\d{4})-(\d{1,2})", period)
    if period_match is None or not 1 <= int(period_match.group(2)) <= 12:
        raise RosterSyncError("目标月份必须使用 YYYY-MM 格式")
    period = f"{period_match.group(1)}-{int(period_match.group(2)):02d}"
    master_path = Path(master_path)
    source_path = Path(source_path)
    output_path = Path(output_path)
    if not master_path.is_file():
        raise RosterSyncError("总表文件不存在")
    if not source_path.is_file():
        raise RosterSyncError("来源文件不存在")

    master = openpyxl.load_workbook(master_path, data_only=False)
    source = openpyxl.load_workbook(source_path, read_only=True, data_only=True)
    changes: list[dict[str, Any]] = []
    try:
        if source_sheet not in source.sheetnames:
            raise RosterSyncError(f"来源文件不存在工作表“{source_sheet}”")
        if target_sheet not in master.sheetnames:
            raise RosterSyncError(f"总表不存在工作表“{target_sheet}”")
        source_ws = source[source_sheet]
        target_ws = master[target_sheet]
        source_header_row = _find_header_row(source_ws, {"员工姓名"})
        target_header_row = _find_header_row(target_ws, {"序号", "姓名"})
        source_headers = _header_map(source_ws, source_header_row)
        target_headers = _header_map(target_ws, target_header_row)
        source_name_column = source_headers["员工姓名"]
        target_name_column = target_headers["姓名"]

        source_records: dict[str, dict[str, Any]] = {}
        source_display_names: dict[str, str] = {}
        for row in range(source_header_row + 1, source_ws.max_row + 1):
            display_name = str(source_ws.cell(row, source_name_column).value or "").strip()
            normalized_name = _normalize_text(display_name)
            if not normalized_name:
                continue
            if normalized_name in source_records:
                raise RosterSyncError(f"来源工作表“{source_sheet}”存在重复姓名：{display_name}")
            source_records[normalized_name] = {
                header: source_ws.cell(row, column).value for header, column in source_headers.items()
            }
            source_display_names[normalized_name] = display_name
        if not source_records:
            raise RosterSyncError(f"来源工作表“{source_sheet}”没有人员数据")

        total_row = next(
            (row for row in range(target_header_row + 1, target_ws.max_row + 1)
             if _normalize_text(target_ws.cell(row, 1).value) == "*合计*"),
            None,
        )
        if total_row is None:
            raise RosterSyncError(f"总表工作表“{target_sheet}”未找到“*合计*”行")
        target_rows: dict[str, int] = {}
        target_display_names: dict[str, str] = {}
        for row in range(target_header_row + 1, total_row):
            display_name = str(target_ws.cell(row, target_name_column).value or "").strip()
            normalized_name = _normalize_text(display_name)
            if not normalized_name:
                continue
            if normalized_name in target_rows:
                raise RosterSyncError(f"总表工作表“{target_sheet}”存在重复姓名：{display_name}")
            target_rows[normalized_name] = row
            target_display_names[normalized_name] = display_name
        missing = [source_display_names[name] for name in source_records if name not in target_rows]
        if missing:
            raise RosterSyncError(f"来源人员在总表中不存在，无法安全补齐身份和银行字段：{'、'.join(missing)}")

        employee_write_headers = {
            "序号", "薪资所属月", "社保所属月", "公积金所属月",
            *SOURCE_TO_TARGET_HEADERS.values(), *DERIVED_HEADERS,
        }
        target_formula_coordinates = {
            (row, target_headers[header])
            for name, row in target_rows.items()
            if name in source_records
            for header in employee_write_headers
            if header in target_headers
        }
        target_formula_coordinates.update(
            (total_row, column)
            for header, column in target_headers.items()
            if header in FINANCIAL_HEADERS
        )
        _reject_formula_cells(target_ws, target_formula_coordinates)
        if "付款通知书" in master.sheetnames:
            notice = master["付款通知书"]
            notice_formula_coordinates: set[tuple[int, int]] = set()
            for row in range(1, notice.max_row + 1):
                label = _normalize_text(notice.cell(row, 1).value).rstrip("：:")
                if label in target_headers or label in {"社保小计", "公积金小计"}:
                    notice_formula_coordinates.update({(row, 5), (row, 6)})
                elif label in {"管理费", "本期总计应付款", "人民币大写"}:
                    notice_formula_coordinates.add((row, 5))
            _reject_formula_cells(notice, notice_formula_coordinates)

        removed_names = [target_display_names[name] for name in target_rows if name not in source_records]
        rows_to_delete = sorted((target_rows[name] for name in target_rows if name not in source_records), reverse=True)
        for row in rows_to_delete:
            target_ws.delete_rows(row, 1)
        if rows_to_delete:
            changes.append({
                "sheet": target_sheet,
                "operation": "delete_non_roster_rows",
                "deleted_rows": sorted(rows_to_delete),
                "removed_names": removed_names,
            })

        total_row -= len(rows_to_delete)
        retained_names: list[str] = []
        for sequence, row in enumerate(range(target_header_row + 1, total_row), start=1):
            display_name = str(target_ws.cell(row, target_name_column).value or "").strip()
            normalized_name = _normalize_text(display_name)
            if not normalized_name:
                continue
            retained_names.append(display_name)
            record = source_records[normalized_name]
            sequence_column = target_headers.get("序号")
            if sequence_column:
                _set_value(target_ws, row, sequence_column, sequence, changes)
            for period_header in ("薪资所属月", "社保所属月", "公积金所属月"):
                column = target_headers.get(period_header)
                if column:
                    _set_value(target_ws, row, column, period, changes)
            for source_header, target_header in SOURCE_TO_TARGET_HEADERS.items():
                source_column = source_headers.get(source_header)
                target_column = target_headers.get(target_header)
                if source_column is None or target_column is None:
                    continue
                source_value = record.get(source_header)
                _set_value(target_ws, row, target_column, 0 if source_value in (None, "") else source_value, changes)
            _update_derived_values(target_ws, row, target_headers, changes)

        _set_period_label(target_ws, "结算月：", period, changes)
        _set_period_label(target_ws, "所属月：", period, changes)
        _update_total_row(target_ws, target_header_row, total_row, target_headers, changes)
        if "付款通知书" in master.sheetnames:
            _update_notice(
                master["付款通知书"], period, len(retained_names), target_ws,
                target_headers, total_row, changes,
            )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_name: str | None = None
        try:
            with tempfile.NamedTemporaryFile(
                prefix=f".{output_path.stem}-", suffix=output_path.suffix or ".xlsx",
                dir=output_path.parent, delete=False,
            ) as temporary:
                temporary_name = temporary.name
            master.save(temporary_name)
            os.replace(temporary_name, output_path)
            temporary_name = None
        finally:
            if temporary_name:
                Path(temporary_name).unlink(missing_ok=True)

        unmapped_source_headers = [
            header for header in source_headers
            if header not in SOURCE_TO_TARGET_HEADERS and header not in {"员工姓名"}
        ]
        return {
            "status": "passed",
            "source_sheet": source_sheet,
            "target_sheet": target_sheet,
            "target_period": period,
            "employee_count": len(retained_names),
            "retained_names": retained_names,
            "removed_names": removed_names,
            "unmapped_source_headers": unmapped_source_headers,
            "change_count": len(changes),
            "changes": changes,
        }
    finally:
        source.close()
        master.close()
