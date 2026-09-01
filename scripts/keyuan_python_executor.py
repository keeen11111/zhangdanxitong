"""Pure-Python executor for the verified 科园 2026.07 payroll batch.

The executor is intentionally boring: it copies the master, applies only the
documented cell/range operations, records every physical change, and validates
the resulting workbook before returning it to the Agent.
"""
from __future__ import annotations

from copy import copy
from datetime import datetime
from io import BytesIO
import json
from pathlib import Path
import re
import shutil
from typing import Any, Mapping

import openpyxl
import pandas as pd
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula


def norm(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).replace("（", "(").replace("）", ")")


def number(value: Any) -> float:
    if value is None or isinstance(value, bool) or str(value).strip() == "":
        return 0.0
    try:
        return float(str(value).replace(",", "").replace("¥", ""))
    except (TypeError, ValueError):
        return 0.0


def parsed_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool) or str(value).strip() == "":
        return None
    try:
        return float(str(value).replace(",", "").replace("¥", ""))
    except (TypeError, ValueError):
        return None


def copy_cell(source: Any, target: Any) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)


class KeyuanPythonError(RuntimeError):
    pass


def _set(changes: list[dict[str, Any]], sheet: Any, row: int, col: int, value: Any, source: str, rule: str) -> None:
    cell = sheet.cell(row, col)
    before = cell.value
    if before == value:
        return
    cell.value = value
    changes.append({"sheet": sheet.title, "cell": cell.coordinate, "before": before, "after": value, "source": source, "rule": rule})


def _copy_row(
    sheet: Any,
    source_row: int,
    target_row: int,
    max_col: int,
    *,
    value_sheet: Any | None = None,
) -> None:
    for col in range(1, max_col + 1):
        source = sheet.cell(source_row, col)
        target = sheet.cell(target_row, col)
        target.value = value_sheet.cell(source_row, col).value if value_sheet is not None else source.value
        copy_cell(source, target)
    if source_row in sheet.row_dimensions:
        sheet.row_dimensions[target_row] = copy(sheet.row_dimensions[source_row])


def _read_book(path: Path) -> openpyxl.Workbook:
    try:
        return openpyxl.load_workbook(path, data_only=False, keep_links=False)
    except (OSError, ValueError, KeyError) as exc:
        raise KeyuanPythonError(f"无法读取工作簿：{path.name}") from exc


def _tax_frame(path: Path) -> pd.DataFrame:
    content = path.read_bytes()
    try:
        if content.startswith(b"PK"):
            return pd.read_excel(BytesIO(content), dtype=object, engine="openpyxl")
        return pd.read_excel(path, dtype=object)
    except (OSError, ValueError, ImportError) as exc:
        raise KeyuanPythonError(f"无法读取个税文件：{path.name}") from exc


def _formula_map(row: int) -> dict[int, str]:
    return {
        1: f"=VLOOKUP(C{row},考勤!B:Q,16,0)",
        27: f"=ROUND(Y{row}*Z{row},-1)",
        28: f"=Y{row}-AA{row}",
        35: f"=SUMIF(其他调差累计!A:A,C{row},其他调差累计!B:B)",
        36: f"=SUMIF(考勤!A:A,工资核算!B{row},考勤!G:G)",
        37: f"=SUMIF(考勤!$A:$A,工资核算!$B{row},考勤!I:I)",
        38: f"=SUMIF(考勤!$A:$A,工资核算!$B{row},考勤!J:J)",
        39: f"=SUMIF(考勤!$A:$A,工资核算!$B{row},考勤!K:K)+SUMIF(考勤!$B:$B,C{row},考勤!$L:$L)",
        40: f"=SUMIF(考勤!$A:$A,工资核算!$B{row},考勤!M:M)",
        41: f"=SUMIF(考勤!$A:$A,工资核算!$B{row},考勤!N:N)",
        42: f"=SUMIF(考勤!$A:$A,工资核算!$B{row},考勤!O:O)",
        43: f"=AM{row}*30",
        44: f"=ROUND(ROUND(Y{row}/21.75,2)*30%*AN{row},2)",
        45: f"=ROUND(ROUND(Y{row}/21.75,2)*AO{row},2)",
        46: f"=ROUND(ROUND(Y{row}/21.75/8,2)*30%*AP{row},2)",
        47: f"=ROUND(ROUND(AA{row}/21.75/8,2)*(AK{row}*150%+AL{row}*300%),2)",
        48: f"=ROUND(Q{row}/23*AJ{row},0)",
        49: f'=IF(U{row}="50元/天",50*AJ{row},IF(U{row}="30元/天",30*AJ{row},IF(U{row}="15元/天",15*AJ{row},0)))',
        50: f"=VLOOKUP(C{row},防暑降温费!$B:$H,6,0)",
        53: f"=IFERROR(VLOOKUP(C{row},考勤!B:P,15,0),0)*150",
        54: f"=ROUND(V{row}+W{row}+X{row}+AA{row}+AE{row}+AF{row}+AG{row}+AH{row}+AI{row}-AQ{row}-AR{row}-AS{row}-AT{row}+AU{row}+AV{row}+AW{row}+AX{row}+AY{row}+AZ{row}+BA{row},2)",
        55: f"=VLOOKUP($B{row},台账!$A:M,3,FALSE)",
        56: f"=ROUND(BE{row}+BF{row}+BG{row}+BH{row},2)",
        57: f"=VLOOKUP($B{row},台账!$A:N,5,FALSE)",
        58: f"=VLOOKUP($B{row},台账!$A:O,11,FALSE)",
        59: f"=VLOOKUP($B{row},台账!$A:P,13,FALSE)",
        60: f"=VLOOKUP($B{row},台账!$A:P,7,FALSE)",
        61: f"=ROUND(BJ{row}+BK{row}+BL{row}+BM{row}+BN{row}+BO{row}+BP{row},2)",
        62: f"=VLOOKUP($B{row},台账!$A:Q,4,FALSE)",
        63: f"=VLOOKUP($B{row},台账!$A:Q,10,FALSE)",
        64: f"=VLOOKUP($B{row},台账!$A:Q,12,FALSE)",
        65: f"=VLOOKUP($B{row},台账!$A:R,6,FALSE)",
        66: f"=VLOOKUP($B{row},台账!$A:S,9,FALSE)",
        67: f"=VLOOKUP($B{row},台账!$A:T,8,FALSE)",
        68: f"=VLOOKUP($B{row},台账!$A:U,14,FALSE)",
        69: f"=ROUND(BB{row}-BD{row}-AZ{row},2)",
        70: f"=VLOOKUP(C{row},个税导出核对!$B$2:$AL$48,37,0)",
        72: f"=VLOOKUP($B{row},台账!$A:AK,15,FALSE)+VLOOKUP($B{row},台账!$A:AK,16,FALSE)+VLOOKUP($B{row},台账!$A:AK,17,FALSE)",
        73: f"=ROUND(BB{row}-BD{row}-BR{row}-BT{row}+BS{row},2)",
        74: f"=IF(BU{row}<0,-BU{row},0)",
        76: f"=ROUND(BU{row}+BV{row}+BW{row},2)",
        80: f"=CC{row}-BU{row}",
        81: f"=VLOOKUP($C{row},个税导出核对!$B:$AA,18,0)",
        82: f"=VLOOKUP($C{row},个税导出核对!$B:$AA,22,0)",
        83: f"=VLOOKUP($C{row},个税导出核对!$B:$AA,23,0)",
        84: f"=VLOOKUP($C{row},个税导出核对!$B:$AA,24,0)",
        85: f"=VLOOKUP($C{row},个税导出核对!$B:$AA,25,0)",
        86: f"=VLOOKUP($C{row},个税导出核对!$B:$AA,26,0)",
        87: f"=VLOOKUP($C{row},个税导出核对!$B$2:$AB$48,27,0)",
        88: f"=SUM(CD{row}:CI{row})",
        89: f"=VLOOKUP(C{row},个税导出核对!$B$2:$U$48,20,0)",
        90: f"=ROUND(IF(CC{row}-CK{row}-CJ{row}<0,0,CC{row}-CK{row}-CJ{row}),2)",
        92: f"=SUMIF(个税导出核对!D:D,BY{row},个税导出核对!AK:AK)",
        93: f"=ROUND(IF(CM{row}-CN{row}<0,0,CM{row}-CN{row}),2)",
    }


def process_keyuan_python(
    master_path: Path,
    salary_path: Path,
    social_paths: Mapping[str, Path],
    tax_paths: Mapping[str, Path],
    output_path: Path,
) -> dict[str, Any]:
    if output_path.resolve() in {Path(master_path).resolve(), Path(salary_path).resolve(), *(Path(p).resolve() for p in social_paths.values()), *(Path(p).resolve() for p in tax_paths.values())}:
        raise KeyuanPythonError("科园执行器只能写入独立结果草稿")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(master_path, output_path)
    book = _read_book(output_path)
    cached_master = openpyxl.load_workbook(master_path, data_only=True, keep_links=False)
    changes: list[dict[str, Any]] = []
    issues: list[dict[str, str]] = []
    salary = openpyxl.load_workbook(salary_path, data_only=True, keep_links=False)
    salary_formulas = openpyxl.load_workbook(salary_path, data_only=False, keep_links=False)
    salary_name = Path(salary_path).name
    try:
        payroll = book["工资核算"]
        payroll.insert_cols(1)
        _set(changes, payroll, 2, 1, "班制", "操作手册", "新增班制列")
        for row in range(4, 44):
            for col, formula in _formula_map(row).items():
                existing = payroll.cell(row, col).value
                if existing not in (None, "") and not (isinstance(existing, str) and existing.startswith("=")):
                    continue
                _set(changes, payroll, row, col, formula, "工资核算模板", "插列后恢复公式")
            if norm(payroll.cell(row, 3).value) == "李楠":
                _set(changes, payroll, row, 48, f"=IF(ROUND(Q{row}/15*AJ{row},0)>400,400,ROUND(Q{row}/15*AJ{row},0))", "操作手册", "李楠特殊AU规则")
                _set(changes, payroll, row, 49, f"=IF(ROUND(500/15*AJ{row},0)>500,500,ROUND(500/15*AJ{row},0))", "操作手册", "李楠特殊AV规则")
            if norm(payroll.cell(row, 3).value) in {"李晓苛", "袁艳"}:
                _set(changes, payroll, row, 49, f"=ROUND(400/23*AJ{row},0)", "操作手册", "特殊AV规则")
            if row >= 16:
                _set(changes, payroll, row, 80, f"=CC{row}-BB{row}", "工资核算模板", "按人员区段恢复累计税额公式")
            for col in range(1, payroll.max_column + 1):
                cell = payroll.cell(row, col)
                if isinstance(cell.value, str) and "[" in cell.value:
                    _set(changes, payroll, row, col, None, "操作手册", "清除未提供的外部公式")
            _set(changes, payroll, row, 95, None, "操作手册", "清除未提供的外部公式")
            _set(changes, payroll, row, 96, None, "操作手册", "清除未提供的外部公式")

        for col in range(2, payroll.max_column + 1):
            cell = payroll.cell(44, col)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                translated = Translator(
                    cell.value,
                    origin=f"{get_column_letter(col - 1)}44",
                ).translate_formula(cell.coordinate)
                _set(changes, payroll, 44, col, translated, "工资核算模板", "插列后恢复合计公式")
        _set(changes, payroll, 49, 70, "=BB44+BJ44+BK44+BL44+BM44+BO44", "工资核算模板", "恢复工资核算汇总公式")
        for col in (54, 62, 63, 64, 65, 66, 67):
            letter = get_column_letter(col)
            _set(changes, payroll, 50, col, f"=SUBTOTAL(9,{letter}4:{letter}43)", "工资核算模板", "恢复工资核算分项合计")
        _set(changes, payroll, 51, 67, "=BO50+BM50+BL50+BK50+BJ50+BB50", "工资核算模板", "恢复工资核算合计")

        for row in range(4, 44):
            cell = payroll.cell(row, 91)
            if isinstance(cell.value, ArrayFormula):
                translated = Translator(cell.value.text, origin=cell.value.ref).translate_formula(cell.coordinate)
                before = (cell.value.text, cell.value.ref)
                cell.value = ArrayFormula(ref=cell.coordinate, text=translated)
                changes.append({"sheet": payroll.title, "cell": cell.coordinate, "before": before, "after": (translated, cell.coordinate), "source": "工资核算模板", "rule": "插列后平移数组公式"})

        bonus = next((salary[name] for name in salary.sheetnames if name.startswith("奖金-")), None)
        if bonus:
            for row in range(4, 44):
                name = norm(payroll.cell(row, 3).value)
                matches = [source_row for source_row in range(4, bonus.max_row + 1) if norm(bonus.cell(source_row, 4).value) == name]
                if len(matches) > 1:
                    issues.append({"item": f"工资核算!C{row}", "detail": f"奖金来源姓名重复：{name}"})
                    continue
                if matches:
                    source_row = matches[0]
                    j_value = bonus.cell(source_row, 10).value
                    k_value = bonus.cell(source_row, 11).value
                    l_value = bonus.cell(source_row, 12).value
                    formula_bonus = salary_formulas[bonus.title]
                    for source_col, cached_value, label in ((10, j_value, "J"), (11, k_value, "K"), (12, l_value, "L")):
                        raw_value = formula_bonus.cell(source_row, source_col).value
                        if isinstance(raw_value, str) and raw_value.startswith("=") and cached_value is None:
                            issues.append({"item": f"奖金来源!{formula_bonus.cell(source_row, source_col).coordinate}", "detail": f"{name} 的奖金{label}公式没有缓存值"})
                    j_number = parsed_number(j_value)
                    k_number = parsed_number(k_value)
                    l_number = parsed_number(l_value)
                    if (j_value not in (None, "") and j_number is None) or (k_value not in (None, "") and k_number is None) or (l_value not in (None, "") and l_number is None):
                        issues.append({"item": f"奖金来源!D{source_row}", "detail": f"{name} 的奖金金额不是可确认的数值"})
                        continue
                    if j_number is not None or k_number is not None:
                        _set(changes, payroll, row, 31, (j_number or 0) + (k_number or 0), salary_name, "奖金J+K")
                    if l_number is not None:
                        _set(changes, payroll, row, 32, l_number, salary_name, "奖金L")

        attendance_source = next((salary[name] for name in salary.sheetnames if name.startswith("考勤-")), None)
        attendance = book["考勤"]
        if attendance_source:
            for row in range(2, attendance.max_row + 1):
                name = norm(attendance.cell(row, 2).value)
                if not name:
                    continue
                source_rows = [source_row for source_row in range(9, attendance_source.max_row + 1) if norm(attendance_source.cell(source_row, 1).value) == name]
                if len(source_rows) > 1:
                    issues.append({"item": f"考勤!B{row}", "detail": f"考勤来源姓名重复：{name}"})
                    continue
                if source_rows:
                    source_row = source_rows[0]
                    values = {7: attendance_source.cell(source_row, 7).value, 13: attendance_source.cell(source_row, 17).value, 9: number(attendance_source.cell(source_row, 30).value) + number(attendance_source.cell(source_row, 32).value), 10: attendance_source.cell(source_row, 33).value, 16: attendance_source.cell(source_row, 34).value}
                    for col, value in values.items():
                        _set(changes, attendance, row, col, value, salary_name, "考勤G/Q/AD+AF/AG/AH")

        duty_source = salary["值班"] if "值班" in salary.sheetnames else None
        duty = book["配送员值班费"]
        counts: dict[str, int] = {}
        if duty_source:
            for row in range(3, duty_source.max_row + 1):
                name = norm(duty_source.cell(row, 3).value)
                if name:
                    counts[name] = counts.get(name, 0) + 1
        for row in range(2, duty.max_row + 1):
            _set(changes, duty, row, 4, counts.get(norm(duty.cell(row, 2).value), 0), salary_name, "值班次数")

        adjustment = book["其他调差累计"]
        for row in range(2, adjustment.max_row + 1):
            for col in range(1, 4):
                if row == adjustment.max_row and col == 1 and adjustment.cell(row, col).value == "合计":
                    continue
                if not (isinstance(adjustment.cell(row, col).value, str) and adjustment.cell(row, col).value.startswith("=")):
                    _set(changes, adjustment, row, col, None, "操作手册", "清空当月调差区")
        adjustment_source = salary["补发补扣"] if "补发补扣" in salary.sheetnames else None
        output_row = 2
        amount_pattern = re.compile(r"(?<!\d)(\d+(?:\.\d{1,2})?)(?=元)")
        if adjustment_source:
            for row in range(1, adjustment_source.max_row + 1):
                name, note = norm(adjustment_source.cell(row, 1).value), str(adjustment_source.cell(row, 2).value or "").strip()
                if not name or not note:
                    continue
                matches = amount_pattern.findall(note)
                if not matches:
                    issues.append({"item": "补发补扣", "detail": f"{name} 的金额无法从说明中识别"})
                    continue
                _set(changes, adjustment, output_row, 1, name, salary_name, "补发补扣")
                _set(changes, adjustment, output_row, 2, float(matches[-1]), salary_name, "补发补扣")
                _set(changes, adjustment, output_row, 3, note, salary_name, "补发补扣")
                output_row += 1

        heat = book["防暑降温费"]
        _set(changes, heat, 1, 11, 23, "操作手册", "当月应出勤23天")
        for row in range(4, 44):
            name = norm(payroll.cell(row, 3).value)
            if name != "李楠":
                formula = payroll.cell(row, 48).value
                if isinstance(formula, str) and formula.startswith("=") and "/21" in formula:
                    _set(changes, payroll, row, 48, formula.replace("/21", "/23"), "操作手册", "AU分母21改23")
            if name in {"李晓苛", "袁艳"}:
                formula = payroll.cell(row, 49).value
                if isinstance(formula, str) and formula.startswith("=") and "/21" in formula:
                    _set(changes, payroll, row, 49, formula.replace("/21", "/23"), "操作手册", "AV分母21改23")

        # Preserve the previous payroll-slip page as a values/formulas snapshot.
        payslip = book["工资条"]
        for row in range(1, 42):
            _copy_row(payslip, row, row + 45, min(74, payslip.max_column))
        _set(changes, payslip, 45, 1, "累计工资条", "原始总表", "沉淀上月工资条")

        oa = book["OA请款及审批"]
        for row in range(13, 7, -1):
            _copy_row(
                oa,
                row,
                row + 1,
                min(11, oa.max_column),
                value_sheet=cached_master["OA请款及审批"] if row == 8 else None,
            )
        _set(changes, oa, 8, 1, "7月", "操作手册", "OA新增当月记录")

        # Rebuild the two external-data sheets by stable headers/identity.
        ledger = book["台账"]
        ledger_rows = {norm(ledger.cell(row, 2).value): row for row in range(3, 44) if norm(ledger.cell(row, 2).value)}
        social_map = {4: 13, 5: 14, 6: 16, 7: 17, 8: 19, 10: 20, 11: 21, 12: 25, 13: 26}
        for source_name, source_path in social_paths.items():
            source_book = openpyxl.load_workbook(Path(source_path), data_only=True, keep_links=False)
            try:
                source = source_book.active
                for source_row in range(5, source.max_row + 1):
                    name = norm(source.cell(source_row, 2).value)
                    target_row = ledger_rows.get(name)
                    if not target_row:
                        if name:
                            issues.append({"item": "台账", "detail": f"社保人员未在总表匹配：{name}"})
                        continue
                    formula_source_book = openpyxl.load_workbook(Path(source_path), data_only=False, read_only=True, keep_links=False)
                    try:
                        formula_source = formula_source_book.active
                        for target_col, source_col in social_map.items():
                            value = source.cell(source_row, source_col).value
                            raw_value = formula_source.cell(source_row, source_col).value
                            if isinstance(raw_value, str) and raw_value.startswith("=") and value is None:
                                issues.append({"item": f"台账!{ledger.cell(target_row, target_col).coordinate}", "detail": f"{source_name} {name} 的来源公式没有缓存值"})
                                continue
                            if value in (None, ""):
                                continue
                            if target_col == 8 and "鹤安长泰" in source_name:
                                value = round(number(source.cell(source_row, 11).value) * 0.004, 2)
                            _set(changes, ledger, target_row, target_col, value, source_name, "按姓名匹配社保公积金")
                    finally:
                        formula_source_book.close()
            finally:
                source_book.close()
        for row in range(3, 43):
            _set(changes, ledger, row, 17, None, "操作手册", "清空本月商保历史值")

        tax = book["个税导出核对"]
        for row in list(range(2, 41)) + list(range(42, 48)):
            for col in range(1, 40):
                _set(changes, tax, row, col, None, "操作手册", "清空当月个税数据区")
        headers = {norm(tax.cell(1, col).value): col for col in range(1, 40) if norm(tax.cell(1, col).value)}
        aliases = {"累计婴幼儿照护费用": "累计3岁以下婴幼儿照护", "累计已预缴税额": "已缴税额", "累计应补(退)税额": "应补(退)税额"}
        next_rows = {"益药科园": 2, "鹤安长泰": 42}
        limits = {"益药科园": 41, "鹤安长泰": 48}
        tax_order = (
            "202608_税款计算_工资薪金所得-益药科园.xls",
            "202608_税款计算_劳务报酬-益药科园.xls",
            "202608_税款计算_劳务报酬-鹤安长泰.xls",
            "202608_税款计算_工资薪金所得-鹤安长泰.xls",
        )
        for source_name in ("益药科园", "鹤安长泰"):
            for tax_name in tax_order:
                tax_path = tax_paths.get(tax_name)
                if tax_path is None:
                    continue
                if source_name not in tax_name:
                    continue
                frame = _tax_frame(Path(tax_path))
                source_headers = {norm(column): column for column in frame.columns}
                for _, record in frame.iterrows():
                    name_column = source_headers.get("姓名")
                    name = norm(record.get(name_column)) if name_column else ""
                    if not name or "合计" in name:
                        continue
                    target_row = next_rows[source_name]
                    if target_row >= limits[source_name]:
                        raise KeyuanPythonError(f"{source_name} 个税记录超过模板预留行")
                    for header, target_col in headers.items():
                        source_col = source_headers.get(header) or source_headers.get(aliases.get(header, ""))
                        if source_col is not None:
                            value = record.get(source_col)
                            _set(changes, tax, target_row, target_col, None if pd.isna(value) else value, tax_name, "按表头匹配个税")
                    next_rows[source_name] += 1
        book.calculation.fullCalcOnLoad = True
        book.calculation.forceFullCalc = True
        book.calculation.calcMode = "auto"
        book.save(output_path)
    finally:
        salary.close()
        salary_formulas.close()
        cached_master.close()
        book.close()

    verify = _read_book(output_path)
    try:
        missing = {"工资核算", "台账", "个税导出核对", "OA请款及审批"}.difference(verify.sheetnames)
        formula_errors = [f"{sheet.title}!{cell.coordinate}" for sheet in verify.worksheets for row in sheet.iter_rows() for cell in row if isinstance(cell.value, str) and "#REF!" in cell.value]
        structural = {"sheet_names": not missing, "payroll_columns": verify["工资核算"].max_column == 99, "payroll_header": verify["工资核算"]["A2"].value == "班制", "formula_references": not formula_errors}
    finally:
        verify.close()
    if not all(structural.values()):
        raise KeyuanPythonError(f"科园结果结构校验失败：{structural}")
    return {"status": "passed" if not issues else "needs_review", "output_path": str(output_path), "change_count": len(changes), "changes": changes, "issues": issues, "structural": structural}


def main() -> int:
    raise SystemExit("该执行器由 Agent 通过受控适配器调用")


if __name__ == "__main__":
    main()
