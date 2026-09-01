from __future__ import annotations

import json
import re
import sys
from copy import copy
from datetime import datetime
from io import BytesIO
from pathlib import Path

import openpyxl
import pandas as pd


def norm(value: object) -> str:
    return re.sub(r"\s+", "", str(value or "")).replace("（", "(").replace("）", ")")


workbook_path = Path(sys.argv[1])
sample_dir = Path(sys.argv[2])
changes: list[dict[str, object]] = []
book = openpyxl.load_workbook(workbook_path, data_only=False)


def set_cell(sheet, row: int, col: int, value: object, source: str, rule: str) -> None:
    cell = sheet.cell(row, col)
    old_value = cell.value
    if old_value == value:
        return
    cell.value = value
    changes.append({
        "sheet": sheet.title,
        "cell": cell.coordinate,
        "before": old_value,
        "after": value,
        "source": source,
        "rule": rule,
    })


def copy_cell_style(source, target) -> None:
    target._style = copy(source._style)
    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)

ledger = book["台账"]
ledger_rows = {norm(ledger.cell(row, 2).value): row for row in range(3, 44) if norm(ledger.cell(row, 2).value)}
social_map = {4: 13, 5: 14, 6: 16, 7: 17, 8: 19, 10: 20, 11: 21, 12: 25, 13: 26}
for source_path in sorted(sample_dir.glob("*2026.07.xlsx")):
    if not re.search(r"益药科园|鹤安长泰", source_path.name):
        continue
    source_book = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    try:
        source = source_book.active
        for source_row in range(5, source.max_row + 1):
            name = norm(source.cell(source_row, 2).value)
            target_row = ledger_rows.get(name)
            if not target_row:
                continue
            for target_col, source_col in social_map.items():
                new_value = source.cell(source_row, source_col).value
                if target_col == 8 and "鹤安长泰" in source_path.name:
                    injury_base = source.cell(source_row, 11).value
                    new_value = round(float(injury_base) * 0.004, 2) if injury_base is not None else None
                old_value = ledger.cell(target_row, target_col).value
                if old_value != new_value:
                    ledger.cell(target_row, target_col).value = new_value
                    changes.append({
                        "sheet": "台账", "cell": ledger.cell(target_row, target_col).coordinate,
                        "before": old_value, "after": new_value, "source": source_path.name,
                        "rule": "按姓名匹配社保公积金；公司住房写入单位住房",
                    })
    finally:
        source_book.close()
for row in ledger.iter_rows(min_row=1, max_row=ledger.max_row, min_col=18, max_col=20):
    for cell in row:
        cell.value = None
for row in range(3, 43):
    set_cell(ledger, row, 17, None, "操作手册", "清空本月无商保人员的历史值")

tax_sheet = book["个税导出核对"]
for row in list(range(2, 41)) + list(range(42, 48)):
    for col in range(1, 40):
        tax_sheet.cell(row, col).value = None
target_headers = {
    norm(tax_sheet.cell(1, col).value): col
    for col in range(1, 40)
    if norm(tax_sheet.cell(1, col).value)
}
tax_aliases = {
    "累计婴幼儿照护费用": "累计3岁以下婴幼儿照护",
    "累计已预缴税额": "已缴税额",
    "累计应补(退)税额": "应补(退)税额",
}
tax_dir = next(sample_dir.glob("回复：202608*"))
def tax_order(path: Path) -> tuple[str, int]:
    company = "益药科园" if "益药科园" in path.name else "鹤安长泰"
    labor = "劳务报酬" in path.name
    order = (1 if labor else 0) if company == "益药科园" else (0 if labor else 1)
    return company, order


tax_files = sorted(
    [path for path in tax_dir.iterdir() if path.is_file() and "税款计算" in path.name],
    key=tax_order,
)
next_rows = {"益药科园": 2, "鹤安长泰": 42}
limits = {"益药科园": 41, "鹤安长泰": 48}
for source_path in tax_files:
    company = "益药科园" if "益药科园" in source_path.name else "鹤安长泰"
    content = source_path.read_bytes()
    frame = (
        pd.read_excel(BytesIO(content), dtype=object, engine="openpyxl")
        if content.startswith(b"PK")
        else pd.read_excel(source_path, dtype=object)
    )
    source_columns = {norm(column): column for column in frame.columns}
    for _, record in frame.iterrows():
        name = norm(record.get(source_columns.get("姓名")))
        if not name or "合计" in name:
            continue
        target_row = next_rows[company]
        if target_row >= limits[company]:
            raise ValueError(f"{company} 个税记录超过模板预留行")
        for header, target_col in target_headers.items():
            source_col = source_columns.get(header) or source_columns.get(tax_aliases.get(header, ""))
            if source_col is None:
                continue
            value = record.get(source_col)
            tax_sheet.cell(target_row, target_col).value = None if pd.isna(value) else value
        changes.append({
            "sheet": "个税导出核对", "cell": f"A{target_row}:AM{target_row}",
            "before": "", "after": name, "source": source_path.name,
            "rule": "按表头匹配；益药在41行上方，鹤安在41行下方；保留41/48行合计",
        })
        next_rows[company] += 1

# The first-column insertion performed by WPS shifts several cross-sheet ranges
# incorrectly. Restore the intended references and the formulas present in the
# approved monthly workbook.
payroll = book["工资核算"]
for row in range(4, 44):
    formulas = {
        62: f"=VLOOKUP($B{row},台账!$A:Q,4,FALSE)",
        63: f"=VLOOKUP($B{row},台账!$A:Q,10,FALSE)",
        64: f"=VLOOKUP($B{row},台账!$A:Q,12,FALSE)",
        65: f"=VLOOKUP($B{row},台账!$A:R,6,FALSE)",
        66: f"=VLOOKUP($B{row},台账!$A:S,9,FALSE)",
        67: f"=VLOOKUP($B{row},台账!$A:T,8,FALSE)",
        72: (
            f"=VLOOKUP($B{row},台账!$A:AK,15,FALSE)+"
            f"VLOOKUP($B{row},台账!$A:AK,16,FALSE)+"
            f"VLOOKUP($B{row},台账!$A:AK,17,FALSE)"
        ),
    }
    for col, formula in formulas.items():
        set_cell(payroll, row, col, formula, "列插入校正", "恢复跨表公式引用")
    if isinstance(payroll.cell(row, 68).value, str) and payroll.cell(row, 68).value.startswith("="):
        set_cell(
            payroll,
            row,
            68,
            f"=VLOOKUP($B{row},台账!$A:U,14,FALSE)",
            "列插入校正",
            "恢复公积金公司公式引用",
        )
    set_cell(payroll, row, 95, None, "操作手册", "清除上月外部核对辅助公式")
    set_cell(payroll, row, 96, None, "操作手册", "清除上月外部核对辅助公式")

row = 43
for col, formula in {
    47: f"=ROUND(ROUND(AA{row}/21.75/8,2)*(AK{row}*150%+AL{row}*300%),2)",
    48: f"=IF(ROUND(Q{row}/15*AJ{row},0)>400,400,ROUND(Q{row}/15*AJ{row},0))",
    49: f"=IF(ROUND(500/15*AJ{row},0)>500,500,ROUND(500/15*AJ{row},0))",
    70: f"=VLOOKUP(C{row},个税导出核对!$B$2:$AL$48,37,0)",
    74: f"=IF(BU{row}<0,-BU{row},0)",
}.items():
    set_cell(payroll, row, col, formula, "工资核算模板", "补齐新增人员公式")
for col in (54, 62, 63, 64, 65, 66, 67):
    column = openpyxl.utils.get_column_letter(col)
    set_cell(payroll, 50, col, f"=SUBTOTAL(9,{column}4:{column}43)", "工资核算模板", "补齐合计公式")
set_cell(payroll, 51, 67, "=BO50+BM50+BL50+BK50+BJ50+BB50", "工资核算模板", "补齐公司成本合计")

# Apply effective internal-transfer salary changes from the monthly personnel
# change sheet. The sample contains 李文慧's 2026-07-01 adjustment to 1050.
salary_path = next(sample_dir.glob("薪资数据-科园-7月薪资*.xlsx"))
salary_book = openpyxl.load_workbook(salary_path, data_only=True, read_only=True)
heat_annual_by_name: dict[str, object] = {}
try:
    transfer = salary_book["入离职、转岗、转正、其他"]
    payroll_rows = {norm(payroll.cell(row, 3).value): row for row in range(4, 44)}
    for source_row in range(24, transfer.max_row + 1):
        name = norm(transfer.cell(source_row, 2).value)
        effective_date = transfer.cell(source_row, 10).value
        monthly_salary = transfer.cell(source_row, 11).value
        if (
            name in payroll_rows
            and isinstance(effective_date, datetime)
            and effective_date.year == 2026
            and effective_date.month == 7
            and isinstance(monthly_salary, (int, float))
        ):
            target_row = payroll_rows[name]
            set_cell(payroll, target_row, 25, monthly_salary, salary_path.name, "当月内部调动月薪")
            set_cell(payroll, target_row, 27, monthly_salary, salary_path.name, "同步调整工资基数")
    heat_source = salary_book["高温费"]
    heat_annual_by_name = {
        norm(heat_source.cell(row, 1).value): heat_source.cell(row, 2).value
        for row in range(2, heat_source.max_row + 1)
        if norm(heat_source.cell(row, 1).value)
    }
finally:
    salary_book.close()

# Match the approved summary layout: 益药 subtotal, 鹤安 row, overall total.
summary = book["工资汇总表"]
if summary["A11"].value != "小计":
    summary.insert_rows(11, 1)
    summary.delete_rows(17, 1)
for col in range(1, 24):
    copy_cell_style(summary.cell(13, min(col, 22)), summary.cell(11, col))
set_cell(summary, 2, 23, "实发工资", "工资汇总模板", "新增实发工资列")
for row in range(3, 11):
    set_cell(summary, row, 23, f"=E{row}-H{row}-I{row}-J{row}-K{row}-T{row}", "工资汇总模板", "计算分部门实发")
set_cell(summary, 11, 1, "小计", "工资汇总模板", "益药小计")
for col in (2, 3):
    set_cell(summary, 11, col, None, "工资汇总模板", "益药小计")
for col in range(4, 24):
    column = openpyxl.utils.get_column_letter(col)
    set_cell(summary, 11, col, f"=SUM({column}3:{column}10)", "工资汇总模板", "益药小计")
for col in range(4, 24):
    if col in (7, 12, 19, 21, 22):
        set_cell(summary, 12, col, None, "工资汇总模板", "鹤安无对应项目")
        continue
    if col == 23:
        formula = "=E12-H12-I12-J12-K12-T12"
    else:
        payroll_column = {
            4: "F", 5: "BB", 6: "BW", 8: "BE", 9: "BF", 10: "BH", 11: "BG",
            13: "BJ", 14: "BK", 15: "BM", 16: "BN", 17: "BO", 18: "BL", 20: "BR",
        }[col]
        function = "COUNTIF" if col == 4 else "SUMIF"
        formula = (
            f"=COUNTIF(工资核算!F:F,C12)"
            if function == "COUNTIF"
            else f"=SUMIF(工资核算!$F:$F,工资汇总表!$C12,工资核算!{payroll_column}:{payroll_column})"
        )
    set_cell(summary, 12, col, formula, "工资汇总模板", "鹤安汇总")
for merged in list(summary.merged_cells.ranges):
    if str(merged) == "A13:C13":
        summary.unmerge_cells(str(merged))
summary.merge_cells("A13:C13")
set_cell(summary, 13, 1, "合计", "工资汇总模板", "两公司合计")
for col in range(4, 23):
    column = openpyxl.utils.get_column_letter(col)
    set_cell(summary, 13, col, f"=SUM({column}11:{column}12)", "工资汇总模板", "两公司合计")
set_cell(summary, 13, 23, "=E13-H13-I13-J13-K13-T13", "工资汇总模板", "两公司实发合计")
for row in (15, 16):
    for col in range(5, 24):
        set_cell(summary, row, col, None, "工资汇总模板", "清理旧校验区")
set_cell(summary, 15, 23, "=W13-工资核算!BU44", "工资汇总模板", "实发校验")

# The approved workbook removes the temporary attendance scratch area after
# importing the five required attendance fields.
attendance = book["考勤"]
for row in range(1, 49):
    for col in range(18, 33):
        set_cell(attendance, row, col, None, "操作手册", "清理考勤导入辅助区")
set_cell(attendance, 1, 18, "备注", "考勤模板", "保留备注表头")

heat = book["防暑降温费"]
new_standard_names = {"李晓苛", "袁艳", "李楠"}
for row in range(2, 42):
    cap = 300 if norm(heat.cell(row, 2).value) in new_standard_names else 100
    set_cell(heat, row, 7, f"=IF(H{row}>{cap},{cap},H{row})", "操作手册", "防暑降温费封顶")
    set_cell(
        heat,
        row,
        10,
        heat_annual_by_name.get(norm(heat.cell(row, 2).value), 0),
        salary_path.name,
        "按姓名写入年度高温费标准，消除外部链接",
    )

payslip = book["工资条"]
for row in range(2, 42):
    set_cell(payslip, row, 6, datetime(2026, 6, 1), "工资条模板", "所属工资期间开始")
    set_cell(payslip, row, 7, datetime(2026, 6, 30), "工资条模板", "所属工资期间结束")
for row in range(47, 87):
    set_cell(payslip, row, 6, datetime(2026, 6, 1), "工资条历史", "上月工资期间开始")
    set_cell(payslip, row, 7, datetime(2026, 6, 30), "工资条历史", "上月工资期间结束")
for coordinate, formula in {
    "BQ42": "=SUM(BQ2:BQ41)",
    "BS42": "=SUM(BS2:BS41)",
    "BU42": "=SUM(BU2:BU41)",
    "BQ43": "=BQ42-工资核算!BO51",
    "BU43": "=BU42-工资核算!BX44",
}.items():
    set_cell(payslip, payslip[coordinate].row, payslip[coordinate].column, formula, "工资条模板", "补齐工资条校验公式")

book.save(workbook_path)
book.close()
print(json.dumps(changes, ensure_ascii=False, default=str))
