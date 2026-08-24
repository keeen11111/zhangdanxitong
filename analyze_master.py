"""深入分析 Master 模板的「工资核算」Sheet —— 导出主表的核心结构。"""
import openpyxl

path = r"C:\Users\CHJHR-yangben\Desktop\上药项目\大药房工资核算总表-v2(1).xlsx"
wb = openpyxl.load_workbook(path)

print("Master 模板所有 Sheet：", wb.sheetnames)
print()

# 重点分析「工资核算」Sheet
ws = wb["工资核算"]
print(f"=== 「工资核算」Sheet: {ws.max_row} 行 × {ws.max_column} 列 ===")
print()

# 表头行（第1行）—— 全部列
print("--- 第1行表头（全部列）---")
for c in range(1, ws.max_column + 1):
    v = ws.cell(row=1, column=c).value
    print(f"  列{c} ({openpyxl.utils.get_column_letter(c)}): {v}")

print()

# 第2行内容（公式/数据样例）
print("--- 第2行内容（公式/数据样例）---")
for c in range(1, ws.max_column + 1):
    cell = ws.cell(row=2, column=c)
    v = cell.value
    header = ws.cell(row=1, column=c).value
    if v is not None:
        print(f"  列{c} ({openpyxl.utils.get_column_letter(c)}) [{header}]: "
              f"值={v}  类型={type(v).__name__}  格式={cell.number_format}")

print()

# 看第3、4行是否还有公式（确认数据起始行）
print("--- 第3行内容 ---")
for c in range(1, ws.max_column + 1):
    cell = ws.cell(row=3, column=c)
    v = cell.value
    header = ws.cell(row=1, column=c).value
    if v is not None:
        print(f"  列{c} [{header}]: 值={v}  类型={type(v).__name__}")

print()

# 统计公式列 vs 静态列
print("--- 公式列 vs 静态列统计 ---")
formula_cols = []
static_cols = []
for c in range(1, ws.max_column + 1):
    v = ws.cell(row=2, column=c).value
    header = ws.cell(row=1, column=c).value
    if isinstance(v, str) and v.startswith("="):
        formula_cols.append((c, header, v))
    elif v is not None:
        static_cols.append((c, header, v))

print(f"公式列 ({len(formula_cols)} 个):")
for c, h, v in formula_cols:
    print(f"  列{c} [{h}]: {v[:80]}...")
print()
print(f"静态数据列 ({len(static_cols)} 个):")
for c, h, v in static_cols:
    print(f"  列{c} [{h}]: {v}")
