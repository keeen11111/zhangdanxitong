"""分析业务文件结构：Sheet 名、表头、行数、公式。
仅读取结构，不完整展示敏感数据（身份证号等脱敏）。
"""
import openpyxl
import pandas as pd
import os

FILES = {
    "Master模板": r"C:\Users\CHJHR-yangben\Desktop\上药项目\大药房工资核算总表-v2(1).xlsx",
    "薪资数据": r"C:\Users\CHJHR-yangben\Desktop\上药项目\薪资数据-科园-6月薪资（7.15发薪）.xlsx",
    "社保-鹤安": r"C:\Users\CHJHR-yangben\Desktop\上药项目\鹤安长泰2026.06.xlsx",
    "社保-益药": r"C:\Users\CHJHR-yangben\Desktop\上药项目\益药科园2026.06.xlsx",
}


def mask_sensitive(val):
    """脱敏：身份证号只保留前4后2。"""
    s = str(val)
    if len(s) >= 15 and s.replace("x", "").replace("X", "").isdigit():
        return s[:4] + "***" + s[-2:]
    return s


def analyze_pandas(label, path):
    """用 pandas 读取数据结构。"""
    print(f"\n{'='*70}")
    print(f"【{label}】{os.path.basename(path)}")
    print(f"路径: {path}")
    if not os.path.exists(path):
        print("  *** 文件不存在 ***")
        return
    try:
        xls = pd.ExcelFile(path)
        print(f"Sheet 数: {len(xls.sheet_names)}")
        for i, sheet_name in enumerate(xls.sheet_names):
            df = pd.read_excel(path, sheet_name=sheet_name, nrows=3)
            print(f"\n  --- Sheet[{i}]: 「{sheet_name}」 ---")
            print(f"  行数(全部): ", end="")
            df_full = pd.read_excel(path, sheet_name=sheet_name)
            print(f"{len(df_full)} 行 × {len(df_full.columns)} 列")
            print(f"  表头: {list(df_full.columns)}")
            # 前 3 行样例（脱敏）
            if len(df_full) > 0:
                print(f"  样例(首行):")
                for col in df_full.columns:
                    val = mask_sensitive(df_full.iloc[0][col])
                    print(f"    {col}: {val}")
    except Exception as e:
        print(f"  读取错误: {e}")


def analyze_master_formulas(path):
    """专门分析 Master 模板的公式列。"""
    print(f"\n{'='*70}")
    print(f"【Master模板公式分析】")
    if not os.path.exists(path):
        print("  *** 文件不存在 ***")
        return
    wb = openpyxl.load_workbook(path)
    print(f"Sheet 列表: {wb.sheetnames}")
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        print(f"\n  --- Sheet: 「{ws_name}」 {ws.max_row}行 × {ws.max_column}列 ---")
        # 打印表头行（第1行）
        headers = []
        for c in range(1, min(ws.max_column + 1, 60)):
            v = ws.cell(row=1, column=c).value
            headers.append(str(v) if v else "")
        print(f"  表头(前60列): {headers}")
        
        # 找公式行（第2行）
        if ws.max_row >= 2:
            print(f"  第2行内容（含公式）:")
            for c in range(1, min(ws.max_column + 1, 60)):
                cell = ws.cell(row=2, column=c)
                v = cell.value
                if v is not None:
                    print(f"    列{c}({headers[c-1] if c-1<len(headers) else '?'}): "
                          f"值={mask_sensitive(v)}  类型={type(v).__name__}  "
                          f"格式={cell.number_format}")


for label, path in FILES.items():
    if label == "Master模板":
        analyze_master_formulas(path)
    else:
        analyze_pandas(label, path)

print(f"\n{'='*70}")
print("分析完成")
