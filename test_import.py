"""测试实体导入流程。"""
from core.entity_engine import (
    import_file_to_entities,
    import_template_to_entities,
    merge_entities,
    get_entity_overview,
    export_to_template,
)

# 第0步：导入模板初始数据
print("=" * 60)
print("测试0: 模板初始数据导入")
tpl_result = import_template_to_entities(
    r"C:\Users\CHJHR-yangben\Desktop\上药项目\大药房工资核算总表-v2(1).xlsx"
)
tpl_logs = tpl_result.pop("_log", [])
for l in tpl_logs:
    print(l)

# 测试薪资数据导入
print("\n" + "=" * 60)
print("测试1: 薪资数据导入")
result1 = import_file_to_entities(
    r"C:\Users\CHJHR-yangben\Desktop\上药项目\薪资数据-科园-6月薪资（7.15发薪）.xlsx",
    "薪资数据-科园-6月薪资（7.15发薪）.xlsx",
)
logs1 = result1.pop("_log", [])
for l in logs1:
    print(l)

# 测试社保数据导入
print("\n" + "=" * 60)
print("测试2: 社保-鹤安导入")
result2 = import_file_to_entities(
    r"C:\Users\CHJHR-yangben\Desktop\上药项目\鹤安长泰2026.06.xlsx",
    "鹤安长泰2026.06.xlsx",
)
logs2 = result2.pop("_log", [])
for l in logs2:
    print(l)

print("\n" + "=" * 60)
print("测试3: 社保-益药导入")
result3 = import_file_to_entities(
    r"C:\Users\CHJHR-yangben\Desktop\上药项目\益药科园2026.06.xlsx",
    "益药科园2026.06.xlsx",
)
logs3 = result3.pop("_log", [])
for l in logs3:
    print(l)

# 合并所有（模板 + 源数据）
print("\n" + "=" * 60)
print("合并所有实体（模板 + 源数据）")
collected = {}
for r in [tpl_result, result1, result2, result3]:
    for eid, records in r.items():
        collected.setdefault(eid, []).extend(records)

merged = merge_entities(collected)
overview = get_entity_overview(merged)
print("\n=== 实体概览 ===")
for o in overview:
    print(f'  {o["icon"]} {o["name"]}: {o["row_count"]}行, {o["field_count"]}字段, 完整度{o["completeness_pct"]}%')

# 测试导出
print("\n" + "=" * 60)
print("测试4: 模板导出")
import os
output_path = os.path.join(os.path.dirname(__file__), "test_export.xlsx")
template_path = r"C:\Users\CHJHR-yangben\Desktop\上药项目\大药房工资核算总表-v2(1).xlsx"
try:
    log = export_to_template(template_path, merged, output_path, freeze_formulas=True)
    for l in log:
        print(l)
except Exception as e:
    print(f"导出失败: {e}")

print("\n完成")
