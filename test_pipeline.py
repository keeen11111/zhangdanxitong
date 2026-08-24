"""端到端测试：6模块流水线完整流程。"""
import os
import json
from core.entity_engine import import_template_to_entities, import_file_to_entities
from core.diff_engine import compute_diff, apply_confirmed_diff
from core.entity_engine import get_entity_overview, export_to_template

TEMPLATE = r"C:\Users\CHJHR-yangben\Desktop\上药项目\大药房工资核算总表-v2(1).xlsx"
SALARY = r"C:\Users\CHJHR-yangben\Desktop\上药项目\薪资数据-科园-6月薪资（7.15发薪）.xlsx"
SOCIAL1 = r"C:\Users\CHJHR-yangben\Desktop\上药项目\鹤安长泰2026.06.xlsx"
SOCIAL2 = r"C:\Users\CHJHR-yangben\Desktop\上药项目\益药科园2026.06.xlsx"

# ============================================================
# 模块A：原始数据库（从模板提取）
# ============================================================
print("=" * 60)
print("模块A：原始数据库（从模板提取）")
base_db = import_template_to_entities(TEMPLATE)
logs = base_db.pop("_log", [])
for l in logs:
    print(f"  {l}")
print(f"\n  概览: {[(k, len(v)) for k, v in base_db.items()]}")

# ============================================================
# 模块B：数据源头表（从源文件提取）
# ============================================================
print("\n" + "=" * 60)
print("模块B：数据源头表（从源文件提取）")
source_data = {}
for fp, fn in [(SALARY, "薪资数据"), (SOCIAL1, "社保-鹤安"), (SOCIAL2, "社保-益药")]:
    print(f"\n  处理: {fn}")
    result = import_file_to_entities(fp, os.path.basename(fp))
    for l in result.pop("_log", []):
        print(f"    {l}")
    for eid, recs in result.items():
        source_data.setdefault(eid, []).extend(recs)
print(f"\n  概览: {[(k, len(v)) for k, v in source_data.items()]}")

# ============================================================
# 模块C：匹配引擎（计算差异）
# ============================================================
print("\n" + "=" * 60)
print("模块C：匹配引擎（计算差异）")
diff = compute_diff(base_db, source_data)
summary = diff["summary"]
print(f"  新增入职: {summary['hire_count']} 人")
print(f"  离职清理: {summary['depart_count']} 人")
print(f"  调岗转正: {summary['transfer_count']} 人")
print(f"  薪资变动: {summary['salary_change_count']} 项")
print(f"  考勤数据: {summary['attendance_count']} 条")
print(f"  社保数据: {summary['social_count']} 条")
print(f"  个税数据: {summary['tax_count']} 条")
print(f"  差异总数: {summary['total_diff']} 项")

# 打印入职/离职/调岗详情
for dt in ["hire", "depart", "transfer"]:
    items = diff.get(dt, [])
    if items:
        labels = {"hire": "新增入职", "depart": "离职", "transfer": "调岗转正"}
        print(f"\n  --- {labels[dt]} ---")
        for i, item in enumerate(items):
            name = item.get("姓名", "")
            detail = item.get("变动内容", item.get("离职日期", item.get("入职日期", "")))
            print(f"    [{i}] {name} - {detail}")

# ============================================================
# 模块D：差异审核（全部确认）
# ============================================================
print("\n" + "=" * 60)
print("模块D：差异审核（全部确认）")
for dt in ["hire", "depart", "transfer", "salary_change", "attendance", "social", "tax"]:
    items = diff.get(dt, [])
    for item in items:
        item["status"] = "confirmed"
    if items:
        print(f"  {dt}: {len(items)} 项已确认")

# ============================================================
# 模块E：最终数据库（拼接）
# ============================================================
print("\n" + "=" * 60)
print("模块E：最终数据库（拼接）")
final_db = apply_confirmed_diff(base_db, diff)
overview = get_entity_overview(final_db)
print("\n  === 最终数据库概览 ===")
for o in overview:
    print(f'    {o["icon"]} {o["name"]}: {o["row_count"]}行, {o["field_count"]}字段, 完整度{o["completeness_pct"]}%')

# ============================================================
# 模块F：模板导出
# ============================================================
print("\n" + "=" * 60)
print("模块F：模板导出")
output_path = os.path.join(os.path.dirname(__file__), "test_pipeline_export.xlsx")
try:
    log = export_to_template(TEMPLATE, final_db, output_path, freeze_formulas=True)
    for l in log:
        print(f"  {l}")
    print(f"\n  导出文件: {output_path}")
except Exception as e:
    print(f"  导出失败: {e}")

print("\n" + "=" * 60)
print("完整流程测试完成！")
