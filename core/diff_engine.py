"""差异引擎（模块C）：原始数据库与数据源头表的匹配计算。

核心职责：
    1. 将原始数据库（模块A）与数据源头表（模块B）进行匹配
    2. 生成差异项（新增/离职/调岗/薪资变动/考勤/社保）
    3. 差异项供模块D审核后拼接为最终数据库

差异类型：
    - hire:       新增入职（源数据有，原始库无）
    - depart:     离职（入离职表标记，或原始库有但源数据无）
    - transfer:   调岗（部门/岗位变化）
    - salary:     薪资变动（基本工资变化）
    - attendance: 考勤数据（新增）
    - social:     社保数据（新增）
    - tax:        个税数据（新增）
"""
from __future__ import annotations

from datetime import datetime
from typing import Any

from core.normalizer import DataNormalizer


_normalizer = DataNormalizer()


def _clean(val: Any, dtype: str = "text") -> Any:
    """清理数值。"""
    if dtype == "text":
        return _normalizer.clean_text(val)
    elif dtype == "amount":
        return _normalizer.clean_amount(val)
    elif dtype == "date":
        return _normalizer.clean_date(val)
    return val


def _get_key(rec: dict, primary_key: str = "工号") -> str:
    """获取记录的主键值，支持姓名回退。"""
    pk = str(rec.get(primary_key, "")).strip()
    if pk:
        return pk
    # 回退到姓名
    name = str(rec.get("姓名", "")).strip()
    return f"_name_{name}" if name else ""


def compute_diff(
    base_db: dict[str, list[dict]],
    source_data: dict[str, list[dict]],
) -> dict[str, list[dict]]:
    """计算原始数据库与数据源头表的差异。

    Args:
        base_db: 原始数据库（模块A），包含 employee_profile, tax_record 等
        source_data: 数据源头表（模块B），包含从源文件提取的实体数据

    Returns:
        差异字典，按类型分组：
        {
            "hire": [...],        # 新增入职
            "depart": [...],      # 离职
            "transfer": [...],    # 调岗
            "salary_change": [...], # 薪资变动
            "attendance": [...],  # 考勤数据（全部新增）
            "social": [...],      # 社保数据（全部新增）
            "tax": [...],         # 个税数据（全部新增）
            "summary": {...}      # 汇总统计
        }
    """
    diff: dict[str, list[dict]] = {
        "hire": [],
        "depart": [],
        "transfer": [],
        "salary_change": [],
        "attendance": [],
        "social": [],
        "tax": [],
    }

    # ---- 1. 员工档案匹配 ----
    base_emps = base_db.get("employee_profile", [])
    source_emps = source_data.get("employee_profile", [])

    # 构建原始库索引（工号→记录，姓名→工号）
    base_by_id: dict[str, dict] = {}
    name_to_id: dict[str, str] = {}
    for rec in base_emps:
        pk = _get_key(rec)
        if pk and not pk.startswith("_name_"):
            base_by_id[pk] = rec
            name = str(rec.get("姓名", "")).strip()
            if name:
                name_to_id[name] = pk

    # 构建源数据索引
    source_by_id: dict[str, dict] = {}
    source_names: set[str] = set()
    for rec in source_emps:
        pk = _get_key(rec)
        if pk and not pk.startswith("_name_"):
            source_by_id[pk] = rec
        name = str(rec.get("姓名", "")).strip()
        if name:
            source_names.add(name)

    # 新员工的薪资与身份证分散在不同源表中，按工号和唯一姓名做确定性关联。
    source_salary_by_id: dict[str, dict] = {}
    for rec in source_data.get("salary_detail", []):
        salary_key = _get_key(rec)
        if not salary_key or salary_key.startswith("_name_"):
            continue
        merged = source_salary_by_id.setdefault(salary_key, {})
        for field, value in rec.items():
            if value is not None and value != "":
                merged[field] = value

    social_ids_by_name: dict[str, set[str]] = {}
    for rec in source_data.get("social_security", []):
        name = str(rec.get("姓名", "")).strip()
        identity = str(rec.get("身份证号", "")).strip()
        if name and identity:
            social_ids_by_name.setdefault(name, set()).add(identity)
    unique_social_id = {
        name: next(iter(identities))
        for name, identities in social_ids_by_name.items()
        if len(identities) == 1
    }

    # 从入离职表提取的变动信息（源数据中的 employee_profile 包含本月状态）
    for src_rec in source_emps:
        status = str(src_rec.get("本月状态", "")).strip()
        pk = _get_key(src_rec)
        # 如果无工号，尝试通过姓名匹配
        if not pk or pk.startswith("_name_"):
            name = str(src_rec.get("姓名", "")).strip()
            if name in name_to_id:
                pk = name_to_id[name]

        if status == "入职":
            name = str(src_rec.get("姓名", "")).strip()
            salary_info = {
                key: value
                for key, value in source_salary_by_id.get(pk, {}).items()
                if key not in {"工号", "姓名", "_source_file", "_source_sheet"}
            }
            profile_info = {
                key: value
                for key, value in src_rec.items()
                if not key.startswith("_") and value is not None and value != ""
            }
            if not profile_info.get("身份证号") and name in unique_social_id:
                profile_info["身份证号"] = unique_social_id[name]
            # 新增入职
            diff["hire"].append({
                "姓名": src_rec.get("姓名", ""),
                "工号": pk if pk and not pk.startswith("_name_") else src_rec.get("工号", ""),
                "公司": src_rec.get("公司", ""),
                "部门": src_rec.get("部门", ""),
                "岗位": src_rec.get("岗位", ""),
                "入职日期": src_rec.get("入职日期", ""),
                "档案信息": profile_info,
                "薪资信息": salary_info,
                "source": "入离职表",
                "status": "pending",  # pending / confirmed / ignored
            })

        elif status == "离职":
            # 离职
            base_rec = base_by_id.get(pk, {})
            diff["depart"].append({
                "姓名": src_rec.get("姓名", ""),
                "工号": pk if pk and not pk.startswith("_name_") else base_rec.get("工号", ""),
                "离职日期": src_rec.get("离职日期", ""),
                "原部门": base_rec.get("部门", ""),
                "source": "入离职表",
                "status": "pending",
            })

        elif status == "调岗":
            # 调岗
            base_rec = base_by_id.get(pk, {})
            old_dept = base_rec.get("部门", "")
            new_dept = src_rec.get("部门", "")
            old_pos = base_rec.get("岗位", "")
            new_pos = src_rec.get("岗位", "")
            changes = []
            if old_dept != new_dept and new_dept:
                changes.append(f"部门: {old_dept} → {new_dept}")
            if old_pos != new_pos and new_pos:
                changes.append(f"岗位: {old_pos} → {new_pos}")
            if changes:
                diff["transfer"].append({
                    "姓名": src_rec.get("姓名", ""),
                    "工号": pk if pk and not pk.startswith("_name_") else base_rec.get("工号", ""),
                    "变动内容": "; ".join(changes),
                    "原部门": old_dept,
                    "新部门": new_dept,
                    "原岗位": old_pos,
                    "新岗位": new_pos,
                    "source": "入离职表",
                    "status": "pending",
                })

        elif status == "转正":
            # 转正（作为调岗的一种特殊形式）
            base_rec = base_by_id.get(pk, {})
            diff["transfer"].append({
                "姓名": src_rec.get("姓名", ""),
                "工号": pk if pk and not pk.startswith("_name_") else base_rec.get("工号", ""),
                "变动内容": f"转正日期: {src_rec.get('转正日期', '')}",
                "原部门": base_rec.get("部门", ""),
                "新部门": base_rec.get("部门", ""),
                "原岗位": base_rec.get("岗位", ""),
                "新岗位": base_rec.get("岗位", ""),
                "转正日期": src_rec.get("转正日期", ""),
                "source": "入离职表",
                "status": "pending",
            })

    # 检查原始库中有但源数据中完全消失的员工（可能离职）
    source_ids = set(source_by_id.keys())
    for pk, base_rec in base_by_id.items():
        name = str(base_rec.get("姓名", "")).strip()
        if pk not in source_ids and name not in source_names:
            # 原始库有，但源数据完全无此人
            # 检查是否已在离职差异中（避免重复）
            already_departed = any(
                d.get("工号") == pk or d.get("姓名") == name
                for d in diff["depart"]
            )
            if not already_departed:
                diff["depart"].append({
                    "姓名": name,
                    "工号": pk,
                    "离职日期": "",
                    "原部门": base_rec.get("部门", ""),
                    "source": "原始库中存在但源数据无此人",
                    "status": "pending",
                })

    # ---- 2. 薪资变动检测 ----
    base_salary = { _get_key(r): r for r in base_db.get("salary_detail", []) if _get_key(r) }
    for src_rec in source_data.get("salary_detail", []):
        pk = _get_key(src_rec)
        if not pk or pk.startswith("_name_"):
            name = str(src_rec.get("姓名", "")).strip()
            if name in name_to_id:
                pk = name_to_id[name]
        if pk in base_salary:
            base_sal = base_salary[pk]
            # 比较基本工资
            old_base = _clean(base_sal.get("月基本薪资"), "amount")
            new_base = _clean(src_rec.get("月基本薪资"), "amount")
            if old_base and new_base and abs(old_base - new_base) > 0.01:
                diff["salary_change"].append({
                    "姓名": src_rec.get("姓名", base_sal.get("姓名", "")),
                    "工号": pk,
                    "变动项": "月基本薪资",
                    "原值": old_base,
                    "新值": new_base,
                    "差额": new_base - old_base,
                    "source": "薪资数据",
                    "status": "pending",
                })

    # ---- 3. 考勤数据（全部作为新增差异） ----
    for rec in source_data.get("attendance_record", []):
        diff["attendance"].append({**rec, "status": "pending"})

    # ---- 4. 社保数据（全部作为新增差异） ----
    for rec in source_data.get("social_security", []):
        diff["social"].append({**rec, "status": "pending"})

    # ---- 5. 个税数据（全部作为新增差异） ----
    for rec in source_data.get("tax_record", []):
        diff["tax"].append({**rec, "status": "pending"})

    # ---- 汇总统计 ----
    diff["summary"] = {
        "hire_count": len(diff["hire"]),
        "depart_count": len(diff["depart"]),
        "transfer_count": len(diff["transfer"]),
        "salary_change_count": len(diff["salary_change"]),
        "attendance_count": len(diff["attendance"]),
        "social_count": len(diff["social"]),
        "tax_count": len(diff["tax"]),
        "total_diff": sum(len(v) for k, v in diff.items() if k != "summary"),
        "computed_at": datetime.now().isoformat(),
    }

    return diff


def apply_confirmed_diff(
    base_db: dict[str, list[dict]],
    diff: dict[str, list[dict]],
) -> dict[str, list[dict]]:
    """将已确认的差异项应用到原始数据库，生成最终数据库（模块E）。

    仅应用 status == "confirmed" 的差异项，ignored 的跳过。
    """
    final_db: dict[str, list[dict]] = {}
    for eid, recs in base_db.items():
        final_db[eid] = [dict(r) for r in recs]

    # 构建员工档案索引
    emp_list = final_db.get("employee_profile", [])
    emp_by_id: dict[str, dict] = {}
    name_to_id: dict[str, str] = {}
    for rec in emp_list:
        pk = _get_key(rec)
        if pk and not pk.startswith("_name_"):
            emp_by_id[pk] = rec
            name = str(rec.get("姓名", "")).strip()
            if name:
                name_to_id[name] = pk

    # ---- 1. 应用入职 ----
    for item in diff.get("hire", []):
        if item.get("status") != "confirmed":
            continue
        pk = item.get("工号", "")
        new_rec = {
            key: value
            for key, value in item.get("档案信息", {}).items()
            if not key.startswith("_")
        }
        new_rec.update({
            "工号": pk,
            "姓名": item.get("姓名", ""),
            "公司": item.get("公司", ""),
            "部门": item.get("部门", ""),
            "岗位": item.get("岗位", ""),
            "入职日期": item.get("入职日期", ""),
            "本月状态": "入职",
        })
        sal_info = item.get("薪资信息", {})
        emp_list.append(new_rec)
        if pk:
            emp_by_id[pk] = new_rec
        if sal_info:
            final_db.setdefault("salary_detail", []).append({
                "工号": pk,
                "姓名": item.get("姓名", ""),
                **sal_info,
            })

    # ---- 2. 应用离职（从员工档案中移除） ----
    departed_keys: set[str] = set()
    for item in diff.get("depart", []):
        if item.get("status") != "confirmed":
            continue
        pk = item.get("工号", "")
        name = item.get("姓名", "")
        if not pk and name in name_to_id:
            pk = name_to_id[name]
        departed_keys.add(pk)
        departed_keys.add(f"_name_{name}")

    if departed_keys:
        final_db["employee_profile"] = [
            rec for rec in emp_list
            if _get_key(rec) not in departed_keys
        ]

    # ---- 3. 应用调岗 ----
    for item in diff.get("transfer", []):
        if item.get("status") != "confirmed":
            continue
        pk = item.get("工号", "")
        name = item.get("姓名", "")
        if not pk and name in name_to_id:
            pk = name_to_id[name]
        if pk in emp_by_id:
            rec = emp_by_id[pk]
            if item.get("新部门"):
                rec["部门"] = item["新部门"]
            if item.get("新岗位"):
                rec["岗位"] = item["新岗位"]
            if item.get("转正日期"):
                rec["转正日期"] = item["转正日期"]

    # ---- 4. 应用薪资变动 ----
    sal_list = final_db.get("salary_detail", [])
    sal_by_id: dict[str, dict] = {}
    for rec in sal_list:
        pk = _get_key(rec)
        if pk and not pk.startswith("_name_"):
            sal_by_id[pk] = rec

    for item in diff.get("salary_change", []):
        if item.get("status") != "confirmed":
            continue
        pk = item.get("工号", "")
        if pk in sal_by_id:
            field = item.get("变动项", "")
            if field:
                sal_by_id[pk][field] = item.get("新值")

    # ---- 5. 考勤/社保/个税（已确认的直接覆盖写入） ----
    for eid, diff_key in [
        ("attendance_record", "attendance"),
        ("social_security", "social"),
        ("tax_record", "tax"),
    ]:
        confirmed = [
            {k: v for k, v in item.items() if k != "status"}
            for item in diff.get(diff_key, [])
            if item.get("status") == "confirmed"
        ]
        if confirmed:
            final_db[eid] = confirmed

    return final_db
