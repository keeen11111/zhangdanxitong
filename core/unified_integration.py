"""统一总表整合中的人员并集规则。"""

from __future__ import annotations

import calendar
from copy import deepcopy
from datetime import date, datetime
import os
import re
import shutil
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.formula import ArrayFormula


_NON_PERSON_NAMES = {
    "合计",
    "总计",
    "小计",
    "汇总",
    "姓名",
    "人员姓名",
    "员工姓名",
    "总经理：",
    "总经理:",
    "序号",
    "无",
    "离职、退休",
    "离职退休",
    "试用期转正",
    "内部调动",
    "其他",
    "旷工",
    "分类",
    "销售目标",
    "销售达成",
    "达成率",
}

_NAME_HEADERS = {"姓名", "人员姓名", "员工姓名", "名字"}
_EMPLOYEE_ID_HEADERS = {"工号", "新工号", "员工工号", "员工编号", "人员编码", "工牌号"}
_ID_CARD_HEADERS = {"身份证号", "身份证号码", "证件号码", "身份证"}
_COMPANY_HEADERS = {"公司", "新公司", "所属公司", "缴费公司", "单位名称", "公司名称"}
_DEPARTMENT_HEADERS = {"部门", "所属部门", "成本中心/部门", "新成本中心/部门", "成本中心"}
_ENTRY_DATE_HEADERS = {"入职日期", "入职时间", "入职日"}
_MEAL_STANDARD_HEADERS = {"餐补标准", "餐补标准随工资发", "餐补", "饭贴"}


def _text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def _valid_name(value: Any) -> bool:
    name = _text(value)
    lowered = name.lower()
    return (
        bool(name)
        and len(name) <= 30
        and not name.isdigit()
        and name not in _NON_PERSON_NAMES
        and "dtype:" not in lowered
        and "name:" not in lowered
        and "销售额达成情况" not in name
        and not any(token in name for token in ("打卡", "不算迟到", "视为迟到", "豁免", "不扣工资"))
        and not name.endswith(("：", ":"))
    )


def _header(value: Any) -> str:
    return "".join(_text(value).split())


def _same_department(master_value: str, change_value: str) -> bool:
    """Compare department values without inferring a hierarchy relationship."""
    master = _header(master_value)
    change = _header(change_value)
    return bool(master and change and master == change)


def _salary_month_cutoff(salary_month: str | None) -> date | None:
    if not salary_month:
        return None
    match = re.fullmatch(r"\s*(\d{4})[.\-/年](\d{1,2})(?:月)?\s*", salary_month)
    if not match:
        return None
    year, month = (int(part) for part in match.groups())
    if month < 1 or month > 12:
        return None
    return date(year, month, calendar.monthrange(year, month)[1])


def _record_effective_date(record: dict[str, Any]) -> date | None:
    for field in (
        "_effective_date",
        "离职日期",
        "转正日期",
        "入职日期",
        "税款所属期止",
        "税款所属期起",
    ):
        raw_value = record.get(field)
        if raw_value in (None, ""):
            continue
        if isinstance(raw_value, datetime):
            return raw_value.date()
        if isinstance(raw_value, date):
            return raw_value
        text = _text(raw_value)
        for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
            try:
                return datetime.strptime(text, pattern).date()
            except ValueError:
                continue
    return None


def _dated_field_resolution(
    field_rows: list[dict[str, Any]],
    field: str,
    cutoff: date | None,
) -> dict[str, Any]:
    source_rows = [row for row in field_rows if not row.get("_roster_authoritative")]
    dated_rows: list[tuple[date, dict[str, Any]]] = []
    missing_rows: list[dict[str, Any]] = []
    future_rows: list[dict[str, Any]] = []
    for row in source_rows:
        effective_date = _record_effective_date(row)
        if effective_date is None:
            missing_rows.append(row)
        elif cutoff and effective_date > cutoff:
            future_rows.append(row)
        else:
            dated_rows.append((effective_date, row))

    result: dict[str, Any] = {
        "status": "none",
        "value": None,
        "effective_date": None,
        "missing_rows": missing_rows,
        "future_rows": future_rows,
        "conflict_rows": [],
    }
    if not dated_rows:
        return result

    latest_date = max(item[0] for item in dated_rows)
    latest_rows = [row for effective_date, row in dated_rows if effective_date == latest_date]
    latest_values: dict[str, Any] = {}
    for row in latest_rows:
        value = row.get(field)
        latest_values.setdefault(_text(value), value)
    result["effective_date"] = latest_date.isoformat()
    if len(latest_values) == 1:
        latest_value = next(iter(latest_values.values()))
        missing_conflicts = [
            row for row in missing_rows if _text(row.get(field)) != _text(latest_value)
        ]
        if not missing_conflicts:
            result.update(status="resolved", value=latest_value)
            return result
        result.update(status="missing_date", conflict_rows=missing_conflicts)
        return result

    result.update(status="conflict", conflict_rows=latest_rows)
    return result


def _valid_employee_id(value: Any) -> bool:
    employee_id = _text(value)
    return (
        bool(employee_id)
        and len(employee_id) <= 30
        and _header(employee_id) not in _EMPLOYEE_ID_HEADERS
        and employee_id.lower() not in {"none", "nan"}
        and "dtype:" not in employee_id.lower()
    )


def discover_people_in_workbook(file_path: str, filename: str) -> list[dict[str, Any]]:
    """兜底扫描工作簿中的姓名列，只用于保证人员不因字段解析失败而丢失。"""
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    discovered: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str]] = set()

    for sheet in workbook.worksheets:
        rows = list(sheet.iter_rows(values_only=True))
        header_rows: list[tuple[
            int,
            int,
            int | None,
            int | None,
            int | None,
            int | None,
            int | None,
            int | None,
        ]] = []
        for row_number, row in enumerate(rows, start=1):
            values = [_header(value) for value in row]
            for name_index, value in enumerate(values, start=1):
                if value not in _NAME_HEADERS:
                    continue

                def nearest(headers: set[str]) -> int | None:
                    matches = [
                        index
                        for index, candidate in enumerate(values, start=1)
                        if candidate in headers and abs(index - name_index) <= 12
                    ]
                    return min(matches, key=lambda index: abs(index - name_index)) if matches else None

                header_rows.append((
                    row_number,
                    name_index,
                    nearest(_EMPLOYEE_ID_HEADERS),
                    nearest(_ID_CARD_HEADERS),
                    nearest(_COMPANY_HEADERS),
                    nearest(_DEPARTMENT_HEADERS),
                    nearest(_ENTRY_DATE_HEADERS),
                    nearest(_MEAL_STANDARD_HEADERS),
                ))

        for (
            header_row,
            name_column,
            employee_column,
            id_card_column,
            company_column,
            department_column,
            entry_date_column,
            meal_standard_column,
        ) in header_rows:
            blank_streak = 0
            for row_number in range(header_row + 1, len(rows) + 1):
                row_values = rows[row_number - 1]

                def value_at(column: int | None) -> Any:
                    if column is None or column > len(row_values):
                        return None
                    return row_values[column - 1]

                name_value = value_at(name_column)
                name = _text(name_value)
                if _header(name_value) in _NAME_HEADERS:
                    continue
                if not name:
                    blank_streak += 1
                    if blank_streak >= 5:
                        break
                    continue
                blank_streak = 0
                if not _valid_name(name):
                    continue

                employee_id = _text(value_at(employee_column))
                id_card = _text(value_at(id_card_column))
                company = _text(value_at(company_column))
                department = _text(value_at(department_column))
                entry_date = _text(value_at(entry_date_column))
                meal_standard = _text(value_at(meal_standard_column))
                key = (name, employee_id.upper(), id_card.upper(), sheet.title)
                if key in seen:
                    continue
                seen.add(key)
                discovered.append({
                    "工号": employee_id or None,
                    "姓名": name,
                    "身份证号": id_card or None,
                    "公司": company or None,
                    "部门": department or None,
                    "入职日期": entry_date or None,
                    "餐补标准": meal_standard or None,
                    "_source_file": filename,
                    "_source_sheet": sheet.title,
                    "_discovery_only": True,
                })

    workbook.close()
    return discovered


def _person_issue(
    issue_type: str,
    name: str,
    message: str,
    source_records: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "issue_type": issue_type,
        "person_name": name,
        "target_field": "人员匹配",
        "message": message,
        "source_files": sorted({_text(row.get("_source_file")) for row in source_records if row.get("_source_file")}),
        "source_sheets": sorted({_text(row.get("_source_sheet")) for row in source_records if row.get("_source_sheet")}),
    }


def _manual_person_group_count(rows: list[dict[str, Any]]) -> int:
    """Conservatively count distinct people represented by a manual-only group."""
    employee_ids = {
        _text(row.get("工号")).upper()
        for row in rows
        if _valid_employee_id(row.get("工号"))
    }
    id_cards = {
        _text(row.get("身份证号")).upper()
        for row in rows
        if _text(row.get("身份证号"))
    }
    return max(1, len(employee_ids), len(id_cards))


def _append_date_resolution_issues(
    issues: list[dict[str, Any]],
    resolution: dict[str, Any],
    name: str,
    field: str,
    source_rows: list[dict[str, Any]],
) -> None:
    if resolution["future_rows"]:
        issues.append({
            **_person_issue(
                "future_effective_date",
                name,
                f"{field}的生效日期晚于本次工资所属月，未纳入本次更新。",
                resolution["future_rows"],
            ),
            "target_field": field,
        })
    if resolution["status"] == "conflict":
        candidate_values = sorted({
            _text(row.get(field))
            for row in resolution["conflict_rows"]
            if _text(row.get(field))
        })
        issues.append({
            **_person_issue(
                "effective_date_conflict",
                name,
                f"{field}在同一最新生效日期存在多个值，目标字段已留空。",
                resolution["conflict_rows"],
            ),
            "target_field": field,
            "effective_date": resolution["effective_date"],
            "candidate_values": candidate_values,
        })
    elif resolution["status"] == "missing_date":
        issues.append({
            **_person_issue(
                "missing_effective_date",
                name,
                f"{field}存在未填写生效日期的冲突记录，未自动覆盖。",
                source_rows,
            ),
            "target_field": field,
        })


def build_complete_roster(
    entities: dict[str, list[dict]],
    salary_month: str | None = None,
) -> dict[str, Any]:
    """把所有实体中的有效人员组成并集，并给关联记录写入内部人员键。

    工号和身份证号是唯一可自动关联的强标识。姓名不能作为关联依据；任何
    仅有姓名或同名但强标识不同的记录均进入待人工处理，避免误写总表。
    """
    cutoff = _salary_month_cutoff(salary_month)
    copied = {
        entity_id: [dict(row) for row in rows if isinstance(row, dict)]
        for entity_id, rows in deepcopy(entities).items()
        if isinstance(rows, list)
    }
    records: list[tuple[str, dict[str, Any]]] = []
    authoritative_profiles = [
        row
        for row in copied.get("employee_profile", [])
        if row.get("_roster_authoritative")
    ]
    admitted_tokens: set[tuple[str, str]] = set()
    unmatched_groups: dict[str, list[dict[str, Any]]] = {}
    identity_conflict_groups: dict[str, list[dict[str, Any]]] = {}
    authoritative_by_name: dict[str, dict[str, set[str]]] = {}
    if authoritative_profiles:
        for row in authoritative_profiles:
            name = _text(row.get("姓名"))
            identity = authoritative_by_name.setdefault(name, {"工号": set(), "身份证号": set()})
            for field in ("工号", "身份证号"):
                value = _text(row.get(field)).upper()
                if field == "工号" and not _valid_employee_id(value):
                    value = ""
                if value:
                    admitted_tokens.add((field, value))
                    identity[field].add(value)
        token_sources: dict[tuple[str, str], set[str]] = {}
        has_confirmation_markers = any(
            row.get("_roster_confirmation")
            for rows in copied.values()
            for row in rows
        )
        for rows in copied.values():
            for row in rows:
                if row.get("_roster_authoritative"):
                    continue
                if has_confirmation_markers and not row.get("_roster_confirmation"):
                    continue
                source_file = _text(row.get("_source_file"))
                if not source_file:
                    continue
                tokens: list[tuple[str, str]] = []
                employee_id = _text(row.get("工号")).upper() if _valid_employee_id(row.get("工号")) else ""
                id_card = _text(row.get("身份证号")).upper()
                if employee_id:
                    tokens.append(("工号", employee_id))
                if id_card:
                    tokens.append(("身份证号", id_card))
                for token in tokens:
                    token_sources.setdefault(token, set()).add(source_file)
        admitted_tokens.update(
            token for token, source_files in token_sources.items() if len(source_files) >= 2
        )

    for entity_id, rows in copied.items():
        for row in rows:
            name = _text(row.get("姓名"))
            employee_id = _text(row.get("工号")).upper() if _valid_employee_id(row.get("工号")) else ""
            id_card = _text(row.get("身份证号")).upper()
            if not (employee_id or id_card):
                if _valid_name(name):
                    unmatched_groups.setdefault(name, []).append(row)
                continue
            if authoritative_profiles:
                master_identity = authoritative_by_name.get(name)
                identity_conflicts = bool(master_identity) and (
                    bool(employee_id and master_identity["工号"] and employee_id not in master_identity["工号"])
                    or bool(id_card and master_identity["身份证号"] and id_card not in master_identity["身份证号"])
                )
                if identity_conflicts and not row.get("_roster_authoritative"):
                    identity_conflict_groups.setdefault(name, []).append(row)
                    continue
                row_tokens = {
                    ("工号", employee_id) if employee_id else None,
                    ("身份证号", id_card) if id_card else None,
                }
                if not any(token in admitted_tokens for token in row_tokens if token):
                    label = name if _valid_name(name) else employee_id or id_card
                    unmatched_groups.setdefault(label, []).append(row)
                    continue
            records.append((entity_id, row))

    parent = list(range(len(records)))

    def find(index: int) -> int:
        while parent[index] != index:
            parent[index] = parent[parent[index]]
            index = parent[index]
        return index

    def union(left: int, right: int) -> None:
        left_root = find(left)
        right_root = find(right)
        if left_root != right_root:
            parent[right_root] = left_root

    for field in ("工号", "身份证号"):
        seen: dict[str, int] = {}
        for index, (_, row) in enumerate(records):
            value = _text(row.get(field)).upper()
            if field == "工号" and not _valid_employee_id(value):
                continue
            if not value:
                continue
            if value in seen:
                union(seen[value], index)
            else:
                seen[value] = index

    issues: list[dict[str, Any]] = [
        _person_issue(
            "unmatched_person",
            name,
            "该人员只出现在一个来源中，无法确认是否属于本月在册；未写入工资总表，请人工核对。",
            rows,
        )
        for name, rows in unmatched_groups.items()
    ]
    issues.extend(
        _person_issue(
            "identity_conflict_with_master",
            name,
            "来源记录与总表同名人员的工号或身份证号冲突，未写入总表，请人工确认是否为新人员或身份变更。",
            rows + [
                row
                for row in authoritative_profiles
                if _text(row.get("姓名")) == name
            ],
        )
        for name, rows in identity_conflict_groups.items()
    )
    by_name: dict[str, list[int]] = {}
    for index, (_, row) in enumerate(records):
        name = _text(row.get("姓名"))
        if _valid_name(name):
            by_name.setdefault(name, []).append(index)

    for name, indexes in by_name.items():
        if len({find(index) for index in indexes}) > 1:
            issues.append(_person_issue(
                "ambiguous_person",
                name,
                "同名记录缺少可互相验证的工号或身份证号，未自动合并；请人工确认身份。",
                [records[index][1] for index in indexes],
            ))

    components: dict[int, list[int]] = {}
    for index in range(len(records)):
        components.setdefault(find(index), []).append(index)

    roster: list[dict[str, Any]] = []
    for component_number, indexes in enumerate(components.values(), start=1):
        component_rows = [records[index][1] for index in indexes]
        profile_rows = [
            records[index][1]
            for index in indexes
            if records[index][0] == "employee_profile"
        ]
        profile: dict[str, Any] = {}
        person_name = _text(next(
            (row.get("姓名") for row in component_rows if _valid_name(row.get("姓名"))),
            "",
        ))
        organizational_fields = {"公司", "部门", "岗位", "本月状态", "离职日期"}
        identity_fields = {"工号", "姓名", "身份证号"}
        profile_fields = {
            field
            for row in profile_rows
            for field in row
            if not field.startswith("_")
            and field not in organizational_fields
            and field not in identity_fields
        }
        for field in profile_fields:
            field_rows = [row for row in profile_rows if row.get(field) not in (None, "")]
            resolution = _dated_field_resolution(field_rows, field, cutoff)
            _append_date_resolution_issues(
                issues, resolution, person_name, field, field_rows
            )
            if resolution["status"] == "resolved":
                profile[field] = resolution["value"]
                continue
            if resolution["status"] in {"conflict", "missing_date"}:
                profile[field] = None
                continue
            values: list[Any] = []
            value_keys: set[str] = set()
            for row in field_rows:
                value = row[field]
                value_key = _text(value)
                if value_key not in value_keys:
                    values.append(value)
                    value_keys.add(value_key)
            authoritative_values = [
                value
                for value in values
                if any(
                    row.get("_roster_authoritative") and _text(row.get(field)) == _text(value)
                    for row in field_rows
                )
            ]
            if len(authoritative_values) == 1:
                profile[field] = authoritative_values[0]
                if any(_text(value) != _text(authoritative_values[0]) for value in values):
                    issues.append({
                        **_person_issue(
                            "conflicting_value",
                            _text(next((row.get("姓名") for row in component_rows if _valid_name(row.get("姓名"))), "")),
                            f"{field}与总表值不一致，已保留总表值，请人工确认。",
                            component_rows,
                        ),
                        "target_field": field,
                        "candidate_values": values,
                    })
            elif len(values) == 1:
                profile[field] = values[0]
            elif len(values) > 1:
                profile[field] = None
                issues.append({
                    **_person_issue(
                        "conflicting_value",
                        _text(next((row.get("姓名") for row in component_rows if _valid_name(row.get("姓名"))), "")),
                        f"{field}在多个来源中不一致，目标字段已留空。",
                        component_rows,
                    ),
                    "target_field": field,
                    "candidate_values": values,
                })

        for field in organizational_fields:
            field_rows = [
                row for row in component_rows if row.get(field) not in (None, "")
            ]
            resolution = _dated_field_resolution(field_rows, field, cutoff)
            _append_date_resolution_issues(
                issues, resolution, person_name, field, field_rows
            )
            if resolution["status"] == "resolved":
                profile[field] = resolution["value"]
                continue
            if resolution["status"] in {"conflict", "missing_date"}:
                profile[field] = None
                continue
            master_values = {
                _text(row.get(field))
                for row in component_rows
                if row.get("_roster_authoritative") and _text(row.get(field))
            }
            change_values = {
                _text(row.get(field))
                for row in component_rows
                if (
                    not row.get("_roster_authoritative")
                    and _text(row.get(field))
                )
            }
            values_conflict = (
                any(
                    not any(_same_department(master, value) for master in master_values)
                    for value in change_values
                )
                if field == "部门"
                else any(value not in master_values for value in change_values)
            )
            if len(master_values) == 1:
                profile[field] = next(iter(master_values))
            elif len(master_values) > 1:
                profile[field] = None
                issues.append({
                    **_person_issue(
                        "conflicting_value",
                        _text(next((row.get("姓名") for row in component_rows if _valid_name(row.get("姓名"))), "")),
                        f"总表中的{field}存在多个值，目标字段已留空。",
                        component_rows,
                    ),
                    "target_field": field,
                    "candidate_values": sorted(master_values),
                })
            if master_values and values_conflict:
                issue_type = "department_transfer" if field == "部门" else "organizational_change"
                issues.append({
                    **_person_issue(
                        issue_type,
                        _text(next((row.get("姓名") for row in component_rows if _valid_name(row.get("姓名"))), "")),
                        f"{field}从总表值“{'、'.join(sorted(master_values))}”变更为“{'、'.join(sorted(change_values))}”，未自动覆盖，请人工确认。",
                        component_rows,
                    ),
                    "target_field": field,
                    "employee_id": _text(profile.get("工号")),
                    "current_value": "、".join(sorted(master_values)),
                    "proposed_value": (
                        next(iter(change_values)) if len(change_values) == 1 else None
                    ),
                    "candidate_values": sorted(change_values),
                })
            elif not master_values and len(change_values) == 1:
                profile[field] = next(iter(change_values))
            elif not master_values and len(change_values) > 1:
                profile[field] = None
                issues.append({
                    **_person_issue(
                        "conflicting_value",
                        _text(next((row.get("姓名") for row in component_rows if _valid_name(row.get("姓名"))), "")),
                        f"{field}在多个来源中不一致，目标字段已留空。",
                        component_rows,
                    ),
                    "target_field": field,
                    "candidate_values": sorted(change_values),
                })

        for field in ("工号", "姓名", "身份证号"):
            authoritative_values = []
            for row in component_rows:
                value = _text(row.get(field))
                if field == "工号" and not _valid_employee_id(value):
                    continue
                if row.get("_roster_authoritative") and value and value not in authoritative_values:
                    authoritative_values.append(value)
            values = []
            for row in component_rows:
                value = _text(row.get(field))
                if field == "工号" and not _valid_employee_id(value):
                    continue
                if value and (field != "姓名" or _valid_name(value)) and value not in values:
                    values.append(value)
            if len(authoritative_values) == 1:
                profile[field] = authoritative_values[0]
            elif len(values) == 1:
                profile[field] = values[0]
            elif len(values) > 1:
                profile[field] = None
                issues.append(_person_issue(
                    "conflicting_identity",
                    _text(next((row.get("姓名") for row in component_rows if _valid_name(row.get("姓名"))), "")),
                    f"同一人员的{field}存在冲突，目标字段已留空。",
                    component_rows,
                ) | {
                    "target_field": field,
                    "candidate_values": values,
                })

        employee_id = _text(profile.get("工号")).upper()
        id_card = _text(profile.get("身份证号")).upper()
        name = _text(profile.get("姓名"))
        person_key = (
            f"employee:{employee_id}"
            if employee_id
            else f"id-card:{id_card}"
            if id_card
            else f"name:{name}:{component_number}"
        )
        profile["_person_key"] = person_key
        profile["_source_files"] = sorted({_text(row.get("_source_file")) for row in component_rows if row.get("_source_file")})
        profile["_source_sheets"] = sorted({_text(row.get("_source_sheet")) for row in component_rows if row.get("_source_sheet")})
        roster.append(profile)
        for index in indexes:
            records[index][1]["_person_key"] = person_key

    copied["employee_profile"] = roster

    for entity_id, rows in list(copied.items()):
        if entity_id == "employee_profile":
            continue
        grouped: dict[str, list[dict[str, Any]]] = {}
        unlinked: list[dict[str, Any]] = []
        for row in rows:
            person_key = _text(row.get("_person_key"))
            if person_key:
                grouped.setdefault(person_key, []).append(row)
            else:
                unlinked.append(row)

        consolidated: list[dict[str, Any]] = []
        for person_key, group in grouped.items():
            merged: dict[str, Any] = {"_person_key": person_key}
            fields = {
                field
                for row in group
                for field in row
                if not field.startswith("_")
            }
            for field in fields:
                field_rows = [row for row in group if row.get(field) not in (None, "")]
                if not field_rows:
                    continue
                person_name = _text(next(
                    (row.get("姓名") for row in group if _valid_name(row.get("姓名"))),
                    "",
                ))
                resolution = _dated_field_resolution(field_rows, field, cutoff)
                _append_date_resolution_issues(
                    issues, resolution, person_name, field, field_rows
                )
                if resolution["status"] == "resolved":
                    merged[field] = resolution["value"]
                    continue
                if resolution["status"] in {"conflict", "missing_date"}:
                    merged[field] = None
                    continue
                values: list[Any] = []
                value_keys: list[str] = []
                for row in field_rows:
                    value = row.get(field)
                    value_key = _text(value)
                    if value not in (None, "") and value_key not in value_keys:
                        values.append(value)
                        value_keys.append(value_key)
                if len(values) == 1:
                    merged[field] = values[0]
                elif len(values) > 1:
                    merged[field] = None
                    issues.append({
                        **_person_issue(
                            "conflicting_value",
                            _text(next((row.get("姓名") for row in group if _valid_name(row.get("姓名"))), "")),
                            f"{field}在多个来源中不一致，目标字段已留空。",
                            group,
                        ),
                        "target_field": field,
                        "candidate_values": values,
                    })
            merged["_source_files"] = sorted({_text(row.get("_source_file")) for row in group if row.get("_source_file")})
            merged["_source_sheets"] = sorted({_text(row.get("_source_sheet")) for row in group if row.get("_source_sheet")})
            consolidated.append(merged)
        copied[entity_id] = consolidated + unlinked

    manual_only_person_count = sum(
        _manual_person_group_count(rows)
        for rows in (*unmatched_groups.values(), *identity_conflict_groups.values())
    )
    input_person_count = len(roster) + manual_only_person_count
    coverage_gap_count = input_person_count - len(roster) - manual_only_person_count

    return {
        "entities": copied,
        "issues": issues,
        "source_person_count": len(components),
        "output_person_count": len(roster),
        "input_person_count": input_person_count,
        "manual_only_person_count": manual_only_person_count,
        "coverage_gap_count": coverage_gap_count,
    }


def combine_entity_sources(
    *sources: dict[str, list[dict]],
    salary_month: str | None = None,
) -> dict[str, Any]:
    """按上传顺序汇集实体，再构建不会漏人的统一人员清单。"""
    combined: dict[str, list[dict]] = {}
    for source in sources:
        for entity_id, rows in source.items():
            if entity_id.startswith("_") or not isinstance(rows, list):
                continue
            combined.setdefault(entity_id, []).extend(
                dict(row) for row in rows if isinstance(row, dict)
            )
    return build_complete_roster(combined, salary_month=salary_month)


def _missing_value_issues(entities: dict[str, list[dict]]) -> list[dict[str, Any]]:
    salary_by_person = {
        _text(row.get("_person_key")): row
        for row in entities.get("salary_detail", [])
        if _text(row.get("_person_key"))
    }
    issues: list[dict[str, Any]] = []
    for profile in entities.get("employee_profile", []):
        person_key = _text(profile.get("_person_key"))
        checks = {
            "工号": profile.get("工号"),
            "姓名": profile.get("姓名"),
            "身份证号": profile.get("身份证号"),
            "月基本薪资": salary_by_person.get(person_key, {}).get("月基本薪资"),
        }
        missing_fields = [field for field, value in checks.items() if value in (None, "")]
        if missing_fields:
            issues.append({
                "issue_type": "missing_value",
                "person_name": _text(profile.get("姓名")),
                "target_field": "、".join(missing_fields),
                "message": f"{'、'.join(missing_fields)}缺失，统一总表对应单元格已留空。",
                "source_files": profile.get("_source_files", []),
                "source_sheets": profile.get("_source_sheets", []),
            })
    return issues


MANUAL_ISSUES_SHEET = "待人工处理"

_ISSUE_DIFFICULTY = {
    "missing_value": "简单",
    "department_transfer": "简单",
    "organizational_change": "一般",
    "conflicting_value": "一般",
    "missing_effective_date": "一般",
    "unmatched_person": "复杂",
    "ambiguous_person": "复杂",
    "identity_conflict_with_master": "复杂",
    "conflicting_identity": "复杂",
    "effective_date_conflict": "复杂",
    "future_effective_date": "一般",
}
_DIFFICULTY_RANK = {"简单": 0, "一般": 1, "复杂": 2}
_ISSUE_TYPE_LABELS = {
    "department_transfer": "转岗",
    "organizational_change": "转岗",
    "missing_value": "数据缺失",
    "conflicting_value": "数据冲突",
    "missing_effective_date": "生效日期缺失",
    "effective_date_conflict": "生效日期冲突",
    "future_effective_date": "未来日期",
    "unmatched_person": "人员匹配",
    "ambiguous_person": "人员匹配",
    "identity_conflict_with_master": "人员匹配",
    "conflicting_identity": "人员匹配",
}


def _normalize_manual_issues(issues: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    suggestions = {
        "简单": "确认后写入总表",
        "一般": "核对候选值和生效日期后处理",
        "复杂": "人工核对原始文件后再处理",
    }
    for issue in issues:
        item = dict(issue)
        difficulty = item.get("difficulty") or _ISSUE_DIFFICULTY.get(
            item.get("issue_type"), "一般"
        )
        item["difficulty"] = difficulty
        item["problem_type"] = item.get("problem_type") or _ISSUE_TYPE_LABELS.get(
            item.get("issue_type"), item.get("issue_type", "其他问题")
        )
        item["suggestion"] = item.get("suggestion") or suggestions[difficulty]
        current_value = item.get("current_value")
        proposed_value = item.get("proposed_value")
        if proposed_value not in (None, "") and current_value not in (None, ""):
            item["recommended_action"] = item.get("recommended_action") or "apply_proposed"
            item["action_options"] = item.get("action_options") or [
                {"value": "apply_proposed", "label": f"采用新值：{proposed_value}"},
                {"value": "keep_current", "label": f"保留原值：{current_value}"},
            ]
            item["copy_text"] = item.get("copy_text") or "\t".join(
                str(value or "")
                for value in (
                    item.get("employee_id"),
                    item.get("person_name"),
                    item.get("target_field"),
                    current_value,
                    proposed_value,
                )
            )
        else:
            item["recommended_action"] = item.get("recommended_action") or "copy_only"
            item["action_options"] = item.get("action_options") or []
            candidate_values = item.get("candidate_values") or []
            if candidate_values:
                item["copy_text"] = item.get("copy_text") or "\t".join(
                    str(value or "")
                    for value in (
                        item.get("employee_id"),
                        item.get("person_name"),
                        item.get("target_field"),
                        "；".join(str(value) for value in candidate_values),
                    )
                )
                normalized.append(item)
                continue
            item["copy_text"] = item.get("copy_text") or "\t".join(
                str(value or "")
                for value in (
                    item.get("employee_id"),
                    item.get("person_name"),
                    item.get("target_field"),
                    proposed_value or current_value,
                )
                if value not in (None, "")
            )
        normalized.append(item)
    normalized.sort(
        key=lambda item: (
            _DIFFICULTY_RANK.get(item["difficulty"], 1),
            str(item.get("problem_type", "")),
            str(item.get("person_name", "")),
            str(item.get("target_field", "")),
        )
    )
    return normalized


_REVIEW_MODIFIED_FILL = PatternFill("solid", fgColor="FFFFC000")
_REVIEW_NEW_FILL = PatternFill("solid", fgColor="FFC6EFCE")


def _review_scalar(value: Any) -> Any:
    if isinstance(value, ArrayFormula):
        return value.text
    return value


def _review_external_formula(value: Any) -> bool:
    scalar = _review_scalar(value)
    return isinstance(scalar, str) and scalar.startswith("=") and "[" in scalar and "]" in scalar


def _review_text(value: Any) -> str:
    value = _review_scalar(value)
    return "" if value is None else str(value).strip()


def _review_headers(sheet: Any) -> tuple[int | None, int | None, int | None, dict[int, str]]:
    id_headers = {"工号", "新工号", "员工工号", "员工编号", "人员编码", "工牌号"}
    name_headers = {"姓名", "人员姓名", "员工姓名", "名字"}
    for row_number in range(1, min(sheet.max_row, 30) + 1):
        columns: dict[int, str] = {}
        id_column = None
        name_column = None
        for column in range(1, sheet.max_column + 1):
            raw = sheet.cell(row_number, column).value
            text = "".join(_review_text(raw).split())
            columns[column] = _review_text(raw)
            if text in id_headers and id_column is None:
                id_column = column
            if text in name_headers and name_column is None:
                name_column = column
        if id_column is not None or name_column is not None:
            return row_number, id_column, name_column, columns
    return None, None, None, {}


def _review_row_key(sheet: Any, row_number: int, id_column: int | None, name_column: int | None) -> str:
    employee_id = _review_text(sheet.cell(row_number, id_column).value) if id_column else ""
    if employee_id and employee_id not in _NON_PERSON_NAMES:
        return f"id:{employee_id.upper()}"
    name = _review_text(sheet.cell(row_number, name_column).value) if name_column else ""
    return f"name:{name}" if _valid_name(name) else ""


def _review_row_map(
    sheet: Any,
    header_row: int | None,
    id_column: int | None,
    name_column: int | None,
) -> dict[str, int]:
    if header_row is None:
        return {}
    rows: dict[str, int] = {}
    for row_number in range(header_row + 1, sheet.max_row + 1):
        key = _review_row_key(sheet, row_number, id_column, name_column)
        if key and key not in rows:
            rows[key] = row_number
    return rows


def create_review_workbook(
    base_path: str,
    final_path: str,
    review_path: str,
    issues: list[dict[str, Any]] | None = None,
) -> dict[str, int]:
    """Create a color-marked copy of the final workbook without changing its values."""
    shutil.copy2(final_path, review_path)
    base = load_workbook(base_path, data_only=False)
    review = load_workbook(review_path, data_only=False)
    event_count = 0
    try:
        for review_sheet in review.worksheets:
            base_sheet = base[review_sheet.title] if review_sheet.title in base.sheetnames else None
            header_row, id_column, name_column, _ = _review_headers(review_sheet)
            base_header_row, base_id_column, base_name_column, _ = _review_headers(base_sheet) if base_sheet else (None, None, None, {})
            final_rows = _review_row_map(review_sheet, header_row, id_column, name_column)
            base_rows = _review_row_map(base_sheet, base_header_row, base_id_column, base_name_column) if base_sheet else {}
            for key, final_row in final_rows.items():
                if key not in base_rows:
                    event_count += 1
                    for column in range(1, review_sheet.max_column + 1):
                        cell = review_sheet.cell(final_row, column)
                        if cell.value in (None, ""):
                            continue
                        cell.fill = _REVIEW_NEW_FILL
                    continue
                base_row = base_rows[key]
                changed_columns: list[int] = []
                for column in range(1, review_sheet.max_column + 1):
                    final_cell = review_sheet.cell(final_row, column)
                    base_cell = base_sheet.cell(base_row, column) if base_sheet else None
                    final_value = _review_scalar(final_cell.value)
                    base_value = _review_scalar(base_cell.value) if base_cell else None
                    if final_value != base_value:
                        if base_cell and _review_external_formula(base_cell.value) and not _review_external_formula(final_cell.value):
                            continue
                        changed_columns.append(column)
                        final_cell.fill = _REVIEW_MODIFIED_FILL
                if changed_columns:
                    event_count += 1
        review.save(review_path)
    finally:
        base.close()
        review.close()
    return {"change_count": event_count, "audit_row_count": 0}


def _populate_manual_issues_sheet(sheet: Any, issues: list[dict[str, Any]]) -> None:
    headers = [
        "序号", "人员姓名", "问题类型", "处理难度", "目标字段", "问题说明",
        "建议操作", "处理状态", "来源文件", "来源工作表", "可选动作", "可复制内容",
    ]
    widths = {
        "A": 8, "B": 14, "C": 20, "D": 12, "E": 24, "F": 48,
        "G": 28, "H": 12, "I": 30, "J": 24, "K": 34, "L": 48,
    }
    normalized_issues = _normalize_manual_issues(issues)
    sheet.append(headers)
    for index, issue in enumerate(normalized_issues, start=1):
        row = [
            index,
            issue.get("person_name", ""),
            issue.get("problem_type", "其他问题"),
            issue.get("difficulty", "一般"),
            issue.get("target_field", ""),
            issue.get("message", ""),
            issue.get("suggestion", ""),
            "待处理",
            "、".join(issue.get("source_files", [])),
            "、".join(issue.get("source_sheets", [])),
            "；".join(option.get("label", "") for option in issue.get("action_options", [])),
            issue.get("copy_text", ""),
        ]
        sheet.append(row)
        estimated_lines = max(
            (len(str(value)) + int(widths[column] * 1.2) - 1) // int(widths[column] * 1.2)
            for column, value in zip(widths, row, strict=True)
        )
        sheet.row_dimensions[index + 1].height = min(180, max(36, estimated_lines * 18 + 8))

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:L{max(1, sheet.max_row)}"
    sheet.sheet_view.showGridLines = False


def write_manual_issues_workbook(output_path: str, issues: list[dict[str, Any]]) -> None:
    """Create a standalone workbook for records requiring manual follow-up."""
    workbook = Workbook()
    try:
        sheet = workbook.active
        sheet.title = MANUAL_ISSUES_SHEET
        _populate_manual_issues_sheet(sheet, issues)
        workbook.save(output_path)
    finally:
        workbook.close()


def remove_manual_issues_sheet(output_path: str) -> None:
    """The payroll workbook must never carry the separate manual-review sheet."""
    workbook = load_workbook(output_path, data_only=False)
    try:
        if MANUAL_ISSUES_SHEET in workbook.sheetnames:
            workbook.remove(workbook[MANUAL_ISSUES_SHEET])
            workbook.save(output_path)
    finally:
        workbook.close()


def _formula_signature(file_path: str) -> dict[tuple[str, str], str]:
    """Return every formula keyed by its worksheet and cell coordinate."""
    workbook = load_workbook(file_path, read_only=True, data_only=False)
    try:
        signature: dict[tuple[str, str], str] = {}
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    value = cell.value
                    formula = value.text if isinstance(value, ArrayFormula) else value
                    if cell.data_type == "f" or (
                        isinstance(formula, str) and formula.startswith("=")
                    ):
                        signature[(sheet.title, cell.coordinate)] = str(formula)
        return signature
    finally:
        workbook.close()


def _verify_unchanged_formula_signature(template_path: str, output_path: str) -> int:
    """Reject an output that drops or rewrites any original master formula."""
    expected = _formula_signature(template_path)
    actual = _formula_signature(output_path)
    missing = sorted(key for key in expected if key not in actual)
    changed = sorted(
        key for key in expected
        if key in actual and expected[key] != actual[key]
    )
    if missing or changed:
        samples = [
            f"{sheet}!{coordinate}"
            for sheet, coordinate in (missing + changed)[:10]
        ]
        raise RuntimeError(
            "公式完整性校验失败："
            f"缺失 {len(missing)} 个、被改写 {len(changed)} 个；"
            f"示例：{'、'.join(samples)}"
        )
    return len(expected)


def _payroll_person_count(file_path: str) -> int:
    """Count the contiguous employee rows in the master payroll worksheet."""
    workbook = load_workbook(file_path, read_only=True, data_only=False)
    try:
        if "工资核算" not in workbook.sheetnames:
            return 0
        payroll = workbook["工资核算"]
        count = 0
        for row_number in range(4, payroll.max_row + 1):
            employee_id = payroll.cell(row=row_number, column=1).value
            employee_name = payroll.cell(row=row_number, column=2).value
            if employee_id in (None, "") and employee_name in (None, ""):
                break
            count += 1
        return count
    finally:
        workbook.close()


def export_unified_workbook(
    template_path: str,
    entities: dict[str, list[dict]],
    output_path: str,
    additional_issues: list[dict[str, Any]] | None = None,
    coverage: dict[str, int] | None = None,
    verify_formulas: bool = True,
) -> dict[str, Any]:
    """生成唯一统一总表，并强制校验来源人员数与输出人员数一致。"""
    from core.entity_engine import export_to_template

    profiles = entities.get("employee_profile", [])
    if profiles and all(_text(profile.get("_person_key")) for profile in profiles):
        prepared_entities = deepcopy(entities)
        result = {
            "entities": prepared_entities,
            "issues": [],
            "source_person_count": len(profiles),
            "output_person_count": len(profiles),
            "input_person_count": int((coverage or {}).get("input_person_count", len(profiles))),
            "manual_only_person_count": int((coverage or {}).get("manual_only_person_count", 0)),
            "coverage_gap_count": int((coverage or {}).get("coverage_gap_count", 0)),
        }
    else:
        result = build_complete_roster(entities)
    issues = (
        list(additional_issues or [])
        + result["issues"]
        + _missing_value_issues(result["entities"])
    )
    issues = _normalize_manual_issues(issues)
    log = export_to_template(template_path, result["entities"], output_path, freeze_formulas=True)
    remove_manual_issues_sheet(output_path)

    template_people = _payroll_person_count(template_path)
    if verify_formulas and template_people == result["output_person_count"]:
        formula_count = _verify_unchanged_formula_signature(template_path, output_path)
        log.append(f"公式完整性校验通过：原有 {formula_count} 个公式全部保留")

    workbook = load_workbook(output_path, read_only=True, data_only=False)
    payroll = workbook["工资核算"]
    exported_people = sum(
        1
        for row in range(4, 4 + result["output_person_count"])
        if payroll.cell(row=row, column=1).value
        or payroll.cell(row=row, column=2).value
        or payroll.cell(row=row, column=76).value
    )
    workbook.close()
    if exported_people != result["output_person_count"]:
        raise RuntimeError(
            f"人员数量校验失败：准备输出 {result['output_person_count']} 人，实际输出 {exported_people} 人"
        )
    if result["coverage_gap_count"] != 0 or (
        result["input_person_count"]
        != result["output_person_count"] + result["manual_only_person_count"]
    ):
        raise RuntimeError(
            "人员覆盖校验失败：输入候选人员未全部进入工资总表或待人工处理清单"
        )

    return {
        **result,
        "issues": issues,
        "issue_count": len(issues),
        "log": log,
    }
