"""Private adapter for payroll masters with dual IDs and repeated personnel blocks.

Column positions come from the uploaded workbook. Source standards are retained
in the onboarding block as well as their operational ledgers. No prior employee's
attendance, tax, bank details or other literal data is used as a default.
"""
from __future__ import annotations

from collections import OrderedDict
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
import math
import re
from typing import Any

from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.formula import ArrayFormula

from core._new_hire_sync import (
    _ALIASES, _Layout, _TEXT_FIELDS, _audit, _cell_value, _formula,
    _insert_person, _issue, _normal, _shift_formula, _shift_reference, _text,
)


@dataclass
class _Block:
    sheet: Any
    header: int
    rows: list[int]
    row: int
    insert: bool
    name_col: int
    id_col: int | None
    columns: dict[str, int]
    existing: bool = False


def _headers(sheet: Any, row: int) -> dict[str, int]:
    return {_normal(c.value): c.column for c in sheet[row] if c.value is not None}


def _payroll_header(book: Any) -> tuple[Any, int, dict[str, int]] | None:
    if '工资核算' not in book:
        return None
    sheet = book['工资核算']
    for row in range(1, min(sheet.max_row, 5) + 1):
        headers = _headers(sheet, row)
        if {'人员编码', '姓名', '新工号', '新公司', '月基本薪资', '月绩效标准', '固浮比'} <= headers.keys():
            return sheet, row, headers
    return None


def _structured_people(book: Any) -> list[int]:
    found = _payroll_header(book)
    if found is None:
        return []
    sheet, header, columns = found
    return [r for r in range(header + 1, sheet.max_row + 1)
            if _text(sheet.cell(r, columns['姓名']).value)
            and _text(sheet.cell(r, columns['新工号']).value)
            and not _formula(sheet.cell(r, columns['新工号']))
            and _normal(sheet.cell(r, columns['姓名']).value) != '姓名']


def _make_block(sheet: Any, header: int, start: int, stop: int,
                name_col: int, id_col: int | None) -> _Block:
    rows = [r for r in range(start, stop)
            if _text(sheet.cell(r, name_col).value)
            and not _formula(sheet.cell(r, name_col))]
    row = max(rows, default=start - 1) + 1
    # Empty identities with only pre-existing formulas/zero placeholders are spares.
    spare = row < stop and not _text(sheet.cell(row, name_col).value)
    if spare:
        spare = all(c.value in (None, '', 0) or _formula(c) for c in sheet[row])
    columns = {}
    for cell in sheet[header]:
        field = _ALIASES.get(_normal(cell.value))
        if field:
            columns[field] = cell.column
    return _Block(sheet, header, rows, row, not spare, name_col, id_col, columns)


def _first_end(sheet: Any, start: int, name_col: int) -> int:
    for row in range(start, sheet.max_row + 1):
        cell = sheet.cell(row, name_col)
        if _formula(cell) or _normal(cell.value) in {'姓名', '总经理', '合计', '小计', '总计'}:
            return row
        if not _text(cell.value) and any(
            m[1] == m[3]
            for c in sheet[row] if _formula(c)
            for m in re.finditer(r'\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)', str(c.value))
            if re.search(r'\b(?:SUBTOTAL|SUM)\(', str(c.value), re.I)
        ):
            return row
    return sheet.max_row + 2


def _retarget_ranges(sheet: Any, row: int) -> None:
    """Extend ranges belonging to a table whose footer receives an inserted row."""
    def shifted(value: Any) -> str:
        ranges = []
        for area in MultiCellRange(str(value)).sorted():
            # Whole-column CF and validation ranges already cover future rows.
            if area.max_row == 1048576:
                ranges.append(str(area))
            else:
                ranges.append(_shift_reference(str(area), sheet.title, sheet.title, row))
        return ' '.join(ranges)
    rules = OrderedDict()
    for cf, items in sheet.conditional_formatting._cf_rules.items():
        new_cf = copy(cf)
        new_cf.sqref = shifted(cf.sqref)
        rules[new_cf] = items
    sheet.conditional_formatting._cf_rules = rules
    for validation in sheet.data_validations.dataValidation:
        validation.sqref = shifted(validation.sqref)
    if sheet.auto_filter.ref:
        sheet.auto_filter.ref = shifted(sheet.auto_filter.ref)


def _insert_block(book: Any, block: _Block, record: dict, updates: list[dict]) -> None:
    if not hasattr(book, '_personnel_row_insertions'):
        book._personnel_row_insertions = []
    book._personnel_row_insertions.append({'sheet': block.sheet.title, 'row': block.row})
    _retarget_ranges(block.sheet, block.row)
    # The common inserter handles ordinary formulas. Array expressions and their
    # declared ranges also follow the row move, including cross-sheet references.
    for sheet in book:
        for row in sheet:
            for cell in row:
                if isinstance(cell.value, ArrayFormula):
                    old = cell.value
                    value = ArrayFormula(
                        ref=_shift_reference(old.ref, sheet.title, block.sheet.title, block.row),
                        text=_shift_formula(old.text, sheet.title, block.sheet.title, block.row,
                                            sheet.title != block.sheet.title or cell.row >= block.row),
                    )
                    if value.text != old.text or value.ref != old.ref:
                        _audit(cell, value, record, updates)
    _insert_person(book, _Layout(block.sheet, block.header, {}, block.rows, block.row), record, updates)


def _extend_spare_references(book: Any, block: _Block, record: dict, updates: list[dict]) -> None:
    """A lookup ending at the former last employee must include a filled spare."""
    previous = max(block.rows, default=block.row - 1)
    pattern = re.compile(r'(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?)' + str(previous) + r'$')
    from openpyxl.formula.tokenizer import Tokenizer
    for sheet in book:
        for cells in sheet:
            for cell in cells:
                value = cell.value.text if isinstance(cell.value, ArrayFormula) else cell.value
                if not isinstance(value, str) or not value.startswith('='):
                    continue
                parts = []
                for token in Tokenizer(value).items:
                    text = token.value
                    qualifier, sep, address = text.rpartition('!')
                    belongs = qualifier.strip("'").replace("''", "'") == block.sheet.title if sep else sheet == block.sheet
                    if belongs and token.type == 'OPERAND' and token.subtype == 'RANGE':
                        area = address if sep else text
                        match = pattern.fullmatch(area)
                        # A one-row horizontal expression must not become a vertical range.
                        if match and int(re.search(r'\d+', area).group()) < previous:
                            area = match[1] + str(block.row)
                            text = qualifier + '!' + area if sep else area
                    parts.append(text)
                result = '=' + ''.join(parts)
                if result != value:
                    if isinstance(cell.value, ArrayFormula):
                        result = ArrayFormula(ref=cell.value.ref, text=result)
                    _audit(cell, result, record, updates)


def _add_column(block: _Block, label: str, record: dict, updates: list[dict]) -> int:
    headers = _headers(block.sheet, block.header)
    if _normal(label) in headers:
        return headers[_normal(label)]
    # Do not overwrite notes or values elsewhere in an existing column.
    column = max((cell.column for cells in block.sheet for cell in cells if cell.value is not None), default=0) + 1
    cell = block.sheet.cell(block.header, column)
    cell._style = copy(block.sheet.cell(block.header, max(1, column - 1))._style)
    _audit(cell, label, record, updates)
    block.sheet.column_dimensions[get_column_letter(column)].width = max(16, len(label) * 2 + 2)
    return column


def _apply_structured_person(book: Any, record: dict, field_sources: dict, policy_action: Any,
                             salary_month: str | None, updates: list[dict], matches: list[dict]) -> tuple[bool, list[dict], set[str]]:
    found = _payroll_header(book)
    if found is None:
        return False, [], set()
    pay, header, headers = found
    values = record['values']
    name, employee_id = _text(values['姓名']), _text(values['工号'])
    people = _structured_people(book)
    if not people:
        return True, [_issue('hire_unsupported_layout', '双工号总表没有可核验的既有人员结构。', record)], set()
    by_name = { _text(pay.cell(r, headers['姓名']).value): r for r in people }
    company = values.get('公司')
    company_rows = [r for r in people if _text(pay.cell(r, headers['新公司']).value) == _text(company)]
    short_companies = {_text(pay.cell(r, headers['所属公司']).value) for r in company_rows} if '所属公司' in headers else set()
    if len(short_companies) != 1:
        return True, [_issue('hire_company_conflict', '无法从总表确定唯一的公司名称对应关系。', record)], set()
    short_company = next(iter(short_companies))
    department_rows = [r for r in company_rows if _text(pay.cell(r, headers.get('新成本中心/部门', 1)).value) == _text(values.get('部门'))]
    old_departments = {_text(pay.cell(r, headers['成本中心/部门']).value) for r in department_rows} if '成本中心/部门' in headers else set()
    old_department = next(iter(old_departments)) if len(old_departments) == 1 else values.get('部门')
    old_row = by_name.get(name)
    code = _cell_value(pay.cell(old_row, headers['人员编码']), '工号') if old_row else employee_id
    valid_ids = {employee_id, _text(code)}
    blocks: dict[str, _Block] = {}
    blocks['payroll'] = _make_block(pay, header, min(people), _first_end(pay, max(people) + 1, headers['姓名']), headers['姓名'], headers['新工号'])
    # Detect section boundaries before choosing a spare. Historical event blocks
    # are separate tables and must not receive a copy of this onboarding notice.
    if '人员异动' in book:
        sheet = book['人员异动']
        hire_headers = [r + 1 for r in range(1, min(10, sheet.max_row) + 1)
                        if any(_normal(c.value).startswith('入职情况') for c in sheet[r])]
        if len(hire_headers) == 1:
            h = hire_headers[0]
            end = next((r for r in range(h + 1, sheet.max_row + 1)
                        if any(_normal(c.value).endswith('情况') and _text(c.value) for c in sheet[r])), sheet.max_row + 2)
            blocks['hire'] = _make_block(sheet, h, h + 1, end, 2, 3)
    if 'hire' not in blocks:
        return True, [_issue('hire_unsupported_layout', '未找到唯一的入职情况分区。', record)], set()
    for role, title, h in [('attendance', '考勤', 1), ('performance', '绩效考核', 1), ('social', '台账', 2), ('heat', '防暑降温费', 1)]:
        if title in book:
            sheet = book[title]
            columns = _headers(sheet, h)
            if '姓名' not in columns:
                return True, [_issue('hire_unsupported_layout', '相关明细表头不完整。', record, target_sheet=title)], set()
            blocks[role] = _make_block(sheet, h, h + 1, _first_end(sheet, h + 1, columns['姓名']), columns['姓名'], 1)
    for role, title in [('tax', '个税申报'), ('tax_check', '个税导出核对')]:
        if title not in book:
            continue
        sheet = book[title]
        ends = [r for r in range(2, sheet.max_row + 1) if _formula(sheet.cell(r, 2)) and 'SUBTOTAL(' in str(sheet.cell(r, 2).value).upper()]
        begin, candidates = 2, []
        for end in ends:
            companies = {_text(pay.cell(by_name[_text(sheet.cell(r, 2).value)], headers['新公司']).value)
                         for r in range(begin, end) if _text(sheet.cell(r, 2).value) in by_name}
            if companies == {_text(company)}:
                candidates.append(_make_block(sheet, 1, begin, end, 2, 1))
            begin = end + 1
        if len(candidates) != 1:
            return True, [_issue('hire_unsupported_layout', '个税公司分区无法唯一确定。', record, target_sheet=title)], set()
        blocks[role] = candidates[0]
    bank_blocks = []
    for sheet in book:
        for h in range(1, min(sheet.max_row, 8) + 1):
            cols = _headers(sheet, h)
            if not {'收款户名', '收款账号', '金额'} <= cols.keys():
                continue
            block = _make_block(sheet, h, h + 1, sheet.max_row + 2, cols['收款户名'], None)
            companies = {_text(pay.cell(by_name[_text(sheet.cell(r, block.name_col).value)], headers['新公司']).value)
                         for r in block.rows if _text(sheet.cell(r, block.name_col).value) in by_name}
            if companies == {_text(company)}:
                bank_blocks.append(block)
    if len(bank_blocks) > 1:
        return True, [_issue('hire_unsupported_layout', '公司对应多个报盘表，无法唯一定位。', record)], set()
    if bank_blocks:
        blocks['bank'] = bank_blocks[0]
    errors = []
    for field in values:
        if policy_action(record['source_sheet'], field) in {'ignore', 'review'}:
            errors.append(_issue('configured_field_review', '匹配策略要求暂停该入职字段。', record, target_field=field))
    for block in blocks.values():
        hits = []
        for row in block.rows:
            old_name = _text(block.sheet.cell(row, block.name_col).value)
            old_id = _text(_cell_value(block.sheet.cell(row, block.id_col), '工号')) if block.id_col else ''
            if old_name == name or old_id in valid_ids:
                if old_name != name or (old_id and old_id not in valid_ids):
                    errors.append(_issue('hire_identity_conflict', '相关表页的工号与姓名冲突。', record, target_sheet=block.sheet.title))
                hits.append(row)
        if len(hits) > 1:
            errors.append(_issue('hire_identity_conflict', '相关表页存在重复人员。', record, target_sheet=block.sheet.title))
        if len(hits) == 1:
            block.row, block.insert, block.existing = hits[0], False, True
        if block.insert and (block.sheet.tables or block.sheet._charts or block.sheet._images
                             or any(m.min_row <= block.row <= m.max_row for m in block.sheet.merged_cells.ranges)):
            errors.append(_issue('hire_unsupported_layout', '新增位置涉及无法安全移动的表格、图形或合并区域。', record, target_sheet=block.sheet.title))
    if errors:
        return True, errors, set()
    # All identity and structure checks finish before the first write.
    for block in blocks.values():
        if block.existing:
            continue
        if block.insert:
            _insert_block(book, block, record, updates)
        else:
            _extend_spare_references(book, block, record, updates)
        if block.rows:
            for cell in block.sheet[max(block.rows)]:
                block.sheet.cell(block.row, cell.column)._style = copy(cell._style)
            block.sheet.row_dimensions[block.row].height = block.sheet.row_dimensions[max(block.rows)].height

    def put(block: _Block, column: int, value: Any, field: str | None = None) -> None:
        cell = block.sheet.cell(block.row, column)
        _audit(cell, value, field_sources.get(field, record), updates)
        if field in _TEXT_FIELDS or column == block.id_col:
            cell.number_format = '@'
        elif isinstance(value, (date, datetime)):
            cell.number_format = 'yyyy-mm-dd'
        if not block.existing and isinstance(value, str) and not value.startswith('=') and len(value) > 7:
            alignment = copy(cell.alignment)
            alignment.wrapText = True
            cell.alignment = alignment
            width = block.sheet.column_dimensions[get_column_letter(column)].width or 13
            units = sum(2 if ord(char) > 255 else 1 for char in value)
            lines = max(1, math.ceil(units / max(1, width - 1)))
            height = (cell.font.sz or 11) * 1.5 * lines + 6
            block.sheet.row_dimensions[block.row].height = max(block.sheet.row_dimensions[block.row].height or 15, height)

    for role, block in blocks.items():
        put(block, block.name_col, name, '姓名')
        if block.id_col:
            put(block, block.id_col, employee_id if role == 'payroll' else code, '工号')
        for field, column in block.columns.items():
            if field not in values or field in {'工号', '姓名', '公司', '部门', '岗位', '月基本薪资', '月绩效标准'}:
                continue
            if not _formula(block.sheet.cell(block.row, column)):
                put(block, column, values[field], field)
        matches.append({'source_file': record['source_file'], 'source_sheet': record['source_sheet'],
                        'target_sheet': block.sheet.title, 'confidence': 1.0,
                        'match_basis': 'explicit_hire_dual_id_and_company'})
    payroll = blocks['payroll']
    for label, value, field in [('人员编码', code, '工号'), ('新公司', company, '公司'), ('所属公司', short_company, '公司'),
                                ('成本中心/部门', old_department, '部门'), ('新成本中心/部门', values.get('部门'), '部门'),
                                ('岗位', values.get('岗位'), '岗位'), ('新岗位', values.get('岗位'), '岗位')]:
        if label in headers and value is not None:
            put(payroll, headers[label], value, field)
    hire = blocks['hire']
    for col, value, field in [(4, company, '公司'), (5, values.get('部门'), '部门'), (6, values.get('岗位'), '岗位'),
                               (8, values.get('入职日期'), '入职日期'), (9, values.get('转正日期'), '转正日期')]:
        if value is not None:
            put(hire, col, value, field)
    if not hire.existing:
        put(hire, 1, max([int(hire.sheet.cell(r, 1).value) for r in hire.rows if isinstance(hire.sheet.cell(r, 1).value, (int, float))], default=0) + 1)
    # Retain every source standard explicitly, so formula-backed payroll columns
    # and short company/dept aliases never lose the original customer information.
    direct_hire = {'姓名': 2, '工号': 3, '公司': 4, '部门': 5, '岗位': 6, '入职日期': 8, '转正日期': 9}
    for field, value in values.items():
        column = direct_hire.get(field) or _add_column(hire, field, record, updates)
        put(hire, column, value, field)
    for role in ('attendance', 'performance', 'heat'):
        if role in blocks:
            block = blocks[role]
            for field, value in [('部门', old_department), ('岗位', values.get('岗位')), ('公司', short_company)]:
                if field in block.columns and value is not None:
                    put(block, block.columns[field], value, field)
    if 'social' in blocks:
        social = blocks['social']
        put(social, 3, short_company, '公司')
        for field, label in [('单位申报基数', '社保基数'), ('公积金基数', '公积金基数')]:
            if field in values:
                put(social, _add_column(social, label, record, updates), values[field], field)
    if {'月基本薪资', '月绩效标准'} <= values.keys():
        total_header = next((h for h in headers if h.startswith('月度合计')), None)
        if total_header:
            total_col, ratio_col = headers[total_header], headers['固浮比']
            base, performance = values['月基本薪资'], values['月绩效标准']
            put(payroll, total_col, f'={base:g}+{performance:g}')
            total_ref = payroll.sheet.cell(payroll.row, total_col).coordinate
            put(payroll, ratio_col, f'=IF({total_ref}=0,0,{base:g}/{total_ref})')
    # Seed formulas only once. A missing monthly input must stay missing; it must
    # not produce a plausible full-month salary simply because a new row exists.
    if not payroll.existing:
        template_row = company_rows[0]
        for src in payroll.sheet[template_row]:
            if not _formula(src):
                continue
            dst = payroll.sheet.cell(payroll.row, src.column)
            text = src.value.text if isinstance(src.value, ArrayFormula) else src.value
            formula = Translator(text, origin=src.coordinate).translate_formula(dst.coordinate)
            if src.column not in {headers['月基本薪资'], headers['月绩效标准']}:
                # Release calculations automatically only when the actual monthly
                # attendance, personal contributions and tax have numeric values.
                required = []
                for role, predicate in (
                    ('attendance', lambda h: h == '出勤天数'),
                    ('performance', lambda h: h == '考核成绩'),
                    ('social', lambda h: h.startswith('个人')),
                    ('tax_check', lambda h: '应补' in h and '税' in h),
                ):
                    if role in blocks:
                        b = blocks[role]
                        required.extend(f"'{b.sheet.title}'!${get_column_letter(c)}${b.row}"
                                        for h, c in _headers(b.sheet, b.header).items() if predicate(h))
                if required:
                    formula = f'=IF(COUNT({",".join(required)})<{len(required)},"",{formula[1:]})'
                else:
                    formula = f'=IF(1,"",{formula[1:]})'
            put(payroll, src.column, ArrayFormula(ref=dst.coordinate, text=formula) if isinstance(src.value, ArrayFormula) else formula)
        for field in ('月基本薪资', '月绩效标准'):
            if field not in values:
                continue
            cell = payroll.sheet.cell(payroll.row, headers[field])
            # Some templates have literal overrides, or round to tens. The
            # customer's exact standard must still survive those variations.
            rounds_tens = isinstance(cell.value, str) and re.search(r',\s*-1\s*\)', cell.value)
            if not _formula(cell) or (rounds_tens and values[field] % 10 != 0):
                source_col = _headers(hire.sheet, hire.header)[_normal(field)]
                put(payroll, headers[field], f"='{hire.sheet.title}'!{get_column_letter(source_col)}{hire.row}", field)
    if 'heat' in blocks:
        heat = blocks['heat']
        if '高温费备注' in values:
            put(heat, _add_column(heat, '高温费备注', record, updates), values['高温费备注'], '高温费备注')
        # An annual standard is not a current-month payment. Do not carry a
        # neighbouring employee's payment formula or zero into a new hire row.
        if not heat.existing:
            for column in (4, 5, 7, 8):
                put(heat, column, None)
    # Dependent transaction tables receive identity; missing bill/tax/bank inputs
    # remain empty instead of copying someone else's data or payable amount.
    for role in ('tax', 'tax_check', 'bank', 'social'):
        if role not in blocks or blocks[role].existing:
            continue
        block = blocks[role]
        for cell in block.sheet[block.row]:
            if _formula(cell):
                put(block, cell.column, None)
    return True, [], set(values)
