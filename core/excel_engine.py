"""Openpyxl 无损 Excel 渲染引擎：公式平移、插行、删行、历史值冻结。

设计原则：
- 公式平移：正则识别 A1 引用，仅平移相对行号；绝对引用（$N）不动；列号不变。
- 插行：复制上一行样式/数字格式/公式（平移 1 行），再覆盖静态数据。
- 删行：从下往上遍历，避免行号塌陷。
- 冻结：从 data_only 工作簿读取计算值，覆盖写回目标行。
"""
from __future__ import annotations

import re
from copy import copy
from typing import Iterable

import openpyxl
from openpyxl.utils import get_column_letter


class ExcelFormulaEngine:
    """基于 openpyxl 的无损 Excel 操作引擎。"""

    # 匹配 A1 引用（不含 R1C1），支持 $A$1 / $A1 / A$1 / A1 四种形式
    # 形如 SUM(A1:B10) 内的区间也会被拆成两个引用分别处理
    _A1_REF_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})")

    # ------------------------------------------------------------------
    # 1. 公式平移
    # ------------------------------------------------------------------
    @classmethod
    def shift_formula(cls, formula_str: str, row_offset: int) -> str:
        """重写公式中的相对行号，绝对引用（$N）保持不动。

        :param formula_str: 形如 "=Z10+AA10" 或 "Z10+AA10"
        :param row_offset:  行偏移量，可为负
        :return:            平移后的公式字符串

        示例：
            shift_formula("=Z10+AA10", 1)        -> "=Z11+AA11"
            shift_formula("=$Z$1+AA10", 1)       -> "=$Z$1+AA11"
            shift_formula("=SUM(A1:B10)", 1)     -> "=SUM(A2:B11)"
        """
        if formula_str is None:
            return formula_str
        s = str(formula_str)
        if s == "":
            return s
        # 仅处理公式（以 = 开头），普通字符串原样返回
        if not s.startswith("="):
            return s

        def _replace(match: re.Match) -> str:
            col_dollar, col_letters, row_dollar, row_digits = match.groups()
            # 行号绝对引用（$N）不参与平移
            if row_dollar == "$":
                return match.group(0)
            new_row = int(row_digits) + row_offset
            if new_row < 1:
                # 行号越界，保留原值（避免破坏公式）
                return match.group(0)
            return f"{col_dollar}{col_letters}{row_dollar}{new_row}"

        return cls._A1_REF_RE.sub(_replace, s)

    # ------------------------------------------------------------------
    # 2. 插入员工行
    # ------------------------------------------------------------------
    @staticmethod
    def insert_employee_row(ws, template_row_idx: int, employee_data: dict) -> int:
        """在 template_row_idx 下方插入新行，复用模板行样式/公式，覆盖静态数据。

        :param ws: openpyxl Worksheet
        :param template_row_idx: 模板行号（1-based）
        :param employee_data: {列号(1-based)或列字母: 值}
        :return: 新插入的行号（template_row_idx + 1）
        """
        if template_row_idx < 1 or template_row_idx >= ws.max_row:
            raise ValueError(
                f"template_row_idx 必须在 [1, {ws.max_row - 1}] 范围内，"
                f"收到 {template_row_idx}"
            )

        new_row_idx = template_row_idx + 1
        ws.insert_rows(new_row_idx)

        # 复制模板行样式与公式到新行
        for col_idx in range(1, ws.max_column + 1):
            src_cell = ws.cell(row=template_row_idx, column=col_idx)
            # 注意：insert_rows 后，原 template_row_idx+1 及之下的单元格已下移，
            # 因此此时 new_row_idx 处是空白单元格，需要从 src_cell 复制
            dst_cell = ws.cell(row=new_row_idx, column=col_idx)

            # 样式（字体/填充/边框/对齐）
            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.border = copy(src_cell.border)
                dst_cell.alignment = copy(src_cell.alignment)
                dst_cell.protection = copy(src_cell.protection)
            # 数字格式（保留 % 等）
            if src_cell.number_format:
                dst_cell.number_format = src_cell.number_format
            # 行高对齐
            if template_row_idx in ws.row_dimensions:
                src_dim = ws.row_dimensions[template_row_idx]
                if src_dim.height is not None:
                    ws.row_dimensions[new_row_idx].height = src_dim.height

            # 公式：复制并平移 1 行
            src_val = src_cell.value
            if isinstance(src_val, str) and src_val.startswith("="):
                dst_cell.value = ExcelFormulaEngine.shift_formula(src_val, 1)
            # 静态值不复制（避免把模板示例数据带进新行）

        # 覆盖写入新员工静态数据
        for key, value in employee_data.items():
            col_idx = ExcelFormulaEngine._resolve_col_idx(key)
            ws.cell(row=new_row_idx, column=col_idx).value = value

        return new_row_idx

    # ------------------------------------------------------------------
    # 3. 删除离职行
    # ------------------------------------------------------------------
    @staticmethod
    def delete_departed_rows(ws, key_col_idx: int, departed_keys: Iterable) -> int:
        """从下往上遍历，删除 key 列匹配 departed_keys 的行。

        :param ws: openpyxl Worksheet
        :param key_col_idx: 1-based 关键字列号
        :param departed_keys: 离职员工 key 集合（字符串/数值）
        :return: 实际删除的行数
        """
        # 归一化为字符串集合，避免类型不一致漏删
        key_set = {str(k).strip() for k in departed_keys if k is not None}

        # 一次性扫描所有行号，再按从大到小排序删行
        rows_to_delete = []
        for row_idx in range(1, ws.max_row + 1):
            cell_val = ws.cell(row=row_idx, column=key_col_idx).value
            if cell_val is None:
                continue
            if str(cell_val).strip() in key_set:
                rows_to_delete.append(row_idx)

        # 从下往上删，避免行号塌陷
        rows_to_delete.sort(reverse=True)
        for row_idx in rows_to_delete:
            ws.delete_rows(row_idx, 1)

        return len(rows_to_delete)

    # ------------------------------------------------------------------
    # 4. 冻结历史值
    # ------------------------------------------------------------------
    @staticmethod
    def freeze_row_values(target_ws, row_idx: int, data_only_ws) -> int:
        """读取 data_only_ws 计算后的静态数值，覆盖写回 target_ws 对应行。

        :param target_ws: 目标工作表（公式版）
        :param row_idx:   待冻结的行号
        :param data_only_ws: data_only=True 工作簿中的同名工作表（带缓存值）
        :return: 写回的单元格数量
        """
        if row_idx < 1 or row_idx > target_ws.max_row:
            raise ValueError(f"row_idx 越界：{row_idx}")

        max_col = max(target_ws.max_column, data_only_ws.max_column)
        written = 0
        for col_idx in range(1, max_col + 1):
            src_cell = data_only_ws.cell(row=row_idx, column=col_idx)
            cached_val = src_cell.value
            if cached_val is None:
                continue
            target_ws.cell(row=row_idx, column=col_idx).value = cached_val
            written += 1
        return written

    # ------------------------------------------------------------------
    # 工具方法
    # ------------------------------------------------------------------
    @staticmethod
    def _resolve_col_idx(key) -> int:
        """把列号(int)或列字母(str)统一解析为 1-based 列号。"""
        if isinstance(key, int):
            if key < 1:
                raise ValueError(f"列号必须 >=1，收到 {key}")
            return key
        if isinstance(key, str):
            key = key.strip()
            # 纯数字字符串当列号
            if key.isdigit():
                return int(key)
            # 列字母：A->1, AA->27
            from openpyxl.utils import column_index_from_string
            return column_index_from_string(key.upper())
        raise TypeError(f"不支持的列标识类型：{type(key)}")
