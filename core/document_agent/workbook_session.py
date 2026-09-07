"""Small physical workbook tools; business decisions stay with the Agent.

No scripts, arbitrary formulas, paths, or model-invented amounts are accepted.
Each batch is checked completely before an atomic replacement of the draft.
"""
from __future__ import annotations

from copy import copy
import math
import os
import re
from pathlib import Path
import shutil
import tempfile
from typing import Any, Literal
from zipfile import BadZipFile

import openpyxl
from openpyxl.formula.translate import Translator
from pydantic import BaseModel, ConfigDict, Field, ValidationError

from .orchestrator import ToolExecutionError


class SourceCellChange(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)
    source_file: str = Field(min_length=1)
    source_sheet: str = Field(min_length=1)
    source_cells: list[str] = Field(min_length=1, max_length=100)
    target_sheet: str = Field(min_length=1)
    target_cell: str = Field(pattern=r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")
    expected_value: str | int | float | bool | None
    aggregation: Literal["copy", "sum"]


class FormulaDivisorChange(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    target_sheet: str = Field(min_length=1)
    target_cell: str = Field(pattern=r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")
    expected_formula: str = Field(min_length=2)
    old_divisor: int = Field(gt=0)
    new_divisor: int = Field(gt=0)


class RowCopyChange(BaseModel):
    """A guarded insert that copies one existing row within a worksheet."""

    model_config = ConfigDict(extra="forbid", strict=True)

    sheet: str = Field(min_length=1)
    source_row: int = Field(ge=1, le=1_048_576)
    insert_before_row: int = Field(ge=1, le=1_048_576)
    expected_source_cells: dict[str, str | int | float | bool | None] = Field(min_length=1, max_length=20)
    copy_max_column: int = Field(ge=1, le=16_384)
    # Required when the insertion point is inside a personnel table. Keys are
    # destination coordinates on the new row (for example A12/B12).
    identity_values: dict[str, str | int] = Field(default_factory=dict, max_length=10)


class RowDeleteChange(BaseModel):
    """A guarded deletion of already-read rows within a worksheet."""

    model_config = ConfigDict(extra="forbid", strict=True)

    sheet: str = Field(min_length=1)
    rows: list[int] = Field(min_length=1, max_length=200)
    # 每个待删行至少给出一个锚点单元格及其当前值，证明该行确实读过且
    # 未发生变化；锚点坐标必须位于对应待删行上。
    expected_cells: dict[str, str | int | float | bool | None] = Field(min_length=1, max_length=400)


class VerifiedCellValueChange(BaseModel):
    """A deterministic, stale-protected non-formula cell update."""

    model_config = ConfigDict(extra="forbid", strict=True)

    target_sheet: str = Field(min_length=1)
    target_cell: str = Field(pattern=r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")
    expected_value: str | int | float | bool | None
    new_value: str | int | float | bool | None


class SummaryRowRollForward(BaseModel):
    """Insert a new current-month summary row and freeze the previous one."""

    model_config = ConfigDict(extra="forbid", strict=True)

    sheet: str = Field(min_length=1)
    current_row: int = Field(ge=1, le=1_048_575)
    expected_current_month: str | int | float | bool | None
    new_month: str = Field(min_length=1)
    start_column: int = Field(ge=1, le=16_384)
    end_column: int = Field(ge=1, le=16_384)


class WorkbookSession:
    def __init__(self, master: Path, draft: Path, sources: dict[str, Path]):
        self.master = master.resolve()
        self.draft = draft.resolve()
        self.sources = {name: path.resolve() for name, path in sources.items()}
        if self.draft in {self.master, *self.sources.values()}:
            raise ToolExecutionError("草稿必须与原件和来源文件分离")

    def prepare_copy(self) -> dict[str, Any]:
        try:
            self.draft.parent.mkdir(parents=True, exist_ok=True)
            if self.draft.exists():
                return {"created": False, "filename": self.draft.name}
            # Exclusive creation avoids resetting work if a call is repeated.
            with self.master.open("rb") as source, self.draft.open("xb") as target:
                shutil.copyfileobj(source, target)
            return {"created": True, "filename": self.draft.name}
        except OSError as exc:
            raise ToolExecutionError("无法创建独立副本，请检查原件和输出目录") from exc

    def apply_source_cells(self, changes: list[dict[str, Any]]) -> list[dict[str, Any]]:
        # 单批上限500格：整批校验通过后原子写入，支持“一次读全后
        # 少量大批写入”的执行模式（23人×20字段级别一批写完）。
        if not 1 <= len(changes) <= 500:
            raise ToolExecutionError("每批必须包含1至500项变更")
        try:
            parsed = [SourceCellChange.model_validate(change) for change in changes]
        except ValidationError as exc:
            raise ToolExecutionError("来源写入参数不合法") from exc
        if not self.draft.is_file():
            raise ToolExecutionError("请先创建原始总表的独立副本")
        try:
            workbook = openpyxl.load_workbook(self.draft, data_only=False)
        except (OSError, ValueError, BadZipFile) as exc:
            raise ToolExecutionError("草稿工作簿不存在或无法读取") from exc
        source_books: dict[tuple[str, bool], Any] = {}
        updates: list[dict[str, Any]] = []
        targets: set[tuple[str, str]] = set()
        try:
            for change in parsed:
                if change.source_file not in self.sources:
                    raise ToolExecutionError("来源文件不在当前任务白名单内")
                target_key = (change.target_sheet, change.target_cell)
                if target_key in targets:
                    raise ToolExecutionError("同一批次不能重复写入同一单元格")
                targets.add(target_key)
                if change.target_sheet not in workbook:
                    raise ToolExecutionError("目标工作表不存在")
                target = workbook[change.target_sheet][change.target_cell]
                if target.data_type == "f":
                    raise ToolExecutionError("不能用来源数值覆盖总表公式")
                if target.value != change.expected_value:
                    raise ToolExecutionError("目标值已变化，请重新读取后再决定")
                values = []
                for cached in (False, True):
                    key = (change.source_file, cached)
                    if key not in source_books:
                        source_books[key] = openpyxl.load_workbook(self.sources[change.source_file], read_only=True, data_only=cached)
                formula_book = source_books[(change.source_file, False)]
                cached_book = source_books[(change.source_file, True)]
                if change.source_sheet not in formula_book:
                    raise ToolExecutionError("来源工作表不存在")
                for coordinate in change.source_cells:
                    if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}", coordinate):
                        raise ToolExecutionError("来源必须是单一单元格坐标")
                    source_cell = formula_book[change.source_sheet][coordinate]
                    value = cached_book[change.source_sheet][coordinate].value
                    if source_cell.data_type == "f" and value is None:
                        raise ToolExecutionError("来源公式没有已计算缓存，不能猜测金额")
                    if source_cell.data_type == "e" or isinstance(value, str) and value.startswith("="):
                        raise ToolExecutionError("来源含错误值或公式文本，不能写入")
                    if value is not None and not isinstance(value, (str, int, float, bool)):
                        raise ToolExecutionError("来源值不是支持的标量类型")
                    if isinstance(value, float) and not math.isfinite(value):
                        raise ToolExecutionError("来源金额不是有限数值")
                    values.append(value)
                if change.aggregation == "sum":
                    if any(isinstance(value, bool) or not isinstance(value, (int, float)) for value in values):
                        raise ToolExecutionError("求和来源必须全部是明确数值，空白不能默认为零")
                    new_value = sum(values)
                else:
                    if len(values) != 1:
                        raise ToolExecutionError("复制操作必须只有一个来源单元格")
                    new_value = values[0]
                updates.append({
                    "target_sheet": change.target_sheet, "target_cell": change.target_cell,
                    "old_value": target.value, "new_value": new_value,
                    "source_file": change.source_file, "source_sheet": change.source_sheet,
                    "source_cells": change.source_cells, "aggregation": change.aggregation,
                })
            for update in updates:
                workbook[update["target_sheet"]][update["target_cell"]] = update["new_value"]
            descriptor, temporary = tempfile.mkstemp(prefix="agent-write-", suffix=".xlsx", dir=self.draft.parent)
            os.close(descriptor)
            temporary_path = Path(temporary)
            try:
                workbook.save(temporary_path)
                os.replace(temporary_path, self.draft)
            finally:
                temporary_path.unlink(missing_ok=True)
            return updates
        except (KeyError, ValueError, AttributeError, OSError, BadZipFile) as exc:
            raise ToolExecutionError("单元格坐标不可写或超出有效范围，本批未写入") from exc
        finally:
            workbook.close()
            for source_book in source_books.values():
                source_book.close()

    @staticmethod
    def _replace_formula_divisor(formula: str, old_divisor: int, new_divisor: int) -> str:
        """Replace exact numeric divisors while preserving quoted Excel strings."""
        old_value, new_value = str(old_divisor), str(new_divisor)
        result: list[str] = []
        index = 0
        in_string = False
        replacements = 0
        while index < len(formula):
            character = formula[index]
            if character == '"':
                result.append(character)
                if in_string and index + 1 < len(formula) and formula[index + 1] == '"':
                    result.append('"')
                    index += 2
                    continue
                in_string = not in_string
                index += 1
                continue
            if not in_string and character == "/":
                cursor = index + 1
                while cursor < len(formula) and formula[cursor].isspace():
                    cursor += 1
                if formula[cursor:cursor + len(old_value)] == old_value:
                    end = cursor + len(old_value)
                    if end == len(formula) or not (formula[end].isdigit() or formula[end] == "."):
                        result.extend((formula[index:cursor], new_value))
                        replacements += 1
                        index = end
                        continue
            result.append(character)
            index += 1
        if not replacements:
            raise ToolExecutionError("公式中没有可替换的指定分母")
        return "".join(result)

    def apply_formula_divisors(self, changes: list[dict[str, Any]]) -> list[dict[str, Any]]:
        if not 1 <= len(changes) <= 500:
            raise ToolExecutionError("每批必须包含1至500项公式参数变更")
        try:
            parsed = [FormulaDivisorChange.model_validate(change) for change in changes]
        except ValidationError as exc:
            raise ToolExecutionError("公式参数变更不合法") from exc
        if not self.draft.is_file():
            raise ToolExecutionError("请先创建原始总表的独立副本")
        try:
            workbook = openpyxl.load_workbook(self.draft, data_only=False)
        except (OSError, ValueError, BadZipFile) as exc:
            raise ToolExecutionError("草稿工作簿不存在或无法读取") from exc
        updates: list[dict[str, Any]] = []
        targets: set[tuple[str, str]] = set()
        try:
            for change in parsed:
                target_key = (change.target_sheet, change.target_cell)
                if target_key in targets:
                    raise ToolExecutionError("同一批次不能重复修改同一公式")
                targets.add(target_key)
                if change.target_sheet not in workbook.sheetnames:
                    raise ToolExecutionError("目标工作表不存在")
                cell = workbook[change.target_sheet][change.target_cell]
                formula = cell.value
                if not isinstance(formula, str) or not formula.startswith("=") or cell.data_type != "f":
                    raise ToolExecutionError("目标单元格不是可修改的公式")
                if formula != change.expected_formula:
                    raise ToolExecutionError("目标公式已变化，请重新读取后再执行")
                replacement = self._replace_formula_divisor(formula, change.old_divisor, change.new_divisor)
                updates.append({
                    "target_sheet": change.target_sheet,
                    "target_cell": change.target_cell,
                    "old_value": formula,
                    "new_value": replacement,
                    "old_divisor": change.old_divisor,
                    "new_divisor": change.new_divisor,
                })
            for update in updates:
                workbook[update["target_sheet"]][update["target_cell"]] = update["new_value"]
            descriptor, temporary = tempfile.mkstemp(prefix="agent-formula-", suffix=".xlsx", dir=self.draft.parent)
            os.close(descriptor)
            temporary_path = Path(temporary)
            try:
                workbook.save(temporary_path)
                os.replace(temporary_path, self.draft)
            finally:
                temporary_path.unlink(missing_ok=True)
            return updates
        except (KeyError, ValueError, AttributeError, OSError, BadZipFile) as exc:
            raise ToolExecutionError("公式参数变更失败，本批未写入") from exc
        finally:
            workbook.close()

    @staticmethod
    def _shift_cell_reference(reference: str, insert_before_row: int) -> str:
        match = re.fullmatch(r"(\$?[A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})", reference)
        if not match:
            return reference
        column, row_marker, row_text = match.groups()
        row = int(row_text)
        if row < insert_before_row:
            return reference
        return f"{column}{row_marker}{row + 1}"

    @staticmethod
    def _sheet_name_from_reference(value: str) -> str:
        name = value.strip()
        if name.startswith("'") and name.endswith("'"):
            name = name[1:-1].replace("''", "'")
        return name

    @classmethod
    def _shift_range_for_row_insert(
        cls, first: str, last: str, insert_before_row: int, *, expand_adjacent: bool,
    ) -> tuple[str, str]:
        """Shift a range and expand one ending immediately above the insert."""
        first_match = re.fullmatch(r"(\$?[A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})", first)
        last_match = re.fullmatch(r"(\$?[A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})", last)
        shifted_first = cls._shift_cell_reference(first, insert_before_row)
        shifted_last = cls._shift_cell_reference(last, insert_before_row)
        if first_match and last_match:
            first_row = int(first_match.group(3))
            last_row = int(last_match.group(3))
            if expand_adjacent and first_row <= last_row and last_row == insert_before_row - 1:
                shifted_last = f"{last_match.group(1)}{last_match.group(2)}{last_row + 1}"
        return shifted_first, shifted_last

    @classmethod
    def _shift_formula_references_for_row_insert(
        cls,
        formula: str,
        *,
        current_sheet: str,
        target_sheet: str,
        insert_before_row: int,
        expand_adjacent_ranges: bool = True,
    ) -> str:
        """Repair references left stale by openpyxl's row insertion."""
        if not isinstance(formula, str) or not formula.startswith("="):
            return formula
        explicit = re.compile(
            r"(?P<sheet>'[^']+'|[^=!'+\-*/(),\s]+)!"
            r"(?P<first>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})"
            r"(?:[:](?P<last>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}))?"
        )
        bare_range = re.compile(
            r"(?<![A-Za-z0-9_!'])"
            r"(?P<first>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})"
            r":(?P<last>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})"
            r"(?![A-Za-z0-9_])(?!(?:\s*)\()"
        )
        bare = re.compile(
            r"(?<![A-Za-z0-9_!'])\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}"
            r"(?![A-Za-z0-9_])(?!(?:\s*)\()"
        )
        protected: list[str] = []

        def rewrite_segment(segment: str) -> str:
            def replace_explicit(match: re.Match[str]) -> str:
                if cls._sheet_name_from_reference(match.group("sheet")) != target_sheet:
                    return match.group(0)
                first = cls._shift_cell_reference(match.group("first"), insert_before_row)
                last = match.group("last")
                if last:
                    first, last = cls._shift_range_for_row_insert(
                        match.group("first"), last, insert_before_row,
                        expand_adjacent=expand_adjacent_ranges,
                    )
                replacement = f"{first}:{last}" if last else first
                marker = f"\x00{len(protected)}\x00"
                protected.append(f"{match.group('sheet')}!{replacement}")
                return marker

            rewritten = explicit.sub(replace_explicit, segment)
            if current_sheet == target_sheet:
                def replace_bare_range(match: re.Match[str]) -> str:
                    first, last = cls._shift_range_for_row_insert(
                        match.group("first"), match.group("last"), insert_before_row,
                        expand_adjacent=expand_adjacent_ranges,
                    )
                    marker = f"\x00{len(protected)}\x00"
                    protected.append(f"{first}:{last}")
                    return marker

                rewritten = bare_range.sub(replace_bare_range, rewritten)
                rewritten = bare.sub(
                    lambda match: cls._shift_cell_reference(match.group(0), insert_before_row),
                    rewritten,
                )
            for index, value in enumerate(protected):
                rewritten = rewritten.replace(f"\x00{index}\x00", value)
            return rewritten

        parts: list[str] = []
        start = 0
        in_string = False
        index = 0
        while index < len(formula):
            if formula[index] == '"':
                if in_string and index + 1 < len(formula) and formula[index + 1] == '"':
                    index += 2
                    continue
                if not in_string:
                    parts.append(rewrite_segment(formula[start:index]))
                    start = index
                else:
                    parts.append(formula[start:index + 1])
                    start = index + 1
                in_string = not in_string
            index += 1
        if not in_string:
            parts.append(rewrite_segment(formula[start:]))
        elif start < len(formula):
            parts.append(formula[start:])
        return "".join(parts)

    @staticmethod
    def _shift_formula_references_for_row_delete(
        formula: str,
        *,
        current_sheet: str,
        target_sheet: str,
        deleted_rows: list[int],
    ) -> str:
        """Repair same-sheet and cross-sheet references after row deletion."""
        if not isinstance(formula, str) or not formula.startswith("="):
            return formula
        deleted = sorted(set(deleted_rows))
        deleted_set = set(deleted)
        cell_pattern = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}"
        explicit = re.compile(
            rf"(?P<sheet>'[^']+'|[^=!'+\-*/(),\s]+)!"
            rf"(?P<first>{cell_pattern})(?:[:](?P<last>{cell_pattern}))?"
        )
        bare_range = re.compile(
            rf"(?<![A-Za-z0-9_!'])(?P<first>{cell_pattern}):(?P<last>{cell_pattern})"
            r"(?![A-Za-z0-9_])(?!(?:\s*)\()"
        )
        bare = re.compile(
            rf"(?<![A-Za-z0-9_!']){cell_pattern}"
            r"(?![A-Za-z0-9_])(?!(?:\s*)\()"
        )
        reference = re.compile(r"(?P<column>\$?[A-Za-z]{1,3})(?P<marker>\$?)(?P<row>[1-9][0-9]{0,6})")
        protected: list[str] = []

        def shift_cell(value: str) -> str:
            match = reference.fullmatch(value)
            if not match:
                return value
            row = int(match.group("row"))
            if row in deleted_set:
                return "#REF!"
            new_row = row - sum(candidate < row for candidate in deleted)
            return f"{match.group('column')}{match.group('marker')}{new_row}"

        def shift_range(first: str, last: str) -> str:
            first_match = reference.fullmatch(first)
            last_match = reference.fullmatch(last)
            if not first_match or not last_match:
                return f"{first}:{last}"
            first_row = int(first_match.group("row"))
            last_row = int(last_match.group("row"))
            if first_row > last_row:
                return f"{shift_cell(first)}:{shift_cell(last)}"
            remaining = [row for row in range(first_row, last_row + 1) if row not in deleted_set]
            if not remaining:
                return "#REF!"
            new_first = remaining[0] - sum(candidate < remaining[0] for candidate in deleted)
            new_last = remaining[-1] - sum(candidate < remaining[-1] for candidate in deleted)
            return (
                f"{first_match.group('column')}{first_match.group('marker')}{new_first}:"
                f"{last_match.group('column')}{last_match.group('marker')}{new_last}"
            )

        def protect(value: str) -> str:
            marker = f"\x00{len(protected)}\x00"
            protected.append(value)
            return marker

        def rewrite_segment(segment: str) -> str:
            def replace_explicit(match: re.Match[str]) -> str:
                if WorkbookSession._sheet_name_from_reference(match.group("sheet")) != target_sheet:
                    return match.group(0)
                changed = (
                    shift_range(match.group("first"), match.group("last"))
                    if match.group("last") else shift_cell(match.group("first"))
                )
                return protect(f"{match.group('sheet')}!{changed}")

            rewritten = explicit.sub(replace_explicit, segment)
            if current_sheet == target_sheet:
                rewritten = bare_range.sub(
                    lambda match: protect(shift_range(match.group("first"), match.group("last"))),
                    rewritten,
                )
                rewritten = bare.sub(lambda match: shift_cell(match.group(0)), rewritten)
            for index, value in enumerate(protected):
                rewritten = rewritten.replace(f"\x00{index}\x00", value)
            return rewritten

        parts: list[str] = []
        start = 0
        in_string = False
        index = 0
        while index < len(formula):
            if formula[index] == '"':
                if in_string and index + 1 < len(formula) and formula[index + 1] == '"':
                    index += 2
                    continue
                if not in_string:
                    parts.append(rewrite_segment(formula[start:index]))
                    start = index
                else:
                    parts.append(formula[start:index + 1])
                    start = index + 1
                in_string = not in_string
            index += 1
        if not in_string:
            parts.append(rewrite_segment(formula[start:]))
        elif start < len(formula):
            parts.append(formula[start:])
        return "".join(parts)

    @staticmethod
    def _shift_range_tokens_for_row_change(
        reference: str,
        *,
        insert_before_row: int | None = None,
        deleted_rows: list[int] | None = None,
        expand_adjacent: bool = True,
    ) -> str | None:
        """Shift local A1 ranges used by tables, filters, validation and formatting."""
        token_pattern = re.compile(
            r"(?P<first_col>\$?[A-Za-z]{1,3})(?P<first_abs>\$?)(?P<first_row>[1-9][0-9]{0,6})"
            r"(?:[:](?P<last_col>\$?[A-Za-z]{1,3})(?P<last_abs>\$?)(?P<last_row>[1-9][0-9]{0,6}))?"
        )
        deleted = sorted(set(deleted_rows or []))
        deleted_set = set(deleted)
        shifted_tokens: list[str] = []
        for raw_token in str(reference or "").split():
            match = token_pattern.fullmatch(raw_token)
            if match is None:
                raise ValueError(f"无法安全平移范围：{raw_token}")
            first_row = int(match.group("first_row"))
            last_row = int(match.group("last_row") or first_row)
            if deleted:
                remaining = [row for row in range(first_row, last_row + 1) if row not in deleted_set]
                if not remaining:
                    continue
                new_first = remaining[0] - sum(row < remaining[0] for row in deleted)
                new_last = remaining[-1] - sum(row < remaining[-1] for row in deleted)
            elif insert_before_row is not None:
                new_first = first_row + (1 if first_row >= insert_before_row else 0)
                new_last = last_row + (
                    1 if last_row >= insert_before_row
                    or expand_adjacent and last_row == insert_before_row - 1
                    else 0
                )
            else:
                new_first, new_last = first_row, last_row
            first = f"{match.group('first_col')}{match.group('first_abs')}{new_first}"
            if match.group("last_row") is None:
                shifted_tokens.append(first)
                continue
            last = f"{match.group('last_col')}{match.group('last_abs')}{new_last}"
            shifted_tokens.append(f"{first}:{last}")
        return " ".join(shifted_tokens) or None

    @classmethod
    def _repair_row_dependent_objects(
        cls,
        workbook: Any,
        *,
        target_sheet: str,
        insert_before_row: int | None = None,
        deleted_rows: list[int] | None = None,
    ) -> None:
        """Repair non-cell workbook references that openpyxl row moves leave stale."""
        if (insert_before_row is None) == (not deleted_rows):
            raise ValueError("必须且只能指定一种行结构变更")

        def shift_range(value: Any, *, expand_adjacent: bool = True) -> str | None:
            return cls._shift_range_tokens_for_row_change(
                str(value or ""),
                insert_before_row=insert_before_row,
                deleted_rows=deleted_rows,
                expand_adjacent=expand_adjacent,
            )

        def shift_formula(value: Any, *, current_sheet: str) -> Any:
            if not isinstance(value, str) or not value:
                return value
            had_equals = value.startswith("=")
            formula = value if had_equals else f"={value}"
            if insert_before_row is not None:
                shifted = cls._shift_formula_references_for_row_insert(
                    formula,
                    current_sheet=current_sheet,
                    target_sheet=target_sheet,
                    insert_before_row=insert_before_row,
                    expand_adjacent_ranges=True,
                )
            else:
                shifted = cls._shift_formula_references_for_row_delete(
                    formula,
                    current_sheet=current_sheet,
                    target_sheet=target_sheet,
                    deleted_rows=list(deleted_rows or []),
                )
            if "#REF!" in shifted and "#REF!" not in formula:
                raise ValueError("行结构变更会破坏工作簿对象引用")
            return shifted if had_equals else shifted[1:]

        target = workbook[target_sheet]
        for table in target.tables.values():
            table_ref = shift_range(table.ref)
            if table_ref is None:
                raise ValueError(f"行结构变更会删除表格 {table.name} 的全部范围")
            table.ref = table_ref
            if table.autoFilter is not None and table.autoFilter.ref:
                table.autoFilter.ref = shift_range(table.autoFilter.ref)
        if target.auto_filter.ref:
            target.auto_filter.ref = shift_range(target.auto_filter.ref)

        for worksheet in workbook.worksheets:
            validations = []
            for validation in worksheet.data_validations.dataValidation:
                if worksheet.title == target_sheet:
                    shifted_sqref = shift_range(validation.sqref)
                    if shifted_sqref is None:
                        continue
                    validation.sqref = shifted_sqref
                validation.formula1 = shift_formula(validation.formula1, current_sheet=worksheet.title)
                validation.formula2 = shift_formula(validation.formula2, current_sheet=worksheet.title)
                validations.append(validation)
            worksheet.data_validations.dataValidation = validations

            conditional_entries = []
            for conditional, rules in list(worksheet.conditional_formatting._cf_rules.items()):
                if worksheet.title == target_sheet:
                    shifted_sqref = shift_range(conditional.sqref)
                    if shifted_sqref is None:
                        continue
                    conditional.sqref = shifted_sqref
                for rule in rules:
                    if rule.formula:
                        rule.formula = [
                            shift_formula(formula, current_sheet=worksheet.title)
                            for formula in rule.formula
                        ]
                conditional_entries.append((conditional, rules))
            worksheet.conditional_formatting._cf_rules.clear()
            for conditional, rules in conditional_entries:
                worksheet.conditional_formatting._cf_rules[conditional] = rules

            for chart in worksheet._charts:
                for series in chart.series:
                    for source_name in ("val", "cat", "xVal", "yVal", "bubbleSize"):
                        source = getattr(series, source_name, None)
                        for reference_name in ("numRef", "strRef"):
                            reference = getattr(source, reference_name, None)
                            if reference is not None and reference.f:
                                reference.f = shift_formula(
                                    reference.f, current_sheet=worksheet.title,
                                )
                    title = getattr(series, "tx", None)
                    title_reference = getattr(title, "strRef", None)
                    if title_reference is not None and title_reference.f:
                        title_reference.f = shift_formula(
                            title_reference.f, current_sheet=worksheet.title,
                        )

                if worksheet.title == target_sheet:
                    anchor = getattr(chart, "anchor", None)
                    for marker_name in ("_from", "to"):
                        marker = getattr(anchor, marker_name, None)
                        if marker is None:
                            continue
                        row = int(marker.row) + 1
                        if insert_before_row is not None and row >= insert_before_row:
                            marker.row += 1
                        elif deleted_rows:
                            marker.row = max(
                                0,
                                marker.row - sum(deleted < row for deleted in set(deleted_rows)),
                            )

        for defined_name in workbook.defined_names.values():
            local_sheet_id = getattr(defined_name, "localSheetId", None)
            current_sheet = (
                workbook.sheetnames[local_sheet_id]
                if isinstance(local_sheet_id, int) and 0 <= local_sheet_id < len(workbook.sheetnames)
                else ""
            )
            defined_name.attr_text = shift_formula(
                defined_name.attr_text,
                current_sheet=current_sheet,
            )

    def insert_and_copy_row(self, change: dict[str, Any]) -> dict[str, Any]:
        """Insert a row and copy a preceding source row without mutating inputs.

        Formula references move relative to the newly inserted destination row.  The
        caller supplies stable source-cell values to protect against stale templates.
        """
        try:
            parsed = RowCopyChange.model_validate(change)
        except ValidationError as exc:
            raise ToolExecutionError("插入复制行参数不合法") from exc
        if parsed.source_row >= parsed.insert_before_row:
            raise ToolExecutionError("只能复制插入位置之前的行，避免行号移动导致误写")
        if not self.draft.is_file():
            raise ToolExecutionError("请先创建原始总表的独立副本")
        try:
            workbook = openpyxl.load_workbook(self.draft, data_only=False)
        except (OSError, ValueError, BadZipFile) as exc:
            raise ToolExecutionError("草稿工作簿不存在或无法读取") from exc
        try:
            if parsed.sheet not in workbook.sheetnames:
                raise ToolExecutionError("目标工作表不存在")
            sheet = workbook[parsed.sheet]
            if parsed.copy_max_column > sheet.max_column:
                raise ToolExecutionError("复制列范围超过当前工作表")
            if any(
                merged.min_row <= parsed.source_row <= merged.max_row
                or merged.min_row <= parsed.insert_before_row <= merged.max_row
                for merged in sheet.merged_cells.ranges
            ):
                raise ToolExecutionError("复制或插入行涉及合并单元格，需人工确认后处理")
            for coordinate, expected_value in parsed.expected_source_cells.items():
                if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}", coordinate):
                    raise ToolExecutionError("行标识必须是单一单元格坐标")
                if sheet[coordinate].row != parsed.source_row:
                    raise ToolExecutionError("行标识必须位于来源行")
                if sheet[coordinate].value != expected_value:
                    raise ToolExecutionError("来源行已变化，请重新读取后再决定")

            identity_header_names = {
                "姓名", "员工姓名", "工号", "员工工号", "员工编号", "人员编号",
                "身份证", "身份证号", "证件号码",
            }
            header_row = None
            identity_columns: dict[int, str] = {}
            for row_index in range(parsed.source_row - 1, max(0, parsed.source_row - 20), -1):
                found = {
                    cell.column: re.sub(r"\s+", "", str(cell.value or ""))
                    for cell in sheet[row_index]
                    if re.sub(r"\s+", "", str(cell.value or "")) in identity_header_names
                }
                if any(value in {"姓名", "员工姓名"} for value in found.values()):
                    header_row = row_index
                    identity_columns = found
                    break
            is_personnel_row = header_row is not None
            if is_personnel_row:
                supplied: dict[int, tuple[str, str | int]] = {}
                for coordinate, value in parsed.identity_values.items():
                    if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}", coordinate):
                        raise ToolExecutionError("人员身份坐标必须是单一单元格")
                    cell = sheet[coordinate]
                    if cell.row != parsed.insert_before_row or cell.column not in identity_columns:
                        raise ToolExecutionError("人员身份只能写入新行的姓名、工号或证件字段")
                    if str(value).strip() == "":
                        raise ToolExecutionError("新增人员身份字段不能为空")
                    supplied[cell.column] = (coordinate, value)
                required_columns = {
                    column for column, header in identity_columns.items()
                    if header in {"姓名", "员工姓名"}
                    or header in {"工号", "员工工号", "员工编号", "人员编号", "身份证", "身份证号", "证件号码"}
                }
                if not required_columns or not required_columns.issubset(supplied):
                    required_headers = "、".join(identity_columns[column] for column in sorted(required_columns))
                    raise ToolExecutionError(f"新增人员必须先提供完整身份字段：{required_headers}")
                new_identity = tuple(str(supplied[column][1]).strip() for column in sorted(required_columns))
                for row_index in range(header_row + 1, sheet.max_row + 1):
                    existing = tuple(
                        str(sheet.cell(row_index, column).value or "").strip()
                        for column in sorted(required_columns)
                    )
                    if existing == new_identity:
                        raise ToolExecutionError("新增人员身份与现有行重复，已拒绝插入")
            elif parsed.identity_values:
                raise ToolExecutionError("当前插入位置未识别为人员表，不能写入人员身份")

            formula_snapshot: list[tuple[str, str, str]] = []
            for worksheet in workbook.worksheets:
                for row in worksheet.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            formula_snapshot.append((worksheet.title, cell.coordinate, cell.value))

            source_cells = [sheet.cell(parsed.source_row, column) for column in range(1, parsed.copy_max_column + 1)]
            source_dimension = copy(sheet.row_dimensions[parsed.source_row])
            sheet.insert_rows(parsed.insert_before_row, amount=1)
            for column, source_cell in enumerate(source_cells, start=1):
                target_cell = sheet.cell(parsed.insert_before_row, column)
                value = source_cell.value
                if isinstance(value, str) and value.startswith("="):
                    value = Translator(value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
                target_cell.value = value
                if source_cell.has_style:
                    target_cell._style = copy(source_cell._style)
                if source_cell.number_format:
                    target_cell.number_format = source_cell.number_format
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.protection = copy(source_cell.protection)
                target_cell.comment = copy(source_cell.comment)
            if is_personnel_row:
                # A personnel template row supplies styles and formulas only.
                # Carrying its constant amounts/months into a new identity is
                # precisely how one person's prior-period data can be assigned
                # to another person.
                for column in range(1, parsed.copy_max_column + 1):
                    target_cell = sheet.cell(parsed.insert_before_row, column)
                    if not (isinstance(target_cell.value, str) and target_cell.value.startswith("=")):
                        target_cell.value = None
                for column, (_coordinate, value) in supplied.items():
                    sheet.cell(parsed.insert_before_row, column).value = value
                persisted_identity = tuple(
                    str(sheet.cell(parsed.insert_before_row, column).value or "").strip()
                    for column in sorted(required_columns)
                )
                if persisted_identity != new_identity:
                    raise ToolExecutionError("新增人员身份写入后校验失败，本批已回滚")
            target_dimension = sheet.row_dimensions[parsed.insert_before_row]
            for attribute in ("height", "hidden", "outlineLevel", "collapsed", "thickTop", "thickBot"):
                setattr(target_dimension, attribute, getattr(source_dimension, attribute))

            # openpyxl moves cells but leaves formula references untouched.
            # Repair formulas in moved rows relative to their new coordinates,
            # and repair references from formulas elsewhere in the workbook.
            for worksheet_name, coordinate, formula in formula_snapshot:
                if worksheet_name == parsed.sheet:
                    old_row = int(re.search(r"[0-9]+$", coordinate).group(0))
                    if old_row >= parsed.insert_before_row:
                        new_coordinate = f"{re.match(r'[A-Za-z]+', coordinate).group(0)}{old_row + 1}"
                        workbook[worksheet_name][new_coordinate].value = self._shift_formula_references_for_row_insert(
                            formula,
                            current_sheet=worksheet_name,
                            target_sheet=parsed.sheet,
                            insert_before_row=parsed.insert_before_row,
                            expand_adjacent_ranges=True,
                        )
                        continue
                workbook[worksheet_name][coordinate].value = self._shift_formula_references_for_row_insert(
                    formula,
                    current_sheet=worksheet_name,
                    target_sheet=parsed.sheet,
                    insert_before_row=parsed.insert_before_row,
                    expand_adjacent_ranges=worksheet_name != parsed.sheet,
                )
            self._repair_row_dependent_objects(
                workbook,
                target_sheet=parsed.sheet,
                insert_before_row=parsed.insert_before_row,
            )
            descriptor, temporary = tempfile.mkstemp(prefix="agent-row-copy-", suffix=".xlsx", dir=self.draft.parent)
            os.close(descriptor)
            temporary_path = Path(temporary)
            try:
                workbook.save(temporary_path)
                os.replace(temporary_path, self.draft)
            finally:
                temporary_path.unlink(missing_ok=True)
            result = {
                "sheet": parsed.sheet,
                "source_row": parsed.source_row,
                "inserted_row": parsed.insert_before_row,
                "copied_columns": parsed.copy_max_column,
            }
            if is_personnel_row:
                result["identity_cells"] = {
                    coordinate: value for _column, (coordinate, value) in supplied.items()
                }
                result["identity_validated"] = True
            return result
        except (KeyError, ValueError, AttributeError, OSError, BadZipFile) as exc:
            raise ToolExecutionError("插入复制行失败，本批未写入") from exc
        finally:
            workbook.close()

    def delete_rows(self, change: dict[str, Any]) -> dict[str, Any]:
        """Delete previously read rows from the draft after anchor verification.

        Rows are removed bottom-up so earlier row numbers stay stable.  Each
        row must be proven through an anchor cell whose current value matches,
        and merged-cell overlaps are rejected like the insert tool.
        """
        try:
            parsed = RowDeleteChange.model_validate(change)
        except ValidationError as exc:
            raise ToolExecutionError("删除行参数不合法") from exc
        if len(set(parsed.rows)) != len(parsed.rows):
            raise ToolExecutionError("待删行号重复，请去重后重试")
        if not self.draft.is_file():
            raise ToolExecutionError("请先创建原始总表的独立副本")
        try:
            workbook = openpyxl.load_workbook(self.draft, data_only=False)
        except (OSError, ValueError, BadZipFile) as exc:
            raise ToolExecutionError("草稿工作簿不存在或无法读取") from exc
        try:
            if parsed.sheet not in workbook.sheetnames:
                raise ToolExecutionError("目标工作表不存在")
            sheet = workbook[parsed.sheet]
            anchor_rows: dict[int, str] = {}
            for coordinate, expected_value in parsed.expected_cells.items():
                if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]{0,6}", coordinate):
                    raise ToolExecutionError("锚点必须是单一单元格坐标")
                anchor_rows[sheet[coordinate].row] = coordinate
                if sheet[coordinate].value != expected_value:
                    raise ToolExecutionError(f"锚点 {coordinate} 的值已变化，请重新读取后再决定")
            for row in parsed.rows:
                if row < 1 or row > sheet.max_row:
                    raise ToolExecutionError(f"第 {row} 行超出工作表范围")
                if row not in anchor_rows:
                    raise ToolExecutionError(f"第 {row} 行缺少锚点单元格，必须先读取并提供该行任一单元格的当前值")
                if any(
                    merged.min_row <= row <= merged.max_row
                    for merged in sheet.merged_cells.ranges
                ):
                    raise ToolExecutionError(f"第 {row} 行涉及合并单元格，需人工确认后处理")
            formula_snapshot: list[tuple[str, str, str]] = []
            for worksheet in workbook.worksheets:
                for cells in worksheet.iter_rows():
                    for cell in cells:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            formula_snapshot.append((worksheet.title, cell.coordinate, cell.value))
            deleted_rows = sorted(parsed.rows, reverse=True)
            for row in deleted_rows:
                sheet.delete_rows(row, amount=1)
            deleted_set = set(deleted_rows)
            for worksheet_name, coordinate, formula in formula_snapshot:
                old_row = int(re.search(r"[0-9]+$", coordinate).group(0))
                if worksheet_name == parsed.sheet:
                    if old_row in deleted_set:
                        continue
                    new_row = old_row - sum(candidate < old_row for candidate in deleted_rows)
                    new_coordinate = f"{re.match(r'[A-Za-z]+', coordinate).group(0)}{new_row}"
                else:
                    new_coordinate = coordinate
                workbook[worksheet_name][new_coordinate].value = self._shift_formula_references_for_row_delete(
                    formula,
                    current_sheet=worksheet_name,
                    target_sheet=parsed.sheet,
                    deleted_rows=deleted_rows,
                )
            self._repair_row_dependent_objects(
                workbook,
                target_sheet=parsed.sheet,
                deleted_rows=deleted_rows,
            )
            descriptor, temporary = tempfile.mkstemp(prefix="agent-row-delete-", suffix=".xlsx", dir=self.draft.parent)
            os.close(descriptor)
            temporary_path = Path(temporary)
            try:
                workbook.save(temporary_path)
                os.replace(temporary_path, self.draft)
            finally:
                temporary_path.unlink(missing_ok=True)
            return {
                "sheet": parsed.sheet,
                "deleted_rows": sorted(parsed.rows),
                "remaining_max_row": sheet.max_row,
            }
        except (KeyError, ValueError, AttributeError, OSError, BadZipFile) as exc:
            raise ToolExecutionError("删除行失败，本批未写入") from exc
        finally:
            workbook.close()

    def apply_verified_values(self, changes: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Apply pre-calculated values in an atomic batch for deterministic phases.

        This is intentionally not exposed as a model tool.  The named executor is
        responsible for deriving each value from approved source data and recording
        its audit evidence before calling this low-level workbook operation.
        """
        if not 1 <= len(changes) <= 500:
            raise ToolExecutionError("每批必须包含1至500项确定性单元格变更")
        try:
            parsed = [VerifiedCellValueChange.model_validate(change) for change in changes]
        except ValidationError as exc:
            raise ToolExecutionError("确定性单元格变更参数不合法") from exc
        if not self.draft.is_file():
            raise ToolExecutionError("请先创建原始总表的独立副本")
        try:
            workbook = openpyxl.load_workbook(self.draft, data_only=False)
        except (OSError, ValueError, BadZipFile) as exc:
            raise ToolExecutionError("草稿工作簿不存在或无法读取") from exc
        updates: list[dict[str, Any]] = []
        targets: set[tuple[str, str]] = set()
        try:
            for change in parsed:
                target_key = (change.target_sheet, change.target_cell)
                if target_key in targets:
                    raise ToolExecutionError("同一批次不能重复写入同一单元格")
                targets.add(target_key)
                if change.target_sheet not in workbook.sheetnames:
                    raise ToolExecutionError("目标工作表不存在")
                target = workbook[change.target_sheet][change.target_cell]
                if target.data_type == "f":
                    raise ToolExecutionError("不能用确定性数值覆盖总表公式")
                if target.value != change.expected_value:
                    raise ToolExecutionError("目标值已变化，请重新读取后再决定")
                if isinstance(change.new_value, str) and change.new_value.startswith("="):
                    raise ToolExecutionError("确定性写入不能创建公式")
                if isinstance(change.new_value, float) and not math.isfinite(change.new_value):
                    raise ToolExecutionError("写入数值必须是有限数")
                updates.append({
                    "target_sheet": change.target_sheet,
                    "target_cell": change.target_cell,
                    "old_value": target.value,
                    "new_value": change.new_value,
                })
            for update in updates:
                workbook[update["target_sheet"]][update["target_cell"]] = update["new_value"]
            descriptor, temporary = tempfile.mkstemp(prefix="agent-values-", suffix=".xlsx", dir=self.draft.parent)
            os.close(descriptor)
            temporary_path = Path(temporary)
            try:
                workbook.save(temporary_path)
                os.replace(temporary_path, self.draft)
            finally:
                temporary_path.unlink(missing_ok=True)
            return updates
        except (KeyError, ValueError, AttributeError, OSError, BadZipFile) as exc:
            raise ToolExecutionError("确定性单元格变更失败，本批未写入") from exc
        finally:
            workbook.close()

    def roll_forward_summary_row(self, change: dict[str, Any]) -> dict[str, Any]:
        """Create the next summary month without letting prior months recalculate."""
        try:
            parsed = SummaryRowRollForward.model_validate(change)
        except ValidationError as exc:
            raise ToolExecutionError("月份汇总滚动参数不合法") from exc
        if parsed.start_column > parsed.end_column:
            raise ToolExecutionError("月份汇总列范围不合法")
        if not self.draft.is_file():
            raise ToolExecutionError("请先创建原始总表的独立副本")
        workbook = None
        try:
            workbook = openpyxl.load_workbook(self.draft, data_only=False)
            cached_workbook = openpyxl.load_workbook(self.draft, data_only=True)
        except (OSError, ValueError, BadZipFile) as exc:
            if workbook is not None:
                workbook.close()
            raise ToolExecutionError("草稿工作簿不存在或无法读取") from exc
        try:
            if parsed.sheet not in workbook.sheetnames:
                raise ToolExecutionError("目标工作表不存在")
            sheet = workbook[parsed.sheet]
            cached_sheet = cached_workbook[parsed.sheet]
            if parsed.end_column > sheet.max_column:
                raise ToolExecutionError("月份汇总列范围超过当前工作表")
            if any(
                merged.min_row <= parsed.current_row <= merged.max_row
                for merged in sheet.merged_cells.ranges
            ):
                raise ToolExecutionError("当前月份行涉及合并单元格，需人工确认后处理")
            if sheet.cell(parsed.current_row, parsed.start_column).value != parsed.expected_current_month:
                raise ToolExecutionError("当前月份行已变化，请重新读取后再决定")
            if any(sheet.cell(row, parsed.start_column).value == parsed.new_month for row in range(1, sheet.max_row + 1)):
                raise ToolExecutionError("目标月份已存在，不能重复新增")
            source_cells = []
            history_values: dict[int, Any] = {}
            workbook_errors = {"#REF!", "#VALUE!", "#N/A", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!"}
            for column in range(parsed.start_column, parsed.end_column + 1):
                source_cell = sheet.cell(parsed.current_row, column)
                cached_value = cached_sheet.cell(parsed.current_row, column).value
                if source_cell.data_type == "f" or isinstance(source_cell.value, str) and source_cell.value.startswith("="):
                    if cached_value is None:
                        raise ToolExecutionError(f"公式 {source_cell.coordinate} 没有已计算缓存，不能冻结历史月份")
                    if isinstance(cached_value, str) and cached_value.upper() in workbook_errors:
                        raise ToolExecutionError(f"公式 {source_cell.coordinate} 的缓存是错误值，不能冻结历史月份")
                    history_value = cached_value
                else:
                    history_value = source_cell.value
                if isinstance(history_value, float) and not math.isfinite(history_value):
                    raise ToolExecutionError("历史冻结值必须是有限数")
                history_values[column] = history_value
                source_cells.append({
                    "coordinate": source_cell.coordinate,
                    "value": source_cell.value,
                    "style": copy(source_cell._style),
                    "font": copy(source_cell.font),
                    "fill": copy(source_cell.fill),
                    "border": copy(source_cell.border),
                    "alignment": copy(source_cell.alignment),
                    "protection": copy(source_cell.protection),
                    "comment": copy(source_cell.comment),
                })

            formula_snapshot: list[tuple[str, str, str]] = []
            for worksheet in workbook.worksheets:
                for row in worksheet.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            formula_snapshot.append((worksheet.title, cell.coordinate, cell.value))
            source_dimension = copy(sheet.row_dimensions[parsed.current_row])
            sheet.insert_rows(parsed.current_row, amount=1)
            for offset, source_cell in enumerate(source_cells):
                column = parsed.start_column + offset
                target_cell = sheet.cell(parsed.current_row, column)
                target_cell.value = source_cell["value"]
                target_cell._style = copy(source_cell["style"])
                target_cell.font = copy(source_cell["font"])
                target_cell.fill = copy(source_cell["fill"])
                target_cell.border = copy(source_cell["border"])
                target_cell.alignment = copy(source_cell["alignment"])
                target_cell.protection = copy(source_cell["protection"])
                target_cell.comment = copy(source_cell["comment"])
            sheet.cell(parsed.current_row, parsed.start_column).value = parsed.new_month
            for column, value in history_values.items():
                sheet.cell(parsed.current_row + 1, column).value = value
            for worksheet_name, coordinate, formula in formula_snapshot:
                old_row = int(re.search(r"[0-9]+$", coordinate).group(0))
                old_column = openpyxl.utils.column_index_from_string(
                    re.match(r"[A-Za-z]+", coordinate).group(0)
                )
                if (
                    worksheet_name == parsed.sheet
                    and old_row == parsed.current_row
                    and parsed.start_column <= old_column <= parsed.end_column
                ):
                    continue
                if worksheet_name == parsed.sheet and old_row >= parsed.current_row:
                    new_coordinate = f"{re.match(r'[A-Za-z]+', coordinate).group(0)}{old_row + 1}"
                else:
                    new_coordinate = coordinate
                workbook[worksheet_name][new_coordinate].value = self._shift_formula_references_for_row_insert(
                    formula,
                    current_sheet=worksheet_name,
                    target_sheet=parsed.sheet,
                    insert_before_row=parsed.current_row,
                    expand_adjacent_ranges=worksheet_name != parsed.sheet,
                )
            self._repair_row_dependent_objects(
                workbook,
                target_sheet=parsed.sheet,
                insert_before_row=parsed.current_row,
            )
            for row in (parsed.current_row, parsed.current_row + 1):
                dimension = sheet.row_dimensions[row]
                for attribute in ("height", "hidden", "outlineLevel", "collapsed", "thickTop", "thickBot"):
                    setattr(dimension, attribute, getattr(source_dimension, attribute))
            descriptor, temporary = tempfile.mkstemp(prefix="agent-month-roll-", suffix=".xlsx", dir=self.draft.parent)
            os.close(descriptor)
            temporary_path = Path(temporary)
            try:
                workbook.save(temporary_path)
                os.replace(temporary_path, self.draft)
            finally:
                temporary_path.unlink(missing_ok=True)
            return {
                "sheet": parsed.sheet,
                "current_row": parsed.current_row,
                "history_row": parsed.current_row + 1,
                "new_month": parsed.new_month,
                "start_column": parsed.start_column,
                "end_column": parsed.end_column,
                "history_values": {
                    sheet.cell(parsed.current_row + 1, column).coordinate: value
                    for column, value in history_values.items()
                },
                "history_snapshot": {
                    sheet.cell(row, column).coordinate: (
                        sheet.cell(row, column).value
                        if isinstance(sheet.cell(row, column).value, (str, int, float, bool, type(None)))
                        else str(sheet.cell(row, column).value)
                    )
                    for row in range(parsed.current_row + 1, sheet.max_row + 1)
                    for column in range(parsed.start_column, parsed.end_column + 1)
                },
            }
        except (KeyError, ValueError, AttributeError, OSError, BadZipFile) as exc:
            raise ToolExecutionError("月份汇总滚动失败，本批未写入") from exc
        finally:
            workbook.close()
            cached_workbook.close()


def source_write_schema() -> dict[str, Any]:
    return {"type": "function", "function": {
        "name": "apply_source_cells",
        "description": "把已读取核对的来源值复制或求和到独立草稿。不能覆盖公式，不支持空白当零；先读取目标旧值。",
        "parameters": {"type": "object", "additionalProperties": False,
                       "properties": {"changes": {"type": "array", "minItems": 1, "maxItems": 200,
                                                  "items": SourceCellChange.model_json_schema()}},
                       "required": ["changes"]},
    }}


def formula_divisor_schema() -> dict[str, Any]:
    return {"type": "function", "function": {
        "name": "apply_formula_divisors",
        "description": "在独立草稿中仅替换已核对公式里的指定数值分母；必须提供未变更的原公式。",
        "parameters": {"type": "object", "additionalProperties": False,
                       "properties": {"changes": {"type": "array", "minItems": 1, "maxItems": 500,
                                                  "items": FormulaDivisorChange.model_json_schema()}},
                       "required": ["changes"]},
    }}


def row_copy_schema() -> dict[str, Any]:
    return {"type": "function", "function": {
        "name": "insert_and_copy_row",
        "description": (
            "在独立草稿中插入一行并复制已核对模板行的格式与公式。"
            "人员表新增行必须同时提供 identity_values：工具会清空模板人员的常量业务值，"
            "先写入并校验新人员姓名/工号/证件字段；身份通过后才能另批写业务字段。"
            "不得用于合并单元格或未知模板行。"
        ),
        "parameters": {"type": "object", "additionalProperties": False,
                       "properties": {"change": RowCopyChange.model_json_schema()},
                       "required": ["change"]},
    }}


def row_delete_schema() -> dict[str, Any]:
    return {"type": "function", "function": {
        "name": "delete_rows",
        "description": (
            "在独立草稿中删除已读取核对过的数据行（如总表明细中不在本次来源里的人员）。"
            "每个待删行必须提供至少一个锚点单元格及其当前值以证明已读取；"
            "行号从大到小内部处理，一次调用可删多行；涉及合并单元格会被拒绝。"
        ),
        "parameters": {"type": "object", "additionalProperties": False,
                       "properties": {"change": RowDeleteChange.model_json_schema()},
                       "required": ["change"]},
    }}


def month_roll_forward_schema() -> dict[str, Any]:
    return {"type": "function", "function": {
        "name": "roll_forward_month",
        "description": (
            "为下一工资月在独立草稿中新增当前月汇总行，并用文件自带的已计算缓存"
            "冻结上月历史行。必须是跨月 Run 的第一个数据写入；不接受模型自报的历史数值。"
        ),
        "parameters": {
            "type": "object", "additionalProperties": False,
            "properties": {"change": SummaryRowRollForward.model_json_schema()},
            "required": ["change"],
        },
    }}
