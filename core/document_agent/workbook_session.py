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
    history_values: dict[str, str | int | float | bool | None] = Field(min_length=1, max_length=500)


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
        if not 1 <= len(changes) <= 200:
            raise ToolExecutionError("每批必须包含1至200项变更")
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
            target_dimension = sheet.row_dimensions[parsed.insert_before_row]
            for attribute in ("height", "hidden", "outlineLevel", "collapsed", "thickTop", "thickBot"):
                setattr(target_dimension, attribute, getattr(source_dimension, attribute))
            descriptor, temporary = tempfile.mkstemp(prefix="agent-row-copy-", suffix=".xlsx", dir=self.draft.parent)
            os.close(descriptor)
            temporary_path = Path(temporary)
            try:
                workbook.save(temporary_path)
                os.replace(temporary_path, self.draft)
            finally:
                temporary_path.unlink(missing_ok=True)
            return {
                "sheet": parsed.sheet,
                "source_row": parsed.source_row,
                "inserted_row": parsed.insert_before_row,
                "copied_columns": parsed.copy_max_column,
            }
        except (KeyError, ValueError, AttributeError, OSError, BadZipFile) as exc:
            raise ToolExecutionError("插入复制行失败，本批未写入") from exc
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
        try:
            workbook = openpyxl.load_workbook(self.draft, data_only=False)
        except (OSError, ValueError, BadZipFile) as exc:
            raise ToolExecutionError("草稿工作簿不存在或无法读取") from exc
        try:
            if parsed.sheet not in workbook.sheetnames:
                raise ToolExecutionError("目标工作表不存在")
            sheet = workbook[parsed.sheet]
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
            expected_cells = {
                sheet.cell(parsed.current_row, column).coordinate
                for column in range(parsed.start_column, parsed.end_column + 1)
            }
            if set(parsed.history_values) != expected_cells:
                raise ToolExecutionError("历史冻结值必须完整覆盖当前月份行")
            for coordinate, value in parsed.history_values.items():
                if isinstance(value, str) and value.startswith("="):
                    raise ToolExecutionError("历史冻结值不能包含公式")
                if isinstance(value, float) and not math.isfinite(value):
                    raise ToolExecutionError("历史冻结值必须是有限数")

            source_cells = [sheet.cell(parsed.current_row, column) for column in range(parsed.start_column, parsed.end_column + 1)]
            source_coordinates = [cell.coordinate for cell in source_cells]
            source_dimension = copy(sheet.row_dimensions[parsed.current_row])
            sheet.insert_rows(parsed.current_row, amount=1)
            for offset, source_cell in enumerate(source_cells):
                column = parsed.start_column + offset
                target_cell = sheet.cell(parsed.current_row, column)
                value = source_cell.value
                if isinstance(value, str) and value.startswith("="):
                    value = Translator(value, origin=source_coordinates[offset]).translate_formula(target_cell.coordinate)
                target_cell.value = value
                target_cell._style = copy(source_cell._style)
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.protection = copy(source_cell.protection)
                target_cell.comment = copy(source_cell.comment)
            sheet.cell(parsed.current_row, parsed.start_column).value = parsed.new_month
            for coordinate, value in parsed.history_values.items():
                column = sheet[coordinate].column
                sheet.cell(parsed.current_row + 1, column).value = value
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
            }
        except (KeyError, ValueError, AttributeError, OSError, BadZipFile) as exc:
            raise ToolExecutionError("月份汇总滚动失败，本批未写入") from exc
        finally:
            workbook.close()


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
        "description": "在独立草稿中插入一行并复制已核对的上一行格式与公式；不得用于合并单元格或未知模板行。",
        "parameters": {"type": "object", "additionalProperties": False,
                       "properties": {"change": RowCopyChange.model_json_schema()},
                       "required": ["change"]},
    }}
