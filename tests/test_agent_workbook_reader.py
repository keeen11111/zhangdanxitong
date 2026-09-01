from pathlib import Path

import openpyxl
import pytest

from core.document_agent.orchestrator import ToolExecutionError
from core.document_agent.workbook_reader import WorkbookReader


@pytest.fixture
def reader(tmp_path):
    master = tmp_path / "总表.xlsx"
    source = tmp_path / "来源.xlsx"
    for path in (master, source):
        wb = openpyxl.Workbook()
        wb.active.title = "人员"
        wb.active.append(["姓名", "金额"])
        wb.active.append(["李楠", 123])
        wb.active["A30"] = "李晓苛"
        wb.active["B30"] = "=B2/21"
        wb.save(path)
        wb.close()
    return WorkbookReader(master, {source.name: source})


def test_search_finds_rows_outside_profile_without_writes(reader):
    before = reader.master.read_bytes()
    result = reader.search_cells("总表.xlsx", "人员", "李晓苛", True)
    assert result["matches"] == [{"sheet": "人员", "cell": "A30", "value": "李晓苛"}]
    assert result["complete"] is True
    assert reader.master.read_bytes() == before


def test_read_returns_formula_and_explicit_missing_cache(reader):
    result = reader.read_source_range("来源.xlsx", "人员", "A30:B30")
    assert result["rows"] == [["李晓苛", "=B2/21"]]
    assert result["cached_rows"] == [["李晓苛", None]]
    assert result["missing_formula_caches"] == ["B30"]


@pytest.mark.parametrize("area", ["A:A", "1:3", "XFE1", "A1048577", "A1:Z1000", "B2:A1"])
def test_invalid_or_unbounded_ranges_fail_closed(reader, area):
    with pytest.raises(ToolExecutionError):
        reader.read_source_range("来源.xlsx", "人员", area)


def test_reader_rejects_unregistered_file_and_missing_sheet(reader):
    with pytest.raises(ToolExecutionError):
        reader.search_cells("../other.xlsx", "人员", "李楠", True)
    with pytest.raises(ToolExecutionError):
        reader.read_range("不存在", "A1:B2")
