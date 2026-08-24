"use client";

import * as React from "react";
import { Trash2, Plus, Search, ClipboardPaste } from "lucide-react";

import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { cn } from "@/lib/utils";

export interface EditableGridProps {
  columns: string[];
  rows: Record<string, unknown>[];
  requiredFields?: string[];
  onChange: (rows: Record<string, unknown>[]) => void;
  onDeleteRow?: (rowIdx: number) => void;
  onAddRow?: () => void;
}

export function EditableGrid({
  columns,
  rows,
  requiredFields = [],
  onChange,
  onDeleteRow,
  onAddRow,
}: EditableGridProps) {
  const [searchTerm, setSearchTerm] = React.useState("");

  const handleCellChange = (
    rowIdx: number,
    col: string,
    value: string
  ) => {
    const newRows = [...rows];
    newRows[rowIdx] = { ...newRows[rowIdx], [col]: value };
    onChange(newRows);
  };

  const handleDeleteRow = (rowIdx: number) => {
    onDeleteRow?.(rowIdx);
  };

  const handleAddRow = () => {
    onAddRow?.();
  };

  // 批量粘贴：监听 paste 事件
  const handlePaste = (e: React.ClipboardEvent<HTMLTableSectionElement>) => {
    e.preventDefault();
    const clipboardData = e.clipboardData.getData("text/plain");
    if (!clipboardData) return;

    // 按行拆分
    const pastedLines = clipboardData.split(/\r?\n/).filter((line) => line.trim());

    // 找到当前焦点所在行（通过 data 属性）
    const target = e.target as HTMLElement;
    const rowEl = target.closest("tr[data-row-idx]") as HTMLTableRowElement | null;
    const startRow = rowEl ? parseInt(rowEl.dataset.rowIdx || "0", 10) : 0;

    // 获取焦点所在列
    const cellEl = target.closest("td[data-col]") as HTMLTableCellElement | null;
    const startCol = cellEl ? cellEl.dataset.col || "" : "";

    const newRows = [...rows];

    pastedLines.forEach((line, lineOffset) => {
      const targetRowIdx = startRow + lineOffset;
      if (targetRowIdx >= newRows.length) return;

      // 按 Tab 拆分列
      const pastedCells = line.split("\t");
      const startColIdx = startCol ? columns.indexOf(startCol) : 0;

      pastedCells.forEach((cellValue, colOffset) => {
        const targetColIdx = startColIdx + colOffset;
        if (targetColIdx >= columns.length) return;
        const colName = columns[targetColIdx];
        newRows[targetRowIdx] = {
          ...newRows[targetRowIdx],
          [colName]: cellValue.trim(),
        };
      });
    });

    onChange(newRows);
  };

  // 筛选行
  const filteredRows = React.useMemo(() => {
    if (!searchTerm.trim()) {
      return rows.map((row, idx) => ({ row, idx }));
    }
    const term = searchTerm.toLowerCase();
    return rows
      .map((row, idx) => ({ row, idx }))
      .filter(({ row }) =>
        columns.some((col) =>
          String(row[col] ?? "").toLowerCase().includes(term)
        )
      );
  }, [rows, searchTerm, columns]);

  // 统计缺失
  const missingCount = React.useMemo(() => {
    if (!requiredFields.length) return 0;
    return rows.filter((row) =>
      requiredFields.some(
        (field) => !row[field] || String(row[field]).trim() === ""
      )
    ).length;
  }, [rows, requiredFields]);

  const isEmpty = rows.length === 0;

  return (
    <div className="flex flex-col gap-2">
      {/* 工具栏 */}
      <div className="flex items-center justify-between gap-2">
        <div className="flex items-center gap-2 text-xs text-slate-500">
          <span>共 {rows.length} 行</span>
          {requiredFields.length > 0 && (
            <>
              <span className="text-slate-300">|</span>
              <span className="text-red-500">
                {requiredCountLabel(missingCount)}
              </span>
            </>
          )}
        </div>
        <div className="flex items-center gap-2">
          <div className="relative">
            <Search className="absolute left-2 top-1/2 -translate-y-1/2 h-3 w-3 text-slate-400" />
            <Input
              placeholder="搜索..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="h-7 w-32 pl-7 text-xs"
            />
          </div>
          <div className="flex items-center gap-1 text-[10px] text-slate-400">
            <ClipboardPaste className="h-3 w-3" />
            <span>支持粘贴</span>
          </div>
        </div>
      </div>

      {/* 表格 */}
      <div className="overflow-auto max-h-[55vh] border rounded-lg">
        <table className="w-full text-xs border-collapse">
          <thead className="sticky top-0 z-10">
            <tr>
              <th className="w-12 bg-slate-100 font-bold text-left px-2 py-1 border-b border-slate-200">
                #
              </th>
              {columns.map((col) => {
                const isRequired = requiredFields.includes(col);
                return (
                  <th
                    key={col}
                    className="min-w-[120px] bg-slate-100 font-bold text-left px-2 py-1 border-b border-slate-200 whitespace-nowrap"
                  >
                    <span className={isRequired ? "text-red-600" : ""}>
                      {col}
                      {isRequired && <span className="text-red-500"> *</span>}
                    </span>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody onPaste={handlePaste}>
            {isEmpty ? (
              <tr>
                <td
                  colSpan={columns.length + 1}
                  className="text-center text-slate-400 py-8"
                >
                  暂无数据，点击下方新增行
                </td>
              </tr>
            ) : (
              filteredRows.map(({ row, idx: rowIdx }) => {
                // 检查该行是否有必填项缺失
                const hasMissing = requiredFields.some(
                  (field) => !row[field] || String(row[field]).trim() === ""
                );
                return (
                  <tr
                    key={rowIdx}
                    data-row-idx={rowIdx}
                    className={cn(
                      "hover:bg-slate-50",
                      hasMissing && "bg-amber-50/40"
                    )}
                  >
                    <td className="w-12 border-b border-slate-100 px-1 py-1 align-middle">
                      <div className="flex items-center gap-1">
                        <span className="text-slate-400">{rowIdx + 1}</span>
                        <button
                          type="button"
                          onClick={() => handleDeleteRow(rowIdx)}
                          className="text-slate-400 hover:text-red-500 transition-colors"
                          aria-label="删除该行"
                        >
                          <Trash2 className="h-3 w-3" />
                        </button>
                      </div>
                    </td>
                    {columns.map((col) => {
                      const value = String(row[col] ?? "");
                      const isEmptyCell = value === "";
                      const isRequired = requiredFields.includes(col);
                      return (
                        <td
                          key={col}
                          data-col={col}
                          className="min-w-[120px] border-b border-slate-100 p-0 align-middle"
                        >
                          <Input
                            value={value}
                            onChange={(e) =>
                              handleCellChange(rowIdx, col, e.target.value)
                            }
                            className={cn(
                              "border-transparent rounded-none h-8 text-xs px-2 py-1",
                              "focus-visible:ring-2 focus-visible:ring-teal-500 focus-visible:ring-offset-0",
                              isEmptyCell && isRequired && "bg-amber-100/60",
                              isEmptyCell && !isRequired && "bg-slate-50"
                            )}
                          />
                        </td>
                      );
                    })}
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
      <Button
        type="button"
        variant="outline"
        onClick={handleAddRow}
        className="w-full border-dashed text-slate-500"
      >
        <Plus className="h-4 w-4 mr-1" />
        新增行
      </Button>
    </div>
  );
}

function requiredCountLabel(missing: number): string {
  if (missing === 0) return "必填项完整";
  return `${missing} 行缺必填项`;
}

export default EditableGrid;
