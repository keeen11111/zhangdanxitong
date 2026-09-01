"use client";

import { useEffect, useState } from "react";
import { CheckCircle2, ChevronDown, Download, FileSpreadsheet } from "lucide-react";

import { Button } from "@/components/ui/button";
import type { AgentRunResult } from "@/lib/api";

function displayValue(value: unknown): string {
  if (value === null || value === undefined || value === "") return "空";
  if (typeof value === "string") return value;
  try { return JSON.stringify(value); } catch { return String(value); }
}

export function AgentResultView({ result, busy, onDownload }: {
  result: AgentRunResult;
  busy: boolean;
  onDownload: () => void;
}) {
  const [expanded, setExpanded] = useState(true);

  useEffect(() => { setExpanded(true); }, [result.run_id, result.completed_at]);
  if (!result.available) return null;

  return (
    <section className="my-6 overflow-hidden rounded-xl border border-emerald-200 bg-white" aria-labelledby="agent-result-heading">
      <div className="flex flex-wrap items-start justify-between gap-4 p-5">
        <div className="flex min-w-0 items-start gap-3">
          <span className="mt-0.5 inline-flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-emerald-50 text-emerald-700"><CheckCircle2 className="h-5 w-5" /></span>
          <div className="min-w-0">
            <h2 id="agent-result-heading" className="text-base font-semibold text-slate-950">更新后表格已生成</h2>
            <p className="mt-1 break-all text-sm text-slate-600">{result.filename}</p>
            <p className="mt-1 text-xs leading-5 text-slate-500">{result.validation.detail || "结果已完成服务端校验"}</p>
          </div>
        </div>
        <div className="flex shrink-0 flex-wrap gap-2">
          <Button type="button" size="sm" variant="outline" onClick={() => setExpanded((value) => !value)} aria-expanded={expanded}>
            <ChevronDown className={`mr-1.5 h-3.5 w-3.5 transition-transform ${expanded ? "rotate-180" : ""}`} />
            {expanded ? "收起修改明细" : "查看处理结果"}
          </Button>
          <Button type="button" size="sm" className="bg-blue-700 text-white hover:bg-blue-800" onClick={onDownload} disabled={busy || !result.can_download}>
            <Download className="mr-1.5 h-3.5 w-3.5" />下载更新后表格
          </Button>
        </div>
      </div>

      {expanded ? <div className="border-t border-slate-200">
        <div className="flex flex-wrap items-center justify-between gap-2 bg-slate-50 px-5 py-3 text-xs text-slate-600">
          <span className="inline-flex items-center gap-1.5"><FileSpreadsheet className="h-3.5 w-3.5" />共 {result.change_count} 项实际修改</span>
          <span>校验状态：{result.validation.status || "unknown"}</span>
        </div>
        {result.changes.length ? <div className="max-h-80 overflow-auto">
          <table className="w-full min-w-[760px] border-collapse text-left text-xs">
            <thead className="sticky top-0 bg-white text-slate-500 shadow-[0_1px_0_rgba(226,232,240,1)]"><tr><th className="px-4 py-2.5 font-medium">Sheet / 位置</th><th className="px-4 py-2.5 font-medium">原值</th><th className="px-4 py-2.5 font-medium">新值</th><th className="px-4 py-2.5 font-medium">来源依据</th></tr></thead>
            <tbody className="divide-y divide-slate-100">{result.changes.map((change, index) => <tr key={`${change.target_sheet || change.sheet}-${change.target_cell || change.inserted_row}-${index}`} className="align-top text-slate-700"><td className="px-4 py-3 font-medium text-slate-900">{change.target_sheet || change.sheet || "-"}{change.target_cell ? `!${change.target_cell}` : change.inserted_row ? ` 第 ${change.inserted_row} 行` : ""}</td><td className="max-w-52 break-words px-4 py-3">{displayValue(change.old_value)}</td><td className="max-w-52 break-words px-4 py-3">{displayValue(change.new_value)}</td><td className="max-w-64 break-words px-4 py-3">{[change.source_file, change.source_sheet, ...(change.source_cells || [])].filter(Boolean).join(" / ") || "经核对的模板操作"}</td></tr>)}</tbody>
          </table>
        </div> : <p className="px-5 py-8 text-center text-sm text-slate-500">本次已生成可下载副本，没有产生可追溯的单元格变更。</p>}
      </div> : null}
    </section>
  );
}
