"use client";

import { CheckCircle2, Download } from "lucide-react";

import { Button } from "@/components/ui/button";
import type { AgentRunResult } from "@/lib/api";

export function AgentResultView({ result, busy, onDownload }: {
  result: AgentRunResult;
  busy: boolean;
  onDownload: () => void;
}) {
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
          <Button type="button" size="sm" className="bg-blue-700 text-white hover:bg-blue-800" onClick={onDownload} disabled={busy || !result.can_download}>
            <Download className="mr-1.5 h-3.5 w-3.5" />下载结果文件
          </Button>
        </div>
      </div>

    </section>
  );
}
