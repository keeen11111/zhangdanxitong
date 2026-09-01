"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { AlertTriangle, ArrowRight, CheckCircle2, ClipboardCheck, Loader2 } from "lucide-react";
import { toast } from "sonner";

import { api, FinancialWorkQueueItem } from "@/lib/api";
import { getWorkQueueStatusDisplay } from "@/lib/work-queue-status";
import { Button } from "@/components/ui/button";

export default function WorkQueuePage() {
  const [items, setItems] = useState<FinancialWorkQueueItem[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    async function load() {
      try {
        const queue = await api.getFinancialWorkQueue();
        setItems(queue.items);
      } catch (error) {
        toast.error(error instanceof Error ? error.message : "待处理事项加载失败");
      } finally {
        setLoading(false);
      }
    }
    void load();
  }, []);

  return <div className="payroll-page max-w-5xl">
    <header className="flex flex-col gap-3 border-b border-slate-200/90 pb-7 sm:flex-row sm:items-end sm:justify-between">
      <div>
        <p className="payroll-kicker">Action queue</p>
        <h1 className="page-heading mt-2">待处理事项</h1>
        <p className="mt-2 text-sm leading-6 text-slate-600">只显示系统无法安全自动完成的批次。先处理人工事项，再确认发布正式账单。</p>
      </div>
      <Button asChild variant="outline" className="bg-white shadow-sm"><Link href="/projects">查看全部账单项目</Link></Button>
    </header>

    {loading ? <div className="flex h-52 items-center justify-center rounded-lg border border-slate-200/90 bg-white text-sm text-slate-500 shadow-[0_1px_2px_rgba(15,23,42,0.03)]" aria-busy="true"><Loader2 className="mr-2 h-5 w-5 animate-spin text-blue-700" />正在加载待处理事项…</div> : null}
    {!loading && !items.length ? <section className="rounded-lg border border-emerald-200 bg-emerald-50 px-5 py-12 text-center" role="status"><CheckCircle2 className="mx-auto h-8 w-8 text-emerald-700" /><h2 className="mt-3 text-sm font-semibold text-emerald-950">当前没有待处理事项</h2><p className="mt-1 text-sm text-emerald-800">已生成的财务批次均已发布，或尚未开始整合。</p><Button asChild className="mt-5 bg-emerald-700 hover:bg-emerald-800"><Link href="/projects">进入账单项目</Link></Button></section> : null}
    {!loading && items.length ? <section className="overflow-hidden rounded-lg border border-slate-200/90 bg-white shadow-[0_1px_2px_rgba(15,23,42,0.03)]" aria-label="待处理批次"><div className="divide-y divide-slate-200/90">{items.map((item) => {
      const status = getWorkQueueStatusDisplay(item.status);
      const Icon = item.status === "review_required" || item.status === "blocked" ? AlertTriangle : ClipboardCheck;
      const needsManualReview = item.issue_count > 0;
      const destination = `/projects/${item.project_id}/agent`;
      return <article key={item.project_id} id={`project-${item.project_id}`} className="px-4 py-5 transition-colors hover:bg-blue-50/30 sm:px-5">
        <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between"><div className="flex min-w-0 gap-3"><span className="mt-0.5 flex h-10 w-10 shrink-0 items-center justify-center rounded-lg bg-blue-50 text-blue-700"><Icon className="h-4 w-4" /></span><div className="min-w-0"><div className="flex flex-wrap items-center gap-2"><h2 className="truncate text-sm font-semibold text-slate-950">{item.project_name}</h2><span className={`rounded-full px-2 py-0.5 text-xs font-medium ${status.className}`}>{status.label}</span></div><p className="mt-1 text-xs text-slate-500">账期：{item.salary_month}{item.issue_count ? ` · ${item.issue_count} 项需要回答` : ""}</p><p className="mt-2 text-sm text-slate-700">当前环节：{status.step}</p></div></div><Button asChild className="w-full shrink-0 bg-blue-700 shadow-sm hover:bg-blue-800 sm:w-auto"><Link href={destination}>{needsManualReview ? "在对话中继续" : item.action_label}<ArrowRight className="ml-2 h-4 w-4" /></Link></Button></div>
      </article>;
    })}</div></section> : null}
  </div>;
}
