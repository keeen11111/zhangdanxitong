"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { useParams, useSearchParams } from "next/navigation";
import { CheckCircle2, Download, FileSpreadsheet, Loader2 } from "lucide-react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { AgentRun, api } from "@/lib/api";

const ACCEPTED_STATUSES = new Set(["passed", "reference_match", "demo_reference_match"]);

export default function AgentAcceptancePage() {
  const { projectId } = useParams<{ projectId: string }>();
  const searchParams = useSearchParams();
  const runId = searchParams.get("run") || "";
  const [run, setRun] = useState<AgentRun | null>(null);
  const [loading, setLoading] = useState(true);
  const [downloading, setDownloading] = useState(false);

  useEffect(() => {
    if (!runId) { setLoading(false); return; }
    let active = true;
    void api.getAgentRun(runId).then((next) => { if (active) setRun(next); })
      .catch((error) => { if (active) toast.error(error instanceof Error ? error.message : "验收批次加载失败"); })
      .finally(() => { if (active) setLoading(false); });
    return () => { active = false; };
  }, [runId]);

  const accepted = Boolean(run?.draft_filename && ACCEPTED_STATUSES.has(String(run.validation?.status || "")));

  async function download() {
    if (!run) return;
    setDownloading(true);
    try {
      const blob = await api.downloadAcceptedAgentFinal(run.run_id);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = run.draft_filename || "正式稿.xlsx";
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
      toast.success("正式稿已开始下载");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "正式稿下载失败");
    } finally { setDownloading(false); }
  }

  if (loading) {
    return <div className="flex min-h-[60vh] items-center justify-center text-sm text-slate-500"><Loader2 className="mr-2 h-4 w-4 animate-spin" />正在加载验收结果...</div>;
  }

  if (!accepted) {
    return <main className="mx-auto max-w-2xl px-6 py-16"><section className="border border-amber-200 bg-amber-50 p-6"><h1 className="text-lg font-semibold text-slate-950">尚未通过验收</h1><p className="mt-2 text-sm leading-6 text-slate-700">该批次仍有未完成步骤或校验项，系统不会提供下载入口。</p><Button asChild className="mt-5" variant="outline"><Link href={`/projects/${projectId}/agent`}>返回 Agent 继续处理</Link></Button></section></main>;
  }

  return <main className="mx-auto max-w-3xl px-6 py-12"><Link href={`/projects/${projectId}/agent`} className="text-sm text-slate-500 hover:text-slate-900">返回 Agent</Link><section className="mt-6 border border-emerald-200 bg-white p-6"><div className="flex items-start gap-3"><CheckCircle2 className="mt-0.5 h-6 w-6 text-emerald-700" /><div><p className="text-sm font-medium text-emerald-800">验收通过</p><h1 className="mt-1 text-xl font-semibold text-slate-950">正式稿可以下载</h1><p className="mt-2 text-sm leading-6 text-slate-600">{run.validation?.detail || "服务端验收已完成。原始上传文件保持不变。"}</p></div></div><div className="mt-6 border-t border-slate-200 pt-5"><div className="flex items-center gap-3"><FileSpreadsheet className="h-5 w-5 text-blue-700" /><p className="min-w-0 break-all text-sm text-slate-800">{run.draft_filename}</p></div><Button type="button" className="mt-5 bg-blue-700 text-white hover:bg-blue-800" onClick={() => void download()} disabled={downloading}>{downloading ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载正式稿</Button></div></section></main>;
}
