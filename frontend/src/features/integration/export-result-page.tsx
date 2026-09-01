"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import { AlertTriangle, ArrowRight, CheckCircle2, Download, FileSpreadsheet, Loader2 } from "lucide-react";
import { toast } from "sonner";

import { api, FinancialWorkbookIntegration, IntegrationProgress, PipelineExportResult, Project } from "@/lib/api";
import { Button } from "@/components/ui/button";
import { downloadWorkbookDirect } from "@/lib/download";
import { cn } from "@/lib/utils";
import { selectLatestIntegrationResult, shouldLoadFinancialIntegrationResult } from "@/lib/integration-result-selection";

function IntegrationProgressCard({ progress }: { progress: IntegrationProgress | null }) {
  if (!progress || progress.status === "idle") return null;
  const isRunning = progress.status === "processing";
  const isProblem = progress.status === "blocked" || progress.status === "failed";
  const heading = isRunning ? "正在更新最终总表" : isProblem ? "本次整合需要处理" : "最终总表已同步";

  return (
    <section className={cn("rounded-lg border p-4 shadow-[0_1px_2px_rgba(15,23,42,0.03)] sm:p-5", isProblem ? "border-red-200 bg-red-50" : isRunning ? "border-blue-200 bg-blue-50" : "border-emerald-200 bg-emerald-50")} aria-live="polite">
      <div className="flex items-start gap-3">
        {isRunning ? <Loader2 className="mt-0.5 h-5 w-5 shrink-0 animate-spin text-blue-700" /> : isProblem ? <AlertTriangle className="mt-0.5 h-5 w-5 shrink-0 text-red-700" /> : <CheckCircle2 className="mt-0.5 h-5 w-5 shrink-0 text-emerald-700" />}
        <div className="min-w-0"><h2 className="text-sm font-semibold text-slate-950">{heading}</h2><p className="mt-1 text-sm text-slate-700">{progress.detail || "正在同步任务状态…"}</p></div>
      </div>
      <ol className="mt-4 grid gap-2 sm:grid-cols-4" aria-label="整合进度">
        {progress.stages.map((stage) => <li key={stage.key} className={cn("rounded-md border px-3 py-2 text-xs", stage.status === "completed" ? "border-emerald-200 bg-white text-emerald-900" : stage.status === "running" ? "border-sky-300 bg-white text-sky-900" : "border-slate-200 bg-white/70 text-slate-500")}><span className="block font-medium">{stage.label}</span><span className="mt-1 block">{stage.status === "completed" ? "已完成" : stage.status === "running" ? "进行中" : "等待中"}</span></li>)}
      </ol>
    </section>
  );
}

function financialIssueTitle(issue: Record<string, unknown>) {
  const titles: Record<string, string> = {
    unprofiled_source_sheet: "来源工作表无法识别",
    empty_source_sheet: "来源工作表为空",
    ambiguous_sheet: "无法确定目标工作表",
    unmatched_sheet: "未匹配到目标工作表",
    insufficient_topic_evidence: "工作表主题证据不足",
    missing_business_key: "缺少可定位的业务主键",
    duplicate_source_record: "来源记录重复",
    ambiguous_record: "总表记录不唯一",
    unknown_record: "总表中没有对应记录",
    conflicting_value: "来源数据存在冲突",
    source_formula_conflict: "来源公式不能安全写入",
  };
  return titles[String(issue.issue_type || "")] || "需要人工确认";
}

function financialIssueSource(issue: Record<string, unknown>) {
  const files = Array.isArray(issue.source_files) ? issue.source_files.filter((value): value is string => typeof value === "string") : [];
  const sheets = Array.isArray(issue.source_sheets) ? issue.source_sheets.filter((value): value is string => typeof value === "string") : [];
  return [files.join("、"), sheets.join("、")].filter(Boolean).join(" · ") || "—";
}

function FinancialWorkbookResult({
  project,
  projectId,
  result,
}: {
  project: Project | null;
  projectId: string;
  result: FinancialWorkbookIntegration;
}) {
  const [downloading, setDownloading] = useState<"formal" | "review" | null>(null);

  async function downloadWorkbook(kind: "formal" | "review") {
    const isReview = kind === "review";
    const filename = isReview ? result.review_filename : result.filename;
    if (!filename) return;
    setDownloading(kind);
    try {
      await downloadWorkbookDirect(
        isReview ? api.financialWorkbookReviewDownloadUrl(projectId) : api.financialWorkbookIntegrationDownloadUrl(projectId),
        filename,
      );
      toast.success(isReview ? "修订稿已开始下载" : "正式稿已开始下载");
    } catch (downloadError) {
      toast.error(downloadError instanceof Error ? downloadError.message : "文件下载失败，请稍后重试");
    } finally {
      setDownloading(null);
    }
  }

  return (
    <div className="payroll-page max-w-6xl pb-8">
      <header className="border-b border-slate-200/90 pb-7"><Link href={`/projects/${projectId}`} className="text-sm text-slate-500 hover:text-slate-900">← 返回整合页</Link><p className="payroll-kicker mt-5">{project?.salary_month || "本月"}财务总表</p><h1 className="page-heading mt-2">整合结果</h1><p className="mt-2 text-sm leading-6 text-slate-600">正式稿、修订稿和待人工处理内容都集中在这里，无需再进入待处理事项。</p></header>

      <section className="rounded-lg border border-emerald-200 bg-emerald-50 p-4 shadow-[0_1px_2px_rgba(15,23,42,0.03)] sm:p-5"><div className="flex items-start gap-3"><CheckCircle2 className="mt-0.5 h-5 w-5 shrink-0 text-emerald-700" /><div><h2 className="text-sm font-semibold text-emerald-950">整合完毕</h2><p className="mt-1 text-sm text-emerald-900">已从 {result.source_file_count} 份来源文件自动更新 {result.auto_update_count} 个单元格，匹配 {result.matches.length} 个工作表。</p></div></div></section>

      <section className="rounded-lg border border-slate-200/90 bg-white p-4 shadow-[0_1px_2px_rgba(15,23,42,0.03)] sm:p-5"><div className="flex items-start gap-3"><FileSpreadsheet className="mt-0.5 h-5 w-5 shrink-0 text-blue-700" /><div><h2 className="text-sm font-semibold text-slate-950">交付文件</h2><p className="mt-1 text-sm leading-6 text-slate-600">正式稿用于交付；修订稿数据一致，并标明本次更新的原值、新值和位置。</p></div></div><div className="mt-4 grid gap-3 lg:grid-cols-2"><article className="rounded-lg border border-slate-200 p-4"><h3 className="text-sm font-semibold text-slate-950">正式稿</h3><p className="mt-1 text-sm leading-6 text-slate-600">已经完成自动更新的公司财务总表，可直接下载使用。</p><Button type="button" className="mt-4 w-full bg-blue-700 shadow-sm hover:bg-blue-800" disabled={downloading !== null} onClick={() => void downloadWorkbook("formal")}>{downloading === "formal" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载正式稿</Button></article><article className="rounded-lg border border-amber-200 bg-amber-50/40 p-4"><h3 className="text-sm font-semibold text-slate-950">修订稿（带变更标记）</h3><p className="mt-1 text-sm leading-6 text-slate-700">更新单元格带原值、新值和来源批注，并包含“修改记录”工作表。</p>{result.review_filename ? <Button type="button" variant="outline" className="mt-4 w-full border-amber-400 bg-white text-amber-950 hover:bg-amber-100" disabled={downloading !== null} onClick={() => void downloadWorkbook("review")}>{downloading === "review" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载修订稿</Button> : <p className="mt-4 text-xs text-amber-900">这是历史结果，尚无修订稿；重新整合后即可生成。</p>}</article></div></section>

      <section className={cn("overflow-hidden rounded-xl border", result.issues.length ? "border-amber-200 bg-white" : "border-emerald-200 bg-emerald-50")}><div className="p-4 sm:p-5"><div className="flex items-start gap-3">{result.issues.length ? <AlertTriangle className="mt-0.5 h-5 w-5 shrink-0 text-amber-700" /> : <CheckCircle2 className="mt-0.5 h-5 w-5 shrink-0 text-emerald-700" />}<div><h2 className="text-sm font-semibold text-slate-950">{result.issues.length ? `待人工处理（${result.issues.length} 项）` : "没有待人工处理事项"}</h2><p className="mt-1 text-sm leading-6 text-slate-700">{result.issues.length ? "以下内容未被自动写入，请按列表逐项核对。" : "所有可识别内容均已完成处理。"}</p></div></div></div>{result.issues.length ? <div className="overflow-x-auto border-t border-amber-200"><table className="min-w-[760px] w-full text-left"><thead className="bg-amber-50 text-xs text-amber-950"><tr><th className="px-4 py-3 font-medium">需要处理</th><th className="px-4 py-3 font-medium">问题说明</th><th className="px-4 py-3 font-medium">来源文件与工作表</th></tr></thead><tbody>{result.issues.map((issue, index) => <tr key={`${String(issue.issue_type || "issue")}-${index}`} className="border-t border-slate-200 align-top"><td className="px-4 py-4 text-sm font-medium text-slate-950">{financialIssueTitle(issue)}</td><td className="max-w-xl px-4 py-4 text-sm leading-6 text-slate-700">{typeof issue.message === "string" ? issue.message : "系统无法安全自动写入，请人工核对。"}</td><td className="px-4 py-4 text-xs leading-5 text-amber-900">{financialIssueSource(issue)}</td></tr>)}</tbody></table></div> : null}</section>
    </div>
  );
}

export function ExportResultPage({ projectId }: { projectId: string }) {
  const [project, setProject] = useState<Project | null>(null);
  const [result, setResult] = useState<PipelineExportResult | null>(null);
  const [financialResult, setFinancialResult] = useState<FinancialWorkbookIntegration | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [progress, setProgress] = useState<IntegrationProgress | null>(null);
  const [downloading, setDownloading] = useState<"formal" | "review" | null>(null);
  const previousProgressStatusRef = useRef<IntegrationProgress["status"] | null>(null);

  const loadResult = useCallback(async () => {
    setLoading(true);
    setError(null);
    setResult(null);
    setFinancialResult(null);
    const [projectResponse, filesResponse, pipelineResponse, financialProgressResponse] = await Promise.allSettled([
      api.getProject(projectId),
      api.listFiles(projectId),
      api.getLatestPipelineExport(projectId),
      api.getFinancialWorkbookIntegrationProgress(projectId),
    ]);
    if (projectResponse.status === "fulfilled") setProject(projectResponse.value);
    const financialProgress = financialProgressResponse.status === "fulfilled" ? financialProgressResponse.value : null;
    const shouldLoadFinancialResult = filesResponse.status === "fulfilled" && shouldLoadFinancialIntegrationResult(filesResponse.value, financialProgress);
    const financialResponse = shouldLoadFinancialResult
      ? await Promise.allSettled([api.getLatestFinancialWorkbookIntegration(projectId)])
      : [];
    const financial = financialResponse[0]?.status === "fulfilled" ? financialResponse[0].value : null;
    const pipeline = pipelineResponse.status === "fulfilled" ? pipelineResponse.value : null;
    const financialError = financialResponse[0]?.status === "rejected" ? financialResponse[0].reason : null;
    if (selectLatestIntegrationResult(financial, pipeline) === "financial" && financial) {
      setFinancialResult(financial);
    } else if (pipeline) {
      setResult(pipeline);
    } else {
      setError(financialError instanceof Error ? financialError.message : "尚未找到可查收的制作结果");
    }
    setLoading(false);
  }, [projectId]);

  useEffect(() => { void loadResult(); }, [loadResult]);

  useEffect(() => {
    if (financialResult) return;
    let active = true;
    async function loadProgress() {
      try {
        const nextProgress = await api.getIntegrationProgress(projectId);
        if (!active) return;
        const previousStatus = previousProgressStatusRef.current;
        previousProgressStatusRef.current = nextProgress.status;
        setProgress(nextProgress);
        if (previousStatus === "processing" && (nextProgress.status === "completed" || nextProgress.status === "completed_with_issues")) {
          void loadResult();
        }
      } catch {
        // The result page remains usable if a server predates the progress API.
      }
    }
    void loadProgress();
    const timer = window.setInterval(() => { void loadProgress(); }, 2000);
    return () => { active = false; window.clearInterval(timer); };
  }, [financialResult, loadResult, projectId]);

  if (loading) return <div className="mx-auto max-w-6xl py-16 text-center text-sm text-slate-500"><Loader2 className="mx-auto mb-3 h-5 w-5 animate-spin" />正在加载整合状态…</div>;
  if (financialResult) return <FinancialWorkbookResult project={project} projectId={projectId} result={financialResult} />;
  if (!result) return <div className="mx-auto max-w-3xl space-y-5"><IntegrationProgressCard progress={progress} /><section className="rounded-lg border border-slate-200/90 bg-white p-6 shadow-[0_1px_2px_rgba(15,23,42,0.03)]"><h1 className="text-xl font-semibold text-slate-950">暂未找到整合结果</h1><p className="mt-2 text-sm text-slate-600">{error || "请返回整合页，完成本月文件整合后再查看。"}</p><div className="mt-5 flex flex-wrap gap-3"><Button type="button" variant="outline" onClick={() => void loadResult()}>重新加载</Button><Button asChild className="bg-blue-700 shadow-sm hover:bg-blue-800"><Link href={`/projects/${projectId}`}>返回整合页</Link></Button></div></section></div>;

  const peopleBalanced = result.coverage_gap_count === 0
    && result.input_person_count === result.output_person_count + result.manual_only_person_count;
  const pendingIssueCount = result.issues.filter((issue) => issue.status !== "confirmed").length;
  const hasIssues = pendingIssueCount > 0;

  async function downloadWorkbook(kind: "formal" | "review") {
    const isReview = kind === "review";
    const filename = isReview ? result.review_filename : result.filename;
    if (!filename) return;
    setDownloading(kind);
    try {
      await downloadWorkbookDirect(
        isReview ? api.reviewExportDownloadUrl(projectId) : api.pipelineExportDownloadUrl(projectId),
        filename,
      );
      toast.success(isReview ? "修改稿已开始下载" : "正式稿已开始下载");
    } catch (downloadError) {
      toast.error(downloadError instanceof Error ? downloadError.message : "文件下载失败，请稍后重试");
    } finally {
      setDownloading(null);
    }
  }

  return (
    <div className="payroll-page max-w-6xl pb-8">
      <header className="border-b border-slate-200/90 pb-7">
        <Link href={`/projects/${projectId}`} className="text-sm text-slate-500 hover:text-slate-900">← 返回整合页</Link>
        <p className="payroll-kicker mt-5">{project?.salary_month || "本月"}工资交付</p>
        <h1 className="page-heading mt-2">最终总表已更新</h1>
        <p className="mt-2 text-sm leading-6 text-slate-600">自动处理结果已写入最终总表。需要人工确认的事项，请在专门的在线处理页完成编辑并直接回写。</p>
      </header>

      <IntegrationProgressCard progress={progress} />

      <section className="rounded-lg border border-slate-200/90 bg-white p-4 shadow-[0_1px_2px_rgba(15,23,42,0.03)] sm:p-5">
        <div className="flex items-center gap-2"><CheckCircle2 className="h-5 w-5 text-emerald-700" /><h2 className="text-base font-semibold text-slate-950">总表更新摘要</h2></div>
        <dl className="mt-5 grid grid-cols-2 gap-px overflow-hidden rounded-lg border border-slate-200 bg-slate-200 lg:grid-cols-5"><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">输入候选人员</dt><dd className="mt-1 text-lg font-semibold tabular-nums text-slate-950">{result.input_person_count}</dd></div><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">总表自动写入</dt><dd className="mt-1 text-lg font-semibold tabular-nums text-slate-950">{result.output_person_count}</dd></div><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">待人工确认</dt><dd className="mt-1 text-lg font-semibold tabular-nums text-amber-900">{result.manual_only_person_count}</dd></div><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">变更文件</dt><dd className="mt-1 text-lg font-semibold tabular-nums text-slate-950">{result.source_file_count}</dd></div><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">覆盖校验</dt><dd className={cn("mt-1 text-sm font-semibold", peopleBalanced ? "text-emerald-800" : "text-red-700")}>{peopleBalanced ? "候选人员全部有去向" : `存在 ${result.coverage_gap_count} 人缺口`}</dd></div></dl>
      </section>

      <section className={cn("rounded-lg border p-4 shadow-[0_1px_2px_rgba(15,23,42,0.03)] sm:p-5", hasIssues ? "border-amber-200 bg-amber-50/50" : "border-emerald-200 bg-emerald-50/50")}>
        <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-start gap-3"><AlertTriangle className={cn("mt-0.5 h-5 w-5", hasIssues ? "text-amber-700" : "text-emerald-700")} /><div><h2 className="text-sm font-semibold text-slate-950">{hasIssues ? "待人工处理" : "人工事项已全部处理"}</h2><p className="mt-1 text-sm leading-6 text-slate-700">{hasIssues ? `还有 ${pendingIssueCount} 项需要核对。进入在线处理页后，编辑结果会直接更新最终总表。` : "无需下载、预览或上传人工处理文件。"}</p></div></div>
          {hasIssues ? <Link href={`/projects/${projectId}/manual-review`} className="inline-flex min-h-10 shrink-0 cursor-pointer items-center justify-center gap-2 rounded-md bg-blue-700 px-4 py-2 text-sm font-semibold text-white shadow-sm hover:bg-blue-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-700 focus-visible:ring-offset-2">处理待人工任务<ArrowRight className="h-4 w-4" /></Link> : null}
        </div>
      </section>

      <section className="rounded-lg border border-slate-200/90 bg-white p-4 shadow-[0_1px_2px_rgba(15,23,42,0.03)] sm:p-5">
        <div className="flex items-start gap-3"><FileSpreadsheet className="mt-0.5 h-5 w-5 shrink-0 text-blue-700" /><div><h2 className="text-sm font-semibold text-slate-950">交付文件</h2><p className="mt-1 text-sm leading-6 text-slate-600">正式稿用于交付；修改稿与正式稿数据一致，专门用于核查本次更新。</p></div></div>
        <div className="mt-4 grid gap-3 lg:grid-cols-2">
          <article className="rounded-lg border border-slate-200 p-4"><h3 className="text-sm font-semibold text-slate-950">正式稿</h3><p className="mt-1 text-sm leading-6 text-slate-600">更新后的最终工资总表，可直接作为本次交付文件。</p><Button type="button" className="mt-4 w-full bg-blue-700 shadow-sm hover:bg-blue-800" disabled={downloading !== null} onClick={() => void downloadWorkbook("formal")}>{downloading === "formal" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载正式稿</Button></article>
          <article className="rounded-lg border border-amber-200 bg-amber-50/40 p-4"><h3 className="text-sm font-semibold text-slate-950">修改稿（带变更标记）</h3><p className="mt-1 text-sm leading-6 text-slate-700">每个更新单元格带“原值 / 新值”批注，另有“修改记录”工作表，列出 Sheet、字段、单元格位置和变更类型。</p>{result.review_filename ? <Button type="button" variant="outline" className="mt-4 w-full border-amber-400 bg-white text-amber-950 hover:bg-amber-100" disabled={downloading !== null} onClick={() => void downloadWorkbook("review")}>{downloading === "review" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载修改稿</Button> : <p className="mt-4 text-xs text-amber-900">当前历史结果没有修改稿；请重新执行本次总表更新后生成。</p>}</article>
        </div>
      </section>
    </div>
  );
}
