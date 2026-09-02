"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { AlertTriangle, CheckCircle2, Clock3, Loader2 } from "lucide-react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { api, FileMeta, FinancialWorkbookIntegration, IntegrationProgress, UploadFailure } from "@/lib/api";
import { financialIntegrationProgressPercent } from "@/lib/financial-workbook-flow";
import { mergeFileRecords } from "@/lib/file-list";
import { WorkbookUploadPanel } from "./workbook-upload-panel";
import { isEncryptedWorkbookFailure, promptForWorkbookPassword } from "./workbook-password";

function selectExcelFiles(files: File[]) {
  const accepted = files.filter((file) => /\.(xlsx|xls)$/i.test(file.name));
  const rejected = files
    .filter((file) => !/\.(xlsx|xls)$/i.test(file.name))
    .map((file) => ({ filename: file.name, message: "仅支持 .xlsx 或 .xls 文件" }));
  return { accepted, rejected };
}

const fallbackProgressStages: IntegrationProgress["stages"] = [
  { key: "prepare", label: "准备总表和来源文件", status: "pending" },
  { key: "match", label: "核对工作表与业务主键", status: "pending" },
  { key: "update", label: "安全写入总表数据", status: "pending" },
  { key: "export", label: "生成正式稿和修订稿", status: "pending" },
];

function FinancialIntegrationLoadingCard({
  progress,
  elapsedSeconds,
  fileCount,
}: {
  progress: IntegrationProgress | null;
  elapsedSeconds: number;
  fileCount: number;
}) {
  const stages = progress?.stages.length ? progress.stages : fallbackProgressStages;
  const percentage = financialIntegrationProgressPercent(stages);

  return (
    <div className="mt-4 overflow-hidden rounded-md border border-teal-200 bg-teal-50/80" role="status" aria-live="polite" aria-busy="true">
      <div className="p-4 sm:p-5">
        <div className="flex items-start gap-3">
          <span className="mt-0.5 flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-teal-700 text-white"><Loader2 className="h-4 w-4 animate-spin" /></span>
          <div><h3 className="text-sm font-semibold text-teal-950">正在核对并更新</h3><p className="mt-1 text-sm text-teal-900">{progress?.detail || `正在处理 ${fileCount} 份来源文件，请勿关闭页面。`}</p></div>
        </div>
        <div className="mt-4">
          <div className="mb-2 flex items-center justify-between text-xs font-medium text-teal-900"><span>整合进度</span><span>{percentage}%</span></div>
          <div className="h-2 overflow-hidden rounded-full bg-teal-100" role="progressbar" aria-label="财务整合进度" aria-valuemin={0} aria-valuemax={100} aria-valuenow={percentage}><div className="h-full rounded-full bg-teal-700 transition-[width] duration-500" style={{ width: `${percentage}%` }} /></div>
        </div>
        <p className="mt-3 flex items-center gap-2 text-xs text-teal-800"><Clock3 className="h-4 w-4" />已处理 {elapsedSeconds} 秒，大文件可能需要数分钟</p>
      </div>
      <ol className="grid border-t border-teal-200 bg-white/70 sm:grid-cols-2" aria-label="财务整合步骤">
        {stages.map((stage, index) => <li key={stage.key} className="flex items-center gap-2 border-b border-teal-100 px-4 py-2.5 text-xs sm:border-r"><span className={stage.status === "completed" ? "flex h-5 w-5 items-center justify-center rounded-full bg-teal-700 text-white" : stage.status === "running" ? "flex h-5 w-5 items-center justify-center rounded-full border-2 border-teal-700 text-teal-700" : "flex h-5 w-5 items-center justify-center rounded-full bg-slate-200 text-slate-500"}>{stage.status === "completed" ? "✓" : stage.status === "running" ? <Loader2 className="h-3 w-3 animate-spin" /> : index + 1}</span><span className={stage.status === "pending" ? "text-slate-500" : "font-medium text-teal-950"}>{stage.label}</span></li>)}
      </ol>
    </div>
  );
}

export function FinancialWorkbookIntegrationPage({ projectId }: { projectId: string }) {
  const [files, setFiles] = useState<FileMeta[]>([]);
  const [result, setResult] = useState<FinancialWorkbookIntegration | null>(null);
  const [masterUploading, setMasterUploading] = useState(false);
  const [sourcesUploading, setSourcesUploading] = useState(false);
  const [integrating, setIntegrating] = useState(false);
  const [failures, setFailures] = useState<UploadFailure[]>([]);
  const [completionOpen, setCompletionOpen] = useState(false);
  const [integrationProgress, setIntegrationProgress] = useState<IntegrationProgress | null>(null);
  const [integrationElapsedSeconds, setIntegrationElapsedSeconds] = useState(0);

  const masters = useMemo(() => files.filter((file) => file.file_type === "financial_master"), [files]);
  const sources = useMemo(() => files.filter((file) => file.file_type === "financial_source"), [files]);
  const legacyFiles = useMemo(
    () => files.filter((file) => file.file_type === "template" || file.file_type === "source"),
    [files],
  );
  const ready = masters.length === 1 && sources.length > 0;
  const workflowStep = result ? 3 : ready ? 2 : 1;

  const loadWorkspace = useCallback(async () => {
    const [filesResult, resultResult] = await Promise.allSettled([
      api.listFiles(projectId),
      api.getLatestFinancialWorkbookIntegration(projectId),
    ]);
    if (filesResult.status === "fulfilled") setFiles(filesResult.value);
    setResult(resultResult.status === "fulfilled" ? resultResult.value : null);
  }, [projectId]);

  useEffect(() => { void loadWorkspace(); }, [loadWorkspace]);

  useEffect(() => {
    if (!integrating) return;
    let active = true;
    const startedAt = Date.now();
    const refreshProgress = async () => {
      try {
        const nextProgress = await api.getFinancialWorkbookIntegrationProgress(projectId);
        if (active && nextProgress.status === "processing") setIntegrationProgress(nextProgress);
      } catch {
        // Keep the visible local progress card active while the endpoint initializes.
      }
    };
    const elapsedTimer = window.setInterval(() => {
      setIntegrationElapsedSeconds(Math.floor((Date.now() - startedAt) / 1000));
    }, 1000);
    void refreshProgress();
    const progressTimer = window.setInterval(() => { void refreshProgress(); }, 1200);
    return () => {
      active = false;
      window.clearInterval(elapsedTimer);
      window.clearInterval(progressTimer);
    };
  }, [integrating, projectId]);

  async function uploadMaster(selected: File[]) {
    const { accepted, rejected } = selectExcelFiles(selected.slice(0, 1));
    setFailures(rejected);
    if (!accepted.length) return toast.error("请选择一份 Excel 财务总表");
    setMasterUploading(true);
    try {
      let uploaded: FileMeta;
      try {
        uploaded = await api.uploadFile(projectId, accepted[0], "financial_master");
      } catch (error) {
        const failure = { filename: accepted[0].name, message: error instanceof Error ? error.message : "" };
        if (!isEncryptedWorkbookFailure(failure)) throw error;
        const password = promptForWorkbookPassword();
        if (password === undefined) return;
        uploaded = await api.uploadFile(projectId, accepted[0], "financial_master", undefined, password);
      }
      setFiles((current) => mergeFileRecords(current, [uploaded]));
      setResult(null);
      toast.success("通用财务总表已上传");
    } catch (error) {
      const message = error instanceof Error ? error.message : "总表上传失败";
      setFailures([{ filename: accepted[0].name, message }]);
      toast.error(message);
    } finally {
      setMasterUploading(false);
    }
  }

  async function uploadSources(selected: File[]) {
    const { accepted, rejected } = selectExcelFiles(selected);
    setFailures(rejected);
    if (!accepted.length) return toast.error("请选择至少一份来源更新文件");
    setSourcesUploading(true);
    try {
      let upload = await api.uploadFilesIndividually(projectId, accepted, "financial_source");
      const encryptedFiles = accepted.filter((file) => upload.failed.some((failure) => failure.filename === file.name && isEncryptedWorkbookFailure(failure)));
      if (encryptedFiles.length) {
        const password = promptForWorkbookPassword();
        if (password !== undefined) {
          const retry = await api.uploadFilesIndividually(projectId, encryptedFiles, "financial_source", password);
          const encryptedNames = new Set(encryptedFiles.map((file) => file.name));
          upload = {
            uploaded: [...upload.uploaded, ...retry.uploaded],
            failed: [...upload.failed.filter((failure) => !encryptedNames.has(failure.filename)), ...retry.failed],
          };
        }
      }
      const allFailures = [...rejected, ...upload.failed];
      setFailures(allFailures);
      if (upload.uploaded.length) {
        setFiles((current) => mergeFileRecords(current, upload.uploaded));
        setResult(null);
        toast.success(`已添加 ${upload.uploaded.length} 份来源更新文件`);
      }
      if (allFailures.length) toast.error(`${allFailures.length} 份文件未上传，请查看原因`);
    } finally {
      setSourcesUploading(false);
    }
  }

  async function removeFile(file: FileMeta) {
    if (!window.confirm(`确认移除“${file.original_name}”？`)) return;
    try {
      await api.deleteFile(projectId, file.id);
      setFiles((current) => current.filter((item) => item.id !== file.id));
      setResult(null);
      toast.success("文件已移除");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "文件移除失败");
    }
  }

  async function integrate() {
    if (!ready) return toast.error("请先上传一份总表和至少一份来源更新文件");
    setIntegrationProgress({
      status: "processing",
      stages: fallbackProgressStages.map((stage, index) => ({ ...stage, status: index === 0 ? "running" : "pending" })),
      detail: "正在准备总表和来源文件",
      updated_at: "",
    });
    setIntegrationElapsedSeconds(0);
    setIntegrating(true);
    try {
      const nextResult = await api.createFinancialWorkbookIntegration(projectId);
      setResult(nextResult);
      setCompletionOpen(true);
      toast.success(`整合完成：自动更新 ${nextResult.auto_update_count} 个单元格`);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "通用财务整合失败");
    } finally {
      setIntegrating(false);
    }
  }

  return <div className="payroll-page pb-8">
    <header className="border-b border-slate-200 pb-5">
      <p className="payroll-kicker">Financial workbook</p>
      <h1 className="mt-2 text-3xl font-bold tracking-tight text-slate-950">用来源文件更新公司财务总表</h1>
      <p className="mt-2 max-w-3xl text-sm leading-6 text-slate-600">适用于不同公司的应收、应付、费用、对账等 Excel 总表。系统不依赖固定 Sheet 名；只有目标、字段和业务主键都能唯一确认时才自动写入。</p>
    </header>

    <ol className="grid gap-3 border-b border-slate-200 pb-5 text-sm sm:grid-cols-3" aria-label="更新步骤">
      {[
        ["上传文件", "上传一份总表和来源更新文件"],
        ["核对并更新", "系统自动核对并写入可确认的数据"],
        ["查看结果", "下载正式稿、修订稿并查看人工事项"],
      ].map(([label, description], index) => {
        const stepNumber = index + 1;
        const completed = stepNumber < workflowStep;
        const active = stepNumber === workflowStep;
        return <li key={label} className="flex items-center gap-3"><span className={`flex h-8 w-8 items-center justify-center rounded-full font-semibold ${completed || active ? "bg-teal-700 text-white" : "bg-slate-200 text-slate-600"}`}>{completed ? "✓" : stepNumber}</span><span><b className="block text-slate-800">{label}</b><small className="text-slate-500">{description}</small></span></li>;
      })}
    </ol>

    {result ? <section className="rounded-xl border border-emerald-200 bg-emerald-50/60 p-4 sm:p-5" aria-live="polite">
      <div className="flex gap-3"><CheckCircle2 className="mt-0.5 h-5 w-5 shrink-0 text-emerald-700" /><div><h2 className="text-sm font-semibold text-emerald-950">整合完毕</h2><p className="mt-1 text-sm text-emerald-900">已自动更新 {result.auto_update_count} 个单元格，生成正式稿和修订稿{result.issues.length ? `；另有 ${result.issues.length} 项待人工处理。` : "，没有待人工处理事项。"}</p></div></div>
    </section> : null}

    <div className="grid gap-5 lg:grid-cols-2">
      <WorkbookUploadPanel title="上传待更新的财务总表" description="上传本公司的唯一总表。它决定导出的 Sheet、布局和公式，系统不会要求任何固定名称。" helper="拖入总表，或点击选择" files={masters} uploading={masterUploading} disabled={masters.length > 0} tone="master" onSelect={(selected) => void uploadMaster(selected)} onRemove={(file) => void removeFile(file)} />
      <WorkbookUploadPanel title="上传来源更新文件" description="可上传多份账单、明细或更新表。它们仅提供更新依据，不会替换总表结构。" helper="拖入来源文件，或点击多选" files={sources} multiple uploading={sourcesUploading} tone="changes" onSelect={(selected) => void uploadSources(selected)} onRemove={(file) => void removeFile(file)} />
    </div>

    {legacyFiles.length ? <section className="rounded-xl border border-sky-200 bg-sky-50 px-4 py-3" role="status"><p className="text-sm font-semibold text-sky-950">发现 {legacyFiles.length} 份历史工资流程文件</p><p className="mt-1 text-xs leading-5 text-sky-900">它们不会混入本次通用财务整合，避免误写。若需继续处理该历史工资项目，可进入旧工作区；新公司财务文件请在当前页面按“总表”和“来源”重新上传。</p><Button asChild size="sm" variant="outline" className="mt-3 border-sky-300 bg-white text-sky-950 hover:bg-sky-100"><Link href={`/projects/${projectId}/legacy-payroll`}>打开旧工资工作区</Link></Button></section> : null}

    {failures.length ? <section className="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3" role="alert"><div className="flex items-center gap-2"><AlertTriangle className="h-4 w-4 text-amber-700" /><h2 className="text-sm font-semibold text-amber-950">以下文件未上传</h2></div><ul className="mt-2 space-y-1 text-xs leading-5 text-amber-900">{failures.map((failure) => <li key={`${failure.filename}-${failure.message}`}><span className="font-medium">{failure.filename}</span>：{failure.message}</li>)}</ul></section> : null}

    <section className="rounded-md border border-slate-200 bg-white p-4 sm:p-5"><div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between"><div><h2 className="text-sm font-semibold text-slate-950">开始整合</h2><p className="mt-1 text-xs leading-5 text-slate-500">可安全确认的数据会写入总表，其他内容会在结果页直接列为待人工处理。</p></div><Button onClick={() => void integrate()} disabled={!ready || integrating || masterUploading || sourcesUploading} className="w-full bg-teal-700 hover:bg-teal-800 sm:w-auto">{integrating ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}{integrating ? "正在核对并更新…" : "开始核对并更新"}</Button></div>{integrating ? <FinancialIntegrationLoadingCard progress={integrationProgress} elapsedSeconds={integrationElapsedSeconds} fileCount={sources.length} /> : null}</section>

    <Dialog open={completionOpen} onOpenChange={setCompletionOpen}><DialogContent><DialogHeader><DialogTitle>整合完毕</DialogTitle><DialogDescription>正式稿和修订稿已经生成{result?.issues.length ? `，另有 ${result.issues.length} 项待人工处理。` : "，没有待人工处理事项。"}</DialogDescription></DialogHeader><DialogFooter><Button type="button" variant="outline" onClick={() => setCompletionOpen(false)}>关闭</Button></DialogFooter></DialogContent></Dialog>
  </div>;
}
