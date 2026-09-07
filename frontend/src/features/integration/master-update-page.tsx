"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { AlertTriangle, CheckCircle2, Clock3, Download, Loader2 } from "lucide-react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { api, FileMeta, IntegrationProgress, IntegrationResult, PipelineExportResult, Project, UploadFailure } from "@/lib/api";
import { downloadWorkbookDirect } from "@/lib/download";
import { mergeFileRecords } from "@/lib/file-list";
import { WorkbookUploadPanel } from "./workbook-upload-panel";
import { isEncryptedWorkbookFailure, promptForWorkbookPassword } from "./workbook-password";

type Result = IntegrationResult | PipelineExportResult;

function completionNoticeKey(projectId: string, filename: string) {
  return `payroll-completion-notice:${projectId}:${filename}`;
}

function excelSelection(selected: File[]) {
  const accepted = selected.filter((file) => /\.(xlsx|xls)$/i.test(file.name));
  const rejected = selected
    .filter((file) => !/\.(xlsx|xls)$/i.test(file.name))
    .map((file) => ({ filename: file.name, message: "仅支持 .xlsx 或 .xls 文件" }));
  return { accepted, rejected };
}

function StepLabel({ current, step, children }: { current: number; step: number; children: React.ReactNode }) {
  const completed = current > step;
  const active = current === step;
  return <li className={completed || active ? "text-teal-800" : "text-slate-400"}><span className={completed || active ? "mr-2 inline-flex h-5 w-5 items-center justify-center rounded-full bg-teal-700 text-xs font-semibold text-white" : "mr-2 inline-flex h-5 w-5 items-center justify-center rounded-full bg-slate-200 text-xs font-semibold text-slate-500"}>{completed ? "✓" : step}</span>{children}</li>;
}

const fallbackIntegrationStages: IntegrationProgress["stages"] = [
  { key: "base", label: "加载总表底板", status: "pending" },
  { key: "source", label: "解析变更文件", status: "pending" },
  { key: "match", label: "核对人员与数据", status: "pending" },
  { key: "export", label: "生成更新后的总表", status: "pending" },
];

function formatRemainingTime(seconds: number) {
  if (seconds <= 0) return "即将完成";
  return seconds <= 60 ? "约 1 分钟" : `约 ${Math.ceil(seconds / 60)} 分钟`;
}

function IntegrationLoadingCard({
  progress,
  elapsedSeconds,
  estimatedSeconds,
  fileCount,
}: {
  progress: IntegrationProgress | null;
  elapsedSeconds: number;
  estimatedSeconds: number;
  fileCount: number;
}) {
  const stages = progress?.stages.length ? progress.stages : fallbackIntegrationStages;
  const runningIndex = stages.findIndex((stage) => stage.status === "running");
  const completedCount = stages.filter((stage) => stage.status === "completed").length;
  const activeIndex = runningIndex >= 0 ? runningIndex : Math.min(completedCount, stages.length - 1);
  const percentage = Math.min(96, completedCount * 25 + (runningIndex >= 0 ? 12 : 4));
  const remainingSeconds = estimatedSeconds - elapsedSeconds;

  return (
    <section className="mt-4 overflow-hidden rounded-md border border-teal-200 bg-teal-50/80" role="status" aria-live="polite">
      <div className="flex flex-col gap-4 p-4 sm:p-5">
        <div className="flex items-start gap-3">
          <span className="mt-0.5 flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-teal-700 text-white"><Loader2 className="h-4 w-4 animate-spin" /></span>
          <div className="min-w-0">
            <h3 className="text-sm font-semibold text-teal-950">正在更新总表</h3>
            <p className="mt-1 text-sm text-teal-900">{progress?.detail || `正在处理 ${fileCount} 份变更文件，请勿关闭此页面。`}</p>
          </div>
        </div>

        <div>
          <div className="mb-2 flex items-center justify-between text-xs font-medium text-teal-900"><span>处理进度</span><span>{percentage}%</span></div>
          <div className="h-2 overflow-hidden rounded-full bg-teal-100" role="progressbar" aria-label="总表更新进度" aria-valuemin={0} aria-valuemax={100} aria-valuenow={percentage}>
            <div className="h-full rounded-full bg-teal-700 transition-[width] duration-500" style={{ width: `${percentage}%` }} />
          </div>
        </div>

        <div className="grid gap-2 border-t border-teal-200 pt-3 text-sm sm:grid-cols-2">
          <p className="flex items-center gap-2 text-teal-950"><Clock3 className="h-4 w-4 text-teal-700" />预计剩余：{formatRemainingTime(remainingSeconds)}</p>
          <p className="text-teal-800 sm:text-right">已完成 {completedCount}/{stages.length} 个处理步骤</p>
        </div>
      </div>
      <ol className="grid divide-y divide-teal-100 border-t border-teal-200 bg-white/70 sm:grid-cols-2 sm:divide-x sm:divide-y-0" aria-label="总表更新处理步骤">
        {stages.map((stage, index) => {
          const completed = stage.status === "completed";
          const running = stage.status === "running" || (!progress && index === activeIndex);
          return <li key={stage.key} className="flex items-center gap-2 px-4 py-2.5 text-xs"><span className={completed ? "flex h-5 w-5 items-center justify-center rounded-full bg-teal-700 text-white" : running ? "flex h-5 w-5 items-center justify-center rounded-full border-2 border-teal-700 text-teal-700" : "flex h-5 w-5 items-center justify-center rounded-full bg-slate-200 text-slate-500"}>{completed ? "✓" : running ? <Loader2 className="h-3 w-3 animate-spin" /> : index + 1}</span><span className={completed || running ? "font-medium text-teal-950" : "text-slate-500"}>{stage.label}</span></li>;
        })}
      </ol>
    </section>
  );
}

export function MasterUpdatePage({ projectId }: { projectId: string }) {
  const [project, setProject] = useState<Project | null>(null);
  const [files, setFiles] = useState<FileMeta[]>([]);
  const [result, setResult] = useState<Result | null>(null);
  const [masterUploading, setMasterUploading] = useState(false);
  const [changesUploading, setChangesUploading] = useState(false);
  const [integrating, setIntegrating] = useState(false);
  const [failures, setFailures] = useState<UploadFailure[]>([]);
  const [completionOpen, setCompletionOpen] = useState(false);
  const [downloading, setDownloading] = useState<"formal" | "review" | null>(null);
  const [changesConfirmed, setChangesConfirmed] = useState(false);
  const [integrationProgress, setIntegrationProgress] = useState<IntegrationProgress | null>(null);
  const [integrationElapsedSeconds, setIntegrationElapsedSeconds] = useState(0);

  const masters = useMemo(() => files.filter((file) => file.file_type === "template" || file.file_type === "financial_master"), [files]);
  const master = masters.length === 1 ? masters[0] : null;
  const changes = useMemo(() => files.filter((file) => file.file_type === "source" || file.file_type === "financial_source"), [files]);
  const step = !master ? 1 : !changesConfirmed ? 2 : 3;
  const estimatedIntegrationSeconds = Math.max(45, 25 + changes.length * 35);
  const personnelCoverage = result?.validation?.personnel_coverage;
  const completionMessage = personnelCoverage
    ? `已自动处理 ${personnelCoverage.processed_person_count}/${personnelCoverage.input_person_count} 人，匹配率 ${Math.round((personnelCoverage.match_rate || 0) * 100)}%；已回填 ${personnelCoverage.matched_field_count}/${personnelCoverage.input_field_count} 项来源信息。`
    : result?.issue_count ? `还有 ${result.issue_count} 项需要处理。` : "全部内容已更新完成。";

  const loadWorkspace = useCallback(async () => {
    const [projectResult, fileResult, exportResult] = await Promise.allSettled([
      api.getProject(projectId),
      api.listFiles(projectId),
      api.getLatestPipelineExport(projectId),
    ]);
    if (projectResult.status === "fulfilled") setProject(projectResult.value);
    if (fileResult.status === "fulfilled") setFiles(fileResult.value);
    if (exportResult.status === "fulfilled") {
      setResult(exportResult.value);
      if (!window.sessionStorage.getItem(completionNoticeKey(projectId, exportResult.value.filename))) {
        setCompletionOpen(true);
      }
    } else {
      setResult(null);
    }
  }, [projectId]);

  useEffect(() => { void loadWorkspace(); }, [loadWorkspace]);

  useEffect(() => {
    if (!integrating) return;
    let active = true;
    const startedAt = Date.now();
    const updateElapsed = () => setIntegrationElapsedSeconds(Math.floor((Date.now() - startedAt) / 1000));
    async function refreshProgress() {
      try {
        const nextProgress = await api.getIntegrationProgress(projectId);
        if (active) setIntegrationProgress(nextProgress);
      } catch {
        // The local stage indicator remains available if the progress endpoint is temporarily unavailable.
      }
    }
    updateElapsed();
    void refreshProgress();
    const elapsedTimer = window.setInterval(updateElapsed, 1000);
    const progressTimer = window.setInterval(() => { void refreshProgress(); }, 1200);
    return () => { active = false; window.clearInterval(elapsedTimer); window.clearInterval(progressTimer); };
  }, [integrating, projectId]);

  async function uploadMaster(selected: File[]) {
    const { accepted, rejected } = excelSelection(selected.slice(0, 1));
    setFailures(rejected);
    if (!accepted.length) return toast.error("请选择一份 Excel 总表");
    setMasterUploading(true);
    try {
      let uploaded: FileMeta;
      try {
        uploaded = await api.uploadFile(projectId, accepted[0], "template");
      } catch (error) {
        const failure = { filename: accepted[0].name, message: error instanceof Error ? error.message : "" };
        if (!isEncryptedWorkbookFailure(failure)) throw error;
        const password = promptForWorkbookPassword();
        if (password === undefined) return;
        uploaded = await api.uploadFile(projectId, accepted[0], "template", undefined, password);
      }
      setFiles((current) => mergeFileRecords(current, [uploaded]));
      setResult(null);
      setChangesConfirmed(false);
      toast.success("总表已上传，请继续上传变更文件");
    } catch (error) {
      const message = error instanceof Error ? error.message : "总表上传失败";
      setFailures([{ filename: accepted[0].name, message }]);
      toast.error(message);
    } finally {
      setMasterUploading(false);
    }
  }

  async function uploadAndMerge(selected: File[]) {
    const { accepted, rejected } = excelSelection(selected);
    if (accepted.length === 1 && !rejected.length) return uploadMaster(accepted);
    setFailures(rejected);
    if (rejected.length || accepted.length < 2) return toast.error("请选择总表和变更文件");
    setMasterUploading(true);
    try {
      let uploaded: FileMeta[];
      try {
        uploaded = await api.uploadFilesAuto(projectId, accepted);
      } catch (error) {
        const failure = { filename: "工作簿", message: error instanceof Error ? error.message : "" };
        if (!isEncryptedWorkbookFailure(failure)) throw error;
        const password = promptForWorkbookPassword();
        if (password === undefined) return;
        uploaded = await api.uploadFilesAuto(projectId, accepted, password);
      }
      const merged = mergeFileRecords(files, uploaded);
      setFiles(merged);
      setResult(null);
      const masterCount = merged.filter((file) => file.file_type === "template").length;
      const sourceCount = merged.filter((file) => file.file_type === "source").length;
      if (masterCount !== 1 || sourceCount === 0) throw new Error("需要一份包含工资核算表页的总表和至少一份变更文件");
      setChangesConfirmed(true);
      await generateWorkbook(true);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "自动合并失败");
    } finally {
      setMasterUploading(false);
    }
  }

  async function uploadChanges(selected: File[]) {
    const { accepted, rejected } = excelSelection(selected);
    setFailures(rejected);
    if (!accepted.length) return toast.error("请选择至少一份 Excel 变更文件");
    setChangesUploading(true);
    try {
      let uploadResult = await api.uploadFilesIndividually(projectId, accepted, "source");
      const encryptedFiles = accepted.filter((file) => uploadResult.failed.some((failure) => failure.filename === file.name && isEncryptedWorkbookFailure(failure)));
      if (encryptedFiles.length) {
        const password = promptForWorkbookPassword();
        if (password !== undefined) {
          const retry = await api.uploadFilesIndividually(projectId, encryptedFiles, "source", password);
          const encryptedNames = new Set(encryptedFiles.map((file) => file.name));
          uploadResult = {
            uploaded: [...uploadResult.uploaded, ...retry.uploaded],
            failed: [...uploadResult.failed.filter((failure) => !encryptedNames.has(failure.filename)), ...retry.failed],
          };
        }
      }
      const allFailures = [...rejected, ...uploadResult.failed];
      setFailures(allFailures);
      if (uploadResult.uploaded.length) {
        setFiles((current) => mergeFileRecords(current, uploadResult.uploaded));
        setResult(null);
        toast.success("变更文件已上传，还可以继续添加");
      }
      if (allFailures.length) toast.error(`${allFailures.length} 份文件未上传，请检查后重试`);
    } finally {
      setChangesUploading(false);
    }
  }

  async function removeFile(file: FileMeta) {
    if (!window.confirm(`确认移除“${file.original_name}”？`)) return;
    try {
      await api.deleteFile(projectId, file.id);
      setFiles((current) => current.filter((item) => item.id !== file.id));
      setResult(null);
      setChangesConfirmed(false);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "文件移除失败");
    }
  }

  async function generateWorkbook(prepared = false) {
    if (!prepared && (!master || !changes.length)) return;
    setIntegrationProgress(null);
    setIntegrationElapsedSeconds(0);
    setIntegrating(true);
    try {
      const integration = await api.integratePipeline(projectId);
      setResult(integration);
      if (integration.filename) {
        window.sessionStorage.removeItem(completionNoticeKey(projectId, integration.filename));
      }
      setCompletionOpen(true);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "更新失败，请检查上传文件");
    } finally {
      setIntegrating(false);
    }
  }

  function closeCompletionNotice() {
    setCompletionOpen(false);
    if (result?.filename) {
      window.sessionStorage.setItem(completionNoticeKey(projectId, result.filename), "dismissed");
    }
  }

  async function downloadWorkbook(kind: "formal" | "review") {
    const isReview = kind === "review";
    const filename = isReview ? result?.review_filename : result?.filename;
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
    <div className="payroll-page max-w-3xl pb-8">
      <header className="border-b border-slate-200/90 pb-7"><p className="payroll-kicker">{project?.salary_month || "本月"}工资总表</p><h1 className="page-heading mt-2">更新总表</h1><p className="mt-2 text-sm text-slate-600">按顺序完成下面 3 步即可。</p></header>

      <ol className="flex flex-col gap-2 border-y border-slate-200 py-3 text-sm sm:flex-row sm:items-center sm:gap-6" aria-label="更新步骤"><StepLabel current={step} step={1}>上传总表</StepLabel><StepLabel current={step} step={2}>上传变更文件</StepLabel><StepLabel current={step} step={3}>更新总表</StepLabel></ol>

      {result?.filename ? <section className="rounded-md border border-emerald-200 bg-emerald-50 p-4"><div className="flex items-start gap-3"><CheckCircle2 className="mt-0.5 h-5 w-5 shrink-0 text-emerald-700" /><div><h2 className="text-sm font-semibold text-emerald-950">最终总表已更新</h2><p className="mt-1 text-xs text-emerald-800">{result.issue_count ? `还有 ${result.issue_count} 项需要人工填写。` : "没有需要人工填写的事项。"}</p></div></div><div className="mt-3 flex flex-col gap-2 border-t border-emerald-200 pt-3 sm:flex-row"><Button type="button" size="sm" variant="outline" className="border-teal-300 bg-white text-teal-900 hover:bg-teal-50" disabled={downloading !== null} onClick={() => void downloadWorkbook("formal")}>{downloading === "formal" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载正式稿</Button>{result.review_filename ? <Button type="button" size="sm" variant="outline" className="border-amber-300 bg-white text-amber-950 hover:bg-amber-50" disabled={downloading !== null} onClick={() => void downloadWorkbook("review")}>{downloading === "review" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载修改稿（含变更批注）</Button> : null}</div></section> : null}

      {step === 1 ? <WorkbookUploadPanel title="上传文件，自动合并" description="同时选择文件夹中的总表和新增人员表，系统自动识别、匹配并生成结果。也可以先上传一份总表。" helper="选择总表和变更文件" files={masters} multiple uploading={masterUploading} disabled={integrating} tone="master" onSelect={(selected) => void uploadAndMerge(selected)} onRemove={(file) => void removeFile(file)} /> : null}

      {step === 2 ? <><section className="flex items-center justify-between gap-3 rounded-md border border-slate-200 bg-slate-50 px-4 py-3"><p className="min-w-0 truncate text-sm text-slate-700">已上传总表：<span className="font-medium">{master?.original_name}</span></p><Button type="button" size="sm" variant="outline" onClick={() => master && void removeFile(master)}>更换</Button></section><WorkbookUploadPanel title="第 2 步：上传变更文件" description="可多次添加，也可一次选择多个文件。" helper="选择变更文件" files={changes} multiple uploading={changesUploading} tone="changes" onSelect={(selected) => void uploadChanges(selected)} onRemove={(file) => void removeFile(file)} /><section className="flex flex-col gap-3 rounded-md border border-slate-200 bg-white p-4 sm:flex-row sm:items-center sm:justify-between"><p className="text-sm text-slate-600">{changes.length ? `已添加 ${changes.length} 份变更文件，确认后再开始更新。` : "请先添加至少 1 份变更文件。"}</p><Button type="button" disabled={!changes.length || changesUploading} onClick={() => setChangesConfirmed(true)} className="bg-teal-700 hover:bg-teal-800">确认变更文件（{changes.length}）</Button></section></> : null}

      {step === 3 ? <section className="rounded-md border border-slate-200 bg-white p-5"><h2 className="text-base font-semibold text-slate-950">第 3 步：开始更新</h2><p className="mt-2 text-sm text-slate-600">总表和已确认的 {changes.length} 份变更文件已准备好。</p>{integrating ? <IntegrationLoadingCard progress={integrationProgress} elapsedSeconds={integrationElapsedSeconds} estimatedSeconds={estimatedIntegrationSeconds} fileCount={changes.length} /> : null}<div className="mt-4 flex flex-col gap-2 sm:flex-row"><Button onClick={() => void generateWorkbook()} disabled={integrating || masterUploading || changesUploading} className="bg-teal-700 hover:bg-teal-800">{integrating ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}{integrating ? "正在更新…" : "开始更新总表"}</Button><Button type="button" variant="outline" disabled={integrating} onClick={() => setChangesConfirmed(false)}>继续添加变更文件</Button></div></section> : null}

      {failures.length ? <section className="rounded-md border border-amber-200 bg-amber-50 px-4 py-3" role="alert"><div className="flex items-center gap-2"><AlertTriangle className="h-4 w-4 text-amber-700" /><h2 className="text-sm font-semibold text-amber-950">有文件未上传</h2></div><ul className="mt-2 space-y-1 text-xs text-amber-900">{failures.map((failure) => <li key={`${failure.filename}-${failure.message}`}>{failure.filename}：{failure.message}</li>)}</ul></section> : null}

      {personnelCoverage ? <section className="rounded-md border border-teal-200 bg-teal-50 p-4 text-sm text-teal-950"><p>{completionMessage}</p><p className="mt-2 text-xs">{result?.validation.message}</p></section> : null}

      <Dialog open={completionOpen} onOpenChange={(open) => { if (open) setCompletionOpen(true); else closeCompletionNotice(); }}><DialogContent><DialogHeader><DialogTitle>总表已更新</DialogTitle><DialogDescription>{completionMessage} 修改稿保留逐项变更批注。{personnelCoverage ? "本次更新人员信息与薪资标准。" : ""}</DialogDescription></DialogHeader><DialogFooter><Button type="button" variant="outline" onClick={closeCompletionNotice}>关闭</Button></DialogFooter></DialogContent></Dialog>
    </div>
  );
}
