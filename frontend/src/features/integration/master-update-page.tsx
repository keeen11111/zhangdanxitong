"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { useRouter } from "next/navigation";
import { AlertTriangle, CheckCircle2, Clock3, Download, Loader2 } from "lucide-react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { api, FileMeta, IntegrationResult, PipelineExportResult, Project, UploadFailure, WorkbookAnalysis } from "@/lib/api";
import { mergeFileRecords } from "@/lib/file-list";
import { estimateIntegrationSeconds, formatEstimateRange } from "@/lib/integration-estimate";
import { MasterWorkbookAnalysis } from "./master-workbook-analysis";
import { WorkbookUploadPanel } from "./workbook-upload-panel";
import { downloadWorkbookDirect } from "@/lib/download";

type Result = IntegrationResult | PipelineExportResult;

function excelSelection(selected: File[]) {
  const accepted = selected.filter((file) => /\.(xlsx|xls)$/i.test(file.name));
  const rejected = selected
    .filter((file) => !/\.(xlsx|xls)$/i.test(file.name))
    .map((file) => ({ filename: file.name, message: "仅支持 .xlsx 或 .xls 文件" }));
  return { accepted, rejected };
}

export function MasterUpdatePage({ projectId }: { projectId: string }) {
  const router = useRouter();
  const [project, setProject] = useState<Project | null>(null);
  const [files, setFiles] = useState<FileMeta[]>([]);
  const [result, setResult] = useState<Result | null>(null);
  const [masterUploading, setMasterUploading] = useState(false);
  const [changesUploading, setChangesUploading] = useState(false);
  const [integrating, setIntegrating] = useState(false);
  const [elapsedSeconds, setElapsedSeconds] = useState(0);
  const [failures, setFailures] = useState<UploadFailure[]>([]);
  const [analysis, setAnalysis] = useState<WorkbookAnalysis | null>(null);
  const [analysisLoading, setAnalysisLoading] = useState(false);
  const [completionOpen, setCompletionOpen] = useState(false);
  const [downloading, setDownloading] = useState<"formal" | "review" | null>(null);

  const masters = useMemo(() => files.filter((file) => file.file_type === "template"), [files]);
  const master = masters.length === 1 ? masters[0] : null;
  const changes = useMemo(() => files.filter((file) => file.file_type === "source"), [files]);
  const estimatedSeconds = useMemo(() => {
    const masterRows = Math.max(0, ...(analysis?.sheets.map((sheet) => sheet.row_count) || [0]));
    const masterColumns = Math.max(0, ...(analysis?.sheets.map((sheet) => sheet.col_count) || [0]));
    return estimateIntegrationSeconds({
      masterRows,
      masterColumns,
      masterSheets: analysis?.sheet_count || 1,
      formulaCount: analysis?.formula_count || 0,
      changeCells: changes.reduce((total, file) => total + file.row_count * file.col_count, 0),
      changeFiles: changes.length,
    });
  }, [analysis, changes]);

  const loadWorkspace = useCallback(async () => {
    const [projectResult, fileResult, exportResult] = await Promise.allSettled([
      api.getProject(projectId),
      api.listFiles(projectId),
      api.getLatestPipelineExport(projectId),
    ]);
    if (projectResult.status === "fulfilled") setProject(projectResult.value);
    if (fileResult.status === "fulfilled") setFiles(fileResult.value);
    setResult(exportResult.status === "fulfilled" ? exportResult.value : null);
  }, [projectId]);

  useEffect(() => { void loadWorkspace(); }, [loadWorkspace]);
  useEffect(() => {
    if (!master) {
      setAnalysis(null);
      return;
    }
    let active = true;
    setAnalysisLoading(true);
    api.analyzeFile(projectId, master.id)
      .then((value) => { if (active) setAnalysis(value); })
      .catch((error) => { if (active) toast.error(error instanceof Error ? error.message : "总表分析失败"); })
      .finally(() => { if (active) setAnalysisLoading(false); });
    return () => { active = false; };
  }, [master, projectId]);
  useEffect(() => {
    if (!integrating) return;
    setElapsedSeconds(0);
    const timer = window.setInterval(() => setElapsedSeconds((value) => value + 1), 1000);
    return () => window.clearInterval(timer);
  }, [integrating]);

  async function uploadMaster(selected: File[]) {
    const { accepted, rejected } = excelSelection(selected.slice(0, 1));
    setFailures(rejected);
    if (!accepted.length) return toast.error("请选择一份 Excel 总表");
    setMasterUploading(true);
    try {
      const uploaded = await api.uploadFile(projectId, accepted[0], "template");
      setFiles((current) => mergeFileRecords(current, [uploaded]));
      setResult(null);
      const details = [
        uploaded.normalization_note,
        uploaded.processing_seconds ? `耗时 ${uploaded.processing_seconds.toFixed(1)} 秒` : null,
      ].filter(Boolean).join("，");
      toast.success(details ? `总表已上传，${details}` : "总表已上传并通过结构校验");
    } catch (error) {
      const message = error instanceof Error ? error.message : "总表上传失败";
      setFailures([{ filename: accepted[0].name, message }]);
      toast.error(message);
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
      const uploadResult = await api.uploadFilesIndividually(projectId, accepted, "source");
      const allFailures = [...rejected, ...uploadResult.failed];
      setFailures(allFailures);
      if (uploadResult.uploaded.length) {
        setFiles((current) => mergeFileRecords(current, uploadResult.uploaded));
        setResult(null);
        toast.success(`已添加 ${uploadResult.uploaded.length} 份变更文件`);
      }
      if (allFailures.length) toast.error(`${allFailures.length} 份文件未上传，请查看原因`);
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
      toast.success("文件已移除");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "文件移除失败");
    }
  }

  async function generateWorkbook() {
    if (!master) return toast.error("请先上传总表");
    if (!changes.length) return toast.error("请至少上传一份变更文件");
    setIntegrating(true);
    try {
      const integration = await api.integratePipeline(projectId);
      setResult(integration);
      setCompletionOpen(true);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "更新失败，请检查上传文件");
    } finally {
      setIntegrating(false);
    }
  }

  const openResult = () => router.push(`/projects/${projectId}/result`);
  async function downloadFormal() {
    if (!result?.filename) return;
    setDownloading("formal");
    try {
      await downloadWorkbookDirect(api.pipelineExportDownloadUrl(projectId), result.filename);
      toast.success("正式稿已开始下载");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "正式稿下载失败");
    } finally {
      setDownloading(null);
    }
  }

  async function downloadReview() {
    if (!result) return;
    setDownloading("review");
    try {
      await downloadWorkbookDirect(
        api.reviewExportDownloadUrl(projectId),
        result.review_filename || `修改稿_${result.filename}`,
      );
      toast.success("修改稿已开始下载");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "修改稿下载失败");
    } finally {
      setDownloading(null);
    }
  }
  const ready = Boolean(masters.length === 1 && changes.length);

  return (
    <div className="mx-auto max-w-6xl space-y-5 pb-8">
      <header><p className="text-sm text-slate-500">{project?.salary_month || "本月"}工资总表更新</p><h1 className="mt-1 text-2xl font-semibold tracking-tight text-slate-950">用变更文件更新现有总表</h1><p className="mt-2 max-w-3xl text-sm leading-6 text-slate-600">总表是唯一基准；变更文件只用于更新明确的数据。人员冲突、部门转移、临时公式和特殊事项不会被系统猜测，会单独导出待人工处理清单。</p></header>

      {result?.filename ? <section className="flex flex-col gap-3 rounded-xl border border-emerald-200 bg-emerald-50/60 p-4"><div className="flex items-start gap-3"><CheckCircle2 className="mt-0.5 h-5 w-5 text-emerald-700" /><div><p className="text-sm font-semibold text-emerald-950">已有更新结果</p><p className="mt-1 text-xs text-emerald-800">正式稿和修改稿已生成 · {result.issue_count} 项待人工处理</p><p className="mt-1 break-all text-xs text-slate-600">正式稿：{result.filename}</p><p className="break-all text-xs text-slate-600">修改稿：{result.review_filename || "点击下载时自动生成"}</p></div></div><div className="flex flex-wrap gap-2"><Button onClick={() => void downloadFormal()} disabled={downloading === "formal"} className="bg-sky-700 hover:bg-sky-800">{downloading === "formal" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载正式稿</Button><Button onClick={() => void downloadReview()} disabled={downloading === "review"} className="border border-amber-300 bg-amber-700 text-white hover:bg-amber-800 disabled:bg-amber-700 disabled:text-white disabled:opacity-75">{downloading === "review" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载修改稿</Button><Button onClick={openResult} variant="outline">查看结果与人工处理</Button></div></section> : null}

      <div className="grid gap-5 lg:grid-cols-2">
        <WorkbookUploadPanel title="上传总表" description="只上传 1 份需要被更新的完整总表。系统会校验“工资核算”Sheet，并保留其格式、公式和其他 Sheet。" helper="拖入总表，或点击选择" files={masters} uploading={masterUploading} disabled={masters.length > 0} tone="master" onSelect={(selected) => void uploadMaster(selected)} onRemove={(file) => void removeFile(file)} />
        <WorkbookUploadPanel title="上传变更文件" description="上传本月新增或变化的 Sheet/明细文件，可多选。它们不会取代总表，只提供更新依据。" helper="拖入全部变更文件，或点击多选" files={changes} multiple uploading={changesUploading} tone="changes" onSelect={(selected) => void uploadChanges(selected)} onRemove={(file) => void removeFile(file)} />
      </div>

      {masters.length > 1 ? <section className="rounded-xl border border-red-200 bg-red-50 px-4 py-3" role="alert"><div className="flex items-center gap-2"><AlertTriangle className="h-4 w-4 text-red-700" /><p className="text-sm font-semibold text-red-950">检测到 {masters.length} 份总表</p></div><p className="mt-1 text-xs leading-5 text-red-800">请保留唯一一份需要更新的总表，其余文件移除后才能开始处理。</p></section> : null}

      <MasterWorkbookAnalysis analysis={analysis} loading={analysisLoading} />

      {failures.length ? <section className="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3" role="alert"><div className="flex items-center gap-2"><AlertTriangle className="h-4 w-4 text-amber-700" /><p className="text-sm font-semibold text-amber-950">以下文件未上传</p></div><ul className="mt-2 space-y-1 text-xs leading-5 text-amber-900">{failures.map((failure) => <li key={`${failure.filename}-${failure.message}`}><span className="font-medium">{failure.filename}</span>：{failure.message}</li>)}</ul></section> : null}

      <section className="flex flex-col gap-4 rounded-xl border border-slate-200 bg-white p-4 sm:flex-row sm:items-center sm:justify-between sm:p-5"><div><p className="text-sm font-semibold text-slate-950">人员核对与总表更新</p><p className="mt-1 flex items-center gap-1.5 text-xs text-slate-500"><Clock3 className="h-3.5 w-3.5" />{integrating ? `已用 ${elapsedSeconds} 秒 · 预计 ${formatEstimateRange(estimatedSeconds)}` : `预计用时：${formatEstimateRange(estimatedSeconds)}`}</p><p className="mt-1 text-xs text-slate-500">系统更新奖金、考勤、值班等明确数据；公式、其他累计调差和特殊人员事项请按清单人工完成。</p></div><Button onClick={() => void generateWorkbook()} disabled={!ready || integrating || masterUploading || changesUploading} className="w-full min-w-44 bg-sky-700 hover:bg-sky-800 sm:w-auto">{integrating && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}{integrating ? "正在更新总表…" : "开始核对并更新"}</Button></section>

      <Dialog open={completionOpen} onOpenChange={setCompletionOpen}><DialogContent><DialogHeader><DialogTitle>基础数据更新完成</DialogTitle><DialogDescription>更新后的总表和 {result?.issue_count || 0} 项人工处理记录已准备好{result?.duration_seconds ? `，实际耗时 ${result.duration_seconds.toFixed(1)} 秒` : ""}。公式与特殊事项尚未自动修改，请到成果页下载人工清单后再定稿。</DialogDescription></DialogHeader><DialogFooter><Button variant="outline" onClick={() => setCompletionOpen(false)}>稍后查看</Button><Button onClick={openResult} className="bg-sky-700 hover:bg-sky-800">查看更新结果</Button></DialogFooter></DialogContent></Dialog>
    </div>
  );
}
