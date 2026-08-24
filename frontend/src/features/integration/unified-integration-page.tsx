"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useRouter } from "next/navigation";
import { CheckCircle2, FileSpreadsheet, Loader2, Search, Trash2, Upload } from "lucide-react";
import { toast } from "sonner";

import { api, FileMeta, IntegrationResult, PipelineExportResult, Project, UploadFailure } from "@/lib/api";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { mergeFileRecords } from "@/lib/file-list";
import { cn } from "@/lib/utils";

type Result = IntegrationResult | PipelineExportResult;

const filenameCollator = new Intl.Collator("zh-CN", { numeric: true, sensitivity: "base" });

function fileSummary(file: FileMeta) {
  return !file.row_count && !file.col_count ? "图片 / 图表附件" : `${file.row_count} 行 · ${file.col_count} 列`;
}

export function UnifiedIntegrationPage({ projectId }: { projectId: string }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const router = useRouter();
  const [project, setProject] = useState<Project | null>(null);
  const [files, setFiles] = useState<FileMeta[]>([]);
  const [result, setResult] = useState<Result | null>(null);
  const [uploading, setUploading] = useState(false);
  const [integrating, setIntegrating] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [uploadFailures, setUploadFailures] = useState<UploadFailure[]>([]);
  const [query, setQuery] = useState("");
  const [completionOpen, setCompletionOpen] = useState(false);

  const visibleFiles = useMemo(() => {
    const keyword = query.trim().toLocaleLowerCase("zh-CN");
    return [...files]
      .sort((left, right) => filenameCollator.compare(left.original_name, right.original_name))
      .filter((file) => !keyword || file.original_name.toLocaleLowerCase("zh-CN").includes(keyword));
  }, [files, query]);

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

  async function uploadFiles(selected: File[]) {
    const excelFiles = selected.filter((file) => /\.(xlsx|xls)$/i.test(file.name));
    const unsupportedFiles = selected.filter((file) => !/\.(xlsx|xls)$/i.test(file.name))
      .map((file) => ({ filename: file.name, message: "仅支持 .xlsx 或 .xls 文件" }));
    if (!excelFiles.length) {
      toast.error("请选择 Excel 文件（.xlsx 或 .xls）");
      setUploadFailures(unsupportedFiles);
      return;
    }
    setUploading(true);
    setUploadFailures([]);
    try {
      const uploadResult = await api.uploadFilesIndividually(projectId, excelFiles);
      const failures = [...unsupportedFiles, ...uploadResult.failed];
      setUploadFailures(failures);
      if (uploadResult.uploaded.length) {
        setResult(null);
        // Count every successful batch immediately. The server refresh below
        // remains the final authority, but a slow refresh must not hide the
        // second (or later) upload from the user.
        setFiles((current) => mergeFileRecords(current, uploadResult.uploaded));
        await loadWorkspace();
        toast.success(`已添加 ${uploadResult.uploaded.length} 个文件`);
      }
      if (failures.length) toast.error(`${failures.length} 个文件未上传，请查看原因`);
    } finally {
      setUploading(false);
    }
  }

  async function deleteFile(file: FileMeta) {
    if (!window.confirm(`确认移除“${file.original_name}”？移除后需要重新整合。`)) return;
    try {
      await api.deleteFile(projectId, file.id);
      setFiles((current) => current.filter((item) => item.id !== file.id));
      setResult(null);
      toast.success(`已移除 ${file.original_name}`);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "文件移除失败");
    }
  }

  async function generateWorkbook() {
    if (!files.length) return toast.error("请至少加入一个 Excel 文件");
    setIntegrating(true);
    try {
      const integration = await api.integratePipeline(projectId);
      setResult(integration);
      setCompletionOpen(true);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "整合失败，请检查上传文件");
    } finally {
      setIntegrating(false);
    }
  }

  const openResult = () => router.push(`/projects/${projectId}/result`);

  return (
    <div className="mx-auto max-w-6xl space-y-5 pb-8">
      <header>
        <p className="text-sm text-slate-500">{project?.salary_month || "本月"}工资整合</p>
        <h1 className="mt-1 text-2xl font-semibold tracking-tight text-slate-950">生成统一工资核算总表</h1>
        <p className="mt-2 max-w-3xl text-sm leading-6 text-slate-600">加入本月全部 Excel 后开始整合。制作完成后在成果页查收总表和单独的人工处理表。</p>
      </header>

      {result?.filename ? (
        <section className="flex flex-col gap-3 rounded-xl border border-emerald-200 bg-emerald-50/60 p-4 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-start gap-3">
            <CheckCircle2 className="mt-0.5 h-5 w-5 shrink-0 text-emerald-700" />
            <div><p className="text-sm font-semibold text-emerald-950">已有制作结果</p><p className="mt-1 text-xs text-emerald-800">{result.filename} · {result.issue_count} 项待人工处理</p></div>
          </div>
          <Button onClick={openResult} className="bg-sky-700 hover:bg-sky-800">查看制作结果</Button>
        </section>
      ) : null}

      <section className="rounded-xl border border-slate-200 bg-white p-4 sm:p-5">
        <input ref={inputRef} className="hidden" type="file" accept=".xlsx,.xls" multiple onChange={(event) => { void uploadFiles(Array.from(event.target.files || [])); event.target.value = ""; }} />
        <button
          type="button" disabled={uploading} onClick={() => inputRef.current?.click()}
          onDragEnter={(event) => { event.preventDefault(); setDragging(true); }} onDragOver={(event) => event.preventDefault()} onDragLeave={() => setDragging(false)}
          onDrop={(event) => { event.preventDefault(); setDragging(false); void uploadFiles(Array.from(event.dataTransfer.files)); }}
          className={cn("flex w-full items-center justify-center gap-3 rounded-lg border border-dashed px-5 text-left transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-sky-700 focus-visible:ring-offset-2 disabled:cursor-wait disabled:opacity-70", files.length ? "min-h-20" : "min-h-36", dragging ? "border-sky-600 bg-sky-50" : "border-slate-300 bg-slate-50 hover:border-sky-500 hover:bg-sky-50/60")}
        >
          {uploading ? <Loader2 className="h-5 w-5 shrink-0 animate-spin text-sky-700" /> : <Upload className="h-5 w-5 shrink-0 text-sky-700" />}
          <span><span className="block text-sm font-medium text-slate-900">{uploading ? "正在添加文件…" : "拖入全部 Excel，或点击选择"}</span><span className="mt-0.5 block text-xs text-slate-500">支持一次选择多个 .xlsx / .xls 文件</span></span>
        </button>

        {files.length ? <div className="mt-4 flex flex-col gap-3 rounded-lg border border-slate-200 bg-slate-50 px-4 py-3 sm:flex-row sm:items-center sm:justify-between"><p className="text-xs leading-5 text-slate-600"><span className="font-medium text-slate-900">{files.length} 个文件已就绪</span> · 完成后进入独立成果页查收</p><Button onClick={generateWorkbook} disabled={integrating || uploading} className="w-full min-w-36 bg-sky-700 hover:bg-sky-800 sm:w-auto">{integrating && <Loader2 className="mr-2 h-4 w-4 animate-spin" />}{integrating ? "正在整合…" : "开始整合"}</Button></div> : null}

        <div className="mt-4" aria-live="polite">
          <div className="flex flex-col gap-3 border-b border-slate-200 pb-3 sm:flex-row sm:items-center sm:justify-between"><div className="flex items-baseline gap-2"><h2 className="text-sm font-medium text-slate-900">已添加文件</h2><span className="text-xs tabular-nums text-slate-500">{files.length} 个</span></div>{files.length > 6 ? <label className="relative block sm:w-64"><Search className="pointer-events-none absolute left-3 top-2.5 h-4 w-4 text-slate-400" /><Input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="搜索文件名" aria-label="搜索文件名" className="h-9 bg-white pl-9" /></label> : null}</div>
          {files.length ? <div className="max-h-[420px] overflow-auto"><table className="w-full table-fixed text-left text-sm"><thead className="sticky top-0 z-10 bg-white text-xs text-slate-500"><tr className="border-b border-slate-200"><th className="w-auto py-2 pr-4 font-medium">文件</th><th className="hidden w-28 py-2 pr-4 font-medium sm:table-cell">用途</th><th className="hidden w-36 py-2 pr-4 font-medium md:table-cell">数据概览</th><th className="w-12 py-2 text-right font-medium"><span className="sr-only">操作</span></th></tr></thead><tbody className="divide-y divide-slate-100">{visibleFiles.map((file) => <tr key={file.id} className="hover:bg-slate-50/80"><td className="py-2.5 pr-4"><div className="flex min-w-0 items-center gap-2.5"><FileSpreadsheet className="h-4 w-4 shrink-0 text-slate-400" /><span className="truncate text-slate-800" title={file.original_name}>{file.original_name}</span></div><p className="mt-1 pl-6 text-xs text-slate-500 sm:hidden">{file.file_type === "template" ? "总表底板" : "来源明细"} · {fileSummary(file)}</p></td><td className="hidden py-2.5 pr-4 text-xs text-slate-600 sm:table-cell">{file.file_type === "template" ? "总表底板" : "来源明细"}</td><td className="hidden py-2.5 pr-4 text-xs tabular-nums text-slate-600 md:table-cell">{fileSummary(file)}</td><td className="py-2.5 text-right"><button type="button" onClick={() => void deleteFile(file)} className="inline-flex h-8 w-8 items-center justify-center rounded-md text-slate-400 hover:bg-red-50 hover:text-red-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-600" aria-label={`移除 ${file.original_name}`}><Trash2 className="h-4 w-4" /></button></td></tr>)}</tbody></table>{!visibleFiles.length ? <p className="py-8 text-center text-sm text-slate-500">没有匹配的文件</p> : null}</div> : <p className="py-5 text-sm text-slate-500">尚未添加文件。</p>}
        </div>
        {uploadFailures.length ? <div className="mt-4 rounded-lg border border-amber-200 bg-amber-50 px-4 py-3" role="alert"><p className="text-sm font-medium text-amber-950">以下文件未上传，请处理后重新选择</p><ul className="mt-2 space-y-1 text-xs leading-5 text-amber-900">{uploadFailures.map((failure) => <li key={`${failure.filename}-${failure.message}`}><span className="font-medium">{failure.filename}</span>：{failure.message}</li>)}</ul></div> : null}
      </section>

      <Dialog open={completionOpen} onOpenChange={setCompletionOpen}>
        <DialogContent>
          <DialogHeader><DialogTitle>制作完成</DialogTitle><DialogDescription>总表和{result?.issue_count || 0}项人工处理记录已准备好，可在成果页预览、改名、直接下载或选择保存位置。</DialogDescription></DialogHeader>
          <DialogFooter><Button variant="outline" onClick={() => setCompletionOpen(false)}>稍后查看</Button><Button onClick={openResult} className="bg-sky-700 hover:bg-sky-800">查看制作结果</Button></DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
