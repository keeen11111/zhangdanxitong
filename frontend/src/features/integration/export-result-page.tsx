"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import Link from "next/link";
import {
  AlertTriangle,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  Download,
  Eye,
  FileSpreadsheet,
  Loader2,
  Save,
  Upload,
} from "lucide-react";
import { toast } from "sonner";

import { api, PipelineExportPreview, PipelineExportResult, Project } from "@/lib/api";
import { downloadWorkbookDirect } from "@/lib/download";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { cn } from "@/lib/utils";

function WorkbookPreview({ preview }: { preview: PipelineExportPreview | null }) {
  if (!preview) return null;
  return (
    <div className="mt-4 overflow-x-auto rounded-lg border border-slate-200 bg-white">
      <p className="border-b border-slate-200 bg-slate-50 px-4 py-2 text-xs text-slate-600">{preview.filename} · {preview.sheet_name} · 前 {preview.rows.length} 行</p>
      <table className="min-w-full text-left text-xs text-slate-700">
        <tbody>
          {preview.rows.map((row, rowIndex) => (
            <tr key={rowIndex} className="border-b border-slate-100 last:border-0">
              {row.map((value, columnIndex) => <td key={columnIndex} className="whitespace-nowrap px-3 py-2">{value == null ? "" : String(value)}</td>)}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export function ExportResultPage({ projectId }: { projectId: string }) {
  const [project, setProject] = useState<Project | null>(null);
  const [result, setResult] = useState<PipelineExportResult | null>(null);
  const [filenameDraft, setFilenameDraft] = useState("");
  const [renaming, setRenaming] = useState(false);
  const [masterPreview, setMasterPreview] = useState<PipelineExportPreview | null>(null);
  const [reviewPreview, setReviewPreview] = useState<PipelineExportPreview | null>(null);
  const [issuesPreview, setIssuesPreview] = useState<PipelineExportPreview | null>(null);
  const [previewing, setPreviewing] = useState<"master" | "review" | "issues" | null>(null);
  const [downloading, setDownloading] = useState<"master" | "review" | "issues" | null>(null);
  const [issuesOpen, setIssuesOpen] = useState(false);
  const [selectedDepartmentIssueIds, setSelectedDepartmentIssueIds] = useState<string[]>([]);
  const [confirmingDepartmentTransfers, setConfirmingDepartmentTransfers] = useState(false);
  const [issueActions, setIssueActions] = useState<Record<string, string>>({});
  const [resolvingIssueId, setResolvingIssueId] = useState<string | null>(null);
  const [manualReviewFile, setManualReviewFile] = useState<File | null>(null);
  const [importingManualReview, setImportingManualReview] = useState(false);
  const manualReviewFileInputRef = useRef<HTMLInputElement>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const loadResult = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const [loadedProject, loadedResult] = await Promise.all([
        api.getProject(projectId),
        api.getLatestPipelineExport(projectId),
      ]);
      setProject(loadedProject);
      setResult(loadedResult);
      setFilenameDraft(loadedResult.filename);
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : "尚未找到可查收的制作结果");
    } finally {
      setLoading(false);
    }
  }, [projectId]);

  useEffect(() => { void loadResult(); }, [loadResult]);

  async function renameWorkbook() {
    if (!result || !filenameDraft.trim() || filenameDraft.trim() === result.filename) return;
    setRenaming(true);
    try {
      const renamed = await api.renamePipelineExport(projectId, filenameDraft);
      setResult((current) => current ? { ...current, filename: renamed.filename } : current);
      setFilenameDraft(renamed.filename);
      setMasterPreview((current) => current ? { ...current, filename: renamed.filename } : current);
      toast.success("总表文件名已保存");
    } catch (renameError) {
      toast.error(renameError instanceof Error ? renameError.message : "文件名保存失败");
    } finally {
      setRenaming(false);
    }
  }

  async function previewMaster() {
    setPreviewing("master");
    try {
      setMasterPreview(await api.previewPipelineExport(projectId));
    } catch (previewError) {
      toast.error(previewError instanceof Error ? previewError.message : "总表预览失败");
    } finally {
      setPreviewing(null);
    }
  }

  async function toggleIssues() {
    const nextOpen = !issuesOpen;
    setIssuesOpen(nextOpen);
    if (!nextOpen || issuesPreview) return;
    await previewIssues();
  }

  async function previewReview() {
    setPreviewing("review");
    try {
      setReviewPreview(await api.previewReviewExport(projectId));
    } catch (previewError) {
      toast.error(previewError instanceof Error ? previewError.message : "修改稿预览失败");
    } finally {
      setPreviewing(null);
    }
  }

  async function previewIssues() {
    setPreviewing("issues");
    try {
      setIssuesPreview(await api.previewManualIssuesExport(projectId));
    } catch (previewError) {
      toast.error(previewError instanceof Error ? previewError.message : "人工处理表预览失败");
    } finally {
      setPreviewing(null);
    }
  }

  async function downloadMasterWorkbook() {
    if (!result) return;
    setDownloading("master");
    try {
      await downloadWorkbookDirect(api.pipelineExportDownloadUrl(projectId), result.filename);
      toast.success("总表已开始下载");
    } catch (downloadError) {
      toast.error(downloadError instanceof Error ? downloadError.message : "下载失败");
    } finally {
      setDownloading(null);
    }
  }

  async function downloadManualIssuesWorkbook() {
    if (!result) return;
    setDownloading("issues");
    try {
      await downloadWorkbookDirect(
        api.manualIssuesExportDownloadUrl(projectId),
        result.issues_filename || "待人工处理.xlsx",
      );
      toast.success("待人工处理表已开始下载");
    } catch (downloadError) {
      toast.error(downloadError instanceof Error ? downloadError.message : "下载失败");
    } finally {
      setDownloading(null);
    }
  }

  async function downloadReviewWorkbook() {
    if (!result) return;
    setDownloading("review");
    try {
      await downloadWorkbookDirect(
        api.reviewExportDownloadUrl(projectId),
        result.review_filename || `修改稿_${result.filename}`,
      );
      toast.success("修改稿已开始下载");
    } catch (downloadError) {
      toast.error(downloadError instanceof Error ? downloadError.message : "修改稿下载失败");
    } finally {
      setDownloading(null);
    }
  }

  function toggleDepartmentIssue(issueId: string) {
    setSelectedDepartmentIssueIds((current) => current.includes(issueId)
      ? current.filter((id) => id !== issueId)
      : [...current, issueId]);
  }

  async function confirmSelectedDepartmentTransfers() {
    if (!selectedDepartmentIssueIds.length) return;
    setConfirmingDepartmentTransfers(true);
    try {
      const updated = await api.confirmDepartmentTransfers(projectId, selectedDepartmentIssueIds);
      setResult(updated);
      setFilenameDraft(updated.filename);
      setSelectedDepartmentIssueIds([]);
      setIssuesPreview(null);
      toast.success("已确认部门变更，并生成新的工资核算总表");
    } catch (confirmError) {
      toast.error(confirmError instanceof Error ? confirmError.message : "部门变更确认失败");
    } finally {
      setConfirmingDepartmentTransfers(false);
    }
  }

  async function resolveIssue(issueId: string, action: "apply_proposed" | "keep_current") {
    setResolvingIssueId(issueId);
    try {
      const updated = await api.resolveManualIssue(projectId, issueId, action);
      setResult(updated);
      setFilenameDraft(updated.filename);
      setIssuesPreview(null);
      toast.success(action === "apply_proposed" ? "已采用推荐值并重新生成结果" : "已保留原值并关闭该事项");
    } catch (resolveError) {
      toast.error(resolveError instanceof Error ? resolveError.message : "人工事项处理失败");
    } finally {
      setResolvingIssueId(null);
    }
  }

  async function importManualReviewWorkbook() {
    if (!manualReviewFile) return;
    setImportingManualReview(true);
    try {
      const updated = await api.importManualReviewWorkbook(projectId, manualReviewFile);
      setResult(updated);
      setFilenameDraft(updated.filename);
      setManualReviewFile(null);
      if (manualReviewFileInputRef.current) manualReviewFileInputRef.current.value = "";
      setMasterPreview(null);
      setReviewPreview(null);
      setIssuesPreview(null);
      toast.success("人工处理结果已更新到总表，并生成最新文件");
    } catch (importError) {
      toast.error(importError instanceof Error ? importError.message : "人工处理表导入失败");
    } finally {
      setImportingManualReview(false);
    }
  }

  if (loading) return <div className="mx-auto max-w-6xl py-16 text-center text-sm text-slate-500"><Loader2 className="mx-auto mb-3 h-5 w-5 animate-spin" />正在加载制作结果…</div>;
  if (!result) return <div className="mx-auto max-w-3xl rounded-xl border border-slate-200 bg-white p-6"><h1 className="text-xl font-semibold text-slate-950">暂未找到制作结果</h1><p className="mt-2 text-sm text-slate-600">{error || "请返回整合页，完成本月文件整合后再查收。"}</p><Button asChild className="mt-5 bg-sky-700 hover:bg-sky-800"><Link href={`/projects/${projectId}`}>返回整合页</Link></Button></div>;

  const peopleBalanced = result.coverage_gap_count === 0
    && result.input_person_count === result.output_person_count + result.manual_only_person_count;

  return (
    <div className="mx-auto max-w-6xl space-y-5 pb-8">
      <header>
        <Link href={`/projects/${projectId}`} className="text-sm text-slate-500 hover:text-slate-900">← 返回整合页</Link>
        <p className="mt-5 text-sm text-slate-500">{project?.salary_month || "本月"}工资交付</p>
        <h1 className="mt-1 text-2xl font-semibold tracking-tight text-slate-950">基础数据更新完成</h1>
        <p className="mt-2 text-sm leading-6 text-slate-600">正式稿用于业务使用，修改稿用于人工审核。修改稿与正式稿内容一致，在变更单元格添加批注，并在“修改记录”表逐项列出原值、新值和位置。</p>
      </header>

      <section className="rounded-xl border border-slate-200 bg-white p-4 sm:p-5">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-2"><CheckCircle2 className="h-5 w-5 text-emerald-700" /><h2 className="text-base font-semibold text-slate-950">工资核算总表</h2></div>
            <p className="mt-2 break-all text-xs text-slate-600">正式稿：{result.filename}</p>
            <p className="break-all text-xs text-slate-600">修改稿：{result.review_filename || "生成中"}</p>
            <div className="mt-3 flex max-w-2xl flex-col gap-2 sm:flex-row">
              <Input value={filenameDraft} onChange={(event) => setFilenameDraft(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") void renameWorkbook(); }} aria-label="总表文件名" className="h-9 bg-white text-sm" />
              <Button type="button" size="sm" variant="outline" disabled={renaming || !filenameDraft.trim() || filenameDraft.trim() === result.filename} onClick={() => void renameWorkbook()} className="shrink-0">{renaming ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Save className="mr-2 h-4 w-4" />}保存文件名</Button>
            </div>
          </div>
          <div className="grid w-full gap-3 sm:grid-cols-2 lg:w-auto lg:min-w-[520px]">
            <div className="rounded-md border border-slate-200 bg-slate-50 p-3">
              <p className="mb-2 text-xs font-semibold text-slate-700">正式稿</p>
              <div className="grid grid-cols-2 gap-2">
                <Button onClick={() => void previewMaster()} variant="outline" disabled={previewing !== null}>{previewing === "master" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Eye className="mr-2 h-4 w-4" />}预览</Button>
                <Button onClick={() => void downloadMasterWorkbook()} disabled={downloading !== null} className="bg-sky-700 hover:bg-sky-800">{downloading === "master" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载正式稿</Button>
              </div>
            </div>
            <div className="rounded-md border border-amber-300 bg-amber-50 p-3">
              <p className="mb-2 text-xs font-semibold text-amber-900">修改稿（批注与修改记录版）</p>
              <div className="grid grid-cols-2 gap-2">
                <Button onClick={() => void previewReview()} variant="outline" disabled={previewing === "review"} className="border-amber-300 bg-white text-amber-900 hover:bg-amber-100">{previewing === "review" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Eye className="mr-2 h-4 w-4" />}预览</Button>
                <Button onClick={() => void downloadReviewWorkbook()} disabled={downloading === "review"} className="bg-amber-700 text-white hover:bg-amber-800 disabled:bg-amber-700 disabled:text-white disabled:opacity-75">{downloading === "review" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载修改稿</Button>
              </div>
            </div>
          </div>
        </div>
        <dl className="mt-5 grid grid-cols-2 gap-px overflow-hidden rounded-lg border border-slate-200 bg-slate-200 lg:grid-cols-5"><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">输入候选人员</dt><dd className="mt-1 text-lg font-semibold tabular-nums text-slate-950">{result.input_person_count}</dd></div><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">总表自动写入</dt><dd className="mt-1 text-lg font-semibold tabular-nums text-slate-950">{result.output_person_count}</dd></div><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">仅人工复核</dt><dd className="mt-1 text-lg font-semibold tabular-nums text-amber-900">{result.manual_only_person_count}</dd></div><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">变更文件</dt><dd className="mt-1 text-lg font-semibold tabular-nums text-slate-950">{result.source_file_count}</dd></div><div className="bg-slate-50 px-4 py-3"><dt className="text-xs text-slate-500">覆盖校验</dt><dd className={cn("mt-1 text-sm font-semibold", peopleBalanced ? "text-emerald-800" : "text-red-700")}>{peopleBalanced ? "候选人员全部有去向" : `存在 ${result.coverage_gap_count} 人缺口`}</dd></div></dl>
        <WorkbookPreview preview={masterPreview} />
        <WorkbookPreview preview={reviewPreview} />
      </section>

      <section className="rounded-xl border border-amber-200 bg-amber-50/50">
        <div className="flex flex-col gap-3 px-4 py-4 sm:flex-row sm:items-center sm:justify-between sm:px-5"><button type="button" onClick={() => void toggleIssues()} className="flex min-w-0 flex-1 items-center justify-between gap-4 text-left focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-amber-700 focus-visible:ring-inset" aria-expanded={issuesOpen}><span className="flex items-center gap-3"><AlertTriangle className="h-5 w-5 shrink-0 text-amber-700" /><span><span className="block text-sm font-semibold text-amber-950">待人工处理</span><span className="mt-1 block text-xs text-amber-800">下载后可在“处理结果、处理值、处理备注”列填写，再导入更新总表 · {result.issue_count} 项</span></span></span>{issuesOpen ? <ChevronUp className="h-5 w-5 shrink-0 text-amber-800" /> : <ChevronDown className="h-5 w-5 shrink-0 text-amber-800" />}</button><div className="flex flex-wrap gap-2"><Button type="button" onClick={() => void confirmSelectedDepartmentTransfers()} disabled={!selectedDepartmentIssueIds.length || confirmingDepartmentTransfers} className="shrink-0 border border-emerald-300 bg-emerald-100 text-emerald-900 hover:bg-emerald-200 disabled:border-emerald-300 disabled:bg-emerald-100 disabled:text-emerald-900 disabled:opacity-100">{confirmingDepartmentTransfers ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}确定更新已选择项</Button><Button onClick={() => void downloadManualIssuesWorkbook()} disabled={downloading !== null} className="shrink-0 border border-amber-300 bg-amber-100 text-amber-900 hover:bg-amber-200 disabled:border-amber-300 disabled:bg-amber-100 disabled:text-amber-900 disabled:opacity-100">{downloading === "issues" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Download className="mr-2 h-4 w-4" />}下载待人工处理</Button></div></div>
        {issuesOpen ? <div className="border-t border-amber-200 px-4 pb-4 pt-4 sm:px-5"><div className="flex flex-col gap-4"><div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between"><div className="min-w-0"><p className="truncate text-sm font-medium text-slate-900">{result.issues_filename || "待人工处理.xlsx"}</p><p className="mt-1 text-xs text-slate-600">仅会回写可通过唯一工号和目标字段安全定位的已处理事项；无法安全定位的复杂项仍保留在清单中。</p></div><Button onClick={() => void previewIssues()} variant="outline" disabled={previewing === "issues"} className="shrink-0">{previewing === "issues" ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Eye className="mr-2 h-4 w-4" />}刷新预览</Button></div><div className="flex flex-col gap-2 rounded-md border border-amber-200 bg-white p-3 sm:flex-row sm:items-center"><label className="min-w-0 flex-1"><span className="sr-only">选择已填写的人工处理表</span><Input ref={manualReviewFileInputRef} type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" onChange={(event) => setManualReviewFile(event.target.files?.[0] || null)} className="h-9 cursor-pointer bg-white text-xs" /></label><Button type="button" onClick={() => void importManualReviewWorkbook()} disabled={!manualReviewFile || importingManualReview} className="shrink-0 bg-emerald-700 hover:bg-emerald-800">{importingManualReview ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Upload className="mr-2 h-4 w-4" />}导入并更新总表</Button></div>{manualReviewFile ? <p className="text-xs text-slate-600">待导入：{manualReviewFile.name}</p> : null}<WorkbookPreview preview={issuesPreview} /></div></div> : null}
      </section>

      <section className="rounded-xl border border-slate-200 bg-white p-4 sm:p-5"><div className="flex items-center justify-between gap-3"><div className="flex items-center gap-2"><FileSpreadsheet className="h-5 w-5 text-slate-500" /><h2 className="text-sm font-semibold text-slate-950">人工问题摘要</h2></div><span className="text-xs text-slate-500">共 {result.issues.length} 项</span></div>{result.issues.length ? <ul className="mt-3 divide-y divide-slate-200">{result.issues.map((issue, index) => {
        const canConfirmDepartment = issue.issue_type === "department_transfer" && issue.issue_id && issue.employee_id && issue.proposed_value;
        const selectedAction = issueActions[issue.issue_id || ""] || issue.recommended_action || "";
        const canResolve = Boolean(issue.issue_id && issue.action_options?.length);
        return <li key={issue.issue_id || `${issue.person_name}-${issue.target_field}-${index}`} className="grid gap-4 py-5 lg:grid-cols-[minmax(0,1fr)_180px] lg:items-stretch">
          <div className="min-w-0">
            <div className="flex flex-wrap items-center gap-x-2 gap-y-1">
              <span className="text-sm font-semibold text-slate-950">{issue.person_name || "未识别人员"}</span>
              <span className="text-xs text-amber-900">{issue.target_field || "人员匹配"}</span>
              <span className={cn("rounded px-1.5 py-0.5 text-[11px] font-medium", issue.difficulty === "简单" ? "bg-emerald-100 text-emerald-800" : issue.difficulty === "复杂" ? "bg-red-100 text-red-800" : "bg-amber-100 text-amber-800")}>{issue.difficulty || "一般"}</span>
              <span className="text-xs text-slate-600">{issue.problem_type || issue.issue_type}</span>
              <span className={cn("text-xs font-medium", issue.issue_type === "department_transfer" ? "text-emerald-800" : "text-slate-500")}>{issue.issue_type === "department_transfer" ? "可确认更新" : "保持待处理"}</span>
            </div>
            <p className="mt-2 text-sm leading-6 text-slate-700">{issue.message}</p>
            {issue.suggestion ? <p className="mt-1 text-xs leading-5 text-slate-600">建议：{issue.suggestion}</p> : null}
            {issue.source_files?.length ? <p className="mt-1 text-xs leading-5 text-slate-500">来源：{issue.source_files.join("、")}</p> : null}
            {canResolve ? <div className="mt-3 flex flex-wrap items-center gap-2 border-t border-slate-100 pt-3"><select aria-label={`选择 ${issue.person_name || "该事项"} 的处理方式`} value={selectedAction} onChange={(event) => setIssueActions((current) => ({ ...current, [issue.issue_id!]: event.target.value }))} className="h-9 max-w-full rounded-md border border-slate-300 bg-white px-2 text-xs text-slate-700">{issue.action_options!.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}</select><Button type="button" size="sm" disabled={resolvingIssueId === issue.issue_id || !["apply_proposed", "keep_current"].includes(selectedAction)} onClick={() => void resolveIssue(issue.issue_id!, selectedAction as "apply_proposed" | "keep_current")} className="bg-emerald-700 hover:bg-emerald-800">{resolvingIssueId === issue.issue_id ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}一键处理</Button></div> : null}
          </div>
          {canConfirmDepartment ? <label className="flex min-h-14 cursor-pointer items-center justify-between gap-3 rounded-md border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-medium text-slate-700 transition-colors hover:border-emerald-300 hover:bg-emerald-50 lg:self-stretch"><span>批量选择</span><input type="checkbox" aria-label={`批量确认 ${issue.person_name} 的部门变更`} checked={selectedDepartmentIssueIds.includes(issue.issue_id!)} onChange={() => toggleDepartmentIssue(issue.issue_id!)} className="h-4 w-4 rounded border-slate-300 text-emerald-700 focus:ring-emerald-700" /></label> : <div className="hidden lg:block" aria-hidden="true" />}
        </li>;
      })}</ul> : <p className="mt-3 text-sm text-slate-600">本批次没有需要人工处理的记录，人工处理表仅保留表头。</p>}</section>
    </div>
  );
}
