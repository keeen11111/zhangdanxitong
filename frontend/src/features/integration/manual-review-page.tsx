"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { AlertTriangle, CheckCircle2, ChevronLeft, Loader2, Save } from "lucide-react";
import { toast } from "sonner";

import { api, FinancialWorkbookIntegration, ManualIssue, ManualReviewDecision, MatchingPolicy, PipelineExportResult, Project } from "@/lib/api";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { cn } from "@/lib/utils";
import { selectLatestIntegrationResult, shouldLoadFinancialIntegrationResult } from "@/lib/integration-result-selection";
import { getManualReviewGuidance } from "@/lib/manual-review-guidance";
import { validManualSheetMappings, ManualSheetMappingSelection } from "@/lib/manual-sheet-mapping";

type ReviewMode = "manual" | "keep_current";

interface ReviewDraft {
  mode: ReviewMode;
  value: string;
  selectedOption?: string;
  remember?: boolean;
}

const emptyDraft: ReviewDraft = { mode: "manual", value: "" };

function financialIssueTitle(issue: Record<string, unknown>) {
  const titles: Record<string, string> = {
    empty_source_sheet: "来源工作表为空",
    ambiguous_sheet: "无法确定对应的目标工作表",
    unmatched_sheet: "未匹配到目标工作表",
    conflicting_value: "来源数据存在冲突",
  };
  return titles[String(issue.issue_type || "")] || "需要人工确认";
}

function financialIssueLocation(issue: Record<string, unknown>) {
  const files = Array.isArray(issue.source_files) ? issue.source_files.filter((value): value is string => typeof value === "string") : [];
  const sheets = Array.isArray(issue.source_sheets) ? issue.source_sheets.filter((value): value is string => typeof value === "string") : [];
  return [files.join("、"), sheets.join("、")].filter(Boolean).join(" · ");
}

const SHEET_MAPPING_ISSUE_TYPES = new Set([
  "ambiguous_sheet",
  "unmatched_sheet",
  "insufficient_topic_evidence",
]);

const SEMANTIC_REVIEW_ISSUE_TYPES = new Set([
  "empty_source_sheet", "unprofiled_source_sheet", "ambiguous_sheet", "unmatched_sheet",
  "insufficient_topic_evidence", "missing_business_key", "source_formula_conflict",
  "formula_target_conflict", "duplicate_source_record", "ambiguous_record",
  "unknown_record", "conflicting_value",
]);

function isSheetMappingIssue(issue: ManualIssue) {
  return SHEET_MAPPING_ISSUE_TYPES.has(issue.issue_type);
}

function isSemanticReviewIssue(issue: ManualIssue) {
  return SEMANTIC_REVIEW_ISSUE_TYPES.has(issue.issue_type);
}

function sourceSheetName(issue: ManualIssue) {
  return issue.source_sheets?.[0] || "这份来源数据";
}

function sourceFileName(issue: ManualIssue) {
  return issue.source_files?.[0] || "来源文件";
}

function sheetMappingQuestion(issue: ManualIssue) {
  return `请确认「${sourceSheetName(issue)}」的数据，应更新到总表的哪个部分？`;
}

function semanticTaskQuestion(issue: ManualIssue) {
  const sheet = sourceSheetName(issue);
  const questions: Record<string, string> = {
    empty_source_sheet: `「${sheet}」没有可更新的数据，本次是否不更新？`,
    unprofiled_source_sheet: `「${sheet}」的表头无法识别，本次是否不更新？`,
    missing_business_key: `「${sheet}」缺少可确认人员/记录的编号，本次是否不更新？`,
    source_formula_conflict: `「${sheet}」包含来源公式，是否暂不使用该公式结果？`,
    formula_target_conflict: `总表目标位置含计算公式，是否保留总表公式？`,
    duplicate_source_record: `「${sheet}」同一记录出现多次，是否本次不更新？`,
    ambiguous_record: `总表中找到多条相同记录，是否本次不更新？`,
    unknown_record: `总表没有对应记录，是否本次不新增？`,
    conflicting_value: `多个来源给出了不同数值，是否本次保留总表？`,
  };
  return questions[issue.issue_type] || `「${sheet}」本次是否不更新？`;
}

function issueKey(issue: ManualIssue, index: number) {
  return `${issue.issue_id || "issue"}-${index}`;
}

function textValue(value: string | number | boolean | null | undefined) {
  return value == null || value === "" ? "—" : String(value);
}

function isDirectlyEditable(issue: ManualIssue) {
  return issue.can_update_online === true;
}

function MatchingPolicyPanel({ projectId, initial }: { projectId: string; initial: MatchingPolicy | null }) {
  const [policy, setPolicy] = useState(initial);
  const [saving, setSaving] = useState(false);
  useEffect(() => setPolicy(initial), [initial]);
  if (!policy) return null;
  async function save() {
    setSaving(true);
    try {
      const next = await api.updateMatchingPolicy(projectId, {
        auto_match_threshold: policy.auto_match_threshold,
        field_weights: policy.field_weights,
        source_rules: policy.source_rules,
      });
      setPolicy(next);
      toast.success("匹配策略已保存，下次整合生效");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "匹配策略保存失败");
    } finally { setSaving(false); }
  }
  return <section className="mt-5 rounded-md border border-slate-200 bg-white p-4 sm:p-5"><div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between"><div><h2 className="text-sm font-semibold text-slate-950">智能匹配策略</h2><p className="mt-1 text-xs leading-5 text-slate-500">仅作用于 {policy.company} · {policy.salary_month}，身份冲突仍必须人工确认。</p></div><Button type="button" size="sm" onClick={() => void save()} disabled={saving} className="bg-slate-900 hover:bg-slate-700">{saving ? <Loader2 className="mr-2 h-3.5 w-3.5 animate-spin" /> : <Save className="mr-2 h-3.5 w-3.5" />}保存策略</Button></div><div className="mt-4 grid gap-4 sm:grid-cols-[12rem_1fr]"><label className="text-xs font-medium text-slate-700">自动关联阈值<input type="number" min={0.85} max={0.99} step={0.01} value={policy.auto_match_threshold} onChange={(event) => setPolicy({ ...policy, auto_match_threshold: Number(event.target.value) })} className="mt-1 h-9 w-full rounded-md border border-slate-300 px-2 text-sm" /></label><div><p className="text-xs font-medium text-slate-700">字段权重</p><div className="mt-1 grid grid-cols-2 gap-2 sm:grid-cols-4">{["姓名", "公司", "部门", "岗位"].map((field) => <label key={field} className="text-xs text-slate-500">{field}<input type="number" min={0} step={0.05} value={policy.field_weights[field] ?? 0} onChange={(event) => setPolicy({ ...policy, field_weights: { ...policy.field_weights, [field]: Number(event.target.value) } })} className="mt-1 h-9 w-full rounded-md border border-slate-300 px-2 text-sm text-slate-800" /></label>)}</div></div></div></section>;
}

function decisionFromDraft(issue: ManualIssue, draft: ReviewDraft): ManualReviewDecision | null {
  if (!issue.issue_id || !isDirectlyEditable(issue)) return null;
  if (draft.mode === "keep_current") {
    return { issue_id: issue.issue_id, outcome: "保留总表", value: null, note: "", remember: draft.remember };
  }
  if (!draft.value.trim()) return null;
  return { issue_id: issue.issue_id, outcome: "更新到总表", value: draft.value, note: "", remember: draft.remember };
}

function SheetMappingReview({
  issues,
  targetSheets,
  isSaving,
  onConfirm,
}: {
  issues: ManualIssue[];
  targetSheets: string[];
  isSaving: boolean;
  onConfirm: (mappings: ManualSheetMappingSelection[]) => void;
}) {
  const [selections, setSelections] = useState<Record<string, string>>({});
  const selectedMappings = validManualSheetMappings(Object.entries(selections).map(([issue_id, target_sheet]) => ({ issue_id, target_sheet })));
  if (!issues.length) return null;

  return (
    <section className="mt-5 overflow-hidden rounded-md border border-slate-200 bg-white">
      <div className="border-b border-slate-200 px-4 py-4 sm:px-5">
        <div className="flex flex-col gap-1 sm:flex-row sm:items-center sm:justify-between"><div><h2 className="text-sm font-semibold text-slate-950">确认更新位置</h2><p className="mt-1 text-sm text-slate-600">勾选本次要更新的来源工作表，并在左侧选择目标位置。未勾选的事项不会写入总表。</p></div><span className="text-xs font-medium text-teal-800">已选 {selectedMappings.length} / {issues.length} 项</span></div>
      </div>
      <div className="divide-y divide-slate-200">
        {issues.map((issue, index) => {
          const issueId = issue.issue_id || `sheet-task-${index}`;
          const rowKey = `${issueId}-${index}`;
          const targetSheet = selections[issueId] || "";
          const checked = Boolean(targetSheet);
          const recommendedSheets = issue.candidate_target_sheets || [];
          return <div key={rowKey} className={cn("grid gap-4 px-4 py-4 transition sm:grid-cols-[19rem_minmax(0,1fr)] sm:px-5", checked ? "bg-teal-50/40" : "bg-white")}>
            <div className="space-y-2">
              <label className="flex items-start gap-2 text-sm font-medium text-slate-900"><input type="checkbox" checked={checked} disabled={!issue.issue_id || !targetSheets.length || isSaving} onChange={(event) => setSelections((current) => { const next = { ...current }; if (!event.target.checked) delete next[issueId]; else if (!next[issueId]) next[issueId] = recommendedSheets[0] || targetSheets[0] || ""; return next; })} className="mt-0.5 h-4 w-4 rounded border-slate-300 text-teal-700 focus:ring-teal-700" aria-label={`选择${sourceSheetName(issue)}更新到总表`} />选择更新</label>
              <label className="block text-xs font-medium text-slate-600" htmlFor={`target-sheet-${rowKey}`}>更新到总表的</label>
              <select id={`target-sheet-${rowKey}`} value={targetSheet} disabled={!checked || isSaving} onChange={(event) => setSelections((current) => ({ ...current, [issueId]: event.target.value }))} className="h-10 w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm text-slate-900 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-700"><option value="">请选择工作表</option>{targetSheets.map((sheet) => <option key={sheet} value={sheet}>{sheet}</option>)}</select>
              {recommendedSheets.length ? <p className="text-xs leading-5 text-teal-800">系统参考：{recommendedSheets.join("、")}</p> : null}
            </div>
            <div><h3 className="text-base font-semibold text-slate-950">{sheetMappingQuestion(issue)}</h3><p className="mt-2 text-sm leading-6 text-slate-600">来源：{sourceFileName(issue)} · {sourceSheetName(issue)}。系统之前没有足够依据自动决定位置，因此没有写入任何数据。</p></div>
          </div>;
        })}
      </div>
      <div className="flex flex-col gap-3 border-t border-slate-200 bg-slate-50 px-4 py-4 sm:flex-row sm:items-center sm:justify-between sm:px-5"><p className="text-xs leading-5 text-slate-600">确认后会一次性生成最新正式稿和修订稿；未勾选的事项会继续保留为待处理。</p><Button type="button" disabled={!selectedMappings.length || isSaving} onClick={() => onConfirm(selectedMappings)} className="bg-teal-700 hover:bg-teal-800">{isSaving ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <CheckCircle2 className="mr-2 h-4 w-4" />}确定更新（{selectedMappings.length}）</Button></div>
    </section>
  );
}

function SemanticIgnoreTasks({
  issues,
  savingIssueId,
  onIgnore,
}: {
  issues: ManualIssue[];
  savingIssueId: string | null;
  onIgnore: (issue: ManualIssue) => void;
}) {
  if (!issues.length) return null;
  return <section className="mt-5 rounded-md border border-slate-200 bg-white"><div className="border-b border-slate-200 px-4 py-4 sm:px-5"><h2 className="text-sm font-semibold text-slate-950">确认本次不更新的数据</h2><p className="mt-1 text-sm text-slate-600">这些来源不能安全写入。若确认本月无需处理，选择“本次不更新”即可从清单移除，且不会改动总表。</p></div><ul className="divide-y divide-slate-200" role="list">{issues.map((issue, index) => { const issueId = issue.issue_id || `semantic-${index}`; const rowKey = `${issueId}-${index}`; return <li key={rowKey} className="flex flex-col gap-3 px-4 py-4 sm:flex-row sm:items-center sm:justify-between sm:px-5"><div><h3 className="text-sm font-medium text-slate-950">{semanticTaskQuestion(issue)}</h3><p className="mt-1 text-xs leading-5 text-slate-500">来源：{sourceFileName(issue)} · {sourceSheetName(issue)}</p></div><Button type="button" variant="outline" disabled={!issue.issue_id || savingIssueId === issueId} onClick={() => onIgnore(issue)}>{savingIssueId === issueId ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}本次不更新</Button></li>; })}</ul></section>;
}

function ManualIssueRow({
  issue,
  index,
  draft,
  onChange,
  isSaving,
  onSubmit,
}: {
  issue: ManualIssue;
  index: number;
  draft: ReviewDraft;
  onChange: (next: ReviewDraft) => void;
  isSaving: boolean;
  onSubmit: () => void;
}) {
  const suggestion = issue.proposed_value == null ? "" : String(issue.proposed_value);
  const editable = isDirectlyEditable(issue);
  const keepingCurrent = draft.mode === "keep_current";
  const actionOptions = issue.action_options || [];
  const candidateOptions = Array.from(new Map(
    (issue.candidate_values || [])
      .filter((value) => value != null && value !== "")
      .map((value) => [String(value), value]),
  ).entries()).filter(([value]) => value !== suggestion && value !== String(issue.current_value ?? ""));
  const keepOption = actionOptions.find((option) => option.value === "keep_current");
  const proposalOption = actionOptions.find((option) => option.value === "apply_proposed");
  const guidance = getManualReviewGuidance(issue);

  function changeChoice(choice: string) {
    if (choice === "keep_current") {
      onChange({ ...draft, mode: "keep_current", value: "", selectedOption: choice });
      return;
    }
    if (choice === "manual") {
      onChange({ ...draft, mode: "manual", value: "", selectedOption: choice });
      return;
    }
    if (choice === "proposed") {
      onChange({ ...draft, mode: "manual", value: suggestion, selectedOption: choice });
      return;
    }
    const candidate = candidateOptions[Number(choice.replace("candidate-", ""))];
    if (candidate) onChange({ ...draft, mode: "manual", value: String(candidate[1]), selectedOption: choice });
  }

  return (
    <tr className="border-b border-slate-200 align-top last:border-0">
      <td className="px-4 py-4"><span className={cn("inline-flex rounded px-2 py-1 text-xs font-medium", editable ? "bg-amber-50 text-amber-900" : "bg-slate-100 text-slate-700")}>{editable ? "待确认" : "需核对"}</span></td>
      <td className="px-4 py-4 text-sm"><p className="font-medium text-slate-950">{issue.person_name || "全表"}</p>{issue.employee_id ? <p className="mt-1 text-xs text-slate-500">工号：{issue.employee_id}</p> : null}</td>
      <td className="max-w-64 px-4 py-4"><p className="text-sm font-medium text-slate-900">{issue.target_field || "待确认事项"}</p><p className="mt-1 break-words text-xs leading-5 text-slate-500">{issue.message}</p>{editable ? <p className="mt-2 rounded-md border border-teal-100 bg-teal-50 px-3 py-2 text-xs leading-5 text-teal-950">{guidance}</p> : null}</td>
      <td className="px-4 py-4"><span className={cn("inline-flex rounded px-2 py-1 font-mono text-xs", issue.target_location ? "bg-teal-50 text-teal-900" : "bg-slate-100 text-slate-700")}>{issue.target_location || "数据已唯一匹配"}</span>{typeof issue.confidence === "number" ? <p className="mt-2 text-xs text-slate-500">匹配置信度：{Math.round(issue.confidence * 100)}%</p> : null}{!editable ? <p className="mt-2 max-w-44 text-xs leading-5 text-amber-900">人员或字段无法唯一匹配，不能安全写入。</p> : null}</td>
      <td className="max-w-40 px-4 py-4 text-sm text-slate-700">{textValue(issue.current_value)}</td>
      <td className="min-w-72 px-4 py-4">{editable ? <div className="space-y-2"><select aria-label={`选择${issue.person_name || "该项"}的处理方式`} value={draft.selectedOption || (keepingCurrent ? "keep_current" : "manual")} onChange={(event) => changeChoice(event.target.value)} className="h-9 w-full rounded-md border border-slate-300 bg-white px-2 text-sm text-slate-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-teal-700"><option value="manual">我来填写</option>{suggestion ? <option value="proposed">采用系统建议：{suggestion}</option> : null}{candidateOptions.map(([value], candidateIndex) => <option key={value} value={`candidate-${candidateIndex}`}>采用来源候选值：{value}</option>)}<option value="keep_current">{keepOption?.label || "保留当前总表值"}</option></select>{keepingCurrent ? <p className="rounded-md bg-slate-50 px-3 py-2 text-xs text-slate-600">将保留当前值，不会修改总表。</p> : <Input value={draft.value} onChange={(event) => onChange({ ...draft, mode: "manual", value: event.target.value, selectedOption: "manual" })} placeholder={`填写${issue.target_field}`} className="h-9 bg-white text-sm" aria-label={`填写${issue.person_name || "该项"}的${issue.target_field}`} />}<label className="flex items-center gap-2 text-xs text-slate-600"><input type="checkbox" checked={Boolean(draft.remember)} onChange={(event) => onChange({ ...draft, remember: event.target.checked })} className="h-4 w-4 rounded border-slate-300 text-teal-700" />记住这次处理方式</label><Button type="button" onClick={onSubmit} disabled={isSaving || (!keepingCurrent && !draft.value.trim())} className="h-9 w-full bg-teal-700 hover:bg-teal-800">{isSaving ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Save className="mr-2 h-4 w-4" />}{keepingCurrent ? "确认不更新" : "确定更新"}</Button></div> : <div className="space-y-2"><p className="text-xs leading-5 text-slate-500">系统无法唯一定位，不能直接改写；你可以确认本次保留总表。</p><Button type="button" variant="outline" onClick={onSubmit} disabled={isSaving}>{isSaving ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : null}本次不更新</Button></div>}</td>
    </tr>
  );
}

export function ManualReviewPage({ projectId }: { projectId: string }) {
  const [project, setProject] = useState<Project | null>(null);
  const [result, setResult] = useState<PipelineExportResult | null>(null);
  const [financialResult, setFinancialResult] = useState<FinancialWorkbookIntegration | null>(null);
  const [drafts, setDrafts] = useState<Record<string, ReviewDraft>>({});
  const [loading, setLoading] = useState(true);
  const [savingIssueId, setSavingIssueId] = useState<string | null>(null);
  const [savingAll, setSavingAll] = useState(false);
  const [savingSheetMapping, setSavingSheetMapping] = useState(false);
  const [ignoringSemanticIssueId, setIgnoringSemanticIssueId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [matchingPolicy, setMatchingPolicy] = useState<MatchingPolicy | null>(null);

  const load = useCallback(async () => {
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
    const latestResult = selectLatestIntegrationResult(financial, pipeline);
    if (latestResult === "financial" && financial) {
      setFinancialResult(financial);
    } else if (latestResult === "pipeline" && pipeline) {
      setResult(pipeline);
    } else {
      setError(financialError instanceof Error ? financialError.message : "暂未找到待人工处理事项");
    }
    setLoading(false);
  }, [projectId]);

  useEffect(() => { void load(); }, [load]);
  useEffect(() => {
    void api.getMatchingPolicy(projectId).then(setMatchingPolicy).catch(() => setMatchingPolicy(null));
  }, [projectId]);

  const pendingIssues = useMemo(() => result?.issues.filter((issue) => issue.status !== "confirmed") || [], [result]);
  const sheetMappingIssues = useMemo(() => pendingIssues.filter(isSheetMappingIssue), [pendingIssues]);
  const semanticIgnoreIssues = useMemo(
    () => pendingIssues.filter((issue) => isSemanticReviewIssue(issue) && !isSheetMappingIssue(issue) && !isDirectlyEditable(issue)),
    [pendingIssues],
  );
  const valueReviewIssues = useMemo(
    () => pendingIssues.filter((issue) => !isSemanticReviewIssue(issue) || isDirectlyEditable(issue)),
    [pendingIssues],
  );
  const editableIssueCount = useMemo(() => pendingIssues.filter(isDirectlyEditable).length, [pendingIssues]);
  const preparedDecisions = useMemo(
    () => pendingIssues.flatMap((issue, index) => {
      const draft = drafts[issueKey(issue, index)];
      return draft ? [decisionFromDraft(issue, draft)].filter((decision): decision is ManualReviewDecision => decision !== null) : [];
    }),
    [drafts, pendingIssues],
  );
  async function saveIssue(issue: ManualIssue, index: number) {
    const key = issueKey(issue, index);
    const draft = drafts[key] || emptyDraft;
    if (!issue.issue_id) return;
    if (isDirectlyEditable(issue) && draft.mode !== "keep_current" && !draft.value.trim()) return toast.error("请先填写处理值，或选择保留当前值");
    const decision: ManualReviewDecision = {
      issue_id: issue.issue_id,
      outcome: !isDirectlyEditable(issue) || draft.mode === "keep_current" ? "保留总表" : "更新到总表",
      value: !isDirectlyEditable(issue) || draft.mode === "keep_current" ? null : draft.value,
      note: "",
      remember: draft.remember,
    };
    setSavingIssueId(key);
    try {
      await api.applyOnlineManualReview(projectId, [decision]);
      await load();
      setDrafts((current) => {
        const next = { ...current };
        delete next[key];
        return next;
      });
      toast.success("该项已更新到最终总表");
    } catch (saveError) {
      toast.error(saveError instanceof Error ? saveError.message : "更新总表失败");
    } finally {
      setSavingIssueId(null);
    }
  }

  async function saveAllPreparedIssues() {
    if (!preparedDecisions.length) return toast.error("请先为至少一项选择处理方式或填写处理值");
    setSavingAll(true);
    try {
      await api.applyOnlineManualReview(projectId, preparedDecisions);
      await load();
      setDrafts({});
      toast.success(`已更新 ${preparedDecisions.length} 项并生成最新总表`);
    } catch (saveError) {
      toast.error(saveError instanceof Error ? saveError.message : "更新总表失败");
    } finally {
      setSavingAll(false);
    }
  }

  async function confirmSheetMappings(mappings: ManualSheetMappingSelection[]) {
    if (!mappings.length) return toast.error("请至少勾选一项并选择目标工作表");
    setSavingSheetMapping(true);
    try {
      await api.applyManualSheetMappings(projectId, mappings);
      await load();
      toast.success(`已批量更新 ${mappings.length} 个工作表，并生成最新总表`);
    } catch (saveError) {
      toast.error(saveError instanceof Error ? saveError.message : "保存更新位置失败");
    } finally {
      setSavingSheetMapping(false);
    }
  }

  async function ignoreSemanticIssue(issue: ManualIssue) {
    if (!issue.issue_id) return;
    setIgnoringSemanticIssueId(issue.issue_id);
    try {
      await api.ignoreSemanticManualIssue(projectId, issue.issue_id);
      await load();
      toast.success(`已标记「${sourceSheetName(issue)}」本次不更新`);
    } catch (saveError) {
      toast.error(saveError instanceof Error ? saveError.message : "保存本次不更新失败");
    } finally {
      setIgnoringSemanticIssueId(null);
    }
  }

  if (loading) return <div className="mx-auto max-w-7xl py-16 text-center text-sm text-slate-500"><Loader2 className="mx-auto mb-3 h-5 w-5 animate-spin" />正在加载人工处理事项…</div>;
  if (financialResult) return <FinancialManualReview project={project} projectId={projectId} result={financialResult} />;
  if (!result) return <div className="mx-auto max-w-3xl rounded-md border border-slate-200 bg-white p-6"><h1 className="text-xl font-semibold text-slate-950">暂未找到人工处理事项</h1><p className="mt-2 text-sm text-slate-600">{error || "请先完成本月整合。"}</p><Button asChild className="mt-5 bg-teal-700 hover:bg-teal-800"><Link href={`/projects/${projectId}/result`}>返回更新状态</Link></Button></div>;

  return (
    <div className="payroll-page max-w-7xl pb-8">
      <header className="border-b border-slate-200/90 pb-7"><Link href={`/projects/${projectId}/result`} className="inline-flex items-center gap-1 text-sm text-slate-500 hover:text-blue-800"><ChevronLeft className="h-4 w-4" />返回更新状态</Link><p className="payroll-kicker mt-5">{project?.salary_month || "本月"}工资总表</p><h1 className="page-heading mt-2">确认更新任务</h1><p className="mt-2 text-sm text-slate-600">不需要理解系统术语。每一项只需确认：数据更新到哪里、采用哪个值，或本次是否保留总表。</p></header>

      <MatchingPolicyPanel projectId={projectId} initial={matchingPolicy} />

      {pendingIssues.length ? <><SheetMappingReview issues={sheetMappingIssues} targetSheets={result.target_sheets || []} isSaving={savingSheetMapping} onConfirm={(mappings) => void confirmSheetMappings(mappings)} /><SemanticIgnoreTasks issues={semanticIgnoreIssues} savingIssueId={ignoringSemanticIssueId} onIgnore={(issue) => void ignoreSemanticIssue(issue)} />{valueReviewIssues.length ? <section className="mt-5 rounded-md border border-slate-200 bg-white"><div className="flex flex-col gap-3 border-b border-slate-200 px-4 py-4 sm:flex-row sm:items-center sm:justify-between sm:px-5"><div><h2 className="text-sm font-semibold text-slate-950">确认具体数据</h2><p className="mt-1 text-xs text-slate-500">共 {valueReviewIssues.length} 项，其中 {editableIssueCount} 项可直接选择或填写；已准备 {preparedDecisions.length} 项。</p></div><Button type="button" disabled={!preparedDecisions.length || savingAll} onClick={() => void saveAllPreparedIssues()} className="bg-teal-700 hover:bg-teal-800">{savingAll ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <Save className="mr-2 h-4 w-4" />}更新总表（{preparedDecisions.length}）</Button></div><div className="overflow-x-auto"><table className="min-w-[1180px] w-full text-left"><thead className="bg-slate-50 text-xs text-slate-600"><tr><th className="px-4 py-3 font-medium">状态</th><th className="px-4 py-3 font-medium">人员</th><th className="px-4 py-3 font-medium">需要处理</th><th className="px-4 py-3 font-medium">写入位置</th><th className="px-4 py-3 font-medium">当前值</th><th className="px-4 py-3 font-medium">选择或填写</th></tr></thead><tbody>{valueReviewIssues.map((issue, index) => { const key = issueKey(issue, index); return <ManualIssueRow key={key} issue={issue} index={index} draft={drafts[key] || emptyDraft} isSaving={savingIssueId === key} onChange={(next) => setDrafts((current) => ({ ...current, [key]: next }))} onSubmit={() => void saveIssue(issue, index)} />; })}</tbody></table></div></section> : null}</> : <section className="mt-5 rounded-md border border-slate-200 bg-white py-12 text-center"><CheckCircle2 className="mx-auto h-6 w-6 text-emerald-700" /><h2 className="mt-3 text-sm font-semibold text-slate-950">已全部更新</h2><p className="mt-1 text-sm text-slate-600">所有人工处理内容已写入最终总表。</p><Button asChild variant="outline" className="mt-4"><Link href={`/projects/${projectId}/result`}>返回更新状态</Link></Button></section>}
    </div>
  );
}

function FinancialManualReview({ project, projectId, result }: { project: Project | null; projectId: string; result: FinancialWorkbookIntegration }) {
  const [matchingPolicy, setMatchingPolicy] = useState<MatchingPolicy | null>(null);
  useEffect(() => {
    void api.getMatchingPolicy(projectId).then(setMatchingPolicy).catch(() => setMatchingPolicy(null));
  }, [projectId]);
  return (
    <div className="mx-auto max-w-5xl pb-8">
      <header className="border-b border-slate-200 pb-5">
        <Link href={`/projects/${projectId}/result`} className="inline-flex items-center gap-1 text-sm text-slate-500 hover:text-teal-800"><ChevronLeft className="h-4 w-4" />返回整合结果</Link>
        <p className="mt-5 text-sm text-slate-500">{project?.salary_month || "本月"}账单项目</p>
        <h1 className="mt-1 text-2xl font-semibold tracking-tight text-slate-950">待人工处理</h1>
        <p className="mt-2 text-sm text-slate-600">请根据来源文件核对每项内容。在项目页修正来源文件或映射后，重新整合即可更新此清单。</p>
      </header>

      <MatchingPolicyPanel projectId={projectId} initial={matchingPolicy} />

      <section className="mt-5 rounded-md border border-amber-200 bg-amber-50/50 p-4 sm:p-5">
        <div className="flex items-start gap-3"><AlertTriangle className="mt-0.5 h-5 w-5 shrink-0 text-amber-700" /><div><h2 className="text-sm font-semibold text-amber-950">共 {result.issues.length} 项待处理</h2><p className="mt-1 text-sm leading-6 text-amber-900">这些事项未被自动写入总表，避免在工作表或记录无法唯一确认时产生错误数据。</p></div></div>
        <Button asChild className="mt-4 bg-teal-700 hover:bg-teal-800"><Link href={`/projects/${projectId}/result`}>返回整合结果</Link></Button>
      </section>

      {result.issues.length ? <section className="mt-5 overflow-hidden rounded-md border border-slate-200 bg-white"><div className="overflow-x-auto"><table className="min-w-[720px] w-full text-left"><thead className="bg-slate-50 text-xs text-slate-600"><tr><th className="px-4 py-3 font-medium">需要处理</th><th className="px-4 py-3 font-medium">问题说明</th><th className="px-4 py-3 font-medium">来源</th></tr></thead><tbody>{result.issues.map((issue, index) => <tr key={`${String(issue.issue_type || "issue")}-${index}`} className="border-t border-slate-200 align-top"><td className="px-4 py-4 text-sm font-medium text-slate-950">{financialIssueTitle(issue)}</td><td className="max-w-xl px-4 py-4 text-sm leading-6 text-slate-700">{typeof issue.message === "string" ? issue.message : "系统没有安全自动写入，请根据来源文件确认后处理。"}</td><td className="px-4 py-4 text-xs leading-5 text-amber-900">{financialIssueLocation(issue) || "—"}</td></tr>)}</tbody></table></div></section> : <section className="mt-5 rounded-md border border-emerald-200 bg-emerald-50 py-12 text-center"><CheckCircle2 className="mx-auto h-6 w-6 text-emerald-700" /><h2 className="mt-3 text-sm font-semibold text-emerald-950">已全部处理</h2><p className="mt-1 text-sm text-emerald-800">重新整合后没有发现需要人工处理的事项。</p></section>}
    </div>
  );
}
