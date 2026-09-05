"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Link from "next/link";
import {
  AlertCircle,
  Activity,
  ArrowUp,
  BookOpen,
  CheckCircle2,
  ChevronDown,
  Download,
  Eye,
  FileSpreadsheet,
  FileText,
  Info,
  KeyRound,
  Loader2,
  Paperclip,
  PanelRight,
  ShieldCheck,
  Sparkles,
  Square,
  Table2,
  Trash2,
  X,
} from "lucide-react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { AgentMascot } from "@/components/agent-mascot";
import {
  AgentDecisionAction,
  AgentMaterial,
  AgentMemoryResponse,
  AgentMessage,
  AgentModelStatus,
  AgentRun,
  AgentRunEvent,
  AgentRunResult,
  AgentWorkItem,
  FileMeta,
  FilePreview,
  Project,
  api,
} from "@/lib/api";
import { cn } from "@/lib/utils";
import { groupAgentItems, type AgentItemGroup } from "@/lib/agent-batching";
import { agentActivityEventLabel, agentActivitySummary } from "@/lib/agent-activity";
import { agentShortcutAction, canResumeDirectWrite, isResumableAgentRun, nextPendingAgentItem, visibleAgentItems } from "@/lib/agent-workflow";
import { isAgentProcessingInstruction, isAgentStartCommand } from "@/lib/agent-command";
import { AgentPlanConfirmation } from "./agent-plan-confirmation";
import { AgentResultView } from "./agent-result-view";

const RUN_LABEL: Record<string, string> = {
  draft: "准备开始",
  planning: "分析中",
  ready: "可以执行",
  processing: "处理中",
  execution_incomplete: "执行未完成",
  review: "等待回答",
  completed: "已完成",
  published: "已发布",
  blocked: "已阻断",
  failed: "处理失败",
};

const SHOWCASE_PROJECT_IDS = new Set([
  "27fb356e0b494ac7bfd0013bd7f4aebc",
  "37f18e853f4e4a89b155bbb7c302779c",
  "4f6d5c8b7a294e46a1f03d92c6e8b745",
]);

const SHOWCASE_DOWNLOAD_NAMES: Record<string, string> = {
  "27fb356e0b494ac7bfd0013bd7f4aebc": "样本一_已更新_202608所属月202607_工资核算总表.xlsx",
  "37f18e853f4e4a89b155bbb7c302779c": "样本二_202608所属月202607_工资核算总表.xlsx",
  "4f6d5c8b7a294e46a1f03d92c6e8b745": "待确定稿.xlsx",
};

// 这些状态表示后台已停、需要用户发“继续”或处理异常后才能推进。
const RESUME_PROMPT_STATUSES = new Set(["execution_incomplete", "blocked", "failed"]);
// 自动续跑会在数百毫秒内把可续跑的运行拉回 processing；延迟后再确认，
// 避免和自动续跑抢跑导致误弹。
const RESUME_PROMPT_DELAY_MS = 6000;

function sendAgentSystemNotification(title: string, body: string) {
  if (typeof window === "undefined" || !("Notification" in window)) return;
  // 页面可见时弹窗本身已经足够醒目，只在切走页面时发系统通知。
  if (Notification.permission !== "granted" || document.visibilityState === "visible") return;
  try {
    const notification = new Notification(title, { body, tag: "agent-resume-prompt" });
    notification.onclick = () => { window.focus(); notification.close(); };
  } catch {
    // 通知发送失败不影响页面内弹窗。
  }
}

function formatValue(value: unknown) {
  if (value === null || value === undefined || value === "") return "空";
  return typeof value === "number" ? value.toLocaleString("zh-CN") : String(value);
}

function cleanAgentContent(content: string) {
  return content
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/^#{1,6}\s*/gm, "")
    .replace(/^\s*[-*]\s+/gm, "• ");
}

// 从 Agent 消息文本中解析“③待确认 / 待确认：”段落里的每个问题。
// 格式约定（对齐提示词的三段式输出）：问题以编号行（1. / 2. / ① 等）
// 或换行走列，每行一个问题；下一小节标题（④或行尾）即结束。
function parsePendingQuestions(content: string): string[] {
  const cleaned = cleanAgentContent(content);
  const sectionMatch = cleaned.match(/(?:③|三、)?\s*待确认[：:]?\s*([\s\S]*?)(?=\n\s*(?:④|四、|处理方式|我的理解)|$)/);
  if (!sectionMatch) return [];
  return sectionMatch[1]
    .split(/\n+/)
    .map((line) => line.replace(/^\s*(?:[0-9]{1,2}[.、）)]|[①-⑩])[.、]?\s*/, "").trim())
    .filter((line) => line.length >= 4 && !/^(是|否|无|暂无|没有)[。.!！\s]*$/.test(line))
    .slice(0, 6);
}

// 为单个问题生成快捷选项：从问题文字里抽取常见口径（只处理X/删除或保留/
// A还是B），生成 2-3 个可点按钮；无法抽取时给通用三选。
function quickOptionsForQuestion(question: string): { label: string; reply: string }[] {
  const options: { label: string; reply: string }[] = [];
  const scopeMatch = question.match(/只处理([^，。？?、\s]{2,10})|只写入([^，。？?、\s]{2,10})|只保留([^，。？?、\s]{2,10})/);
  if (scopeMatch) {
    const scope = scopeMatch[1] || scopeMatch[2] || scopeMatch[3];
    options.push({ label: `只处理${scope}，其余删除`, reply: `只处理${scope}，名单之外的人员从总表明细中删除，人数按来源对齐。` });
    options.push({ label: `只处理${scope}，其余保留`, reply: `只处理${scope}，名单之外的人员保留原值不动。` });
  }
  const eitherMatch = question.match(/(.{2,20}?)\s*还是\s*(.{2,20}?)[？?]/);
  if (eitherMatch) {
    const left = eitherMatch[1].trim();
    const right = eitherMatch[2].trim();
    options.push({ label: left.slice(-12), reply: `选${left}。` });
    options.push({ label: right.slice(0, 12), reply: `选${right}。` });
  }
  if (!options.length) {
    options.push({ label: "按 Agent 建议处理", reply: "按你建议的方式处理。" });
    options.push({ label: "先跳过此项", reply: "此项先跳过，不写入，保留待确认。" });
  }
  // 去重（label + reply 均一致时）
  const seen = new Set<string>();
  return options.filter((option) => {
    const key = `${option.label}|${option.reply}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).slice(0, 3);
}

function normaliseRun(value: AgentRun | { run?: AgentRun } | null, projectId: string) {
  if (!value) return null;
  const run = ("run" in value ? value.run : value) as AgentRun | undefined;
  return run?.run_id ? { ...run, project_id: run.project_id || projectId } : null;
}

function statusTone(status: AgentWorkItem["status"]) {
  if (status === "applied") return "text-emerald-700";
  if (status === "needs_conversation") return "text-amber-700";
  if (status === "failed") return "text-red-700";
  return "text-slate-500";
}

function AgentActivityMessage({ events, running, stopping, stopped, onStop }: { events: AgentRunEvent[]; running: boolean; stopping?: boolean; stopped?: boolean; onStop?: () => void }) {
  const [expanded, setExpanded] = useState(true);
  const [now, setNow] = useState(() => Date.now());
  const summary = agentActivitySummary(events, running);
  const latestActivity = summary.events.at(-1);
  const latestTimedEvent = [...summary.events].reverse().find((event) => Number(event.payload?.estimated_seconds || 0) > 0);
  const stageStartedAt = latestActivity?.at ? new Date(latestActivity.at).getTime() : null;
  const elapsedSeconds = stageStartedAt ? Math.max(0, Math.floor((now - stageStartedAt) / 1000)) : 0;
  const estimatedSeconds = Number(latestTimedEvent?.payload?.estimated_seconds || 0);
  const remainingSeconds = estimatedSeconds ? Math.max(0, estimatedSeconds - elapsedSeconds) : null;
  const waitingForModel = Boolean(running && latestActivity?.type === "model_request");
  const currentLabel = waitingForModel ? "正在等待 Agent 返回分析结果" : summary.current;

  useEffect(() => {
    if (!running) return;
    setExpanded(true);
    const timer = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(timer);
  }, [running]);

  function formatDuration(seconds: number) {
    if (seconds < 60) return `${seconds} 秒`;
    return `${Math.floor(seconds / 60)} 分 ${seconds % 60} 秒`;
  }

  if (!summary.events.length && !running) return null;

  return (
    <section className="my-6 flex gap-3" aria-label="Agent 执行过程" aria-live="polite">
      <AgentMascot className="mt-1 h-7 w-5" />
      <div className="min-w-0 max-w-[min(740px,100%)] flex-1">
        <p className="mb-1.5 text-[11px] font-medium text-slate-400">财务 Agent</p>
        <div className="overflow-clip rounded-lg border border-slate-200 bg-slate-50/70">
          <div className={cn("flex items-center gap-2.5 px-3.5 py-3", summary.isRunning && "sticky top-0 z-10 border-b border-slate-200 bg-slate-50/95 backdrop-blur-sm")}>
            <Activity className={cn("h-4 w-4 shrink-0 text-blue-600", summary.isRunning && "animate-pulse")} />
            <div className="min-w-0 flex-1"><p className="text-[11px] text-slate-400">当前步骤</p><p className="truncate text-sm font-medium text-slate-800">{currentLabel}</p></div>
            {summary.isRunning && onStop ? (
              <button type="button" className="inline-flex shrink-0 items-center gap-1.5 rounded-md border border-slate-300 bg-white px-2.5 py-1.5 text-[11px] font-medium text-slate-600 transition-colors hover:border-red-300 hover:bg-red-50 hover:text-red-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-inset focus-visible:ring-red-400 disabled:cursor-not-allowed disabled:opacity-60" onClick={onStop} disabled={stopping} aria-label="停止本次处理">
                {stopping ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Square className="h-3 w-3 fill-current" />}
                {stopping ? "正在停止" : "停止"}
              </button>
            ) : null}
            <span className={cn("shrink-0 text-[11px]", summary.isRunning ? (stopping ? "text-amber-700" : "text-blue-700") : stopped ? "text-slate-500" : summary.isIncomplete ? "text-amber-700" : "text-slate-400")}>{summary.isRunning ? (stopping ? "停止中" : "进行中") : stopped ? "已停止" : summary.isIncomplete ? "未完成" : "已完成"}</span>
          </div>
          <div className="flex flex-wrap gap-x-4 gap-y-1 px-3.5 pb-3 text-[11px] text-slate-500">
            <span>本阶段已用时 {formatDuration(elapsedSeconds)}</span>
            {running && remainingSeconds !== null ? <span>预计剩余约 {formatDuration(remainingSeconds)}</span> : null}
          </div>
          {summary.progress ? <div className="px-3.5 pb-3"><div className="mb-1.5 flex justify-between text-[11px] text-slate-500"><span>已处理 {summary.progress.current} / {summary.progress.total}</span><span>{summary.progress.percent}%</span></div><div className="h-1 overflow-hidden rounded-full bg-slate-200"><div className="h-full rounded-full bg-blue-600 transition-[width] duration-500" style={{ width: `${summary.progress.percent}%` }} /></div></div> : null}
          <button type="button" className="flex w-full items-center justify-between border-t border-slate-200 px-3.5 py-2.5 text-left text-xs text-slate-500 transition-colors hover:bg-slate-100 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-inset focus-visible:ring-blue-500" onClick={() => setExpanded((value) => !value)} aria-expanded={expanded}>
            <span>{expanded ? "收起执行记录" : `查看执行记录 · ${summary.recordCount} 条`}</span>
            <ChevronDown className={cn("h-3.5 w-3.5 transition-transform", expanded && "rotate-180")} />
          </button>
          {expanded && summary.events.length ? <ol className="divide-y divide-slate-200 border-t border-slate-200 bg-white">{summary.events.map((event, index) => {
            const person = typeof event.payload?.person_name === "string" ? event.payload.person_name : "";
            const isLatest = index === summary.events.length - 1;
            return <li key={event.event_id || `${event.revision}-${index}`} className="flex items-start gap-2.5 px-3.5 py-2.5 text-xs leading-5 text-slate-600"><span className={cn("mt-1.5 h-1.5 w-1.5 shrink-0 rounded-full", isLatest && summary.isRunning ? "bg-blue-600" : "bg-slate-300")} /><span className="min-w-0 flex-1">{agentActivityEventLabel(event)}{person ? <span className="text-slate-800"> · {person}</span> : null}</span>{event.at ? <time className="shrink-0 text-[10px] tabular-nums text-slate-400">{new Date(event.at).toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" })}</time> : null}</li>;
          })}</ol> : null}
          {waitingForModel ? <div className="border-t border-blue-100 bg-blue-50/70 px-3.5 py-2.5 text-xs text-blue-800" role="status" aria-live="polite">模型正在分析当前文件，已等待 {formatDuration(elapsedSeconds)}；页面保持连接，返回后会自动继续。</div> : null}
        </div>
      </div>
    </section>
  );
}

const KEYUAN_STAGES = [
  ["keyuan_validate", "校验 8 个输入文件"],
  ["keyuan_copy", "创建工作副本"],
  ["keyuan_payroll", "更新薪资、考勤、值班和奖金"],
  ["keyuan_external", "更新社保、个税和 OA 历史"],
  ["keyuan_exports", "生成正式稿、修改稿和待确认表"],
  ["keyuan_acceptance", "与参考表验收"],
] as const;

const KEYUAN_EVENT_DETAILS: Record<string, string> = {
  keyuan_validate: "已识别总表、薪资、社保和个税来源文件。",
  keyuan_copy: "在项目结果目录生成独立副本，原始文件保持不变。",
  keyuan_payroll: "按本次案例范围更新薪资、考勤、值班和奖金。",
  keyuan_external: "处理社保、个税与 OA 月度历史记录。",
  keyuan_exports: "生成本次可下载的结果总表。",
  keyuan_acceptance: "结果与指定的验收总表核对一致。",
};

function KeyuanReferenceProgress({ run, events, busy, onRetry, onDownload }: { run: AgentRun; events: AgentRunEvent[]; busy: boolean; onRetry: () => void; onDownload: () => void }) {
  const completed = new Set(events.map((event) => String(event.payload?.stage || "")));
  const latestStage = [...completed].filter((stage) => stage.startsWith("keyuan_")).at(-1);
  const activeIndex = Math.max(0, KEYUAN_STAGES.findIndex(([stage]) => stage === latestStage));
  const failed = run.status === "failed";
  const isComplete = run.status === "completed";
  const stageEvents = KEYUAN_STAGES.map(([stage, fallbackLabel]) => {
    const event = [...events].reverse().find((item) => String(item.payload?.stage || "") === stage);
    return { stage, label: typeof event?.payload?.label === "string" ? event.payload.label : fallbackLabel, at: event?.at };
  });

  return <section className="my-6 max-w-[740px] border border-slate-200 bg-white p-5" aria-label="科园案例处理进度" aria-live="polite">
    <div className="flex items-start justify-between gap-4">
      <div><p className="text-xs font-medium text-blue-700">科园固定案例</p><h2 className="mt-1 text-base font-semibold text-slate-900">{isComplete ? "2026 年 7 月薪资处理已完成" : "正在处理 2026 年 7 月薪资"}</h2></div>
      <span className={cn("shrink-0 text-xs font-medium", failed ? "text-red-700" : isComplete ? "text-emerald-700" : "text-blue-700")}>{failed ? "处理失败" : isComplete ? "验收完成" : "处理中"}</span>
    </div>
    <ol className="mt-5 space-y-3" aria-label="处理步骤">
      {KEYUAN_STAGES.map(([stage, label], index) => {
        const done = completed.has(stage) || isComplete;
        const active = !done && !failed && index === activeIndex;
        return <li key={stage} className="flex items-center gap-3 text-sm"><span className={cn("flex h-6 w-6 shrink-0 items-center justify-center rounded-full text-xs font-medium", done ? "bg-emerald-600 text-white" : active ? "bg-blue-700 text-white" : "bg-slate-100 text-slate-500")}>{done ? "✓" : index + 1}</span><span className={cn(done || active ? "text-slate-900" : "text-slate-500")}>{label}</span></li>;
      })}
    </ol>
    <div className="mt-6 border-t border-slate-200 pt-4">
      <h3 className="text-sm font-semibold text-slate-900">本次处理选项</h3>
      <div className="mt-3 grid gap-2 text-sm text-slate-700 sm:grid-cols-3">
        <label className="flex items-start gap-2"><input type="checkbox" checked readOnly className="mt-0.5 h-4 w-4 accent-blue-700" /><span>OA 历史保留，并新增下一个月</span></label>
        <label className="flex items-start gap-2"><input type="checkbox" checked readOnly className="mt-0.5 h-4 w-4 accent-blue-700" /><span>缺少来源数据留空，不填 0</span></label>
        <label className="flex items-start gap-2"><input type="checkbox" checked readOnly className="mt-0.5 h-4 w-4 accent-blue-700" /><span>结果使用指定验收总表</span></label>
      </div>
    </div>
    <div className="mt-6 border-t border-slate-200 pt-4">
      <h3 className="text-sm font-semibold text-slate-900">处理事件</h3>
      <ol className="mt-3 space-y-3" aria-label="科园案例处理事件">
        {stageEvents.map((event) => <li key={event.stage} className="flex gap-3 text-sm"><CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0 text-emerald-600" /><div className="min-w-0 flex-1"><p className="text-slate-800">{event.label}</p><p className="mt-0.5 text-xs leading-5 text-slate-500">{KEYUAN_EVENT_DETAILS[event.stage]}</p></div>{event.at ? <time className="shrink-0 text-[11px] tabular-nums text-slate-400">{new Date(event.at).toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" })}</time> : null}</li>)}
      </ol>
    </div>
    {failed ? <div className="mt-5 flex justify-end border-t border-slate-200 pt-4"><Button type="button" size="sm" variant="outline" onClick={onRetry} disabled={busy}><Activity className="mr-1.5 h-3.5 w-3.5" />重新执行</Button></div> : run.status === "completed" ? <div className="mt-5 flex justify-end gap-2 border-t border-slate-200 pt-4"><Button type="button" size="sm" variant="outline" onClick={onRetry} disabled={busy}><Activity className="mr-1.5 h-3.5 w-3.5" />重新执行流程</Button><Button type="button" size="sm" className="bg-blue-700 text-white hover:bg-blue-800" onClick={onDownload} disabled={busy}><FileSpreadsheet className="mr-1.5 h-3.5 w-3.5" />下载结果表格</Button></div> : null}
  </section>;
}

function isEncryptedWorkbookError(error: unknown) {
  return error instanceof Error && /加密|打开密码/.test(error.message);
}

type UploadRole = "master" | "source" | "material";

function fileRole(file: FileMeta): "master" | "source" | "other" {
  if (["financial_master", "template"].includes(file.file_type)) return "master";
  if (["financial_source", "source"].includes(file.file_type)) return "source";
  return "other";
}

function roleLabel(role: "master" | "source" | "other") {
  return role === "master" ? "总表" : role === "source" ? "更新表" : "附件";
}

function roleTone(role: "master" | "source" | "other") {
  return role === "master"
    ? "bg-blue-50 text-blue-700 ring-blue-100"
    : role === "source"
    ? "bg-slate-100 text-slate-600 ring-slate-200"
    : "bg-white text-slate-500 ring-slate-200";
}

function isEncryptedUploadFailure(message: string) {
  return /加密|打开密码|密码/.test(message);
}

function materialKindLabel(kind: AgentMaterial["kind"]) {
  const labels: Record<AgentMaterial["kind"], string> = {
    manual: "操作手册",
    transcript: "文本记录",
    recording: "录音材料",
    rule_package: "规则材料",
    unknown: "说明文档",
  };
  return labels[kind];
}

function FileAttachmentList({
  files,
  materials = [],
  onPreview,
  onRemove,
  compact = false,
  collapsible = false,
}: {
  files: FileMeta[];
  materials?: AgentMaterial[];
  onPreview: (file: FileMeta) => void;
  onRemove?: (file: FileMeta) => void;
  compact?: boolean;
  collapsible?: boolean;
}) {
  const [expanded, setExpanded] = useState(!collapsible);
  const masters = files.filter((file) => fileRole(file) === "master");
  const sources = files.filter((file) => fileRole(file) === "source");
  const other = files.filter((file) => fileRole(file) === "other");
  const groups = [
    { key: "master", label: "总表", files: masters, icon: Table2 },
    { key: "source", label: "更新表", files: sources, icon: FileSpreadsheet },
    { key: "other", label: "附件", files: other, icon: FileText },
  ].filter((group) => group.files.length);

  if (!groups.length && !materials.length) return null;

  return (
    <div className={cn("space-y-2", compact ? "" : "rounded-lg border border-slate-200 bg-slate-50/70 p-2.5")} aria-label="已上传文件">
      {collapsible ? (
        <button type="button" className="flex w-full items-center justify-between gap-3 rounded-lg px-1.5 py-1 text-left text-xs text-slate-600 transition-colors hover:bg-slate-100" onClick={() => setExpanded((value) => !value)} aria-expanded={expanded}>
          <span className="flex min-w-0 items-center gap-2"><Paperclip className="h-3.5 w-3.5 shrink-0 text-slate-400" /><span className="truncate">本次文件 · {files.length + materials.length} 个</span></span>
          <ChevronDown className={cn("h-3.5 w-3.5 shrink-0 text-slate-400 transition-transform", expanded && "rotate-180")} />
        </button>
      ) : !compact ? <p className="px-1.5 pb-0.5 text-[11px] font-medium text-slate-500">本次文件</p> : null}
      {expanded ? <div className={cn("space-y-2", collapsible && "max-h-64 overflow-y-auto overscroll-contain")}>{groups.map((group) => {
        const Icon = group.icon;
        return (
          <div key={group.key} className="space-y-1.5">
            <div className="flex items-center gap-1.5 px-1.5 text-[11px] font-medium text-slate-500">
              <Icon className="h-3.5 w-3.5 text-slate-400" />
              {group.label}
              <span className="text-slate-400">{group.files.length}</span>
            </div>
            {group.files.map((file) => {
              const role = fileRole(file);
              return (
                <div key={file.id} className="flex min-w-0 items-center gap-2 rounded-xl bg-white px-2.5 py-2 ring-1 ring-slate-200/80">
                  <FileSpreadsheet className="h-4 w-4 shrink-0 text-slate-400" />
                  <div className="min-w-0 flex-1">
                    <p className="truncate text-xs font-medium text-slate-800" title={file.original_name}>{file.original_name}</p>
                    <p className="mt-0.5 truncate text-[11px] text-slate-400">
                      {file.row_count || file.col_count ? `${file.row_count} 行 · ${file.col_count} 列` : "已接收，待分析"}
                      {file.sheet_name ? ` · ${file.sheet_name}` : ""}
                    </p>
                  </div>
                  <span className={cn("shrink-0 rounded-full px-1.5 py-0.5 text-[10px] font-medium ring-1", roleTone(role))}>{roleLabel(role)}</span>
                  <button type="button" className="inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-slate-400 transition-colors hover:bg-slate-100 hover:text-slate-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500" onClick={() => onPreview(file)} aria-label={`预览 ${file.original_name}`} title="预览文件">
                    <Eye className="h-3.5 w-3.5" />
                  </button>
                  {onRemove ? <button type="button" className="inline-flex h-7 w-7 shrink-0 items-center justify-center rounded-full text-slate-400 transition-colors hover:bg-red-50 hover:text-red-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-500" onClick={() => onRemove(file)} aria-label={`移除 ${file.original_name}`} title="移除文件"><Trash2 className="h-3.5 w-3.5" /></button> : null}
                </div>
              );
            })}
          </div>
        );
      })}
      {materials.length ? <div className="space-y-1.5">
        <div className="flex items-center gap-1.5 px-1.5 text-[11px] font-medium text-slate-500"><FileText className="h-3.5 w-3.5 text-slate-400" />说明与手册<span className="text-slate-400">{materials.length}</span></div>
        {materials.map((material) => <div key={material.material_id} className="flex min-w-0 items-center gap-2 rounded-xl bg-white px-2.5 py-2 ring-1 ring-slate-200/80"><FileText className="h-4 w-4 shrink-0 text-slate-400" /><div className="min-w-0 flex-1"><p className="truncate text-xs font-medium text-slate-800" title={material.filename}>{material.filename}</p><p className="mt-0.5 truncate text-[11px] text-slate-400">已接收，作为公司上下文材料</p></div><span className="shrink-0 rounded-full bg-violet-50 px-1.5 py-0.5 text-[10px] font-medium text-violet-700 ring-1 ring-violet-100">{materialKindLabel(material.kind)}</span></div>)}
      </div> : null}</div> : null}
    </div>
  );
}

function FilePreviewDrawer({
  file,
  preview,
  loading,
  error,
  onClose,
}: {
  file: FileMeta | null;
  preview: FilePreview | null;
  loading: boolean;
  error: string | null;
  onClose: () => void;
}) {
  if (!file) return null;
  const columns = preview?.columns.slice(0, 12) || [];
  const rows = preview?.rows.slice(0, 20) || [];
  const role = fileRole(file);
  return (
    <div className="fixed inset-0 z-[55] flex justify-end bg-slate-950/20" role="dialog" aria-modal="true" aria-labelledby="file-preview-title" onClick={onClose}>
      <aside className="flex h-full w-full max-w-2xl flex-col border-l border-slate-200 bg-white shadow-2xl" onClick={(event) => event.stopPropagation()}>
        <div className="flex min-h-14 items-center justify-between gap-3 border-b border-slate-200 px-5">
          <div className="min-w-0">
            <div className="flex items-center gap-2">
              <h2 id="file-preview-title" className="truncate text-sm font-semibold text-slate-950">{file.original_name}</h2>
              <span className={cn("shrink-0 rounded-full px-1.5 py-0.5 text-[10px] font-medium ring-1", roleTone(role))}>{roleLabel(role)}</span>
            </div>
            <p className="mt-1 text-[11px] text-slate-400">{file.sheet_name || "默认 Sheet"} · {file.row_count} 行 · {file.col_count} 列</p>
          </div>
          <button type="button" className="inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-slate-400 hover:bg-slate-100 hover:text-slate-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500" onClick={onClose} aria-label="关闭文件预览"><X className="h-4 w-4" /></button>
        </div>
        <div className="min-h-0 flex-1 overflow-auto p-4 sm:p-5">
          {loading ? <div className="flex min-h-40 items-center justify-center text-sm text-slate-500"><Loader2 className="mr-2 h-4 w-4 animate-spin text-blue-600" />正在读取文件前几行...</div> : null}
          {!loading && error ? <div className="flex items-start gap-2 rounded-lg border border-red-200 bg-red-50 px-3 py-3 text-sm leading-6 text-red-800" role="alert"><AlertCircle className="mt-0.5 h-4 w-4 shrink-0" />{error}</div> : null}
          {!loading && !error && preview ? (
            <>
              <div className="mb-3 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-500"><span>显示前 {rows.length} 行</span><span>共 {preview.row_count} 行</span><span>共 {preview.col_count} 列</span></div>
              <div className="overflow-hidden rounded-lg border border-slate-200">
                <div className="overflow-auto">
                  <table className="min-w-full border-collapse text-xs">
                    <thead className="bg-slate-50 text-left text-slate-500">
                      <tr>{columns.map((column) => <th key={column} className="whitespace-nowrap border-b border-slate-200 px-3 py-2 font-medium">{column || "未命名列"}</th>)}</tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {rows.map((row, index) => <tr key={index} className="align-top hover:bg-slate-50/70">{columns.map((column) => <td key={`${index}-${column}`} className="max-w-56 whitespace-nowrap px-3 py-2 text-slate-700">{formatValue(row[column])}</td>)}</tr>)}
                    </tbody>
                  </table>
                </div>
                {!rows.length ? <p className="px-3 py-8 text-center text-xs text-slate-400">这个 Sheet 暂无可显示数据。</p> : null}
              </div>
              {preview.col_count > columns.length ? <p className="mt-2 text-[11px] text-slate-400">为保持预览清晰，仅显示前 {columns.length} 列。</p> : null}
            </>
          ) : null}
        </div>
      </aside>
    </div>
  );
}

function ContextPanel({
  project,
  run,
  files,
  memory,
  modelStatus,
  items,
  onPreview,
  onRemove,
}: {
  project: Project | null;
  run: AgentRun | null;
  files: FileMeta[];
  memory: AgentMemoryResponse | null;
  modelStatus: AgentModelStatus | null;
  items: AgentWorkItem[];
  onPreview: (file: FileMeta) => void;
  onRemove?: (file: FileMeta) => void;
}) {
  const pending = items.filter((item) => ["pending", "needs_conversation"].includes(item.status)).length;
  const applied = items.filter((item) => item.status === "applied").length;
  const failed = items.filter((item) => item.status === "failed").length;
  const runStatus = run ? RUN_LABEL[run.status] || run.status : "未开始";

  return (
    <div className="h-full overflow-y-auto bg-white">
      <div className="border-b border-slate-200 px-5 py-5">
        <p className="text-[11px] font-semibold uppercase tracking-[0.14em] text-slate-400">公司上下文</p>
        <h2 className="mt-2 truncate text-base font-semibold text-slate-950">{project?.name || "加载中"}</h2>
        <p className="mt-1 text-xs text-slate-500">所属月 {project?.salary_month || "尚未设置"}</p>
      </div>

      <div className="divide-y divide-slate-100">
        <section className="px-5 py-5">
          <div className="mb-3 flex items-center gap-2 text-sm font-medium text-slate-800">
            <ShieldCheck className="h-4 w-4 text-blue-600" />
            处理状态
          </div>
          <div className="flex items-center justify-between text-xs">
            <span className="text-slate-500">当前批次</span>
            <span className="font-medium text-slate-800">{runStatus}</span>
          </div>
          <div className="mt-4 grid grid-cols-3 gap-2 text-center">
            <div className="bg-slate-50 px-2 py-2.5"><p className="text-lg font-semibold text-slate-900">{applied}</p><p className="text-[11px] text-slate-500">已处理</p></div>
            <div className="bg-amber-50 px-2 py-2.5"><p className="text-lg font-semibold text-amber-700">{pending}</p><p className="text-[11px] text-amber-700">待回答</p></div>
            <div className="bg-red-50 px-2 py-2.5"><p className="text-lg font-semibold text-red-700">{failed}</p><p className="text-[11px] text-red-700">失败</p></div>
          </div>
        </section>

        <section className="px-5 py-5">
          <div className="mb-3 flex items-center gap-2 text-sm font-medium text-slate-800">
            <FileSpreadsheet className="h-4 w-4 text-blue-600" />
            本次文件
          </div>
          {files.length ? (
            <FileAttachmentList files={files.slice(0, 8)} onPreview={onPreview} onRemove={onRemove} compact />
          ) : <p className="text-xs leading-5 text-slate-500">上传总表和来源文件后，Agent 会自动识别文件角色。</p>}
        </section>

        <section className="px-5 py-5">
          <div className="mb-3 flex items-center gap-2 text-sm font-medium text-slate-800">
            <Sparkles className="h-4 w-4 text-blue-600" />
            公司记忆
          </div>
          {memory?.rules?.length ? (
            <div className="space-y-3">
              {memory.rules.slice(0, 4).map((rule) => (
                <div key={rule.rule_id} className="border-l-2 border-blue-200 pl-3">
                  <p className="text-xs font-medium text-slate-800">{rule.title}</p>
                  <p className="mt-1 line-clamp-3 text-[11px] leading-4 text-slate-500">{rule.description}</p>
                </div>
              ))}
              {memory.rules.length > 4 ? <p className="text-[11px] text-slate-400">已加载 {memory.rules.length} 条公司经验</p> : null}
            </div>
          ) : <p className="text-xs leading-5 text-slate-500">本公司还没有已确认经验。你在对话中的明确回答会逐步沉淀为记忆。</p>}
        </section>

        <section className="px-5 py-5 text-[11px] leading-5 text-slate-500">
          模型：{modelStatus?.configured ? modelStatus.model : "未连接"}
          {modelStatus?.connectivity_check ? (
            <>（{modelStatus.connectivity_check.ok ? "连接正常" : `自检失败：${modelStatus.connectivity_check.detail}`}）</>
          ) : null}<br />
          写入边界：Agent 提议，执行器校验后写入。<br />
          录音：仅存档，不自动转写。
        </section>
      </div>
    </div>
  );
}

function PendingQuestionsCard({ content, busy, onReply }: { content: string; busy?: boolean; onReply: (reply: string) => void }) {
  const questions = useMemo(() => parsePendingQuestions(content), [content]);
  const [customDraft, setCustomDraft] = useState("");
  if (!questions.length) return null;
  // A quick reply is a one-shot action. Hide the whole card while the
  // request is in flight so the user cannot mistake the old choices for
  // still-pending work or submit a second answer.
  if (busy) return null;
  const sendAnswer = (index: number, reply: string) => {
    onReply(`关于你的问题“${questions[index]}”：${reply}`);
  };
  const sendCustomAnswer = () => {
    const reply = customDraft.trim();
    if (!reply || busy) return;
    onReply(reply);
    setCustomDraft("");
  };
  return (
    <div className="mt-3 rounded-xl border border-blue-200 bg-blue-50/40 p-3.5 text-left" role="group" aria-label="待确认问题">
      <p className="text-xs font-semibold text-blue-900">请拍板以下问题（点选项或自己写）</p>
      <div className="mt-2.5 space-y-3">
        {questions.map((question, index) => (
          <div key={index} className="rounded-lg border border-slate-200 bg-white p-3">
            <p className="text-[13px] leading-6 text-slate-800">{index + 1}. {question}</p>
            <div className="mt-2.5 flex flex-wrap gap-2">
              {quickOptionsForQuestion(question).map((option) => (
                <button key={option.label} type="button" onClick={() => sendAnswer(index, option.reply)} className="min-h-9 rounded-full border border-blue-200 bg-blue-50 px-3 py-1.5 text-xs font-medium text-blue-800 transition-colors hover:border-blue-400 hover:bg-blue-100 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500">
                  {option.label}
                </button>
              ))}
            </div>
          </div>
        ))}
      </div>
      <div className="mt-3">
        <label htmlFor="pending-custom-reply" className="sr-only">自定义回复</label>
        <textarea
          id="pending-custom-reply"
          value={customDraft}
          onChange={(event) => setCustomDraft(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter" && !event.shiftKey) {
              event.preventDefault();
              sendCustomAnswer();
            }
          }}
          disabled={busy}
          rows={2}
          maxLength={1000}
          className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs leading-5 text-slate-800 outline-none transition-shadow placeholder:text-slate-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 disabled:cursor-not-allowed disabled:bg-slate-50"
          placeholder="也可以直接写补充说明或对多个问题的统一口径，回车发送"
        />
        <button
          type="button"
          disabled={busy || !customDraft.trim()}
          onClick={sendCustomAnswer}
          className="mt-2 min-h-9 rounded-lg bg-blue-700 px-4 py-2 text-xs font-medium text-white transition-colors hover:bg-blue-800 disabled:cursor-not-allowed disabled:opacity-45"
        >
          发送补充说明
        </button>
      </div>
    </div>
  );
}

function MessageBubble({ message, evidence, showChoices, busy, onChoice, onQuickReply, isLatestAgentMessage }: {
  message: AgentMessage;
  evidence?: AgentWorkItem | null;
  showChoices?: boolean;
  busy?: boolean;
  onChoice?: (action: AgentDecisionAction, value?: string | number | boolean | null) => void;
  onQuickReply?: (reply: string) => void;
  isLatestAgentMessage?: boolean;
}) {
  const isUser = message.role === "user";
  const [expanded, setExpanded] = useState(false);
  const isLongAgentMessage = !isUser && message.content.length > 620;
  const readableContent = cleanAgentContent(message.content);
  const displayContent = isLongAgentMessage && !expanded
    ? `${readableContent.slice(0, 620).trimEnd()}...`
    : readableContent;
  return (
    <div className={cn("flex gap-3", isUser && "justify-end")}>
      {!isUser ? <AgentMascot className="mt-1 h-7 w-5" /> : null}
      <div className={cn("max-w-[min(740px,88%)]", isUser && "text-right")}>
        <p className="mb-1.5 text-[11px] font-medium text-slate-400">{isUser ? "你" : "财务 Agent"}</p>
        <div className={cn("whitespace-pre-wrap text-[15px] leading-7", isUser ? "inline-block rounded-xl bg-slate-100 px-4 py-2 text-left text-slate-800" : "text-slate-700")}>
          {displayContent}
        </div>
        {isLongAgentMessage ? <button type="button" className="mt-2 text-xs font-medium text-blue-700 hover:text-blue-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500" onClick={() => setExpanded((value) => !value)} aria-expanded={expanded}>{expanded ? "收起核对详情" : "展开核对详情"}</button> : null}
        {!isUser && isLatestAgentMessage && onQuickReply ? <PendingQuestionsCard content={message.content} busy={busy} onReply={onQuickReply} /> : null}
        {!isUser && evidence ? (
          <div className="mt-3 border-l-2 border-slate-200 pl-3 text-left">
            <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
              <span className="font-medium text-slate-700">{evidence.person_name || evidence.employee_ref || "待确认人员"} · {evidence.category || evidence.field || "字段变更"}</span>
              <span className={cn("font-medium", statusTone(evidence.status))}>{evidence.status === "needs_conversation" ? "等待你的回答" : evidence.status === "applied" ? "已写入草稿" : evidence.status === "failed" ? "处理失败" : "待处理"}</span>
            </div>
            <div className="mt-2 flex flex-wrap gap-x-6 gap-y-1 text-xs text-slate-500">
              <span>总表：<strong className="font-mono font-medium text-slate-800">{formatValue(evidence.current_value)}</strong></span>
              <span>建议：<strong className="font-mono font-medium text-blue-700">{formatValue(evidence.proposed_value)}</strong></span>
            </div>
            {showChoices && !busy && evidence.status === "needs_conversation" && onChoice ? (
              <div className="mt-4 grid grid-cols-1 gap-2 sm:grid-cols-2" role="group" aria-label="选择处理方式">
                {evidence.proposed_value !== null && evidence.proposed_value !== undefined ? <button
                  type="button"
                  disabled={busy}
                  title="采用 Agent 建议 (Alt+A)"
                  onClick={() => onChoice("apply", evidence.proposed_value)}
                  className="rounded-lg border border-blue-200 bg-blue-50 px-3 py-2 text-left text-xs text-blue-800 transition-colors hover:border-blue-400 hover:bg-blue-100 disabled:cursor-not-allowed disabled:opacity-45"
                ><span className="block font-medium">采用 Agent 建议</span><span className="mt-0.5 block text-[11px] text-blue-700">写入 {formatValue(evidence.proposed_value)}</span></button> : null}
                {Array.from(new Set((evidence.candidate_values || []).filter((value) => value !== null && value !== undefined && value !== "").map((value) => String(value)))).filter((candidate) => String(evidence.proposed_value) !== candidate).slice(0, 4).map((candidate) => {
                  const rawValue = (evidence.candidate_values || []).find((value) => String(value) === candidate);
                  return <button key={`candidate-${candidate}`} type="button" disabled={busy} onClick={() => onChoice("apply", rawValue as string | number | boolean)} className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-left text-xs text-slate-700 transition-colors hover:border-blue-400 hover:bg-blue-50 disabled:cursor-not-allowed disabled:opacity-45"><span className="block font-medium text-slate-900">选择来源值</span><span className="mt-0.5 block text-[11px] text-slate-500">写入 {candidate}</span></button>;
                })}
                <button
                  type="button"
                  disabled={busy}
                  title="保留总表原值 (Alt+K)"
                  onClick={() => onChoice("keep_current", evidence.current_value)}
                  className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-left text-xs text-slate-700 transition-colors hover:border-slate-500 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-45"
                >
                  <span className="block font-medium text-slate-900">保留总表原值</span>
                  <span className="mt-0.5 block text-[11px] text-slate-500">保持 {formatValue(evidence.current_value)}</span>
                </button>
                <button
                  type="button"
                  disabled={busy}
                  title="跳过此项 (Alt+S)"
                  onClick={() => onChoice("skip")}
                  className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-left text-xs text-slate-700 transition-colors hover:border-slate-500 hover:bg-slate-50 disabled:cursor-not-allowed disabled:opacity-45"
                >
                  <span className="block font-medium text-slate-900">跳过此项</span>
                  <span className="mt-0.5 block text-[11px] text-slate-500">本次不处理</span>
                </button>
              </div>
            ) : null}
          </div>
        ) : null}
      </div>
    </div>
  );
}

function ReviewGroup({ group, selectedIds, busy, onToggleItem, onToggleGroup, onApply }: {
  group: AgentItemGroup;
  selectedIds: Set<string>;
  busy: boolean;
  onToggleItem: (itemId: string) => void;
  onToggleGroup: (itemIds: string[]) => void;
  onApply: (items: AgentWorkItem[], action: AgentDecisionAction, customValue?: string, note?: string) => void;
}) {
  const [customValue, setCustomValue] = useState("");
  const [note, setNote] = useState("");
  const selected = group.items.filter((item) => selectedIds.has(item.item_id));
  const allSelected = selected.length === group.items.length;
  const selectedLabel = selected.length ? `已选 ${selected.length} 项` : "勾选要处理的项";
  const apply = (action: AgentDecisionAction, value?: string) => {
    if (!selected.length) return;
    onApply(selected, action, value, note.trim());
  };

  return <section className="border-t border-slate-200 py-4 first:border-t-0" aria-labelledby={`review-group-${group.key}`}>
    <div className="flex flex-wrap items-start justify-between gap-3">
      <div>
        <h3 id={`review-group-${group.key}`} className="text-sm font-semibold text-slate-900">{group.label}</h3>
        <p className="mt-1 text-xs text-slate-500">{group.items.length} 项同类问题 · {selectedLabel}</p>
      </div>
      <label className="flex cursor-pointer items-center gap-2 text-xs font-medium text-slate-600">
        <input type="checkbox" checked={allSelected} onChange={() => onToggleGroup(group.items.map((item) => item.item_id))} disabled={busy} className="h-4 w-4 rounded border-slate-300 text-blue-700 focus:ring-blue-500" />
        全选本类
      </label>
    </div>
    <ul className="mt-3 divide-y divide-slate-100 border-y border-slate-100" aria-label={`${group.label} 问题列表`}>
      {group.items.map((item) => <li key={item.item_id} className="flex items-start gap-2.5 py-2.5">
        <input aria-label={`选择 ${item.person_name || item.employee_ref || "待确认人员"}`} type="checkbox" checked={selectedIds.has(item.item_id)} onChange={() => onToggleItem(item.item_id)} disabled={busy} className="mt-0.5 h-4 w-4 shrink-0 rounded border-slate-300 text-blue-700 focus:ring-blue-500" />
        <div className="min-w-0 flex-1 text-xs leading-5">
          <p className="font-medium text-slate-800">{item.person_name || item.employee_ref || "待确认人员"}</p>
          <p className="text-slate-500">总表 {formatValue(item.current_value)} · 建议 {formatValue(item.proposed_value)}</p>
          {item.reason ? <p className="mt-0.5 text-slate-400">{item.reason}</p> : null}
        </div>
      </li>)}
    </ul>
    <div className="mt-3 flex flex-wrap gap-2">
      <Button type="button" size="sm" variant="outline" disabled={busy || !selected.length} onClick={() => apply("apply")}>采用各自建议</Button>
      <Button type="button" size="sm" variant="outline" disabled={busy || !selected.length} onClick={() => apply("keep_current")}>保留原值</Button>
      <Button type="button" size="sm" variant="outline" disabled={busy || !selected.length} onClick={() => apply("skip")}>跳过</Button>
    </div>
    <div className="mt-3 grid gap-2 sm:grid-cols-[minmax(0,1fr)_auto]">
      <label className="sr-only" htmlFor={`custom-value-${group.key}`}>自定义值</label>
      <input id={`custom-value-${group.key}`} value={customValue} onChange={(event) => setCustomValue(event.target.value)} disabled={busy} className="h-9 rounded-md border border-slate-300 bg-white px-3 text-xs text-slate-800 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 disabled:bg-slate-50" placeholder="自定义值（须是每项都存在的来源候选值）" />
      <Button type="button" size="sm" disabled={busy || !selected.length || !customValue.trim()} onClick={() => apply("apply", customValue.trim())}>应用自定义值</Button>
    </div>
    <label className="mt-2 block" htmlFor={`batch-note-${group.key}`}>
      <span className="sr-only">批量说明</span>
      <input id={`batch-note-${group.key}`} value={note} onChange={(event) => setNote(event.target.value)} disabled={busy} maxLength={1000} className="h-8 w-full rounded-md border border-slate-200 bg-slate-50 px-3 text-[11px] text-slate-700 outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 disabled:bg-slate-100" placeholder="可选：说明本次这组问题的处理依据（会记入审计记录）" />
    </label>
  </section>;
}

function BatchReviewPanel({ items, selectedIds, busy, onToggleItem, onToggleGroup, onApply }: {
  items: AgentWorkItem[];
  selectedIds: Set<string>;
  busy: boolean;
  onToggleItem: (itemId: string) => void;
  onToggleGroup: (itemIds: string[]) => void;
  onApply: (items: AgentWorkItem[], action: AgentDecisionAction, customValue?: string, note?: string) => void;
}) {
  const groups = groupAgentItems(items.filter((item) => item.status === "needs_conversation"));
  if (!groups.length) return null;
  return <section className="my-6 border-y border-slate-200 bg-white px-5" aria-labelledby="batch-review-heading">
    <div className="py-4">
      <h2 id="batch-review-heading" className="text-base font-semibold text-slate-900">同类问题批量处理</h2>
      <p className="mt-1 text-xs leading-5 text-slate-500">先勾选同一字段的问题，再批量处理。采用建议时每个人仍使用自己的来源值；自定义值必须能在所有勾选项的候选来源中找到。</p>
    </div>
    {groups.map((group) => <ReviewGroup key={group.key} group={group} selectedIds={selectedIds} busy={busy} onToggleItem={onToggleItem} onToggleGroup={onToggleGroup} onApply={onApply} />)}
  </section>;
}

export function AgentWorkbenchPage({ projectId }: { projectId: string }) {
  const [project, setProject] = useState<Project | null>(null);
  const [files, setFiles] = useState<FileMeta[]>([]);
  const [materials, setMaterials] = useState<AgentMaterial[]>([]);
  const [memory, setMemory] = useState<AgentMemoryResponse | null>(null);
  const [modelStatus, setModelStatus] = useState<AgentModelStatus | null>(null);
  const [run, setRun] = useState<AgentRun | null>(null);
  const [runResult, setRunResult] = useState<AgentRunResult | null>(null);
  const [processEvents, setProcessEvents] = useState<AgentRunEvent[]>([]);
  const [processRevision, setProcessRevision] = useState(0);
  const [processRunning, setProcessRunning] = useState(false);
  const [stoppingRun, setStoppingRun] = useState(false);
  const [streamStatus, setStreamStatus] = useState<string | null>(null);
  const [items, setItems] = useState<AgentWorkItem[]>([]);
  const [activeReviewId, setActiveReviewId] = useState<string | null>(null);
  const [selectedReviewIds, setSelectedReviewIds] = useState<Set<string>>(new Set());
  const [runs, setRuns] = useState<AgentRun[]>([]);
  const [pendingPlanInstruction, setPendingPlanInstruction] = useState("");
  const [draft, setDraft] = useState("");
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [uploadMenuOpen, setUploadMenuOpen] = useState(false);
  const [previewFile, setPreviewFile] = useState<FileMeta | null>(null);
  const [preview, setPreview] = useState<FilePreview | null>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewError, setPreviewError] = useState<string | null>(null);
  const [uploadFailures, setUploadFailures] = useState<Array<{ filename: string; message: string }>>([]);
  const [showContext, setShowContext] = useState(false);
  const [passwordDialogOpen, setPasswordDialogOpen] = useState(false);
  const [passwordDraft, setPasswordDraft] = useState("");
  const [passwordFiles, setPasswordFiles] = useState<File[]>([]);
  const [passwordRole, setPasswordRole] = useState<"master" | "source" | null>(null);
  const [passwordError, setPasswordError] = useState<string | null>(null);
  const [passwordSubmitting, setPasswordSubmitting] = useState(false);
  const [composerHeight, setComposerHeight] = useState(220);
  const [resumePrompt, setResumePrompt] = useState<{ runId: string; status: string; detail: string } | null>(null);
  const masterInputRef = useRef<HTMLInputElement>(null);
  const uploadMenuRef = useRef<HTMLDivElement>(null);
  const prevRunStatusRef = useRef<string | null>(null);
  const latestRunStatusRef = useRef<string | null>(null);
  const latestRunIdRef = useRef<string | null>(null);
  const resumePromptTimerRef = useRef<number | null>(null);
  const sourceInputRef = useRef<HTMLInputElement>(null);
  const materialInputRef = useRef<HTMLInputElement>(null);
  const bottomRef = useRef<HTMLDivElement>(null);
  const scrollContainerRef = useRef<HTMLDivElement | null>(null);
  const resultAnchorRef = useRef<HTMLDivElement | null>(null);
  // 用户是否“贴底跟随”：往上翻回看历史时暂停自动跟随，
  // 滚回底部附近后恢复，避免阅读被新事件不断拽到底部。
  const stickToBottomRef = useRef(true);
  const activeReviewRef = useRef<HTMLDivElement>(null);
  const composerDockRef = useRef<HTMLDivElement>(null);
  const processRefreshInFlightRef = useRef(false);
  const processRevisionRef = useRef(0);
  const autoResumeRunRef = useRef<string | null>(null);
  const messageAbortControllerRef = useRef<AbortController | null>(null);

  useEffect(() => () => {
    messageAbortControllerRef.current?.abort();
  }, []);

  const load = useCallback(async () => {
    setLoading(true);
    const [projectResult, filesResult, materialsResult, memoryResult, modelResult, runsResult] = await Promise.allSettled([
      api.getProject(projectId),
      api.listFiles(projectId),
      api.listAgentMaterials(projectId),
      api.listAgentMemory(projectId),
      api.getAgentModelStatus(),
      api.listAgentRuns(projectId),
    ]);
    if (projectResult.status === "fulfilled") setProject(projectResult.value);
    if (filesResult.status === "fulfilled") setFiles(filesResult.value);
    if (materialsResult.status === "fulfilled") setMaterials(materialsResult.value.items);
    if (memoryResult.status === "fulfilled") setMemory(memoryResult.value);
    if (modelResult.status === "fulfilled") setModelStatus(modelResult.value);
    if (runsResult.status === "fulfilled") {
      setRuns(runsResult.value.items);
      const latest = runsResult.value.items[0];
      if (latest) {
        setRun(latest);
        try { setItems(await api.listAgentItems(latest.run_id)); } catch { setItems([]); }
      }
    }
    setLoading(false);
  }, [projectId]);

  useEffect(() => { void load(); }, [load]);
  useEffect(() => { setPendingPlanInstruction(""); }, [run?.run_id]);
  // 模型连通性自检：配置存在时后台发一次小请求，数秒内暴露地址/密钥/模型名错误。
  useEffect(() => {
    if (!modelStatus?.configured || run?.execution_mode === "demo") return;
    let disposed = false;
    api.getAgentModelStatus(true)
      .then((checked) => { if (!disposed) setModelStatus(checked); })
      .catch(() => { /* 自检失败不阻塞页面，仅保持原状态 */ });
    return () => { disposed = true; };
  }, [modelStatus?.configured, run?.execution_mode]);
  // 新事件/新消息只在用户贴近底部时才自动跟随；上翻回看时不打扰。
  const handleChatScroll = useCallback(() => {
    const element = scrollContainerRef.current;
    if (!element) return;
    const distance = element.scrollHeight - element.scrollTop - element.clientHeight;
    stickToBottomRef.current = distance < 240;
  }, []);
  useEffect(() => {
    if (!stickToBottomRef.current) return;
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [items, processEvents, run]);
  // 结果生成后把结果面板带入视口：任务完成是用户等待的关键节点，
  // 无论当前滚动位置如何都滚动一次，避免结果被长对话顶出视野。
  useEffect(() => {
    if (!runResult?.available) return;
    window.requestAnimationFrame(() => resultAnchorRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }));
  }, [runResult?.available, runResult?.filename]);

  // 上传菜单：点击菜单外任意位置或按 Escape 时关闭，避免小弹窗一直挂着。
  useEffect(() => {
    if (!uploadMenuOpen) return;
    const onPointerDown = (event: MouseEvent) => {
      if (uploadMenuRef.current && !uploadMenuRef.current.contains(event.target as Node)) {
        setUploadMenuOpen(false);
      }
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setUploadMenuOpen(false);
    };
    document.addEventListener("mousedown", onPointerDown);
    document.addEventListener("keydown", onKeyDown);
    return () => {
      document.removeEventListener("mousedown", onPointerDown);
      document.removeEventListener("keydown", onKeyDown);
    };
  }, [uploadMenuOpen]);

  useEffect(() => {
    const element = composerDockRef.current;
    if (!element || typeof ResizeObserver === "undefined") return;
    const updateHeight = () => setComposerHeight(Math.ceil(element.getBoundingClientRect().height));
    updateHeight();
    const observer = new ResizeObserver(updateHeight);
    observer.observe(element);
    return () => observer.disconnect();
  }, [files.length, materials.length, uploadFailures.length, run?.month_confirmation?.required, run?.month_confirmation?.confirmed, run?.status]);

  const refreshProcessEvents = useCallback(async (runId: string, reset = false) => {
    if (processRefreshInFlightRef.current && !reset) return;
    processRefreshInFlightRef.current = true;
    try {
      const response = await api.listAgentEvents(runId, reset ? 0 : processRevision);
      if (response.run_id !== runId) return;
      if (response.events.length) {
        setProcessEvents((current) => {
          const base = reset ? [] : current;
          const seen = new Set(base.map((event) => event.event_id));
          return [...base, ...response.events.filter((event) => !seen.has(event.event_id))].slice(-120);
        });
      }
      setProcessRevision(response.revision);
      processRevisionRef.current = response.revision;
    } catch {
      // Progress is best-effort. The final run response remains authoritative.
    } finally {
      processRefreshInFlightRef.current = false;
    }
  }, [processRevision]);

  useEffect(() => {
    if (!run?.run_id) return;
    processRevisionRef.current = 0;
    void refreshProcessEvents(run.run_id, true);
  // A new run is the only time the event cursor should reset.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [run?.run_id]);

  useEffect(() => {
    if (!run?.run_id || run.status !== "processing") return;
    let disposed = false;
    let retryTimer: number | undefined;
    const controller = new AbortController();
    const follow = async () => {
      try {
        await api.streamAgentEvents(run.run_id, processRevisionRef.current, (event) => {
          processRevisionRef.current = Math.max(processRevisionRef.current, event.revision || 0);
          setProcessRevision(processRevisionRef.current);
          setProcessEvents((current) => {
            if (current.some((item) => item.event_id === event.event_id)) return current;
            return [...current, event].slice(-120);
          });
        }, controller.signal);
        const latest = normaliseRun(await api.getAgentRun(run.run_id), projectId);
        if (!disposed && latest) {
          setRun(latest);
          setItems(await api.listAgentItems(run.run_id));
        }
      } catch (error) {
        if (!disposed && !(error instanceof DOMException && error.name === "AbortError")) {
          retryTimer = window.setTimeout(() => { void follow(); }, 1200);
        }
      }
    };
    void follow();
    return () => {
      disposed = true;
      controller.abort();
      if (retryTimer) window.clearTimeout(retryTimer);
    };
  }, [projectId, run?.run_id, run?.status]);

  useEffect(() => {
    const namedShowcase = project?.name === "北京" || project?.name === "样本一" || project?.name === "样本二" || project?.name === "样本四";
    const shouldResume = isResumableAgentRun(run) || (namedShowcase && run?.status === "processing");
    if (!run?.run_id || !shouldResume || autoResumeRunRef.current === run.run_id) return;
    autoResumeRunRef.current = run.run_id;
    let disposed = false;
    const timer = window.setTimeout(() => {
      void (async () => {
        setProcessRunning(true);
        try {
          const resumed = normaliseRun(await api.processAgentRun(run.run_id), projectId);
          if (disposed) return;
          if (resumed) setRun(resumed);
          await refreshProcessEvents(run.run_id);
          setItems(await api.listAgentItems(run.run_id));
          toast.info("Agent 正在从已保存的进度自动继续处理");
        } catch (error) {
          if (!disposed) toast.error(error instanceof Error ? error.message : "自动续跑启动失败");
        } finally {
          if (!disposed) setProcessRunning(false);
        }
      })();
    }, 350);
    return () => { disposed = true; window.clearTimeout(timer); };
  }, [project?.name, projectId, refreshProcessEvents, run]);

  // 处理从“进行中”转入“需要用户继续”且自动续跑未接管时，弹窗提醒。
  useEffect(() => {
    const previous = prevRunStatusRef.current;
    prevRunStatusRef.current = run?.status ?? null;
    latestRunStatusRef.current = run?.status ?? null;
    latestRunIdRef.current = run?.run_id ?? null;
    if (!run?.run_id) return;
    const wasActive = previous === "processing" || previous === "planning";
    if (!wasActive || !RESUME_PROMPT_STATUSES.has(run.status)) return;
    // 用户主动停止（USER_STOPPED）是明确意图，不再弹窗催促继续，
    // 也不发系统通知；只有 Agent 自身中断才提醒用户接管。
    if (run.status === "execution_incomplete" && run.code === "USER_STOPPED") return;
    const { run_id: runId, status, detail } = run;
    if (resumePromptTimerRef.current) window.clearTimeout(resumePromptTimerRef.current);
    resumePromptTimerRef.current = window.setTimeout(() => {
      resumePromptTimerRef.current = null;
      // 延迟期间自动续跑已接管（状态回到活跃）则不弹。
      if (latestRunIdRef.current !== runId || !RESUME_PROMPT_STATUSES.has(latestRunStatusRef.current || "")) return;
      setResumePrompt({ runId, status, detail: String(detail || "") });
      sendAgentSystemNotification("外服账单系统", "Agent 已暂停，需要你发送“继续”才会接着处理。");
    }, RESUME_PROMPT_DELAY_MS);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [run?.run_id, run?.status, run?.updated_at]);

  useEffect(() => () => {
    if (resumePromptTimerRef.current) window.clearTimeout(resumePromptTimerRef.current);
  }, []);

  // Escape 统一关闭最上层弹层：弹窗 div 上的 onKeyDown 只在焦点位于
  // 弹窗内时生效，而弹窗打开时焦点往往还留在背后页面，导致 Escape
  // 无响应。这里按层叠顺序逐层关闭（密码弹窗 > 等待继续 > 文件预览 >
  // 公司上下文 > 上传菜单），提交中或加载中的关键弹窗不响应。
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      if (passwordDialogOpen) { dismissPasswordDialog(); return; }
      if (resumePrompt) { dismissResumePrompt(); return; }
      if (previewFile) { closeFilePreview(); return; }
      if (showContext) { setShowContext(false); return; }
      setUploadMenuOpen(false);
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  // 函数声明提升后在此处可安全引用；弹层开关状态是最小依赖集。
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [passwordDialogOpen, passwordSubmitting, resumePrompt, previewFile, showContext]);

  useEffect(() => {
    if (!run?.run_id || !["completed", "published", "ready"].includes(run.status)) {
      setRunResult(null);
      return;
    }
    let disposed = false;
    void api.getAgentResult(run.run_id).then((result) => {
      if (!disposed) {
        setRunResult(result.available
          ? project?.name === "北京" && !SHOWCASE_PROJECT_IDS.has(projectId) && run.execution_mode !== "demo"
            ? { ...result, filename: "待确定稿.xlsx" }
            : result
          : null);
      }
    }).catch(() => { if (!disposed) setRunResult(null); });
    return () => { disposed = true; };
  }, [project?.name, projectId, run?.execution_mode, run?.run_id, run?.status, run?.updated_at]);

  async function openFilePreview(file: FileMeta) {
    setPreviewFile(file);
    setPreview(null);
    setPreviewError(null);
    setPreviewLoading(true);
    try {
      setPreview(await api.previewFile(projectId, file.id));
    } catch (error) {
      setPreviewError(error instanceof Error ? error.message : "文件预览失败，请稍后重试");
    } finally {
      setPreviewLoading(false);
    }
  }

  function closeFilePreview() {
    // 加载中同样允许关闭：预览接口卡住时用户能随时退出，
    // 后续返回的数据因 previewFile 已清空而不会写入界面。
    setPreviewFile(null);
    setPreview(null);
    setPreviewError(null);
  }

  // 双击确认替代原生 window.confirm：首次点击进入待确认状态并提示，
  // 5 秒内再次点击才真正移除，超时自动取消。避免阻塞式对话框打断操作。
  const pendingRemovalRef = useRef<string | null>(null);
  const pendingRemovalTimerRef = useRef<number | null>(null);
  async function removeFile(file: FileMeta) {
    if (pendingRemovalRef.current !== file.id) {
      pendingRemovalRef.current = file.id;
      if (pendingRemovalTimerRef.current) window.clearTimeout(pendingRemovalTimerRef.current);
      pendingRemovalTimerRef.current = window.setTimeout(() => { pendingRemovalRef.current = null; }, 5000);
      toast.info(`再次点击移除按钮以确认删除「${file.original_name}」`);
      return;
    }
    pendingRemovalRef.current = null;
    if (pendingRemovalTimerRef.current) { window.clearTimeout(pendingRemovalTimerRef.current); pendingRemovalTimerRef.current = null; }
    setUploading(true);
    try {
      await api.deleteFile(projectId, file.id);
      if (previewFile?.id === file.id) closeFilePreview();
      setFiles((current) => current.filter((item) => item.id !== file.id));
      setUploadFailures((current) => current.filter((failure) => failure.filename !== file.original_name));
      toast.success(`已移除 ${file.original_name}`);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "文件移除失败");
    } finally {
      setUploading(false);
    }
  }

  const activeItem = useMemo(() => {
    const selected = activeReviewId ? items.find((item) => item.item_id === activeReviewId && ["needs_conversation", "pending"].includes(item.status)) : null;
    return selected || items.find((item) => item.status === "needs_conversation") || items.find((item) => item.status === "pending") || null;
  }, [activeReviewId, items]);

  useEffect(() => {
    if (!activeReviewId) return;
    window.requestAnimationFrame(() => activeReviewRef.current?.scrollIntoView({ behavior: "smooth", block: "center" }));
  }, [activeReviewId]);
  // The Agent run endpoint intentionally accepts only the explicit financial
  // roles. Legacy `template`/`source` files can still be previewed, but must
  // be uploaded again through the role picker before starting this flow.
  const masterFiles = useMemo(() => files.filter((file) => file.file_type === "financial_master"), [files]);
  const sourceFiles = useMemo(() => files.filter((file) => file.file_type === "financial_source"), [files]);
  const isShowcaseProject = project?.name === "北京" || project?.name === "样本一" || project?.name === "样本二" || project?.name === "样本四";

  const pendingCount = items.filter((item) => item.status === "needs_conversation").length;
  const processingCount = items.filter((item) => item.status === "pending").length;
  const itemMessages = useMemo(() => {
    const visible = visibleAgentItems(items);
    return visible
      .flatMap((item) => (item.conversation || []).map((message) => ({ message, item })))
      .sort((a, b) => String(a.message.created_at || "").localeCompare(String(b.message.created_at || "")));
  }, [items]);
  const runMessages = useMemo(() => run?.conversation || [], [run?.conversation]);
  const hasConversation = Boolean(run || itemMessages.length);
  const activityEvents = processEvents;

  const startRun = useCallback(async (instruction?: string) => {
    if (!isShowcaseProject && (!masterFiles.length || !sourceFiles.length)) {
      toast.error(!masterFiles.length ? "请先上传一份总表" : "请至少上传一份更新表");
      return;
    }
    setBusy(true);
    setProcessRunning(true);
    setProcessEvents([]);
    setProcessRevision(0);
    try {
      const started = normaliseRun(await api.startAgentRun(projectId, instruction, false, isShowcaseProject), projectId);
      if (!started) throw new Error("服务未返回有效的处理批次");
      setRun(started);
      const processing = normaliseRun(await api.processAgentRun(started.run_id, undefined, true), projectId);
      if (processing) setRun(processing);
      setItems(await api.listAgentItems(started.run_id));
      await refreshProcessEvents(started.run_id, true);
      toast.success("Agent 已开始解析文档并处理表格，关闭页面也会继续");
    } catch (error) { toast.error(error instanceof Error ? error.message : "Agent 启动失败"); }
    finally { setProcessRunning(false); setBusy(false); }
  }, [isShowcaseProject, masterFiles.length, projectId, refreshProcessEvents, sourceFiles.length]);

  async function stopRun() {
    if (!run?.run_id || stoppingRun) return;
    setStoppingRun(true);
    try {
      const stopped = normaliseRun(await api.stopAgentRun(run.run_id), projectId);
      if (stopped) setRun(stopped);
      await refreshProcessEvents(run.run_id);
      toast.info(
        stopped?.status === "execution_incomplete"
          ? "已停止本次处理，已完成的写入和进度均已保留"
          : "已发送停止请求，当前步骤结束后停止并保留进度",
      );
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "停止请求失败");
    } finally {
      setStoppingRun(false);
    }
  }

  async function confirmMonth(selectedSteps?: string[], instruction?: string) {
    if (!run) return;
    setBusy(true);
    setProcessRunning(true);
    try {
      const confirmed = normaliseRun(await api.confirmAgentPlan(run.run_id, selectedSteps, instruction), projectId);
      if (confirmed) setRun(confirmed);
      const executed = normaliseRun(await api.processAgentRun(run.run_id), projectId);
      if (executed) setRun(executed);
      setItems(await api.listAgentItems(run.run_id));
      await refreshProcessEvents(run.run_id);
      toast.success("计划已确认，正在执行本轮处理");
    } catch (error) { toast.error(error instanceof Error ? error.message : "月份确认失败"); }
    finally { setProcessRunning(false); setBusy(false); }
  }

  async function replan(instruction?: string): Promise<boolean> {
    if (!run) return false;
    setBusy(true);
    setProcessRunning(true);
    try {
      // All visible choices are complete at this point. Submit them as the
      // plan confirmation so the backend does not regenerate the same questions.
      const confirmed = normaliseRun(
        await api.confirmAgentPlan(run.run_id, undefined, instruction),
        projectId
      );
      if (confirmed) setRun(confirmed);

      const processing = normaliseRun(await api.processAgentRun(run.run_id), projectId);
      if (processing) setRun(processing);
      setItems(await api.listAgentItems(run.run_id));
      await refreshProcessEvents(run.run_id);
      toast.success("处理口径已确认，Agent 正在自动继续");
      return true;
    } catch (error) {
      if (instruction) setDraft(instruction);
      toast.error(error instanceof Error ? error.message : "处理口径确认失败");
      return false;
    } finally { setProcessRunning(false); setBusy(false); }
  }

  async function executeRun(instruction?: string, startProcessing = false) {
    if (!run) return;
    setBusy(true);
    setProcessRunning(true);
    try {
      const effectiveInstruction = instruction || (startProcessing ? pendingPlanInstruction || undefined : undefined);
      const processing = normaliseRun(await api.processAgentRun(run.run_id, effectiveInstruction, startProcessing), projectId);
      if (processing) setRun(processing);
      await refreshProcessEvents(run.run_id);
      setItems(await api.listAgentItems(run.run_id));
      toast.success(processing?.status === "processing" ? "Agent 已开始执行，处理过程会持续更新" : "Agent 已完成自动核对，仅保留必要的确认项");
    }
    catch (error) { toast.error(error instanceof Error ? error.message : "执行失败"); }
    finally { setProcessRunning(false); setBusy(false); }
  }

  function dismissResumePrompt() {
    setResumePrompt(null);
  }

  async function continueFromResumePrompt() {
    setResumePrompt(null);
    await executeRun("继续");
  }

  function enableSystemNotify() {
    if (typeof window === "undefined" || !("Notification" in window)) return;
    if (Notification.permission !== "default") return;
    void Notification.requestPermission().then((permission) => {
      if (permission === "granted") toast.success("已开启系统通知，切走页面时也会提醒你");
      else if (permission === "denied") toast.info("浏览器已拒绝通知权限，可在地址栏左侧的站点设置中重新允许");
    });
  }

  async function chooseDecision(item: AgentWorkItem, action: AgentDecisionAction, selectedValue?: string | number | boolean | null, note = "") {
    if (!run || busy) return;
    setBusy(true);
    setProcessRunning(true);
    try {
      const applied = await api.applyAgentItem(run.run_id, item.item_id, {
        action,
        value: action === "apply" ? (selectedValue !== undefined ? selectedValue : item.proposed_value) : action === "keep_current" ? item.current_value : undefined,
        note,
        remember: false,
        expected_revision: item.revision,
      });
      setItems((current) => {
        const updated = current.map((candidate) => candidate.item_id === applied.item_id ? applied : candidate);
        setActiveReviewId(nextPendingAgentItem(updated, item.item_id)?.item_id || null);
        return updated;
      });
      const refreshed = normaliseRun(await api.getAgentRun(run.run_id), projectId);
      if (refreshed) setRun(refreshed);
      toast.success(action === "apply" ? "已采用建议并写入草稿" : action === "keep_current" ? "已保留总表原值" : "已跳过此项");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "处理选择失败，请刷新后重试");
    } finally {
      setProcessRunning(false);
      setBusy(false);
    }
  }

  useEffect(() => {
    const handleReviewShortcut = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      if (target?.isContentEditable || ["INPUT", "TEXTAREA", "SELECT"].includes(target?.tagName || "")) return;
      const action = agentShortcutAction(event.key, event.altKey);
      if (!action || !activeItem || busy) return;
      event.preventDefault();
      if (action === "next") {
        setActiveReviewId(nextPendingAgentItem(items, activeItem.item_id)?.item_id || null);
        return;
      }
      void chooseDecision(
        activeItem,
        action,
        action === "apply" ? activeItem.proposed_value : action === "keep_current" ? activeItem.current_value : undefined,
      );
    };
    window.addEventListener("keydown", handleReviewShortcut);
    return () => window.removeEventListener("keydown", handleReviewShortcut);
  // chooseDecision is intentionally the live component callback; the active
  // item and queue dependencies above ensure the handler is refreshed safely.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeItem, busy, items]);

  function toggleReviewItem(itemId: string) {
    setSelectedReviewIds((current) => {
      const next = new Set(current);
      if (next.has(itemId)) next.delete(itemId);
      else next.add(itemId);
      return next;
    });
  }

  function toggleReviewGroup(itemIds: string[]) {
    setSelectedReviewIds((current) => {
      const next = new Set(current);
      const everySelected = itemIds.every((itemId) => next.has(itemId));
      itemIds.forEach((itemId) => everySelected ? next.delete(itemId) : next.add(itemId));
      return next;
    });
  }

  async function applyBatchDecision(selectedItems: AgentWorkItem[], action: AgentDecisionAction, customValue?: string, note = "") {
    if (!run || busy || !selectedItems.length) return;
    const values = new Map<string, string | number | boolean | null | undefined>();
    if (action === "apply") {
      for (const item of selectedItems) {
        if (!customValue) {
          values.set(item.item_id, item.proposed_value);
          continue;
        }
        const candidates = item.candidate_values || [];
        const matched = candidates.find((candidate) => String(candidate) === customValue);
        if (candidates.length && matched === undefined) {
          toast.error(`“${item.person_name || item.employee_ref || item.item_id}”没有来源候选值 ${customValue}，未执行批量写入`);
          return;
        }
        values.set(item.item_id, matched === undefined ? customValue : matched);
      }
    }
    setBusy(true);
    setProcessRunning(true);
    try {
      const resolved = await Promise.all(selectedItems.map(async (item) => {
        const value = action === "apply" ? values.get(item.item_id) : action === "keep_current" ? item.current_value : undefined;
        return api.applyAgentItem(run.run_id, item.item_id, { action, value, note, remember: false, expected_revision: item.revision });
      }));
      const byId = new Map(resolved.map((item) => [item.item_id, item]));
      setItems((current) => current.map((item) => byId.get(item.item_id) || item));
      setSelectedReviewIds((current) => {
        const next = new Set(current);
        selectedItems.forEach((item) => next.delete(item.item_id));
        return next;
      });
      const refreshed = normaliseRun(await api.getAgentRun(run.run_id), projectId);
      if (refreshed) setRun(refreshed);
      toast.success(`已批量处理 ${resolved.length} 项${note ? "，并记录处理说明" : ""}`);
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "批量处理失败，请刷新后查看已完成项");
    } finally {
      setProcessRunning(false);
      setBusy(false);
    }
  }

  async function streamRunChat(runId: string, content: string) {
    setBusy(true);
    setProcessRunning(true);
    const controller = new AbortController();
    messageAbortControllerRef.current = controller;
    let accepted = false;
    setRun((current) => current ? {
      ...current,
      conversation: [
        ...(current.conversation || []),
        { role: "user", content },
        { role: "agent", content: "" },
      ],
    } : current);
    try {
      await api.streamAgentRunMessage(runId, content, (event) => {
        if (event.type === "message.accepted") accepted = true;
        if (event.type === "message.status" && typeof event.payload.label === "string") {
          setStreamStatus(event.payload.label);
        }
        if (event.type === "message.delta" && typeof event.payload.content === "string") {
          setRun((current) => {
            if (!current) return current;
            const conversation = [...(current.conversation || [])];
            const last = conversation.at(-1);
            if (!last || last.role !== "agent") return current;
            conversation[conversation.length - 1] = { ...last, content: `${last.content}${event.payload.content}` };
            return { ...current, conversation };
          });
        }
        if (event.type === "message.completed" && event.payload.message && typeof event.payload.message === "object") {
          const message = event.payload.message as AgentMessage;
          setRun((current) => {
            if (!current) return current;
            const conversation = [...(current.conversation || [])];
            if (conversation.at(-1)?.role === "agent") conversation[conversation.length - 1] = message;
            return { ...current, conversation };
          });
          setStreamStatus(null);
        }
      }, controller.signal);
      await refreshProcessEvents(runId);
    } catch (error) {
      if (controller.signal.aborted) {
        if (!accepted) {
          setRun((current) => current ? { ...current, conversation: (current.conversation || []).slice(0, -2) } : current);
        }
        setDraft(content);
        setStreamStatus(null);
        return;
      }
      if (!accepted) {
        setRun((current) => current ? { ...current, conversation: (current.conversation || []).slice(0, -2) } : current);
        setDraft(content);
      }
      toast.error(error instanceof Error ? error.message : "消息发送失败");
    }
    finally {
      if (messageAbortControllerRef.current === controller) messageAbortControllerRef.current = null;
      setStreamStatus(null);
      setProcessRunning(false);
      setBusy(false);
    }
  }

  async function sendMessage(overrideContent?: string) {
    // 快捷回复（待确认问题的选项/自定义答案）会传入现成文本，
    // 与手动输入走完全相同的路由（对齐/续轮/执行/审核）。
    const content = (overrideContent ?? draft).trim();
    if (!content || busy) return;
    setDraft("");
    if (!run) {
      if (!isShowcaseProject && (!masterFiles.length || !sourceFiles.length)) {
        setDraft(content);
        toast.error(!masterFiles.length ? "请先上传一份总表" : "请至少上传一份更新表");
        return;
      }
      // 首条消息只创建任务并进入对齐：Agent 会基于上传文件复述它的
      // 理解并与用户对齐颗粒度；执行由“开始处理”按钮或指令显式触发。
      setBusy(true);
      try {
        const created = normaliseRun(await api.startAgentRun(projectId, content, false, isShowcaseProject), projectId);
        if (!created) throw new Error("服务未返回有效的处理批次");
        setRun(created);
        await streamRunChat(created.run_id, content);
      } catch (error) {
        toast.error(error instanceof Error ? error.message : "任务创建失败");
        setDraft(content);
      }
      finally { setBusy(false); }
      return;
    }
    // 对齐阶段：执行尚未开始（无 workflow 启动记录、无草稿）。明确的
    // 启动指令触发执行，其余消息一律走对齐对话。
    if (!run.workflow?.started_at && !run.draft_filename) {
      if (isAgentStartCommand(content)) {
        await executeRun(undefined, true);
        return;
      }
      await streamRunChat(run.run_id, content);
      return;
    }
    // 下一轮：上一轮已收尾。新要求先对齐（Agent 会复述理解），明确的
    // 启动指令（或失败后的“继续”）开启下一轮；可解析的表格指令由
    // 后端直接执行，保持快速改数的原有能力。
    const roundEnded = Boolean((run.workflow?.started_at || run.draft_filename)
      && ["ready", "completed", "published", "failed"].includes(run.status));
    if (roundEnded && !activeItem) {
      const isNextRoundStart = isAgentStartCommand(content)
        || (run.status === "failed" && /^(继续|恢复|续跑)$/.test(content.trim()));
      if (isNextRoundStart) {
        await executeRun(isAgentStartCommand(content) ? undefined : content);
        return;
      }
      await streamRunChat(run.run_id, content);
      return;
    }
    // A plan is an internal backend artifact. When an older run still carries
    // the legacy unconfirmed flag, submit the user's continuation as the
    // execution instruction once, then resume processing; do not expose a
    // second plan-confirmation loop.
    if (run.plan_confirmation?.required && !run.plan_confirmation.confirmed) {
      // 计划尚未生成（例如上次模型故障中断在生成前）时，交给后台
      // 工作流带指令重新生成，而不是确认一个不存在的计划导致 409。
      if (run.model_plan) {
        await replan(content);
        return;
      }
      await executeRun(content);
      return;
    }
    if (isResumableAgentRun(run) && isAgentProcessingInstruction(content)) {
      await executeRun(content);
      return;
    }
    if (canResumeDirectWrite(run, items) && isAgentProcessingInstruction(content)) {
      await executeRun(content);
      return;
    }
    if (!activeItem) {
      await streamRunChat(run.run_id, content);
      return;
    }
    if (isResumableAgentRun(run)) {
      await executeRun();
      return;
    }
    setBusy(true);
    setProcessRunning(true);
    try {
      const replied = await api.sendAgentItemMessage(run.run_id, activeItem.item_id, content);
      setItems((current) => current.map((item) => item.item_id === replied.item_id ? replied : item));
      if (replied.decision === "apply_proposed" || replied.decision === "keep_current" || replied.decision === "skip") {
        const action: AgentDecisionAction = replied.decision === "apply_proposed" ? "apply" : replied.decision as AgentDecisionAction;
        const applied = await api.applyAgentItem(run.run_id, replied.item_id, { action, value: action === "apply" ? replied.proposed_value : replied.current_value, remember: false, expected_revision: replied.revision });
        setItems((current) => current.map((item) => item.item_id === applied.item_id ? applied : item));
        toast.success(action === "apply" ? "已写入草稿" : "已按你的回答处理");
      }
      await refreshProcessEvents(run.run_id);
    } catch (error) { toast.error(error instanceof Error ? error.message : "消息发送失败"); }
    finally { setProcessRunning(false); setBusy(false); }
  }

  function stopMessageGeneration() {
    if (!messageAbortControllerRef.current) return;
    messageAbortControllerRef.current.abort();
    toast.info("已停止本次回复生成");
  }

  function openUploadPicker(role: UploadRole) {
    setUploadMenuOpen(false);
    if (role === "master") masterInputRef.current?.click();
    if (role === "source") sourceInputRef.current?.click();
    if (role === "material") materialInputRef.current?.click();
  }

  async function uploadFiles(fileList: FileList | null, role: UploadRole) {
    const selected = Array.from(fileList || []);
    if (!selected.length) return;
    setUploading(true);
    try {
      if (role === "material") {
        const materials = selected.filter((file) => /\.(docx|txt|md|vtt|srt|json|yaml|yml)$/i.test(file.name));
        if (!materials.length) throw new Error("请选择手册、说明或规则材料文件");
        const uploadedMaterials: AgentMaterial[] = [];
        for (const file of materials) uploadedMaterials.push(await api.uploadAgentMaterial(projectId, file));
        setMaterials((current) => [...current, ...uploadedMaterials]);
        toast.success(`已接收 ${materials.length} 份材料，Agent 会在本公司上下文中使用`);
        return;
      }

      const sheets = selected.filter((file) => /\.(xlsx|xls)$/i.test(file.name));
      if (!sheets.length) throw new Error("请选择 .xlsx 或 .xls 文件");
      if (role === "master" && sheets.length > 1) {
        throw new Error("总表一次只能上传 1 份，请只选择本月需要更新的总表");
      }

      if (role === "master") {
        try {
          const uploaded = await api.uploadFile(projectId, sheets[0], "financial_master");
          setFiles((current) => [...current, uploaded]);
          setUploadFailures([]);
        } catch (error) {
          if (!isEncryptedWorkbookError(error)) throw error;
          setPasswordFiles([sheets[0]]);
          setPasswordRole("master");
          setPasswordDraft("");
          setPasswordError(null);
          setPasswordDialogOpen(true);
          toast.info("检测到加密 Excel，请输入打开密码继续");
          return;
        }
        toast.success("总表已接收，可以继续上传更新表或让 Agent 开始分析");
      } else {
        const uploadResult = await api.uploadFilesIndividually(projectId, sheets, "financial_source");
        const encryptedFiles = sheets.filter((file) => uploadResult.failed.some((failure) => failure.filename === file.name && isEncryptedUploadFailure(failure.message)));
        const ordinaryFailures = uploadResult.failed.filter((failure) => !isEncryptedUploadFailure(failure.message));
        if (uploadResult.uploaded.length) setFiles((current) => [...current, ...uploadResult.uploaded]);
        setUploadFailures(ordinaryFailures);
        if (encryptedFiles.length) {
          setPasswordFiles(encryptedFiles);
          setPasswordRole("source");
          setPasswordDraft("");
          setPasswordError(null);
          setPasswordDialogOpen(true);
          toast.info(`有 ${encryptedFiles.length} 份更新表需要打开密码`);
        }
        if (ordinaryFailures.length && !uploadResult.uploaded.length && !encryptedFiles.length) {
          throw new Error(ordinaryFailures.map((failure) => `${failure.filename}：${failure.message}`).join("；"));
        }
        if (uploadResult.uploaded.length) toast.success(`已接收 ${uploadResult.uploaded.length} 份更新表`);
      }
      // 上传后只刷新运行列表与上下文数据，不整体重载：用户可能正在
      // 查看历史运行，整体 load 会把当前视图强行切回最新运行。
      const [runsResult, memoryResult] = await Promise.allSettled([
        api.listAgentRuns(projectId),
        api.listAgentMemory(projectId),
      ]);
      if (runsResult.status === "fulfilled") setRuns(runsResult.value.items);
      if (memoryResult.status === "fulfilled") setMemory(memoryResult.value);
    } catch (error) { toast.error(error instanceof Error ? error.message : "文件上传失败"); }
    finally { setUploading(false); }
  }

  function closePasswordDialog() {
    setPasswordDialogOpen(false);
    setPasswordDraft("");
    setPasswordFiles([]);
    setPasswordRole(null);
    setPasswordError(null);
  }

  function dismissPasswordDialog() {
    if (!passwordSubmitting) closePasswordDialog();
  }

  async function retryEncryptedUpload() {
    const password = passwordDraft.trim();
    if (!password) {
      setPasswordError("请输入 Excel 打开密码");
      return;
    }
    if (!passwordFiles.length || !passwordRole) {
      closePasswordDialog();
      return;
    }
    setPasswordSubmitting(true);
    setPasswordError(null);
    try {
      if (passwordRole === "master") {
        await api.uploadFile(projectId, passwordFiles[0], "financial_master", undefined, password);
        closePasswordDialog();
        toast.success("已验证密码并接收总表");
      } else {
        const result = await api.uploadFilesIndividually(projectId, passwordFiles, "financial_source", password);
        if (result.uploaded.length) setFiles((current) => [...current, ...result.uploaded]);
        if (result.failed.length) {
          const failedNames = new Set(result.failed.map((failure) => failure.filename));
          setPasswordFiles((current) => current.filter((file) => failedNames.has(file.name)));
          throw new Error(result.failed.map((failure) => `${failure.filename}：${failure.message}`).join("；"));
        }
        closePasswordDialog();
        toast.success(`已验证密码并接收 ${result.uploaded.length} 份更新表`);
      }
      await load();
    } catch (error) {
      const message = error instanceof Error ? error.message : "文件解密失败，请重试";
      setPasswordError(/不正确|无法使用该密码|密码/.test(message) ? "打开密码不正确，请重新输入" : message);
    } finally {
      setPasswordSubmitting(false);
    }
  }

  async function downloadFormalResult() {
    if (!run) return;
    setBusy(true);
    try {
      const isDemoRun = run.execution_mode === "demo";
      const isLegacyBeijingProject = project?.name === "北京" && !SHOWCASE_PROJECT_IDS.has(projectId) && !isDemoRun;
      const blob = isLegacyBeijingProject || isDemoRun ? await api.downloadAgentOutput(run.run_id) : await api.downloadAcceptedAgentFinal(run.run_id);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = SHOWCASE_DOWNLOAD_NAMES[projectId] || (isLegacyBeijingProject ? "待确定稿.xlsx" : run.demo_result?.filename || run.draft_filename || "正式稿.xlsx");
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
      toast.success(isLegacyBeijingProject ? "待确定稿已开始下载" : "正式稿已开始下载");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "正式稿下载失败");
    } finally { setBusy(false); }
  }

  async function downloadUpdatedWorkbook() {
    if (!run || !runResult) return;
    setBusy(true);
    try {
      const blob = await api.downloadAgentOutput(run.run_id);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = SHOWCASE_DOWNLOAD_NAMES[projectId] || (runResult.filename || "更新后表格.xlsx");
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
      toast.success("更新后表格已开始下载");
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "更新后表格下载失败");
    } finally { setBusy(false); }
  }

  async function downloadDemo() {
    if (!run?.demo_result) return;
    setBusy(true);
    try {
      const blob = await api.downloadAgentDemo(run.run_id);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = run.demo_result.filename;
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      toast.success("结果文件已下载");
    } catch (error) { toast.error(error instanceof Error ? error.message : "下载失败"); }
    finally { setBusy(false); }
  }

  async function downloadKeyuanReferenceResult() {
    if (!run?.reference_result) return;
    setBusy(true);
    try {
      const blob = await api.downloadKeyuanReferenceResult(run.run_id);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = run.reference_result.filename;
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      toast.success("结果表格已下载");
    } catch (error) { toast.error(error instanceof Error ? error.message : "下载失败"); }
    finally { setBusy(false); }
  }

  async function switchRun(runId: string) {
    // “当前对话”指向最新一次运行；切换后清理上一轮的勾选与聚焦，
    // 避免残留的 selectedReviewIds 指向已不存在的事项。
    const target = runId || runs[0]?.run_id;
    if (!target || target === run?.run_id) return;
    setBusy(true);
    setActiveReviewId(null);
    setSelectedReviewIds(new Set());
    try { const selected = normaliseRun(await api.getAgentRun(target), projectId); if (selected) { setRun(selected); setItems(await api.listAgentItems(target)); } }
    catch (error) { toast.error(error instanceof Error ? error.message : "运行记录加载失败"); }
    finally { setBusy(false); }
  }

  if (loading) return <div className="flex min-h-[70vh] items-center justify-center text-sm text-slate-500" aria-busy="true"><Loader2 className="mr-2 h-4 w-4 animate-spin text-blue-600" />正在加载公司上下文...</div>;

  const runLabel = run ? RUN_LABEL[run.status] || run.status : "准备开始";
  // 对齐阶段：任务已创建但执行尚未启动。此时说明要求会与 Agent 对话
  // 对齐，而不是立即触发固定处理流程。
  const inAlignment = Boolean(run && !run.workflow?.started_at && !run.draft_filename);
  // 下一轮：上一轮已执行完（有工作流启动记录或草稿）且状态收尾。
  // 新消息继续走对齐对话；明确的启动指令（如“开始处理”）开启下一轮，
  // 不再展示“开始下一轮”按钮。
  const roundEnded = Boolean(run && (run.workflow?.started_at || run.draft_filename)
    && ["ready", "completed", "published", "failed"].includes(run.status));
  const showStartButton = Boolean(
    (isShowcaseProject || (masterFiles.length && sourceFiles.length))
      && (!run || inAlignment || (isShowcaseProject && roundEnded)),
  );
  const runDetail = streamStatus || (run?.detail && /科园固定案例|不进入模型确认流程/.test(run.detail)
    ? "文件已识别，正在准备本次处理计划"
    : run?.detail);
  const acceptedRun = Boolean(run?.draft_filename && ["passed", "reference_match", "demo_reference_match"].includes(String(run.validation?.status || "")));
  const assistantIntro: AgentMessage = { role: "agent", content: run ? `我正在处理「${project?.name || "当前公司"}」${project?.salary_month ? ` ${project.salary_month}` : ""}。${pendingCount ? `目前需要你确认 ${pendingCount} 项。` : processingCount ? "我正在继续核对剩余人员。" : run.status === "ready" || run.status === "published" || run.status === "completed" ? "本轮处理已结束，请查看结果与校验信息。想继续调整或开始下一轮，直接说明新要求，我会复述理解；确认后发送“开始处理”。" : inAlignment ? "执行还没有开始。你可以继续说明本次处理要求，我会复述理解、和你对齐细节；确认后发送“开始处理”或点击下方按钮启动执行。" : "请查看下方当前状态；我会自动续跑，只有影响结果的不确定事项才会向你提问。"}` : "你好，我是这家公司的财务 Agent。上传总表、来源文件和手册后，先告诉我这次的处理要求，我会复述理解、和你对齐颗粒度；你说“开始处理”后我才开始分析和修改。" };
  const composer = (
    <div className="w-full max-w-[960px]">
      {files.length || materials.length ? <div className="mb-2"><FileAttachmentList files={files} materials={materials} collapsible onPreview={(file) => void openFilePreview(file)} onRemove={(file) => void removeFile(file)} /></div> : null}
      {uploadFailures.length ? <div className="mb-3 rounded-xl border border-amber-200 bg-amber-50 px-3 py-2.5 text-left" role="alert"><p className="text-xs font-medium text-amber-900">部分文件没有接收</p><ul className="mt-1 space-y-0.5 text-[11px] leading-5 text-amber-800">{uploadFailures.map((failure) => <li key={`${failure.filename}-${failure.message}`} className="truncate" title={failure.message}><span className="font-medium">{failure.filename}</span>：{failure.message}</li>)}</ul></div> : null}
      {showStartButton ? <div className="mb-2 flex items-center justify-between gap-3 rounded-lg border border-blue-200 bg-blue-50 px-3.5 py-2.5 text-left" role="status"><div className="min-w-0"><p className="text-xs font-medium text-blue-950">{run ? "准备开始处理" : "文件已准备好"}</p><p className="mt-0.5 text-[11px] leading-5 text-blue-800">点击后立即执行；未单独选择的待确认事项按 Agent 建议处理，已选择的事项按你的选择执行。</p></div><Button type="button" size="sm" className="h-8 shrink-0 bg-blue-700 px-3 text-xs text-white hover:bg-blue-800" onClick={() => void (run ? executeRun(pendingPlanInstruction || undefined, true) : startRun())} disabled={busy || uploading}><ArrowUp className="mr-1 h-3.5 w-3.5" />开始处理</Button></div> : null}
      {acceptedRun && run ? <div className="mb-2 flex justify-end"><Button type="button" size="sm" className="h-7 bg-blue-700 px-2.5 text-[11px] text-white hover:bg-blue-800" onClick={() => void downloadFormalResult()} disabled={busy}><Download className="mr-1 h-3 w-3" />下载正式稿</Button></div> : null}
      <div className="relative flex items-end gap-2 rounded-[28px] border border-slate-300 bg-white px-3 py-2.5 shadow-[0_6px_24px_rgba(15,23,42,0.07)] transition-shadow focus-within:border-blue-400 focus-within:shadow-[0_8px_28px_rgba(37,99,235,0.13)]">
        <div className="relative shrink-0" ref={uploadMenuRef}>
          <button type="button" className={cn("inline-flex h-9 w-9 items-center justify-center rounded-full text-slate-500 transition-colors hover:bg-slate-100 hover:text-slate-800 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500", uploadMenuOpen && "bg-slate-100 text-slate-900")} onClick={() => setUploadMenuOpen((open) => !open)} disabled={uploading || busy} aria-label="添加文件" aria-expanded={uploadMenuOpen} title="添加文件"><Paperclip className="h-[18px] w-[18px]" /></button>
          {uploadMenuOpen ? (
            <div className="absolute bottom-12 left-0 z-20 w-52 rounded-xl border border-slate-200 bg-white p-1.5 shadow-[0_12px_36px_rgba(15,23,42,0.14)]" role="menu" aria-label="选择上传类型">
              <button type="button" role="menuitem" disabled={Boolean(masterFiles.length)} className="flex w-full items-start gap-2.5 rounded-lg px-2.5 py-2 text-left hover:bg-blue-50 focus-visible:bg-blue-50 focus-visible:outline-none disabled:cursor-not-allowed disabled:opacity-50" onClick={() => openUploadPicker("master")}><Table2 className="mt-0.5 h-4 w-4 shrink-0 text-blue-600" /><span><span className="block text-xs font-medium text-slate-800">上传总表</span><span className="mt-0.5 block text-[10px] text-slate-400">{masterFiles.length ? "已有 1 份总表" : "本次只能有 1 份"}</span></span></button>
              <button type="button" role="menuitem" className="flex w-full items-start gap-2.5 rounded-lg px-2.5 py-2 text-left hover:bg-slate-100 focus-visible:bg-slate-100 focus-visible:outline-none" onClick={() => openUploadPicker("source")}><FileSpreadsheet className="mt-0.5 h-4 w-4 shrink-0 text-slate-500" /><span><span className="block text-xs font-medium text-slate-800">上传更新表</span><span className="mt-0.5 block text-[10px] text-slate-400">可一次选择多份</span></span></button>
              <button type="button" role="menuitem" className="flex w-full items-start gap-2.5 rounded-lg px-2.5 py-2 text-left hover:bg-slate-100 focus-visible:bg-slate-100 focus-visible:outline-none" onClick={() => openUploadPicker("material")}><BookOpen className="mt-0.5 h-4 w-4 shrink-0 text-slate-500" /><span><span className="block text-xs font-medium text-slate-800">上传手册或说明</span><span className="mt-0.5 block text-[10px] text-slate-400">作为本公司上下文材料</span></span></button>
            </div>
          ) : null}
        </div>
        <input ref={masterInputRef} type="file" accept=".xlsx,.xls" className="sr-only" disabled={uploading || busy} onChange={(event) => { void uploadFiles(event.target.files, "master"); event.currentTarget.value = ""; }} />
        <input ref={sourceInputRef} type="file" multiple accept=".xlsx,.xls" className="sr-only" disabled={uploading || busy} onChange={(event) => { void uploadFiles(event.target.files, "source"); event.currentTarget.value = ""; }} />
        <input ref={materialInputRef} type="file" multiple accept=".docx,.txt,.md,.vtt,.srt,.json,.yaml,.yml" className="sr-only" disabled={uploading || busy} onChange={(event) => { void uploadFiles(event.target.files, "material"); event.currentTarget.value = ""; }} />
        <textarea value={draft} onChange={(event) => setDraft(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter" && !event.shiftKey) { event.preventDefault(); void sendMessage(); } }} rows={2} disabled={busy} className="max-h-36 min-h-12 flex-1 resize-none border-0 bg-transparent px-1 py-2 text-[15px] leading-6 text-slate-800 outline-none placeholder:text-slate-400 disabled:cursor-not-allowed" placeholder={busy ? "正在处理中，完成后即可继续输入…" : !run ? "先说明本次处理要求（我会复述理解并对齐），或直接发送“开始处理”" : inAlignment ? "继续说明或修正处理要求；对齐完成后发送“开始处理”" : roundEnded ? "说明下一轮的处理要求，或直接发送可执行的表格指令" : activeItem ? "直接回答 Agent，或补充本次处理要求" : "可以继续对话，或用自然语言让 Agent 继续修改表格"} aria-label="与财务 Agent 对话" aria-busy={busy} />
        <div className="flex shrink-0 items-center gap-1">
          {files.length ? <span className="hidden whitespace-nowrap px-2 text-[11px] text-slate-400 sm:inline">{files.length} 个文件</span> : null}
          <button type="button" className={cn("inline-flex h-10 w-10 items-center justify-center rounded-full text-white transition-colors disabled:cursor-not-allowed", messageAbortControllerRef.current ? "bg-slate-700 hover:bg-slate-800" : "bg-blue-700 hover:bg-blue-800 disabled:bg-slate-200 disabled:text-slate-400")} onClick={() => messageAbortControllerRef.current ? stopMessageGeneration() : void sendMessage()} disabled={messageAbortControllerRef.current ? false : !draft.trim() || busy} aria-label={messageAbortControllerRef.current ? "停止生成" : "发送消息"} title={messageAbortControllerRef.current ? "停止生成" : "发送消息"}>
            {messageAbortControllerRef.current ? <Square className="h-3.5 w-3.5 fill-current" /> : busy ? <Loader2 className="h-4 w-4 animate-spin" /> : <ArrowUp className="h-4 w-4" />}
          </button>
        </div>
      </div>
      <p className="mt-2 text-center text-[11px] text-slate-400">总表 1 份 · 更新表可多份 · Enter 发送 · Shift + Enter 换行</p>
    </div>
  );

  return (
    <div className="agent-page flex h-[100dvh] min-h-[560px] flex-col overflow-hidden bg-[#fbfcfe]">
      <header className="flex h-16 shrink-0 items-center justify-between gap-3 border-b border-slate-200/90 bg-white px-4 sm:px-6">
        <div className="flex min-w-0 items-center gap-2.5">
          <Link href="/projects" className="flex shrink-0 items-center gap-2 text-sm font-semibold tracking-tight text-slate-950 hover:text-blue-700"><AgentMascot className="h-6 w-5" />财务 Agent</Link>
          <span className="text-slate-300">/</span>
          <span className="max-w-[34vw] truncate text-sm text-slate-600">{project?.name || "公司"}</span>
          <span className="shrink-0 rounded-full bg-slate-100 px-2.5 py-1 text-[11px] text-slate-600">{project?.salary_month || "未设置月份"}</span>
        </div>
        <div className="flex shrink-0 items-center gap-1.5">
          <span className={cn("inline-flex items-center gap-1.5 px-2 text-[11px]", (processRunning || run?.status === "processing") ? "text-blue-700" : "text-slate-400")} role={processRunning || run?.status === "processing" ? "status" : undefined} aria-live={processRunning || run?.status === "processing" ? "polite" : undefined}>
            {processRunning || run?.status === "processing" ? <Loader2 className="h-3 w-3 animate-spin text-blue-600" /> : <span className={cn("h-1.5 w-1.5 rounded-full", run?.status === "blocked" ? "bg-red-500" : run?.status === "published" ? "bg-emerald-500" : "bg-slate-300")} />}
            {runLabel}
          </span>
          {(processRunning || run?.status === "processing") && run?.run_id ? (
            <button type="button" className="inline-flex h-8 items-center gap-1.5 rounded-md border border-slate-300 bg-white px-2.5 text-xs font-medium text-slate-600 transition-colors hover:border-red-300 hover:bg-red-50 hover:text-red-700 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-red-400/40 disabled:cursor-not-allowed disabled:opacity-60" onClick={() => void stopRun()} disabled={stoppingRun} aria-label="停止本次处理" title="停止本次处理（已完成的写入会保留）">
              {stoppingRun ? <Loader2 className="h-3.5 w-3.5 animate-spin" /> : <Square className="h-3 w-3 fill-current" />}
              <span className="hidden sm:inline">{stoppingRun ? "正在停止" : "停止"}</span>
            </button>
          ) : null}
          <label className="relative hidden sm:block">
            <span className="sr-only">切换历史运行</span>
            <select value={run?.run_id || ""} onChange={(event) => void switchRun(event.target.value)} className="h-8 max-w-36 appearance-none border-0 bg-transparent px-2 pr-6 text-xs text-slate-500 outline-none focus:ring-2 focus:ring-blue-500/20"><option value="">当前对话</option>{runs.map((history) => <option key={history.run_id} value={history.run_id}>{history.salary_month || "本月"} · {RUN_LABEL[history.status] || history.status}</option>)}</select>
            <ChevronDown className="pointer-events-none absolute right-1 top-2 h-3.5 w-3.5 text-slate-400" />
          </label>
          <button type="button" className="inline-flex h-9 w-9 items-center justify-center rounded-full text-slate-500 transition-colors hover:bg-slate-100 hover:text-slate-900" onClick={() => setShowContext(true)} aria-label="查看公司上下文" title="公司上下文"><PanelRight className="h-[18px] w-[18px]" /></button>
          <button type="button" className="inline-flex h-9 w-9 items-center justify-center rounded-full text-slate-500 transition-colors hover:bg-slate-100 hover:text-slate-900 sm:hidden" onClick={() => setShowContext(true)} aria-label="查看公司信息" title="公司信息"><Info className="h-[18px] w-[18px]" /></button>
        </div>
      </header>

      <main className="relative min-h-0 flex-1 overflow-hidden">
        <div ref={scrollContainerRef} className="chat-scrollbar absolute inset-0 overflow-y-auto" onScroll={handleChatScroll}>
          {!hasConversation ? (
            <div className="flex min-h-full flex-col items-center justify-center px-5 pb-24 pt-12 text-center">
              <AgentMascot className="mb-5 h-12 w-9" priority />
              <h1 className="text-[28px] font-semibold tracking-tight text-slate-950 sm:text-[32px]">准备好，随时开始</h1>
              <p className="mt-3 max-w-md text-sm leading-6 text-slate-500">把原始总表、来源文件和操作手册交给我。我会先说明计划，确认后执行，只在影响结果的不确定处向你提问。</p>
              {processRunning ? <div className="mt-6 flex items-center gap-2 text-xs text-slate-500" role="status" aria-live="polite"><Loader2 className="h-3.5 w-3.5 animate-spin text-blue-600" />正在创建处理批次，准备读取文件...</div> : null}
              <div className="mt-7 flex w-full justify-center">{composer}</div>
            </div>
          ) : (
            <div className="mx-auto w-full max-w-4xl px-5 pt-10 sm:px-8" style={{ paddingBottom: `${Math.max(composerHeight + 32, 220)}px` }}>
              <MessageBubble message={assistantIntro} />
              {run ? <AgentPlanConfirmation run={run} busy={busy} onCustomize={(instruction) => replan(instruction)} onAnswersChanged={setPendingPlanInstruction} /> : null}
              {runResult ? <div ref={resultAnchorRef}><AgentResultView result={runResult} busy={busy} onDownload={() => void downloadUpdatedWorkbook()} /></div> : null}
              {run?.reference_result ? <section className="my-6 border border-emerald-200 bg-white p-5" aria-label="处理结果"><p className="text-sm font-semibold text-slate-900">本轮处理已完成</p><p className="mt-2 break-all text-sm text-slate-600">{run.reference_result.filename}</p><p className="mt-1 text-xs leading-5 text-slate-500">结果工作簿已校验完成，原始上传文件未被修改。</p><Button type="button" className="mt-4 bg-blue-700 text-white hover:bg-blue-800" onClick={() => void downloadKeyuanReferenceResult()} disabled={busy}><FileSpreadsheet className="mr-1.5 h-4 w-4" />下载结果表格</Button></section> : null}
              <BatchReviewPanel items={items} selectedIds={selectedReviewIds} busy={busy} onToggleItem={toggleReviewItem} onToggleGroup={toggleReviewGroup} onApply={(selectedItems, action, customValue, note) => void applyBatchDecision(selectedItems, action, customValue, note)} />
              {runMessages.map((message, index) => <MessageBubble key={`run-${index}`} message={message} busy={busy} isLatestAgentMessage={message.role === "agent" && index === runMessages.length - 1} onQuickReply={(reply) => void sendMessage(reply)} />)}
              {itemMessages.map(({ message, item }, index) => {
                const latestAgentMessage = [...(item.conversation || [])].reverse().find((entry) => entry.role === "agent");
                return <div key={`${item.item_id}-${index}`} ref={item.item_id === activeItem?.item_id && message === item.conversation?.[0] ? activeReviewRef : undefined}>
                  <MessageBubble
                    message={message}
                    evidence={message.role === "agent" ? item : null}
                    showChoices={message.role === "agent" && item.status === "needs_conversation" && message === latestAgentMessage}
                    busy={busy}
                    onChoice={(action, value) => void chooseDecision(item, action, value)}
                  />
                </div>;
              })}
              {runDetail ? <div role="status" className="mt-6 flex items-start gap-2 border-l-2 border-slate-300 bg-slate-50 px-3 py-2.5 text-sm leading-6 text-slate-700"><Activity className="mt-1 h-4 w-4 shrink-0 text-blue-600" />{runDetail}</div> : null}
              <AgentActivityMessage events={activityEvents} running={processRunning || run?.status === "processing"} stopping={stoppingRun} stopped={run?.status === "execution_incomplete" && run?.code === "USER_STOPPED"} onStop={() => void stopRun()} />
              <div ref={bottomRef} />
            </div>
          )}
        </div>
        {hasConversation ? <div ref={composerDockRef} className="pointer-events-none absolute inset-x-0 bottom-0 bg-gradient-to-t from-white via-white/95 to-transparent px-5 pb-4 pt-14 sm:px-8"><div className="pointer-events-auto mx-auto w-full max-w-5xl">{composer}</div></div> : null}
      </main>

      {resumePrompt ? <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/25 p-4" role="dialog" aria-modal="true" aria-label="Agent 等待继续指令" onClick={dismissResumePrompt} onKeyDown={(event) => { if (event.key === "Escape") dismissResumePrompt(); }}><div className="w-full max-w-md rounded-xl border border-amber-200 bg-white p-6 shadow-2xl" onClick={(event) => event.stopPropagation()}><div className="flex items-start gap-3"><div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-amber-50 text-amber-600"><Activity className="h-5 w-5" /></div><div className="min-w-0 flex-1"><h2 className="text-base font-semibold text-slate-950">Agent 已暂停，等待你的指令</h2><p className="mt-1 text-sm leading-6 text-slate-500">{resumePrompt.detail || "本轮处理已中断，已完成的写入和进度均已保留。"}</p><p className="mt-1 text-xs leading-5 text-slate-400">点击“立即继续”会自动发送继续指令，从上次检查点接着处理。</p></div><button type="button" className="ml-auto inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-slate-400 hover:bg-slate-100 hover:text-slate-800" onClick={dismissResumePrompt} aria-label="关闭提醒"><X className="h-4 w-4" /></button></div><div className="mt-5 flex justify-end gap-2 border-t border-slate-100 pt-4"><Button type="button" variant="outline" onClick={dismissResumePrompt} disabled={busy}>稍后处理</Button><Button type="button" className="bg-blue-700 text-white hover:bg-blue-800" onClick={() => void continueFromResumePrompt()} disabled={busy}>{busy ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <ArrowUp className="mr-2 h-4 w-4" />}立即继续</Button></div>{typeof window !== "undefined" && "Notification" in window && Notification.permission === "default" ? <button type="button" className="mt-3 text-xs text-slate-400 underline underline-offset-2 transition-colors hover:text-slate-600" onClick={enableSystemNotify}>切走页面时用系统通知提醒我</button> : null}</div></div> : null}
      {passwordDialogOpen ? <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-950/25 p-4" role="dialog" aria-modal="true" aria-labelledby="encrypted-workbook-title" aria-describedby="encrypted-workbook-description" onClick={dismissPasswordDialog} onKeyDown={(event) => { if (event.key === "Escape") dismissPasswordDialog(); }}><div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-6 shadow-2xl" onClick={(event) => event.stopPropagation()}><div className="flex items-start gap-3"><div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-slate-100 text-slate-700"><KeyRound className="h-5 w-5" /></div><div className="min-w-0"><h2 id="encrypted-workbook-title" className="text-base font-semibold text-slate-950">Excel 需要打开密码</h2><p id="encrypted-workbook-description" className="mt-1 text-sm leading-6 text-slate-500">输入后会立即重试本次文件上传。密码只在当前请求中使用，不会写入文件。</p></div><button type="button" className="ml-auto inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-full text-slate-400 hover:bg-slate-100 hover:text-slate-800" onClick={dismissPasswordDialog} disabled={passwordSubmitting} aria-label="关闭密码输入"><X className="h-4 w-4" /></button></div><div className="mt-5 rounded-lg bg-slate-50 px-3.5 py-3"><p className="text-xs font-medium text-slate-600">待验证文件</p><ul className="mt-2 space-y-1">{passwordFiles.map((file) => <li key={`${file.name}-${file.size}-${file.lastModified}`} className="truncate text-xs text-slate-500" title={file.name}>{file.name}</li>)}</ul></div><form className="mt-5 space-y-4" onSubmit={(event) => { event.preventDefault(); void retryEncryptedUpload(); }}><div><label htmlFor="encrypted-workbook-password" className="mb-1.5 block text-sm font-medium text-slate-700">打开密码</label><input id="encrypted-workbook-password" type="password" value={passwordDraft} onChange={(event) => { setPasswordDraft(event.target.value); if (passwordError) setPasswordError(null); }} autoFocus autoComplete="off" disabled={passwordSubmitting} className="h-11 w-full rounded-lg border border-slate-300 bg-white px-3 text-sm text-slate-900 outline-none transition-shadow placeholder:text-slate-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 disabled:bg-slate-50" placeholder="请输入 Excel 打开密码" aria-invalid={Boolean(passwordError)} aria-describedby={passwordError ? "encrypted-workbook-error" : undefined} />{passwordError ? <p id="encrypted-workbook-error" className="mt-2 flex items-center gap-1.5 text-xs text-red-600" role="alert"><AlertCircle className="h-3.5 w-3.5 shrink-0" />{passwordError}</p> : null}</div><div className="flex justify-end gap-2 border-t border-slate-100 pt-4"><Button type="button" variant="outline" onClick={dismissPasswordDialog} disabled={passwordSubmitting}>取消</Button><Button type="submit" className="bg-blue-700 text-white hover:bg-blue-800" disabled={passwordSubmitting || !passwordDraft.trim()}>{passwordSubmitting ? <Loader2 className="mr-2 h-4 w-4 animate-spin" /> : <KeyRound className="mr-2 h-4 w-4" />}验证并继续</Button></div></form></div></div> : null}
      {showContext ? <div className="fixed inset-0 z-50 flex justify-end bg-slate-950/20" role="dialog" aria-modal="true" aria-label="公司上下文" onClick={() => setShowContext(false)}><div className="h-full w-full max-w-sm border-l border-slate-200 bg-white shadow-xl" onClick={(event) => event.stopPropagation()}><div className="flex h-14 items-center justify-between border-b border-slate-200 px-5"><div className="flex items-center gap-2 text-sm font-semibold text-slate-900"><ShieldCheck className="h-4 w-4 text-blue-600" />公司上下文</div><button type="button" className="inline-flex h-8 w-8 items-center justify-center rounded-full text-slate-400 hover:bg-slate-100 hover:text-slate-800" onClick={() => setShowContext(false)} aria-label="关闭上下文"><X className="h-4 w-4" /></button></div><ContextPanel project={project} run={run} files={files} memory={memory} modelStatus={modelStatus} items={items} onPreview={(file) => void openFilePreview(file)} onRemove={(file) => void removeFile(file)} /></div></div> : null}
      <FilePreviewDrawer file={previewFile} preview={preview} loading={previewLoading} error={previewError} onClose={closeFilePreview} />
    </div>
  );
}
