import type { AgentDecisionAction, AgentRun, AgentWorkItem } from "./api";

export type AgentTimelineKind = "welcome" | "files" | "plan" | "progress" | "review" | "complete" | "blocked";

export type AgentTimelineEvent = {
  id: string;
  kind: AgentTimelineKind;
  title: string;
  body: string;
  itemId?: string;
  tone?: "neutral" | "success" | "warning" | "danger";
};

export type AgentShortcutAction = AgentDecisionAction | "next";

/** Resolve the documented review shortcuts without coupling them to the UI. */
export function agentShortcutAction(key: string, altKey: boolean): AgentShortcutAction | null {
  if (!altKey) return null;
  const normalized = key.toLowerCase();
  if (normalized === "a") return "apply";
  if (normalized === "k") return "keep_current";
  if (normalized === "s") return "skip";
  if (normalized === "n") return "next";
  return null;
}

export function isPendingAgentItem(item: AgentWorkItem) {
  return item.status === "pending" || item.status === "needs_conversation";
}

export function isResumableAgentRun(run: AgentRun | null) {
  return run?.status === "execution_incomplete";
}

export function canResumeDirectWrite(run: AgentRun | null, items: AgentWorkItem[]) {
  return Boolean(
    run?.status === "review"
    && run.plan_confirmation?.confirmed
    && run.execution_result
    && !items.some(isPendingAgentItem),
  );
}

export function nextPendingAgentItem(items: AgentWorkItem[], currentId?: string | null) {
  const pending = items.filter(isPendingAgentItem);
  if (!pending.length) return null;
  const currentIndex = currentId ? pending.findIndex((item) => item.item_id === currentId) : -1;
  return pending[(currentIndex + 1 + pending.length) % pending.length] || pending[0];
}

export function agentProgress(items: AgentWorkItem[]) {
  if (!items.length) return 0;
  const resolved = items.filter((item) => !isPendingAgentItem(item)).length;
  return Math.round((resolved / items.length) * 100);
}

export function agentRunLabel(run: AgentRun | null) {
  if (!run) return "尚未开始";
  if (run.status === "published") return "已发布";
  if (run.status === "blocked") return "需确认后继续";
  if (run.status === "review") return "等待逐人确认";
  if (run.status === "ready" || run.status === "completed") return "可以发布";
  if (run.status === "processing" || run.status === "planning") return "Agent 正在处理";
  return "待开始";
}

export function timelineForRun(run: AgentRun | null, items: AgentWorkItem[], fileCount: number): AgentTimelineEvent[] {
  if (!run) {
    return [{
      id: "welcome",
      kind: "welcome",
      title: "先把本月文件交给我",
      body: "上传总表、变更文件、手册或人工整理的文字材料，再告诉我本次要求。我会先识别文件和月份，确认后才会写入总表。录音仅存档，不自动转写。",
    }];
  }
  const events: AgentTimelineEvent[] = [
    {
      id: "files",
      kind: "files",
      title: "文件已完成识别",
      body: fileCount ? `已纳入 ${fileCount} 份文件，正在按公司规则建立人员处理项。` : "已建立处理批次，等待文件识别结果。",
      tone: "neutral",
    },
    {
      id: "plan",
      kind: "plan",
      title: run.detail || "处理计划已生成",
      body: run.rule_version ? `本次使用规则版本 ${run.rule_version}。优先处理低风险变更，冲突项逐人停下确认。` : "Agent 会先检查月份、总表定位和人员匹配，再生成可审计的写入计划。",
      tone: run.status === "blocked" ? "warning" : "neutral",
    },
  ];
  const resolved = items.filter((item) => !isPendingAgentItem(item)).length;
  if (items.length) {
    events.push({
      id: "progress",
      kind: "progress",
      title: resolved ? `已处理 ${resolved}/${items.length} 人` : "人员处理项已建立",
      body: run.status === "blocked" ? (run.detail || "请先确认阻断原因") : "Agent 会先自主核对并处理明确事项，只在证据不足或存在冲突时请求确认。",
      tone: run.status === "blocked" ? "warning" : resolved === items.length ? "success" : "neutral",
    });
  }
  items.filter(isPendingAgentItem).forEach((item) => {
    events.push({
      id: `review-${item.item_id}`,
      kind: "review",
      title: `${item.person_name || item.employee_ref || "一名人员"} 需要确认`,
      body: item.reason || item.recommendation || "这个人员存在无法安全自动判断的变更。",
      itemId: item.item_id,
      tone: item.risk_level === "high" ? "danger" : "warning",
    });
  });
  if (run.status === "published") {
    events.push({
      id: "complete",
      kind: "complete",
      title: "正式结果已发布",
      body: "本次处理的决策、规则版本和文件结果已经归档，可从历史记录下载审计文件。",
      tone: "success",
    });
  }
  return events;
}
