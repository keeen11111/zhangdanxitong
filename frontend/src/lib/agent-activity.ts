import type { AgentRunEvent } from "./api";

const visibleEventTypes = new Set([
  "run_started", "progress", "model_request", "model_response", "tool_call",
  "tool_result", "work_item_updated", "validation", "needs_user_input",
  "run_blocked", "run_completed", "run_failed",
]);
const resumableExecutionCodes = new Set([
  "MAX_TURNS_EXCEEDED",
  "EMPTY_MODEL_RESPONSE",
  "MODEL_PROVIDER_ERROR",
]);

export function isVisibleAgentActivity(event: AgentRunEvent) {
  return visibleEventTypes.has(event.type);
}

export function agentActivityEventLabel(event: AgentRunEvent) {
  const payload = event.payload || {};
  const label = typeof payload.label === "string" ? payload.label : "";
  if (label) return label;
  if (event.type === "run_started") return "已创建处理批次";
  if (event.type === "model_request") return "正在请求 Agent 分析";
  if (event.type === "model_response") return "已收到 Agent 分析结果";
  if (event.type === "tool_call") {
    const name = String(payload.name || "受控工具");
    const names: Record<string, string> = {
      inspect_workbook: "正在读取工作簿结构",
      classify_file: "正在确认文件角色",
      find_table: "正在定位数据表",
      read_range: "正在读取相关单元格",
      match_person: "正在匹配人员",
      propose_changes: "正在生成更新建议",
      prepare_workbook_copy: "正在创建可回退的更新副本",
      apply_source_cells: "正在按来源表写入更新",
      apply_cell_changes: "正在写入草稿",
      apply_formula_divisors: "正在更新公式参数",
      insert_and_copy_row: "正在新增并复制表格行",
      validate_workbook: "正在校验工作簿",
    };
    return names[name] || "正在执行受控检查";
  }
  if (event.type === "tool_result") {
    if (payload.status !== "succeeded") return "受控操作未通过校验";
    const name = String(payload.name || "");
    const completed: Record<string, string> = {
      inspect_workbook: "已读取工作簿结构，正在确定下一步",
      inspect_source_file: "已读取来源文件，正在核对目标字段",
      find_table: "已定位数据表，正在核对字段",
      read_range: "已读取相关单元格，正在分析下一步",
      read_source_range: "已读取来源数据，正在核对写入依据",
      match_person: "已完成人员匹配，正在核对变更",
      prepare_workbook_copy: "已创建独立副本，正在写入已核对变更",
      apply_source_cells: "已写入来源变更，正在继续校验",
      apply_formula_divisors: "已更新公式参数，正在继续校验",
      insert_and_copy_row: "已插入并复制表格行，正在继续校验",
      validate_workbook: "已完成工作簿校验，正在整理结果",
    };
    return completed[name] || "已完成受控操作，正在继续处理";
  }
  if (event.type === "work_item_updated") return "人员状态已更新";
  if (event.type === "validation") return "校验已完成";
  if (event.type === "needs_user_input") {
    if (payload.code === "MAX_TURNS_EXCEEDED") return "执行未完成：模型读取达到本轮上限";
    if (payload.code === "EMPTY_MODEL_RESPONSE") return "模型空响应已自动重试，当前进度已保留";
    if (payload.code === "MODEL_PROVIDER_ERROR") return "模型服务请求失败，当前进度已保留";
    return "等待你的确认";
  }
  if (event.type === "run_blocked") return "处理被阻断，需要补充信息";
  if (event.type === "run_failed") return typeof payload.detail === "string" ? payload.detail : "处理失败";
  if (event.type === "run_completed") return "处理批次已完成";
  return "Agent 正在处理";
}

export function agentActivitySummary(events: AgentRunEvent[], isRunning: boolean) {
  const visibleEvents = events.filter(isVisibleAgentActivity);
  const latestProgress = [...visibleEvents].reverse().find((event) => event.type === "progress");
  // Person-level progress is high volume during batch runs. Keep the latest
  // checkpoint as one timeline entry while preserving tools and exceptions.
  const summarizedEvents = latestProgress
    ? [...visibleEvents.filter((event) => event.type !== "progress"), latestProgress].sort((a, b) => a.revision - b.revision)
    : visibleEvents;
  const latest = summarizedEvents.at(-1);
  const progress = latestProgress?.payload || {};
  const current = typeof progress.current === "number" ? progress.current : null;
  const total = typeof progress.total === "number" ? progress.total : null;

  return {
    current: latest ? agentActivityEventLabel(latest) : isRunning ? "正在准备处理" : "暂无执行记录",
    events: summarizedEvents,
    recordCount: summarizedEvents.length,
    isIncomplete: latest?.type === "needs_user_input" && resumableExecutionCodes.has(String(latest.payload?.code || "")),
    isRunning,
    progress: current !== null && total && total > 0 ? { current, total, percent: Math.min(100, Math.round((current / total) * 100)) } : null,
  };
}
