import assert from "node:assert/strict";
import test from "node:test";

import { agentActivitySummary } from "./agent-activity.ts";
import type { AgentRunEvent } from "./api";

const event = (type: string, payload: Record<string, unknown> = {}): AgentRunEvent => ({
  event_id: type,
  run_id: "run",
  revision: 1,
  type,
  payload,
});

test("shows the latest tool activity as the live assistant status", () => {
  const summary = agentActivitySummary([
    event("run_started"),
    event("tool_call", { name: "match_person" }),
  ], true);

  assert.equal(summary.current, "正在匹配人员");
  assert.equal(summary.recordCount, 2);
  assert.equal(summary.isRunning, true);
});

test("keeps a visible processing status before the first event arrives", () => {
  const summary = agentActivitySummary([], true);

  assert.equal(summary.current, "正在准备处理");
  assert.equal(summary.isRunning, true);
});

test("describes workbook write tools instead of showing a generic check", () => {
  const summary = agentActivitySummary([
    event("tool_call", { name: "apply_source_cells" }),
  ], true);

  assert.equal(summary.current, "正在按来源表写入更新");
});

test("keeps model requests visibly distinct from completed work", () => {
  const summary = agentActivitySummary([
    event("model_request", { turn: 2, attempt: 1 }),
  ], true);

  assert.equal(summary.current, "正在请求 Agent 分析");
  assert.equal(summary.isRunning, true);
});

test("keeps a completed tool result specific while the agent is still working", () => {
  const summary = agentActivitySummary([
    event("tool_call", { name: "read_range" }),
    event("tool_result", { name: "read_range", status: "succeeded" }),
  ], true);

  assert.equal(summary.current, "已读取相关单元格，正在分析下一步");
});

test("collapses repeated progress events into one summary while keeping failures", () => {
  const progressEvents = Array.from({ length: 100 }, (_, index) => event("progress", {
    stage: "person", current: index + 1, total: 100, label: `已处理第 ${index + 1} 人`,
  }));
  const summary = agentActivitySummary([
    event("run_started"),
    ...progressEvents,
    event("work_item_updated", { person_name: "需要确认的人" }),
  ], false);

  assert.equal(summary.events.filter((entry) => entry.type === "progress").length, 1);
  assert.equal(summary.events.some((entry) => entry.type === "work_item_updated"), true);
  assert.equal(summary.progress?.current, 100);
  assert.equal(summary.recordCount, 3);
});

test("turns a completed stream into a concise finished status", () => {
  const summary = agentActivitySummary([
    event("tool_call", { name: "validate_workbook" }),
    event("run_completed"),
  ], false);

  assert.equal(summary.current, "处理批次已完成");
  assert.equal(summary.isRunning, false);
});

test("does not describe a turn limit as a completed user confirmation", () => {
  const summary = agentActivitySummary([
    event("tool_result", { name: "read_range", status: "succeeded" }),
    event("needs_user_input", { code: "MAX_TURNS_EXCEEDED" }),
  ], false);

  assert.equal(summary.current, "执行未完成：模型读取达到本轮上限");
  assert.equal(summary.recordCount, 2);
  assert.equal(summary.isIncomplete, true);
});

test("describes an empty model response as resumable work", () => {
  const summary = agentActivitySummary([
    event("model_response", { content: "", tool_call_count: 0 }),
    event("needs_user_input", { code: "EMPTY_MODEL_RESPONSE" }),
  ], false);

  assert.equal(summary.current, "模型空响应已自动重试，当前进度已保留");
  assert.equal(summary.isIncomplete, true);
});

test("describes a provider interruption as resumable work", () => {
  const summary = agentActivitySummary([
    event("needs_user_input", { code: "MODEL_PROVIDER_ERROR" }),
  ], false);

  assert.equal(summary.current, "模型服务请求失败，当前进度已保留");
  assert.equal(summary.isIncomplete, true);
});
