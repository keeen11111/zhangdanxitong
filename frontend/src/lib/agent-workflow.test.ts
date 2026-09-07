import assert from "node:assert/strict";
import test from "node:test";

import { agentProgress, agentShortcutAction, canResumeDirectWrite, defaultTargetSalaryMonth, isResumableAgentRun, nextPendingAgentItem, timelineForRun, visibleAgentItems } from "./agent-workflow.ts";
import type { AgentRun, AgentWorkItem } from "./api";

const items: AgentWorkItem[] = [
  { item_id: "a", status: "applied", person_name: "甲" },
  { item_id: "b", status: "needs_conversation", person_name: "乙" },
  { item_id: "c", status: "pending", person_name: "丙" },
];

test("progress counts only resolved people", () => {
  assert.equal(agentProgress(items), 33);
});

test("next pending item wraps through the queue", () => {
  assert.equal(nextPendingAgentItem(items, "b")?.item_id, "c");
  assert.equal(nextPendingAgentItem(items, "c")?.item_id, "b");
  assert.equal(nextPendingAgentItem([{ item_id: "a", status: "applied" }]), null);
});

test("maps review shortcuts only when Alt is held", () => {
  assert.equal(agentShortcutAction("a", true), "apply");
  assert.equal(agentShortcutAction("K", true), "keep_current");
  assert.equal(agentShortcutAction("s", true), "skip");
  assert.equal(agentShortcutAction("n", true), "next");
  assert.equal(agentShortcutAction("a", false), null);
  assert.equal(agentShortcutAction("x", true), null);
});

test("timeline turns unresolved people into direct review prompts", () => {
  const events = timelineForRun({ run_id: "run", project_id: "p", status: "review" }, items, 3);
  assert.equal(events.filter((event) => event.kind === "review").length, 2);
  assert.equal(events.find((event) => event.itemId === "b")?.tone, "warning");
});

test("an incomplete execution remains resumable without review items", () => {
  assert.equal(isResumableAgentRun({ run_id: "run", project_id: "p", status: "execution_incomplete" }), true);
  assert.equal(isResumableAgentRun({ run_id: "run", project_id: "p", status: "review" }), false);
});

test("a new run defaults to the month after the latest accepted run", () => {
  const runs: AgentRun[] = [
    {
      run_id: "may", project_id: "p", status: "completed", salary_month: "2026.05",
      validation: { status: "structurally_valid" },
    },
    {
      run_id: "april", project_id: "p", status: "completed", salary_month: "2026.04",
      validation: { status: "passed" },
    },
  ];

  assert.equal(defaultTargetSalaryMonth("2026.04", runs), "2026-06");
  assert.equal(defaultTargetSalaryMonth("2026.12", []), "2026-12");
});

test("the newest accepted salary month wins even when run history is reordered", () => {
  const runs: AgentRun[] = [
    {
      run_id: "april-reopened", project_id: "p", status: "published", salary_month: "2026.04",
      validation: { status: "passed" }, updated_at: "2026-07-01T00:00:00Z",
    },
    {
      run_id: "may", project_id: "p", status: "completed", salary_month: "2026.05",
      validation: { status: "structurally_valid" }, updated_at: "2026-06-01T00:00:00Z",
    },
  ];

  assert.equal(defaultTargetSalaryMonth("2026.04", runs), "2026-06");
});

test("keeps unresolved items visible when a different item is resolved", () => {
  const reviewItems = [
    { item_id: "resolved", status: "applied" as const },
    { item_id: "unselected", status: "needs_conversation" as const },
    { item_id: "pending", status: "pending" as const },
  ];

  assert.deepEqual(visibleAgentItems(reviewItems).map((item) => item.item_id), ["resolved", "unselected", "pending"]);
});

test("a confirmed review run without pending people can resume direct writing", () => {
  const run = {
    run_id: "run", project_id: "p", status: "review" as const,
    plan_confirmation: { required: true, confirmed: true },
    execution_result: { status: "completed" },
  };

  assert.equal(canResumeDirectWrite(run, []), true);
  assert.equal(canResumeDirectWrite(run, [{ item_id: "a", status: "needs_conversation" }]), false);
  assert.equal(canResumeDirectWrite({ ...run, plan_confirmation: { required: true, confirmed: false } }, []), false);
});

test("the known workbook case still uses the normal plan timeline", () => {
  const events = timelineForRun({ run_id: "run", project_id: "p", status: "planning", execution_mode: "keyuan_reference" }, [], 8);
  assert.equal(events.some((event) => event.kind === "plan"), true);
});
