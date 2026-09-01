import assert from "node:assert/strict";
import test from "node:test";

import { buildPlanResponse, presentPlanQuestions, planConfirmationAction } from "./agent-plan-confirmation-state.ts";

test("automatically processes a plan when it has no unresolved questions", () => {
  assert.equal(planConfirmationAction({ hasQuestions: false }), "auto_process");
});

test("turns recurring planning questions into explicit choices instead of a free-text blocker", () => {
  const questions = presentPlanQuestions([
    "当前master文件名为202607且样例显示薪资月为202606，而本次指令指定salary_month为2026.08。请确认是否按项目月份处理？",
    "手册要求【其他调整累计】每月清空重填，但未提供202608对应的补发补扣来源文件。请确认本月是否留空？",
    "【工资核算】AD/AE列奖金数据来源于奖金-7月Sheet，而本次处理薪资月为2026.08。请确认是否仍使用7月奖金表？",
  ]);

  assert.deepEqual(questions.map((question) => question.title), ["处理月份", "补发补扣", "奖金来源"]);
  assert.ok(questions.every((question) => question.options.length === 2));
});

test("groups repeated issues into one category choice", () => {
  const questions = presentPlanQuestions([
    "考勤!G2 是公式，基础处理器跳过写入",
    "考勤!I2 是公式，基础处理器跳过写入",
    "考勤!J2 是公式，基础处理器跳过写入",
  ]);
  assert.equal(questions.length, 1);
  assert.equal(questions[0].title, "考勤数据");
  assert.match(questions[0].question, /共 3 项/);
  assert.equal(questions[0].options.length, 2);
});

test("builds one auditable response from the selected answers", () => {
  const questions = presentPlanQuestions(["当前master文件名为202607，而本次指令指定salary_month为2026.08。请确认处理月份。"], "2026.08");
  const instruction = buildPlanResponse(questions, { [questions[0].id]: "use_project_month" }, "仅处理来源中有记录的人员");

  assert.match(instruction, /按项目月份 2026\.08 处理/);
  assert.match(instruction, /仅处理来源中有记录的人员/);
});

test("confirms the plan once every genuine question has an answer", () => {
  assert.equal(planConfirmationAction({ hasQuestions: true, hasUnansweredQuestions: true }), "answer_questions");
  assert.equal(planConfirmationAction({ hasQuestions: true, hasUnansweredQuestions: false }), "confirm_plan");
});
