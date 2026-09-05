import assert from "node:assert/strict";
import test from "node:test";

import { isAgentProcessingInstruction, isAgentStartCommand } from "./agent-command.ts";

test("recognizes natural-language workbook actions", () => {
  assert.equal(isAgentProcessingInstruction("继续"), true);
  assert.equal(isAgentProcessingInstruction("继续处理剩余项"), true);
  assert.equal(isAgentProcessingInstruction("把防暑降温费 K1 改成 23"), true);
  assert.equal(isAgentProcessingInstruction("请按手册更新表格"), true);
  assert.equal(isAgentProcessingInstruction("跳过解析手册，直接进行更新"), true);
  assert.equal(isAgentProcessingInstruction("不要解析手册，直接处理文件"), true);
});

test("keeps questions in normal chat", () => {
  assert.equal(isAgentProcessingInstruction("修改了什么？"), false);
  assert.equal(isAgentProcessingInstruction("为什么还没完成"), false);
  assert.equal(isAgentProcessingInstruction("为什么要跳过解析？"), false);
});

test("only explicit start commands launch execution during alignment", () => {
  assert.equal(isAgentStartCommand("开始处理"), true);
  assert.equal(isAgentStartCommand("开始执行！"), true);
  assert.equal(isAgentStartCommand("  开始。"), true);
  assert.equal(isAgentStartCommand("run"), true);
  // 说明性要求不是启动指令，必须留在对齐对话里。
  assert.equal(isAgentStartCommand("把奖金按更新表更新，缺人的先留着"), false);
  assert.equal(isAgentStartCommand("请按手册更新表格"), false);
  assert.equal(isAgentStartCommand("先不要开始，我再说两句"), false);
  assert.equal(isAgentStartCommand("什么时候开始处理？"), false);
});
