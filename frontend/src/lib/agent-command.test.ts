import assert from "node:assert/strict";
import test from "node:test";

import { isAgentProcessingInstruction } from "./agent-command.ts";

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
