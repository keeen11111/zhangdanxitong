import assert from "node:assert/strict";
import test from "node:test";

import {
  FINANCIAL_INTEGRATION_TIMEOUT_MS,
  financialIntegrationProgressPercent,
} from "./financial-workbook-flow.ts";

test("financial integration keeps waiting for long-running workbooks", () => {
  assert.equal(FINANCIAL_INTEGRATION_TIMEOUT_MS, 15 * 60 * 1000);
});

test("financial integration progress reaches completion with all stages done", () => {
  const stages = [
    { key: "prepare", label: "准备", status: "completed" as const },
    { key: "match", label: "核对", status: "completed" as const },
    { key: "update", label: "写入", status: "completed" as const },
    { key: "export", label: "导出", status: "completed" as const },
  ];

  assert.equal(financialIntegrationProgressPercent(stages), 100);
});

test("financial integration progress stays visibly active while a stage is running", () => {
  const stages = [
    { key: "prepare", label: "准备", status: "completed" as const },
    { key: "match", label: "核对", status: "running" as const },
    { key: "update", label: "写入", status: "pending" as const },
    { key: "export", label: "导出", status: "pending" as const },
  ];

  assert.equal(financialIntegrationProgressPercent(stages), 37);
});
