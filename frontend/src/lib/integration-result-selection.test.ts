import assert from "node:assert/strict";
import test from "node:test";

import {
  hasFinancialWorkbookFiles,
  selectLatestIntegrationResult,
  shouldLoadFinancialIntegrationResult,
} from "./integration-result-selection.ts";

test("does not probe generic integration results for legacy payroll files", () => {
  assert.equal(hasFinancialWorkbookFiles([{ file_type: "template" }, { file_type: "source" }]), false);
});

test("recognizes projects that can have a generic integration result", () => {
  assert.equal(hasFinancialWorkbookFiles([{ file_type: "financial_master" }]), true);
  assert.equal(hasFinancialWorkbookFiles([{ file_type: "financial_source" }]), true);
});

test("waits for generic integration progress before loading its latest result", () => {
  const files = [{ file_type: "financial_master" }];
  assert.equal(shouldLoadFinancialIntegrationResult(files, { status: "idle" }), false);
  assert.equal(shouldLoadFinancialIntegrationResult(files, { status: "completed" }), true);
  assert.equal(shouldLoadFinancialIntegrationResult([{ file_type: "template" }], { status: "completed" }), false);
});

test("selects the newer payroll result over an older generic result", () => {
  assert.equal(
    selectLatestIntegrationResult(
      { completed_at: "2026-08-26T11:00:00+08:00" },
      { completed_at: "2026-08-26T11:05:00+08:00" },
    ),
    "pipeline",
  );
});

test("keeps the generic result when it is the only available result", () => {
  assert.equal(
    selectLatestIntegrationResult({ completed_at: "2026-08-26T11:00:00+08:00" }, null),
    "financial",
  );
});

test("keeps the generic result when it is newer than the payroll result", () => {
  assert.equal(
    selectLatestIntegrationResult(
      { completed_at: "2026-08-26T11:05:00+08:00" },
      { completed_at: "2026-08-26T11:00:00+08:00" },
    ),
    "financial",
  );
});
