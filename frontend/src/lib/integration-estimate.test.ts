import assert from "node:assert/strict";
import test from "node:test";

import { estimateIntegrationSeconds, formatEstimateRange } from "./integration-estimate.ts";

test("estimate grows with workbook structure and change volume", () => {
  const small = estimateIntegrationSeconds({
    masterRows: 50,
    masterColumns: 80,
    masterSheets: 5,
    formulaCount: 200,
    changeCells: 2_000,
    changeFiles: 1,
  });
  const large = estimateIntegrationSeconds({
    masterRows: 500,
    masterColumns: 120,
    masterSheets: 25,
    formulaCount: 20_000,
    changeCells: 150_000,
    changeFiles: 20,
  });

  assert.ok(large > small);
  assert.ok(small >= 15);
  assert.ok(large <= 600);
});

test("estimate range is readable in seconds and minutes", () => {
  assert.equal(formatEstimateRange(30), "约 25–40 秒");
  assert.equal(formatEstimateRange(90), "约 1–2 分钟");
});

test("a typical payroll master stays in the observed sub-minute range", () => {
  const estimate = estimateIntegrationSeconds({
    masterRows: 289,
    masterColumns: 98,
    masterSheets: 21,
    formulaCount: 4_873,
    changeCells: 91 * 38,
    changeFiles: 1,
  });

  assert.ok(estimate >= 15 && estimate <= 25);
});
