import assert from "node:assert/strict";
import test from "node:test";

import { getManualReviewGuidance } from "./manual-review-guidance.ts";

test("shows source candidates before asking an operator to choose a value", () => {
  const guidance = getManualReviewGuidance({
    target_field: "实际出勤天数",
    candidate_values: [21, 23],
    source_files: ["7月考勤.xlsx"],
    source_sheets: ["考勤明细"],
  });

  assert.match(guidance, /21、23/);
  assert.match(guidance, /7月考勤.xlsx/);
  assert.match(guidance, /考勤明细/);
});

test("explains the expected date format for manual date entry", () => {
  assert.match(
    getManualReviewGuidance({ target_field: "入职日期" }),
    /YYYY-MM-DD/,
  );
});
