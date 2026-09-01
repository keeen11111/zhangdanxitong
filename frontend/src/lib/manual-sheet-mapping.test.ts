import assert from "node:assert/strict";
import test from "node:test";

import { validManualSheetMappings } from "./manual-sheet-mapping.ts";

test("keeps only checked mappings with a selected target sheet", () => {
  assert.deepEqual(
    validManualSheetMappings([
      { issue_id: "issue-1", target_sheet: "考勤" },
      { issue_id: "issue-2", target_sheet: "" },
      { issue_id: "", target_sheet: "工资" },
    ]),
    [{ issue_id: "issue-1", target_sheet: "考勤" }],
  );
});
