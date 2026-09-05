import assert from "node:assert/strict";
import test from "node:test";

import { mergeRecentProject } from "./project-list.ts";

const project = (id: string) => ({
  id,
  name: id,
  salary_month: "2026.09",
  status: "import",
  created_at: "2026-09-04T00:00:00Z",
  updated_at: "2026-09-04T00:00:00Z",
  file_count: 0,
  has_result: false,
  result_completed_at: null,
  pending_issue_count: 0,
});

test("a newly created project is placed at the front of the recent list", () => {
  const merged = mergeRecentProject([project("old-a"), project("old-b")], project("new"));

  assert.deepEqual(merged.map((item) => item.id), ["new", "old-a", "old-b"]);
});

test("project creation does not duplicate an existing sidebar item", () => {
  const merged = mergeRecentProject([project("old-a"), project("old-b")], project("old-b"));

  assert.deepEqual(merged.map((item) => item.id), ["old-b", "old-a"]);
});
