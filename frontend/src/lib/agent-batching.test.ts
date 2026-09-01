import assert from "node:assert/strict";
import test from "node:test";

import type { AgentWorkItem } from "./api";
import { agentItemGroupKey, groupAgentItems, selectedGroupItemIds } from "./agent-batching.ts";

const item = (item_id: string, category?: string, field?: string): AgentWorkItem => ({
  item_id,
  status: "needs_conversation",
  category,
  field,
});

test("groups review items by business category across different target cells", () => {
  const groups = groupAgentItems([
    item("a", "奖金", "金额"),
    item("b", "奖金", "金额"),
    item("c", "考勤", "出勤天数"),
  ]);
  assert.deepEqual(groups.map((group) => [group.key, group.items.length]), [
    ["奖金", 2],
    ["考勤", 1],
  ]);
});

test("uses a stable fallback for items without a category or field", () => {
  const value = item("a");
  assert.equal(agentItemGroupKey(value), "其他待确认事项");
  assert.deepEqual(selectedGroupItemIds(groupAgentItems([value]), "其他待确认事项"), ["a"]);
});
