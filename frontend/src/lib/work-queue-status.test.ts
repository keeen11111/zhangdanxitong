import assert from "node:assert/strict";
import test from "node:test";

import { getWorkQueueStatusDisplay } from "./work-queue-status.ts";

test("all queue items are shown as incomplete with their current step", () => {
  assert.deepEqual(getWorkQueueStatusDisplay("blocked"), {
    label: "未完成",
    step: "需重新整合",
    className: "bg-red-100 text-red-800",
  });

  assert.deepEqual(getWorkQueueStatusDisplay("review_required"), {
    label: "未完成",
    step: "需修正来源或映射",
    className: "bg-amber-100 text-amber-800",
  });

  assert.deepEqual(getWorkQueueStatusDisplay("ready_for_release"), {
    label: "未完成",
    step: "待确认发布",
    className: "bg-teal-100 text-teal-800",
  });
});
