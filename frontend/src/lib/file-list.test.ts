import assert from "node:assert/strict";
import test from "node:test";

import { mergeFileRecords } from "./file-list.ts";

test("second upload is added even when it has the same filename", () => {
  const current = [{ id: "first", original_name: "工资表.xlsx" }];
  const secondUpload = [{ id: "second", original_name: "工资表.xlsx" }];

  const merged = mergeFileRecords(current, secondUpload);

  assert.equal(merged.length, 2);
  assert.deepEqual(merged.map((file) => file.id), ["first", "second"]);
});

test("the same server record is not counted twice", () => {
  const current = [{ id: "first", original_name: "工资表.xlsx" }];
  const repeatedResponse = [{ id: "first", original_name: "工资表-更新.xlsx" }];

  const merged = mergeFileRecords(current, repeatedResponse);

  assert.equal(merged.length, 1);
  assert.equal(merged[0].original_name, "工资表-更新.xlsx");
});
