import assert from "node:assert/strict";
import test from "node:test";

import { fetchWithTimeout, withRequestTimeout } from "./request-timeout.ts";

test("a stalled upload stops with an actionable timeout message", async () => {
  await assert.rejects(
    withRequestTimeout(
      (signal) => new Promise((_resolve, reject) => {
        signal.addEventListener("abort", () => reject(new Error("aborted")), { once: true });
      }),
      10,
      "上传等待超时，请重试",
    ),
    /上传等待超时，请重试/,
  );
});

test("a completed request is returned before the timeout", async () => {
  const result = await withRequestTimeout(async () => "uploaded", 100, "不应超时");

  assert.equal(result, "uploaded");
});

test("a stalled API request is aborted instead of leaving the page loading forever", async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = ((_input: RequestInfo | URL, init?: RequestInit) =>
    new Promise<Response>((_resolve, reject) => {
      init?.signal?.addEventListener("abort", () => reject(new Error("aborted")), { once: true });
    })) as typeof fetch;

  try {
    await assert.rejects(
      fetchWithTimeout("/api/pipeline/demo/export/latest", {}, 10, "加载结果超时，请重试"),
      /加载结果超时，请重试/,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});
