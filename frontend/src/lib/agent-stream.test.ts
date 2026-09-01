import assert from "node:assert/strict";
import test from "node:test";

import { parseSseFrames } from "./agent-stream.ts";

test("parses complete SSE frames and preserves an incomplete tail", () => {
  const parsed = parseSseFrames(
    'event: message.delta\ndata: {"content":"你好"}\n\nevent: message.completed\ndata: {"revision":3}\n\nevent: message',
  );

  assert.deepEqual(parsed.events, [
    { type: "message.delta", payload: { content: "你好" } },
    { type: "message.completed", payload: { revision: 3 } },
  ]);
  assert.equal(parsed.remaining, "event: message");
});

test("ignores malformed and unknown SSE frames", () => {
  const parsed = parseSseFrames("event: message.delta\ndata: not-json\n\n\n");

  assert.deepEqual(parsed.events, []);
  assert.equal(parsed.remaining, "");
});

test("parses event ids, comments, and multiline data", () => {
  const parsed = parseSseFrames(
    ': keep-alive\n\nid: 17\nevent: progress\ndata: {"label":\ndata: "已写入"}\n\n',
  );

  assert.deepEqual(parsed.events, [
    { id: "17", type: "progress", payload: { label: "已写入" } },
  ]);
  assert.equal(parsed.remaining, "");
});
