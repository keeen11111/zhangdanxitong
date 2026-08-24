import assert from "node:assert/strict";
import test from "node:test";

import { triggerBrowserDownload } from "./browser-download.ts";

test("direct download clicks a named browser download and releases the blob url", () => {
  let clicked = false;
  let attached = false;
  let removed = false;
  let revoked = "";
  const anchor = { href: "", download: "", click: () => { clicked = true; } };

  triggerBrowserDownload(new Blob(["workbook"]), "工资核算.xlsx", {
    createAnchor: () => anchor,
    attachAnchor: () => { attached = true; },
    removeAnchor: () => { removed = true; },
    createObjectUrl: () => "blob:workbook",
    revokeObjectUrl: (url) => { revoked = url; },
    scheduleCleanup: (cleanup) => {
      assert.equal(attached, true);
      assert.equal(clicked, true);
      assert.equal(removed, false);
      assert.equal(revoked, "");
      cleanup();
    },
  });

  assert.equal(anchor.href, "blob:workbook");
  assert.equal(anchor.download, "工资核算.xlsx");
  assert.equal(attached, true);
  assert.equal(clicked, true);
  assert.equal(removed, true);
  assert.equal(revoked, "blob:workbook");
});
