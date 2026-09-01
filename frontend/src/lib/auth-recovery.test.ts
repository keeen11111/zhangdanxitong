import assert from "node:assert/strict";
import test from "node:test";

import { recoverFromUnauthorized } from "./auth-recovery.ts";

test("a non-authentication error leaves the current session unchanged", () => {
  let cleared = false;
  let redirectTarget: string | null = null;

  const recovered = recoverFromUnauthorized(500, "/projects", {
    clear: () => { cleared = true; },
    redirect: (target) => { redirectTarget = target; },
  });

  assert.equal(recovered, false);
  assert.equal(cleared, false);
  assert.equal(redirectTarget, null);
});

test("an unauthorized response clears the expired session and returns to login", () => {
  let cleared = false;
  let redirectTarget: string | null = null;

  const recovered = recoverFromUnauthorized(401, "/projects/company/agent", {
    clear: () => { cleared = true; },
    redirect: (target) => { redirectTarget = target; },
  });

  assert.equal(recovered, true);
  assert.equal(cleared, true);
  assert.equal(redirectTarget, "/login");
});

test("a rejected login clears stale data without redirecting from the login page", () => {
  let cleared = false;
  let redirected = false;

  const recovered = recoverFromUnauthorized(401, "/login", {
    clear: () => { cleared = true; },
    redirect: () => { redirected = true; },
  });

  assert.equal(recovered, true);
  assert.equal(cleared, true);
  assert.equal(redirected, false);
});
