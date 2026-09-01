# Implementation Plan: 科园 2026.08 全量执行

## Phase 1: Execution State And Reporting

1. Add an explicit incomplete execution result and resumable phase state.
   - Acceptance: a turn-limit failure is visible as incomplete, retains completed phases, and cannot be published.
   - Verify: focused orchestrator and API tests.

2. Correct activity labels and the workbench status panel.
   - Acceptance: event count is not called completed actions; turn limit has a clear label and a retry action.
   - Verify: TypeScript tests, lint, and browser check.

## Phase 2: Safe Workbook Operations

3. Add audited structural operations for row insertion/copy, explicit value updates, and scoped formula divisor replacement.
   - Acceptance: OA, AU, AV, and 防暑降温费 updates preserve formulas and non-target cells.
   - Verify: fixture workbooks and formula integrity tests.

4. Add source-table matching and batch writes for bonus, duty, attendance, and ledger fields.
   - Acceptance: only unique employee matches with valid source fields write; all other records become review items.
   - Verify: per-operation fixture tests with mismatched names and missing values.

5. Add bounded rebuild support for other adjustments and tax export regions.
   - Acceptance: only the declared monthly data region changes; protected total formulas remain intact.
   - Verify: row-boundary and formula-protection tests.

## Phase 3: Validation And Release

6. Generate a unified review list and validation report, then connect the phase executor to the Agent run endpoint.
   - Acceptance: each planned step has an audit result and release remains blocked until all validation gates pass.
   - Verify: end-to-end API run against the supplied fixtures.

## Checkpoints

- After Phase 1: existing run reports the correct reason and can safely resume.
- After Phase 2: all supported workbook writes pass focused regression tests.
- After Phase 3: full Python tests, frontend lint/build, and browser verification pass.

## Risks

- Source sheet layouts can vary by month. Mitigation: table/header detection with evidence capture and fail-closed review items.
- Legacy `.xls` tax files need a compatible read path. Mitigation: validate parser availability before implementing tax mapping; do not silently omit a file.
- Formula and row operations can alter references. Mitigation: preserve styles/formulas through `openpyxl`, capture before/after formulas, and check protected totals.
