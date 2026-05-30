
# budg69

Google Apps Script web app for budget management and support-budget reporting.

## Project Layout

- `code.gs`
  Main backend for the primary budget app.
- `support_module.gs`
  Support-budget backend logic.
- `support_report.gs`
  Support report generation and PDF export helpers.
- `support_agg_reader.gs`
  Support aggregate reader.
- `Index.html`
  Main web app entry page.
- `SupportIndex.html`
  Support page entry.
- `Styles.html`
  Shared styles.
- `JavaScript_*.html`
  Frontend modules and UI helpers.
- `appsscript.json`
  Apps Script manifest.

## Local Naming Rules

- Server-side Apps Script files use `.gs`
- HTML templates and includes use `.html`
- `createTemplateFromFile('Index')` maps to `Index.html`
- `include('Styles')` maps to `Styles.html`

## Current Frontend Structure

- `JavaScript_core.html`
  Data loading, filter application, page bootstrap.
- `JavaScript_ui.html`
  Card rendering, history view, dashboard view, UI state display.
- `JavaScript_modal_manager.html`
  Modal lifecycle and shared backdrop management.
- `client_rpc_adapter.html`
  Promise-based RPC wrappers with retry support.
- `JavaScript_export_enhancements.html`
  Export-specific bindings for dashboard and work-detail modals.
- `JavaScript_fallback_shims.html`
  Lightweight compatibility shims for alerts and work-detail fallback behavior.
- `JavaScript_workdetail_modal.html`
  Work-detail modal creation/opening wrapper.
- `JavaScript_workdetail_ui.html`
  Work-detail table rendering tweaks and action cleanup.
- `JavaScript_card_export_visibility.html`
  Hides stray export buttons in card contexts while preserving modal exports.

## Files Excluded From Apps Script Push

Defined in `.claspignore`:

- `debug.gs`
- `debug_relaxed.gs`
- local tooling/config files such as `.clasp.json`

## Apps Script / Clasp

1. Create `.clasp.json` from `.clasp.json.example`
2. Put in your real `scriptId`
3. Push with:

```powershell
clasp push
```

## Recommended Pre-Push Checklist

1. Confirm `Index.html` and `SupportIndex.html` still match `doGet()` in `code.gs`
2. Confirm `.claspignore` excludes local-only files
3. Review `appsscript.json`
4. Test main page load
5. Test support page load
6. Test expense save flow
7. Test history modal
8. Test dashboard open/export flow

## Manual Regression Checklist

Use this checklist before and after refactor work. The goal is to confirm that the current behavior still works, not to improve the flow while testing it.

### Main Page Load

- [ ] Open the main page from `doGet()`
- [ ] Confirm the page finishes bootstrapping without a blank screen
- [ ] Confirm cards or primary budget content render
- [ ] Confirm the current user header area renders
- [ ] Check whether the UI behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Support Page Load

- [ ] Open the support page with `?page=support`
- [ ] Confirm support content renders without getting stuck in a loading state
- [ ] Confirm the support list or cards appear
- [ ] Confirm the support header controls render and respond
- [ ] Check whether the UI behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Filter And Search

- [ ] Use the main page filters/search inputs
- [ ] Confirm filtered results update
- [ ] Confirm counts, totals, and warning badges stay in sync
- [ ] Clear filters and confirm the default list returns
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Open Expense Modal

- [ ] Open the expense modal from a main budget item
- [ ] Confirm the modal shows the correct item context
- [ ] Confirm required fields and default values render correctly
- [ ] Close and reopen the modal once to catch duplicate binding issues
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Save Expense

- [ ] Record one expense from the main flow
- [ ] Confirm success feedback appears once
- [ ] Confirm the modal closes cleanly
- [ ] Confirm the list refreshes with updated used/remaining values
- [ ] If possible, verify the transaction log entry was created correctly
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### History Modal

- [ ] Open history from a main item
- [ ] Confirm records load for the selected item
- [ ] Confirm the selected item context shown in the modal is correct
- [ ] Close and reopen history from another item
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Transfer Budget

- [ ] Open the transfer budget flow
- [ ] Confirm source and target item selection works
- [ ] Submit a valid transfer
- [ ] Confirm budget and remaining values refresh correctly on both items
- [ ] Confirm the result is reflected in related views if shown
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Dashboard Open

- [ ] Open the dashboard view
- [ ] Confirm summary values render
- [ ] Confirm charts/tables/cards load without broken layout
- [ ] Open any work-detail drilldown that is part of the dashboard flow
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Dashboard Export

- [ ] Run the dashboard export flow
- [ ] Confirm the export starts from the intended view
- [ ] Confirm the generated output contains the expected header/content blocks
- [ ] Confirm no extra export buttons or layout artifacts appear in the output
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Edit Expense

- [ ] Open a transaction that supports edit
- [ ] Submit a small edit to amount, date, or description
- [ ] Confirm the edited values are reflected in history and remaining budget
- [ ] Confirm edit status or metadata updates as expected
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

### Cancel Expense

- [ ] Open a transaction that supports cancel
- [ ] Cancel the transaction with a test reason if required
- [ ] Confirm used/remaining values roll back correctly
- [ ] Confirm history/log status reflects the cancel action
- [ ] Check whether the flow behaves normally:
  - [ ] opens successfully
  - [ ] shows expected data
  - [ ] does not show unexpected fallback behavior, duplicate handlers, or visible runtime errors

## Known Fragile Areas

These areas deserve extra attention during refactor and verification:

- Main page bootstrapping
  - `Index.html` contains multiple early safety/shim blocks before the regular app modules load
- Support page bootstrapping
  - `SupportIndex.html` depends on shared modules plus support-specific client behavior and fallbacks
- Modal and history actions
  - Main and support flows both use layered handlers, fallback paths, and global function routing
- Export and print flows
  - Dashboard/support export paths rely on dedicated templates, print-specific CSS, and extra JS coordination
- Shared frontend state
  - Several modules write directly to `window.State` and other global flags
- CSS override layers
  - `Styles.html` contains multiple override sections and heavy `!important` usage

## Day 1 Refactor Baseline

Day 1 of the refactor plan is documentation-only. No runtime logic should change in this step.

Definition of done:

- [ ] `README.md` contains a usable manual regression checklist
- [ ] `README.md` identifies fragile areas worth watching during refactor
- [ ] `REFACTOR_PLAN.md` stays aligned with the current repository structure
- [ ] No server-side or frontend behavior is intentionally changed in this step

## Before Refactor

Run through this short checklist before starting any refactor commit:

- [ ] Check the current working tree and note any unrelated in-progress changes
- [ ] Confirm which files are intentionally in scope for the next commit
- [ ] Confirm `doGet()` still routes to the expected templates:
  - [ ] main page -> `Index.html`
  - [ ] support page -> `SupportIndex.html`
- [ ] Choose at least one write flow to use as a smoke test after the change:
  - [ ] main expense save
  - [ ] support expense save
  - [ ] transfer budget
- [ ] Re-read the relevant section of `REFACTOR_PLAN.md` before moving files or changing responsibilities
- [ ] Confirm whether the next change is documentation-only, structure-only, or behavior-changing

## After Each Refactor Commit

Run a minimal smoke test after every refactor-oriented commit, even if the change looks internal:

- [ ] Open the main page
- [ ] Open the support page
- [ ] Run one expense save flow
- [ ] Open one history modal
- [ ] Confirm there are no obvious duplicate event handlers, blank sections, or broken modals
- [ ] If the commit touched dashboard/export code, also run:
  - [ ] dashboard open
  - [ ] dashboard export

## Out Of Scope For Commit 1

Commit 1 is for documentation and alignment only. It should not include these changes:

- [ ] No `.gs` backend file split yet
- [ ] No frontend architecture rewrite
- [ ] No removal of shims, stubs, or fallback layers
- [ ] No CSS cleanup or large visual changes
- [ ] No renaming of server functions used by `google.script.run`
- [ ] No intentional runtime behavior changes

## Day 2 Working Agreement

Day 2 turns the refactor notes into a repeatable team workflow.

Definition of done:

- [ ] `README.md` includes a clear pre-refactor checklist
- [ ] `README.md` includes a clear post-commit smoke-test checklist
- [ ] `README.md` states that Commit 1 remains documentation-only
- [ ] The repo now has a shared working agreement for the next refactor commits

## Refactor Status

Completed:

- Renamed source files from `.txt` to `.gs` / `.html`
- Rebuilt `Index.html`
- Extracted core/UI/modal responsibilities
- Moved expense save flow into `JavaScript_core.html`
- Removed `record-expense-hotfix.html`
- Reduced work-detail code from patch-heavy scripts into narrower modules
- Removed empty include modules that no longer contributed runtime behavior
- Added `appsscript.json`
- Added `clasp` helper files

Still worth improving:

- Consider merging work-detail modules into a single first-class feature module
- Add browser-level verification after the recent frontend refactor

