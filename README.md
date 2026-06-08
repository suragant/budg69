
# budg69

Google Apps Script web app for budget management and support-budget reporting.

## Project Layout

### Backend (Google Apps Script `.gs` files)

- `app_config.gs` — Configuration constants
- `app_sheet_utils.gs` — Spreadsheet helpers, header resolution, shared utilities
- `app_errors.gs` — Error logging and wrapper
- `app_auth.gs` — User/auth/access helpers
- `app_web.gs` — `doGet()` and `include()`
- `budget_service.gs` — Budget list, dashboard, alerts, transfers
- `transaction_service.gs` — Expense CRUD, history, edit, cancel, transaction log
- `admin_service.gs` — Email alerts, PDF export
- `support_module.gs` — Support-budget backend logic
- `support_report.gs` — Support report generation and PDF export helpers
- `support_agg_reader.gs` — Support aggregate reader
- `code.gs` — Minimal stub (functions moved to service files)

### Frontend (HTML includes)

- `Index.html` — Main web app entry page
- `SupportIndex.html` — Support page entry
- `Styles.html` — Shared styles (2026 UI refresh)
- `client_rpc_adapter.html` — Canonical RPC layer (`rpcWithRetry`)
- `JavaScript_core.html` — Data loading, filter application, page bootstrap
- `JavaScript_ui.html` — Card rendering, history view, dashboard view
- `JavaScript_modal_manager.html` — Modal lifecycle and backdrop management
- `JavaScript_actions.html` — Event delegation for main page buttons
- `JavaScript_helpers.html` — Shared utilities (`escapeHtml`, etc.)
- `JavaScript_alerts_manager.html` — Alert modal and badge management
- `JavaScript_export_enhancements.html` — Export bindings (PDF, dashboard)
- `JavaScript_workdetail_modal.html` — Work-detail modal open/create
- `JavaScript_workdetail_ui.html` — Work-detail table rendering
- `JavaScript_card_export_visibility.html` — Hide stray export buttons in cards
- `JavaScript_fallback_shims.html` — Compatibility shims and bootstrap safety net
- `ExpenseEdit.html` — Edit/cancel expense modals
- `ViewToggle.html` — Card/tile view toggle and sorting
- `support_client.html` — Support page client logic
- `showSupportView.html` — Support view rendering
- `support_inline_quantity.html` — Inline quantity editing for support
- `support_export.html` — Support export helpers
- `TransferBudgetModal.html` — Budget transfer modal

### Config

- `appsscript.json` — Apps Script manifest
- `REFACTOR_PLAN.md` — Phased refactor roadmap

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

## Refactor Complete

The refactor is complete. See `REFACTOR_PLAN.md` for the full roadmap.

### What Changed

- Backend split into 7 service files by responsibility
- Shared helpers in `app_sheet_utils.gs` reduce duplication
- Frontend RPC centralized through `rpcWithRetry`
- Dead code and duplicate CSS removed
- Thai text encoding repaired

### What's Left

- Manual regression testing (use the checklist above)
- Consider further CSS `!important` reduction
- Consider removing the safe event interception block in `Index.html`

## Refactor Status

Completed:

- Phase 0: Documentation, regression checklist, REFACTOR_PLAN.md
- Phase 1: Backend file split — `code.gs` split into 7 service files
- Phase 2: Shared backend helpers — unified utilities in `app_sheet_utils.gs`
- Phase 3: Frontend RPC centralization — `rpcWithRetry` as single RPC path, escape dedup
- Phase 4: Dead code removal — IE meta, debug utilities, unused variables, redundant retries
- Phase 5: CSS cleanup — removed 111 lines of duplicate definitions
- Phase 6: Thai text repair — fixed 2 corrupted strings in `JavaScript_ui.html`

Still worth improving:

- Consider merging work-detail modules into a single first-class feature module
- Phase 7 candidate: further CSS `!important` reduction (low priority)
- Consider removing the safe event interception block in `Index.html` (currently kept as safety net)

