# BG69 Refactor Plan

Refactor roadmap for the existing Google Apps Script codebase.

Constraints:
- Keep Google Apps Script
- Keep `SpreadsheetApp` as the data store
- Keep HTML + vanilla JS frontend
- Keep `google.script.run` integration style
- Prefer incremental migration over full rewrite

## Goals

- Make changes easier to isolate
- Reduce cross-feature breakage
- Remove compatibility patches gradually
- Preserve current behavior while improving structure

## Phase 0: Baseline And Safety Net

Purpose: define the current behavior before moving files or changing structure.

### Checklist

- [ ] Expand the regression checklist from `README.md` into a fuller manual test sheet
- [ ] Confirm both entry routes still work:
  - [ ] main page via `doGet()`
  - [ ] support page via `?page=support`
- [ ] Verify these user flows end to end:
  - [ ] main page load
  - [ ] support page load
  - [ ] filters
  - [ ] record expense
  - [ ] history modal
  - [ ] transfer budget
  - [ ] dashboard view
  - [ ] dashboard export
  - [ ] edit expense
  - [ ] cancel expense
- [ ] Document current shims and fallback behaviors before touching them
- [ ] Capture known fragile spots and reproduction steps

### File Checklist

- [ ] [README.md](/C:/Users/Surag/Documents/bg69/bg69/README.md)
  - [ ] Turn the current pre-push checklist into a more explicit regression section
- [ ] [Index.html](/C:/Users/Surag/Documents/bg69/bg69/Index.html)
  - [ ] Inventory the shim blocks at the top of the file
- [ ] [SupportIndex.html](/C:/Users/Surag/Documents/bg69/bg69/SupportIndex.html)
  - [ ] Confirm page bootstrapping assumptions

### Deliverables

- Manual regression checklist
- Shim inventory
- Known-risk notes

## Phase 1: Backend File Split Without Behavior Change

Purpose: split the backend by responsibility while keeping the same public GAS function names.

### Target Backend Structure

- `app_config.gs`
- `app_sheet_utils.gs`
- `app_errors.gs`
- `app_auth.gs`
- `app_web.gs`
- `budget_service.gs`
- `transaction_service.gs`
- keep support files in place first, then normalize in Phase 2

### Checklist

- [ ] Move configuration constants into `app_config.gs`
- [ ] Move spreadsheet and header helper functions into `app_sheet_utils.gs`
- [ ] Move error logging and error wrapper logic into `app_errors.gs`
- [ ] Move user/auth/access helpers into `app_auth.gs`
- [ ] Move `doGet()` and `include()` into `app_web.gs`
- [ ] Move budget list, dashboard, alert, and transfer functions into `budget_service.gs`
- [ ] Move expense, history, edit, cancel, and transaction-log functions into `transaction_service.gs`
- [ ] Keep all existing exported function names unchanged
- [ ] Re-run Phase 0 regression after each move batch

### File Checklist

- [ ] [code.gs](/C:/Users/Surag/Documents/bg69/bg69/code.gs)
  - [ ] Move `CONFIG`
  - [ ] Move spreadsheet cache and sheet resolution helpers
  - [ ] Move error handling helpers
  - [ ] Move `doGet()` and `include()`
  - [ ] Move auth helpers
  - [ ] Move budget/domain functions into service files
  - [ ] Leave only temporary compatibility wrappers if needed
- [ ] [support_module.gs](/C:/Users/Surag/Documents/bg69/bg69/support_module.gs)
  - [ ] Do not refactor deeply yet
  - [ ] Note shared helper candidates for Phase 2
- [ ] New files:
  - [ ] [app_config.gs](/C:/Users/Surag/Documents/bg69/bg69/app_config.gs)
  - [ ] [app_sheet_utils.gs](/C:/Users/Surag/Documents/bg69/bg69/app_sheet_utils.gs)
  - [ ] [app_errors.gs](/C:/Users/Surag/Documents/bg69/bg69/app_errors.gs)
  - [ ] [app_auth.gs](/C:/Users/Surag/Documents/bg69/bg69/app_auth.gs)
  - [ ] [app_web.gs](/C:/Users/Surag/Documents/bg69/bg69/app_web.gs)
  - [ ] [budget_service.gs](/C:/Users/Surag/Documents/bg69/bg69/budget_service.gs)
  - [ ] [transaction_service.gs](/C:/Users/Surag/Documents/bg69/bg69/transaction_service.gs)

### Deliverables

- Smaller `.gs` modules by responsibility
- Existing frontend still calling the same server-side function names

## Phase 2: Shared Backend Helpers For Main And Support

Purpose: reduce duplication between budget and support flows.

### Checklist

- [ ] Identify duplicated logic between `code.gs` split files and `support_module.gs`
- [ ] Create shared helpers for:
  - [ ] item id normalization
  - [ ] number parsing
  - [ ] lock acquisition
  - [ ] row lookup by item id
  - [ ] transaction log append
  - [ ] access checks
- [ ] Separate shared mechanics from support-specific business rules
- [ ] Update support flow to call shared helpers rather than copying logic
- [ ] Re-run regression on both main and support flows

### File Checklist

- [ ] [support_module.gs](/C:/Users/Surag/Documents/bg69/bg69/support_module.gs)
  - [ ] Replace duplicated utility logic with shared helpers
- [ ] [support_report.gs](/C:/Users/Surag/Documents/bg69/bg69/support_report.gs)
  - [ ] Verify assumptions after helper consolidation
- [ ] [support_agg_reader.gs](/C:/Users/Surag/Documents/bg69/bg69/support_agg_reader.gs)
  - [ ] Verify read-path assumptions after helper consolidation
- [ ] [app_sheet_utils.gs](/C:/Users/Surag/Documents/bg69/bg69/app_sheet_utils.gs)
  - [ ] Expand as the shared utility home if appropriate
- [ ] New optional shared files if needed:
  - [ ] [shared_budget_helpers.gs](/C:/Users/Surag/Documents/bg69/bg69/shared_budget_helpers.gs)
  - [ ] [shared_transaction_helpers.gs](/C:/Users/Surag/Documents/bg69/bg69/shared_transaction_helpers.gs)

### Deliverables

- One source of truth for common backend mechanics
- Less drift between main and support logic

## Phase 3: Frontend Module Boundaries

Purpose: keep the current HTML include model, but make the JS responsibilities cleaner.

### Target Frontend Split

- shared rpc layer
- shared state/store layer
- shared UI helpers
- main page controller
- support page controller
- feature modules for modal/export/dashboard/work-detail

### Checklist

- [ ] Standardize one RPC entry path and reduce direct `google.script.run` calls
- [ ] Create a clearer shared state surface instead of ad hoc `window.State` writes
- [ ] Separate main page initialization from support page initialization
- [ ] Move page-specific behavior out of shared modules where possible
- [ ] Reduce direct cross-calls through global functions
- [ ] Keep existing HTML includes working during the transition

### File Checklist

- [ ] [client_rpc_adapter.html](/C:/Users/Surag/Documents/bg69/bg69/client_rpc_adapter.html)
  - [ ] Make this the preferred shared RPC path
- [ ] [JavaScript_core.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_core.html)
  - [ ] Narrow to bootstrap, load, and state update responsibilities
- [ ] [JavaScript_ui.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_ui.html)
  - [ ] Narrow to rendering and visual state only
- [ ] [JavaScript_actions.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_actions.html)
  - [ ] Keep event routing simple and explicit
- [ ] [support_client.html](/C:/Users/Surag/Documents/bg69/bg69/support_client.html)
  - [ ] Reduce support-specific fallback complexity where shared modules can take over
- [ ] [JavaScript_modal_manager.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_modal_manager.html)
  - [ ] Keep modal lifecycle centralized
- [ ] [JavaScript_alerts_manager.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_alerts_manager.html)
  - [ ] Verify it reads alerts from shared state only
- [ ] [JavaScript_export_enhancements.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_export_enhancements.html)
  - [ ] Keep export concerns isolated from page bootstrap
- [ ] [JavaScript_workdetail_modal.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_workdetail_modal.html)
  - [ ] Keep fetch/open lifecycle separate from rendering
- [ ] [JavaScript_workdetail_ui.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_workdetail_ui.html)
  - [ ] Keep rendering-only responsibility

### Deliverables

- Cleaner JS ownership boundaries
- Less global coupling between main/support pages

## Phase 4: Remove Compatibility Patches Gradually

Purpose: move shims off the hot path once replacement flow is stable.

### Checklist

- [ ] Inventory all shim, stub, fallback, and retry layers
- [ ] Classify each as:
  - [ ] required for now
  - [ ] temporary
  - [ ] removable
- [ ] Replace inline onclick dependencies where practical
- [ ] Remove fallback paths only after replacement flow is verified
- [ ] Remove dead code after each cleanup

### File Checklist

- [ ] [Index.html](/C:/Users/Surag/Documents/bg69/bg69/Index.html)
  - [ ] Review safe iframe write shim
  - [ ] Review init client stub
  - [ ] Review pointer/event interception block
- [ ] [support_client.html](/C:/Users/Surag/Documents/bg69/bg69/support_client.html)
  - [ ] Review `window.__support_on_btn`
  - [ ] Review fallback card actions
  - [ ] Review history handler retries
- [ ] [JavaScript_fallback_shims.html](/C:/Users/Surag/Documents/bg69/bg69/JavaScript_fallback_shims.html)
  - [ ] Review all remaining compatibility behavior
- [ ] [ExpenseEdit.html](/C:/Users/Surag/Documents/bg69/bg69/ExpenseEdit.html)
  - [ ] Check whether this can be reduced after shared modal/service cleanup
- [ ] [ViewToggle.html](/C:/Users/Surag/Documents/bg69/bg69/ViewToggle.html)
  - [ ] Check whether repeated init guards are still needed

### Deliverables

- Smaller runtime patch surface
- More predictable page lifecycle

## Phase 5: CSS And UI Ownership Cleanup

Purpose: reduce override battles and make styling safer to edit.

### Checklist

- [ ] Split shared styles by concern:
  - [ ] base
  - [ ] layout
  - [ ] cards
  - [ ] forms
  - [ ] modals
  - [ ] dashboard
  - [ ] print/export
  - [ ] support-specific
- [ ] Reduce `!important` usage in normal runtime styles
- [ ] Keep print/export overrides isolated
- [ ] Normalize component naming and ownership

### File Checklist

- [ ] [Styles.html](/C:/Users/Surag/Documents/bg69/bg69/Styles.html)
  - [ ] Separate legacy rules from current rules
  - [ ] Group runtime vs print/export sections
  - [ ] Reduce duplicate `.budget-card`, `.btn`, and `body` definitions
- [ ] [Index.html](/C:/Users/Surag/Documents/bg69/bg69/Index.html)
  - [ ] Remove inline style dependencies where practical
- [ ] [SupportIndex.html](/C:/Users/Surag/Documents/bg69/bg69/SupportIndex.html)
  - [ ] Remove inline style dependencies where practical
- [ ] Export-related templates:
  - [ ] [exportDashboardCoordinator.html](/C:/Users/Surag/Documents/bg69/bg69/exportDashboardCoordinator.html)
  - [ ] [exportDashboardHeader.html](/C:/Users/Surag/Documents/bg69/bg69/exportDashboardHeader.html)
  - [ ] [exportDashboardTable.html](/C:/Users/Surag/Documents/bg69/bg69/exportDashboardTable.html)
  - [ ] [SupportReportCaptureTemplate.html](/C:/Users/Surag/Documents/bg69/bg69/SupportReportCaptureTemplate.html)

### Deliverables

- More predictable styling
- Less fear when changing UI components

## Phase 6: Encoding And Text Cleanup

Purpose: remove text corruption and improve maintainability.

### Checklist

- [ ] Normalize repo text files to UTF-8
- [ ] Fix corrupted Thai strings in backend and frontend files
- [ ] Consolidate repeated labels and messages where worthwhile
- [ ] Re-check exports, alerts, and modal messages after conversion

### File Checklist

- [ ] [code.gs](/C:/Users/Surag/Documents/bg69/bg69/code.gs)
  - [ ] Fix corrupted user-facing Thai strings
- [ ] [Index.html](/C:/Users/Surag/Documents/bg69/bg69/Index.html)
  - [ ] Fix corrupted metadata/title text
- [ ] [SupportIndex.html](/C:/Users/Surag/Documents/bg69/bg69/SupportIndex.html)
  - [ ] Fix corrupted labels
- [ ] [support_module.gs](/C:/Users/Surag/Documents/bg69/bg69/support_module.gs)
  - [ ] Fix corrupted sheet headers/messages
- [ ] [support_client.html](/C:/Users/Surag/Documents/bg69/bg69/support_client.html)
  - [ ] Fix corrupted UI text
- [ ] [Styles.html](/C:/Users/Surag/Documents/bg69/bg69/Styles.html)
  - [ ] Verify comments and labels if present

### Deliverables

- Readable Thai text throughout the app
- Safer validation and alert maintenance

## Phase 7: Hardening And Regression Pass

Purpose: make sure the cleaned structure still behaves correctly.

### Checklist

- [ ] Run the full regression list from Phase 0
- [ ] Verify permission behavior for:
  - [ ] admin
  - [ ] normal user
  - [ ] viewer
- [ ] Verify transaction log integrity after create/edit/cancel
- [ ] Verify support and main flows still honor access boundaries
- [ ] Verify dashboard/export/report flows
- [ ] Remove dead code discovered during the pass
- [ ] Update README with the new structure

### File Checklist

- [ ] [README.md](/C:/Users/Surag/Documents/bg69/bg69/README.md)
  - [ ] Update project layout after the split
  - [ ] Update test checklist
- [ ] [appsscript.json](/C:/Users/Surag/Documents/bg69/bg69/appsscript.json)
  - [ ] Confirm manifest still matches runtime needs
- [ ] all touched `.gs` and `.html`
  - [ ] verify no stale include or moved-function assumptions remain

### Deliverables

- Verified refactor baseline
- Updated docs

## Suggested Commit Sequence

This is the safest first-pass commit order.

### Commit 1

`docs: add phased refactor plan and regression checklist`

Scope:
- add this plan file
- optionally expand `README.md` with a stronger manual test checklist

### Commit 2

`refactor(gas): extract config, web, auth, and error helpers`

Scope:
- create `app_config.gs`
- create `app_web.gs`
- create `app_auth.gs`
- create `app_errors.gs`
- move only low-risk shared helpers out of `code.gs`
- no behavior changes

### Commit 3

`refactor(gas): extract sheet utilities and budget service`

Scope:
- create `app_sheet_utils.gs`
- create `budget_service.gs`
- move read-heavy budget/dashboard/alert functions
- preserve server function names

### Commit 4

`refactor(gas): extract transaction service`

Scope:
- create `transaction_service.gs`
- move record/edit/cancel/history/log functions
- verify lock behavior and transaction logging

### Commit 5

`refactor(gas): consolidate shared main and support helpers`

Scope:
- reduce duplication across support and main backend logic

### Commit 6

`refactor(frontend): centralize rpc and page bootstrap responsibilities`

Scope:
- reduce direct scattered `google.script.run`
- narrow `JavaScript_core.html` and `support_client.html`

### Commit 7

`refactor(frontend): remove obsolete shims and simplify event handling`

Scope:
- remove validated dead fallback code
- keep only the shims still justified by runtime behavior

### Commit 8

`refactor(ui): normalize styles and separate print/export rules`

Scope:
- clean `Styles.html`
- reduce duplicate selectors and `!important` reliance

### Commit 9

`fix(i18n): normalize utf8 text and repair corrupted thai strings`

Scope:
- text cleanup pass
- verify alerts, labels, sheet headers, and exports

### Commit 10

`docs: update project structure after refactor`

Scope:
- update `README.md`
- note remaining known risks and deferred cleanup

## Recommended First Working Slice

If starting now, do this first:

1. Commit the docs plan
2. Extract `CONFIG`, `doGet()`, `include()`, auth helpers, and error helpers
3. Re-test main page load and support page load
4. Re-test one write flow: record expense

That gives a low-risk first structural win without touching the most fragile UI paths yet.
