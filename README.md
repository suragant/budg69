
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

