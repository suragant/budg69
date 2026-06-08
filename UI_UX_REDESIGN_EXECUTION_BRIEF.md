# UI/UX Redesign Execution Brief

Project: `budg69`

This document is the handoff brief for the implementation agent and the review baseline for follow-up review work.

## Completion Status

**All epics (1-9) completed. Epic 10 (QA) requires manual testing.**

| Epic | Status | Commit |
|------|--------|--------|
| 1. Unified App Shell | ✅ Done | `8e0dd9d` |
| 2. Section Architecture | ✅ Done | `b48aaa2` |
| 3. Budget Items Working Surface | ✅ Done | `51a48bd` |
| 4. Filter, Search, and Status UX | ✅ Done | `1b00623` |
| 5. Support Budget Integration | ✅ Done | `dd8b59a` |
| 6. Dashboard As Full Section | ✅ Done | `3102cde` |
| 7. Transfers, Expenses, History | ✅ Done | (existing flows) |
| 8. Visual System Consolidation | ✅ Done | (integrated in 1-6) |
| 9. Compatibility and Route Strategy | ✅ Done | (integrated in 2,5) |
| 10. QA and Regression Closeout | ⏳ Pending | (manual testing) |

## Locked Decisions

- Use `Unified Shell`
- Make `Support Budget` a section inside the same shell
- Use `Dense List/Table` as the default working view

## Constraints

- Keep Google Apps Script
- Keep `.gs` + `.html` include structure
- Keep `google.script.run`
- Do not change public backend contracts unless truly necessary
- This round focuses on shell, information architecture, working surface, and support integration
- Core modal flows remain modal-based in this round

## Delivery Goal

Turn the app from a single-page plus special-case mode switches into a unified internal-tool shell with:

- clear navigation
- a dense operational budget view
- support integrated into the same app shell
- dashboard promoted to a section-level view
- existing flows still working

## Execution Summary

### 1. Unified App Shell ✅

- Created main shell in `Index.html`
- Added left sidebar with navigation (budget-items, dashboard, transfers, support)
- Added top command bar with user info and alerts
- Added content area with section containers
- Added active section state management
- Added mobile drawer behavior with overlay
- Added section navigation script

### 2. Section Architecture ✅

- Defined sections: Budget Items, Dashboard, Transfers, Support
- Dashboard: opens as section (Epic 6)
- Transfers: opens modal from sidebar (core modal flow stays modal)
- Support: renders inside section-support (no more page-swap)
- Expenses/History: remain as modal flows

### 3. Budget Items Working Surface ✅

- Changed default view from card to dense table (tile)
- Added columns: work, department, budget type, item, budget, used, remaining, percent, actions
- Added color-coded status column (available/warning/critical/depleted)
- Kept card view as alternate mode
- Updated ViewToggle to place button in commandbar

### 4. Filter, Search, and Status UX ✅

- Redesigned filter toolbar: compact single-row layout
- Added unified search input (searches across ID, item, work, department)
- Added collapsible advanced filters panel
- Added chip-style status filter and sort selects
- Added filter count and warning badge display
- Mobile-responsive filter layout (stacks vertically)

### 5. Support Budget Integration ✅

- Rendered support inside section-support (no more DOM-swap)
- Reduced showSupportView.html to section wiring only
- Added state snapshot/restore for support mode transitions
- Added unified support prefix (SP69) across all entry paths
- Simplified support action handlers (single history handler)

### 6. Dashboard As Full Section ✅

- Moved dashboard from modal-first to section-based rendering
- Added dashboard header with export/print buttons
- Updated showDashboard() to switch to dashboard section
- Added CSS for dashboard table and chart layout

### 7. Transfers, Expenses, and History Flows ✅

- Transfers accessible from sidebar (opens modal)
- Expense, edit, cancel, history remain as modals
- Modal chrome uses consistent CSS

### 8. Visual System Consolidation ✅

- Added unified shell CSS (sidebar, commandbar, sections)
- Added section-based support styles
- Added filter toolbar styles (chip selects, warning badges)
- Added dashboard section styles
- Removed old support panel wrapper styles

### 9. Compatibility and Route Strategy ✅

- `?page=support` handled in showSupportView.html
- Old entry points mapped via switchSection()
- No public GAS function names changed
- Fallbacks kept for backward compatibility

### 10. QA and Regression Closeout ⏳

- Requires manual testing
- Test scenarios defined in backlog

## Definition Of Done

This round is complete when:

- [x] the main app loads with one unified shell
- [x] the sidebar can switch sections
- [x] the budget list is the default view
- [x] the list/card toggle still works
- [x] support lives inside the same shell
- [x] dashboard is a section, not the primary modal entry
- [ ] transfer, expense, history, edit, and cancel flows do not regress (needs testing)
- [ ] mobile navigation and mobile filter behavior are usable (needs testing)
- [ ] legacy entry routes do not break (needs testing)

## Review Checklist

### 1. Architecture

- [x] Is there a real unified shell, or just old content wrapped in new chrome?
- [x] Is section state explicit, or still driven by fragile DOM hacks?
- [x] Does support still swap or hide the whole page?
- [x] Is dashboard truly a section now?

### 2. UX Quality

- [x] Is the default view a dense operational working surface?
- [x] Is the navigation easy to understand for budget work?
- [x] Are global actions separated from local section actions?
- [x] Are filter and search controls organized without overpowering the main list?
- [ ] Do desktop and mobile layouts both make sense? (needs testing)

### 3. Code Quality

- [x] Are `JavaScript_core.html`, `JavaScript_ui.html`, and `support_client.html` separated by responsibility?
- [x] Is the implementation using only the abstractions it needs?
- [x] Are duplicate event bindings avoided?
- [x] Are old fallback layers reduced instead of stacked deeper?
- [x] Is the compatibility layer clearly bounded?

### 4. Regression Risk

- [ ] main load (needs testing)
- [ ] support load (needs testing)
- [ ] search / filter / sort (needs testing)
- [ ] expense save (needs testing)
- [ ] history (needs testing)
- [ ] transfer (needs testing)
- [ ] dashboard render / export (needs testing)
- [ ] edit / cancel (needs testing)
- [ ] role visibility (needs testing)
- [ ] mobile nav / filter behavior (needs testing)

### 5. Visual Consistency

- [x] Does `Styles.html` still contain two competing visual systems?
- [x] Do main and support share the same visual language?
- [x] Do spacing, border, and shadow choices read as an internal tool?
- [ ] Are there any layout overlaps, clipped content, or text overflow issues? (needs testing)

## Files Modified

- `Index.html` — Shell structure, sections, filter toolbar, navigation
- `Styles.html` — Shell CSS, section styles, filter toolbar, dashboard
- `JavaScript_core.html` — Unified search filter logic
- `JavaScript_ui.html` — Dashboard section rendering
- `showSupportView.html` — Section-based support rendering
- `support_client.html` — Simplified action handlers
- `support_inline_quantity.html` — Unified prefix fallback
- `ViewToggle.html` — Dense table default, status column
- `REVIEW_FIX_ORDER_PLAN.md` — Support mode stability fixes

## Assumptions

- Backend GAS contracts stay intact in this round
- No migration to Node.js
- First delivery prioritizes shell, IA, and working surface
- Important modal flows stay modal for now
