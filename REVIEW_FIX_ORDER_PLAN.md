# Review Fix Order Plan

**Project:** budg69

This document turns the latest code review into a recommended fix order for the implementation agent.

---

## Summary

The main remaining risk after the refactor is not the backend split itself. It is the shared frontend state and compatibility behavior around support mode.

Fix priority should follow this order:

1. restore correct main state after leaving support
2. stop corrupting browser history during support transitions
3. unify support ID normalization across all entry paths
4. run focused regression around support/main switching

---

## Fix Order

### Fix 1: Restore Main State After Leaving Support

**Priority:** P1

**Problem**

Support loading currently overwrites shared client state:

- `State.allItems`
- `State.filteredItems`

When support view closes, the DOM is restored, but the main budget state is not restored or reloaded. This can break:

- expense open
- history open
- later filter behavior
- any lookup path that depends on `State.allItems`

**Target**

When leaving support mode, the app must end in one of these safe states:

- restore the previous main dataset from a preserved snapshot, or
- explicitly reload main data before allowing main interactions again

**Recommended Implementation Rule**

- Do not let support-mode data permanently replace main-mode state
- Keep support-specific data separate from main data, or snapshot/restore main state on mode transition
- `closeSupportView()` must restore both:
  - layout
  - app data/state

**Acceptance**

- open main page
- enter support
- load support items
- leave support
- open a main item expense modal successfully
- open a main item history modal successfully
- filters still operate on main items, not support items

---

### Fix 2: Make Support History Navigation Idempotent

**Priority:** P1

**Problem**

`showSupportView()` always calls `history.pushState(...)`, even when:

- restoring support view from `popstate`
- opening support because the URL already contains `?page=support`

This can create duplicate history entries and broken Back/Forward behavior.

**Target**

Support navigation should distinguish between:

- user-initiated navigation into support
- restoring an existing support state

**Recommended Implementation Rule**

- only push a new history entry on explicit user navigation into support
- do not push when:
  - handling `popstate`
  - hydrating initial load from `?page=support`
- keep one clear mapping between:
  - main view
  - support view
  - browser history state

**Acceptance**

- from main page, open support once
- browser history gains one support entry, not multiple
- Back returns to main view cleanly
- Forward reopens support cleanly
- loading directly with `?page=support` does not create extra phantom history steps

---

### Fix 3: Unify Support Prefix And Item ID Normalization

**Priority:** P2

**Problem**

Support compatibility logic is inconsistent:

- `SupportIndex.html` sets support prefix to `SP69`
- main `Index.html` does not set that prefix
- support client fallback normalization can default to `BG69`
- support API routing in main code depends on `_supportDefaultPrefix` or `SP...`

This creates a risk that support items are normalized differently depending on entry path.

**Target**

There must be exactly one canonical support item prefix rule used by:

- support page
- main page support mode
- client-side normalization
- support API routing checks

**Recommended Implementation Rule**

- define the support prefix once and load it in every path that can render support items
- do not let support normalization fall back to main-budget prefix
- keep support ID normalization and support API detection aligned

**Acceptance**

- support item IDs resolve consistently from:
  - main shell support path
  - `?page=support`
  - button actions
  - history actions
  - expense save path
- no support item is routed through the main-budget API by mistake

---

### Fix 4: Regression Pass Around Main/Support Boundary

**Priority:** P2

**Problem**

The highest-risk flows now sit at the boundary between:

- main budget mode
- support mode
- browser history
- shared state

**Target**

Add a focused regression pass after the fixes above before touching lower-priority cleanup.

**Required Test Scenarios**

- main load
- support open from main
- support load
- support history open
- support expense open/save
- leave support back to main
- main expense open after returning
- main history open after returning
- Back/Forward behavior with support route
- direct load via `?page=support`

**Acceptance**

- no stale support state leaks into main mode
- no duplicate history entries
- no ID normalization mismatch between routes

---

### Fix 5: Secondary Cleanup After Functional Stabilization

**Priority:** P3

**Problem**

There is still layered compatibility behavior around support actions:

- global safe event interception
- support container delegation
- inline onclick fallback
- history-specific MutationObserver patch

This is not the first bug to fix, but it remains a maintenance risk.

**Target**

After the functional fixes above, reduce duplicate action-routing layers where possible.

**Recommended Implementation Rule**

- keep one primary routing path for support actions
- keep fallback behavior only where there is a proven runtime need
- do not stack new handlers on top of the current compatibility layers

**Acceptance**

- support record/history actions still work
- no obvious duplicate handler behavior
- fewer overlapping routing layers than before

---

## Recommended Execution Sequence

1. Fix main-state restoration after leaving support
2. Fix support history/state transition behavior in browser history
3. Unify support prefix and item normalization across entry paths
4. Run targeted regression on support/main transitions
5. Only then clean up overlapping support action handlers

---

## Review Gate

Do not accept the next implementation round until all of these are true:

- [x] leaving support restores a correct main working state
- [x] Back/Forward behavior is stable
- [x] support ID normalization is identical across all routes
- [x] main actions still work after a support round-trip
