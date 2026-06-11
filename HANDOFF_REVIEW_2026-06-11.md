# Handoff Review - 2026-06-11

## Context

This project uses Google Apps Script for the backend and HTML plus vanilla JS for the frontend.

- Backend: `.gs` files using `SpreadsheetApp` as the database
- Frontend: HTML + vanilla JS using `google.script.run`
- Important constraint: keep GAS style and do not convert to Node.js

I did not modify application code during this review.

Existing user changes already present in the working tree:

- [README.md](C:/Users/Surag/Projects/01-products/budg69/README.md)
- [REFACTOR_PLAN.md](C:/Users/Surag/Projects/01-products/budg69/REFACTOR_PLAN.md)

## Priority Findings

### 1. `reverseTransfer` allows unauthorized or repeated reversals

File:

- [transaction_service.gs](C:/Users/Surag/Projects/01-products/budg69/transaction_service.gs:528)

Problem:

- The function checks that the log entry is a transfer, but it does not verify that the caller is the original actor or an admin.
- It also does not block reversal of a transfer row that has already been cancelled.

Impact:

- Any authenticated caller who can invoke the server method may be able to reverse another user's transfer.
- The same transfer may be reversed multiple times, causing budget corruption.

Recommended fix:

- Reject when `logEntry.status === 'CANCELLED'`
- Apply owner/admin authorization checks similar to `cancelExpense` and `editExpense`
- Ensure sibling transfer rows are not already cancelled before applying reversal

### 2. `recordSupportExpenseSupport` has no server-side access control

File:

- [support_module.gs](C:/Users/Surag/Projects/01-products/budg69/support_module.gs:151)

Problem:

- Unlike `recordExpense` and `transferBudget`, this function does not call `getUserPermission()` or validate role/department access before writing.

Impact:

- A caller may be able to write support expenses without proper authorization.

Recommended fix:

- Add user lookup at the top of the function
- Reject missing users and `viewer` users
- Validate access to the target row's department with `hasAccessToRow()`

### 3. `recordSupportExpenseSupport` writes quantity before validating the budget

File:

- [support_module.gs](C:/Users/Surag/Projects/01-products/budg69/support_module.gs:179)

Problem:

- `usedQty` is updated before checking whether the monetary amount would exceed budget.
- If budget validation fails, the function returns an error but the quantity update has already been committed.

Impact:

- Quantity and money fields can diverge, leaving the sheet in a partially updated state.

Recommended fix:

- Compute all new values first
- Validate both quantity and monetary updates before any write
- Write all related fields only after every validation passes

### 4. Sensitive read methods are missing server-side authorization filters

Files:

- [support_module.gs](C:/Users/Surag/Projects/01-products/budg69/support_module.gs:225)
- [transaction_service.gs](C:/Users/Surag/Projects/01-products/budg69/transaction_service.gs:235)

Problem:

- `getSupportQuarterlyReport` does not check the caller's permission or scope the output by department.
- `getTransactionHistory` returns history for any `itemId` without verifying whether the caller should see that item's department.

Impact:

- Users may be able to read cross-department data by invoking server methods directly.

Recommended fix:

- Require `getUserPermission()` in both methods
- Resolve the target row department and gate with `hasAccessToRow()`
- Filter quarterly report aggregates to only rows the current user is allowed to access

### 5. Support expense cancellation reverses money but not quantity

File:

- [transaction_service.gs](C:/Users/Surag/Projects/01-products/budg69/transaction_service.gs:332)

Problem:

- When cancelling a support expense, the code reverses `usedMoney` and `remainingMoney`.
- It does not reverse the logged quantity even though support entries record quantity in the transaction log.

Impact:

- Quantity totals remain inflated after a cancellation.

Recommended fix:

- Read the original transaction quantity from the log entry
- Subtract that quantity from the support row during cancellation
- Apply the same consistency rules to edit flows if quantity changes are supported later

## Suggested Next Steps

1. Add permission checks to every write path first.
2. Fix partial-write behavior in `recordSupportExpenseSupport`.
3. Harden cancellation and reversal flows against duplicate actions.
4. Add server-side scoping to read endpoints.
5. Run manual regression tests on support and transfer flows.

## AI Agent Backlog

Use this backlog in order. Each item is sized so another agent can pick it up and complete it with a small, reviewable patch.

### Backlog 1. Harden `recordSupportExpenseSupport`

Goal:

- Prevent unauthorized writes and partial updates in the support expense flow.

Scope:

- [support_module.gs](C:/Users/Surag/Projects/01-products/budg69/support_module.gs:151)

Tasks:

1. Add `getUserPermission()` at the start of the function.
2. Reject users with no permission record.
3. Reject `viewer` users.
4. Resolve the target row first, then read its department.
5. Enforce `hasAccessToRow(currentUser, department)` before any write.
6. Compute `newUsedQty`, `newUsedMoney`, and `newRemainingMoney` before writing.
7. If any validation fails, return without changing the sheet.
8. Only write quantity and money fields after all validations pass.

Acceptance criteria:

- Unauthorized users cannot record support expenses.
- A budget validation failure leaves both quantity and money unchanged.
- Successful writes still append a transaction log row.

### Backlog 2. Fix support cancellation consistency

Goal:

- Make support cancellation reverse both money and quantity.

Scope:

- [transaction_service.gs](C:/Users/Surag/Projects/01-products/budg69/transaction_service.gs:299)

Tasks:

1. In the support branch of `cancelExpense`, read `logEntry.quantity`.
2. Subtract the quantity from the support row's used quantity column.
3. Prevent quantity from going negative.
4. Keep quantity and monetary reversal inside the same guarded flow.
5. Confirm the transaction log status update still happens only after successful reversal.

Acceptance criteria:

- Cancelling a support expense restores both `usedQty` and money totals.
- If reversal would make quantity negative, the function rejects safely.

### Backlog 3. Lock down `reverseTransfer`

Goal:

- Prevent unauthorized and repeated reversal of transfer records.

Scope:

- [transaction_service.gs](C:/Users/Surag/Projects/01-products/budg69/transaction_service.gs:528)

Tasks:

1. Reject rows whose `status` is already `CANCELLED`.
2. Add owner/admin authorization logic similar to `cancelExpense`.
3. Verify sibling transfer rows are not already cancelled before applying reversal.
4. Ensure a reversal cannot be applied twice from either transfer row.
5. Keep existing budget integrity checks for negative destination budget/remaining.

Acceptance criteria:

- Only the original user or an admin can reverse a transfer.
- A cancelled transfer cannot be reversed again.
- Reversal from one transfer row blocks reversal from its sibling row as well.

### Backlog 4. Add server-side authorization to read methods

Goal:

- Prevent cross-department data access through direct server method calls.

Scope:

- [transaction_service.gs](C:/Users/Surag/Projects/01-products/budg69/transaction_service.gs:235)
- [support_module.gs](C:/Users/Surag/Projects/01-products/budg69/support_module.gs:225)

Tasks:

1. Add `getUserPermission()` to `getTransactionHistory`.
2. Resolve the item's department before returning history.
3. Reject or return empty when the caller lacks access.
4. Add `getUserPermission()` to `getSupportQuarterlyReport`.
5. Filter `itemsMap`, `byArea`, and `byExpenseType` to only authorized rows/items.
6. Decide whether unauthorized access should return an error or an empty response and keep that behavior consistent.

Acceptance criteria:

- Users cannot fetch another department's transaction history by passing a raw `itemId`.
- Quarterly support reporting only includes rows visible to the caller.

### Backlog 5. Review related edit flows for support parity

Goal:

- Check whether support-related edit flows have the same consistency gaps as cancellation.

Scope:

- [transaction_service.gs](C:/Users/Surag/Projects/01-products/budg69/transaction_service.gs:410)

Tasks:

1. Review `editExpense` support logic for quantity consistency.
2. Confirm whether support edits are supposed to allow quantity changes.
3. If quantity is intentionally immutable, document that assumption in code comments or notes.
4. If quantity should be editable, create a follow-up patch plan before implementing.

Acceptance criteria:

- The team has an explicit decision on support quantity behavior during edits.
- No hidden inconsistency remains between support create, edit, and cancel flows.

### Backlog 6. Manual regression pass

Goal:

- Verify the patched flows from an end-user perspective.

Scope:

- Support expense flow
- Budget expense flow
- Transfer and reverse transfer flow
- Transaction history visibility

Tasks:

1. Test support write success.
2. Test support write failure due to budget overflow.
3. Test support cancellation.
4. Test normal budget cancellation.
5. Test transfer reversal once.
6. Test transfer reversal twice.
7. Test cross-department access with a non-admin user.

Acceptance criteria:

- All guarded flows reject safely.
- No sheet field changes occur on failed operations.
- Authorized operations still work as expected.

## Manual Regression Checklist

1. Record a normal support expense and verify quantity, used, and remaining all update correctly.
2. Attempt a support expense that exceeds budget and verify nothing is written.
3. Cancel a support expense and verify both money and quantity are reversed.
4. Try reversing the same transfer twice and verify the second attempt is rejected.
5. Try reversing or viewing data as a user from another department and verify access is denied.

## Assumptions

- The web app is published as a standard Apps Script web app.
- Server functions can be invoked by client code by name, so server-side authorization is required even if the UI hides actions.
