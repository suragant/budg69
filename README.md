# Budget System (GAS Version)

## Structure
- Code.gs → entry point
- Expense.gs → add/edit/cancel logic
- Auth.gs → permission
- Index.html → UI

## Flow
Frontend → google.script.run → GAS → Google Sheet

## Important Rules
- edit = mark old row as EDITED + create new log
- cancel = mark row CANCELLED + refund budget
- history = hide REVERSAL
