# F5 Manual Recovery Follow-Up — 2026-03-12

Purpose: verify the live F5 manual recovery flow against the next real morning print batch.

## What To Check
- Run `Build F5 Manual Recovery` after the first real F5 issue or controlled test case.
- Confirm the sheet shows only actionable lines in `Manual Lines To Deduct`.
- Confirm `Hidden Review` is `0` for the clean case. If not, capture the exact order and reason.
- For inferred rows, confirm the SKU/qty matches the visible F5 partial row plus the ShipStation order.
- Manually deduct the listed lines in WASP from `MMH Kelowna / SHOPIFY`.
- Run `Finalize F5 Manual Recovery`.

## Expected Result
- Existing failed/partial F5 rows are repaired visually.
- Orders that never logged visibly get a manual recovery row added.
- `F5 Manual Recovery` hides itself after finalize.
- `F5 Shipment Ledger` and `F5 Recovery State` stay hidden.
- F5/void polling resumes automatically.

## Specific Validation Points
- Check whether `#94265`-style inferred rows behave correctly in live use.
- Confirm no already-deducted orders appear in the manual action list.
- Confirm no double-deduction happens after finalize if F5 later sees the same shipment.
- Confirm the top banner and counts stay understandable for the operator.

## If Something Is Off
- Capture the `F5 Manual Recovery` sheet screenshot.
- Capture the matching `Activity` / `F5 Shipping` rows.
- Note whether the problem is:
  - wrong manual line
  - hidden review false positive
  - finalize did not repair visible rows
  - polling did not resume
