# Katana-WASP Inventory Sync

## What
Apps Script integration syncing Katana MRP inventory events to WASP InventoryCloud. Triggered by Katana webhooks and WASP callouts, processes 5 flows (product sync, stock adjustment, auto-poll, PO batch, transfers).

## Tech Stack
- **Runtime**: Google Apps Script (V8) — MUST use `var` only, no `const`/`let`/`for-of`
- **Source of Truth**: Katana MRP API
- **Target**: WASP InventoryCloud API
- **Logging**: Google Sheets tracker (Debug=errors, Activity=F2 results) + Slack notifications
- **Testing**: Python test suite (pytest)

## Code Locations
- **Apps Script (local)**: `C:\Users\Admin\Documents\claude-projects\wasp-katana\src\` (11 canonical .gs files)
- **Apps Script (remote)**: scriptId `1vMGYLeN5iCcLzW9mF0DJB3bT-My_yj1Tv8MVVnaVWKLb4m582ZqyT4Lg`
- **Python tests**: `C:\Users\Admin\Downloads\wasp-katana\`
- **Build plan**: `./build-plan.md`
- **Tracker sheet**: WASP-Katana-tracker (Google Sheets)
- **Command Center**: Sheet `1ZpwOKBJ1brRWVDG2hb9ZZSGjdsBv4c-623LwDt_zLRs`
- **Old files** (archived): `C:\Users\Admin\Downloads\wasp-katana-scripts\` (includes _v2/_updated/_COMPLETE variants — do not use)

## Current Phase
Active Development — F2 deployed and working. Log cleanup done. Pending: WASP callout Lot/DateCode fix, full deploy of all files.

## Flows
| Flow | Name | Status |
|------|------|--------|
| F1 | Product Sync (Katana → WASP) | Working |
| F2 | Stock Adjustment (WASP → Katana) | Working (remove OK, add needs callout Lot fix) |
| F3 | Auto Poll (Katana ST → WASP) | Working (B-WAX qty error = data issue) |
| F4 | PO Batch | Working |
| F5 | Transfer / ShipStation | Working |

## Key Files
| File | Purpose |
|------|---------|
| 00_Config.gs | Configuration, API keys, site mapping |
| 01_Main.gs | Entry point, event routing |
| 02_Handlers.gs | Flow handler dispatching |
| 03_WaspCallouts.gs | F2: WASP callout handlers, batch grouping, lot/expiry |
| 04_KatanaAPI.gs | Katana API wrapper functions |
| 05_WaspAPI.gs | WASP API wrapper functions |
| 06_ShipStation.gs | ShipStation integration |
| 07_PickMappings.gs | Pick order → Katana SO mapping |
| 07_F3_AutoPoll.gs | F3: Stock transfer polling + WASP sync |
| 08_Logging.gs | logToSheet(), logF2Activity() |
| 09_Utils.gs | Shared utilities |

## Critical API Notes
- WASP write endpoints require **ARRAY** payload `[{...}]` not object `{...}`
- Katana batch tracking: `batch_stocks?batch_id=` pattern
- WASP `DateCode` field = expiry date
- WASP callout MUST include `Lot` and `DateCode` for correct lot attribution
- Katana requires `batch_tracking` enabled on product for lot numbers
- ALL files must use `var` — `const`/`let` causes silent file load failure
