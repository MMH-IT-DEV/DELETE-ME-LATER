# Handoff: F4 MO Lot-Tracked Ingredient Deduction Bug

**Date**: 2026-02-23
**Priority**: CRITICAL
**Status**: Unresolved — needs research and fix

---

## Problem Statement

When a Manufacturing Order (MO) is marked "Done" in Katana, the F4 handler (`handleManufacturingOrderDone` in `src/02_Handlers.gs:759`) should remove consumed ingredients from WASP InventoryCloud. **Batch/lot-tracked ingredients consistently fail** with WASP error `-57041` ("Lot is missing. Date Code is missing.") because the handler cannot resolve the batch number.

**Latest failure**: LUP-1 with batch UFC401B assigned in Katana UI — still failed to deduct.

---

## What Was Already Tried (This Session)

### 1. Katana batch_transactions → batch_stocks resolution
- `fetchKatanaMOIngredients(moId)` with `&include=batch_transactions` WORKS — returns `[{batch_id: 3099449, quantity: 1}]`
- But `batch_transactions` only contains `batch_id` + `quantity` — NO `batch_number` field
- `fetchKatanaBatchStock(batchId)` was added to resolve batch_id → batch_number
- **PROBLEM**: Katana PERMANENTLY DELETES batch records when stock reaches zero. `GET /batch_stocks/3099449` returns 404.
- Three fallback approaches tried: direct lookup, `include_deleted=true`, query by batch_id — ALL return 404/empty
- **Conclusion**: Cannot recover batch_number from Katana after the batch is consumed

### 2. WASP lot lookup via advancedinventorysearch
- Old `waspLookupItemLots()` used `infosearch` endpoint — only returns item METADATA (TrackedByLot: true), NO inventory/lot data
- **Rewritten** to use `POST /public-api/ic/item/advancedinventorysearch` with `{SearchPattern: itemNumber, PageSize: 100, PageNumber: 1}`
- **Diagnostic confirmed it works**: For LUP-4 at PRODUCTION, returned `Lot: UFC408A, DateCode: 2029-02-28T00:00:00Z`
- Also added `waspLookupItemLotAndDate()` which returns `{lot, dateCode}` object

### 3. Proactive WASP fallback in F4 handler
- Added code at `02_Handlers.gs:906-928` that checks: if `extractIngredientBatchNumber()` returns null AND `batch_transactions` exist (meaning item IS batch-tracked), proactively call `waspLookupItemLots(ingSku, FLOWS.MO_INGREDIENT_LOCATION)` BEFORE attempting removal
- Should avoid the wasted `-57041` call and get lot from WASP directly

### 4. 15-second delay
- `Utilities.sleep(15000)` at line 819 waits for Katana to propagate batch assignments
- Confirmed: batch_transactions ARE populated (delay is working)

### Result: STILL FAILS
Despite all the above, a new MO test with LUP-1 + batch UFC401B still failed.

---

## Investigation Areas (Priority Order)

### A. Were all 3 modified files deployed to GAS?
The user may have only deployed some files. ALL THREE must be in the live GAS project:
- `04_KatanaAPI.gs` — `fetchKatanaBatchStock()` with 3-tier fallback
- `05_WaspAPI.gs` — `waspLookupItemLots()` rewritten to use `advancedinventorysearch`
- `02_Handlers.gs` — proactive WASP fallback at lines 906-928 + `expiration_date` field fix

**How to check**: In GAS editor, open each file and look for the specific changes. Or run `testWaspLotLookup()` from `09_Utils.gs` — if it still uses `infosearch`, the file wasn't deployed.

### B. Does advancedinventorysearch find LUP-1 at PRODUCTION?
The diagnostic only tested **LUP-4**. The failing item is **LUP-1**. Possible issues:
- LUP-1 may not have stock at PRODUCTION in WASP
- `SearchPattern: "LUP-1"` may return too many results or no exact match (it's a search, not a filter)
- The rewritten `waspLookupItemLots` matches on `rowItem !== itemNumber` — if WASP returns `LUP-1` with different casing or extra chars, it won't match

**Action**: Run `testWaspLotLookup()` but change the test item from LUP-4 to LUP-1 and check the raw response.

### C. Check Debug/Activity logs for diagnostic entries
The proactive fallback logs two event types:
- `F4_BATCH_WASP_FALLBACK` — if WASP lot lookup SUCCEEDS (means fix is working)
- `F4_BATCH_DIAG` — if BOTH Katana and WASP lot lookups FAIL

Also check for:
- `F4_INGREDIENT_FAIL` — logged when removeResult fails (line 971-975)
- `WASP_LOT_LOOKUP_FAIL` or `WASP_LOT_LOOKUP_ERROR` — from waspLookupItemLots itself

If NONE of these appear in the Debug tab, it means the updated code is NOT running (deployment issue).

### D. Check the exact WASP error on the failing ingredient
The Activity tab / F4 flow tab should show the error for LUP-1. Possible errors:
- `-57041` = "Lot is missing" — WASP lot lookup didn't fire or returned null
- `-46002` = "Insufficient quantity" — lot was found but WASP doesn't have enough stock
- `-57010` = "Item not found" — LUP-1 doesn't exist in WASP

### E. FLOWS.MO_INGREDIENT_LOCATION value
The proactive lookup passes `FLOWS.MO_INGREDIENT_LOCATION` which maps to `LOCATIONS.PRODUCTION` = `'PRODUCTION'`. Verify this matches the actual WASP location where LUP-1 stock lives. If stock is at a different location (e.g., RECEIVING-DOCK, UNSORTED), the lookup will find nothing.

### F. advancedinventorysearch SearchPattern behavior
The endpoint is a SEARCH not an exact filter. `SearchPattern: "LUP-1"` may return:
- LUP-1, LUP-10, LUP-100, LUP-1-SPECIAL, etc.
- The code filters with `rowItem !== itemNumber` after search — but if the WASP response field names are different than expected, the filter may fail

Relevant code in `05_WaspAPI.gs:174-188`:
```javascript
var rowItem = row.ItemNumber || row.itemNumber || '';
if (rowItem !== itemNumber) continue;
var rowLoc = row.LocationCode || row.locationCode || '';
if (locationCode && rowLoc !== locationCode) continue;
var rowLot = row.Lot || row.lot || row.LotNumber || '';
if (rowLot) return rowLot;
```

### G. Could be a code path issue
The proactive fallback at line 906-928 checks `if (!ingLot)` and then `if (bt.length > 0)`. But:
- What if `extractIngredientBatchNumber()` returns an EMPTY STRING `''` instead of `null`?
  - `''` is falsy in JS so `!ingLot` would be `true` — this should work
- What if `ing.batch_transactions` uses a different key name in the actual API response?
  - Code checks `ing.batch_transactions || ing.batchTransactions || []`
  - But the diagnostic confirmed `batch_transactions` is the correct key

### H. Multiple lots at same location
`waspLookupItemLots()` returns the FIRST lot found at the location. If LUP-1 has multiple lots at PRODUCTION with different batch numbers, WASP may reject the removal if the wrong lot is picked. Consider:
- The lot returned may not have sufficient quantity
- The lot returned may not match what Katana consumed

---

## Key Files & Functions

| File | Function | Line | Purpose |
|------|----------|------|---------|
| `02_Handlers.gs` | `handleManufacturingOrderDone()` | 759 | Main F4 handler |
| `02_Handlers.gs` | `extractIngredientBatchNumber()` | 542 | Extract lot from Katana data |
| `02_Handlers.gs` | `extractIngredientExpiryDate()` | 593 | Extract expiry from Katana data |
| `02_Handlers.gs` | Proactive WASP fallback | 906-928 | WASP lot lookup when Katana fails |
| `02_Handlers.gs` | -57041 retry fallback | 954-967 | Second WASP lot lookup after failed remove |
| `04_KatanaAPI.gs` | `fetchKatanaBatchStock()` | 86 | 3-tier Katana batch resolution |
| `04_KatanaAPI.gs` | `fetchKatanaMOIngredients()` | 115 | Fetches recipe rows with batch_transactions |
| `05_WaspAPI.gs` | `waspLookupItemLots()` | 154 | WASP lot lookup via advancedinventorysearch |
| `05_WaspAPI.gs` | `waspLookupItemLotAndDate()` | 202 | Returns lot + dateCode from WASP |
| `05_WaspAPI.gs` | `waspRemoveInventoryWithLot()` | 129 | Remove with lot/dateCode |
| `09_Utils.gs` | `testMOBatchData()` | (end) | Diagnostic: inspect MO batch data |
| `09_Utils.gs` | `testWaspLotLookup()` | (end) | Diagnostic: test WASP lot lookup |
| `00_Config.gs` | `FLOWS.MO_INGREDIENT_LOCATION` | 104 | = 'PRODUCTION' |

---

## F4 Handler Flow (Ingredient Removal)

```
1. Fetch MO header → get moId, moOrderNo, outputSku
2. Utilities.sleep(15000) — wait for Katana to propagate batch assignments
3. fetchKatanaMOIngredients(moId) — includes batch_transactions
4. For each ingredient:
   a. fetchKatanaVariant(ing.variant_id) → ingSku
   b. extractIngredientBatchNumber(ing) → ingLot
      - Checks batch_transactions[0].batch_number (empty — Katana only has batch_id+qty)
      - Checks batch_stock embedded object (empty)
      - Resolves batch_id via fetchKatanaBatchStock() (404 — depleted)
      - Returns null
   c. Proactive WASP fallback (lines 906-928):
      - If ingLot is null AND batch_transactions exist:
        - waspLookupItemLots(ingSku, 'PRODUCTION') → should return lot from advancedinventorysearch
        - If found: ingLot = waspLot (SUCCESS PATH)
        - If not found: logs F4_BATCH_DIAG (FAILURE — needs investigation)
   d. If ingLot exists:
      - waspRemoveInventoryWithLot(ingSku, qty, 'PRODUCTION', ingLot, notes)
   e. If ingLot is null:
      - waspRemoveInventory(ingSku, qty, 'PRODUCTION', notes) → fails with -57041
      - -57041 retry: waspLookupItemLots(ingSku, 'PRODUCTION') → retry with lot
```

---

## Diagnostic Functions Available

### testMOBatchData() in 09_Utils.gs
- Scans recent MOs, finds one by order_no prefix
- Fetches recipe rows WITH and WITHOUT `&include=batch_transactions`
- For each batch_id found, resolves via `fetchKatanaBatchStock()`
- Also queries `batch_stocks?variant_id=X` directly
- Run from GAS editor: select function dropdown → testMOBatchData → Run → View → Logs

### testWaspLotLookup() in 09_Utils.gs
- Tests `advancedinventorysearch` for a hardcoded item (currently LUP-4 at PRODUCTION)
- **Change the test item to LUP-1** to test the actual failing SKU
- Shows raw API response, matched rows, lot found

---

## Confirmed Facts

1. Katana `batch_transactions` on recipe rows contains `{batch_id, quantity}` — NO `batch_number`
2. Katana permanently deletes batch records when stock = 0. No `include_deleted` recovery.
3. `manufacturing_order.done` fires AFTER user assigns batch in Katana popup — batch_transactions IS populated
4. WASP `advancedinventorysearch` returns inventory rows with `Lot`, `DateCode`, `ItemNumber`, `LocationCode`
5. WASP `infosearch` only returns item catalog metadata — NO inventory/lot data
6. Katana field name for expiry: `expiration_date` (NOT `expiry_date` or `best_before_date`)
7. Katana API pagination uses `limit` param (NOT `per_page`), max 250
8. WASP error `-57041` = lot-tracked item missing Lot/DateCode in API call
9. WASP error `-46002` = insufficient quantity (stock doesn't exist at that location)

---

## Deployment Details

- GAS Script ID: `1vMGYLeN5iCcLzW9mF0DJB3bT-My_yj1Tv8MVVnaVWKLb4m582ZqyT4Lg`
- Deploy command: `node scripts/push-to-apps-script.js` from project root
- After pushing: GAS Editor → Deploy → Manage deployments → Edit → Version: latest → Save
- Active webhook URL: `https://script.google.com/macros/s/AKfycbzSojwLpeEo6irjUpCoUnRB9f_n8NW3vUWkShU8zt6kiHffbKkeJVxXNdspA0N28tv75w/exec`
- Debug sheet: `1eX7MCU-Is5CMmROL1PfuhGoB73yRF7dYdyXHqzMYOUQ`

---

## Recommended Fix Strategy

1. **First**: Verify deployment — confirm all 3 files are live in GAS
2. **Second**: Run `testWaspLotLookup()` with LUP-1 to confirm advancedinventorysearch finds it
3. **Third**: Check Debug tab for `F4_BATCH_WASP_FALLBACK` or `F4_BATCH_DIAG` entries from the latest MO test
4. **Fourth**: If WASP lookup works for LUP-1 but F4 still fails, add more granular logging at each step of the fallback chain to pinpoint where the data drops
5. **Fifth**: Consider that `waspRemoveInventoryWithLot` may also need `dateCode` — currently the proactive fallback gets lot but NOT dateCode. Use `waspLookupItemLotAndDate()` instead of `waspLookupItemLots()` to get both values, and pass dateCode to the remove call.

### Likely Missing Piece: DateCode
Looking at `waspRemoveInventoryWithLot()` in `05_WaspAPI.gs:129-147`:
```javascript
function waspRemoveInventoryWithLot(itemNumber, quantity, locationCode, lotNumber, notes, siteName, dateCode) {
  ...
  var payload = [{
    ItemNumber: itemNumber,
    Quantity: quantity,
    SiteName: siteName || CONFIG.WASP_SITE,
    LocationCode: locationCode,
    Lot: lotNumber || '',
    Notes: notes || ''
  }];
  // WASP requires DateCode for lot-tracked items
  if (dateCode) {
    payload[0].DateCode = dateCode;
  }
  ...
}
```

The proactive fallback at line 911 calls `waspLookupItemLots()` which returns ONLY the lot string — no dateCode. But WASP may REQUIRE DateCode in the remove payload. The function signature accepts `dateCode` as 7th parameter, but the fallback doesn't pass it.

**FIX**: Change the fallback to use `waspLookupItemLotAndDate()` and pass both lot AND dateCode to `waspRemoveInventoryWithLot()`.
