# Agent Briefing: Update WASP Item Descriptions from Katana Names

## Objective

WASP InventoryCloud items are missing descriptions. The `ItemDescription` field in WASP
needs to be populated with the product/material **Name** from Katana MRP for every item
that exists in both systems.

## Data Source

**Katana inventory export** (contains the source names):
`C:\Users\Admin\Downloads\InventoryItems-2026-03-19-13_00.xlsx`

Columns: Name | Variant code / SKU | Category | Default supplier | Units of measure |
Default storage bin | Average cost | Value in stock | In stock | Expected | Committed |
Safety stock | Calculated stock | Location

- **305 items** total
- **SKU** is in column "Variant code / SKU" (e.g., `NI-CUTTER`, `NI-PALLERT JACK`)
- **Name** is in column "Name" (e.g., `EDGE PROTECTOR CUTTER`, `INDUSTRIAL PALLET TRUCK`)
- The Name column is what needs to go into WASP's `ItemDescription` field

## Codebase Location

The sync script is the right place to work from. It already has WASP and Katana API access.

**Sync script files**: `C:\Users\Admin\Documents\claude-projects\latest-wasp-katana update\sync\script\`

Key files:
- `03_WaspAPI.js` — WASP API functions (`waspFetch`, `getWaspBase`, `getStoredWaspToken_`)
- `04_SyncEngine.js` — `waspCreateItem()` function (creates items with description)
- `05_PushEngine.js` — Push engine (processes Katana→WASP syncs)
- `02_KatanaAPI.js` — Katana API functions
- `test.js` — One-time test/utility functions (good place for the update script)
- `.clasp.json` — Script ID for clasp push

## WASP API Reference

Full reference at: `C:\Users\Admin\Documents\claude-projects\latest-wasp-katana update\docs\api\wasp-api-reference.md`

### Key endpoints for this task:

**Search for items (to get current data):**
```
POST /public-api/ic/item/advancedinfosearch
Payload: { SearchPattern: "SKU", PageSize: 100, PageNumber: 1 }
```
Returns: `{ Data: [{ ItemNumber, ItemDescription, StockingUnit, Cost, ... }] }`
NOTE: SearchPattern is substring match — filter results by exact ItemNumber.

**Also available:**
```
POST /public-api/ic/item/infosearch
Payload: { ItemNumber: "SKU", AltItemNumber: "", ItemDescription: "" }
```
Returns more fields including TrackbyInfo, DimensionInfo, CustomFields.

**Update items:**
```
POST /public-api/ic/item/updateInventoryItems
Payload: [{ ItemNumber: "SKU", StockingUnit: "EA", PurchaseUnit: "EA", SalesUnit: "EA",
  ItemDescription: "NEW DESCRIPTION",
  DimensionInfo: { DimensionUnit: "Inch", WeightUnit: "Pound", VolumeUnit: "Cubic Inch", ... }
}]
```

### CRITICAL: updateInventoryItems behavior

**Known issue**: This endpoint returns `SuccessfullResults: 0` in many cases. We tested it
extensively and it consistently returned 0 results for item updates (lot tracking, cost).
It may not work for description updates either.

**Possible workaround**: The `createInventoryItems` endpoint successfully sets `ItemDescription`
when creating new items. For EXISTING items, it returns an "already exists" error. But the
description might NOT be updated on existing items via create.

**Recommended approach**:
1. First try `updateInventoryItems` with a single test item to see if descriptions can be updated
2. The payload MUST include `StockingUnit`, `PurchaseUnit`, `SalesUnit`, AND `DimensionInfo`
   (WASP validates these even on update — missing fields cause `-57066` error)
3. If update works → batch process all 305 items
4. If update doesn't work → investigate alternative approaches:
   - Check if WASP has an import/CSV upload feature for item data
   - Contact WASP support for the correct update endpoint behavior
   - Use the WASP UI to update descriptions manually (last resort)

### Authentication

```javascript
var token = getStoredWaspToken_();  // from ScriptProperties WASP_API_TOKEN
var base = getWaspBase();            // e.g., https://mymagichealer.waspinventorycloud.com/public-api/
```

All requests need:
```
Headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }
Method: POST
muteHttpExceptions: true
```

## Implementation Plan

### Step 1: Build SKU → Name mapping from Katana

Two options:
a) Read from the Excel file (requires parsing in GAS — possible with Sheets import)
b) Use the Katana API directly (`fetchAllKatanaData()` already returns items with names)

Option (b) is easier since the sync script already has `fetchAllKatanaData()` which returns:
```javascript
{ items: [{ sku: "NI-CUTTER", name: "EDGE PROTECTOR CUTTER", ... }] }
```

### Step 2: Fetch current WASP item data

Use `advancedinfosearch` to get all WASP items with their current descriptions.
Compare against Katana names to find items needing updates.

### Step 3: Update descriptions

For each item where WASP ItemDescription is empty or different from Katana Name:
- Call `updateInventoryItems` with the new description
- Include all required fields (UOM, DimensionInfo) from the existing item data
- Log success/failure per item

### Step 4: Rate limiting

WASP doesn't publish rate limits, but the codebase uses 300ms delay between calls.
For 305 items: ~90 seconds total. Well within GAS 6-minute execution limit.

## GAS Compatibility Rules

- **var only** — no const, no let
- **No arrow functions** — use `function(x) { }` not `(x) => {}`
- **No template literals** — use string concatenation
- All code must be Google Apps Script V8 compatible

## Deployment

After writing the code:
```bash
cd "C:\Users\Admin\Documents\claude-projects\latest-wasp-katana update\sync\script"
clasp push
```

Then run the function from the GAS editor (sync script):
`https://script.google.com/u/0/home/projects/1xRjhF03JdKNxn-3wcqUyOT39HunMgRhEGZzxvzug9_lDCot6vpn_MEeE/edit`

## Testing

1. Pick ONE test item (e.g., `NI-CUTTER` → description should be `EDGE PROTECTOR CUTTER`)
2. Try `updateInventoryItems` with just that one item
3. Check the response — does it show `SuccessfullResults: 1`?
4. Verify in WASP UI that the description was updated
5. If it works → run the full batch
6. If not → report findings and try alternatives

## What NOT to do

- Do NOT modify the sync engine or push engine code
- Do NOT change any existing API functions
- Write all code in `test.js` as a standalone function
- Do NOT create new items — only UPDATE existing items
- Clean up temporary functions after the task is complete
