# Agent Handoff — Katana Tab WASP Sync Control

**Date**: 2026-02-20
**Active Problem**: WASP `ic/item/create` endpoint returning 404 errors

---

## Project Overview

Google Apps Script system that syncs inventory between **Katana MRP** and **WASP InventoryCloud** via a Google Sheets control panel. The spreadsheet has 4 tabs: Katana, Wasp, Batch, Sync History.

**Key files** (all in `sync-engine/`):
- `01_Code.gs` — Menu, triggers, onEdit handler, runFullSync orchestrator
- `02_KatanaAPI.gs` — Katana API data fetching (read-only, not modified recently)
- `03_WaspAPI.gs` — WASP API wrappers (waspApiCall, getWaspBase, add/remove/adjust)
- `04_SyncEngine.gs` — waspCreateItem, autoCreateMissingItems (disabled), logSyncHistory
- `04b_RawTabs.gs` — writeKatanaTab, writeWaspTab, writeBatchTab, writeSkuLookup
- `05_PushEngine.gs` — pushToWasp, pushKatanaSyncRows_, pushBatchPendingRows_

**Python test scripts** (in `tests/`):
- `sync_all_v5.py` — The script that successfully created 266/275 items in WASP initially

---

## What We're Building

A "WASP Status" column (col J) on the Katana tab that lets users manually control which Katana items get created in WASP. Replaces the old auto-create behavior.

### Workflow
1. **Sync Now** → Katana tab shows all items with statuses:
   - `Synced` (green) — item exists in WASP
   - `NEW` (orange) — item not in WASP yet
2. User clicks **dropdown** in col J → selects **"Push"** → row turns yellow
3. User clicks **Push to WASP** → item created in WASP + stock added → status becomes `Synced`
4. If push fails → `ERROR` (red) → user can select Push again to retry

### Katana Tab Layout (10 columns)
```
A: SKU | B: Name | C: Location | D: Type | E: Lot Tracked | F: Batch | G: Expiry | H: Qty | I: Match | J: WASP Status
```

---

## User Decisions (confirmed)

| Decision | Answer |
|----------|--------|
| Dropdown options | Push (single user action) + system statuses (NEW, Synced, ERROR) |
| Default for new items | NEW (orange) |
| After successful push | "Synced" (no timestamp) |
| Error retry | User can select Push again on ERROR rows |
| WASP Category/Taxable | Don't send (skip) |
| Location routing | Smart: Materials→PRODUCTION, Products→UNSORTED, Storage WH→SW-STORAGE |
| Batch push | Per SKU (one Push on any row pushes all batches for that SKU) |
| Storage Warehouse | Route to WASP site "Storage Warehouse" / location "SW-STORAGE" |
| Push scope | All tabs at once (single Push to WASP button) |
| 0-qty items | Show them, allow pushing (creates item, skips stock ADD) |

---

## ACTIVE BUG: WASP ic/item/create Returning 404

### Symptom
Every `waspCreateItem()` call returns:
```
"The resource you are looking for has been removed, had its name changed, or is temporarily unavailable."
```
This is a standard IIS 404 error page, NOT a WASP business error.

### What We Know
- The GAS function calls: `getWaspBase() + 'ic/item/create'`
- `getWaspBase()` returns: `'https://' + instance + '.waspinventorycloud.com/public-api/'`
- So the URL is: `https://mymagichealer.waspinventorycloud.com/public-api/ic/item/create`

### CORRECTED ANALYSIS (Feb 20): Both URLs Are Identical
```python
# Python sync_all_v5.py line 52:
url = f"{self.base_url}/public-api/{endpoint}"
# Full URL: https://mymagichealer.waspinventorycloud.com/public-api/ic/item/create
```

```javascript
// GAS 04_SyncEngine.gs:
var url = getWaspBase() + 'ic/item/create';
// Full URL: https://mymagichealer.waspinventorycloud.com/public-api/ic/item/create
```

**CORRECTED**: The previous analysis was WRONG — Python DOES include `/public-api/` in the URL (line 52 of sync_all_v5.py). Both produce the identical URL. The URL is NOT the issue.

### Most Likely Cause: Payload Differences
The 404 is an IIS routing error. Some ASP.NET WebAPI apps return 404 when payload fields don't match the expected model (model binding failure). The GAS payload includes fields Python does NOT send:
- `SiteName`, `LocationCode`, `LotTracking`, `TrackDateCode`, `StockingUnit`

### Diagnostic Test Function Added
`testWaspCreateEndpoint()` added to 04_SyncEngine.gs (and menu). Tests 6 combinations:
1. Python-style minimal payload + /public-api/
2. GAS full payload + /public-api/
3. Bare minimal payload + /public-api/
4. Python-style + NO /public-api/ prefix
5. Array-wrapped payload
6. Known working endpoint (advancedinventorysearch) as control

**Run from GAS editor → View → Logs to see which test returns HTTP 200.**

### Diagnostic Results (Feb 20)
- Tests 1-3 (all payloads with /public-api/): **ALL 404** — payload doesn't matter
- Test 4 (no /public-api/): HTTP 200 but returned HTML error page, not API
- Test 5 (array-wrapped): **404**
- Test 6 (advancedinventorysearch control): **200 JSON** — other endpoints still work

**CONFIRMED: `ic/item/create` endpoint has been REMOVED from WASP API.**
Not a payload issue — bare minimum 3-field payload also 404s. Endpoint worked on Feb 13 (Python sync). WASP changed their API between Feb 13-20.

### Next Step: Endpoint Discovery
`discoverWaspCreateEndpoint()` added — probes 19 candidate endpoint names.
Run from GAS menu → "Discover Wasp Create Endpoint" → View → Logs.

### Fallback Options If No Endpoint Found
1. Contact WASP support (Jason) about ic/item/create removal
2. Use WASP CSV Import (12-Inventory format) as workaround
3. Create items manually in WASP UI, then use API only for stock add/remove

### Current GAS Payload (04_SyncEngine.gs line 165-184)
```javascript
{
    ItemNumber: sku,
    Description: name,           // Changed from ItemDescription
    SiteName: site,
    LocationCode: location,
    ItemType: 'Inventory',       // Changed from 0
    Active: true,                // Added
    LotTracking: true/false,
    TrackDateCode: true,         // Only if lot tracked
    StockingUnit: uom,           // Only if uom truthy
    SalePrice: salesPrice,       // Changed from SalesPrice
    Cost: cost                   // Only if > 0
}
```

### Python Payload (tests/sync_all_v5.py line 193-206)
```python
{
    "ItemNumber": sku,
    "Description": description,
    "Category": "FINISHED GOODS",    # GAS doesn't send this
    "ItemType": "Inventory",
    "Cost": purchase_price,
    "SalePrice": sales_price,
    "Price": sales_price,            # GAS doesn't send this
    "ListPrice": sales_price,        # GAS doesn't send this
    "Active": True,
    "Taxable": True,                 # GAS doesn't send this
    "AltItemNumber": barcode,        # GAS doesn't send this
    "Notes": "Synced from Katana..." # GAS doesn't send this
}
```

**NOTE**: Python does NOT send `SiteName`, `LocationCode`, `LotTracking`, `TrackDateCode`, or `StockingUnit` in the create call. These might be causing the 404 if the endpoint doesn't accept them.

---

## Bugs Fixed in This Session

### 1. Early Return in pushToWasp (FIXED)
`pushToWasp()` returned early with "No pending rows to push" if the Wasp tab had no pending rows, preventing Katana Push rows from ever executing.
- **Fix**: Removed the early return (05_PushEngine.gs ~line 56)

### 2. False "Synced" on 0-qty Items (FIXED)
If CREATE failed but all rows had qty 0, the ADD loop was skipped entirely, `skuSuccess` stayed `true`, and status became "Synced" even though item was never created.
- **Fix**: Added `itemCreated` and `addAttempted` tracking. If CREATE fails and no ADD attempted, marks ERROR. If CREATE fails and has qty rows, skips ADD and marks ERROR.

### 3. "Invalid" Dropdown Warning (FIXED)
Dropdown only had `['Push']` so cells with system values (NEW, Synced, ERROR) showed "Invalid: Input must be an item on the specified list" red warning.
- **Fix**: Changed to `requireValueInList(['Push', 'NEW', 'Synced', 'ERROR'], true)`

### 4. SYNC/SKIP → Push Rename (DONE)
Renamed from free-text SYNC/SKIP to dropdown with "Push" as only user action. Removed SKIP entirely (user just leaves items as NEW).

### 5. Smart Location Routing (DONE)
Was hardcoded `MMH Kelowna / UNSORTED` for all items. Now routes based on Katana location + item type:
- Storage Warehouse items → `Storage Warehouse / SW-STORAGE`
- Materials/Intermediates at MMH → `MMH Kelowna / PRODUCTION`
- Products/other at MMH → `MMH Kelowna / UNSORTED`

---

## Changes Made to Each File

### 04_SyncEngine.gs
- `waspCreateItem()`: `ItemDescription` → `Description`, `ItemType: 0` → `'Inventory'`, added `Active: true`, `SalesPrice` → `SalePrice`

### 04b_RawTabs.gs (writeKatanaTab)
- preserveMap: only preserves `'PUSH'` (removed SKIP)
- Dropdown: `['Push', 'NEW', 'Synced', 'ERROR']` with `setAllowInvalid(true)`
- Conditional formatting: Push (yellow), NEW (orange), Synced (green), ERROR (red) — removed SKIP rule
- Separator rows get dropdown cleared

### 05_PushEngine.gs
- Removed early return when Wasp tab has no pending rows
- `pushKatanaSyncRows_`: SYNC → PUSH detection
- Smart location routing per row (reads col C Location + col D Type)
- CREATE failure handling: tracks `itemCreated` + `addAttempted`
- After push: just `'Synced'` (no timestamp)

### 01_Code.gs (onEdit)
- Katana handler: SYNC → PUSH, removed SKIP handling
- Sets cell to `'Push'` + yellow row highlight

---

## Other Context

- **GAS V8 rules**: `var` only (no const/let), no arrow functions, no template literals
- **WASP API**: `waspApiCall()` sends POST with Bearer token, checks `code === 200` and no `HasError:true`
- **Katana API**: Products, Materials, Intermediates. Items have: sku, name, type, uom, locationName, qtyInStock, avgCost, salesPrice
- **SKU Lookup hidden sheet**: 7 cols (SKU, Name, UOM, Lot Tracked, Type, Price, Cost) — used by push engine for UOM/Price/Cost
- **autoCreateMissingItems**: Commented out in runFullSync (01_Code.gs ~line 232). User controls creation via Push dropdown.
- **WASP has no item update API**: `ic/item/update` exists but only for limited fields (LotTracking)

---

## File Paths
```
C:\Users\Admin\Documents\claude-projects\wasp-katana\
├── sync-engine\          # Google Apps Script files
│   ├── 01_Code.gs
│   ├── 02_KatanaAPI.gs
│   ├── 03_WaspAPI.gs
│   ├── 04_SyncEngine.gs
│   ├── 04b_RawTabs.gs
│   └── 05_PushEngine.gs
├── tests\                # Python test scripts
│   └── sync_all_v5.py    # Working WASP item create reference
├── knowledge\
│   ├── learnings.md
│   └── decisions.md
└── state\
    ├── tasks.md
    └── handoff.md         # THIS FILE
```
