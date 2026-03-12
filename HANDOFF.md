# Katana-WASP Integration — Handoff Document
**Last Updated: 2026-02-27**

---

## What This Project Is

Google Apps Script (GAS) integration that syncs inventory between **Katana MRP** and **WASP InventoryCloud**.
When inventory moves in either system, the other system is automatically updated.
Code lives locally in `src/` and is deployed to Google Apps Script via `clasp push`.

**GAS Script ID:** `1vMGYLeN5iCcLzW9mF0DJB3bT-My_yj1Tv8MVVnaVWKLb4m582ZqyT4Lg`
**Debug Sheet:** `1eX7MCU-Is5CMmROL1PfuhGoB73yRF7dYdyXHqzMYOUQ`
**Sync Sheet:** `1FiG8G3J-IbKoCzOiQ4aVCg6N1JBS01w76igmpkECJSI`
**WASP Base URL:** `https://mymagichealer.waspinventorycloud.com`

---

## Folder Structure

```
wasp-katana/
├── src/                    ← Main GAS project (clasp push to deploy)
│   ├── 00_Config.gs        ← CONFIG, LOCATIONS, FLOWS, SITE_TO_KATANA_LOCATION
│   ├── 01_Router.gs        ← doPost entry point + routeWebhook (NEW - created this session)
│   ├── 02_Handlers.gs      ← F1/F3/F4/F5 Katana webhook handlers
│   ├── 03_WaspCallouts.gs  ← F2: WASP→Katana (add/remove/move handlers + batch queue)
│   ├── 04_KatanaAPI.gs     ← Katana REST wrappers
│   ├── 05_WaspAPI.gs       ← WASP REST wrappers
│   ├── 06_ShipStation.gs   ← ShipStation API
│   ├── 07_PickMappings.gs  ← Pick order → Katana SO mapping
│   ├── 07_F3_AutoPoll.gs   ← F3: polls Katana stock transfers → WASP
│   ├── 08_Logging.gs       ← logActivity, logFlowDetail, logAdjustment
│   ├── 09_Utils.gs         ← setupAdjustmentsLogTab(), Slack, test functions
│   ├── 10_InventorySync.gs ← Full reset sync (Zero WASP → Re-add from Katana)
│   ├── 11_SyncHelpers.gs   ← buildKatanaMap, retry wrappers
│   ├── 12_SyncSheetFormat.gs ← Sheet formatting
│   └── 13_Adjustments.gs   ← Manual corrections from sheet tab
├── engin-src/              ← Older sheet-based sync engine (separate GAS project)
│   ├── 01_Code.txt
│   ├── 02_Utils.txt
│   ├── 03_KatanaAPI.txt
│   ├── 04b_RawTabs.txt
│   └── 05_PushEngine.txt   ← Fixed this session (logAdjustment_ call sites)
├── knowledge/              ← Reference docs, API notes
├── HANDOFF.md              ← This file
└── CLAUDE.md               ← Claude AI instructions
```

---

## The 5 Sync Flows

| Flow | Direction | Trigger | Status |
|------|-----------|---------|--------|
| F1 | Katana PO received → WASP RECEIVING-DOCK | Katana webhook | ✅ Working |
| F2 | WASP qty add/remove/move → Katana SA | WASP callout → GAS doPost | ✅ Working (see issues below) |
| F3 | Katana stock transfers → WASP | GAS time trigger (auto-poll) | ✅ Working |
| F4 | Katana MO complete → WASP | Katana webhook | ✅ Working |
| F5 | Katana SO (ShipStation) → WASP | Katana webhook | ✅ Working |

---

## What Was Done This Session (Feb 26-27, 2026)

### 1. Created `src/01_Router.gs` (NEW FILE — critical fix)
**Problem:** `doPost` and `routeWebhook` only existed in the remote GAS editor, not in local `src/` files. Every `clasp push` deleted them, so WASP callouts returned HTTP 200 but nothing was processed.

**Fix:** Created `src/01_Router.gs` with `doPost` and `routeWebhook`.

Key routing logic:
- `quantity_added` → `handleWaspQuantityAdded`
- `quantity_removed` → `handleWaspQuantityRemoved`
- `item_moved` / `quantity_moved` → `handleWaspItemMoved`
- `purchase_order.*` / `sales_order.*` / `manufacturing_order.*` → Katana handlers

Note: Both WASP and Katana use `"action"` field. Katana values always have a dot (`sales_order.delivered`), WASP values never do (`quantity_removed`) — no conflict.

---

### 2. Fixed SKU Field in F2 Handlers (`src/03_WaspCallouts.gs`)
**Problem:** WASP callout sends `"ItemNumber"`, not `"AssetTag"`. All handlers were looking for `payload.AssetTag` → SKU was always undefined.

**Fix:** All 3 handlers now use `payload.ItemNumber || payload.AssetTag`.

---

### 3. Fixed Silent Failure in F2 (`src/03_WaspCallouts.gs`)
**Problem:** When SKU lookup failed (saRows.length === 0), function returned silently with no log entry.

**Fix:** Added `logActivity` call to the empty-saRows case so failures are always visible in the Activity tab.

---

### 4. Multi-Batch Same-SKU Support (`src/03_WaspCallouts.gs`)
Added `usedBatchKeys` guard — prevents the same lot being included twice in one batched SA when multiple callouts arrive for the same SKU+lot within the 4-second window.

---

### 5. Improved SA# and Reason Fields (`src/03_WaspCallouts.gs` + `src/04_KatanaAPI.gs`)
**SA# (stock_adjustment_number):**
- Single item → SKU (e.g. `"4OZSEAL"`)
- Multiple items → `"WASP Adjustment"`

**Reason field (now descriptive):**
- Single: `"4OZSEAL -3"`
- Multi: `"2OZSEAL -1, AGJAR-4 -2"` (up to 3 items shown, then "...+N more")

Added `reason` as 6th parameter to `createKatanaBatchStockAdjustment` in `04_KatanaAPI.gs`.

---

### 6. Move Handler (`src/03_WaspCallouts.gs`)
`handleWaspItemMoved` — logs to Adjustments Log only, no Katana SA created.
- Records: SKU, from→to location, lot, quantity, user
- Adjustments Log row: Action="Move", light blue background
- Reads: `FromLocationCode`/`FromLocation`, `ToLocationCode`/`ToLocation`/`LocationCode`

---

### 7. Adjustments Log Tab (`src/08_Logging.gs` + `src/09_Utils.gs`)
`logAdjustment(source, action, user, sku, itemName, site, location, lot, expiry, diff, katanaSaNum, saId, status)`

Columns: Timestamp, Source, Action, User, SKU, Item Name, Site, Location, Lot, Expiry, Qty Change, Katana SA#, Status

Color coding by Action:
- Add → green
- Remove → pink/red
- Move → light blue
- Sync → amber

SA# is a clickable hyperlink to Katana when `saId` is present.

`setupAdjustmentsLogTab()` in `09_Utils.gs` — run once from GAS editor to create/reset the tab.

---

### 8. Fixed `engin-src/05_PushEngine.txt`
5 stale `logAdjustment_` call sites had the old function signature. All fixed to match new 12-parameter signature:
`(ss, source, action, user, sku, itemName, site, location, lot, expiry, diff, status)`

Lines fixed: ~919, ~990, ~1608, ~1652, ~1660

---

## Pending / Not Yet Done

### MUST DO before F2 is fully reliable:

**A) Deploy `src/` to GAS**
```
clasp push
```
Then in GAS editor: Deploy → Manage Deployments → New Version → update the web app deployment to use the new version.

**B) Run `setupAdjustmentsLogTab()` once**
In GAS editor, run this function once to create the Adjustments Log tab with correct headers and formatting.

**C) Fix lot field in WASP callout templates**
Current template uses `{trans.Lot}` but the lot number that arrives doesn't match what WASP UI displays.
- Try changing to `{trans.LotNumber}` in the WASP callout Request Body for Add and Remove callouts.
- After changing, do a test remove on a lot-tracked item and verify the lot in the Adjustments Log matches what WASP shows.

**D) Verify user/assignee field**
Check if `{trans.Assignee}` correctly captures the operator name, or if it needs to be `{trans.ModifiedBy}` or `{trans.UserName}`.

**E) Test Add flow end-to-end**
Do a WASP Add and confirm:
- SA appears in Katana with correct lot+expiry
- Row appears in Adjustments Log (green, Action=Add)

**F) Set up Move callout in WASP**
The callout trigger must be:
- When: **any item** (NOT "any location")
- Is: **move**

Request Body:
```json
{
  "action": "item_moved",
  "ItemNumber": "{item.AssetTag}",
  "Quantity": "{trans.AssetTransQuantity}",
  "SiteName": "{trans.SiteName}",
  "FromLocationCode": "{trans.FromLocationCode}",
  "ToLocationCode": "{trans.ToLocationCode}",
  "Lot": "{trans.LotNumber}",
  "DateCode": "{trans.DateCode}",
  "Assignee": "{trans.Assignee}"
}
```

**G) engin-src deployment**
Copy updated `05_PushEngine.txt` content into the sync sheet GAS editor (separate project from src/).

---

## Critical Technical Rules (GAS)

- **`var` only** — no `const`, `let`, arrow functions, template literals, `for-of`
- WASP write endpoints need **ARRAY** payload `[{...}]` not object `{...}`
- `DateCode` in WASP = expiry date field
- Katana requires `batch_tracking` enabled on product for lot numbers to work
- `clasp push` REPLACES all remote GAS files — never add functions in the remote editor only
- SYNC_CACHE_SECONDS=120 prevents feedback loops between WASP↔Katana
- 4-second sliding window in F2 groups rapid callouts into one Katana SA

---

## Known Bugs / Limitations

| Issue | Status |
|-------|--------|
| Lot mismatch: `{trans.Lot}` ≠ WASP UI lot display | Pending — try `{trans.LotNumber}` |
| WASP `ic/item/create` endpoint returns 404 (removed ~Feb 2026) | Pending — endpoint unknown |
| F2 "add" path: if lot not in payload, does WASP inventory lookup | Working but depends on lot field fix |
| engin-src `ic/item/create` 404 bug | Pending — needs correct WASP endpoint |

---

## WASP Callout Templates (current working versions)

### Katana Sync - Remove
```json
{
  "action": "quantity_removed",
  "ItemNumber": "{trans.AssetTag}",
  "Quantity": "{trans.AssetTransQuantity}",
  "SiteName": "{trans.SiteName}",
  "LocationCode": "{trans.LocationCode}",
  "Lot": "{trans.LotNumber}",
  "DateCode": "{trans.DateCode}",
  "Notes": "{trans.Notes}",
  "Assignee": "{trans.Assignee}"
}
```
Trigger: any item / is: remove

### Katana Sync - Add
```json
{
  "action": "quantity_added",
  "ItemNumber": "{trans.AssetTag}",
  "Quantity": "{trans.AssetTransQuantity}",
  "SiteName": "{trans.SiteName}",
  "LocationCode": "{trans.LocationCode}",
  "Lot": "{trans.LotNumber}",
  "DateCode": "{trans.DateCode}",
  "Notes": "{trans.Notes}",
  "Assignee": "{trans.Assignee}"
}
```
Trigger: any item / is: add (or quantity added)

### Katana Sync - Move (TO BE CREATED)
```json
{
  "action": "item_moved",
  "ItemNumber": "{item.AssetTag}",
  "Quantity": "{trans.AssetTransQuantity}",
  "SiteName": "{trans.SiteName}",
  "FromLocationCode": "{trans.FromLocationCode}",
  "ToLocationCode": "{trans.ToLocationCode}",
  "Lot": "{trans.LotNumber}",
  "DateCode": "{trans.DateCode}",
  "Assignee": "{trans.Assignee}"
}
```
Trigger: **any item** (not "any location") / is: **move**

---

## Logging Architecture

| Function | Output | When |
|----------|--------|------|
| `logToSheet(type, meta, detail)` | Logger.log only (not sheet) | Debug/trace |
| `logActivity(flow, detail, status, label, subItems)` | Activity tab | Every F2 operation |
| `logFlowDetail(flow, execId, parent, children)` | F1-F5 tabs | Per-flow audit trail |
| `logAdjustment(source, action, user, sku, ...)` | Adjustments Log tab | Add/Remove/Move/Sync |

