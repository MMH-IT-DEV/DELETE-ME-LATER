# wasp-katana — Agent Context File
## Read this before touching anything

---

## What This Project Does
Google Apps Script (GAS) project that syncs inventory data from **Katana MRP** (source of truth) to **WASP InventoryCloud** (mirror).

Syncs: **Cost**, **UOM (StockingUnit)**, **Category**, **Inventory Quantity**

Deployed via **clasp** from two local folders:
- `src/` — main production project
- `engin-src/` — separate engineering Apps Script project

**Push command:** `cd src && clasp push` (or `cd engin-src && clasp push`)

---

## How to Run the Daily Sync
Run **`syncCostsDailyTrigger`** from the Apps Script editor.
- This calls `syncAllItemMetadata(false)` — syncs Cost + UOM + Category for all ~300 items
- Also runs automatically at 3am daily via Apps Script trigger
- To set up the trigger: run `setupCostSyncTrigger()` once

To do a dry run first: run `syncAllItemMetadata(true)` — logs what would change without applying.

---

## Key Files

| File | What's in it |
|---|---|
| `src/00_Config.js` | All config: API keys, URLs, site mappings, UOM_CONVERSIONS |
| `src/04_KatanaAPI.js` | Katana REST API wrapper |
| `src/05_WaspAPI.js` | WASP REST API functions incl. `waspUpdateItemCost()` |
| `src/09_Utils.js` | ALL sync logic: `syncAllItemMetadata`, `buildKatanaUomMap_`, `buildWaspUomMap_`, `normalizeUomForCompare_`, debug helpers |
| `src/02_Handlers.js` | Webhook handlers: F1=PO receive, F2=stock adjust, F3=pick order |

---

## WASP API — Critical Knowledge (Do Not Guess)

### Working payload for `POST /public-api/ic/item/updateInventoryItems`
```json
[{
  "ItemNumber":          "CAR-4OZ",
  "Cost":                3.9876,
  "StockingUnit":        "Each",
  "PurchaseUnit":        "Each",
  "SalesUnit":           "Each",
  "CategoryDescription": "FINISHED GOODS",
  "DimensionInfo": {
    "DimensionUnit": "Inch",
    "WeightUnit":    "Pound",
    "VolumeUnit":    "Cubic Inch"
  }
}]
```

### Rules that took hours of debugging to discover:
1. **Payload must be a JSON array** `[{...}]` — not a plain object
2. **StockingUnit is always required** — omit it → error -57072 "base unit of measure does not exist"
3. **DimensionInfo must have exactly 3 fields**: DimensionUnit, WeightUnit, VolumeUnit. **Never add** Height/Width/Depth/Weight/MaxVolume — those cause updates to silently fail
4. **`TotalResults:0, SuccessfullResults:0, HasError:false` = SUCCESS** — WASP returns misleading metadata. The update IS applied. Verified by read-back. Treat it as success.
5. **CategoryDescription** must exactly match a WASP category string (all caps). Mismatch → -64001. On -64001: retry without CategoryDescription so cost+UOM still apply.
6. **Do not use `advancedinfosearch` to look up by ItemNumber** — it searches descriptions. Use `infosearch` instead.

### Error codes:
| Code | Meaning |
|---|---|
| -57066 | DimensionUnit empty — add DimensionInfo |
| -57072 | StockingUnit empty — always include it |
| -64001 | Category doesn't exist in WASP — retry without CategoryDescription |

### Success check:
```javascript
var success = code === 200 && !body.includes('HasError\":true');
```

### Useful read endpoint — look up item by ItemNumber:
```
POST /public-api/ic/item/infosearch
{ "SearchPattern": "CAR-4OZ", "PageNumber": 1, "PageSize": 5 }
```
Returns: Cost, StockingUnit, CategoryDescription, DimensionInfo, etc.

---

## Katana API — Key Field Names
- Products and materials both use **`uom`** (not `unit_of_measure`)
- Category field is **`category_name`** (not `category`)
- Inventory cost is in **`average_cost`** on `/inventory` records
- Cost fallback for zero-cost items: `fetchLastAverageCostFromMovements(variantId)` in 09_Utils.js

---

## UOM Normalization
`normalizeUomForCompare_()` in `09_Utils.js` is used **only for comparison** (detecting mismatches). Raw Katana UOM is written to WASP as-is.

Current normalization:
- `ea`, `each` → `each`
- `pc`, `pcs`, `piece`, `pieces` → `pc` ← **separate from "each" — intentional**
- `g`, `gram` → `g` | `kg` → `kg` | `ml` → `ml` | `oz` → `oz` | `lb` → `lb`

Items in `UOM_CONVERSIONS` (00_Config.js) have intentional unit differences and are skipped for UOM sync.

---

## WASP Categories (configured in WASP UI — cannot be created via API)
```
DEVICES, EQUIPMENT, FINISHED GOODS, INTERMEDIATE PRODUCT,
MIEDEMA HONEY FARM, NON-INVENTORY, PACKAGING MATERIALS,
RAW MATERIALS, TOOLS/HARDWARE
```
To add a new category: do it manually in the WASP UI first, then sync will pick it up.

---

## Sync Flow (`syncAllItemMetadata` in 09_Utils.js)
1. `buildKatanaUomMap_()` → `{sku: {variantId, uom, category}}` from Katana products+materials
2. Paginate `/inventory` endpoint → costMap `{variantId: average_cost}`
3. `buildWaspUomMap_()` → `{sku: {uom, category}}` via WASP infosearch (paginated, PageSize:100)
4. For each SKU: compute waspCost (fallback to last movement if 0), detect UOM+category mismatch
5. Build payload — always include StockingUnit; include CategoryDescription only when changed
6. Batch POST to `updateInventoryItems` in batches of 50
7. Batch fail → retry individually → -64001 triggers fallback retry without CategoryDescription

---

## Debug Functions in 09_Utils.js
| Function | Use |
|---|---|
| `debugWaspSingleUpdate()` | Test updateInventoryItems on CAR-4OZ with read-back |
| `debugWaspCategory()` | Test category update approaches |
| `debugKatanaItemFields()` | Dump raw Katana product/material field names |
| `auditItemUoms()` | Compare all Katana vs WASP UOMs and categories |
| `testSyncOneCost()` | Sync cost for one SKU (edit SKU_TO_TEST inside) |

---

## Known Limitations
- **Category sync**: works for categories that exactly match WASP's configured list. Items with unrecognized Katana categories get cost+UOM updated but category skipped.
- **~84 items** currently have Katana category strings that don't match WASP exactly — investigate with `auditItemUoms()` to see which categories are mismatched.
- **Inventory quantity sync** is a separate flow (`10_InventorySync.js`) — not part of the daily metadata sync.
