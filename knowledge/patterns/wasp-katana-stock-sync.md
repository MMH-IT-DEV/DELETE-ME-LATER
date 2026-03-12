# WASP-Katana Stock Sync — Operations Guide

## Purpose
Make WASP InventoryCloud match Katana MRP inventory exactly. This is a full reset: zero WASP, then re-add everything from Katana's per-location inventory data.

## When to Use
- After major inventory discrepancies between WASP and Katana
- After bulk imports/corrections in Katana
- Periodic full sync (quarterly, etc.)

## Prerequisites
- Katana API key configured in GAS script properties
- WASP API token configured in GAS script properties
- 98_ItemComparison.gs deployed to Google Apps Script
- Comparison sheet: `1mqdZ1Yp9fIzpSMxlJYgfbYGSe1ii2N836L1stc4AtNU`

---

## Step-by-Step Procedure

### Step 1: Build the Re-Add Plan
Run `reAddPlan_BUILD()` from the GAS script editor.

**What it does:**
- Fetches all inventory from Katana (`GET /v1/inventory`) per location per variant
- Fetches all Katana locations (`GET /v1/locations`)
- Fetches all variants for SKU + product type mapping
- Maps each Katana location to a WASP site + location
- Writes rows to the "Re-Add Plan" tab

**Location mapping (hardcoded in the function):**
| Katana Location | WASP Site | WASP Location | Notes |
|-----------------|-----------|---------------|-------|
| MMH Kelowna (product) | MMH Kelowna | SHIPPING-DOCK | Finished goods |
| MMH Kelowna (material/unknown) | MMH Kelowna | RECEIVING-DOCK | Raw materials, packaging |
| Storage Warehouse | Storage Warehouse | SW-STORAGE | All item types |
| Shopify | MMH Kelowna | SHOPIFY | Channel-committed stock |
| Amazon USA | — | — | SKIPPED (FBA stock, not at facility) |

**Review the plan before proceeding:**
- Check row count makes sense
- Verify location mapping is correct
- Look for suspicious quantities (e.g., B-WAX 1.2M)

### Step 2: Add Batch/Lot Data
Run `reAddPlan_ADD_BATCHES()` from the GAS script editor.

**What it does:**
- Fetches batch_stocks from Katana (`GET /v1/batch_stocks?variant_id=X`) for each variant
- Fills in Lot and DateCode columns for batch-tracked items
- Splits multi-lot items into separate rows (appended at bottom)

**Expected results:**
- ~15 SKUs updated with lot data (varies by catalog)
- ~10 extra rows added for multi-lot items
- ~110 "errors" — these are just non-batch-tracked items (normal)

**Known limitation:** The ADD_BATCHES function assigns lot splits to the location of the FIRST row for that SKU. If an item exists at multiple locations, the split rows may get the wrong location. Check the appended rows at the bottom of the sheet and correct locations manually if needed.

### Step 3: Zero WASP Inventory
Run `zeroWasp_FROM_EXPORT()` from the GAS script editor.

**What it does:**
- Reads the WASP export sheet (`1j0dfwTzlmGoAu8axju88xsRSyA1qJwUp7tpbr3nsYso`)
- Removes all inventory from WASP using "Total In House" quantities
- Uses `formatDateCode()` for safe Date object handling

**Important:** Export fresh WASP data to the export sheet BEFORE running this. The function reads current stock from that sheet.

**Alternative:** If WASP is already at zero (e.g., from a previous sync), skip this step.

### Step 4: Execute the Re-Add Plan
Run `reAddPlan_EXECUTE()` from the GAS script editor.

**What it does:**
- Reads every row in the "Re-Add Plan" tab
- For items WITH Lot column: calls `waspAddInventoryWithLot()`
- For items WITHOUT Lot: calls `waspAddInventory()`
- Marks each row DONE or ERROR in the Status column

**Takes 3-5 minutes** for ~275 rows (rate-limited to avoid WASP throttling).

### Step 5: Handle Failures
Run `reAddPlan_RETRY_FAILED()` to retry ERROR rows.

**Common error types:**

| Error Code | Meaning | Fix |
|------------|---------|-----|
| -57041 "Date Code missing" | Item is lot-tracked in WASP but no lot/datecode provided | Either: (a) disable lot tracking in WASP for items Katana doesn't batch-track, or (b) fill in Lot + DateCode columns in the plan |
| -57010 "Asset Tag not found" | Item doesn't exist in WASP | Create the item in WASP admin first, then retry |
| -46002 empty ResultList | Item has no inventory to remove (zero step) | Ignore — item already at zero |

**Lot tracking rules:**
- If Katana batch-tracks the item → keep lot tracking ON in WASP, provide lot + datecode
- If Katana does NOT batch-track the item → turn lot tracking OFF in WASP, retry as [no lot]
- Check Katana item page → "Material tracking" → "Batch / lot numbers" to verify

**Date handling bug:** Google Sheets stores dates as Date objects. The `formatDateCode()` function in 98_ItemComparison.gs handles this correctly. If writing datecodes to the sheet via MCP or manually, they may be interpreted as dates — this is OK as long as the code uses `formatDateCode()`.

### Step 6: Post-Sync Cleanup
1. **Check for duplicate stock** from ADD_BATCHES location bug (lot splits at wrong location)
2. **Verify totals** — spot-check a few items in WASP vs Katana
3. **Skip rows** — items you intentionally don't want in WASP (e.g., Amazon FBA stock) can be left as ERROR

---

## Key Functions Reference

| Function | File | Purpose |
|----------|------|---------|
| `reAddPlan_BUILD()` | 98_ItemComparison.gs | Build plan from Katana inventory |
| `reAddPlan_ADD_BATCHES()` | 98_ItemComparison.gs | Fetch lot/expiry data from Katana |
| `reAddPlan_EXECUTE()` | 98_ItemComparison.gs | Push plan to WASP |
| `reAddPlan_RETRY_FAILED()` | 98_ItemComparison.gs | Retry only ERROR rows |
| `zeroWasp_FROM_EXPORT()` | 98_ItemComparison.gs | Zero all WASP inventory from export |
| `zeroWasp_EXPORT_PREVIEW()` | 98_ItemComparison.gs | Preview what zero would do |
| `formatDateCode()` | 98_ItemComparison.gs | Safe Date → YYYY-MM-DD conversion |

## Key Sheets

| Sheet | ID | Purpose |
|-------|----|---------|
| Comparison/Plan | `1mqdZ1Yp9fIzpSMxlJYgfbYGSe1ii2N836L1stc4AtNU` | Re-Add Plan tab, WASP Raw tab |
| WASP Export | `1j0dfwTzlmGoAu8axju88xsRSyA1qJwUp7tpbr3nsYso` | Fresh WASP inventory export |

## Re-Add Plan Columns

| Col | Header | Description |
|-----|--------|-------------|
| A | SKU | Item number (e.g., UFC-4OZ) |
| B | Name | Product name (may be empty) |
| C | Qty | Quantity to add |
| D | Katana Location | Source location in Katana |
| E | Item Type | product, material, or unknown |
| F | Katana UOM | Unit of measure from Katana |
| G | WASP Site | Target WASP site |
| H | WASP Location | Target location within site |
| I | Lot | Lot/batch number (blank for non-lotted) |
| J | DateCode | Expiry/date code YYYY-MM-DD (blank for non-lotted) |
| K | Status | PENDING, DONE, ERROR, or SKIP |

---

## Troubleshooting

**"Date Code is missing" on items that HAVE lot data in the sheet:**
The dateCode column contains a Date object that the old code mishandled. Ensure `formatDateCode()` is used (not `String().split('T')`). Redeploy 98_ItemComparison.gs.

**"Quantity is missing or invalid" in WASP CSV import:**
Total Available column must be > 0. Use Total In House value.

**Function times out (>6 min GAS limit):**
The execute function processes all rows sequentially. If you have 500+ rows, it may time out. Run again — it skips DONE rows and continues where it left off.

**ADD_BATCHES puts lot splits at wrong location:**
Known bug. The multi-lot split rows inherit the location from the first row of that SKU. After ADD_BATCHES, manually check the appended rows (bottom of sheet) and correct their WASP Site / WASP Location columns.

**Items at zero but WASP shows phantom stock:**
Re-export WASP data to the export sheet and verify. Old exports may have stale data from previous bulk imports.
