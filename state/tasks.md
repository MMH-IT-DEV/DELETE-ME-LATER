# Tasks — Katana-WASP

## In Progress
- [ ] F5 ShipStation — CRITICAL: itemCount:0 on ALL shipments (no WASP deductions)
  - [x] Debug scan: confirmed 200+ orders affected across 14 SHIP_NOTIFY batches (Feb 17)
  - [x] Fix: robust URL construction (trim, remove existing params, re-append) in 06_ShipStation.gs
  - [x] Fix: debug logging (SS_DEBUG_RESPONSE) captures response structure, fetch URL, field names
  - [x] Fix: fallback field names (shipmentItems || items || orderItems) in processShipment + processVoid
  - [x] Fix: PO_CREATED severity (INFO not ERROR) in 08_Logging.gs
  - [ ] Deploy updated 06_ShipStation.gs + 08_Logging.gs to GAS
  - [ ] Wait for next SHIP_NOTIFY webhook → check Debug for SS_DEBUG_RESPONSE entry
  - [ ] If itemField shows "none" or "null" → need to inspect actual ShipStation response keys
  - [ ] Once items flowing: verify WASP SHOPIFY deductions working

## In Progress — F5 Deployment
- [ ] F5 ShipStation Migration (replaces Katana SO flow)
  - [x] Code built: 06_ShipStation.gs, 01_Main.gs, 08_Logging.gs
  - [x] Deployed + SHIP_NOTIFY webhook registered
  - [x] Void poll trigger running (every 5 min)
  - [ ] Test: label voided → items returned to SHOPIFY, Activity row updated
  - [ ] Test: label re-printed → deduct again, Activity row updated back

## Queued — Bug Fixes (from Debug scan Feb 17)
- [ ] F3 STORAGE location: WASP error -57009 "STORAGE does not exist" — check WASP for correct location name for Storage Warehouse site, update F3_SITE_MAP in 07_F3_AutoPoll.gs
- [ ] F4 Lot/DateCode missing (-57041): EGG-X, FGJ-US-1, FGJ-US-2 — need manual WASP stock fix (add items to PRODUCTION in WASP UI)
- [x] F4 Insufficient qty (-46002): PL-YELLOW-1, PL-YELLOW-2, AGJL-1 — now shows visible warning "WASP stock out of sync" in Activity + flow tabs (Feb 18)
- [x] F4 expiry date pre-flight: `extractIngredientExpiryDate()` + pre-flight check logs warning for batch-tracked ingredients missing expiry. Processing continues (non-blocking). (Feb 18)
- [x] F3 Repeated polling: ST-217 (B-WAX), ST-223, ST-224, ST-229 keep retrying failed transfers — FIXED: permanent error classification + max 3 retries + escalating backoff + "Failed"/"Skip" statuses

## RESOLVED — F4 Lot-Tracked Ingredient Deduction (Feb 23)
- [x] Diagnostic: confirmed Katana batch_transactions has batch_id only (no batch_number)
- [x] Diagnostic: confirmed Katana permanently deletes depleted batch records (404)
- [x] Fix: fetchKatanaBatchStock() 3-tier fallback (all fail for depleted batches)
- [x] Fix: waspLookupItemLots() rewritten from infosearch → advancedinventorysearch
- [x] Fix: advancedinventorysearch confirmed working (LUP-4 → Lot: UFC408A at PRODUCTION)
- [x] Fix: proactive WASP lot fallback in F4 handler (lines 906-928)
- [x] Fix: extractIngredientExpiryDate() now checks expiration_date (Katana field name)
- [x] **ROOT CAUSE 1**: advancedinventorysearch SearchPattern doesn't filter — returns ALL rows. LUP-1 was on page 2. Fixed: pagination loop (max 20 pages) in both lookup functions.
- [x] **ROOT CAUSE 2**: proactive fallback + -57041 retry used waspLookupItemLots (lot only) — WASP requires DateCode too. Fixed: switched to waspLookupItemLotAndDate, passes both lot+dateCode.
- [x] testWaspLotLookup() confirmed: LUP-1 found on page 2, lot=UFC401B, dateCode=2029-02-28T00:00:00Z
- [x] **VERIFIED**: MO-7264 completed — LUP-1 x1 consumed with lot:UFC401B exp:2029-02-28. All 6 ingredients consumed, output produced.

## Queued — Testing (v2 fixes)
- [ ] F4 lot fallback: trigger MO with lot-tracked ingredient (FGJ-MS-1)
- [ ] Issue #7: Retest F4 MO with lot-tracked ingredients
- [ ] Issue #12: Test SO cancellation (cancel Shopify order, watch Debug tab)

## In Progress — Lot Tracking & Stock Reorganization (Feb 20)
- [x] Added `writeLotTrackingTab()` — creates "Lot Tracking" tab with Katana vs WASP comparison
- [x] Added menu item "Write Lot Tracking Tab" in 01_Code.gs
- [ ] Deploy updated 04_SyncEngine.gs + 01_Code.gs to GAS
- [ ] Run "Write Lot Tracking Tab" from menu → verify tab shows all SKUs with match/mismatch
- [ ] Review mismatches: fix WASP lot tracking to match Katana (via ic/item/update or WASP UI)
- [ ] Stock reorganization: after lot tracking aligned, reorganize stock locations

## RESOLVED — WASP Item Auto-Create (Feb 24)
- [x] Correct endpoint found: `ic/item/createInventoryItems` (array payload)
- [x] `waspCreateItem()` rewritten in 04_SyncEngine with correct payload format
- [x] `normalizeUomForWasp()` added — maps Katana UOMs to WASP abbreviations (EA, pack, PCS, BOX, KG, g, ROLLS)
- [x] Inline creation in 05_PushEngine with DimensionInfo, UOM normalization, Katana enrichment (price, category)
- [x] WASP callouts configured: item added + item removed → GAS webhook URL
- [x] Item creation working: NI-GLASSHW created with pack UOM, cost from Katana, DimensionInfo
- [x] Stock ADD after create: `isCatalogOnly=false`, batch/lot tracking wired to ADD step
- [x] "Already exists" handling: CREATE failure proceeds to ADD stock

## In Progress — Activity Tab Edit + Retry (Feb 26)
- [x] Retry feature: retryMarkedItems() with header-level checkbox, multi-flow F1-F6 (Overstory, Feb 25)
- [x] onEdit trigger: FIXED — renamed onEdit→onActivityEdit, installable trigger via installActivityEditTrigger() (was broken: simple triggers don't fire on standalone web app scripts) (Feb 26)
- [x] pushActivityEdits(): parses old vs new text, computes delta, calls WASP API (Overstory, Feb 26)
- [x] Emoji cleanup: removed all emojis from Slack notifications (02_Handlers.gs), pick order complete (03_WaspCallouts.gs), audit comments (09_Utils.gs) (Feb 26)
- [x] F5 sub-item format: standardized Delivered action to location-only, void detail uses → SHOPIFY (Overstory, Feb 26)
- **NOTE**: 13_Adjustments.gs DELETED from live GAS by user (Feb 26) — edits not deployed
- [ ] **DEPLOY**: push updated 03_WaspCallouts.gs + 09_Utils.gs to GAS (13_Adjustments removed)
- [ ] **SETUP**: Run installRetryMenu() once in GAS editor — installs BOTH menu trigger AND Activity edit trigger
- [ ] **TEST**: Edit a sub-item cell (col D) on Activity tab → row should turn yellow + Status=Pending
- [ ] **TEST**: Click Katana-WASP menu → Push Activity Edits → verify WASP API call + status update
- [ ] **TEST**: Check retry checkbox on failed row → Retry Failed → verify WASP call + status update

## In Progress — Multi-Batch + Flow Format Alignment (Feb 26)
- [x] 08_Logging.gs: Enhanced logActivity() and logFlowDetail() with `isParent`/`nested`/`batchCount` flags for multi-batch tree indentation (`│   ├─` nested sub-rows)
- [x] 02_Handlers.gs F4: Replaced `'  ↳'` hack with proper `isParent`/`nested` flags on both Activity and Flow Detail logging
- [x] 02_Handlers.gs F1: Added multi-batch support — loops batch_transactions when >1, processes each with own qty/lot/expiry, parent+nested logging
- [x] 07_F3_AutoPoll.gs F3: Added multi-batch support — loops batch_transactions when >1, processes each with own qty/lot/expiry, parent+nested logging
- [x] 03_WaspCallouts.gs F2: Added "add"/"remove" action word to sub-items, standardized `lot:` (no space) format
- [x] F5 (06_ShipStation.gs): Verified format matches reference — no changes needed
- [ ] **DEPLOY**: push updated 08_Logging.gs + 02_Handlers.gs + 07_F3_AutoPoll.gs + 03_WaspCallouts.gs to GAS
- [ ] **TEST**: Trigger F4 MO with multi-batch ingredient → verify parent row + nested batch sub-rows in Activity + F4 tab
- [ ] **TEST**: Trigger F1 PO with multi-batch item → verify parent + nested rows
- [ ] **TEST**: Trigger F3 transfer with multi-batch → verify parent + nested rows
- [ ] **TEST**: Verify each flow tab only shows entries for that flow (F1→F1, F2→F2, etc.)

## In Progress — Manual Adjustments + F4 Batch Fix (Feb 24)
- [x] `13_Adjustments.gs` created: Adjustments sheet + `processAdjustments()` for manual stock corrections
- [x] `retryFailedMO(wkId)`: scans Activity Log, parses failed sub-items, generates adjustment rows
- [x] `listFailedMOs()`: lists all Partial/Failed F4 entries for review
- [x] F4 batch fallback: parse lot/expiry from MO name `(UFC410B 02/29)` when `mo.batch_number` is empty
- [x] F4 output lot: removed `stage === 'FINISHED'` restriction — sends lot for all stages
- [x] push-to-apps-script.js: added 10_, 11_, 12_, 13_ files to deploy list
- [ ] **DEPLOY**: `node scripts/push-to-apps-script.js` to push all files to GAS
- [ ] **TEST**: Run `getAdjustmentsSheet()` — creates Adjustments tab in debug spreadsheet
- [ ] **TEST**: Run `retryFailedMO('WK-1418')` — generates adjustment rows for MO-7226 failures
- [ ] **TEST**: Run `processAdjustments()` — pushes corrections to WASP
- [ ] **TEST**: Next MO with batch in name → verify output gets lot from MO name

## In Progress — 4-Tab Inventory Control Panel Rewrite (Feb 19)
- [x] 04b_RawTabs.gs rewrite: writeKatanaTab (7→10 cols), writeWaspTab (14 cols, editable), writeBatchTab, deleteOldTabs
- [x] 04_SyncEngine.gs rewrite: autoCreateMissingItems, waspCreateItem, enhanced logSyncHistory (5 cols), removed dashboard/reconcile
- [x] 05_PushEngine.gs (NEW): pushToWasp, computeRowActions_, executeAction_ — detect pending rows, compute deltas, push to WASP API
- [x] 01_Code.gs rewrite: simplified menu (Sync Now + Push to WASP), onEdit trigger for Pending detection, runFullSync orchestrator
- [x] Katana tab WASP Sync Control: 10-col layout (added Batch, Expiry, WASP Status), SYNC/SKIP user-controlled item creation, pushKatanaSyncRows_ in PushEngine, autoCreateMissingItems disabled (Feb 19)
- [x] Katana tab dropdown + smart routing: "Push" dropdown (single option), NEW/Synced/ERROR system-set, smart location routing (Materials→PRODUCTION, Products→UNSORTED, Storage WH→SW-STORAGE), per-SKU push, no timestamp on Synced, error retry allowed (Feb 19)
- [x] Fix: pushToWasp clears entire row background on success (not just status column)
- [x] Fix: Status conditional formatting uses whenTextContains('Synced') for timestamp-suffixed values
- [x] Deploy: pushed to GAS, removed old files, new deployment created (Feb 19)
- [x] Feature: onEdit expanded to ALL editable columns (Name, Site, Location, Lot Tracked, Qty) with multi-column pending detection (Feb 19)
- [x] Feature: Lot Tracked editable from sheet — col O (Orig Lot Tracked) hidden, pushToWasp detects change → calls waspUpdateItemLotTracking (Feb 19)
- [x] Feature: Row deletion detection via manifest — writeWaspTab saves manifest to ScriptProperties, pushToWasp compares sheet vs manifest, generates REMOVE actions for deleted rows (Feb 19)
- [x] Feature: Row insertion works as Pending ADD — new rows (no orig values) detected by computeRowActions_ Case 1 (Feb 19)
- [x] Fix: Pending preservation now uses orig site/location for key matching (not user-edited values) — pending rows survive hourly sync correctly (Feb 19)
- [x] Fix: Pending restoration now also restores user-edited Site, Location, and Lot Tracked values (Feb 19)
- [x] Wasp tab now 15 cols (was 14): added hidden col O = Orig Lot Tracked (Feb 19)
- [x] Batch tab now 18 cols (was 14): added hidden cols O-Q (Orig W.Qty, Orig W.Site, Orig W.Location) + col R (Sync Status). Pending preservation, conditional formatting, hidden columns all working (Feb 19)
- [x] onEdit handles BOTH Wasp (15 cols) and Batch (18 cols) tabs via config pattern — editable cols: Batch W.Site(7), W.Location(8), W.Qty(9) (Feb 19)
- [x] pushBatchPendingRows_ and computeBatchRowActions_ added to 05_PushEngine.gs — Batch pending rows push to WASP API with lot/dateCode, supports qty change, site/location move, new row ADD (Feb 19)
- [x] Katana tab Qty editable (12→14 cols): Orig Qty hidden col 13, Adj Status col 14, onEdit detects Qty changes, clearErrorsToPush handles both WASP Status + Adj Status (Feb 23)
- [x] pushKatanaAdjustments_ added to 05_PushEngine.gs: scans Adj Status=Pending, computes delta, pushes ADD/REMOVE to WASP, smart location routing (Feb 23)
- [x] logAdjustment_ function added to 04b_RawTabs.gs: creates/appends to Adjustments tab with 11-col layout (Timestamp, Source, SKU, Location, Lot, Old Qty, New Qty, Diff, System, Status, Note) (Feb 23)
- [x] Adjustment logging wired into all 3 push flows: Wasp push (line 128), Batch push (line 900), Katana adj (lines 1403/1447/1455) in 05_PushEngine.gs (Feb 23)
- [x] KATANA_COLS updated to 14 in pushKatanaSyncRows_ (line 1106) (Feb 23)
- [x] Excel demo: docs/adjustments-tab-demo.xlsx created with sample data and formatted cells (Feb 23)
- [ ] Deploy: push updated 01_Code.gs + 04_SyncEngine.gs + 04b_RawTabs.gs + 05_PushEngine.gs to GAS
- [ ] Test: Katana tab — edit Qty → Adj Status = Pending, yellow highlight
- [ ] Test: Push to WASP → Katana adj pushed, Adj Status = Pushed, logged to Adjustments tab
- [ ] Test: Wasp tab push → logged to Adjustments tab
- [ ] Test: Batch tab push → logged to Adjustments tab
- [ ] Test: Hourly sync preserves Pending adj rows (not overwritten)
- [ ] Note: WASP `ic/item/update` endpoint not needed — user prefers manual WASP UI for lot tracking changes (zero stock → WASP UI → resync)

## In Progress — Wasp Tab UX Improvements (Feb 19)
- [x] FIX: (unknown) items should default to site="MMH Kelowna", location="UNSORTED" in writeWaspTab
- [ ] FEATURE: Delete item from WASP catalog — new DELETE_ITEM action when user deletes a 0-qty row from sheet (manifest comparison). Need WASP delete endpoint.
- [x] FEATURE: SKU Lookup hidden sheet — written during sync with all Katana+WASP SKUs, Name, UOM, Lot Tracked. Used by onEdit for auto-populate.
- [x] FEATURE: New row auto-populate — onEdit detects SKU typed in col A on new Wasp row, looks up SKU Lookup sheet, fills Name/UOM/LotTracked automatically
- [x] FEATURE: New row CREATE + ADD — pushToWasp checks if SKU exists in WASP, if not: creates item via ic/item/create THEN adds stock at chosen location. CREATE error handler always proceeds to ADD (item may already exist).
- [x] FIX: Lot (col F) and DateCode (col G) editable — added to onEdit origColMap with orig tracking (cols O, P)
- [x] FIX: Lot Tracked column shows wrong value — `fetchAllWaspData()` returned `false` instead of `undefined` when API doesn't include LotTracking field, causing writeWaspTab to trust it over lotTrackedSkus heuristic. Fixed: `ltVal` stays `undefined` when field absent.
- [x] FIX: DELETE handler gated lot/dateCode on `delIsLotTracked` — removed gating, always sends lot/dateCode. Also added Date object handling for delDateCode.

## Queued — Push Engine Bugs (from testing Feb 19)
- [x] BUG: Push REMOVE fails with -46002 when Orig Qty is stale → FIXED: -46002 retry handler fetches real WASP qty via advancedinventorysearch, retries REMOVE with actual qty (Feb 19)
- [x] BUG: Push ADD/REMOVE fails with -57041 when item is lot-tracked in WASP but sheet has no lot info → FIXED: -57041 retry handler — ADD retries with NO-LOT + future dateCode, REMOVE looks up actual lot from WASP then retries (Feb 19)
- [x] BUG: ERROR rows persist → FIXED: Added "Clear Errors → Push" menu item. Bulk-clears ERROR→Push (Katana) / ERROR→Pending (Wasp/Batch). User can also manually select Push on individual ERROR rows (onEdit already handles this). (Feb 20)
- [x] BUG: Lot/DateCode edits don't trigger Pending → ALREADY FIXED in local code: cols 6→15, 7→16 in onEdit origColMap + trackedPairs. Just needs deploy. (Feb 20 verified)
- [ ] BUG: Push uses delta approach (newQty - origQty) which is fragile when Orig Qty is out of sync with real WASP stock. Consider: fetch real WASP qty before computing delta, or use WASP `item/adjust` endpoint with absolute target qty.
- [ ] Test: Run "Sync Now" → verify 4 tabs created (Katana, Wasp, Batch, Sync History), old tabs deleted
- [ ] Test: Edit Qty on Wasp tab → verify Pending status + yellow background
- [ ] Test: Edit Name on Wasp tab → verify Pending (no orig comparison, always marks)
- [ ] Test: Edit Lot Tracked Yes→No → verify Pending + Push calls waspUpdateItemLotTracking
- [ ] Test: Click "Push to WASP" → verify WASP API calls, Synced status, Sync History log
- [ ] Test: Edit Location on Wasp tab → Push → verify MOVE (REMOVE old + ADD new) in WASP
- [ ] Test: Insert new row with SKU/Qty → Push → verify ADD in WASP
- [ ] Test: Delete a row → Push → verify REMOVE from WASP (manifest comparison)
- [ ] Test: Restore edited value to original → verify Pending clears only if ALL columns match originals
- [ ] Test: Hourly sync with pending rows → verify pending rows preserved (including site/location/lot tracked edits)

## In Progress — Inventory Sync System (older)
- [x] Build 10_InventorySync.gs — entry points, orchestrator, sheet writers (Feb 17)
- [x] Build 11_SyncHelpers.gs — Katana/WASP map builders, delta calc, adjustments (Feb 17)
- [x] Add SYNC_CONFIG + SYNC_LOCATION_MAP to 00_Config.gs (Feb 17)
- [x] Update push-to-apps-script.js with 2 new files (Feb 17)
- [ ] Deploy: `node scripts/push-to-apps-script.js` + new GAS deployment
- [ ] Phase 0: Run syncTestWaspRead() → verify advancedinventorysearch endpoint
- [ ] Phase 0: Run syncDiscoverLocations() → verify WASP site/location names
- [ ] Phase 1: Run syncDryRun() → review Sync-Deltas tab
- [ ] Phase 2: Set MAX_EXECUTE_ITEMS=5, run syncExecute() → verify 5 items in WASP UI
- [ ] Phase 3: Set MAX_EXECUTE_ITEMS=999, run syncExecute() → full sync

## DONE — Lot/Batch Display + Qty Coloring (Feb 23)
- [x] 08_Logging.gs: RichTextValue qty coloring in logActivity() and logFlowDetail() — green for add, red for remove
- [x] 02_Handlers.gs: F1 lot display in Activity sub-items, F4 qtyColor red (ingredients) and green (output)
- [x] 03_WaspCallouts.gs: F2 qtyColor green (add) and red (remove) in Activity + flow detail
- [x] 07_F3_AutoPoll.gs: F3 lot + expiry in Activity sub-items (neutral — no color for transfers)
- [x] 06_ShipStation.gs: F5 lot fallback (-57041 retry), qtyColor red (ship) and green (void)
- [x] F4 multi-lot fix: when batch_transactions has 2+ entries, loop each with its own qty+lot instead of sending total qty to first lot (MO-7208 EGG-X: 9 from HID260213 + 191 from HID260218)
- **Deploy**: push all 5 files to GAS, create new deployment

## Queued — Audit & Backfill
- [ ] Run webhookAudit_RUN() after ShipStation cutover — MISSING should drop
- [ ] Run backfillMissedSOs() if audit shows missed SOs (97_WebhookAudit.gs)

## Queued — Cleanup
- [ ] Fix B-WAX qty 1,222,576.6 in Katana (bad data)
- [ ] Clean 3 "TEST DELETE ME" StockTransfers entries (ST-233, ST-225, ST-234)

## Queued — Integration
- [ ] Fix STItems duplicate row bug in 07_F3_AutoPoll.gs (check before append)
- [ ] Add Lot/DateCode fields to WASP quantity_added callout template (WASP admin UI)
- [ ] Issue #13 (deferred): Fix WASP item costs (currently 0.01 placeholder)
- [ ] WH-8OZ-W F2 adjustment failures — Katana rejects SA, needs item-sync fix (D24)

## Queued — Future Options
- [ ] WASP Pick Order for F2 grouping — Pick Orders group removals into one `pick_complete` callout (avoids 1-item-per-SA problem). Limitation: Pick Orders are removal-only, can't add items. Feature request submitted to WASP (Jason, Feb 2026). Park until WASP adds grouped Add callout.
- [ ] Item Sync System — auto-sync Katana items to WASP (add/update). Track in katana-wasp-issues/item-sync.md. WH-8OZ-W needs lot data fix first.
- [x] F4 MO stock corrections — `auditDuplicateMOs()` + `executeDuplicateMOCorrections()` added to 09_Utils.gs. Scans Activity sheet for duplicate F4 entries, generates WASP removal commands. (Feb 18)
  - [ ] Deploy + run `auditDuplicateMOs()` to review duplicates, then `executeDuplicateMOCorrections()` to apply

## Queued — Batch Fix: F5 Logging Consistency (Feb 19)
- [x] F5 sub-items: Status col E empty → FIXED: always pass itemResults to logActivity (not gated on length > 1)
- [x] F5 sub-items: Error col F empty on failures → FIXED: auto-error-builder now runs since subItems always passed
- [x] F5 sub-items: "shipped" text in Details → FIXED: action now shows SHOPIFY location (Overstory, Feb 26)
- [x] F5 sub-items: emojis in Details → FIXED: all emojis removed from src/ (Feb 26)
- [ ] F5 headers: time-only timestamp (8:03) → full date (2026-02-19 08:03)
- [x] F5 headers: emojis in Details → FIXED: all emojis removed (Feb 26)
- [x] F5 header Error col: empty → FIXED: failCount appended to headerDetails + auto-error-builder populates col F
- [x] F5 single-item orders: sub-items hidden → FIXED: always pass itemResults (removed > 1 check)
- [x] F5 void sub-items: missing status/action fields → FIXED: added action:'SHOPIFY', status:'Voided', qtyColor:'green' (Overstory, Feb 25)
- [x] Exec ID race condition: duplicate WK-IDs → FIXED: getNextExecId() switched from getScriptLock() to getDocumentLock() (Feb 23)
- [x] F4 MO duplicate consumption on retry → FIXED: smart retry with getMOConsumedMap/saveMOConsumedMap tracks consumed SKUs per MO, skips already-consumed on reprocess (Feb 23)
- [x] F4 batch diagnostic logging → FIXED: F4_BATCH_DIAG logged when extractIngredientBatchNumber returns null (Feb 23)
- [x] Deployed to GAS (Feb 23) — https://script.google.com/macros/s/AKfycbzSojwLpeEo6irjUpCoUnRB9f_n8NW3vUWkShU8zt6kiHffbKkeJVxXNdspA0N28tv75w/exec
- **Files**: 06_ShipStation.gs (F5), 08_Logging.gs (exec ID), 02_Handlers.gs (F4 retry + batch diag)

## Blocked
- [ ] F2 grouping fix (items should group into one SA) — low priority

## Queued — Manual GAS Editor Steps (Phase 3 deploy — 6-col logging overhaul)
- [ ] Deploy all updated files: `node scripts/push-to-apps-script.js` from project root
- [ ] Create NEW deployment in GAS editor after pushing files
- [ ] Run `setupFlowTabFormatting()` once in **08_Logging.gs** — sets 6-col widths + updates row 3 headers on F1-F5 + Activity + deletes cols G-J + clears old flow tab data rows 4+
- [ ] Run `setupActivityConditionalFormatting()` once in **08_Logging.gs** — 6 flow header colors + 16 status colors on col E + error col F highlighting
- [ ] Run `clearActivityBackgrounds()` once in **08_Logging.gs** — removes old setBackground colors from existing Activity data (now handled by conditional formatting)
- [ ] Run `setupFlowSummaryFormulas()` once in **09_Utils.gs** — updates E1 formula for new status column E
- [ ] Run `setupF3Formatting()` once in **07_F3_AutoPoll.gs** — adds Failed/Skip conditional format rules

## Queued — Post-Deploy Cleanup
- [ ] Run `auditDuplicateMOs()` — 9 duplicate MOs found (MO-7183x4, MO-7184x2, MO-7185x3, MO-7175x2, MO-7189x2, MO-7171x2, MO-7164x2, MO-7163x2, MO-7176x2)
- [ ] Run `auditDuplicatePOs()` — PO-546 received twice (AGJL-1 x11520 excess at RECEIVING-DOCK)
- [ ] Run `cleanupTestRows()` — 7+ "TEST DELETE ME" entries in Activity log
- [ ] Check UFC-1OZ stock at SHOPIFY in WASP — 16+ F5 shipments failing (zero stock?)
- [ ] Delete STArchive tab (6 stale rows)

## Done
- [x] **Diff column formulas + conditional formatting** — 04b_RawTabs.gs: Diff col now uses `=IF(Adjust="",0,Adjust)` formula (auto-mirrors user input). Conditional formatting rules replace 500+ per-row API calls: green/red on Adjust+Diff, MATCH/MISMATCH/ONLY/ZERO colors on Match, blue on Lot Tracked "Yes". Per-row loop kept only for separators, ZERO whole-row grey, and Original col grey. (Feb 19)
- [x] **6-column logging overhaul** — Activity tab + all flow tabs migrated to 6 cols (ID, Time, Flow, Details, Status, Error). All emojis removed. 16 status values with conditional formatting. Sub-items use `status` field. Files changed: 08_Logging.gs (core), 02_Handlers.gs (F1/F4/F5-cancel/F6), 03_WaspCallouts.gs (F2), 06_ShipStation.gs (F5 ship/void), 07_F3_AutoPoll.gs (F3). (Feb 18)
- [x] sales_order.deleted removed from handledEvents — stops queue flooding with unfetchable deleted SOs (Feb 18)
- [x] handleSalesOrderCancelled graceful skip when SO fetch fails — returns 'skipped' not 'error' (Feb 18)
- [x] F3 skip Amazon USA and Shopify virtual locations — F3_SKIP_LOCATIONS array, immediate "Skip" status (Feb 18)
- [x] Exec ID race condition: `getNextExecId()` now uses ScriptProperties atomic counter + script lock. Bootstraps from sheet on first use. Fallback to WK-T{timestamp} if lock timeout. (Feb 18)
- [x] F1 PO idempotency: `isPOAlreadyReceived()` checks Activity log for existing Received F1 entry. Added permanent check after cache dedup in `handlePurchaseOrderReceived()`. (Feb 18)
- [x] PO audit/correction: `auditDuplicatePOs()` scans Activity for duplicate F1 entries. Added to 09_Utils.gs. (Feb 18)
- [x] Activity cleanup: `cleanupTestRows()` deletes "TEST DELETE ME" rows + their sub-items from Activity sheet. Added to 09_Utils.gs. (Feb 18)
- [x] F4 expiry pre-flight: `extractIngredientExpiryDate()` helper + pre-flight check in `handleManufacturingOrderDone()`. Logs warning for missing expiry, continues processing. (Feb 18)
- [x] F4 -46002 warning: visible "WASP stock out of sync" message in Activity error column + flow tab error column for insufficient qty items. Still counts as success. (Feb 18)
- [x] F4 MO audit/correction: `auditDuplicateMOs()` scans Activity for duplicate F4 Complete entries, `executeDuplicateMOCorrections()` applies WASP removals. Added to 09_Utils.gs. (Feb 18)
- [x] F4 idempotency: `isMOAlreadyCompleted()` checks Activity log for existing Complete F4 entry before processing MO. Failed/Partial allow reprocessing. (Feb 18)
- [x] F1 ingredient routing: materials/intermediates auto-route to PRODUCTION instead of RECEIVING-DOCK. Uses `getKatanaItemType()` per PO row. Products keep default destination. (Feb 18)
- [x] Item-sync tracking file created at katana-wasp-issues/item-sync.md — WH-8OZ-W documented (Feb 18)
- [x] `fetchKatanaProduct()` added to 04_KatanaAPI.gs for item type lookups (Feb 18)
- [x] Timestamp fix: 'HH:mm' → 'yyyy-MM-dd HH:mm' in logActivity() and logFlowDetail() (Feb 18)
- [x] ShipStation lock: added script lock for SS webhook processing in 01_Main.gs (Feb 18)
- [x] SO updated case fix: added .toUpperCase() for status comparison in handleSalesOrderUpdated() (Feb 18)
- [x] Hardcoded API keys removed from setupApiKeys() in 09_Utils.gs (Feb 18)
- [x] Silent WASP error fix: `.error` → `.response` fallback with `parseWaspError()` across all 5 flows (F1/F2/F4/F5), 10 instances in 02_Handlers.gs (Feb 17, v111)
- [x] F3 retry fix: permanent error classification (immediate fail) + transient retry with [Rx] counter + escalating backoff (1h→6h→24h) + max 3 retries + "Failed"/"Skip" statuses (Feb 17)
- [x] Flow tab polish: sub-item status removal (col E blank on sub-rows, color by error), 'returned' added to green status list, setupFlowTabFormatting() clears old data + deletes excess cols G-J (Feb 17)
- [x] Phase 2 log overhaul: Debug removal, flow tab 6-col tree format, WASP error code mapping, Shopify/ST links, exec ID bug fix, fake lot fix, emoji fix (Feb 17, all 6 src files edited locally — pending deploy)
- [x] Flow tab formatting: column widths + frozen rows for F1-F5 + Activity, error cleaning in logFlowDetail(), truncation 150→200 (Feb 17)
- [x] WebhookQueue old rows cleaned: deleted 3 pre-deploy rows with full JSON in Summary column (Feb 17)
- [x] Tab redesign: all 8 tabs standardized (navy headers, column widths, error formatting, status colors) (Feb 17)
- [x] WebhookQueue readability: restructured to 7 cols (hidden Payload, human-readable Summary+Result) (Feb 17)
- [x] Debug/Tracker sheet tab cleanup: deleted STArchive + PickOrderMappings, cleared Debug (19K rows) (Feb 17)
- [x] F5 ShipStation code built: 06_ShipStation.gs, 01_Main.gs, 08_Logging.gs (Feb 13)
- [x] V2 deploy: all files pushed, new deployment created, webhook URL updated (Feb 13)
- [x] G-OIL cleanup: removed 1600 duplicate units (lot 08002/25, DateCode 2028-02-13) from RECEIVING-DOCK (Feb 13)
- [x] F4 lot fallback: waspLookupItemLots() retry on -57041 in handleManufacturingOrderDone (Feb 13)
- [x] SO pick order dedup: handle -70010 "Another order already has this Order Number" as success (Feb 13)
- [x] Queue optimization: filter sales_order.updated to only queue CANCELLED/VOIDED (Feb 13)
- [x] Queue race condition: SpreadsheetApp.flush() after sheet creation (Feb 13)
- [x] WebhookQueue header fix: deleted displaced row 1 (Feb 13)
- [x] WASP stock sync complete: 266/275 rows added, 9 intentional skips (Feb 13)
- [x] Webhook queue system deployed + trigger running (Feb 13)
- [x] Disable lot tracking on finished products not batch-tracked in Katana (Feb 13)
- [x] formatDateCode fix in EXECUTE and RETRY functions (Feb 13)
- [x] SHOPIFY location created, PICK_FROM_LOCATION updated (Feb 13)
- [x] All earlier fixes (Feb 9-13)
