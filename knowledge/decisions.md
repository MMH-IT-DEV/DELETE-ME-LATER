# Decisions — Katana-WASP

## D1: Apps Script over Python for sync runtime
**Date**: Jan 2026
**Choice**: Google Apps Script
**Reason**: Direct integration with Google Sheets for logging/tracking, native Sheets triggers, no separate hosting needed. Python used only for testing.

## D2: Katana as source of truth
**Date**: Jan 2026
**Choice**: Katana MRP is authoritative for inventory data
**Reason**: Katana is the manufacturing system — all stock movements originate there. WASP receives and reflects, never initiates.

## D3: Lot/expiry tracking via existing waspAddInventoryWithLot
**Date**: Feb 5, 2026
**Choice**: Use existing `waspAddInventoryWithLot()` function rather than building new endpoint
**Reason**: Function already exists in 05_WaspAPI.gs, handles the ARRAY payload format correctly, and is proven in other flows.

## D4: Activity Log with RichTextValue clickable links
**Date**: Feb 9, 2026
**Choice**: Use `SpreadsheetApp.newRichTextValue()` for inline clickable links on SO/MO/PO numbers in Activity log
**Reason**: Links go directly to Katana web pages (e.g., `factory.katanamrp.com/salesorder/{id}`). No extra columns needed — link is embedded in the Details cell text. Only order numbers link to Katana (not SKUs to WASP).

## D5: Flow-based header colors instead of status-based
**Date**: Feb 9, 2026
**Choice**: Color Activity log header rows by flow (F1 blue, F2 gold, F3 green, F4 purple, F5 cyan) instead of by status (green/red/yellow)
**Reason**: User wants to visually scan the log and immediately identify which flow each entry belongs to. Status is already shown via icons (✅/❌/⚠) in the Details text.

## D6: Confirm scope before coding
**Date**: Feb 9, 2026
**Choice**: Always ask "Is this everything, or are there more changes?" before implementing code changes
**Reason**: User feedback — implemented Activity log changes across 5 files without asking if there was more (Debug log redesign and flow tab enhancements were also wanted). Batching related changes together is more efficient than deploying incrementally.

## D7: Debug log designed for agent troubleshooting
**Date**: Feb 9, 2026
**Choice**: Debug log has 7 columns (Timestamp, Severity, Flow, Exec ID, Message, Context, Response) with auto-classified severity/flow
**Reason**: User specified "enough information for the lead agent to solve it." Structured context (SKU:X | Qty:Y | Loc:Z) and clean messages let an agent diagnose issues without needing raw JSON payloads.

## D8: Sheet queue over CacheService for webhook processing
**Date**: Feb 13, 2026
**Choice**: Use Google Sheets "WebhookQueue" tab instead of CacheService for queuing Katana webhooks
**Reason**: CacheService is fastest (~50ms) but "not guaranteed to persist" — silent data loss risk for webhooks. Sheet appendRow is ~200ms (well within Katana's 10s timeout) and provides persistence + auditability. Queue tab doubles as an ops dashboard showing pending/done/error counts.

## D9: WASP callouts stay synchronous, only Katana events queued
**Date**: Feb 13, 2026
**Choice**: Keep WASP callout events (quantity_added, quantity_removed, pick_complete) processing synchronously in doPost(); only queue Katana events
**Reason**: WASP callouts need the 4-second batch grouping window timing that happens within the request lifecycle. Katana events are the bottleneck (5-10s processing each, causing 26% drops during batch windows).

## D10: SHOPIFY location for F5 order fulfillment
**Date**: Feb 13, 2026
**Choice**: Create a new SHOPIFY location in WASP (MMH Kelowna site). Change F5 PICK_FROM_LOCATION from SHIPPING-DOCK to SHOPIFY.
**Reason**: Katana has a virtual "Shopify" location for channel-committed stock. Shopify orders should pull from SHOPIFY location in WASP to match. This separates Shopify-committed stock from general warehouse stock at SHIPPING-DOCK.

## D11: Skip Amazon USA stock in WASP sync
**Date**: Feb 13, 2026
**Choice**: Filter out Amazon USA location entirely during WASP stock sync
**Reason**: Amazon FBA stock is at Amazon's fulfillment centers, not at MMH facility. Including it in WASP would create phantom inventory. Amazon manages its own stock — no WASP tracking needed.

## D12: Disable lot tracking on finished products not batch-tracked in Katana
**Date**: Feb 13, 2026
**Choice**: Disable WASP lot tracking on finished products (CAR-2OZ, LCP-1/2/4, LTG-1/2/4, LUP-1/2/4, LUY-1/2, MS-IP-4, MT-1-W, US-IP-4, VP-50-W, WH-8OZ-W, GP-50-W) that Katana does NOT batch-track
**Reason**: If WASP lot-tracks items that Katana doesn't, every API call (add/remove) fails with -57041 "Date Code missing." This blocks not just the one-time sync but all ongoing webhook flows (F1, F2, F4, F5). Disabling lot tracking aligns WASP with Katana's tracking model.

## D13: Keep lot tracking on raw materials that Katana batch-tracks
**Date**: Feb 13, 2026
**Choice**: Keep WASP lot tracking on EGG-X, G-OIL, O-OIL and other raw materials where Katana has "Batch / lot numbers" enabled
**Reason**: These materials need lot traceability for production (used in BOMs). Katana's batch_stocks API provides lot numbers and expiry dates that WASP needs for inventory accuracy.

## D14: WASP lot lookup fallback for F4 ingredient removal
**Date**: Feb 13, 2026
**Choice**: When Katana doesn't provide batch data for MO ingredients, query WASP (`waspLookupItemLots()`) for the lot at PRODUCTION location before removing. Pattern: try without lot → if -57041 → look up from WASP → retry with lot.
**Reason**: Katana's `batch_transactions` include field is unreliable for some ingredients. WASP already knows which lots are at each location. Querying WASP is more reliable than trying to parse Katana's inconsistent batch allocation structures.

## D15: Filter sales_order.updated before queuing
**Date**: Feb 13, 2026
**Choice**: In doPost(), check `payload.object.status` for `sales_order.updated` events. Only queue if CANCELLED or VOIDED. Skip all other updated events immediately.
**Reason**: ~80% of webhook traffic is `sales_order.updated` events (status changes like NOT_SHIPPED → PARTIALLY_DELIVERED) that are immediately ignored by the handler. Only cancel detection needs these. Filtering before queuing reduces queue bloat from ~100 rows/hour to ~20 rows/hour.

## D16: Replace Katana SO with ShipStation for F5
**Date**: Feb 13, 2026
**Choice**: ShipStation SHIP_NOTIFY webhook + 5-min void poll replaces Katana SO webhooks entirely
**Reason**: Katana fires duplicate webhooks (same event re-sent 15-20 min later), SO dedup cache expires between re-fires. ShipStation is more reliable — fires once per shipment, has void tracking, and handles Shopify-only orders. Direct deduction from SHOPIFY location (no pick orders). Voided labels detected via poll → items returned to SHOPIFY. Original Activity row updated in-place for void/reprint (no duplicate rows).

## D17: ShipStation F5 design choices
**Date**: Feb 13, 2026
**Choice**: Direct deduct (no pick orders), same GAS deployment endpoint, 5-min void poll, Shopify channel only, clean cutover, original timestamp preserved, icon-only status (no color change for voids), single shipment per order
**Reason**: Simplest implementation that covers the full label lifecycle (create → void → reprint). Pick orders are unnecessary since the label IS the pick confirmation. 5-min poll balances responsiveness vs API rate limits. In-place Activity row updates keep the log clean.

## D18: Grouped tree format for flow tabs (replacing flat rows)
**Date**: Feb 17, 2026
**Choice**: Rewrite all F1-F5 flow tabs from flat 10-column per-item rows to 6-column grouped tree format (Exec ID, Time, Ref#, Details, Status, Error) with header + `├─`/`└─` sub-items
**Reason**: Flat rows had redundant data (exec ID, time, ref# repeated on every item line). Tree format matches Activity tab's pattern, is scannable, and reduces visual noise. 6 columns means wider Details and Error columns — easier to read at a glance. All tabs share the same column layout, simplifying formatting functions.

## D19: Remove Debug sheet writes
**Date**: Feb 17, 2026
**Choice**: Simplify `logToSheet()` to Logger.log only, remove `getDebugSheet()`, make `clearDebugSheet()` a no-op
**Reason**: Debug sheet accumulated 19K+ rows, caused slow writes, and was redundant — Activity tab shows all operations with better formatting, flow tabs show per-item detail. Logger.log still available for GAS execution log (auto-expires). Keeping `clearDebugSheet()` as no-op prevents errors from existing triggers.

## D20: WASP error code mapping in cleanErrorMessage
**Date**: Feb 17, 2026
**Choice**: Map known WASP error codes (-46002, -57009, -57041, -70010) to short human-readable messages instead of just stripping prefixes
**Reason**: Raw WASP errors like "ItemNumber: FGJ-US-1 is fail, message: Date Code is missing for Lot tracked item" are long and hard to scan. Mapped versions like "Lot/DateCode required by WASP" are instantly understandable. Pattern matching fallback handles variations not caught by error code.

## D21: F1 auto-route materials/ingredients to PRODUCTION
**Date**: Feb 18, 2026
**Choice**: F1 PO receiving checks item type (material/intermediate vs product) and routes materials directly to PRODUCTION instead of RECEIVING-DOCK
**Reason**: Ingredients coming from POs are consumed during manufacturing (F4) at PRODUCTION. Routing them directly eliminates the need for a separate RECEIVING-DOCK to PRODUCTION transfer (F6 staging step). Uses Katana product type field via `getKatanaItemType()`. Products/finished goods still go to RECEIVING-DOCK (default).

## D22: F4 idempotency via Activity log check
**Date**: Feb 18, 2026
**Choice**: Add `isMOAlreadyCompleted()` check in `handleManufacturingOrderDone()` that searches Activity sheet for existing F4 entry with same MO order_no and "Complete" status. Failed/Partial entries allow reprocessing.
**Reason**: Katana sends duplicate MO.done webhooks. Cache-based dedup (300s) expires between retries. Activity log is the permanent record — checking it prevents duplicate stock additions even across long intervals. Batch-read of columns A-D is efficient (~100ms for 1000 rows).

## D23: Keep FG at PROD-RECEIVING only
**Date**: Feb 18, 2026
**Choice**: F4 finished goods stay at PROD-RECEIVING. No auto-move to SHOPIFY or other locations.
**Reason**: User decision. PROD-RECEIVING is a staging area for QC/packaging. The move from PROD-RECEIVING to a sellable location (SHOPIFY) is a separate manual/transfer process, not part of the MO completion flow.

## D24: Item sync system — future build
**Date**: Feb 18, 2026
**Choice**: Track item sync needs in `katana-wasp-issues/item-sync.md`. Build a dedicated Katana→WASP item catalog sync system later.
**Reason**: WH-8OZ-W exists in both systems but WASP has no lot data. Other items may have similar issues. A proper item sync system (triggered by Katana product webhooks) would catch these automatically. Deferred to focus on fixing active flow bugs first.

## D25: F4 expiry pre-flight — warn, don't abort
**Date**: Feb 18, 2026
**Choice**: Log warning when batch-tracked MO ingredients are missing expiry dates, but continue processing the MO. Do not abort.
**Reason**: User confirmed batch_transactions includes expiry. Missing expiry is a data quality issue in Katana, not a system error. Aborting would block production. WASP accepts blank expiry for lot-tracked items. Warning logged via `MO_MISSING_EXPIRY` for operator to fix in Katana.

## D26: F4 -46002 shows visible warning, not silent success
**Date**: Feb 18, 2026
**Choice**: -46002 (insufficient qty) items still count as "success" but now show visible warning text "WASP stock out of sync" in Activity error column and flow tab error column.
**Reason**: User wants to see which items have WASP stock mismatches at a glance. Previously these were silently absorbed — ingOk was set true and error was left blank. Now operators can identify which items need manual WASP stock fixes (PL-YELLOW-1, PL-YELLOW-2, AGJL-1, etc).

## D27: MO duplicate correction via audit + execute functions
**Date**: Feb 18, 2026
**Choice**: Two-step approach: `auditDuplicateMOs()` scans and reports, `executeDuplicateMOCorrections()` applies. Both in 09_Utils.gs.
**Reason**: Corrections are destructive (remove inventory from WASP). Two-step lets operator review the audit output (Logger.log) before committing. Parses SKU+qty from Activity details text. Affected MOs: MO-7183, MO-7184, MO-7171, MO-7175, MO-7185, MO-7189.

## D28: Exec ID atomic counter via ScriptProperties
**Date**: Feb 18, 2026
**Choice**: Replace sheet-read approach in `getNextExecId()` with `PropertiesService.getScriptProperties()` + `LockService.getScriptLock()`. Bootstrap from sheet on first use (counter=0).
**Reason**: Activity log showed 13+ duplicate exec IDs (WK-1203, WK-1205, WK-1268 etc.) caused by concurrent F5+F2 processing reading same last ID from sheet. ScriptProperties persists across invocations, lock ensures atomicity. Fallback to WK-T{timestamp} if lock times out.

## D29: F1 PO permanent idempotency via Activity log
**Date**: Feb 18, 2026
**Choice**: Add `isPOAlreadyReceived()` check in `handlePurchaseOrderReceived()` after the 600s cache dedup. Mirrors the F4 `isMOAlreadyCompleted()` pattern.
**Reason**: PO-546 was received twice (23,040 AGJL-1 instead of 11,520) because Katana re-fired the webhook after cache expired. Cache-only dedup is insufficient — Activity log is the permanent record.

## D30: Remove sales_order.deleted from webhook queue
**Date**: Feb 18, 2026
**Choice**: Remove `sales_order.deleted` from `handledEvents` in doPost() and routeWebhook(). Change `handleSalesOrderCancelled` to gracefully skip when SO fetch fails (returns `skipped` instead of `error`).
**Reason**: Deleted SOs can't be fetched from Katana API (404). Queue was flooding with 30+ error rows per day. Since ShipStation handles all shipping deductions (D16), Katana SO deleted events require no WASP action. Cancelled SOs still handled via `sales_order.cancelled` and `sales_order.updated` (CANCELLED/VOIDED status).

## D31: F3 skip Amazon USA and Shopify virtual locations
**Date**: Feb 18, 2026
**Choice**: Add `F3_SKIP_LOCATIONS` array in 07_F3_AutoPoll.gs. Transfers involving Amazon USA or Shopify are immediately marked "Skip" with no WASP API calls.
**Reason**: Amazon FBA stock is at Amazon's fulfillment centers (D11). Shopify is a virtual channel location in Katana, not a physical warehouse. Neither has a WASP equivalent. Previously these failed with "Unknown source/dest" and polluted the StockTransfers tab as permanent failures.

## D32: F2 log shows WASP location not just site name
**Date**: Feb 18, 2026
**Choice**: F2 Activity log and flow tab now show specific WASP location (SHIPPING-DOCK, RECEIVING-DOCK, etc.) instead of just site name (MMH Kelowna). Format: `WASP → Katana @ SHIPPING-DOCK @ MMH Kelowna`. Also shows lot number in Activity header when single-item SA has a lot.
**Reason**: User needs to see exactly which WASP location triggered the adjustment. The WASP callout payload includes `LocationCode` — it was captured but not used in logging. Also filter out unresolved WASP template variables (`{trans.Lot}`, `{trans.DateCode}`) that appear when items don't have lot tracking in WASP.

## D33: Unified 6-column Activity tab — no emojis, status column
**Date**: Feb 18, 2026
**Choice**: Activity tab migrated from 4 columns (ID, Time, Flow, Details with embedded emoji+status) to 6 columns (ID, Time, Flow, Details, Status, Error). All emoji icons removed. Conditional formatting on Status column (E) replaces emoji parsing. 16 status values with colors: green (Complete/Received/Shipped/Synced/Consumed/Produced/Added/Deducted/Returned/Staged/Picked), red (Failed), yellow (Partial), amber (Skipped), grey (Voided/Cancelled). Sub-items use `item.status` field for column E. Flow tabs match same pattern.
**Reason**: User wanted all tabs "as optimized as flow tabs" — no double icons, same clean column-based format. Emojis caused parsing issues in idempotency checks and were inconsistent across flows. Status column is machine-readable for future automation.

## D34: F4 low stock — Approach A (remove what's available, log shortage)
**Date**: Feb 18, 2026
**Choice**: For F4 MO ingredients, -46002 (insufficient qty) items get status 'Skipped' with error 'Not in WASP'. Counted as success (Katana is source of truth). Still plan Approach A in future: query WASP stock before removing, take what's available, log shortage with "WASP: X (short Y)".
**Reason**: Katana tracks actual consumption. WASP stock being out of sync doesn't mean the ingredient wasn't consumed. The operator needs visibility into which items WASP doesn't track, not a blocking error.

## D36: Diff column formula + conditional formatting over static values + per-row formatting
**Date**: Feb 19, 2026
**Choice**: Replace hardcoded `0` in Diff column with `=IF(Adjust="",0,Adjust)` formula. Replace 500+ per-row `setBackground()`/`setFontColor()` calls with 11 sheet-wide conditional formatting rules applied in a single `setConditionalFormatRules()` call.
**Reason**: Static `0` meant Diff never reflected user-typed Adjust values (the whole point of the Diff column). Per-row formatting was the performance bottleneck — each `getRange().setBackground()` is a Sheets API call. Conditional formatting rules apply automatically to all cells in the range (including future edits) with 1 API call total. Kept per-row formatting only for separators (3-6 rows, must be whole-row colored), ZERO rows (whole-row grey), and Original column (grey tint) — these can't be expressed as column-level conditional rules.

## D35: F5 ShipStation — 'Deducted' not 'shipped' status
**Date**: Feb 18, 2026
**Choice**: F5 sub-items use 'Deducted' status (not 'shipped') for successfully removed items. Void returns use 'Returned'. Cancel returns use 'Returned'. Header status from FLOW_STATUS_TEXT['F5'] = 'Shipped' for the overall operation.
**Reason**: 'shipped' describes the logistics action (ShipStation), not what happened in WASP. 'Deducted' accurately describes the inventory action (stock removed from SHOPIFY location). Consistent with F4 'Consumed'/'Produced' pattern where status describes the inventory operation.

## D37: Two-channel reconciliation output (Wasp Raw Adjust + Wasp Updates tab)
**Date**: Feb 19, 2026
**Choice**: Reconciliation writes to TWO places: (1) Adjust column on Wasp Raw (per-row delta for existing WASP rows), (2) Wasp Updates tab (complete action list including CREATEs and lot additions). User reviews both, then clicks Execute.
**Reason**: Wasp Raw Adjust gives visual per-row feedback (green/red via Diff formula) but can't show items that don't exist in WASP yet. Wasp Updates tab is the complete execution list including CREATEs for new items and REVIEW flags for wrong-location stock. The Adjust column is for review; the Updates tab is what actually executes.

## D38: WASP ONLY items untouched during reconciliation
**Date**: Feb 19, 2026
**Choice**: Items in WASP but not in Katana are completely skipped — no removal, no modification, Adjust=0.
**Reason**: User decision. WASP may have items that Katana doesn't track (e.g., packaging, equipment, NI-* items). Removing them would lose tracking data. Katana is source of truth for items it tracks; WASP-only items are managed manually.

## D39: KATANA ONLY items auto-created in WASP (replaces REVIEW-only approach)
**Date**: Feb 19, 2026
**Choice**: Katana items missing from WASP are auto-created every sync. Lot tracking copied from Katana batch_tracking setting. Materials/Intermediates → PRODUCTION location. Products → MMH Kelowna site with NO location (human assigns later). Unknown types → same as Products.
**Reason**: User requested full automation. Manual REVIEW was too slow for 200+ items. Lot tracking is correctly inherited from Katana, and Products without a location prompt human placement.

## D40: 4-Tab Inventory Control Panel (replaces 7-tab layout)
**Date**: Feb 19, 2026
**Choice**: Simplified to 4 tabs: Katana (read-only), Wasp (editable control panel), Batch (lot comparison), Sync History (unified log). Deleted: Dashboard, Wasp Updates, Sync, old Katana Raw/Wasp Raw/Batch Detail.
**Reason**: User wanted to manage WASP inventory directly from Google Sheets. The Wasp tab is the single source of edits — change Qty, Location, or Site, then Push to WASP. Reduces cognitive load and removes redundant tabs.

## D41: Wasp tab as editable control panel with onEdit detection
**Date**: Feb 19, 2026
**Choice**: Wasp tab has hidden Orig Site/Orig Location columns + visible Orig Qty column. onEdit trigger marks changed rows as "Pending" with yellow background. Push button scans for Pending rows, computes delta, pushes to WASP API.
**Reason**: Batch-then-push is safer than immediate execution. Pending visual feedback (yellow + status) lets user review before pushing. Hidden orig columns enable MOVE detection without cluttering the view.

## D42: Pending rows preserved across hourly refresh
**Date**: Feb 19, 2026
**Choice**: When hourly sync refreshes Wasp tab, pending rows are saved before clearing and restored after writing new data. Pending edits survive refresh.
**Reason**: User would lose work if pending edits were wiped every hour. Skip-and-restore pattern ensures no data loss while keeping the rest of the tab fresh.

## D43: Any editable column change marks Pending
**Date**: Feb 19, 2026
**Choice**: onEdit responds to columns B(Name), C(Site), D(Location), E(Lot Tracked), J(Qty). Columns with orig tracking (C/D/E/J) compare before marking. Name (B) always marks Pending. Clearing Pending requires ALL tracked columns to match originals.
**Reason**: User wants full control panel — any sheet edit should be treated as an intended change. Multi-column check prevents accidental clearing when one column is restored but another is still changed.

## D44: Lot Tracked editable via hidden Orig Lot Tracked column
**Date**: Feb 19, 2026
**Choice**: Added 15th hidden column (O=Orig Lot Tracked) to Wasp tab. Push engine compares col E vs col O, calls waspUpdateItemLotTracking() on change.
**Reason**: Needed original value to detect changes. Alternatives (fetch from API during push, always-send) were either slow or wasteful. Hidden column is consistent with existing orig pattern for Site/Location/Qty.

## D45: Manifest-based deleted row detection
**Date**: Feb 19, 2026
**Choice**: writeWaspTab saves a compact manifest to ScriptProperties (chunked for 9KB limit). pushToWasp compares current sheet rows against manifest. Missing rows generate REMOVE actions. No real-time detection — only on push.
**Reason**: GAS onEdit doesn't fire for row deletions. onChange trigger exists but doesn't identify which rows were deleted. Manifest comparison in pushToWasp is reliable and keeps the architecture simple. Deleted rows reappear on next sync if user doesn't push first (by design — Sync refreshes from WASP).

## D46: Katana tab WASP Sync Control — user-controlled item creation
**Date**: Feb 19, 2026
**Choice**: Replace autoCreateMissingItems with manual SYNC/SKIP workflow on Katana tab. New items default to "NEW" (not synced). User types "SYNC" to create in WASP at MMH Kelowna/UNSORTED. "SKIP" permanently ignores. Status preserved across hourly syncs by SKU. Items detected in WASP auto-set to "Synced". Katana tab expanded from 7 to 10 columns (added Batch, Expiry, WASP Status).
**Reason**: Auto-creating all Katana items in WASP was too aggressive — user wanted manual control over which items enter WASP. Per-item SYNC/SKIP gives granular control. Batch+Expiry columns enable lot-tracked push with correct lot number and dateCode. Status column provides at-a-glance visibility.

## D47: Katana tab Push dropdown + smart location routing
**Date**: Feb 19, 2026
**Choice**: Replaced free-text SYNC/SKIP with dropdown. Single dropdown option: "Push". System-set statuses: NEW (orange), Synced (green), ERROR (red). No "Ignore/Skip" option — just Push or leave as NEW. Smart location routing: Materials/Intermediates → PRODUCTION, Products → UNSORTED, Storage Warehouse items → SW-STORAGE. Push is per-SKU (all batch rows pushed together). After push: "Synced" (no timestamp). User can retry ERROR by selecting Push again. No Category/Taxable fields sent to WASP.
**Reason**: User wanted simpler UX — one action (Push) instead of two (SYNC/SKIP). Dropdown prevents typos. Smart routing eliminates manual UNSORTED-to-PRODUCTION sorting for materials. Storage Warehouse routing fixes mis-routing bug (was creating all at MMH Kelowna).

## D48: Katana Qty editable with delta-based WASP adjustment
**Date**: Feb 23, 2026
**Choice**: Make Katana tab Qty (col 9) editable. Editing triggers onEdit → compares against Orig Qty (col 13) → sets Adj Status (col 14) to "Pending". Push computes delta = editedQty - waspQty (not origQty) and sends ADD or REMOVE to WASP.
**Reason**: User needs to correct WASP stock from the Katana tab when quantities are out of sync. Delta uses waspQty (col 10) as base because that's the real WASP state — origQty (col 9) is the Katana state, not what WASP has. Matches the existing Wasp/Batch tab edit pattern.

## D49: Adjustments tab as append-only audit log
**Date**: Feb 23, 2026
**Choice**: All qty changes across Katana, Wasp, and Batch tabs are logged to a shared "Adjustments" tab with 11 columns (Timestamp, Source, SKU, Location, Lot, Old Qty, New Qty, Diff, System, Status, Note). Tab auto-created on first log entry.
**Reason**: User wants full audit trail of every stock change pushed to WASP. Append-only design ensures no data loss. Source column identifies which tab initiated the change. Tab creation on-demand avoids cluttering sheets that never use the feature.
