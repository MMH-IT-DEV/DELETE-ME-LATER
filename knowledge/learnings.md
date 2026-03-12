# Learnings — Katana-WASP

## API Behavior
- WASP write endpoints require ARRAY payload format `[{...}]`, not object — silent failure if wrong
- WASP `DateCode` field maps to expiry date (not obvious from field name)
- Katana `batch_stocks?batch_id=` pattern for batch tracking queries
- Katana API requires `batch_tracking` to be enabled on the product for lot numbers to appear
- `waspAddInventoryWithLot()` in 05_WaspAPI.gs already handles lot+expiry writes

## Architecture
- Main entry point (01_Main.gs) routes `quantity_added` events to F2 handler
- F2 handler lives in 03_WaspCallouts.gs (the "callout" pattern)
- Stock adjustment row structure: documented in LEARNINGS_CAPTURE.md
- F3 Auto Poll (07_F3_AutoPoll.gs) has working batch tracking pattern to follow

## Multi-Batch Logging Pattern
- 08_Logging.gs logActivity() and logFlowDetail() support nested batch sub-rows via flags on sub-items
- `isParent: true` + `batchCount: N` → renders `├─ SKU x{totalQty}  LOCATION (N batches)` with grey qty
- `nested: true` → renders `│   ├─ x{qty}  lot:{lot}  exp:{exp}` (deeper indent, no SKU)
- Multi-batch processing pattern: check `batch_transactions.length > 1`, loop each entry, push results with `multiBatch/multiBatchFirst/multiBatchLast/multiBatchTotal` flags
- F1, F3, F4 all support multi-batch; F2 is single-item per SA; F5 is single-lot per component
- F2 sub-items include action word "add"/"remove" at end of action string
- F4 output lot: first tries `mo.batch_number` from Katana API, then parses MO name `(LOT)` or `[LOT]` via regex `[\(\[]([^\)\]]+)[\)\]]`
- F4 output expiry: if MO name has `MM/YY` after lot → uses last day of that month; otherwise defaults to +3 years from MO completion date
- F4 output sub-item now shows both `lot:` and `exp:` in Activity + Flow Detail tabs
- F5 void: `updateActivityRow()` flips qty colors (red→green), status backgrounds (green→amber) on original row
- F5 void: `updateFlowDetailRow()` updates original F5 tab row in-place instead of creating a new row
- `updateFlowDetailRow(flow, orderNumber, newStatus, newSubStatus, newQtyColor)` — generic function in 08_Logging.gs for updating existing flow detail rows

## Deployment
- Apps Script deployment is manual (copy-paste to script editor)
- Test with Python test suite before deploying: `python test_stock_adjustment.py`
- Tracker spreadsheet shows flow status per-row
- Subagents (coder/sonnet) CANNOT handle large update_script_content payloads — they truncate files >200 lines
- Use Node.js push script (`push-to-apps-script.js`) to deploy all files via MCP proxy directly
- push-to-apps-script.js reads all .gs files from disk, fetches appsscript.json from live, POSTs to MCP proxy
- After pushing files, run `setupApiKeys()` then `checkConfig()` in script editor before deploying
- `setupApiKeys()` is in 09_Utils.gs — stores 3 API keys in ScriptProperties (run once, then delete the function)

## Google MCP
- All 70+ MCP tools verified working: Sheets (17), Docs Basic (6), Docs Enhanced (13), Drive Files (12), Drive Folders (6), Apps Script (10)
- IT Claude Desktop MCP scriptId: `1rMXgG60z2Rs8rw0tTDjLwr1Oc3fEdhvgE4QYCsUFYwLzEvoz44LZwYnT`
- Apps Script management tools NOW WORKING — fixed by creating own GCP project `794058818989` (IT-Claude-MCP) and linking it to the script. Default workspace project (284275201478) didn't allow API enablement
- To fix Apps Script API access: create own GCP project → enable Apps Script/Drive/Docs/Sheets APIs → set OAuth consent screen (Internal) → change GCP project in script settings → re-authorize → re-deploy
- MCP proxy lives at `C:\Users\Admin\mcp-google-script-2\index.js`, URL in `SCRIPT_URL` const
- WASP-Katana Apps Script scriptId: `1vMGYLeN5iCcLzW9mF0DJB3bT-My_yj1Tv8MVVnaVWKLb4m582ZqyT4Lg`
- Sheet values for update_sheet must be 2D array even for single cell: `[["value"]]`
- folderPath uses `/` separator and auto-creates nested folders
- Comprehensive reference skill at: `Documents/claude-skills/integrations/google-mcp-reference.md`

## Google Apps Script Gotchas (Feb 6, 2026)
- MUST use `var` instead of `const`/`let` — files with modern JS silently fail to load, making ALL functions in that file `undefined`
- MUST use `for (var i = 0; ...)` instead of `for-of` loops — same silent failure
- When a file fails to load, there's NO error — functions just become `undefined` and cascade across the project
- Must deploy ALL files together — partial updates cause cross-reference errors
- Must create NEW deployment after updating files — just saving isn't enough

## F2 Stock Adjustment (Feb 5-6, 2026)
- F2 remove flow: works, creates negative qty stock adjustments in Katana
- F2 add flow: works, but lot lookup returns wrong lot if not in callout payload
- `getWaspLotInfo()` returns FIRST lot found for item, not necessarily the correct one — always pass Lot/DateCode in callout
- WASP callout must include `"Lot": "{trans.Lot}"` and `"DateCode": "{trans.DateCode}"` for correct lot attribution
- Batch grouping (4-second window via CacheService) doesn't work well — concurrent webhooks fight over locks
- `SITE_TO_KATANA_LOCATION` in 00_Config.gs maps WASP SiteName → Katana location name (case-sensitive match)
- F3's `F3_SITE_MAP` maps in OPPOSITE direction (Katana → WASP) — don't confuse them

## Logging Cleanup (Feb 6, 2026)
- Removed 30+ verbose logs across all files — Debug sheet now shows errors only
- Log types kept: anything with ERROR, FAILED, NOT_FOUND in the name
- Log types removed: CONFIG_CHECK, RAW_REQUEST, all MO_*, SO_*, PO_*, WASP_PAYLOAD, BATCH_* processing logs, F3_POLL_*, F3_LINE_SYNC, PICK_MAPPING_* (non-error)
- Renamed batch error logs to F2_ prefix: F2_SITE_NOT_MAPPED, F2_LOCATION_NOT_FOUND, F2_SKU_NOT_FOUND
- F3 file had `fetchKatanaVariant` renamed to `fetchKatanaVariantF3` to avoid conflict with 04_KatanaAPI.gs

## Activity Log Format (Feb 6-9, 2026)
- Universal format across ALL flows: `[ref#]  [summary]  [status icon+text]  [context/direction]`
- ALL sub-rows shown (both success AND errors) — user wants to see every product moved
- Sub-row format: `    ├─ ✅ SKU x{qty} {action}` with tree lines (├─ and └─)
- Error sub-row: `    └─ ❌ SKU x{qty}: clean error message`
- Column layout: ID | Time (HH:mm) | Flow (F1-F5 label) | Details (one-line format)
- F1 PO: `PO-532  3 items  ✅ Received  → RECEIVING-DOCK`
- F2 SA single: `CAR-2OZ x20  ✅ Synced  WASP → Katana`
- F2 SA batch: `3 items  ✅ Synced  WASP → Katana`
- F3 ST: `ST-221  3 items  ✅ Synced  MMH Kelowna → Storage Warehouse`
- F4 MO: `MO-7093  UFC-1OZ x900  ⚠ 3 errors  {batch} → PROD-RECEIVING`
- F4 sub-rows: `✅ UFC-1OZ added to PROD-RECEIVING` / `❌ LUP-1 x900: Failed`
- F5 SO: `#90590  1 item  ✅ Shipped  → SHIPPING-DOCK`
- F5 sub-rows: `✅ CAR-2OZ x1` / `✅ LUP-1 x2`
- **Header row colors by FLOW** (not status): F1 #cce5ff (blue), F2 #fff3cd (gold), F3 #d4edda (green), F4 #f3e5f5 (purple), F5 #d1ecf1 (cyan)
- Sub-row colors: success #e8f5e9 (light green), error #ffebee (light pink)
- **Clickable links**: Reference numbers (SO#, MO#, PO#) are RichTextValue hyperlinks → Katana web pages
- Katana web URL: `factory.katanamrp.com/salesorder/{id}` (SO confirmed), MO/PO/ST paths in KATANA_WEB_PATHS config
- `SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(start, end, url).build()` for inline links
- Error messages CLEANED — `cleanErrorMessage()` strips `ItemNumber: X is fail, message:` prefix
- Template layout on tracker Activity tab: sheet ID 1eX7MCU-Is5CMmROL1PfuhGoB73yRF7dYdyXHqzMYOUQ

## Activity Sheet Formatting Fix (Feb 6, 2026)
- NEVER use hidden column D with color codes for conditional formatting — causes sticky backgrounds when rows are deleted
- Instead: conditional formatting on column C (Details) matching actual content text
- Matching rules on column C: "✅ Received", "✅ Synced", "✅ Complete", "✅ Shipped" → green; "❌ Failed" → red; "⚠️" → yellow; "    ✅" (4 spaces) → light green; "    ❌" (4 spaces) → light pink
- When content is deleted, formatting condition no longer matches → background reverts to white
- For LIVE code: logActivity() should use sheet.getRange().setBackground() for static colors (more reliable than conditional formatting)
- Column D exists but is hidden (1px wide) and empty — can be removed or repurposed

## Tracker Tab Structure (Feb 6, 2026)
- ALL 8 tabs follow CC design pattern: Row 1 title (#1c2333 13pt), Row 2 subtitle (#1c2333 gray 9pt), Row 3 headers (#2d3748 white bold 10pt), Rows 1-3 frozen
- Activity tab = main dashboard (chronological, all flows, tree lines for sub-items)
- 5 dedicated flow tabs: F1 Receiving, F2 Adjustments, F3 Transfers, F4 Manufacturing, F5 Shipping
- Flow tabs use flat tabular format (one row per item, columns specific to each flow)
- Activity uses tree lines: `├─` for middle items, `└─` for last item in a group
- Flow column (B): descriptive labels "F1 Receiving" etc. with per-flow colors (blue/gold/green/purple/cyan)
- F1 columns: Time, PO#, SKU, Qty, Lot, Expiry, Location, Status
- F2 columns: Time, Type, SKU, Qty, Lot, WASP Site, Direction, Status
- F3 columns: Time, ST#, SKU, Qty, Lot, From, To, Status
- F4 columns: Time, MO#, Action (Produced/Consumed), SKU, Qty, Batch, Location, Status
- F5 columns: Time, SO#, SKU, Qty, Status, Location
- StockTransfers: Katana ID, ST Number, Status, From, To, Items, Created, Updated, WASP Status, Synced, Notes (+ summary row 4)
- STItems: Katana ID, ST#, SKU, Qty, Lot, Expiry, From, To, Status, Error
- logActivity() writes to Activity tab; logFlowDetail() writes to flow-specific tab

## STItems Duplicate Bug (Feb 6, 2026)
- F3 poll was writing duplicate error rows to STItems on every run (ST-217 B-WAX appeared 15 times)
- Root cause: 07_F3_AutoPoll.gs appends to STItems without checking if the line already exists
- Fix needed: check for existing Katana ID + SKU before appending, or update existing row

## MCP Tool Serialization (Feb 6, 2026)
- Deferred MCP tools can lose schema context mid-session causing "Expected array/object, received string" errors
- Fix: run ToolSearch to reload the tool (e.g., `select:mcp__google-script__update_sheet`), then retry
- This happens with array/object params: update_sheet values, append_rows rows, format_sheet options
- Always reload tools after a serialization failure before retrying

## Execution ID System (Feb 6, 2026)
- Format: `WK-XXXX` (e.g., WK-001, WK-002) — auto-incrementing per execution
- Activity tab column A = ID, cross-referenced in all flow tabs as "Exec ID" column A
- logActivity() generates next ID: read last row col A, parse number, increment
- Flow tabs link back to Activity via shared Exec ID for tracing any execution across views
- Activity tab now has 4 columns: ID, Time, Flow, Details
- Flow tabs have Exec ID as first column followed by flow-specific columns

## Katana → WASP Item Import (Feb 6, 2026)
- Katana has 312 variants total: 232 materials, 72 products, 8 services
- Materials endpoint (`GET /v1/materials`) provides names + categories that variants endpoint alone doesn't
- Variant endpoint (`GET /v1/variants`) returns all SKUs but materials have null product_id (use material_id instead)
- 102 of 164 materials have null category — default NI-* prefix items to NON-INVENTORY
- Material UOMs from Katana: pc (65), EA (23), pack (22), pcs (18), PC (9), box (8), ROLLS (7), g (4)
- Normalize UOMs: pc/pcs/PCS/PC/1/EA/g/bucket → Each; pack/PK/6pack → PK; box/BOX → BOX; kg → KG; ROLLS stays ROLLS
- WASP has no item update API — `ic/item/update` and `ic/item/edit` both 404. Existing items must be updated via WASP UI
- WASP UOM creation is UI-only — no API endpoint for creating units of measure
- Final WASP UOM set: Each (exists, 240 items), PK (create, 30), BOX (create, 20), ROLLS (create, 7), KG (create, 2)
- 10 of 17 existing WASP items need UOM updates (5 PCS→Each, 2 Each→BOX, 1 Pound→Each, 1 dozen→Each, 1 PCS→Each)
- User decided: grams items → Each in WASP (not G), bucket → Each
- WASP `ic/item/infosearch` with empty SearchPattern returned 0 items — may need specific pattern or pagination
- WASP item create endpoint: `POST /public-api/ic/item/create` — NOT array format (unlike transaction endpoints)
- Description format for WASP: `[SKU] NAME` matching existing items like `[NI-B221210] SHIPPING BOX / 22x12x10`
- TrackLot and TrackDateCode fields on WASP item create enable lot/date code tracking checkboxes
- Subagents cannot handle update_script_content with large payloads — use Node.js push script instead
- Project files now organized: src/ (.gs), scripts/ (.js utilities), tests/ (.py), config/ (SQL exports, tokens)

## WASP Token & Auth (Feb 6, 2026)
- WASP token is a base64-encoded blob containing: Refresh_token, .issued, .expires, Token (inner JWT), client_id, UserId, roles, claims
- The full base64 blob is used as Bearer token (not just the inner Token field)
- Inner Token field uses URL-safe base64 (`-` and `_` instead of `+` and `/`)
- If inner token alone is sent, WASP returns "not a valid Base-64 string"
- If inner token is converted to standard base64, WASP returns "Failed to get claims identity"
- `invalid_token` (HTTP 400) means the full token is recognized as a token but rejected — could be expired, revoked, or corrupted
- Token generation UI: WASP Admin > User Settings > API User
- WASP has NO token refresh endpoint — /token, /public-api/token, /api/token, /public-api/auth/refresh all return 404
- WASP public-api GET endpoints mostly return HTML "Resource not found" — only POST transaction endpoints return real API responses
- Node.js `fetch` and `curl` behave identically with WASP tokens — no client-side difference
- WASP Inventory Import CSV for adding stock: Item Number, Quantity, Cost, Site, Location, Lot, Date Code

## Debug Log Redesign (Feb 9, 2026)
- Debug tab: 7 columns — Timestamp, Severity, Flow, Exec ID, Message, Context, Response
- Severity auto-classified from eventType: ERROR (default), WARN (SKIP/BUSY/ALREADY/NO_ITEMS), ALERT (ALERT/SLOW/LIMIT)
- Flow auto-detected from eventType prefix: F1_ or PO_ → F1, F2_ or SA_ or BATCH_ → F2, etc., fallback SYS
- Context field uses structured format: `SKU:CAR-2OZ | Qty:5 | Loc:PRODUCTION | MO:12345`
- buildDebugMessage() maps event types to clean actionable messages (not raw JSON)
- buildDebugContext() extracts key fields from data object into pipe-separated pairs
- Row coloring by severity: ERROR=#f8d7da (pink), WARN=#fff3cd (gold), ALERT=#ffe0b2 (orange)
- Dark header row: #1c2333 white bold

## Flow Tab Enhancements (Feb 9, 2026)
- All flow tabs now have Error column as last column (J for F1-F4, H for F5)
- logFlowDetail() writes one row per item with flow-specific columns + error + clickable links
- Status cell colored: green (#d4edda) for success/synced/complete, red (#f8d7da) for failed/error
- Order number column gets RichTextValue clickable link to Katana web page
- F5 column order fixed: Exec ID, Time, SO#, SKU, Qty, Location, Status, Error (was Status/Location swapped)
- Summary formulas in E1 via setupFlowSummaryFormulas() — can't use MCP update_sheet for formulas (treats "=" as literal)
- setupFlowSummaryFormulas() is a one-time setup function in 09_Utils.gs, run from Apps Script editor after deploy

## WASP CSV Import (Feb 9, 2026)
- WASP import form: "12-Inventory" with comma delimiter
- Required columns: Item Number, Quantity, Location, Site, Unit Cost (mandatory — cannot be 0)
- Lot-tracked items also require: Lot, Date Code — import fails with "Lot is missing; Date Code is missing" if omitted
- Items in WASP with lot tracking enabled: all raw materials (oils, waxes, herbs, EOs), wound care products (WH-8OZ-W, GP-50-W, VP-50-W, MT-1-W), some finished goods (CAR-2OZ)
- Items WITHOUT lot tracking: packaging components (boxes, labels, seals, jars, lids, leaflets, inserts)
- Katana API `variants` endpoint only returns ~46 product variants — does NOT include raw materials
- Katana API `materials` endpoint returned 0 results — may not exist in v1 API or requires different auth
- Workaround for costs: use 0.01 placeholder, need to sync real costs later from Katana CSV export
- WASP location structure: Site > Location (e.g., MMH Kelowna > PRODUCTION, SHIPPING-DOCK, RECEIVING-DOCK, PROD-RECEIVING)
- Item categorization: raw materials + packaging → PRODUCTION, finished goods → SHIPPING-DOCK, NI-* equipment → stays at default location
- Download tab as CSV: Google Sheets File → Download → CSV for current sheet

## Issue Fixes (Feb 9, 2026)
- Issue #1: F2 exec ID traceability — pre-allocate WK-XXX via getNextExecId() BEFORE creating Katana SA, pass as `additional_info`. logActivity() accepts optional 7th param `preExecId`.
- Issue #6: MO dedup — CacheService key `mo_done_{moId}` with 300-second TTL. Katana fires MO.done webhook 2-3 times per MO.
- Issue #9: PO ship-to routing — `KATANA_LOCATION_TO_WASP` map in 00_Config.gs maps Katana location names to {site, location}. PO has `location_id` field. WASP API functions now accept optional `siteName` last param.
- Issue #10: SO delivered F5/F3 detection — `getPickOrderBySoId()` reverse lookup in 07_PickMappings.gs. If pick mapping exists → F5, else F3. If pick status="completed" → skip WASP removal (already deducted during pick).
- Issue #11: SO cancelled — new `handleSalesOrderCancelled()` adds items back to SHIPPING-DOCK. Registered as `sales_order.cancelled` event in router. Cleans up pick order mapping.
- Issue #12: Katana cancel webhook unknown — registered `sales_order.cancelled`, need to test if it fires. Fallback: listen for `sales_order.updated` and check status field.
- WASP API functions (waspAddInventory, waspAddInventoryWithLot, waspRemoveInventory) now accept optional `siteName` as last param — defaults to CONFIG.WASP_SITE if not provided.
- WASP lot-tracked items REQUIRE `Lot` field in remove payload — without it, removal fails silently for lot-tracked items (non-lot items work fine)
- Added `waspRemoveInventoryWithLot()` for lot-aware removals and `waspLookupItemLots()` to query item lots before removing
- F4 handler now looks up item lots at PRODUCTION before removing ingredients — if lot found, uses lot-aware removal
- WASP item lookup API endpoint: `/public-api/items/getbyitemnumber` (POST, payload `{ItemNumbers: [...]}`) — needs verification, may need different endpoint

## F3 Stock Transfer Retry (Feb 9, 2026)
- F3 poll runs every 5 min but was re-processing failed STs every cycle, flooding Activity log
- Fix: check WASP status before syncing — skip 'Synced', throttle 'Error' to retry once per hour
- `getSTLastAttempt()` reads WASP_SYNCED timestamp column. `updateSTWaspStatus()` now always sets timestamp.
- `isRetry` flag suppresses Activity log entries on repeat failures — only first failure and successes logged
- To force retry a failed ST: manually set WASP Status to 'Pending' on StockTransfers tab

## Activity Tab Conditional Formatting (Feb 9, 2026)
- Switched from `setBackground()` to conditional formatting rules — colors auto-clear when rows are deleted
- `setupActivityConditionalFormatting()` — one-time setup, run from Apps Script editor after deploy
- `clearActivityBackgrounds()` — removes old sticky `setBackground` colors from existing data rows
- Rules: header rows colored by Flow label (column C), sub-items colored by ✅/❌ in Details (column D)
- Header detection: column A not empty (has WK-xxx), sub-item detection: column A empty

## SO Handler Data Unwrapping Bug (Feb 9, 2026)
- Katana API wraps single-resource responses in `{ data: { ... } }` — SO endpoint returns `{ data: { order_no, sales_order_rows, ... } }`
- `handleSalesOrderCreated()` correctly unwraps: `var so = soData.data ? soData.data : soData;`
- `handleSalesOrderDelivered()` and `handleSalesOrderCancelled()` were NOT unwrapping — read `soData.rows` and `soData.order_no` directly (always undefined)
- Result: 0 items processed, Activity log showed "0 items" or no entry at all
- Fix: add `var so = soData.data ? soData.data : soData;` and use `so.` instead of `soData.` for all SO field access
- Fallback: if no embedded rows, fetch separately via `fetchKatanaSalesOrderRows(soId)`
- Variant API endpoint returns resource directly (no `data` wrapper) — `variant.sku` works without unwrapping

## SKIP_SKUS Filter (Feb 9, 2026)
- OP-300 (Order Protection) is a Shopify service item — auto-fulfilled, no physical inventory, should never go to WASP
- Added `SKIP_SKUS = ['OP-300']` in 00_Config.gs — filtered in SO Created (pick order), SO Delivered, and SO Cancelled handlers
- Check with `SKIP_SKUS.indexOf(sku) === -1` before processing any SO line item

## F5 Activity Log Timing (Feb 9, 2026)
- `sales_order.created` fires for ALL new Shopify orders — don't log to Activity (warehouse pick order only)
- Activity log for F5 Shipping should ONLY appear in `handleSalesOrderDelivered()` — when items are actually shipped
- If only SKIP_SKU items (OP-*) are delivered (partial delivery), skip Activity log entirely — no trackable items
- SKIP_SKU_PREFIXES pattern (not exact match): `['OP-']` catches OP-195, OP-300, and any future OP-* variants
- `isSkippedSku(sku)` helper in 00_Config.gs — prefix match, used across all SO handlers

## SO Updated Fallback (Feb 9, 2026)
- `sales_order.cancelled` is NOT registered in Katana webhook settings — code exists but event never fires
- Fallback: `handleSalesOrderUpdated()` checks `payload.object.status` for 'cancelled'/'voided' and routes to cancel handler
- `sales_order.updated` IS registered in Katana webhooks — fires on any SO status change
- User still needs to manually add `sales_order.cancelled` to Katana webhook settings for direct detection

## F2 Echo/Feedback Loop Prevention (Feb 10, 2026)
- Every WASP inventory change (add/remove) fires a callout to our webhook
- If F1/F3/F4/F5 write to WASP, the callout triggers F2 creating redundant Katana SAs
- Fix: `markSyncedToWasp(sku, location, action)` BEFORE the WASP API call (not after!)
- The F2 handler checks `wasRecentlySyncedToWasp()` and skips if marker exists
- Race condition: callout can arrive BEFORE marker is set if marked after API call
- Marking BEFORE is safe: if the WASP call fails, no callout arrives anyway
- `SYNC_CACHE_SECONDS` increased from 60→120 for delayed callouts
- All 4 flows now mark before WASP calls: F1 (PO add), F3 (ST remove+add), F4 (MO remove+add), F5 (SO remove)

## SO Delivered Dedup (Feb 10, 2026)
- Katana fires `sales_order.delivered` multiple times for the same SO (like MO.done)
- Without dedup, same SO processed 3-4 times → 3-4 F5 entries + 3-4 F2 echoes
- Fix: CacheService key `so_delivered_{soId}` with 300-second TTL (same pattern as MO dedup)

## PO & SO Created Dedup (Feb 13, 2026)
- PO-547 was processed 4 times over 18 minutes (WK-722 through WK-725) — Katana retries on timeout
- Added PO received dedup: cache key `po_received_{poNumber}` with 600s (10 min) TTL. Dedup happens AFTER fetching PO header (need PO number from API), BEFORE WASP writes
- SO Created was firing duplicate WASP pick orders — WASP rejects with "Another order already has this Order Number" but clutters Debug log
- Added SO created dedup: cache key `so_created_{soId}` with 300s (5 min) TTL. Dedup before Katana API call — saves the fetch entirely
- Now all 4 webhook-driven handlers have dedup: MO done (300s), SO delivered (300s), SO created (300s), PO received (600s)

## Webhook Queue System (Feb 13, 2026)
- doPost() was processing everything synchronously (5-10s per webhook) — during batch delivery windows (30+ SOs in minutes), 26% of SO deliveries were dropped
- Katana has a 10-second timeout — if doPost() doesn't return 200 within 10s, Katana retries (4 attempts: initial + retries at 30s, 120s, 900s)
- Fix: split doPost() — WASP callout events process immediately (need batch timing window), Katana events queue to WebhookQueue sheet for async processing
- Queue write is ~200ms (sheet appendRow) vs old 5-10s synchronous processing — well within Katana's 10s timeout
- processWebhookQueue() runs every 1 min via ScriptApp time trigger, processes up to 10 PENDING items per run
- Uses LockService.tryLock(0) to skip if another instance is running — no contention
- Queue status flow: PENDING → PROCESSING → DONE/ERROR
- cleanupProcessedQueue() deletes DONE/ERROR rows older than 24h (runs after each batch)
- Sheet queue chosen over CacheService because cache "not guaranteed to persist" — risky for webhook data that must not be silently lost
- setupWebhookQueueTrigger() — run once after deploy to create the 1-min trigger (deletes old triggers first)
- backfillMissedSOs() reads Webhook Audit tab, finds MISSING+RECENT SO rows, builds fake payload `{action: 'sales_order.delivered', object: {id: soId}}`, calls routeWebhook()

## Katana Inventory API (Feb 13, 2026)
- `GET /v1/inventory` returns per-location stock: { variant_id, location_id, quantity_in_stock, quantity_committed, quantity_expected }
- Paginates with `limit` (default 50) and `page` params — X-Pagination header has total_records/total_pages
- `GET /v1/locations` returns all warehouse locations with id + name
- Katana product `type` field differentiates: 'product' (finished goods) vs 'material' (raw materials)
- WASP stock sync strategy: zero all WASP inventory, then re-add from Katana per-location data
- Location mapping for re-add: MMH Kelowna + product → SHIPPING-DOCK, MMH Kelowna + material → PRODUCTION, Storage Warehouse → SW-STORAGE
- `batch_stocks?variant_id=X` returns per-batch stock with batch_number, in_stock, best_before fields — needed for lot-tracked WASP items
- Re-add plan written to "Re-Add Plan" tab in comparison sheet for review before execution
- Katana has 4 locations: MMH Kelowna, Storage Warehouse, Amazon USA (FBA — skip), Shopify (channel-committed)
- Many Katana items return type 'unknown' (not just 'product' or 'material') — treat as material/packaging → RECEIVING-DOCK
- Katana batch_stocks totals may differ from per-location inventory totals (committed stock, in-transit, etc.)
- ADD_BATCHES multi-lot split rows inherit location from FIRST row of that SKU — bug causes wrong location assignment for multi-location items

## WASP Stock Sync (Feb 13, 2026)
- WASP items with lot tracking enabled REQUIRE both Lot AND DateCode for add/remove API calls — error -57041 otherwise
- WASP error -57010 "Asset Tag not found" means the item doesn't exist in WASP
- WASP error -57041 on remove actually STILL processes the removal — WASP is tolerant on remove but not on add
- Items lot-tracked in WASP but NOT batch-tracked in Katana → disable lot tracking in WASP (otherwise all future webhook syncs will also fail)
- GAS Date object handling bug: `String(dateObj).split('T')` splits at 'T' in "GMT"/"Thu" not at ISO position — use `formatDateCode()` instead
- Google Sheets stores date-like strings ("2026-01-30") as Date objects — getValues() returns Date, not string
- `formatDateCode()` handles Date objects via `instanceof Date` check with manual y-m-d formatting
- WASP export sheet "Total In House" (col F) is the real per-location qty, not "Total Available" (col C)
- Full sync procedure documented in: knowledge/patterns/wasp-katana-stock-sync.md
- SHOPIFY location created in WASP for Shopify channel-committed stock — F5 PICK_FROM_LOCATION updated to SHOPIFY

## Webhook Queue Reliability (Feb 13, 2026)
- Katana "Send a test" button shows "Test unsuccessful" even when webhooks work — the test payload has no matching event type, returns `{status: 'ignored'}` which Katana interprets as failure. Real events work fine.
- WebhookQueue sheet race condition: if two webhooks arrive simultaneously when the sheet is first created, data can be written before the header row. Fix: `SpreadsheetApp.flush()` after creating sheet + headers.
- WASP `-70010: "Another order already has this Order Number"` means the pick order was already created — treat as success, not failure.
- `sales_order.updated` is ~80% of all Katana webhook traffic but only needed for cancel detection. Filter by `payload.object.status === 'CANCELLED'` before queuing to reduce bloat.
- CacheService dedup can miss across separate GAS execution contexts — queue-level dedup is more reliable than in-handler dedup alone.

## F4 MO Ingredient Lot Tracking (Feb 13, 2026)
- Katana's `manufacturing_order_recipe_rows?include=batch_transactions` should provide batch allocation data, but `extractIngredientBatchNumber()` may still return null if Katana doesn't track the batch for that specific ingredient consumption.
- When Katana doesn't provide batch info but WASP requires it (-57041), use `waspLookupItemLots()` as fallback to look up the lot from WASP inventory at the removal location (PRODUCTION).
- Pattern: try without lot → if -57041 → look up lot from WASP → retry with lot.

## Hooks & Settings
- Windows does NOT expand `$HOME` in hook commands — always use full path `C:/Users/Admin/`
- Hook exit code 2 means "block tool" — Python's file-not-found also exits 2, causing accidental blocking
- Long JSON command strings get line-wrapped when copy-pasting code blocks — always write files directly via Write tool instead of asking user to copy-paste
- settings.json hooks must have all command strings on a single line (no newlines in JSON string values)

## ShipStation API (Feb 13, 2026)
- Webhook events: ORDER_NOTIFY, ITEM_ORDER_NOTIFY, SHIP_NOTIFY, ITEM_SHIP_NOTIFY, FULFILLMENT_SHIPPED, FULFILLMENT_REJECTED
- SHIP_NOTIFY payload is minimal: just `resource_url` + `resource_type`. Must GET resource_url to get actual shipment data
- resource_url format: `https://ssapiX.shipstation.com/shipments?storeID=123&batchId=456`
- **No webhook for voided labels** — must poll API to detect voids
- List Shipments API: `GET /shipments` with `voidDateStart/End` filters to find voided labels
- Each shipment has `voidDate` field — if populated, label was voided
- `includeShipmentItems=true` parameter returns item details per shipment
- ShipStation API uses Basic Auth: base64(API_KEY:API_SECRET)
- API keys already in GAS ScriptProperties: SHIPSTATION_API_KEY, SHIPSTATION_API_SECRET
- Base URL: https://ssapi.shipstation.com
- Katana duplicate webhook problem: same event re-fired 15-20 min later, cache-based dedup expires between fires
- ShipStation fires SHIP_NOTIFY once per shipment — no duplicate problem
- **CRITICAL**: SHIP_NOTIFY resource_url does NOT include shipmentItems by default — must append `&includeShipmentItems=true` to the fetch URL
- The void poll endpoint already supports `includeShipmentItems=true` as a query param — items returned correctly in poll
- GAS deployment gotcha: pushing code via Apps Script API only updates HEAD. Must Manage deployments → Edit → Version: latest → Save to serve new code at the same URL
- Bundle SKUs in Katana: VB-UFC-DUO (FLARE CARE DUO, is_auto_assembly: true, product_id: 15926773), VB-UHP (ULTIMATE HEALING PACK, is_auto_assembly: true, product_id: 15482245) — these are auto-assembly products, not standalone inventory items
- Katana webhooks can be managed via API: GET /webhooks (list), PATCH /webhook/{id} (update), DELETE /webhook/{id} (delete)

## F5 ShipStation itemCount:0 Bug (Feb 17, 2026)
- ALL SHIP_NOTIFY shipments return `shipmentItems: null/undefined` despite `includeShipmentItems=true` appended to fetch URL
- 200+ orders across 14 batches affected — zero WASP SHOPIFY deductions since ShipStation deploy
- URL construction verified correct: `?storeID=X&batchId=Y&includeShipmentItems=true`
- Added SS_DEBUG_RESPONSE logging to capture: fetchUrl, shipment object keys, itemField detection
- Added fallback: `shipment.shipmentItems || shipment.items || shipment.orderItems || []`
- Possible causes: ShipStation API changed field name, region endpoint difference, URL encoding issue
- After deploy: check Debug tab for SS_DEBUG_RESPONSE entry — `shipmentKeys` field reveals actual structure

## F3 STORAGE Location Bug (Feb 17, 2026)
- F3_SITE_MAP maps "storage warehouse" → location 'STORAGE', but WASP returns -57009 "STORAGE does not exist"
- 20+ SKUs affected across ST-224 and ST-229 stock transfers
- Need to check WASP admin for the correct location name at the Storage Warehouse site

## Debug Tab Behavior (Feb 17, 2026)
- After clearing Debug tab (19K rows), GAS sheet.appendRow() writes to row 19256 (not row 2) — cleared rows still count for getLastRow()
- To truly reset, need to DELETE rows not just clear content
- logToSheet severity auto-classification: defaults to ERROR; must add explicit rules for INFO events like PO_CREATED
- PO_CREATED was logged 3x per event — Katana retry behavior (initial + retries at 30s, 120s)

## Tab Formatting & Error Cleanup (Feb 17, 2026)
- F1-F5 flow tabs had NO column widths set — all auto-sized to content, making Error column too narrow to read
- `setupFlowTabFormatting()` is a one-time setup function (like `setupActivityConditionalFormatting()`) — sets widths on all tabs at once
- F1-F4: 10 cols [70,50,85,110,90,80,120,120,80,300]; F5: 8 cols [70,50,85,110,55,120,80,300]; Activity: 4 cols [70,50,120,550]
- Error column = 300px wide on all flow tabs — enough to read cleaned error messages
- `logFlowDetail()` was using raw `substring(0,150)` instead of `cleanErrorMessage()` — WASP prefix "ItemNumber: XXX is fail, message:" was not stripped in flow tabs
- `cleanErrorMessage()` truncation bumped from 150→200 chars to match Debug tab
- Subagent deployment confirmed broken for large payloads — always use `node scripts/push-to-apps-script.js`

## Phase 2 Log Overhaul (Feb 17, 2026)
- Flow tabs restructured from flat 10-column per-item rows to grouped tree format (6 cols: Exec ID, Time, Ref#, Details, Status, Error)
- New `logFlowDetail(flow, execId, header, subItems)` signature — header row + `├─`/`└─` sub-items, same tree pattern as Activity tab
- header param: `{ref, detail, status, error, linkText, linkUrl}` — linkText/linkUrl for clickable RichTextValue on Ref# cell
- subItems param: `[{sku, qty, detail, status, error}]` — each becomes a tree-line row
- `logToSheet()` simplified to Logger.log only — Debug sheet no longer written to (was slow + noisy, all useful data in Activity/Flow tabs)
- `clearDebugSheet()` made no-op — function kept so existing triggers/callers don't error
- WASP error code mapping in `cleanErrorMessage()`: -46002=insufficient qty, -57009=location not found, -57041=lot required, -70010=duplicate order
- Pattern matching fallback: "date code is missing", "insufficient"/"not enough", "location not found", "item not found", "lot not found"
- F5 ShipStation exec ID bug: `logActivity()` return value was not captured — `processShipment` used parent `execId` (UUID) instead of `WK-XXX`
- F4 fake lot removed: `mo.batch_number || ('MO-' + moId)` generated noise in Activity — changed to `mo.batch_number || ''`
- Shopify search links on F5: `https://admin.shopify.com/store/mymagichealer/orders?query={orderNumber}` — encodeURIComponent for safety
- F3 ST links: `getKatanaWebUrl('st', stId)` — KATANA_WEB_PATHS already has `st: '/stocktransfers/'`
- ⚠ vs ⚠️: bare ⚠ (U+26A0) renders as text emoji in some contexts; ⚠️ (U+26A0 U+FE0F) forces color emoji rendering
- `setupFlowTabFormatting()` updated: now sets 6-col widths [70,50,100,400,80,300] for ALL F1-F5 tabs + writes row 3 headers
- `setupFlowSummaryFormulas()` updated: status column now E (was I for F1-F4, G for F5) across all flow tabs
- `testAllLogging()` completely rewritten: 6 test calls with new 4-param format instead of 13 old-style calls

## Silent WASP Error Bug (Feb 17, 2026)
- `waspApiCall()` returns `{success, code, response}` — there is NO `.error` property
- All error logging across F1/F2/F4/F5 was reading `.result.error` which is always `undefined` → errors silently lost
- Fix: fallback pattern `res.result.error || parseWaspError(res.result.response, action, sku)` — future-proof if `.error` is ever added
- `parseWaspError()` is defined in `07_F3_AutoPoll.gs` but available globally in GAS (all .gs files share scope)
- 10 total instances fixed in `02_Handlers.gs`: lines 121, 143 (F1), 442, 464 (F2/SO delivery), ~745, ~762, ~790, ~804 (F4), 1127, 1148 (F5)
- This was the root cause of "❌ with no error message" in Activity/flow tabs for all failed WASP operations

## F3 Retry & Error Classification (Feb 17, 2026)
- F3 poller retried failed STs forever (once per hour) with no max — all 9 Error STs were zombie-retrying
- STItems tab didn't exist — `getSyncedItemsForST()` returned `{}`, so every retry re-attempted all items
- Deploy clears GAS script cache → `F3_LAST_SYNC_TIMESTAMP` resets → poller re-fetches ALL STs since MIN_CREATED_DATE
- ALL 9 current Error STs had permanent errors: STORAGE doesn't exist, Remove failed (insufficient qty), unmapped locations (Shopify, Amazon USA)
- Fix: error classification (permanent vs transient) + retry count `[Rx]` in Notes + escalating backoff (1h→6h→24h) + max 3 retries
- Permanent errors: `does not exist`, `not found`, `Remove failed`, `Insufficient`, `Unknown source/dest` — fail immediately, never retry
- Unmapped locations (`Unknown source`/`Unknown dest`) now write "Failed" not "Error" — permanent, obvious in sheet
- "Skip" status supported — user can manually set WASP Status to "Skip" in StockTransfers sheet to suppress any ST
- "Failed" and "Skip" render as grey in conditional formatting (vs red for Error) — visually distinct

## F1/F4 Flow Fixes (Feb 18, 2026)
- Katana product `type` field determines routing: 'material' and 'intermediate' → PRODUCTION, 'product' → default location
- Variant endpoint may embed `product.type` if product data is included — check embedded data first to avoid extra API call
- `fetchKatanaProduct(productId)` added as fallback when variant doesn't embed product type
- F4 idempotency: batch-read Activity sheet (4 columns, all rows) is fast (~100ms for 1000 rows). Search from bottom up since recent entries are most likely to match.
- F4 idempotency only blocks `✅ Complete` entries — `❌ Failed` and `⚠️ Partial` allow reprocessing (operator may have fixed data)
- SYNC_LOCATION_MAP already defines material→PRODUCTION and product→SHOPIFY routing — F1 now uses same item type logic
- WH-8OZ-W confirmed: exists in both Katana (batch AS081525-1, 140 PCS, exp 2027-08-29) and WASP (1751 at SW-STORAGE, 140 at SHIPPING-DOCK) but WASP has NO lot data — F2 adjustments fail
- Katana batch_transactions on MO ingredients include `expiry_date` and `best_before_date` fields — confirmed by user
- `extractIngredientExpiryDate()` mirrors `extractIngredientBatchNumber()` pattern: direct fields → batch_transactions → stock_allocations → picked_batches
- F4 expiry pre-flight is non-blocking (warn only) — WASP accepts blank expiry for lot-tracked items, and aborting would block production
- -46002 (insufficient qty) items: previously silent success, now show "WASP stock out of sync" warning. Still counted as success since Katana is source of truth
- Items needing manual WASP stock fix at PRODUCTION: EGG-X, FGJ-US-1, FGJ-US-2, PL-YELLOW-1, PL-YELLOW-2, AGJL-1
- `auditDuplicateMOs()` parses SKU+qty from Activity details text using regex `MO-\d+\s+(\S+)\s+x(\d+)` — depends on `logActivity` format staying consistent
- Two-step correction pattern (audit then execute) is safest for destructive WASP operations — operator reviews Logger output before committing
- `getNextExecId()` race condition: when F5 and F2 process concurrently, both read same last WK-ID from sheet and return same next ID. 13+ duplicates found in Activity log. Fix: ScriptProperties atomic counter + script lock
- Duplicate MOs are worse than originally tracked: 9 MOs (not 6), with MO-7183 processed 4x at different quantities (541 and 1050)
- PO receiving also has duplicate issue: PO-546 AGJL-1 x11520 received twice (23,040 total). F1 only had 600s cache dedup, no permanent check.
- `isPOAlreadyReceived()` mirrors `isMOAlreadyCompleted()` — both scan Activity sheet columns A-D from bottom up for flow+status match
- UFC-1OZ fails on every F5 shipment — likely zero stock at SHOPIFY location in WASP. 16+ affected orders.
- F3 transfers all failing for EGG-X (ST-218 through ST-222) and TTFC-*/UFC-1OZ (ST-232) — source location stock issues
- F2 batch grouping: 4-second window catches some items (UFC-4OZ x3) but many arrive as individual x1. F5 burst pattern may need longer window.
- `cleanupTestRows()` traverses backwards to preserve indices when deleting. Uses two-pass: first finds "TEST DELETE ME" headers, then finds their sub-item rows.

## Inventory Sync System Architecture (Feb 17, 2026)
- Built 2 new GAS files: 10_InventorySync.gs (507 lines, 10 functions) and 11_SyncHelpers.gs (454 lines, 5 functions)
- Composite key format: `SKU|SiteName|LocationCode` (pipe separator) for both Katana and WASP inventory maps
- SYNC_LOCATION_MAP translates `KatanaLocation|itemType` → `{site, location}` for WASP destination routing
- waspAdjustInventory() in 05_WaspAPI.gs hardcodes CONFIG.WASP_SITE — sync uses waspApiCall() directly with explicit SiteName per-item
- advancedinventorysearch endpoint (`/public-api/ic/item/advancedinventorysearch`) is UNVERIFIED — Phase 0 syncTestWaspRead() tests it
- Fallback: per-SKU inventorysearch with dynamic site list extracted from SYNC_LOCATION_MAP (not hardcoded)
- item/adjust endpoint accepts signed Quantity (positive=add, negative=remove) — simpler than separate add/remove calls
- Safety: 80% abort threshold, MAX_EXECUTE_ITEMS limit, DRY_RUN flag, pre-sync snapshot sheet, retry with exponential backoff
- Status values from syncApplyAdjustments(): 'OK', 'ERROR', 'DRY_RUN' — must match case in orchestrator status check
- All sync functions use "sync" prefix to avoid GAS global namespace collisions with existing 98_ItemComparison.gs functions

## Standalone Sync Dashboard (Feb 18, 2026)
- New standalone Google Sheet "Katana-Wasp Inventory Sync" (ID: 1FiG8G3J-IbKoCzOiQ4aVCg6N1JBS01w76igmpkECJSI)
- Apps Script project "Katana-Wasp Sync Engine" (scriptId: 1xRjhF03JdKNxn-3wcqUyOT39HunMgRhEGZzxvzug9_lDCot6vpn_MEeE)
- 6 tabs: Dashboard (per-location), Katana Raw, Wasp Raw, Batch Detail, Wasp Updates, Sync Log
- 6 code files: 01_Code, 02_KatanaAPI, 03_WaspAPI, 04_SyncEngine, 05_WaspUpdater, appsscript.json
- Dashboard is per-location (not SKU-aggregate) — matches Katana inventory entries to WASP via SYNC_LOCATION_MAP
- Batch Detail tab fetches Katana batch_stocks per variant — shows batch_number, in_stock, best_before
- Wasp Raw now includes Lot and DateCode columns from advancedinventorysearch response
- Wasp Updates tab: user fills SKU + Action (ADD/REMOVE/ADJUST dropdown) + Qty + Site + Location + Lot + DateCode, then runs "Execute Wasp Updates" from menu
- ADJUST action does NOT support lot-tracked items — must use ADD/REMOVE for those (WASP error -57041)
- OAuth scope changed from spreadsheets.currentonly to spreadsheets (needed because sheet is opened by ID in triggers)
- User must authorize the script first: open spreadsheet → Extensions → Apps Script → Run runFullSync → Authorize
- API keys stored in Script Properties: KATANA_API_KEY, WASP_API_TOKEN, WASP_INSTANCE

## Cross-Tab Scan Findings (Feb 18, 2026)
- Katana fires 3 rapid duplicates (within ~2 min) PLUS a delayed refire ~17 min later. Cache (300s/600s) catches the rapid ones but not the 17-min refire. Activity log idempotency is the permanent fix (isMOAlreadyCompleted / isPOAlreadyReceived)
- `sales_order.deleted` events always fail — Katana API returns 404 for deleted SOs. Handler tried to fetch SO details → "Failed to fetch SO details". Fix: removed from handledEvents entirely (ShipStation handles shipping per D16)
- WebhookQueue had 30+ sales_order.deleted error rows — flooding the queue. After fix, these are ignored before queuing
- F3 transfers: 0/17 successful. Root causes: STORAGE location wrong in WASP, EGG-X/TTFC/UFC insufficient stock at source, Amazon USA and Shopify are virtual locations
- F3_SKIP_LOCATIONS added: Amazon USA and Shopify transfers immediately "Skip" instead of "Failed"
- WH-8OZ-W F2 adjustments always fail — Katana rejects the stock adjustment. Needs item-sync fix (D24)
- STItems tab rows are raw error data without headers — sheet may have been recreated without header row
- F2 `{trans.Lot}` literal was from old code version — already fixed in current src, entries are historical
- StockTransfers has 3 "TEST DELETE ME" entries (ST-233, ST-225, ST-234) — need manual cleanup

## F2 Error Display Bug (Feb 18, 2026)
- `createKatanaBatchStockAdjustment` returns `{success: false, code: N, response: text}` on API error. Error info is in `.response`, NOT `.error`. `.error` is only set on network exception (catch block)
- F2 batch logging used `(result.error || '')` — always blank for Katana API rejections. Fixed to `(result.error || parseKatanaError(result.response) || 'Katana SA rejected')`
- Same `.error` vs `.response` bug pattern was fixed in 02_Handlers.gs (v111) for F1/F4/F5. F2 (03_WaspCallouts.gs) was missed because it's a separate file
- `parseKatanaError()` maps Katana JSON error responses to short messages: "Batch tracking required", "Item not in Katana", "Location error", or first 80 chars of raw message
- WASP callout `LocationCode` field contains the specific WASP location (e.g., SHIPPING-DOCK) — was captured in payload but not used in F2 logging. Now shown in Activity and flow tab
- WASP template variables `{trans.Lot}` and `{trans.DateCode}` appear as literals when items don't have lot tracking in WASP. Filtered by checking `indexOf('{') >= 0`

## Sync Dashboard Raw Tabs — Diff Formula + Conditional Formatting (Feb 19, 2026)
- Diff column uses `=IF(J{row}="",0,J{row})` formula (Katana) / `=IF(K{row}="",0,K{row})` (WASP) — auto-mirrors Adjust column values
- Formulas batch-written via `setFormulas()` on a single column range after `setValues()` for data — separator rows get empty string
- `setupConditionalRules_()` creates all color rules in 1 `setConditionalFormatRules()` call per sheet (replaces 500+ per-row formatting calls)
- Conditional formatting rule types: `whenNumberGreaterThan(0)` / `whenNumberLessThan(0)` for Adjust+Diff, `whenTextEqualTo()` for Match+LotTracked
- Per-row formatting kept only for: separator rows (whole-row colored bar), ZERO match rows (whole-row grey), Original column (grey tint)
- `clearRawSheet()` now also clears conditional formatting via `sheet.setConditionalFormatRules([])` to prevent rule accumulation on re-sync
- `stageChanges()` reads Adjust column directly (Katana: index 9 = col J, WASP: index 10 = col K) — Diff formula in adjacent column doesn't affect staging
- GAS `setFormulas()` requires 2D array of strings, empty string `''` for cells without formulas
- **Date auto-format bug**: Wasp Raw numeric columns (Original, Qty, Adjust, Diff) displayed numbers as dates (12588 → `1934-06-16`, 0 → `1899-12-30`). Root cause: column-level date format persists after `clearFormat()` (which only clears cell-level). Fix: explicit `setNumberFormat('0.##')` on cols I-L after `setValues()`. Katana Raw was unaffected (no DateCode column triggering auto-detect)

## Logging Overhaul — 6-Column Unified Format (Feb 18, 2026)
- Activity tab migrated from 4 columns (ID, Time, Flow, Details) to 6 columns (ID, Time, Flow, Details, Status, Error)
- ALL emoji icons (✅❌⚠️⊘) removed from logging — status column + conditional formatting replace them
- `logActivity()` now has 8 parameters: flow, details, status, context, subItems, linkInfo, preExecId, headerError
- Sub-items require `status` field: 'Added', 'Consumed', 'Skipped', 'Produced', 'Deducted', 'Returned', 'Failed', 'Synced', 'Staged', 'Picked', 'Complete'
- Sub-item `action` field is for detail/location info only — NO status phrases or icons
- `headerError` (8th param) auto-builds from subItems if not provided (counts skipped/failed/warnings)
- `logFlowDetail()` sub-items write status to column E and error to column F (matching Activity)
- `updateActivityRow()` reads/writes Status from column E instead of parsing emojis in column D
- Idempotency checks (`isMOAlreadyCompleted`, `isPOAlreadyReceived`) read 5 columns, check column E for status text with fallback to old emoji format
- `setupActivityConditionalFormatting()` has 16 status values: 11 green (Complete, Received, Shipped, Synced, Consumed, Produced, Added, Deducted, Returned, Staged, Picked), 1 red (Failed), 1 yellow (Partial), 1 amber (Skipped), 2 grey (Voided, Cancelled)
- F1 sub-items: status='Added', action=WASP location (PRODUCTION or RECEIVING-DOCK)
- F2 sub-items: status='Synced'/'Failed', action=WASP location + lot info
- F3 sub-items: status='Synced'/'Failed', action=fromLocation → toLocation (WASP location names, not site names)
- F4 ingredients: status='Consumed'/'Skipped'/'Failed', -46002 maps to 'Skipped' + 'Not in WASP'
- F4 output: status='Produced'/'Failed', action=output location + lot
- F5 ship: status='Deducted'/'Failed', action=SHOPIFY + bundle info
- F5 void: status='Returned'/'Failed' in flow detail, Activity row updated to 'Voided'
- F5 cancel: status='Returned'/'Failed', detail includes 'CANCELLED'
- F6 staging: status='Staged'/'Failed', action=move path + lot
- `parseWaspError()` in 07_F3_AutoPoll.gs accepts 4th param `location` for context-aware short messages: 'No stock at SW-RECEIVING'
- Flow tab header rows colored by flow (FLOW_COLORS map), status cell colored by value, error cell light red when non-empty

## Stock Reconciliation System (Feb 19, 2026)
- `reconcileToWasp()` orchestrator: fetches fresh Katana + WASP + batch data, builds target/current state maps, computes deltas, fills Adjust column, writes Wasp Updates tab
- Target state key format: `SKU|site|location` (non-lot) or `SKU|site|location|lot` (lot-tracked) — must match current state keys
- buildTargetState_ aggregates non-lot-tracked items per SKU+location FIRST, then applies SYNC_LOCATION_MAP — avoids double-counting multi-variant items
- buildCurrentState_ skips non-lot WASP rows for lot-tracked SKUs — forces lot-level matching only
- computeReconcileDeltas_ generates 3 action types: ADD (Katana has more), REMOVE (Katana has less), CREATE (SKU not in WASP at all)
- Wrong-location detection: if WASP has stock at a location that doesn't match SYNC_LOCATION_MAP for the SKU's type → REVIEW flag
- WASP ONLY items (not in Katana) are completely skipped — never modified
- Katana items with qty=0 skipped at buildTargetState_ level (only positive stock enters target)
- fillWaspRawAdjust_ does reverse-lookup: for each WASP row, finds which Katana location maps to this WASP site+location
- Wasp Updates tab format: SKU, Action, Qty, Site, Location, Lot, DateCode, Notes, Status, Error (10 cols)
- REVIEW status rows skipped by executeWaspUpdates() — user must manually clear REVIEW to approve
- CREATE action in executeWaspUpdates() posts to `ic/item/create` with non-array payload (unlike transaction endpoints)
- Timer (setupReconcileTimer) only refreshes the delta view — never auto-executes writes to WASP
- SpreadsheetApp.getUi() throws when called from time-based triggers — wrap all ui.alert() in try/catch for dual-context compatibility (menu + timer)
- GAS date auto-format bug: numeric columns can render as dates (e.g., 12588 → 1934-06-16). setNumberFormat('0.##') prevents it, but getValues() may return Date objects from already-corrupted cells. Always defend with `instanceof Date` check before parseFloat.
- O(n*m) SKU existence checks killed by pre-building a waspSkuSet object for O(1) lookups

## 4-Tab Rewrite (Feb 19)
- onEdit simple trigger fires ONLY for user UI edits, NOT for script setValues() — perfect for change detection
- onEdit cannot access UrlFetchApp or ScriptProperties — only basic spreadsheet operations
- CRITICAL: Simple onEdit triggers ONLY work on bound scripts (created from within the spreadsheet). The tracker GAS project is a standalone web app (doPost/doGet) — must use installable triggers via ScriptApp.newTrigger('handler').forSpreadsheet(ss).onEdit().create() for edit detection on the debug spreadsheet (Feb 26)
- Hidden columns (width=1 + hideColumns) still readable by script via getValues() — good for storing Orig Site/Orig Location
- Pending row preservation pattern: read existing Status before clearing sheet, build pendingMap keyed by SKU|Site|Location|Lot, restore after writing new data
- Wasp tab separator labels include counts: "MMH Kelowna > SHOPIFY (45 items, 125,000 units)" — computed before building finalRows
- pushToWasp reads ALL backgrounds via getBackgrounds(), modifies in-memory, writes back in batch — avoids per-row setBackground calls
- Status conditional formatting uses whenTextContains('Synced') not whenTextEqualTo because status includes timestamp: "Synced 2026-02-19 12:30"
- autoCreateMissingItems determines lot tracking from batchData.lotTrackedSkus (copied from Katana batch_tracking)
- WASP item create endpoint: `ic/item/create` with non-array payload, includes LotTracking boolean field
- 05_SyncActions.gs and 05_WaspUpdater.gs replaced by 05_PushEngine.gs — remove from GAS project on deploy
- Wasp tab expanded to 15 cols (was 14): col O = Orig Lot Tracked (hidden). Needed for detecting lot tracking changes on push.
- onEdit expanded to all editable columns via ORIG_COL_MAP pattern: `{2:null, 3:11, 4:12, 5:15, 10:9}` — null means "no orig, always Pending"
- Multi-column pending clear: when user restores a column to original, must check ALL tracked columns before clearing Pending status
- Manifest system for deleted row detection: writeWaspTab saves compact JSON to ScriptProperties (chunked 8KB), pushToWasp compares sheet keys vs manifest, generates REMOVE for missing rows
- ScriptProperties 9KB per-value limit: manifest chunked into WASP_MANIFEST_0, _1, etc. at 8000 chars each. ~300 items = ~3 chunks.
- Pending preservation fix: pendingMap key must use ORIG site/location (not user-edited values) so fresh data rows match the key
- WASP item update endpoint: `ic/item/update` with ItemNumber + LotTracking fields (needs verification — may differ from actual API)
- GAS ORIG_COL_MAP in onEdit: uses object lookup `hasOwnProperty(col)` for O(1) column filter instead of chained `!==` comparisons
- Batch tab 18-col layout: SKU(1), K.Location(2), K.Qty(3), K.UOM(4), K.Batch(5), K.Expiry(6), W.Site(7), W.Location(8), W.Qty(9), W.UOM(10), W.Lot(11), W.DateCode(12), Diff(13), Status(14), Orig W.Qty(15 hidden), Orig W.Site(16 hidden), Orig W.Location(17 hidden), Sync Status(18)
- Batch tab editable cols: W.Site(7)→Orig(16), W.Location(8)→Orig(17), W.Qty(9)→Orig(15). onEdit config: `origColMap:{7:16, 8:17, 9:15}`, trackedPairs: `[[9,15],[7,16],[8,17]]`
- Batch rows are always lot-tracked — all push actions include lot+dateCode from W.Lot(11) and W.DateCode(12)
- pushBatchPendingRows_ mirrors pushToWasp pattern: read 18 cols, find Pending in col 18 (idx 17), compute actions, execute, update Sync Status + orig values + backgrounds
- Batch push "new row" case: user fills W.Site/W.Location/W.Qty on a KATANA ONLY row (orig cols empty) → ADD action with lot from W.Lot column
- Batch pending preservation key: `SKU|W.Lot|W.Site|W.Location` (lot is part of key since batch tab is per-lot)
- WASP error -57041 means "Lot is missing. Date Code is missing." — item is lot-tracked in WASP but API call didn't include lot/dateCode
- -57041 retry handler in executeAction_: ADD retries with lot='NO-LOT' + dateCode=2yrs future; REMOVE looks up actual lot from WASP via advancedinventorysearch then retries
- Scenario: user enables lot tracking in WASP admin, resyncs, then edits location/qty → push fails with -57041 because sheet row has no lot info. Retry handler auto-fixes.
- WASP lotTracking field: added to fetchAllWaspData() item mapping with fallbacks: LotTracking, TrackLot, IsLotTracked, LotTrackingEnabled. Augments batchData.lotTrackedSkus alongside existing lot-based detection.
- CRITICAL: `item.LotTracking || item.TrackLot || ... || false` evaluates to `false` (not `undefined`) when none of the fields exist in the API response. Then `writeWaspTab` trusts `false` over the `lotTrackedSkus` heuristic because `false !== undefined` is true. Fix: compute `ltVal` via hasOwnProperty-style checks, leave as `undefined` when field absent.
- WASP `advancedinventorysearch` endpoint returns INVENTORY rows (qty, site, location, lot) but does NOT include item-level settings like LotTracking. Separate item search endpoints (`ic/item/search`, `ic/item/itemsearch`, `ic/item/detailsearch`) all fail (404 or empty). Best heuristic: if any WASP inventory row for a SKU has a `lot` value → that SKU is lot-tracked.
- DELETE handler dateCode field needs same `instanceof Date` handling as all other dateCode fields — Google Sheets `getValues()` returns Date objects for date-formatted cells.
- `replace_all` for fixing lot gating (`isLotTracked ? lot : ''` → `lot`) only catches exact variable name matches. The DELETE section used `delIsLotTracked ? delLot : ''` which survived because different variable names.

## WASP ic/item/create 404 Investigation (Feb 20, 2026)
- Previous handoff analysis was WRONG: Python sync_all_v5.py line 52 uses `f"{self.base_url}/public-api/{endpoint}"` — DOES include /public-api/
- Both Python and GAS produce identical URL: `https://mymagichealer.waspinventorycloud.com/public-api/ic/item/create`
- IIS 404 can be caused by ASP.NET WebAPI model binding failure when payload contains unexpected fields — the server can't find a matching action method
- GAS sends fields Python doesn't: SiteName, LocationCode, LotTracking, TrackDateCode, StockingUnit
- Python sends fields GAS doesn't: Category, Taxable, Price, ListPrice, AltItemNumber, Notes
- Diagnostic test function `testWaspCreateEndpoint()` tests 6 combinations to isolate the cause
- `waspFetch` vs `waspApiCall`: both use POST + JSON + Bearer token, only difference is 429 retry logic in waspFetch (shouldn't affect 404)
- **CONFIRMED (Feb 20 diagnostic)**: `ic/item/create` endpoint REMOVED from WASP API. ALL payload variants return 404. Not payload-related — bare 3-field payload also fails. Endpoint worked Feb 13 (Python sync), gone by Feb 20.
- Other ic/ endpoints still work: `ic/item/advancedinventorysearch` (200 JSON), `ic/item/delete` (exists)
- Without /public-api/ prefix: WASP returns HTTP 200 with HTML "Resource not found" page (web app, not API) — a false positive
- `discoverWaspCreateEndpoint()` probes 19 candidate endpoint names to find replacement
- Fallback options: contact WASP support, CSV import (12-Inventory format), or manual creation in WASP UI

## WASP Item Create — Correct Endpoint (Feb 24, 2026)
- **RESOLVED**: `ic/item/create` was never the right endpoint — the correct one is `ic/item/createInventoryItems`
- Documentation: https://help.waspinventorycloud.com/Help/Details?apiId=POST-public-api-ic-item-createInventoryItems
- Payload is an ARRAY of `ItemInventoryInfo` objects (like transaction endpoints)
- Key fields: `ItemNumber`, `ItemDescription`, `StockingUnit`
- Lot tracking: `TrackbyInfo: { TrackedByLot: true, TrackedByDateCode: true }` (not flat boolean)
- Location assignment: `ItemLocationSettings: [{ SiteName, LocationCode, PrimaryLocation: true }]`
- Site assignment: `ItemSiteSettings: [{ SiteName }]`
- Also available: `ic/item/updateInventoryItems` (POST, array, same pattern)
- Also available: `ic/item/getinventoryitem/{itemNumber}` for single item lookup
- Also available: `ic/item/getmultipleinventoryitem` for batch lookup (up to 20)
- `waspCreateItem()` in 04_SyncEngine rewritten to use correct endpoint + payload format
- **StockingUnit must use WASP abbreviation** (not full name): EA, dz, in, cu.in, lb, KG, PCS, BOX, g, pack, ROLLS
- 'Each' does NOT work — must send 'EA'. 'PK' does NOT work — must send 'pack'
- TaxCode 'None' does NOT exist — omit TaxCode field entirely (WASP defaults)
- CostMethod expects numeric enum (e.g. 10) not string 'Average' — omit field (WASP defaults)
- Minimum working payload: ItemNumber + ItemDescription + StockingUnit + PurchaseUnit + SalesUnit (all 3 UOM fields)
- `createWaspCatalogItem_()` added inline in pushKatanaSyncRows_ (05_PushEngine) for reliability

## Lot Tracking Reference Tab (Feb 20, 2026)
- `writeLotTrackingTab()` in 04_SyncEngine.gs creates a "Lot Tracking" tab comparing Katana vs WASP lot settings
- Katana lot tracking determined from TWO sources: product/material `batch_tracking` API field + `batch_stocks` record existence (heuristic for items with active batches)
- `fetchWaspItemLotSettings()` tries 3 endpoints: `ic/item/search`, `ic/item/itemsearch`, `ic/item/detailsearch` — returns per-SKU LotTracking boolean
- Status values: MATCH (green), MISMATCH (red), NOT IN WASP (orange), WASP ONLY (blue)
- Katana DateCode column = same as Katana Lot — WASP TrackDateCode should always match LotTracking for Katana items
- WASP `ic/item/update` can change LotTracking per-item: `{ItemNumber: sku, LotTracking: true/false}`
- Function makes extra Katana API calls (products/materials/variants) to get `batch_tracking` field directly — slower but more accurate than batch_stocks-only heuristic

## Lot Tracking Source of Truth Fix (Feb 20, 2026)
- `lotTrackedSkus` was unreliable: built from batch_stocks (only items with active batches), then WASP could DELETE Katana entries via `delete batchData.lotTrackedSkus[sku]`
- Items with batch_tracking=Yes in Katana but 0 stock showed as "No" on Wasp tab (e.g., FBA-FRAGILE, FGJ-US-1, FGJ-US-2, VB-WITCH-HAZEL-TB)
- Fix: `fetchAllKatanaData()` now extracts `batch_tracking` from products/materials API → builds `batchTrackingSkus` map → returned in result
- `runFullSync()` now merges `katanaData.batchTrackingSkus` into `batchData.lotTrackedSkus` FIRST (before any WASP data)
- Removed the WASP override that deleted Katana entries (`delete batchData.lotTrackedSkus[sku]`) — Katana is always source of truth
- `advancedinventorysearch` does NOT return item-level LotTracking field — `fetchWaspItemLotSettings()` endpoints may all fail (404/empty)
- Katana product fields tried: `batch_tracking`, `batch_tracked`, `is_batch_tracked`, `track_batches` — defensive fallback chain

## WASP Stock Sync XLSX Script (Feb 20, 2026)
- Script: `scripts/wasp-stock-sync.py` — reads xlsx, rebuilds WASP tab from Katana, marks actions
- Skill: `claude-skills/integrations/wasp-stock-sync.md` — usage reference
- Batch tab TOTAL rows MUST be filtered: check `status contains 'TOTAL'` AND `K.Location is empty` — both needed, either alone misses cases
- Batch tab TOTAL rows caused doubled quantities (19 SKUs exactly 2x Katana) — the TOTAL summary rows had qty values that passed through without the filter
- Lot-tracked WASP aggregate rows (no lot number) must be replaced with per-batch rows from Katana Batch tab — 14 SKUs needed this breakdown (LUP-1 had 12 batches)
- NO-LOT placeholder: lot-tracked items without Katana batch data use 'NO-LOT' as lot value (INSTRIO, UFC-1OZ pattern)
- Duplicate aggregate rows possible (LUY-2 had TWO aggregate rows in WASP) — deduplicate by tracking replaced SKUs
- Action markers in col Q: ADD NEW (green #92D050), ADJUST (yellow #FFFF00), REMOVE (red #FF6B6B), blank = no action
- REMOVE section at bottom of WASP tab: old aggregate rows that were replaced by batch rows — user must delete these from real WASP
- Batch tab sync: fill W side for KATANA ONLY rows (we're adding to WASP), mark REMOVED for WASP ONLY aggregate rows
- None sort crash: `sorted(batches, key=lambda x: x['loc'])` crashes if loc is None — use `str(x['loc'] or '')`
- Negative Katana qty (e.g., B-GREEN-1 at -78): preserved as-is, indicates Katana data issue (overconsumption)
- WASP tab structure: 16 data cols (A:SKU through P:OrigDateCode) + Q:Adjusted marker
- Batch tab structure: 20 cols (A:SKU through T:Orig W.DateCode)
- Separator rows group by Site > Location with item count and total qty
- Group order: PRODUCTION → UNSORTED → SW-STORAGE (then any extras)
- Auto-backup created before any modifications: `filename - backup.xlsx`
- Items only in Katana but not WASP: skip (user says "doesn't matter")
- Cross-verification: sum per-SKU totals and compare Katana vs WASP — must show ALL MATCH

## Mismatch Column Bug — String Concatenation (Feb 23, 2026)
- ROOT CAUSE: Katana API returns quantity_in_stock as STRING like "3800.00000000000000000000" — code stores as-is without parseFloat
- buildTotalBySku_ does `map[sku] += (item[qtyField] || 0)` — JavaScript string concatenation: `0 + "3800" = "03800"` (string, not number!)
- Multi-location SKUs get garbage: `"03800.000000000000000000000.0000000000"` → NaN in Math.round → always MISMATCH
- Single-location SKUs worked by accident: `"03800.00..."` is a valid numeric string when coerced by `*100`
- WASP side was fine because fetchAllWaspData already uses parseFloat() for TotalAvailable/TotalInHouse
- Fix #1 (critical): `map[sku] += (parseFloat(item[qtyField]) || 0)` in buildTotalBySku_ (04b_RawTabs.gs)
- Fix #2 (defensive): `var rawQty = parseFloat(inv.quantity_in_stock) || 0` in fetchAllKatanaData (02_KatanaAPI.gs)
- Fix #3 (lot double-counting): filterWaspItemsForTotals_() removes aggregate WASP rows when per-lot rows exist
- All 3 fixes applied to local engine-src files via PowerShell (UTF-16 encoding handled)

## Per-Row Match Logic (Feb 23)
- Old Match column compared SKU-total Katana qty vs SKU-total WASP qty — showed MISMATCH for O-OIL lot 4300-6022 (3800=3800) because OTHER lots for same SKU had differences
- New approach: lot-tracked rows compare per-lot (SKU+batch → qty), non-lot rows keep SKU-total comparison
- `computeRowMatch_()` added as wrapper around `computeCrossMatch_()` — delegates to per-lot for batch rows, SKU-total for non-batch
- WASP Qty column added to Katana tab (col J) showing the comparison number — lot rows show WASP per-lot qty, non-lot rows show WASP SKU total
- Katana tab went from 11 → 12 columns: SKU|Name|Location|Type|LotTracked|Batch|Expiry|UOM|Qty|**WASP Qty**|Match|WASP Status
- `waspBySkuLot` lookup (Katana tab): built from waspData.items, key=SKU|lot, value=sum of qtyAvailable
- `katanaBySkuLot` lookup (WASP tab): built from batchData.batchesBySku, key=SKU|batchNumber, value=sum of batch.qty
- Column shift required updates in: 04b_RawTabs.txt (writeKatanaTab), 01_Code.txt (onEdit, clearErrorsToPush), 05_PushEngine.txt (pushKatanaSyncRows_)
- Encoding: engine-src files are UTF-16 LE — must convert to UTF-8 for Edit tool, then back to UTF-16 after
- Non-lot items also need per-row matching: Katana location = WASP site name, so per-site comparison works
- `waspBySkuSite` lookup: built from filterWaspItemsForTotals_(waspData.items), key=SKU|siteName, avoids double-counting
- `katanaBySkuSite` lookup: built from katanaData.items, key=SKU|locationName
- WASP tab per-site: row.site maps directly to Katana location name (MMH Kelowna, Storage Warehouse, Shopify)
- Three comparison levels: per-lot (batch rows), per-site (non-lot with location), SKU-total (fallback)

## Debug Log Fixes (Feb 23, 2026)
- Exec ID duplicate bug: `getNextExecId()` used `getScriptLock()` — same lock type as ShipStation handler's outer lock. Inner `finally { lock.releaseLock() }` released the OUTER lock, causing concurrent webhooks to overlap. Fixed by switching to `getDocumentLock()` in 08_Logging.gs
- Smart MO retry: `getMOConsumedMap(moRef)` / `saveMOConsumedMap(moRef, map)` in ScriptProperties tracks consumed ingredient SKUs per MO. On retry, skips already-consumed items (status "Skip-OK"). Map uses `{ sku: count }` to handle duplicate ingredients correctly. Cleaned up on full success.
- Batch diagnostic logging: when `extractIngredientBatchNumber(ing)` returns null, logs `F4_BATCH_DIAG` with raw ingredient keys, batch_transactions presence/length, and first 400 chars of raw batch data. This will reveal the actual Katana API response structure for `&include=batch_transactions`.
- F5 Error column fix: `logActivity()` auto-builds headerError from sub-items, but F5 was only passing sub-items when `itemResults.length > 1`. Changed to always pass `itemResults` so single-item failures get Error column populated.
- F5 itemCount: flow detail tab now uses `itemResults.length` (post-bundle-expansion count) instead of `shipmentItems.length` (pre-expansion). Header still uses `shipmentItems.length` since that matches the user-visible order.
- F5 headerDetails: added error count suffix (`2 errors`) when `failCount > 0`, matching F4's pattern for consistency across flows.
- Katana batch_transactions field: API returns `{ batch_id, quantity }` only — NO `batch_number` field. Must resolve batch_id via `GET /batch_stocks/{batch_id}` to get the actual batch number string.
- `fetchKatanaBatchStock(batchId)` added to 04_KatanaAPI.gs — single batch stock lookup by ID, returns batch_number + expiration_date (NOT expiry_date).
- Katana DELETES depleted batch records permanently — `batch_stocks/{id}` returns 404 for consumed batches, even with `include_deleted=true`. Cannot resolve batch_number from Katana after full consumption.
- Webhook timing confirmed: `manufacturing_order.done` fires AFTER user assigns batch in Katana popup (not before). So batch_transactions IS populated — the issue was purely field name resolution.
- WASP `advancedinventorysearch` endpoint returns inventory rows WITH Lot and DateCode fields — this is the correct endpoint for WASP lot lookups (NOT `infosearch` which only returns item metadata).
- `waspLookupItemLots()` rewritten to use `advancedinventorysearch` — matches ItemNumber + LocationCode, returns first lot found. Old `infosearch` approach never returned lot data.
- F4 batch resolution fallback chain: (1) Katana batch_transactions → batch_stocks/{id} (works for non-depleted), (2) batch_stocks with include_deleted (doesn't work — Katana permanently deletes), (3) WASP advancedinventorysearch lot lookup (works — finds lot at PRODUCTION), (4) WASP -57041 retry with lot lookup (existing fallback).
- Katana batch_stocks field name: `expiration_date` (NOT `expiry_date` or `best_before_date`) — must check all three in extraction code.
- **CRITICAL**: WASP remove for lot-tracked items requires BOTH Lot AND DateCode — `waspRemoveInventoryWithLot()` only adds DateCode to payload if the 7th param is provided. The proactive fallback and -57041 retry were using `waspLookupItemLots()` (returns lot string only) instead of `waspLookupItemLotAndDate()` (returns `{lot, dateCode}`). This was the root cause of persistent -57041 failures even after lot lookup was working. Fixed Feb 23.
- **CRITICAL**: WASP `advancedinventorysearch` SearchPattern does NOT filter by ItemNumber. It returns ALL inventory rows in sequential order regardless of the search term. LUP-4 was found only because it happened to be in the first 100 rows. LUP-1 was NOT found on page 1 (items like GP-50-W, CAR-2OZ, B-PURPLE-2 returned instead). Fix: paginate through all pages with client-side ItemNumber filtering (max 20 pages safety limit). Fixed Feb 23.
- Katana API pagination uses `limit` parameter (NOT `per_page`), max 250 per page. No `order_no` filter on manufacturing_orders — must list and match client-side.
- Idempotency hardened: `isMOAlreadyCompleted()` and `isPOAlreadyReceived()` now block ANY existing entry (Complete/Partial/Failed), not just success. No more duplicate processing from Katana re-webhooks.

## Katana Tab Qty Editing + Adjustments Log (Feb 23, 2026)
- Katana tab expanded from 12 to 14 columns: col 13 = Orig Qty (hidden), col 14 = Adj Status
- onEdit handles BOTH col 12 (WASP Status "Push") and col 9 (Qty edits for adjustments) independently — two separate code paths in the same Katana handler
- Adj Status values: Pending (yellow), Pushed (green), ERROR (red) — mirrors Wasp/Batch Sync Status pattern
- pushKatanaAdjustments_ computes delta = editedQty - waspQty (col 10), not editedQty - origQty — ensures WASP ends up at the correct absolute value
- Smart location routing in pushKatanaAdjustments_: Storage Warehouse→SW-STORAGE, Material/Intermediate→PRODUCTION, else→UNSORTED
- logAdjustment_ auto-creates "Adjustments" tab on first call with headers, conditional formatting, and column widths — no setup function needed
- logAdjustment_ now writes to DEBUG_SHEET_ID (via ScriptProperties 'DEBUG_SHEET_ID') — same sheet as webhook handlers. Falls back to active spreadsheet if not set.
- Adjustments tab: 12-col layout — Timestamp, Source, User, SKU, Item Name, Location, Lot/Batch, Old Qty, New Qty, Diff, Status, Note
- Sources: "WASP" (callout/detected), "Katana" (detected), "Google Sheet" (sheet pushes) — full row coloring per source
- Only manual adjustments logged — F1/F3/F4/F5 operational flows removed from adjustment logging
- User field: Google Sheet = Session.getActiveUser().getEmail(); WASP = payload.UserName (needs callout template update); Katana = empty (detected)
- logAdjustment_ called from 3 push flows: Wasp push (after successful row push), Batch push (after successful row push), Katana adj (on push success, error, and no-delta skip)
- Katana adj rows with no delta (Math.abs(editedQty - waspQty) <= 0.01) are silently cleared (Adj Status='', restore background) — no push, no log
- clearErrorsToPush updated to handle 14 cols: checks both col 11 (WASP Status ERROR→Push) and col 13 (Adj Status ERROR→Pending)
- adjPendingMap in writeKatanaTab: preserves Pending/ERROR adj rows across hourly sync, keyed by SKU|Location|Batch
- UTF-16LE file editing: PowerShell scripts needed when Edit tool can't match exact whitespace patterns in UTF-16 encoded files. Write .ps1 script to disk, execute with `powershell -File`, avoids $variable interpolation issues in inline commands
- Engine-src files (04b_RawTabs.txt, 05_PushEngine.txt) are UTF-16LE (BOM: 0xFF 0xFE); 01_Code.txt is plain UTF-8

## Qty Coloring + Lot Display Enhancement (Feb 23, 2026)
- RichTextValue `setTextStyle(start, end, style)` can color a substring within a cell after `setValues()` writes the full text
- Must find the "x{qty}" substring position with `indexOf()` before applying — position depends on SKU name length + prefix characters
- `qtyColor` field on sub-items: 'green' (#008000) for additions, 'red' (#cc0000) for removals. Neutral transfers (F3) have no color.
- Applied in both `logActivity()` (Activity tab col D) and `logFlowDetail()` (flow tab col D) — same pattern, different variable names
- F5 lot fallback: when WASP returns -57041 on processShipment, looks up lot via `waspLookupItemLotAndDate()` and retries with `waspRemoveInventoryWithLot()`
- F1 PO receiving: lot number is the PO number (sanitized) — now shown in Activity sub-items as `lot:PO-XXX`
- F3 transfers: lot + expiry now shown in Activity sub-items (was only in flow detail tab). Uses `successItems[s].lot` and `.expiry` fields
- F4 ingredients: qtyColor='red' (consumed), output: qtyColor='green' (produced). Both Activity and flow detail sub-items.
- F5 ship: qtyColor='red' (deducted from SHOPIFY). Void: qtyColor='green' (returned to stock).
- F2 add: qtyColor='green', F2 remove: qtyColor='red'. Applied to both Activity and flow detail sub-items.
- **F4 multi-lot ingredients**: Katana `batch_transactions` can have MULTIPLE entries (e.g., EGG-X: 9 from HID260213 + 191 from HID260218). `extractIngredientBatchNumber(ing)` only returns `bt[0]` — the first lot. Sending total qty (200) to one lot fails with -46002 (insufficient). Fix: when `bt.length > 1`, loop each batch_transaction with its own `quantity` and resolved lot/expiry, making separate WASP removal calls per batch allocation.

## WASP Item Create — Full Payload (Feb 24, 2026)
- **DimensionInfo required**: WASP validates dimension UOMs on create — sending empty defaults causes -57066 "dimension unit of measure does not exist". Fix: always include `DimensionInfo: { DimensionUnit: 'in', WeightUnit: 'lb', VolumeUnit: 'cu.in' }`
- **Cost from Katana**: `purchase_price` field on Katana variants — use as WASP `Cost`. Katana API: `/variants` has `purchase_price`, `/products` has `category`, `/materials` has `category`
- **CategoryDescription**: WASP accepts category as string (e.g. "NON-INVENTORY", "FINISHED GOOD") via `CategoryDescription` field on createInventoryItems
- **catalogOnly flag**: Was hardcoded `true` in pushKatanaSyncRows_ — must be `false` for Push items to also ADD stock with batch/lot tracking after CREATE
- **Item already exists**: CREATE may fail for existing items. Check error for "already exist" or "-57" codes and still proceed to ADD stock

## Manual Adjustments (Feb 24, 2026)
- Created `13_Adjustments.gs` — processes manual stock corrections from "Adjustments" sheet in debug spreadsheet
- Sheet layout: Action (ADD/REMOVE/ADJUST) | SKU | Qty | Site | Location | Lot | Expiry | Notes | Status | Error
- `processAdjustments()` reads PENDING rows, calls WASP API, updates Status, logs to Activity as F2
- Uses `markSyncedToWasp()` to prevent echo loops (WASP callout fires back after adjustment)
- `addAdjustmentRow()` helper for programmatic retry scenarios
- Added 13_Adjustments.gs to push-to-apps-script.js file list (also added 10_, 11_, 12_)
- `parseWaspError()` lives in 07_F3_AutoPoll.gs — reused by adjustments, no duplicate needed
- `cleanErrorMessage()` lives in 08_Logging.gs — maps WASP error codes to human-readable text

## Auto-Sync New Items (Feb 26)
- engin-src/ files (Command Center sheet) now auto-sync new Katana items to WASP during hourly runFullSync()
- Excluded categories (stay as NEW): DEVICES, DIGITAL PRODUCT, TOOLS/HARDWARE, EQUIPMENT
- Category flows: Katana product/material → productMap.category → variantMap.category → items[].category → writeKatanaTab rows
- AUTO_SYNC_SKIP_CATEGORIES defined in 04b_RawTabs.txt — edit this array to change exclusions
- pushKatanaSyncRows_(ss) called automatically after writeKatanaTab in runFullSync()
- push-to-apps-script.js path bug fixed: was reading from scripts/ instead of src/ (caused all deploys to fail)

## Systems Health Tracker (Feb–Mar 2026)
- Project folder: systems-health-scr/ — clasp project for 2026_Systems-Health-Tracker GAS
- Script ID: 1TBkee_JgNKnHxxeCbWSp3uJyh5Wg5o_GLowbp0k51DP98GjzzDWN4eS5
- GAS V8 CRITICAL: var only, no const/let, no arrow functions, no template literals — silent load failures otherwise
- sheet-setup.js owns `setupHealthSheet()` and all chip/status styling helpers
- health-checks.js defines menu ("Health Monitor"), URL checks, and heartbeat receive handler
- it-security-tracker.js: slack functions prefixed `secSend*` to avoid collision; no onOpen (health-checks.js owns it)
- `sheet.clear()` does NOT remove data validations in GAS — must call `.clearDataValidations()` explicitly
- Heartbeat helper `sendHeartbeatToMonitor_(name, status, details)` added to: engin-src/01_Code.js, src/09_Utils.gs
- HEALTH_MONITOR_URL must be set as Script Property in each script that sends heartbeats

## Shopify Flow Heartbeats (Mar 2026)
- "Send HTTP Request" is a native Shopify Flow action — no third-party needed
- Only available on Shopify Grow, Advanced, or Plus plans (NOT Basic)
- Supports POST, custom headers, Liquid-templated JSON body ({{order.name}}, {{order.email}}, etc.)
- Timeout: 30s — Apps Script doPost must respond within 30s
- Conditional flows (condition gates actions): add HTTP step on BOTH Then AND Otherwise branches — use condition_met:true/false in body
- Non-conditional flows (specific event trigger): add HTTP step as FIRST action (before any other steps)
- Setup guide: systems-health-scr/docs/shopify-flow-heartbeat-guide.txt
- 6 flows tracked in System Registry: 3 internal (fraud/promo), 3 ShipStation API flows
