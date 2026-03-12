# Browser Agent Execution Prompt

You are making the Shopify Flow changes now. Work carefully in the Shopify admin UI. Change only the gift-related workflows described below. Capture screenshots before and after every edited flow.

## Objective

Implement the minimum safe fix set for gift-item workflows so that:

- Shop App orders are not processed by the wrong general gift flow
- gift insertion does not happen before fraud or hold review clears
- Shop App 4oz eligibility is based on qualifying SKU quantity, not total order quantity
- eligible orders get the gift once
- ineligible, high-risk, and held orders do not get the gift

## Flows In Scope

1. `Add gift when 2+ 4oz products ordered`
2. `Shop App - Gift Trigger - Bundles`
3. `Shop App - Gift Trigger - 4oz Jars`
4. `Shopify High-Risk Order Hold - Slack Alert`

Also inspect any helper flow you create or edit for post-review gift gating.

## Operating Rules

- Make changes one flow at a time.
- Take a screenshot before editing each flow.
- Take a screenshot after saving each flow.
- If Shopify Flow UI limitations prevent a clean implementation, stop and document the exact blocker.
- Do not assume a field is equivalent unless the UI label clearly matches.
- Do not leave two active flows able to add the same gift variant to the same order.

## Required End State

There must be one safe path to add the gift.

That path must enforce all of the following before `Add order line item`:

- correct sales channel logic
- correct qualifying SKU quantity logic
- order is not high-risk
- order is not on hold
- gift variant is not already on the order

If the UI makes a perfect single-path design impractical, use the safest achievable version and document the compromise.

## Execution Steps

### Step 1: Record the current state

For each in-scope flow:

- open the flow
- capture full screenshots of trigger, conditions, and actions
- record whether it is Active

### Step 2: Remove the incorrect Shop App exclusion

Open `Add gift when 2+ 4oz products ordered`.

Replace the first condition:

- remove `shop.name is not equal to Shop App`
- use `order.channelInformation.app.title is not equal to Shop App`

If the field path is not available in that condition builder:

- stop and report the blocker
- do not substitute another field without evidence

### Step 3: Eliminate overlapping gift ownership

Preferred implementation:

- only one active flow path can add `Universal Flare Care Gift - 1oz`

Use this decision rule:

- if a single consolidated post-review gift flow is feasible in Shopify Flow, build that and disable direct gift insertion in the overlapping source-specific flows
- if consolidation is not feasible quickly, keep the fewest possible active gift-insertion flows and ensure their eligibility scopes do not overlap

Minimum acceptable outcome:

- `Add gift when 2+ 4oz products ordered` must not be able to add the same gift to Shop App orders
- Shop App gift flows must not overlap each other for the same order type

### Step 4: Move gift insertion behind review clearance

Implement gift insertion so it occurs only after fraud review has had a chance to complete.

Preferred design:

- trigger the gift decision from `Order risk analyzed`
- require risk is not high
- require order is not on hold

If needed, use an intermediate tag/state approach:

- source-specific eligibility flow adds a neutral marker only
- a post-risk flow checks the marker, verifies safety conditions, then adds the gift

Do not keep the current design where gifts are directly inserted on raw `Order created` without a review guard.

### Step 5: Fix Shop App 4oz quantity logic

Open `Shop App - Gift Trigger - 4oz Jars`.

Replace the total-order quantity logic:

- remove `order.currentSubtotalLineItemsQuantity >= 2`
- add a `Run code` step or equivalent logic that counts only the qualifying 4oz variant IDs
- require qualifying quantity `>= 2`

Reuse the variant-counting approach from `Add gift when 2+ 4oz products ordered` if possible.

### Step 6: Add explicit duplicate protection

Before `Add order line item`, add the strongest available duplicate guard:

- check whether the gift variant is already present on the order

If Shopify Flow does not expose a clean "line item contains variant" check, add a marker after insertion:

- tag or metafield such as `gift_added`

If using a marker:

- also add a pre-check that the marker is not already present

Document clearly if this is only a partial race-condition mitigation.

### Step 7: Verify free-gift behavior

Inspect the product or action settings for `Universal Flare Care Gift - 1oz`.

Determine whether the variant is priced at `$0.00`.

If the variant is not `$0.00`:

- enable `Add for free` on every gift insertion action

If the variant is already `$0.00`:

- note that in the report

### Step 8: Save and document final state

For every edited flow:

- save the flow
- capture final screenshots
- record exact trigger and condition fields in the final version

## Validation Tasks After Edits

If test order creation is available, run these tests. If not, state that execution testing could not be completed from the browser session.

1. Non-Shop-App order with 2 qualifying 4oz items
   Expected: gift added once

2. Shop App order with 2 qualifying 4oz items
   Expected: gift added once

3. Shop App order with 1 qualifying 4oz item and 1 unrelated item
   Expected: no gift

4. Eligible Shop App bundle order
   Expected: gift added once and free

5. High-risk order meeting gift criteria
   Expected: no gift

6. Held order meeting gift criteria
   Expected: no gift while hold remains

7. Duplicate-path regression check
   Expected: only one gift line item, not two

For each test, collect:

- order ID
- channel/source
- risk outcome
- hold status
- resulting line items
- tags or markers
- relevant flow run history

## Output Format

Return a concise implementation report with these sections:

1. Changes made
2. Final flow ownership model
3. Any blockers or compromises
4. Validation results
5. Remaining risks

## Stop Conditions

Stop and report immediately if:

- the required channel field is unavailable where needed
- the post-review design cannot be implemented cleanly in Shopify Flow
- another undisclosed active flow also adds the same gift variant
- the gift variant pricing cannot be verified
