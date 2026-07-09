# Bed Ministry — Inventory Management System — Changelog

## v3.6 (2026-07-05) — Starter kit: replication & per-city flexibility
Everything a new city needs to stand up their own instance, with zero behavior
change for a workbook that doesn't opt in (all new config defaults to current
behavior; San Antonio remains on v3.5 until it chooses to upgrade).

**Form automation (kills the #1 replication fragility):**
- `FORM_SCHEMA` — the exact 46-question request-form schema, reconciled against
  the live San Antonio form (captured via "Dump live form schema to log" on a
  private copy) and then genericized: SA-specific help text (a ministry email,
  a local area code) replaced with neutral wording. Carries page structure,
  help text, required flags, and the per-child "No → skip to Caregiver"
  navigation. The 7 "Need to enter another child?" items intentionally share
  one identical title (matching how Google builds the form); the response sheet
  is what disambiguates them as "… 2", "… 3".
- Setup > "★ Create & link Bed Request Form" (`setupForm`) builds the Google
  Form from `FORM_SCHEMA` — including the per-child page flow ("Need to enter
  another child?" No → jump to Contact & Delivery) — and links it to the
  workbook. Refuses to run when a form is already linked.
- Health check now verifies the linked form's questions against `FORM_SCHEMA`
  (same normalization `onFormSubmit` uses), so a renamed question becomes a
  visible warning instead of a silently dropped field.

**Ministry Options (per-city flexibility):**
- New Config block (columns D/E, rows 10–22): ministry name + YES/NO toggles
  for Cribs, Toddler/Twin/Bunk beds, Bears, Books, Plaques, Bedrails, Themed
  Bedding. Blank = YES, so existing workbooks behave exactly as before.
- Disabled bed types are dropped from the generated form; disabled extras
  forecast 0 demand, default to 0 on pick lists, are hidden from the pick-list
  dialog, and are greyed on Inventory Position. Columns are never removed —
  flipping a toggle back on later needs no migration (run "Refresh Ministry
  Options" after any change).
- Pick-list PDF title now uses the ministry name from Config (blank = "BED
  MINISTRY"). San Antonio sets "Child of God Bed Ministry" there when it
  adopts v3.6.
- NOTE: catalog item names are NOT safely renameable (shortage formulas
  pattern-match names; pick lists render the code-side catalog). Treat the 39
  slots as fixed roles; the Launch Guide documents this.

**One-click setup / New City wizard:**
- Setup > "★ Set up EVERYTHING" (`setupEverything`) runs the full setup chain
  in the documented order with progress toasts — including `ensureCoreSheets_`,
  which creates all core sheets with base headers and a generic Dashboard
  scaffold when starting from a BLANK workbook (existing sheets are never
  restructured).
- "🚀 New City Setup" (`newCityWizard`) — for a church that just copied the
  template: prompts for ministry name, alert emails, and options; creates the
  form if the copy didn't carry it; reinstalls both triggers (triggers never
  survive "Make a copy"); checks timezone; ends with a health check.
- Existing setup functions refactored into silent cores + alerting menu
  wrappers (identical menu behavior).

**Version visibility / update path:**
- `SYSTEM_VERSION` constant; health check and the new "About / version" menu
  item show it with the starter-kit repo URL. Updating a city = paste the new
  Code.js over the old one; all state lives in the sheet.
- Health check also warns on sheet/script timezone mismatch.

**Public dashboard:** workbook ID now comes from Script Property `WORKBOOK_ID`
instead of a hardcoded constant; COGBM naming genericized. (San Antonio's live
dashboard keeps its current code until the v3.6 upgrade; set the property
before deploying.)

**Scope note:** the `@OnlyCurrentDoc` annotation was removed because
`FormApp.create()` needs the forms scope — first run after upgrading prompts
for re-authorization.

## v3.5 (2026-07-02)
Source cleanup — no behavior changes.
- Moved this changelog out of the `Code.js` header.
- Removed historical column-shift annotations ("was AL, shifted +2") from
  constants; comments now describe only the current layout.
- Renamed `LAST_COL` → `FORM_ROW_WIDTH` to make its meaning explicit (the
  60-column row that `onFormSubmit` builds — NOT the full v3 row width;
  archiving uses `V3_LAST_COL` = 115).
- Removed orphaned section banners for functions that no longer exist.
- Removed `setupDropdowns` and its menu item — fully redundant with
  "Set up data validation", which installs the same Status dropdown plus all
  other rules.
- Kept `updateInventoryOutFormulas()` as an explicit legacy alias for
  `setupReconciliationLogic()` in case an old installable trigger or manual
  run still references it by name.
- Deleted the stale `Bed Ministry Apps Script.txt` (a v3.2 copy; git history
  is the source of truth).

## v3.4
Pick-list substitutions & omissions at fulfillment (deployed live & verified).
- Pack-time dialog lets the volunteer swap mattress thickness per sleeping
  surface, choose comforter with/without sheets per child, and adjust or omit
  quantities for extras (pillows, protectors, bedrails, plaques, bears, books).
- Whatever is chosen is recorded as consumption, so inventory stays straight.
- Chosen fulfillment config is persisted as JSON (column DK/115) and pre-fills
  the dialog on a re-run; it travels with the row to Completed Deliveries.
- Demand formulas use actual consumption for Packed rows and the standard
  estimate for Active rows, so omissions don't overstate demand.
- If request data (beds/children) changes on a row that already had a pick
  list, the stale pick snapshot is cleared and a Packed row reverts to Active.

## v3.3 — Hardening / foolproofing
- CRITICAL FIX: archiving/restoring now carries ALL v3 columns (Pack Config,
  demand BJ–BW, consumption BX–CJ) instead of truncating at column 60.
  Previously, consumption data was lost when a request was completed, so
  inventory On Hand silently rebounded after every delivery. Archived rows are
  now frozen to static values (a stable historical record + big speed-up,
  since hundreds of live SUMIFS no longer recalc on Completed Deliveries).
- FIX: Mattress Protectors demand (BT) now matches the picker (1 per non-crib
  sleeping surface) instead of child count.
- FIX: Waiting List formula ranges standardized to a single bound so Packed
  reservations past row 999 are no longer dropped.
- Failures are now VISIBLE: dropped/failed form submissions email the
  maintainer (recipients in Config!O2) instead of logging silently.
- Form intake tolerates question-wording drift and reports missing fields by
  email rather than silently dropping a request.
- `handleEdit` now processes multi-row status edits (paste / fill-down).
- New menu item: Setup > Check system health (verifies triggers, TOTALS row,
  named range, protections, notification recipients).
- PDF export validates content type before saving.
- Removed the simple `onEdit(e)` wrapper (double-fired the archive alongside
  the installable trigger, creating duplicate rows). Edits are handled ONLY by
  the installable trigger (Bed Ministry > Setup > Install edit trigger).

Deploy notes (v3.3): paste script, save, reload, then run (Setup submenu):
Set up reconciliation logic → Repair: Waiting List formulas → Install form
trigger → Install edit trigger → Check system health. Then enter notification
email(s) in Config!O2 and do a physical recount in the Reconciliation Log to
re-anchor inventory.

## v3.0
Major restructuring: gendered + age-tiered comforters, mattress size variants,
constructed/purchased frame variants, per-family book allocation, age-capped
bears, plaques.
- Component count expanded from 14 → 39 items.
- Frames split into Constructed/Purchased variants.
- Mattresses split into 4 size variants (6"/8" × 38x74/39x75).
- Bedding split: separate Sheets, Mattress Protectors, Bedrails, Plaques, and
  18 comforter variants (with/without sheets × boys/girls/neutral ×
  toddler/young/teen).
- Books split by age tier: Toddler, Young, Teen.
- Bears no longer age-tiered (single line), age-capped at 10.
- Pack-time prompt expanded: constructed/purchased per bed type.
- Pick list groups items by category, includes frame component checklist for
  constructed beds.
- Per-child comforter allocation with gender + age matching and neutral
  fallback; prefer with-sheets variant over without.
- Per-family book allocation based on age tiers of children.

Upgrade path to v3.0:
1. Paste script, save, reload sheet
2. Run: Bed Ministry > Setup > Set up v3.0 inventory structure
3. Run: Set up reconciliation logic
4. Run: Set up data validation
5. Run: Set up formula protections
6. Begin physical recount (all On Hand = 0 until recounts entered)

## v2.4
Fixed compounding reconciliation bug (snapshot model).

## v2.3
Added Packed status, Pick Lists folder, book titles on Config.

## v2.2
Added Pillows and Toddler Books columns (14 components).
