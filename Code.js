// ============================================================
//  Bed Ministry — Inventory Management System v3.6
//
//  Tracks bed requests (Google Form → Waiting List), pick-list
//  generation with pack-time substitutions/omissions, delivery
//  archiving, and 39-component inventory reconciliation.
//
//  Version history: see CHANGELOG.md in the repo.
//
//  DEPLOY (after pasting a new script version): save, reload the
//  sheet, then run from Bed Ministry > Setup:
//    Set up reconciliation logic → Repair: Waiting List formulas →
//    Install form trigger → Install edit trigger → Check system health.
//
//  NEW CITY? Copy the template workbook, then run
//  Bed Ministry > 🚀 New City Setup — it walks through everything.
// ============================================================

const SYSTEM_VERSION = "3.6";
// Where updates live. Cities compare their version (health check / About)
// against the CHANGELOG here and paste the new Code.js over the old one —
// all data lives in the sheet, so updating the script never touches data.
const STARTER_KIT_REPO_URL = "https://github.com/Glitterphobe/bed-ministry-starter-kit";

// NOTE: no @OnlyCurrentDoc annotation — setupForm() creates a NEW Google Form
// via FormApp.create(), which the current-document-only scope would forbid.
// Adding the annotation back will break "Create & link Bed Request Form".


// ── Sheet name constants ────────────────────────────────────────────────────
const ACTIVE_SHEET      = "Waiting List";
const ARCHIVE_SHEET     = "Completed Deliveries";
const CANCELLED_SHEET   = "Cancelled Requests";
const DASHBOARD_SHEET   = "Dashboard";
const RESPONSES_SHEET   = "Form Responses 1";
const CONFIG_SHEET      = "Config";
const INCOMING_SHEET    = "Incoming Items";
const INVENTORY_SHEET   = "Inventory Position";
const RECON_SHEET       = "Reconciliation Log";

// ── Column constants ─────────────────────────────────────────────────────────
const STATUS_COL        = 3;   // Column C
const HEADER_ROW        = 1;
const TOTALS_ROW_LABEL  = "TOTALS";
const DELIVERY_DATE_COL = 51;  // Column AY
const FORM_ROW_WIDTH    = 60;  // Width of the row onFormSubmit builds (cols A–BH).
                               // NOT the full row width — v3 demand/consumption
                               // data runs out to column 114. Archiving/restoring
                               // use V3_LAST_COL (defined below) so that data is
                               // never truncated.

// Single source of truth for the last Waiting List data row referenced by
// inventory/shortage formulas. Previously some formulas used 999 and others
// 9999; the smaller bound silently dropped Packed reservations past row 999.
const WL_LAST_DATA_ROW  = 9999;

// Waiting List bed-type columns
const COL_CRIBS         = 5;   // E
const COL_TODDLERS      = 6;   // F
const COL_TWINS         = 7;   // G
const COL_BUNKS         = 8;   // H

// Summary formula columns
const COL_NO_BEDS       = 4;   // D
const COL_8_MATT        = 15;  // O
const COL_6_MATT        = 16;  // P
const COL_CRIB_MATT     = 17;  // Q
const COL_PILLOWS       = 18;  // R
const COL_TODDLER_BOOKS = 20;  // T
const COL_YOUNG_BOOKS   = 21;  // U
const COL_YOUNG_BEARS   = 22;  // V
const COL_OLDER_BOOKS   = 23;  // W

// Pick list header columns
const COL_APN           = 2;   // B
const COL_CAREGIVER     = 40;  // AN
const COL_CAREGIVER_PH  = 41;  // AO
const COL_COMPLEX       = 42;  // AP
const COL_ADDRESS       = 43;  // AQ
const COL_APT           = 44;  // AR
const COL_CITY          = 45;  // AS
const COL_ZIP           = 46;  // AT
const COL_DELIVERY_CONTACT    = 53; // BA
const COL_DELIVERY_CONTACT_PH = 54; // BB
const COL_DELIVERY_ORG        = 55; // BC
const COL_NOTES         = 56;  // BD

// Shortage/bunk columns at the end of the form-built row
const COL_SHORTAGE_IMPACT  = 58; // BF
const COL_CONSTRAINED      = 59; // BG
const COL_BUNK_CONFIG      = 60; // BH

// Text columns struck through when a row is styled Cancelled
const TEXT_COLS_FOR_STRIKETHROUGH = [2, 40, 41, 43, 49, 56]; // APN, Caregiver, Phone, Address, CW, Notes

// Child age columns (8 children; gender sits one column right of each age)
const CHILD_AGE_COLS_A1 = ["X", "Z", "AB", "AD", "AF", "AH", "AJ", "AL"];


// ── TOTALS row self-healing (v3.1 hardening) ────────────────────────────────
// The Waiting List TOTALS row sums columns D..W with a delete-proof pattern:
//     =SUM(D1:INDIRECT("D"&ROW()-1))
// INDIRECT rebuilds the range from text, so deleting data rows can never turn
// it into #REF!, and it auto-extends when a new row is inserted above TOTALS.
// ensureTotalsRow_ guarantees the row exists, carries the right label, and
// holds valid formulas — it is called after every add/complete/cancel/restore.
const TOTALS_FIRST_SUM_COL = 4;   // D (No. Beds)
const TOTALS_LAST_SUM_COL  = 23;  // W (Older Books)

function ensureTotalsRow_(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_SHEET);
  if (!sheet) return -1;

  // Locate TOTALS by a forgiving match (trim + uppercase) so "Totals" or a
  // stray trailing space can never hide the anchor row.
  var totalsRow = -1;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var colB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < colB.length; i++) {
      if (String(colB[i][0]).trim().toUpperCase() === TOTALS_ROW_LABEL) { totalsRow = i + 2; break; }
    }
  }

  // Recreate it at the bottom if it is missing.
  var created = false;
  if (totalsRow === -1) {
    totalsRow = Math.max(sheet.getLastRow() + 1, 2);
    created = true;
  }

  // Always re-assert the label; (re)write any formula that is missing or errored.
  sheet.getRange(totalsRow, 2).setValue(TOTALS_ROW_LABEL);
  for (var col = TOTALS_FIRST_SUM_COL; col <= TOTALS_LAST_SUM_COL; col++) {
    var L = sheet.getRange(1, col).getA1Notation().replace(/[0-9]+/g, "");
    var want = '=SUM(' + L + '1:INDIRECT("' + L + '"&ROW()-1))';
    var cell = sheet.getRange(totalsRow, col);
    var f = cell.getFormula();
    var v = cell.getValue();
    var broken = (typeof v === "string" && v.charAt(0) === "#");
    if (created || f === "" || broken) cell.setFormula(want);
  }
  if (created) {
    sheet.getRange(totalsRow, 1, 1, FORM_ROW_WIDTH).setFontWeight("bold").setBackground("#D9D9D9");
  }
  return totalsRow;
}

function repairTotalsRow() {
  var r = ensureTotalsRow_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_SHEET));
  SpreadsheetApp.getUi().alert("\u2713 TOTALS row ensured at row " + r + " (sums columns D\u2013W).");
}


// ═══════════════════════════════════════════════════════════════════════════
//  v3.0 INVENTORY CATALOG
// ═══════════════════════════════════════════════════════════════════════════

// The 39-item ordered catalog. Order matters — drives Config list, Inventory
// Position row order, Incoming Items column order. Grouping is for pick list
// presentation only.

const V3_CATALOG = [
  // Frames
  { name: "Twin Frame - Constructed",               category: "FRAMES" },
  { name: "Twin Frame - Purchased",                 category: "FRAMES" },
  { name: "Toddler Frame - Constructed",            category: "FRAMES" },
  { name: "Toddler Frame - Purchased",              category: "FRAMES" },
  { name: "Bunk Frame - Constructed",               category: "FRAMES" },
  { name: "Bunk Frame - Purchased",                 category: "FRAMES" },
  { name: "Crib",                                   category: "FRAMES" },
  // Mattresses
  { name: "Mattress 6in 38x74",                     category: "MATTRESSES" },
  { name: "Mattress 6in 39x75",                     category: "MATTRESSES" },
  { name: "Mattress 8in 38x74",                     category: "MATTRESSES" },
  { name: "Mattress 8in 39x75",                     category: "MATTRESSES" },
  { name: "Crib Mattress",                          category: "MATTRESSES" },
  // Bedding basics
  { name: "Pillow",                                 category: "BEDDING" },
  { name: "Twin Sheets",                            category: "BEDDING" },
  { name: "Twin Mattress Protector",                category: "BEDDING" },
  // Hardware
  { name: "Bedrail",                                category: "HARDWARE" },
  { name: "Plaque",                                 category: "HARDWARE" },
  // Comforters w/o Sheets (9)
  { name: "Comforter w/o Sheets - Boys Toddler",    category: "COMFORTERS" },
  { name: "Comforter w/o Sheets - Boys Young",      category: "COMFORTERS" },
  { name: "Comforter w/o Sheets - Boys Teen",       category: "COMFORTERS" },
  { name: "Comforter w/o Sheets - Girls Toddler",   category: "COMFORTERS" },
  { name: "Comforter w/o Sheets - Girls Young",     category: "COMFORTERS" },
  { name: "Comforter w/o Sheets - Girls Teen",      category: "COMFORTERS" },
  { name: "Comforter w/o Sheets - Neutral Toddler", category: "COMFORTERS" },
  { name: "Comforter w/o Sheets - Neutral Young",   category: "COMFORTERS" },
  { name: "Comforter w/o Sheets - Neutral Teen",    category: "COMFORTERS" },
  // Comforters with Sheets (9)
  { name: "Comforter w/Sheets - Boys Toddler",      category: "COMFORTERS" },
  { name: "Comforter w/Sheets - Boys Young",        category: "COMFORTERS" },
  { name: "Comforter w/Sheets - Boys Teen",         category: "COMFORTERS" },
  { name: "Comforter w/Sheets - Girls Toddler",     category: "COMFORTERS" },
  { name: "Comforter w/Sheets - Girls Young",       category: "COMFORTERS" },
  { name: "Comforter w/Sheets - Girls Teen",        category: "COMFORTERS" },
  { name: "Comforter w/Sheets - Neutral Toddler",   category: "COMFORTERS" },
  { name: "Comforter w/Sheets - Neutral Young",     category: "COMFORTERS" },
  { name: "Comforter w/Sheets - Neutral Teen",      category: "COMFORTERS" },
  // Bears & Books
  { name: "Bear",                                   category: "BEARS & BOOKS" },
  { name: "Toddler Book",                           category: "BEARS & BOOKS" },
  { name: "Young Book",                             category: "BEARS & BOOKS" },
  { name: "Teen Book",                              category: "BEARS & BOOKS" },
];

// v3.4 — Pack-time fulfillment overrides (substitutions / omissions).
// The volunteer may, at pick-list time, swap mattress thickness per sleeping
// surface (both directions), choose comforter with/without sheets per child, and
// set a quantity ≥ 0 (i.e. adjust or omit) for the "extras" below. Whatever is
// chosen is what gets recorded as consumption, so inventory stays straight.
// Items here render as number inputs (min 0) in the dialog; default = computed qty.
const V3_QTY_ADJUSTABLE = [
  { name: "Pillow",                  label: "Pillows" },
  { name: "Twin Mattress Protector", label: "Mattress protectors" },
  { name: "Bedrail",                 label: "Bedrails" },
  { name: "Plaque",                  label: "Plaques" },
  { name: "Bear",                    label: "Bears" },
  { name: "Toddler Book",            label: "Toddler book" },
  { name: "Young Book",              label: "Young book" },
  { name: "Teen Book",               label: "Teen book" },
];

// Convenience lookups — derived from V3_CATALOG
const V3_ITEM_COUNT = V3_CATALOG.length;   // 39
const V3_INV_FIRST_ROW = 4;
const V3_INV_LAST_ROW  = V3_INV_FIRST_ROW + V3_ITEM_COUNT - 1;  // 42
const V3_INC_FIRST_COL = 4;   // Column D on Incoming Items (after Date, SourceType, SourceDetail)
const V3_INC_LAST_COL  = V3_INC_FIRST_COL + V3_ITEM_COUNT - 1;  // 42 (column AP)
const V3_INC_NOTES_COL = V3_INC_LAST_COL + 1;                    // 43 (column AQ)

// Age tier boundaries (inclusive)
const V3_AGE_TODDLER_MAX = 3;    // 0-3
const V3_AGE_YOUNG_MAX   = 10;   // 4-10
const V3_AGE_TEEN_MAX    = 18;   // 11-18
const V3_BEAR_AGE_MAX    = 10;   // Bears only for ages 0-10

function v3AgeTier(age) {
  const a = Number(age);
  if (isNaN(a) || a < 0) return null;
  if (a <= V3_AGE_TODDLER_MAX) return "Toddler";
  if (a <= V3_AGE_YOUNG_MAX)   return "Young";
  if (a <= V3_AGE_TEEN_MAX)    return "Teen";
  return "Teen";  // 18+ treated as Teen
}

function v3GenderLabel(g) {
  const s = String(g || "").trim().toUpperCase();
  if (s === "M" || s === "BOY"  || s === "BOYS")  return "Boys";
  if (s === "F" || s === "GIRL" || s === "GIRLS") return "Girls";
  return "Neutral";
}

// v3.0 Waiting List extended column positions (BI–BW)
const V3_COL_PACK_CONFIG       = 61;  // BI — pack-time constructed/purchased JSON
const V3_COL_COMF_FIRST        = 62;  // BJ — first of 9 comforter demand cols (ignores w/-or-w/o-sheets)
const V3_COL_COMF_LAST         = 70;  // BR — last of 9
const V3_COL_SHEETS_DEMAND     = 71;  // BS
const V3_COL_MATT_PROT_DEMAND  = 72;  // BT
const V3_COL_BEDRAILS_MAX      = 73;  // BU
const V3_COL_PLAQUES_MAX       = 74;  // BV
const V3_COL_BEARS_DEMAND      = 75;  // BW

// v3.0 per-item actual consumption columns (BX onwards, 39 columns)
// Populated by generatePickList when a request is packed/delivered.
// These are what the snapshot-model Out formulas read.
const V3_COL_CONSUMPTION_FIRST = 76;                                // BX
const V3_COL_CONSUMPTION_LAST  = V3_COL_CONSUMPTION_FIRST + V3_ITEM_COUNT - 1;  // DJ (col 114)
// v3.4: chosen fulfillment config (substitutions/omissions) stored as JSON in the
// first column AFTER the consumption block. Appended at the end so existing column
// references (consumption block, protection map, inventory formulas) never shift.
const V3_COL_FULFILL_CONFIG    = V3_COL_CONSUMPTION_LAST + 1;       // 115 (DK)
const V3_LAST_COL              = V3_COL_FULFILL_CONFIG;             // 115 — copyRowFull_ archives through here

// Ordered list of 9 comforter demand columns (gender × age, ignoring sheets)
// in the order: Boys-Toddler, Boys-Young, Boys-Teen, Girls-Toddler, Girls-Young,
// Girls-Teen, Neutral-Toddler, Neutral-Young, Neutral-Teen
const V3_COMF_DEMAND_LABELS = [
  { gender: "Boys",    age: "Toddler" },
  { gender: "Boys",    age: "Young"   },
  { gender: "Boys",    age: "Teen"    },
  { gender: "Girls",   age: "Toddler" },
  { gender: "Girls",   age: "Young"   },
  { gender: "Girls",   age: "Teen"    },
  { gender: "Neutral", age: "Toddler" },
  { gender: "Neutral", age: "Young"   },
  { gender: "Neutral", age: "Teen"    },
];

// Constructed frame component checklist (for pick list display only — not inventory)
const V3_CONSTRUCTED_FRAME_PARTS = [
  "Headboard",
  "Footboard",
  "Left Side Rail",
  "Right Side Rail",
  "Slats",
  "Bolts / Hardware",
];


// ═══════════════════════════════════════════════════════════════════════════
//  v3.6 BED REQUEST FORM SCHEMA
//  The exact question set onFormSubmit expects, in form order. setupForm()
//  builds a linked Google Form from this; verifyFormSchema_() (health check)
//  compares the live form against it so a renamed question is caught before
//  it silently drops a field.
//
//  Reconciled 2026-07-06 against the live San Antonio form (via
//  logLiveFormSchema on a private copy), then genericized: San-Antonio-specific
//  help text (a ministry email, a local area code) was replaced with neutral
//  wording so the generated form carries no instance data.
//
//  Item fields:
//    title    — exact question title (matched by onFormSubmit via normKey)
//    type     — TEXT | DATE | MC (multiple choice) | PARAGRAPH | PAGE
//    choices  — MC options
//    helpText — optional italic hint under the question
//    required — optional; marks the field required on the generated form
//    bedType  — bed-count questions only; dropped when that bed type is
//               disabled in Ministry Options
//    nav      — MC child-flow control: { no: "<pageId>" } routes a "No"
//               answer straight to that page (skipping the rest of the
//               children); "Yes" continues to the next page
//    pageId   — PAGE items only; nav `no` targets reference these
//
//  IMPORTANT: every "Need to enter another child?" item has the IDENTICAL
//  title — that's how the live form is built. (The response SHEET disambiguates
//  the columns as "... 2", "... 3"; the FORM items do not.) onFormSubmit never
//  reads these — they are pure navigation — so duplicate titles are harmless,
//  and keeping them identical is what stops verifyFormSchema_ from crying
//  "renamed question" on a real form.
// ═══════════════════════════════════════════════════════════════════════════

// Child pages 2–8 follow one repeating shape; build them programmatically so
// the schema stays readable and the titles can't drift out of sync.
function v3BuildChildPages_() {
  const items = [];
  for (let n = 2; n <= 8; n++) {
    const page = { type: "PAGE", title: "Child " + n };
    if (n === 8) page.helpText = "If there are more than 8 children, reach out to your ministry lead.";
    items.push(page);
    items.push({ title: "Child " + n + " Age",    type: "TEXT" });
    items.push({ title: "Child " + n + " Gender",  type: "MC", choices: ["M", "F"] });
    // No "Need to enter another child?" after Child 8 — the form flows straight
    // to Caregiver Information from there.
    if (n < 8) {
      items.push({ title: "Need to enter another child?", type: "MC",
                   choices: ["Yes", "No"], nav: { no: "CAREGIVER" } });
    }
  }
  return items;
}

const FORM_SCHEMA = [
  { type: "PAGE", title: "Request Information" },
  { title: "Date of Request",                                     type: "DATE", required: true },
  { title: "Tracking Number (Care Portal/APN#)",                  type: "TEXT", required: true },

  { type: "PAGE", title: "Beds Needed" },
  { title: "How Many Twin Beds?",                                 type: "TEXT", required: true, bedType: "Twin" },
  { title: "How Many Bunk Beds? (1 bunk = top and bottom bed)",   type: "TEXT", required: true, bedType: "Bunk" },
  { title: "How many Toddler Beds?",                              type: "TEXT", required: true, bedType: "Toddler" },
  { title: "How many Cribs?",                                     type: "TEXT", required: true, bedType: "Crib" },

  { type: "PAGE", title: "Children" },
  { title: "How many Children need beds?", type: "MC", required: true,
    choices: ["1", "2", "3", "4", "5", "6", "7", "8"],
    helpText: "If your request is for more than 8 children, contact your ministry lead." },

  { type: "PAGE", title: "Child 1" },
  { title: "Child 1 Age",                                         type: "TEXT", required: true },
  { title: "Child 1 Gender",                                      type: "MC", choices: ["M", "F"], required: true },
  { title: "Need to enter another child?",                        type: "MC", choices: ["Yes", "No"], required: true, nav: { no: "CAREGIVER" } },

  // Child pages 2–8 (see v3BuildChildPages_)
  ...v3BuildChildPages_(),

  { type: "PAGE", pageId: "CAREGIVER", title: "Caregiver Information" },
  { title: "Caregiver Name (First and Last)",                     type: "TEXT", required: true },
  { title: "Caregiver Phone",                                     type: "TEXT", required: true, helpText: "Format guidance:  (555)-555-1234" },

  { type: "PAGE", title: "Case Worker Information" },
  { title: "Case Worker Name",                                    type: "TEXT", required: true, helpText: "Full name of the referring case worker." },
  { title: "Case Worker Phone",                                   type: "TEXT", required: true, helpText: "Format guidance:  (555)-555-1234" },

  { type: "PAGE", title: "Delivery Address" },
  { title: "Complex/Building Name",                               type: "TEXT", helpText: "If an apartment complex, include the complex name and building number (if applicable) here." },
  { title: "Street Address",                                      type: "TEXT", required: true, helpText: "Street number and name only, no city or ZIP" },
  { title: "Apt/Unit #",                                          type: "TEXT", helpText: "e.g. Apt 204 or #1103" },
  { title: "Floor #",                                             type: "TEXT", helpText: "Ground floor = 1" },
  { title: "City",                                                type: "TEXT", required: true },
  { title: "ZIP Code",                                            type: "TEXT", required: true, helpText: "5-digit ZIP. Used for delivery team routing." },
  { title: "Gate Code",                                           type: "TEXT", helpText: "If the complex or neighborhood is gated, please provide a gate code for the delivery team." },
  { title: "Do you have a Delivery Team assigned already?",       type: "MC", choices: ["Yes", "No"], required: true, helpText: "Some referrals arrive with a delivery team already assigned." },

  { type: "PAGE", title: "Delivery Team", helpText: "Information about the delivery team, if already available." },
  { title: "Delivery Contact Name",                               type: "TEXT" },
  { title: "Delivery Contact Phone",                              type: "TEXT", helpText: "Format guidance:  (555)-555-1234" },
  { title: "Delivery Organization",                               type: "TEXT" },

  { type: "PAGE", title: "Additional Info", helpText: "Capture any relevant information here that hasn't been included in the responses so far." },
  { title: "Notes",                                               type: "PARAGRAPH" },
];


// ═══════════════════════════════════════════════════════════════════════════
//  v3.6 MINISTRY OPTIONS — per-city configuration (starter-kit flexibility)
//  Lives on Config in a block written by setupMinistryOptions(). Every value
//  BLANK means YES/default, so a workbook that has never run the setup (e.g.
//  the San Antonio original) behaves exactly as before.
//
//  Ministry name  → Config!E11 (pick-list PDF title; blank = "Bed Ministry")
//  Toggles        → Config!E13:E21, YES/NO, blank = YES:
//    row 13 Cribs · 14 Toddler Beds · 15 Twin Beds · 16 Bunk Beds
//    row 17 Bears · 18 Books · 19 Plaques · 20 Bedrails · 21 Themed Bedding
//
//  A disabled bed type is dropped from the generated form. Disabled extras
//  (bears/books/plaques/bedrails) demand-forecast at 0, default to 0 on pick
//  lists, and are hidden from the pick-list dialog — but their columns stay,
//  so flipping a toggle back on later needs no migration.
// ═══════════════════════════════════════════════════════════════════════════

const OPT_BLOCK_HEADER_ROW = 10;   // Config!D10 "Ministry Options"
const OPT_NAME_ROW         = 11;   // Config!D11/E11 Ministry Name
const OPT_FIRST_ROW        = 13;   // first toggle row
const OPT_COL_LABEL        = 4;    // D
const OPT_COL_VALUE        = 5;    // E
const MINISTRY_OPTIONS = [
  { key: "cribs",    label: "Cribs",         bedType: "Crib" },
  { key: "toddlers", label: "Toddler Beds",  bedType: "Toddler" },
  { key: "twins",    label: "Twin Beds",     bedType: "Twin" },
  { key: "bunks",    label: "Bunk Beds",     bedType: "Bunk" },
  { key: "bears",    label: "Bears" },
  { key: "books",    label: "Books" },
  { key: "plaques",  label: "Plaques" },
  { key: "bedrails", label: "Bedrails" },
  { key: "themed",   label: "Themed Bedding" },
];

// Catalog items greyed out on Inventory Position when an option is off.
const OPT_CATALOG_ITEMS = {
  cribs:    ["Crib", "Crib Mattress"],
  toddlers: ["Toddler Frame - Constructed", "Toddler Frame - Purchased"],
  twins:    ["Twin Frame - Constructed", "Twin Frame - Purchased"],
  bunks:    ["Bunk Frame - Constructed", "Bunk Frame - Purchased"],
  bears:    ["Bear"],
  books:    ["Toddler Book", "Young Book", "Teen Book"],
  plaques:  ["Plaque"],
  bedrails: ["Bedrail"],
};

// Read the toggles. Blank/unrecognized = enabled, so pre-v3.6 workbooks and
// the San Antonio original run with everything on. Best-effort: never throws
// (falls back to all-enabled), safe inside triggers. Cached per execution —
// row-formula writers call this inside loops over every Waiting List row.
let _ministryOptionsCache = null;
function getMinistryOptions_() {
  if (_ministryOptionsCache) return _ministryOptionsCache;
  const opts = {};
  MINISTRY_OPTIONS.forEach(o => { opts[o.key] = true; });
  try {
    const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    if (!cfg) return opts;
    const vals = cfg.getRange(OPT_FIRST_ROW, OPT_COL_VALUE, MINISTRY_OPTIONS.length, 1).getValues();
    MINISTRY_OPTIONS.forEach((o, i) => {
      if (String(vals[i][0] || "").trim().toUpperCase() === "NO") opts[o.key] = false;
    });
  } catch (e) { /* all-enabled fallback */ }
  _ministryOptionsCache = opts;
  return opts;
}

// Ministry display name for the pick-list PDF title (and future headings).
function getMinistryName_() {
  try {
    const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    const v = cfg ? String(cfg.getRange(OPT_NAME_ROW, OPT_COL_VALUE).getValue() || "").trim() : "";
    if (v) return v;
  } catch (e) {}
  return "Bed Ministry";
}



// ═══════════════════════════════════════════════════════════════════════════
//  MAINTAINER NOTIFICATIONS (v3.3) — make failures visible, never silent
//  Recipients are read from Config!O2 (comma / semicolon / space separated).
//  Sending is free via MailApp (Gmail daily quota). If O2 is blank, falls back
//  to the script owner's email so an alert is never lost. Best-effort: these
//  never throw, so they're safe to call from inside triggers.
// ═══════════════════════════════════════════════════════════════════════════

function getNotificationRecipients_() {
  try {
    const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    let raw = cfg ? String(cfg.getRange("O2").getValue() || "").trim() : "";
    if (!raw) raw = Session.getEffectiveUser().getEmail() || "";
    return raw.split(/[,;\s]+/).map(s => s.trim()).filter(Boolean);
  } catch (e) {
    return [];
  }
}

function notifyMaintainer_(subject, body) {
  try {
    const to = getNotificationRecipients_();
    if (!to.length) { Logger.log("notifyMaintainer_ (no recipients): " + subject); return; }
    MailApp.sendEmail({ to: to.join(","), subject: "[Bed Ministry] " + subject, body: body });
  } catch (e) {
    Logger.log("notifyMaintainer_ failed: " + e.message + " | subject: " + subject);
  }
}


// ═══════════════════════════════════════════════════════════════════════════
//  EDIT TRIGGER (installable) — with Delivery Date safeguard
// ═══════════════════════════════════════════════════════════════════════════

function handleEdit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== ACTIVE_SHEET) return;

    const firstCol = e.range.getColumn();
    const lastCol  = firstCol + e.range.getNumColumns() - 1;
    const firstRow = e.range.getRow();
    const numRows  = e.range.getNumRows();
    const ss       = SpreadsheetApp.getActiveSpreadsheet();

    // What did this edit touch? (Cheap range checks so unrelated edits bail fast.)
    const touchesStatus  = !(STATUS_COL < firstCol || STATUS_COL > lastCol);
    // Request fields whose change invalidates an already-generated pick list:
    //   bed counts E\u2013H (COL_CRIBS..COL_BUNKS) and children age/gender X\u2013AM (24\u201339).
    const touchesRequest =
      rangesIntersect_(firstCol, lastCol, COL_CRIBS, COL_BUNKS) ||
      rangesIntersect_(firstCol, lastCol, 24, 39);

    if (!touchesStatus && !touchesRequest) return;

    // (1) v3.4 safeguard \u2014 if request data changed on a row that already had a pick
    //     list generated, clear the stale snapshot (BI / BX\u2013DJ / config) so the summary
    //     columns revert to the live estimate and the row must be re-packed.
    if (touchesRequest) {
      for (let i = 0; i < numRows; i++) {
        const row = firstRow + i;
        if (row <= HEADER_ROW) continue;
        clearStalePickList_(sheet, ss, row);
      }
    }

    // (2) Status changes \u2014 snapshot each row's new Status first, then process from the
    //     BOTTOM up (archiving deletes rows, so bottom-to-top keeps later rows valid).
    if (touchesStatus) {
      const edits = [];
      for (let i = 0; i < numRows; i++) {
        const row = firstRow + i;
        if (row <= HEADER_ROW) continue;
        edits.push({ row: row, value: sheet.getRange(row, STATUS_COL).getValue() });
      }
      for (let i = edits.length - 1; i >= 0; i--) {
        handleStatusChange_(sheet, edits[i].row, edits[i].value, ss);
      }
    }
  } catch (err) {
    // UI is not available in trigger context, so log and email rather than alert.
    Logger.log("Edit handler error: " + err.message);
    notifyMaintainer_("Edit handler error",
      "An error occurred while handling an edit on the Waiting List:\n\n" +
      err.message + "\n\nThe edit may not have been fully processed \u2014 please check.");
  }
}

// True if integer ranges [aLo,aHi] and [bLo,bHi] overlap.
function rangesIntersect_(aLo, aHi, bLo, bHi) {
  return aLo <= bHi && bLo <= aHi;
}

// v3.4 \u2014 When a request's beds/children change after a pick list was generated, the
// saved consumption snapshot (BX\u2013DJ), Pack Config (BI) and fulfillment config no longer
// match the request. Clear them so the O\u2013W summary reverts to the live estimate and the
// row is re-packed. A Packed row is reverted to Active so it isn't left reserving stale
// inventory. No-op when no pick list was generated (BI blank) or on the TOTALS row.
// (Programmatic writes from the script don't fire the edit trigger, so this only runs on
//  genuine user edits.)
function clearStalePickList_(sheet, ss, row) {
  const colB = String(sheet.getRange(row, 2).getValue()).trim().toUpperCase();
  if (colB === TOTALS_ROW_LABEL) return;
  const bi = String(sheet.getRange(row, V3_COL_PACK_CONFIG).getValue() || "").trim();
  if (!bi) return;  // no pick list generated for this row \u2192 nothing to clear

  sheet.getRange(row, V3_COL_PACK_CONFIG).clearContent();
  sheet.getRange(row, V3_COL_CONSUMPTION_FIRST, 1, V3_ITEM_COUNT).clearContent();
  sheet.getRange(row, V3_COL_FULFILL_CONFIG).clearContent();

  if (sheet.getRange(row, STATUS_COL).getValue() === "Packed") {
    sheet.getRange(row, STATUS_COL).setValue("Active");
    styleRowActive_(sheet, row);
  }
  ss.toast("Request on row " + row + " changed \u2014 regenerate the pick list.",
           "Pick list cleared", 5);
}

// Process a single row whose Status just changed. Archiving is serialised with a
// document lock + re-read guard so each status change is acted on EXACTLY ONCE,
// even if triggers overlap \u2014 this is what prevents duplicate rows in the archives.
function handleStatusChange_(sheet, row, newValue, ss) {
  if (newValue === "Completed" || newValue === "Cancelled") {
    const lock = LockService.getDocumentLock();
    try { lock.waitLock(15000); } catch (lockErr) { return; }
    try {
      const current = sheet.getRange(row, STATUS_COL).getValue();
      if (current !== newValue) return;  // already handled by another execution

      if (newValue === "Completed") {
        const deliveryDate = sheet.getRange(row, DELIVERY_DATE_COL).getValue();
        if (!deliveryDate) {
          sheet.getRange(row, STATUS_COL).setValue("Active");
          styleRowActive_(sheet, row);
          ss.toast("Enter a Delivery Date (column AY) before marking row " + row +
                   " as Completed.", "\u26A0 Not archived", 5);
          return;
        }
        archiveRow_(sheet, row);
      } else {
        archiveCancelledRow_(sheet, row);
      }
    } finally {
      lock.releaseLock();
    }
  } else if (newValue === "Packed") {
    styleRowPacked_(sheet, row);
  } else if (newValue === "Active") {
    styleRowActive_(sheet, row);
  }
}

// NOTE: the simple onEdit(e) wrapper was REMOVED in v3.2. A function literally
// named onEdit runs as a simple trigger on every edit; combined with the
// installable handleEdit trigger it fired the archive twice and created
// duplicate rows. Edits are now handled ONLY by the installable trigger
// (Bed Ministry > Setup > Install edit trigger).


// ═══════════════════════════════════════════════════════════════════════════
//  FORM SUBMIT HANDLER
// ═══════════════════════════════════════════════════════════════════════════

function onFormSubmit(e) {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getSheetByName(ACTIVE_SHEET);
    const responses   = e.namedValues || {};

    if (!activeSheet) {
      Logger.log("Waiting List sheet not found.");
      notifyMaintainer_("Form submission failed — 'Waiting List' sheet not found",
        "A form was submitted but the script could not find the Waiting List tab, so the " +
        "request was NOT added. Check that the tab exists and is named exactly 'Waiting List'.");
      return;
    }

    ensureTotalsRow_(activeSheet);  // self-heal: guarantee the TOTALS anchor exists

    // v3.3: tolerate form-question wording drift (capitalization, extra spaces,
    // punctuation such as a trailing "?"). Build a normalized lookup once; get()
    // falls back to it when an exact key match fails, so a small wording change
    // no longer silently drops a field.
    const normKey = s => String(s).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
    const normResponses = {};
    Object.keys(responses).forEach(k => { normResponses[normKey(k)] = responses[k]; });

    function get(key, defaultVal = "") {
      let val = responses[key];
      if (!(val && val[0])) val = normResponses[normKey(key)];
      return val && val[0] ? String(val[0]).trim() : defaultVal;
    }

    function getNum(key, defaultVal = 0) {
      const val = get(key, String(defaultVal));
      const num = parseInt(val, 10);
      return isNaN(num) ? defaultVal : num;
    }

    function fmtPhone(raw) {
      if (!raw) return "";
      const digits = raw.replace(/\D/g, "");
      if (digits.length === 10) return `(${digits.slice(0,3)})-${digits.slice(3,6)}-${digits.slice(6)}`;
      if (digits.length === 11 && digits[0] === "1") return `(${digits.slice(1,4)})-${digits.slice(4,7)}-${digits.slice(7)}`;
      return raw;
    }

    const cribs    = getNum("How many Cribs?");
    const toddlers = getNum("How many Toddler Beds?");
    const twins    = getNum("How Many Twin Beds?");
    const bunks    = getNum("How Many Bunk Beds? (1 bunk = top and bottom bed)");
    const totalBeds = cribs + toddlers + twins + (bunks * 2);

    if (totalBeds === 0) {
      const apn0 = get("Tracking Number (Care Portal/APN#)") || "(none)";
      Logger.log("WARNING: Zero beds. APN: " + apn0);
      notifyMaintainer_("Form submission counted 0 beds — please review",
        "A bed request came in but the script read 0 beds, so it was NOT added to the " +
        "Waiting List.\n\nTracking #: " + apn0 + "\n\n" +
        "This usually means the form was submitted with no bed quantities, OR a form " +
        "question was renamed so the script can no longer read the counts.\n\n" +
        "Form fields received in this submission:\n  - " +
        Object.keys(responses).join("\n  - ") + "\n\n" +
        "Open the newest row on the Form Responses tab and add the request manually if needed.");
      return;
    }

    const numChildren = getNum("How many Children need beds?", 0);
    const childData = [];
    for (let i = 1; i <= 8; i++) {
      childData.push({
        age:    i <= numChildren ? getNum(`Child ${i} Age`) || "" : "",
        gender: i <= numChildren ? get(`Child ${i} Gender`)     : "",
      });
    }

    // Build row — 60 columns total
    const row = new Array(FORM_ROW_WIDTH).fill("");

    row[0]  = get("Date of Request") ? new Date(get("Date of Request")) : new Date();
    row[1]  = get("Tracking Number (Care Portal/APN#)");
    row[2]  = "Active";
    row[3]  = ""; // D: formula
    row[4]  = cribs;
    row[5]  = toddlers;
    row[6]  = twins;
    row[7]  = bunks;
    // Cols 9-14 (I-N): administrative
    // Cols 15-17 (O-Q): mattress formulas
    // Col 18 (R): Pillows formula
    // Col 19 (S): Themed Bedding? (blank)
    // Col 20 (T): Toddler Books formula
    // Cols 21-23 (U-W): Young Books / Young Bears / Older Books formulas
    // Child data at cols 24-39 (X-AM): age/gender pairs
    for (let i = 0; i < 8; i++) {
      row[23 + (i * 2)] = childData[i].age;
      row[24 + (i * 2)] = childData[i].gender;
    }
    row[39] = get("Caregiver Name (First and Last)");       // col 40/AN
    row[40] = fmtPhone(get("Caregiver Phone"));              // col 41/AO
    row[41] = get("Complex/Building Name");                  // col 42/AP
    row[42] = get("Street Address");                         // col 43/AQ
    row[43] = get("Apt/Unit #");                             // col 44/AR
    row[44] = get("City");                                   // col 45/AS
    row[45] = get("ZIP Code");                               // col 46/AT
    row[46] = get("Gate Code");                              // col 47/AU
    row[47] = getNum("Floor #") || "";                       // col 48/AV
    row[48] = get("Case Worker Name");                       // col 49/AW
    row[49] = fmtPhone(get("Case Worker Phone"));            // col 50/AX
    // Col 51/AY: Delivery Date (blank)
    // Col 52/AZ: Caseworker Notified (blank)
    const deliveryAssigned = get("Do you have a Delivery Team assigned already?");
    if (deliveryAssigned === "Yes") {
      row[52] = get("Delivery Contact Name");                // col 53/BA
      row[53] = fmtPhone(get("Delivery Contact Phone"));     // col 54/BB
      row[54] = get("Delivery Organization");                // col 55/BC
    }
    row[55] = get("Notes");                                  // col 56/BD
    row[56] = false;                                         // col 57/BE — Logged
    // Cols 58-60 (BF/BG/BH): Shortage Impact, Constrained Items, Bunk Config (formulas)

    // Find insert position (above TOTALS). Halt with a clear error if the
    // TOTALS row is missing or mislabelled — silently appending to the bottom
    // would break SUM formulas.
    const lastRow = activeSheet.getLastRow();
    let insertRow = -1;
    for (let r = 2; r <= lastRow; r++) {
      if (String(activeSheet.getRange(r, 2).getValue()) === TOTALS_ROW_LABEL) {
        insertRow = r;
        break;
      }
    }
    if (insertRow === -1) {
      const msg = "Form submission could not be processed: the '" + TOTALS_ROW_LABEL +
                  "' row is missing from the Waiting List tab. " +
                  "The new request was NOT added. Please restore the TOTALS row " +
                  "(column B) and re-submit the form, or manually copy the form " +
                  "response from the Form Responses 1 tab.";
      Logger.log(msg);
      try {
        SpreadsheetApp.getUi().alert(msg);
      } catch (e) {
        // UI may not be available in trigger context — log only
      }
      notifyMaintainer_("Form submission halted — TOTALS row missing",
        msg + "\n\nTracking #: " + (get("Tracking Number (Care Portal/APN#)") || "(none)"));
      ss.toast("⚠ Form submission halted — TOTALS row missing. See Form Responses 1.",
               "Error", 10);
      return;
    }
    activeSheet.insertRowBefore(insertRow);
    activeSheet.getRange(insertRow, 1, 1, row.length).setValues([row]);

    writeRowFormulas_(activeSheet, insertRow);
    styleRowActive_(activeSheet, insertRow);
    activeSheet.getRange(insertRow, 1).setNumberFormat("MM/DD/YYYY");
    ss.toast(`✓ New bed request added to Waiting List (row ${insertRow})`, "Form Submitted", 4);

  } catch (err) {
    Logger.log("Form submission error: " + err.message);
    notifyMaintainer_("Form submission ERROR — request may not have been added",
      "An error occurred while processing a form submission:\n\n" + err.message +
      "\n\nPlease check the newest row on the Form Responses tab and add the request " +
      "to the Waiting List manually if it is missing.");
  }
}


// ═══════════════════════════════════════════════════════════════════════════
//  WRITE ROW FORMULAS
// ═══════════════════════════════════════════════════════════════════════════

function writeRowFormulas_(sheet, r) {
  // D: No. Beds — unchanged
  sheet.getRange(r, COL_NO_BEDS).setFormula(`=E${r}+F${r}+G${r}+(H${r}*2)`);
  
  // ── O–W summary columns (v3.4) ──
  // Show the STANDARD estimate until a pick list is generated (Pack Config / BI blank),
  // then the ACTUAL allocated quantities from the consumption block — so substitutions
  // and omissions are visible on the request line and freeze into Completed Deliveries.
  const biCol = getColumnA1_(V3_COL_PACK_CONFIG);                 // "BI"
  const gen   = "$" + biCol + r + '<>""';                         // a pick list has been generated
  const cc = name => getColumnA1_(V3_COL_CONSUMPTION_FIRST + V3_CATALOG.findIndex(it => it.name === name)) + r;
  const a6    = "(" + cc("Mattress 6in 38x74") + "+" + cc("Mattress 6in 39x75") + ")";
  const a8    = "(" + cc("Mattress 8in 38x74") + "+" + cc("Mattress 8in 39x75") + ")";
  const aCrib = cc("Crib Mattress");
  const aPill = cc("Pillow");
  const aBear = cc("Bear");
  const aTBk  = cc("Toddler Book");
  const aYBk  = cc("Young Book");
  const aZBk  = cc("Teen Book");

  // O: 8" Mattresses
  sheet.getRange(r, COL_8_MATT).setFormula(`=IF(${gen},${a8},IFERROR(G${r},0)+IFERROR(H${r},0))`);
  // P: 6" Mattresses
  sheet.getRange(r, COL_6_MATT).setFormula(`=IF(${gen},${a6},IFERROR(F${r},0)+IFERROR(H${r},0))`);
  // Q: Crib Mattresses
  sheet.getRange(r, COL_CRIB_MATT).setFormula(`=IF(${gen},${aCrib},IFERROR(E${r},0))`);
  // R: Pillows — standard: O + P + Q; actual: Pillow consumption
  sheet.getRange(r, COL_PILLOWS).setFormula(
    `=IF(${gen},${aPill},IFERROR(O${r},0)+IFERROR(P${r},0)+IFERROR(Q${r},0))`
  );

  // Age-based book/bear standard formulas (used until a pick list is generated).
  // v3.6: a city that has switched Books/Bears off in Ministry Options forecasts
  // 0 for them (actual consumption still shows if something was packed anyway).
  const opts = getMinistryOptions_();
  const ageCols = CHILD_AGE_COLS_A1; // X, Z, AB, AD, AF, AH, AJ, AL
  const tStd = !opts.books ? "0" : `IF(${ageCols.map(c => `COUNTIFS(${c}${r},">=0",${c}${r},"<=${V3_AGE_TODDLER_MAX}")`).join("+")}>0,1,0)`;
  const uStd = !opts.books ? "0" : `IF(${ageCols.map(c => `COUNTIFS(${c}${r},">=4",${c}${r},"<=${V3_AGE_YOUNG_MAX}")`).join("+")}>0,1,0)`;
  const vStd = !opts.bears ? "0" : `${ageCols.map(c => `COUNTIFS(${c}${r},">=0",${c}${r},"<=${V3_BEAR_AGE_MAX}")`).join("+")}`;
  const wStd = !opts.books ? "0" : `IF(${ageCols.map(c => `COUNTIFS(${c}${r},">=${V3_AGE_YOUNG_MAX + 1}",${c}${r},"<=${V3_AGE_TEEN_MAX}")`).join("+")}>0,1,0)`;

  // T: Toddler Books
  sheet.getRange(r, COL_TODDLER_BOOKS).setFormula(`=IF(${gen},${aTBk},${tStd})`);
  // U: Young Books
  sheet.getRange(r, COL_YOUNG_BOOKS).setFormula(`=IF(${gen},${aYBk},${uStd})`);
  // V: Bears
  sheet.getRange(r, COL_YOUNG_BEARS).setFormula(`=IF(${gen},${aBear},${vStd})`);
  // W: Teen Books
  sheet.getRange(r, COL_OLDER_BOOKS).setFormula(`=IF(${gen},${aZBk},${wStd})`);

  // BF: Shortage Impact (v3.0) — "⚠ CONSTRAINED" if this request's bed types
  // require any category that's currently short. Simpler and more reliable
  // than the per-item approach: checks the shortage section (rows 47-85) by
  // bed-type category using the same SUMIFS pattern as BG below.
  const inv = "'Inventory Position'";
  sheet.getRange(r, COL_SHORTAGE_IMPACT).setFormula(
    `=IF(C${r}<>"Active","",` +
    `IF(OR(` +
    `AND(OR(G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Twin Frame*")>0),` +
    `AND(F${r}>0,SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Toddler Frame*")>0),` +
    `AND(H${r}>0,SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Bunk Frame*")>0),` +
    `AND(E${r}>0,SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Crib")>0),` +
    `AND(OR(F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Mattress*")>0),` +
    `AND(E${r}>0,SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Crib Mattress")>0),` +
    `AND(OR(E${r}>0,F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Pillow")>0),` +
    `AND(OR(E${r}>0,F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Comforter*")>0),` +
    `AND(OR(F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Twin Sheets")>0),` +
    `AND(OR(F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Twin Mattress Protector")>0),` +
    `AND(OR(F${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Bedrail")>0),` +
    `AND(OR(F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Plaque")>0)` +
    `),"\u26A0 CONSTRAINED","\u2713 CLEAR"))`
  );

  // BG: Constrained Items (v3.0) — list of short items this request uses,
  // derived from the v3.0 shortage section at rows 47-85. Simplified from the
  // legacy per-item check because v3.0 has 39 components — too many to enumerate.
  // Shows category-level summary instead of specific items for readability.
  sheet.getRange(r, COL_CONSTRAINED).setFormula(
    `=IF(C${r}<>"Active","",IF(BF${r}="\u2713 CLEAR","",` +
    `TRIM(` +
    `IF(AND(OR(G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Twin Frame*")>0),"Twin Frames, ","")&` +
    `IF(AND(F${r}>0,SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Toddler Frame*")>0),"Toddler Frames, ","")&` +
    `IF(AND(H${r}>0,SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Bunk Frame*")>0),"Bunk Frames, ","")&` +
    `IF(AND(E${r}>0,SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Crib")>0),"Cribs, ","")&` +
    `IF(AND(OR(F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Mattress*")>0),"Mattresses, ","")&` +
    `IF(AND(OR(E${r}>0,F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Comforter*")>0),"Comforters, ","")&` +
    `IF(AND(OR(F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Twin Sheets")>0),"Sheets, ","")&` +
    `IF(AND(OR(F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Twin Mattress Protector")>0),"Mattress Protectors, ","")&` +
    `IF(AND(OR(F${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Bedrail")>0),"Bedrails, ","")&` +
    `IF(AND(OR(F${r}>0,G${r}>0,H${r}>0),SUMIFS(${inv}!$D$47:$D$85,${inv}!$A$47:$A$85,"Plaque")>0),"Plaques, ","")` +
    `)))`
  );

  // BH: Bunk Config — populated by generatePickList
  // BI: Pack Config — populated by generatePickList

  // ── v3.0: Also populate the demand columns BJ–BW for this row ──
  // This is critical — without this, new rows from the form won't have
  // comforter/sheets/bedrails/plaques/bears demand formulas.
  v3WriteWaitingListFormulas(sheet, r);
}


// ═══════════════════════════════════════════════════════════════════════════
//  FULL-WIDTH, VALUE-FROZEN ROW COPY (v3.3)
//  Copies an entire request row across ALL v3 columns (A..V3_LAST_COL = 114),
//  not just FORM_ROW_WIDTH (60), so Pack Config (BI), demand (BJ–BW) and — critically —
//  per-item consumption (BX–CJ) travel with the row. Computed cells are then
//  frozen to static values, so archived rows are a stable snapshot, inventory
//  reads reliable consumption numbers, and hundreds of live SUMIFS no longer
//  recalculate on the archive tabs. Returns the populated destination range.
// ═══════════════════════════════════════════════════════════════════════════

function copyRowFull_(srcSheet, srcRow, destSheet, destRow) {
  const src  = srcSheet.getRange(srcRow, 1, 1, V3_LAST_COL);
  const dest = destSheet.getRange(destRow, 1, 1, V3_LAST_COL);
  src.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false); // values + formats
  dest.setValues(src.getValues());  // overwrite formulas with their computed values
  return dest;
}


// ═══════════════════════════════════════════════════════════════════════════
//  ARCHIVE ROW — full v3 width, value-frozen (v3.3)
// ═══════════════════════════════════════════════════════════════════════════

function archiveRow_(activeSheet, rowNum) {
  const ss        = activeSheet.getParent();
  const archSheet = ss.getSheetByName(ARCHIVE_SHEET);

  if (!archSheet) {
    SpreadsheetApp.getUi().alert('Could not find the "Completed Deliveries" sheet.');
    return;
  }

  const archData = archSheet.getRange("A:A").getValues();
  let archLastRow = 1;
  for (let i = archData.length - 1; i >= 1; i--) {
    if (archData[i][0] !== "" && archData[i][0] !== null && archData[i][0] !== false) {
      archLastRow = i + 1;
      break;
    }
  }
  const destRow   = archLastRow + 1;

  // v3.3: copy the FULL row (all 114 cols) and FREEZE to static values so
  // consumption data (BX–CJ) travels to Completed Deliveries. This is what
  // keeps inventory On Hand correct after delivery (v3.2 truncated at col 60).
  const destRange = copyRowFull_(activeSheet, rowNum, archSheet, destRow);

  const archiveNote = "Archived: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy h:mm a");
  archSheet.getRange(destRow, STATUS_COL).setNote(archiveNote);
  destRange.setBackground("#EAF3DE").setFontColor("#1E4620");

  const colBValue = String(activeSheet.getRange(rowNum, 2).getValue()).trim().toUpperCase();
  if (colBValue !== TOTALS_ROW_LABEL) {
    activeSheet.deleteRow(rowNum);
  }

  ensureTotalsRow_(activeSheet);
  ss.toast("✓ Row archived to Completed Deliveries", "Done", 3);
}


// ═══════════════════════════════════════════════════════════════════════════
//  ARCHIVE CANCELLED ROW — moves to Cancelled Requests tab
// ═══════════════════════════════════════════════════════════════════════════

function archiveCancelledRow_(activeSheet, rowNum) {
  const ss           = activeSheet.getParent();
  const cancelSheet  = ss.getSheetByName(CANCELLED_SHEET);

  if (!cancelSheet) {
    SpreadsheetApp.getUi().alert('Could not find the "Cancelled Requests" sheet.');
    return;
  }

  // Find next empty row in Cancelled Requests
  const cancelData = cancelSheet.getRange("A:A").getValues();
  let cancelLastRow = 1;
  for (let i = cancelData.length - 1; i >= 1; i--) {
    if (cancelData[i][0] !== "" && cancelData[i][0] !== null && cancelData[i][0] !== false) {
      cancelLastRow = i + 1;
      break;
    }
  }
  const destRow   = cancelLastRow + 1;

  // v3.3: full-width, value-frozen copy (carries Pack Config + demand + consumption)
  const destRange = copyRowFull_(activeSheet, rowNum, cancelSheet, destRow);

  // Stamp a note on the status cell
  const cancelNote = "Cancelled: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy h:mm a");
  cancelSheet.getRange(destRow, STATUS_COL).setNote(cancelNote);

  // Apply cancelled styling to the archived row
  styleRowCancelled_(cancelSheet, destRow);

  // Delete from Waiting List
  const colBValue = String(activeSheet.getRange(rowNum, 2).getValue()).trim().toUpperCase();
  if (colBValue !== TOTALS_ROW_LABEL) {
    activeSheet.deleteRow(rowNum);
  }

  ensureTotalsRow_(activeSheet);
  ss.toast("✓ Request cancelled and moved to Cancelled Requests tab", "Cancelled", 3);
}


// ═══════════════════════════════════════════════════════════════════════════
//  RESTORE LAST CANCELLED ROW
// ═══════════════════════════════════════════════════════════════════════════

function restoreLastCancelled() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cancelSheet = ss.getSheetByName(CANCELLED_SHEET);
  const activeSheet = ss.getSheetByName(ACTIVE_SHEET);

  if (!cancelSheet || !activeSheet) {
    ui.alert("Could not find required sheets.");
    return;
  }

  ensureTotalsRow_(activeSheet);

  // Find last non-empty row in Cancelled Requests
  const cancelData = cancelSheet.getRange("A:A").getValues();
  let cancelLastRow = 1;
  for (let i = cancelData.length - 1; i >= 1; i--) {
    if (cancelData[i][0] !== "" && cancelData[i][0] !== null && cancelData[i][0] !== false) {
      cancelLastRow = i + 1;
      break;
    }
  }

  if (cancelLastRow <= 1) {
    ui.alert("No cancelled rows found.");
    return;
  }

  const apn = cancelSheet.getRange(cancelLastRow, COL_APN).getValue();
  const statusNote = cancelSheet.getRange(cancelLastRow, STATUS_COL).getNote() || "(no timestamp)";

  const response = ui.alert(
    "Restore Cancelled Row",
    `Restore "${apn}" (${statusNote}) to the Waiting List as Active?\n\nThis will remove it from Cancelled Requests.`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Find insert position above TOTALS
  const lastRow = activeSheet.getLastRow();
  let insertRow = lastRow + 1;
  for (let r = 2; r <= lastRow; r++) {
    if (String(activeSheet.getRange(r, 2).getValue()) === TOTALS_ROW_LABEL) {
      insertRow = r;
      break;
    }
  }

  activeSheet.insertRowBefore(insertRow);
  // v3.3: full-width copy so consumption/demand data returns with the row
  copyRowFull_(cancelSheet, cancelLastRow, activeSheet, insertRow);

  // Reset to Active
  activeSheet.getRange(insertRow, STATUS_COL).setValue("Active");
  writeRowFormulas_(activeSheet, insertRow);
  styleRowActive_(activeSheet, insertRow);

  const restoreNote = activeSheet.getRange(insertRow, STATUS_COL).getNote() || "";
  activeSheet.getRange(insertRow, STATUS_COL).setNote(
    restoreNote + "\nRestored from Cancelled: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy h:mm a")
  );

  cancelSheet.deleteRow(cancelLastRow);
  ss.toast(`✓ "${apn}" restored to Waiting List as Active`, "Restored", 4);
}




function styleRowCancelled_(sheet, rowNum) {
  const lastCol = sheet.getLastColumn();
  sheet.getRange(rowNum, 1, 1, lastCol).setBackground("#FCE4D6").setFontColor("#C00000");
  TEXT_COLS_FOR_STRIKETHROUGH.forEach(col => {
    if (col <= lastCol) {
      sheet.getRange(rowNum, col).setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(true).build());
    }
  });
  sheet.getRange(rowNum, STATUS_COL).setFontWeight("bold").setHorizontalAlignment("center");
}

function styleRowActive_(sheet, rowNum) {
  const lastCol = sheet.getLastColumn();
  const range   = sheet.getRange(rowNum, 1, 1, lastCol);
  range.setBackground("#FFFFFF").setFontColor("#000000");
  range.setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(false).build());
  sheet.getRange(rowNum, STATUS_COL).setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#155724").setBackground("#D4EDDA");
}

function styleRowPacked_(sheet, rowNum) {
  const lastCol = sheet.getLastColumn();
  const range   = sheet.getRange(rowNum, 1, 1, lastCol);
  // Amber highlight across the row to signal "reserved / ready to deliver"
  range.setBackground("#FFF3CD").setFontColor("#5A4520");
  range.setTextStyle(SpreadsheetApp.newTextStyle().setStrikethrough(false).build());
  sheet.getRange(rowNum, STATUS_COL).setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#5A4520").setBackground("#FFE08A");
}


// ═══════════════════════════════════════════════════════════════════════════
//  RESTORE LAST ARCHIVED ROW
// ═══════════════════════════════════════════════════════════════════════════

function restoreLastArchived() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const archSheet = ss.getSheetByName(ARCHIVE_SHEET);
  const activeSheet = ss.getSheetByName(ACTIVE_SHEET);

  if (!archSheet || !activeSheet) {
    ui.alert("Could not find required sheets.");
    return;
  }

  ensureTotalsRow_(activeSheet);

  const archData = archSheet.getRange("A:A").getValues();
  let archLastRow = 1;
  for (let i = archData.length - 1; i >= 1; i--) {
    if (archData[i][0] !== "" && archData[i][0] !== null && archData[i][0] !== false) {
      archLastRow = i + 1;
      break;
    }
  }

  if (archLastRow <= 1) {
    ui.alert("No archived rows found.");
    return;
  }

  const apn = archSheet.getRange(archLastRow, COL_APN).getValue();
  const statusNote = archSheet.getRange(archLastRow, STATUS_COL).getNote() || "(no timestamp)";

  const response = ui.alert(
    "Restore Archived Row",
    `Restore "${apn}" (${statusNote}) to the Waiting List?\n\nThis will remove it from Completed Deliveries.`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const lastRow = activeSheet.getLastRow();
  let insertRow = lastRow + 1;
  for (let r = 2; r <= lastRow; r++) {
    if (String(activeSheet.getRange(r, 2).getValue()) === TOTALS_ROW_LABEL) {
      insertRow = r;
      break;
    }
  }

  activeSheet.insertRowBefore(insertRow);

  // v3.3: full-width copy so consumption/demand data returns with the row
  copyRowFull_(archSheet, archLastRow, activeSheet, insertRow);

  activeSheet.getRange(insertRow, STATUS_COL).setValue("Active");
  writeRowFormulas_(activeSheet, insertRow);
  styleRowActive_(activeSheet, insertRow);

  const restoreNote = activeSheet.getRange(insertRow, STATUS_COL).getNote() || "";
  activeSheet.getRange(insertRow, STATUS_COL).setNote(
    restoreNote + "\nRestored: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy h:mm a")
  );

  archSheet.deleteRow(archLastRow);
  ss.toast(`✓ "${apn}" restored to Waiting List`, "Restored", 4);
}


// ═══════════════════════════════════════════════════════════════════════════
//  PICK LIST GENERATION
// ═══════════════════════════════════════════════════════════════════════════

function generatePickList() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const wl  = ss.getSheetByName(ACTIVE_SHEET);
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const cfg = ss.getSheetByName(CONFIG_SHEET);

  if (!wl || !inv) {
    ui.alert("Could not find required sheets.");
    return;
  }

  // ── Determine which row to operate on ──
  let rowNum = SpreadsheetApp.getActiveRange().getRow();
  const selectedSheet = SpreadsheetApp.getActiveSheet().getName();
  if (selectedSheet !== ACTIVE_SHEET || rowNum <= HEADER_ROW) {
    const resp = ui.prompt("Generate Pick List",
      "Enter the row number on the Waiting List:", ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    rowNum = parseInt(resp.getResponseText(), 10);
    if (isNaN(rowNum) || rowNum <= HEADER_ROW) {
      ui.alert("Invalid row number."); return;
    }
  }

  const status = wl.getRange(rowNum, STATUS_COL).getValue();
  if (status !== "Active" && status !== "Packed") {
    ui.alert("This request is not Active or Packed (status: " + status +
             "). Pick lists can only be generated for Active or Packed requests.");
    return;
  }

  const req = readPickListRequest_(wl, rowNum);

  // ── Bed count validation vs. children (menu context — ui.alert is fine here) ──
  const totalBeds = req.twins + req.toddlers + (req.bunks * 2) + req.cribs;
  if (totalBeds === 0) {
    ui.alert("This request has no beds. Nothing to pick.");
    return;
  }
  if (req.children.length > totalBeds) {
    const resp = ui.alert("More children than beds",
      "Row " + rowNum + " has " + req.children.length + " children but only " +
      totalBeds + " sleeping surfaces (bunks count as 2). Continue anyway?",
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  // ── Open the pack-time input dialog (frames + optional substitutions/omissions) ──
  // The dialog collects choices client-side and calls finalizePickList(rowNum, json),
  // which runs the allocation, builds the PDF and writes consumption/config. We can't
  // collect dropdowns via ui.prompt, and a google.script.run callback can't open
  // server-side UI — so the result PDF link is rendered inside the dialog itself.
  const html = buildPickListDialogHtml_(rowNum, req, inv, cfg);
  ui.showModalDialog(html, "Generate Pick List — " + req.apn);
}

// Read a Waiting List request row into a normalized object used by both the dialog
// builder and finalizePickList. (No UI — safe to call from a google.script.run context.)
function readPickListRequest_(wl, rowNum) {
  const rowData = wl.getRange(rowNum, 1, 1, V3_COL_CONSUMPTION_LAST).getValues()[0];
  // Age cols: X(24), Z(26), AB(28), AD(30), AF(32), AH(34), AJ(36), AL(38)
  // Gender cols: Y(25), AA(27), AC(29), AE(31), AG(33), AI(35), AK(37), AM(39)
  const ageCols    = [24, 26, 28, 30, 32, 34, 36, 38];
  const genderCols = [25, 27, 29, 31, 33, 35, 37, 39];
  const children = [];
  for (let i = 0; i < ageCols.length; i++) {
    const age = rowData[ageCols[i] - 1];
    const gen = rowData[genderCols[i] - 1];
    if (age !== "" && age !== null && age !== undefined) {
      children.push({
        age: Number(age),
        gender: v3GenderLabel(gen),
        genderRaw: String(gen || "").trim(),
        tier: v3AgeTier(age),
      });
    }
  }
  // Previously-saved fulfillment config lives in the appended column (115), which is
  // beyond rowData's width — read it separately for re-run prefill.
  let savedFulfillJson = "";
  try { savedFulfillJson = String(wl.getRange(rowNum, V3_COL_FULFILL_CONFIG).getValue() || ""); } catch (e) {}

  return {
    rowData:  rowData,
    apn:      rowData[COL_APN - 1]      || "Unknown",
    twins:    Number(rowData[COL_TWINS - 1])    || 0,
    bunks:    Number(rowData[COL_BUNKS - 1])    || 0,
    toddlers: Number(rowData[COL_TODDLERS - 1]) || 0,
    cribs:    Number(rowData[COL_CRIBS - 1])    || 0,
    children: children,
    savedFulfillJson: savedFulfillJson,
  };
}

// Build the ordered bedList ({type, isConstructed}) from the dialog's frame choices.
// choices.frames is an array of {type:'Twin'|'Toddler'|'Bunk'|'Crib', constructed:bool}.
// Returns null if the selections don't match the request's bed counts.
function buildBedListFromChoices_(choices, counts) {
  const frames = (choices && choices.frames) || [];
  const bedList = [];
  frames.forEach(f => {
    bedList.push({ type: f.type, isConstructed: f.type === "Crib" ? false : !!f.constructed });
  });
  const n = t => bedList.filter(b => b.type === t).length;
  if (n("Twin") !== counts.twins || n("Toddler") !== counts.toddlers ||
      n("Bunk") !== counts.bunks || n("Crib") !== counts.cribs) {
    return null;
  }
  return bedList;
}

// Serialize bedList to the BI Pack Config string ("T:C,TD:P,BK:C,CR:P").
function packConfigStr_(bedList) {
  return bedList.map(b => {
    const t = b.type === "Twin" ? "T" :
              b.type === "Toddler" ? "TD" :
              b.type === "Bunk" ? "BK" : "CR";
    return t + ":" + (b.isConstructed ? "C" : "P");
  }).join(",");
}

// Expand a bedList into ordered sleeping surfaces (bunks → 2). Used by BOTH the
// dialog builder and finalizePickList so surface indices line up with the choices.
// Each surface: {bedType, isConstructed, mattressType:'6in'|'8in'|'crib', position?,
//   label, substitutable}. label is the human caption shown in the dialog.
function buildSleepingSurfaces_(bedList) {
  const surfaces = [];
  const counts = { Twin: 0, Toddler: 0, Bunk: 0, Crib: 0 };
  bedList.forEach(b => {
    counts[b.type] = (counts[b.type] || 0) + 1;
    const num = counts[b.type];
    if (b.type === "Crib") {
      surfaces.push({ bedType: "Crib", isConstructed: false, mattressType: "crib", label: "Crib #" + num, substitutable: false });
    } else if (b.type === "Twin") {
      surfaces.push({ bedType: "Twin", isConstructed: b.isConstructed, mattressType: "8in", label: "Twin Bed #" + num, substitutable: true });
    } else if (b.type === "Toddler") {
      surfaces.push({ bedType: "Toddler", isConstructed: b.isConstructed, mattressType: "6in", label: "Toddler Bed #" + num, substitutable: true });
    } else if (b.type === "Bunk") {
      surfaces.push({ bedType: "Bunk", isConstructed: b.isConstructed, mattressType: "6in", position: "top", label: "Bunk #" + num + " — top", substitutable: true });
      surfaces.push({ bedType: "Bunk", isConstructed: b.isConstructed, mattressType: "8in", position: "bottom", label: "Bunk #" + num + " — bottom", substitutable: true });
    }
  });
  return surfaces;
}

// Parse the saved fulfillment-config JSON (column 115) into {choices,notes} or null.
function readSavedFulfillConfig_(jsonStr) {
  if (!jsonStr) return null;
  try {
    const o = JSON.parse(jsonStr);
    return (o && o.choices) ? o : null;
  } catch (e) { return null; }
}

// Build the pick-list input dialog (frames + optional substitutions/omissions).
// Returns an HtmlOutput for ui.showModalDialog.
function buildPickListDialogHtml_(rowNum, req, inv, cfg) {
  // On-hand for annotations.
  const onHand = {};
  try {
    const vals = inv.getRange(V3_INV_FIRST_ROW, 6, V3_ITEM_COUNT, 1).getValues();
    V3_CATALOG.forEach((it, i) => { onHand[it.name] = Number(vals[i][0]) || 0; });
  } catch (e) {}
  const total6 = (onHand["Mattress 6in 38x74"] || 0) + (onHand["Mattress 6in 39x75"] || 0);
  const total8 = (onHand["Mattress 8in 38x74"] || 0) + (onHand["Mattress 8in 39x75"] || 0);

  // Ordered frame list — twins, toddlers, bunks, cribs (matches finalize's bedList order).
  const frames = [];
  for (let i = 0; i < req.twins; i++)    frames.push({ type: "Twin",    label: "Twin Bed #" + (i + 1),    hasCP: true });
  for (let i = 0; i < req.toddlers; i++) frames.push({ type: "Toddler", label: "Toddler Bed #" + (i + 1), hasCP: true });
  for (let i = 0; i < req.bunks; i++)    frames.push({ type: "Bunk",    label: "Bunk Bed #" + (i + 1),    hasCP: true });
  for (let i = 0; i < req.cribs; i++)    frames.push({ type: "Crib",    label: "Crib #" + (i + 1),        hasCP: false });

  // Sleeping surfaces for mattress dropdowns — same order; C/P doesn't change the list.
  const previewBedList = frames.map(f => ({ type: f.type, isConstructed: false }));
  const allSurfaces = buildSleepingSurfaces_(previewBedList);
  const surfaces = allSurfaces
    .map((s, idx) => ({ idx: idx, label: s.label, defaultThickness: s.mattressType, substitutable: s.substitutable }))
    .filter(s => s.substitutable);

  // Children for comforter dropdowns (index = child order, matches finalize ci).
  const numChildrenToBed = Math.min(req.children.length, allSurfaces.length);
  const children = req.children.slice(0, numChildrenToBed).map((c, i) => ({
    idx: i,
    label: "Child " + (i + 1) + " (age " + c.age + (c.genderRaw ? ", " + c.genderRaw : "") + ")"
  }));

  // Quantity items with computed standard defaults where independent of frame C/P.
  // v3.6: items switched off in Ministry Options are hidden (null) — the allocator
  // then defaults them to 0.
  const opts = getMinistryOptions_();
  const surfAll = req.twins + req.toddlers + req.bunks * 2 + req.cribs;
  const bears = req.children.filter(c => c.age >= 0 && c.age <= V3_BEAR_AGE_MAX).length;
  const tiers = { Toddler: false, Young: false, Teen: false };
  req.children.forEach(c => { if (c.tier) tiers[c.tier] = true; });
  const qtyDef = {
    "Pillow": surfAll,
    "Twin Mattress Protector": Math.max(0, surfAll - req.cribs),
    "Bedrail": opts.bedrails ? "" : null,   // "" = depends on C/P → blank = standard
    "Plaque":  opts.plaques  ? "" : null,
    "Bear":    opts.bears    ? bears : null,
    "Toddler Book": (opts.books && tiers.Toddler) ? 1 : null,
    "Young Book":   (opts.books && tiers.Young) ? 1 : null,
    "Teen Book":    (opts.books && tiers.Teen) ? 1 : null,
  };
  const qtyItems = V3_QTY_ADJUSTABLE
    .filter(q => qtyDef[q.name] !== null)   // hide absent book tiers + disabled options
    .map(q => ({ name: q.name, label: q.label, def: qtyDef[q.name], onHand: onHand[q.name] || 0 }));

  // Re-run prefill from a saved config, but only if bed counts still match.
  const saved = readSavedFulfillConfig_(req.savedFulfillJson);
  let prefillFrames = null, savedChoices = null, hasOverrides = false;
  if (saved && saved.choices) {
    const sc = saved.choices;
    const cnt = arr => {
      const c = { Twin: 0, Toddler: 0, Bunk: 0, Crib: 0 };
      (arr || []).forEach(f => { if (c[f.type] !== undefined) c[f.type]++; });
      return c;
    };
    const sf = cnt(sc.frames);
    if (sf.Twin === req.twins && sf.Toddler === req.toddlers &&
        sf.Bunk === req.bunks && sf.Crib === req.cribs) {
      prefillFrames = sc.frames;
      savedChoices = sc;
      hasOverrides = !!((sc.mattress && Object.keys(sc.mattress).length) ||
                        (sc.comforter && Object.keys(sc.comforter).length) ||
                        (sc.qty && Object.keys(sc.qty).length));
    }
  }

  const model = {
    rowNum: rowNum,
    frames: frames.map(f => ({ type: f.type, hasCP: f.hasCP })),
    surfaces: surfaces.map(s => ({ idx: s.idx, defaultThickness: s.defaultThickness })),
    children: children.map(c => ({ idx: c.idx })),
    qtyItems: qtyItems.map(q => ({ name: q.name })),
  };

  return HtmlService
    .createHtmlOutput(pickListDialogHtml_({
      apn: req.apn, frames: frames, surfaces: surfaces, children: children,
      qtyItems: qtyItems, total6: total6, total8: total8,
      gateDefault: hasOverrides ? "yes" : "no",
      prefillFrames: prefillFrames, savedChoices: savedChoices, model: model
    }))
    .setWidth(580).setHeight(660);
}

// Render the dialog HTML string. Controls are built server-side (easy prefill); a small
// client script toggles the override section and submits choices via google.script.run.
function pickListDialogHtml_(d) {
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const sel = (cond) => cond ? " selected" : "";

  // Frame controls
  let framesHtml = "";
  d.frames.forEach((f, i) => {
    let control;
    if (f.hasCP) {
      const pf = d.prefillFrames && d.prefillFrames[i];
      const isC = pf ? !!pf.constructed : false;
      control =
        '<select data-fi="' + i + '">' +
          '<option value="C"' + sel(isC) + '>Constructed</option>' +
          '<option value="P"' + sel(!isC) + '>Purchased</option>' +
        '</select>';
    } else {
      control = '<span class="na">no frame choice</span>';
    }
    framesHtml += '<div class="row"><label>' + esc(f.label) + '</label>' + control + '</div>';
  });

  // Mattress controls
  let mattHtml = "";
  d.surfaces.forEach(s => {
    const saved = d.savedChoices && d.savedChoices.mattress && d.savedChoices.mattress[s.idx];
    const cur = (saved === "6in" || saved === "8in") ? saved : s.defaultThickness;
    mattHtml += '<div class="row"><label>' + esc(s.label) + '</label>' +
      '<select data-mi="' + s.idx + '">' +
        '<option value="6in"' + sel(cur === "6in") + '>6 inch' + (s.defaultThickness === "6in" ? " (standard)" : "") + '</option>' +
        '<option value="8in"' + sel(cur === "8in") + '>8 inch' + (s.defaultThickness === "8in" ? " (standard)" : "") + '</option>' +
      '</select></div>';
  });
  if (mattHtml) {
    mattHtml = '<div class="hint">On hand — 6": ' + d.total6 + ' · 8": ' + d.total8 + '</div>' + mattHtml;
  }

  // Comforter controls
  let comfHtml = "";
  d.children.forEach(c => {
    const saved = d.savedChoices && d.savedChoices.comforter && d.savedChoices.comforter[c.idx];
    const without = saved === "without";
    comfHtml += '<div class="row"><label>' + esc(c.label) + '</label>' +
      '<select data-ci="' + c.idx + '">' +
        '<option value="with"' + sel(!without) + '>With sheets</option>' +
        '<option value="without"' + sel(without) + '>Without sheets (+ Twin Sheets)</option>' +
      '</select></div>';
  });

  // Quantity controls
  let qtyHtml = "";
  d.qtyItems.forEach((q, qi) => {
    let val = (q.def === null || q.def === undefined) ? "" : q.def;
    if (d.savedChoices && d.savedChoices.qty &&
        Object.prototype.hasOwnProperty.call(d.savedChoices.qty, q.name)) {
      val = d.savedChoices.qty[q.name];
    }
    qtyHtml += '<div class="row"><label>' + esc(q.label) +
      ' <span class="oh">(on hand: ' + q.onHand + ')</span></label>' +
      '<input type="number" min="0" step="1" data-qi="' + qi + '" value="' + val +
      '" placeholder="standard"></div>';
  });

  const modelJson = JSON.stringify(d.model);

  return '' +
'<style>' +
'  body{font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#1A3C5E;margin:0;padding:12px;}' +
'  h3{margin:6px 0 2px;font-size:14px;color:#1A3C5E;}' +
'  .sect{border:1px solid #D7E0EA;border-radius:6px;padding:8px 10px;margin:8px 0;}' +
'  .sect h4{margin:0 0 6px;font-size:12px;color:#4A6B8A;text-transform:uppercase;letter-spacing:.04em;}' +
'  .row{display:flex;align-items:center;justify-content:space-between;margin:4px 0;gap:8px;}' +
'  .row label{flex:1;}' +
'  select,input[type=number]{font-size:13px;padding:3px 4px;min-width:140px;}' +
'  input[type=number]{width:80px;min-width:0;}' +
'  .na{color:#999;font-style:italic;}' +
'  .oh{color:#888;font-weight:normal;font-size:11px;}' +
'  .hint{color:#666;font-size:11px;margin:2px 0 6px;}' +
'  .gate{background:#EAF3DE;border:1px solid #BFD8A0;border-radius:6px;padding:8px 10px;margin:8px 0;}' +
'  #overrides{display:none;}' +
'  .btns{display:flex;gap:8px;justify-content:flex-end;margin-top:10px;}' +
'  button{font-size:13px;padding:6px 14px;border-radius:5px;border:1px solid #1A3C5E;cursor:pointer;}' +
'  .primary{background:#1A3C5E;color:#fff;}' +
'  .secondary{background:#fff;color:#1A3C5E;}' +
'  #result{margin-top:10px;}' +
'  #result a{font-weight:bold;color:#1A3C5E;font-size:15px;}' +
'  .err{color:#C62828;}' +
'  .warnbox{background:#FFF8E1;border:1px solid #E0C870;border-radius:5px;padding:6px 8px;margin-top:6px;font-size:12px;}' +
'  .busy{color:#4A6B8A;}' +
'</style>' +
'<div id="form">' +
'  <div class="sect"><h4>Frames</h4>' + framesHtml + '</div>' +
'  <div class="gate">' +
'    <label style="font-weight:bold;">Do you need to modify the standard configuration for this request?</label>' +
'    <select id="gate" onchange="toggleOverrides()">' +
'      <option value="no"'  + sel(d.gateDefault === "no")  + '>No — use the standard configuration</option>' +
'      <option value="yes"' + sel(d.gateDefault === "yes") + '>Yes — substitute or omit items</option>' +
'    </select>' +
'  </div>' +
'  <div id="overrides">' +
     (mattHtml ? '<div class="sect"><h4>Mattresses</h4>' + mattHtml + '</div>' : '') +
     (comfHtml ? '<div class="sect"><h4>Comforters</h4>' + comfHtml + '</div>' : '') +
     (qtyHtml  ? '<div class="sect"><h4>Quantities (0 = omit; blank = standard)</h4>' + qtyHtml + '</div>' : '') +
'  </div>' +
'  <div class="btns">' +
'    <button class="secondary" onclick="google.script.host.close()">Cancel</button>' +
'    <button class="primary" id="go" onclick="submitChoices()">Generate Pick List</button>' +
'  </div>' +
'</div>' +
'<div id="result"></div>' +
'<script>' +
'  var MODEL = ' + modelJson + ';' +
'  function toggleOverrides(){' +
'    document.getElementById("overrides").style.display =' +
'      document.getElementById("gate").value === "yes" ? "block" : "none";' +
'  }' +
'  toggleOverrides();' +
'  function gather(){' +
'    var c = { frames: [], mattress:{}, comforter:{}, qty:{} };' +
'    MODEL.frames.forEach(function(f,i){' +
'      if (f.hasCP){' +
'        var s = document.querySelector(\'[data-fi="\'+i+\'"]\');' +
'        c.frames.push({ type: f.type, constructed: s.value === "C" });' +
'      } else { c.frames.push({ type: f.type }); }' +
'    });' +
'    if (document.getElementById("gate").value === "yes"){' +
'      MODEL.surfaces.forEach(function(s){' +
'        var el = document.querySelector(\'[data-mi="\'+s.idx+\'"]\');' +
'        if (el && el.value && el.value !== s.defaultThickness) c.mattress[s.idx] = el.value;' +
'      });' +
'      MODEL.children.forEach(function(ch){' +
'        var el = document.querySelector(\'[data-ci="\'+ch.idx+\'"]\');' +
'        if (el && el.value === "without") c.comforter[ch.idx] = "without";' +
'      });' +
'      MODEL.qtyItems.forEach(function(q,qi){' +
'        var el = document.querySelector(\'[data-qi="\'+qi+\'"]\');' +
'        if (el && el.value !== ""){ var n = Math.floor(Number(el.value));' +
'          if (!isNaN(n)){ if (n<0) n=0; c.qty[q.name] = n; } }' +
'      });' +
'    }' +
'    return c;' +
'  }' +
'  function submitChoices(){' +
'    document.getElementById("go").disabled = true;' +
'    document.getElementById("result").innerHTML = \'<span class="busy">Generating pick list…</span>\';' +
'    google.script.run' +
'      .withSuccessHandler(onDone)' +
'      .withFailureHandler(onErr)' +
'      .finalizePickList(MODEL.rowNum, JSON.stringify(gather()));' +
'  }' +
'  function onErr(e){' +
'    document.getElementById("go").disabled = false;' +
'    document.getElementById("result").innerHTML = \'<span class="err">Error: \'+ (e && e.message ? e.message : e) +\'</span>\';' +
'  }' +
'  function onDone(res){' +
'    var r = document.getElementById("result");' +
'    if (!res || !res.ok){' +
'      document.getElementById("go").disabled = false;' +
'      r.innerHTML = \'<span class="err">\'+ (res && res.error ? res.error : "Something went wrong.") +\'</span>\';' +
'      return;' +
'    }' +
'    var html = "";' +
'    if (res.url){ html += \'<p>✓ Pick list saved. <a href="\'+res.url+\'" target="_blank">📄 Open Pick List PDF</a></p>\'; }' +
'    else if (res.pdfError){ html += \'<p class="err">\'+res.pdfError+\'</p>\'; }' +
'    if (res.notes && res.notes.length){' +
'      html += \'<div class="warnbox"><b>Substitutions / changes:</b><ul>\';' +
'      res.notes.forEach(function(n){ html += "<li>"+n+"</li>"; });' +
'      html += "</ul></div>";' +
'    }' +
'    if (res.shortages && res.shortages.length){' +
'      html += \'<div class="warnbox err"><b>⚠ Short on inventory:</b><ul>\';' +
'      res.shortages.forEach(function(s){ html += "<li>"+ s.shortBy +" "+ s.item +"</li>"; });' +
'      html += "</ul></div>";' +
'    }' +
'    html += \'<div class="btns"><button class="primary" onclick="google.script.host.close()">Done</button></div>\';' +
'    document.getElementById("form").style.display = "none";' +
'    r.innerHTML = html;' +
'  }' +
'</script>';
}

// ═══════════════════════════════════════════════════════════════════════════
//  FINALIZE PICK LIST — called from the dialog via google.script.run.
//  Runs allocation (honoring substitutions/omissions), writes consumption + the
//  chosen-config JSON, builds the PDF, and returns {ok,url,notes,shortages} or
//  {ok:false,error}. Invoked via google.script.run, so it MUST NOT call any ui.*.
// ═══════════════════════════════════════════════════════════════════════════
function finalizePickList(rowNum, choicesJson) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const wl  = ss.getSheetByName(ACTIVE_SHEET);
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const cfg = ss.getSheetByName(CONFIG_SHEET);
  try {
    if (!wl || !inv) return { ok: false, error: "Could not find required sheets." };
    const choices = (typeof choicesJson === "string")
      ? JSON.parse(choicesJson || "{}") : (choicesJson || {});

    const status = wl.getRange(rowNum, STATUS_COL).getValue();
    if (status !== "Active" && status !== "Packed") {
      return { ok: false, error: "Request is no longer Active or Packed (status: " + status + ")." };
    }

    const req = readPickListRequest_(wl, rowNum);
    const rowData  = req.rowData;
    const apn      = req.apn;
    const twins    = req.twins;
    const bunks    = req.bunks;
    const toddlers = req.toddlers;
    const cribs    = req.cribs;
    const children = req.children;

    // Build bedList from the dialog's frame choices and persist Pack Config (BI).
    const bedList = buildBedListFromChoices_(choices, { twins, toddlers, bunks, cribs });
    if (!bedList) return { ok: false, error: "Frame selections were incomplete — please re-open the pick list." };
    wl.getRange(rowNum, V3_COL_PACK_CONFIG).setValue(packConfigStr_(bedList));

  // (Request data, children, bedList and Pack Config are all prepared above from the
  //  dialog's choices; allocation continues below honoring any substitutions/omissions.)

  // ── Read inventory On Hand for all 39 items ──
  const onHandVals = inv.getRange(V3_INV_FIRST_ROW, 6, V3_ITEM_COUNT, 1).getValues();
  const onHand = {};
  V3_CATALOG.forEach((item, idx) => {
    onHand[item.name] = Number(onHandVals[idx][0]) || 0;
  });
  // Mutable remaining quantities during allocation
  const avail = Object.assign({}, onHand);

  // ── Read book titles from Config!B24:B26 ──
  let bookTitles = { "Toddler": "", "Young": "", "Teen": "" };
  if (cfg) {
    try {
      const tv = cfg.getRange("B24:B26").getValues();
      bookTitles.Toddler = String(tv[0][0] || "").trim();
      bookTitles.Young   = String(tv[1][0] || "").trim();
      bookTitles.Teen    = String(tv[2][0] || "").trim();
    } catch(e) { /* titles remain blank */ }
  }

  // ── Allocate items ──
  // consumption: { itemName: qty } — what the pick list consumes
  const consumption = {};
  V3_CATALOG.forEach(i => { consumption[i.name] = 0; });

  const take = (itemName, qty) => {
    if (qty <= 0) return;
    consumption[itemName] = (consumption[itemName] || 0) + qty;
    avail[itemName] = (avail[itemName] || 0) - qty;
  };

  const shortages = [];  // [{item, shortBy}]
  const notes = [];      // Human-readable notes about substitutions/omissions

  // ── Pack-time overrides chosen in the dialog (empty object = standard config) ──
  const qtyChoice       = (choices && choices.qty)       || {};   // { itemName: number≥0 }
  const mattressChoice  = (choices && choices.mattress)  || {};   // { surfaceIndex: '6in'|'8in' }
  const comforterChoice = (choices && choices.comforter) || {};   // { childIndex: 'with'|'without' }
  // Resolve a quantity for an adjustable "extra": chosen value if present (≥0), else default.
  const qtyFor = (itemName, def) => {
    if (Object.prototype.hasOwnProperty.call(qtyChoice, itemName)) {
      let v = Math.floor(Number(qtyChoice[itemName]));
      if (!isNaN(v)) {
        if (v < 0) v = 0;
        if (v !== def) notes.push(itemName + ": " + def + " → " + v + " (adjusted).");
        return v;
      }
    }
    return def;
  };

  // Build per-bed allocation with sleeping-surface expansion for bunks (shared helper
  // so indices match the dialog's mattress dropdowns).
  const sleepingSurfaces = buildSleepingSurfaces_(bedList);

  // ── FRAMES ──
  bedList.forEach(b => {
    if (b.type === "Crib") {
      take("Crib", 1);
    } else if (b.type === "Twin") {
      take(b.isConstructed ? "Twin Frame - Constructed" : "Twin Frame - Purchased", 1);
    } else if (b.type === "Toddler") {
      take(b.isConstructed ? "Toddler Frame - Constructed" : "Toddler Frame - Purchased", 1);
    } else if (b.type === "Bunk") {
      take(b.isConstructed ? "Bunk Frame - Constructed" : "Bunk Frame - Purchased", 1);
    }
  });

  // ── MATTRESSES (thickness substitutable per surface, both directions) ──
  // Thickness defaults to the bed's standard (toddler/bunk-top = 6", twin/bunk-bottom
  // = 8") but the volunteer may override either way per surface. Size SKU stays
  // automatic: purchased → 38x74 (so 39x75 is NEVER put on a purchased bed, regardless
  // of the thickness choice); constructed → prefer 39x75, fall back to 38x74.
  // Crib mattresses: just "Crib Mattress" (not substitutable).
  sleepingSurfaces.forEach((surf, si) => {
    if (surf.mattressType === "crib") {
      take("Crib Mattress", 1);
      return;
    }
    let thickness = surf.mattressType; // '6in' or '8in'
    const ch = mattressChoice[si];
    if (ch === "6in" || ch === "8in") {
      if (ch !== thickness) {
        notes.push(surf.label + ": " + thickness + " mattress → " + ch + " (substituted).");
      }
      thickness = ch;
    }
    if (!surf.isConstructed) {
      // Purchased bed: must be 38x74
      take(`Mattress ${thickness} 38x74`, 1);
    } else {
      // Constructed bed: prefer 39x75, fall back to 38x74 if none available
      const pref = `Mattress ${thickness} 39x75`;
      const alt  = `Mattress ${thickness} 38x74`;
      if (avail[pref] >= 1) {
        take(pref, 1);
      } else {
        take(alt, 1);
        notes.push(`No ${pref} available; using ${alt} on constructed bed.`);
      }
    }
  });

  // ── PILLOWS ── (default 1 per sleeping surface; adjustable) ──
  take("Pillow", qtyFor("Pillow", sleepingSurfaces.length));

  // ── COMFORTERS + SHEETS + MATTRESS PROTECTORS ──
  // Per child with gender+age match. Default prefers the "with sheets" variant; the
  // volunteer may flip a child to "without sheets" (which then also takes a Twin Sheets).
  // Neutral fallback is retained. Mattress protectors: default 1 per non-crib surface;
  // adjustable.
  take("Twin Mattress Protector",
       qtyFor("Twin Mattress Protector", sleepingSurfaces.length - (cribs)));
  // (cribs don't need twin mattress protectors; if cribs get their own
  // mattress protector in future, add a Crib Mattress Protector item.)

  // For each child (up to number of sleeping surfaces), allocate bedding
  const numChildrenToBed = Math.min(children.length, sleepingSurfaces.length);
  for (let ci = 0; ci < numChildrenToBed; ci++) {
    const child = children[ci];
    if (!child.tier) continue;

    // Sheets preference: default with-sheets; volunteer may flip this child to without.
    const pref = comforterChoice[ci];           // 'with' | 'without' | undefined
    const withFirst = (pref !== "without");
    const wsG = `Comforter w/Sheets - ${child.gender} ${child.tier}`;
    const wsN = `Comforter w/Sheets - Neutral ${child.tier}`;
    const woG = `Comforter w/o Sheets - ${child.gender} ${child.tier}`;
    const woN = `Comforter w/o Sheets - Neutral ${child.tier}`;
    // Within each sheets-state, prefer gendered then neutral; order the two states
    // by the chosen preference.
    const tryOrder = [];
    const pushPair = (g, n) => {
      tryOrder.push(g);
      if (child.gender !== "Neutral") tryOrder.push(n);
    };
    if (withFirst) { pushPair(wsG, wsN); pushPair(woG, woN); }
    else           { pushPair(woG, woN); pushPair(wsG, wsN); }

    let picked = null;
    for (const candidate of tryOrder) {
      if (avail[candidate] >= 1) {
        picked = candidate;
        take(candidate, 1);
        break;
      }
    }

    if (picked === null) {
      // No inventory for any variant — take the most-preferred candidate
      // and flag as a shortage. Assigning `picked` ensures the w/o-Sheets
      // companion logic below still runs correctly.
      const fallback = tryOrder[0];
      picked = fallback;
      take(fallback, 1);
      shortages.push({ item: fallback, shortBy: 1, reason: "child " + (ci+1) });
    }

    // If the allocated variant is w/o sheets, we need to add 1 Twin Sheets
    if (picked && picked.indexOf("w/o Sheets") >= 0) {
      take("Twin Sheets", 1);
    }

    // Note when the volunteer's sheets preference couldn't be honored (stock).
    if (pref === "without" && picked.indexOf("w/o Sheets") < 0) {
      notes.push("Child " + (ci + 1) + ": wanted comforter without sheets, only with-sheets available.");
    } else if (pref === "with" && picked.indexOf("w/o Sheets") >= 0) {
      notes.push("Child " + (ci + 1) + ": wanted comforter with sheets, only without-sheets available.");
    }
  }

  // v3.6: extras switched off in Ministry Options default to 0 (the dialog hides
  // them, so qtyFor's default is what gets used).
  const opts = getMinistryOptions_();

  // ── BEDRAILS ── (default: 2 per constructed toddler/bunk bed; adjustable) ──
  let bedrailCount = 0;
  bedList.forEach(b => {
    if (b.isConstructed && (b.type === "Toddler" || b.type === "Bunk")) {
      bedrailCount += 2;
    }
  });
  bedrailCount = qtyFor("Bedrail", opts.bedrails ? bedrailCount : 0);
  if (bedrailCount > 0) take("Bedrail", bedrailCount);

  // ── PLAQUES ── (default: 2/constructed bunk, 1/other constructed bed; adjustable) ──
  let plaqueCount = 0;
  bedList.forEach(b => {
    if (!b.isConstructed) return;
    if (b.type === "Bunk") plaqueCount += 2;
    else if (b.type === "Twin" || b.type === "Toddler") plaqueCount += 1;
  });
  plaqueCount = qtyFor("Plaque", opts.plaques ? plaqueCount : 0);
  if (plaqueCount > 0) take("Plaque", plaqueCount);

  // ── BEARS ── (default: 1 per child age 0-10; adjustable, 0 = omit) ──
  let bearCount = 0;
  children.forEach(c => { if (c.age >= 0 && c.age <= V3_BEAR_AGE_MAX) bearCount++; });
  bearCount = qtyFor("Bear", opts.bears ? bearCount : 0);
  if (bearCount > 0) take("Bear", bearCount);

  // ── BOOKS ── (default: per family, 1 each tier present; adjustable, 0 = omit) ──
  const hasTodlerChild = children.some(c => c.tier === "Toddler");
  const hasYoungChild  = children.some(c => c.tier === "Young");
  const hasTeenChild   = children.some(c => c.tier === "Teen");
  const toddlerBooks = qtyFor("Toddler Book", (opts.books && hasTodlerChild) ? 1 : 0);
  const youngBooks   = qtyFor("Young Book",   (opts.books && hasYoungChild) ? 1 : 0);
  const teenBooks    = qtyFor("Teen Book",    (opts.books && hasTeenChild) ? 1 : 0);
  if (toddlerBooks > 0) take("Toddler Book", toddlerBooks);
  if (youngBooks   > 0) take("Young Book", youngBooks);
  if (teenBooks    > 0) take("Teen Book", teenBooks);

  // ── Compute shortages based on original onHand ──
  V3_CATALOG.forEach(item => {
    const used = consumption[item.name];
    const had  = onHand[item.name];
    if (used > had) {
      // Check if already flagged from comforter allocation
      const already = shortages.find(s => s.item === item.name);
      if (!already) {
        shortages.push({ item: item.name, shortBy: used - had });
      } else if (already.shortBy < used - had) {
        already.shortBy = used - had;
      }
    }
  });

  // ── Build pick list sheet ──
  const existing = ss.getSheetByName("Pick List");
  if (existing) ss.deleteSheet(existing);
  const ps = ss.insertSheet("Pick List");
  ps.setColumnWidth(1, 50);
  ps.setColumnWidth(2, 80);
  ps.setColumnWidth(3, 480);

  let row = 1;

  // Title — ministry name from Ministry Options (Config!E11); falls back to
  // the generic "BED MINISTRY". San Antonio sets its name when it adopts v3.6.
  ps.getRange(row, 1, 1, 3).merge()
    .setValue(getMinistryName_().toUpperCase() + " — PICK LIST")
    .setFontSize(16).setFontWeight("bold").setFontColor("#1A3C5E")
    .setHorizontalAlignment("center");
  row += 2;

  // APN + caregiver + address
  const caregiver      = rowData[COL_CAREGIVER - 1]      || "";
  const caregiverPh    = rowData[COL_CAREGIVER_PH - 1]   || "";
  const complex        = rowData[COL_COMPLEX - 1]        || "";
  const address        = rowData[COL_ADDRESS - 1]        || "";
  const apt            = rowData[COL_APT - 1]            || "";
  const city           = rowData[COL_CITY - 1]           || "";
  const zip            = rowData[COL_ZIP - 1]            || "";
  const delivContact   = rowData[COL_DELIVERY_CONTACT - 1]    || "";
  const delivContactPh = rowData[COL_DELIVERY_CONTACT_PH - 1] || "";
  const delivOrg       = rowData[COL_DELIVERY_ORG - 1]   || "";
  const notesText      = String(rowData[COL_NOTES - 1]   || "");
  const fullAddress    = [complex, address, apt, city, zip].filter(Boolean).join(", ");

  const info = [
    ["APN #:", apn],
    ["Caregiver:", `${caregiver}  ${caregiverPh}`],
    ["Delivery Address:", fullAddress],
    ["Delivery Contact:", `${delivContact}  ${delivContactPh}`],
    ["Delivery Org:", delivOrg],
  ];
  if (children.length > 0) {
    const childStr = children.map(c =>
      c.genderRaw ? `${c.age} (${c.genderRaw})` : `${c.age}`
    ).join(",   ");
    info.push([`Children (${children.length}):`, childStr]);
  }

  info.forEach(([lbl, val]) => {
    ps.getRange(row, 1, 1, 2).merge();
    ps.getRange(row, 1).setValue(lbl).setFontWeight("bold").setFontSize(11);
    ps.getRange(row, 3).setValue(val).setFontSize(11).setWrap(true);
    row++;
  });

  // Notes block (if any)
  if (notesText) {
    row += 1;
    ps.getRange(row, 1, 1, 3).merge()
      .setValue("Notes")
      .setFontSize(11).setFontWeight("bold")
      .setFontColor("#FFFFFF").setBackground("#4A6B8A");
    row++;
    ps.getRange(row, 1, 1, 3).merge()
      .setValue(notesText).setFontSize(11).setWrap(true)
      .setVerticalAlignment("top").setBackground("#FFF8E1");
    ps.setRowHeight(row, Math.max(36, Math.min(120, 20 + Math.ceil(notesText.length / 80) * 16)));
    row++;
  }

  // Bed summary
  row += 1;
  const bedSummaryParts = [];
  bedList.forEach(b => {
    const typ = b.type;
    const cp = b.isConstructed ? "Constructed" : "Purchased";
    if (b.type === "Crib") {
      bedSummaryParts.push("Crib");
    } else {
      bedSummaryParts.push(`${cp} ${typ}`);
    }
  });
  ps.getRange(row, 1, 1, 3).merge()
    .setValue("Beds: " + bedSummaryParts.join(", "))
    .setFontSize(12).setFontWeight("bold").setFontColor("#4A6B8A");
  row += 2;

  // Categorized pick list
  const categoryOrder = ["FRAMES", "MATTRESSES", "BEDDING", "COMFORTERS",
                         "HARDWARE", "BEARS & BOOKS"];
  categoryOrder.forEach(category => {
    const itemsInCat = V3_CATALOG.filter(i => i.category === category);
    const itemsWithQty = itemsInCat.filter(i => consumption[i.name] > 0);
    if (itemsWithQty.length === 0) return;

    // Category header
    ps.getRange(row, 1, 1, 3).merge()
      .setValue(category)
      .setFontSize(11).setFontWeight("bold").setFontColor("#FFFFFF")
      .setBackground("#1A3C5E").setHorizontalAlignment("left");
    row++;

    // Column sub-header
    ps.getRange(row, 1).setValue("☐").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center");
    ps.getRange(row, 2).setValue("Qty").setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center");
    ps.getRange(row, 3).setValue("Item").setFontSize(10).setFontWeight("bold");
    ps.getRange(row, 1, 1, 3).setBackground("#4A6B8A").setFontColor("#FFFFFF");
    row++;

    // Items
    let altRow = false;
    itemsWithQty.forEach(item => {
      const qty = consumption[item.name];
      ps.getRange(row, 1).setValue("☐").setFontSize(14).setHorizontalAlignment("center");
      ps.getRange(row, 2).setValue(qty).setFontSize(11).setHorizontalAlignment("center").setFontWeight("bold");

      // For books, append the title from Config if available
      let display = item.name;
      if (item.name === "Toddler Book" && bookTitles.Toddler) {
        display += ` — "${bookTitles.Toddler}"`;
      } else if (item.name === "Young Book" && bookTitles.Young) {
        display += ` — "${bookTitles.Young}"`;
      } else if (item.name === "Teen Book" && bookTitles.Teen) {
        display += ` — "${bookTitles.Teen}"`;
      }

      ps.getRange(row, 3).setValue(display).setFontSize(11).setWrap(true);
      if (altRow) ps.getRange(row, 1, 1, 3).setBackground("#F2F5F8");

      // Highlight if this item is short
      if (shortages.find(s => s.item === item.name)) {
        ps.getRange(row, 1, 1, 3).setBackground("#F8D7DA").setFontColor("#721C24");
      }
      row++;
      altRow = !altRow;
    });
    row++;
  });

  // Substitution notes
  if (notes.length > 0) {
    row += 1;
    ps.getRange(row, 1, 1, 3).merge()
      .setValue("SUBSTITUTIONS / NOTES")
      .setFontSize(11).setFontWeight("bold").setFontColor("#FFFFFF")
      .setBackground("#D4943A");
    row++;
    notes.forEach(n => {
      ps.getRange(row, 1, 1, 3).merge()
        .setValue("• " + n).setFontSize(10).setWrap(true)
        .setBackground("#FFF8E1");
      row++;
    });
    row++;
  }

  // Constructed frame component checklist
  const constructedBeds = bedList.filter(b => b.isConstructed);
  if (constructedBeds.length > 0) {
    row += 1;
    ps.getRange(row, 1, 1, 3).merge()
      .setValue("CONSTRUCTED FRAME COMPONENTS (verify complete frame)")
      .setFontSize(11).setFontWeight("bold").setFontColor("#FFFFFF")
      .setBackground("#1A3C5E").setHorizontalAlignment("left");
    row++;

    // Track counters per bed type
    const typeCounts = { Twin: 0, Toddler: 0, Bunk: 0 };
    constructedBeds.forEach(b => {
      typeCounts[b.type]++;
      const num = typeCounts[b.type];
      if (b.type === "Bunk") {
        // Two sleeping surfaces, two component lists
        ["Top", "Bottom"].forEach(lvl => {
          ps.getRange(row, 1, 1, 3).merge()
            .setValue(`Constructed Bunk #${num} — ${lvl}:`)
            .setFontSize(10).setFontWeight("bold").setBackground("#EBF4FF");
          row++;
          V3_CONSTRUCTED_FRAME_PARTS.forEach(part => {
            ps.getRange(row, 1).setValue("☐").setFontSize(12).setHorizontalAlignment("center");
            ps.getRange(row, 3).setValue(part).setFontSize(10);
            row++;
          });
        });
      } else {
        ps.getRange(row, 1, 1, 3).merge()
          .setValue(`Constructed ${b.type} #${num}:`)
          .setFontSize(10).setFontWeight("bold").setBackground("#EBF4FF");
        row++;
        V3_CONSTRUCTED_FRAME_PARTS.forEach(part => {
          ps.getRange(row, 1).setValue("☐").setFontSize(12).setHorizontalAlignment("center");
          ps.getRange(row, 3).setValue(part).setFontSize(10);
          row++;
        });
      }
    });
    row++;
  }

  // Shortage warning
  if (shortages.length > 0) {
    row += 1;
    const warnText = "⚠ WARNING: Insufficient inventory. Short: " +
      shortages.map(s => `${s.shortBy} ${s.item}`).join(", ");
    ps.getRange(row, 1, 1, 3).merge()
      .setValue(warnText)
      .setFontSize(11).setFontWeight("bold").setFontColor("#C62828")
      .setBackground("#F8D7DA").setWrap(true);
    row++;
  }

  // Footer
  row += 2;
  ps.getRange(row, 1, 1, 3).merge()
    .setValue("Generated: " + Utilities.formatDate(new Date(),
              Session.getScriptTimeZone(), "MM/dd/yyyy h:mm a"))
    .setFontSize(9).setFontColor("#999999").setHorizontalAlignment("center");

  // ── Write consumption to Waiting List columns BX–DJ ──
  const consumptionRow = V3_CATALOG.map(item => consumption[item.name] || 0);
  wl.getRange(rowNum, V3_COL_CONSUMPTION_FIRST, 1, V3_ITEM_COUNT)
    .setValues([consumptionRow]);

  // ── Persist the chosen fulfillment config (substitutions/omissions) as JSON ──
  // Stored in the appended column so it travels to Completed Deliveries and can
  // pre-fill the dialog on a re-run.
  wl.getRange(rowNum, V3_COL_FULFILL_CONFIG)
    .setValue(JSON.stringify({ v: 1, choices: choices, notes: notes }));

  // ── Export PDF ──
  SpreadsheetApp.flush();
  const sheetId = ps.getSheetId();
  const ssId = ss.getId();
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
    `format=pdf&gid=${sheetId}&size=letter&portrait=true&fitw=true&` +
    `top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&` +
    `sheetnames=false&printtitle=false&gridlines=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });

  // v3.3: a 200 response can still carry an HTML error page instead of a PDF —
  // verify the content type really is a PDF before saving it as one.
  const blobOk = response.getResponseCode() === 200 &&
    String(response.getBlob().getContentType() || "").toLowerCase().indexOf("pdf") >= 0;

  if (blobOk) {
    const blob = response.getBlob().setName(
      `Pick_List_${apn}_${Utilities.formatDate(new Date(),
       Session.getScriptTimeZone(), "yyyyMMdd_HHmm")}.pdf`
    );

    // Save to Pick Lists subfolder next to workbook
    let targetFolder = null;
    try {
      const wbFile = DriveApp.getFileById(ssId);
      const parents = wbFile.getParents();
      const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
      const sub = parent.getFoldersByName("Pick Lists");
      targetFolder = sub.hasNext() ? sub.next() : parent.createFolder("Pick Lists");
    } catch(e) {
      targetFolder = DriveApp.getRootFolder();
    }

    const file = targetFolder.createFile(blob);
    const fileUrl = file.getUrl();

    try { ss.deleteSheet(ps); } catch(e) { /* ignore */ }

    ss.toast(`✓ Pick list generated for ${apn}`, "Done", 4);
    // Return result to the dialog (we're in a google.script.run context — no ui.*).
    return { ok: true, url: fileUrl, notes: notes, shortages: shortages };
  } else {
    // Leave the "Pick List" tab in place so it can be printed/viewed manually.
    return {
      ok: true, url: null, notes: notes, shortages: shortages,
      pdfError: "Could not generate the PDF (HTTP " + response.getResponseCode() +
        "). The 'Pick List' tab was left in the workbook so you can print or view it manually."
    };
  }
  } catch (err) {
    return { ok: false, error: String((err && err.message) || err) };
  }
}



// ═══════════════════════════════════════════════════════════════════════════
//  SETUP FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
//  SETUP: BOOK TITLES SECTION ON CONFIG TAB
// ═══════════════════════════════════════════════════════════════════════════
//  Adds a Book Titles section (rows 22-26) to the Config tab if not present.
//  Idempotent — safe to run multiple times.

function setupBookTitles() {
  const ui = SpreadsheetApp.getUi();
  const msg = setupBookTitlesCore_(SpreadsheetApp.getActiveSpreadsheet());
  ui.alert(msg);
}

// Silent core \u2014 returns a status string. Used by the menu wrapper above and by
// setupEverything() so the full chain runs without eight interleaved alerts.
function setupBookTitlesCore_(ss) {
  const cfg = ss.getSheetByName(CONFIG_SHEET);
  if (!cfg) return "Config sheet not found.";

  // Check if already set up
  const existing = cfg.getRange("A22").getValue();
  if (String(existing) === "Book Titles") {
    return 'Book Titles section already exists on Config (row 22). Edit titles in cells B24\u2013B26.';
  }

  // Section header (row 22, merged A22:C22)
  cfg.getRange("A22:C22").merge();
  cfg.getRange("A22").setValue("Book Titles")
    .setFontWeight("bold").setFontSize(11).setFontColor("#FFFFFF")
    .setBackground("#1A3C5E").setHorizontalAlignment("left")
    .setFontFamily("Arial");

  // Column headers (row 23)
  cfg.getRange("A23").setValue("Book Type")
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
    .setHorizontalAlignment("center").setFontFamily("Arial").setFontSize(11);
  cfg.getRange("B23").setValue("Title")
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
    .setHorizontalAlignment("center").setFontFamily("Arial").setFontSize(11);

  // Data rows (24-26)
  const bookTypes = ["Toddler Books", "Young Books", "Older Books"];
  bookTypes.forEach((bt, i) => {
    const row = 24 + i;
    cfg.getRange(row, 1).setValue(bt).setFontWeight("bold").setFontFamily("Arial").setFontSize(10);
    // Leave B column blank for user to fill in
    cfg.getRange(row, 2).setFontFamily("Arial").setFontSize(10)
       .setBackground("#FFF8E1")  // light amber to signal "editable"
       .setBorder(true, true, true, true, false, false, "#C8D1DA", SpreadsheetApp.BorderStyle.SOLID);
    if (i % 2 === 1) {
      cfg.getRange(row, 1, 1, 2).setBackground("#F2F5F8");
      cfg.getRange(row, 2).setBackground("#FFF8E1"); // keep B amber
    }
  });

  // Instruction note (row 27)
  cfg.getRange("A27:F27").merge();
  cfg.getRange("A27").setValue("Edit cells B24\u2013B26 to set the current title for each book type. Leave blank if no title to display on pick lists.")
    .setFontSize(9).setFontStyle("italic").setFontColor("#777777").setWrap(true)
    .setFontFamily("Arial");
  cfg.setRowHeight(27, 28);

  return '\u2713 Book Titles section added to Config (rows 22\u201326). Edit titles in cells B24\u2013B26. Then re-run \"Set up formula protections\" to unlock those cells for editing.';
}


// ═══════════════════════════════════════════════════════════════════════════
//  v3.0 SETUP: FULL INVENTORY STRUCTURE RESTRUCTURE
//  One-shot setup that expands the workbook from 14 → 39 components and
//  adds all v3.0 formula columns to the Waiting List.
//
//  Stages:
//    1. Config tab      — replace components list, clear tracking start date
//    2. Incoming Items  — rebuild header row with 39 quantity columns
//    3. Inventory Position — prep header row for 39-component snapshot grid
//    4. Waiting List    — add v3.0 formula columns (BI–BW) and headers
//    5. Completed Deliveries / Cancelled Requests — add matching columns
//
//  After this runs, the user should run:
//    - Set up reconciliation logic (rebuilds snapshot formulas over 39 rows)
//    - Set up data validation    (refreshes all dropdowns/rules)
//    - Set up formula protections
//
//  Idempotent — safe to re-run.
// ═══════════════════════════════════════════════════════════════════════════

function setupV3Structure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    "Set up v3.0 inventory structure",
    "This will restructure the workbook to support the 39-item inventory " +
    "with gendered/age-tiered comforters, mattress size variants, and " +
    "constructed/purchased frame variants.\n\n" +
    "Existing data is preserved. New columns and rows are added.\n" +
    "The Tracking Start Date will be CLEARED so all On Hand counts read 0 " +
    "until you enter fresh physical counts in the Reconciliation Log.\n\n" +
    "Continue?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const log = [];
  try {
    log.push(v3Stage1_Config(ss));
    log.push(v3Stage2_IncomingItems(ss));
    log.push(v3Stage3_InventoryPositionHeader(ss));
    log.push(v3Stage4_WaitingList(ss));
    log.push(v3Stage5_ArchiveTabs(ss));
  } catch (e) {
    ui.alert(
      "Setup encountered an error",
      "Error: " + e.message + "\n\n" +
      "Stages completed:\n" + log.join("\n") + "\n\n" +
      "Re-run this setup to continue. Each stage is idempotent.",
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert(
    "v3.0 structure setup complete",
    "Stages completed:\n" + log.join("\n") + "\n\n" +
    "Next steps (run these from the Setup submenu):\n" +
    "  1. Set up reconciliation logic\n" +
    "  2. Set up data validation\n" +
    "  3. Set up formula protections\n\n" +
    "Then begin the physical recount by entering counts in the Reconciliation Log.",
    ui.ButtonSet.OK
  );
}


// ═══════════════════════════════════════════════════════════════════════════
//  v3.0 ONE-TIME REPAIR FUNCTION
//
//  Fixes the Inventory Position tab if rows 20-34 have blank component names
//  or missing formulas. Caused by a bug in an earlier v3.0 setup that cleared
//  rows 20-34 after writing component names to them. Safe to run multiple
//  times. Does NOT touch the portfolio shortage section (rows 45+).
// ═══════════════════════════════════════════════════════════════════════════

function v3RepairInventoryNames() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) { ui.alert('Inventory Position sheet not found.'); return; }

  // Step 1: Write all 39 component names to col A rows 4-42
  const names = V3_CATALOG.map(c => [c.name]);
  inv.getRange(V3_INV_FIRST_ROW, 1, names.length, 1).setValues(names)
    .setFontFamily("Arial").setFontSize(10).setFontWeight("bold");

  // Step 2: Rewrite per-component formulas for any rows missing them
  const cd       = "'Completed Deliveries'";
  const wl       = "'Waiting List'";
  const cfgDate  = "Config!$B$19";
  const rl       = "'Reconciliation Log'";
  const pkMaskWL = `(${wl}!$C$2:$C$9999="Packed")`;

  let rowsRepaired = 0;
  V3_CATALOG.forEach((item, idx) => {
    const r = V3_INV_FIRST_ROW + idx;
    const currentB = inv.getRange(r, 2).getFormula();
    if (currentB && currentB.startsWith("=")) return;  // Row B already OK, skip

    const effDate    = `IF(B${r}="",${cfgDate},B${r})`;
    const incColLtr  = getColumnA1_(V3_INC_FIRST_COL + idx);
    const usedColLtr = getColumnA1_(V3_COL_CONSUMPTION_FIRST + idx);

    inv.getRange(r, 2).setFormula(
      `=IFERROR(IF(MAXIFS(${rl}!A$2:A$9999,${rl}!B$2:B$9999,A${r},${rl}!D$2:D$9999,"<>")=0,"",` +
      `MAXIFS(${rl}!A$2:A$9999,${rl}!B$2:B$9999,A${r},${rl}!D$2:D$9999,"<>")),"")`
    );
    inv.getRange(r, 2).setNumberFormat("MM/DD/YYYY");

    inv.getRange(r, 3).setFormula(
      `=IF(B${r}="",0,SUMPRODUCT(` +
      `(${rl}!B$2:B$9999=A${r})*(${rl}!A$2:A$9999=B${r})*` +
      `(${rl}!D$2:D$9999<>"")*${rl}!D$2:D$9999))`
    );

    inv.getRange(r, 4).setFormula(
      `=IF(${cfgDate}="",0,SUMIFS('Incoming Items'!${incColLtr}$2:${incColLtr}$9999,` +
      `'Incoming Items'!$A$2:$A$9999,">="&${effDate}))`
    );

    inv.getRange(r, 5).setFormula(
      `=IF(${cfgDate}="",0,` +
      `SUMIFS(${cd}!${usedColLtr}$2:${usedColLtr}$9999,` +
        `${cd}!$AY$2:$AY$9999,">="&${effDate})+` +
      `SUMPRODUCT(${pkMaskWL}*${wl}!${usedColLtr}$2:${usedColLtr}$9999))`
    );

    inv.getRange(r, 6).setFormula(`=IF(${cfgDate}="",0,C${r}+D${r}-E${r})`);
    inv.getRange(r, 7).setFormula(
      `=IF(F${r}<=0,"\u26A0 OUT",IF(F${r}<=3,"\u26A0 LOW","\u2713 OK"))`
    );

    rowsRepaired++;
  });

  ui.alert(
    "\u2713 Inventory Position repair complete.\n\n" +
    "• Wrote all 39 component names to column A (rows 4\u201342).\n" +
    "• Rewrote snapshot formulas for " + rowsRepaired + " row(s) that were missing them.\n\n" +
    "No changes to the portfolio shortage section (rows 45\u201385) or any other tab."
  );
}


// ═══════════════════════════════════════════════════════════════════════════
//  v3.0 REPAIR: WAITING LIST FORMULAS
//
//  Retroactively writes all v3.0 formulas (legacy D/O-W + BF/BG + BJ-BW)
//  to any existing Active or Packed rows that are missing them. Use after
//  patching the script to fix rows that came in from form submissions
//  before the fix was deployed.
//
//  Safe to re-run. Only touches rows where status is Active or Packed;
//  skips Completed, Cancelled, and TOTALS rows.
// ═══════════════════════════════════════════════════════════════════════════

function v3RepairWaitingListFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const wl = ss.getSheetByName(ACTIVE_SHEET);
  if (!wl) { ui.alert('Waiting List sheet not found.'); return; }

  const lastRow = wl.getLastRow();
  let repaired = 0;
  let skipped = 0;

  for (let r = 2; r <= lastRow; r++) {
    const b = String(wl.getRange(r, 2).getValue());
    const c = String(wl.getRange(r, 3).getValue());

    // Skip TOTALS row and non-Active/Packed rows
    if (b === TOTALS_ROW_LABEL) continue;
    if (c !== "Active" && c !== "Packed") { skipped++; continue; }

    // Rewrite all formulas for this row
    writeRowFormulas_(wl, r);
    repaired++;
  }

  ui.alert(
    "\u2713 Waiting List formula repair complete.\n\n" +
    "• Rewrote formulas for " + repaired + " Active/Packed row(s).\n" +
    "• Skipped " + skipped + " row(s) (Completed/Cancelled/blank).\n\n" +
    "The updated BF (Shortage Impact) and BG (Constrained Items) formulas now " +
    "reference the v3.0 shortage section at rows 47\u201385, and all v3.0 demand " +
    "columns (BJ\u2013BW) are populated."
  );
}

function v3Stage1_Config(ss) {
  const cfg = ss.getSheetByName(CONFIG_SHEET);
  if (!cfg) throw new Error("Config sheet not found");

  // Ensure A1 says "Component"
  cfg.getRange("A1").setValue("Component")
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
    .setFontFamily("Arial").setFontSize(11);

  // Clear A2:A50 first to remove old component list
  cfg.getRange("A2:A50").clearContent();

  // Write new 39-item list at A2:A40
  const names = V3_CATALOG.map(c => [c.name]);
  cfg.getRange(2, 1, names.length, 1).setValues(names)
    .setFontFamily("Arial").setFontSize(10);

  // Clear the Tracking Start Date so all On Hand = 0 until recount
  cfg.getRange("B19").clearContent();

  return "  ✓ Stage 1: Config components list (39 items); Tracking Start Date cleared";
}

function v3Stage2_IncomingItems(ss) {
  const inc = ss.getSheetByName(INCOMING_SHEET);
  if (!inc) throw new Error("Incoming Items sheet not found");

  // Desired header: A=Date, B=Source Type, C=Source Detail, D..AP = 39 items, AQ = Notes
  const headers = ["Date Received", "Source Type", "Source Detail"]
    .concat(V3_CATALOG.map(c => c.name))
    .concat(["Notes"]);

  // Write headers
  inc.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
    .setFontFamily("Arial").setFontSize(10).setHorizontalAlignment("center")
    .setWrap(true);

  inc.setRowHeight(1, 48);

  // Set reasonable column widths
  inc.setColumnWidth(1, 110);   // Date
  inc.setColumnWidth(2, 110);   // Source Type
  inc.setColumnWidth(3, 150);   // Source Detail
  for (let i = V3_INC_FIRST_COL; i <= V3_INC_LAST_COL; i++) {
    inc.setColumnWidth(i, 100);
  }
  inc.setColumnWidth(V3_INC_NOTES_COL, 250);

  inc.setFrozenRows(1);

  return "  ✓ Stage 2: Incoming Items (" + headers.length + " columns: 3 fixed + 39 items + Notes)";
}

function v3Stage3_InventoryPositionHeader(ss) {
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) throw new Error("Inventory Position sheet not found");

  // Write the component names at col A, rows 4..42 (39 items)
  const names = V3_CATALOG.map(c => [c.name]);
  inv.getRange(V3_INV_FIRST_ROW, 1, names.length, 1).setValues(names)
    .setFontFamily("Arial").setFontSize(10).setFontWeight("bold");

  // Clear any leftover component rows below row 42 from the old 14-item layout
  // (safe — v3.0 only uses rows 4-42 for components; rows 43+ get rebuilt by
  // setupReconciliationLogic into the shortage section starting at row 45).
  for (let r = V3_INV_LAST_ROW + 1; r <= V3_INV_LAST_ROW + 2; r++) {
    inv.getRange(r, 1, 1, 7).clearContent();
  }

  // Note: earlier versions had a loop here that cleared rows 20-34 to wipe
  // the old v2.4 shortage section. That was a BUG — in v3.0 those rows hold
  // newly-written component names (items 17-31), and the clear wiped them.
  // setupReconciliationLogic now handles clearing rows 45+ independently.

  return "  \u2713 Stage 3: Inventory Position component names (" + V3_ITEM_COUNT + " rows)";
}

function v3Stage4_WaitingList(ss) {
  const wl = ss.getSheetByName(ACTIVE_SHEET);
  if (!wl) throw new Error("Waiting List sheet not found");

  // Part A: Write new demand column headers for BI–BW
  const newHeaders = [["Pack Config"]];
  V3_COMF_DEMAND_LABELS.forEach(({gender, age}) => {
    newHeaders[0].push("Comforter — " + gender + " " + age);
  });
  newHeaders[0].push("Sheets Demand (max)");
  newHeaders[0].push("Mattress Protectors Demand");
  newHeaders[0].push("Bedrails Max");
  newHeaders[0].push("Plaques Max");
  newHeaders[0].push("Bears Demand");

  wl.getRange(1, V3_COL_PACK_CONFIG, 1, newHeaders[0].length).setValues(newHeaders)
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
    .setFontFamily("Arial").setFontSize(10).setHorizontalAlignment("center")
    .setWrap(true);

  // Part B: Write per-item consumption column headers for BX–CJ (39 cols)
  const consumptionHeaders = V3_CATALOG.map(c => "Used: " + c.name);
  wl.getRange(1, V3_COL_CONSUMPTION_FIRST, 1, consumptionHeaders.length)
    .setValues([consumptionHeaders])
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#4A6B8A")
    .setFontFamily("Arial").setFontSize(9).setHorizontalAlignment("center")
    .setWrap(true);

  // Set column widths
  wl.setColumnWidth(V3_COL_PACK_CONFIG, 160);
  for (let c = V3_COL_COMF_FIRST; c <= V3_COL_COMF_LAST; c++) {
    wl.setColumnWidth(c, 90);
  }
  wl.setColumnWidth(V3_COL_SHEETS_DEMAND, 90);
  wl.setColumnWidth(V3_COL_MATT_PROT_DEMAND, 100);
  wl.setColumnWidth(V3_COL_BEDRAILS_MAX, 90);
  wl.setColumnWidth(V3_COL_PLAQUES_MAX, 90);
  wl.setColumnWidth(V3_COL_BEARS_DEMAND, 90);
  for (let c = V3_COL_CONSUMPTION_FIRST; c <= V3_COL_CONSUMPTION_LAST; c++) {
    wl.setColumnWidth(c, 80);
  }

  // Part C: Write formulas for all existing data rows (rows 2 to last row, excluding TOTALS)
  const lastRow = wl.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    const colB = String(wl.getRange(r, 2).getValue());
    if (colB === TOTALS_ROW_LABEL) continue;
    v3WriteWaitingListFormulas(wl, r);
  }

  return "  ✓ Stage 4: Waiting List extended to column " + getColumnA1_(V3_COL_CONSUMPTION_LAST) +
         " (" + newHeaders[0].length + " demand cols + " + V3_ITEM_COUNT + " consumption cols)";
}

function v3WriteWaitingListFormulas(wl, r) {
  // Children ages and genders at cols X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM
  // Age: X,Z,AB,AD,AF,AH,AJ,AL  (cols 24,26,28,30,32,34,36,38)
  // Gender: Y,AA,AC,AE,AG,AI,AK,AM (cols 25,27,29,31,33,35,37,39)
  const ageCols    = ["X", "Z", "AB", "AD", "AF", "AH", "AJ", "AL"];
  const genderCols = ["Y", "AA", "AC", "AE", "AG", "AI", "AK", "AM"];

  // ── 9 Comforter demand formulas (BJ–BR) ──
  // For each (gender, age) combo, count children matching both.
  // Gender is Boys (M), Girls (F), or Neutral (blank/unknown — won't happen per req).
  //   Ages: Toddler 0-3, Young 4-10, Teen 11-18.
  const ageRanges = {
    "Toddler": { min: 0, max: 3 },
    "Young":   { min: 4, max: 10 },
    "Teen":    { min: 11, max: 18 },
  };
  const genderLetters = { "Boys": "M", "Girls": "F" };

  V3_COMF_DEMAND_LABELS.forEach(({gender, age}, idx) => {
    const col = V3_COL_COMF_FIRST + idx;
    const range = ageRanges[age];

    let terms = [];
    for (let i = 0; i < ageCols.length; i++) {
      const ac = ageCols[i] + r;
      const gc = genderCols[i] + r;
      if (gender === "Neutral") {
        // Neutral demand = 0 under current data model (gender is required)
        // Reserved for future use if neutral children are ever entered
        terms.push(`IF(AND(${ac}<>"",${gc}="",${ac}>=${range.min},${ac}<=${range.max}),1,0)`);
      } else {
        const letter = genderLetters[gender];
        terms.push(`IF(AND(${gc}="${letter}",${ac}>=${range.min},${ac}<=${range.max}),1,0)`);
      }
    }
    const formula = "=" + terms.join("+");
    wl.getRange(r, col).setFormula(formula);
  });

  // ── Sheets Demand (BS) — WORST-CASE max (v3.3) ──
  // True sheet need is allocation-dependent: the picker prefers with-sheets
  // comforters and only adds a loose Twin Sheet when a without-sheets comforter
  // is used, so actual sheets are usually far fewer than 1 per child. We keep a
  // worst-case max (= 1 per child = sum of comforter demand) here because
  // over-forecasting sheets is the safe direction. The header is labelled
  // "Sheets Demand (max)" so it's not mistaken for exact need.
  const comfSumRange = "BJ" + r + ":BR" + r;
  wl.getRange(r, V3_COL_SHEETS_DEMAND).setFormula(`=SUM(${comfSumRange})`);

  // ── Mattress Protectors Demand (BT) — 1 per non-crib sleeping surface (v3.3) ──
  // Matches the pick-list allocator exactly: twins + toddlers + bunks×2; cribs
  // excluded (they don't take a twin mattress protector). Previously this was
  // (wrongly) the child count, which over/under-counted vs. what's actually used.
  wl.getRange(r, V3_COL_MATT_PROT_DEMAND).setFormula(`=G${r}+F${r}+H${r}*2`);

  // v3.6: Ministry Options gate the extras' demand forecasts (blank toggles = on,
  // so pre-v3.6 workbooks are unchanged).
  const opts = getMinistryOptions_();

  // ── Bedrails Max (BU) ──
  // Max possible = toddler beds × 2 + bunk beds × 2 (if all constructed)
  // Actual depends on constructed/purchased choice made at pack time
  wl.getRange(r, V3_COL_BEDRAILS_MAX).setFormula(opts.bedrails ? `=(F${r}+H${r})*2` : "=0");

  // ── Plaques Max (BV) ──
  // Max possible = twin + toddler + bunk*2 (assuming all constructed)
  wl.getRange(r, V3_COL_PLAQUES_MAX).setFormula(opts.plaques ? `=G${r}+F${r}+H${r}*2` : "=0");

  // ── Bears Demand (BW) — children ages 0-10 ──
  let bearTerms = [];
  for (let i = 0; i < ageCols.length; i++) {
    const ac = ageCols[i] + r;
    bearTerms.push(`IF(AND(${ac}<>"",${ac}>=0,${ac}<=${V3_BEAR_AGE_MAX}),1,0)`);
  }
  wl.getRange(r, V3_COL_BEARS_DEMAND).setFormula(opts.bears ? "=" + bearTerms.join("+") : "=0");
}

function v3Stage5_ArchiveTabs(ss) {
  // Add matching columns to Completed Deliveries and Cancelled Requests
  // so archived rows preserve their v3.0 demand and consumption data.
  const demandHeaders = ["Pack Config"];
  V3_COMF_DEMAND_LABELS.forEach(({gender, age}) => {
    demandHeaders.push("Comforter — " + gender + " " + age);
  });
  demandHeaders.push("Sheets Demand (max)");
  demandHeaders.push("Mattress Protectors Demand");
  demandHeaders.push("Bedrails Max");
  demandHeaders.push("Plaques Max");
  demandHeaders.push("Bears Demand");

  const consumptionHeaders = V3_CATALOG.map(c => "Used: " + c.name);

  let updated = 0;
  [ARCHIVE_SHEET, CANCELLED_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    sh.getRange(1, V3_COL_PACK_CONFIG, 1, demandHeaders.length).setValues([demandHeaders])
      .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
      .setFontFamily("Arial").setFontSize(10).setHorizontalAlignment("center")
      .setWrap(true);

    sh.getRange(1, V3_COL_CONSUMPTION_FIRST, 1, consumptionHeaders.length)
      .setValues([consumptionHeaders])
      .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#4A6B8A")
      .setFontFamily("Arial").setFontSize(9).setHorizontalAlignment("center")
      .setWrap(true);

    updated++;
  });

  return "  ✓ Stage 5: Archive tabs (" + updated + " sheets) now have matching v3.0 columns";
}




// ═══════════════════════════════════════════════════════════════════════════
//  v3.0 SETUP: RECONCILIATION LOGIC
//  Rebuilds the Inventory Position snapshot grid for all 39 components.
//  Because v3.0 uses 1-to-1 columns (one Incoming Items col per item, one
//  consumption col per item on Waiting List/Completed Deliveries), the
//  formulas are much simpler than v2.4: no per-item expression maps needed.
// ═══════════════════════════════════════════════════════════════════════════

function setupReconciliationLogic() {
  const ui = SpreadsheetApp.getUi();
  try {
    setupReconciliationLogicCore_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {
    ui.alert(e.message);
    return;
  }
  ui.alert(
    '✓ Reconciliation logic set up for v3.0.\n\n' +
    '• 39 components in Inventory Position snapshot grid (rows 4–42)\n' +
    '• Portfolio shortage section rebuilt at rows 45–85\n' +
    '• Reconciliation Log Calculated Count formulas updated\n\n' +
    'Next: run "Set up formula protections".'
  );
}

// Silent core — throws on missing sheets. Shared by the menu wrapper and
// setupEverything().
function setupReconciliationLogicCore_(ss) {
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const rec = ss.getSheetByName(RECON_SHEET);

  if (!inv) throw new Error('Inventory Position sheet not found.');
  if (!rec) throw new Error('Reconciliation Log sheet not found.');

  const cd       = "'Completed Deliveries'";
  const wl       = "'Waiting List'";
  const cfgDate  = "Config!$B$19";
  const rl       = "'Reconciliation Log'";
  const pkMaskWL = `(${wl}!$C$2:$C$9999="Packed")`;

  // ── 1. Rewrite Inventory Position header row ──────────────────────────────
  const headers = ["Component", "Last Recon Date", "Last Physical Count",
                   "In Since Recon", "Out Since Recon", "On Hand", "Status"];
  headers.forEach((h, i) => {
    inv.getRange(3, i + 1).setValue(h)
      .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
      .setFontFamily("Arial").setFontSize(11).setHorizontalAlignment("center");
  });

  // ── 2. Per-component snapshot formulas for all 39 items ───────────────────
  // In v3.0, item i uses:
  //   - Incoming Items column  (V3_INC_FIRST_COL + i)
  //   - Waiting List / CD consumption column (V3_COL_CONSUMPTION_FIRST + i)
  // This means the formulas are identical per-component except for the column
  // letters — no complex expression mapping like v2.4 required.

  V3_CATALOG.forEach((item, idx) => {
    const r          = V3_INV_FIRST_ROW + idx;
    const effDate    = `IF(B${r}="",${cfgDate},B${r})`;
    const incColLtr  = getColumnA1_(V3_INC_FIRST_COL + idx);
    const usedColLtr = getColumnA1_(V3_COL_CONSUMPTION_FIRST + idx);

    // A — component name (already written by Stage 3)

    // B — Last Recon Date (MAXIFS with D<>"" so partial rows don't trigger)
    inv.getRange(r, 2).setFormula(
      `=IFERROR(IF(MAXIFS(${rl}!A$2:A$9999,${rl}!B$2:B$9999,A${r},${rl}!D$2:D$9999,"<>")=0,"",` +
      `MAXIFS(${rl}!A$2:A$9999,${rl}!B$2:B$9999,A${r},${rl}!D$2:D$9999,"<>")),"")`
    );
    inv.getRange(r, 2).setNumberFormat("MM/DD/YYYY");

    // C — Last Physical Count
    inv.getRange(r, 3).setFormula(
      `=IF(B${r}="",0,SUMPRODUCT(` +
      `(${rl}!B$2:B$9999=A${r})*(${rl}!A$2:A$9999=B${r})*` +
      `(${rl}!D$2:D$9999<>"")*${rl}!D$2:D$9999))`
    );

    // D — In Since Recon (from Incoming Items)
    inv.getRange(r, 4).setFormula(
      `=IF(${cfgDate}="",0,SUMIFS('Incoming Items'!${incColLtr}$2:${incColLtr}$9999,` +
      `'Incoming Items'!$A$2:$A$9999,">="&${effDate}))`
    );

    // E — Out Since Recon (Completed Deliveries consumption after effDate +
    //                      currently Packed reservations on Waiting List)
    inv.getRange(r, 5).setFormula(
      `=IF(${cfgDate}="",0,` +
      `SUMIFS(${cd}!${usedColLtr}$2:${usedColLtr}$9999,` +
        `${cd}!$AY$2:$AY$9999,">="&${effDate})+` +
      `SUMPRODUCT(${pkMaskWL}*${wl}!${usedColLtr}$2:${usedColLtr}$9999))`
    );

    // F — On Hand
    inv.getRange(r, 6).setFormula(`=IF(${cfgDate}="",0,C${r}+D${r}-E${r})`);

    // G — Status
    inv.getRange(r, 7).setFormula(
      `=IF(F${r}<=0,"\u26A0 OUT",IF(F${r}<=3,"\u26A0 LOW","\u2713 OK"))`
    );
  });

  // ── 3. Portfolio shortage section (rebuilt at rows 45+) ───────────────────
  // One row per item where demand formula is meaningful. Demand is read from
  // the Waiting List (Active + Packed rows) using the corresponding demand
  // column (BI–BW) or consumption column fallback for items without a
  // dedicated demand col.

  // Clear any existing content in rows 45-120 first (also break apart any
  // merges that might exist from a prior deployment)
  inv.getRange(45, 1, 76, 7).breakApart().clearContent();

  // Section header
  inv.getRange(45, 1, 1, 7).merge()
    .setValue("Portfolio-Level Shortage Summary")
    .setFontWeight("bold").setFontSize(12).setFontColor("#FFFFFF")
    .setBackground("#1A3C5E").setHorizontalAlignment("center");

  // Column headers row 46
  const shHeaders = ["Component", "Demand", "On Hand", "Shortage", "Status"];
  shHeaders.forEach((h, i) => {
    inv.getRange(46, i + 1).setValue(h)
      .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#4A6B8A")
      .setFontFamily("Arial").setFontSize(11).setHorizontalAlignment("center");
  });

  // For each of 39 items, write a shortage row starting at row 47.
  // Demand formula comes from Waiting List:
  //   - For the 9 comforter variants (Boys/Girls/Neutral × Toddler/Young/Teen)
  //     without regard to sheets-state: demand is sum of BOTH with-sheets and
  //     without-sheets usage — but since the picker prefers with-sheets, we
  //     sum the aggregate comforter demand and display it on BOTH rows.
  //   - For items with a dedicated max-demand column (Bedrails, Plaques, etc),
  //     use that column directly.
  //   - For items driven by bed counts (frames, mattresses) we use consumption
  //     column for actual use when packed, or max possible demand based on
  //     bed-type columns for Active requests.

  // Helper: for each catalog item, produce the demand formula.
  // v3.4 — one consistent rule for EVERY item so substitutions/omissions stay truthful:
  //   demand = SUMPRODUCT((Status="Active") * forecastExpr)        ← estimate for not-yet-packed
  //          + SUMPRODUCT((Status="Packed") * actualConsumption)   ← exactly what was allocated
  // This removes the phantom 6"/understated 8" demand after a cross-bucket mattress swap and
  // the overstated bear/comforter demand after an omission (a Packed row now contributes its
  // ACTUAL consumption, not a bed-count/age estimate). On Hand was already correct.
  const ageA1 = CHILD_AGE_COLS_A1; // ['X','Z','AB','AD','AF','AH','AJ','AL']
  const bookTierExpr = (perCol) =>
    "(" + ageA1.map(c => perCol(`${wl}!${c}$2:${c}$9999`)).join("+") + ">0)";

  const demandFormula = (item, idx) => {
    const name = item.name;
    const activeMask = `(${wl}!$C$2:$C$9999="Active")`;
    const packedMask = `(${wl}!$C$2:$C$9999="Packed")`;
    const consCol = (() => {
      const L = getColumnA1_(V3_COL_CONSUMPTION_FIRST + idx);
      return `${wl}!${L}$2:${L}$9999`;
    })();

    // forecastExpr = estimated demand for ACTIVE rows. Purchased frames and any
    // unmatched item stay "0" (they only count once actually packed → consumption).
    let forecastExpr = "0";

    const comfMatch = name.match(/^Comforter.*-\s+(Boys|Girls|Neutral)\s+(Toddler|Young|Teen)$/);
    if (comfMatch) {
      const demandIdx = V3_COMF_DEMAND_LABELS.findIndex(
        l => l.gender === comfMatch[1] && l.age === comfMatch[2]
      );
      if (demandIdx >= 0) {
        const demandCol = getColumnA1_(V3_COL_COMF_FIRST + demandIdx);
        forecastExpr = `${wl}!${demandCol}$2:${demandCol}$9999`;
      }
    } else if (name === "Twin Sheets")              { forecastExpr = `${wl}!BS$2:BS$9999`; }
    else if (name === "Twin Mattress Protector")    { forecastExpr = `${wl}!BT$2:BT$9999`; }
    else if (name === "Bedrail")                    { forecastExpr = `${wl}!BU$2:BU$9999`; }
    else if (name === "Plaque")                     { forecastExpr = `${wl}!BV$2:BV$9999`; }
    else if (name === "Bear")                       { forecastExpr = `${wl}!BW$2:BW$9999`; }
    else if (name === "Twin Frame - Constructed")   { forecastExpr = `(${wl}!G$2:G$9999+${wl}!H$2:H$9999*2)`; }
    else if (name === "Toddler Frame - Constructed"){ forecastExpr = `${wl}!F$2:F$9999`; }
    else if (name === "Bunk Frame - Constructed")   { forecastExpr = `${wl}!H$2:H$9999`; }
    else if (name === "Crib")                       { forecastExpr = `${wl}!E$2:E$9999`; }
    else if (name === "Crib Mattress")              { forecastExpr = `${wl}!E$2:E$9999`; }
    else if (name === "Pillow")                     { forecastExpr = `(${wl}!G$2:G$9999+${wl}!F$2:F$9999+${wl}!E$2:E$9999+${wl}!H$2:H$9999*2)`; }
    else if (name === "Mattress 6in 38x74" || name === "Mattress 6in 39x75") { forecastExpr = `(${wl}!H$2:H$9999+${wl}!F$2:F$9999)`; }
    else if (name === "Mattress 8in 38x74" || name === "Mattress 8in 39x75") { forecastExpr = `(${wl}!G$2:G$9999+${wl}!H$2:H$9999)`; }
    else if (name === "Toddler Book") { forecastExpr = bookTierExpr(c => `(${c}<=3)*(${c}<>"")`); }
    else if (name === "Young Book")   { forecastExpr = bookTierExpr(c => `(${c}>=4)*(${c}<=10)`); }
    else if (name === "Teen Book")    { forecastExpr = bookTierExpr(c => `(${c}>=11)*(${c}<=18)`); }

    return `=SUMPRODUCT(${activeMask}*(${forecastExpr}))+SUMPRODUCT(${packedMask}*${consCol})`;
  };

  // Write shortage rows 47..85 (39 items)
  V3_CATALOG.forEach((item, idx) => {
    const sRow = 47 + idx;
    const invRow = V3_INV_FIRST_ROW + idx;
    inv.getRange(sRow, 1).setValue(item.name).setFontFamily("Arial").setFontSize(10);
    inv.getRange(sRow, 2).setFormula(demandFormula(item, idx));
    inv.getRange(sRow, 3).setFormula(`=F${invRow}`);
    inv.getRange(sRow, 4).setFormula(`=MAX(0,B${sRow}-C${sRow})`);
    inv.getRange(sRow, 5).setFormula(`=IF(D${sRow}>0,"\u26A0 SHORT","\u2713 OK")`);

    // Alternating row shading
    if (idx % 2 === 1) {
      inv.getRange(sRow, 1, 1, 5).setBackground("#F2F5F8");
    }
  });

  // ── 4. Reconciliation Log Calculated Count (col C) ────────────────────────
  // In v3.0 the formula is much simpler because each component has its own
  // dedicated consumption column on CD and Waiting List.
  //
  // Logic per row r:
  //   prevDate  = most recent prior recon for this component (strictly before A_r)
  //   prevCount = physical count on prevDate (0 if no prior)
  //   baseDate  = prevDate, or Config tracking start date if no prior recon
  //   endDate   = A_r (count date), or today if A_r is blank
  //   itemIdx   = 1-39 mapping to Config!A2:A40
  //   incColLtr = Incoming Items column letter (D + itemIdx - 1)
  //   usedColLtr = Consumption column letter (BX + itemIdx - 1)
  //   inSince   = Incoming Items for this component from baseDate to endDate
  //   outCd     = Completed Deliveries consumption from baseDate to endDate
  //   outWl     = Currently Packed consumption
  //   Result    = prevCount + inSince - outCd - outWl
  //
  // Column letter math:
  //   incColLtr: column V3_INC_FIRST_COL (4=D) through V3_INC_LAST_COL (42=AP)
  //   usedColLtr: column V3_COL_CONSUMPTION_FIRST (76=BX) through 114=CJ
  // We compute letters at formula time using CHAR() for single-letter cols and
  // ADDRESS() for multi-letter. ADDRESS(1,col,4) returns the column letter string.

  for (let r = 2; r <= 101; r++) {
    const formula =
      `=IF(OR(B${r}="",${cfgDate}=""),"",` +
      `IFERROR(LET(` +
        `prevDate,IFERROR(MAXIFS(A$2:A$9999,B$2:B$9999,B${r},` +
          `A$2:A$9999,"<"&IF(A${r}="",TODAY()+1,A${r}),D$2:D$9999,"<>"),0),` +
        `prevCount,IF(prevDate=0,0,SUMPRODUCT(` +
          `(B$2:B$9999=B${r})*(A$2:A$9999=prevDate)*(D$2:D$9999<>"")*D$2:D$9999)),` +
        `baseDate,IF(prevDate=0,${cfgDate},prevDate),` +
        `endDate,IF(A${r}="",TODAY(),A${r}),` +
        // itemIdx: 1-39 matching Config!A2:A40
        `itemIdx,MATCH(B${r},Config!$A$2:$A$40,0),` +
        // incColNum: absolute col number for Incoming Items (V3_INC_FIRST_COL + itemIdx - 1)
        `incColNum,${V3_INC_FIRST_COL}+itemIdx-1,` +
        // usedColNum: absolute col number for consumption col
        `usedColNum,${V3_COL_CONSUMPTION_FIRST}+itemIdx-1,` +
        // incColLtr / usedColLtr via SUBSTITUTE + ADDRESS
        `incColLtr,SUBSTITUTE(ADDRESS(1,incColNum,4),"1",""),` +
        `usedColLtr,SUBSTITUTE(ADDRESS(1,usedColNum,4),"1",""),` +
        // inSince
        `inSince,IFERROR(SUMIFS(` +
          `INDIRECT("'Incoming Items'!"&incColLtr&"2:"&incColLtr&"9999"),` +
          `'Incoming Items'!A$2:A$9999,">="&baseDate,` +
          `'Incoming Items'!A$2:A$9999,"<="&endDate),0),` +
        // outCd: deliveries from baseDate to endDate, this component
        `outCd,IFERROR(SUMIFS(` +
          `INDIRECT("${cd}!"&usedColLtr&"2:"&usedColLtr&"9999"),` +
          `${cd}!$AY$2:$AY$9999,">="&baseDate,` +
          `${cd}!$AY$2:$AY$9999,"<="&endDate),0),` +
        // outWl: currently Packed reservations
        `outWl,IFERROR(SUMPRODUCT(` +
          `(${wl}!$C$2:$C$9999="Packed")*` +
          `INDIRECT("${wl}!"&usedColLtr&"2:"&usedColLtr&"9999")),0),` +
        // Result
        `prevCount+inSince-outCd-outWl` +
      `),0))`;
    rec.getRange(r, 3).setFormula(formula);

    // E — Adjustment: Physical Count minus Calculated Count.
    // Blank when either value is missing so incomplete rows don't show noise.
    const adjFormula =
      `=IF(OR(D${r}="",C${r}=""),"",D${r}-C${r})`;
    rec.getRange(r, 5).setFormula(adjFormula);
  }

  // ── 5. Conditional formatting ─────────────────────────────────────────────
  inv.clearConditionalFormatRules();
  const rules = [];

  const onHandRange = inv.getRange(`F${V3_INV_FIRST_ROW}:F${V3_INV_LAST_ROW}`);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(0)
    .setBackground("#F8D7DA").setFontColor("#721C24").setBold(true)
    .setRanges([onHandRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 3)
    .setBackground("#FFF3CD").setFontColor("#856404").setBold(true)
    .setRanges([onHandRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(3)
    .setBackground("#D4EDDA").setFontColor("#155724").setBold(true)
    .setRanges([onHandRange]).build());

  const statusRange = inv.getRange(`G${V3_INV_FIRST_ROW}:G${V3_INV_LAST_ROW}`);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("OUT")
    .setBackground("#F8D7DA").setFontColor("#721C24").setBold(true)
    .setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("LOW")
    .setBackground("#FFF3CD").setFontColor("#856404").setBold(true)
    .setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("OK")
    .setBackground("#D4EDDA").setFontColor("#155724").setBold(true)
    .setRanges([statusRange]).build());

  // Shortage section conditional formatting (rows 47-85)
  const shortageColD = inv.getRange(`D47:D${47 + V3_ITEM_COUNT - 1}`);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground("#F8D7DA").setFontColor("#721C24").setBold(true)
    .setRanges([shortageColD]).build());

  const shortageColE = inv.getRange(`E47:E${47 + V3_ITEM_COUNT - 1}`);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("SHORT")
    .setBackground("#F8D7DA").setFontColor("#721C24").setBold(true)
    .setRanges([shortageColE]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("OK")
    .setBackground("#D4EDDA").setFontColor("#155724").setBold(true)
    .setRanges([shortageColE]).build());

  inv.setConditionalFormatRules(rules);
}



// Alias kept so any old installable trigger or saved manual run that still
// references this name keeps working. Do not delete without confirming no
// trigger in the live sheet points at it (Extensions > Apps Script > Triggers).
function updateInventoryOutFormulas() {
  setupReconciliationLogic();
}


function setupProtections() {
  setupProtectionsCore_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert(
    "\u2713 Protections set up for v3.0.\n\n" +
    "Uses sheet-level protection with editable-range exceptions \u2014 " +
    "completes in seconds instead of timing out.\n\n" +
    "Formula columns are locked on Waiting List, Completed Deliveries, " +
    "and Cancelled Requests. Config is locked except for Tracking Start " +
    "Date, Book Titles, and Ministry Options. Inventory Position is fully locked."
  );
}

// Silent core — shared by the menu wrapper and setupEverything().
function setupProtectionsCore_(ss) {
  const owner = Session.getEffectiveUser();

  // ── Helper: remove any existing protections from a sheet first ──
  // (prevents duplicate protections stacking when re-running setup)
  function clearSheetProtections(sheet) {
    const sheetProts = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    sheetProts.forEach(p => { try { p.remove(); } catch (e) {} });
    const rangeProts = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    rangeProts.forEach(p => { try { p.remove(); } catch (e) {} });
  }

  // ── Helper: protect entire sheet, allow specified ranges as editable ──
  function protectSheetWithExceptions(sheet, description, editableRangeStrings) {
    const p = sheet.protect().setDescription(description);
    p.addEditor(owner);
    // Remove all editors except owner so the protection actually restricts others
    p.getEditors()
      .filter(e => e.getEmail() !== owner.getEmail())
      .forEach(e => {
        try { p.removeEditor(e); } catch (err) {}
      });
    if (p.canDomainEdit()) p.setDomainEdit(false);
    if (editableRangeStrings && editableRangeStrings.length > 0) {
      p.setUnprotectedRanges(editableRangeStrings.map(r => sheet.getRange(r)));
    }
    return p;
  }

  // Compute once: the last column letter (column CJ = 114 in v3.0)
  const wlLastColLtr = getColumnA1_(V3_COL_CONSUMPTION_LAST);  // "CJ"
  const incNotesColLtr = getColumnA1_(V3_INC_NOTES_COL);        // "AQ"

  // ── Config tab: protect entire sheet except Tracking Start Date + Book Titles ──
  const cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  if (cfgSheet) {
    clearSheetProtections(cfgSheet);
    protectSheetWithExceptions(cfgSheet, "Config — ministry lead only", [
      "B19",      // Tracking Start Date
      "B24:B26",  // Book Titles
      "E11",      // Ministry Name (v3.6 Ministry Options)
      "E13:E21",  // Ministry Options toggles
    ]);
  }

  // ── Inventory Position: fully protected (all formulas + snapshot grid) ──
  const invSheet = ss.getSheetByName(INVENTORY_SHEET);
  if (invSheet) {
    clearSheetProtections(invSheet);
    protectSheetWithExceptions(invSheet, "Inventory Position — owner only");
  }

  // ── Incoming Items: header row protected, everything else editable ──
  const incSheet = ss.getSheetByName(INCOMING_SHEET);
  if (incSheet) {
    clearSheetProtections(incSheet);
    protectSheetWithExceptions(incSheet, "Incoming Items — header protected",
      ["A2:" + incNotesColLtr + "9999"]);
  }

  // ── Reconciliation Log: inputs editable (date, component, count, counted by, notes)
  //    Calculated Count (C) and Adjustment (E) are formula columns — protected ──
  const recSheet = ss.getSheetByName(RECON_SHEET);
  if (recSheet) {
    clearSheetProtections(recSheet);
    protectSheetWithExceptions(recSheet, "Reconciliation Log — formulas protected", [
      "A2:B9999",   // Date, Component (editable)
      "D2:D9999",   // Physical Count (editable)
      "F2:G9999",   // Counted By, Notes (editable)
    ]);
  }

  // ── Waiting List: protect the sheet, allow editing of user-input columns only ──
  // Input columns (editable):
  //   A (Date Responded), B (APN), C (Status),
  //   E–N (Cribs, Toddlers, Twins, Bunks, then admin fields)
  //   S (Themed Bedding flag — col 19)
  //   X, Y, Z, AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM
  //     (8 child age + 8 gender, cols 24–39)
  //   AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ, BA, BB, BC, BD, BE
  //     (cols 40–57: caregiver, address, caseworker fields, delivery date,
  //      delivery contact, notes, logged-flag)
  //   BH (Bunk Config — col 60), BI (Pack Config — col 61)
  // Everything else is formula (D, O-R, T-W, BF-BG, BJ-BW, BX-CJ) → protected.
  const wlSheet = ss.getSheetByName(ACTIVE_SHEET);
  if (wlSheet) {
    clearSheetProtections(wlSheet);
    protectSheetWithExceptions(wlSheet, "Waiting List — formula columns protected", [
      "A2:C9999",   // Date, APN, Status
      "E2:N9999",   // Bed counts (cribs, toddlers, twins, bunks) + admin fields
      "S2:S9999",   // Themed Bedding? flag
      "X2:AM9999",  // Children age/gender (8 pairs)
      "AN2:BE9999", // Caregiver through delivery details and Notes
      "BH2:BI9999", // Bunk Config, Pack Config (pack-time inputs)
    ]);
  }

  // ── Completed Deliveries: same editable-ranges as Waiting List ──
  // Rows flow in here via the archive script; volunteers may occasionally
  // need to correct a delivery date or note, so inputs stay editable.
  const cdSheet = ss.getSheetByName(ARCHIVE_SHEET);
  if (cdSheet) {
    clearSheetProtections(cdSheet);
    protectSheetWithExceptions(cdSheet, "Completed Deliveries — formulas protected", [
      "A2:C9999",
      "E2:N9999",
      "S2:S9999",
      "X2:AM9999",
      "AN2:BE9999",
      "BH2:BI9999",
    ]);
  }

  // ── Cancelled Requests: same as Completed Deliveries ──
  const cancelledSheet = ss.getSheetByName(CANCELLED_SHEET);
  if (cancelledSheet) {
    clearSheetProtections(cancelledSheet);
    protectSheetWithExceptions(cancelledSheet, "Cancelled Requests — formulas protected", [
      "A2:C9999",
      "E2:N9999",
      "S2:S9999",
      "X2:AM9999",
      "AN2:BE9999",
      "BH2:BI9999",
    ]);
  }

}

function getColumnA1_(colNum) {
  let result = "";
  while (colNum > 0) {
    colNum--;
    result = String.fromCharCode(65 + (colNum % 26)) + result;
    colNum = Math.floor(colNum / 26);
  }
  return result;
}

function setupDataValidation() {
  setupDataValidationCore_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert("\u2713 Data validation set up for v3.0 (39 items on Incoming Items, 39-item Recon Log dropdown).");
}

// Silent core — shared by the menu wrapper and setupEverything().
function setupDataValidationCore_(ss) {
  const incSheet = ss.getSheetByName(INCOMING_SHEET);
  if (incSheet) {
    const lastRow = Math.max(incSheet.getLastRow(), 100);
    const sourceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Constructed", "Purchased", "Donated"], true)
      .setAllowInvalid(false).setHelpText("Select Constructed, Purchased, or Donated").build();
    incSheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(sourceRule);

    const numRule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 999).setAllowInvalid(false)
      .setHelpText("Enter a whole number between 0 and 999").build();
    // v3.0: Incoming Items has 39 quantity columns (D through AP)
    incSheet.getRange(2, V3_INC_FIRST_COL, lastRow - 1, V3_ITEM_COUNT).setDataValidation(numRule);

    const dateRule = SpreadsheetApp.newDataValidation()
      .requireDate().setAllowInvalid(false).setHelpText("Enter a valid date").build();
    incSheet.getRange(2, 1, lastRow - 1, 1).setDataValidation(dateRule);
  }

  const recSheet = ss.getSheetByName(RECON_SHEET);
  const cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  if (recSheet && cfgSheet) {
    const lastRow = Math.max(recSheet.getLastRow(), 50);
    // v3.0: Components range is A2:A40 (39 items)
    const compRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(cfgSheet.getRange("A2:A40"), true)
      .setAllowInvalid(false).setHelpText("Select a component from the list").build();
    recSheet.getRange(2, 2, lastRow - 1, 1).setDataValidation(compRule);

    const dateRule = SpreadsheetApp.newDataValidation()
      .requireDate().setAllowInvalid(false).setHelpText("Enter a valid date").build();
    recSheet.getRange(2, 1, lastRow - 1, 1).setDataValidation(dateRule);

    const countRule = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false)
      .setHelpText("Enter the physical count (whole number)").build();
    recSheet.getRange(2, 4, lastRow - 1, 1).setDataValidation(countRule);
  }

  const wlSheet = ss.getSheetByName(ACTIVE_SHEET);
  if (wlSheet) {
    const lastRow = Math.max(wlSheet.getLastRow(), 100);

    // Status dropdown
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Active", "Packed", "Completed", "Cancelled"], true)
      .setAllowInvalid(false).setHelpText("Select the current status of this request.").build();
    wlSheet.getRange(2, STATUS_COL, lastRow - 1, 1).setDataValidation(statusRule);

    // Gender dropdowns — child gender columns: Y(25), AA(27), AC(29), AE(31), AG(33), AI(35), AK(37), AM(39)
    const genderRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["M", "F"], true)
      .setAllowInvalid(true).setHelpText("M or F").build();
    const genderCols = [25, 27, 29, 31, 33, 35, 37, 39]; // Y, AA, AC, AE, AG, AI, AK, AM
    genderCols.forEach(col => {
      wlSheet.getRange(2, col, lastRow - 1, 1).setDataValidation(genderRule);
    });
  }

}

function setupNamedRanges() {
  try {
    setupNamedRangesCore_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {
    SpreadsheetApp.getUi().alert(e.message);
    return;
  }
  SpreadsheetApp.getUi().alert('\u2713 Named range "Components" created (A2:A40, 39 components).');
}

// Silent core — shared by the menu wrapper and setupEverything().
function setupNamedRangesCore_(ss) {
  const cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!cfgSheet) throw new Error("Config sheet not found.");

  ss.getNamedRanges().forEach(nr => {
    if (nr.getName() === "Components") nr.remove();
  });

  // v3.0: 39 components at A2:A40
  ss.setNamedRange("Components", cfgSheet.getRange("A2:A40"));
}

function installEditTrigger() {
  installEditTriggerCore_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert("✓ Installable edit trigger created.");
}

function installEditTriggerCore_(ss) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "handleEdit")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("handleEdit").forSpreadsheet(ss).onEdit().create();
}

function installFormTrigger() {
  installFormTriggerCore_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert("✓ Form submit trigger installed.");
}

function installFormTriggerCore_(ss) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "onFormSubmit")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("onFormSubmit").forSpreadsheet(ss).onFormSubmit().create();
}


// ═══════════════════════════════════════════════════════════════════════════
//  DASHBOARD CHART
// ═══════════════════════════════════════════════════════════════════════════

function buildDashboardChart() {
  try {
    buildDashboardChartCore_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (e) {
    SpreadsheetApp.getUi().alert(e.message);
    return;
  }
  SpreadsheetApp.getUi().alert("\u2713 Deliveries by Year chart built on the Dashboard.");
}

// Silent core — shared by the menu wrapper and setupEverything().
function buildDashboardChartCore_(ss) {
  const dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  if (!dashboard) throw new Error('Could not find the "Dashboard" sheet.');

  dashboard.getCharts().forEach(c => dashboard.removeChart(c));

  const DATA_START_ROW = 1;
  const DATA_END_ROW   = 13;
  const YEAR_COL       = 11;
  const FAM_COL        = 12;
  const BEDS_COL       = 13;

  const yearRange = dashboard.getRange(DATA_START_ROW + 1, YEAR_COL, DATA_END_ROW - 1, 1);
  const bedsRange = dashboard.getRange(DATA_START_ROW,     BEDS_COL, DATA_END_ROW,     1);
  const famRange  = dashboard.getRange(DATA_START_ROW,     FAM_COL,  DATA_END_ROW,     1);

  const chart = dashboard.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(bedsRange).addRange(famRange)
    .setOption('title', '')
    .setOption('legend', { position: 'bottom' })
    .setOption('series', {
      0: { color: '#2E75B6', labelInLegend: 'Total Beds' },
      1: { color: '#70AD47', labelInLegend: 'Families Served' }
    })
    .setOption('hAxis', { title: '', ticks: yearRange.getValues().flat().filter(v => v !== ''), format: '0' })
    .setOption('vAxis', { title: 'Count', format: '#,##0', minValue: 0 })
    .setOption('bar', { groupWidth: '70%' })
    .setOption('backgroundColor', 'transparent')
    .setOption('chartArea', { left: 60, top: 20, width: '85%', height: '75%' })
    .setNumHeaders(1).setTransposeRowsAndColumns(false)
    .setPosition(15, 2, 0, 0)
    .build();

  dashboard.insertChart(chart);
}


// ═══════════════════════════════════════════════════════════════════════════
//  SYSTEM HEALTH CHECK (v3.3)
//  Read-only. Verifies the moving parts that tend to drift or break silently —
//  installable triggers, the TOTALS anchor, the Components named range, sheet
//  protections, and notification recipients — and reports a clear pass/fail.
//  Run from Bed Ministry > Setup > Check system health.
// ═══════════════════════════════════════════════════════════════════════════

function runHealthCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const lines = [];
  const ok   = s => lines.push("✓ " + s);
  const bad  = s => lines.push("✗ " + s);
  const warn = s => lines.push("⚠ " + s);

  // 0. Version + timezone (v3.6)
  lines.push("System version: v" + SYSTEM_VERSION + " — updates: " + STARTER_KIT_REPO_URL);
  const sheetTz  = ss.getSpreadsheetTimeZone();
  const scriptTz = Session.getScriptTimeZone();
  if (sheetTz === scriptTz) ok("Timezone: " + sheetTz + " (sheet and script agree).");
  else warn("Timezone mismatch — sheet is " + sheetTz + " but the script is " + scriptTz +
            ". Fix in File > Settings and Extensions > Apps Script > Project Settings.");

  // 1. Installable triggers
  let handlers = [];
  try {
    handlers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  } catch (e) {
    warn("Could not read triggers: " + e.message);
  }
  if (handlers.indexOf("handleEdit") >= 0) ok("Edit trigger (handleEdit) installed.");
  else bad("Edit trigger MISSING — run Setup > Install edit trigger. Without it, " +
           "status changes will not archive.");
  if (handlers.indexOf("onFormSubmit") >= 0) ok("Form trigger (onFormSubmit) installed.");
  else bad("Form trigger MISSING — run Setup > Install form trigger. Without it, new " +
           "form requests will not appear on the Waiting List.");

  // 2. TOTALS anchor on the Waiting List
  const wl = ss.getSheetByName(ACTIVE_SHEET);
  if (!wl) {
    bad("Waiting List tab not found.");
  } else {
    let found = false;
    const last = wl.getLastRow();
    if (last >= 2) {
      const colB = wl.getRange(2, 2, last - 1, 1).getValues();
      found = colB.some(r => String(r[0]).trim().toUpperCase() === TOTALS_ROW_LABEL);
    }
    if (found) ok("TOTALS row present on the Waiting List.");
    else bad("TOTALS row MISSING — run Setup > Repair: Rebuild TOTALS row.");
  }

  // 3. Components named range
  const comp = ss.getNamedRanges().find(nr => nr.getName() === "Components");
  if (!comp) {
    bad("Named range 'Components' MISSING — run Setup > Set up named ranges.");
  } else {
    let a1 = "";
    try { a1 = comp.getRange().getA1Notation(); } catch (e) { a1 = "(unreadable)"; }
    if (a1.indexOf("A2:A40") >= 0) ok("Named range 'Components' = " + a1 + ".");
    else warn("Named range 'Components' is " + a1 + " (expected A2:A40).");
  }

  // 4. Protections on the key sheets
  [ACTIVE_SHEET, ARCHIVE_SHEET, CANCELLED_SHEET, CONFIG_SHEET, INVENTORY_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { warn("Sheet '" + name + "' not found."); return; }
    const prot = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (prot.length > 0) ok("'" + name + "' is protected.");
    else warn("'" + name + "' is NOT protected — run Setup > Set up formula protections.");
  });

  // 5. Notification recipients
  const cfg = ss.getSheetByName(CONFIG_SHEET);
  const o2  = cfg ? String(cfg.getRange("O2").getValue() || "").trim() : "";
  if (o2) ok("Notification recipients set in Config!O2: " + o2 + ".");
  else warn("Config!O2 is empty — failure alerts will go to the script owner (" +
            getNotificationRecipients_().join(", ") + "). Add the ministry email(s) " +
            "to Config!O2 to be sure they reach the right people.");

  // 6. Linked form vs FORM_SCHEMA (v3.6) — the #1 silent failure mode is a
  //    renamed form question that onFormSubmit can no longer match.
  try {
    const formWarnings = verifyFormSchema_(ss);
    if (formWarnings.length === 0) ok("Linked form questions match FORM_SCHEMA.");
    else formWarnings.forEach(w => warn(w));
  } catch (e) {
    warn("Form schema check skipped: " + e.message);
  }

  // 7. Ministry Options summary (v3.6)
  const opts = getMinistryOptions_();
  const offLabels = MINISTRY_OPTIONS.filter(o => !opts[o.key]).map(o => o.label);
  if (offLabels.length) {
    lines.push("Ministry Options — disabled: " + offLabels.join(", ") +
               " (change on Config, then run Refresh Ministry Options).");
  }

  // 8. Reminder: backup editor (sole-owner risk)
  lines.push("");
  lines.push("Tip: you are the sole owner. Consider sharing edit access with one " +
             "trusted backup person so the sheet is never locked to a single account.");

  ui.alert("Bed Ministry — System Health Check", lines.join("\n"), ui.ButtonSet.OK);
}


// ═══════════════════════════════════════════════════════════════════════════
//  v3.6 STARTER KIT — core-sheet bootstrap, form automation, ministry
//  options, one-click setup, and the New City wizard.
// ═══════════════════════════════════════════════════════════════════════════

// Base Waiting List headers (columns A–BH). Written only when a sheet is
// brand new / has an empty header row — never overwrites an existing layout.
const WL_BASE_HEADERS = [
  "Date Responded to Request", "Tracking # (APN)", "Status", "No. Beds",
  "Cribs", "Toddlers", "Twins", "Bunks",
  "Twin Frames Only", "Mattress Only", "Bedding Only", "Bedding Plus Mattress",
  "Bed/Bedding Picked Up by Delivery Team", "Date Material Picked Up",
  '8" Mattresses', '6" Mattresses', "Crib Mattresses", "Pillows",
  "Themed Bedding?", "Toddler Books", "Young Books", "Young Bears", "Older Books",
  "Child 1\nAge", "Child 1\nGender", "Child 2\nAge", "Child 2\nGender",
  "Child 3\nAge", "Child 3\nGender", "Child 4\nAge", "Child 4\nGender",
  "Child 5\nAge", "Child 5\nGender", "Child 6\nAge", "Child 6\nGender",
  "Child 7\nAge", "Child 7\nGender", "Child 8\nAge", "Child 8\nGender",
  "Caregiver", "Caregiver Phone", "Complex / Building Name", "Street Address",
  "Apt / Unit #", "City", "ZIP", "Gate Code", "Floor #",
  "Case Worker", "Case Worker Phone", "Delivery Date",
  "Caseworker Notified of Delivery", "Delivery Contact", "Delivery Contact Phone",
  "Delivery Organization", "Notes", "Logged",
  "Shortage Impact", "Constrained Items", "Bunk Config",
];

// Create any missing core sheets and give brand-new ones their base layout.
// Existing sheets are never restructured — safe on the San Antonio original,
// essential when building the starter-kit template from a blank workbook.
function ensureCoreSheets_(ss) {
  const created = [];
  const getOrCreate = name => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); created.push(name); }
    return sh;
  };

  const headerStyle = rng => rng
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
    .setFontFamily("Arial").setFontSize(10).setHorizontalAlignment("center")
    .setWrap(true);

  // Request tabs — Waiting List + both archives share the 60-column base layout.
  [ACTIVE_SHEET, ARCHIVE_SHEET, CANCELLED_SHEET].forEach(name => {
    const sh = getOrCreate(name);
    if (String(sh.getRange(1, 1).getValue() || "") === "") {
      if (sh.getMaxColumns() < V3_LAST_COL) {
        sh.insertColumnsAfter(sh.getMaxColumns(), V3_LAST_COL - sh.getMaxColumns());
      }
      headerStyle(sh.getRange(1, 1, 1, WL_BASE_HEADERS.length).setValues([WL_BASE_HEADERS]));
      sh.setFrozenRows(1);
    }
  });

  // Config / Incoming Items / Inventory Position — content comes from the v3
  // stages; here we only guarantee the sheets exist (plus the notification
  // header, which no stage writes).
  const cfg = getOrCreate(CONFIG_SHEET);
  if (String(cfg.getRange("O1").getValue() || "") === "") {
    cfg.getRange("O1").setValue("Notification Recipients")
      .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
      .setFontFamily("Arial").setFontSize(11);
  }
  getOrCreate(INCOMING_SHEET);
  getOrCreate(INVENTORY_SHEET);

  // Reconciliation Log header
  const rec = getOrCreate(RECON_SHEET);
  if (String(rec.getRange(1, 1).getValue() || "") === "") {
    headerStyle(rec.getRange(1, 1, 1, 7).setValues([[
      "Date", "Component", "Calculated Count", "Physical Count",
      "Adjustment", "Counted By", "Notes",
    ]]));
    rec.setFrozenRows(1);
  }

  // Dashboard — impact scaffold with generic formulas (no historical offsets;
  // a city that starts tracking from day one needs none).
  const dash = getOrCreate(DASHBOARD_SHEET);
  if (String(dash.getRange("B4").getValue() || "") === "") {
    buildDashboardScaffold_(dash);
  }

  // TOTALS anchor on the Waiting List
  ensureTotalsRow_(ss.getSheetByName(ACTIVE_SHEET));

  return created;
}

// Write the Dashboard's summary cells and by-year table on a fresh sheet.
// Formulas reference the year cells in column K (SA's originals hardcoded
// each year), so the same scaffold works for any city and any start year.
function buildDashboardScaffold_(dash) {
  const cd = "'Completed Deliveries'";

  dash.getRange("B2").setValue(getMinistryName_() + " — Impact Dashboard")
    .setFontSize(16).setFontWeight("bold").setFontColor("#1A3C5E");
  dash.getRange("B3").setValue("  All-Time Totals")
    .setFontSize(12).setFontWeight("bold").setFontColor("#4A6B8A");

  const kpiLabels = ["Families\nServed", "Total Beds\nDelivered", "Books\nGiven", "Bears\nGiven", "Years\nServing"];
  dash.getRange(4, 2, 1, 5).setValues([kpiLabels])
    .setFontWeight("bold").setHorizontalAlignment("center").setWrap(true)
    .setBackground("#1A3C5E").setFontColor("#FFFFFF");
  dash.getRange("B5").setFormula(`=COUNTA(${cd}!A2:A9990)-COUNTIF(${cd}!A2:A9990,"TOTALS")`);
  dash.getRange("C5").setFormula(`=SUM(${cd}!D2:D9990)`);
  dash.getRange("D5").setFormula(`=SUM(${cd}!T2:T9990)+SUM(${cd}!U2:U9990)+SUM(${cd}!W2:W9990)`);
  dash.getRange("E5").setFormula(`=SUM(${cd}!V2:V9990)`);
  dash.getRange("F5").setFormula(`=IFERROR(YEAR(MAX(${cd}!A2:A9990))-YEAR(MIN(${cd}!A2:A9990))+1,"")`);
  dash.getRange(5, 2, 1, 5).setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");

  dash.getRange("B8").setValue("  Bed Types Delivered (All Time)")
    .setFontSize(12).setFontWeight("bold").setFontColor("#4A6B8A");
  dash.getRange(9, 2, 1, 7).setValues([[
    "Twins", "Bunks", "Toddler Beds", "Cribs", "Frames Only", "Mattress Only", "Bedding Only",
  ]]).setFontWeight("bold").setHorizontalAlignment("center")
    .setBackground("#4A6B8A").setFontColor("#FFFFFF");
  const bedTypeCols = ["G", "H", "F", "E", "I", "J", "K"];
  bedTypeCols.forEach((col, i) => {
    dash.getRange(10, 2 + i).setFormula(`=SUM(${cd}!${col}2:${col}9990)`);
  });
  dash.getRange(10, 2, 1, 7).setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center");

  dash.getRange("B14").setValue("  Deliveries by Year")
    .setFontSize(12).setFontWeight("bold").setFontColor("#4A6B8A");

  // By-year table (K:Q): 12 years starting from the current year.
  dash.getRange(1, 11, 1, 7).setValues([[
    "Year", "Families", "Total Beds", "Twins", "Bunks", "Toddlers", "Cribs",
  ]]).setFontWeight("bold");
  const startYear = new Date().getFullYear();
  for (let i = 0; i < 12; i++) {
    const r = 2 + i;
    dash.getRange(r, 11).setValue(startYear + i);
    dash.getRange(r, 12).setFormula(
      `=SUMPRODUCT(IFERROR((YEAR(${cd}!$AY$2:$AY$9990)=$K${r})*ISNUMBER(${cd}!$A$2:$A$9990),0))`);
    const yearCols = { 13: "D", 14: "G", 15: "H", 16: "F", 17: "E" };
    Object.keys(yearCols).forEach(c => {
      dash.getRange(r, Number(c)).setFormula(
        `=SUMPRODUCT(IFERROR((YEAR(${cd}!$AY$2:$AY$9990)=$K${r}),0)*IFERROR(${cd}!${yearCols[c]}$2:${yearCols[c]}$9990,0))`);
    });
  }
}

// ── Ministry Options block on Config ────────────────────────────────────────

function setupMinistryOptions() {
  const msg = setupMinistryOptionsCore_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert(msg);
}

function setupMinistryOptionsCore_(ss) {
  const cfg = ss.getSheetByName(CONFIG_SHEET);
  if (!cfg) return "Config sheet not found.";

  if (String(cfg.getRange(OPT_BLOCK_HEADER_ROW, OPT_COL_LABEL).getValue()) === "Ministry Options") {
    return "Ministry Options section already exists on Config (rows " +
           OPT_BLOCK_HEADER_ROW + "–" + (OPT_FIRST_ROW + MINISTRY_OPTIONS.length - 1) +
           "). Edit the YES/NO cells in column E, then run " +
           '"Refresh Ministry Options".';
  }

  // Section header
  cfg.getRange(OPT_BLOCK_HEADER_ROW, OPT_COL_LABEL, 1, 3).merge();
  cfg.getRange(OPT_BLOCK_HEADER_ROW, OPT_COL_LABEL).setValue("Ministry Options")
    .setFontWeight("bold").setFontSize(11).setFontColor("#FFFFFF")
    .setBackground("#1A3C5E").setFontFamily("Arial");

  // Ministry name row
  cfg.getRange(OPT_NAME_ROW, OPT_COL_LABEL).setValue("Ministry Name")
    .setFontWeight("bold").setFontFamily("Arial").setFontSize(10);
  cfg.getRange(OPT_NAME_ROW, OPT_COL_VALUE)
    .setBackground("#FFF8E1")
    .setNote('Shown on pick-list PDFs. Blank = "Bed Ministry".');

  // Column headers
  cfg.getRange(OPT_FIRST_ROW - 1, OPT_COL_LABEL).setValue("Option")
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
    .setFontFamily("Arial").setFontSize(10);
  cfg.getRange(OPT_FIRST_ROW - 1, OPT_COL_VALUE).setValue("Enabled? (YES/NO)")
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1A3C5E")
    .setFontFamily("Arial").setFontSize(10);

  // Toggle rows — blank means YES, so nothing changes until a city opts out.
  const yesNo = SpreadsheetApp.newDataValidation()
    .requireValueInList(["YES", "NO"], true).setAllowInvalid(false)
    .setHelpText("YES (or blank) = your ministry provides this; NO = hide it.").build();
  MINISTRY_OPTIONS.forEach((o, i) => {
    const r = OPT_FIRST_ROW + i;
    cfg.getRange(r, OPT_COL_LABEL).setValue(o.label)
      .setFontFamily("Arial").setFontSize(10);
    cfg.getRange(r, OPT_COL_VALUE).setBackground("#FFF8E1").setDataValidation(yesNo);
    if (i % 2 === 1) cfg.getRange(r, OPT_COL_LABEL).setBackground("#F2F5F8");
  });

  // Instruction note
  const noteRow = OPT_FIRST_ROW + MINISTRY_OPTIONS.length;
  cfg.getRange(noteRow, OPT_COL_LABEL, 1, 5).merge();
  cfg.getRange(noteRow, OPT_COL_LABEL)
    .setValue('Blank = YES. After changing a toggle, run Bed Ministry > Setup > "Refresh Ministry Options" ' +
              "so demand formulas and the inventory view match.")
    .setFontSize(9).setFontStyle("italic").setFontColor("#777777").setWrap(true)
    .setFontFamily("Arial");

  return "✓ Ministry Options section added to Config (columns D/E, rows " +
         OPT_BLOCK_HEADER_ROW + "–" + noteRow + "). Set the ministry name in E" +
         OPT_NAME_ROW + " and YES/NO toggles in E" + OPT_FIRST_ROW + ":E" +
         (OPT_FIRST_ROW + MINISTRY_OPTIONS.length - 1) + ".";
}

// Re-apply everything that depends on the toggles: refresh demand formulas on
// Active/Packed rows and grey out disabled items on Inventory Position.
// Run after changing any YES/NO cell. Columns are never removed, so a toggle
// can be flipped back on at any time — just run this again.
function refreshMinistryOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ministryOptionsCache = null;
  const opts = getMinistryOptions_();

  // 1. Rewrite formulas for Active/Packed Waiting List rows
  const wl = ss.getSheetByName(ACTIVE_SHEET);
  let rewritten = 0;
  if (wl) {
    const lastRow = wl.getLastRow();
    for (let r = 2; r <= lastRow; r++) {
      const b = String(wl.getRange(r, 2).getValue());
      const c = String(wl.getRange(r, 3).getValue());
      if (b === TOTALS_ROW_LABEL) continue;
      if (c !== "Active" && c !== "Packed") continue;
      writeRowFormulas_(wl, r);
      rewritten++;
    }
  }

  // 2. Grey out disabled items on Inventory Position (snapshot grid + the
  //    shortage section); restore enabled ones.
  refreshMinistryOptionsSilent_(ss);

  const offLabels = MINISTRY_OPTIONS.filter(o => !opts[o.key]).map(o => o.label);
  SpreadsheetApp.getUi().alert(
    "✓ Ministry Options applied.\n\n" +
    "Disabled: " + (offLabels.length ? offLabels.join(", ") : "(none — everything enabled)") + "\n" +
    "Waiting List rows refreshed: " + rewritten + "\n\n" +
    "Disabled items forecast 0 demand, default to 0 on pick lists, and are " +
    "greyed on Inventory Position. Their columns remain, so flipping a " +
    "toggle back on later needs no migration."
  );
}

// Inventory greyout without the alert — used by the wizard and refresh.
function refreshMinistryOptionsSilent_(ss) {
  try {
    const opts = getMinistryOptions_();
    const inv = ss.getSheetByName(INVENTORY_SHEET);
    if (!inv) return;
    const disabledNames = {};
    Object.keys(OPT_CATALOG_ITEMS).forEach(key => {
      if (!opts[key]) OPT_CATALOG_ITEMS[key].forEach(n => { disabledNames[n] = true; });
    });
    V3_CATALOG.forEach((item, idx) => {
      const color = disabledNames[item.name] ? "#B7B7B7" : "#000000";
      inv.getRange(V3_INV_FIRST_ROW + idx, 1, 1, 7).setFontColor(color);
      inv.getRange(47 + idx, 1, 1, 5).setFontColor(color);
    });
  } catch (e) { /* cosmetic */ }
}

// ── Form automation ─────────────────────────────────────────────────────────

// Create the Bed Request Form from FORM_SCHEMA and link it to this workbook.
// Refuses to run when a form is already linked (protects a live deployment).
function setupForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let linkedUrl = null;
  try { linkedUrl = ss.getFormUrl(); } catch (e) {}
  if (linkedUrl) {
    ui.alert(
      "A form is already linked to this workbook:\n\n" + linkedUrl + "\n\n" +
      "Nothing was created. To rebuild the form, unlink the existing one " +
      "first (open it → Responses → unlink), or edit it by hand and run " +
      '"Check system health" to verify the questions still match.'
    );
    return;
  }

  const resp = ui.alert(
    "Create & link Bed Request Form",
    "This creates a new Google Form with the standard bed-request questions " +
    "(bed types disabled in Ministry Options are left out) and links its " +
    "responses to this workbook.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const beforeNames = ss.getSheets().map(sh => sh.getName());
  const form = buildFormFromSchema_(ss);

  // Link responses to this workbook, then normalize the response tab name so
  // guide screenshots and error messages match ("Form Responses 1").
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  try {
    const fresh = SpreadsheetApp.openById(ss.getId());
    const newSheet = fresh.getSheets().find(sh =>
      beforeNames.indexOf(sh.getName()) < 0 && /^Form Responses/.test(sh.getName()));
    if (newSheet && newSheet.getName() !== RESPONSES_SHEET &&
        !fresh.getSheetByName(RESPONSES_SHEET)) {
      newSheet.setName(RESPONSES_SHEET);
    }
  } catch (e) { /* cosmetic only */ }

  ui.alert(
    "✓ Bed Request Form created and linked.\n\n" +
    "Edit it here:\n" + form.getEditUrl() + "\n\n" +
    "Share THIS link with request submitters:\n" + form.getPublishedUrl() + "\n\n" +
    'Now run Setup > "Install form trigger" so submissions flow onto the ' +
    "Waiting List (the New City wizard does this for you)."
  );
}

// Build the form itself. Separate from setupForm so the wizard can reuse it.
function buildFormFromSchema_(ss) {
  const opts = getMinistryOptions_();
  const form = FormApp.create(getMinistryName_() + " — Bed Request Form");
  form.setDescription(
    "Request beds for children in need. The tracking number comes from your " +
    "referral system (e.g., CarePortal)."
  );

  const pageBreaks = {};   // pageId → PageBreakItem
  const navItems = [];     // { item: MultipleChoiceItem, nav: {yes, no} }

  FORM_SCHEMA.forEach(q => {
    if (q.type === "PAGE") {
      const pb = form.addPageBreakItem().setTitle(q.title);
      if (q.helpText) pb.setHelpText(q.helpText);
      if (q.pageId) pageBreaks[q.pageId] = pb;   // only nav targets are keyed
      return;
    }
    if (q.bedType) {
      const opt = MINISTRY_OPTIONS.find(o => o.bedType === q.bedType);
      if (opt && !opts[opt.key]) return;   // bed type disabled → drop question
    }
    let item;
    if (q.type === "DATE") {
      item = form.addDateItem().setTitle(q.title);
    } else if (q.type === "PARAGRAPH") {
      item = form.addParagraphTextItem().setTitle(q.title);
    } else if (q.type === "MC") {
      item = form.addMultipleChoiceItem().setTitle(q.title);
      if (q.nav) navItems.push({ item: item, nav: q.nav });   // choices set below
      else item.setChoiceValues(q.choices || []);
    } else {
      item = form.addTextItem().setTitle(q.title);
    }
    if (q.helpText) item.setHelpText(q.helpText);
    if (q.required) item.setRequired(true);
  });

  // Wire up the child-flow navigation now that the CAREGIVER page break exists:
  // "Yes" continues to the next child page, "No" jumps to Caregiver Information.
  navItems.forEach(({ item, nav }) => {
    const target = pageBreaks[nav.no];
    item.setChoices([
      item.createChoice("Yes", FormApp.PageNavigationType.CONTINUE),
      target ? item.createChoice("No", target)
             : item.createChoice("No", FormApp.PageNavigationType.SUBMIT),
    ]);
  });

  return form;
}

// Dump the linked form's schema as JSON — run this on a COPY of a live
// workbook to capture exact wording for FORM_SCHEMA (never needed on a fresh
// setupForm build; they match by construction).
function logLiveFormSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let url = null;
  try { url = ss.getFormUrl(); } catch (e) {}
  if (!url) { ui.alert("No form is linked to this workbook."); return; }

  const form = FormApp.openByUrl(url);
  const items = form.getItems().map((it, i) => {
    const entry = { index: i, type: String(it.getType()), title: it.getTitle() };
    const help = it.getHelpText();
    if (help) entry.helpText = help;
    try {
      if (it.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
        const mc = it.asMultipleChoiceItem();
        entry.choices = mc.getChoices().map(c => c.getValue());
        entry.required = mc.isRequired();
      } else if (it.getType() === FormApp.ItemType.LIST) {
        const li = it.asListItem();
        entry.choices = li.getChoices().map(c => c.getValue());
        entry.required = li.isRequired();
      } else if (it.getType() === FormApp.ItemType.TEXT) {
        entry.required = it.asTextItem().isRequired();
      } else if (it.getType() === FormApp.ItemType.DATE) {
        entry.required = it.asDateItem().isRequired();
      } else if (it.getType() === FormApp.ItemType.PARAGRAPH_TEXT) {
        entry.required = it.asParagraphTextItem().isRequired();
      }
    } catch (e) { entry.readError = e.message; }
    return entry;
  });

  const json = JSON.stringify(items, null, 2);
  Logger.log(json);
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:Arial;font-size:12px;">Live form schema (also in the ' +
    "execution log). Copy it out and reconcile with FORM_SCHEMA in Code.js:</p>" +
    '<textarea style="width:100%;height:340px;font-family:monospace;font-size:11px;">' +
    json.replace(/&/g, "&amp;").replace(/</g, "&lt;") +
    "</textarea>"
  ).setWidth(640).setHeight(430);
  ui.showModalDialog(html, "Live Form Schema");
}

// Health-check helper: compare the linked form's question titles against
// FORM_SCHEMA (normalized the same way onFormSubmit matches them). Returns
// warning strings; empty array = all good.
function verifyFormSchema_(ss) {
  const warnings = [];
  let url = null;
  try { url = ss.getFormUrl(); } catch (e) {}
  if (!url) {
    warnings.push("No form linked to this workbook — new requests cannot arrive. " +
                  'Run Setup > "Create & link Bed Request Form".');
    return warnings;
  }

  let liveTitles;
  try {
    const skipTypes = [FormApp.ItemType.PAGE_BREAK, FormApp.ItemType.SECTION_HEADER,
                       FormApp.ItemType.IMAGE, FormApp.ItemType.VIDEO];
    liveTitles = FormApp.openByUrl(url).getItems()
      .filter(it => skipTypes.indexOf(it.getType()) < 0)
      .map(it => it.getTitle());
  } catch (e) {
    warnings.push("Could not open the linked form to verify questions: " + e.message);
    return warnings;
  }

  const normKey = s => String(s).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  const liveSet = {};
  liveTitles.forEach(t => { liveSet[normKey(t)] = t; });

  const opts = getMinistryOptions_();
  FORM_SCHEMA.forEach(q => {
    if (q.type === "PAGE") return;
    if (q.bedType) {
      const opt = MINISTRY_OPTIONS.find(o => o.bedType === q.bedType);
      if (opt && !opts[opt.key]) return;   // intentionally absent
    }
    if (!liveSet[normKey(q.title)]) {
      warnings.push('Form question missing or renamed: "' + q.title + '" — ' +
                    "onFormSubmit will drop this field. Restore the exact wording.");
    }
  });
  return warnings;
}

// ── One-click setup + New City wizard ───────────────────────────────────────

// Run the entire setup chain in the documented order. Idempotent; used to
// build the starter-kit template from a blank workbook and to repair a copy.
function setupEverything() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    "Set up EVERYTHING",
    "Runs the full setup chain: core sheets, v3 structure (39-item catalog), " +
    "reconciliation logic, book titles, fulfillment-config column, Ministry " +
    "Options, data validation, named ranges, protections, and the Dashboard " +
    "chart.\n\nSafe to re-run — each step is idempotent. On a workbook that " +
    "is already set up it only fills gaps.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const log = [];
  const step = (label, fn) => {
    ss.toast(label + "…", "Setup", 5);
    const out = fn();
    log.push("✓ " + label + (typeof out === "string" && out ? " — " + out : ""));
  };

  try {
    step("Core sheets", () => {
      const created = ensureCoreSheets_(ss);
      return created.length ? "created: " + created.join(", ") : "all present";
    });
    // v3Stage1_Config clears the Tracking Start Date (correct for a first-time
    // migration, wrong for a repair re-run on an operating city) — preserve it.
    const cfgSheet = ss.getSheetByName(CONFIG_SHEET);
    const savedTrackingDate = cfgSheet ? cfgSheet.getRange("B19").getValue() : "";
    step("v3 structure — Config catalog",       () => v3Stage1_Config(ss).trim());
    if (cfgSheet && savedTrackingDate) cfgSheet.getRange("B19").setValue(savedTrackingDate);
    step("v3 structure — Incoming Items",       () => v3Stage2_IncomingItems(ss).trim());
    step("v3 structure — Inventory Position",   () => v3Stage3_InventoryPositionHeader(ss).trim());
    step("v3 structure — Waiting List columns", () => v3Stage4_WaitingList(ss).trim());
    step("v3 structure — archive tabs",         () => v3Stage5_ArchiveTabs(ss).trim());
    step("Reconciliation logic",                () => { setupReconciliationLogicCore_(ss); });
    step("Book Titles section",                 () => setupBookTitlesCore_(ss));
    step("Fulfillment-config column",           () => setupFulfillmentConfigColumnCore_(ss).join("; "));
    step("Ministry Options section",            () => setupMinistryOptionsCore_(ss));
    step("Data validation",                     () => { setupDataValidationCore_(ss); });
    step("Named ranges",                        () => { setupNamedRangesCore_(ss); });
    step("Formula protections",                 () => { setupProtectionsCore_(ss); });
    step("Dashboard chart",                     () => { buildDashboardChartCore_(ss); });
  } catch (e) {
    ui.alert(
      "Setup stopped on an error",
      "Error: " + e.message + "\n\nCompleted so far:\n" + log.join("\n") +
      "\n\nFix the issue and re-run — completed steps are safe to repeat.",
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert(
    "✓ Everything is set up",
    log.join("\n") + "\n\nStill to do (the New City wizard covers all of it):\n" +
    "  • Create & link the Bed Request Form\n" +
    "  • Install the form + edit triggers\n" +
    "  • Fill in Config: notification emails, ministry name, options",
    ui.ButtonSet.OK
  );
}

// Guided setup for a church that just copied the template workbook.
// Collects config, verifies the form link, installs triggers (copies don't
// inherit them), and finishes with a health check.
function newCityWizard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const intro = ui.alert(
    "🚀 New City Setup",
    "Welcome! This wizard configures your copy of the Bed Ministry workbook:\n\n" +
    "  1. Your ministry's name\n" +
    "  2. Who receives system alerts\n" +
    "  3. What your ministry provides (beds, bears, books…)\n" +
    "  4. The request form + automation triggers\n\n" +
    "Takes about 5 minutes. You can re-run it any time.\n\nReady?",
    ui.ButtonSet.YES_NO
  );
  if (intro !== ui.Button.YES) return;

  const cfg = ss.getSheetByName(CONFIG_SHEET);
  if (!cfg) {
    ui.alert('No Config sheet found. Run Setup > "Set up EVERYTHING" first, then re-run this wizard.');
    return;
  }
  setupMinistryOptionsCore_(ss);   // make sure the options block exists

  // 1. Ministry name
  const nameResp = ui.prompt(
    "Step 1 of 4 — Ministry name",
    'What is your ministry called? Shown on pick lists, e.g. "Riverside Bed Ministry".\n' +
    'Leave blank for plain "Bed Ministry":',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;
  const name = nameResp.getResponseText().trim();
  if (name) cfg.getRange(OPT_NAME_ROW, OPT_COL_VALUE).setValue(name);

  // 2. Notification recipients
  const emailResp = ui.prompt(
    "Step 2 of 4 — Alert emails",
    "Email address(es) that should receive system alerts (form errors, archive " +
    "failures). Separate multiple addresses with commas:",
    ui.ButtonSet.OK_CANCEL
  );
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  const emails = emailResp.getResponseText().trim();
  if (emails) cfg.getRange("O2").setValue(emails);

  // 3. Ministry options
  const optList = MINISTRY_OPTIONS.map((o, i) => "  " + (i + 1) + ". " + o.label).join("\n");
  const optResp = ui.prompt(
    "Step 3 of 4 — What do you provide?",
    "Everything is ON by default:\n\n" + optList + "\n\n" +
    "If there is anything your ministry does NOT provide, enter its number(s), " +
    "separated by commas (e.g. 5,6 if you don't give bears or books). Leave " +
    "blank to keep everything on:",
    ui.ButtonSet.OK_CANCEL
  );
  if (optResp.getSelectedButton() !== ui.Button.OK) return;
  const offNums = optResp.getResponseText().split(/[\s,;]+/)
    .map(t => parseInt(t, 10)).filter(n => n >= 1 && n <= MINISTRY_OPTIONS.length);
  MINISTRY_OPTIONS.forEach((o, i) => {
    cfg.getRange(OPT_FIRST_ROW + i, OPT_COL_VALUE)
      .setValue(offNums.indexOf(i + 1) >= 0 ? "NO" : "YES");
  });
  _ministryOptionsCache = null;

  // 4. Form + triggers. Copies of a workbook normally carry the linked form
  //    along, but that is Google behavior rather than a guarantee — so check,
  //    and fall back to creating one from FORM_SCHEMA.
  let formUrl = null;
  try { formUrl = ss.getFormUrl(); } catch (e) {}
  if (!formUrl) {
    const mk = ui.alert(
      "Step 4 of 4 — Request form",
      "No request form is linked to this workbook yet. Create one now with the " +
      "standard questions? (Recommended — you can adjust wording and design in " +
      "Google Forms afterwards.)",
      ui.ButtonSet.YES_NO
    );
    if (mk === ui.Button.YES) {
      const form = buildFormFromSchema_(ss);
      form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      SpreadsheetApp.flush();
      ui.alert(
        "✓ Form created.\n\nEdit: " + form.getEditUrl() +
        "\n\nShare with submitters: " + form.getPublishedUrl()
      );
    }
  }

  // Triggers never survive "Make a copy" — install fresh ones.
  installFormTriggerCore_(ss);
  installEditTriggerCore_(ss);

  // Timezone sanity note (both the sheet and the script carry one).
  const sheetTz  = ss.getSpreadsheetTimeZone();
  const scriptTz = Session.getScriptTimeZone();
  let tzNote = "Timezone: " + sheetTz;
  if (sheetTz !== scriptTz) tzNote += " (script: " + scriptTz + " — mismatch!)";

  ui.alert(
    "✓ New City Setup complete",
    "Ministry: " + getMinistryName_() + "\n" +
    "Alerts to: " + (emails || "(script owner)") + "\n" +
    "Disabled options: " +
      (offNums.length ? offNums.map(n => MINISTRY_OPTIONS[n - 1].label).join(", ") : "(none)") + "\n" +
    tzNote + "\n\n" +
    "If " + sheetTz + " is not your local timezone, fix BOTH:\n" +
    "  • Sheet: File > Settings > Time zone\n" +
    "  • Script: Extensions > Apps Script > Project Settings\n\n" +
    "A health check will now run to confirm everything is wired up.",
    ui.ButtonSet.OK
  );
  refreshMinistryOptionsSilent_(ss);
  runHealthCheck();
}

// About / version — cities compare against the repo CHANGELOG to see whether
// an update is available.
function aboutBedMinistry() {
  SpreadsheetApp.getUi().alert(
    "Bed Ministry System",
    "Version: v" + SYSTEM_VERSION + "\n\n" +
    "Updates & documentation:\n" + STARTER_KIT_REPO_URL + "\n\n" +
    "To update: open the repo's CHANGELOG to see what changed, then paste the " +
    "new Code.js over this script (Extensions > Apps Script). All your data " +
    "lives in the sheet — updating the script never touches it.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ═══════════════════════════════════════════════════════════════════════════
//  MENU
// ═══════════════════════════════════════════════════════════════════════════

// v3.4 — Ensure the appended Fulfillment Config column (115) exists with a header on
// the Waiting List AND both archive sheets, so copyRowFull_ (which archives through
// V3_LAST_COL) never overruns a sheet's column count. Idempotent: safe to re-run.
function setupFulfillmentConfigColumn() {
  const report = setupFulfillmentConfigColumnCore_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert(
    "\u2713 Fulfillment-config column ready (column " + V3_COL_FULFILL_CONFIG + ").\n\n" +
    report.join("\n") +
    "\n\nIt stores each delivery's substitution/omission choices as JSON and travels to " +
    "Completed Deliveries with the archived row."
  );
}

// Silent core — returns the per-sheet report. Shared by the menu wrapper
// and setupEverything().
function setupFulfillmentConfigColumnCore_(ss) {
  const header = "Fulfillment Config (JSON)";
  const report = [];
  [ACTIVE_SHEET, ARCHIVE_SHEET, CANCELLED_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { report.push("• " + name + ": NOT FOUND (skipped)"); return; }
    const maxCols = sh.getMaxColumns();
    if (maxCols < V3_LAST_COL) {
      sh.insertColumnsAfter(maxCols, V3_LAST_COL - maxCols);
      report.push("• " + name + ": added " + (V3_LAST_COL - maxCols) + " column(s)");
    } else {
      report.push("• " + name + ": already wide enough");
    }
    sh.getRange(HEADER_ROW, V3_COL_FULFILL_CONFIG).setValue(header).setFontWeight("bold");
  });
  return report;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Bed Ministry")
    .addItem("📋 Generate Pick List",         "generatePickList")
    .addItem("🩺 Check system health",         "runHealthCheck")
    .addSeparator()
    .addItem("↩️ Restore last archived row",   "restoreLastArchived")
    .addItem("↩️ Restore last cancelled row",  "restoreLastCancelled")
    .addSeparator()
    .addItem("🚀 New City Setup (run after copying)", "newCityWizard")
    .addSubMenu(ui.createMenu("⚙️ Setup")
      .addItem("★ Set up EVERYTHING (new workbook)",  "setupEverything")
      .addItem("★ Set up v3.0 inventory structure",   "setupV3Structure")
      .addItem("★ Create & link Bed Request Form",    "setupForm")
      .addSeparator()
      .addItem("🔧 Repair: fix blank component names", "v3RepairInventoryNames")
      .addItem("🔧 Repair: Waiting List formulas",     "v3RepairWaitingListFormulas")
      .addItem("Repair: Rebuild TOTALS row",           "repairTotalsRow")
      .addSeparator()
      .addItem("Set up data validation",             "setupDataValidation")
      .addItem("Set up named ranges",                "setupNamedRanges")
      .addItem("Set up formula protections",         "setupProtections")
      .addSeparator()
      .addItem("Add Book Titles section to Config",  "setupBookTitles")
      .addItem("Set up Ministry Options section",    "setupMinistryOptions")
      .addItem("Refresh Ministry Options",           "refreshMinistryOptions")
      .addItem("Set up fulfillment-config column",   "setupFulfillmentConfigColumn")
      .addItem("Set up reconciliation logic",         "setupReconciliationLogic")
      .addSeparator()
      .addItem("Install form trigger",               "installFormTrigger")
      .addItem("Install edit trigger",               "installEditTrigger")
      .addItem("Build Dashboard chart",              "buildDashboardChart")
      .addSeparator()
      .addItem("Dump live form schema to log",       "logLiveFormSchema")
      .addItem("🩺 Check system health",              "runHealthCheck")
      .addItem("ℹ️ About / version",                  "aboutBedMinistry")
    )
    .addToUi();
}