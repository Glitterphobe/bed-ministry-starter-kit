/* ============================================================
   BED MINISTRY PUBLIC DASHBOARD ENDPOINT  —  STANDALONE VERSION
   This runs as its OWN Apps Script project in your Google Drive,
   separate from the workbook. It reads the Dashboard sheet's
   finished totals by the workbook's ID and serves them as public
   JSON. No raw rows, no names, no addresses.

   >>> BEFORE IT WORKS: set a Script Property named WORKBOOK_ID
       to your Bed Ministry workbook's ID (the long string in the
       workbook URL between /d/ and /edit):
       Project Settings (gear icon) > Script Properties > Add.
   ============================================================ */

// Workbook ID comes from Script Properties so this same code serves any
// city's deployment — nothing to edit in the source.
function getWorkbookId_() {
  var id = PropertiesService.getScriptProperties().getProperty('WORKBOOK_ID');
  if (!id) {
    throw new Error(
      'Set a Script Property named WORKBOOK_ID to your Bed Ministry ' +
      'workbook ID (Project Settings > Script Properties).');
  }
  return id;
}


/**
 * Web app entry point. Returns aggregate dashboard data as JSON.
 * Supports a ?callback= parameter for JSONP (used by website embeds
 * so the browser doesn't block the request across domains).
 */
function doGet(e) {
  var json = bedMinistryBuildDashboardJson_();

  if (e && e.parameter && e.parameter.callback) {
    var safeName = String(e.parameter.callback).replace(/[^\w$.]/g, '');
    return ContentService
      .createTextOutput(safeName + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Reads the Dashboard sheet and packages the public totals.
 * Cached for 1 hour so the website stays fast and we stay well
 * under Google's daily request limits.
 */
function bedMinistryBuildDashboardJson_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('bed_ministry_dashboard_json');
  if (cached) {
    return cached;
  }

  // The one line that changed from the bound version: open the
  // workbook by ID instead of assuming we live inside it.
  var ss = SpreadsheetApp.openById(getWorkbookId_());
  var sheet = ss.getSheetByName('Dashboard');

  function num(a1) {
    var v = sheet.getRange(a1).getValue();
    return (typeof v === 'number' && isFinite(v)) ? v : 0;
  }

  var years = sheet.getRange('K2:K13').getValues();
  var beds  = sheet.getRange('M2:M13').getValues();
  var bedsByYear = [];
  for (var i = 0; i < years.length; i++) {
    var y = years[i][0];
    if (typeof y === 'number' && y >= 2000) {
      var b = beds[i][0];
      bedsByYear.push({ year: y, beds: (typeof b === 'number' && isFinite(b)) ? b : 0 });
    }
  }

  var data = {
    familiesServed: num('B5'),
    bedsDelivered:  num('C5'),
    booksGiven:     num('D5'),
    bearsGiven:     num('E5'),
    yearsServing:   num('F5'),
    bedTypes: {
      twins:   num('B10'),
      bunks:   num('C10'),
      toddler: num('D10'),
      cribs:   num('E10')
    },
    bedsByYear: bedsByYear,
    updated: new Date().toISOString()
  };

  var json = JSON.stringify(data);
  cache.put('bed_ministry_dashboard_json', json, 3600); // 3600 sec = 1 hour
  return json;
}

/**
 * Run this once from the editor to confirm the numbers look right
 * (and to trigger the one-time authorization). Check View > Logs.
 */
function bedMinistryTestDashboardJson() {
  Logger.log(bedMinistryBuildDashboardJson_());
}
