// ============================================================
// JobTrackerSync.gs
// Reconciles Awarded sheet rows against Liz's Job Tracker
// version 1.0 [03/31-07:45AM] by Claude Sonnet 4.6
// ============================================================

const SYNC_CONFIG = {
  AWARDED_SHEET_ID: '1H9A-5kdSAzmsxjpsuOhTmGLwLQye7XZRLcFDizpR-F4',
  AWARDED_TAB_NAME: 'Awarded',
  AWARDED_ADDRESS_COL: 31,        // AE
  AWARDED_STATUS_COL:  19,        // S
  AWARDED_DATA_START_ROW: 2,

  JOB_TRACKER_ID: '1kZ5WOpQljDEOePDjJg3DMT3E5twdRUZT-w3pSmiQkew',
  JOB_TRACKER_ADDRESS_COL: 4,    // D
  JOB_TRACKER_DATA_START_ROW: 4,

  STATUS_MATCHED: 'Tracker',
  STATUS_MISSING: 'Not in Tracker',
};

// -----------------------------------------------------------
// Helpers
// -----------------------------------------------------------

function jts_extractStreetAddress_(fullAddress) {
  if (!fullAddress) return '';
  return String(fullAddress)
    .split(',')[0]
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ');
}

function jts_fuzzyMatch_(a, b) {
  if (!a || !b) return false;
  return a.includes(b) || b.includes(a);
}

// -----------------------------------------------------------
// Main Sync
// -----------------------------------------------------------

function syncJobTracker() {
  // --- Load all addresses from Liz's Job Tracker ---
  var trackerSS    = SpreadsheetApp.openById(SYNC_CONFIG.JOB_TRACKER_ID);
  var trackerSheet = trackerSS.getSheets()[0];
  var trackerLastRow = trackerSheet.getLastRow();

  if (trackerLastRow < SYNC_CONFIG.JOB_TRACKER_DATA_START_ROW) {
    Logger.log('Job Tracker sheet appears empty -- aborting sync.');
    return;
  }

  var trackerNumRows = trackerLastRow - SYNC_CONFIG.JOB_TRACKER_DATA_START_ROW + 1;
  var trackerAddresses = trackerSheet
    .getRange(SYNC_CONFIG.JOB_TRACKER_DATA_START_ROW, SYNC_CONFIG.JOB_TRACKER_ADDRESS_COL, trackerNumRows, 1)
    .getValues()
    .flat()
    .map(jts_extractStreetAddress_)
    .filter(function(a) { return a.length > 0; });

  // --- Load Awarded rows ---
  var awardedSS    = SpreadsheetApp.openById(SYNC_CONFIG.AWARDED_SHEET_ID);
  var awardedSheet = awardedSS.getSheetByName(SYNC_CONFIG.AWARDED_TAB_NAME);
  var awardedLastRow = awardedSheet.getLastRow();

  if (awardedLastRow < SYNC_CONFIG.AWARDED_DATA_START_ROW) {
    Logger.log('Awarded sheet has no data rows -- aborting sync.');
    return;
  }

  var numRows = awardedLastRow - SYNC_CONFIG.AWARDED_DATA_START_ROW + 1;

  var statusRange  = awardedSheet.getRange(SYNC_CONFIG.AWARDED_DATA_START_ROW, SYNC_CONFIG.AWARDED_STATUS_COL,  numRows, 1);
  var addressRange = awardedSheet.getRange(SYNC_CONFIG.AWARDED_DATA_START_ROW, SYNC_CONFIG.AWARDED_ADDRESS_COL, numRows, 1);

  var statusValues  = statusRange.getValues();
  var addressValues = addressRange.getValues();

  var updates = [];

  for (var i = 0; i < numRows; i++) {
    var currentStatus = String(statusValues[i][0]).trim();

    // Already reconciled -- leave it alone
    if (currentStatus === SYNC_CONFIG.STATUS_MATCHED) {
      updates.push([currentStatus]);
      continue;
    }

    var rawAddress = addressValues[i][0];

    // No address -- skip without changing status
    if (!rawAddress || String(rawAddress).trim() === '') {
      updates.push([currentStatus]);
      continue;
    }

    var awardedStreet = jts_extractStreetAddress_(rawAddress);
    var found = trackerAddresses.some(function(ta) {
      return jts_fuzzyMatch_(awardedStreet, ta);
    });

    updates.push([found ? SYNC_CONFIG.STATUS_MATCHED : SYNC_CONFIG.STATUS_MISSING]);
  }

  // Batch write -- single API call
  statusRange.setValues(updates);
  Logger.log('syncJobTracker complete -- ' + numRows + ' rows evaluated.');
}

// -----------------------------------------------------------
// Trigger Installer
// Run once from the Apps Script editor, then wire into Menus.gs
// -----------------------------------------------------------

function installJobTrackerSyncTriggers_() {
  // Remove any existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncJobTracker') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 7:00 AM daily
  ScriptApp.newTrigger('syncJobTracker')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  // 5:00 PM daily
  ScriptApp.newTrigger('syncJobTracker')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Sync triggers installed: 7:00 AM & 5:00 PM daily.',
    'Job Tracker Sync',
    5
  );
}