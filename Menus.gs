/**
 * Menus.gs
 * Version: 01/26-11:52AM EST by Claude Opus 4.1
 *
 * CHANGES:
 * - Added robust trigger management system to prevent trigger loss
 * - Added master trigger handler that routes to all sub-handlers safely
 * - Added trigger health check and auto-repair functionality
 * - Added error logging to diagnose trigger issues
 * - All onEdit handlers now wrapped in try-catch for stability
 * - Preserved all existing menu items and functions
 */

// ============================================
// MASTER TRIGGER HANDLER - SINGLE POINT OF ENTRY
// ============================================

/**
 * MASTER onEdit trigger handler
 * This is the ONLY onEdit trigger that should be installed.
 * It safely routes to all sub-handlers with error protection.
 */
function masterOnEditHandler_(e) {
  // Validate event object
  if (!e || !e.source || !e.range) {
    console.log('[MasterHandler] Invalid event object - skipping');
    return;
  }
  
  const handlerResults = [];
  
  // 1. Stage Automation (move rows, create folders, format links)
  try {
    if (typeof handleEditMove_ === 'function') {
      handleEditMove_(e);
      handlerResults.push({ handler: 'Stage Automation', status: 'OK' });
    }
  } catch (err) {
    handlerResults.push({ handler: 'Stage Automation', status: 'ERROR', error: err.message });
    logTriggerError_('handleEditMove_', err, e);
  }
  
  // 2. Draft Creator (create Gmail drafts)
  try {
    if (typeof handleEditDraft_V2 === 'function') {
      handleEditDraft_V2(e);
      handlerResults.push({ handler: 'Draft Creator', status: 'OK' });
    }
  } catch (err) {
    handlerResults.push({ handler: 'Draft Creator', status: 'ERROR', error: err.message });
    logTriggerError_('handleEditDraft_V2', err, e);
  }
  
  // 3. Awning Ruby Generator (lean-to and A-frame)
  try {
    if (typeof handleEditAwningRuby_ === 'function') {
      handleEditAwningRuby_(e);
      handlerResults.push({ handler: 'Awning Ruby', status: 'OK' });
    }
  } catch (err) {
    handlerResults.push({ handler: 'Awning Ruby', status: 'ERROR', error: err.message });
    logTriggerError_('handleEditAwningRuby_', err, e);
  }
  // 5. Formula Protection (auto-restore protected formulas)
  try {
    if (typeof handleEditFormula_ === 'function') {
      handleEditFormula_(e);
      handlerResults.push({ handler: 'Formula Protection', status: 'OK' });
    }
  } catch (err) {
    handlerResults.push({ handler: 'Formula Protection', status: 'ERROR', error: err.message });
    logTriggerError_('handleEditFormula_', err, e);
  }
  // 4. Follow-up Draft Creator (if still needed - may be redundant with V2)
  try {
    if (typeof handleEditDraft_FU === 'function') {
      handleEditDraft_FU(e);
      handlerResults.push({ handler: 'Follow-up Draft', status: 'OK' });
    }
  } catch (err) {
    handlerResults.push({ handler: 'Follow-up Draft', status: 'ERROR', error: err.message });
    logTriggerError_('handleEditDraft_FU', err, e);
  }
  
  // Log summary if any errors occurred
  const errors = handlerResults.filter(r => r.status === 'ERROR');
  if (errors.length > 0) {
    console.error('[MasterHandler] Errors in handlers:', JSON.stringify(errors));
  }
}

/**
 * Log trigger errors to a hidden sheet for diagnosis
 */
function logTriggerError_(handlerName, error, event) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('_TriggerErrors');
    
    // Create log sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet('_TriggerErrors');
      logSheet.appendRow(['Timestamp', 'Handler', 'Error', 'Sheet', 'Cell', 'Value']);
      logSheet.hideSheet(); // Keep it hidden from normal view
    }
    
    // Add error entry
    const timestamp = new Date().toISOString();
    const sheetName = event && event.range ? event.range.getSheet().getName() : 'Unknown';
    const cell = event && event.range ? event.range.getA1Notation() : 'Unknown';
    const value = event && event.value !== undefined ? String(event.value).substring(0, 100) : 'N/A';
    
    logSheet.appendRow([timestamp, handlerName, error.message || String(error), sheetName, cell, value]);
    
    // Keep only last 500 errors to prevent sheet bloat
    const lastRow = logSheet.getLastRow();
    if (lastRow > 501) {
      logSheet.deleteRows(2, lastRow - 501);
    }
    
  } catch (logErr) {
    // If logging fails, at least output to console
    console.error('[TriggerErrorLog] Failed to log error:', logErr.message);
  }
}

// ============================================
// TRIGGER MANAGEMENT FUNCTIONS
// ============================================

/**
 * Install the MASTER onEdit trigger (replaces all individual triggers)
 * This is the RECOMMENDED way to set up triggers
 */
function installMasterTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  
  // Remove ALL existing onEdit triggers (clean slate)
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  
  // Install the single master handler
  ScriptApp.newTrigger('masterOnEditHandler_')
    .forSpreadsheet(ssId)
    .onEdit()
    .create();
  
  // Verify installation
  const newTriggers = ScriptApp.getProjectTriggers();
  const masterTrigger = newTriggers.find(t => t.getHandlerFunction() === 'masterOnEditHandler_');
  
  if (masterTrigger) {
    ss.toast(
      `✅ Master trigger installed!\nRemoved ${removed} old triggers.\nAll automations now run through one handler.`,
      'Trigger Setup Complete',
      5
    );
    console.log('[TriggerSetup] Master trigger installed successfully');
  } else {
    ss.toast('❌ Failed to install master trigger!', 'Error', 5);
    console.error('[TriggerSetup] Master trigger installation failed');
  }
}

/**
 * Check trigger health and report status
 */
function checkTriggerHealthMenu_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  
  let report = '=== TRIGGER HEALTH REPORT ===\n\n';
  report += `Total triggers: ${triggers.length}\n\n`;
  
  // Categorize triggers
  const onEditTriggers = [];
  const timeBasedTriggers = [];
  const otherTriggers = [];
  
  triggers.forEach(trigger => {
    const info = {
      handler: trigger.getHandlerFunction(),
      type: trigger.getEventType().toString(),
      id: trigger.getUniqueId()
    };
    
    if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT) {
      onEditTriggers.push(info);
    } else if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      timeBasedTriggers.push(info);
    } else {
      otherTriggers.push(info);
    }
  });
  
  report += `ON_EDIT Triggers (${onEditTriggers.length}):\n`;
  if (onEditTriggers.length === 0) {
    report += '  ❌ NONE - Automations will NOT work!\n';
  } else {
    onEditTriggers.forEach(t => {
      const isMaster = t.handler === 'masterOnEditHandler_';
      report += `  ${isMaster ? '✅' : '⚠️'} ${t.handler}\n`;
    });
    
    if (onEditTriggers.length > 1) {
      report += '\n  ⚠️ WARNING: Multiple onEdit triggers may cause conflicts!\n';
      report += '     Run "Install Master Trigger" to consolidate.\n';
    }
  }
  
  report += `\nTime-Based Triggers (${timeBasedTriggers.length}):\n`;
  if (timeBasedTriggers.length === 0) {
    report += '  (none)\n';
  } else {
    timeBasedTriggers.forEach(t => {
      report += `  • ${t.handler}\n`;
    });
  }
  
  // Check for expected time-based triggers
  const expectedTimeBased = ['er_processNewEmails', 'runMileageSync_', 'checkEmptyFoldersDaily_', 'emailWeeklySchedulePDF_'];
  const missingTimeBased = expectedTimeBased.filter(
    expected => !timeBasedTriggers.some(t => t.handler === expected)
  );
  
  if (missingTimeBased.length > 0) {
    report += '\n  ℹ️ Optional time-based triggers not installed:\n';
    missingTimeBased.forEach(m => {
      report += `     • ${m}\n`;
    });
  }
  
  // Check error log
  const logSheet = ss.getSheetByName('_TriggerErrors');
  if (logSheet) {
    const lastRow = logSheet.getLastRow();
    const recentErrors = lastRow > 1 ? lastRow - 1 : 0;
    report += `\nError Log: ${recentErrors} recorded errors\n`;
    
    if (recentErrors > 0) {
      // Get last 5 errors
      const startRow = Math.max(2, lastRow - 4);
      const numRows = Math.min(5, lastRow - 1);
      const errors = logSheet.getRange(startRow, 1, numRows, 4).getValues();
      
      report += 'Recent errors:\n';
      errors.reverse().forEach(row => {
        report += `  • ${row[0]}: ${row[1]} - ${row[2]}\n`;
      });
    }
  } else {
    report += '\nError Log: No errors recorded yet\n';
  }
  
  // Show report
  const ui = SpreadsheetApp.getUi();
  ui.alert('Trigger Health Check', report, ui.ButtonSet.OK);
  
  console.log(report);
}

/**
 * Auto-repair triggers if master trigger is missing
 * Can be run manually or scheduled
 */
function autoRepairTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasMasterTrigger = triggers.some(
    t => t.getHandlerFunction() === 'masterOnEditHandler_' && 
         t.getEventType() === ScriptApp.EventType.ON_EDIT
  );
  
  if (!hasMasterTrigger) {
    console.log('[AutoRepair] Master trigger missing - reinstalling...');
    installMasterTrigger_();
    return true;
  }
  
  console.log('[AutoRepair] Master trigger is healthy');
  return false;
}

/**
 * Install a daily trigger health check (runs at 5 AM)
 */
function installDailyTriggerRepair_() {
  // Remove existing auto-repair triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoRepairTriggers_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Install daily health check at 5 AM
  ScriptApp.newTrigger('autoRepairTriggers_')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Daily trigger auto-repair installed (5 AM)',
    'Auto-Repair Enabled',
    5
  );
}

/**
 * View the error log sheet
 */
function viewTriggerErrorLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('_TriggerErrors');
  
  if (!logSheet) {
    SpreadsheetApp.getUi().alert('No error log exists yet. This is good - no errors have been recorded!');
    return;
  }
  
  // Unhide and activate the sheet
  logSheet.showSheet();
  ss.setActiveSheet(logSheet);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Error log is now visible. You can hide it again from the sheet menu.',
    'Error Log',
    5
  );
}

/**
 * Clear the error log
 */
function clearTriggerErrorLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('_TriggerErrors');
  
  if (!logSheet) {
    SpreadsheetApp.getUi().alert('No error log exists.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Error Log',
    'Are you sure you want to clear all recorded trigger errors?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      logSheet.deleteRows(2, lastRow - 1);
    }
    ui.alert('Error log cleared.');
  }
}

// ============================================
// MENU CREATION
// ============================================

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

    // ====== MAIN TRIGGER MENU (Most Important!) ======
    ui.createMenu('🔧 Triggers')
      .addItem('✅ Install Master Trigger (RECOMMENDED)', 'installMasterTrigger_')
      .addItem('🔍 Check Trigger Health', 'checkTriggerHealthMenu_')
      .addItem('🔄 Auto-Repair Triggers Now', 'autoRepairTriggers_')
      .addSeparator()
      .addItem('⏰ Install Daily Auto-Repair (5 AM)', 'installDailyTriggerRepair_')
      .addSeparator()
      .addItem('📋 View Error Log', 'viewTriggerErrorLog_')
      .addItem('🗑️ Clear Error Log', 'clearTriggerErrorLog_')
      .addSeparator()
      .addItem('📊 View All Triggers', 'listAllTriggers_')
      .addItem('⚠️ Remove All Triggers', 'removeAllTriggers_')
      .addToUi();

    // Stage Automation menu
    ui.createMenu('Setup (Move)')
      .addItem('Install On-Edit Trigger (Stage)', 'installTriggerMove_')
      .addSeparator()
      .addItem('Test Drive Access', 'testDriveAccess_')
      .addItem('Test Calendar Access', 'testCalendarAccess_')
      .addItem('Validate Sheet Structure', 'validateSheetStructure_')
      .addSeparator()
      .addItem('🔍 Check Empty Folders Now', 'checkEmptyFoldersNow_')
      .addItem('⏰ Install Daily Folder Check (7am)', 'installDailyFolderCheckTrigger_')
      .addToUi();

    // Draft Creator menu
    ui.createMenu('Setup (Drafts)')
      .addItem('Install On-Edit Trigger (Drafts V2)', 'installTriggerDrafts_V2')
      .addItem('Create Drafts For All Rows (V2)', 'createDraftsForAllRows_V2')
      .addSeparator()
      .addItem('📊 Go to Re-cover Calculations', 'goToRecoverCalculations_')
      .addSeparator()
      .addItem('🗺️ Create Schedule Map Draft', 'v2_createPlotMapDraft_')
      .addToUi();

    // Mileage Log menu
    ui.createMenu('Setup (Mileage)')
      .addItem('Install Mileage Sync Trigger', 'installTriggerMileage_')
      .addItem('Run Mileage Sync Now', 'runMileageSync_')
      .addItem('Populate All Historical Events', 'populateAllMileage_')
      .addSeparator()
      .addItem('Test Distance Calculation', 'testDistanceCalculation_')
      .addItem('Clear Mileage Log', 'clearMileageLog_')
      .addToUi();

    // Lean-to Ruby generator menu
    ui.createMenu('Setup (Ruby)')
      .addItem('Install On-Edit Trigger (Lean-to Ruby)', 'installTriggerLeanToRuby_')
      .addItem('Generate Ruby for Current Row', 'testGenerateRubyCurrentRow_')
      .addItem('Generate Ruby for All Lean-to Rows', 'generateRubyForAllLeantoRows_')
      .addSeparator()
      .addItem('📋 Copy Ruby Code (Selected Row)', 'copyRubyForSelectedRow_')
      .addToUi();

    // QuickBooks menu
    ui.createMenu('Setup (QuickBooks)')
      // === SETUP (Run these first, in order) ===
      .addItem('📋 Setup Instructions', 'showQuickBooksSetup_')
      .addItem('🔧 Get Web App URL (Run First!)', 'getScriptUrl')
      .addItem('⚙️ Configure Environment', 'configureEnvironment')
      .addItem('🔗 Authorize QuickBooks', 'authorize')
      .addSeparator()
      // === DAILY USE (After authorization) ===
      .addItem('📊 Send Estimate (Current Row)', 'sendEstimateCurrentRow_')
      .addItem('💰 Convert Estimate to Invoice', 'convertEstimateToInvoice')
      .addSeparator()
      // === TROUBLESHOOTING ===
      .addItem('✅ Test Connection', 'testQuickBooksConnection_')
      .addItem('🔍 Show Configuration', 'showRedirectUri_')
      .addItem('🔄 Reset Authorization', 'resetAuth')
      .addToUi();

    // Email Reader Menu
    ui.createMenu('Email Reader')
      .addItem('🔍 Run Diagnostic Check', 'er_diagnosticCheck')
      .addSeparator()
      .addItem('📧 Process "Add lead" Emails', 'er_processAddLeadEmails')
      .addItem('▶️ Process ALL Pending Emails', 'er_processNewEmails')
      .addItem('🧪 Test Email Processing', 'er_testProcessing')
      .addSeparator()
      .addItem('⚙️ Setup Auto-Check (Ruby Only)', 'er_installTrigger')
      .addItem('🛑 Remove Auto-Check', 'er_removeTrigger')
      .addSeparator()
      .addItem('🏥 Install Daily Health Check', 'installTriggerHealthCheck_')
      .addItem('🧪 Test Health Check Now', 'testTriggerHealthCheck_')
      .addItem('🛑 Remove Health Check', 'removeTriggerHealthCheck_')
      .addToUi();

    // Job Tracker Sync menu
    ui.createMenu('📋 Job Tracker')
      .addItem('🔄 Sync Now', 'syncJobTracker')
      .addItem('⏰ Install Sync Triggers (7am & 5pm)', 'installJobTrackerSyncTriggers_')
      .addToUi();

    // Formula Protection menu
    ui.createMenu('🔢 Formulas')
      .addItem('Install Formula Protection', 'installTriggerFormula_')
      .addItem('Disable Formula Protection', 'uninstallTriggerFormula_')
      .addSeparator()
      .addItem('Restore All Formulas Now', 'fr_restoreAllFormulasNow_')
      .addItem('View Protected Formulas', 'fr_viewProtectedFormulas_')
      .addItem('View Pending Restorations', 'fr_viewPendingRestorations_')
      .addToUi();

// System utilities menu
    ui.createMenu('⚙️ System')
      .addItem('View All Triggers', 'listAllTriggers_')
      .addItem('Remove All Triggers', 'removeAllTriggers_')
      .addToUi();

    // Dialpad SMS menu
    ui.createMenu('📱 SMS')
      .addItem('Send Swatches',       'dp_sendSwatches_')
      .addItem('Request Photos',      'dp_requestPhotos_')
      .addItem('Send Proposal Link',  'dp_sendProposalLink_')
      .addItem('Custom Message',      'dp_sendCustomMessage_')
      .addToUi();

    // Weekly Schedule PDF menu
    ui.createMenu('📅 Schedule')
      .addItem('Email Weekly Schedule PDF Now', 'emailWeeklySchedulePDF_')
      .addItem('Install Monday Auto-Email Trigger', 'installWeeklyScheduleTrigger_')
      .addToUi();

    console.log('Menus created successfully');
    
    // Auto-check trigger health on open (silent)
    try {
      const triggers = ScriptApp.getProjectTriggers();
      const hasOnEdit = triggers.some(t => t.getEventType() === ScriptApp.EventType.ON_EDIT);
      if (!hasOnEdit) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          '⚠️ No onEdit triggers installed!\nGo to 🔧 Triggers → Install Master Trigger',
          'Warning',
          10
        );
      }
    } catch (checkErr) {
      // Silently ignore check errors
    }
    
  } catch (error) {
    console.error('Error creating menus:', error);
  }
}

// ============================================
// EXISTING UTILITY FUNCTIONS (preserved)
// ============================================

/**
 * Navigate to Re-cover sheet with the selected customer populated in K2
 * Works from Leads, F/U, or Awarded sheets
 */
function goToRecoverCalculations_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();
  
  // Only allow from Leads, F/U, or Awarded
  const allowedSheets = ['Leads', 'F/U', 'Awarded'];
  if (!allowedSheets.includes(sheetName)) {
    ui.alert('Wrong Sheet', 
      'Please select a row in Leads, F/U, or Awarded sheet first.', 
      ui.ButtonSet.OK);
    return;
  }
  
  const row = activeSheet.getActiveCell().getRow();
  
  if (row === 1) {
    ui.alert('Header Row', 
      'Please select a data row, not the header.', 
      ui.ButtonSet.OK);
    return;
  }
  
  // Get the Display Name from column F (index 6)
  const displayName = activeSheet.getRange(row, 6).getDisplayValue();
  
  if (!displayName) {
    ui.alert('No Display Name', 
      'This row has no Display Name in column F.\n\nPlease enter a Display Name first.', 
      ui.ButtonSet.OK);
    return;
  }
  
  // Get the Re-cover sheet
  const recoverSheet = ss.getSheetByName('Re-cover');
  
  if (!recoverSheet) {
    ui.alert('Sheet Not Found', 
      'Could not find the "Re-cover" sheet.', 
      ui.ButtonSet.OK);
    return;
  }
  
  // Set the Display Name in K2 (the selector cell)
  recoverSheet.getRange('K2').setValue(displayName);
  
  // Activate the Re-cover sheet and select K2 so user sees the selection
  ss.setActiveSheet(recoverSheet);
  recoverSheet.setActiveRange(recoverSheet.getRange('K2'));
  
  // Brief pause to let formulas recalculate
  SpreadsheetApp.flush();
  
  // Toast confirmation
  SpreadsheetApp.getActive().toast(
    `Loaded: ${displayName}\n\nRe-cover calculations now showing for this customer.`,
    '📊 Re-cover Sheet',
    5
  );
}

/**
 * Show QuickBooks setup instructions
 */
function showQuickBooksSetup_() {
  const ui = SpreadsheetApp.getUi();
  
  const instructions = `
📋 QUICKBOOKS SETUP INSTRUCTIONS
================================

STEP 1: Deploy as Web App
-------------------------
1. In Script Editor: Deploy → New Deployment
2. Select type: Web App
3. Execute as: Me
4. Who has access: Anyone
5. Click "Deploy" and copy the Web App URL

STEP 2: Get Redirect URI
-------------------------
1. In the menu: Setup (QuickBooks) → Get Web App URL
2. Copy the URL shown in the dialog
3. This is your QBO_REDIRECT_URI

STEP 3: Configure Script Properties
------------------------------------
In Script Editor: Project Settings → Script Properties
Add these properties:
- QBO_CLIENT_ID: From your Intuit Developer app
- QBO_CLIENT_SECRET: From your Intuit Developer app
- QBO_REDIRECT_URI: From Step 2
- QBO_ENVIRONMENT: "production" or "sandbox"
- QBO_REALM_ID: (set automatically during auth)

STEP 4: Configure Intuit App
-----------------------------
In Intuit Developer Dashboard:
1. Go to your app settings
2. Add the Redirect URI from Step 2 to "Redirect URIs"
3. MUST match EXACTLY (including trailing /exec)
4. Save changes

STEP 5: Authorize
-----------------
1. Setup (QuickBooks) → Authorize QuickBooks
2. Click "Authorize QuickBooks" button
3. Sign in and select your company
4. Click "Connect"

STEP 6: Test
------------
Setup (QuickBooks) → Test Connection

TROUBLESHOOTING
---------------
If you get "undefined didn't connect":
✓ Check Redirect URI matches in ALL 3 places:
  - Script Properties
  - Intuit app settings
  - Web app deployment URL
✓ Ensure web app is deployed
✓ Verify Client ID and Secret are correct
✓ Check that app is not in development mode
`;
  
  ui.alert('QuickBooks Setup Guide', instructions, ui.ButtonSet.OK);
}

/**
 * System utility: List all triggers
 */
function listAllTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  const ui = SpreadsheetApp.getUi();
  
  if (triggers.length === 0) {
    ui.alert('No triggers found.\n\nGo to 🔧 Triggers → Install Master Trigger to set up automations.');
    return;
  }
  
  const triggerInfo = triggers.map((trigger, index) => {
    const eventType = trigger.getEventType().toString();
    const handler = trigger.getHandlerFunction();
    return `${index + 1}. ${handler}\n   Type: ${eventType}`;
  }).join('\n\n');
  
  ui.alert('Active Triggers', triggerInfo, ui.ButtonSet.OK);
}

/**
 * System utility: Remove all triggers
 */
function removeAllTriggers_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Remove All Triggers', 
    '⚠️ WARNING: This will remove ALL triggers!\n\nAll automations will STOP working until you reinstall triggers.\n\nAre you sure?', 
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    ui.alert('Success', `Removed ${triggers.length} triggers.\n\nRemember to reinstall triggers when ready:\n🔧 Triggers → Install Master Trigger`, ui.ButtonSet.OK);
  }
}
// ─── Weekly Schedule PDF Emailer ────────────────────────────────────────────
// version1 [03/09-4:10PM] by Claude Sonnet 4.6

function emailWeeklySchedulePDF_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Employee Weekly Schedule');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "Employee Weekly Schedule" not found.');
    return;
  }

  var ssId    = ss.getId();
  var sheetId = sheet.getSheetId();
  var url     = 'https://docs.google.com/spreadsheets/d/' + ssId +
                '/export?format=pdf' +
                '&gid=' + sheetId +
                '&portrait=true' +
                '&fitw=true' +
                '&gridlines=false' +
                '&printtitle=false' +
                '&sheetnames=false' +
                '&pagenum=UNDEFINED' +
                '&fzr=false';

  var token    = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  var dateStr  = Utilities.formatDate(new Date(), 'America/New_York', 'MMMM d, yyyy');
  var pdfBlob  = response.getBlob().setName(
    'Employee_Weekly_Schedule_' + 
    Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd') + '.pdf'
  );

  GmailApp.sendEmail(
    'Gino@WalkerAwning.com',
    'Employee Weekly Schedule — ' + dateStr,
    'Hi Gino,\n\nAttached is the Employee Weekly Schedule for the week of ' + dateStr + '.\n\nWalker Awning Automation',
    { attachments: [pdfBlob] }
  );

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Schedule PDF emailed to Gino@WalkerAwning.com', 'Email Sent', 5);
  Logger.log('Weekly schedule PDF emailed successfully.');
}

function installWeeklyScheduleTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'emailWeeklySchedulePDF_') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('emailWeeklySchedulePDF_')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .nearMinute(0)
    .inTimezone('America/New_York')
    .create();

  SpreadsheetApp.getUi().alert('✅ Monday 8:00 AM trigger installed.\nYou will receive the PDF automatically every Monday.');
}