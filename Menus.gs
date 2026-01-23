/**
 * Menus.gs
 * Version: 01/20-12:05PM EST by Claude Opus 4.1
 *
 * CHANGES:
 * - Added "Go to Re-cover Calculations" menu item under Setup (Drafts)
 * - Added "Copy Ruby Code (Selected Row)" to Setup (Ruby) menu
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

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

    // System utilities menu
    ui.createMenu('⚙️ System')
      .addItem('View All Triggers', 'listAllTriggers_')
      .addItem('Remove All Triggers', 'removeAllTriggers_')
      .addToUi();

    console.log('Menus created successfully');
  } catch (error) {
    console.error('Error creating menus:', error);
  }
}

/**
 * Navigate to Re-cover sheet with the selected customer populated in K2
 * Works from Leads, F/U, or Awarded sheets
 * Version: 01/20-12:05PM EST by Claude Opus 4.1
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
    ui.alert('No triggers found');
    return;
  }
  
  const triggerInfo = triggers.map((trigger, index) => {
    return `${index + 1}. ${trigger.getHandlerFunction()} - ${trigger.getEventType()}`;
  }).join('\n');
  
  ui.alert('Active Triggers', triggerInfo, ui.ButtonSet.OK);
}

/**
 * System utility: Remove all triggers
 */
function removeAllTriggers_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Remove All Triggers', 
    'Are you sure you want to remove ALL triggers? This cannot be undone.', 
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    ui.alert('Success', `Removed ${triggers.length} triggers.`, ui.ButtonSet.OK);
  }
}