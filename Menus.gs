/**
 * Menus.gs
 * Version: 1/16 9am EST by Claude Sonnet 4.5
 *
 * CHANGES:
 * - Updated Email Reader menu with separate options for Ruby vs Add lead
 * - Removed Gemini Leads menu (consolidated into Email Reader)
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

    // Email Reader Menu - UPDATED
    ui.createMenu('Email Reader')
      .addItem('🔍 Run Diagnostic Check', 'er_diagnosticCheck')
      .addSeparator()
      .addItem('📧 Process "Add lead" Emails', 'er_processAddLeadEmails')
      .addItem('▶️ Process ALL Pending Emails', 'er_processNewEmails')
      .addItem('🧪 Test Email Processing', 'er_testProcessing')
      .addSeparator()
      .addItem('⚙️ Setup Auto-Check (Ruby Only)', 'er_installTrigger')
      .addItem('🛑 Remove Auto-Check', 'er_removeTrigger')
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