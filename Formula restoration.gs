/**
 * ============================================================================
 * FORMULA RESTORATION SYSTEM
 * ============================================================================
 * Version# 01/28-10:45AM by Claude Opus 4.5
 * 
 * PURPOSE:
 * Monitors protected formula cells and automatically restores them after 5 
 * minutes if they are manually overwritten. This allows temporary manual 
 * entry for quick adjustments while ensuring formulas are eventually restored.
 * 
 * HOW IT WORKS:
 * 1. User overwrites a protected formula cell with a manual value
 * 2. System detects the change and schedules restoration in 5 minutes
 * 3. User sees a toast notification about the pending restoration
 * 4. After 5 minutes, the original formula is automatically restored
 * 
 * ADDING NEW FORMULAS:
 * To protect a new formula, add an entry to the appropriate category in
 * FORMULA_CONFIG.FORMULAS below. Follow the existing format:
 *   'CellAddress': 'FormulaString'
 * 
 * MANUAL CONTROLS:
 * - "Restore All Formulas Now" - Immediately restores all protected formulas
 * - "View Protected Formulas" - Shows list of all protected formulas
 * 
 * IMPORTANT: This file must NOT define onOpen(). Use Menus.gs (single onOpen).
 * ============================================================================
 */

/*** =========================================================================
 * FORMULA RESTORATION CONFIGURATION
 * =========================================================================
 * 
 * FORMULAS STRUCTURE:
 * Organized by Sheet → Category → Individual Formulas
 * 
 * NAMING CONVENTION:
 * - Sheet names must match exactly (case-sensitive)
 * - Categories are for organization/documentation only
 * - Cell addresses use A1 notation (e.g., 'O2', 'B12')
 * 
 * FORMULA TYPES LEGEND:
 * - Complete: New build/full installation calculations
 * - Re-cover: Fabric replacement only calculations
 * - Pricing: Cost and pricing formulas
 * - Structural: Frame and support member calculations
 * - Fabric: Material and yardage calculations
 * ========================================================================= */

const FORMULA_CONFIG = {
  
  // -------------------------------------------------------------------------
  // GLOBAL SETTINGS
  // -------------------------------------------------------------------------
  RESTORE_DELAY_MS: 300000,  // 5 minutes = 300,000 milliseconds
  ENABLE_LOGGING: true,      // Set to false to disable console logging
  
  // -------------------------------------------------------------------------
  // PROTECTED FORMULAS BY SHEET
  // -------------------------------------------------------------------------
  // Add new sheets as needed following this structure
  
  FORMULAS: {
    
    // =======================================================================
    // SHEET: Re-cover
    // =======================================================================
    // This sheet handles calculations for both re-cover jobs and complete
    // (new build) awning installations.
    
    'Re-cover': {
      
      // ---------------------------------------------------------------------
      // CATEGORY: Structural Member Counter (Complete/New Build)
      // ---------------------------------------------------------------------
      // These formulas calculate the number of structural pipe members
      // needed for complete awning installations based on dimensions,
      // frame type (lean-to vs a-frame), and wrap style.
      //
      // Type: Complete (new build calculations)
      // Dependencies: M14 (Wrapped/Hanging), M16 (# front bars), 
      //               M18 (# wings), N2 (Length), N3 (Width), 
      //               N14 (Fabric type), N15 (Slope)
      // ---------------------------------------------------------------------
      
      // O2: Length Members (horizontal pipes running along the length)
      // Calculates: head pipe, front bar tops/bottoms, wrap pipes, stringers
      // - Lean-to (1 front bar): base of 2 + 1 bottom + wrap + stringer
      // - A-frame (2 front bars): base of 4 + 2 bottoms + wrap + stringers
      'O2': '=IF(VALUE(LEFT(M16,1))>1, 4, 2) + VALUE(LEFT(M16,1)) + IF(M14="Wrapped", VALUE(LEFT(M16,1)), 0) + IF(VALUE(LEFT(M16,1))>1, IF(N15>18, 2, 0), IF(N15>9, 1, 0))',
      
      // O3: Projection/Width Members (trusses running along the width)
      // Calculates: base trusses (fabric-dependent), wing wraps, double trusses, diagonals
      // - Sunbrella: spaced every 3.5 feet
      // - Vinyl/Other: spaced every 5 feet
      // - +1 buffer added until double truss spacing logic is finalized
      'O3': '=IF(N14="Sunbrella", ROUNDUP((N2/3.5)+2), ROUNDUP((N2/5)+2)) + IF(M14="Wrapped", VALUE(LEFT(M18,1)), 0) + ROUNDUP(N2/10) + IF(N3>10, ROUNDUP(N2/10), 0) + 1'
      
      // ---------------------------------------------------------------------
      // CATEGORY: [Future - Fabric Calculations]
      // ---------------------------------------------------------------------
      // Add fabric yardage formulas here when ready
      // Type: Re-cover or Complete
      // 
      // Example:
      // 'P10': '=ROUNDUP((N2*N3)/9, 2)'  // Basic square footage to yardage
      
      // ---------------------------------------------------------------------
      // CATEGORY: [Future - Pricing Formulas]
      // ---------------------------------------------------------------------
      // Add pricing formulas here when ready
      // Type: Pricing
      //
      // Example:
      // 'B12': '=IF(B11="Vinyl", 110, IF(B11="Sunbrella", 115, 0))'  // Fabric rate
      
    }
    
    // =======================================================================
    // SHEET: [Future Sheet Name]
    // =======================================================================
    // Copy this template when adding formulas for a new sheet:
    //
    // 'SheetName': {
    //   
    //   // CATEGORY: Category Name
    //   // Description of what these formulas do
    //   // Type: Complete | Re-cover | Pricing | Structural | Fabric
    //   
    //   'A1': '=FORMULA_HERE',
    //   'B2': '=ANOTHER_FORMULA'
    // }
    
  }
};


/* ****************************************************************************
 * ============================================================================
 * CORE FUNCTIONS - Edit Handler & Restoration Logic
 * ============================================================================
 * These functions handle the automatic detection and restoration of formulas.
 * Generally, you should not need to modify these.
 * ****************************************************************************/

/**
 * Main edit handler - detects when protected formulas are changed
 * Called automatically on every edit via installable trigger
 */
function handleEditFormula_(e) {
  if (!e || !e.source || !e.range) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  // Check if this sheet has protected formulas
  if (!FORMULA_CONFIG.FORMULAS.hasOwnProperty(sheetName)) return;
  
  const sheetFormulas = FORMULA_CONFIG.FORMULAS[sheetName];
  const r = e.range;
  
  // Ignore multi-cell edits
  if (r.getNumRows() !== 1 || r.getNumColumns() !== 1) return;
  
  const cellA1 = r.getA1Notation();
  
  // Check if this cell has a protected formula
  if (!sheetFormulas.hasOwnProperty(cellA1)) return;
  
  const currentFormula = r.getFormula();
  const expectedFormula = sheetFormulas[cellA1];
  
  // If formula matches, nothing to do
  if (currentFormula === expectedFormula) return;
  
  // Formula was changed or removed - schedule restoration
  if (FORMULA_CONFIG.ENABLE_LOGGING) {
    fr_logOperation_('Formula removed/changed', {
      sheet: sheetName,
      cell: cellA1,
      hadFormula: currentFormula ? true : false,
      newValue: String(r.getValue()).substring(0, 50)
    });
  }
  
  // Delete any existing restoration trigger for this cell
  fr_deleteRestoreTrigger_(cellA1);
  
  // Store restoration info
  const props = PropertiesService.getScriptProperties();
  const restoreKey = `FR_RESTORE_${sheetName}_${cellA1}`;
  const timeKey = `FR_TIME_${sheetName}_${cellA1}`;
  
  props.setProperty(restoreKey, expectedFormula);
  props.setProperty(timeKey, new Date().getTime().toString());
  
  // Create time-based trigger to restore after delay
  ScriptApp.newTrigger('fr_restoreScheduled_')
    .timeBased()
    .after(FORMULA_CONFIG.RESTORE_DELAY_MS)
    .create();
  
  if (FORMULA_CONFIG.ENABLE_LOGGING) {
    fr_logOperation_('Restoration scheduled', {
      sheet: sheetName,
      cell: cellA1,
      restoreIn: '5 minutes'
    });
  }
  
  // Show toast to user
  SpreadsheetApp.getActive().toast(
    `Formula in ${sheetName}!${cellA1} will be restored in 5 minutes`,
    'Formula Protection',
    5
  );
}

/**
 * Scheduled restoration handler - runs after delay to restore formulas
 * Called automatically by time-based trigger
 */
function fr_restoreScheduled_() {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  const ss = SpreadsheetApp.getActive();
  
  const now = new Date().getTime();
  let restored = 0;
  
  // Find all cells that need restoration
  for (const key in allProps) {
    if (!key.startsWith('FR_RESTORE_')) continue;
    
    // Parse the key: FR_RESTORE_SheetName_CellA1
    const keyParts = key.replace('FR_RESTORE_', '').split('_');
    const cellA1 = keyParts.pop();  // Last part is cell address
    const sheetName = keyParts.join('_');  // Remaining parts are sheet name (handles underscores in sheet names)
    
    const formula = allProps[key];
    const timeKey = `FR_TIME_${sheetName}_${cellA1}`;
    const scheduledTime = parseInt(allProps[timeKey] || '0');
    
    // Check if enough time has passed (with 10 second buffer)
    if (now - scheduledTime >= FORMULA_CONFIG.RESTORE_DELAY_MS - 10000) {
      const sheet = ss.getSheetByName(sheetName);
      
      if (!sheet) {
        if (FORMULA_CONFIG.ENABLE_LOGGING) {
          fr_logOperation_('Restoration failed - sheet not found', { sheet: sheetName, cell: cellA1 });
        }
        // Clean up properties anyway
        props.deleteProperty(key);
        props.deleteProperty(timeKey);
        continue;
      }
      
      try {
        // Check current value - only restore if still not the formula
        const cell = sheet.getRange(cellA1);
        const currentFormula = cell.getFormula();
        
        if (currentFormula !== formula) {
          cell.setFormula(formula);
          restored++;
          
          if (FORMULA_CONFIG.ENABLE_LOGGING) {
            fr_logOperation_('Formula restored', { sheet: sheetName, cell: cellA1 });
          }
        }
        
        // Clean up properties
        props.deleteProperty(key);
        props.deleteProperty(timeKey);
        
      } catch (err) {
        if (FORMULA_CONFIG.ENABLE_LOGGING) {
          fr_logOperation_('Restoration error', { sheet: sheetName, cell: cellA1, error: err.message });
        }
        // Clean up properties anyway to prevent infinite retries
        props.deleteProperty(key);
        props.deleteProperty(timeKey);
      }
    }
  }
  
  // Clean up triggers
  fr_cleanupTriggers_();
  
  if (restored > 0) {
    SpreadsheetApp.getActive().toast(
      `Restored ${restored} formula(s)`,
      'Formula Protection',
      3
    );
  }
}


/* ****************************************************************************
 * ============================================================================
 * MENU FUNCTIONS - Manual Controls
 * ============================================================================
 * These functions are called from the menu for manual formula management.
 * ****************************************************************************/

/**
 * Immediately restore all protected formulas across all sheets
 * Use this to quickly reset all formulas without waiting for the 5-minute delay
 */
function fr_restoreAllFormulasNow_() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  let totalRestored = 0;
  let totalCorrect = 0;
  let totalErrors = 0;
  const results = [];
  
  // Loop through each sheet in the config
  for (const sheetName in FORMULA_CONFIG.FORMULAS) {
    const sheet = ss.getSheetByName(sheetName);
    const sheetFormulas = FORMULA_CONFIG.FORMULAS[sheetName];
    
    if (!sheet) {
      results.push(`❌ ${sheetName}: Sheet not found`);
      totalErrors++;
      continue;
    }
    
    const sheetResults = [];
    
    // Loop through each formula in the sheet
    for (const cellA1 in sheetFormulas) {
      const formula = sheetFormulas[cellA1];
      
      try {
        const cell = sheet.getRange(cellA1);
        const currentFormula = cell.getFormula();
        
        if (currentFormula !== formula) {
          cell.setFormula(formula);
          sheetResults.push(`  ✅ ${cellA1}: Restored`);
          totalRestored++;
        } else {
          sheetResults.push(`  ✓ ${cellA1}: Already correct`);
          totalCorrect++;
        }
      } catch (err) {
        sheetResults.push(`  ❌ ${cellA1}: Error - ${err.message}`);
        totalErrors++;
      }
    }
    
    results.push(`📄 ${sheetName}:`);
    results.push(...sheetResults);
  }
  
  // Clear any pending restorations
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  for (const key in allProps) {
    if (key.startsWith('FR_RESTORE_') || key.startsWith('FR_TIME_')) {
      props.deleteProperty(key);
    }
  }
  
  // Clean up triggers
  fr_cleanupTriggers_();
  
  // Show results
  const summary = `FORMULA RESTORATION COMPLETE\n\n` +
    `✅ Restored: ${totalRestored}\n` +
    `✓ Already correct: ${totalCorrect}\n` +
    `❌ Errors: ${totalErrors}\n\n` +
    `DETAILS:\n${results.join('\n')}`;
  
  ui.alert('Formula Restoration', summary, ui.ButtonSet.OK);
  
  if (FORMULA_CONFIG.ENABLE_LOGGING) {
    fr_logOperation_('Manual restore all completed', { 
      restored: totalRestored, 
      correct: totalCorrect, 
      errors: totalErrors 
    });
  }
}

/**
 * Display all protected formulas organized by sheet and category
 * Useful for reviewing what formulas are being monitored
 */
function fr_viewProtectedFormulas_() {
  const ui = SpreadsheetApp.getUi();
  const output = [];
  
  let totalFormulas = 0;
  
  for (const sheetName in FORMULA_CONFIG.FORMULAS) {
    const sheetFormulas = FORMULA_CONFIG.FORMULAS[sheetName];
    const formulaCount = Object.keys(sheetFormulas).length;
    totalFormulas += formulaCount;
    
    output.push(`═══════════════════════════════════`);
    output.push(`📄 SHEET: ${sheetName} (${formulaCount} formulas)`);
    output.push(`═══════════════════════════════════`);
    
    for (const cellA1 in sheetFormulas) {
      const formula = sheetFormulas[cellA1];
      // Truncate long formulas for display
      const displayFormula = formula.length > 60 
        ? formula.substring(0, 60) + '...' 
        : formula;
      output.push(`\n${cellA1}:`);
      output.push(`${displayFormula}`);
    }
    
    output.push('');
  }
  
  const header = `PROTECTED FORMULAS\n` +
    `Total: ${totalFormulas} formula(s) across ${Object.keys(FORMULA_CONFIG.FORMULAS).length} sheet(s)\n` +
    `Restore delay: ${FORMULA_CONFIG.RESTORE_DELAY_MS / 60000} minutes\n\n`;
  
  ui.alert('Protected Formulas', header + output.join('\n'), ui.ButtonSet.OK);
}

/**
 * Check status of any pending formula restorations
 * Shows which formulas are scheduled to be restored and when
 */
function fr_viewPendingRestorations_() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  
  const pending = [];
  const now = new Date().getTime();
  
  for (const key in allProps) {
    if (!key.startsWith('FR_RESTORE_')) continue;
    
    const keyParts = key.replace('FR_RESTORE_', '').split('_');
    const cellA1 = keyParts.pop();
    const sheetName = keyParts.join('_');
    
    const timeKey = `FR_TIME_${sheetName}_${cellA1}`;
    const scheduledTime = parseInt(allProps[timeKey] || '0');
    const restoreTime = scheduledTime + FORMULA_CONFIG.RESTORE_DELAY_MS;
    const remainingMs = restoreTime - now;
    const remainingMin = Math.ceil(remainingMs / 60000);
    
    pending.push(`• ${sheetName}!${cellA1}: ~${remainingMin} minute(s) remaining`);
  }
  
  if (pending.length === 0) {
    ui.alert('Pending Restorations', 'No formulas are currently scheduled for restoration.', ui.ButtonSet.OK);
  } else {
    ui.alert('Pending Restorations', 
      `${pending.length} formula(s) scheduled for restoration:\n\n${pending.join('\n')}`,
      ui.ButtonSet.OK
    );
  }
}


/* ****************************************************************************
 * ============================================================================
 * HELPER FUNCTIONS
 * ============================================================================
 * Internal utility functions used by the core logic.
 * ****************************************************************************/

/**
 * Delete restoration trigger for specific cell (placeholder for future enhancement)
 */
function fr_deleteRestoreTrigger_(cellA1) {
  // Currently just cleans up old triggers
  // Future: could target specific cells if needed
  fr_cleanupTriggers_();
}

/**
 * Clean up restoration triggers when no more pending restorations exist
 */
function fr_cleanupTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  
  // Check if there are any pending restorations
  let hasPending = false;
  for (const key in allProps) {
    if (key.startsWith('FR_RESTORE_')) {
      hasPending = true;
      break;
    }
  }
  
  // If no pending restorations, remove all fr_restoreScheduled_ triggers
  if (!hasPending) {
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'fr_restoreScheduled_') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  }
}

/**
 * Logging helper - writes to console log with timestamp
 */
function fr_logOperation_(operation, details) {
  if (!FORMULA_CONFIG.ENABLE_LOGGING) return;
  try {
    console.log(`[Formula Restoration ${new Date().toISOString()}] ${operation}:`, JSON.stringify(details));
  } catch (err) {
    console.log(`[Formula Restoration] Logging error:`, err.message);
  }
}


/* ****************************************************************************
 * ============================================================================
 * TRIGGER MANAGEMENT
 * ============================================================================
 * Functions for installing and managing the edit trigger.
 * ****************************************************************************/

/**
 * Install the formula protection edit trigger
 * Call this from the menu after adding new formulas or on initial setup
 */
function installTriggerFormula_() {
  const ssId = SpreadsheetApp.getActive().getId();
  
  // Remove existing formula triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'handleEditFormula_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // Create new trigger
  ScriptApp.newTrigger('handleEditFormula_')
    .forSpreadsheet(ssId)
    .onEdit()
    .create();
  
  // Count total formulas being protected
  let totalFormulas = 0;
  let sheetCount = 0;
  for (const sheetName in FORMULA_CONFIG.FORMULAS) {
    totalFormulas += Object.keys(FORMULA_CONFIG.FORMULAS[sheetName]).length;
    sheetCount++;
  }
  
  SpreadsheetApp.getActive().toast(
    `Formula protection installed!\nMonitoring ${totalFormulas} formula(s) across ${sheetCount} sheet(s)`,
    'Setup Complete',
    5
  );
  
  if (FORMULA_CONFIG.ENABLE_LOGGING) {
    fr_logOperation_('Trigger installed', { 
      totalFormulas, 
      sheetCount,
      sheets: Object.keys(FORMULA_CONFIG.FORMULAS)
    });
  }
}

/**
 * Remove the formula protection trigger
 * Use this if you need to temporarily disable formula protection
 */
function uninstallTriggerFormula_() {
  let removed = 0;
  
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'handleEditFormula_' || 
        t.getHandlerFunction() === 'fr_restoreScheduled_') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  
  // Clear any pending restorations
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();
  for (const key in allProps) {
    if (key.startsWith('FR_RESTORE_') || key.startsWith('FR_TIME_')) {
      props.deleteProperty(key);
    }
  }
  
  SpreadsheetApp.getActive().toast(
    `Formula protection disabled.\nRemoved ${removed} trigger(s).`,
    'Protection Disabled',
    5
  );
  
  if (FORMULA_CONFIG.ENABLE_LOGGING) {
    fr_logOperation_('Trigger uninstalled', { triggersRemoved: removed });
  }
}