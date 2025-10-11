/**
 * STAGE AUTOMATION (standalone)
 * Version: 01/10-09:15PM EST by Claude Opus 4.1
 * Moves rows between sheets, creates Drive folders, writes Google Earth links,
 * formats hyperlink cells, PRESERVES rich links/notes on row moves,
 * includes column M formula automation, expanded link generation, email linking,
 * daily empty folder checking, and AUTO-SORT by Stage column.
 *
 * IMPORTANT: This file must NOT define onOpen(). Use Menus.gs (single onOpen).
 *
 * AUTOMATED FEATURES:
 * - Display Name (col F) → Creates/finds Drive folder in Photos folder & makes F a hyperlink (Leads, F/U, Awarded)
 * - Customer Email (col I) → Formats as hyperlink to Gmail search for that email address (Leads, F/U, Awarded)
 * - Address (col J) → Creates Google Maps directions link from shop to jobsite AND triggers Earth link (col Q) (Leads, F/U, Awarded)
 * - Column M → Auto-populates job description based on Job Type (R) with dimensions and fabric details (Leads only)
 * - QB URL (col P) → Formats as "QB" hyperlink (ALL SHEETS: Leads, F/U, Awarded, Heaven, Purgatory)
 * - Earth Link (col Q) → Auto-generated when Address changes (Leads, F/U, Awarded)
 * - Stage (col D) → Moves rows between sheets based on stage value (Leads, F/U, Awarded → Purgatory/Heaven/etc.)
 * - Stage Auto-Sort → Automatically sorts by Stage A-Z when changed (Leads, F/U, Awarded, Heaven)
 * - Split CSV (col A) → Auto-splits comma-separated values across columns (Leads sheet only)
 * - Calendar Event → "2. Sched" stage creates next-day ALL-DAY event in "Appointments with Customers" calendar
 * - Column C Clear → Automatically cleared when rows move to F/U or Awarded sheets
 * - F/U Email Link (col B) → When moving to F/U, automatically searches for and links sent email in column B
 * - Empty Folder Check → Daily at 7am, checks columns F in Leads/F/U/Awarded for empty folders (red highlight)
 * - Co-exists with Draft Creator → Stages not handled by this script pass through to Draft Creator handler
 */

/*** MOVE/LINK/FOLDER AUTOMATION CONFIGURATION ***/
const MOVE_CONFIG = {
  SHEETS: {
    LEADS: 'Leads',
    FU: 'F/U',
    AWARDED: 'Awarded',
    PURG: 'Purgatory',
    RECOVER: 'Re-cover',
    HEAVEN: 'Heaven'
  },
  COLS: {
    COMMENTS: 3,   // C - Comments (cleared on move to F/U or Awarded)
    STAGE: 4,      // D - Stage
    NAME: 5,       // E - Customer Name
    DISPLAY: 6,    // F - Quote Display Name (hyperlinked to Drive folder)
    TYPE: 7,       // G - Customer type
    PHONE: 8,      // H - Customer Phone Number
    EMAIL: 9,      // I - Customer Email (hyperlinked to Gmail search)
    ADDRESS: 10,   // J - Project address (hyperlinked to Google Maps directions)
    DESC: 11,      // K - Job Description
    JOB_DESC_FORMULA: 13, // M - Auto-generated job description based on Job Type
    QUOTE: 14,     // N - Quote price
    CALCS: 15,     // O - Calculations
    QB_URL: 16,    // P - Link to QuickBooks quote (hyperlink as "QB")
    EARTH_LINK: 17,// Q - Link to Google Earth
    JOB_TYPE: 18,  // R - Type of Job
    // Awning specifications T through AC
    LENGTH: 20,    // T - Length of awning
    WIDTH: 21,     // U - Width of awning
    FRONT_BAR: 22, // V - Front bar of awning
    SHELF: 23,     // W - Shelf of awning
    WING_HEIGHT: 24,// X - Wing Height of awning
    NUM_WINGS: 25, // Y - # of Wings of awning
    VALANCE: 26,   // Z - Valance Style of awning
    FRAME: 27,     // AA - Frame type of awning
    FABRIC: 28,    // AB - Fabric of awning
    AWNING_TYPE: 29 // AC - Type of awning
  },
  PHOTOS_FOLDER_ID: '1rkg4olfmU7fw5WIYLJe7-3aUvR1cMWuK',

  // Shop address for directions (URL encoded)
  SHOP_ADDRESS_ENCODED: 'Walker+Awning,+5190+NW+10th+Terrace,+Fort+Lauderdale,+FL+33309',

  // Stage mappings for flexible matching (lowercase)
  STAGE_MAPPINGS: {
    quote_sent: ['quote sent', 'quoted', 'estimate sent', 'sent'],
    awarded:    ['awarded', 'won', 'accepted', 'approved', 'waiting on deposit', 'deposit'],
    archive:    ['archive', 'archived', 'closed'],
    lost:       ['lost', 'declined', 'rejected', 'cancelled'],
    heaven:     ['heaven', 'installed', 'complete', 'completed', 'done', 'pending review'],
    revise:     ['revise', 'revision', 'needs revision'],
    schedule:   ['2. sched', 'schedule', 'sched'],
    print_folder: ['print folder', 'print packet', 'print']
  },

  // Auto-sort configuration
  AUTO_SORT: {
    ENABLED: true,
    SHEETS: ['Leads', 'F/U', 'Awarded', 'Heaven'], // Sheets that auto-sort by Stage
    SORT_COLUMN: 4 // Column D (Stage)
  },

  // Enable/disable logging
  ENABLE_LOGGING: true,

  // Split-to-columns feature (Leads!A only)
  SPLIT: {
    ENABLED: true,
    COLUMN: 1 // Column A
  },

  // Calendar settings
  CALENDAR_NAME: 'Appointments with Customers'
};

/*** MAIN EDIT HANDLER ***/
function handleEditMove_(e) {
  if (!e || !e.source || !e.range) return;

  const S = MOVE_CONFIG;
  const sheet = e.source.getActiveSheet();
  const r = e.range;

  // Only single-cell edits, not header
  const row = r.getRow(), col = r.getColumn();
  if (row === 1 || r.getNumRows() !== 1 || r.getNumColumns() !== 1) return;

  // Allow handling on Leads, F/U, Awarded, Heaven, Purgatory
  const allowedSheets = [S.SHEETS.LEADS, S.SHEETS.FU, S.SHEETS.AWARDED, S.SHEETS.HEAVEN, S.SHEETS.PURG];
  const sheetName = sheet.getName();
  if (!allowedSheets.includes(sheetName)) return;

  const isLeads   = sheetName === S.SHEETS.LEADS;
  const isFU      = sheetName === S.SHEETS.FU;
  const isAwarded = sheetName === S.SHEETS.AWARDED;

  // Track if we should auto-sort at the end
  const shouldAutoSort = S.AUTO_SORT.ENABLED && 
                         col === S.COLS.STAGE && 
                         S.AUTO_SORT.SHEETS.includes(sheetName);

  // Split-to-columns: Leads!A only
  if (isLeads && S.SPLIT.ENABLED && col === S.SPLIT.COLUMN) {
    const val = String(r.getValue() || '').trim();
    if (val && val.indexOf(',') !== -1) {
      sheet.getRange(row, S.SPLIT.COLUMN)
           .splitTextToColumns(SpreadsheetApp.TextToColumnsDelimiter.COMMA);
      return;
    }
  }

  if (S.ENABLE_LOGGING) {
    m_logOperation_('Edit detected', {sheet: sheetName, row, col, value: r.getValue()});
  }

  // Link generation features: Now works on Leads, F/U, and Awarded
  const linkGenerationSheets = [S.SHEETS.LEADS, S.SHEETS.FU, S.SHEETS.AWARDED];
  if (linkGenerationSheets.includes(sheetName)) {
    if (col === S.COLS.DISPLAY)  { handleDisplayNameChange_(sheet, row, r.getValue()); return; }
    if (col === S.COLS.EMAIL)    { handleEmailChange_(sheet, row, r.getValue());       return; }
    if (col === S.COLS.ADDRESS)  { handleAddressChange_(sheet, row, r.getValue());     return; }
  }
  
  // Column M auto-population: Leads only
  if (isLeads && (col === S.COLS.JOB_TYPE || col === S.COLS.LENGTH || col === S.COLS.WIDTH || 
      col === S.COLS.VALANCE || col === S.COLS.FABRIC)) {
    updateJobDescription_(sheet, row);
    return;
  }

  // QB URL formatting: ALL allowed sheets (Leads, F/U, Awarded, Heaven, Purgatory)
  if (col === S.COLS.QB_URL) { handleQbUrlChange_(sheet, row, r.getValue()); return; }

  // Stage moves: allowed on Leads, F/U, and Awarded only
  // Only return if stage was actually handled by handleStageChange_
  if ((isLeads || isFU || isAwarded) && col === S.COLS.STAGE) {
    const handled = handleStageChange_(e, sheet, row, r.getValue());
    
    // Auto-sort after stage change (whether moved or not)
    if (shouldAutoSort) {
      // Small delay to ensure all operations complete
      Utilities.sleep(100);
      m_autoSortByStage_(sheet);
    }
    
    if (handled) return; // Only return if this handler processed the stage
    // If not handled, allow other handlers (like Draft Creator) to run
  }
  
  // Auto-sort if Stage column was edited but no move occurred
  if (shouldAutoSort && col === S.COLS.STAGE) {
    Utilities.sleep(100);
    m_autoSortByStage_(sheet);
  }
}

/*** FEATURE: Auto-sort sheet by Stage column (A to Z) + Apply formatting ***/
function m_autoSortByStage_(sheet) {
  const S = MOVE_CONFIG;
  
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 2) return; // Need at least 2 data rows to sort
    
    // Get the full data range (excluding header row 1)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    
    // Sort by Stage column (D = column 4) in ascending order (A to Z)
    dataRange.sort({
      column: S.COLS.STAGE,
      ascending: true
    });
    
    // Apply formatting to all data rows
    dataRange
      .setHorizontalAlignment('center')
      .setFontFamily('Roboto')
      .setFontSize(10)
      .setVerticalAlignment('middle')
      .setBackground('#ffffff'); // Reset fill color to white
    
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Auto-sorted and formatted', {
        sheet: sheet.getName(),
        rows: lastRow - 1
      });
    }
    
  } catch (err) {
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Auto-sort/format error', {
        sheet: sheet.getName(),
        error: err.message
      });
    }
    // Don't throw - sorting failure shouldn't break other functionality
  }
}

/*** FEATURE: Auto-populate Job Description in Column M based on Job Type (Leads only) ***/
function updateJobDescription_(sheet, row) {
  const S = MOVE_CONFIG;
  
  try {
    // Get values from relevant columns
    const jobType = String(sheet.getRange(row, S.COLS.JOB_TYPE).getValue() || '').trim();
    const length = sheet.getRange(row, S.COLS.LENGTH).getValue();
    const width = sheet.getRange(row, S.COLS.WIDTH).getValue();
    const valance = String(sheet.getRange(row, S.COLS.VALANCE).getValue() || '').trim();
    const fabric = String(sheet.getRange(row, S.COLS.FABRIC).getValue() || '').trim();
    
    let description = '';
    
    if (jobType.toLowerCase() === 're-cover') {
      description = 'Replace existing awning fabric.\n';
      if (length && width) {
        description += `Approximate Dimensions: ${length}' x ${width}'\n`;
      }
      if (valance) {
        description += `Valance Style: ${valance}\n`;
      }
      if (fabric) {
        description += `Awning fabric: ${fabric}\n`;
        // Add fabric-specific notes based on fabric type
        if (fabric.toLowerCase() === 'sunbrella') {
          description += '(Mayfield and Silica collection are additional)';
        } else if (fabric.toLowerCase() === 'vinyl') {
          description += '(Patio 500 or Coastline Plus)';
        }
      }
      
    } else if (jobType.toLowerCase() === 'comp') {
      description = 'Design, fabricate and Install a new awning frame and cover\n';
      if (length && width) {
        description += `Approximate Dimensions: ${length}' x ${width}'\n`;
      }
      if (fabric) {
        description += `Fabric: ${fabric}\n`;
        // Add fabric-specific notes
        if (fabric.toLowerCase() === 'sunbrella') {
          description += '(Mayfield and Silica collection are additional)\n';
        } else if (fabric.toLowerCase() === 'vinyl') {
          description += '(Patio 500 or Coastline Plus)\n';
        }
      }
      
      description += "Color: Client's choice of color\n";
      
      if (valance) {
        description += `Valance Style: ${valance}\n`;
      }
      
      description += 'Metal: Schedule 40 framework\n';
      description += "Painted frames: Client's choice of color";
    }
    
    // Set the value in column M
    const target = sheet.getRange(row, S.COLS.JOB_DESC_FORMULA);
    if (description) {
      target.setValue(description.trim());
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Job description updated', {row, jobType, description: description.substring(0, 100) + '...'});
      }
    } else {
      target.clearContent();
    }
    
  } catch (err) {
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Job description update error', {row, error: err.message});
    }
  }
}

/*** FEATURE: Create/find Drive folder and hyperlink the Display Name (F) - Now works on Leads, F/U, Awarded ***/
function handleDisplayNameChange_(sheet, row, displayName) {
  const S = MOVE_CONFIG;
  const name = String(displayName || '').trim();

  try {
    if (!name) {
      // If Display Name is empty, clear F
      sheet.getRange(row, S.COLS.DISPLAY).clearContent();
      return;
    }

    // Validate parent folder access
    let parent;
    try {
      parent = DriveApp.getFolderById(S.PHOTOS_FOLDER_ID);
    } catch (err) {
      throw new Error(`Cannot access parent folder: ${err.message}`);
    }

    // Find or create folder
    const existing = parent.getFoldersByName(name);
    const folder = existing.hasNext() ? existing.next() : parent.createFolder(name);

    // Hyperlink the Display Name cell (F) directly to the folder
    const url = folder.getUrl();
    const target = sheet.getRange(row, S.COLS.DISPLAY);
    const rich = SpreadsheetApp.newRichTextValue().setText(name).setLinkUrl(url).build();
    target.setRichTextValue(rich);

    if (S.ENABLE_LOGGING) m_logOperation_('Display name linked to folder', {name, url, row, sheet: sheet.getName()});

  } catch (err) {
    sheet.getRange(row, S.COLS.DISPLAY).setValue(`Error: ${err.message || err}`);
    SpreadsheetApp.getActive().toast(`Drive folder failed: ${err.message}`, 'Drive Error', 5);
    if (S.ENABLE_LOGGING) m_logOperation_('Folder link error', {name, error: err.message, row, sheet: sheet.getName()});
  }
}

/*** FEATURE: Format Email (I) as hyperlink to Gmail search - Now works on Leads, F/U, Awarded ***/
function handleEmailChange_(sheet, row, email) {
  const S = MOVE_CONFIG;
  const emailAddr = String(email || '').trim();
  const target = sheet.getRange(row, S.COLS.EMAIL);
  
  if (!emailAddr) {
    target.clearContent();
    return;
  }
  
  try {
    // Encode email for URL (@ becomes %40)
    const encodedEmail = encodeURIComponent(emailAddr);
    const gmailSearchUrl = `https://mail.google.com/mail/u/0/#search/${encodedEmail}`;
    
    // Create hyperlink with email as display text
    const rich = SpreadsheetApp.newRichTextValue()
      .setText(emailAddr)
      .setLinkUrl(gmailSearchUrl)
      .build();
    target.setRichTextValue(rich);
    
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Email linked to Gmail search', {email: emailAddr, url: gmailSearchUrl, row, sheet: sheet.getName()});
    }
  } catch (err) {
    target.setValue(`Error: ${err.message}`);
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Email hyperlink error', {error: err.message, row, sheet: sheet.getName()});
    }
  }
}

/*** FEATURE: Create Google Maps directions link AND Google Earth link when Address changes - Now works on Leads, F/U, Awarded ***/
function handleAddressChange_(sheet, row, address) {
  const S = MOVE_CONFIG;
  const addr = String(address || '').trim();

  if (!addr) {
    sheet.getRange(row, S.COLS.EARTH_LINK).clearContent();
    return;
  }

  // PART 1: Maps directions link in Address (J)
  try {
    const target = sheet.getRange(row, S.COLS.ADDRESS);
    const origin = S.SHOP_ADDRESS_ENCODED;
    const dest = encodeURIComponent(addr);
    const directionsUrl = `https://www.google.com/maps/dir/${origin}/${dest}`;

    const rich = SpreadsheetApp.newRichTextValue().setText(addr).setLinkUrl(directionsUrl).build();
    target.setRichTextValue(rich);

    if (S.ENABLE_LOGGING) m_logOperation_('Address linked to Maps directions', {row, address: addr, url: directionsUrl, sheet: sheet.getName()});
  } catch (err) {
    sheet.getRange(row, S.COLS.ADDRESS).setValue(`Error: ${err.message}`);
    if (S.ENABLE_LOGGING) m_logOperation_('Address hyperlink error', {row, address: addr, error: err.message, sheet: sheet.getName()});
  }

  // PART 2: Earth link in Q
  createEarthLink_(sheet, row, addr);
}

/*** FEATURE: Create Google Earth link with geocoded coordinates - Now works on Leads, F/U, Awarded ***/
function createEarthLink_(sheet, row, address) {
  const S = MOVE_CONFIG;
  const addr = String(address || '').trim();
  const target = sheet.getRange(row, S.COLS.EARTH_LINK);

  if (!addr) { target.clearContent(); return; }

  try {
    const coords = m_geocodeAddress_(addr);

    if (!coords) {
      const cleanAddr = m_cleanAddressForGoogleEarth_(addr);
      const fallbackUrl = `https://earth.google.com/web/search/${encodeURIComponent(cleanAddr)}/`;
      const rich = SpreadsheetApp.newRichTextValue().setText('Earth (No coords)').setLinkUrl(fallbackUrl).build();
      target.setRichTextValue(rich);
      if (S.ENABLE_LOGGING) m_logOperation_('Earth link fallback', {address: addr, reason: 'Geocoding failed', row, sheet: sheet.getName()});
      return;
    }

    // Neighborhood view with your preferred parameters
    const url = `https://earth.google.com/web/@${coords.lat},${coords.lng},10a,2574d,1y,0h,0t,0r`;
    const rich = SpreadsheetApp.newRichTextValue().setText('Earth').setLinkUrl(url).build();
    target.setRichTextValue(rich);

    if (S.ENABLE_LOGGING) m_logOperation_('Earth link created', {address: addr, lat: coords.lat, lng: coords.lng, url, row, sheet: sheet.getName()});

  } catch (err) {
    target.setValue(`Error: ${err.message}`);
    if (S.ENABLE_LOGGING) m_logOperation_('Earth link error', {address: addr, error: err.message, row, sheet: sheet.getName()});
  }
}

/*** FEATURE: Format QB URL (P) as "QB" hyperlink - Works on ALL sheets ***/
function handleQbUrlChange_(sheet, row, _val) {
  const S = MOVE_CONFIG;
  const url = m_getUrlFromCell_(sheet, row, S.COLS.QB_URL);
  const target = sheet.getRange(row, S.COLS.QB_URL);
  if (!url) return;

  try {
    const rich = SpreadsheetApp.newRichTextValue().setText('QB').setLinkUrl(url).build();
    target.setRichTextValue(rich);
    if (S.ENABLE_LOGGING) m_logOperation_('QB URL formatted', {url, row, sheet: sheet.getName()});
  } catch (err) {
    target.setValue(`Error: ${err.message}`);
    if (S.ENABLE_LOGGING) m_logOperation_('QB URL format error', {error: err.message, row, sheet: sheet.getName()});
  }
}

/*** FEATURE: Move rows based on Stage (edited in Leads, F/U, or Awarded) ***/
function handleStageChange_(e, sheet, row, newStage) {
  const S = MOVE_CONFIG;
  const stage = String(newStage || '').trim().toLowerCase();
  if (!stage) return false; // Not handled

  const ss = e.source;

  // Determine which sheet we're on
  const isLeads   = sheet.getName() === S.SHEETS.LEADS;
  const isFU      = sheet.getName() === S.SHEETS.FU;
  const isAwarded = sheet.getName() === S.SHEETS.AWARDED;

  // Validate all destination sheets exist
  const sheets = {
    leads:  ss.getSheetByName(S.SHEETS.LEADS),
    fu:     ss.getSheetByName(S.SHEETS.FU),
    awarded:ss.getSheetByName(S.SHEETS.AWARDED),
    purg:   ss.getSheetByName(S.SHEETS.PURG),
    heaven: ss.getSheetByName(S.SHEETS.HEAVEN)
  };
  
  if (!sheets.fu || !sheets.awarded || !sheets.purg || !sheets.heaven) {
    SpreadsheetApp.getActive().toast('Missing destination sheets!', 'Error', 5);
    return false; // Not handled due to error
  }

  // Special case: "2. Sched" in Leads creates a calendar event (no move)
  if (isLeads && m_stageMatches_(stage, S.STAGE_MAPPINGS.schedule)) {
    const success = createNextDayGinoEvent_(sheet, row);
    SpreadsheetApp.getActive().toast(
      success ? 'All-day calendar event created for tomorrow' : 'Failed to create calendar event',
      'Calendar Event',
      5
    );
    return true; // Handled - do not move the row
  }

  // Special case: "Print Folder" creates print packet (works on F/U and Awarded)
  if ((isFU || isAwarded) && m_stageMatches_(stage, S.STAGE_MAPPINGS.print_folder)) {
    const result = m_createPrintPacket_(sheet, row);
    SpreadsheetApp.getActive().toast(result.message, 'Print Packet', 5);
    return true; // Handled - do not move the row
  }

  // Special case: "Quote sent" links sent email in column B (works on all sheets)
  if (m_stageMatches_(stage, S.STAGE_MAPPINGS.quote_sent)) {
    const displayName = sheet.getRange(row, S.COLS.DISPLAY).getDisplayValue() || '';
    if (displayName) {
      m_linkQuoteSentEmail_(sheet, row, displayName);
    }
    
    // Only move to F/U from Leads or Awarded, NOT from F/U itself
    if (isFU) {
      return true; // Handled - linked email but no move
    }
    // Continue to move logic below for Leads/Awarded
  }

  // Determine destination based on stage mappings
  let dest = null, reason = '';
  
  if (m_stageMatches_(stage, S.STAGE_MAPPINGS.quote_sent)) {
    // Only move to F/U from Leads or Awarded, NOT from F/U itself
    if (isLeads || isAwarded) {
      dest = sheets.fu;      
      reason = 'Quote sent → moved to F/U';
    }
    // Already handled email linking above, so just return if in F/U
  } else if (m_stageMatches_(stage, S.STAGE_MAPPINGS.awarded)) {
    dest = sheets.awarded; 
    reason = 'Awarded → moved to Awarded';
  } else if (m_stageMatches_(stage, S.STAGE_MAPPINGS.archive)) {
    dest = sheets.purg;    
    reason = 'Archived → moved to Purgatory';
  } else if (m_stageMatches_(stage, S.STAGE_MAPPINGS.lost)) {
    dest = sheets.purg;    
    reason = 'Lost → moved to Purgatory';
  } else if (m_stageMatches_(stage, S.STAGE_MAPPINGS.heaven)) {
    dest = sheets.heaven;  
    reason = 'Pending Review/Complete → moved to Heaven';
  } else if (m_stageMatches_(stage, S.STAGE_MAPPINGS.revise)) {
    // Revise can move from F/U or Awarded back to Leads
    if (isFU || isAwarded) {
      dest = sheets.leads;   
      reason = 'Revise → moved back to Leads';
    }
    // If already in Leads, no move needed (stays in Leads)
  } else {
    return false; // No mapping → not handled, allow other handlers to run
  }

  // Only move if there's a destination
  if (dest) {
    const customerName = sheet.getRange(row, S.COLS.NAME).getDisplayValue() || 'Unknown';
    const displayName = sheet.getRange(row, S.COLS.DISPLAY).getDisplayValue() || '';
    
    const res = moveRow_(sheet, row, dest);

    if (res.success) {
      SpreadsheetApp.getActive().toast(`${customerName}: ${reason}`, 'Row Moved', 3);
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Row moved', {
          customer: customerName, stage, from: sheet.getName(), to: dest.getName(), originalRow: row
        });
      }
      
      // If moved to F/U, search for and link the sent email
      if (dest.getName() === S.SHEETS.FU && displayName) {
        m_searchAndLinkSentEmail_(dest, displayName);
      }
      
      // Auto-sort destination sheet if it's in the auto-sort list
      if (S.AUTO_SORT.ENABLED && S.AUTO_SORT.SHEETS.includes(dest.getName())) {
        Utilities.sleep(100);
        m_autoSortByStage_(dest);
      }
      
      return true; // Handled successfully
    } else {
      SpreadsheetApp.getActive().toast(`Move failed: ${res.error}`, 'Move Error', 5);
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Row move failed', {customer: customerName, stage, error: res.error, row});
      }
      return false; // Not handled due to error
    }
  }
  
  return false; // No destination, not handled
}

/*** Move a row safely with a lock, PRESERVING rich text links and notes, CLEARING column C for F/U and Awarded ***/
function moveRow_(srcSheet, srcRow, destSheet) {
  const S = MOVE_CONFIG;
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) throw new Error('Could not obtain lock (another operation in progress)');

    const lastCol = srcSheet.getLastColumn();
    if (lastCol === 0) throw new Error('Source sheet appears empty');

    // Ensure dest has enough columns
    const destLastCol = destSheet.getLastColumn();
    if (destLastCol < lastCol) destSheet.insertColumnsAfter(destLastCol, lastCol - destLastCol);

    // Preserve values & formatting
    const srcRange = srcSheet.getRange(srcRow, 1, 1, lastCol);
    const values        = srcRange.getValues();
    const rich          = srcRange.getRichTextValues();
    const notes         = srcRange.getNotes();
    const backgrounds   = srcRange.getBackgrounds();
    const fontColors    = srcRange.getFontColors();
    const fontFamilies  = srcRange.getFontFamilies();
    const fontSizes     = srcRange.getFontSizes();
    const fontWeights   = srcRange.getFontWeights();
    const fontStyles    = srcRange.getFontStyles();

    // Clear column C (comments) if moving to F/U or Awarded
    const destSheetName = destSheet.getName();
    if (destSheetName === S.SHEETS.FU || destSheetName === S.SHEETS.AWARDED) {
      values[0][S.COLS.COMMENTS - 1] = ''; // Clear column C value
      rich[0][S.COLS.COMMENTS - 1] = SpreadsheetApp.newRichTextValue().setText('').build(); // Clear rich text
      notes[0][S.COLS.COMMENTS - 1] = ''; // Clear notes
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Column C cleared on move', {destSheet: destSheetName, srcRow});
      }
    }

    destSheet.appendRow(values[0]);
    const destRow = destSheet.getLastRow();
    const destRange = destSheet.getRange(destRow, 1, 1, lastCol);

    destRange.setValues(values);
    destRange.setRichTextValues(rich);
    destRange.setNotes(notes);
    destRange.setBackgrounds(backgrounds);
    destRange.setFontColors(fontColors);
    destRange.setFontFamilies(fontFamilies);
    destRange.setFontSizes(fontSizes);
    destRange.setFontWeights(fontWeights);
    destRange.setFontStyles(fontStyles);

    srcSheet.deleteRow(srcRow);
    return { success: true };

  } catch (err) {
    return { success: false, error: err.message || String(err) };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/*** FEATURE: Search for sent email and link in column B when row is moved to F/U ***/
function m_searchAndLinkSentEmail_(sheet, displayName) {
  const S = MOVE_CONFIG;
  
  try {
    // Construct the expected email subject
    const searchSubject = `Your awning quote from Walker Awning - ${displayName}`;
    
    // Search in sent mail for this subject
    const searchQuery = `in:sent subject:"${searchSubject}"`;
    
    // Search Gmail for the most recent matching email
    const threads = GmailApp.search(searchQuery, 0, 1);
    
    if (threads.length === 0) {
      // No email found - update column B with message
      const lastRow = sheet.getLastRow();
      const logCell = sheet.getRange(lastRow, 2); // Column B
      logCell.setValue(`📧 No sent email found for: "${searchSubject}"`);
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Email search - not found', {displayName, searchSubject});
      }
      return;
    }
    
    // Get the most recent thread's first message
    const thread = threads[0];
    const messages = thread.getMessages();
    if (messages.length === 0) {
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Email search - thread empty', {displayName});
      }
      return;
    }
    
    const message = messages[0];
    const messageId = message.getId();
    
    // Construct Gmail URL to the sent email
    const gmailUrl = `https://mail.google.com/mail/u/0/#sent/${messageId}`;
    
    // Create rich text link for column B
    const linkText = '✉️ Sent: ' + searchSubject;
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(linkText)
      .setLinkUrl(0, linkText.length, gmailUrl)
      .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
      .build();
    
    // Update column B in the last row (newly moved row)
    const lastRow = sheet.getLastRow();
    const logCell = sheet.getRange(lastRow, 2); // Column B
    logCell.setRichTextValue(richText);
    
    if (S.ENABLE_LOGGING) {
      const messageDate = message.getDate();
      const formattedDate = Utilities.formatDate(messageDate, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
      m_logOperation_('Email linked in F/U column B', {
        displayName, 
        messageDate: formattedDate, 
        url: gmailUrl,
        row: lastRow
      });
    }
    
  } catch (err) {
    // If there's an error, just log it and don't fail the whole move
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Email search error', {displayName, error: err.message});
    }
    
    // Optionally put error in column B
    try {
      const lastRow = sheet.getLastRow();
      const logCell = sheet.getRange(lastRow, 2); // Column B
      logCell.setValue('Error searching for sent email: ' + (err.message || err.toString()));
    } catch (_) {
      // Ignore if we can't even write the error
    }
  }
}

/*** FEATURE: Link "Quote sent" email in column B when stage is set to "Quote sent" ***/
function m_linkQuoteSentEmail_(sheet, row, displayName) {
  const S = MOVE_CONFIG;
  
  try {
    // Clear column B first
    const logCell = sheet.getRange(row, 2); // Column B
    logCell.clearContent();
    
    // Get customer email for fallback searches
    const customerEmail = String(sheet.getRange(row, S.COLS.EMAIL).getValue() || '').trim();
    
    // PRIMARY SEARCH: "Your awning quote from Walker Awning - [displayName]"
    const primarySubject = `Your awning quote from Walker Awning - ${displayName}`;
    const primaryQuery = `in:sent subject:"${primarySubject}"`;
    let threads = GmailApp.search(primaryQuery, 0, 1);
    let searchMethod = 'primary';
    
    // SECONDARY SEARCH: If primary fails and we have email, search for "Awning Proposal" + email
    if (threads.length === 0 && customerEmail) {
      const secondaryQuery = `subject:"Awning Proposal" (to:${customerEmail} OR from:${customerEmail})`;
      threads = GmailApp.search(secondaryQuery, 0, 1);
      searchMethod = 'secondary';
    }
    
    // TERTIARY SEARCH: If both fail, find latest email from Gino to customer
    if (threads.length === 0 && customerEmail) {
      const tertiaryQuery = `from:gino@walkerawning.com to:${customerEmail}`;
      threads = GmailApp.search(tertiaryQuery, 0, 1);
      searchMethod = 'tertiary';
    }
    
    if (threads.length === 0) {
      const msg = customerEmail 
        ? `📧 No email found. Tried:\n1) "${primarySubject}"\n2) "Awning Proposal" with ${customerEmail}\n3) Any email from Gino to ${customerEmail}`
        : `📧 No email found: "${primarySubject}"`;
      logCell.setValue(msg);
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Quote sent email - not found', {displayName, customerEmail, row});
      }
      return;
    }
    
    // Get the most recent thread's first message
    const thread = threads[0];
    const messages = thread.getMessages();
    if (messages.length === 0) {
      logCell.setValue('Email thread found but contains no messages');
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Quote sent email - thread empty', {displayName, row});
      }
      return;
    }
    
    const message = messages[0];
    const messageId = message.getId();
    const emailSubject = message.getSubject();
    
    // Determine if it's in sent or inbox
    const isDraft = message.isDraft();
    const folder = isDraft ? 'drafts' : (searchMethod === 'tertiary' || searchMethod === 'secondary' ? 'all' : 'sent');
    
    // Construct Gmail URL
    const gmailUrl = `https://mail.google.com/mail/u/0/#${folder}/${messageId}`;
    
    // Create rich text link for column B with search method indicator
    let subjectText;
    if (searchMethod === 'primary') {
      subjectText = primarySubject;
    } else if (searchMethod === 'secondary') {
      subjectText = `Awning Proposal (${customerEmail})`;
    } else {
      // Tertiary - show actual email subject
      subjectText = emailSubject || `Latest email to ${customerEmail}`;
    }
    
    const linkText = '✉️ Quote Email: ' + subjectText;
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(linkText)
      .setLinkUrl(0, linkText.length, gmailUrl)
      .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
      .build();
    
    // Update column B (overwrites existing content)
    logCell.setRichTextValue(richText);
    
    if (S.ENABLE_LOGGING) {
      const messageDate = message.getDate();
      const formattedDate = Utilities.formatDate(messageDate, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
      m_logOperation_('Quote sent email linked in column B', {
        displayName,
        searchMethod,
        emailSubject,
        messageDate: formattedDate, 
        url: gmailUrl,
        row
      });
    }
    
  } catch (err) {
    logCell.setValue('Error linking quote email: ' + (err.message || err.toString()));
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Quote sent email link error', {displayName, error: err.message, row});
    }
  }
}

/*** FEATURE: Create print packet from email attachments when "Print Folder" is selected ***/
function m_createPrintPacket_(sheet, row) {
  const S = MOVE_CONFIG;
  
  // ALWAYS clear column B first to ensure clean slate
  const logCell = sheet.getRange(row, 2); // Column B
  logCell.clearContent();
  
  try {
    // Debug: Write initial message
    logCell.setValue('🔄 Creating print packet...');
    SpreadsheetApp.flush(); // Force write
    
    // Get customer data
    const customerName = String(sheet.getRange(row, S.COLS.NAME).getValue() || '').trim();
    const displayName = String(sheet.getRange(row, S.COLS.DISPLAY).getDisplayValue() || '').trim();
    const jobType = String(sheet.getRange(row, S.COLS.JOB_TYPE).getValue() || '').trim();
    
    if (!displayName) {
      logCell.setValue('❌ Error: No display name (col F) found');
      return { message: 'Error: No display name found' };
    }
    
    if (!customerName) {
      logCell.setValue('❌ Error: No customer name (col E) found');
      return { message: 'Error: No customer name found' };
    }
    
    // Search for email with just display name (no job type suffix for better match rate)
    const subject = `Proposal Review: ${displayName}`;
    
    logCell.setValue(`🔍 Searching for email: "${subject}"...`);
    SpreadsheetApp.flush();
    
    // Search for the email
    const searchQuery = `subject:"${subject}"`;
    const threads = GmailApp.search(searchQuery, 0, 1);
    
    if (threads.length === 0) {
      logCell.setValue(`❌ No email found with subject:\n"${subject}"`);
      return { message: `No email found with subject: "${subject}"` };
    }
    
    // Get attachments from the email
    const thread = threads[0];
    const messages = thread.getMessages();
    if (messages.length === 0) {
      logCell.setValue('❌ Email thread found but contains no messages');
      return { message: 'Email thread found but no messages' };
    }
    
    logCell.setValue('📎 Extracting image attachments...');
    SpreadsheetApp.flush();
    
    // Get all image attachments from all messages in the thread
    let imageAttachments = [];
    for (const message of messages) {
      const attachments = message.getAttachments();
      for (const attachment of attachments) {
        const contentType = attachment.getContentType();
        // Only include image attachments
        if (contentType && contentType.startsWith('image/')) {
          imageAttachments.push(attachment);
        }
      }
    }
    
    if (imageAttachments.length === 0) {
      logCell.setValue('❌ No image attachments found in email');
      return { message: 'No image attachments found in email' };
    }
    
    logCell.setValue(`📁 Getting Photos folder from col F...`);
    SpreadsheetApp.flush();
    
    // Get Photos folder from column F
    const photosCell = sheet.getRange(row, S.COLS.DISPLAY);
    let photosFolder = null;
    
    try {
      // Try to extract folder ID from rich text link
      const richText = photosCell.getRichTextValue();
      if (richText) {
        const folderUrl = richText.getLinkUrl();
        if (folderUrl) {
          const match = folderUrl.match(/[-\w]{25,}/);
          if (match) {
            photosFolder = DriveApp.getFolderById(match[0]);
          }
        }
      }
      
      // Fallback: try to get from formula
      if (!photosFolder) {
        const formula = photosCell.getFormula();
        if (formula) {
          const match = formula.match(/[-\w]{25,}/);
          if (match) {
            photosFolder = DriveApp.getFolderById(match[0]);
          }
        }
      }
    } catch (err) {
      logCell.setValue('❌ Error: Cannot access Photos folder from column F');
      return { message: 'Error: Cannot access Photos folder' };
    }
    
    if (!photosFolder) {
      logCell.setValue('❌ Error: No Photos folder link found in column F');
      return { message: 'Error: No Photos folder found' };
    }
    
    logCell.setValue('📄 Creating Google Doc with images...');
    SpreadsheetApp.flush();
    
    // Create subfolder with date
    const now = new Date();
    const monthYear = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/yy');
    const subfolderName = `Customer Folder - ${customerName} - ${monthYear}`;
    
    // Check if subfolder already exists
    const existingFolders = photosFolder.getFoldersByName(subfolderName);
    const subfolder = existingFolders.hasNext() 
      ? existingFolders.next() 
      : photosFolder.createFolder(subfolderName);
    
    // Create Google Doc with images
    const doc = DocumentApp.create(`Print Packet - ${displayName}`);
    const body = doc.getBody();
    
    // Add title
    body.appendParagraph(`Print Packet: ${displayName}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    body.appendParagraph(`Customer: ${customerName}`)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    if (jobType) {
      body.appendParagraph(`Job Type: ${jobType}`)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
    
    body.appendParagraph(''); // Spacer
    
    // Add each image
    for (let i = 0; i < imageAttachments.length; i++) {
      const attachment = imageAttachments[i];
      try {
        const blob = attachment.copyBlob();
        const image = body.appendImage(blob);
        
        // Scale image to fit page (6.5 inches wide for letter size with margins)
        const maxWidth = 468; // 6.5 inches * 72 points per inch
        if (image.getWidth() > maxWidth) {
          const ratio = maxWidth / image.getWidth();
          image.setWidth(maxWidth);
          image.setHeight(image.getHeight() * ratio);
        }
        
        body.appendParagraph(''); // Spacer between images
      } catch (imgErr) {
        body.appendParagraph(`Error loading image ${i + 1}: ${attachment.getName()}`);
      }
    }
    
    doc.saveAndClose();
    
    // Move doc to subfolder
    const docFile = DriveApp.getFileById(doc.getId());
    subfolder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);
    
    // Create link in column B (OVERWRITES any existing content)
    const docUrl = doc.getUrl();
    const linkText = `📄 Print Packet (${imageAttachments.length} images)`;
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(linkText)
      .setLinkUrl(0, linkText.length, docUrl)
      .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
      .build();
    
    logCell.setRichTextValue(richText); // This overwrites
    
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Print packet created', {
        customerName,
        displayName,
        imageCount: imageAttachments.length,
        subfolder: subfolderName,
        docUrl
      });
    }
    
    return { message: `✅ Print packet created with ${imageAttachments.length} images. Click link in column B.` };
    
  } catch (err) {
    logCell.setValue('❌ Error: ' + (err.message || err.toString()));
    
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Print packet error', {row, error: err.message});
    }
    
    return { message: '❌ Error: ' + (err.message || err.toString()) };
  }
}

/*** CALENDAR EVENT CREATION - Updated for "Appointments with Customers" calendar and ALL-DAY events ***/
function createNextDayGinoEvent_(sheet, row) {
  const S = MOVE_CONFIG;
  
  try {
    const name    = String(sheet.getRange(row, S.COLS.NAME).getDisplayValue() || '').trim();
    const address = String(sheet.getRange(row, S.COLS.ADDRESS).getDisplayValue() || '').trim();
    const phone   = String(sheet.getRange(row, S.COLS.PHONE).getDisplayValue() || '').trim();

    if (!name) {
      SpreadsheetApp.getActive().toast('Customer name (column E) is required', 'Missing Data', 5);
      return false;
    }

    // Find the "Appointments with Customers" calendar
    let targetCalendar;
    try {
      const calendars = CalendarApp.getAllCalendars();
      targetCalendar = calendars.find(cal => cal.getName() === S.CALENDAR_NAME);
      
      if (!targetCalendar) {
        // Fallback to default calendar if specific calendar not found
        targetCalendar = CalendarApp.getDefaultCalendar();
        SpreadsheetApp.getActive().toast(`Calendar "${S.CALENDAR_NAME}" not found, using default calendar`, 'Calendar Warning', 3);
      }
    } catch (err) {
      // If there's any issue finding calendars, use default
      targetCalendar = CalendarApp.getDefaultCalendar();
      SpreadsheetApp.getActive().toast('Using default calendar due to access issue', 'Calendar Warning', 3);
    }

    // Calculate tomorrow as an all-day event
    const now = new Date();
    const tomorrow = new Date(now);
    tomorrow.setDate(tomorrow.getDate() + 1);

    // Build event description
    let description = '';
    if (phone) description += `Phone: ${phone}\n`;
    if (address) description += `Address: ${address}`;

    // Create ALL-DAY event
    const event = targetCalendar.createAllDayEvent(
      `Gino - ${name}`,
      tomorrow,
      {
        location: address || '',
        description: description.trim()
      }
    );

    if (S.ENABLE_LOGGING) {
      m_logOperation_('All-day calendar event created', {
        calendar: targetCalendar.getName(),
        title: `Gino - ${name}`,
        date: tomorrow,
        location: address
      });
    }

    return true;
    
  } catch (err) {
    SpreadsheetApp.getActive().toast(`Calendar error: ${err.message}`, 'Calendar Error', 5);
    if (S.ENABLE_LOGGING) {
      m_logOperation_('Calendar event failed', {error: err.message, row: row});
    }
    return false;
  }
}

/*** TEST CALENDAR ACCESS FOR MOVE MENU ***/
function testCalendarAccess_() {
  // Simple calendar test for Move menu
  try {
    const calendars = CalendarApp.getAllCalendars();
    SpreadsheetApp.getActive().toast(`Found ${calendars.length} calendars`, 'Calendar OK', 3);
  } catch (err) {
    SpreadsheetApp.getActive().toast(`Calendar error: ${err.message}`, 'Error', 5);
  }
}

/*** HELPERS ***/
function m_stageMatches_(stage, options) { 
  return options ? options.some(opt => stage === opt) : false; 
}

function m_looksLikeUrl_(s) { 
  return /^https?:\/\/\S+/i.test(String(s || '').trim()); 
}

function m_getUrlFromCell_(sheet, row, col) {
  const cell = sheet.getRange(row, col);

  // HYPERLINK formula
  const formula = cell.getFormula();
  if (formula) {
    let match = formula.match(/HYPERLINK\(\s*"([^"]+)"\s*,/i);
    if (!match) match = formula.match(/HYPERLINK\(\s*([^,]+)\s*,/i);
    if (match && match[1]) {
      const raw = String(match[1]).replace(/^"/, '').replace(/"$/, '');
      if (m_looksLikeUrl_(raw)) return raw;
    }
  }

  // Rich text link
  try {
    const rt = cell.getRichTextValue();
    if (rt) {
      const link = rt.getLinkUrl();
      if (m_looksLikeUrl_(link)) return link;
    }
  } catch (_) {}

  // Plain value
  const val = String(cell.getValue() || '').trim();
  if (m_looksLikeUrl_(val)) return val;

  return '';
}

function m_cleanAddressForGoogleEarth_(address) {
  let cleaned = String(address || '').trim();
  cleaned = cleaned.replace(/#.+$/, '');
  cleaned = cleaned.replace(/[#\[\]{}]/g, '');
  cleaned = cleaned.replace(/\s+/g, ' ');
  cleaned = cleaned.replace(/^[,.\s]+|[,.\s]+$/g, '');
  cleaned = cleaned.replace(/\b(apt|apartment|unit|suite|ste)\s*[#]?\s*(\w+)/gi, '$1 $2');
  cleaned = cleaned.replace(/[;]+/g, ',');
  cleaned = cleaned.replace(/(\b\w+\b)(\s*,\s*\1)+/gi, '$1');
  if (!/\b(USA|United States|America)\b/i.test(cleaned)) {
    if (/\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY)\b/.test(cleaned)) {
      cleaned += ', USA';
    }
  }
  return cleaned.trim();
}

function m_geocodeAddress_(address) {
  try {
    const cleanAddr = m_cleanAddressForGoogleEarth_(address);
    const response = Maps.newGeocoder().geocode(cleanAddr);
    if (response.status === 'OK' && response.results && response.results.length > 0) {
      const location = response.results[0].geometry.location;
      if (location.lat < 20 || location.lat > 30 || location.lng < -85 || location.lng > -75) {
        const regionalAddr = cleanAddr.includes('FL') ? cleanAddr : cleanAddr + ', Florida, USA';
        const retryResponse = Maps.newGeocoder().geocode(regionalAddr);
        if (retryResponse.status === 'OK' && retryResponse.results && retryResponse.results.length > 0) {
          const rLoc = retryResponse.results[0].geometry.location;
          return { lat: rLoc.lat, lng: rLoc.lng };
        }
      }
      return { lat: location.lat, lng: location.lng };
    }
    return null;
  } catch (err) {
    if (MOVE_CONFIG.ENABLE_LOGGING) console.log('Geocoding error:', err.message, 'for address:', address);
    return null;
  }
}

function m_extractCoordinates_(address) {
  const coordPattern = /[-+]?([1-8]?\d(\.\d+)?|90(\.0+)?),\s*[-+]?(180(\.0+)?|((1[0-7]\d)|([1-9]?\d))(\.\d+)?)/;
  const match = String(address || '').match(coordPattern);
  if (match) {
    const parts = match[0].split(',');
    return { lat: parseFloat(parts[0].trim()), lng: parseFloat(parts[1].trim()) };
  }
  return null;
}

function m_logOperation_(operation, details) {
  if (!MOVE_CONFIG.ENABLE_LOGGING) return;
  try { console.log(`[${new Date().toISOString()}] ${operation}:`, JSON.stringify(details)); }
  catch (err) { console.log(`[${new Date().toISOString()}] Logging error:`, err.message); }
}

/*** UTILITIES (Menu items call these) ***/
function testDriveAccess_() {
  try {
    const f = DriveApp.getFolderById(MOVE_CONFIG.PHOTOS_FOLDER_ID);
    SpreadsheetApp.getActive().toast(`Drive OK: ${f.getName()}`, 'Drive Test', 3);
    return true;
  } catch (err) {
    SpreadsheetApp.getActive().toast(`Drive access failed: ${err.message}`, 'Drive Test Failed', 5);
    return false;
  }
}

function validateSheetStructure_() {
  const ss = SpreadsheetApp.getActive();
  const S = MOVE_CONFIG;
  const issues = [];

  const missing = [];
  Object.entries(S.SHEETS).forEach(([key, name]) => {
    if (!ss.getSheetByName(name)) missing.push(`${key}: "${name}"`);
  });
  if (missing.length) issues.push(`Missing sheets: ${missing.join(', ')}`);

  const leads = ss.getSheetByName(S.SHEETS.LEADS);
  if (leads) {
    const headers = leads.getRange(1, 1, 1, leads.getLastColumn()).getValues()[0];
    const expectedHeaders = {
      [S.COLS.STAGE - 1]: 'Stage',
      [S.COLS.NAME - 1]: 'Name',
      [S.COLS.DISPLAY - 1]: 'Display',
      [S.COLS.EMAIL - 1]: 'Email',
      [S.COLS.ADDRESS - 1]: 'Address',
      [S.COLS.JOB_DESC_FORMULA - 1]: 'Job Desc',
      [S.COLS.QB_URL - 1]: 'QB',
      [S.COLS.EARTH_LINK - 1]: 'Earth'
    };
    
    // Check if columns exist but skip header name validation (headers may have emojis)
    const maxCol = Math.max(...Object.keys(expectedHeaders).map(k => parseInt(k)));
    if (headers.length <= maxCol) {
      issues.push(`Not enough columns: expected at least ${maxCol + 1} columns`);
    }
  }

  try { DriveApp.getFolderById(S.PHOTOS_FOLDER_ID); }
  catch (err) { issues.push(`Cannot access Drive folder: ${err.message}`); }

  if (issues.length) {
    SpreadsheetApp.getUi().alert('Validation Issues Found', issues.join('\n\n'), SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }

  SpreadsheetApp.getActive().toast('All validation checks passed! Script ready to use.', 'Validation OK', 3);
  return true;
}

/*** TRIGGER INSTALLER (called from Menus.gs) ***/
function installTriggerMove_() {
  const ssId = SpreadsheetApp.getActive().getId();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'handleEditMove_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('handleEditMove_').forSpreadsheet(ssId).onEdit().create();
  SpreadsheetApp.getActive().toast('Stage automation trigger installed!', 'Setup Complete', 3);
  validateSheetStructure_();
}

/*** DAILY FOLDER CHECKER - Highlights empty folders in red ***/

/**
 * Check empty folders daily at 7am (time-driven trigger)
 */
function checkEmptyFoldersDaily_() {
  const S = MOVE_CONFIG;
  const ss = SpreadsheetApp.getActive();
  const sheetsToCheck = [S.SHEETS.LEADS, S.SHEETS.FU, S.SHEETS.AWARDED];
  
  let totalChecked = 0;
  let totalEmpty = 0;
  let totalErrors = 0;
  
  if (S.ENABLE_LOGGING) {
    m_logOperation_('Daily folder check started', { time: new Date().toISOString() });
  }
  
  sheetsToCheck.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const result = m_checkEmptyFoldersInSheet_(sheet);
    totalChecked += result.checked;
    totalEmpty += result.empty;
    totalErrors += result.errors;
  });
  
  if (S.ENABLE_LOGGING) {
    m_logOperation_('Daily folder check completed', { 
      totalChecked, 
      totalEmpty, 
      totalErrors 
    });
  }
}

/**
 * Check empty folders immediately (manual trigger from menu)
 */
function checkEmptyFoldersNow_() {
  const S = MOVE_CONFIG;
  const ss = SpreadsheetApp.getActive();
  const sheetsToCheck = [S.SHEETS.LEADS, S.SHEETS.FU, S.SHEETS.AWARDED];
  
  let totalChecked = 0;
  let totalEmpty = 0;
  let totalErrors = 0;
  
  sheetsToCheck.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const result = m_checkEmptyFoldersInSheet_(sheet);
    totalChecked += result.checked;
    totalEmpty += result.empty;
    totalErrors += result.errors;
  });
  
  SpreadsheetApp.getActive().toast(
    `Checked ${totalChecked} folders:\n${totalEmpty} empty (red)\n${totalErrors} errors (orange)`,
    'Folder Check Complete',
    5
  );
}

/**
 * Check all folders in a single sheet and highlight empty ones
 */
function m_checkEmptyFoldersInSheet_(sheet) {
  const S = MOVE_CONFIG;
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    return { checked: 0, empty: 0, errors: 0 };
  }
  
  let checked = 0;
  let empty = 0;
  let errors = 0;
  
  // Get all display name cells at once for efficiency
  const displayRange = sheet.getRange(2, S.COLS.DISPLAY, lastRow - 1, 1);
  const richTextValues = displayRange.getRichTextValues();
  const backgrounds = [];
  
  for (let i = 0; i < richTextValues.length; i++) {
    const row = i + 2; // Actual sheet row
    const richText = richTextValues[i][0];
    
    // Skip empty cells
    if (!richText || !richText.getText()) {
      backgrounds.push(['#ffffff']); // White background for empty
      continue;
    }
    
    // Try to extract folder URL
    const folderUrl = richText.getLinkUrl();
    
    if (!folderUrl) {
      backgrounds.push(['#ffffff']); // No link = white background
      continue;
    }
    
    // Extract folder ID from URL
    const match = folderUrl.match(/[-\w]{25,}/);
    if (!match) {
      backgrounds.push(['#ffffff']); // Invalid URL = white background
      continue;
    }
    
    const folderId = match[0];
    
    try {
      // Check if folder exists and has files
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();
      
      checked++;
      
      if (files.hasNext()) {
        // Folder has files - white background (or clear any existing color)
        backgrounds.push(['#ffffff']);
      } else {
        // Folder is empty - red background
        backgrounds.push(['#ff0000']);
        empty++;
      }
      
    } catch (err) {
      // Access error - orange background
      backgrounds.push(['#ff9900']);
      errors++;
      
      if (S.ENABLE_LOGGING) {
        m_logOperation_('Folder access error', {
          sheet: sheet.getName(),
          row,
          folderId,
          error: err.message
        });
      }
    }
  }
  
  // Apply all backgrounds at once for efficiency
  displayRange.setBackgrounds(backgrounds);
  
  return { checked, empty, errors };
}

/**
 * Install daily folder check trigger (7am)
 */
function installDailyFolderCheckTrigger_() {
  // Remove any existing daily folder check triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkEmptyFoldersDaily_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // Create new trigger for 7am daily
  ScriptApp.newTrigger('checkEmptyFoldersDaily_')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
  
  SpreadsheetApp.getActive().toast(
    'Daily folder check installed!\nRuns every day at 7:00 AM',
    'Trigger Installed',
    5
  );
  
  if (MOVE_CONFIG.ENABLE_LOGGING) {
    m_logOperation_('Daily folder check trigger installed', { time: new Date().toISOString() });
  }
}