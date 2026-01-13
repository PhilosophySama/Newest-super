/**
 * Draft Creator.gs
 * version# 01/05-10:15PM EST by Claude Opus 4.1
 *
 * PURPOSE
 * - Create Gmail drafts when Stage (col D) becomes TARGET_STAGE ("qDraft") on allowed sheets.
 * - Search and link existing emails when Stage (col D) becomes "Liz" on allowed sheets.
 * - Create customer follow-up drafts when Stage (col D) becomes "Email customer" on allowed sheets.
 * - Create customer handoff drafts when Stage (col D) becomes "Cust Handoff" on allowed sheets.
 * - Create rough quote drafts/messages when Stage (col D) becomes "Rough quote" on allowed sheets.
 * - Create customer info request drafts when Stage (col D) becomes "Customer Info" on Leads, F/U, and Awarded sheets.
 * - Create COI request drafts when Stage (col D) becomes "COI Req" on F/U, Awarded, and Heaven sheets only.
 * - Under PICS, embed a Re-cover!A1:K14 snapshot (HTML only).
 *
 * CHANGES IN THIS VERSION:
 * - Removed 5-second delay - drafts now create BEFORE row moves
 * - "Email customer" now uses Job Type (R) instead of Job Description (K)
 * - "Email customer" signature moved after QuickBooks button
 *
 * NOTES
 * - HTML export requires enabling "Google Sheets API" in Advanced Google Services + GCP Console.
 * - Column F is both Display Name + PICS link (hyperlink supported).
 */

const DRAFTS_V2 = {
  SPREADSHEET_ID: 'REPLACE_ME_WITH_YOUR_SHEET_ID',

  SHEETS: { LEADS: 'Leads', RECOVER: 'Re-cover' },

  // ========================================
  // STAGE TRIGGERS (Change these to customize)
  // ========================================
  TARGET_STAGE: 'qDraft',              // Main draft creation trigger
  LIZ_STAGE: 'Liz',  // Stage value for email search
  REVISE_STAGE: 'Revise',  // Stage value for email search + move to Leads
  CUSTOMER_STAGE: 'Email customer',    // Customer follow-up trigger
  HANDOFF_STAGE: 'Cust Handoff',       // Customer handoff trigger
  ROUGH_QUOTE_STAGE: 'Rough quote',    // Rough quote trigger
  CUSTOMER_INFO_STAGE: 'Customer Info', // Info request trigger
  COI_STAGE: 'COI Req',                // COI request trigger

  COLS: {
    LOG_B: 'B',
    STAGE: 'D',

    CUSTOMER_NAME: 'E',
    DISPLAY_NAME:  'F',   // Display name; also holds the PICS hyperlink
    FOLDER_URL:    'F',   // PICS folder URL lives here

    CUSTOMER_TYPE: 'G',
    PHONE:         'H',
    EMAIL:         'I',
    ADDRESS:       'J',
    JOB_DESC:      'K',

    QUOTE_PRICE:   'N',
    CALCS:         'O',
    QB_URL:        'P',
    GE_URL:        'Q',
    JOB_TYPE:      'R',

    LEN:           'T',
    WIDTH:         'U',
    FRONT_BAR:     'V',
    SHELF:         'W',
    WING_HEIGHT:   'X',
    NUM_WINGS:     'Y',
    VALANCE_STYLE: 'Z',
    FRAME_TYPE:    'AA',
    FABRIC:        'AB',
    AWNING_TYPE:   'AC',
  },

  // Google Drive File IDs for Customer Info attachments
  CUSTOMER_INFO_ATTACHMENTS: {
    SUNBRELLA_FILE_ID: '1_SUNBRELLA_PLACEHOLDER_',
    VINYL_FERRARI_FILE_ID: '1_FERRARI_PLACEHOLDER_',
    VINYL_PATIO500_FILE_ID: '1_PATIO500_PLACEHOLDER_',
    VINYL_COASTLINE_FILE_ID: '1_COASTLINE_PLACEHOLDER_'
  },

  RECOVER: {
    SELECT_CELL_A1:    'K2',
    SNAPSHOT_RANGE_A1: 'A1:K14',
    WAIT_MS:           2000,              // give calculations time to settle
    DEBUG:             true
  },

  EMAIL: {
    TO: ['Liz@WalkerAwning.com'],
    CC: [],
    BCC: [],
    SUBJECT_PREFIX: 'Proposal Review',
    LINK_LABELS: { PHOTOS:'PICS', EARTH:'Google Earth', QUICKBOOKS:'Quickbooks', ROUTE_MAP:'Route Map' },
    SUBJECT_TEMPLATE: '${prefix}: ${displayName} - ${jobType}',
    SUBJECT_MAX_LENGTH: 120,
    SKIP_IF_DRAFT_EXISTS: true
  },

  EMAIL_SEARCH: {
    MAX_RESULTS: 10,  // Max emails to search through
    SEARCH_DAYS: 90   // Search emails from last N days
  },

  CUSTOMER_EMAIL: {
    SUBJECT_TEMPLATE: 'Your awning quote from Walker Awning - ${displayName}',
    BODY_TEMPLATE: `Hello \${firstName},

Thank you for considering Walker Awning! Linked is your custom quote for the awning project we discussed.

Project Details:
- Location: \${address}
- Description: \${jobType}

If you have any questions or would like to proceed with this project, please reach out via text or call.`,
    HTML_BODY_TEMPLATE: `<div style="font-family: Arial, sans-serif; color: #333;">
<p>Hello \${firstName},</p>

<p>Thank you for considering Walker Awning! Linked is your custom quote for the awning project we discussed.</p>

<p><strong>Project Details:</strong></p>
<ul style="line-height: 1.8;">
  <li><strong>Location:</strong> \${address}</li>
  <li><strong>Description:</strong> \${jobType}</li>
</ul>

<p>If you have any questions or would like to proceed with this project, please reach out via text or call.</p>
</div>`
  },

  HANDOFF_EMAIL: {
    SUBJECT_TEMPLATE: 'Re: Your Walker Awning Project - ${displayName}',
    BODY_TEMPLATE: `Hello \${firstName},

Michael is no longer with our team, and I'll be taking over your project. To help me get up to speed, could you please send me any photos, rough dimensions, and other key details about your situation?

I wasn't able to find Mike's notes on this (haha), so your input will be a huge help.

Thanks,`,
    HTML_BODY_TEMPLATE: `<div style="font-family: Arial, sans-serif; color: #333;">
<p>Hello \${firstName},</p>

<p>Michael is no longer with our team, and I'll be taking over your project. To help me get up to speed, could you please send me any photos, rough dimensions, and other key details about your situation?</p>

<p>I wasn't able to find Mike's notes on this (haha), so your input will be a huge help.</p>

<p>Thanks,</p>
</div>`
  },

  ROUGH_QUOTE_EMAIL: {
    SUBJECT_TEMPLATE: 'Quick quote for your awning - ${displayName}',
    MIN_WIDTH_RECOVER: 5,  // Minimum width for re-cover calculations
    SIGNATURE: '\n\nRegards,\nGino\nWalker Awning'
  },

  COI_REQUEST: {
    ATTACHMENT_FILE_ID: '1G1yVE4Ys7JI03h8QA20ONHM6Y3nDPhgY',  // W9 2025.pdf from Google Drive
    RECIPIENTS: [
      'ccardozo@keyescoverage.com',
      'ealvarado@keyescoverage.com',
      'jsirias@keyescoverage.com',
      'gonzalo@keyescoverage.com'
    ]
  },

  // Google Maps Static API configuration
  MAPS_CONFIG: {
    SHOP_ADDRESS: '5190 NW 10th Terrace, Fort Lauderdale, FL 33309',
    MAP_WIDTH: 640,
    MAP_HEIGHT: 400,
    MAP_TYPE: 'roadmap',
    ROUTE_COLOR: '0x4285F4',
    ROUTE_WEIGHT: 6,
    MARKER_ORIGIN_COLOR: 'green',
    MARKER_DEST_COLOR: 'red'
  },

  RETRY: { MAX_ATTEMPTS: 3, DELAYS_MS: [5000, 15000, 30000] }
};

/**
 * Helper function to find Google Drive file IDs by exact file names
 * Run this ONCE to get the file IDs for CUSTOMER_INFO_ATTACHMENTS
 */
function findCustomerInfoAttachmentFileIds() {
  const fileNames = {
    SUNBRELLA: '2025 Sunbrella Colors.pdf',
    VINYL_FERRARI: '2025 - Vinyl - Ferrari.jpg',
    VINYL_PATIO500: '2025 - Vinyl - Patio 500.jpg',
    VINYL_COASTLINE: '2025 - Vinyl - Coastline Plus.jpg'
  };
  
  const results = {};
  
  for (const [key, fileName] of Object.entries(fileNames)) {
    try {
      const files = DriveApp.getFilesByName(fileName);
      if (files.hasNext()) {
        const file = files.next();
        results[key] = file.getId();
        Logger.log(`${key}: ${file.getId()} (${fileName})`);
      } else {
        Logger.log(`${key}: FILE NOT FOUND - ${fileName}`);
        results[key] = 'FILE_NOT_FOUND';
      }
    } catch (err) {
      Logger.log(`${key}: ERROR - ${err.message}`);
      results[key] = 'ERROR';
    }
  }
  
  Logger.log('\n\n=== COPY THIS INTO YOUR DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS CONFIG ===');
  Logger.log(`  SUNBRELLA_FILE_ID: '${results.SUNBRELLA}',`);
  Logger.log(`  VINYL_FERRARI_FILE_ID: '${results.VINYL_FERRARI}',`);
  Logger.log(`  VINYL_PATIO500_FILE_ID: '${results.VINYL_PATIO500}',`);
  Logger.log(`  VINYL_COASTLINE_FILE_ID: '${results.VINYL_COASTLINE}'`);
  
  return results;
}

/** Install onEdit trigger (clean re-install). */
function installTriggerDrafts_V2() {
  v2_validateConfig_();
  const handler = 'handleEditDraft_V2';
  
  // Remove existing triggers
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler)
    .forEach(t => ScriptApp.deleteTrigger(t));
  
  // Install onEdit trigger
  ScriptApp.newTrigger(handler)
    .forSpreadsheet(v2_getSpreadsheet_().getId())
    .onEdit()
    .create();
  
  SpreadsheetApp.getActive().toast('Drafts V2 trigger installed', 'Draft Creator', 4);
}

/** onEdit (installable) — handles all stage-based draft creation */
function handleEditDraft_V2(e) {
  try {
    if (!e || !e.source || !e.range) return;
    const sh = e.range.getSheet();
    const sheetName = sh.getName();
    
    // Allow Leads, F/U, Awarded, and Heaven sheets
    if (sheetName !== 'Leads' && sheetName !== 'F/U' && sheetName !== 'Awarded' && sheetName !== 'Heaven') return;

    if (e.range.getRow() === 1) return;
    if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

    const stageCol = d_colLetterToIndex_(DRAFTS_V2.COLS.STAGE);
    if (e.range.getColumn() !== stageCol) return;

    const newVal = String((e.value != null ? e.value : e.range.getValue()) || '').trim();
    const newValLower = newVal.toLowerCase();
    const row = e.range.getRow();
    
    // Handle "Liz" stage - search for existing email
    if (newValLower === String(DRAFTS_V2.LIZ_STAGE).toLowerCase()) {
      const result = v2_searchAndLinkEmail_(sh, row);
      SpreadsheetApp.getActive().toast(result.message, 'Email Search', 5);
      return;
    }

    // Handle "Revise" stage - search for email
    if (newValLower === String(DRAFTS_V2.REVISE_STAGE).toLowerCase()) {
      const result = v2_searchAndLinkEmail_(sh, row);
      SpreadsheetApp.getActive().toast(result.message, 'Email Search', 5);
      return;
    }
    
    // Handle "Email customer" stage - create customer follow-up draft
    if (newValLower === String(DRAFTS_V2.CUSTOMER_STAGE).toLowerCase()) {
      const result = v2_createCustomerDraft_(sh, row);
      SpreadsheetApp.getActive().toast(result.toast, 'Customer Draft', 5);
      return;
    }
    
    // Handle "Cust Handoff" stage - create customer handoff draft
    if (newValLower === String(DRAFTS_V2.HANDOFF_STAGE).toLowerCase()) {
      const result = v2_createHandoffDraft_(sh, row);
      SpreadsheetApp.getActive().toast(result.toast, 'Handoff Draft', 5);
      return;
    }
    
    // Handle "Rough quote" stage - create rough quote draft/message
    if (newValLower === String(DRAFTS_V2.ROUGH_QUOTE_STAGE).toLowerCase()) {
      const result = v2_createRoughQuote_(sh, row);
      SpreadsheetApp.getActive().toast(result.toast, 'Rough Quote', 5);
      return;
    }
    
    // Handle "Customer Info" stage - create customer info request draft/message
    if (newValLower === String(DRAFTS_V2.CUSTOMER_INFO_STAGE).toLowerCase()) {
      const result = v2_createCustomerInfoDraft_(sh, row);
      SpreadsheetApp.getActive().toast(result.toast, 'Customer Info', 5);
      return;
    }
    
    // Handle "COI Req" stage - create COI request draft
    if (newValLower === String(DRAFTS_V2.COI_STAGE).toLowerCase()) {
      const result = v2_createCOIDraft_(sh, row);
      SpreadsheetApp.getActive().toast(result.toast, 'COI Request', 5);
      return;
    }
    
    // Handle "qDraft" stage - create main draft
    if (newValLower === String(DRAFTS_V2.TARGET_STAGE).toLowerCase()) {
      const result = v2_createDraftForRow_(sh, row, DRAFTS_V2.EMAIL.SKIP_IF_DRAFT_EXISTS);
      SpreadsheetApp.getActive().toast(result.toast, 'Draft Creator', 5);
    }
  } catch (err) {
    console.error('Handler error:', err);
    SpreadsheetApp.getActive().toast('Error: ' + d_shortErr_(err), 'Draft Creator', 8);
  }
}

/**
 * Search Gmail for existing emails/drafts and link in column B
 */
function v2_searchAndLinkEmail_(sh, row) {
  try {
    const lastCol = sh.getLastColumn();
    const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const idx = (L) => d_colLetterToIndex_(L) - 1;
    
    // Get display name for search
    const displayName = d_safeString_(vals[idx(DRAFTS_V2.COLS.DISPLAY_NAME)]);
    if (!displayName) {
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue('No display name found in column F for email search');
      return { message: 'No display name found for search' };
    }
    
    // Search for emails/drafts with display name
    const searchQuery = `subject:"${displayName}" newer_than:${DRAFTS_V2.EMAIL_SEARCH.SEARCH_DAYS}d`;
    
    // First check drafts
    const drafts = GmailApp.search(`in:drafts ${searchQuery}`, 0, DRAFTS_V2.EMAIL_SEARCH.MAX_RESULTS);
    if (drafts.length > 0) {
      // Found a draft - use the most recent one
      const thread = drafts[0];
      const messages = thread.getMessages();
      const message = messages[messages.length - 1];
      const messageId = message.getId();
      const gmailUrl = `https://mail.google.com/mail/u/0/#drafts?compose=${encodeURIComponent(messageId)}`;
      
      // Create link in column B
      const linkText = '📝 Draft ';
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(linkText)
        .setLinkUrl(0, linkText.length, gmailUrl)
        .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
        .build();
      
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setRichTextValue(richText);
      
      const messageDate = message.getDate();
      const formattedDate = Utilities.formatDate(messageDate, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
      return { message: `Found and linked draft from ${formattedDate}` };
    }
    
    // No draft found, search sent emails
    const threads = GmailApp.search(`in:sent ${searchQuery}`, 0, DRAFTS_V2.EMAIL_SEARCH.MAX_RESULTS);
    if (threads.length === 0) {
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(`No emails or drafts found for "${displayName}" in last ${DRAFTS_V2.EMAIL_SEARCH.SEARCH_DAYS} days`);
      return { message: 'No emails found' };
    }
    
    // Found an email - use the most recent one
    const thread = threads[0];
    const messages = thread.getMessages();
    const message = messages[messages.length - 1];
    const messageId = message.getId();
    const gmailUrl = `https://mail.google.com/mail/u/0/#inbox/${thread.getId()}`;
    
    // Determine if it's a draft or sent email
    const isDraft = message.isDraft();
    
    // Create link in column B
    const linkText = isDraft ? '📝 Draft ' : '✉️ Liz ';
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(linkText)
      .setLinkUrl(0, linkText.length, gmailUrl)
      .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
      .build();
    
    // Update column B (overwrites existing content)
    const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    logCell.setRichTextValue(richText);
    
    // Return success message with thread details
    const messageDate = message.getDate();
    const formattedDate = Utilities.formatDate(messageDate, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm');
    return { 
      message: `Found and linked ${isDraft ? 'draft' : 'email'} from ${formattedDate}` 
    };
    
  } catch (err) {
    const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    logCell.setValue('Error searching email: ' + d_shortErr_(err));
    return { message: 'Error searching for email: ' + d_shortErr_(err) };
  }
}

/**
 * Create customer follow-up draft when "Email customer" is selected
 */
function v2_createCustomerDraft_(sh, row) {
  try {
    const lastCol = sh.getLastColumn();
    const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const rtv = sh.getRange(row, 1, 1, lastCol).getRichTextValues()[0];
    const idx = (L) => d_colLetterToIndex_(L) - 1;
    
    // Get customer data
    const customerEmail = d_safeString_(vals[idx(DRAFTS_V2.COLS.EMAIL)]);
    const displayName = d_safeString_(vals[idx(DRAFTS_V2.COLS.DISPLAY_NAME)]) || 'Unnamed Lead';
    const customerName = d_safeString_(vals[idx(DRAFTS_V2.COLS.CUSTOMER_NAME)]);
    const firstName = customerName ? customerName.split(' ')[0] : 'there';
    const address = d_safeString_(vals[idx(DRAFTS_V2.COLS.ADDRESS)]);
    const jobType = d_safeString_(vals[idx(DRAFTS_V2.COLS.JOB_TYPE)]); // Changed from JOB_DESC to JOB_TYPE (column R)
    const qbUrl = d_extractUrlFromCell_(vals[idx(DRAFTS_V2.COLS.QB_URL)], rtv[idx(DRAFTS_V2.COLS.QB_URL)], sh, row, DRAFTS_V2.COLS.QB_URL);
    
    // Validate customer email
    if (!customerEmail || !d_isValidEmail_(customerEmail)) {
      const msg = 'No valid customer email found in column I';
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
    // Create subject
    let subject = d_templateSafe_(DRAFTS_V2.CUSTOMER_EMAIL.SUBJECT_TEMPLATE, { 
      displayName, 
      customerName 
    });
    
    // Create plain text body
    let plainBody = DRAFTS_V2.CUSTOMER_EMAIL.BODY_TEMPLATE;
    plainBody = plainBody.replace(/\$\{firstName\}/g, firstName);
    plainBody = plainBody.replace(/\$\{address\}/g, address || 'Not specified');
    plainBody = plainBody.replace(/\$\{jobType\}/g, jobType || 'Not specified'); // Changed from jobDescription
    plainBody = plainBody.replace(/\$\{displayName\}/g, displayName);
    
    // Add QuickBooks link if available, then add signature
    if (qbUrl) {
      plainBody += `\n\nView your quote online: ${qbUrl}\n\nBest regards,\nWalker Awning Team`;
    } else {
      plainBody += '\n\nBest regards,\nWalker Awning Team';
    }
    
    // Create HTML body
    let htmlBody = DRAFTS_V2.CUSTOMER_EMAIL.HTML_BODY_TEMPLATE;
    htmlBody = htmlBody.replace(/\$\{firstName\}/g, d_htmlEscape_(firstName));
    htmlBody = htmlBody.replace(/\$\{address\}/g, d_htmlEscape_(address || 'Not specified'));
    htmlBody = htmlBody.replace(/\$\{jobType\}/g, d_htmlEscape_(jobType || 'Not specified')); // Changed from jobDescription
    htmlBody = htmlBody.replace(/\$\{displayName\}/g, d_htmlEscape_(displayName));
    
    // Add QuickBooks link to HTML if available, then add signature after button
    if (qbUrl) {
      htmlBody += `<p><a href="${d_htmlEscape_(qbUrl)}" style="background-color: #4CAF50; color: white; padding: 12px 24px; text-decoration: none; display: inline-block; border-radius: 4px; font-weight: bold;">View Your Quote Online</a></p>

<p>Best regards,<br>
Walker Awning Team</p>
</div>`;
    } else {
      htmlBody += `<p>Best regards,<br>
Walker Awning Team</p>
</div>`;
    }
    
    // Create draft
    const options = {
      htmlBody: htmlBody
    };
    
    try {
      const draft = d_withRetry_(() => GmailApp.createDraft(customerEmail, subject, plainBody, options));
      const draftMessageId = draft.getMessage().getId();
      const draftUrl = 'https://mail.google.com/mail/u/0/#drafts?compose=' + encodeURIComponent(draftMessageId);
      
      // Create rich text link for column B
      const linkText = '📧 Quote';
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(linkText)
        .setLinkUrl(0, linkText.length, draftUrl)
        .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
        .build();
      
      // Update column B
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setRichTextValue(richText);
      
      return { toast: 'Customer follow-up draft created & linked in column B' };
      
    } catch (err) {
      const msg = d_specificErrorMessage_(err);
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
  } catch (err) {
    const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    logCell.setValue('Error creating customer draft: ' + d_shortErr_(err));
    return { toast: 'Error creating customer draft: ' + d_shortErr_(err) };
  }
}

/**
 * Create customer handoff draft when "Cust Handoff" is selected
 */
function v2_createHandoffDraft_(sh, row) {
  try {
    const lastCol = sh.getLastColumn();
    const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const idx = (L) => d_colLetterToIndex_(L) - 1;
    
    // Get customer data
    const customerEmail = d_safeString_(vals[idx(DRAFTS_V2.COLS.EMAIL)]);
    const displayName = d_safeString_(vals[idx(DRAFTS_V2.COLS.DISPLAY_NAME)]) || 'Unnamed Lead';
    const customerName = d_safeString_(vals[idx(DRAFTS_V2.COLS.CUSTOMER_NAME)]);
    
    // Extract first name from customer name
    const firstName = customerName ? customerName.split(' ')[0] : 'there';
    
    // Validate customer email
    if (!customerEmail || !d_isValidEmail_(customerEmail)) {
      const msg = 'No valid customer email found in column I';
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
    // Create subject
    let subject = d_templateSafe_(DRAFTS_V2.HANDOFF_EMAIL.SUBJECT_TEMPLATE, { 
      displayName
    });
    
    // Create plain text body
    let plainBody = DRAFTS_V2.HANDOFF_EMAIL.BODY_TEMPLATE;
    plainBody = plainBody.replace(/\$\{firstName\}/g, firstName);
    
    // Create HTML body
    let htmlBody = DRAFTS_V2.HANDOFF_EMAIL.HTML_BODY_TEMPLATE;
    htmlBody = htmlBody.replace(/\$\{firstName\}/g, d_htmlEscape_(firstName));
    
    // Create draft
    const options = {
      htmlBody: htmlBody
    };
    
    try {
      const draft = d_withRetry_(() => GmailApp.createDraft(customerEmail, subject, plainBody, options));
      const draftMessageId = draft.getMessage().getId();
      const draftUrl = 'https://mail.google.com/mail/u/0/#drafts?compose=' + encodeURIComponent(draftMessageId);
      
      // Create rich text link for column B
      const linkText = '🤝 Handoff Draft: ' + displayName;
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(linkText)
        .setLinkUrl(0, linkText.length, draftUrl)
        .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
        .build();
      
      // Update column B
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setRichTextValue(richText);
      
      return { toast: 'Customer handoff draft created & linked in column B' };
      
    } catch (err) {
      const msg = d_specificErrorMessage_(err);
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
  } catch (err) {
    const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    logCell.setValue('Error creating handoff draft: ' + d_shortErr_(err));
    return { toast: 'Error creating handoff draft: ' + d_shortErr_(err) };
  }
}

/**
 * Create customer info request draft when "Customer Info" is selected
 */
function v2_createCustomerInfoDraft_(sh, row) {
  try {
    // Customer Info only works on Leads, F/U, and Awarded sheets
    const sheetName = sh.getName();
    if (sheetName !== 'Leads' && sheetName !== 'F/U' && sheetName !== 'Awarded') {
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue('Customer Info Request only available on Leads, F/U, and Awarded sheets');
      return { toast: 'Customer Info Request only available on Leads, F/U, and Awarded sheets' };
    }
    
    const lastCol = sh.getLastColumn();
    const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const idx = (L) => d_colLetterToIndex_(L) - 1;
    
    // Get customer data
    const customerEmail = d_safeString_(vals[idx(DRAFTS_V2.COLS.EMAIL)]);
    const customerName = d_safeString_(vals[idx(DRAFTS_V2.COLS.CUSTOMER_NAME)]);
    const firstName = customerName ? customerName.split(' ')[0] : 'there';
    
    // Check what's missing (columns C, E, F, H, I, J, R, T, U)
    const fabricColor = d_safeString_(vals[2]); // Column C
    const name = d_safeString_(vals[idx(DRAFTS_V2.COLS.CUSTOMER_NAME)]); // E
    const displayName = d_safeString_(vals[idx(DRAFTS_V2.COLS.DISPLAY_NAME)]); // F
    const phone = d_safeString_(vals[idx(DRAFTS_V2.COLS.PHONE)]); // H
    const email = d_safeString_(vals[idx(DRAFTS_V2.COLS.EMAIL)]); // I
    const address = d_safeString_(vals[idx(DRAFTS_V2.COLS.ADDRESS)]); // J
    const jobType = d_safeString_(vals[idx(DRAFTS_V2.COLS.JOB_TYPE)]); // R
    const length = d_safeString_(vals[idx(DRAFTS_V2.COLS.LEN)]); // T
    const width = d_safeString_(vals[idx(DRAFTS_V2.COLS.WIDTH)]); // U
    
    // Get valance style (Z) and fabric (AB) for conditional attachments
    const valanceStyle = d_safeString_(vals[idx(DRAFTS_V2.COLS.VALANCE_STYLE)]); // Z
    const fabric = d_safeString_(vals[idx(DRAFTS_V2.COLS.FABRIC)]); // AB
    
    // Build list of missing items
    const missingItems = [];
    if (!fabricColor) missingItems.push('Fabric Color');
    if (!name) missingItems.push('Your Name');
    if (!displayName) missingItems.push('Display Name');
    if (!phone) missingItems.push('Best Phone Number');
    if (!email) missingItems.push('Email Address');
    if (!address) missingItems.push('Project Address');
    if (!jobType) missingItems.push('What kind of awning you are looking to get');
    if (!length || !width) missingItems.push('Rough dimensions of the awning (Length x Width)');
    
    // If nothing is missing, don't send
    if (missingItems.length === 0) {
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue('All customer info already complete - no email needed');
      return { toast: 'All customer info already complete' };
    }
    
    // Create subject
    const subject = 'Info request for awning';
    
    // Build bullet list
    let bulletList = '';
    let htmlBulletList = '';
    
    missingItems.forEach(item => {
      bulletList += `• ${item}\n`;
      htmlBulletList += `  <li>${d_htmlEscape_(item)}</li>\n`;
    });
    
    // Determine if we need Sunbrella link based on conditions
    const needsSunbrellaLink = 
      (valanceStyle.toLowerCase() === 'wrapped' && fabric.toLowerCase() === 'vinyl') ||
      (valanceStyle.toLowerCase() === 'hanging' && fabric.toLowerCase() === 'vinyl');
    
    // Create plain text body
    let plainBody = `Hello ${firstName},

May I just get a bit more info from you?

Please provide:
${bulletList}`;

    // Add Sunbrella link to plain text if needed
    if (needsSunbrellaLink) {
      plainBody += `\nSunbrella colors here: https://www.sunbrella.com/browse-fabrics/fabrics-by-use/shade-awnings-pergolas\n`;
    }

    plainBody += `\nAlso please send me some pics of the frame of the awning and I will get an estimate to you right away.

Thank you so much and call me if you have any questions.

Best Regards,
Gino Carneiro
Walker Awning`;
    
    // Create HTML body
    let htmlBody = `<div style="font-family: Arial, sans-serif; color: #333;">
<p>Hello ${d_htmlEscape_(firstName)},</p>

<p>May I just get a bit more info from you?</p>

<p><strong>Please provide:</strong></p>
<ul style="line-height: 1.8;">
${htmlBulletList}</ul>`;

    // Add Sunbrella link to HTML if needed
    if (needsSunbrellaLink) {
      htmlBody += `\n<p><a href="https://www.sunbrella.com/browse-fabrics/fabrics-by-use/shade-awnings-pergolas" target="_blank" style="color: #0066cc; text-decoration: underline;">Sunbrella colors here</a></p>`;
    }

    htmlBody += `\n<p>Also please send me some pics of the frame of the awning and I will get an estimate to you right away.</p>

<p>Thank you so much and call me if you have any questions.</p>

<p>Best Regards,<br>
Gino Carneiro<br>
Walker Awning</p>
</div>`;
    
    // Check if email exists
    if (customerEmail && d_isValidEmail_(customerEmail)) {
      // Create Gmail draft options
      const options = {
        htmlBody: htmlBody
      };
      
      // Handle attachments based on conditions
      const attachments = [];
      
      try {
        // Condition 1: If AB = Sunbrella, attach 2025 Sunbrella Colors.pdf
        if (fabric.toLowerCase() === 'sunbrella') {
          if (DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS && DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.SUNBRELLA_FILE_ID) {
            const sunbrellaFile = DriveApp.getFileById(DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.SUNBRELLA_FILE_ID);
            attachments.push(sunbrellaFile.getBlob());
          }
        }
        
        // Condition 2: If Z = wrapped AND AB = Vinyl, attach 3 vinyl files
        if (valanceStyle.toLowerCase() === 'wrapped' && fabric.toLowerCase() === 'vinyl') {
          if (DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS && DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_FERRARI_FILE_ID) {
            const ferrariFile = DriveApp.getFileById(DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_FERRARI_FILE_ID);
            attachments.push(ferrariFile.getBlob());
          }
          if (DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS && DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_PATIO500_FILE_ID) {
            const patio500File = DriveApp.getFileById(DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_PATIO500_FILE_ID);
            attachments.push(patio500File.getBlob());
          }
          if (DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS && DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_COASTLINE_FILE_ID) {
            const coastlineFile = DriveApp.getFileById(DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_COASTLINE_FILE_ID);
            attachments.push(coastlineFile.getBlob());
          }
        }
        
        // Condition 3: If Z = hanging AND AB = Vinyl, attach 2 vinyl files
        if (valanceStyle.toLowerCase() === 'hanging' && fabric.toLowerCase() === 'vinyl') {
          if (DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS && DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_PATIO500_FILE_ID) {
            const patio500File = DriveApp.getFileById(DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_PATIO500_FILE_ID);
            attachments.push(patio500File.getBlob());
          }
          if (DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS && DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_COASTLINE_FILE_ID) {
            const coastlineFile = DriveApp.getFileById(DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS.VINYL_COASTLINE_FILE_ID);
            attachments.push(coastlineFile.getBlob());
          }
        }
        
        // Add attachments to options if any exist
        if (attachments.length > 0) {
          options.attachments = attachments;
        }
        
      } catch (attachErr) {
        // Log error but continue - draft will be created without attachments
        console.error('Could not attach files:', attachErr);
        const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
        logCell.setValue('Warning: Some attachments failed. Check file IDs in DRAFTS_V2.CUSTOMER_INFO_ATTACHMENTS config.');
      }
      
      // Create draft
      try {
        const draft = d_withRetry_(() => GmailApp.createDraft(customerEmail, subject, plainBody, options));
        const draftMessageId = draft.getMessage().getId();
        const draftUrl = 'https://mail.google.com/mail/u/0/#drafts?compose=' + encodeURIComponent(draftMessageId);
        
        // Create rich text link for column B
        const linkText = '📋 Info';
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(linkText)
          .setLinkUrl(0, linkText.length, draftUrl)
          .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
          .build();
        
        // Update column B
        const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
        logCell.setRichTextValue(richText);
        
        return { toast: 'Info request draft created & linked in column B' };
        
      } catch (err) {
        const msg = d_specificErrorMessage_(err);
        const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
        logCell.setValue(msg);
        return { toast: msg };
      }
      
    } else {
      // No email - just put the text message in column B
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue('📱 TEXT MESSAGE (copy below):\n\n' + plainBody);
      return { toast: 'Text message script placed in column B for copying' };
    }
    
  } catch (err) {
    const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    logCell.setValue('Error creating info request: ' + d_shortErr_(err));
    return { toast: 'Error creating info request: ' + d_shortErr_(err) };
  }
}

/**
 * Create COI request draft when "COI Req" is selected
 */
function v2_createCOIDraft_(sh, row) {
  try {
    // COI Request only works on F/U, Awarded, and Heaven sheets
    const sheetName = sh.getName();
    if (sheetName !== 'F/U' && sheetName !== 'Awarded' && sheetName !== 'Heaven') {
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue('COI Request only available on F/U, Awarded, and Heaven sheets');
      return { toast: 'COI Request only available on F/U, Awarded, and Heaven sheets' };
    }
    
    const lastCol = sh.getLastColumn();
    const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const idx = (L) => d_colLetterToIndex_(L) - 1;
    
    // Get customer data
    const customerName = d_safeString_(vals[idx(DRAFTS_V2.COLS.CUSTOMER_NAME)]);
    const customerEmail = d_safeString_(vals[idx(DRAFTS_V2.COLS.EMAIL)]);
    const address = d_safeString_(vals[idx(DRAFTS_V2.COLS.ADDRESS)]);
    const displayName = d_safeString_(vals[idx(DRAFTS_V2.COLS.DISPLAY_NAME)]) || 'Unnamed Lead';
    
    // Validate required fields
    if (!customerName) {
      const msg = 'No customer name found in column E';
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
    if (!customerEmail || !d_isValidEmail_(customerEmail)) {
      const msg = 'No valid customer email found in column I';
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
    if (!address) {
      const msg = 'No address found in column J';
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
    if (!displayName || displayName === 'Unnamed Lead') {
      const msg = 'No display name found in column F';
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
    // Create subject
    const subject = `COI Request: ${displayName}`;
    
    // Create plain text body
    const plainBody = `Hello wonderful Keyes team,

     May I please have the COI for:
${displayName}
${address}

Please forward (with attached W9) to ${customerName} at ${customerEmail} and CC Gino@WalkerAwning.com.

So they get all the docs in one thread.

Thank you all in advance for your timely replies and great work!

Best regards,
Gino Carneiro`;
    
    // Create HTML body with bold inserted data
    const htmlBody = `<div style="font-family: Arial, sans-serif; color: #333;">
<p>Hello wonderful Keyes team,</p>

<p style="margin-left: 20px;">May I please have the COI for:<br>
<strong>${d_htmlEscape_(displayName)}</strong><br>
<strong>${d_htmlEscape_(address)}</strong></p>

<p>Please forward (with attached W9) to <strong>${d_htmlEscape_(customerName)}</strong> at <strong>${d_htmlEscape_(customerEmail)}</strong> and CC Gino@WalkerAwning.com.</p>

<p>So they get all the docs in one thread.</p>

<p>Thank you all in advance for your timely replies and great work!</p>

<p>Best regards,<br>
Gino Carneiro</p>
</div>`;
    
    // Create draft options
    const options = {
      htmlBody: htmlBody
    };
    
    // Attach W9 PDF from Drive
    if (DRAFTS_V2.COI_REQUEST && DRAFTS_V2.COI_REQUEST.ATTACHMENT_FILE_ID) {
      try {
        const file = DriveApp.getFileById(DRAFTS_V2.COI_REQUEST.ATTACHMENT_FILE_ID);
        options.attachments = [file.getBlob()];
      } catch (attachErr) {
        // Log error but continue - draft will be created without attachment
        console.error('Could not attach W9 file:', attachErr);
        const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
        logCell.setValue('Warning: W9 attachment failed. Check file ID. Draft created without attachment.');
        // Don't return - still create the draft
      }
    }
    
    // Create draft
    try {
      const draft = d_withRetry_(() => GmailApp.createDraft(
        DRAFTS_V2.COI_REQUEST.RECIPIENTS.join(','), 
        subject, 
        plainBody, 
        options
      ));
      const draftMessageId = draft.getMessage().getId();
      const draftUrl = 'https://mail.google.com/mail/u/0/#drafts?compose=' + encodeURIComponent(draftMessageId);
      
      // Create rich text link for column B
      const linkText = '📋 COI Request';
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(linkText)
        .setLinkUrl(0, linkText.length, draftUrl)
        .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
        .build();
      
      // Update column B
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setRichTextValue(richText);
      
      return { toast: 'COI request draft created & linked in column B' };
      
    } catch (err) {
      const msg = d_specificErrorMessage_(err);
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
  } catch (err) {
    const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    logCell.setValue('Error creating COI draft: ' + d_shortErr_(err));
    return { toast: 'Error creating COI draft: ' + d_shortErr_(err) };
  }
}

/**
 * Create rough quote draft/message when "Rough quote" is selected
 */
function v2_createRoughQuote_(sh, row) {
  try {
    const lastCol = sh.getLastColumn();
    const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const idx = (L) => d_colLetterToIndex_(L) - 1;
    
    // Get data
    const customerEmail = d_safeString_(vals[idx(DRAFTS_V2.COLS.EMAIL)]);
    const displayName = d_safeString_(vals[idx(DRAFTS_V2.COLS.DISPLAY_NAME)]) || 'Unnamed Lead';
    const customerName = d_safeString_(vals[idx(DRAFTS_V2.COLS.CUSTOMER_NAME)]);
    const firstName = customerName ? customerName.split(' ')[0] : 'there';
    
    // Get dimensions and job info
    const length = parseFloat(vals[idx(DRAFTS_V2.COLS.LEN)]) || 0;
    const width = parseFloat(vals[idx(DRAFTS_V2.COLS.WIDTH)]) || 0;
    const jobType = d_safeString_(vals[idx(DRAFTS_V2.COLS.JOB_TYPE)]);
    const fabric = d_safeString_(vals[idx(DRAFTS_V2.COLS.FABRIC)]);
    
    // Validate dimensions
    if (length <= 0 || width <= 0) {
      const msg = 'Missing or invalid dimensions (Length/Width)';
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue(msg);
      return { toast: msg };
    }
    
    // Calculate price based on job type
    let priceText = '';
    const jobTypeLower = jobType.toLowerCase();
    
    if (jobTypeLower === 'complete') {
      const minPrice = Math.round(length * width * 50);
      const maxPrice = Math.round(length * width * 60);
      priceText = `${minPrice.toLocaleString()} - ${maxPrice.toLocaleString()}`;
      
    } else if (jobTypeLower === 're-cover') {
      // Use minimum width of 5 for calculations
      const calcWidth = Math.max(width, DRAFTS_V2.ROUGH_QUOTE_EMAIL.MIN_WIDTH_RECOVER);
      const totalFeet = Math.ceil(length / 5) * calcWidth;
      const yards = totalFeet / 3;
      const minPrice = Math.round(yards * 105);
      const maxPrice = Math.round(yards * 115);
      priceText = `${minPrice.toLocaleString()} - ${maxPrice.toLocaleString()}`;
      
    } else if (jobTypeLower === 'aluminum canopy') {
      const minPrice = Math.round(length * width * 100);
      const maxPrice = Math.round(length * width * 150);
      priceText = `${minPrice.toLocaleString()} - ${maxPrice.toLocaleString()}`;
      
    } else {
      priceText = 'price to be determined';
    }
    
    // Build message body
    const fabricText = fabric ? ` ${fabric}` : '';
    const dimensionText = `${length}' x ${width}'`;
    
    const messageBody = `Hello ${firstName},\n\n` +
      `For the ${dimensionText}${fabricText} awning ${jobType}, it will be around ${priceText}. Please let me know what you think.` +
      DRAFTS_V2.ROUGH_QUOTE_EMAIL.SIGNATURE;
    
    const htmlBody = `<div style="font-family: Arial, sans-serif; color: #333;">
<p>Hello ${d_htmlEscape_(firstName)},</p>
<p>For the ${d_htmlEscape_(dimensionText)}${d_htmlEscape_(fabricText)} awning ${d_htmlEscape_(jobType)}, it will be around <strong>${d_htmlEscape_(priceText)}</strong>. Sound good?</p>
<p style="white-space: pre-line;">${d_htmlEscape_(DRAFTS_V2.ROUGH_QUOTE_EMAIL.SIGNATURE)}</p>
</div>`;
    
    // Check if email exists
    if (customerEmail && d_isValidEmail_(customerEmail)) {
      // Create Gmail draft
      const subject = d_templateSafe_(DRAFTS_V2.ROUGH_QUOTE_EMAIL.SUBJECT_TEMPLATE, { displayName });
      
      try {
        const draft = d_withRetry_(() => GmailApp.createDraft(customerEmail, subject, messageBody, { htmlBody }));
        const draftMessageId = draft.getMessage().getId();
        const draftUrl = 'https://mail.google.com/mail/u/0/#drafts?compose=' + encodeURIComponent(draftMessageId);
        
        // Create rich text link for column B
        const linkText = '💬 Rough: ';
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(linkText)
          .setLinkUrl(0, linkText.length, draftUrl)
          .setTextStyle(0, linkText.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
          .build();
        
        // Update column B
        const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
        logCell.setRichTextValue(richText);
        
        return { toast: 'Rough quote draft created & linked in column B' };
        
      } catch (err) {
        const msg = d_specificErrorMessage_(err);
        const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
        logCell.setValue(msg);
        return { toast: msg };
      }
      
    } else {
      // No email - just put the text message in column B
      const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
      logCell.setValue('📱 TEXT MESSAGE (copy below):\n\n' + messageBody);
      return { toast: 'Text message script placed in column B for copying' };
    }
    
  } catch (err) {
    const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    logCell.setValue('Error creating rough quote: ' + d_shortErr_(err));
    return { toast: 'Error creating rough quote: ' + d_shortErr_(err) };
  }
}

/** Backfill all rows where Stage == TARGET_STAGE. */
function createDraftsForAllRows_V2() {
  const ss = v2_getSpreadsheet_();
  const sh = ss.getSheetByName(DRAFTS_V2.SHEETS.LEADS);
  if (!sh) throw new Error('Leads sheet not found.');

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const idx = (L) => d_colLetterToIndex_(L) - 1;
  const stageIdx = idx(DRAFTS_V2.COLS.STAGE);
  const logBColIndex = d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B);

  const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
  const values = dataRange.getValues();
  const rtv    = dataRange.getRichTextValues();

  const logBRange = sh.getRange(2, logBColIndex, lastRow - 1, 1);
  const logBRTV = logBRange.getRichTextValues();
  const outB = logBRTV.map(r => [r[0]]);

  let created = 0, skipped = 0, failed = 0;

  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 2;
    const val = String(values[i][stageIdx] || '').trim().toLowerCase();
    if (val !== String(DRAFTS_V2.TARGET_STAGE).toLowerCase()) continue;

    const existingLink = d_firstLinkInRichText_(logBRTV[i][0]);
    if (DRAFTS_V2.EMAIL.SKIP_IF_DRAFT_EXISTS && d_isGmailDraftUrl_(existingLink)) { skipped++; continue; }

    try {
      const r = v2_createDraftForRow_(sh, rowNum, false, values[i], rtv[i],
        (richText)=>{ outB[i][0] = richText; });
      if (r.ok) created++; else failed++;
    } catch (err) {
      failed++;
      outB[i][0] = SpreadsheetApp.newRichTextValue().setText('Error: ' + d_shortErr_(err)).build();
    }
  }

  logBRange.setRichTextValues(outB);
  SpreadsheetApp.getActive().toast(`Backfill → created:${created} | skipped:${skipped} | failed:${failed}`, 'Draft Creator', 7);
}

/* =========================
 * Row → Draft creation
 * ========================= */
function v2_createDraftForRow_(sh, row, respectExisting, rowValsOpt, rowRtvOpt, batchReceiverOpt) {
  const lastCol = sh.getLastColumn();
  const vals = rowValsOpt || sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const rtv  = rowRtvOpt  || sh.getRange(row, 1, 1, lastCol).getRichTextValues()[0];
  const idx = (L) => d_colLetterToIndex_(L) - 1;

  const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));

  // Notes (diagnostics)
  const notes = [];

  // Idempotency
  if (respectExisting) {
    const existing = d_firstLinkInRichText_(logCell.getRichTextValue());
    if (d_isGmailDraftUrl_(existing)) return { ok:true, skipped:true, toast:'Skipped (existing draft link in B).' };
  }

  const displayName = d_safeString_(vals[idx(DRAFTS_V2.COLS.DISPLAY_NAME)]) || 'Unnamed Lead';
  const jobType     = d_safeString_(vals[idx(DRAFTS_V2.COLS.JOB_TYPE)]);

  // Subject
  const prefix = (DRAFTS_V2.EMAIL.SUBJECT_PREFIX || 'Proposal Review').trim();
  let subject  = d_templateSafe_(DRAFTS_V2.EMAIL.SUBJECT_TEMPLATE, { prefix, displayName, jobType });
  if (DRAFTS_V2.EMAIL.SUBJECT_MAX_LENGTH && subject.length > DRAFTS_V2.EMAIL.SUBJECT_MAX_LENGTH) {
    subject = subject.substring(0, DRAFTS_V2.EMAIL.SUBJECT_MAX_LENGTH);
  }

  // URLs
  const photoUrl = d_extractUrlFromCell_(vals[idx(DRAFTS_V2.COLS.FOLDER_URL)], rtv[idx(DRAFTS_V2.COLS.FOLDER_URL)], sh, row, DRAFTS_V2.COLS.FOLDER_URL);
  const geUrl    = d_extractUrlFromCell_(vals[idx(DRAFTS_V2.COLS.GE_URL)],     rtv[idx(DRAFTS_V2.COLS.GE_URL)],     sh, row, DRAFTS_V2.COLS.GE_URL);
  const qbUrl    = d_extractUrlFromCell_(vals[idx(DRAFTS_V2.COLS.QB_URL)],     rtv[idx(DRAFTS_V2.COLS.QB_URL)],     sh, row, DRAFTS_V2.COLS.QB_URL);
// NEW: Generate route map data (includes distance/duration)
  let routeMapData = null;
  const address = d_safeString_(vals[idx(DRAFTS_V2.COLS.ADDRESS)]);
  if (address) {
    try {
      routeMapData = d_generateRouteMapUrl_(address);
      if (!routeMapData || !routeMapData.mapUrl) notes.push('Route map failed');
    } catch (mapErr) {
      console.error('Route map generation error:', mapErr);
      notes.push('Route map error');
    }
  } else {
    notes.push('No address for route map');
  }

  // Generate satellite aerial view
  let satelliteMapUrl = null;
  if (address) {
    try {
      satelliteMapUrl = d_generateSatelliteMapUrl_(address);
      if (!satelliteMapUrl) notes.push('Satellite map failed');
    } catch (satErr) {
      console.error('Satellite map generation error:', satErr);
      notes.push('Satellite map error');
    }
  }

  // Re-cover HTML snapshot
  let recoverHtml = null;
  try {
    const ss = sh.getParent();
    console.log('Starting Re-cover HTML export for spreadsheet: ' + ss.getId());
    
    // Just wait for calculations, don't try to change active range (fails in triggers)
    Utilities.sleep(DRAFTS_V2.RECOVER.WAIT_MS || 2000);
    
    recoverHtml = v2_exportRecoverRangeAsHtml_(ss);
    if (!recoverHtml) {
      notes.push('Re-cover HTML empty');
      console.log('recoverHtml is NULL or empty');
    } else {
      console.log('recoverHtml generated, length: ' + recoverHtml.length);
    }
  } catch (htmlErr) {
    if (DRAFTS_V2.RECOVER && DRAFTS_V2.RECOVER.DEBUG) {
      console.error('HTML export failed:', htmlErr);
    }
    notes.push('HTML fail');
  }

  // Body
  const html = v2_buildHtmlBody_({ photoUrl, geUrl, qbUrl, recoverHtml, routeMapData, address, satelliteMapUrl });
  const plain = v2_buildPlainBody_({ photoUrl, geUrl, qbUrl });

  // Recipients
  const to  = (DRAFTS_V2.EMAIL.TO  || []).filter(Boolean);
  const cc  = (DRAFTS_V2.EMAIL.CC  || []).filter(Boolean);
  const bcc = (DRAFTS_V2.EMAIL.BCC || []).filter(Boolean);
  if (!to.length) { const msg='Error: No TO recipients configured.'; d_writeB_(sh,row,msg,batchReceiverOpt); return {ok:false,toast:msg}; }
  const invalid = [].concat(to,cc,bcc).filter(x=>!d_isValidEmail_(x));
  if (invalid.length) { const msg='Error: Invalid email address: '+invalid[0]; d_writeB_(sh,row,msg,batchReceiverOpt); return {ok:false,toast:msg}; }

  const options = {};
  if (html) options.htmlBody = html;
  if (cc.length)  options.cc  = cc.join(',');
  if (bcc.length) options.bcc = bcc.join(',');

  try {
    const draft = d_withRetry_(()=> GmailApp.createDraft(to.join(','), subject, plain, options));
    const draftMessageId = draft.getMessage().getId();
    const draftUrl = 'https://mail.google.com/mail/u/0/#drafts?compose=' + encodeURIComponent(draftMessageId);

    // Rich text in B: link + optional diagnostics if snapshot failed
    const base = '✅ Est Draft';
    const suffix = (!recoverHtml && notes.length) ? '\n' + notes.join(' | ') : '';
    const rich = SpreadsheetApp.newRichTextValue()
      .setText(base + suffix)
      .setLinkUrl(0, base.length, draftUrl)
      .setTextStyle(0, base.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
      .build();

    if (batchReceiverOpt) batchReceiverOpt(rich); else logCell.setRichTextValue(rich);

    return { ok:true, toast:'Draft created & linked in column B.' };
  } catch (err) {
    const msg = d_specificErrorMessage_(err);
    d_writeB_(sh, row, msg, batchReceiverOpt);
    return { ok:false, toast:msg };
  }
}

/**
 * HTML export of Re-cover range (FIXED VERSION)
 * Uses proper A1 notation in the API call
 */
function v2_exportRecoverRangeAsHtml_(ss) {
  try {
    const spreadsheetId = ss.getId();
    const sheet = ss.getSheetByName(DRAFTS_V2.SHEETS.RECOVER);
    if (!sheet) return null;

    const rangeA1 = DRAFTS_V2.RECOVER.SNAPSHOT_RANGE_A1; // e.g., "A1:K14"
    const sheetName = sheet.getName();
    
    // Use proper A1 notation: SheetName!A1:K14
    const fullRange = `${sheetName}!${rangeA1}`;
    
    const endpoint = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}` +
      `?ranges=${encodeURIComponent(fullRange)}` +
      `&includeGridData=true` +
      `&fields=sheets(data(rowData(values(effectiveValue,formattedValue,hyperlink,effectiveFormat)),rowMetadata,columnMetadata),merges)`;

    const resp = UrlFetchApp.fetch(endpoint, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      if (DRAFTS_V2.RECOVER && DRAFTS_V2.RECOVER.DEBUG) {
        console.error('API call failed:', resp.getResponseCode(), resp.getContentText());
      }
      return null;
    }

    const json = JSON.parse(resp.getContentText());
    const sheetData = (json && json.sheets && json.sheets[0]) ? json.sheets[0] : null;
    if (!sheetData || !sheetData.data || !sheetData.data[0]) return null;

    return v2_sheetsDataToHtml_(sheetData, sheet);
  } catch (err) {
    if (DRAFTS_V2.RECOVER && DRAFTS_V2.RECOVER.DEBUG) {
      console.error('HTML export error:', err);
    }
    return null;
  }
}

function v2_sheetsDataToHtml_(sheetData, sheet) {
  const data = sheetData.data[0];
  const rowData = data.rowData || [];

  const merges = sheetData.merges || [];
  const numRows = rowData.length;
  if (numRows === 0) return null;
  
  const numCols = rowData[0].values ? rowData[0].values.length : 0;
  if (numCols === 0) return null;

  const skipCell = Array.from({length: numRows}, ()=> Array(numCols).fill(false));
  const mergedTopLeft = Array.from({length: numRows}, ()=> Array(numCols).fill(null));

  merges.forEach(m => {
    const c0 = m.startColumnIndex || 0;
    const c1 = (m.endColumnIndex || 0) - 1;
    const r0 = m.startRowIndex || 0;
    const r1 = (m.endRowIndex || 0) - 1;

    if (c0 < 0 || c1 >= numCols || r0 < 0 || r1 >= numRows) return;

    const colspan = c1 - c0 + 1;
    const rowspan = r1 - r0 + 1;
    mergedTopLeft[r0][c0] = { rowspan, colspan, r0, c0 };

    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        if (!(r === r0 && c === c0)) skipCell[r][c] = true;
      }
    }
  });

  const colMeta = data.columnMetadata || [];
  const rowMeta = data.rowMetadata || [];

  const colWidths = Array.from({length: numCols}, (_, c) => {
    const m = colMeta[c] && colMeta[c].pixelSize;
    return m || 100;
  });
  const rowHeights = Array.from({length: numRows}, (_, r) => {
    const m = rowMeta[r] && rowMeta[r].pixelSize;
    return m || 21;
  });

  const toHex = (c) => {
    if (!c) return null;
    const rgb = (c.rgbColor) ? c.rgbColor : c;
    const R = Math.round((rgb.red || 0) * 255);
    const G = Math.round((rgb.green || 0) * 255);
    const B = Math.round((rgb.blue || 0) * 255);
    const hh = (n)=> ('0' + Math.max(0, Math.min(255, n)).toString(16)).slice(-2);
    return `#${hh(R)}${hh(G)}${hh(B)}`;
  };

  const borderCss = (b) => {
    if (!b) return null;
    const styleMap = { DOTTED:'dotted', DASHED:'dashed', SOLID:'solid', SOLID_MEDIUM:'solid', SOLID_THICK:'solid', DOUBLE:'double', NONE:'none' };
    const widthMap = { DOTTED:'1px', DASHED:'1px', SOLID:'1px', SOLID_MEDIUM:'2px', SOLID_THICK:'3px', DOUBLE:'3px', NONE:'0' };
    const cssStyle = styleMap[b.style || 'SOLID'] || 'solid';
    const width = widthMap[b.style || 'SOLID'] || '1px';
    const color = toHex(b.colorStyle || b.color) || '#000000';
    if (cssStyle === 'none' || width === '0') return 'border:none;';
    return `border:${width} ${cssStyle} ${color};`;
  };

  let html = '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;border-spacing:0;font-family:Arial,sans-serif;">';

  for (let r = 0; r < numRows; r++) {
    html += `<tr style="height:${rowHeights[r]}px;">`;
    const row = rowData[r] || {};
    const cells = row.values || [];

    for (let c = 0; c < numCols; c++) {
      if (skipCell[r][c]) continue;

      const cell = cells[c] || {};
      const eff = cell.effectiveFormat || {};
      const fmt = eff.textFormat || {};

      const bgHex = toHex(eff.backgroundColorStyle || eff.backgroundColor) || '#ffffff';
      const alignH = eff.horizontalAlignment || 'left';
      const alignV = eff.verticalAlignment || 'middle';

      const v = (cell.formattedValue != null) ? String(cell.formattedValue) :
                (cell.effectiveValue != null ? String(Object.values(cell.effectiveValue)[0]) : '');
      const link = cell.hyperlink || null;

      const borders = eff.borders || {};
      const topCss = borderCss(borders.top);
      const bottomCss = borderCss(borders.bottom);
      const leftCss = borderCss(borders.left);
      const rightCss = borderCss(borders.right);
      const defaultBorder = 'border:1px solid #000;';

      const widthCss = `width:${colWidths[c]}px;`;
      const colorCss = (fmt.foregroundColorStyle || fmt.foregroundColor) ? `color:${toHex(fmt.foregroundColorStyle || fmt.foregroundColor)};` : '';
      const boldCss  = fmt.bold ? 'font-weight:bold;' : '';
      const italicCss= fmt.italic ? 'font-style:italic;' : '';
      const uCss     = fmt.underline ? 'text-decoration:underline;' : '';
      const sizeCss  = fmt.fontSize ? `font-size:${fmt.fontSize}px;` : '';
      const fontCss  = fmt.fontFamily ? `font-family:${fmt.fontFamily},Arial,sans-serif;` : '';
      const alignCss = `text-align:${alignH};vertical-align:${alignV};`;
      const padCss   = 'padding:4px 8px;';

      const borderCssAll = (topCss || bottomCss || leftCss || rightCss)
        ? `${topCss || ''}${rightCss || ''}${bottomCss || ''}${leftCss || ''}`
        : defaultBorder;

      let tdAttrs = `style="${widthCss}background:${bgHex};${colorCss}${boldCss}${italicCss}${uCss}${sizeCss}${fontCss}${alignCss}${padCss}${borderCssAll}"`;

      const mergeInfo = mergedTopLeft[r][c];
      if (mergeInfo) {
        if (mergeInfo.rowspan > 1) tdAttrs += ` rowspan="${mergeInfo.rowspan}"`;
        if (mergeInfo.colspan > 1) tdAttrs += ` colspan="${mergeInfo.colspan}"`;
      }

      const safeVal = d_htmlEscape_(v || '');
      const content = link
        ? `<a href="${d_htmlEscape_(link)}" target="_blank" style="color:inherit;text-decoration:underline;">${safeVal}</a>`
        : safeVal;

      html += `<td ${tdAttrs}>${content}</td>`;
    }
    html += '</tr>';
  }

  html += '</table>';
  return html;
}

/* ================
 * Email body builders
 * ================ */
function v2_buildHtmlBody_(data) {
  const esc = d_htmlEscape_;
  const L = DRAFTS_V2.EMAIL.LINK_LABELS || { PHOTOS:'PICS', EARTH:'Google Earth', QUICKBOOKS:'Quickbooks' };
  let html = '<div style="font-size:24pt;">';

  html += data.photoUrl
    ? '<a href="' + esc(data.photoUrl) + '" target="_blank" style="font-size:24pt;">' + esc(L.PHOTOS) + '</a>'
    : '<span style="font-size:24pt;">' + esc(L.PHOTOS) + ' (No URL)</span>';
  html += '<br><br>';

  if (data.recoverHtml) {
    html += data.recoverHtml + '<br>';
  }

  html += data.geUrl
    ? '<a href="' + esc(data.geUrl) + '" target="_blank" style="font-size:24pt;">' + esc(L.EARTH) + '</a>'
    : '<span style="font-size:24pt;">' + esc(L.EARTH) + ' (No URL)</span>';
  html += '<br>';

  // Satellite aerial view image - clickable to Google Earth
  if (data.satelliteMapUrl && data.geUrl) {
    html += '<a href="' + esc(data.geUrl) + '" target="_blank">';
    html += '<img src="' + esc(data.satelliteMapUrl) + '" alt="Aerial View" style="max-width:100%; border:1px solid #ccc; border-radius:8px;">';
    html += '</a><br>';
  }

  // Route Map - simple display
  if (data.routeMapData && data.routeMapData.mapUrl && data.address) {
    const mapsDirectionsUrl = 'https://www.google.com/maps/dir/' + encodeURIComponent(DRAFTS_V2.MAPS_CONFIG.SHOP_ADDRESS) + '/' + encodeURIComponent(data.address);
    
    // Distance and Time - plain text
    if (data.routeMapData.distance || data.routeMapData.duration) {
      let infoText = '';
      if (data.routeMapData.distance) {
        infoText += data.routeMapData.distance;
      }
      if (data.routeMapData.distance && data.routeMapData.duration) {
        infoText += '  |  ';
      }
      if (data.routeMapData.duration) {
        infoText += data.routeMapData.duration;
      }
      html += '<span style="font-size:24pt;">' + esc(infoText) + '</span><br>';
    }
    
    // Map image - clickable
    html += '<a href="' + esc(mapsDirectionsUrl) + '" target="_blank">';
    html += '<img src="' + esc(data.routeMapData.mapUrl) + '" alt="Route Map" style="max-width:100%; border:1px solid #ccc; border-radius:8px;">';
    html += '</a><br><br>';
  }

  html += data.qbUrl
    ? '<a href="' + esc(data.qbUrl) + '" target="_blank" style="font-size:24pt;">' + esc(L.QUICKBOOKS) + '</a>'
    : '<span style="font-size:24pt;">' + esc(L.QUICKBOOKS) + ' (No URL)</span>';

  html += '</div>';
  return html;
}

function v2_buildPlainBody_(data) {
  const L = DRAFTS_V2.EMAIL.LINK_LABELS || { PHOTOS:'PICS', EARTH:'Google Earth', QUICKBOOKS:'Quickbooks' };
  const parts = [];
  parts.push(data.photoUrl ? (L.PHOTOS + '\n' + data.photoUrl) : (L.PHOTOS + ' (No URL)'));
  parts.push(data.geUrl   ? (L.EARTH  + '\n' + data.geUrl)   : (L.EARTH  + ' (No URL)'));
  parts.push(data.qbUrl   ? (L.QUICKBOOKS + '\n' + data.qbUrl) : (L.QUICKBOOKS + ' (No URL)'));
  return parts.join('\n\n');
}

/* =====================
 * Helpers
 * ===================== */
function d_extractUrlFromCell_(rawVal, richVal, sh, row, colLetter) {
  try {
    const cell = sh.getRange(row, d_colLetterToIndex_(colLetter));
    const rtv = cell.getRichTextValue();
    if (rtv) {
      const runs = rtv.getRuns();
      if (runs && runs.length) {
        for (let i = 0; i < runs.length; i++) {
          const u = runs[i].getLinkUrl();
          if (u) return String(u).trim();
        }
      }
      if (rtv.getLinkUrl && rtv.getLinkUrl()) return String(rtv.getLinkUrl()).trim();
      const text = rtv.getText && rtv.getText();
      if (text) for (let j = 0; j < text.length; j++) {
        const u = rtv.getLinkUrl(j);
        if (u) return String(u).trim();
      }
    }
  } catch (_) {}

  try {
    if (richVal && typeof richVal.getLinkUrl === 'function') {
      const whole = richVal.getLinkUrl();
      if (whole) return String(whole).trim();
    }
    if (richVal && typeof richVal.getText === 'function' && typeof richVal.getLinkUrl === 'function') {
      const txt = richVal.getText();
      if (txt) for (let j = 0; j < txt.length; j++) {
        const u = richVal.getLinkUrl(j);
        if (u) return String(u).trim();
      }
    }
  } catch (_) {}

  const asString = String(rawVal || '').trim();
  if (d_looksLikeUrl_(asString)) return asString;

  try {
    const cell2 = sh.getRange(row, d_colLetterToIndex_(colLetter));
    const f = cell2.getFormula();
    if (f && /^=HYPERLINK\(/i.test(f)) {
      const m = f.match(/^=HYPERLINK\(\s*"([^"]+)"/i);
      if (m && m[1]) return m[1].trim();
    }
  } catch (_) {}

  return '';
}

function d_looksLikeUrl_(s){ return /^https?:\/\/\S+/i.test(String(s||'')); }
function d_safeString_(v){ if (v==null) return ''; return v instanceof Date ? v.toLocaleString() : String(v).trim(); }
function d_htmlEscape_(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

function d_colLetterToIndex_(letter){
  let col = 0; const up = String(letter||'').toUpperCase();
  for (let i=0;i<up.length;i++) col = col * 26 + (up.charCodeAt(i) - 64);
  return col;
}

function v2_getSpreadsheet_() {
  if (DRAFTS_V2.SPREADSHEET_ID && DRAFTS_V2.SPREADSHEET_ID !== 'REPLACE_ME_WITH_YOUR_SHEET_ID') {
    return SpreadsheetApp.openById(DRAFTS_V2.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function d_writeB_(sh, row, textOrRich, batchReceiverOpt) {
  if (batchReceiverOpt) {
    batchReceiverOpt(typeof textOrRich === 'string'
      ? SpreadsheetApp.newRichTextValue().setText(textOrRich).build()
      : textOrRich);
  } else {
    const cell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    if (typeof textOrRich === 'string') cell.setValue(textOrRich); else cell.setRichTextValue(textOrRich);
  }
}

function d_withRetry_(fn) {
  const max = DRAFTS_V2.RETRY.MAX_ATTEMPTS || 3;
  const delays = DRAFTS_V2.RETRY.DELAYS_MS || [5000, 15000, 30000];
  for (let attempt=1; attempt<=max; attempt++){
    try { return fn(); }
    catch(err){
      if (attempt>=max) throw err;
      const msg = String(err);
      const transient = /Service invoked too many times|Rate Limit|Quota exceeded|Internal error|Timeout|temporary/i.test(msg);
      if (!transient) throw err;
      Utilities.sleep(delays[Math.min(attempt-1, delays.length-1)]);
    }
  }
  throw new Error('Unknown retry failure');
}

function d_specificErrorMessage_(err){
  const msg = d_shortErr_(err);
  if (/quota|too many times|rate limit/i.test(msg)) return 'Gmail quota exceeded - try again later';
  if (/Invalid\s+email|Bad Request.*Invalid argument.*email/i.test(msg)) return 'Invalid email address detected';
  if (/timeout|timed out|Network/i.test(msg)) return 'Network timeout - retrying...';
  return 'Draft error — ' + msg;
}
function d_shortErr_(err){ const s = (err && err.message) ? err.message : String(err||''); return s.length>200 ? s.slice(0,200)+'…' : s; }

function v2_validateConfig_() {
  const to = (DRAFTS_V2.EMAIL.TO || []).filter(Boolean);
  if (!to.length) throw new Error('Drafts V2 config error: EMAIL.TO is empty.');
  const all = [].concat(to, (DRAFTS_V2.EMAIL.CC||[]), (DRAFTS_V2.EMAIL.BCC||[])).filter(Boolean);
  const bad = all.filter(x => !d_isValidEmail_(x));
  if (bad.length) throw new Error('Drafts V2 config error: Invalid email in recipients: ' + bad[0]);
}
function d_isValidEmail_(s){ return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s||'').trim()); }

function d_firstLinkInRichText_(rtv){
  try{
    if (!rtv) return '';
    const runs = rtv.getRuns();
    if (runs && runs.length) for (let k=0;k<runs.length;k++){
      const u = runs[k].getLinkUrl(); if (u) return String(u);
    }
    if (rtv.getLinkUrl){ const u = rtv.getLinkUrl(); if (u) return String(u); }
    return '';
  } catch(_){ return ''; }
}

function d_isGmailDraftUrl_(u) {
  return typeof u === 'string' &&
         /^https:\/\/mail\.google\.com\/mail\/u\/\d+\/#(?:drafts|inbox|all)\?compose=/.test(u);
}
function d_templateSafe_(tpl, data) {
  try {
    let out = String(tpl || '');
    out = out.replace(/\$\{prefix\}/g, String(data.prefix || ''));
    out = out.replace(/\$\{displayName\}/g, String(data.displayName || ''));
    out = out.replace(/\$\{jobType\}/g, String(data.jobType || ''));
    out = out.replace(/\s*-\s*$/, '');
    return out.replace(/\s{2,}/g, ' ').trim();
  } catch (_) {
    return String(tpl || '');
  }
}
function d_timestamp_() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function d_safeLogErrorToSheet_(sh, row, err) {
  try {
    if (!sh || !row) {
      console.error('Error logging failed - no sheet/row:', err);
      return;
    }
    const logCell = sh.getRange(row, d_colLetterToIndex_(DRAFTS_V2.COLS.LOG_B));
    logCell.setValue('Error: ' + d_shortErr_(err));
  } catch (logErr) {
    console.error('Failed to log to sheet:', logErr, 'Original error:', err);
  }
}
/**
 * Generate Google Maps Static API URL for route image with distance/time data
 * Returns object: { mapUrl, distance, duration } or null on failure
 */
function d_generateRouteMapUrl_(destinationAddress) {
  const M = DRAFTS_V2.MAPS_CONFIG;
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_MAPS_API_KEY');
  if (!apiKey) {
    console.error('GOOGLE_MAPS_API_KEY not found in Script Properties');
    return null;
  }
  
  if (!destinationAddress) {
    console.error('No destination address provided for route map');
    return null;
  }
  
  try {
    const directionsUrl = 'https://maps.googleapis.com/maps/api/directions/json?' +
      'origin=' + encodeURIComponent(M.SHOP_ADDRESS) +
      '&destination=' + encodeURIComponent(destinationAddress) +
      '&key=' + apiKey;
    
    const response = UrlFetchApp.fetch(directionsUrl, { muteHttpExceptions: true });
    const directionsData = JSON.parse(response.getContentText());
    
    if (directionsData.status !== 'OK' || !directionsData.routes || directionsData.routes.length === 0) {
      console.error('Directions API error:', directionsData.status);
      // Return simple map without distance/time
      return { 
        mapUrl: d_generateSimpleMapUrl_(destinationAddress, apiKey),
        distance: null,
        duration: null
      };
    }
    
    const route = directionsData.routes[0];
    const leg = route.legs[0];
    const encodedPolyline = route.overview_polyline.points;
    
    // Extract distance and duration
    const distance = leg.distance ? leg.distance.text : null;
    const duration = leg.duration ? leg.duration.text : null;
    
    const staticMapUrl = 'https://maps.googleapis.com/maps/api/staticmap?' +
      'size=' + M.MAP_WIDTH + 'x' + M.MAP_HEIGHT +
      '&maptype=' + M.MAP_TYPE +
      '&markers=color:' + M.MARKER_ORIGIN_COLOR + '|label:A|' + encodeURIComponent(M.SHOP_ADDRESS) +
      '&markers=color:' + M.MARKER_DEST_COLOR + '|label:B|' + encodeURIComponent(destinationAddress) +
      '&path=color:' + M.ROUTE_COLOR + '|weight:' + M.ROUTE_WEIGHT + '|enc:' + encodedPolyline +
      '&key=' + apiKey;
    
    return { 
      mapUrl: staticMapUrl, 
      distance: distance,
      duration: duration
    };
    
  } catch (err) {
    console.error('Error generating route map URL:', err);
    return { 
      mapUrl: d_generateSimpleMapUrl_(destinationAddress, apiKey),
      distance: null,
      duration: null
    };
  }
}

/**
 * Fallback: Generate simple map URL with just markers (no route line)
 */
function d_generateSimpleMapUrl_(destinationAddress, apiKey) {
  const M = DRAFTS_V2.MAPS_CONFIG;
  
  return 'https://maps.googleapis.com/maps/api/staticmap?' +
    'size=' + M.MAP_WIDTH + 'x' + M.MAP_HEIGHT +
    '&maptype=' + M.MAP_TYPE +
    '&markers=color:' + M.MARKER_ORIGIN_COLOR + '|label:A|' + encodeURIComponent(M.SHOP_ADDRESS) +
    '&markers=color:' + M.MARKER_DEST_COLOR + '|label:B|' + encodeURIComponent(destinationAddress) +
    '&key=' + apiKey;
}

/**
 * Test function - run this to verify Maps API works
 */
function testMapsApiConnection() {
  const testAddress = '1600 Amphitheatre Parkway, Mountain View, CA';
  const mapUrl = d_generateRouteMapUrl_(testAddress);
  
  if (mapUrl) {
    Logger.log('SUCCESS! Map URL generated:');
    Logger.log(mapUrl);
    SpreadsheetApp.getActive().toast('Maps API working! Check logs for URL.', 'Test Success', 5);
  } else {
    Logger.log('FAILED: Could not generate map URL. Check API key in Script Properties.');
    SpreadsheetApp.getActive().toast('Maps API failed. Check logs.', 'Test Failed', 5);
  }
}
/**
 * Generate Google Maps Static API satellite view URL for a location
 * Returns the satellite image URL or null on failure
 */
function d_generateSatelliteMapUrl_(address) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_MAPS_API_KEY');
  if (!apiKey || !address) {
    return null;
  }
  
  try {
    // Satellite view centered on the address, zoomed in for property detail
    const satelliteUrl = 'https://maps.googleapis.com/maps/api/staticmap?' +
      'center=' + encodeURIComponent(address) +
      '&zoom=19' +  // High zoom for property detail
      '&size=640x400' +
      '&maptype=satellite' +
      '&markers=color:red|' + encodeURIComponent(address) +
      '&key=' + apiKey;
    
    return satelliteUrl;
  } catch (err) {
    console.error('Error generating satellite map URL:', err);
    return null;
  }
}
/** end-of-file */