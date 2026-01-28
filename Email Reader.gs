/**
 * EMAIL READER AUTOMATION - UNIFIED LEAD INGESTION
 * Version: 01/20-09:20AM EST by Claude Opus 4.1
 * 
 * FEATURES:
 * - Ruby emails: Automatically detected and processed
 * - "Add lead" emails: Automatically processed (same schedule as Ruby)
 * - Writes CSV to column A, Stage Automation splits it
 * - ALL processed emails get "LeadProcessed" label
 * - Uses LABELS (not read/unread) to prevent duplicates
 * - Three-tier fallback system (AI → Pattern → Structured → Manual Review)
 * 
 * CHANGES IN THIS VERSION:
 * - Changed to write 29-column CSV to column A
 * - Stage Automation handles the split
 * - Simplified flow - no more direct column writes
 */

const EMAIL_READER_CONFIG = {
  TARGET_SHEET: 'Leads',
  
  SEARCHES: [
    {
      name: 'Ruby Mail',
      query: 'from:noreply@ruby.com -label:LeadProcessed',
      enabled: true,
      sourceType: 'ruby',
      settings: {
        stage: '1. F/U',
        category: 'Res',
        useAI: true,
        useFallback: true
      }
    },
    {
      name: 'Add Lead',
      query: 'label:Add lead -label:LeadProcessed',
      enabled: true,
      sourceType: 'addlead',
      settings: {
        stage: '1. F/U',
        category: 'Res',
        useAI: true,
        useFallback: true
      }
    }
  ],
  
  PROCESSED_LABEL: 'LeadProcessed',
  ADD_LEAD_LABEL: 'Add lead',
  
  // Column count for CSV (A through AC = 29 columns)
  TOTAL_COLUMNS: 29,
  
  // Column indices (1-based)
  COLS: {
    DATE: 1,           // A
    LINK_B: 2,         // B
    COMMENTS: 3,       // C
    STAGE: 4,          // D
    NAME: 5,           // E
    DISPLAY: 6,        // F
    TYPE: 7,           // G
    PHONE: 8,          // H
    EMAIL: 9,          // I
    ADDRESS: 10,       // J
    DESC: 11,          // K
    L: 12,             // L (blank)
    JOB_DESC_M: 13,    // M
    QUOTE: 14,         // N
    CALCS: 15,         // O
    QB_URL: 16,        // P
    EARTH_LINK: 17,    // Q
    JOB_TYPE: 18,      // R
    S: 19,             // S (blank)
    LENGTH: 20,        // T
    WIDTH: 21,         // U
    FRONT_BAR: 22,     // V
    SHELF: 23,         // W
    WING_HEIGHT: 24,   // X
    NUM_WINGS: 25,     // Y
    VALANCE: 26,       // Z
    FRAME: 27,         // AA
    FABRIC: 28,        // AB
    AWNING_TYPE: 29    // AC
  },

  MAX_EMAILS_PER_RUN: 10,
  WAIT_FOR_AI_MS: 8000,
  MAX_AI_RETRIES: 3,
  ENABLE_LOGGING: true,
  
  // Ruby-specific patterns (for fallback parsing)
  RUBY_PATTERNS: {
    first: /First:\s*([^\n]+)/i,
    last: /Last:\s*([^\n]+)/i,
    phone: /Phone Number:\s*([^\n]+)/i,
    company: /Company:\s*([^\n]+)/i,
    regarding: /Regarding:\s*([^\n]+)/i,
    address: /Project Address:\s*([^\n]+)/i,
    email: /Email:\s*([^\n]+)/i,
    actions: /Actions:\s*([^\n]+)/i
  },
  
  // Generic patterns (for any email fallback)
  GENERIC_PATTERNS: {
    phone: /(\+?1[-.\s]?)?(\()?(\d{3})(\))?[-.\s]?(\d{3})[-.\s]?(\d{4})/,
    email: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/,
    address: /(\d+\s+[a-zA-Z0-9\s]+(?:st|street|ave|avenue|rd|road|ln|lane|blvd|boulevard|dr|drive|ct|court|way|pl|place|cir|circle)[,\s]+[a-zA-Z\s]+(?:,\s*)?(?:FL|Florida)?(?:\s+\d{5})?)/i
  }
};

/**
 * Build a 29-column CSV string from extracted data
 * Columns: A-AC (Date through Awning Type)
 */
function er_buildCSV_(data) {
  const C = EMAIL_READER_CONFIG;
  
  // Create array of 29 empty strings
  const values = new Array(C.TOTAL_COLUMNS).fill('');
  
  // Fill in the values we have (0-indexed)
  values[0] = data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd');  // A - Date
  values[1] = '';  // B - Will be overwritten with link after split
  values[2] = '';  // C - Comments
  values[3] = data.stage || '1. F/U';  // D - Stage
  values[4] = data.name || '';  // E - Customer Name
  values[5] = data.company || '';  // F - Display Name
  values[6] = data.type || 'Res';  // G - Type
  values[7] = data.phone || '';  // H - Phone
  values[8] = data.email || '';  // I - Email
  values[9] = data.address || '';  // J - Address (NO COMMAS!)
  values[10] = data.regarding || '';  // K - Job Description
  // L through AC remain blank unless we have awning specs
  
  // Clean all values - remove commas and extra whitespace
  const cleanedValues = values.map(v => {
    return String(v)
      .replace(/,/g, ' ')  // Remove ALL commas
      .replace(/\s+/g, ' ')  // Collapse whitespace
      .trim();
  });
  
  return cleanedValues.join(',');
}

/**
 * DIAGNOSTIC FUNCTION - Quick check, no popup
 */
function er_diagnosticCheck() {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const rubyCount = GmailApp.search('from:noreply@ruby.com -label:LeadProcessed', 0, 50).length;
    const addLeadCount = GmailApp.search(`label:"${C.ADD_LEAD_LABEL}" -label:LeadProcessed`, 0, 50).length;
    
    const triggers = ScriptApp.getProjectTriggers();
    const hasTrigger = triggers.some(t => t.getHandlerFunction() === 'er_processNewEmails');
    
    SpreadsheetApp.getActive().toast(
      `Ruby pending: ${rubyCount}\nAdd lead pending: ${addLeadCount}\nTrigger: ${hasTrigger ? '✅' : '❌'}`,
      'Email Reader Status',
      5
    );
    
    er_log_('Diagnostic check', { rubyCount, addLeadCount, hasTrigger });
    
  } catch (err) {
    SpreadsheetApp.getActive().toast(`Error: ${err.message}`, 'Diagnostic Error', 5);
  }
}

/**
 * Determine if email is from Ruby
 */
function er_isRubyEmail_(message) {
  const from = message.getFrom().toLowerCase();
  return from.includes('noreply@ruby.com');
}

/**
 * Process a single email - writes CSV to column A
 * Stage Automation will split it across columns
 */
function er_processEmail_(message, searchConfig, processedLabel) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const thread = message.getThread();
    
    // Check if already processed using LABEL
    const labels = thread.getLabels();
    if (labels.some(l => l.getName() === C.PROCESSED_LABEL)) {
      er_log_('Email already has LeadProcessed label - skipping', { messageId: message.getId() });
      return false;
    }
    
    const emailData = {
      subject: message.getSubject(),
      from: message.getFrom(),
      date: message.getDate(),
      body: message.getPlainBody(),
      htmlBody: message.getBody(),
      messageId: message.getId(),
      isRuby: er_isRubyEmail_(message)
    };
    
    const gmailUrl = `https://mail.google.com/mail/u/0/#inbox/${emailData.messageId}`;
    
    if (C.ENABLE_LOGGING) {
      er_log_('Processing email', {
        from: emailData.from,
        subject: emailData.subject,
        sourceType: emailData.isRuby ? 'Ruby' : 'Add lead',
        bodyLength: emailData.body.length,
        messageId: emailData.messageId
      });
    }
    
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(C.TARGET_SHEET);
    
    if (!sheet) {
      throw new Error(`Sheet "${C.TARGET_SHEET}" not found`);
    }
    
    // Extract data from email
    let extractedData = null;
    let parseMethod = 'none';
    
    // Try pattern parsing first (more reliable than AI for structured emails)
    extractedData = er_extractWithPatterns_(emailData);
    if (extractedData && (extractedData.name || extractedData.phone || extractedData.email)) {
      parseMethod = 'Pattern';
    }
    
    // If pattern parsing didn't get much, try structured extraction
    if (!extractedData || (!extractedData.name && !extractedData.phone && !extractedData.email)) {
      extractedData = er_extractStructured_(emailData);
      if (extractedData && (extractedData.name || extractedData.phone || extractedData.email)) {
        parseMethod = 'Structured';
      }
    }
    
    // If still nothing, create minimal data for manual review
    if (!extractedData || (!extractedData.name && !extractedData.phone && !extractedData.email && !extractedData.address)) {
      extractedData = {
        date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
        stage: '1. F/U',
        type: 'Res',
        regarding: emailData.subject || 'NEEDS MANUAL REVIEW'
      };
      parseMethod = 'Manual Review';
    }
    
    // Build CSV string (29 columns)
    const csvString = er_buildCSV_(extractedData);
    
    // Write CSV to column A of new row
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;

    sheet.getRange(newRow, 1).setValue(csvString);
    SpreadsheetApp.flush();

    // DO THE SPLIT OURSELVES (Stage Automation won't trigger on programmatic writes)
    sheet.getRange(newRow, 1).splitTextToColumns(SpreadsheetApp.TextToColumnsDelimiter.COMMA);
    SpreadsheetApp.flush();
    
    // Now add the email link to column B
    const linkCell = sheet.getRange(newRow, C.COLS.LINK_B);
    const sourceLabel = emailData.isRuby ? '[Ruby]' : '[Add lead]';
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(sourceLabel)
      .setLinkUrl(gmailUrl)
      .build();
    linkCell.setRichTextValue(richText);
    
    // If manual review needed, highlight row and add note
    if (parseMethod === 'Manual Review') {
      sheet.getRange(newRow, 1, 1, C.TOTAL_COLUMNS).setBackground('#ffebee');
      sheet.getRange(newRow, C.COLS.COMMENTS).setNote(
        `Email Content:\n\n${emailData.body.substring(0, 500)}${emailData.body.length > 500 ? '...' : ''}`
      );
    }
    
    // ALWAYS add "LeadProcessed" label after processing
    thread.addLabel(processedLabel);
    er_log_('Added LeadProcessed label', { messageId: emailData.messageId });
    
    // Remove "Add lead" label if present (only for non-Ruby emails)
    if (!emailData.isRuby) {
      try {
        const addLeadLabel = GmailApp.getUserLabelByName(C.ADD_LEAD_LABEL);
        if (addLeadLabel) {
          const labels = thread.getLabels();
          const hasAddLeadLabel = labels.some(l => l.getName() === C.ADD_LEAD_LABEL);
          if (hasAddLeadLabel) {
            thread.removeLabel(addLeadLabel);
            er_log_('Removed "Add lead" label', { messageId: emailData.messageId });
          }
        }
      } catch (err) {
        er_log_('Could not remove "Add lead" label', { 
          error: err.toString(),
          messageId: emailData.messageId 
        });
      }
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Email processed successfully', {
        row: newRow,
        method: parseMethod,
        sourceType: emailData.isRuby ? 'Ruby' : 'Add lead',
        csvPreview: csvString.substring(0, 100)
      });
    }
    
    return true;
    
  } catch (err) {
    er_log_('Email processing error', { 
      error: err.toString(),
      stack: err.stack 
    });
    throw err;
  }
}

/**
 * Extract data using pattern matching
 */
function er_extractWithPatterns_(emailData) {
  const C = EMAIL_READER_CONFIG;
  const body = emailData.body;
  
  const extracted = {
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
    stage: '1. F/U',
    type: 'Res',
    name: '',
    company: '',
    phone: '',
    email: '',
    address: '',
    regarding: ''
  };
  
  // Try Ruby-specific patterns first
  const P = C.RUBY_PATTERNS;
  
  let firstName = '';
  let lastName = '';
  
  const firstMatch = body.match(P.first);
  if (firstMatch) firstName = firstMatch[1].trim();
  
  const lastMatch = body.match(P.last);
  if (lastMatch) lastName = lastMatch[1].trim();
  
  extracted.name = `${firstName} ${lastName}`.trim();
  
  const phoneMatch = body.match(P.phone);
  if (phoneMatch) {
    const digits = phoneMatch[1].trim().replace(/\D/g, '');
    if (digits.length === 10) {
      extracted.phone = digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    } else if (digits.length === 11 && digits.startsWith('1')) {
      extracted.phone = digits.substring(1).replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    }
  }
  
  const companyMatch = body.match(P.company);
  if (companyMatch) extracted.company = companyMatch[1].trim();
  
  const emailMatch = body.match(P.email);
  if (emailMatch) {
    // Extract just the email address
    const emailAddr = emailMatch[1].match(C.GENERIC_PATTERNS.email);
    if (emailAddr) extracted.email = emailAddr[0];
  }
  
  const addressMatch = body.match(P.address);
  if (addressMatch) {
    // Remove commas from address!
    extracted.address = addressMatch[1].trim().replace(/,/g, ' ').replace(/\s+/g, ' ');
  }
  
  const regardingMatch = body.match(P.regarding);
  if (regardingMatch) extracted.regarding = regardingMatch[1].trim();
  
  // If Ruby patterns didn't find phone, try generic
  if (!extracted.phone) {
    const genericPhone = body.match(C.GENERIC_PATTERNS.phone);
    if (genericPhone) {
      const digits = genericPhone[0].replace(/\D/g, '');
      if (digits.length === 10) {
        extracted.phone = digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
      } else if (digits.length === 11) {
        extracted.phone = digits.substring(1).replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
      }
    }
  }
  
  // If no email found, try generic
  if (!extracted.email) {
    const genericEmail = body.match(C.GENERIC_PATTERNS.email);
    if (genericEmail) extracted.email = genericEmail[0];
  }
  
  // If no address found, try generic
  if (!extracted.address) {
    const genericAddress = body.match(C.GENERIC_PATTERNS.address);
    if (genericAddress) {
      extracted.address = genericAddress[0].replace(/,/g, ' ').replace(/\s+/g, ' ');
    }
  }
  
  // If no regarding, use subject
  if (!extracted.regarding) {
    extracted.regarding = emailData.subject || '';
  }
  
  // If no name found, try to get from "From" field
  if (!extracted.name) {
    const fromMatch = emailData.from.match(/^([^<]+)</);
    if (fromMatch) {
      extracted.name = fromMatch[1].trim();
    }
  }
  
  return extracted;
}

/**
 * Extract data using structured key:value parsing
 */
function er_extractStructured_(emailData) {
  const C = EMAIL_READER_CONFIG;
  const body = emailData.body;
  
  const extracted = {
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
    stage: '1. F/U',
    type: 'Res',
    name: '',
    company: '',
    phone: '',
    email: '',
    address: '',
    regarding: ''
  };
  
  const lines = body.split('\n');
  const data = {};
  
  for (const line of lines) {
    const colonIndex = line.indexOf(':');
    if (colonIndex > 0 && colonIndex < 30) {
      const key = line.substring(0, colonIndex).trim().toLowerCase();
      const value = line.substring(colonIndex + 1).trim();
      if (value) data[key] = value;
    }
  }
  
  // Map extracted data
  const firstName = data['first'] || '';
  const lastName = data['last'] || '';
  extracted.name = `${firstName} ${lastName}`.trim() || data['name'] || data['contact'] || '';
  
  extracted.company = data['company'] || '';
  
  // Clean phone
  const rawPhone = data['phone number'] || data['phone'] || '';
  const phoneDigits = rawPhone.replace(/\D/g, '');
  if (phoneDigits.length === 10) {
    extracted.phone = phoneDigits.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
  } else if (phoneDigits.length === 11 && phoneDigits.startsWith('1')) {
    extracted.phone = phoneDigits.substring(1).replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
  }
  
  // Clean email
  const rawEmail = data['email'] || '';
  const emailMatch = rawEmail.match(C.GENERIC_PATTERNS.email);
  extracted.email = emailMatch ? emailMatch[0] : '';
  
  // If no email in fields, try body
  if (!extracted.email) {
    const bodyEmail = emailData.body.match(C.GENERIC_PATTERNS.email);
    if (bodyEmail) extracted.email = bodyEmail[0];
  }
  
  // Address (remove commas!)
  extracted.address = (data['project address'] || data['address'] || '').replace(/,/g, ' ').replace(/\s+/g, ' ');
  
  extracted.regarding = data['regarding'] || emailData.subject || '';
  
  // If no name, try from field
  if (!extracted.name) {
    const fromMatch = emailData.from.match(/^([^<]+)</);
    if (fromMatch) {
      extracted.name = fromMatch[1].trim();
    }
  }
  
  return extracted;
}

/**
 * Logging function
 */
function er_log_(operation, details) {
  if (!EMAIL_READER_CONFIG.ENABLE_LOGGING) return;
  try {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    console.log(`[${timestamp}] [EmailReader] ${operation}:`, JSON.stringify(details));
  } catch (err) {
    console.log(`[EmailReader] Log error:`, err.message);
  }
}

/**
 * Install trigger - runs every 15 minutes for BOTH Ruby and "Add lead"
 */
function er_installTrigger() {
  er_removeTrigger();
  
  ScriptApp.newTrigger('er_processNewEmails')
    .timeBased()
    .everyMinutes(15)
    .create();
  
  SpreadsheetApp.getActive().toast(
    'Auto-check installed (every 15 min)\nRuby + Add lead emails',
    '✅ Email Reader',
    5
  );
}

/**
 * Remove trigger
 */
function er_removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'er_processNewEmails' ||
        trigger.getHandlerFunction() === 'er_processRubyOnly') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * MAIN: Process ALL pending emails (Ruby + Add lead) - runs automatically
 */
function er_processNewEmails() {
  const C = EMAIL_READER_CONFIG;
  
  try {
    // Ensure labels exist
    let processedLabel = GmailApp.getUserLabelByName(C.PROCESSED_LABEL);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(C.PROCESSED_LABEL);
    }
    
    let addLeadLabel = GmailApp.getUserLabelByName(C.ADD_LEAD_LABEL);
    if (!addLeadLabel) {
      addLeadLabel = GmailApp.createLabel(C.ADD_LEAD_LABEL);
    }
    
    let totalProcessed = 0;
    let rubyProcessed = 0;
    let addLeadProcessed = 0;
    
    for (const search of C.SEARCHES) {
      if (!search.enabled) continue;
      
      const threads = GmailApp.search(search.query, 0, C.MAX_EMAILS_PER_RUN);
      
      er_log_('Search completed', { 
        name: search.name,
        query: search.query,
        threadsFound: threads.length 
      });
      
      for (const thread of threads) {
        const messages = thread.getMessages();
        for (const message of messages) {
          // For Ruby search, only process Ruby emails
          if (search.sourceType === 'ruby' && !message.getFrom().includes('noreply@ruby.com')) {
            continue;
          }
          
          const success = er_processEmail_(message, search, processedLabel);
          if (success) {
            totalProcessed++;
            if (search.sourceType === 'ruby') {
              rubyProcessed++;
            } else {
              addLeadProcessed++;
            }
          }
        }
      }
    }
    
    if (totalProcessed > 0) {
      let toastMsg = `✅ Processed ${totalProcessed} email(s)`;
      if (rubyProcessed > 0) toastMsg += `\n• ${rubyProcessed} Ruby`;
      if (addLeadProcessed > 0) toastMsg += `\n• ${addLeadProcessed} Add lead`;
      
      SpreadsheetApp.getActive().toast(toastMsg, 'Email Reader', 5);
    }
    
    er_log_('Batch complete', { totalProcessed, rubyProcessed, addLeadProcessed });
    
    return totalProcessed;
    
  } catch (err) {
    er_log_('Batch processing error', { error: err.toString() });
    SpreadsheetApp.getActive().toast(`Error: ${err.message}`, 'Email Reader Error', 5);
    throw err;
  }
}

/**
 * Process only "Add lead" emails manually
 */
function er_processAddLeadEmails() {
  const C = EMAIL_READER_CONFIG;
  
  try {
    let processedLabel = GmailApp.getUserLabelByName(C.PROCESSED_LABEL);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(C.PROCESSED_LABEL);
    }
    
    const searchConfig = C.SEARCHES.find(s => s.sourceType === 'addlead');
    if (!searchConfig) {
      SpreadsheetApp.getActive().toast('Add lead search not configured', 'Error', 5);
      return;
    }
    
    const threads = GmailApp.search(searchConfig.query, 0, C.MAX_EMAILS_PER_RUN);
    
    if (threads.length === 0) {
      SpreadsheetApp.getActive().toast('No "Add lead" emails found', 'Info', 3);
      return;
    }
    
    let processed = 0;
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      for (const message of messages) {
        const success = er_processEmail_(message, searchConfig, processedLabel);
        if (success) processed++;
      }
    }
    
    SpreadsheetApp.getActive().toast(`✅ Processed ${processed} "Add lead" email(s)`, 'Complete', 5);
    
  } catch (err) {
    SpreadsheetApp.getActive().toast(`Error: ${err.message}`, 'Error', 5);
  }
}

/**
 * Test processing - no popup, just runs and shows toast
 */
function er_testProcessing() {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const rubyThreads = GmailApp.search('from:noreply@ruby.com -label:LeadProcessed', 0, 1);
    const addLeadThreads = GmailApp.search(`label:"${C.ADD_LEAD_LABEL}" -label:LeadProcessed`, 0, 1);
    
    let testThread = null;
    let sourceType = '';
    
    if (rubyThreads.length > 0) {
      testThread = rubyThreads[0];
      sourceType = 'Ruby';
    } else if (addLeadThreads.length > 0) {
      testThread = addLeadThreads[0];
      sourceType = 'Add lead';
    }
    
    if (!testThread) {
      SpreadsheetApp.getActive().toast(
        'No unprocessed emails found.\nLabel an email with "Add lead" to test.',
        'No Emails',
        5
      );
      return;
    }
    
    let processedLabel = GmailApp.getUserLabelByName(C.PROCESSED_LABEL);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(C.PROCESSED_LABEL);
    }
    
    const message = testThread.getMessages()[0];
    const searchConfig = sourceType === 'Ruby' ? C.SEARCHES[0] : C.SEARCHES[1];
    
    SpreadsheetApp.getActive().toast(
      `Processing ${sourceType} email:\n${message.getSubject().substring(0, 40)}...`,
      'Testing',
      3
    );
    
    er_processEmail_(message, searchConfig, processedLabel);
    
    SpreadsheetApp.getActive().toast(
      `✅ Test complete!\nCheck Leads sheet for new row.\nEmail labeled "LeadProcessed"`,
      'Success',
      5
    );
    
  } catch (err) {
    SpreadsheetApp.getActive().toast(`Error: ${err.message}`, 'Test Failed', 5);
  }
}
/**
 * TRIGGER HEALTH CHECK - Emails you if triggers are missing
 * Version: 01/20-09:20AM EST by Claude Opus 4.1
 */

const HEALTH_CHECK_CONFIG = {
  EXPECTED_TRIGGERS: [
    'masterOnEditHandler_',    // Master trigger handles all onEdit events
    'er_processNewEmails',
    'checkEmptyFoldersDaily_',
    'runMileageSync_',
    'checkTriggerHealth_'
  ],
  ALERT_EMAIL: Session.getActiveUser().getEmail(),
  SUBJECT: '⚠️ Walker Awning: Trigger Alert'
};

/**
 * Check if all expected triggers are installed - runs daily
 * Only sends email if something is WRONG
 */
function checkTriggerHealth_() {
  const installed = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const missing = HEALTH_CHECK_CONFIG.EXPECTED_TRIGGERS.filter(t => !installed.includes(t));
  
  if (missing.length === 0) {
    // All good - no email needed
    er_log_('Trigger health check passed', { installed: installed.length });
    return;
  }
  
  // Something is missing - send alert email
  const body = `
One or more automation triggers are missing from your Walker Awning spreadsheet.

MISSING TRIGGERS:
${missing.map(t => '  ❌ ' + t).join('\n')}

CURRENTLY INSTALLED:
${installed.map(t => '  ✅ ' + t).join('\n')}

HOW TO FIX:
1. Open your Google Sheet
2. Go to the appropriate Setup menu and reinstall the trigger:
   - handleEditDraft_V2 → Setup (Drafts) → Install On-Edit Trigger
   - handleEditMove_ → Setup (Move) → Install On-Edit Trigger
   - handleEditAwningRuby_ → Setup (Ruby) → Install On-Edit Trigger
   - er_processNewEmails → Email Reader → Setup Auto-Check
   - checkEmptyFoldersDaily_ → Setup (Move) → Install Daily Folder Check

This is an automated alert from your Walker Awning automation system.
  `.trim();
  
  MailApp.sendEmail({
    to: HEALTH_CHECK_CONFIG.ALERT_EMAIL,
    subject: HEALTH_CHECK_CONFIG.SUBJECT,
    body: body
  });
  
  er_log_('Trigger health alert sent', { missing, to: HEALTH_CHECK_CONFIG.ALERT_EMAIL });
}

/**
 * Install the daily health check trigger (runs at 7 AM)
 */
function installTriggerHealthCheck_() {
  // Remove existing health check triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkTriggerHealth_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // Install new trigger at 7 AM daily
  ScriptApp.newTrigger('checkTriggerHealth_')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
  
  SpreadsheetApp.getActive().toast(
    'Daily trigger health check installed (7 AM)',
    '✅ Health Check',
    5
  );
}

/**
 * Remove the health check trigger
 */
function removeTriggerHealthCheck_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'checkTriggerHealth_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  SpreadsheetApp.getActive().toast(
    'Health check trigger removed',
    'Removed',
    3
  );
}

/**
 * Test the health check (runs immediately, will email if triggers missing)
 */
function testTriggerHealthCheck_() {
  checkTriggerHealth_();
  SpreadsheetApp.getActive().toast(
    'Health check complete - email sent only if triggers missing',
    'Test Complete',
    5
  );
}