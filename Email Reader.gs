/**
 * EMAIL READER AUTOMATION - UNIFIED LEAD INGESTION
 * Version# [12/31-07:45PM EST] by Claude Opus 4.1
 * 
 * FEATURES:
 * - Ruby emails: Automatically detected and processed
 * - "Add lead" emails: Automatically processed (same schedule as Ruby)
 * - Both use identical AI prompt for consistent results
 * - ALL processed emails get "LeadProcessed" label
 * - Uses LABELS (not read/unread) to prevent duplicates
 * - Three-tier fallback system (AI → Pattern → Structured → Manual Review)
 * 
 * CHANGES IN THIS VERSION:
 * - Removed diagnostic dialog popup
 * - Diagnostic now just shows toast + logs to console
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
        autoSplit: true,
        markAsRead: false,
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
        autoSplit: false,
        markAsRead: false,
        useAI: true,
        useFallback: true
      }
    }
  ],
  
  PROCESSED_LABEL: 'LeadProcessed',
  ADD_LEAD_LABEL: 'Add lead',
  
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
    LINK_L: 12,        // L
    JOB_DESC_M: 13,    // M
    QUOTE: 14,         // N
    CALCS: 15,         // O
    QB_URL: 16,        // P
    EARTH_LINK: 17,    // Q
    JOB_TYPE: 18,      // R
    MISC_S: 19,        // S
    LENGTH: 20,        // T
    WIDTH: 21          // U
  },
  
  // UNIFIED AI FORMULA - Used for BOTH Ruby emails AND "Add lead" emails
  AI_FORMULA: `=AI("Parse this email and extract lead information. Extract: First name, Last name, Phone, Company, Email, Address, Regarding/Subject, Actions/Notes. Return exactly 21 comma-separated values: Date(MM/DD), blank, blank, '1. F/U', Full Name, Company, 'Res', Phone, Email, Address(no commas), Regarding, Actions, then 9 blanks. If a field is not found, leave it blank but keep the comma. Example: 01/20,,,1. F/U,John Smith,ABC Corp,Res,954-555-1234,john@email.com,123 Main St,Quote request,Call back,,,,,,,,",B{ROW})`,

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
    email: /[^\s@]+@[^\s@]+\.[^\s@]+/,
    address: /(\d+\s+[a-z\s]+(?:st|street|ave|avenue|rd|road|ln|lane|blvd|boulevard|dr|drive|ct|court))/i
  }
};

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
 * Process a single email - SAME PROCESS for Ruby and "Add lead"
 * ALWAYS adds "LeadProcessed" label after processing
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
    const emailContent = emailData.body;
    
    if (C.ENABLE_LOGGING) {
      er_log_('Processing email', {
        from: emailData.from,
        subject: emailData.subject,
        sourceType: emailData.isRuby ? 'Ruby' : 'Add lead',
        bodyLength: emailContent.length,
        messageId: emailData.messageId
      });
    }
    
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(C.TARGET_SHEET);
    
    if (!sheet) {
      throw new Error(`Sheet "${C.TARGET_SHEET}" not found`);
    }
    
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    // Write email content to column B for AI to parse
    const contentCell = sheet.getRange(newRow, C.COLS.LINK_B);
    contentCell.setValue(emailContent);
    SpreadsheetApp.flush();
    
    let parseSuccess = false;
    let parseMethod = 'none';
    
    // Try AI parsing first
    if (searchConfig.settings.useAI) {
      parseSuccess = er_tryAIParsing_(sheet, newRow, emailData, gmailUrl);
      if (parseSuccess) parseMethod = 'AI';
    }
    
    // Try pattern parsing if AI failed
    if (!parseSuccess && searchConfig.settings.useFallback) {
      parseSuccess = er_tryPatternParsing_(sheet, newRow, emailData, gmailUrl);
      if (parseSuccess) parseMethod = 'Pattern';
    }
    
    // Try structured extraction as last resort
    if (!parseSuccess) {
      parseSuccess = er_tryStructuredExtraction_(sheet, newRow, emailData, gmailUrl);
      if (parseSuccess) parseMethod = 'Structured';
    }
    
    // If still failed, save for manual review
    if (!parseSuccess) {
      er_saveForManualReview_(sheet, newRow, emailData, gmailUrl);
      parseMethod = 'Manual Review';
    }
    
    // ALWAYS add "LeadProcessed" label after processing
    thread.addLabel(processedLabel);
    er_log_('Added LeadProcessed label', { messageId: emailData.messageId });
    
    // Remove "Add lead" label if present
    try {
      const addLeadLabel = GmailApp.getUserLabelByName(C.ADD_LEAD_LABEL);
      if (addLeadLabel && thread.hasLabel(addLeadLabel)) {
        thread.removeLabel(addLeadLabel);
        er_log_('Removed "Add lead" label', { messageId: emailData.messageId });
      }
    } catch (err) {
      er_log_('Could not remove "Add lead" label', { 
        error: err.toString(),
        messageId: emailData.messageId 
      });
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Email processed successfully', {
        row: newRow,
        method: parseMethod,
        sourceType: emailData.isRuby ? 'Ruby' : 'Add lead',
        labelAdded: 'LeadProcessed'
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
 * Try AI parsing - Same formula for all emails
 */
function er_tryAIParsing_(sheet, row, emailData, gmailUrl) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const formula = C.AI_FORMULA.replace('{ROW}', row);
    const formulaCell = sheet.getRange(row, C.COLS.DATE);
    
    er_log_('Setting AI formula', { 
      row, 
      sourceType: emailData.isRuby ? 'Ruby' : 'Add lead'
    });
    
    formulaCell.setFormula(formula);
    SpreadsheetApp.flush();
    
    for (let attempt = 1; attempt <= C.MAX_AI_RETRIES; attempt++) {
      Utilities.sleep(C.WAIT_FOR_AI_MS);
      
      const result = formulaCell.getValue();
      
      er_log_('AI attempt result', { 
        row, 
        attempt, 
        resultLength: String(result).length,
        resultPreview: String(result).substring(0, 100)
      });
      
      if (result && String(result).trim() !== '' && !String(result).includes('#ERROR') && !String(result).includes('Loading')) {
        const values = String(result).split(',');
        
        if (values.length >= 12) {
          // Pad to 21 columns
          while (values.length < 21) values.push('');
          if (values.length > 21) values.length = 21;
          
          const cleanedValues = values.map(v => 
            String(v).trim()
              .replace(/^["']|["']$/g, '')
              .replace(/\s+/g, ' ')
          );
          
          const targetRange = sheet.getRange(row, 1, 1, 21);
          targetRange.setValues([cleanedValues]);
          
          // Create link in column B
          const linkCell = sheet.getRange(row, C.COLS.LINK_B);
          const sourceLabel = emailData.isRuby ? '[Ruby]' : '[Add lead]';
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(sourceLabel)
            .setLinkUrl(gmailUrl)
            .build();
          linkCell.setRichTextValue(richText);
          
          er_log_('AI parsing successful', { 
            row, 
            attempt,
            sourceType: emailData.isRuby ? 'Ruby' : 'Add lead',
            fieldsFound: cleanedValues.filter(v => v).length 
          });
          
          return true;
        }
      }
      
      if (attempt < C.MAX_AI_RETRIES) {
        Utilities.sleep(3000);
      }
    }
    
    er_log_('AI parsing failed after all retries', { row });
    return false;
    
  } catch (err) {
    er_log_('AI parsing error', { row, error: err.toString() });
    return false;
  }
}

/**
 * Pattern parsing - fallback for both Ruby and Add lead
 */
function er_tryPatternParsing_(sheet, row, emailData, gmailUrl) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const body = emailData.body;
    const isRuby = emailData.isRuby;
    
    const extracted = {
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
      first: '',
      last: '',
      phone: '',
      company: '',
      email: '',
      address: '',
      regarding: '',
      actions: ''
    };
    
    // Try Ruby-specific patterns first
    const P = C.RUBY_PATTERNS;
    
    const firstMatch = body.match(P.first);
    if (firstMatch) extracted.first = firstMatch[1].trim();
    
    const lastMatch = body.match(P.last);
    if (lastMatch) extracted.last = lastMatch[1].trim();
    
    const phoneMatch = body.match(P.phone);
    if (phoneMatch) {
      extracted.phone = phoneMatch[1].trim().replace(/\D/g, '').replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
    }
    
    const companyMatch = body.match(P.company);
    if (companyMatch) extracted.company = companyMatch[1].trim();
    
    const emailMatch = body.match(P.email);
    if (emailMatch) extracted.email = emailMatch[1].trim();
    
    const addressMatch = body.match(P.address);
    if (addressMatch) {
      extracted.address = addressMatch[1].trim().replace(/,/g, ' ').replace(/\s+/g, ' ');
    }
    
    const regardingMatch = body.match(P.regarding);
    if (regardingMatch) extracted.regarding = regardingMatch[1].trim();
    
    const actionsMatch = body.match(P.actions);
    if (actionsMatch) extracted.actions = actionsMatch[1].trim();
    
    // If Ruby patterns didn't find much, try generic patterns
    if (!extracted.phone) {
      const genericPhone = body.match(C.GENERIC_PATTERNS.phone);
      if (genericPhone) {
        const digits = genericPhone[0].replace(/\D/g, '');
        if (digits.length === 10) {
          extracted.phone = digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
        } else if (digits.length === 11) {
          extracted.phone = digits.replace(/(\d{1})(\d{3})(\d{3})(\d{4})/, '$2-$3-$4');
        }
      }
    }
    
    if (!extracted.email) {
      const genericEmail = body.match(C.GENERIC_PATTERNS.email);
      if (genericEmail) extracted.email = genericEmail[0];
    }
    
    if (!extracted.address) {
      const genericAddress = body.match(C.GENERIC_PATTERNS.address);
      if (genericAddress) {
        extracted.address = genericAddress[0].replace(/,/g, ' ').replace(/\s+/g, ' ');
      }
    }
    
    if (!extracted.regarding) {
      extracted.regarding = emailData.subject;
    }
    
    const fullName = `${extracted.first} ${extracted.last}`.trim();
    
    const rowData = [
      extracted.date,
      '',
      '',
      '1. F/U',
      fullName,
      extracted.company,
      'Res',
      extracted.phone,
      extracted.email,
      extracted.address,
      extracted.regarding,
      extracted.actions,
      '', '', '', '', '', '', '', '', ''
    ];
    
    const hasData = fullName || extracted.phone || extracted.email || extracted.address;
    
    if (hasData) {
      const targetRange = sheet.getRange(row, 1, 1, 21);
      targetRange.setValues([rowData]);
      
      const linkCell = sheet.getRange(row, C.COLS.LINK_B);
      const sourceLabel = isRuby ? '[Ruby]' : '[Add lead]';
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(sourceLabel)
        .setLinkUrl(gmailUrl)
        .build();
      linkCell.setRichTextValue(richText);
      
      er_log_('Pattern parsing successful', { 
        row, 
        sourceType: isRuby ? 'Ruby' : 'Add lead',
        name: fullName
      });
      
      return true;
    }
    
    return false;
    
  } catch (err) {
    er_log_('Pattern parsing error', { row, error: err.toString() });
    return false;
  }
}

/**
 * Structured extraction (generic fallback)
 */
function er_tryStructuredExtraction_(sheet, row, emailData, gmailUrl) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const lines = emailData.body.split('\n');
    const data = {};
    
    for (const line of lines) {
      const colonIndex = line.indexOf(':');
      if (colonIndex > 0 && colonIndex < 30) {
        const key = line.substring(0, colonIndex).trim().toLowerCase();
        const value = line.substring(colonIndex + 1).trim();
        if (value) data[key] = value;
      }
    }
    
    const firstName = data['first'] || '';
    const lastName = data['last'] || '';
    const fullName = `${firstName} ${lastName}`.trim();
    
    const rowData = [
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
      '',
      '',
      '1. F/U',
      fullName || data['name'] || data['contact'] || '',
      data['company'] || '',
      'Res',
      (data['phone number'] || data['phone'] || '').replace(/\D/g, '').replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3'),
      data['email'] || '',
      (data['project address'] || data['address'] || '').replace(/,/g, ' '),
      data['regarding'] || emailData.subject || '',
      data['actions'] || '',
      '', '', '', '', '', '', '', '', ''
    ];
    
    if (rowData[4] || rowData[7] || rowData[8]) {
      const targetRange = sheet.getRange(row, 1, 1, 21);
      targetRange.setValues([rowData]);
      
      const linkCell = sheet.getRange(row, C.COLS.LINK_B);
      const sourceLabel = emailData.isRuby ? '[Ruby]' : '[Add lead]';
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(sourceLabel)
        .setLinkUrl(gmailUrl)
        .build();
      linkCell.setRichTextValue(richText);
      
      return true;
    }
    
    return false;
    
  } catch (err) {
    er_log_('Structured extraction error', { row, error: err.toString() });
    return false;
  }
}

/**
 * Save for manual review when all parsing fails
 */
function er_saveForManualReview_(sheet, row, emailData, gmailUrl) {
  const C = EMAIL_READER_CONFIG;
  const sourceLabel = emailData.isRuby ? 'Ruby' : 'Add lead';
  
  const rowData = [
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
    '',
    `⚠️ NEEDS MANUAL PARSING - ${sourceLabel}`,
    '1. F/U',
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
  ];
  
  const targetRange = sheet.getRange(row, 1, 1, 21);
  targetRange.setValues([rowData]);
  
  const linkCell = sheet.getRange(row, C.COLS.LINK_B);
  const richText = SpreadsheetApp.newRichTextValue()
    .setText('[Click to View Email]')
    .setLinkUrl(gmailUrl)
    .build();
  linkCell.setRichTextValue(richText);
  
  sheet.getRange(row, 1, 1, 21).setBackground('#ffebee');
  
  sheet.getRange(row, C.COLS.COMMENTS).setNote(
    `${sourceLabel} Email Content:\n\n${emailData.body.substring(0, 500)}${emailData.body.length > 500 ? '...' : ''}`
  );
  
  er_log_('Saved for manual review', { row, sourceType: sourceLabel });
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