/**
 * EMAIL READER AUTOMATION - UNIFIED LEAD INGESTION (VERSION 3.0.1)
 * Version: 01/20-Enhanced with separate stricter AI for manual leads
 * 
 * FEATURES:
 * - Ruby emails: UNCHANGED - uses proven working AI formula
 * - Manual leads: NEW stricter AI prompt that prefers conservative extraction
 * - Processes any email labeled "add lead"
 * - Format-aware parsing (Ruby-specific vs. generic)
 * - Three-tier fallback system
 * - Comprehensive diagnostics
 * - Automatic label cleanup
 * - Better error handling and logging
 */

const EMAIL_READER_CONFIG = {
  TARGET_SHEET: 'Leads',
  
  SEARCHES: [
    {
      name: 'Ruby Mail',
      query: 'from:noreply@ruby.com is:unread',
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
      name: 'Manual Leads',
      query: 'label:add lead is:unread',
      enabled: true,
      sourceType: 'manual',
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
  ADD_LEAD_LABEL: 'add lead',
  
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
  
  // AI formula for Ruby emails - PROVEN WORKING - DO NOT MODIFY
  AI_FORMULA_RUBY: `=AI("This is a Ruby answering service email. Parse this Ruby email and extract: First name, Last name, Phone, Company, Email, Address, Regarding, Actions. Return exactly 21 comma-separated values: Date(MM/DD), blank, blank, '1. F/U', Full Name, Company, 'Res', Phone, Email, Address(no commas), Regarding, Actions, then 9 blanks. Example: 01/20,,,1. F/U,John Smith,ABC Corp,Res,954-555-1234,john@email.com,123 Main St,Quote request,Call back,,,,,,,,",B{ROW})`,
  
  // AI formula for manual leads - CONSERVATIVE extraction, prefers empty over wrong
  AI_FORMULA_GENERIC: `=AI("Extract business contact info from email. Be CONSERVATIVE - only extract fields clearly present. Return exactly 21 comma-separated values: Date(MM/DD), blank, blank, '1. F/U', Contact Name (from signature/From field only), Company, 'Res', Phone (###-###-####), Email, Address (no commas), Subject Line, Notes, then 9 blanks. RULE: Leave field empty if not certain. For Contact Name only use names found in signature blocks or sender info - never use greetings. Example: 01/20,,,1. F/U,Jane Doe,ABC Company,Res,954-555-1234,jane@email.com,456 Oak Ave,Need quote on canopy project,,,,,,,,",B{ROW})`,

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
  
  // Generic patterns (for manual leads)
  GENERIC_PATTERNS: {
    phone: /(\+?1[-.\s]?)?(\()?(\d{3})(\))?[-.\s]?(\d{3})[-.\s]?(\d{4})/,
    email: /[^\s@]+@[^\s@]+\.[^\s@]+/,
    address: /(\d+\s+[a-z\s]+(?:st|street|ave|avenue|rd|road|ln|lane|blvd|boulevard|dr|drive|ct|court))/i
  }
};

/**
 * DIAGNOSTIC FUNCTION - Run this first to check setup
 */
function er_diagnosticCheck() {
  const ui = SpreadsheetApp.getUi();
  const C = EMAIL_READER_CONFIG;
  const results = [];
  
  // 1. Check if Leads sheet exists
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(C.TARGET_SHEET);
  results.push(sheet ? '✅ Leads sheet found' : '❌ Leads sheet NOT found');
  
  // 2. Check for Ruby emails
  try {
    const threads = GmailApp.search('from:noreply@ruby.com', 0, 1);
    results.push(threads.length > 0 ? `✅ Found ${threads.length} Ruby email(s)` : '⚠️ No Ruby emails found');
  } catch (err) {
    results.push('❌ Gmail search error: ' + err.message);
  }
  
  // 3. Check for add lead label
  try {
    const threads = GmailApp.search(`label:${C.ADD_LEAD_LABEL}`, 0, 1);
    results.push(threads.length > 0 ? `✅ Found ${threads.length} email(s) with "add lead" label` : '⚠️ No "add lead" labeled emails');
  } catch (err) {
    results.push('❌ Add lead label search error: ' + err.message);
  }
  
  // 4. Check labels
  try {
    let processedLabel = GmailApp.getUserLabelByName(C.PROCESSED_LABEL);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(C.PROCESSED_LABEL);
      results.push('✅ Created LeadProcessed label');
    } else {
      results.push('✅ LeadProcessed label exists');
    }
    
    let addLeadLabel = GmailApp.getUserLabelByName(C.ADD_LEAD_LABEL);
    if (!addLeadLabel) {
      addLeadLabel = GmailApp.createLabel(C.ADD_LEAD_LABEL);
      results.push('✅ Created add_lead label');
    } else {
      results.push('✅ add lead label exists');
    }
  } catch (err) {
    results.push('❌ Label error: ' + err.message);
  }
  
  // 5. Check triggers
  const triggers = ScriptApp.getProjectTriggers();
  const emailReaderTrigger = triggers.find(t => t.getHandlerFunction() === 'er_processNewEmails');
  results.push(emailReaderTrigger ? '✅ Trigger installed' : '⚠️ No trigger found (run Setup Auto-Check)');
  
  // 6. Check unread emails by source
  try {
    const rubyThreads = GmailApp.search('from:noreply@ruby.com is:unread', 0, 5);
    const manualThreads = GmailApp.search(`label:${C.ADD_LEAD_LABEL} is:unread`, 0, 5);
    results.push(`📧 Found ${rubyThreads.length} UNREAD Ruby email(s)`);
    results.push(`📧 Found ${manualThreads.length} UNREAD "add lead" email(s)`);
  } catch (err) {
    results.push('❌ Unread search error: ' + err.message);
  }
  
  // Show results
  ui.alert('Email Reader Diagnostics', results.join('\n'), ui.ButtonSet.OK);
  
  er_log_('Diagnostic check completed', { results });
}

/**
 * Determine if email is from Ruby
 */
function er_isRubyEmail_(message) {
  const from = message.getFrom().toLowerCase();
  return from.includes('noreply@ruby.com');
}

/**
 * Process a single email with source awareness
 */
function er_processEmail_(message, searchConfig, processedLabel) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    // Check if already processed
    const labels = message.getThread().getLabels();
    if (labels.some(l => l.getName() === C.PROCESSED_LABEL)) {
      er_log_('Email already processed (has label)', { messageId: message.getId() });
      return;
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
        sourceType: emailData.isRuby ? 'Ruby' : 'Manual',
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
    
    // Try AI parsing first (format-aware)
    if (searchConfig.settings.useAI) {
      parseSuccess = er_tryAIParsing_(sheet, newRow, emailData, gmailUrl, searchConfig);
      if (parseSuccess) parseMethod = 'AI';
    }
    
    // Try pattern parsing if AI failed
    if (!parseSuccess && searchConfig.settings.useFallback) {
      parseSuccess = er_tryPatternParsing_(sheet, newRow, emailData, gmailUrl, searchConfig);
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
    
    // Mark as processed
    const thread = message.getThread();
    thread.addLabel(processedLabel);
    
    // Remove "add lead" label after processing (applies to all leads)
    // This is done regardless of source type for clean label management
    try {
      const addLeadLabel = GmailApp.getUserLabelByName(C.ADD_LEAD_LABEL);
      if (addLeadLabel && thread.hasLabel(addLeadLabel)) {
        thread.removeLabel(addLeadLabel);
        er_log_('Removed add lead label', { 
          messageId: emailData.messageId,
          sourceType: emailData.isRuby ? 'Ruby' : 'Manual'
        });
      }
    } catch (err) {
      er_log_('Could not remove add lead label', { 
        error: err.toString(),
        messageId: emailData.messageId 
      });
    }
    
    if (searchConfig.settings.markAsRead) {
      thread.markRead();
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Email processed successfully', {
        row: newRow,
        method: parseMethod,
        success: parseSuccess,
        sourceType: emailData.isRuby ? 'Ruby' : 'Manual'
      });
    }
    
  } catch (err) {
    er_log_('Email processing error', { 
      error: err.toString(),
      stack: err.stack 
    });
    throw err;
  }
}

/**
 * Try AI parsing with format awareness
 */
function er_tryAIParsing_(sheet, row, emailData, gmailUrl, searchConfig) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    // Choose formula based on source type
    const isRuby = emailData.isRuby;
    const formulaTemplate = isRuby ? C.AI_FORMULA_RUBY : C.AI_FORMULA_GENERIC;
    const formula = formulaTemplate.replace('{ROW}', row);
    
    const formulaCell = sheet.getRange(row, C.COLS.DATE);
    
    er_log_('Setting AI formula', { 
      row, 
      sourceType: isRuby ? 'Ruby' : 'Manual',
      formula: formula.substring(0, 100) + '...' 
    });
    
    formulaCell.setFormula(formula);
    SpreadsheetApp.flush();
    
    for (let attempt = 1; attempt <= C.MAX_AI_RETRIES; attempt++) {
      Utilities.sleep(C.WAIT_FOR_AI_MS);
      
      const result = formulaCell.getValue();
      
      er_log_('AI attempt result', { 
        row, 
        attempt, 
        resultType: typeof result,
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
          const sourceLabel = emailData.isRuby ? '[Ruby Email]' : '[Manual Lead]';
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(sourceLabel)
            .setLinkUrl(gmailUrl)
            .build();
          linkCell.setRichTextValue(richText);
          
          er_log_('AI parsing successful', { 
            row, 
            attempt,
            sourceType: emailData.isRuby ? 'Ruby' : 'Manual',
            fieldsFound: cleanedValues.filter(v => v).length 
          });
          
          return true;
        } else {
          er_log_('AI result has too few fields', { 
            row, 
            attempt,
            fieldCount: values.length 
          });
        }
      }
      
      if (attempt < C.MAX_AI_RETRIES) {
        er_log_('AI retry needed', { row, attempt });
        Utilities.sleep(3000);
      }
    }
    
    er_log_('AI parsing failed after all retries', { row });
    return false;
    
  } catch (err) {
    er_log_('AI parsing error', { 
      row,
      error: err.toString(),
      stack: err.stack 
    });
    return false;
  }
}

/**
 * Pattern parsing with source awareness
 */
function er_tryPatternParsing_(sheet, row, emailData, gmailUrl, searchConfig) {
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
    
    if (isRuby) {
      // Use Ruby-specific patterns
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
      
    } else {
      // Use generic patterns for manual leads
      const GP = C.GENERIC_PATTERNS;
      
      // Try to extract name (look for capitalized words at start of lines)
      const nameMatch = body.match(/^([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)/m);
      if (nameMatch) {
        const parts = nameMatch[1].split(' ');
        extracted.first = parts[0] || '';
        extracted.last = parts.slice(1).join(' ') || '';
      }
      
      // Extract email
      const emailMatch = body.match(GP.email);
      if (emailMatch) extracted.email = emailMatch[0];
      
      // Extract phone
      const phoneMatch = body.match(GP.phone);
      if (phoneMatch) {
        const digits = phoneMatch[0].replace(/\D/g, '');
        if (digits.length === 10) {
          extracted.phone = digits.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
        } else if (digits.length === 11) {
          extracted.phone = digits.replace(/(\d{1})(\d{3})(\d{3})(\d{4})/, '$2-$3-$4');
        }
      }
      
      // Extract company (look for common patterns)
      const companyMatch = body.match(/(company|organization|business):\s*([^\n]+)/i);
      if (companyMatch) extracted.company = companyMatch[2].trim();
      
      // Extract address
      const addressMatch = body.match(GP.address);
      if (addressMatch) {
        extracted.address = addressMatch[0].replace(/,/g, ' ').replace(/\s+/g, ' ');
      }
      
      // Use subject as regarding if no better info found
      extracted.regarding = emailData.subject;
    }
    
    const fullName = `${extracted.first} ${extracted.last}`.trim();
    
    const rowData = [
      extracted.date,        // A
      '',                   // B - Link
      '',                   // C
      '1. F/U',             // D
      fullName,             // E
      extracted.company,    // F
      'Res',                // G
      extracted.phone,      // H
      extracted.email,      // I
      extracted.address,    // J
      extracted.regarding,  // K
      extracted.actions,    // L
      '', '', '', '', '', '', '', '', '' // M-U
    ];
    
    const hasData = fullName || extracted.phone || extracted.email || extracted.address;
    
    if (hasData) {
      const targetRange = sheet.getRange(row, 1, 1, 21);
      targetRange.setValues([rowData]);
      
      const linkCell = sheet.getRange(row, C.COLS.LINK_B);
      const sourceLabel = isRuby ? '[Ruby Email]' : '[Manual Lead]';
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(sourceLabel)
        .setLinkUrl(gmailUrl)
        .build();
      linkCell.setRichTextValue(richText);
      
      er_log_('Pattern parsing successful', { 
        row, 
        sourceType: isRuby ? 'Ruby' : 'Manual',
        name: fullName,
        extracted: Object.keys(extracted).filter(k => extracted[k])
      });
      
      return true;
    }
    
    er_log_('Pattern parsing found no data', { row });
    return false;
    
  } catch (err) {
    er_log_('Pattern parsing error', { 
      row,
      error: err.toString() 
    });
    return false;
  }
}

/**
 * Structured extraction (generic)
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
        
        if (value) {
          data[key] = value;
        }
      }
    }
    
    const firstName = data['first'] || '';
    const lastName = data['last'] || '';
    const fullName = `${firstName} ${lastName}`.trim();
    
    const rowData = [
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
      '',
      emailData.isRuby ? 'Ruby Email - Parsed' : 'Manual Lead - Parsed',
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
      const sourceLabel = emailData.isRuby ? '[Ruby Email]' : '[Manual Lead]';
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(sourceLabel)
        .setLinkUrl(gmailUrl)
        .build();
      linkCell.setRichTextValue(richText);
      
      er_log_('Structured extraction successful', { 
        row,
        sourceType: emailData.isRuby ? 'Ruby' : 'Manual'
      });
      return true;
    }
    
    return false;
    
  } catch (err) {
    er_log_('Structured extraction error', { 
      row,
      error: err.toString() 
    });
    return false;
  }
}

/**
 * Save for manual review
 */
function er_saveForManualReview_(sheet, row, emailData, gmailUrl) {
  const C = EMAIL_READER_CONFIG;
  const sourceLabel = emailData.isRuby ? 'Ruby' : 'Manual';
  
  const rowData = [
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
    '',
    `⚠️ NEEDS MANUAL PARSING - ${sourceLabel} Email`,
    '1. F/U',
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
  ];
  
  const targetRange = sheet.getRange(row, 1, 1, 21);
  targetRange.setValues([rowData]);
  
  const linkCell = sheet.getRange(row, C.COLS.LINK_B);
  const richText = SpreadsheetApp.newRichTextValue()
    .setText('[Click to View Email in Gmail]')
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
 * Install trigger
 */
function er_installTrigger() {
  er_removeTrigger();
  
  ScriptApp.newTrigger('er_processNewEmails')
    .timeBased()
    .everyMinutes(15)
    .create();
  
  SpreadsheetApp.getUi().alert('✅ Email Reader trigger installed!\n\nWill check for new Ruby and Manual leads every 15 minutes.\n\nRun "Run Email Reader Now" to test immediately.');
}

/**
 * Remove trigger
 */
function er_removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'er_processNewEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Main processing function - handles both Ruby and Manual leads
 */
function er_processNewEmails() {
  const C = EMAIL_READER_CONFIG;
  
  try {
    // Ensure labels exist
    let processedLabel = GmailApp.getUserLabelByName(C.PROCESSED_LABEL);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(C.PROCESSED_LABEL);
      er_log_('Created processed label', { labelName: C.PROCESSED_LABEL });
    }
    
    let addLeadLabel = GmailApp.getUserLabelByName(C.ADD_LEAD_LABEL);
    if (!addLeadLabel) {
      addLeadLabel = GmailApp.createLabel(C.ADD_LEAD_LABEL);
      er_log_('Created add lead label', { labelName: C.ADD_LEAD_LABEL });
    }
    
    let totalProcessed = 0;
    let totalSkipped = 0;
    
    for (const search of C.SEARCHES) {
      if (!search.enabled) continue;
      
      // Search for unread emails
      const threads = GmailApp.search(search.query, 0, C.MAX_EMAILS_PER_RUN);
      
      er_log_('Search completed', { 
        query: search.query,
        sourceType: search.sourceType,
        threadsFound: threads.length 
      });
      
      for (const thread of threads) {
        // Check if thread already has the processed label
        const threadLabels = thread.getLabels();
        if (threadLabels.some(l => l.getName() === C.PROCESSED_LABEL)) {
          totalSkipped++;
          continue;
        }
        
        const messages = thread.getMessages();
        for (const message of messages) {
          // For Ruby search, skip if message is not from Ruby
          if (search.sourceType === 'ruby' && !message.getFrom().includes('noreply@ruby.com')) {
            continue;
          }
          
          er_processEmail_(message, search, processedLabel);
          totalProcessed++;
        }
      }
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Batch complete', { 
        totalProcessed,
        totalSkipped
      });
    }
    
    if (totalProcessed > 0) {
      SpreadsheetApp.getActive().toast(
        `Processed ${totalProcessed} new lead(s)`,
        'Email Reader',
        3
      );
    }
    
    return totalProcessed;
    
  } catch (err) {
    er_log_('Batch processing error', { 
      error: err.toString(),
      stack: err.stack 
    });
    
    SpreadsheetApp.getActive().toast(
      `Email Reader error: ${err.message}`,
      'Error',
      5
    );
    
    throw err;
  }
}

/**
 * Test processing with detailed output
 */
function er_testProcessing() {
  const C = EMAIL_READER_CONFIG;
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Look for either Ruby or add lead emails
    const rubyThreads = GmailApp.search('from:noreply@ruby.com is:unread', 0, 1);
    const manualThreads = GmailApp.search(`label:${C.ADD_LEAD_LABEL} is:unread`, 0, 1);
    
    let testThread = null;
    let sourceType = '';
    
    if (rubyThreads.length > 0) {
      testThread = rubyThreads[0];
      sourceType = 'Ruby';
    } else if (manualThreads.length > 0) {
      testThread = manualThreads[0];
      sourceType = 'Manual';
    }
    
    if (!testThread) {
      ui.alert(
        'No Unread Leads Found',
        'No unread emails found.\n\n' +
        'To test:\n' +
        '1. For Ruby emails: Find a Ruby email and mark it as unread\n' +
        '2. For Manual leads: Label an email with "add lead" and mark as unread\n' +
        '3. Run this test again',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Ensure label exists
    let processedLabel = GmailApp.getUserLabelByName(C.PROCESSED_LABEL);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(C.PROCESSED_LABEL);
    }
    
    const message = testThread.getMessages()[0];
    
    // Determine which search config to use
    const searchConfig = sourceType === 'Ruby' ? C.SEARCHES[0] : C.SEARCHES[1];
    
    ui.alert(
      `Found ${sourceType} Email`,
      `Subject: ${message.getSubject()}\n` +
      `From: ${message.getFrom()}\n` +
      `Date: ${message.getDate()}\n\n` +
      'Processing now...',
      ui.ButtonSet.OK
    );
    
    er_processEmail_(message, searchConfig, processedLabel);
    
    ui.alert(
      'Test Complete!',
      'Check the Leads sheet for the new row.\n\n' +
      'If the row has errors or is marked for manual review:\n' +
      '1. Check View > Logs for detailed error messages\n' +
      '2. Run "Diagnostic Check" from the Email Reader menu\n' +
      `3. Verify the ${sourceType} email format matches expected format`,
      ui.ButtonSet.OK
    );
    
  } catch (err) {
    ui.alert(
      'Test Failed',
      `Error: ${err.toString()}\n\n` +
      'Check View > Logs for detailed error information',
      ui.ButtonSet.OK
    );
    er_log_('Test error', { 
      error: err.toString(),
      stack: err.stack 
    });
  }
}
