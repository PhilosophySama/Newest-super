/**
 * EMAIL READER AUTOMATION - RUBY PARSER
 * Version: 01/13-03:00PM EST by Claude Opus 4.1
 * 
 * Enhanced Ruby email parsing with field-specific extraction
 */

const EMAIL_READER_CONFIG = {
  TARGET_SHEET: 'Leads',
  
  SEARCHES: [
    {
      name: 'Ruby Mail',
      query: 'from:noreply@ruby.com is:unread -label:LeadProcessed',  // ← FIXED: Excludes already processed emails
      enabled: true,
      settings: {
        stage: '1. F/U',
        category: 'Res',
        autoSplit: true,
        markAsRead: false,
        useAI: true,
        useFallback: true
      }
    }
  ],
  
  PROCESSED_LABEL: 'LeadProcessed',
  
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
  
  AI_FORMULA_TEMPLATE: `=AI("Parse Ruby email. Extract fields in format 'First:', 'Last:', 'Phone Number:', 'Company:', 'Regarding:', 'Project Address:', 'Email:', 'Actions:'. Return EXACTLY 21 comma-separated values:
A: Today MM/DD
B: blank
C: blank
D: 1. F/U
E: Combine First+space+Last
F: Company value
G: Res
H: Phone Number value
I: Email value (or blank)
J: Project Address (remove commas)
K: Regarding value
L: Actions value
M-U: blank (9 blanks)
Example: 01/13,,,1. F/U,John Smith,ABC Company,Res,954-555-1234,john@email.com,123 Main St,Need awning quote,Follow up,,,,,,,,
",B{ROW})`,

  MAX_EMAILS_PER_RUN: 10,
  WAIT_FOR_AI_MS: 10000,
  MAX_AI_RETRIES: 3,
  ENABLE_LOGGING: true,
  
  RUBY_PATTERNS: {
    first: /First:\s*([^\n]+)/i,
    last: /Last:\s*([^\n]+)/i,
    phone: /Phone Number:\s*([^\n]+)/i,
    company: /Company:\s*([^\n]+)/i,
    regarding: /Regarding:\s*([^\n]+)/i,
    address: /Project Address:\s*([^\n]+)/i,
    email: /Email:\s*([^\n]+)/i,
    actions: /Actions:\s*([^\n]+)/i
  }
};

function er_processEmail_(message, searchConfig, processedLabel) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const emailData = {
      subject: message.getSubject(),
      from: message.getFrom(),
      date: message.getDate(),
      body: message.getPlainBody(),
      htmlBody: message.getBody(),
      messageId: message.getId()
    };
    
    const gmailUrl = `https://mail.google.com/mail/u/0/#inbox/${emailData.messageId}`;
    const emailContent = emailData.body;
    
    if (C.ENABLE_LOGGING) {
      er_log_('Processing Ruby email', {
        from: emailData.from,
        subject: emailData.subject,
        bodyLength: emailContent.length
      });
    }
    
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(C.TARGET_SHEET);
    
    if (!sheet) {
      throw new Error(`Sheet "${C.TARGET_SHEET}" not found`);
    }
    
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    const contentCell = sheet.getRange(newRow, C.COLS.LINK_B);
    contentCell.setValue(emailContent);
    
    let parseSuccess = false;
    
    if (searchConfig.settings.useAI) {
      parseSuccess = er_tryEnhancedAIParsing_(sheet, newRow, emailContent, gmailUrl);
    }
    
    if (!parseSuccess && searchConfig.settings.useFallback) {
      parseSuccess = er_tryRubyPatternParsing_(sheet, newRow, emailContent, gmailUrl, emailData);
    }
    
    if (!parseSuccess) {
      parseSuccess = er_tryStructuredExtraction_(sheet, newRow, emailContent, gmailUrl, emailData);
    }
    
    if (!parseSuccess) {
      er_saveForManualReview_(sheet, newRow, emailContent, gmailUrl);
    }
    
    const thread = message.getThread();
    thread.addLabel(processedLabel);
    
    if (searchConfig.settings.markAsRead) {
      thread.markRead();
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Email processed', {
        row: newRow,
        success: parseSuccess
      });
    }
    
  } catch (err) {
    er_log_('Email processing error', { error: err.toString() });
    throw err;
  }
}

function er_tryEnhancedAIParsing_(sheet, row, emailContent, gmailUrl) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const formulaCell = sheet.getRange(row, C.COLS.DATE);
    const formula = C.AI_FORMULA_TEMPLATE.replace('{ROW}', row);
    formulaCell.setFormula(formula);
    
    for (let attempt = 1; attempt <= C.MAX_AI_RETRIES; attempt++) {
      SpreadsheetApp.flush();
      Utilities.sleep(C.WAIT_FOR_AI_MS);
      
      const result = formulaCell.getValue();
      
      if (result && String(result).trim() !== '' && !String(result).includes('#ERROR')) {
        const values = String(result).split(',');
        
        if (values.length >= 12) {
          while (values.length < 21) values.push('');
          if (values.length > 21) values.length = 21;
          
          const cleanedValues = values.map(v => 
            String(v).trim()
              .replace(/^["']|["']$/g, '')
              .replace(/\s+/g, ' ')
          );
          
          const targetRange = sheet.getRange(row, 1, 1, 21);
          targetRange.setValues([cleanedValues]);
          
          const linkCell = sheet.getRange(row, C.COLS.LINK_B);
          const richText = SpreadsheetApp.newRichTextValue()
            .setText('[Ruby Email]')
            .setLinkUrl(gmailUrl)
            .build();
          linkCell.setRichTextValue(richText);
          
          if (C.ENABLE_LOGGING) {
            er_log_('AI parsing successful', { 
              row, 
              attempt,
              fieldsFound: cleanedValues.filter(v => v).length 
            });
          }
          
          return true;
        }
      }
      
      if (attempt < C.MAX_AI_RETRIES) {
        if (C.ENABLE_LOGGING) {
          er_log_('AI retry', { row, attempt, result });
        }
        Utilities.sleep(3000);
      }
    }
    
    return false;
    
  } catch (err) {
    er_log_('Enhanced AI parsing error', { error: err.toString() });
    return false;
  }
}

function er_tryRubyPatternParsing_(sheet, row, emailContent, gmailUrl, emailData) {
  const C = EMAIL_READER_CONFIG;
  const P = C.RUBY_PATTERNS;
  
  try {
    const body = emailData.body;
    
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
      '',                   // M
      '',                   // N
      '',                   // O
      '',                   // P
      '',                   // Q
      '',                   // R
      '',                   // S
      '',                   // T
      ''                    // U
    ];
    
    const hasData = fullName || extracted.phone || extracted.email || extracted.address;
    
    if (hasData) {
      const targetRange = sheet.getRange(row, 1, 1, 21);
      targetRange.setValues([rowData]);
      
      const linkCell = sheet.getRange(row, C.COLS.LINK_B);
      const richText = SpreadsheetApp.newRichTextValue()
        .setText('[Ruby Email]')
        .setLinkUrl(gmailUrl)
        .build();
      linkCell.setRichTextValue(richText);
      
      if (C.ENABLE_LOGGING) {
        er_log_('Pattern parsing successful', { row, name: fullName });
      }
      
      return true;
    }
    
    return false;
    
  } catch (err) {
    er_log_('Pattern parsing error', { error: err.toString() });
    return false;
  }
}

function er_tryStructuredExtraction_(sheet, row, emailContent, gmailUrl, emailData) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const lines = emailContent.split('\n');
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
      'Ruby Email - Parsed',
      '1. F/U',
      fullName || data['name'] || data['contact'] || '',
      data['company'] || '',
      'Res',
      (data['phone number'] || data['phone'] || '').replace(/\D/g, '').replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3'),
      data['email'] || '',
      (data['project address'] || data['address'] || '').replace(/,/g, ' '),
      data['regarding'] || '',
      data['actions'] || '',
      '', '', '', '', '', '', '', '', ''
    ];
    
    if (rowData[4] || rowData[7] || rowData[8]) {
      const targetRange = sheet.getRange(row, 1, 1, 21);
      targetRange.setValues([rowData]);
      
      const linkCell = sheet.getRange(row, C.COLS.LINK_B);
      const richText = SpreadsheetApp.newRichTextValue()
        .setText('[Ruby Email]')
        .setLinkUrl(gmailUrl)
        .build();
      linkCell.setRichTextValue(richText);
      
      return true;
    }
    
    return false;
    
  } catch (err) {
    er_log_('Structured extraction error', { error: err.toString() });
    return false;
  }
}

function er_saveForManualReview_(sheet, row, emailContent, gmailUrl) {
  const C = EMAIL_READER_CONFIG;
  
  const rowData = [
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd'),
    '',
    '⚠️ NEEDS MANUAL PARSING - Ruby Email',
    '1. F/U',
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
  ];
  
  const targetRange = sheet.getRange(row, 1, 1, 21);
  targetRange.setValues([rowData]);
  
  const linkCell = sheet.getRange(row, C.COLS.LINK_B);
  const richText = SpreadsheetApp.newRichTextValue()
    .setText('[Click to View Ruby Email in Gmail]')
    .setLinkUrl(gmailUrl)
    .build();
  linkCell.setRichTextValue(richText);
  
  sheet.getRange(row, 1, 1, 21).setBackground('#ffebee');
  
  sheet.getRange(row, C.COLS.COMMENTS).setNote(
    `Ruby Email Content:\n\n${emailContent.substring(0, 500)}${emailContent.length > 500 ? '...' : ''}`
  );
  
  if (C.ENABLE_LOGGING) {
    er_log_('Saved for manual review', { row });
  }
}

function er_log_(operation, details) {
  if (!EMAIL_READER_CONFIG.ENABLE_LOGGING) return;
  try {
    console.log(`[EmailReader] ${operation}:`, JSON.stringify(details));
  } catch (err) {
    console.log(`[EmailReader] Log error:`, err.message);
  }
}

function er_installTrigger() {
  er_removeTrigger();
  
  ScriptApp.newTrigger('er_processNewEmails')
    .timeBased()
    .everyMinutes(15)
    .create();
  
  SpreadsheetApp.getUi().alert('Email Reader trigger installed! Will check for new Ruby emails every 15 minutes.');
}

function er_removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'er_processNewEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

function er_processNewEmails() {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const processedLabel = GmailApp.getUserLabelByName(C.PROCESSED_LABEL) || 
                          GmailApp.createLabel(C.PROCESSED_LABEL);
    
    let totalProcessed = 0;
    
    for (const search of C.SEARCHES) {
      if (!search.enabled) continue;
      
      const threads = GmailApp.search(search.query, 0, C.MAX_EMAILS_PER_RUN);
      
      for (const thread of threads) {
        const messages = thread.getMessages();
        for (const message of messages) {
          er_processEmail_(message, search, processedLabel);
          totalProcessed++;
        }
      }
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Batch complete', { totalProcessed });
    }
    
    return totalProcessed;
    
  } catch (err) {
    er_log_('Batch processing error', { error: err.toString() });
    throw err;
  }
}

function er_testProcessing() {
  const C = EMAIL_READER_CONFIG;
  
  try {
    const threads = GmailApp.search('from:noreply@ruby.com', 0, 1);
    
    if (threads.length === 0) {
      SpreadsheetApp.getUi().alert('No Ruby emails found to test with.');
      return;
    }
    
    const processedLabel = GmailApp.getUserLabelByName(C.PROCESSED_LABEL) || 
                          GmailApp.createLabel(C.PROCESSED_LABEL);
    
    const message = threads[0].getMessages()[0];
    er_processEmail_(message, C.SEARCHES[0], processedLabel);
    
    SpreadsheetApp.getUi().alert('Test complete! Check the Leads sheet for the new row.');
    
  } catch (err) {
    SpreadsheetApp.getUi().alert('Test failed: ' + err.toString());
    er_log_('Test error', { error: err.toString() });
  }
}