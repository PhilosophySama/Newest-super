/**
 * EMAIL READER AUTOMATION
 * Version: 01/13-11:45PM EST by Claude Opus 4.1
 * 
 * Scans emails for specific senders/subjects and processes them automatically.
 * Creates new rows in Leads sheet with AI-powered data extraction.
 *
 * FEATURES:
 * - Scans for Ruby Mail emails (noreply@ruby.com)
 * - Extracts email content and pastes into column B
 * - Adds AI formula in column A to parse lead data
 * - Auto-splits formula output across columns A-U
 * - Marks processed emails with label to avoid duplicates
 * - Runs on time-based trigger (every 5 minutes)
 * - Logs all operations for debugging
 *
 * CO-EXISTENCE:
 * - Config: EMAIL_READER_CONFIG
 * - Handler: processEmailsReader_
 * - Installer: installTriggerEmailReader_
 * - Helpers: er_*
 */

const EMAIL_READER_CONFIG = {
  // Target sheet for new leads
  TARGET_SHEET: 'Leads',
  
  // Email search configurations
  SEARCHES: [
    {
      name: 'Ruby Mail',
      query: 'from:noreply@ruby.com',
      enabled: true,
      // Settings for this email type
      settings: {
        stage: '1. F/U',  // Default stage if not ITB
        category: 'Res',  // Default category
        autoSplit: true,  // Auto-split the AI output
        markAsRead: false // Keep unread for manual review
      }
    }
    // Add more search configurations here as needed
  ],
  
  // Gmail label to mark processed emails (creates if doesn't exist)
  PROCESSED_LABEL: 'LeadProcessed',
  
  // Column configuration for Leads sheet
  COLS: {
    AI_FORMULA: 1,      // A - AI formula output
    EMAIL_CONTENT: 2,   // B - Raw email content
    COMMENTS: 3,        // C - Comments
    STAGE: 4,           // D - Stage
    NAME: 5,            // E - Customer Name
    DISPLAY: 6,         // F - Display Name
    TYPE: 7,            // G - Customer type
    PHONE: 8,           // H - Phone
    EMAIL: 9,           // I - Email
    ADDRESS: 10,        // J - Address
    DESC: 11,           // K - Job Description
    // ... columns continue through U (21 total)
  },
  
  // AI Formula template
  AI_FORMULA_TEMPLATE: `=ai("Goal: Convert an email screenshot of a lead into a single text line that I can paste into Google Sheets.
The line must contain exactly 21 columns (A → U) separated by commas.
Each column must be filled in according to the rules below.
If a value is missing, leave the column blank but still include the comma separator.
Do not include commas inside the text itself. Replace or omit them, otherwise the split will break.

General Rules
Format: Each column must be separated by a single comma.
No internal commas: If the text (like addresses or comments) normally contains commas, remove them.
Order: Always follow the exact column order A → U.
Output: Provide only the single copy/paste line, nothing extra.

Column Mapping (A → U)
A: Today's date in MM/DD format
B: (leave blank)
C: (leave blank)
D: Stage → If the lead is an invitation to bid, insert \\"4. ITB\\" - Draft Otherwise insert \\"1. F/U\\"
E: Contact name
F: (leave blank)
G: Category → Res = Residential, Comm = Commercial, GC = General Contractor
H: Phone number
I: Email address
J: Job address (no commas). Format on two lines: Line 1: Street address Line 2: City State ZIP
K: Job description
L: Comments
M: (leave blank)
N: (leave blank)
O: (leave blank)
P: (leave blank)
Q: (leave blank)
R: Type → (e.g. Re-Cover, New Fabric Awning, Aluminum Canopy)
S: (leave blank)
T: Length of awning (ft)
U: Width/Projection of awning (ft)",B)`,
  
  // Processing settings
  MAX_EMAILS_PER_RUN: 10,        // Limit emails processed per execution
  WAIT_FOR_AI_MS: 5000,          // Wait time for AI formula to calculate
  ENABLE_LOGGING: true,          // Enable detailed logging
  
  // Error handling
  RETRY: {
    MAX_ATTEMPTS: 3,
    DELAYS_MS: [1000, 3000, 5000]
  }
};

/**
 * Main email processing function - runs on time trigger
 */
function processEmailsReader_() {
  const C = EMAIL_READER_CONFIG;
  
  if (C.ENABLE_LOGGING) {
    er_log_('Email reader started', { time: new Date().toISOString() });
  }
  
  try {
    // Get or create the processed label
    const processedLabel = er_getOrCreateLabel_(C.PROCESSED_LABEL);
    
    let totalProcessed = 0;
    let totalErrors = 0;
    
    // Process each search configuration
    for (const searchConfig of C.SEARCHES) {
      if (!searchConfig.enabled) continue;
      
      const result = er_processSearchConfig_(searchConfig, processedLabel);
      totalProcessed += result.processed;
      totalErrors += result.errors;
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Email reader completed', { 
        totalProcessed, 
        totalErrors,
        time: new Date().toISOString() 
      });
    }
    
    // Show toast if spreadsheet is open
    if (totalProcessed > 0) {
      try {
        SpreadsheetApp.getActive().toast(
          `Processed ${totalProcessed} new lead email(s)`,
          'Email Reader',
          3
        );
      } catch (e) {
        // Spreadsheet not open, skip toast
      }
    }
    
  } catch (err) {
    if (C.ENABLE_LOGGING) {
      er_log_('Email reader error', { error: err.message });
    }
    throw err;
  }
}

/**
 * Process emails for a specific search configuration
 */
function er_processSearchConfig_(searchConfig, processedLabel) {
  const C = EMAIL_READER_CONFIG;
  let processed = 0;
  let errors = 0;
  
  try {
    // Build search query - exclude already processed emails
    const query = `${searchConfig.query} -label:${C.PROCESSED_LABEL}`;
    
    if (C.ENABLE_LOGGING) {
      er_log_('Searching emails', { 
        name: searchConfig.name,
        query: query 
      });
    }
    
    // Search for matching emails
    const threads = GmailApp.search(query, 0, C.MAX_EMAILS_PER_RUN);
    
    if (threads.length === 0) {
      if (C.ENABLE_LOGGING) {
        er_log_('No new emails found', { name: searchConfig.name });
      }
      return { processed: 0, errors: 0 };
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Found emails to process', { 
        name: searchConfig.name,
        count: threads.length 
      });
    }
    
    // Process each thread
    for (const thread of threads) {
      try {
        const messages = thread.getMessages();
        
        // Process the most recent message in the thread
        if (messages.length > 0) {
          const message = messages[messages.length - 1];
          er_processEmail_(message, searchConfig, processedLabel);
          processed++;
        }
        
      } catch (emailErr) {
        errors++;
        if (C.ENABLE_LOGGING) {
          er_log_('Error processing email', { 
            error: emailErr.message,
            threadId: thread.getId() 
          });
        }
      }
    }
    
  } catch (searchErr) {
    if (C.ENABLE_LOGGING) {
      er_log_('Search error', { 
        name: searchConfig.name,
        error: searchErr.message 
      });
    }
    errors++;
  }
  
  return { processed, errors };
}

/**
 * Process a single email message
 */
function er_processEmail_(message, searchConfig, processedLabel) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    // Get email content
    const subject = message.getSubject();
    const from = message.getFrom();
    const date = message.getDate();
    const body = message.getPlainBody();
    
    // Get attachments info
    const attachments = message.getAttachments();
    const attachmentInfo = attachments.length > 0
      ? `\n\n[Attachments: ${attachments.map(a => a.getName()).join(', ')}]`
      : '';
    
    // Build email content string
    const emailContent = `From: ${from}
Subject: ${subject}
Date: ${Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm')}
${attachmentInfo}

${body}`;
    
    if (C.ENABLE_LOGGING) {
      er_log_('Processing email', { 
        from: from,
        subject: subject,
        date: date,
        bodyLength: body.length 
      });
    }
    
    // Get the target sheet
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(C.TARGET_SHEET);
    
    if (!sheet) {
      throw new Error(`Sheet "${C.TARGET_SHEET}" not found`);
    }
    
    // Add new row at the bottom
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    // Put email content in column B
    const contentCell = sheet.getRange(newRow, C.COLS.EMAIL_CONTENT);
    contentCell.setValue(emailContent);
    
    // Add AI formula in column A
    const formulaCell = sheet.getRange(newRow, C.COLS.AI_FORMULA);
    formulaCell.setFormula(C.AI_FORMULA_TEMPLATE);
    
    if (C.ENABLE_LOGGING) {
      er_log_('Added row to sheet', { 
        sheet: C.TARGET_SHEET,
        row: newRow,
        from: from 
      });
    }
    
    // Force calculation and wait for AI formula
    SpreadsheetApp.flush();
    Utilities.sleep(C.WAIT_FOR_AI_MS);
    
    // Auto-split if enabled
    if (searchConfig.settings.autoSplit) {
      er_autoSplitRow_(sheet, newRow);
    }
    
    // Mark email as processed
    const thread = message.getThread();
    thread.addLabel(processedLabel);
    
    // Optionally mark as read
    if (searchConfig.settings.markAsRead) {
      thread.markRead();
    }
    
    if (C.ENABLE_LOGGING) {
      er_log_('Email processed successfully', { 
        row: newRow,
        from: from,
        subject: subject 
      });
    }
    
  } catch (err) {
    if (C.ENABLE_LOGGING) {
      er_log_('Email processing error', { 
        error: err.message,
        messageId: message.getId() 
      });
    }
    throw err;
  }
}

/**
 * Auto-split column A by commas across columns A-U
 */
function er_autoSplitRow_(sheet, row) {
  const C = EMAIL_READER_CONFIG;
  
  try {
    // Get the value from column A (AI formula output)
    const formulaCell = sheet.getRange(row, C.COLS.AI_FORMULA);
    const formulaValue = formulaCell.getValue();
    
    if (!formulaValue || String(formulaValue).trim() === '') {
      if (C.ENABLE_LOGGING) {
        er_log_('Cannot split - formula result empty', { row });
      }
      return;
    }
    
    // Split by comma
    const values = String(formulaValue).split(',');
    
    if (values.length !== 21) {
      if (C.ENABLE_LOGGING) {
        er_log_('Warning: Expected 21 columns, got ' + values.length, { 
          row,
          actual: values.length 
        });
      }
    }
    
    // Pad or trim to 21 columns
    while (values.length < 21) values.push('');
    if (values.length > 21) values.length = 21;
    
    // Write the split values across columns A-U
    const targetRange = sheet.getRange(row, 1, 1, 21);
    targetRange.setValues([values.map(v => String(v).trim())]);
    
    if (C.ENABLE_LOGGING) {
      er_log_('Row split successfully', { 
        row,
        columns: values.length 
      });
    }
    
  } catch (err) {
    if (C.ENABLE_LOGGING) {
      er_log_('Auto-split error', { 
        row,
        error: err.message 
      });
    }
    // Don't throw - splitting is optional
  }
}

/**
 * Get or create Gmail label
 */
function er_getOrCreateLabel_(labelName) {
  try {
    let label = GmailApp.getUserLabelByName(labelName);
    
    if (!label) {
      label = GmailApp.createLabel(labelName);
      if (EMAIL_READER_CONFIG.ENABLE_LOGGING) {
        er_log_('Created Gmail label', { name: labelName });
      }
    }
    
    return label;
    
  } catch (err) {
    if (EMAIL_READER_CONFIG.ENABLE_LOGGING) {
      er_log_('Label error', { 
        name: labelName,
        error: err.message 
      });
    }
    throw err;
  }
}

/**
 * Logging helper
 */
function er_log_(operation, details) {
  if (!EMAIL_READER_CONFIG.ENABLE_LOGGING) return;
  
  try {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] [EmailReader] ${operation}:`, JSON.stringify(details));
  } catch (err) {
    console.log(`[EmailReader] Logging error:`, err.message);
  }
}

/**
 * Install time-based trigger for email reader
 */
function installTriggerEmailReader_() {
  const C = EMAIL_READER_CONFIG;
  
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processEmailsReader_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger - runs every 5 minutes
  ScriptApp.newTrigger('processEmailsReader_')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  SpreadsheetApp.getActive().toast(
    'Email reader trigger installed!\nChecks emails every 5 minutes',
    'Email Reader Setup',
    5
  );
  
  if (C.ENABLE_LOGGING) {
    er_log_('Trigger installed', { 
      handler: 'processEmailsReader_',
      interval: 'every 5 minutes' 
    });
  }
}

/**
 * Manual run for testing - processes emails immediately
 */
function runEmailReaderNow_() {
  SpreadsheetApp.getActive().toast('Processing emails...', 'Email Reader', 2);
  
  try {
    processEmailsReader_();
    SpreadsheetApp.getActive().toast('Email processing complete!', 'Email Reader', 3);
  } catch (err) {
    SpreadsheetApp.getActive().toast(
      `Error: ${err.message}`,
      'Email Reader Error',
      5
    );
  }
}

/**
 * Test function - check for Ruby Mail emails without processing
 */
function testEmailReaderSearch_() {
  const C = EMAIL_READER_CONFIG;
  const ui = SpreadsheetApp.getUi();
  
  try {
    const query = 'from:noreply@ruby.com';
    const threads = GmailApp.search(query, 0, 5);
    
    if (threads.length === 0) {
      ui.alert('Email Reader Test', 'No emails found from noreply@ruby.com', ui.ButtonSet.OK);
      return;
    }
    
    const results = threads.map((thread, index) => {
      const messages = thread.getMessages();
      const lastMessage = messages[messages.length - 1];
      return `${index + 1}. ${lastMessage.getSubject()}\n   From: ${lastMessage.getFrom()}\n   Date: ${lastMessage.getDate()}`;
    }).join('\n\n');
    
    ui.alert('Email Reader Test', 
      `Found ${threads.length} email(s):\n\n${results}`,
      ui.ButtonSet.OK);
    
  } catch (err) {
    ui.alert('Email Reader Test Error', err.message, ui.ButtonSet.OK);
  }
}

/**
 * Clear processed label from all emails (for testing)
 */
function clearProcessedLabels_() {
  const C = EMAIL_READER_CONFIG;
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Clear Processed Labels',
    `This will remove the "${C.PROCESSED_LABEL}" label from all emails.\n\nThis allows emails to be reprocessed. Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const label = GmailApp.getUserLabelByName(C.PROCESSED_LABEL);
    
    if (!label) {
      ui.alert('No Label Found', `Label "${C.PROCESSED_LABEL}" does not exist.`, ui.ButtonSet.OK);
      return;
    }
    
    const threads = label.getThreads();
    
    for (const thread of threads) {
      thread.removeLabel(label);
    }
    
    ui.alert('Success', 
      `Cleared "${C.PROCESSED_LABEL}" label from ${threads.length} email thread(s).`,
      ui.ButtonSet.OK);
    
  } catch (err) {
    ui.alert('Error', `Failed to clear labels: ${err.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Show current configuration
 */
function showEmailReaderConfig_() {
  const C = EMAIL_READER_CONFIG;
  const ui = SpreadsheetApp.getUi();
  
  const searches = C.SEARCHES
    .map(s => `• ${s.name} (${s.enabled ? 'enabled' : 'disabled'})\n  Query: ${s.query}`)
    .join('\n\n');
  
  const config = `EMAIL READER CONFIGURATION
========================

Target Sheet: ${C.TARGET_SHEET}
Processed Label: ${C.PROCESSED_LABEL}
Max Emails Per Run: ${C.MAX_EMAILS_PER_RUN}
AI Wait Time: ${C.WAIT_FOR_AI_MS}ms
Logging: ${C.ENABLE_LOGGING ? 'Enabled' : 'Disabled'}

SEARCH CONFIGURATIONS:
${searches}

TRIGGER STATUS:
${er_getTriggerStatus_()}`;
  
  ui.alert('Email Reader Configuration', config, ui.ButtonSet.OK);
}

/**
 * Get trigger status
 */
function er_getTriggerStatus_() {
  const triggers = ScriptApp.getProjectTriggers().filter(
    t => t.getHandlerFunction() === 'processEmailsReader_'
  );
  
  if (triggers.length === 0) {
    return 'No trigger installed (run Install Trigger to activate)';
  }
  
  const trigger = triggers[0];
  return `Active - Runs every 5 minutes\nLast modified: ${trigger.getHandlerFunction()}`;
}

/**
 * Add new email search configuration
 * Example usage: addEmailSearch_('Contact Form', 'subject:Contact Form', true, {...})
 */
function addEmailSearch_(name, query, enabled, settings) {
  EMAIL_READER_CONFIG.SEARCHES.push({
    name: name,
    query: query,
    enabled: enabled,
    settings: settings || {
      stage: '1. F/U',
      category: 'Res',
      autoSplit: true,
      markAsRead: false
    }
  });
}

/** end-of-file */