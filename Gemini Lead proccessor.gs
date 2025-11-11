/**
 * GEMINI LEAD PROCESSOR
 * Version: 1.0 - 11/05/2025
 * 
 * Watches for emails labeled "Add lead" and processes them through Gemini API
 * to extract lead information as a single CSV string for column A of Leads sheet.
 * 
 * FEATURES:
 * - Monitors Gmail for "Add lead" label
 * - Sends email content to Gemini with structured prompt
 * - Writes CSV string to column A (Stage Automation splits it)
 * - Removes "Add lead" label and adds "Processed" label after success
 * - Error handling with detailed logging
 * 
 * SETUP REQUIRED:
 * 1. Get Gemini API key from Google AI Studio (ai.google.dev)
 * 2. Add to Script Properties: GEMINI_API_KEY
 * 3. Create Gmail label: "Add lead"
 * 4. Run: Setup (Gemini Leads) → Install Trigger
 */

const GEMINI_LEAD_CONFIG = {
  // Sheet configuration
  TARGET_SHEET: 'Leads',
  TARGET_COLUMN: 1, // Column A - where CSV string goes
  
  // Gmail labels
  WATCH_LABEL: 'Add lead',
  PROCESSED_LABEL: 'Lead Processed',
  ERROR_LABEL: 'Lead Error',
  
  // Gemini API configuration
  GEMINI_MODEL: 'gemini-1.5-pro-latest',
  GEMINI_API_ENDPOINT: 'https://generativelanguage.googleapis.com/v1beta/models/',
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 2000,
  
  // Processing limits
  MAX_EMAILS_PER_RUN: 10,
  
  // Feature flags
  ENABLE_LOGGING: true,
  REMOVE_WATCH_LABEL_AFTER_PROCESSING: true,
  
  // The exact prompt for Gemini
  GEMINI_PROMPT: `Goal: Convert an email screenshot of a lead into a single text line that I can paste into Google Sheets. The line must contain exactly 21 columns (A → U) separated by commas. Each column must be filled in according to the rules below. If a value is missing, leave the column blank but still include the comma separator. Do not include commas inside the text itself. Replace or omit them, otherwise the split will break.

General Rules
Format: Each column must be separated by a single comma.
No internal commas: If the text (like addresses or comments) normally contains commas, remove them.
Order: Always follow the exact column order A → U.
Output: Provide only the single copy/paste line, nothing extra.

Column Mapping (A → U)
A: Today's date in MM/DD format
B: (leave blank)
C: (leave blank)
D: Stage → If the lead is an invitation to bid, insert "4. ITB" - Draft. Otherwise insert "1. F/U"
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
U: Width/Projection of awning (ft)

Now process this email and provide ONLY the comma-separated line with no explanation:`
};

/*****************************************************************************
 * MAIN PROCESSING FUNCTION - Called by trigger
 *****************************************************************************/

/**
 * Main function to process emails labeled "Add lead"
 * Called every 15 minutes by time-based trigger
 */
function gl_processAddLeadEmails() {
  const C = GEMINI_LEAD_CONFIG;
  
  try {
    gl_log_('Starting Add Lead email processing', { timestamp: new Date() });
    
    // Get or create necessary labels
    const watchLabel = gl_getOrCreateLabel_(C.WATCH_LABEL);
    const processedLabel = gl_getOrCreateLabel_(C.PROCESSED_LABEL);
    const errorLabel = gl_getOrCreateLabel_(C.ERROR_LABEL);
    
    if (!watchLabel) {
      throw new Error(`Could not find or create label: ${C.WATCH_LABEL}`);
    }
    
    // Search for threads with "Add lead" label that aren't already processed
    const threads = GmailApp.search(`label:"${C.WATCH_LABEL}" -label:"${C.PROCESSED_LABEL}"`, 0, C.MAX_EMAILS_PER_RUN);
    
    gl_log_('Found threads to process', { count: threads.length });
    
    if (threads.length === 0) {
      gl_log_('No new Add Lead emails found', {});
      return;
    }
    
    let successCount = 0;
    let errorCount = 0;
    
    // Process each thread
    for (const thread of threads) {
      try {
        const success = gl_processThread_(thread, processedLabel, errorLabel);
        if (success) {
          successCount++;
        } else {
          errorCount++;
        }
      } catch (err) {
        gl_log_('Thread processing error', { 
          threadId: thread.getId(),
          error: err.message 
        });
        errorCount++;
        
        // Add error label
        thread.addLabel(errorLabel);
      }
    }
    
    // Show summary toast
    if (successCount > 0 || errorCount > 0) {
      SpreadsheetApp.getActive().toast(
        `Processed ${successCount} lead(s)\n${errorCount > 0 ? `${errorCount} error(s)` : ''}`,
        'Add Lead Processing',
        5
      );
    }
    
    gl_log_('Processing complete', { successCount, errorCount });
    
  } catch (err) {
    gl_log_('Fatal processing error', { 
      error: err.message,
      stack: err.stack 
    });
    
    SpreadsheetApp.getActive().toast(
      `Add Lead processor error: ${err.message}`,
      'Error',
      10
    );
    
    throw err;
  }
}

/**
 * Process a single thread
 */
function gl_processThread_(thread, processedLabel, errorLabel) {
  const C = GEMINI_LEAD_CONFIG;
  
  try {
    // Get the first message from the thread
    const messages = thread.getMessages();
    if (messages.length === 0) {
      throw new Error('Thread has no messages');
    }
    
    const message = messages[0];
    const messageId = message.getId();
    
    gl_log_('Processing message', {
      messageId,
      subject: message.getSubject(),
      from: message.getFrom()
    });
    
    // Extract email content
    const emailContent = gl_extractEmailContent_(message);
    
    // Send to Gemini for processing
    const csvString = gl_processWithGemini_(emailContent);
    
    if (!csvString) {
      throw new Error('Gemini returned empty result');
    }
    
    // Write to sheet
    gl_writeToSheet_(csvString, messageId);
    
    // Mark as processed
    thread.addLabel(processedLabel);
    
    // Remove "Add lead" label if configured
    if (C.REMOVE_WATCH_LABEL_AFTER_PROCESSING) {
      const watchLabel = GmailApp.getUserLabelByName(C.WATCH_LABEL);
      if (watchLabel) {
        thread.removeLabel(watchLabel);
      }
    }
    
    gl_log_('Thread processed successfully', { messageId });
    
    return true;
    
  } catch (err) {
    gl_log_('Thread processing failed', {
      threadId: thread.getId(),
      error: err.message
    });
    
    // Add error label but keep "Add lead" for retry
    thread.addLabel(errorLabel);
    
    return false;
  }
}

/*****************************************************************************
 * EMAIL CONTENT EXTRACTION
 *****************************************************************************/

/**
 * Extract content from email message
 * Handles both plain text and HTML emails
 */
function gl_extractEmailContent_(message) {
  const C = GEMINI_LEAD_CONFIG;
  
  try {
    // Get email metadata
    const subject = message.getSubject();
    const from = message.getFrom();
    const date = message.getDate();
    
    // Get body content (prefer plain text, fallback to HTML)
    let body = message.getPlainBody();
    
    if (!body || body.trim().length < 50) {
      // Try HTML body if plain text is too short
      body = message.getBody();
      // Simple HTML strip (not perfect but works for most cases)
      body = body.replace(/<[^>]*>/g, ' ')
                 .replace(/&nbsp;/g, ' ')
                 .replace(/&amp;/g, '&')
                 .replace(/&lt;/g, '<')
                 .replace(/&gt;/g, '>')
                 .replace(/\s+/g, ' ')
                 .trim();
    }
    
    // Build comprehensive content for Gemini
    const content = `
EMAIL SUBJECT: ${subject}
FROM: ${from}
DATE: ${Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yyyy')}

EMAIL CONTENT:
${body}
`.trim();
    
    gl_log_('Extracted email content', {
      subject,
      contentLength: content.length
    });
    
    return content;
    
  } catch (err) {
    gl_log_('Content extraction error', { error: err.message });
    throw new Error(`Failed to extract email content: ${err.message}`);
  }
}

/*****************************************************************************
 * GEMINI API INTEGRATION
 *****************************************************************************/

/**
 * Process email content with Gemini API
 * Returns CSV string with 21 comma-separated values
 */
function gl_processWithGemini_(emailContent) {
  const C = GEMINI_LEAD_CONFIG;
  
  // Get API key from script properties
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY not found in Script Properties. Please add it first.');
  }
  
  const fullPrompt = `${C.GEMINI_PROMPT}\n\n${emailContent}`;
  
  gl_log_('Sending to Gemini', { 
    promptLength: fullPrompt.length,
    model: C.GEMINI_MODEL
  });
  
  // Try processing with retries
  for (let attempt = 1; attempt <= C.MAX_RETRIES; attempt++) {
    try {
      const result = gl_callGeminiAPI_(apiKey, fullPrompt);
      
      if (result) {
        // Validate result has 21 columns
        const columns = result.split(',');
        
        if (columns.length !== 21) {
          gl_log_('Gemini returned wrong column count', {
            attempt,
            expected: 21,
            actual: columns.length,
            result: result.substring(0, 200)
          });
          
          if (attempt < C.MAX_RETRIES) {
            Utilities.sleep(C.RETRY_DELAY_MS);
            continue;
          }
          
          throw new Error(`Gemini returned ${columns.length} columns, expected 21`);
        }
        
        gl_log_('Gemini processing successful', {
          attempt,
          columnCount: columns.length,
          resultPreview: result.substring(0, 100)
        });
        
        return result;
      }
      
    } catch (err) {
      gl_log_('Gemini API call failed', {
        attempt,
        error: err.message
      });
      
      if (attempt < C.MAX_RETRIES) {
        Utilities.sleep(C.RETRY_DELAY_MS);
      } else {
        throw err;
      }
    }
  }
  
  throw new Error('Failed to process with Gemini after all retries');
}

/**
 * Make actual API call to Gemini
 */
function gl_callGeminiAPI_(apiKey, prompt) {
  const C = GEMINI_LEAD_CONFIG;
  
  const url = `${C.GEMINI_API_ENDPOINT}${C.GEMINI_MODEL}:generateContent?key=${apiKey}`;
  
  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      temperature: 0.2,  // Lower temperature for more consistent output
      topK: 40,
      topP: 0.95,
      maxOutputTokens: 1024,
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error(`Gemini API error ${responseCode}: ${responseText}`);
  }
  
  const data = JSON.parse(responseText);
  
  if (!data.candidates || data.candidates.length === 0) {
    throw new Error('No response from Gemini');
  }
  
  const candidate = data.candidates[0];
  
  if (!candidate.content || !candidate.content.parts || candidate.content.parts.length === 0) {
    throw new Error('Empty response from Gemini');
  }
  
  const result = candidate.content.parts[0].text.trim();
  
  return result;
}

/*****************************************************************************
 * SHEET WRITING
 *****************************************************************************/

/**
 * Write CSV string to column A of Leads sheet
 * Stage Automation will auto-split it across columns
 */
function gl_writeToSheet_(csvString, messageId) {
  const C = GEMINI_LEAD_CONFIG;
  
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(C.TARGET_SHEET);
    
    if (!sheet) {
      throw new Error(`Sheet "${C.TARGET_SHEET}" not found`);
    }
    
    // Find next empty row
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    // Write CSV string to column A
    const cell = sheet.getRange(newRow, C.TARGET_COLUMN);
    cell.setValue(csvString);
    
    gl_log_('Wrote to sheet', {
      sheet: C.TARGET_SHEET,
      row: newRow,
      messageId,
      csvPreview: csvString.substring(0, 100)
    });
    
    // Flush to ensure Stage Automation can process it
    SpreadsheetApp.flush();
    
    // Small delay to let Stage Automation split the columns
    Utilities.sleep(500);
    
  } catch (err) {
    gl_log_('Sheet write error', { error: err.message });
    throw new Error(`Failed to write to sheet: ${err.message}`);
  }
}

/*****************************************************************************
 * LABEL MANAGEMENT
 *****************************************************************************/

/**
 * Get existing label or create it if it doesn't exist
 */
function gl_getOrCreateLabel_(labelName) {
  try {
    let label = GmailApp.getUserLabelByName(labelName);
    
    if (!label) {
      gl_log_('Creating label', { labelName });
      label = GmailApp.createLabel(labelName);
    }
    
    return label;
    
  } catch (err) {
    gl_log_('Label creation error', { 
      labelName,
      error: err.message 
    });
    return null;
  }
}

/*****************************************************************************
 * UTILITIES & LOGGING
 *****************************************************************************/

/**
 * Log operation with timestamp
 */
function gl_log_(operation, details) {
  if (!GEMINI_LEAD_CONFIG.ENABLE_LOGGING) return;
  
  try {
    const timestamp = Utilities.formatDate(
      new Date(), 
      Session.getScriptTimeZone(), 
      'yyyy-MM-dd HH:mm:ss'
    );
    console.log(`[${timestamp}] [GeminiLeads] ${operation}:`, JSON.stringify(details));
  } catch (err) {
    console.log(`[GeminiLeads] Log error:`, err.message);
  }
}

/*****************************************************************************
 * TRIGGER MANAGEMENT (Called from Menu)
 *****************************************************************************/

/**
 * Install time-based trigger to check for "Add lead" emails every 15 minutes
 */
function gl_installTrigger() {
  gl_removeTrigger();
  
  ScriptApp.newTrigger('gl_processAddLeadEmails')
    .timeBased()
    .everyMinutes(15)
    .create();
  
  SpreadsheetApp.getUi().alert(
    '✅ Gemini Lead Processor Installed!',
    'The system will now check for "Add lead" labeled emails every 15 minutes.\n\n' +
    'To test:\n' +
    '1. Label an email with "Add lead"\n' +
    '2. Wait up to 15 minutes (or run "Process Now" from menu)\n' +
    '3. Check the Leads sheet for new row in column A',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  gl_log_('Trigger installed', { function: 'gl_processAddLeadEmails' });
}

/**
 * Remove existing trigger
 */
function gl_removeTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'gl_processAddLeadEmails') {
      ScriptApp.deleteTrigger(trigger);
      gl_log_('Trigger removed', { triggerId: trigger.getUniqueId() });
    }
  }
}

/*****************************************************************************
 * TESTING & DIAGNOSTICS (Called from Menu)
 *****************************************************************************/

/**
 * Run diagnostics to check setup
 */
function gl_diagnosticCheck() {
  const ui = SpreadsheetApp.getUi();
  const C = GEMINI_LEAD_CONFIG;
  const results = [];
  
  // 1. Check API key
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  results.push(apiKey ? '✅ Gemini API key found' : '❌ GEMINI_API_KEY not set in Script Properties');
  
  // 2. Check target sheet
  const sheet = SpreadsheetApp.getActive().getSheetByName(C.TARGET_SHEET);
  results.push(sheet ? `✅ Target sheet "${C.TARGET_SHEET}" found` : `❌ Target sheet "${C.TARGET_SHEET}" not found`);
  
  // 3. Check labels
  const watchLabel = GmailApp.getUserLabelByName(C.WATCH_LABEL);
  results.push(watchLabel ? `✅ Gmail label "${C.WATCH_LABEL}" exists` : `⚠️ Gmail label "${C.WATCH_LABEL}" will be created automatically`);
  
  // 4. Check for emails with label
  try {
    const threads = GmailApp.search(`label:"${C.WATCH_LABEL}"`, 0, 1);
    results.push(`📧 Found ${threads.length} email(s) with "${C.WATCH_LABEL}" label`);
  } catch (err) {
    results.push(`⚠️ Cannot search for labeled emails: ${err.message}`);
  }
  
  // 5. Check trigger
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some(t => t.getHandlerFunction() === 'gl_processAddLeadEmails');
  results.push(hasTrigger ? '✅ Trigger installed' : '⚠️ No trigger found (run Setup Auto-Check)');
  
  // 6. Test API connection (if key exists)
  if (apiKey) {
    try {
      const testResult = gl_callGeminiAPI_(apiKey, 'Say "test" in one word');
      results.push(testResult ? '✅ Gemini API connection works' : '⚠️ API returned empty response');
    } catch (err) {
      results.push(`❌ Gemini API error: ${err.message}`);
    }
  }
  
  ui.alert('Gemini Lead Processor Diagnostics', results.join('\n'), ui.ButtonSet.OK);
  
  gl_log_('Diagnostic check completed', { results });
}

/**
 * Process one email immediately for testing
 */
function gl_testProcessOneEmail() {
  const ui = SpreadsheetApp.getUi();
  const C = GEMINI_LEAD_CONFIG;
  
  try {
    // Find one email with "Add lead" label
    const threads = GmailApp.search(`label:"${C.WATCH_LABEL}"`, 0, 1);
    
    if (threads.length === 0) {
      ui.alert(
        'No Test Email Found',
        `No emails found with "${C.WATCH_LABEL}" label.\n\n` +
        'To test:\n' +
        '1. Find an email in Gmail\n' +
        '2. Add the "Add lead" label to it\n' +
        '3. Run this test again',
        ui.ButtonSet.OK
      );
      return;
    }
    
    ui.alert(
      'Processing Test Email',
      `Found email: ${threads[0].getFirstMessageSubject()}\n\n` +
      'Processing now...',
      ui.ButtonSet.OK
    );
    
    // Get labels
    const processedLabel = gl_getOrCreateLabel_(C.PROCESSED_LABEL);
    const errorLabel = gl_getOrCreateLabel_(C.ERROR_LABEL);
    
    // Process the thread
    const success = gl_processThread_(threads[0], processedLabel, errorLabel);
    
    if (success) {
      ui.alert(
        '✅ Test Successful!',
        'Check the Leads sheet for the new row.\n\n' +
        'The CSV string should be in column A, and Stage Automation should split it across columns.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '❌ Test Failed',
        'Check View > Logs for detailed error information.\n\n' +
        'The email was labeled with "Lead Error" for review.',
        ui.ButtonSet.OK
      );
    }
    
  } catch (err) {
    ui.alert(
      'Test Error',
      `Error: ${err.message}\n\n` +
      'Check View > Logs for details',
      ui.ButtonSet.OK
    );
    
    gl_log_('Test error', { 
      error: err.message,
      stack: err.stack 
    });
  }
}

/**
 * Show configuration info
 */
function gl_showConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const C = GEMINI_LEAD_CONFIG;
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  const config = `
GEMINI LEAD PROCESSOR CONFIGURATION
====================================

Target Sheet: ${C.TARGET_SHEET}
Target Column: A (${C.TARGET_COLUMN})

Gmail Labels:
• Watch: "${C.WATCH_LABEL}"
• Processed: "${C.PROCESSED_LABEL}"
• Error: "${C.ERROR_LABEL}"

Gemini API:
• Model: ${C.GEMINI_MODEL}
• API Key: ${apiKey ? 'Set ✅' : 'NOT SET ❌'}
• Max Retries: ${C.MAX_RETRIES}

Processing:
• Max Emails/Run: ${C.MAX_EMAILS_PER_RUN}
• Remove Label After Processing: ${C.REMOVE_WATCH_LABEL_AFTER_PROCESSING}

SETUP INSTRUCTIONS:
===================
1. Get API key from: https://ai.google.dev
2. Add to Script Properties:
   - Key: GEMINI_API_KEY
   - Value: [your API key]
3. Run: Setup (Gemini Leads) → Install Trigger
4. Label emails with "${C.WATCH_LABEL}"
`;
  
  ui.alert('Configuration', config, ui.ButtonSet.OK);
}