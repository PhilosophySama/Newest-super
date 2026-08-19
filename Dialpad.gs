// Dialpad.gs
// version10 [06/22-05:00PM] by Claude Opus 4.8
// Walker Awning - PM 2.0
//
// =============================================================
// NOTES: Dialpad API features relevant to Google Workspace
// (customer database in Google Sheets)
// =============================================================
//  1. Send SMS to a single number ............... [UTILIZED]
//  2. Send MMS / attachments .................... [NOT UTILIZED - PDFs exceed 500 KiB cap]
//  3. Preset / standard message bodies .......... [UTILIZED - Customer Info request]
//  4. Send bulk SMS (up to 50 recipients) ....... [NOT UTILIZED]
//  5. Read SMS delivery status .................. [NOT UTILIZED]
//  6. Initiate outbound call (API) .............. [NOT UTILIZED]
//  7. dialpad:// click-to-call cell links ....... [NOT UTILIZED]
//  8. Pull call / contact history ............... [NOT UTILIZED]
//  9. Inbound webhooks (needs web app) .......... [NOT UTILIZED]
//
//  This file is intentionally minimal: prove ONE feature (manual
//  SMS to the selected row) works end-to-end before adding more.
//  On failure it surfaces the real HTTP code + Dialpad response.
// =============================================================

// ----- CONFIG -----
var DIALPAD_CONFIG = {
  PHONE_COL:     8,   // H - Customer Phone Number
  NAME_COL:      5,   // E - Customer Name
  HEADER_ROW:    1,
  SHEETS:        ['Leads', 'F/U', 'Awarded', 'Heaven', 'Re-cover'],
  API_KEY_PROP:  'Dialpad_Walker_Awning',
  USER_ID_PROP:  'DIALPAD_USER_ID',
  FROM_NUM_PROP: 'DIALPAD_FROM_NUMBER',
  API_BASE:      'https://dialpad.com/api/v2'
};

// ----- API HELPERS -----

function dp_getApiKey_() {
  var key = PropertiesService.getScriptProperties().getProperty(DIALPAD_CONFIG.API_KEY_PROP);
  if (!key) throw new Error('API key not found in Script Properties under "' + DIALPAD_CONFIG.API_KEY_PROP + '".');
  return key;
}

function dp_apiRequest_(method, endpoint, payload) {
  var separator = endpoint.indexOf('?') === -1 ? '?' : '&';
  var url = DIALPAD_CONFIG.API_BASE + endpoint + separator + 'apikey=' + dp_getApiKey_();

  var options = {
    method:             method,
    headers:            { 'Content-Type': 'application/json' },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);

  var response = UrlFetchApp.fetch(url, options);
  var code     = response.getResponseCode();
  var body     = response.getContentText();

  if (code < 200 || code >= 300) {
    // Surface the real status + response so we can see WHY it failed.
    throw new Error('Dialpad API error (HTTP ' + code + '):\n' + body);
  }
  return JSON.parse(body);
}

// ----- PHONE NORMALIZATION -----

function dp_normalisePhone_(raw) {
  var digits = String(raw).replace(/\D/g, '');
  if (digits.length === 10) digits = '1' + digits;
  if (digits.length < 11)   return null;
  return '+' + digits;
}

// ----- CORE SMS SEND -----

function dp_sendSms(toNumber, message) {
  var e164 = dp_normalisePhone_(toNumber);
  if (!e164) throw new Error('Invalid phone number: ' + toNumber);

  var props      = PropertiesService.getScriptProperties();
  var userId     = props.getProperty(DIALPAD_CONFIG.USER_ID_PROP);
  var fromNumber = props.getProperty(DIALPAD_CONFIG.FROM_NUM_PROP);
  if (!userId)     throw new Error(DIALPAD_CONFIG.USER_ID_PROP + ' missing from Script Properties.');
  if (!fromNumber) throw new Error(DIALPAD_CONFIG.FROM_NUM_PROP + ' missing from Script Properties.');

  return dp_apiRequest_('post', '/sms', {
    to_numbers:  [e164],
    text:        message,
    user_id:     userId,
    from_number: fromNumber
  });
}

// ----- MANUAL SMS TO SELECTED ROW (the one feature we are proving) -----

function dp_sendCustomMessage_() {
  var ui    = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveCell().getRow();

  if (row <= DIALPAD_CONFIG.HEADER_ROW) {
    ui.alert('Please select a data row first (not the header).');
    return;
  }
  if (DIALPAD_CONFIG.SHEETS.indexOf(sheet.getName()) === -1) {
    ui.alert('Please run this from one of the project sheets: ' + DIALPAD_CONFIG.SHEETS.join(', '));
    return;
  }

  var phone = sheet.getRange(row, DIALPAD_CONFIG.PHONE_COL).getDisplayValue();
  var name  = sheet.getRange(row, DIALPAD_CONFIG.NAME_COL).getDisplayValue();
  if (!phone || phone.trim() === '') {
    ui.alert('No phone number found in column H for row ' + row + '.');
    return;
  }

  var prompt = ui.prompt('Send SMS to ' + (name || phone), 'Enter your message:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  var text = prompt.getResponseText();
  if (!text) { ui.alert('Empty message - nothing sent.'); return; }

  try {
    var result = dp_sendSms(phone, text);
    ui.alert('Message sent to ' + (name || phone) + ' (' + phone + ').\n\nResponse:\n' +
             JSON.stringify(result, null, 2).substring(0, 600));
  } catch (err) {
    // Full error shown on purpose so we can diagnose.
    ui.alert('Send FAILED.\n\n' + err.message);
  }
}

// ----- CONNECTION TEST (sends to a number you specify in code) -----

function dp_testSend() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Dialpad test send', 'Enter a phone number to text yourself:', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  try {
    var result = dp_sendSms(prompt.getResponseText(), 'Test from Walker Awning automation.');
    ui.alert('Test OK.\n\nResponse:\n' + JSON.stringify(result, null, 2).substring(0, 600));
  } catch (err) {
    ui.alert('Test FAILED.\n\n' + err.message);
  }
}
// ----- CUSTOMER INFO REQUEST AS SMS (swatches LINKED, not attached) -----
// Called by Draft Creator's v2_createCustomerInfoDraft_ so it piggybacks the
// existing "Customer Info" stage. Swatches go out as a Drive folder link
// because the brochure PDFs exceed the 500 KiB MMS cap. Personalized with
// the first name. Never throws - a text failure must not break the email flow.
function dp_sendCustomerInfoSms_(firstName, phone, missingItems, opt_sheet, opt_row) {
  var e164 = dp_normalisePhone_(phone);
  if (!e164) {
    Logger.log('dp_sendCustomerInfoSms_: no valid phone, SMS skipped.');
    return;
  }

  var who = firstName || 'there';
  var swatchLink = 'https://drive.google.com/drive/folders/1v63q5JkdYre2fyqNr5GNmbdT3vE3NhDT';

  // Build the body with line breaks so it reads as a list, not one blob.
  var msg = 'Hi ' + who + ', it\'s Gino from Walker Awning!';

  if (missingItems && missingItems.length > 0) {
    msg += '\n\nTo move your quote forward, could you reply with:';
    for (var i = 0; i < missingItems.length; i++) {
      // Trim a trailing "?" so bulleted questions don't end up with "??".
      var item = String(missingItems[i]).replace(/\?+\s*$/, '');
      msg += '\n\u2022 ' + item;
    }
  }

  msg += '\n\nHere are our fabric swatches \u2014 just open the folder that fits your project:' +
         '\n' + swatchLink +
         '\n\nThank you!';

  // Open the shared editable review dialog (single path for all texts).
  dp_openReviewDialog_(msg, e164, who, opt_sheet, opt_row);
}
// ----- EDITABLE REVIEW DIALOG (shows the text body, lets you edit, then send) -----
// Built as an array of lines (joined) to keep quotes clean and avoid template
// literals, per Apps Script compatibility. <?= ?> values are injected server-side.
var DP_SMS_DIALOG_HTML = [
  '<!DOCTYPE html>',
  '<html><head><base target="_top"><style>',
  'body{font-family:Arial,sans-serif;margin:0;padding:16px;color:#222;}',
  '.to{font-size:13px;color:#555;margin-bottom:8px;}',
  '.to b{color:#222;}',
  'textarea{width:100%;box-sizing:border-box;height:240px;font-size:14px;',
  'font-family:Arial,sans-serif;padding:8px;border:1px solid #bbb;border-radius:6px;resize:vertical;}',
  '.muted{font-size:12px;color:#888;margin-top:6px;}',
  '.err{color:#c5221f;font-size:13px;margin-top:10px;display:none;}',
  '.row{margin-top:12px;text-align:right;}',
  'button{font-size:14px;padding:8px 16px;border-radius:6px;border:1px solid #bbb;cursor:pointer;margin-left:8px;}',
  '.send{background:#1a73e8;color:#fff;border-color:#1a73e8;}',
  '.cancel{background:#f1f3f4;}',
  '</style></head><body>',
  '<div class="to">Text to <b><?= who ?></b> &middot; <?= phone ?></div>',
  '<textarea id="msg"><?= message ?></textarea>',
  '<input type="hidden" id="phone" value="<?= phone ?>">',
  '<input type="hidden" id="markSheet" value="<?= markSheet ?>">',
  '<input type="hidden" id="markRow" value="<?= markRow ?>">',
  '<div class="muted">Edit the message above if needed, then tap Send.</div>',
  '<div class="err" id="err"></div>',
  '<div class="row">',
  '<button class="cancel" onclick="google.script.host.close()">Cancel</button>',
  '<button class="send" id="sendBtn" onclick="doSend()">Send</button>',
  '</div>',
  '<script>',
  'function doSend(){',
  '  var btn=document.getElementById("sendBtn");',
  '  btn.disabled=true;btn.textContent="Sending...";',
  '  var text=document.getElementById("msg").value;',
  '  var phone=document.getElementById("phone").value;',
  '  var ms=document.getElementById("markSheet").value;',
  '  var mr=document.getElementById("markRow").value;',
  '  google.script.run',
  '    .withSuccessHandler(function(){google.script.host.close();})',
  '    .withFailureHandler(function(e){',
  '      btn.disabled=false;btn.textContent="Send";',
  '      var d=document.getElementById("err");',
  '      d.style.display="block";',
  '      d.textContent="Failed: "+((e&&e.message)?e.message:e);',
  '    })',
  '    .dp_sendSmsFromDialog(phone,text,ms,mr);',
  '}',
  '</script>',
  '</body></html>'
].join('\n');

// Called BY THE DIALOG via google.script.run. Must be PUBLIC (no trailing
// underscore) - google.script.run cannot call trailing-underscore functions.
function dp_sendSmsFromDialog(phone, text, markSheet, markRow) {
  var e164 = dp_normalisePhone_(phone);
  if (!e164)              throw new Error('Invalid phone number: ' + phone);
  if (!text || !text.trim()) throw new Error('Message is empty.');

  try {
    dp_sendSms(e164, text);
  } catch (err) {
    // Mark failure in D (if we were given a target), then rethrow so the
    // dialog shows the red error and stays open.
    if (markSheet && markRow) dp_markStageCellByName_(markSheet, Number(markRow), false);
    throw err;
  }

  // Success → mark D (only when a target was passed; Customer Info passes none yet).
  if (markSheet && markRow) dp_markStageCellByName_(markSheet, Number(markRow), true);
  try {
    if (markSheet && markRow) {
      var logSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(markSheet);
      if (logSh) {
        var logStage = String(logSh.getRange(Number(markRow), 4).getDisplayValue() || '').replace(/[\u2705\u274C]/g, '').trim();
        al_logActivity_(markSheet, 'Automation', String(logSh.getRange(Number(markRow), DP_TX.NAME).getDisplayValue() || ''), String(logSh.getRange(Number(markRow), 6).getDisplayValue() || ''), 'Text', logStage, '', null, '', phone);
      }
    } else {
      al_logActivity_('', 'Automation', '', '', 'Text', 'Info request', '', null, '', phone);
    }
  } catch (_) {}
  SpreadsheetApp.getActiveSpreadsheet().toast('Text sent', 'SMS Sent', 4);
  return 'sent';
}

// Same as dp_markStageCell_ but resolves the sheet by name (the dialog only
// has the name as a string, not a live Sheet object).
function dp_markStageCellByName_(sheetName, row, ok) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return;
    dp_markStageCell_(sheet, row, ok);
  } catch (e) {}
}
// =============================================================
// STAGE-TRIGGERED TEXTS — all route through the review dialog.
// Single dispatcher: master onEdit calls dp_dispatchStageText_.
// Nothing auto-sends; no batching; one row at a time.
// =============================================================

// Column landmarks (1-based) used by the text builders.
var DP_TX = {
  NAME:  5,   // E - Customer Name (full name; first word used)
  PHONE: 8,   // H - Phone
  EMAIL: 9,   // I - Email
  ADDR:  10   // J - Project Address
};

var DP_REVIEW_URL = 'https://g.co/kgs/sb6TX7T';

// First name from the full-name cell (column E).
function dp_firstName_(sheet, row) {
  var full = String(sheet.getRange(row, DP_TX.NAME).getDisplayValue() || '').trim();
  if (!full) return 'there';
  return full.split(/\s+/)[0];
}

// MASTER ENTRY POINT for stage-triggered texts.
// Wire this into your master onEdit handler (see notes). Matches the stage
// (case-insensitive) to the right script and opens the review dialog.
// Returns true if a stage text was handled, false otherwise.
function dp_dispatchStageText_(sheet, row, stageRaw) {
  var sheetName = sheet.getName();

  // Read the ACTUAL cell contents, not the passed-in event value. A paste of
  // "FU2! ✅" can hand this function a value missing the emoji, so trust the
  // cell itself as the source of truth.
  var cellText = String(sheet.getRange(row, 4).getDisplayValue() || '');

  // A ✅/❌ anywhere in the cell means "already handled" - either from a prior
  // send, or pasted in by hand as a manual indicator. Never fire in that case.
  if (/[\u2705\u274C]/.test(cellText)) return false;

  var stage = cellText.trim().toLowerCase();

  // Trailing "!" = instant send, skip the review dialog. Strip it for matching
  // so "fu1!" still matches the "fu1" branch below.
  var bang = /!+\s*$/.test(stage);
  if (bang) stage = stage.replace(/!+\s*$/, '').trim();

  // Cheap guard: only these sheets have text stages. Bail before reading cells.
  if (sheetName !== 'Leads' && sheetName !== 'F/U' && sheetName !== 'Awarded' && sheetName !== 'Heaven') return false;

  // FU stages fire on both Leads and F/U (identical behavior).
  var isFuSheet = (sheetName === 'Leads' || sheetName === 'F/U');

  var first = dp_firstName_(sheet, row);
  var addr  = String(sheet.getRange(row, DP_TX.ADDR).getDisplayValue() || '').trim();
  var phone = sheet.getRange(row, DP_TX.PHONE).getDisplayValue();

  var body = null;

  if (isFuSheet && stage === 'fu1') {
    body = 'Hello ' + first + ',\n' +
           'Did you have any questions for me or are you interested in moving forward with us?\n' +
           'Subject property: \n' + addr + '\n\n' +
           'Gino\nWalker Awning\nGino@WalkerAwning.com';

  } else if (isFuSheet && stage === 'fu2') {
    body = 'Hi ' + first + ',\n' +
           'Just checking to see if you\'re still interested in working with us?\n' +
           'Subject property: \n' + addr + '\n\n' +
           'Gino\nWalker Awning\nGino@WalkerAwning.com';

  } else if (isFuSheet && stage === 'fu3') {
    body = 'Hello ' + first + ',\n' +
           'Are you still interested in our awning services?\n' +
           'If not, text "Stop".\n' +
           'Subject property: \n' + addr + '\n\n' +
           'Gino\nWalker Awning\nGino@WalkerAwning.com';

  } else if (sheetName === 'Awarded' && stage === 'scheduled') {
    body = 'Hello ' + first + ',\n' +
           'Your awning is ready!\n' +
           'Are there any installation\n' +
           'restrictions during business hours?\n\n' +
           'Gino\nWalker Awning\nGino@WalkerAwning.com';

  } else if (sheetName === 'Heaven' && stage === 'request review') {
    body = 'Hello ' + first + ',\n' +
           'Thank you for working with us and I hope we were able to meet your expectations!\n' +
           'Would you kindly take a moment to give us a review.\n' +
           DP_REVIEW_URL + '\n\n' +
           'Gino\nWalker Awning\nGino@WalkerAwning.com';
  }

  if (body === null) return false;  // not a text stage on this sheet

  var e164 = dp_normalisePhone_(phone);
  if (!e164) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No valid phone in column H - text skipped.', 'SMS Skipped', 5);
    return true;  // it WAS a text stage; we just couldn't send
  }

  if (bang) {
    // Instant send - no preview, no second chance. Failures stay visible.
    try {
      dp_sendSms(e164, body);
      dp_markStageCell_(sheet, row, true);
      try { al_logActivity_(sheetName, 'Automation', String(sheet.getRange(row, DP_TX.NAME).getDisplayValue() || ''), String(sheet.getRange(row, 6).getDisplayValue() || ''), 'Text', stage.toUpperCase(), '', null, '', e164); } catch (_) {}
      SpreadsheetApp.getActiveSpreadsheet().toast('Text sent to ' + first, 'SMS Sent', 4);
    } catch (err) {
      dp_markStageCell_(sheet, row, false, err.message);
      SpreadsheetApp.getActiveSpreadsheet().toast('Text FAILED (hover the D cell for details): ' + err.message, 'SMS Error', 10);
    }
    return true;
  }

  // Dialog path: pass sheet+row so the mark can be written after an actual send.
  dp_openReviewDialog_(body, e164, first, sheet, row);
  return true;
}

// Append a status mark to column D (Stage) without re-firing onEdit.
// ok=true → " ✅", ok=false → " ❌". Strips any prior mark first.
function dp_markStageCell_(sheet, row, ok, opt_reason) {
  try {
    var cell = sheet.getRange(row, 4); // D
    var base = String(cell.getDisplayValue() || '').replace(/[\u2705\u274C]/g, '').replace(/\s+$/, '');
    cell.setValue(base + (ok ? ' \u2705' : ' \u274C'));

    // On failure, stash the reason as a hover note. On success, clear any old note.
    if (ok) {
      cell.clearNote();
    } else if (opt_reason) {
      var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d h:mm a');
      cell.setNote('Text failed ' + stamp + '\n\n' + String(opt_reason));
    }
  } catch (e) {
    // Marking is cosmetic - never let it break the send flow.
  }
}

// Opens the editable review dialog for any pre-built body.
// (Same dialog the Customer Info flow uses - shared, not duplicated.)
function dp_openReviewDialog_(message, e164, who, opt_sheet, opt_row) {
  var tmpl = HtmlService.createTemplate(DP_SMS_DIALOG_HTML);
  tmpl.message  = message;
  tmpl.phone    = e164;
  tmpl.who      = who;
  // Stash sheet name + row so the send handler can mark column D afterward.
  tmpl.markSheet = (opt_sheet && opt_row) ? opt_sheet.getName() : '';
  tmpl.markRow   = (opt_sheet && opt_row) ? opt_row : 0;
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(460).setHeight(460),
    'Review & Send Text'
  );
}
// ----- MASTER-HANDLER ENTRY POINT FOR STAGE-TRIGGERED TEXTS -----
// Wired into masterOnEditHandler_ (Menus.gs) as handler #6. Reacts ONLY to
// single-cell edits of column D (Stage). Hands off to dp_dispatchStageText_,
// which matches sheet+stage and opens the review dialog. Never moves/sorts.
function dp_handleEditText_(e) {
  if (!e || !e.range) return;
  var range = e.range;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return; // ignore multi-cell
  if (range.getRow() === 1) return;        // ignore header
  if (range.getColumn() !== 4) return;     // column D = Stage only
  dp_dispatchStageText_(range.getSheet(), range.getRow(), range.getValue());
}