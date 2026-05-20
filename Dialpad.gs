// Dialpad.gs
// version1 [04/08-02:45PM] by Claude Sonnet 4.6

// ─── CONFIG ────────────────────────────────────────────────────────────────
var DIALPAD_CONFIG = {
  PHONE_COL:       8,
  STAGE_COL:       4,
  HEADER_ROW:      1,
  SHEETS:          ['Leads', 'F/U', 'Awarded', 'Heaven', 'Re-cover'],
  API_KEY_PROP:    'Dialpad_Walker_Awning',
  USER_ID_PROP:    'DIALPAD_USER_ID',
  FROM_NUM_PROP:   'DIALPAD_FROM_NUMBER',
  API_BASE:        'https://dialpad.com/api/v2',
  URL_SCHEME:      'dialpad://'
};

// ─── API HELPER ────────────────────────────────────────────────────────────

function dp_getApiKey_() {
  var key = PropertiesService.getScriptProperties().getProperty(DIALPAD_CONFIG.API_KEY_PROP);
  if (!key) throw new Error('Dialpad API key not found in Script Properties under "' + DIALPAD_CONFIG.API_KEY_PROP + '"');
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
    throw new Error('Dialpad API error ' + code + ': ' + body);
  }
  return JSON.parse(body);
}

// ─── PHONE HELPERS ─────────────────────────────────────────────────────────

function dp_normalisePhone_(raw) {
  var digits = String(raw).replace(/\D/g, '');
  if (digits.length === 10) digits = '1' + digits;
  if (digits.length < 11)   return null;
  return '+' + digits;
}

function dp_linkCell_(cell) {
  var val     = cell.getValue();
  var formula = cell.getFormula();

  // If already a Dialpad hyperlink, extract the display text from it
  if (formula && formula.indexOf('HYPERLINK') !== -1 && (formula.indexOf('dialpad.com') !== -1 || formula.indexOf('dialpad://') !== -1)) {
    var match = formula.match(/,"(.+?)"\)/);
    val = match ? match[1] : val;
  } else if (formula) {
    return; // Don't touch non-Dialpad formulas
  }

  if (!val) return;
  var e164 = dp_normalisePhone_(val);
  if (!e164) return;
  var url = DIALPAD_CONFIG.URL_SCHEME + e164;
  cell.setFormula('=HYPERLINK("' + url + '","' + val + '")');
}
// ─── SMS SENDING (API) ─────────────────────────────────────────────────────

/**
 * Send an SMS via Dialpad API.
 * @param {string} toNumber  - raw phone string (will be normalised)
 * @param {string} message   - text body
 */
function dp_sendSms(toNumber, message) {
  var e164 = dp_normalisePhone_(toNumber);
  if (!e164) throw new Error('Invalid phone number: ' + toNumber);
  var props      = PropertiesService.getScriptProperties();
  var userId     = props.getProperty(DIALPAD_CONFIG.USER_ID_PROP);
  var fromNumber = props.getProperty(DIALPAD_CONFIG.FROM_NUM_PROP);
  if (!userId || !fromNumber) throw new Error('DIALPAD_USER_ID or DIALPAD_FROM_NUMBER missing from Script Properties.');
  return dp_apiRequest_('post', '/sms', {
    to_numbers:      [e164],
    text:            message,
    user_id:         userId,
    from_number:     fromNumber
  });
}

/**
 * Send SMS to the customer in the currently selected row.
 * Prompts for a message first.
 */
function dp_sendCustomMessage_() {
  var ui      = SpreadsheetApp.getUi();
  var sheet   = SpreadsheetApp.getActiveSheet();
  var row     = sheet.getActiveCell().getRow();

  if (row <= DIALPAD_CONFIG.HEADER_ROW) {
    ui.alert('Please select a data row first.'); return;
  }
  if (DIALPAD_CONFIG.SHEETS.indexOf(sheet.getName()) === -1) {
    ui.alert('Please run this from one of the project sheets.'); return;
  }

  var phone = sheet.getRange(row, DIALPAD_CONFIG.PHONE_COL).getValue();
  var name  = sheet.getRange(row, 5).getValue(); // col E = Customer Name

  if (!phone) { ui.alert('No phone number in this row.'); return; }

  var result = ui.prompt(
    'Send SMS to ' + name,
    'Enter your message:',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  try {
    dp_sendSms(phone, result.getResponseText());
    ui.alert('✅ Message sent to ' + name + ' (' + phone + ')');
  } catch (err) {
    ui.alert('❌ Failed to send: ' + err.message);
  }
}

// ─── PRESET MESSAGE SENDERS ────────────────────────────────────────────────

function dp_getRowContext_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveCell().getRow();
  if (row <= DIALPAD_CONFIG.HEADER_ROW) throw new Error('Please select a data row first.');
  if (DIALPAD_CONFIG.SHEETS.indexOf(sheet.getName()) === -1) throw new Error('Wrong sheet.');
  return {
    name:    sheet.getRange(row, 5).getValue(),   // col E
    phone:   sheet.getRange(row, 8).getValue(),   // col H
    address: sheet.getRange(row, 10).getValue(),  // col J
    price:   sheet.getRange(row, 14).getValue()   // col N
  };
}

function dp_sendSwatches_() {
  try {
    var ctx = dp_getRowContext_();
    dp_sendSms(ctx.phone,
      'Hi ' + ctx.name + ', this is Walker Awning! We\'d love to send you some fabric swatches to help you choose the perfect look for your project. Just confirm your address and we\'ll get them out to you.');
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Swatches message sent to ' + ctx.name, 'SMS Sent', 4);
  } catch(err) { SpreadsheetApp.getUi().alert('❌ ' + err.message); }
}

function dp_requestPhotos_() {
  try {
    var ctx = dp_getRowContext_();
    dp_sendSms(ctx.phone,
      'Hi ' + ctx.name + ', it\'s Walker Awning! Could you send us a few photos of the area where the awning will be installed? It helps us give you the most accurate quote. Thank you!');
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Photo request sent to ' + ctx.name, 'SMS Sent', 4);
  } catch(err) { SpreadsheetApp.getUi().alert('❌ ' + err.message); }
}

function dp_sendProposalLink_() {
  try {
    var ctx   = dp_getRowContext_();
    var sheet = SpreadsheetApp.getActiveSheet();
    var row   = sheet.getActiveCell().getRow();
    var qbUrl = sheet.getRange(row, 16).getValue(); // col P = QB URL
    if (!qbUrl) { SpreadsheetApp.getUi().alert('No QuickBooks URL found in column P for this row.'); return; }
    dp_sendSms(ctx.phone,
      'Hi ' + ctx.name + ', your proposal from Walker Awning is ready! You can view it here: ' + qbUrl + ' — feel free to call or text us with any questions.');
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Proposal link sent to ' + ctx.name, 'SMS Sent', 4);
  } catch(err) { SpreadsheetApp.getUi().alert('❌ ' + err.message); }
}

// ─── CALL LOG LOOKUP ───────────────────────────────────────────────────────

/**
 * Pull recent call/SMS history for the selected row's phone number.
 */
function dp_viewContactHistory_() {
  try {
    var ctx  = dp_getRowContext_();
    var e164 = dp_normalisePhone_(ctx.phone);
    if (!e164) { SpreadsheetApp.getUi().alert('Invalid phone number.'); return; }

    var result = dp_apiRequest_('get', '/contacts?phone=' + encodeURIComponent(e164));
    SpreadsheetApp.getUi().alert(
      'Contact History: ' + ctx.name,
      JSON.stringify(result, null, 2).substring(0, 1000),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(err) { SpreadsheetApp.getUi().alert('❌ ' + err.message); }
}

// ─── BULK HYPERLINK APPLY ──────────────────────────────────────────────────

function dp_applyAllDialpadLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  DIALPAD_CONFIG.SHEETS.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow <= DIALPAD_CONFIG.HEADER_ROW) return;
    var numRows = lastRow - DIALPAD_CONFIG.HEADER_ROW;
    var cells   = sheet.getRange(DIALPAD_CONFIG.HEADER_ROW + 1, DIALPAD_CONFIG.PHONE_COL, numRows, 1);
    for (var r = 1; r <= numRows; r++) {
      dp_linkCell_(cells.getCell(r, 1));
    }
  });
  SpreadsheetApp.getUi().alert('✅ Dialpad links applied to all sheets.');
}

// ─── ON-EDIT (auto-link new phone numbers as typed) ────────────────────────

function dp_handleEdit_(e) {
  if (!e) return;
  var range = e.range;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
  if (range.getRow() <= DIALPAD_CONFIG.HEADER_ROW)             return;

  var sheet     = range.getSheet();
  var sheetName = sheet.getName();
  var row       = range.getRow();
  var col       = range.getColumn();

  if (DIALPAD_CONFIG.SHEETS.indexOf(sheetName) === -1) return;

  // Auto-link phone numbers as typed
  if (col === DIALPAD_CONFIG.PHONE_COL) {
    dp_linkCell_(range);
  }

  // Text Customer SMS prompt (col D = 4, Leads sheet only)
  Logger.log('dp_handleEdit_ col=' + col + ' sheet=' + sheetName + ' value=' + e.value);
  if (col === 4 && sheetName === 'Leads' && e.value === 'Text Customer') {
    Logger.log('Text Customer detected - firing SMS prompt');
    dp_promptTextCustomerSms_(sheet, row);
  }
}

// ─── TRIGGER INSTALLER ─────────────────────────────────────────────────────

function dp_installTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dp_handleEdit_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dp_handleEdit_').forSpreadsheet(ss).onEdit().create();
  Logger.log('Dialpad onEdit trigger installed.');
}

// ─── Text Customer SMS PROMPT ──────────────────────────────────────────────

function dp_promptTextCustomerSms_(sheet, row) {
  var name      = sheet.getRange(row, 5).getDisplayValue();
  var phoneCell = sheet.getRange(row, 8);
  var phone     = phoneCell.getDisplayValue() || String(phoneCell.getValue());

  Logger.log('dp_promptTextCustomerSms_ row=' + row + ' name=' + name + ' phone=' + phone);

  if (!phone || phone.trim() === '') {
    Logger.log('No phone number found - skipping');
    return;
  }

  var msg =
    'Hi ' + name + ', this is Gino from Walker Awning! ' +
    'I\'d love to learn more about your project and get you a quote. ' +
    'When is a good time to connect?';

  try {
    dp_sendSms(phone, msg);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '✅ SMS sent to ' + name + ' (' + phone + ')',
      'SMS Sent', 5
    );
  } catch (err) {
    Logger.log('SMS failed: ' + err.message);
  }
}

// ─── CONNECTION TEST ───────────────────────────────────────────────────────
function dp_testSend() {
  try {
    var result = dp_sendSms('+19548265915', 'Test from Walker Awning automation');
    SpreadsheetApp.getUi().alert('✅ API response:\n\n' + JSON.stringify(result, null, 2).substring(0, 800));
  } catch(err) {
    SpreadsheetApp.getUi().alert('❌ ' + err.message);
  }
}
function dp_removeAllDialpadLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Leads', 'F/U', 'Awarded', 'Heaven', 'Re-cover'].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    var cells = sheet.getRange(2, 8, lastRow - 1, 1);
    for (var r = 1; r <= lastRow - 1; r++) {
      var cell = cells.getCell(r, 1);
      var formula = cell.getFormula();
      if (formula && formula.indexOf('HYPERLINK') !== -1 && formula.indexOf('dialpad') !== -1) {
        var match = formula.match(/,"(.+?)"\)/);
        var display = match ? match[1] : cell.getValue();
        cell.setValue(display);
      }
    }
  });
  SpreadsheetApp.getUi().alert('✅ Dialpad hyperlinks removed from all sheets.');
}
function dp_debugPhoneCell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads');
  var row   = sheet.getActiveCell().getRow();
  var cell  = sheet.getRange(row, 8);
  SpreadsheetApp.getUi().alert(
    'Row: '          + row                  + '\n' +
    'getValue: '     + cell.getValue()       + '\n' +
    'getDisplayValue: ' + cell.getDisplayValue() + '\n' +
    'getFormula: '   + cell.getFormula()
  );
}