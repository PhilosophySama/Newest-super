// ============================================================
// Dialpad
// version2 [04/08-12:00AM] by Claude Sonnet 4.6
// Handles outbound SMS via Dialpad API v2
// ============================================================

var DIALPAD_CONFIG = {
  API_BASE:     'https://dialpad.com/api/v2',
  FROM_NUMBER:  '+19542710686',
  API_KEY_PROP: 'DIALPAD_API_KEY'
};

function handleEditDialpadPhone_(e) {
  if (!e || !e.range) return;
  var sheet     = e.range.getSheet();
  var col       = e.range.getColumn();
  var row       = e.range.getRow();
  var sheetName = sheet.getName();
  var allowed   = ['Leads', 'F/U', 'Awarded'];
  if (allowed.indexOf(sheetName) === -1) return;
  if (col !== 8) return;
  if (row <= 1) return;
  var raw = e.range.getValue();
  if (!raw) return;
  var digits = String(raw).replace(/\D/g, '');
  if (digits.length === 10) digits = '1' + digits;
  if (digits.length !== 11 || digits.charAt(0) !== '1') return;
  var display = '(' + digits.substr(1,3) + ') ' + digits.substr(4,3) + '-' + digits.substr(7,4);
  var url = 'https://dialpad.com/main/sms/%2B' + digits;
  e.range.setFormula('=HYPERLINK("' + url + '","' + display + '")');
}

function dp_formatPhone_(raw) {
  if (!raw) return null;
  var digits = String(raw).replace(/\D/g, '');
  if (digits.length === 10) digits = '1' + digits;
  if (digits.length === 11 && digits.charAt(0) === '1') return '+' + digits;
  return null;
}

function dp_sendSms_(toNumber, messageText) {
  var apiKey = PropertiesService.getScriptProperties().getProperty(DIALPAD_CONFIG.API_KEY_PROP);
  if (!apiKey) { SpreadsheetApp.getUi().alert('DIALPAD_API_KEY not found in Script Properties.'); return false; }
  var formattedTo = dp_formatPhone_(toNumber);
  if (!formattedTo) { SpreadsheetApp.getUi().alert('Could not parse phone number: ' + toNumber); return false; }
  var payload = { from_number: DIALPAD_CONFIG.FROM_NUMBER, to_numbers: [formattedTo], text: messageText };
  var options = {
    method: 'post', contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + apiKey, 'Accept': 'application/json' },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch(DIALPAD_CONFIG.API_BASE + '/sms', options);
    var code = response.getResponseCode();
    var body = JSON.parse(response.getContentText());
    if (code === 200 || code === 201) return true;
    SpreadsheetApp.getUi().alert('Dialpad API error ' + code + ':\n' + JSON.stringify(body, null, 2));
    return false;
  } catch (e) {
    SpreadsheetApp.getUi().alert('Request failed: ' + e.message);
    return false;
  }
}

function dp_getRowData_() {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row    = sheet.getActiveCell().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert('Please click on a customer row first.'); return null; }
  var values = sheet.getRange(row, 1, 1, 42).getValues()[0];
  var data = {
    row: row, stage: values[3], name: values[4], phone: values[7],
    email: values[8], address: values[9], quotePrice: values[13],
    qbLink: values[15], folderLink: ''
  };
  var fCell = sheet.getRange(row, 6);
  var formula = fCell.getFormula();
  if (formula && formula.toUpperCase().indexOf('HYPERLINK') !== -1) {
    var match = formula.match(/HYPERLINK\("([^"]+)"/i);
    data.folderLink = match ? match[1] : '';
  } else {
    var rt = fCell.getRichTextValue();
    if (rt) {
      var runs = rt.getRuns();
      for (var i = 0; i < runs.length; i++) {
        var url = runs[i].getLinkUrl();
        if (url) { data.folderLink = url; break; }
      }
    }
    if (!data.folderLink) data.folderLink = fCell.getValue();
  }
  if (!data.phone) { SpreadsheetApp.getUi().alert('No phone number found in col H.'); return null; }
  if (!data.name)  { SpreadsheetApp.getUi().alert('No customer name found in col E.'); return null; }
  return data;
}

function dp_sendSwatches_() {
  var customer = dp_getRowData_(); if (!customer) return;
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Send Swatches to ' + customer.name, 'Paste the Google Drive swatch link:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var link = response.getResponseText().trim();
  if (!link) { ui.alert('No link entered.'); return; }
  var message = 'Hi ' + customer.name.split(' ')[0] + '! Here are some fabric swatch options for your awning project: ' + link + '\n\nLet us know if you have any questions! \u2013 Walker Awning';
  if (ui.alert('Confirm Send', 'Sending to ' + customer.phone + ':\n\n' + message, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  if (dp_sendSms_(customer.phone, message)) ui.alert('\u2705 Swatches sent to ' + customer.name + '!');
}

function dp_requestPhotos_() {
  var customer = dp_getRowData_(); if (!customer) return;
  var ui = SpreadsheetApp.getUi();
  var message = 'Hi ' + customer.name.split(' ')[0] + '! Could you please send us some photos of your existing awning/space? This will help us finalize your project. Thank you! \u2013 Walker Awning';
  if (customer.folderLink && customer.folderLink.indexOf('http') === 0) message += '\n\nYou can upload them here: ' + customer.folderLink;
  if (ui.alert('Confirm Send', 'Sending to ' + customer.phone + ':\n\n' + message, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  if (dp_sendSms_(customer.phone, message)) ui.alert('\u2705 Photo request sent to ' + customer.name + '!');
}

function dp_sendProposalLink_() {
  var customer = dp_getRowData_(); if (!customer) return;
  var ui = SpreadsheetApp.getUi();
  var proposalLink = '';
  try {
    var threads = GmailApp.search('in:sent subject:"awning proposal" ' + customer.name.split(' ')[0], 0, 5);
    for (var t = 0; t < threads.length; t++) {
      var msgs = threads[t].getMessages();
      for (var m = 0; m < msgs.length; m++) {
        var match = msgs[m].getBody().match(/href="(https:\/\/[^"]+)"/i);
        if (match) { proposalLink = match[1]; break; }
      }
      if (proposalLink) break;
    }
  } catch(e) { Logger.log('Gmail search error: ' + e.message); }
  var response = ui.prompt('Proposal Link for ' + customer.name, proposalLink ? 'Found link. Confirm or replace:' : 'No link found. Paste manually:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var finalLink = response.getResponseText().trim() || proposalLink;
  if (!finalLink) { ui.alert('No link provided.'); return; }
  var message = 'Hi ' + customer.name.split(' ')[0] + '! Here is your awning proposal from Walker Awning: ' + finalLink + '\n\nPlease review and feel free to reach out with any questions!';
  if (ui.alert('Confirm Send', 'Sending to ' + customer.phone + ':\n\n' + message, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  if (dp_sendSms_(customer.phone, message)) ui.alert('\u2705 Proposal link sent to ' + customer.name + '!');
}

function dp_sendCustomMessage_() {
  var customer = dp_getRowData_(); if (!customer) return;
  var ui = SpreadsheetApp.getUi();
  var templateChoice = ui.alert('Message Type for ' + customer.name, 'Choose a starting template:\n\n[OK] = Walkthrough / Site Visit\n[Cancel] = Write my own', ui.ButtonSet.OK_CANCEL);
  var defaultText = templateChoice === ui.Button.OK
    ? 'Hi ' + customer.name.split(' ')[0] + '! We\'d like to schedule a walkthrough for your awning project. What days/times work best for you? \u2013 Walker Awning'
    : '';
  var response = ui.prompt('Message to ' + customer.name + ' (' + customer.phone + ')', defaultText || 'Write your message here:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var message = response.getResponseText().trim();
  if (!message) { ui.alert('No message entered.'); return; }
  if (ui.alert('Confirm Send', 'Sending to ' + customer.phone + ':\n\n' + message, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  if (dp_sendSms_(customer.phone, message)) ui.alert('\u2705 Message sent to ' + customer.name + '!');
}

function dp_testApiCall_() {
  var result = dp_sendSms_('+19542710686', 'Test from Apps Script - ignore');
  Logger.log('Result: ' + result);
}