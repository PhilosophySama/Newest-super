// ============================================================
// Dialpad
// version1 [04/06-03:45PM] by Claude Sonnet 4.6
// Handles outbound SMS via Dialpad API v2
// ============================================================

var DIALPAD_CONFIG = {
  API_BASE:    'https://dialpad.com/api/v2',
  FROM_NUMBER: '+19542710686',           // Gino's Dialpad number
  API_KEY_PROP: 'DIALPAD_API_KEY'        // Stored in Script Properties
};

// ── onEdit: auto-hyperlink phone numbers in col H ──────────
function handleEditDialpadPhone_(e) {
  if (!e || !e.range) return;

  var sheet     = e.range.getSheet();
  var col       = e.range.getColumn();
  var row       = e.range.getRow();
  var sheetName = sheet.getName();

  // Only act on col H, data rows, allowed sheets
  var allowed = ['Leads', 'F/U', 'Awarded'];
  if (allowed.indexOf(sheetName) === -1) return;
  if (col !== 8) return;  // col H
  if (row <= 1) return;   // skip header

  var raw = e.range.getValue();
  if (!raw) return;

  // Strip to digits only
  var digits = String(raw).replace(/\D/g, '');
  if (digits.length === 10) digits = '1' + digits;
  if (digits.length !== 11 || digits.charAt(0) !== '1') return;

  // Format display as (XXX) XXX-XXXX
  var display = '(' + digits.substr(1,3) + ') ' + digits.substr(4,3) + '-' + digits.substr(7,4);

  // Dialpad SMS URL - opens existing convo or starts new one
  var url = 'https://dialpad.com/main/sms/%2B' + digits;

  // Write HYPERLINK formula
  e.range.setFormula('=HYPERLINK("' + url + '","' + display + '")');
}

// ── Core: format any phone string to E164 ──────────────────
function dp_formatPhone_(raw) {
  if (!raw) return null;
  var digits = String(raw).replace(/\D/g, '');
  if (digits.length === 10) digits = '1' + digits;
  if (digits.length === 11 && digits.charAt(0) === '1') {
    return '+' + digits;
  }
  return null; // unrecognizable format
}

// ── Core: send SMS via Dialpad API ─────────────────────────
function dp_sendSms_(toNumber, messageText) {
  var apiKey = PropertiesService.getScriptProperties()
                 .getProperty(DIALPAD_CONFIG.API_KEY_PROP);
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('DIALPAD_API_KEY not found in Script Properties.');
    return false;
  }

  var formattedTo = dp_formatPhone_(toNumber);
  if (!formattedTo) {
    SpreadsheetApp.getUi().alert('Could not parse phone number: ' + toNumber);
    return false;
  }

  var payload = {
    from_number:  DIALPAD_CONFIG.FROM_NUMBER,
    to_numbers:   [formattedTo],
    text:         messageText
  };

  var options = {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Accept':        'application/json'
    },
    payload:          JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(DIALPAD_CONFIG.API_BASE + '/sms', options);
    var code     = response.getResponseCode();
    var body     = JSON.parse(response.getContentText());

    if (code === 200 || code === 201) {
      return true;
    } else {
      SpreadsheetApp.getUi().alert(
        'Dialpad API error ' + code + ':\n' + JSON.stringify(body, null, 2)
      );
      return false;
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Request failed: ' + e.message);
    return false;
  }
}

// ── Helper: read active row customer data ──────────────────
function dp_getRowData_() {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row    = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please click on a customer row first.');
    return null;
  }

  var values = sheet.getRange(row, 1, 1, 42).getValues()[0];

  // Column index (0-based): D=3, E=4, F=5, H=7, I=8, J=9, N=13, P=15
  var data = {
    row:          row,
    stage:        values[3],   // col D
    name:         values[4],   // col E
    phone:        values[7],   // col H
    email:        values[8],   // col I
    address:      values[9],   // col J
    quotePrice:   values[13],  // col N
    qbLink:       values[15],  // col P
    folderLink:   ''           // col F - need rich text extraction
  };

  // Extract folder hyperlink from col F (rich text or HYPERLINK formula)
  var fCell   = sheet.getRange(row, 6);
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

  if (!data.phone) {
    SpreadsheetApp.getUi().alert('No phone number found in col H for this row.');
    return null;
  }
  if (!data.name) {
    SpreadsheetApp.getUi().alert('No customer name found in col E for this row.');
    return null;
  }

  return data;
}

// ── Use Case 1: Send Swatches ───────────────────────────────
function dp_sendSwatches_() {
  var customer = dp_getRowData_();
  if (!customer) return;

  var ui       = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Send Swatches to ' + customer.name,
    'Paste the Google Drive swatch link to send:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;
  var link = response.getResponseText().trim();
  if (!link) { ui.alert('No link entered.'); return; }

  var message = 'Hi ' + customer.name.split(' ')[0] + '! Here are some fabric swatch options for your awning project: ' + link + '\n\nLet us know if you have any questions! – Walker Awning';

  var confirm = ui.alert(
    'Confirm Send',
    'Sending to ' + customer.phone + ':\n\n' + message,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  var sent = dp_sendSms_(customer.phone, message);
  if (sent) ui.alert('✅ Swatches sent to ' + customer.name + '!');
}

// ── Use Case 2: Request Customer Photos ────────────────────
function dp_requestPhotos_() {
  var customer = dp_getRowData_();
  if (!customer) return;

  var ui      = SpreadsheetApp.getUi();
  var message = 'Hi ' + customer.name.split(' ')[0] + '! Could you please send us some photos of your existing awning/space? This will help us finalize your project details. Thank you! – Walker Awning';

  // Append their folder link if available
  if (customer.folderLink && customer.folderLink.indexOf('http') === 0) {
    message += '\n\nYou can also upload them here: ' + customer.folderLink;
  }

  var confirm = ui.alert(
    'Confirm Send',
    'Sending to ' + customer.phone + ':\n\n' + message,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  var sent = dp_sendSms_(customer.phone, message);
  if (sent) ui.alert('✅ Photo request sent to ' + customer.name + '!');
}

// ── Use Case 3: Send Proposal Link ─────────────────────────
function dp_sendProposalLink_() {
  var customer = dp_getRowData_();
  if (!customer) return;

  var ui = SpreadsheetApp.getUi();

  // Search Sent mail for proposal email
  var proposalLink = '';
  try {
    var firstName = customer.name.split(' ')[0];
    var threads   = GmailApp.search(
      'in:sent subject:"awning proposal" ' + firstName,
      0, 5
    );

    for (var t = 0; t < threads.length; t++) {
      var msgs = threads[t].getMessages();
      for (var m = 0; m < msgs.length; m++) {
        var body = msgs[m].getBody();
        // Look for a hyperlink in the email body
        var linkMatch = body.match(/href="(https:\/\/[^"]+)"/i);
        if (linkMatch) {
          proposalLink = linkMatch[1];
          break;
        }
      }
      if (proposalLink) break;
    }
  } catch (e) {
    Logger.log('Gmail search error: ' + e.message);
  }

  // Allow manual override/confirmation
  var promptMsg = proposalLink
    ? 'Found this proposal link. Confirm or replace it:'
    : 'No proposal email found automatically. Paste the link manually:';

  var response = ui.prompt(
    'Proposal Link for ' + customer.name,
    promptMsg,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var finalLink = response.getResponseText().trim() || proposalLink;
  if (!finalLink) { ui.alert('No link provided.'); return; }

  var message = 'Hi ' + customer.name.split(' ')[0] + '! Here is your awning proposal from Walker Awning: ' + finalLink + '\n\nPlease review and feel free to reach out with any questions!';

  var confirm = ui.alert(
    'Confirm Send',
    'Sending to ' + customer.phone + ':\n\n' + message,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  var sent = dp_sendSms_(customer.phone, message);
  if (sent) ui.alert('✅ Proposal link sent to ' + customer.name + '!');
}

// ── Use Case 4: Custom / Coordination Message ───────────────
function dp_sendCustomMessage_() {
  var customer = dp_getRowData_();
  if (!customer) return;

  var ui = SpreadsheetApp.getUi();

  // Offer quick templates
  var templateChoice = ui.alert(
    'Message Type for ' + customer.name,
    'Choose a starting template:\n\n' +
    '[OK] = Walkthrough / Site Visit\n' +
    '[Cancel] = I\'ll write my own',
    ui.ButtonSet.OK_CANCEL
  );

  var defaultText = '';
  if (templateChoice === ui.Button.OK) {
    defaultText = 'Hi ' + customer.name.split(' ')[0] + '! We\'d like to schedule a walkthrough for your awning project. What days/times work best for you? – Walker Awning';
  }

  var response = ui.prompt(
    'Custom Message to ' + customer.name + ' (' + customer.phone + ')',
    'Edit or write your message below:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var message = response.getResponseText().trim();
  if (!message) { ui.alert('No message entered.'); return; }

  var confirm = ui.alert(
    'Confirm Send',
    'Sending to ' + customer.phone + ':\n\n' + message,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  var sent = dp_sendSms_(customer.phone, message);
  if (sent) ui.alert('✅ Message sent to ' + customer.name + '!');
}