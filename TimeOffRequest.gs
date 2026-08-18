/**
 * TimeOffRequest.gs — v4 (clickable form image, one-page layout)
 * TIME-OFF REQUEST — Walker Canvas Awnings, Inc.
 * ------------------------------------------------
 * Menu item lives in Menus.gs (🔧 Tools → 🏖️ New Time-Off Request).
 *
 * Creates a Gmail DRAFT to Liz where the filled + signed form appears
 * as a single clickable image in the email body. Clicking it opens the
 * PDF (stored in the "Time-Off Requests" Drive folder) for printing.
 *
 * Set ATTACH_PDF to true below if you also want a PDF copy attached.
 *
 * NOTE: No onOpen() here on purpose — Menus.gs owns the menus.
 */

// ============================== CONFIG ==============================

var CONFIG = {
  EMPLOYEE_NAME: 'Gino Carneiro',
  RECIPIENT: 'Liz@WalkerAwning.com',
  SUBJECT: 'Time-off request form (Gino)',
  SIGNATURE_FILE_ID: '1VO9vvGw2C9lncv8vYWO7F2raBm_EFW3x',  // signature PNG in Drive
  ATTACH_PDF: false,  // true = also attach a PDF copy of the form
  CALENDAR_NAME: 'Vacations & Absences',  // shared calendar to mark absences on
  EVENT_TITLE: 'Gino Out (Pending approval)'  // all-day event title
};

// ============================== DIALOG ==============================

function showTimeOffDialog() {
  var html = HtmlService.createHtmlOutputFromFile('TimeOffDialog')
    .setWidth(420)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, 'New Time-Off Request');
}

// ============================== HOLIDAYS ==============================

/** Returns company-observed holidays for a given year as 'yyyy-MM-dd' strings. */
function companyHolidays_(year) {
  var list = [];

  // ---- Fixed dates ----
  list.push(year + '-01-01');                    // New Year's Day
  list.push(year + '-07-04');                    // Independence Day
  list.push(year + '-12-25');                    // Christmas Day

  // ---- Floating dates ----
  list.push(nthWeekdayOfMonth_(year, 5, 1, -1)); // Memorial Day (last Mon of May)
  list.push(nthWeekdayOfMonth_(year, 9, 1, 1));  // Labor Day (1st Mon of Sep)
  list.push(nthWeekdayOfMonth_(year, 11, 4, 4)); // Thanksgiving (4th Thu of Nov)

  // Add more if the company observes them, e.g.:
  // list.push(nthWeekdayOfMonth_(year, 1, 1, 3));  // MLK Day (3rd Mon of Jan)
  // list.push(year + '-06-19');                    // Juneteenth

  return list;
}

function nthWeekdayOfMonth_(year, month, weekday, n) {
  var d;
  if (n === -1) {
    d = new Date(year, month, 0);
    while (d.getDay() !== weekday) d.setDate(d.getDate() - 1);
  } else {
    d = new Date(year, month - 1, 1);
    while (d.getDay() !== weekday) d.setDate(d.getDate() + 1);
    d.setDate(d.getDate() + (n - 1) * 7);
  }
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Counts Mon–Fri days between two 'yyyy-MM-dd' dates inclusive, minus holidays. */
function businessDaysBetween_(startStr, endStr) {
  var start = parseYmd_(startStr);
  var end = parseYmd_(endStr);
  if (end < start) return 0;

  var holidays = {};
  for (var y = start.getFullYear(); y <= end.getFullYear(); y++) {
    companyHolidays_(y).forEach(function (h) { holidays[h] = true; });
  }

  var count = 0;
  var d = new Date(start);
  while (d <= end) {
    var dow = d.getDay();
    var ymd = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (dow !== 0 && dow !== 6 && !holidays[ymd]) count++;
    d.setDate(d.getDate() + 1);
  }
  return count;
}

function parseYmd_(s) {
  var p = s.split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function formatPretty_(s) {
  return Utilities.formatDate(parseYmd_(s), Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

/** e.g. "Monday 8/24/2026" */
function formatDowShort_(s) {
  var d = parseYmd_(s);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEEE') + ' ' +
         Utilities.formatDate(d, Session.getScriptTimeZone(), 'M/d/yyyy');
}

/** e.g. "Monday, 08/24/2026" (used on the form) */
function formatDowLong_(s) {
  var d = parseYmd_(s);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEEE, MM/dd/yyyy');
}

// ============================== MAIN ==============================

/**
 * Called by the dialog. data = {startDate, endDate, type, otherText, comments}
 * Returns {message, url} shown in the dialog.
 */
function submitTimeOffRequest(data) {
  if (!data.startDate || !data.endDate) throw new Error('Start and end dates are required.');
  if (parseYmd_(data.endDate) < parseYmd_(data.startDate)) {
    throw new Error('End date is before start date.');
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
  var numDays = businessDaysBetween_(data.startDate, data.endDate);

  var v = {
    name: CONFIG.EMPLOYEE_NAME,
    requestDate: today,
    numDays: numDays,
    startDate: formatDowLong_(data.startDate),
    endDate: formatDowLong_(data.endDate),
    type: data.type || '',
    otherText: data.otherText || '',
    comments: data.comments || '',
    signDate: today
  };

  // Signature blob for inline embedding (falls back to italic name)
  var sigBlob = null;
  try {
    if (CONFIG.SIGNATURE_FILE_ID) {
      sigBlob = DriveApp.getFileById(CONFIG.SIGNATURE_FILE_ID).getBlob().setName('sig');
    }
  } catch (e) { sigBlob = null; }

  // --- Save PDF to Drive so the email image can link to it ---
  var pdfBlob = buildFormPdf_(v, sigBlob);
  var folder = getOrCreateFolder_('Time-Off Requests');
  var pdfFile = folder.createFile(pdfBlob);
  try { pdfFile.addViewer(CONFIG.RECIPIENT); } catch (shareErr) {}

  // --- Email body: greeting + one clickable image of the form ---
  var dateLine = formatDowShort_(data.startDate) + ' - ' + formatDowShort_(data.endDate);

  var greeting =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;">' +
    'Hello Liz,<br>' +
    '&nbsp;&nbsp;&nbsp;&nbsp;May I please have the following days off?<br>' +
    '<b>' + dateLine + '</b><br><br>' +
    '</div>';

  var options = {};
  var imgBlob = getFormImageBlob_(pdfFile.getId());

  if (imgBlob) {
    options.htmlBody = greeting +
      '<a href="' + pdfFile.getUrl() + '">' +
      '<img src="cid:formimg" style="max-width:700px;width:100%;border:1px solid #ddd;" ' +
      'alt="Time Request Form (click to open and print)"></a>';
    options.inlineImages = { formimg: imgBlob };
  } else {
    // Fallback: render the form as HTML if Drive image generation fails
    options.htmlBody = greeting + buildFormHtml_(v, sigBlob ? 'cid:sigimg' : null);
    if (sigBlob) options.inlineImages = { sigimg: sigBlob };
  }
  if (CONFIG.ATTACH_PDF) options.attachments = [pdfBlob];

  var plainText =
    'Hello Liz,\n    May I please have the following days off?\n' +
    dateLine + '\n\nForm PDF: ' + pdfFile.getUrl();

  var subject = CONFIG.SUBJECT + ' - ' + today;
  var draft = GmailApp.createDraft(CONFIG.RECIPIENT, subject, plainText, options);
  var draftUrl = 'https://mail.google.com/mail/u/0/#drafts/' + draft.getMessage().getId();

  // Mark the absence on the shared Vacations & Absences calendar
  var calMsg = '';
  try {
    calMsg = markVacationCalendar_(data.startDate, data.endDate, subject);
  } catch (calErr) {
    calMsg = '⚠️ Calendar error: ' + calErr.message;
  }

  return {
    message: 'Draft created! (' + numDays + ' business day(s) counted.) ' + calMsg,
    url: draftUrl
  };
}

/**
 * Creates all-day "Gino Out" events on the shared calendar, covering
 * business days only — a range spanning a weekend/holiday becomes
 * separate events so non-work days stay unmarked.
 */
function markVacationCalendar_(startStr, endStr, subject) {
  var paperTrail = 'Approval paper trail — search Gmail for this request:\n' +
    'https://mail.google.com/mail/u/0/#search/' + encodeURIComponent('"' + subject + '"');
  var cal = null;
  var cals = CalendarApp.getAllCalendars();
  for (var i = 0; i < cals.length; i++) {
    if (cals[i].getName() === CONFIG.CALENDAR_NAME) { cal = cals[i]; break; }
  }
  if (!cal) return '⚠️ Calendar "' + CONFIG.CALENDAR_NAME + '" not found — no event created.';

  var start = parseYmd_(startStr), end = parseYmd_(endStr);
  var holidays = {};
  for (var y = start.getFullYear(); y <= end.getFullYear(); y++) {
    companyHolidays_(y).forEach(function (h) { holidays[h] = true; });
  }

  // Split the range into contiguous business-day segments
  var segments = [], segStart = null, prev = null;
  var d = new Date(start);
  while (d <= end) {
    var ymd = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var isBiz = d.getDay() !== 0 && d.getDay() !== 6 && !holidays[ymd];
    if (isBiz && !segStart) segStart = new Date(d);
    if (!isBiz && segStart) { segments.push([segStart, new Date(prev)]); segStart = null; }
    prev = new Date(d);
    d.setDate(d.getDate() + 1);
  }
  if (segStart) segments.push([segStart, new Date(prev)]);

  if (!segments.length) return '⚠️ No business days in range — no event created.';

  segments.forEach(function (seg) {
    var endExclusive = new Date(seg[1]);
    endExclusive.setDate(endExclusive.getDate() + 1);  // all-day end is exclusive
    cal.createAllDayEvent(CONFIG.EVENT_TITLE, seg[0], endExclusive, { description: paperTrail });
  });

  return '📅 Marked on ' + CONFIG.CALENDAR_NAME + '.';
}

/** Finds or creates a Drive folder by name. */
function getOrCreateFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/** Asks Drive to render a PNG snapshot of the PDF. Retries while it generates. */
function getFormImageBlob_(fileId) {
  try {
    var token = ScriptApp.getOAuthToken();
    for (var i = 0; i < 6; i++) {
      var meta = JSON.parse(UrlFetchApp.fetch(
        'https://www.googleapis.com/drive/v3/files/' + fileId + '?fields=thumbnailLink',
        { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
      ).getContentText());
      if (meta.thumbnailLink) {
        var url = meta.thumbnailLink.replace(/=s\d+.*$/, '=s1600');
        var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) {
          return resp.getBlob().setName('form.png');
        }
      }
      Utilities.sleep(1500);
    }
  } catch (e) {}
  return null;
}

// ============================== FORM HTML ==============================

/**
 * Builds the form as email-safe HTML (all inline styles, tables).
 * sigSrc: 'cid:sigimg' for email, a data: URI for PDF, or null for italic name.
 */
function buildFormHtml_(v, sigSrc) {
  var sigHtml = sigSrc
    ? '<img src="' + sigSrc + '" style="height:40px;" alt="signature">'
    : '<span style="font-style:italic;font-size:22px;">' + esc_(v.name) + '</span>';

  var BAR = 'background-color:#1f3a5f;color:#ffffff;font-weight:bold;text-align:center;' +
            'padding:6px;font-size:12px;font-family:Arial,Helvetica,sans-serif;';
  var LBL = 'font-weight:bold;font-family:Arial,Helvetica,sans-serif;font-size:12.5px;white-space:nowrap;';
  var BOX = 'border:1px solid #8a97a8;padding:7px 10px;background-color:#ffffff;font-size:13px;';
  var SEC = 'border:1px solid #b9c2cf;border-collapse:collapse;margin-top:12px;';
  var CB  = 'font-size:14px;';

  function infoRow(label, value, first, last) {
    var padT = first ? '12px' : '5px';
    var padB = last ? '12px' : '5px';
    return '<tr><td style="padding:' + padT + ' 18px ' + padB + ' 18px;' + LBL + 'width:240px;">' + label + '</td>' +
           '<td style="padding:' + padT + ' 18px ' + padB + ' 0;"><div style="' + BOX + '">' + value + '</div></td></tr>';
  }

  var col1 = ['Vacation', 'Sick', 'Jury Duty', 'Bereavement Leave'];
  var col2 = ['Leave Without Pay', 'Family Emergency', 'Other'];

  function cbCell(t, padT, padB) {
    if (t === null) return '<td style="padding:' + padT + ' 18px ' + padB + ' 40px;">&nbsp;</td>';
    var checked = (v.type === t);
    var box = checked ? '&#9745;' : '&#9744;';
    var label = t;
    if (t === 'Other' && checked && v.otherText) label += ': <u>' + esc_(v.otherText) + '</u>';
    return '<td width="50%" style="padding:' + padT + ' 18px ' + padB + ' 40px;' + CB +
           (checked ? 'font-weight:bold;' : '') + '">' + box + ' ' + label + '</td>';
  }

  var typeRows = '';
  for (var i = 0; i < col1.length; i++) {
    var padT = i === 0 ? '10px' : '4px';
    var padB = i === col1.length - 1 ? '10px' : '4px';
    typeRows += '<tr>' + cbCell(col1[i], padT, padB) +
                cbCell(i < col2.length ? col2[i] : null, padT, padB) + '</tr>';
  }

  return '' +
  '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;max-width:700px;"><tr>' +
  '<td style="vertical-align:top;">' +
  '<div style="font-size:24px;font-weight:bold;color:#1f3a5f;font-family:Arial,Helvetica,sans-serif;">Walker Canvas Awnings, Inc.</div>' +
  '<div style="color:#555;font-size:11px;padding-top:4px;font-family:Arial,Helvetica,sans-serif;">5190 NW 10 Terrace &bull; Fort Lauderdale, FL 33309</div>' +
  '</td>' +
  '<td style="vertical-align:top;text-align:right;">' +
  '<div style="font-size:19px;font-weight:bold;color:#1f3a5f;font-family:Arial,Helvetica,sans-serif;">TIME REQUEST FORM</div>' +
  '</td></tr></table>' +
  '<div style="border-bottom:3px solid #1f3a5f;margin-top:10px;max-width:700px;"></div>' +

  '<table width="100%" cellpadding="0" cellspacing="0" style="' + SEC + 'max-width:700px;">' +
  '<tr><td colspan="2" style="' + BAR + '">EMPLOYEE INFORMATION</td></tr>' +
  infoRow('Name:', esc_(v.name), true, false) +
  infoRow('Date of Request:', v.requestDate, false, false) +
  infoRow('Number of Days Requesting:', v.numDays, false, false) +
  infoRow('Start Date:', v.startDate, false, false) +
  infoRow('End Date:', v.endDate, false, true) +
  '</table>' +

  '<table width="100%" cellpadding="0" cellspacing="0" style="' + SEC + 'max-width:700px;">' +
  '<tr><td colspan="2" style="' + BAR + '">TYPE OF REQUEST</td></tr>' +
  typeRows +
  '</table>' +

  '<table width="100%" cellpadding="0" cellspacing="0" style="' + SEC + 'max-width:700px;">' +
  '<tr><td colspan="4" style="' + BAR + '">COMMENTS</td></tr>' +
  '<tr><td colspan="4" style="padding:10px 18px 8px 18px;">' +
  '<div style="border:1px solid #8a97a8;min-height:60px;padding:8px;background-color:#ffffff;font-size:13px;">' +
  esc_(v.comments).replace(/\n/g, '<br>') + '&nbsp;</div></td></tr>' +
  '<tr>' +
  '<td style="padding:3px 8px 10px 18px;' + LBL + '">Total Days Used to Date:</td>' +
  '<td style="padding:3px 18px 10px 0;"><div style="' + BOX + 'width:100px;">&nbsp;</div></td>' +
  '<td style="padding:3px 8px 10px 18px;' + LBL + '">Remaining Days:</td>' +
  '<td style="padding:3px 18px 10px 0;"><div style="' + BOX + 'width:100px;">&nbsp;</div></td>' +
  '</tr></table>' +

  '<table width="100%" cellpadding="0" cellspacing="0" style="' + SEC + 'max-width:700px;background-color:#f4f6f9;">' +
  '<tr><td colspan="4" style="padding:10px 18px 4px 18px;font-style:italic;font-size:11.5px;color:#333;">' +
  'I understand that time away from work is subject to management approval and company policies. ' +
  'I also acknowledge that the total days used and total days remaining are correct.</td></tr>' +
  '<tr>' +
  '<td style="padding:8px 8px 12px 18px;' + LBL + 'width:150px;">Employee Signature:</td>' +
  '<td style="padding:8px 18px 12px 0;"><div style="border-bottom:1.5px solid #000;width:250px;text-align:center;">' +
  sigHtml + '</div></td>' +
  '<td style="padding:8px 8px 12px 10px;' + LBL + 'width:44px;">Date:</td>' +
  '<td style="padding:8px 18px 12px 0;"><div style="' + BOX + 'width:110px;">' + v.signDate + '</div></td>' +
  '</tr></table>' +

  '<table width="100%" cellpadding="0" cellspacing="0" style="' + SEC + 'max-width:700px;">' +
  '<tr><td colspan="4" style="' + BAR + '">APPROVAL &mdash; OFFICE USE ONLY</td></tr>' +
  '<tr>' +
  '<td style="padding:12px 8px 4px 18px;' + LBL + 'width:150px;">Approved:</td>' +
  '<td colspan="3" style="padding:12px 18px 4px 0;font-size:14px;">&#9744; Yes &nbsp;&nbsp;&nbsp;&nbsp; &#9744; No</td>' +
  '</tr><tr>' +
  '<td style="padding:8px 8px 14px 18px;' + LBL + '">Manager Approval:</td>' +
  '<td style="padding:8px 18px 14px 0;"><div style="border-bottom:1.5px solid #000;width:250px;">&nbsp;</div></td>' +
  '<td style="padding:8px 8px 14px 10px;' + LBL + 'width:44px;">Date:</td>' +
  '<td style="padding:8px 18px 14px 0;"><div style="' + BOX + 'width:110px;">&nbsp;</div></td>' +
  '</tr></table>';
}

// ============================== PDF ==============================

function buildFormPdf_(v, sigBlob) {
  var sigSrc = null;
  if (sigBlob) {
    sigSrc = 'data:' + sigBlob.getContentType() + ';base64,' +
             Utilities.base64Encode(sigBlob.getBytes());
  }

  var html = '<html><head><meta charset="utf-8"></head>' +
             '<body style="font-family:Arial,Helvetica,sans-serif;color:#1a1a1a;margin:22px;">' +
             buildFormHtml_(v, sigSrc) +
             '</body></html>';

  return Utilities.newBlob(html, MimeType.HTML, 'form.html')
    .getAs(MimeType.PDF)
    .setName('Time Request Form - ' + v.name + '.pdf');
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}