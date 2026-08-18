/**
 * Gino.gs
 * version# [08/17-1:05PM EST]
 * by Claude Opus 4.1
 *
 * V2 BACKLOG (per Gino):
 * - "Entered/Exited sheet" tracking (installable onOpen + Drive Activity API view sweep)
 * - Drive photo/file-added tracking (backshelved)
 */
const AL_CONFIG = {
  LOG_ID_PROP: 'AL_LOG_ID',
  LAST_SWEEP_PROP: 'AL_LAST_MAIL_SWEEP',
  BACKFILL_CURSOR_PROP: 'AL_BACKFILL_CURSOR',
  BACKFILL_MONTHS: 1, // how far back the one-time backfills reach
  REPORT_URL: 'https://lookerstudio.google.com/reporting/c3fa4070-6ef5-410a-a670-2b35bad39253',
  LOG_FILE_NAME: "Gino's Diary",
  LOG_SHEET_NAME: 'Log',
  HEADERS: ['Stamp', 'Day', 'Date', 'Time', 'Sheet', 'User', 'Name', 'Display Name', 'Activity', 'Purpose', 'Notes'],
  // Standard subjects: matched (lowercase, partial) against the real subject.
  // First match wins; no match falls back to the actual subject line.
  SUBJECTS: [
    // Wrappers first — these contain other subjects inside them.
    { match: '(price update)',                       label: 'Price update' },
    { match: 'following up:',                        label: 'Follow-up' },
    // Base subjects
    { match: 'your awning quote from walker awning', label: 'Awning quote' },
    { match: 'proposal review',                      label: 'Proposal Review' },
    { match: 'awning proposal',                      label: 'Awning Proposal' },
    { match: 're: your walker awning project',       label: 'Handoff' },
    { match: 'quick quote for your awning',          label: 'Rough quote' },
    { match: 'info request for awning',              label: 'Info request' },
    { match: 'coi request',                          label: 'COI request' },
    { match: 'po: samples for',                      label: 'Samples PO' },
    { match: 'walker awning (50% deposit)',          label: 'Deposit invoice' },
    { match: 'quote solicitation for',               label: 'Quote solicitation' },
    { match: 'scheduling items',                     label: 'Scheduling' },
    { match: 'awning follow-up',                     label: 'Follow-up' },
    { match: 'employee weekly schedule',             label: 'Weekly schedule' }
  ],
  // Vendor/internal emails: matched by SUBJECT PREFIX, not recipient.
  // 'strip' is removed from the front of the subject to recover the display name.
  VENDOR_SUBJECTS: [
    { match: 'po: samples for ',        label: 'Samples PO',        strip: 'po: samples for ' },
    { match: 'quote solicitation for ', label: 'Quote solicitation', strip: 'quote solicitation for ' },
    { match: 'coi request: ',           label: 'COI request',       strip: 'coi request: ' },
    { match: 'proposal review: ',       label: 'Proposal Review',   strip: 'proposal review: ' }
  ],
  SHEETS: ['Leads', 'F/U', 'Awarded', 'Heaven', 'Purgatory', 'Re-cover'],
  TABLE_SHEETS: ['Leads', 'F/U', 'Awarded', 'Heaven', 'Purgatory'],
  STAGE_COL: 4,   // D
  NAME_COL: 5,    // E
  DISPLAY_COL: 6, // F
  EMAIL_COL: 9,   // I
  TZ: 'America/New_York'
};

/* ---------- SETUP (menu: "Setup Activity Log") ---------- */
function al_setupLog() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(AL_CONFIG.LOG_ID_PROP);
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (err) { id = null; }
  }
  if (!id) {
    ss = SpreadsheetApp.create(AL_CONFIG.LOG_FILE_NAME);
    props.setProperty(AL_CONFIG.LOG_ID_PROP, ss.getId());
  }
  let sh = ss.getSheetByName(AL_CONFIG.LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(AL_CONFIG.LOG_SHEET_NAME);
  }
  if (sh.getRange(1, 1).getValue() !== AL_CONFIG.HEADERS[0]) {
    sh.getRange(1, 1, 1, AL_CONFIG.HEADERS.length).setValues([AL_CONFIG.HEADERS]);
    sh.setFrozenRows(1);
  }
  SpreadsheetApp.getUi().alert('Activity Log ready:\n' + ss.getUrl());
}

/* ---------- TRIGGER INSTALLER (menu: "Install Log Triggers") ---------- */
function al_installLogTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'al_sweepSentMail_' || fn === 'al_sortLog_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('al_sweepSentMail_').timeBased().everyDays(1).atHour(17).create();
  ScriptApp.newTrigger('al_sortLog_').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('Log triggers installed: mail sweep 5PM, sort 7AM (weekdays only; weekend runs self-skip).');
}

/* ---------- OPEN LOG (menu: "Open Log") ---------- */
function al_openLog() {
  const id = PropertiesService.getScriptProperties().getProperty(AL_CONFIG.LOG_ID_PROP);
  if (!id) {
    SpreadsheetApp.getUi().alert('Run "Setup Activity Log" first.');
    return;
  }
  const url = 'https://docs.google.com/spreadsheets/d/' + id;
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:10px;"><a href="' + url +
    '" target="_blank" onclick="google.script.host.close()">Open Activity Log ↗</a></div>'
  ).setWidth(260).setHeight(70);
  SpreadsheetApp.getUi().showModalDialog(html, 'Activity Log');
}

/* ---------- CORE WRITER ---------- */
function al_logActivity_(sheetName, user, name, display, activity, purpose, notes, when) {
  const id = PropertiesService.getScriptProperties().getProperty(AL_CONFIG.LOG_ID_PROP);
  if (!id) return; // log not set up yet; never block main handlers
  const sh = SpreadsheetApp.openById(id).getSheetByName(AL_CONFIG.LOG_SHEET_NAME);
  const now = (when instanceof Date) ? when : new Date();
  sh.appendRow([
    now,
    Utilities.formatDate(now, AL_CONFIG.TZ, 'EEE'),
    Utilities.formatDate(now, AL_CONFIG.TZ, 'MM/dd'),
    Utilities.formatDate(now, AL_CONFIG.TZ, 'h:mm a'),
    sheetName, user, name, display, activity, purpose || '', notes || ''
  ]);
}

// Maps a real subject line to its standard label; falls back to the actual subject.
function al_standardSubject_(subject) {
  const s = String(subject || '').toLowerCase();
  for (let i = 0; i < AL_CONFIG.SUBJECTS.length; i++) {
    if (s.indexOf(AL_CONFIG.SUBJECTS[i].match) !== -1) return AL_CONFIG.SUBJECTS[i].label;
  }
  return String(subject || '(no subject)');
}

/* ---------- EDIT LOGGER (called from masterOnEditHandler_) ---------- */
function al_handleEditLog_(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return; // ignore multi-cell
    const sh = e.range.getSheet();
    const sheetName = sh.getName();
    if (AL_CONFIG.SHEETS.indexOf(sheetName) === -1) return;

    const user = al_getUser_(e);

    // Re-cover: calculation sheet, no table format
    if (sheetName === 'Re-cover') {
      const who = sh.getRange('K2').getDisplayValue();
      al_logActivity_(sheetName, user, who, '', 'Calc', who, '');
      return;
    }

    const row = e.range.getRow();
    if (row === 1) return; // ignore header row
    const col = e.range.getColumn();

    const name = sh.getRange(row, AL_CONFIG.NAME_COL).getDisplayValue();
    const display = sh.getRange(row, AL_CONFIG.DISPLAY_COL).getDisplayValue();
    const newVal = e.range.getDisplayValue();

    let activity, purpose;
    if (col === AL_CONFIG.STAGE_COL) {
      activity = 'Stage';
      purpose = newVal; // just the status it was changed to
    } else {
      const header = sh.getRange(1, col).getDisplayValue() || al_colLetter_(col);
      activity = 'Edit';
      purpose = header + ' to ' + newVal;
    }
    al_logActivity_(sheetName, user, name, display, activity, purpose, '');
  } catch (err) {
    // Logging must never break stage automation or draft creation.
  }
}

function al_getUser_(e) {
  let email = '';
  try { email = (e && e.user && e.user.getEmail()) || ''; } catch (err) {}
  if (!email) {
    try { email = Session.getActiveUser().getEmail() || ''; } catch (err) {}
  }
  return email || 'not-Gino';
}

function al_colLetter_(col) {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

/* ---------- SENT MAIL SWEEP (5PM weekdays) ---------- */
function al_sweepSentMail_() {
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, AL_CONFIG.TZ, 'u')); // 1=Mon..7=Sun
  if (dow > 5) return; // weekends self-skip

  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty(AL_CONFIG.LAST_SWEEP_PROP)) || (Date.now() - 24 * 60 * 60 * 1000);
  const me = Session.getEffectiveUser().getEmail();

  const emailMap = al_buildEmailMap_();
  const threads = GmailApp.search('in:sent after:' + Math.floor(last / 1000), 0, 50);

  threads.forEach(function (th) {
    th.getMessages().forEach(function (msg) {
      if (msg.getDate().getTime() <= last) return;
      const subjRaw = msg.getSubject() || '(no subject)';
      const subj = subjRaw.replace(/"/g, '""');
      const link = '=HYPERLINK("https://mail.google.com/mail/u/0/#all/' + msg.getId() + '","' + subj + '")';

      // Vendor/internal email? Match by subject; log once and move on.
      const vend = al_matchVendorSubject_(subjRaw);
      if (vend) {
        al_logActivity_(vend.sheet, me, vend.name, vend.display, 'Email', vend.label, link);
        return;
      }

      const recipients = (msg.getTo() + ',' + msg.getCc()).toLowerCase();
      Object.keys(emailMap).forEach(function (addr) {
        if (addr && recipients.indexOf(addr) !== -1) {
          const rec = emailMap[addr];
          al_logActivity_(rec.sheet, me, rec.name, rec.display, 'Email', al_standardSubject_(subjRaw), link);
        }
      });
    });
  });

  props.setProperty(AL_CONFIG.LAST_SWEEP_PROP, String(Date.now()));
}

// Builds { emailAddress: {sheet, name, display} } across the 5 table sheets.
// Supports comma-separated addresses in column I.
function al_buildEmailMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = {};
  AL_CONFIG.TABLE_SHEETS.forEach(function (sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const data = sh.getRange(2, AL_CONFIG.NAME_COL, lastRow - 1, AL_CONFIG.EMAIL_COL - AL_CONFIG.NAME_COL + 1).getDisplayValues();
    data.forEach(function (r) {
      const name = r[0];                                  // E
      const display = r[AL_CONFIG.DISPLAY_COL - AL_CONFIG.NAME_COL]; // F
      const emails = r[AL_CONFIG.EMAIL_COL - AL_CONFIG.NAME_COL];    // I
      if (!emails) return;
      emails.split(',').forEach(function (addr) {
        addr = addr.trim().toLowerCase();
        if (addr) map[addr] = { sheet: sheetName, name: name, display: display };
      });
    });
  });
  return map;
}

/* ---------- SORT LOG (7AM weekdays) ---------- */
function al_sortLog_() {
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, AL_CONFIG.TZ, 'u'));
  if (dow > 5) return;

  const id = PropertiesService.getScriptProperties().getProperty(AL_CONFIG.LOG_ID_PROP);
  if (!id) return;
  const sh = SpreadsheetApp.openById(id).getSheetByName(AL_CONFIG.LOG_SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return;

  const range = sh.getRange(2, 1, lastRow - 1, AL_CONFIG.HEADERS.length);
  const rows = range.getValues();
  const formulas = sh.getRange(2, 11, lastRow - 1, 1).getFormulas(); // preserve Notes hyperlinks (col K)

  const parsed = rows.map(function (r, i) {
    const stamp = (r[0] instanceof Date) ? r[0] : new Date(r[0]); // Stamp column
    return { r: r, f: formulas[i][0], t: isNaN(stamp) ? 0 : stamp.getTime() };
  });
  parsed.sort(function (a, b) { return a.t - b.t; });

  range.setValues(parsed.map(function (p) { return p.r; }));
  parsed.forEach(function (p, i) {
    if (p.f) sh.getRange(i + 2, 11).setFormula(p.f);
  });
}
/* ---------- ONE-TIME BACKFILLS (menu items; safe to run repeatedly) ---------- */

// Sweeps Sent mail back BACKFILL_MONTHS, in batches of 100 threads.
// Run the menu item repeatedly until it says "COMPLETE".
function al_backfillSentMail() {
  const props = PropertiesService.getScriptProperties();
  const BATCH = 100;
  const start = Number(props.getProperty(AL_CONFIG.BACKFILL_CURSOR_PROP)) || 0;

  const from = new Date();
  from.setMonth(from.getMonth() - AL_CONFIG.BACKFILL_MONTHS);
  const afterStr = Utilities.formatDate(from, AL_CONFIG.TZ, 'yyyy/MM/dd');
  const beforeStr = Utilities.formatDate(new Date(), AL_CONFIG.TZ, 'yyyy/MM/dd'); // excludes today (5PM sweep owns today)

  const me = Session.getEffectiveUser().getEmail();
  const emailMap = al_buildEmailMap_();
  const threads = GmailApp.search('in:sent after:' + afterStr + ' before:' + beforeStr, start, BATCH);

  let logged = 0;
  threads.forEach(function (th) {
    th.getMessages().forEach(function (msg) {
      const subjRaw = msg.getSubject() || '(no subject)';
      const subj = subjRaw.replace(/"/g, '""');
      const link = '=HYPERLINK("https://mail.google.com/mail/u/0/#all/' + msg.getId() + '","' + subj + '")';

      const vend = al_matchVendorSubject_(subjRaw);
      if (vend) {
        al_logActivity_(vend.sheet, me, vend.name, vend.display, 'Email', vend.label, link, msg.getDate());
        logged++;
        return;
      }

      const recipients = (msg.getTo() + ',' + msg.getCc()).toLowerCase();
      Object.keys(emailMap).forEach(function (addr) {
        if (addr && recipients.indexOf(addr) !== -1) {
          const rec = emailMap[addr];
          al_logActivity_(rec.sheet, me, rec.name, rec.display, 'Email', al_standardSubject_(subjRaw), link, msg.getDate());
          logged++;
        }
      });
    });
  });

  if (threads.length < BATCH) {
    props.deleteProperty(AL_CONFIG.BACKFILL_CURSOR_PROP);
    SpreadsheetApp.getUi().alert('Email backfill COMPLETE.\nThis run: ' + threads.length + ' threads scanned, ' + logged + ' matches logged.\nRun 7AM sort (or wait for it) to file rows chronologically.');
  } else {
    props.setProperty(AL_CONFIG.BACKFILL_CURSOR_PROP, String(start + BATCH));
    SpreadsheetApp.getUi().alert('Email backfill: batch done (' + logged + ' matches).\nMore remain - click this menu item again to continue.');
  }
}

// Logs past "Gino - " events from the Appointments with Customers calendar.
function al_backfillCalendar() {
  const from = new Date();
  from.setMonth(from.getMonth() - AL_CONFIG.BACKFILL_MONTHS);
  const to = new Date();

  const calendars = CalendarApp.getAllCalendars();
  const cal = calendars.find(function (c) { return c.getName() === 'Appointments with Customers'; });
  if (!cal) {
    SpreadsheetApp.getUi().alert('Calendar "Appointments with Customers" not found.');
    return;
  }

  const events = cal.getEvents(from, to);
  let logged = 0;
  events.forEach(function (ev) {
    const title = String(ev.getTitle() || '');
    if (title.indexOf('Gino - ') !== 0) return;
    const custName = title.substring(7).trim();
    const found = al_findCustomerSheet_(custName);
    al_logActivity_(found.sheet, 'gino@walkerawning.com', custName, found.display, 'Appointment', 'Site visit', ev.getLocation() || '', ev.getStartTime());
    logged++;
  });

  SpreadsheetApp.getUi().alert('Calendar backfill COMPLETE.\n' + logged + ' "Gino - " events logged.');
}

// Finds which table sheet a customer name (col E) lives on. Returns {sheet, display}.
function al_findCustomerSheet_(custName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = String(custName || '').trim().toLowerCase();
  if (!target) return { sheet: '', display: '' };
  for (let i = 0; i < AL_CONFIG.TABLE_SHEETS.length; i++) {
    const sh = ss.getSheetByName(AL_CONFIG.TABLE_SHEETS[i]);
    if (!sh) continue;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;
    const data = sh.getRange(2, AL_CONFIG.NAME_COL, lastRow - 1, AL_CONFIG.DISPLAY_COL - AL_CONFIG.NAME_COL + 1).getDisplayValues();
    for (let j = 0; j < data.length; j++) {
      if (String(data[j][0]).trim().toLowerCase() === target) {
        return { sheet: AL_CONFIG.TABLE_SHEETS[i], display: data[j][AL_CONFIG.DISPLAY_COL - AL_CONFIG.NAME_COL] };
      }
    }
  }
  return { sheet: '', display: '' };
}
/* ---------- OPEN VISUALS (menu: "See Visuals") ---------- */
function al_openVisuals() {
  const url = AL_CONFIG.REPORT_URL;
  if (!url || url.indexOf('http') !== 0) {
    SpreadsheetApp.getUi().alert('Paste your Looker Studio report URL into AL_CONFIG.REPORT_URL first.');
    return;
  }
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:10px;"><a href="' + url +
    '" target="_blank" onclick="google.script.host.close()">Open Visuals ↗</a></div>'
  ).setWidth(260).setHeight(70);
  SpreadsheetApp.getUi().showModalDialog(html, 'Visuals');
}
// Vendor/internal emails: match by subject prefix, recover display name, find its sheet.
// Returns { label, name, display, sheet } or null.
function al_matchVendorSubject_(subject) {
  const s = String(subject || '').toLowerCase();
  for (let i = 0; i < AL_CONFIG.VENDOR_SUBJECTS.length; i++) {
    const v = AL_CONFIG.VENDOR_SUBJECTS[i];
    // Tolerate reply/forward prefixes ("Re: ", "Fwd: ") before the pattern
    const at = s.indexOf(v.match);
    if (at === 0 || (at > 0 && /^((re|fwd|fw)\s*:\s*)+$/.test(s.substring(0, at)))) {
      // Recover the display name from the original-case subject
      let rest = String(subject).substring(at + v.strip.length).trim();
      // "Proposal Review: [F] - [R]" — drop the trailing job type
      const dash = rest.indexOf(' - ');
      if (dash > 0) rest = rest.substring(0, dash).trim();
      const found = al_findByDisplayName_(rest);
      return { label: v.label, name: found.name, display: rest, sheet: found.sheet };
    }
  }
  return null;
}

// Finds which table sheet a DISPLAY NAME (col F) lives on. Returns {sheet, name}.
function al_findByDisplayName_(display) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = String(display || '').trim().toLowerCase();
  if (!target) return { sheet: '', name: '' };
  for (let i = 0; i < AL_CONFIG.TABLE_SHEETS.length; i++) {
    const sh = ss.getSheetByName(AL_CONFIG.TABLE_SHEETS[i]);
    if (!sh) continue;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;
    const data = sh.getRange(2, AL_CONFIG.NAME_COL, lastRow - 1, AL_CONFIG.DISPLAY_COL - AL_CONFIG.NAME_COL + 1).getDisplayValues();
    for (let j = 0; j < data.length; j++) {
      if (String(data[j][AL_CONFIG.DISPLAY_COL - AL_CONFIG.NAME_COL]).trim().toLowerCase() === target) {
        return { sheet: AL_CONFIG.TABLE_SHEETS[i], name: data[j][0] };
      }
    }
  }
  return { sheet: '', name: '' };
}