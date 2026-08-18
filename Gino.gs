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
  LOG_FILE_NAME: "Gino's Diary",
  LOG_SHEET_NAME: 'Log',
  HEADERS: ['Day', 'Date', 'Time', 'Sheet', 'User', 'Name', 'Display Name', 'Activity', 'Notes'],
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
function al_logActivity_(sheetName, user, name, display, activity, notes) {
  const id = PropertiesService.getScriptProperties().getProperty(AL_CONFIG.LOG_ID_PROP);
  if (!id) return; // log not set up yet; never block main handlers
  const sh = SpreadsheetApp.openById(id).getSheetByName(AL_CONFIG.LOG_SHEET_NAME);
  const now = new Date();
  sh.appendRow([
    Utilities.formatDate(now, AL_CONFIG.TZ, 'EEE'),
    Utilities.formatDate(now, AL_CONFIG.TZ, 'MM/dd'),
    Utilities.formatDate(now, AL_CONFIG.TZ, 'h:mm a'),
    sheetName, user, name, display, activity, notes || ''
  ]);
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
      al_logActivity_(sheetName, user, who, '', 'Calculation adjustment for ' + who, '');
      return;
    }

    const row = e.range.getRow();
    if (row === 1) return; // ignore header row
    const col = e.range.getColumn();

    const name = sh.getRange(row, AL_CONFIG.NAME_COL).getDisplayValue();
    const display = sh.getRange(row, AL_CONFIG.DISPLAY_COL).getDisplayValue();
    const newVal = e.range.getDisplayValue();

    let activity;
    if (col === AL_CONFIG.STAGE_COL) {
      activity = newVal; // objective: just the stage it was changed to
    } else {
      const header = sh.getRange(1, col).getDisplayValue() || al_colLetter_(col);
      activity = header + ' to ' + newVal;
    }
    al_logActivity_(sheetName, user, name, display, activity, '');
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
      const recipients = (msg.getTo() + ',' + msg.getCc()).toLowerCase();
      Object.keys(emailMap).forEach(function (addr) {
        if (addr && recipients.indexOf(addr) !== -1) {
          const rec = emailMap[addr];
          const subj = (msg.getSubject() || '(no subject)').replace(/"/g, '""');
          const link = '=HYPERLINK("https://mail.google.com/mail/u/0/#all/' + msg.getId() + '","' + subj + '")';
          al_logActivity_(rec.sheet, me, rec.name, rec.display, 'Emailed customer', link);
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
  const formulas = sh.getRange(2, 9, lastRow - 1, 1).getFormulas(); // preserve Notes hyperlinks

  const parsed = rows.map(function (r, i) {
    const stamp = new Date(new Date().getFullYear() + '/' + r[1] + ' ' + r[2]); // mm/dd + time
    return { r: r, f: formulas[i][0], t: isNaN(stamp) ? 0 : stamp.getTime() };
  });
  parsed.sort(function (a, b) { return a.t - b.t; });

  range.setValues(parsed.map(function (p) { return p.r; }));
  parsed.forEach(function (p, i) {
    if (p.f) sh.getRange(i + 2, 9).setFormula(p.f);
  });
}