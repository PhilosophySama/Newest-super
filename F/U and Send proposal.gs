/**
 * Lead_Followup_Draft_V2.gs
 * version# 09/23-05:05PM EST by Chatgpt5
 *
 * PURPOSE
 * - Create Gmail follow-up drafts when Stage (col D) becomes "F/U" on the Leads sheet.
 * - Draft link is written back into col B as a clickable hyperlink.
 *
 * CO-EXISTENCE
 * - Config object name: DRAFTS_FU
 * - Trigger handler: handleEditDraft_FU
 * - Trigger installer: installTriggerDrafts_FU
 * - Backfill runner: createDraftsForAllRows_FU
 * - Helper prefixes: fu_*
 * - NO onOpen() here. Keep the single onOpen() in Menus.gs.
 */

const DRAFTS_FU = {
  SPREADSHEET_ID: 'REPLACE_ME_WITH_YOUR_SHEET_ID',
  SHEET: 'Leads',
  TARGET_STAGE: 'F/U',
  COLS: {
    LOG_B: 'B',
    STAGE: 'D',
    CUSTOMER_NAME: 'E',
    EMAIL: 'I'
  },
  EMAIL: {
    SUBJECT: 'Awning Follow-up',
    SKIP_IF_DRAFT_EXISTS: true
  }
};

/** Install onEdit trigger (clean re-install). */
function installTriggerDrafts_FU() {
  const handler = 'handleEditDraft_FU';
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger(handler)
    .forSpreadsheet(fu_getSpreadsheet_().getId())
    .onEdit()
    .create();
  SpreadsheetApp.getActive().toast('Follow-up draft trigger installed', 'Draft Creator', 4);
}

/** onEdit handler — only Leads!D and only when becomes "F/U". */
function handleEditDraft_FU(e) {
  try {
    if (!e || !e.source || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== DRAFTS_FU.SHEET) return;
    if (e.range.getRow() === 1) return;
    if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

    const stageCol = fu_colLetterToIndex_(DRAFTS_FU.COLS.STAGE);
    if (e.range.getColumn() !== stageCol) return;

    const newVal = String((e.value != null ? e.value : e.range.getValue()) || '').trim();
    if (newVal !== DRAFTS_FU.TARGET_STAGE) return;

    const row = e.range.getRow();

    // Idempotency
    const logCell = sh.getRange(row, fu_colLetterToIndex_(DRAFTS_FU.COLS.LOG_B));
    const existingLink = fu_firstLinkInRichText_(logCell.getRichTextValue());
    if (DRAFTS_FU.EMAIL.SKIP_IF_DRAFT_EXISTS && fu_isGmailDraftUrl_(existingLink)) {
      SpreadsheetApp.getActive().toast('Skipped (existing draft link in B).', 'Follow-up Draft', 4);
      return;
    }

    const res = fu_createDraftForRow_(sh, row);
    SpreadsheetApp.getActive().toast(res.toast, 'Follow-up Draft', 5);
  } catch (err) {
    SpreadsheetApp.getActive().toast('Draft handler error: ' + err, 'Follow-up Draft', 8);
  }
}

/** Backfill all rows where Stage == "F/U". */
function createDraftsForAllRows_FU() {
  const ss = fu_getSpreadsheet_();
  const sh = ss.getSheetByName(DRAFTS_FU.SHEET);
  if (!sh) throw new Error('Leads sheet not found.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const idx = (L) => fu_colLetterToIndex_(L) - 1;
  const stageIdx = idx(DRAFTS_FU.COLS.STAGE);
  const logBColIndex = fu_colLetterToIndex_(DRAFTS_FU.COLS.LOG_B);

  const dataRange = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn());
  const values = dataRange.getValues();
  const rtv    = dataRange.getRichTextValues();

  const logBRange = sh.getRange(2, logBColIndex, lastRow - 1, 1);
  const logBRTV = logBRange.getRichTextValues();
  const outB = logBRTV.map(r => [r[0]]);

  let created = 0, skipped = 0, failed = 0;

  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 2;
    const val = String(values[i][stageIdx] || '').trim();
    if (val !== DRAFTS_FU.TARGET_STAGE) continue;

    const existingLink = fu_firstLinkInRichText_(logBRTV[i][0]);
    if (DRAFTS_FU.EMAIL.SKIP_IF_DRAFT_EXISTS && fu_isGmailDraftUrl_(existingLink)) { skipped++; continue; }

    try {
      const r = fu_createDraftForRow_(sh, rowNum, values[i], rtv[i],
        (richText)=>{ outB[i][0] = richText; });
      if (r.ok) created++; else failed++;
    } catch (err) {
      failed++;
      outB[i][0] = SpreadsheetApp.newRichTextValue().setText('Error: ' + fu_shortErr_(err)).build();
    }
  }

  logBRange.setRichTextValues(outB);
  SpreadsheetApp.getActive().toast(`Backfill → created:${created} | skipped:${skipped} | failed:${failed}`, 'Follow-up Draft', 7);
}

/* =========================
 * Row → Draft creation
 * ========================= */
function fu_createDraftForRow_(sh, row, rowValsOpt, rowRtvOpt, batchReceiverOpt) {
  const vals = rowValsOpt || sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = (L) => fu_colLetterToIndex_(L) - 1;

  const logCell = sh.getRange(row, fu_colLetterToIndex_(DRAFTS_FU.COLS.LOG_B));

  const customerName = fu_safeString_(vals[idx(DRAFTS_FU.COLS.CUSTOMER_NAME)]);
  const email = fu_safeString_(vals[idx(DRAFTS_FU.COLS.EMAIL)]);

  if (!email) {
    fu_writeB_(sh, row, 'Error: No customer email', batchReceiverOpt);
    return { ok:false, toast:'Missing email' };
  }

  const subject = DRAFTS_FU.EMAIL.SUBJECT;
  const body =
    `Hello ${customerName},\n\n` +
    `My name is Gino, Michael is no longer with the company. I'll be helping you out with your awning.\n` +
    `May I please have some pics and/or rough dimensions for your awning?\n` +
    `Also did you know what color you wanted?\n\n`;

  try {
    const draft = GmailApp.createDraft(email, subject, body);
    const draftMessageId = draft.getMessage().getId();
    const draftUrl = 'https://mail.google.com/mail/u/0/#drafts?compose=' + encodeURIComponent(draftMessageId);

    const base = '✅ Draft created';
    const rich = SpreadsheetApp.newRichTextValue()
      .setText(base)
      .setLinkUrl(0, base.length, draftUrl)
      .setTextStyle(0, base.length, SpreadsheetApp.newTextStyle().setUnderline(true).build())
      .build();

    if (batchReceiverOpt) batchReceiverOpt(rich); else logCell.setRichTextValue(rich);

    return { ok:true, toast:'Draft created & linked in column B.' };
  } catch (err) {
    const msg = fu_shortErr_(err);
    fu_writeB_(sh, row, 'Error: ' + msg, batchReceiverOpt);
    return { ok:false, toast:msg };
  }
}

/* =====================
 * Helpers
 * ===================== */
function fu_safeString_(v){ if (v==null) return ''; return v instanceof Date ? v.toLocaleString() : String(v).trim(); }
function fu_colLetterToIndex_(letter){
  let col = 0; const up = String(letter||'').toUpperCase();
  for (let i=0;i<up.length;i++) col = col * 26 + (up.charCodeAt(i) - 64);
  return col;
}
function fu_getSpreadsheet_() {
  if (DRAFTS_FU.SPREADSHEET_ID && DRAFTS_FU.SPREADSHEET_ID !== 'REPLACE_ME_WITH_YOUR_SHEET_ID') {
    return SpreadsheetApp.openById(DRAFTS_FU.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}
function fu_writeB_(sh, row, textOrRich, batchReceiverOpt) {
  if (batchReceiverOpt) {
    batchReceiverOpt(typeof textOrRich === 'string'
      ? SpreadsheetApp.newRichTextValue().setText(textOrRich).build()
      : textOrRich);
  } else {
    const cell = sh.getRange(row, fu_colLetterToIndex_(DRAFTS_FU.COLS.LOG_B));
    if (typeof textOrRich === 'string') cell.setValue(textOrRich); else cell.setRichTextValue(textOrRich);
  }
}
function fu_firstLinkInRichText_(rtv){
  try{
    if (!rtv) return '';
    const runs = rtv.getRuns();
    if (runs && runs.length) for (let k=0;k<runs.length;k++){
      const u = runs[k].getLinkUrl(); if (u) return String(u);
    }
    if (rtv.getLinkUrl){ const u = rtv.getLinkUrl(); if (u) return String(u); }
    return '';
  } catch(_){ return ''; }
}
function fu_isGmailDraftUrl_(u) {
  return typeof u === 'string' &&
         /^https:\/\/mail\.google\.com\/mail\/u\/\d+\/#drafts\?compose=/.test(u);
}
function fu_shortErr_(err){ const s = (err && err.message) ? err.message : String(err||''); return s.length>200 ? s.slice(0,200)+'…' : s; }

/** end-of-file */