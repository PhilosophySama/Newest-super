/**
 * Mileage Log Automation Script
 * Version 8.1 [09/10-11:00AM EST] by Claude Opus 4.1
 *
 * Features:
 * - Fills gaps between first (A2) and last logged entry
 * - Catches up from last entry to TODAY only
 * - Never adds events beyond today's date
 * - Skips events without addresses or at office location
 * - Makes dates clickable links to Google Calendar
 * - Makes addresses clickable links to Google Maps directions
 * - Adds ordinal numbers (1st, 2nd, 3rd) to trip sequence
 * - Adds timestamp when entries are created
 * - Batch writes for performance (including hyperlinks)
 */

// ========================
// CONFIGURATION
// ========================

const MILEAGE_CONFIG = {
  CALENDAR_NAME: 'Appointments with Customers',
  TARGET_SHEET: '2 Mile log',
  EVENT_PREFIX: 'Gino - ',
  SHOP_ADDRESS: '5190 NW 10th Terrace, Fort Lauderdale, FL 33309',
  SHOP_NAME: 'Walker Awning',

  COLUMNS: {
    DATE: 1,           // A - Date of driv
    MILES: 2,          // B - Miles
    TRIP_TYPE: 3,      // C - Trip
    TO_ADDRESS: 4,     // D - To
    CLIENT_NAME: 5,    // E - Client nam
    PURPOSE: 6,        // F - Purpose of visit
    AMT: 7,            // G - AMT
    NOTES: 8,          // H - Automated
    TIMESTAMP: 9       // I - Time stamp
  },

  DEFAULT_TRIP_TYPE: 'One way',

  AMT_FORMULA: (row) =>
    `=IF(C${row}="One way",0.67*B${row},IF(C${row}="Round trip",(0.67*B${row})*2,"$0.00"))`
};

// ========================
// MAIN
// ========================

function installTriggerMileage_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete ALL old mileage triggers first
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const fn = trigger.getHandlerFunction();
    if (fn.startsWith("ml_") || fn.includes("Mileage") || fn.includes("runMileage")) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Install the new correct one
  ScriptApp.newTrigger('runMileageSync_').timeBased().everyDays(1).atHour(6).create();
  ss.toast('Daily mileage sync trigger installed (6 AM)', 'Success ✓', 5);
}

function runMileageSync_() {
  try {
    const sheet = ml_getMileageSheet();
    ml_validateSheet(sheet);

    const { firstDate, lastDate } = ml_getDateBoundaries(sheet);
    const existingDates = ml_getExistingDates(sheet);

    const today = new Date();
    today.setHours(23, 59, 59, 999); // End of today

    let endDate = lastDate > today ? today : lastDate;
    
    const gapEvents = ml_getCalendarEventsChunked(firstDate, endDate);
    const gapEntries = ml_processEvents(gapEvents, existingDates, firstDate);

    let catchupEntries = [];
    if (lastDate < today) {
      const catchupStart = new Date(lastDate);
      catchupStart.setDate(catchupStart.getDate() + 1);
      catchupStart.setHours(0, 0, 0, 0);
      
      const catchupEvents = ml_getCalendarEventsChunked(catchupStart, today);
      catchupEntries = ml_processEvents(catchupEvents, existingDates, firstDate);
    }

    const allNewEntries = [...gapEntries, ...catchupEntries];

    if (allNewEntries.length > 0) {
      ml_addEntriesToSheet(sheet, allNewEntries);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `✓ Added ${allNewEntries.length} entries (${gapEntries.length} gaps, ${catchupEntries.length} catch-up)`,
        'Success', 5
      );
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('No new entries found', 'Complete', 3);
    }
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${err.message}`, 'Error ✗', 8);
    console.error('[MileageSync] Error:', err);
  }
}

// ========================
// PROCESSING
// ========================

function ml_processEvents(events, existingDates, firstLogDate) {
  const grouped = ml_groupEventsByDate(events);
  const entries = [];
  const timestamp = ml_getTimestamp();

  for (const [dateKey, dayEvents] of Object.entries(grouped)) {
    const dayDate = new Date(dayEvents[0].getStartTime());
    dayDate.setHours(0, 0, 0, 0);

    if (dayDate < firstLogDate) continue;
    if (existingDates.has(ml_formatDate(dayDate))) continue;

    entries.push(...ml_processDayEvents(dayEvents, timestamp));
  }
  return entries;
}

function ml_groupEventsByDate(events) {
  const grouped = {};
  events.forEach(ev => {
    const dateKey = ml_formatDate(ev.getStartTime());
    if (!grouped[dateKey]) grouped[dateKey] = [];
    grouped[dateKey].push(ev);
  });
  return grouped;
}

function ml_processDayEvents(dayEvents, timestamp) {
  const entries = [];
  const validEvents = dayEvents.filter(ev => {
    const addr = ml_getEventAddress(ev);
    return addr && !ml_isOfficeAddress(addr);
  });
  if (validEvents.length === 0) return entries;

  validEvents.sort((a, b) => a.getStartTime() - b.getStartTime());

  let tripNumber = 1;

  for (let i = 0; i < validEvents.length; i++) {
    const ev = validEvents[i];
    const evDate = ev.getStartTime();
    const name = ev.getTitle().replace(MILEAGE_CONFIG.EVENT_PREFIX, '').trim();
    const addr = ml_getEventAddress(ev);

    let fromAddr, miles, note;

    if (i === 0) {
      fromAddr = MILEAGE_CONFIG.SHOP_ADDRESS;
      miles = ml_calculateDistance(fromAddr, addr);
      note = `${ml_getOrdinal(tripNumber)} - from office`;
    } else {
      const prevAddr = ml_getEventAddress(validEvents[i - 1]);
      fromAddr = prevAddr;
      miles = ml_calculateDistance(fromAddr, addr);
      const prevName = validEvents[i - 1].getTitle().replace(MILEAGE_CONFIG.EVENT_PREFIX, '').trim();
      note = `${ml_getOrdinal(tripNumber)} - from ${prevName}`;
    }

    entries.push([evDate, miles, 'One way', addr, name, 'Customer appointment', , ml_flagDistance(note, miles), timestamp, fromAddr]);
    tripNumber++;

    if (i === validEvents.length - 1) {
      const returnMiles = ml_calculateDistance(addr, MILEAGE_CONFIG.SHOP_ADDRESS);
      entries.push([
        evDate, returnMiles, 'One way',
        MILEAGE_CONFIG.SHOP_ADDRESS, 'Return trip', 'Return to office',
        , ml_flagDistance(`${ml_getOrdinal(tripNumber)} - to office`, returnMiles),
        timestamp, addr
      ]);
    }
  }
  return entries;
}

// ========================
// UTILITIES
// ========================

function ml_getOrdinal(n) {
  const s = ["th", "st", "nd", "rd"];
  const v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

function ml_getMileageSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MILEAGE_CONFIG.TARGET_SHEET);
  if (!sheet) throw new Error(`Sheet "${MILEAGE_CONFIG.TARGET_SHEET}" not found`);
  return sheet;
}

function ml_validateSheet(sheet) {
  const headers = sheet.getRange(1, 1, 1, 9).getValues()[0];
  const expected = ['Date of driv', 'Miles', 'Trip', 'To', 'Client nam', 'Purpose of visit', 'AMT', 'Automated', 'Time stamp'];
  expected.forEach((h, i) => {
    if (!headers[i] || !String(headers[i]).toLowerCase().includes(h.toLowerCase())) {
      throw new Error(`Missing or incorrect header in column ${i + 1}: expected "${h}"`);
    }
  });
}

function ml_getDateBoundaries(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) throw new Error('No entries found in sheet');
  
  const firstDate = sheet.getRange(2, MILEAGE_CONFIG.COLUMNS.DATE).getValue();
  const lastDataRow = lastRow - 1;
  const lastDate = sheet.getRange(lastDataRow, MILEAGE_CONFIG.COLUMNS.DATE).getValue();
  
  if (!(firstDate instanceof Date) || !(lastDate instanceof Date)) throw new Error('Invalid dates in sheet');
  
  firstDate.setHours(0, 0, 0, 0);
  lastDate.setHours(0, 0, 0, 0);
  
  return { firstDate, lastDate };
}

function ml_getExistingDates(sheet) {
  const dateSet = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return dateSet;
  
  const numRows = lastRow - 2;
  const dates = sheet.getRange(2, MILEAGE_CONFIG.COLUMNS.DATE, numRows, 1).getValues();
  
  dates.forEach(row => {
    if (row[0] && row[0] instanceof Date) {
      dateSet.add(ml_formatDate(row[0]));
    }
  });
  return dateSet;
}

function ml_getCalendarEventsChunked(start, end) {
  const results = [];
  let cursor = new Date(start);
  while (cursor < end) {
    const chunkEnd = new Date(cursor);
    chunkEnd.setMonth(chunkEnd.getMonth() + 1);
    if (chunkEnd > end) chunkEnd.setTime(end.getTime());
    results.push(...ml_getCalendarEvents(cursor, chunkEnd));
    cursor = chunkEnd;
  }
  return results;
}

function ml_getCalendarEvents(start, end) {
  const cal = CalendarApp.getCalendarsByName(MILEAGE_CONFIG.CALENDAR_NAME)[0];
  if (!cal) throw new Error(`Calendar "${MILEAGE_CONFIG.CALENDAR_NAME}" not found`);
  return cal.getEvents(start, end).filter(ev => ev.getTitle().startsWith(MILEAGE_CONFIG.EVENT_PREFIX));
}

function ml_getEventAddress(event) {
  return event.getLocation() || ml_extractAddress(event.getDescription());
}

function ml_extractAddress(desc) {
  if (!desc) return null;
  const lines = desc.split('\n').map(l => l.trim()).filter(Boolean);
  for (const line of lines) {
    if (line.toLowerCase().startsWith('address:')) return line.replace(/address:/i, '').trim();
    if (line.includes('FL') && /\d/.test(line) && !line.toLowerCase().includes('phone')) return line;
  }
  return null;
}

function ml_isOfficeAddress(address) {
  if (!address) return false;
  const cleanAddr = address.toLowerCase().replace(/[^a-z0-9]/g, '');
  const cleanOffice = MILEAGE_CONFIG.SHOP_ADDRESS.toLowerCase().replace(/[^a-z0-9]/g, '');
  return cleanAddr.includes('5190nw10') || cleanAddr === cleanOffice || cleanAddr.includes('walkerawning');
}

function ml_calculateDistance(from, to) {
  try {
    const dir = Maps.newDirectionFinder().setOrigin(from).setDestination(to).setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();
    if (!dir.routes?.[0]?.legs?.[0]) throw new Error('No route found');
    return parseFloat((dir.routes[0].legs[0].distance.value * 0.000621371).toFixed(1));
  } catch (err) {
    console.error(`Distance calculation failed: ${from} → ${to}`, err);
    return 0;
  }
}

function ml_flagDistance(note, miles) {
  return miles > 140 ? `${note} ⚠️ Distance unusually high (check address!)` : note;
}

function ml_formatDate(date) {
  if (!(date instanceof Date)) return '';
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${m}/${d}/${date.getFullYear()}`;
}

function ml_createCalendarLink(date) {
  if (!(date instanceof Date)) return '';
  return `https://calendar.google.com/calendar/r/day/${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
}

function ml_createMapsLink(fromAddress, toAddress) {
  return `https://www.google.com/maps/dir/${encodeURIComponent(fromAddress)}/${encodeURIComponent(toAddress)}`;
}

function ml_getTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yy hh:mm a');
}

// ========================
// BATCH WRITER (optimized)
// ========================

function ml_addEntriesToSheet(sheet, entries) {
  if (!entries || entries.length === 0) return;

  entries.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  const totalRow = sheet.getLastRow();
  sheet.insertRowsBefore(totalRow, entries.length);

  const values = entries.map((e, i) => [
    e[0], e[1], e[2], e[3], e[4], e[5],
    MILEAGE_CONFIG.AMT_FORMULA(totalRow + i),
    e[7], e[8]
  ]);
  sheet.getRange(totalRow, 1, entries.length, 9).setValues(values);

  const dateRichValues = [];
  const addrRichValues = [];

  for (let i = 0; i < entries.length; i++) {
    const date = entries[i][0];
    const toAddress = entries[i][3];
    const fromAddress = entries[i][9];

    const dateRich = SpreadsheetApp.newRichTextValue()
      .setText(ml_formatDate(date))
      .setLinkUrl(ml_createCalendarLink(date))
      .build();
    dateRichValues.push([dateRich]);

    if (toAddress && fromAddress) {
      const addrRich = SpreadsheetApp.newRichTextValue()
        .setText(toAddress)
        .setLinkUrl(ml_createMapsLink(fromAddress, toAddress))
        .build();
      addrRichValues.push([addrRich]);
    } else {
      addrRichValues.push([SpreadsheetApp.newRichTextValue().setText(toAddress || "").build()]);
    }
  }

  sheet.getRange(totalRow, MILEAGE_CONFIG.COLUMNS.DATE, entries.length, 1).setRichTextValues(dateRichValues);
  sheet.getRange(totalRow, MILEAGE_CONFIG.COLUMNS.TO_ADDRESS, entries.length, 1).setRichTextValues(addrRichValues);

  const tripRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['One way', 'Round trip', 'N/A'], true)
    .build();
  sheet.getRange(totalRow, MILEAGE_CONFIG.COLUMNS.TRIP_TYPE, entries.length, 1).setDataValidation(tripRule);

  const newTotalRow = totalRow + entries.length;
  sheet.getRange(newTotalRow, MILEAGE_CONFIG.COLUMNS.AMT).setFormula(`=SUM(G2:G${newTotalRow - 1})`);
}