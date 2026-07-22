/**
 * ====== CONFIGURATION - EDIT THESE 3 LINES TO MATCH YOUR FORM ======
 * Use the EXACT question text as it appears as a column header
 * in your Form Responses sheet.
 */
const EMAIL_COLUMN_HEADER = 'Email:';
const LAST_NAME_HEADER   = 'Last Name:';
const FIRST_NAME_HEADER  = 'First Name:';
const MIDDLE_NAME_HEADER = 'Middle Name:';
const EVENT_NAME = 'Job Fair 2026';

const ID_HEADER = 'Unique ID';
const STATUS_HEADER = 'Status';
const CHECKIN_TIME_HEADER = 'Check-in Time';
const EMAIL_STATUS_HEADER = 'Email Status';
const INTERVIEW_RESULT_HEADER = 'Interview Result';

function ensureHeaders() {
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const toAdd = [ID_HEADER, STATUS_HEADER, CHECKIN_TIME_HEADER, EMAIL_STATUS_HEADER, INTERVIEW_RESULT_HEADER];
  let lastCol = sheet.getLastColumn();
  toAdd.forEach(h => {
    if (headers.indexOf(h) === -1) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(h);
    }
  });
}

function getResponseSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function getHeaderRow(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function colIndex(headers, name) {
  const idx = headers.indexOf(name);
  if (idx === -1) throw new Error('Column not found: ' + name);
  return idx + 1;
}

// Fires automatically on each new Form submission
function onFormSubmitHandler(e) {
  ensureHeaders();
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const row = e.range.getRow();

  const emailCol = colIndex(headers, EMAIL_COLUMN_HEADER);
  const lastCol  = colIndex(headers, LAST_NAME_HEADER);
  const firstCol = colIndex(headers, FIRST_NAME_HEADER);
  const midCol   = colIndex(headers, MIDDLE_NAME_HEADER);
  const idCol    = colIndex(headers, ID_HEADER);
  const statusCol = colIndex(headers, STATUS_HEADER);
  const emailStatusCol = colIndex(headers, EMAIL_STATUS_HEADER);

  const email = sheet.getRange(row, emailCol).getValue();
  const firstName = sheet.getRange(row, firstCol).getValue();
  const middleName = sheet.getRange(row, midCol).getValue();
  const lastName = sheet.getRange(row, lastCol).getValue();

  // Builds "First Middle Last", skipping middle name if blank
  const fullName = [firstName, middleName, lastName].filter(String).join(' ');

  const uniqueId = Utilities.getUuid();

  sheet.getRange(row, idCol).setValue(uniqueId);
  sheet.getRange(row, statusCol).setValue('Not Checked In');

  try {
    sendTicketEmail(email, fullName, uniqueId);
    sheet.getRange(row, emailStatusCol).setValue('Sent');
  } catch (err) {
    sheet.getRange(row, emailStatusCol).setValue('Pending - ' + err.message);
  }
}

function sendTicketEmail(email, name, uniqueId) {
  const qrBlob = generateQRCodeBlob(uniqueId);

  const htmlBody = `
    <p>Hi ${name},</p>
    <p>You're pre-registered for <b>${EVENT_NAME}</b>. Please show this QR code at check-in (a screenshot works fine):</p>
    <p><img src="cid:qr"/></p>
    <p>Reference ID: <b>${uniqueId}</b></p>
    <p>See you there!</p>
  `;

  GmailApp.sendEmail(email, `Your ${EVENT_NAME} Entry Pass`, 'Please view this email in HTML to see your QR code.', {
    htmlBody: htmlBody,
    inlineImages: { qr: qrBlob }
  });
}

// Safety net: resends any emails that failed (e.g. hit the daily quota)
function retryPendingEmails() {
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const emailCol = colIndex(headers, EMAIL_COLUMN_HEADER);
  const lastCol  = colIndex(headers, LAST_NAME_HEADER);
  const firstCol = colIndex(headers, FIRST_NAME_HEADER);
  const midCol   = colIndex(headers, MIDDLE_NAME_HEADER);
  const idCol    = colIndex(headers, ID_HEADER);
  const emailStatusCol = colIndex(headers, EMAIL_STATUS_HEADER);

  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][emailStatusCol - 1]).indexOf('Pending') === 0) {
      const fullName = [data[r][firstCol - 1], data[r][midCol - 1], data[r][lastCol - 1]].filter(String).join(' ');
      try {
        sendTicketEmail(data[r][emailCol - 1], fullName, data[r][idCol - 1]);
        sheet.getRange(r + 1, emailStatusCol).setValue('Sent');
      } catch (err) {
        sheet.getRange(r + 1, emailStatusCol).setValue('Pending - ' + err.message);
      }
    }
  }
}

// ====== VERIFIER WEB APP ======
function doGet(e) {
  let result;
  try {
    if (e.parameter.action === 'checkin' && e.parameter.id) {
      result = checkInAttendee(e.parameter.id);
    } else if (e.parameter.action === 'hots_info' && e.parameter.id) {
      result = getHotsInfo(e.parameter.id);
    } else if (e.parameter.action === 'hots' && e.parameter.id && e.parameter.company && e.parameter.position) {
      result = registerHOTS(e.parameter.id, e.parameter.company, e.parameter.position);
    } else if (e.parameter.action === 'stats') {
      result = getStats();
    } else if (e.parameter.action === 'list') {
      result = getCheckedInList();
    } else if (e.parameter.action === 'vacancies') {
      result = getVacancyList();
    } else if (e.parameter.action === 'all') {
      result = getAllRegistrants();
    } else {
      result = { status: 'ready', message: EVENT_NAME + ' API is running' };
    }
  } catch (err) {
    result = { status: 'error', message: err.message };
  }
  if (e.parameter.callback) {
    return ContentService.createTextOutput(e.parameter.callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function checkInAttendee(scannedId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!scannedId) return { status: 'invalid' };
    const sheet = getResponseSheet();
    const headers = getHeaderRow(sheet);
    const idCol = colIndex(headers, ID_HEADER);
    const lastCol  = colIndex(headers, LAST_NAME_HEADER);
    const firstCol = colIndex(headers, FIRST_NAME_HEADER);
    const midCol   = colIndex(headers, MIDDLE_NAME_HEADER);
    const statusCol = colIndex(headers, STATUS_HEADER);
    const timeCol = colIndex(headers, CHECKIN_TIME_HEADER);

    const idRange = sheet.getRange(2, idCol, Math.max(sheet.getLastRow() - 1, 1), 1);
    const match = idRange.createTextFinder(scannedId).matchEntireCell(true).findNext();
    if (!match) return { status: 'invalid' };

    const row = match.getRow();
    const fullName = [sheet.getRange(row, firstCol).getValue(), sheet.getRange(row, midCol).getValue(), sheet.getRange(row, lastCol).getValue()].filter(String).join(' ');
    const currentStatus = sheet.getRange(row, statusCol).getValue();

    if (currentStatus === 'Checked In') {
      const time = sheet.getRange(row, timeCol).getValue();
      return { status: 'duplicate', name: fullName, time: Utilities.formatDate(new Date(time), Session.getScriptTimeZone(), 'MMM d, h:mm a') };
    }
    const now = new Date();
    sheet.getRange(row, statusCol).setValue('Checked In');
    sheet.getRange(row, timeCol).setValue(now);
    return { status: 'success', name: fullName, time: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, h:mm a') };
  } finally {
    lock.releaseLock();
  }
}

// Generates a QR code image with automatic retries — isolates all QR logic in one place
function generateQRCodeBlob(uniqueId) {
  const qrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(uniqueId) + '&size=300';
  const maxRetries = 3;
  let lastError;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        return response.getBlob().setName('qr.png');
      }
      lastError = new Error('QR API returned status ' + response.getResponseCode());
    } catch (err) {
      lastError = err;
    }
    Utilities.sleep(1000 * attempt); // wait longer each retry: 1s, 2s, 3s
  }
  throw lastError;
}

function generateQRCodeBlob(uniqueId) {
  const qrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(uniqueId) + '&size=300';
  const maxRetries = 3;
  let lastError;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        return response.getBlob().setName('qr.png');
      }
      lastError = new Error('QR API returned status ' + response.getResponseCode());
    } catch (err) {
      lastError = err;
    }
    Utilities.sleep(1000 * attempt);
  }
  throw lastError;
}

function registerHOTS(scannedId, company, position) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!scannedId || !company || !position) return { status: 'invalid' };
    const sheet = getResponseSheet();
    const headers = getHeaderRow(sheet);
    const idCol = colIndex(headers, ID_HEADER);
    const lastCol = colIndex(headers, LAST_NAME_HEADER);
    const firstCol = colIndex(headers, FIRST_NAME_HEADER);
    const midCol = colIndex(headers, MIDDLE_NAME_HEADER);
    const statusCol = colIndex(headers, STATUS_HEADER);
    const resultCol = colIndex(headers, INTERVIEW_RESULT_HEADER);

    const idRange = sheet.getRange(2, idCol, Math.max(sheet.getLastRow() - 1, 1), 1);
    const match = idRange.createTextFinder(scannedId).matchEntireCell(true).findNext();
    if (!match) return { status: 'invalid' };

    const row = match.getRow();
    const fullName = [sheet.getRange(row, firstCol).getValue(), sheet.getRange(row, midCol).getValue(), sheet.getRange(row, lastCol).getValue()].filter(String).join(' ');
    const currentStatus = sheet.getRange(row, statusCol).getValue();
    if (currentStatus !== 'Checked In') return { status: 'not_checked_in', name: fullName };

    const now = new Date();
    const hotsSheet = getHotsSheet();
    hotsSheet.appendRow([now, scannedId, fullName, company, position]);

    const hotsLastRow = hotsSheet.getLastRow();
    let countForThisId = 0;
    if (hotsLastRow >= 2) {
      hotsSheet.getRange(2, 2, hotsLastRow - 1, 1).getValues().forEach(r => { if (r[0] === scannedId) countForThisId++; });
    }
    sheet.getRange(row, resultCol).setValue('HOTS x' + countForThisId);

    return {
      status: 'success', name: fullName, company: company, position: position,
      time: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, h:mm a'), count: countForThisId
    };
  } finally {
    lock.releaseLock();
  }
}

function getStats() {
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const statusCol = colIndex(headers, STATUS_HEADER);
  const lastRow = sheet.getLastRow();
  let checkedIn = 0;
  if (lastRow >= 2) {
    sheet.getRange(2, statusCol, lastRow - 1, 1).getValues().forEach(r => { if (r[0] === 'Checked In') checkedIn++; });
  }
  const hotsSheet = getHotsSheet();
  const hotsLastRow = hotsSheet.getLastRow();
  const hots = hotsLastRow >= 2 ? hotsLastRow - 1 : 0;
  return { status: 'success', checkedIn: checkedIn, hots: hots };
}

function getCheckedInList() {
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const statusCol = colIndex(headers, STATUS_HEADER);
  const timeCol = colIndex(headers, CHECKIN_TIME_HEADER);
  const lastCol = colIndex(headers, LAST_NAME_HEADER);
  const firstCol = colIndex(headers, FIRST_NAME_HEADER);
  const midCol = colIndex(headers, MIDDLE_NAME_HEADER);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'success', list: [] };

  const numRows = lastRow - 1;
  const statusValues = sheet.getRange(2, statusCol, numRows, 1).getValues();
  const timeValues = sheet.getRange(2, timeCol, numRows, 1).getValues();
  const lastValues = sheet.getRange(2, lastCol, numRows, 1).getValues();
  const firstValues = sheet.getRange(2, firstCol, numRows, 1).getValues();
  const midValues = sheet.getRange(2, midCol, numRows, 1).getValues();

  const list = [];
  for (let i = 0; i < numRows; i++) {
    if (statusValues[i][0] === 'Checked In') {
      const fullName = [firstValues[i][0], midValues[i][0], lastValues[i][0]].filter(String).join(' ');
      const timeVal = timeValues[i][0];
      const timeStr = timeVal ? Utilities.formatDate(new Date(timeVal), Session.getScriptTimeZone(), 'MMM d, h:mm a') : '';
      list.push({ name: fullName, time: timeStr });
    }
  }
  list.reverse(); // most recently checked-in first
  return { status: 'success', list: list };
}

const HOTS_SHEET_NAME = 'HOTS Log';

function getHotsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HOTS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HOTS_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Unique ID', 'Full Name', 'Company/Agency', 'Position']);
  }
  return sheet;
}

// Looks up a jobseeker and any prior HOTS records, WITHOUT recording anything yet
function getHotsInfo(scannedId) {
  if (!scannedId) return { status: 'invalid' };
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const idCol = colIndex(headers, ID_HEADER);
  const lastCol = colIndex(headers, LAST_NAME_HEADER);
  const firstCol = colIndex(headers, FIRST_NAME_HEADER);
  const midCol = colIndex(headers, MIDDLE_NAME_HEADER);
  const statusCol = colIndex(headers, STATUS_HEADER);

  const idRange = sheet.getRange(2, idCol, Math.max(sheet.getLastRow() - 1, 1), 1);
  const match = idRange.createTextFinder(scannedId).matchEntireCell(true).findNext();
  if (!match) return { status: 'invalid' };

  const row = match.getRow();
  const fullName = [sheet.getRange(row, firstCol).getValue(), sheet.getRange(row, midCol).getValue(), sheet.getRange(row, lastCol).getValue()].filter(String).join(' ');
  const currentStatus = sheet.getRange(row, statusCol).getValue();

  if (currentStatus !== 'Checked In') {
    return { status: 'not_checked_in', name: fullName };
  }

  const hotsSheet = getHotsSheet();
  const hotsLastRow = hotsSheet.getLastRow();
  const priorRecords = [];
  if (hotsLastRow >= 2) {
    const data = hotsSheet.getRange(2, 1, hotsLastRow - 1, 5).getValues();
    data.forEach(r => {
      if (r[1] === scannedId) priorRecords.push({ company: r[3], position: r[4] });
    });
  }

  return { status: 'ok', name: fullName, priorRecords: priorRecords };
}

// Addition of Vacancy List
const VACANCY_SHEET_NAME = 'Vacancy List';
const VACANCY_COMPANY_HEADER = 'Company Name';  // EDIT to match your actual column header
const VACANCY_POSITION_HEADER = 'Position';     // EDIT to match your actual column header

function getVacancyList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VACANCY_SHEET_NAME);
  if (!sheet) return { status: 'success', vacancies: [] };

  const headers = getHeaderRow(sheet);
  const companyCol = headers.indexOf(VACANCY_COMPANY_HEADER) + 1;
  const positionCol = headers.indexOf(VACANCY_POSITION_HEADER) + 1;
  if (!companyCol || !positionCol) return { status: 'success', vacancies: [] };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'success', vacancies: [] };

  const companies = sheet.getRange(2, companyCol, lastRow - 1, 1).getValues();
  const positions = sheet.getRange(2, positionCol, lastRow - 1, 1).getValues();

  const vacancies = [];
  for (let i = 0; i < companies.length; i++) {
    const c = String(companies[i][0]).trim();
    const p = String(positions[i][0]).trim();
    if (c && p) vacancies.push({ company: c, position: p });
  }
  return { status: 'success', vacancies: vacancies };
}

// Get all registrants as Fallback
function getAllRegistrants() {
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const idCol = colIndex(headers, ID_HEADER);
  const statusCol = colIndex(headers, STATUS_HEADER);
  const timeCol = colIndex(headers, CHECKIN_TIME_HEADER);
  const lastCol = colIndex(headers, LAST_NAME_HEADER);
  const firstCol = colIndex(headers, FIRST_NAME_HEADER);
  const midCol = colIndex(headers, MIDDLE_NAME_HEADER);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'success', registrants: [] };

  const numRows = lastRow - 1;
  const ids = sheet.getRange(2, idCol, numRows, 1).getValues();
  const statuses = sheet.getRange(2, statusCol, numRows, 1).getValues();
  const times = sheet.getRange(2, timeCol, numRows, 1).getValues();
  const lasts = sheet.getRange(2, lastCol, numRows, 1).getValues();
  const firsts = sheet.getRange(2, firstCol, numRows, 1).getValues();
  const mids = sheet.getRange(2, midCol, numRows, 1).getValues();

  const registrants = [];
  for (let i = 0; i < numRows; i++) {
    const id = ids[i][0];
    if (!id) continue;
    const fullName = [firsts[i][0], mids[i][0], lasts[i][0]].filter(String).join(' ');
    const timeVal = times[i][0];
    const timeStr = timeVal ? Utilities.formatDate(new Date(timeVal), Session.getScriptTimeZone(), 'MMM d, h:mm a') : '';
    registrants.push({ id: id, name: fullName, status: statuses[i][0] || 'Not Checked In', time: timeStr });
  }
  return { status: 'success', registrants: registrants };
}
