/**
 * ====================================================================
 * CONFIGURATION — edit these to match your Form's exact column headers
 * ====================================================================
 */
const EMAIL_COLUMN_HEADER = 'Email:';
const LAST_NAME_HEADER    = 'Last Name:';
const FIRST_NAME_HEADER   = 'First Name:';
const MIDDLE_NAME_HEADER  = 'Middle Name:';
const GENDER_HEADER       = 'Gender:';
const BIRTHDATE_HEADER    = 'Birthdate:';
const BARANGAY_HEADER     = 'Barangay:';
const MUNICIPALITY_HEADER = 'Municipality/City:';
const PROVINCE_HEADER     = 'Province:';
const EVENT_NAME = 'Job Fair 2026';

// Columns this script manages on the main response sheet
const ID_HEADER = 'Unique ID';
const STATUS_HEADER = 'Status';
const CHECKIN_TIME_HEADER = 'Check-in Time';
const EMAIL_STATUS_HEADER = 'Email Status';
const INTERVIEW_RESULT_HEADER = 'Interview Result';

// Other sheets in this workbook
const INTERVIEW_SHEET_NAME = 'Interview Log';
const VACANCY_SHEET_NAME = 'Vacancy List';
const VACANCY_COMPANY_HEADER = 'Company Name';   // EDIT to match your actual column header
const VACANCY_POSITION_HEADER = 'Position';      // EDIT to match your actual column header


/**
 * ====================================================================
 * SHEET / HEADER HELPERS
 * ====================================================================
 */
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

function getInterviewSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INTERVIEW_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INTERVIEW_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Unique ID', 'Full Name', 'Full Address', 'Sex', 'Age', 'Position', 'Company', 'Interview Status']);
  }
  return sheet;
}

function calculateAge(birthdate) {
  if (!birthdate) return '';
  const bd = new Date(birthdate);
  if (isNaN(bd.getTime())) return '';
  const today = new Date();
  let age = today.getFullYear() - bd.getFullYear();
  const m = today.getMonth() - bd.getMonth();
  if (m < 0 || (m === 0 && today.getDate() < bd.getDate())) age--;
  return age;
}


/**
 * ====================================================================
 * REGISTRATION — fires on each new Form submission
 * ====================================================================
 */
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


/**
 * ====================================================================
 * VERIFIER WEB APP — entry point
 * ====================================================================
 */
function doGet(e) {
  let result;
  try {
    if (e.parameter.action === 'checkin' && e.parameter.id) {
      result = checkInAttendee(e.parameter.id);
    } else if (e.parameter.action === 'hots_info' && e.parameter.id) {
      result = getInterviewInfo(e.parameter.id);
    } else if (e.parameter.action === 'hots' && e.parameter.id && e.parameter.company && e.parameter.position && e.parameter.interviewStatus) {
      result = registerInterview(e.parameter.id, e.parameter.company, e.parameter.position, e.parameter.interviewStatus);
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


/**
 * ====================================================================
 * CHECK-IN
 * ====================================================================
 */
function checkInAttendee(scannedId) {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000);
  if (!gotLock) return { status: 'error', message: 'System busy, please try again in a moment.' };
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
    const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]; // ONE read instead of four
    const fullName = [rowValues[firstCol - 1], rowValues[midCol - 1], rowValues[lastCol - 1]].filter(String).join(' ');
    const currentStatus = rowValues[statusCol - 1];

    if (currentStatus === 'Checked In') {
      const time = rowValues[timeCol - 1];
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


/**
 * ====================================================================
 * INTERVIEW LOG (Qualified / Not Qualified / Hired On The Spot / Near Hires)
 * ====================================================================
 */

// Looks up a jobseeker and any prior interview records, WITHOUT recording anything yet
function getInterviewInfo(scannedId) {
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
  if (currentStatus !== 'Checked In') return { status: 'not_checked_in', name: fullName };

  const interviewSheet = getInterviewSheet();
  const lastRow = interviewSheet.getLastRow();
  const priorRecords = [];
  if (lastRow >= 2) {
    interviewSheet.getRange(2, 1, lastRow - 1, 9).getValues().forEach(r => {
      if (r[1] === scannedId) priorRecords.push({ company: r[7], position: r[6], interviewStatus: r[8] });
    });
  }
  return { status: 'ok', name: fullName, priorRecords: priorRecords };
}

// Records one interview outcome. Sex, Age, and Full Address are pulled automatically
// from the jobseeker's own registration data — staff only supply Company, Position, and Status.
function registerInterview(scannedId, company, position, interviewStatus) {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000);
  if (!gotLock) return { status: 'error', message: 'System busy, please try again in a moment.' };
  try {
    if (!scannedId || !company || !position || !interviewStatus) return { status: 'invalid' };
    const sheet = getResponseSheet();
    const headers = getHeaderRow(sheet);
    const idCol = colIndex(headers, ID_HEADER);
    const lastCol = colIndex(headers, LAST_NAME_HEADER);
    const firstCol = colIndex(headers, FIRST_NAME_HEADER);
    const midCol = colIndex(headers, MIDDLE_NAME_HEADER);
    const statusCol = colIndex(headers, STATUS_HEADER);
    const resultCol = colIndex(headers, INTERVIEW_RESULT_HEADER);
    const genderCol = colIndex(headers, GENDER_HEADER);
    const birthdateCol = colIndex(headers, BIRTHDATE_HEADER);
    const barangayCol = colIndex(headers, BARANGAY_HEADER);
    const municipalityCol = colIndex(headers, MUNICIPALITY_HEADER);
    const provinceCol = colIndex(headers, PROVINCE_HEADER);

    const idRange = sheet.getRange(2, idCol, Math.max(sheet.getLastRow() - 1, 1), 1);
    const match = idRange.createTextFinder(scannedId).matchEntireCell(true).findNext();
    if (!match) return { status: 'invalid' };

    const row = match.getRow();
    const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0]; // ONE read instead of eight
    const fullName = [rowValues[firstCol - 1], rowValues[midCol - 1], rowValues[lastCol - 1]].filter(String).join(' ');
    const currentStatus = rowValues[statusCol - 1];
    if (currentStatus !== 'Checked In') return { status: 'not_checked_in', name: fullName };

    const sex = rowValues[genderCol - 1];
    const age = calculateAge(rowValues[birthdateCol - 1]);
    const fullAddress = [rowValues[barangayCol - 1], rowValues[municipalityCol - 1], rowValues[provinceCol - 1]].filter(String).join(', ');

    const now = new Date();
    const interviewSheet = getInterviewSheet();
    interviewSheet.appendRow([now, scannedId, fullName, fullAddress, sex, age, position, company, interviewStatus]);

    const lastRow = interviewSheet.getLastRow();
    let totalCountForId = 0;
    if (lastRow >= 2) {
      interviewSheet.getRange(2, 2, lastRow - 1, 1).getValues().forEach(r => { if (r[0] === scannedId) totalCountForId++; });
    }
    sheet.getRange(row, resultCol).setValue(interviewStatus + (totalCountForId > 1 ? ' (x' + totalCountForId + ')' : ''));

    return {
      status: 'success', name: fullName, company: company, position: position, interviewStatus: interviewStatus,
      time: Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMM d, h:mm a')
    };
  } finally {
    lock.releaseLock();
  }
}


/**
 * ====================================================================
 * DASHBOARD STATS / LISTS
 * ====================================================================
 */
function getStats() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('stats_cache');
  if (cached) return JSON.parse(cached);

  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const statusCol = colIndex(headers, STATUS_HEADER);
  const lastRow = sheet.getLastRow();
  let checkedIn = 0;
  if (lastRow >= 2) {
    sheet.getRange(2, statusCol, lastRow - 1, 1).getValues().forEach(r => { if (r[0] === 'Checked In') checkedIn++; });
  }
  const interviewSheet = getInterviewSheet();
  const interviewLastRow = interviewSheet.getLastRow();
  let hots = 0;
  if (interviewLastRow >= 2) {
    interviewSheet.getRange(2, 9, interviewLastRow - 1, 1).getValues().forEach(r => { if (r[0] === 'Hired On The Spot') hots++; });
  }
  const result = { status: 'success', checkedIn: checkedIn, hots: hots };
  cache.put('stats_cache', JSON.stringify(result), 10);
  return result;
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

// Full registrant list (checked-in or not) — powers the manual check-in/HOTS fallback in search
function getAllRegistrants() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('all_registrants_cache');
  if (cached) return JSON.parse(cached);

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
  const result = { status: 'success', registrants: registrants };
  cache.put('all_registrants_cache', JSON.stringify(result), 10);
  return result;
}


/**
 * ====================================================================
 * VACANCY LIST (for the Company/Position autocomplete in the Interview form)
 * ====================================================================
 */
function getVacancyList() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('vacancy_cache');
  if (cached) return JSON.parse(cached);

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
  const result = { status: 'success', vacancies: vacancies };
  cache.put('vacancy_cache', JSON.stringify(result), 300);
  return result;
}
