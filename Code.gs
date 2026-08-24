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

function sendTicketEmail(email, name, uniqueId, referenceNo) {
  const qrBlob = generateQRCodeBlob(uniqueId);
  // Applicant-controlled values are escaped before HTML interpolation.
  const safeName = escapeHtml_(name);

  const htmlBody = `
    <p>Hi ${safeName},</p>
    <p>You're pre-registered for <b>${EVENT_NAME}</b>. Please show this QR code at check-in (a screenshot works fine):</p>
    <p><img src="cid:qr"/></p>
    <p>Reference ID: <b>${escapeHtml_(referenceNo || uniqueId)}</b></p>
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
    } else if (e.parameter.action === 'ticket' && e.parameter.email) {
      result = lookupTicket(e.parameter.email);
    } else if (e.parameter.action === 'form_choices') {
      result = getPublicFormChoices();
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


/**
 * ====================================================================
 * APPLICANT REGISTRATION FORM BACKEND (doPost)
 * Receives submissions from index.html / api.js
 *
 * Frontend → POST JSON { action:'register', fields:{...}, files:{...} }
 * Files are data URLs (base64) and are stored in Google Drive; the
 * Drive file link is written into the corresponding sheet column.
 * ====================================================================
 */

// Canonical sheet headers, in the exact required order.
const REGISTRATION_HEADERS = [
  'Timestamp',
  'Which additional PESO assistance programs would you like to enroll in?',
  'Data Subject Rights',
  'Location for Interview:',
  'Are you a First-time Jobseeker?',
  'Employment Preference:',
  'Last Name:',
  'First Name:',
  'Middle Name:',
  'Birthdate:',
  'Gender:',
  'Civil Status:',
  'Barangay:',
  'Municipality/City:',
  'Province:',
  'Contact No.:',
  'Email:',
  'Returning OFW?',
  'Returning Worker',
  'Interested in Skills Training?',
  'If yes, please indicate your preferred training program.',
  'Kindly upload a copy of your resume.',
  'Do you have a disability?',
  'Type of Disability:',
  'Kindly upload a copy of your PWD ID.',
  'Local Positions:',
  'Overseas Opportunities:',
  'What is your highest educational attainment?',
  'What is the highest degree, course, program, or strand you have enrolled in or attended?',
  'What are your related work experiences?',
  'What related skills, trainings, or certificates do you possess?',
  ID_HEADER,
  STATUS_HEADER,
  CHECKIN_TIME_HEADER,
  EMAIL_STATUS_HEADER,
  INTERVIEW_RESULT_HEADER
];

const UPLOAD_FOLDER_NAME = 'Job Fair Applicant Uploads';
const MAX_UPLOAD_BYTES = 8 * 1024 * 1024; // 8 MB per file
const REFERENCE_NUMBER_PREFIX = 'JF26-';

/* --------------------------------------------------------------------
 * REMEDIATION CONSTANTS (audit D-1/D-2/D-3/E-1/E-2/E-3)
 * Allowed values mirror index.html / api.js exactly.
 * KEEP IN SYNC: any option list change must be applied to BOTH sides.
 * -------------------------------------------------------------------- */
const MAX_PAYLOAD_CHARS = 15000000;          // ~15 MB request body cap (legit max ≈ 13.4 MB base64)
const MAX_FIELDS = 40;                       // canonical payload carries 30 fields
const MAX_STRING_LENGTH = 1000;              // hard server-side cap per field

const EMAIL_COOLDOWN_WINDOW_SECS = 3600;     // per-email registration cooldown
const EMAIL_COOLDOWN_MAX_PER_WINDOW = 3;     // accepted submissions per email per hour

const TICKET_LOOKUP_WINDOW_SECS = 60;        // per-email ticket lookup cooldown
const TICKET_LOOKUP_MAX_PER_WINDOW = 5;

// Columns written as real Date values; every OTHER column of a new row is
// forced to plain-text ('@') format so leading = + - @ TAB characters are
// stored literally and can never execute as spreadsheet formulas.
const TEXT_FORMAT_EXEMPT_HEADERS = { 'Timestamp': true, 'Birthdate:': true };

/* --------------------------------------------------------------------
 * FORM CHOICES CONFIGURATION
 * Managed selectable fields. Canonical keys/labels/types are FIXED by the
 * application; administrators manage only choices (via the "Form Choices"
 * sheet). These defaults are also the initial seed and the fail-open
 * fallback if the configuration sheet cannot be read.
 * KEEP IN SYNC with index.html static fallback options.
 * -------------------------------------------------------------------- */
const FORM_CHOICES_SHEET_NAME = 'Form Choices';
const FORM_CHOICES_HEADERS = ['Choice ID', 'Field Key', 'Choice Value', 'Display Label', 'Active', 'Sort Order', 'Created At', 'Updated At'];
const FORM_CHOICES_CACHE_KEY = 'form_choices_cache_v1';
const FORM_CHOICES_CACHE_SECS = 300;
const CHOICES_AUDIT_SHEET_NAME = 'Choices Audit Log';
const CHOICES_AUDIT_HEADERS = ['Timestamp', 'Action', 'Field Key', 'Choice ID', 'Old Value', 'New Value', 'Note'];
const FORM_CHOICES_ADMIN_PROP = 'FORM_CHOICES_ADMIN_KEY';
const ADMIN_FAIL_LOCK_WINDOW_SECS = 900;
const ADMIN_FAIL_MAX = 10;

const FORM_CHOICE_FIELDS = {
  interviewLocation: { label: 'Location for Interview:', inputType: 'select' },
  gender:            { label: 'Gender:', inputType: 'select' },
  civilStatus:       { label: 'Civil Status:', inputType: 'select' },
  empPref:           { label: 'Employment Preference:', inputType: 'checkbox' },
  pesoPrograms:      { label: 'Which additional PESO assistance programs would you like to enroll in?', inputType: 'checkbox', protectedValues: ['None'] },
  education:         { label: 'What is your highest educational attainment?', inputType: 'select' },
  firstTime:         { label: 'Are you a First-time Jobseeker?', inputType: 'radio', locked: true },
  returningOfw:      { label: 'Returning OFW?', inputType: 'radio', locked: true },
  returningWorker:   { label: 'Returning Worker', inputType: 'radio', locked: true },
  skillsTraining:    { label: 'Interested in Skills Training?', inputType: 'radio', locked: true },
  disability:        { label: 'Do you have a disability?', inputType: 'radio', locked: true }
};

// Yes/No answers drive conditional logic (PWD fields, training program) and
// PESO "None" drives exclusivity — those VALUES are application control
// values and are seeded as protected; yes/no fields are fully locked.
const DEFAULT_FORM_CHOICES = {
  interviewLocation: [
    'PESO Cagayan - Provincial Capitol, Tuguegarao City',
    'Provincial Capitol Grounds, Tuguegarao City'
  ],
  gender: ['Male', 'Female'],
  civilStatus: ['Single', 'Married', 'Widowed', 'Separated', 'Annulled'],
  empPref: ['Local', 'Overseas'],
  education: [
    'No Formal Education',
    'Elementary Level',
    'Elementary Graduate',
    'High School Level',
    'High School Graduate',
    'Senior High School Level',
    'Senior High School Graduate',
    'Vocational / Technical Course',
    'College Level',
    'College Graduate',
    'Post Graduate Level',
    'Post Graduate'
  ],
  pesoPrograms: [
    'Job Referral and Placement',
    'Special Program for Employment of Students (SPES)',
    'Government Internship Program (GIP)',
    'JobStart Philippines Program',
    'TESDA Skills Training / Scholarship',
    'Pre-Employment Counseling / Career Guidance',
    'Livelihood / Self-Employment Assistance',
    'None'
  ],
  firstTime: ['Yes', 'No'],
  returningOfw: ['Yes', 'No'],
  returningWorker: ['Yes', 'No'],
  skillsTraining: ['Yes', 'No'],
  disability: ['Yes', 'No']
};

const UPLOAD_RULES = {
  'resume': {
    label: 'Resume',
    extensions: ['pdf', 'doc', 'docx'],
    mimeTypes: ['application/pdf', 'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document']
  },
  'pwdId': {
    label: 'PWD ID',
    extensions: ['jpg', 'jpeg', 'png', 'pdf'],
    mimeTypes: ['image/jpeg', 'image/png', 'application/pdf']
  }
};

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let payload;
  try {
    if (!e || !e.postData || typeof e.postData.contents !== 'string') {
      return jsonOut({ status: 'error', code: 'malformed', message: 'Empty request.' });
    }
    if (e.postData.contents.length > MAX_PAYLOAD_CHARS) {
      console.warn('doPost: rejected oversized payload (' + e.postData.contents.length + ' chars)');
      return jsonOut({ status: 'error', code: 'malformed', message: 'Request too large.' });
    }
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    console.error('doPost: malformed payload');
    return jsonOut({ status: 'error', code: 'malformed', message: 'Invalid request.' });
  }

  if (!payload || typeof payload.action !== 'string') {
    return jsonOut({ status: 'error', code: 'malformed', message: 'Unknown action.' });
  }

  // Administrative Form Choices mutations — server-authorized, never public.
  if (FORM_CHOICE_ADMIN_ACTIONS[payload.action]) {
    try {
      return jsonOut(handleFormChoiceMutation_(payload));
    } catch (err) {
      console.error('doPost: form choice mutation failed: ' + err.message);
      return jsonOut({ status: 'error', code: 'internal', message: 'Unexpected server error.' });
    }
  }

  if (payload.action !== 'register' || typeof payload.fields !== 'object' || payload.fields === null) {
    return jsonOut({ status: 'error', code: 'malformed', message: 'Unknown action.' });
  }
  if (Object.keys(payload.fields).length > MAX_FIELDS) {
    console.warn('doPost: rejected payload with excessive fields');
    return jsonOut({ status: 'error', code: 'malformed', message: 'Too many fields.' });
  }

  // Honeypot — hidden form field real applicants can never fill.
  // Rejected before any Sheet/Drive/Gmail work and without an explanatory
  // message so bots get no signal to adapt to.
  if (String(payload.honeypot || '').trim() !== '') {
    console.warn('doPost: honeypot triggered — rejecting.');
    return jsonOut({ status: 'error', code: 'rejected', message: 'Submission not accepted.' });
  }

  try {
    return jsonOut(registerApplicant(payload));
  } catch (err) {
    console.error('doPost: unexpected error: ' + err.message);
    return jsonOut({ status: 'error', code: 'internal', message: 'Unexpected server error.' });
  }
}

function registerApplicant(data) {
  const lock = LockService.getScriptLock();
  // tryLock (not waitLock): a timeout is reported as 'busy' instead of
  // falling into the generic catch and masquerading as a malformed request.
  if (!lock.tryLock(30000)) {
    return { status: 'error', code: 'busy', message: 'System busy, please try again in a moment.' };
  }
  try {
    ensureRegistrationHeaders();
    const sheet = getResponseSheet();
    const headers = getHeaderRow(sheet);

    const f = {};
    Object.keys(data.fields || {}).forEach(k => {
      const v = String(data.fields[k] == null ? '' : data.fields[k]).trim();
      f[String(k)] = v.length > MAX_STRING_LENGTH ? v.slice(0, MAX_STRING_LENGTH) : v;
    });

    const invalid = validateRegistrationFields_(f);
    if (invalid.length) {
      return { status: 'invalid', code: 'validation', message: 'Missing or invalid: ' + invalid.join(', ') };
    }

    const dup = findDuplicateRow_(sheet, headers, f[EMAIL_COLUMN_HEADER], f[LAST_NAME_HEADER], f[FIRST_NAME_HEADER]);
    if (dup) {
      // Deliberately minimal: reference number only — the internal UUID,
      // sheet row, and Drive IDs are never exposed on the duplicate path.
      return {
        status: 'duplicate',
        message: 'An applicant with this email and name has already been registered.',
        referenceNo: dup.referenceNo
      };
    }

    if (!rateLimit_('reg', f[EMAIL_COLUMN_HEADER], EMAIL_COOLDOWN_MAX_PER_WINDOW, EMAIL_COOLDOWN_WINDOW_SECS)) {
      return { status: 'error', code: 'rate_limited', message: 'Too many submissions for this email right now. Please try again later.' };
    }

    const now = new Date();
    const uniqueId = Utilities.getUuid();

    // Uploads are tracked so a later failure can clean up earlier files.
    const createdFiles = [];
    let resumeValue = '';
    let pwdIdValue = '';
    try {
      if (data.files && data.files.resume) {
        const resumeFile = saveUploadFile_(data.files.resume, uniqueId, 'resume');
        resumeValue = resumeFile.getUrl();
        createdFiles.push(resumeFile);
      }
      if (data.files && data.files.pwdId) {
        const pwdFile = saveUploadFile_(data.files.pwdId, uniqueId, 'pwdId');
        pwdIdValue = pwdFile.getUrl();
        createdFiles.push(pwdFile);
      }
    } catch (uploadErr) {
      trashCreatedFiles_(createdFiles);
      console.error('Upload failed: ' + uploadErr.message);
      return { status: 'error', code: 'upload_failed', message: uploadErr.message };
    }

    const birthRaw = f[BIRTHDATE_HEADER];
    const birthDate = new Date(birthRaw);
    const values = {};
    values['Timestamp'] = now;
    values['Which additional PESO assistance programs would you like to enroll in?'] = f['Which additional PESO assistance programs would you like to enroll in?'] || '';
    values['Data Subject Rights'] = f['Data Subject Rights'];
    values['Location for Interview:'] = f['Location for Interview:'];
    values['Are you a First-time Jobseeker?'] = f['Are you a First-time Jobseeker?'];
    values['Employment Preference:'] = f['Employment Preference:'];
    values[LAST_NAME_HEADER] = f[LAST_NAME_HEADER];
    values[FIRST_NAME_HEADER] = f[FIRST_NAME_HEADER];
    values[MIDDLE_NAME_HEADER] = f[MIDDLE_NAME_HEADER] || '';
    values[BIRTHDATE_HEADER] = isNaN(birthDate.getTime()) ? birthRaw : birthDate;
    values[GENDER_HEADER] = f[GENDER_HEADER];
    values['Civil Status:'] = f['Civil Status:'];
    values[BARANGAY_HEADER] = f[BARANGAY_HEADER];
    values[MUNICIPALITY_HEADER] = f[MUNICIPALITY_HEADER];
    values[PROVINCE_HEADER] = f[PROVINCE_HEADER];
    values['Contact No.:'] = f['Contact No.:'] || '';
    values[EMAIL_COLUMN_HEADER] = f[EMAIL_COLUMN_HEADER];
    values['Returning OFW?'] = f['Returning OFW?'] || '';
    values['Returning Worker'] = f['Returning Worker'] || '';
    values['Interested in Skills Training?'] = f['Interested in Skills Training?'] || '';
    values['If yes, please indicate your preferred training program.'] = f['If yes, please indicate your preferred training program.'] || '';
    values['Kindly upload a copy of your resume.'] = resumeValue;
    values['Do you have a disability?'] = f['Do you have a disability?'] || '';
    values['Type of Disability:'] = f['Type of Disability:'] || '';
    values['Kindly upload a copy of your PWD ID.'] = pwdIdValue;
    values['Local Positions:'] = f['Local Positions:'] || '';
    values['Overseas Opportunities:'] = f['Overseas Opportunities:'] || '';
    values['What is your highest educational attainment?'] = f['What is your highest educational attainment?'] || '';
    values['What is the highest degree, course, program, or strand you have enrolled in or attended?'] = f['What is the highest degree, course, program, or strand you have enrolled in or attended?'] || '';
    values['What are your related work experiences?'] = f['What are your related work experiences?'] || '';
    values['What related skills, trainings, or certificates do you possess?'] = f['What related skills, trainings, or certificates do you possess?'] || '';
    values[ID_HEADER] = uniqueId;
    values[STATUS_HEADER] = 'Not Checked In';
    values[CHECKIN_TIME_HEADER] = '';
    values[EMAIL_STATUS_HEADER] = 'Pending';
    values[INTERVIEW_RESULT_HEADER] = 'Pending';

    const rowArr = headers.map(h => (values.hasOwnProperty(h) ? values[h] : ''));

    // Text-safe write: applicant columns are forced to plain-text format so
    // leading = + - @ TAB characters are stored literally, never as formulas.
    // Reference number is derived from the ACTUAL written row.
    let rowIndex;
    try {
      rowIndex = writeRegistrationRow_(sheet, headers, rowArr);
    } catch (storeErr) {
      trashCreatedFiles_(createdFiles);
      console.error('Sheet write failed: ' + storeErr.message);
      return { status: 'error', code: 'storage_failed', message: 'We could not save your registration. Please try again.' };
    }
    const referenceNo = REFERENCE_NUMBER_PREFIX + ('00000' + Math.max(rowIndex - 1, 1)).slice(-5);

    clearRegistrationCaches_();

    const fullName = [f[FIRST_NAME_HEADER], f[MIDDLE_NAME_HEADER], f[LAST_NAME_HEADER]].filter(String).join(' ');

    let emailSent = false;
    try {
      sendTicketEmail(f[EMAIL_COLUMN_HEADER], fullName, uniqueId, referenceNo);
      sheet.getRange(rowIndex, colIndex(headers, EMAIL_STATUS_HEADER)).setValue('Sent');
      emailSent = true;
    } catch (emailErr) {
      console.error('Ticket email failed: ' + emailErr.message);
      sheet.getRange(rowIndex, colIndex(headers, EMAIL_STATUS_HEADER)).setValue('Pending - ' + emailErr.message);
    }

    return {
      status: 'success',
      uniqueId: uniqueId,
      referenceNo: referenceNo,
      emailSent: emailSent,
      name: fullName
    };
  } finally {
    lock.releaseLock();
  }
}

// Adds any canonical registration headers that are not yet present
// (appended at the end so existing columns are never moved or renamed).
function ensureRegistrationHeaders() {
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  let lastCol = sheet.getLastColumn();
  REGISTRATION_HEADERS.forEach(h => {
    if (headers.indexOf(h) === -1) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(h);
    }
  });
}

// One-time utility for a FRESH spreadsheet/tab: writes the canonical
// header order to row 1 of the first sheet. Refuses to overwrite a
// sheet that already contains data rows.
function setupRegistrationSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() > 1) {
    throw new Error('Sheet already contains data — refusing to rewrite headers.');
  }
  sheet.getRange(1, 1, 1, REGISTRATION_HEADERS.length).setValues([REGISTRATION_HEADERS]);
  sheet.setFrozenRows(1);
}

// Obvious-duplicate protection: same email AND same first+last name.
function findDuplicateRow_(sheet, headers, email, lastName, firstName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const numRows = lastRow - 1;

  const norm = v => String(v == null ? '' : v).trim().toLowerCase();
  const targetEmail = norm(email);
  const targetLast = norm(lastName);
  const targetFirst = norm(firstName);

  const emails = sheet.getRange(2, colIndex(headers, EMAIL_COLUMN_HEADER), numRows, 1).getValues();
  const lasts = sheet.getRange(2, colIndex(headers, LAST_NAME_HEADER), numRows, 1).getValues();
  const firsts = sheet.getRange(2, colIndex(headers, FIRST_NAME_HEADER), numRows, 1).getValues();
  const ids = sheet.getRange(2, colIndex(headers, ID_HEADER), numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    if (norm(emails[i][0]) === targetEmail &&
        norm(lasts[i][0]) === targetLast &&
        norm(firsts[i][0]) === targetFirst) {
      return {
        row: i + 2,
        uniqueId: ids[i][0],
        referenceNo: REFERENCE_NUMBER_PREFIX + ('00000' + (i + 1)).slice(-5)
      };
    }
  }
  return null;
}

// Expects { filename, mimeType, dataUrl }. Validates structure, size,
// extension and MIME server-side (client checks are advisory only).
// Returns the Drive File object so callers can clean up on later failure.
function saveUploadFile_(upload, uniqueId, kind) {
  const rules = UPLOAD_RULES[kind];
  if (!rules) throw new Error('Unsupported upload kind.');

  const fail = msg => { throw new Error(rules.label + ' ' + msg); };

  if (!upload || typeof upload.dataUrl !== 'string') {
    fail('file is missing or unreadable.');
  }
  const match = /^data:([^;]+);base64,([A-Za-z0-9+/=\s]+)$/.exec(upload.dataUrl);
  if (!match) fail('file data is malformed.');
  const bytes = Utilities.base64Decode(match[2].replace(/\s+/g, ''));
  if (!bytes.length) fail('file is empty.');
  if (bytes.length > MAX_UPLOAD_BYTES) fail('file exceeds the maximum size of 8 MB.');

  const filename = String(upload.filename || '');
  const dot = filename.lastIndexOf('.');
  const ext = dot === -1 ? '' : filename.slice(dot + 1).toLowerCase();
  if (rules.extensions.indexOf(ext) === -1) {
    fail('file type is not allowed. Allowed types: ' + rules.extensions.join(', ').toUpperCase() + '.');
  }

  const mimeType = String(upload.mimeType || '').toLowerCase().split(';')[0];
  if (mimeType && rules.mimeTypes.indexOf(mimeType) === -1) {
    fail('file format is not allowed.');
  }

  const safeName = filename.replace(/[^\w.\- ]/g, '_').slice(-60);
  const blob = Utilities.newBlob(bytes, mimeType || rules.mimeTypes[0], kind + '-' + uniqueId.slice(0, 8) + '-' + safeName);
  const file = getOrCreateUploadFolder_().createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareErr) {
    // Domain policies may forbid link sharing — the file still exists in Drive.
  }
  return file;
}

function getOrCreateUploadFolder_() {
  const it = DriveApp.getFoldersByName(UPLOAD_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(UPLOAD_FOLDER_NAME);
}

function clearRegistrationCaches_() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('stats_cache');
    cache.remove('all_registrants_cache');
  } catch (err) {
    // Cache is best-effort only
  }
}


/**
 * ====================================================================
 * REMEDIATION HELPERS
 * ====================================================================
 */

// HTML-escapes applicant-controlled values before email/HTML interpolation.
function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Server-side validation. Allowed values mirror the registration form
// exactly (see REMEDIATION CONSTANTS). Returns an array of problem labels.
function validateRegistrationFields_(f) {
  const bad = [];

  [LAST_NAME_HEADER, FIRST_NAME_HEADER, BIRTHDATE_HEADER, GENDER_HEADER,
   'Civil Status:', BARANGAY_HEADER, MUNICIPALITY_HEADER, PROVINCE_HEADER,
   'Contact No.:', EMAIL_COLUMN_HEADER, 'Location for Interview:',
   'Are you a First-time Jobseeker?', 'Employment Preference:', 'Data Subject Rights'
  ].forEach(h => { if (!f[h]) bad.push(h.replace(/:$/, '')); });

  if (f['Data Subject Rights'] && f['Data Subject Rights'] !== 'Consented') bad.push('Data Subject Rights');

  if (f[EMAIL_COLUMN_HEADER] && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(f[EMAIL_COLUMN_HEADER])) bad.push('valid Email');

  if (f['Contact No.:']) {
    const digits = f['Contact No.:'].replace(/[^0-9]/g, '');
    if (!/^[0-9+\-\s()]{7,20}$/.test(f['Contact No.:']) || digits.length < 9 || digits.length > 13) {
      bad.push('valid Contact No.');
    }
  }

  if (f[BIRTHDATE_HEADER]) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(f[BIRTHDATE_HEADER])) {
      bad.push('valid Birthdate');
    } else {
      const bd = new Date(f[BIRTHDATE_HEADER] + 'T00:00:00Z');
      const today = new Date();
      if (isNaN(bd.getTime())) {
        bad.push('valid Birthdate');
      } else if (bd.getTime() > Date.UTC(today.getFullYear(), today.getMonth(), today.getDate())) {
        bad.push('Birthdate in the future');
      } else if (bd.getUTCFullYear() < 1900) {
        bad.push('Birthdate year');
      }
    }
  }

  // Choice-managed fields are validated against ACTIVE configured choices
  // (Form Choices sheet). Falls back to seeded defaults only if the
  // configuration cannot be read. The browser is never the authority.
  const active = getActiveChoiceValueMap_();

  if (f[GENDER_HEADER] && active.gender && active.gender.indexOf(f[GENDER_HEADER]) === -1) bad.push('Gender value');
  if (f['Civil Status:'] && active.civilStatus && active.civilStatus.indexOf(f['Civil Status:']) === -1) bad.push('Civil Status value');
  if (f['Location for Interview:'] && active.interviewLocation && active.interviewLocation.indexOf(f['Location for Interview:']) === -1) bad.push('Location for Interview value');

  ['firstTime', 'returningOfw', 'returningWorker', 'skillsTraining', 'disability'].forEach(key => {
    const val = f[FORM_CHOICE_FIELDS[key].label];
    if (val && active[key] && active[key].indexOf(val) === -1) {
      bad.push(FORM_CHOICE_FIELDS[key].label.replace(/\?$/, '') + ' value');
    }
  });

  if (f['What is your highest educational attainment?'] && active.education && active.education.indexOf(f['What is your highest educational attainment?']) === -1) bad.push('educational attainment value');

  if (f['Employment Preference:']) {
    const prefs = f['Employment Preference:'].split(',').map(s => s.trim());
    if (!prefs.length || prefs.some(p => active.empPref && active.empPref.indexOf(p) === -1)) {
      bad.push('Employment Preference value');
    }
  }

  if (f['Which additional PESO assistance programs would you like to enroll in?']) {
    const progs = f['Which additional PESO assistance programs would you like to enroll in?'].split(';').map(s => s.trim());
    if (progs.some(p => active.pesoPrograms && active.pesoPrograms.indexOf(p) === -1)) bad.push('PESO assistance program value');
  }

  return bad;
}

// Sliding-window counter in CacheService, keyed by a hashed identifier.
// Used for registration cooldown and ticket-lookup throttling.
// Fails open so a CacheService outage cannot block legitimate applicants.
function rateLimit_(keyPrefix, value, maxPerWindow, windowSecs) {
  try {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value).toLowerCase(), Utilities.Charset.UTF_8);
    const hex = digest.map(b => ((b & 0xff) + 0x100).toString(16).slice(1)).join('');
    const key = keyPrefix + '_' + hex.slice(0, 24);
    const cache = CacheService.getScriptCache();
    const current = cache.get(key);
    const count = current ? parseInt(current, 10) : 0;
    if (count >= maxPerWindow) return false;
    cache.put(key, String(count + 1), windowSecs);
    return true;
  } catch (err) {
    console.error('rateLimit_ failed: ' + err.message);
    return true;
  }
}

// Writes one registration row WITHOUT formula interpretation.
// Strategy: create the target row, set plain-text ('@') number format on
// every column except Timestamp/Birthdate (which carry real Date values),
// then setValues — Sheets stores strings literally under '@' format.
function writeRegistrationRow_(sheet, headers, rowArr) {
  const rowIndex = sheet.getLastRow() + 1;

  let runStart = -1;
  for (let i = 0; i <= headers.length; i++) {
    const isText = i < headers.length && !TEXT_FORMAT_EXEMPT_HEADERS[headers[i]];
    if (isText && runStart === -1) runStart = i;
    if (!isText && runStart !== -1) {
      sheet.getRange(rowIndex, runStart + 1, 1, i - runStart).setNumberFormat('@');
      runStart = -1;
    }
  }

  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowArr]);
  return rowIndex;
}

// Best-effort rollback for Drive files created by a failed submission.
// Cleanup failures are logged individually, never silently swallowed.
function trashCreatedFiles_(files) {
  (files || []).forEach(file => {
    try {
      file.setTrashed(true);
    } catch (err) {
      console.error('Upload cleanup failed for ' + file.getId() + ': ' + err.message);
    }
  });
}

// Ticket lookup for ticket.html (action:'ticket').
// Returns only what the pass page renders: name + ticket UUID (the QR
// payload — identical to what the confirmation email already delivers to
// the same address). Rate-limited per email; no row numbers or Drive IDs.
function lookupTicket(rawEmail) {
  const email = String(rawEmail || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { status: 'error', message: 'Please enter a valid email address.' };
  }
  if (!rateLimit_('lookup', email, TICKET_LOOKUP_MAX_PER_WINDOW, TICKET_LOOKUP_WINDOW_SECS)) {
    return { status: 'error', message: 'Too many lookups right now. Please try again in a minute.' };
  }

  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'error', message: 'No registration found for this email address.' };

  const numRows = lastRow - 1;
  const norm = v => String(v == null ? '' : v).trim().toLowerCase();

  const emails = sheet.getRange(2, colIndex(headers, EMAIL_COLUMN_HEADER), numRows, 1).getValues();
  const firsts = sheet.getRange(2, colIndex(headers, FIRST_NAME_HEADER), numRows, 1).getValues();
  const mids = sheet.getRange(2, colIndex(headers, MIDDLE_NAME_HEADER), numRows, 1).getValues();
  const lasts = sheet.getRange(2, colIndex(headers, LAST_NAME_HEADER), numRows, 1).getValues();
  const ids = sheet.getRange(2, colIndex(headers, ID_HEADER), numRows, 1).getValues();
  const emailStatuses = sheet.getRange(2, colIndex(headers, EMAIL_STATUS_HEADER), numRows, 1).getValues();

  let found = -1;
  for (let i = numRows - 1; i >= 0; i--) {
    if (norm(emails[i][0]) === email) { found = i; break; }
  }
  if (found === -1) return { status: 'error', message: 'No registration found for this email address.' };

  // Pass not generated yet (quota/retry window) — let the page retry.
  const emailStatus = String(emailStatuses[found][0] || '');
  if (emailStatus.indexOf('Pending') === 0 || !ids[found][0]) {
    return { status: 'pending', message: 'Your pass is being generated. Please wait a moment.' };
  }

  const fullName = [firsts[found][0], mids[found][0], lasts[found][0]].filter(String).join(' ');
  return { status: 'success', name: fullName, id: ids[found][0] };
}


/**
 * ====================================================================
 * FORM CHOICES MANAGEMENT
 * Source of truth: "Form Choices" sheet. Public READ of ACTIVE choices;
 * admin WRITE requires a server-side shared key stored in Script
 * Properties (never in the repo / GitHub Pages).
 * ====================================================================
 */

const FORM_CHOICE_ADMIN_ACTIONS = {
  list_form_choice: true,
  add_form_choice: true,
  update_form_choice: true,
  toggle_form_choice: true,
  reorder_form_choice: true
};

// One-time provisioning: run manually from the Apps Script editor.
// The key is written to Script Properties and logged once (execution logs
// are visible only to the script owner). Enter this key in
// admin/form-choices.html to unlock administrative operations.
function generateFormChoicesAdminKey() {
  const key = 'fc_' + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(FORM_CHOICES_ADMIN_PROP, key);
  CacheService.getScriptCache().remove('admin_fail');
  Logger.log('FORM CHOICES ADMIN KEY (stored to Script Properties — copy it now): ' + key);
  return key;
}

function getFormChoicesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FORM_CHOICES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FORM_CHOICES_SHEET_NAME);
    sheet.getRange(1, 1, 1, FORM_CHOICES_HEADERS.length).setValues([FORM_CHOICES_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getChoicesAuditSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CHOICES_AUDIT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CHOICES_AUDIT_SHEET_NAME);
    sheet.getRange(1, 1, 1, CHOICES_AUDIT_HEADERS.length).setValues([CHOICES_AUDIT_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function choicesAuditLog_(action, fieldKey, choiceId, oldValue, newValue, note) {
  try {
    getChoicesAuditSheet_().appendRow([new Date(), action, fieldKey || '', choiceId || '', oldValue || '', newValue || '', note || '']);
  } catch (err) {
    console.error('Choices audit log failed: ' + err.message);
  }
}

// Seeds current hard-coded choices on first use so the dynamic form shows
// exactly the same options as before migration. Idempotent + lock-guarded.
function ensureFormChoicesSeeded_() {
  const sheet = getFormChoicesSheet_();
  if (sheet.getLastRow() >= 2) return sheet;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return sheet;
  try {
    if (sheet.getLastRow() >= 2) return sheet;
    const now = new Date();
    const rows = [];
    Object.keys(DEFAULT_FORM_CHOICES).forEach(fieldKey => {
      DEFAULT_FORM_CHOICES[fieldKey].forEach((value, i) => {
        rows.push([
          Utilities.getUuid(), fieldKey, value, value, true, (i + 1) * 10, now, now
        ]);
      });
    });
    if (rows.length) sheet.getRange(2, 1, rows.length, FORM_CHOICES_HEADERS.length).setValues(rows);
    clearFormChoicesCache_();
  } finally {
    lock.releaseLock();
  }
  return sheet;
}

function readFormChoiceRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, FORM_CHOICES_HEADERS.length).getValues()
    .map((r, i) => ({
      row: i + 2,
      id: String(r[0]),
      fieldKey: String(r[1]),
      value: String(r[2] == null ? '' : r[2]).trim(),
      label: String(r[3] == null ? '' : r[3]).trim(),
      active: r[4] === true || String(r[4]).toUpperCase() === 'TRUE',
      sort: Number(r[5]) || 0,
      createdAt: r[6],
      updatedAt: r[7]
    }))
    .filter(c => c.id && c.fieldKey);
}

function buildPublicFormChoices_(rows) {
  const fields = {};
  Object.keys(FORM_CHOICE_FIELDS).forEach(k => {
    fields[k] = {
      label: FORM_CHOICE_FIELDS[k].label,
      inputType: FORM_CHOICE_FIELDS[k].inputType,
      locked: !!FORM_CHOICE_FIELDS[k].locked,
      protectedValues: FORM_CHOICE_FIELDS[k].protectedValues || []
    };
  });
  const choices = {};
  Object.keys(FORM_CHOICE_FIELDS).forEach(k => { choices[k] = []; });
  rows.filter(r => r.active && choices[r.fieldKey])
      .sort((a, b) => (a.sort - b.sort) || String(a.createdAt).localeCompare(String(b.createdAt)))
      .forEach(r => choices[r.fieldKey].push({ value: r.value, label: r.label }));
  return { fields: fields, choices: choices };
}

// Public READ — active choices only, cached briefly; invalidated on mutation.
function getPublicFormChoices() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(FORM_CHOICES_CACHE_KEY);
    if (cached) return JSON.parse(cached);
    ensureFormChoicesSeeded_();
    const payload = buildPublicFormChoices_(readFormChoiceRows_(getFormChoicesSheet_()));
    const out = { status: 'success', fields: payload.fields, choices: payload.choices };
    cache.put(FORM_CHOICES_CACHE_KEY, JSON.stringify(out), FORM_CHOICES_CACHE_SECS);
    return out;
  } catch (err) {
    console.error('getPublicFormChoices failed: ' + err.message);
    // Fail-open fallback: seeded defaults keep the form usable.
    const choices = {};
    Object.keys(DEFAULT_FORM_CHOICES).forEach(k => {
      choices[k] = DEFAULT_FORM_CHOICES[k].map(v => ({ value: v, label: v }));
    });
    return { status: 'success', degraded: true, fields: buildPublicFormChoices_([]).fields, choices: choices };
  }
}

// Values-only view used by registration validation.
function getActiveChoiceValueMap_() {
  try {
    const all = getPublicFormChoices();
    const values = {};
    Object.keys(all.choices || {}).forEach(k => {
      values[k] = all.choices[k].map(c => c.value);
    });
    return values;
  } catch (err) {
    console.error('getActiveChoiceValueMap_ failed: ' + err.message);
    const values = {};
    Object.keys(DEFAULT_FORM_CHOICES).forEach(k => { values[k] = DEFAULT_FORM_CHOICES[k].slice(); });
    return values;
  }
}

function clearFormChoicesCache_() {
  try { CacheService.getScriptCache().remove(FORM_CHOICES_CACHE_KEY); } catch (err) {}
}

// Server-side authorization: shared admin key provisioned via
// generateFormChoicesAdminKey(); never shipped in any client file.
// Failed attempts are rate-limited to blunt brute-forcing.
function authorizeAdmin_(presentedKey) {
  const expected = PropertiesService.getScriptProperties().getProperty(FORM_CHOICES_ADMIN_PROP);
  if (!expected) {
    return { ok: false, code: 'not_provisioned', message: 'Administrative access is not provisioned yet.' };
  }
  const cache = CacheService.getScriptCache();
  const fails = parseInt(cache.get('admin_fail') || '0', 10);
  if (fails >= ADMIN_FAIL_MAX) {
    return { ok: false, code: 'locked', message: 'Too many failed attempts. Try again later.' };
  }
  const given = String(presentedKey == null ? '' : presentedKey);
  if (given.length !== expected.length || given !== expected) {
    cache.put('admin_fail', String(fails + 1), ADMIN_FAIL_LOCK_WINDOW_SECS);
    console.warn('authorizeAdmin_: failed attempt (' + (fails + 1) + ')');
    return { ok: false, code: 'unauthorized', message: 'Unauthorized.' };
  }
  cache.remove('admin_fail');
  return { ok: true };
}

function normalizeChoiceValue_(v) {
  return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').toLowerCase();
}

function findChoiceRow_(rows, choiceId) {
  return rows.filter(r => r.id === String(choiceId))[0] || null;
}

function handleFormChoiceMutation_(p) {
  const auth = authorizeAdmin_(p.adminKey);
  if (!auth.ok) {
    choicesAuditLog_('auth_failed:' + p.action, String(p.fieldKey || ''), '', '', '', auth.code);
    return { status: 'error', code: auth.code, message: auth.message };
  }

  // Read-only listing needs no lock.
  if (p.action === 'list_form_choice') return listFormChoicesAdmin_();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { status: 'error', code: 'busy', message: 'System busy, please try again in a moment.' };
  }
  try {
    ensureFormChoicesSeeded_();
    const sheet = getFormChoicesSheet_();
    const rows = readFormChoiceRows_(sheet);

    switch (p.action) {
      case 'add_form_choice':      return addFormChoice_(sheet, rows, p);
      case 'update_form_choice':   return updateFormChoice_(sheet, rows, p);
      case 'toggle_form_choice':   return toggleFormChoice_(sheet, rows, p);
      case 'reorder_form_choice':  return reorderFormChoice_(sheet, rows, p);
      default:
        return { status: 'error', code: 'malformed', message: 'Unknown action.' };
    }
  } finally {
    lock.releaseLock();
  }
}

// Authenticated management listing: ALL rows (active + inactive) with the
// metadata the admin UI needs. Delivered via POST so the key never appears
// in a URL.
function listFormChoicesAdmin_() {
  ensureFormChoicesSeeded_();
  const rows = readFormChoiceRows_(getFormChoicesSheet_());
  const fields = {};
  Object.keys(FORM_CHOICE_FIELDS).forEach(k => {
    fields[k] = {
      label: FORM_CHOICE_FIELDS[k].label,
      inputType: FORM_CHOICE_FIELDS[k].inputType,
      locked: !!FORM_CHOICE_FIELDS[k].locked,
      protectedValues: FORM_CHOICE_FIELDS[k].protectedValues || []
    };
  });
  const allChoices = {};
  Object.keys(FORM_CHOICE_FIELDS).forEach(k => { allChoices[k] = []; });
  rows.sort((a, b) => (a.sort - b.sort) || String(a.createdAt).localeCompare(String(b.createdAt)))
      .forEach(r => {
        if (allChoices[r.fieldKey]) {
          allChoices[r.fieldKey].push({ id: r.id, value: r.value, label: r.label, active: r.active, sort: r.sort });
        }
      });
  return { status: 'success', fields: fields, allChoices: allChoices };
}

function resolveMutationContext_(rows, p) {
  const fieldKey = String(p.fieldKey || '');
  const field = FORM_CHOICE_FIELDS[fieldKey];
  if (!field) return { error: { status: 'error', code: 'validation', message: 'Unknown field.' } };
  return { fieldKey: fieldKey, field: field };
}

function addFormChoice_(sheet, rows, p) {
  const ctx = resolveMutationContext_(rows, p);
  if (ctx.error) return ctx.error;
  const fieldKey = ctx.fieldKey, field = ctx.field;

  if (field.locked) return { status: 'error', code: 'forbidden', message: 'This field is controlled by application logic and cannot be modified.' };

  const label = String(p.displayLabel || '').trim();
  if (!label) return { status: 'error', code: 'validation', message: 'Display label is required.' };

  // Choice Value defaults to the label; both trimmed, internal whitespace collapsed.
  const value = String(p.choiceValue == null || String(p.choiceValue).trim() === '' ? label : p.choiceValue).trim().replace(/\s+/g, ' ').slice(0, MAX_STRING_LENGTH);
  const displayLabel = label.slice(0, MAX_STRING_LENGTH);
  if (!value) return { status: 'error', code: 'validation', message: 'Choice value is required.' };

  const norm = normalizeChoiceValue_(value);
  if (rows.some(r => r.fieldKey === fieldKey && normalizeChoiceValue_(r.value) === norm)) {
    return { status: 'error', code: 'duplicate', message: 'Choice already exists for this field.' };
  }

  const maxSort = rows.filter(r => r.fieldKey === fieldKey).reduce((m, r) => Math.max(m, r.sort), 0);
  const sort = Number.isFinite(parseInt(p.sortOrder, 10)) && p.sortOrder !== '' && p.sortOrder != null
    ? parseInt(p.sortOrder, 10)
    : maxSort + 10;
  const active = p.active === undefined ? true : !!p.active;
  const now = new Date();
  const id = Utilities.getUuid();

  sheet.appendRow([id, fieldKey, value, displayLabel, active, sort, now, now]);
  clearFormChoicesCache_();
  choicesAuditLog_('add', fieldKey, id, '', value, active ? 'active' : 'inactive');
  return { status: 'success', code: 'added', message: 'Choice added successfully.', choiceId: id };
}

function updateFormChoice_(sheet, rows, p) {
  const target = findChoiceRow_(rows, p.choiceId);
  if (!target) return { status: 'error', code: 'not_found', message: 'Choice not found.' };
  const field = FORM_CHOICE_FIELDS[target.fieldKey];
  if (!field) return { status: 'error', code: 'validation', message: 'Unknown field.' };
  if (field.locked) return { status: 'error', code: 'forbidden', message: 'This field is controlled by application logic and cannot be modified.' };

  const newLabel = p.displayLabel === undefined ? target.label : String(p.displayLabel).trim().slice(0, MAX_STRING_LENGTH);
  if (!newLabel) return { status: 'error', code: 'validation', message: 'Display label is required.' };

  let newValue = target.value;
  if (p.choiceValue !== undefined && String(p.choiceValue).trim() !== '') {
    newValue = String(p.choiceValue).trim().replace(/\s+/g, ' ').slice(0, MAX_STRING_LENGTH);
  }

  const protectedVals = field.protectedValues || [];
  if (newValue !== target.value) {
    if (protectedVals.indexOf(target.value) !== -1 || protectedVals.indexOf(newValue) !== -1) {
      return { status: 'error', code: 'forbidden', message: 'Cannot change this protected choice value.' };
    }
    if (p.confirmValueChange !== true) {
      return { status: 'error', code: 'confirm_required', message: 'Changing the Choice Value requires explicit confirmation.' };
    }
    const norm = normalizeChoiceValue_(newValue);
    if (rows.some(r => r.id !== target.id && r.fieldKey === target.fieldKey && normalizeChoiceValue_(r.value) === norm)) {
      return { status: 'error', code: 'duplicate', message: 'Choice already exists for this field.' };
    }
  }

  const newActive = p.active === undefined ? target.active : !!p.active;
  if (target.active && !newActive && protectedVals.indexOf(target.value) !== -1) {
    return { status: 'error', code: 'forbidden', message: 'Cannot deactivate this protected choice.' };
  }
  let newSort = target.sort;
  if (p.sortOrder !== undefined && p.sortOrder !== '' && Number.isFinite(parseInt(p.sortOrder, 10))) {
    newSort = parseInt(p.sortOrder, 10);
  }

  sheet.getRange(target.row, 3, 1, 6).setValues([[newValue, newLabel, newActive, newSort, target.createdAt, new Date()]]);
  clearFormChoicesCache_();
  choicesAuditLog_('update', target.fieldKey, target.id, target.value, newValue, 'label/active/sort updated');
  return { status: 'success', code: 'updated', message: 'Choice updated successfully.' };
}

function toggleFormChoice_(sheet, rows, p) {
  const target = findChoiceRow_(rows, p.choiceId);
  if (!target) return { status: 'error', code: 'not_found', message: 'Choice not found.' };
  const field = FORM_CHOICE_FIELDS[target.fieldKey];
  if (!field) return { status: 'error', code: 'validation', message: 'Unknown field.' };
  if (field.locked) return { status: 'error', code: 'forbidden', message: 'This field is controlled by application logic and cannot be modified.' };

  const protectedVals = field.protectedValues || [];
  const deactivating = target.active; // toggle semantics
  if (deactivating && protectedVals.indexOf(target.value) !== -1) {
    return { status: 'error', code: 'forbidden', message: 'Cannot deactivate this protected choice.' };
  }

  const newActive = !target.active;
  sheet.getRange(target.row, 5).setValue(newActive);
  sheet.getRange(target.row, 8).setValue(new Date());
  clearFormChoicesCache_();
  choicesAuditLog_(newActive ? 'activate' : 'deactivate', target.fieldKey, target.id, target.value, target.value, '');
  return {
    status: 'success',
    code: newActive ? 'activated' : 'deactivated',
    message: newActive ? 'Choice activated successfully.' : 'Choice deactivated successfully.',
    active: newActive
  };
}

function reorderFormChoice_(sheet, rows, p) {
  const direction = String(p.direction || '').toLowerCase();
  if (direction !== 'up' && direction !== 'down') {
    return { status: 'error', code: 'validation', message: 'Direction must be "up" or "down".' };
  }
  const target = findChoiceRow_(rows, p.choiceId);
  if (!target) return { status: 'error', code: 'not_found', message: 'Choice not found.' };
  const field = FORM_CHOICE_FIELDS[target.fieldKey];
  if (!field) return { status: 'error', code: 'validation', message: 'Unknown field.' };
  if (field.locked) return { status: 'error', code: 'forbidden', message: 'This field is controlled by application logic and cannot be modified.' };

  const siblings = rows.filter(r => r.fieldKey === target.fieldKey)
                       .sort((a, b) => (a.sort - b.sort) || String(a.createdAt).localeCompare(String(b.createdAt)));
  const idx = siblings.findIndex(r => r.id === target.id);
  const swapIdx = direction === 'up' ? idx - 1 : idx + 1;
  if (swapIdx < 0 || swapIdx >= siblings.length) {
    return { status: 'success', code: 'noop', message: 'Already at the ' + (direction === 'up' ? 'top' : 'bottom') + '.' };
  }

  const other = siblings[swapIdx];
  const now = new Date();
  sheet.getRange(target.row, 6).setValue(other.sort);
  sheet.getRange(target.row, 8).setValue(now);
  sheet.getRange(other.row, 6).setValue(target.sort);
  sheet.getRange(other.row, 8).setValue(now);
  clearFormChoicesCache_();
  choicesAuditLog_('reorder', target.fieldKey, target.id, String(target.sort), String(other.sort), direction);
  return { status: 'success', code: 'reordered', message: 'Choice moved ' + direction + '.' };
}