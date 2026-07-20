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
    } else if (e.parameter.action === 'hots' && e.parameter.id) {
      result = registerHOTS(e.parameter.id);
    } else if (e.parameter.action === 'stats') {
      result = getStats();
    } else if (e.parameter.action === 'list') {
      result = getCheckedInList();
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
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
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

function registerHOTS(scannedId) {
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
    const resultCol = colIndex(headers, INTERVIEW_RESULT_HEADER);

    const idRange = sheet.getRange(2, idCol, Math.max(sheet.getLastRow() - 1, 1), 1);
    const match = idRange.createTextFinder(scannedId).matchEntireCell(true).findNext();
    if (!match) return { status: 'invalid' };

    const row = match.getRow();
    const fullName = [sheet.getRange(row, firstCol).getValue(), sheet.getRange(row, midCol).getValue(), sheet.getRange(row, lastCol).getValue()].filter(String).join(' ');
    const currentStatus = sheet.getRange(row, statusCol).getValue();

    // NEW: block HOTS registration unless the person has already been checked in
    if (currentStatus !== 'Checked In') {
      return { status: 'not_checked_in', name: fullName };
    }

    const currentResult = sheet.getRange(row, resultCol).getValue();
    if (currentResult === 'HOTS') {
      return { status: 'duplicate', name: fullName, time: 'already marked HOTS' };
    }

    sheet.getRange(row, resultCol).setValue('HOTS');
    return { status: 'success', name: fullName, time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a') };
  } finally {
    lock.releaseLock();
  }
}

function getStats() {
  const sheet = getResponseSheet();
  const headers = getHeaderRow(sheet);
  const statusCol = colIndex(headers, STATUS_HEADER);
  const resultCol = colIndex(headers, INTERVIEW_RESULT_HEADER);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'success', checkedIn: 0, hots: 0 };

  const statusValues = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
  const resultValues = sheet.getRange(2, resultCol, lastRow - 1, 1).getValues();

  let checkedIn = 0, hots = 0;
  for (let i = 0; i < statusValues.length; i++) {
    if (statusValues[i][0] === 'Checked In') checkedIn++;
    if (resultValues[i][0] === 'HOTS') hots++;
  }
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
