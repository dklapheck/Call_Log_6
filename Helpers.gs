function getRequiredSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Required sheet not found: ' + sheetName);
  return sheet;
}

function normalizeId_(value) {
  if (value === null || typeof value === 'undefined' || value === '') return '';
  if (typeof value === 'number' && isFinite(value)) {
    return Math.floor(value) === value ? String(value) : String(value).trim();
  }
  return String(value).trim().replace(/\.0$/, '');
}

function findHeaderColumn_(sheet, headerText) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) throw new Error('No columns were found on sheet: ' + sheet.getName());
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const matches = [];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === headerText) matches.push(i + 1);
  }
  if (matches.length === 0) {
    throw new Error('Could not find the header "' + headerText + '" on ' + sheet.getName() + '.');
  }
  if (matches.length > 1) {
    throw new Error('More than one column is labeled "' + headerText + '" on ' + sheet.getName() + '.');
  }
  return matches[0];
}

function findOptionalHeaderColumn_(sheet, headerText) {
  try { return findHeaderColumn_(sheet, headerText); }
  catch (error) { return null; }
}

function getHeaderIndexMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    const key = String(header || '').trim();
    if (key && typeof map[key] === 'undefined') map[key] = index;
  });
  return map;
}

function readSheetTable_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { headers: [], headerMap: {}, rows: [] };
  return { headers: values[0], headerMap: getHeaderIndexMap_(values[0]), rows: values.slice(1) };
}

function getSourceSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getRequiredSheet_(ss, APP_CONFIG.sheets.settings);
  const url = String(settings.getRange(APP_CONFIG.source.urlCell).getDisplayValue()).trim();
  if (!/^https:\/\/docs\.google\.com\/spreadsheets\//i.test(url)) {
    throw new Error('Paste the Google Sheets URL for 6Student Data Imports into Instructions and Settings!' + APP_CONFIG.source.urlCell + '.');
  }
  return SpreadsheetApp.openByUrl(url);
}

function withRosterLock_(work) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try { return work(); }
  finally { lock.releaseLock(); }
}

function ensureSheetRows_(sheet, requiredLastRow) {
  if (requiredLastRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredLastRow - sheet.getMaxRows());
  }
}

function ensureSheetColumns_(sheet, requiredLastColumn) {
  if (requiredLastColumn > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredLastColumn - sheet.getMaxColumns());
  }
}

function findStudentRosterRow_(sheet, studentId) {
  const idColumn = findHeaderColumn_(sheet, 'Student Number');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const ids = sheet.getRange(2, idColumn, lastRow - 1, 1).getDisplayValues();
  const wanted = normalizeId_(studentId);
  for (let i = 0; i < ids.length; i++) {
    if (normalizeId_(ids[i][0]) === wanted) return i + 2;
  }
  return null;
}

function getStudentName_(rosterSheet, rosterRow) {
  const firstNameColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.firstNameHeader);
  const lastNameColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.lastNameHeader);
  const firstName = rosterSheet.getRange(rosterRow, firstNameColumn).getDisplayValue().trim();
  const lastName = rosterSheet.getRange(rosterRow, lastNameColumn).getDisplayValue().trim();
  return (firstName + ' ' + lastName).trim();
}

function findQuestionRow_(sheet, questionText) {
  const match = sheet.getRange('B7:B50').createTextFinder(questionText).matchEntireCell(true).findNext();
  if (!match) throw new Error('Could not find this Call Entry question:\n' + questionText);
  return match.getRow();
}

function saveFailedSccAttempt_(ss, rosterSheet, rosterRow, studentName, contactNote) {
  const attemptHeaders = ['Attempt 1', 'Attempt 2', 'Attempt 3', 'Attempt 4', 'Attempt 5'];
  const statuses = ['First Attempt', 'Second Attempt', 'Third Attempt', 'Fourth Attempt', 'Fifth Attempt'];

  for (let i = 0; i < attemptHeaders.length; i++) {
    const col = findHeaderColumn_(rosterSheet, attemptHeaders[i]);
    if (rosterSheet.getRange(rosterRow, col).getDisplayValue().trim() === contactNote) {
      ss.toast('This contact attempt is already saved for ' + studentName + '.', 'Duplicate Not Saved', 5);
      return { saved: true, duplicate: true };
    }
  }

  let target = null;
  let attempt = null;
  for (let i = 0; i < attemptHeaders.length; i++) {
    const col = findHeaderColumn_(rosterSheet, attemptHeaders[i]);
    if (!rosterSheet.getRange(rosterRow, col).getDisplayValue().trim()) {
      target = col; attempt = i; break;
    }
  }
  if (target === null) {
    ss.toast('No empty Attempt 1–5 column is available for ' + studentName + '.', 'Attempt Not Saved', 8);
    return { saved: false };
  }

  rosterSheet.getRange(rosterRow, target).setValue(contactNote);
  rosterSheet.getRange(rosterRow, findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.completionHeader))
    .setValue(statuses[attempt]);
  SpreadsheetApp.flush();
  ss.toast('Attempt ' + (attempt + 1) + ' saved for ' + studentName + '.', 'Contact Attempt Saved', 4);
  return { saved: true, duplicate: false };
}

function appendHistoryEntry_(existingText, entry) {
  const oldText = String(existingText || '').trim();
  if (!oldText) return entry;
  if (oldText.split('\n').map(function(v) { return v.trim(); }).indexOf(entry.trim()) !== -1) return oldText;
  return oldText.replace(/\s+$/g, '') + '\n' + entry;
}

function clearRowsBelowHeader_(sheet) {
  const rows = Math.max(sheet.getLastRow() - 1, 1);
  sheet.getRange(2, 1, rows, sheet.getLastColumn()).clearContent();
}

function writeRowsBelowHeader_(sheet, rows) {
  clearRowsBelowHeader_(sheet);
  if (!rows.length) return;
  ensureSheetRows_(sheet, rows.length + 1);
  ensureSheetColumns_(sheet, rows[0].length);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
