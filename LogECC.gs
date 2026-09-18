const ECC_HANDOFF_PREFIX_ = 'ECC_HANDOFF_V1:';

function saveEccEntry() {
  return withRosterLock_(function() {
    const result = saveEccEntry_(false);
    if (result) SpreadsheetApp.getActiveSpreadsheet().toast('ECC entry saved for ' + result.studentNumber + '.', 'ECC Saved', 4);
    return result;
  });
}

function saveEccEntryAndOpenPowerSchool() {
  return withRosterLock_(function() {
    const result = saveEccEntry_(true);
    if (result) sendEccHandoff_(result);
    return result;
  });
}

function saveEccEntry_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = getRequiredSheet_(ss, APP_CONFIG.sheets.eccEntry);
  const ecc = getRequiredSheet_(ss, APP_CONFIG.sheets.ecc);
  const scc = getRequiredSheet_(ss, APP_CONFIG.sheets.scc);
  const studentData = getRequiredSheet_(ss, APP_CONFIG.sheets.studentData);

  const studentNumber = normalizeId_(form.getRange('B2').getValue());
  const date = form.getRange('B3').getValue();
  const result = String(form.getRange('B4').getDisplayValue()).trim();
  const note = String(form.getRange('B5').getValue() || '').trim();
  const todo = String(form.getRange('B6').getValue() || '').trim();
  const oneNote = String(form.getRange('B7').getValue() || '').trim();

  if (!studentNumber || !(date instanceof Date) || ['Conversation','Attempt'].indexOf(result) === -1 || !note) {
    SpreadsheetApp.getUi().alert('Enter a Student Number, date, Result (Conversation or Attempt), and ECC note.');
    return null;
  }

  let row = findStudentRosterRow_(ecc, studentNumber);
  if (!row) row = createEccRow_(ecc, scc, studentData, studentNumber);
  if (!row) throw new Error('Student Number ' + studentNumber + ' was not found in SCC or StudentData.');

  const dateCol = findHeaderColumn_(ecc, 'ECC Date');
  const recentCol = findHeaderColumn_(ecc, 'Recent ECC Notes');
  const oldCol = findHeaderColumn_(ecc, 'Old ECC Dates and Notes');
  const attempt1Col = findHeaderColumn_(ecc, 'Attempt 1');
  const attempt2Col = findHeaderColumn_(ecc, 'Attempt 2');
  const linkCol = findOptionalHeaderColumn_(ecc, 'OneNote Link');
  const dateLabel = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
  const text = note + (todo ? '\nTo do: ' + todo : '');
  const rendered = '-- ' + dateLabel + ': ' + text;

  if (result === 'Conversation') {
    const oldRecent = String(ecc.getRange(row, recentCol).getValue() || '').trim();
    const history = String(ecc.getRange(row, oldCol).getValue() || '');
    if (oldRecent === text || history.indexOf(rendered) !== -1) throw new Error('This ECC conversation appears to have been saved already.');

    if (oldRecent) {
      const previousDate = ecc.getRange(row, dateCol).getValue();
      const previousLabel = previousDate instanceof Date
        ? Utilities.formatDate(previousDate, ss.getSpreadsheetTimeZone(), 'M/d/yyyy')
        : 'Earlier';
      ecc.getRange(row, oldCol).setValue(appendHistoryEntry_(history, '-- ' + previousLabel + ': ' + oldRecent));
    }
    ecc.getRange(row, dateCol).setValue(date);
    ecc.getRange(row, recentCol).setValue(text);
  } else {
    const first = ecc.getRange(row, attempt1Col);
    const second = ecc.getRange(row, attempt2Col);
    const history = String(ecc.getRange(row, oldCol).getValue() || '');
    if (String(first.getValue() || '').indexOf(rendered) !== -1 ||
        String(second.getValue() || '').indexOf(rendered) !== -1 ||
        history.indexOf(rendered) !== -1) {
      throw new Error('This ECC attempt appears to have been saved already.');
    }
    if (!first.getValue()) first.setValue(rendered);
    else if (!second.getValue()) second.setValue(rendered);
    else ecc.getRange(row, oldCol).setValue(appendHistoryEntry_(history, rendered));
  }

  if (linkCol && oneNote) ecc.getRange(row, linkCol).setValue(oneNote);
  SpreadsheetApp.flush();
  return { studentNumber: studentNumber, date: dateLabel, note: text };
}

function createEccRow_(ecc, scc, studentData, studentNumber) {
  const sccRow = findStudentRosterRow_(scc, studentNumber);
  const dataRow = findStudentRosterRow_(studentData, studentNumber);
  if (!sccRow && !dataRow) return null;

  const row = ecc.getLastRow() + 1;
  ensureSheetRows_(ecc, row);
  if (row > 2) {
    ecc.getRange(row - 1, 1, 1, ecc.getLastColumn())
      .copyTo(ecc.getRange(row, 1, 1, ecc.getLastColumn()), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  }

  const eh = getHeaderIndexMap_(ecc.getRange(1,1,1,ecc.getLastColumn()).getDisplayValues()[0]);
  const sh = getHeaderIndexMap_(scc.getRange(1,1,1,scc.getLastColumn()).getDisplayValues()[0]);
  const dh = getHeaderIndexMap_(studentData.getRange(1,1,1,studentData.getLastColumn()).getDisplayValues()[0]);

  function sourceValue(sccHeader, dataHeader) {
    if (sccRow && typeof sh[sccHeader] !== 'undefined') return scc.getRange(sccRow, sh[sccHeader] + 1).getValue();
    if (dataRow && dataHeader && typeof dh[dataHeader] !== 'undefined') return studentData.getRange(dataRow, dh[dataHeader] + 1).getValue();
    return '';
  }

  const values = {
    'Prefered Name': sourceValue('Prefered Name','Preferred Display Name'),
    'Student Number': Number(studentNumber) || studentNumber,
    'SCC Notes': sourceValue('SCC Notes',''),
    'Small Group': sourceValue('Small Group',''),
    'Student Email': sourceValue('Student Email','Effective Student Email'),
    'LAST NAME': sourceValue('LAST NAME','Source Last Name'),
    'FIRST NAME': sourceValue('FIRST NAME','Source First Name'),
    'GRADE': sourceValue('GRADE','Grade'),
    'START DATE': sourceValue('START DATE','Enroll Date'),
    'SCHOOL': sourceValue('SCHOOL','School'),
    'SPED': sourceValue('SPED','SPED/504')
  };

  Object.keys(values).forEach(function(header) {
    if (typeof eh[header] !== 'undefined') ecc.getRange(row, eh[header] + 1).setValue(values[header]);
  });
  if (typeof eh['Name'] !== 'undefined') ecc.getRange(row, eh['Name'] + 1).setFormula('=L' + row + '&" "&K' + row);
  if (typeof eh['Subject Line'] !== 'undefined') ecc.getRange(row, eh['Subject Line'] + 1).setFormula('=LEFT(L' + row + ',1)&" "&K' + row + '&", "&B' + row + '&", "&O' + row);
  return row;
}

function archiveAndOpenPowerSchoolECC() {
  const data = getCurrentECCRowData_();
  if (!data) return;
  archiveECCData_(data, false);
  sendEccHandoff_({ studentNumber: data.studentNumber, date: data.date, note: data.note });
}

function archiveCurrentECCNote() {
  const data = getCurrentECCRowData_();
  if (data) archiveECCData_(data, true);
}

function getCurrentECCRowData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== APP_CONFIG.sheets.ecc) {
    SpreadsheetApp.getUi().alert('Please select a student row on the ECC tab first.');
    return null;
  }
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please select a student row, not the header row.');
    return null;
  }

  const studentCol = findHeaderColumn_(sheet, 'Student Number');
  const dateCol = findHeaderColumn_(sheet, 'ECC Date');
  const noteCol = findHeaderColumn_(sheet, 'Recent ECC Notes');
  const oldCol = findHeaderColumn_(sheet, 'Old ECC Dates and Notes');

  const studentNumber = normalizeId_(sheet.getRange(row, studentCol).getValue());
  const date = sheet.getRange(row, dateCol).getDisplayValue().trim();
  const note = sheet.getRange(row, noteCol).getDisplayValue().trim();
  if (!studentNumber || !date || !note) {
    SpreadsheetApp.getUi().alert('The selected ECC row needs a Student Number, ECC Date, and Recent ECC Notes.');
    return null;
  }

  return {
    ss: ss,
    studentNumber: studentNumber,
    date: date,
    note: note,
    oldText: sheet.getRange(row, oldCol).getDisplayValue(),
    oldCell: sheet.getRange(row, oldCol)
  };
}

function archiveECCData_(data, showToast) {
  const entry = '-- ' + data.date + ': ' + data.note;
  const updated = appendHistoryEntry_(data.oldText, entry);
  if (updated === String(data.oldText || '').trim()) {
    if (showToast) data.ss.toast('This ECC date/note is already in the archive.', 'ECC Tools', 4);
    return true;
  }
  data.oldCell.setValue(updated).setWrap(true);
  if (showToast) data.ss.toast('ECC note archived for student ' + data.studentNumber + '.', 'ECC Tools', 4);
  return true;
}

function sendEccHandoff_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const encoded = Utilities.base64EncodeWebSafe(JSON.stringify({
    v: 1,
    studentNumber: payload.studentNumber,
    date: payload.date,
    note: payload.note
  }), Utilities.Charset.UTF_8).replace(/=+$/g, '');
  SpreadsheetApp.flush();
  ss.toast(ECC_HANDOFF_PREFIX_ + encoded, 'ECC Tools', 10);
}
