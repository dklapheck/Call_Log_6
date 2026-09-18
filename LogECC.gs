const ECC_HANDOFF_PREFIX_ = 'ECC_HANDOFF_V1:';

function setupEccBatchLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRequiredSheet_(ss, APP_CONFIG.sheets.ecc);

  let checkInCol = findOptionalHeaderColumn_(sheet, 'ECC Check-In');
  const legacyRecentCol = findOptionalHeaderColumn_(sheet, 'Recent ECC Notes');
  if (!checkInCol && legacyRecentCol) {
    sheet.getRange(1, legacyRecentCol).setValue('ECC Check-In');
    checkInCol = legacyRecentCol;
  }

  let historyCol = findOptionalHeaderColumn_(sheet, 'ECC History');
  const legacyHistoryCol = findOptionalHeaderColumn_(sheet, 'Old ECC Dates and Notes');
  if (!historyCol && legacyHistoryCol) {
    sheet.getRange(1, legacyHistoryCol).setValue('ECC History');
    historyCol = legacyHistoryCol;
  }

  if (!checkInCol) {
    checkInCol = Math.min(5, sheet.getLastColumn() + 1);
    ensureSheetColumns_(sheet, checkInCol);
    sheet.getRange(1, checkInCol).setValue('ECC Check-In');
  }

  if (!historyCol) {
    historyCol = checkInCol + 1;
    sheet.insertColumnAfter(checkInCol);
    sheet.getRange(1, historyCol).setValue('ECC History');
  }

  const desired = [
    'ECC Academics',
    'ECC Engagement',
    'ECC Follow-Up / To Do',
    'ECC Type',
    'Ready to Log'
  ];

  const missing = desired.filter(function(header) {
    return findOptionalHeaderColumn_(sheet, header) === null;
  });

  if (missing.length) {
    sheet.insertColumnsAfter(historyCol, missing.length);
    missing.forEach(function(header, i) {
      sheet.getRange(1, historyCol + 1 + i).setValue(header);
    });
  }

  const cols = getEccBatchColumns_(sheet);
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Conversation', 'Attempt'], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, cols.type, maxRows, 1).setDataValidation(typeRule);
  sheet.getRange(2, cols.ready, maxRows, 1).insertCheckboxes();

  [cols.checkIn, cols.history, cols.academics, cols.engagement, cols.followUp].forEach(function(col) {
    sheet.getRange(2, col, maxRows, 1).setWrap(true).setVerticalAlignment('top');
  });

  sheet.getRange(1, cols.ready).setNote(
    'Check this box for every student you want to log, then use Teacher Tools > Log Ready ECC Rows.'
  );

  ss.toast(
    'ECC is ready for batch entry. Column F/history formatting was preserved.',
    'ECC Batch Layout Ready',
    6
  );
}

function logReadyEccRows() {
  return withRosterLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getRequiredSheet_(ss, APP_CONFIG.sheets.ecc);
    const cols = getEccBatchColumns_(sheet);
    const lastRow = sheet.getLastRow();

    let logged = 0;
    const problems = [];

    for (let row = 2; row <= lastRow; row++) {
      if (sheet.getRange(row, cols.ready).getValue() !== true) continue;

      try {
        const result = logEccRow_(sheet, row, cols, true);
        if (result) logged++;
      } catch (error) {
        problems.push('Row ' + row + ': ' + error.message);
      }
    }

    SpreadsheetApp.flush();

    let message = logged + ' ECC row' + (logged === 1 ? '' : 's') + ' logged.';
    if (problems.length) {
      message += '\n\nNot logged:\n' + problems.slice(0, 8).join('\n');
      if (problems.length > 8) message += '\n…and ' + (problems.length - 8) + ' more.';
    }

    SpreadsheetApp.getUi().alert('ECC Batch Update', message, SpreadsheetApp.getUi().ButtonSet.OK);
  });
}

function logCurrentEccRowAndOpenPowerSchool() {
  return withRosterLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    if (sheet.getName() !== APP_CONFIG.sheets.ecc) {
      SpreadsheetApp.getUi().alert('Select the student row on the ECC tab first.');
      return;
    }

    const row = sheet.getActiveCell().getRow();
    if (row <= 1) {
      SpreadsheetApp.getUi().alert('Select a student row, not the header row.');
      return;
    }

    const cols = getEccBatchColumns_(sheet);
    const result = logEccRow_(sheet, row, cols, true);
    if (result) sendEccHandoff_(result);
  });
}

function getEccBatchColumns_(sheet) {
  function requiredAny(names) {
    for (let i = 0; i < names.length; i++) {
      const col = findOptionalHeaderColumn_(sheet, names[i]);
      if (col !== null) return col;
    }
    throw new Error(
      'ECC is missing "' + names[0] +
      '". Run Teacher Tools > Set Up / Repair ECC Batch Columns.'
    );
  }

  return {
    student: requiredAny(['Student Number']),
    date: requiredAny(['ECC Date']),
    checkIn: requiredAny(['ECC Check-In', 'Recent ECC Notes']),
    history: requiredAny(['ECC History', 'Old ECC Dates and Notes']),
    academics: requiredAny(['ECC Academics']),
    engagement: requiredAny(['ECC Engagement']),
    followUp: requiredAny(['ECC Follow-Up / To Do']),
    type: requiredAny(['ECC Type']),
    ready: requiredAny(['Ready to Log']),
    attempt1: findOptionalHeaderColumn_(sheet, 'Attempt 1'),
    attempt2: findOptionalHeaderColumn_(sheet, 'Attempt 2')
  };
}

function logEccRow_(sheet, row, cols, clearAfter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentNumber = normalizeId_(sheet.getRange(row, cols.student).getValue());

  if (!studentNumber) throw new Error('Student Number is blank.');

  const date = sheet.getRange(row, cols.date).getValue();
  if (!(date instanceof Date)) throw new Error('ECC Date is required.');

  const checkIn = String(sheet.getRange(row, cols.checkIn).getValue() || '').trim();
  const academics = String(sheet.getRange(row, cols.academics).getValue() || '').trim();
  const engagement = String(sheet.getRange(row, cols.engagement).getValue() || '').trim();
  const followUp = String(sheet.getRange(row, cols.followUp).getValue() || '').trim();
  const typeValue = String(sheet.getRange(row, cols.type).getDisplayValue() || '').trim();
  const type = typeValue || 'Conversation';

  if (['Conversation', 'Attempt'].indexOf(type) === -1) {
    throw new Error('ECC Type must be Conversation or Attempt.');
  }

  if (!checkIn && !academics && !engagement && !followUp) {
    throw new Error('Enter at least one ECC note field.');
  }

  const dateLabel = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
  const lines = ['-- ' + dateLabel + ' [' + type + ']'];

  if (checkIn) lines.push('Check-in: ' + checkIn);
  if (academics) lines.push('Academics: ' + academics);
  if (engagement) lines.push('Engagement: ' + engagement);
  if (followUp) lines.push('Follow-up / To do: ' + followUp);

  const rendered = lines.join('\n');
  const historyCell = sheet.getRange(row, cols.history);
  const existingHistory = String(historyCell.getValue() || '');

  if (eccHistoryContains_(existingHistory, rendered)) {
    throw new Error('This exact ECC entry is already in the history.');
  }

  historyCell
    .setValue(appendEccHistory_(existingHistory, rendered))
    .setWrap(true)
    .setVerticalAlignment('top');

  if (type === 'Attempt') {
    saveEccAttemptSummary_(sheet, row, cols, rendered);
  }

  if (clearAfter) {
    sheet.getRange(row, cols.date).clearContent();
    sheet.getRange(row, cols.checkIn).clearContent();
    sheet.getRange(row, cols.academics).clearContent();
    sheet.getRange(row, cols.engagement).clearContent();
    sheet.getRange(row, cols.followUp).clearContent();
    sheet.getRange(row, cols.type).clearContent();
    sheet.getRange(row, cols.ready).setValue(false);
  }

  return {
    studentNumber: studentNumber,
    date: dateLabel,
    note: rendered
  };
}

function saveEccAttemptSummary_(sheet, row, cols, rendered) {
  if (cols.attempt1) {
    const first = sheet.getRange(row, cols.attempt1);
    if (!first.getValue()) {
      first.setValue(rendered).setWrap(true);
      return;
    }
    if (String(first.getValue()) === rendered) return;
  }

  if (cols.attempt2) {
    const second = sheet.getRange(row, cols.attempt2);
    if (!second.getValue()) {
      second.setValue(rendered).setWrap(true);
      return;
    }
  }
}

function appendEccHistory_(existingText, entry) {
  const oldText = String(existingText || '').replace(/\s+$/g, '');
  return oldText ? oldText + '\n\n' + entry : entry;
}

function eccHistoryContains_(history, entry) {
  return String(history || '').indexOf(entry) !== -1;
}

function sendEccHandoff_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const encoded = Utilities.base64EncodeWebSafe(
    JSON.stringify({
      v: 1,
      studentNumber: payload.studentNumber,
      date: payload.date,
      note: payload.note
    }),
    Utilities.Charset.UTF_8
  ).replace(/=+$/g, '');

  SpreadsheetApp.flush();
  ss.toast(ECC_HANDOFF_PREFIX_ + encoded, 'ECC Tools', 10);
}
