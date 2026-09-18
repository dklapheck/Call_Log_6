const ECC_HANDOFF_PREFIX_ = 'ECC_HANDOFF_V1:';

// The four recent-note fields remain visible after logging. History is kept
// separately so a teacher can review old notes while writing the next note.
function setupEccBatchLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRequiredSheet_(ss, APP_CONFIG.sheets.ecc);
  const cols = getEccBatchColumns_(sheet);

  // Append controls after the existing data; never insert beside E:I.
  let lastColumn = sheet.getLastColumn();
  if (!cols.type) {
    ensureSheetColumns_(sheet, ++lastColumn);
    sheet.getRange(1, lastColumn).setValue('ECC Type');
    cols.type = lastColumn;
  }
  if (!cols.ready) {
    ensureSheetColumns_(sheet, ++lastColumn);
    sheet.getRange(1, lastColumn).setValue('Ready to Log');
    cols.ready = lastColumn;
  }

  const rows = sheet.getMaxRows() - 1;
  if (rows > 0) {
    const typeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Conversation', 'Attempt'], true)
      .setAllowInvalid(false)
      .build();
    const readyRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    sheet.getRange(2, cols.type, rows, 1).setDataValidation(typeRule);
    sheet.getRange(2, cols.ready, rows, 1).setDataValidation(readyRule);
    [cols.overall, cols.classes, cols.grades, cols.socially, cols.history]
      .filter(Boolean)
      .forEach(function(col) {
        sheet.getRange(2, col, rows, 1)
          .setWrap(true).setVerticalAlignment('top');
      });
  }

  sheet.getRange(1, cols.ready).setNote(
    'Check each student to archive, then use Teacher Tools > Log Ready ECC Rows. ' +
    'The recent notes stay visible; already archived entries are not added twice.'
  );
  ss.toast('ECC batch controls are ready. Existing notes and columns were preserved.',
    'ECC Batch Layout Ready', 6);
}

function logReadyEccRows() {
  return withRosterLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getRequiredSheet_(ss, APP_CONFIG.sheets.ecc);
    let cols;
    try {
      cols = getEccBatchColumns_(sheet);
      if (!cols.ready) throw new Error(
        'Run Teacher Tools > Set Up / Repair ECC Batch Columns first.'
      );
    } catch (error) {
      logAutomationEvent_('ERROR', 'ECC Batch Log', '', error.message, getErrorDetails_(error));
      ss.toast(error.message + ' See Automation Log.', 'ECC Batch Update', 8);
      return;
    }

    let logged = 0;
    let alreadyArchived = 0;
    let failed = 0;
    for (let row = 2; row <= sheet.getLastRow(); row++) {
      if (sheet.getRange(row, cols.ready).getValue() !== true) continue;
      const studentNumber = normalizeId_(sheet.getRange(row, cols.student).getValue());
      try {
        const result = logEccRow_(sheet, row, cols);
        if (result.duplicate) alreadyArchived++;
        else logged++;
        sheet.getRange(row, cols.ready).setValue(false);
        logAutomationEvent_('INFO', 'ECC Batch Log', result.studentNumber,
          result.duplicate ? 'ECC entry was already archived.' : 'ECC entry archived.',
          'Row ' + row);
      } catch (error) {
        failed++;
        logAutomationEvent_('ERROR', 'ECC Batch Log', studentNumber,
          'ECC row was not logged.', 'Row ' + row + '\n' + getErrorDetails_(error));
      }
    }

    SpreadsheetApp.flush();
    ss.toast(logged + ' archived, ' + alreadyArchived + ' already archived' +
      (failed ? ', ' + failed + ' need review. See Automation Log.' : '.'),
      'ECC Batch Update', 8);
  });
}

function logCurrentEccRowAndOpenPowerSchool() {
  return withRosterLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    if (sheet.getName() !== APP_CONFIG.sheets.ecc) {
      const message = 'Select a student row on the ECC tab first.';
      logAutomationEvent_('ERROR', 'ECC Handoff', '', message, '');
      ss.toast(message + ' See Automation Log.', 'ECC Handoff', 7);
      return;
    }

    const row = sheet.getActiveCell().getRow();
    if (row <= 1) {
      const message = 'Select a student row, not the header row.';
      logAutomationEvent_('ERROR', 'ECC Handoff', '', message, 'Row ' + row);
      ss.toast(message + ' See Automation Log.', 'ECC Handoff', 7);
      return;
    }

    let studentNumber = '';
    try {
      const cols = getEccBatchColumns_(sheet);
      studentNumber = normalizeId_(sheet.getRange(row, cols.student).getValue());
      const result = logEccRow_(sheet, row, cols);
      if (cols.ready) sheet.getRange(row, cols.ready).setValue(false);
      // An already archived entry can be handed off again if a tab failed to open.
      sendEccHandoff_(result);
    } catch (error) {
      logAutomationEvent_('ERROR', 'ECC Handoff', studentNumber,
        'Could not prepare the PowerSchool handoff.',
        'Row ' + row + '\n' + getErrorDetails_(error));
      ss.toast('ECC handoff failed. See Automation Log.', 'ECC Handoff', 8);
    }
  });
}

function normalizeEccHeader_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function getEccBatchColumns_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  function find(names) {
    for (const name of names) {
      const wanted = normalizeEccHeader_(name);
      const matches = [];
      headers.forEach(function(value, index) {
        if (normalizeEccHeader_(value) === wanted) matches.push(index + 1);
      });
      if (matches.length > 1) throw new Error('Duplicate ECC header: ' + name);
      if (matches.length) return matches[0];
    }
    return null;
  }
  function requireAny(names) {
    const column = find(names);
    if (!column) throw new Error('ECC is missing the "' + names[0] + '" column.');
    return column;
  }

  const cols = {
    student: requireAny(['Student Number']),
    date: requireAny(['ECC Date']),
    overall: requireAny(['Recent ECC Notes Overall', 'Recent ECC Notes', 'ECC Check-In']),
    classes: find(['Recent ECC Notes Classes']),
    grades: find(['Recent ECC Notes Grades']),
    socially: find(['Recent ECC Notes Socially']),
    history: requireAny(['Old ECC Dates and Notes', 'ECC History']),
    attempt1: find(['Attempt 1']),
    attempt2: find(['Attempt 2']),
    type: find(['ECC Type']),
    ready: find(['Ready to Log'])
  };
  const categoryColumns = [cols.classes, cols.grades, cols.socially];
  if (categoryColumns.some(Boolean) && categoryColumns.some(function(col) { return !col; })) {
    throw new Error('ECC needs all three Recent ECC Notes columns: Classes, Grades, Socially.');
  }
  return cols;
}

function logEccRow_(sheet, row, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentNumber = normalizeId_(sheet.getRange(row, cols.student).getValue());
  if (!studentNumber) throw new Error('Student Number is blank.');

  const date = sheet.getRange(row, cols.date).getValue();
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    throw new Error('ECC Date must be a valid date.');
  }

  function read(col) {
    return col ? String(sheet.getRange(row, col).getDisplayValue() || '').trim() : '';
  }
  const overall = read(cols.overall);
  const classes = read(cols.classes);
  const grades = read(cols.grades);
  const socially = read(cols.socially);
  if (!overall && !classes && !grades && !socially) {
    throw new Error('Enter a recent ECC note in Overall, Classes, Grades, or Socially.');
  }

  const type = read(cols.type) || 'Conversation';
  if (['Conversation', 'Attempt'].indexOf(type) === -1) {
    throw new Error('ECC Type must be Conversation or Attempt.');
  }

  const dateLabel = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
  const lines = ['-- ' + dateLabel + ' [' + type + ']'];
  if (overall) lines.push('Overall: ' + overall);
  if (classes) lines.push('Classes: ' + classes);
  if (grades) lines.push('Grades: ' + grades);
  if (socially) lines.push('Socially: ' + socially);
  const rendered = lines.join('\n');

  const historyCell = sheet.getRange(row, cols.history);
  const existingHistory = String(historyCell.getValue() || '');
  const duplicate = eccHistoryContains_(existingHistory, rendered);
  // Check attempt capacity before updating history to avoid partial logging.
  const attemptCell = type === 'Attempt'
    ? getEccAttemptTarget_(sheet, row, cols, rendered)
    : null;

  if (!duplicate) {
    historyCell.setValue(appendEccHistory_(existingHistory, rendered))
      .setWrap(true).setVerticalAlignment('top');
  }
  if (attemptCell && String(attemptCell.getValue() || '') !== rendered) {
    attemptCell.setValue(rendered).setWrap(true);
  }

  return {
    studentNumber: studentNumber,
    date: dateLabel,
    note: rendered,
    duplicate: duplicate
  };
}

function getEccAttemptTarget_(sheet, row, cols, rendered) {
  const cells = [cols.attempt1, cols.attempt2]
    .filter(Boolean)
    .map(function(col) { return sheet.getRange(row, col); });
  for (const cell of cells) {
    if (String(cell.getValue() || '') === rendered) return cell;
  }
  for (const cell of cells) {
    if (!String(cell.getValue() || '').trim()) return cell;
  }
  throw new Error('Both ECC attempt columns are full; this attempt was not archived.');
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
  try {
    const encoded = Utilities.base64EncodeWebSafe(
      JSON.stringify({
        v: 1,
        studentNumber: payload.studentNumber,
        date: payload.date,
        note: payload.note
      }),
      Utilities.Charset.UTF_8
    ).replace(/=+$/g, '');
    const marker = ECC_HANDOFF_PREFIX_ + encoded;
    logAutomationEvent_('INFO', 'ECC Handoff', payload.studentNumber,
      'PowerSchool handoff created.',
      'Date: ' + payload.date + '\nHandoff marker:\n' + marker);
    SpreadsheetApp.flush();
    ss.toast(marker, 'ECC Tools', 10);
  } catch (error) {
    logAutomationEvent_('ERROR', 'ECC Handoff',
      payload && payload.studentNumber ? payload.studentNumber : '',
      'Could not create the PowerSchool handoff.', getErrorDetails_(error));
    ss.toast('ECC handoff failed. See Automation Log.', 'ECC Handoff', 8);
  }
}
