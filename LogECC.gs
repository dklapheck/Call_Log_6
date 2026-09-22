// The current-row workflow keeps recent notes visible and archives only the selected student's entry.
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
    outcome: type,
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
    const workflow = 'ECC ' + payload.outcome;
    const settings = getPowerSchoolContactSettings_(ss, workflow);
    const requestId = Utilities.getUuid();
    showPowerSchoolHandoffDialog_('ecc', {
      v: 1,
      requestId: requestId,
      studentNumber: payload.studentNumber,
      date: payload.date,
      note: payload.note,
      outcome: workflow,
      settings: settings
    }, 'Open PowerSchool ECC Log');
    logAutomationEvent_('INFO', 'ECC Handoff', '',
      'PowerSchool handoff prepared.',
      'Request ID: ' + requestId);
    SpreadsheetApp.flush();
  } catch (error) {
    logAutomationEvent_('ERROR', 'ECC Handoff', '',
      'Could not create the PowerSchool handoff.', 'Dialog creation failed.');
    ss.toast('ECC handoff failed. See Automation Log.', 'ECC Handoff', 8);
  }
}

// Keep the original menu item and any assigned sheet button working.
function archiveAndOpenPowerSchoolECC() {
  return logCurrentEccRowAndOpenPowerSchool();
}

// The old archive-only menu item must not launch PowerSchool.
function archiveCurrentECCNote() {
  return withRosterLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    if (sheet.getName() !== APP_CONFIG.sheets.ecc || sheet.getActiveCell().getRow() <= 1) {
      ss.toast('Select a student row on the ECC tab first.', 'ECC Not Archived', 6);
      return;
    }

    const row = sheet.getActiveCell().getRow();
    let studentNumber = '';
    try {
      const cols = getEccBatchColumns_(sheet);
      studentNumber = normalizeId_(sheet.getRange(row, cols.student).getValue());
      const result = logEccRow_(sheet, row, cols);
      if (cols.ready) sheet.getRange(row, cols.ready).setValue(false);
      logAutomationEvent_('INFO', 'ECC Archive', result.studentNumber,
        result.duplicate ? 'ECC entry was already archived.' : 'ECC entry archived.',
        'Row ' + row);
      ss.toast(result.duplicate ? 'ECC entry was already archived.' : 'ECC note archived.',
        'ECC Tools', 5);
    } catch (error) {
      logAutomationEvent_('ERROR', 'ECC Archive', studentNumber,
        'Could not archive the ECC note.', 'Row ' + row + '\n' + getErrorDetails_(error));
      ss.toast('ECC archive failed. See Automation Log.', 'ECC Not Archived', 8);
    }
  });
}
