function saveSccToRoster() {
  return withRosterLock_(function() {
    saveSccEntry_();
  });
}

function logSccInPowerSchool() {
  const entry = getSccPowerSchoolEntry_('SCC Not Logged');
  if (!entry) return;

  sendSccHandoff_(entry.studentId, entry.sccNote, entry.workflow);
}

function getSccPowerSchoolEntry_(failureTitle) {
  const activeRange = typeof SpreadsheetApp.getActiveRange === 'function'
    ? SpreadsheetApp.getActiveRange()
    : null;
  if (activeRange && activeRange.getSheet().getName() === SCC_CONFIG.sheets.roster) {
    return getSelectedSccRosterEntry_(activeRange, failureTitle);
  }
  return getSccCallEntryData_(failureTitle);
}

function getSelectedSccRosterEntry_(activeRange, failureTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = activeRange.getSheet();
  const headerRow = SCC_CONFIG.roster.headerRow || 1;

  if (activeRange.getNumRows() !== 1 || activeRange.getNumColumns() !== 1 ||
      activeRange.getRow() <= headerRow) {
    ss.toast('Select one SCC Notes or Attempt 1–5 cell for a student.', failureTitle, 6);
    return null;
  }

  const selectedHeader = rosterSheet.getRange(headerRow, activeRange.getColumn())
    .getDisplayValue().trim();
  const attemptHeaders = SCC_CONFIG.roster.attemptHeaders ||
    ['Attempt 1', 'Attempt 2', 'Attempt 3', 'Attempt 4', 'Attempt 5'];
  const isSuccess = selectedHeader === SCC_CONFIG.roster.notesHeader;
  const isAttempt = attemptHeaders.indexOf(selectedHeader) !== -1;

  if (!isSuccess && !isAttempt) {
    ss.toast('Select a cell under SCC Notes or Attempt 1–5.', failureTitle, 6);
    return null;
  }

  const note = activeRange.getDisplayValue().trim();
  if (!note) {
    ss.toast(selectedHeader + ' is blank for this student.', failureTitle, 5);
    return null;
  }

  const studentIdColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.studentIdHeader);
  const studentId = rosterSheet.getRange(activeRange.getRow(), studentIdColumn).getValue();
  if (!normalizeId_(studentId)) {
    ss.toast('The selected SCC row has no Student Number.', failureTitle, 5);
    return null;
  }

  return {
    ss: ss,
    studentId: studentId,
    sccNote: note,
    successfulContact: isSuccess,
    workflow: isSuccess ? 'SCC Success' : 'SCC Attempt',
    sourceHeader: selectedHeader
  };
}

// Backward-compatible name for any existing drawing/button assignment.
// This action now logs only; it no longer saves to the SCC roster.
function saveSccAndOpenPowerSchool() {
  return logSccInPowerSchool();
}

function getSccCallEntryData_(failureTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const callSheet = getRequiredSheet_(ss, SCC_CONFIG.sheets.callEntry);
  const studentId = callSheet.getRange(SCC_CONFIG.callEntry.studentId).getValue();

  if (!studentId) {
    ss.toast('Select a student first.', failureTitle, 4);
    return null;
  }

  const contactRow = findQuestionRow_(callSheet, 'Successful contact?');
  const successfulContact = callSheet.getRange(contactRow, 4).getValue() === true;
  const unsuccessfulContact = callSheet.getRange(contactRow, 6).getValue() === true;

  if (successfulContact === unsuccessfulContact) {
    ss.toast('Choose either Yes or No for Successful contact?.', failureTitle, 5);
    return null;
  }

  const sccNote = callSheet.getRange(SCC_CONFIG.callEntry.note).getDisplayValue().trim();
  if (!sccNote) {
    ss.toast('There is no contact note for the selected student.', failureTitle, 4);
    return null;
  }

  return {
    ss: ss,
    studentId: studentId,
    sccNote: sccNote,
    sccToDo: callSheet.getRange(SCC_CONFIG.callEntry.toDo).getDisplayValue().trim(),
    successfulContact: successfulContact,
    workflow: successfulContact ? 'SCC Success' : 'SCC Attempt'
  };
}

function saveSccEntry_() {
  const entry = getSccCallEntryData_('SCC Not Saved');
  if (!entry) return;

  const rosterSheet = getRequiredSheet_(entry.ss, SCC_CONFIG.sheets.roster);
  const rosterRow = findStudentRosterRow_(rosterSheet, entry.studentId);
  if (!rosterRow) {
    entry.ss.toast('Student ID ' + entry.studentId + ' was not found on the SCC sheet.', 'SCC Not Saved', 5);
    return;
  }

  const studentName = getStudentName_(rosterSheet, rosterRow);
  if (!entry.successfulContact) {
    saveFailedSccAttempt_(entry.ss, rosterSheet, rosterRow, studentName, entry.sccNote);
    return;
  }

  const notesColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.notesHeader);
  const toDoColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.toDoHeader);
  const completionColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.completionHeader);

  const existingNote = rosterSheet.getRange(rosterRow, notesColumn).getDisplayValue().trim();
  if (existingNote === entry.sccNote &&
      rosterSheet.getRange(rosterRow, completionColumn).getDisplayValue() === SCC_CONFIG.roster.completedValue) {
    entry.ss.toast('This SCC is already saved for ' + studentName + '.', 'Duplicate Not Saved', 5);
    return;
  }

  rosterSheet.getRange(rosterRow, notesColumn).setValue(entry.sccNote);
  rosterSheet.getRange(rosterRow, toDoColumn).setValue(entry.sccToDo);
  rosterSheet.getRange(rosterRow, completionColumn).setValue(SCC_CONFIG.roster.completedValue);
  SpreadsheetApp.flush();

  entry.ss.toast('SCC saved and marked Completed for ' + studentName + '.', 'SCC Saved', 4);
}

function sendSccHandoff_(studentId, note, workflow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const settings = getPowerSchoolContactSettings_(ss, workflow);
    const encoded = Utilities.base64EncodeWebSafe(
      JSON.stringify({ v: 1, studentNumber: normalizeId_(studentId), note: note,
        outcome: workflow, settings: settings }),
      Utilities.Charset.UTF_8
    ).replace(/=+$/g, '');
    const marker = 'SCC_HANDOFF_V1:' + encoded;
    logAutomationEvent_('INFO', 'SCC Handoff', studentId,
      'PowerSchool handoff created.', 'Handoff marker:\n' + marker);
    SpreadsheetApp.flush();
    ss.toast(marker, 'SCC Tools', 10);
  } catch (error) {
    logAutomationEvent_('ERROR', 'SCC Handoff', studentId,
      'Could not create the PowerSchool handoff.', getErrorDetails_(error));
    ss.toast('SCC handoff failed. See Automation Log.', 'SCC Tools', 8);
  }
}
