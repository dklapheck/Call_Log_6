function saveSccToRoster() {
  return withRosterLock_(function() {
    saveSccEntry_(false);
  });
}

function saveSccAndOpenPowerSchool() {
  return withRosterLock_(function() {
    saveSccEntry_(true);
  });
}

function saveSccEntry_(openPowerSchool) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
    const callSheet = getRequiredSheet_(ss, SCC_CONFIG.sheets.callEntry);
    const rosterSheet = getRequiredSheet_(ss, SCC_CONFIG.sheets.roster);

    const studentId = callSheet.getRange(SCC_CONFIG.callEntry.studentId).getValue();
    if (!studentId) {
      ss.toast('Select a student before saving.', 'SCC Not Saved', 4);
      return;
    }

    const rosterRow = findStudentRosterRow_(rosterSheet, studentId);
    if (!rosterRow) {
      ss.toast('Student ID ' + studentId + ' was not found on the SCC sheet.', 'SCC Not Saved', 5);
      return;
    }

    const studentName = getStudentName_(rosterSheet, rosterRow);
    const contactRow = findQuestionRow_(callSheet, 'Successful contact?');
    const successfulContact = callSheet.getRange(contactRow, 4).getValue() === true;
    const unsuccessfulContact = callSheet.getRange(contactRow, 6).getValue() === true;

    if (successfulContact === unsuccessfulContact) {
      ss.toast('Choose either Yes or No for Successful contact?.', 'SCC Not Saved', 5);
      return;
    }

    const sccNote = callSheet.getRange(SCC_CONFIG.callEntry.note).getDisplayValue().trim();
    if (!sccNote) {
      ss.toast('There is no contact note to save for ' + studentName + '.', 'SCC Not Saved', 4);
      return;
    }

    if (unsuccessfulContact) {
      const result = saveFailedSccAttempt_(ss, rosterSheet, rosterRow, studentName, sccNote);
      if (openPowerSchool && result.saved) sendSccHandoff_(studentId, sccNote);
      return;
    }

    const sccToDo = callSheet.getRange(SCC_CONFIG.callEntry.toDo).getDisplayValue().trim();
    const notesColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.notesHeader);
    const toDoColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.toDoHeader);
    const completionColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.completionHeader);

    const existingNote = rosterSheet.getRange(rosterRow, notesColumn).getDisplayValue().trim();
    if (existingNote === sccNote && rosterSheet.getRange(rosterRow, completionColumn).getDisplayValue() === SCC_CONFIG.roster.completedValue) {
      ss.toast('This SCC is already saved for ' + studentName + '.', 'Duplicate Not Saved', 5);
      if (openPowerSchool) sendSccHandoff_(studentId, sccNote);
      return;
    }

    rosterSheet.getRange(rosterRow, notesColumn).setValue(sccNote);
    rosterSheet.getRange(rosterRow, toDoColumn).setValue(sccToDo);
    rosterSheet.getRange(rosterRow, completionColumn).setValue(SCC_CONFIG.roster.completedValue);
    SpreadsheetApp.flush();

    ss.toast('SCC saved and marked Completed for ' + studentName + '.', 'SCC Saved', 4);
    if (openPowerSchool) sendSccHandoff_(studentId, sccNote);
}

function sendSccHandoff_(studentId, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const encoded = Utilities.base64EncodeWebSafe(
      JSON.stringify({ v: 1, studentNumber: normalizeId_(studentId), note: note }),
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
