function openStudentDemographics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeRange = typeof SpreadsheetApp.getActiveRange === 'function'
    ? SpreadsheetApp.getActiveRange()
    : null;

  if (!activeRange) {
    ss.toast('Select a student first.', 'PowerSchool Demographics', 5);
    return;
  }

  let studentId = '';
  if (activeRange.getSheet().getName() === SCC_CONFIG.sheets.callEntry) {
    const callSheet = getRequiredSheet_(ss, SCC_CONFIG.sheets.callEntry);
    studentId = normalizeId_(
      callSheet.getRange(SCC_CONFIG.callEntry.studentId).getValue()
    );
  } else {
    if (activeRange.getNumRows() !== 1 || activeRange.getNumColumns() !== 1) {
      ss.toast('Select one cell containing a Student Number.', 'PowerSchool Demographics', 6);
      return;
    }
    studentId = normalizeId_(activeRange.getValue());
    if (!/^\d{5,12}$/.test(studentId)) {
      ss.toast('The selected cell does not contain a valid Student Number.',
        'PowerSchool Demographics', 6);
      return;
    }
  }

  if (!studentId) {
    ss.toast('Select a student on Call Entry before opening Demographics.',
      'PowerSchool Demographics', 5);
    return;
  }

  sendDemographicsHandoff_(ss, studentId);
}

function sendDemographicsHandoff_(ss, studentId) {
  try {
    const encoded = Utilities.base64EncodeWebSafe(
      JSON.stringify({
        v: 1,
        requestId: Utilities.getUuid(),
        studentNumber: studentId
      }),
      Utilities.Charset.UTF_8
    ).replace(/=+$/g, '');
    const marker = 'DEMOGRAPHICS_HANDOFF_V1:' + encoded;
    logAutomationEvent_('INFO', 'Demographics Handoff', studentId,
      'PowerSchool demographics handoff created.', 'Handoff marker:\n' + marker);
    SpreadsheetApp.flush();
    ss.toast(marker, 'PowerSchool Demographics', 10);
  } catch (error) {
    logAutomationEvent_('ERROR', 'Demographics Handoff', studentId,
      'Could not create the PowerSchool demographics handoff.', getErrorDetails_(error));
    ss.toast('Demographics handoff failed. See Automation Log.', 'PowerSchool Demographics', 8);
  }
}
