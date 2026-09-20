function onSelectionChange(e) {
  const range = e && e.range;
  if (!range || range.getSheet().getName() !== SCC_CONFIG.sheets.callEntry ||
      range.getA1Notation() !== SCC_CONFIG.callEntry.demographicsAction) {
    return;
  }

  openSelectedStudentDemographics_();
}

function openSelectedStudentDemographics_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const callSheet = getRequiredSheet_(ss, SCC_CONFIG.sheets.callEntry);
  const studentId = normalizeId_(
    callSheet.getRange(SCC_CONFIG.callEntry.studentId).getValue()
  );

  if (!studentId) {
    ss.toast('Select a student before opening Demographics.', 'PowerSchool Demographics', 5);
    return;
  }

  try {
    const encoded = Utilities.base64EncodeWebSafe(
      JSON.stringify({ v: 1, studentNumber: studentId }),
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
