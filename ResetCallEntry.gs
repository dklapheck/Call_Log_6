function resetCallEntry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getRequiredSheet_(ss, SCC_CONFIG.sheets.callEntry);
  sheet.getRange(SCC_CONFIG.callEntry.formRange).uncheck();
  sheet.getRangeList([SCC_CONFIG.callEntry.studentSelector].concat(SCC_CONFIG.callEntry.resetRanges)).clearContent();
  sheet.getRange(SCC_CONFIG.callEntry.studentSelector).activate();
  ss.toast('Call Entry has been reset.', 'Reset', 3);
}
