const AUTOMATION_LOG_MAX_ROWS_ = 300;

function getAutomationLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(APP_CONFIG.sheets.automationLog);

  if (!sheet) {
    sheet = ss.insertSheet(APP_CONFIG.sheets.automationLog);
    sheet.getRange('A1:F1').setValues([[
      'Timestamp',
      'Level',
      'Action',
      'Student Number',
      'Message',
      'Details'
    ]]);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:F1').setFontWeight('bold');
    sheet.setColumnWidth(1, 155);
    sheet.setColumnWidth(2, 75);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 260);
    sheet.setColumnWidth(6, 520);
    sheet.getRange('E:F').setWrap(true).setVerticalAlignment('top');
  }

  return sheet;
}

function logAutomationEvent_(level, action, studentNumber, message, details) {
  try {
    const sheet = getAutomationLogSheet_();
    sheet.appendRow([
      new Date(),
      String(level || 'INFO'),
      String(action || ''),
      normalizeId_(studentNumber),
      String(message || ''),
      String(details || '')
    ]);

    const row = sheet.getLastRow();
    sheet.getRange(row, 1).setNumberFormat('m/d/yyyy h:mm:ss am/pm');
    sheet.getRange(row, 5, 1, 2).setWrap(true).setVerticalAlignment('top');

    const dataRows = sheet.getLastRow() - 1;
    if (dataRows > AUTOMATION_LOG_MAX_ROWS_) {
      sheet.deleteRows(2, dataRows - AUTOMATION_LOG_MAX_ROWS_);
    }
  } catch (loggingError) {
    console.error('Could not write Automation Log:', loggingError);
  }
}

function getErrorDetails_(error) {
  if (!error) return 'Unknown error.';
  const message = error.message || String(error);
  const stack = error.stack ? '\n\n' + error.stack : '';
  return message + stack;
}

function openAutomationLog() {
  const sheet = getAutomationLogSheet_();
  sheet.showSheet();
  sheet.activate();
  sheet.getRange(sheet.getLastRow(), 1).activate();
}

function clearAutomationLog() {
  const sheet = getAutomationLogSheet_();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Automation Log cleared.', 'Automation Log', 4);
}
