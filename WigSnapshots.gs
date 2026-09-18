function saveWigSnapshot() {
  try {
    const row = snapshotWig_();
    SpreadsheetApp.getUi().alert('WIG columns B:G saved as values on row ' + row + '.');
  } catch (error) {
    SpreadsheetApp.getUi().alert(error.message);
  }
}

function scheduledWigSnapshot() {
  snapshotWig_();
}

function enableTuesdayWigSnapshots() {
  const fn = 'scheduledWigSnapshot';
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === fn) ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger(fn)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(16)
    .create();
}

function snapshotWig_() {
  return withRosterLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dash = getRequiredSheet_(ss, APP_CONFIG.sheets.dash);
    const wig = getRequiredSheet_(ss, APP_CONFIG.sheets.wig);
    const sourceRow = APP_CONFIG.dash.wigRow;
    const date = dash.getRange(sourceRow, 1).getValue();
    if (!(date instanceof Date)) throw new Error('Dash A' + sourceRow + ' has no valid week date.');

    const key = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    const count = Math.max(wig.getLastRow() - 1, 1);
    const dates = wig.getRange(2, 1, count, 1).getValues();
    let match = -1;

    for (let i = 0; i < dates.length; i++) {
      if (dates[i][0] instanceof Date &&
          Utilities.formatDate(dates[i][0], ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') === key) {
        match = i;
        break;
      }
    }

    if (match < 0) throw new Error('Add the date ' + key + ' to WIG column A first.');
    const rowNumber = match + 2;
    const target = wig.getRange(rowNumber, 2, 1, 6);

    if (target.getValues()[0].some(function(value) { return value !== ''; })) {
      throw new Error('WIG row ' + rowNumber + ' already has a snapshot. It was not changed.');
    }

    const values = dash.getRange(sourceRow, 2, 1, 6).getValues()[0];
    if (typeof values[0] !== 'number' ||
        typeof values[1] !== 'number' ||
        typeof values[3] !== 'number' ||
        typeof values[4] !== 'number' ||
        typeof values[5] !== 'number') {
      throw new Error('Review the live WIG values on Dash row ' + sourceRow + ' before saving.');
    }

    target.setValues([values]);
    target.setNote(
      'Snapshot saved ' +
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d/yyyy h:mm a') +
      '.'
    );
    return rowNumber;
  });
}
