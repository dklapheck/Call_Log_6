// Colors the ECC Date column on the ECC tab by how long ago the date was.
// Run applyEccDateColors() once from the Apps Script editor. The rules stay in
// the sheet and recalculate every day on their own (they use TODAY()).
//
//   Empty ECC Date (student row)  -> maroon, white text
//   21+ days ago                  -> red
//   17-20 days ago                -> deep gold
//   14-16 days ago                -> gold
//   10-13 days ago                -> light yellow
//   7-9 days ago                  -> pale yellow
//   0-6 days ago (or in future)   -> green
//
// Re-running replaces the earlier colors on this column. Other formatting
// rules on the sheet, and custom-formula rules on other columns, are untouched.

function applyEccDateColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ECC');
  if (!sheet) throw new Error('ECC tab not found.');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  function findColumn(name) {
    const wanted = name.toLowerCase();
    const matches = [];
    headers.forEach(function(value, index) {
      if (String(value).replace(/\s+/g, ' ').trim().toLowerCase() === wanted) matches.push(index + 1);
    });
    if (matches.length !== 1) {
      throw new Error('Expected exactly one "' + name + '" column on ECC; found ' + matches.length + '.');
    }
    return matches[0];
  }
  function letter(col) {
    return sheet.getRange(1, col).getA1Notation().replace(/\d+/, '');
  }

  const dateCol = findColumn('ECC Date');
  const studentCol = findColumn('Student Number');
  const d = '$' + letter(dateCol) + '2';
  const s = '$' + letter(studentCol) + '2';

  const lastRow = sheet.getMaxRows();
  const target = sheet.getRange(2, dateCol, lastRow - 1, 1);

  // Remove this column's earlier custom-formula rules so re-running does not stack them.
  const kept = sheet.getConditionalFormatRules().filter(function(rule) {
    const condition = rule.getBooleanCondition();
    if (!condition || condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return true;
    return !rule.getRanges().some(function(range) {
      return range.getColumn() === dateCol && range.getNumColumns() === 1;
    });
  });

  function rule(formula, background, fontColor) {
    const builder = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(background)
      .setRanges([target]);
    if (fontColor) builder.setFontColor(fontColor);
    return builder.build();
  }

  const age = 'TODAY()-' + d;
  const isDate = 'ISNUMBER(' + d + ')';
  const rules = [
    rule('=AND(' + d + '="",' + s + '<>"")', '#7B1113', '#FFFFFF'),   // maroon: empty
    rule('=AND(' + isDate + ',' + age + '>=21)', '#E06666'),          // red: past 20 days
    rule('=AND(' + isDate + ',' + age + '>=17,' + age + '<=20)', '#F1C232'),
    rule('=AND(' + isDate + ',' + age + '>=14,' + age + '<=16)', '#FFD966'),
    rule('=AND(' + isDate + ',' + age + '>=10,' + age + '<=13)', '#FFE599'),
    rule('=AND(' + isDate + ',' + age + '>=7,' + age + '<=9)', '#FFF2CC'),
    rule('=AND(' + isDate + ',' + age + '<=6)', '#B6D7A8')            // green: recent
  ];

  sheet.setConditionalFormatRules(kept.concat(rules));
  ss.toast('ECC Date colors applied.', 'ECC Tools', 5);
}
