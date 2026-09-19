// Editable PowerSchool contact choices in Instructions and Settings!A34:G38.
// Read on every handoff, so the next call uses the latest saved cells.
function getPowerSchoolContactSettings_(ss, workflow) {
  const sheet = ss.getSheetByName('Instructions and Settings');
  if (!sheet) throw new Error('Instructions and Settings tab is missing.');

  const rows = sheet.getRange('A35:G38').getDisplayValues();
  const row = rows.find(function(values) {
    return String(values[0] || '').trim() === workflow;
  });
  if (!row) throw new Error('PowerSchool settings row is missing: ' + workflow);

  const typeValue = String(row[1] || '').trim();
  const subtypeValue = String(row[3] || '').trim();
  if (Boolean(typeValue) !== Boolean(subtypeValue)) {
    throw new Error('Fill both Log Type value and Subtype value for ' + workflow +
      ' in Instructions and Settings, or leave both blank for manual selection.');
  }

  let extraDropdowns = [];
  const extraText = String(row[5] || '').trim();
  if (extraText) {
    try { extraDropdowns = JSON.parse(extraText); }
    catch (_) { throw new Error('Additional dropdowns for ' + workflow + ' must be valid JSON.'); }
    if (!Array.isArray(extraDropdowns) || extraDropdowns.length > 8 ||
        !extraDropdowns.every(function(entry) {
          return entry && typeof entry.name === 'string' &&
            /^[A-Za-z0-9_:-]{1,80}$/.test(entry.name) &&
            !/student|pupil|person|parent|guardian|contact|teacher|staff|school|section|course|email|phone|address|frn|date|month|day|year|calendar|time/i.test(entry.name) &&
            typeof entry.value === 'string' && entry.value.length > 0 && entry.value.length <= 120;
        })) {
      throw new Error('Additional dropdowns for ' + workflow +
        ' must be a JSON array of up to 8 {"name":"...","value":"..."} pairs.');
    }
  }

  return {
    workflow: workflow,
    typeValue: typeValue,
    subtypeValue: subtypeValue,
    extraDropdowns: extraDropdowns
  };
}
