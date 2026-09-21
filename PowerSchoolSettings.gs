// Editable PowerSchool contact choices in Instructions and Settings!A34:K38.
// Read on every handoff, so the next call uses the latest saved cells.
function getPowerSchoolContactSettings_(ss, workflow) {
  const sheet = ss.getSheetByName('Instructions and Settings');
  if (!sheet) throw new Error('Instructions and Settings tab is missing.');

  const block = sheet.getRange('A34:K38').getDisplayValues();
  const headers = block[0].map(function(value) { return String(value || '').trim(); });
  const rows = block.slice(1);
  const workflowColumn = headers.indexOf('Workflow');
  const row = rows.find(function(values) {
    return String(values[workflowColumn] || '').trim() === workflow;
  });
  if (!row) throw new Error('PowerSchool settings row is missing: ' + workflow);

  function valueFor_(header) {
    const index = headers.indexOf(header);
    return index === -1 ? '' : String(row[index] || '').trim();
  }

  const typeValue = valueFor_('Log Type value');
  const subtypeValue = valueFor_('Subtype value');
  if (Boolean(typeValue) !== Boolean(subtypeValue)) {
    throw new Error('Fill both Log Type value and Subtype value for ' + workflow +
      ' in Instructions and Settings, or leave both blank for manual selection.');
  }

  let extraDropdowns = [];
  const extraText = valueFor_('Additional dropdowns (JSON)');
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

  const dateFields = [
    valueFor_('Date & Time field'),
    valueFor_('Incident Date field'),
    valueFor_('Action Date field')
  ].filter(Boolean);
  if (!dateFields.length) {
    const legacyDateField = valueFor_('Log date field');
    if (legacyDateField) dateFields.push(legacyDateField);
  }
  if (dateFields.length > 6 || dateFields.some(function(name) {
    return !/^[A-Za-z0-9_$:=.-]{1,120}$/.test(name);
  })) {
    throw new Error('A log date field for ' + workflow + ' is not a safe PowerSchool control name.');
  }

  let tagMap = {};
  const tagMapText = valueFor_('Attempt tags (JSON)');
  if (tagMapText) {
    try { tagMap = JSON.parse(tagMapText); }
    catch (_) { throw new Error('Attempt tags for ' + workflow + ' must be valid JSON.'); }
    if (!tagMap || Array.isArray(tagMap) || typeof tagMap !== 'object' ||
        !Object.keys(tagMap).every(function(key) {
          return /^[1-6]$/.test(key) && typeof tagMap[key] === 'string' &&
            tagMap[key].length > 0 && tagMap[key].length <= 120 && !/[\t\r\n]/.test(tagMap[key]);
        })) {
      throw new Error('Attempt tags for ' + workflow +
        ' must map attempt numbers 1–6 to one-line PowerSchool labels.');
    }
  }

  const tagLabel = valueFor_('Tag label');
  if (tagLabel.length > 120 || /[\t\r\n]/.test(tagLabel)) {
    throw new Error('Tag label for ' + workflow + ' must be one line and no more than 120 characters.');
  }

  return {
    workflow: workflow,
    typeValue: typeValue,
    subtypeValue: subtypeValue,
    extraDropdowns: extraDropdowns,
    dateFields: dateFields,
    dateField: dateFields[0] || '',
    tagMap: tagMap,
    tagLabel: tagLabel
  };
}
