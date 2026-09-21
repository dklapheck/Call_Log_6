'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'PowerSchoolSettings.gs'), 'utf8');
const DATE_FIELDS = ['entryLogDate', 'UF-008019-1', 'UF-008041-1'];
const TAGS = {
  1: 'Attempt 1 (34)', 2: 'Attempt 2 (35)', 3: 'Attempt 3 (36)',
  4: 'Attempt 4 (37)', 5: 'Attempt 5 (38)', 6: 'Attempt 6 (39)'
};

function makeSettings() {
  const headers = [
    'Workflow', 'Log Type value', 'Log Type label', 'Subtype value',
    'Subtype label', 'Additional dropdowns (JSON)', 'Guidance',
    'Date & Time field', 'Incident Date field', 'Action Date field',
    'Attempt tags (JSON)'
  ];
  const rows = [
    ['SCC Success', '', '', '', '', '', '', ...DATE_FIELDS, ''],
    ['SCC Attempt', '', '', '', '', '', '', ...DATE_FIELDS, JSON.stringify(TAGS)],
    ['ECC Conversation', '1187', 'Student Contact', 'GE:ECC', 'ECC', '', '', ...DATE_FIELDS, ''],
    ['ECC Attempt', '1187', 'Student Contact', 'GE:ECC', 'ECC', '', '', ...DATE_FIELDS, '']
  ];
  const spreadsheet = {
    getSheetByName(name) {
      return name === 'Instructions and Settings'
        ? { getRange(range) {
            assert.equal(range, 'A34:K38');
            return { getDisplayValues: () => [headers, ...rows] };
          } }
        : null;
    }
  };
  const context = vm.createContext({});
  vm.runInContext(source, context);
  return { rows, spreadsheet, read: workflow => context.getPowerSchoolContactSettings_(spreadsheet, workflow) };
}

test('SCC success and attempt read separate live Settings rows', () => {
  const settings = makeSettings();
  settings.rows[0][1] = 'success_type';
  settings.rows[0][3] = 'success_subtype';
  settings.rows[1][1] = 'attempt_type';
  settings.rows[1][3] = 'attempt_subtype';
  assert.equal(settings.read('SCC Success').subtypeValue, 'success_subtype');
  assert.equal(settings.read('SCC Attempt').typeValue, 'attempt_type');
  settings.rows[0][3] = 'changed_on_website';
  assert.equal(settings.read('SCC Success').subtypeValue, 'changed_on_website');
});

test('three date fields and six Attempt tags come from the live workflow row', () => {
  const settings = makeSettings();
  assert.deepEqual(Array.from(settings.read('SCC Success').dateFields), DATE_FIELDS);
  assert.equal(settings.read('SCC Attempt').tagMap['4'], 'Attempt 4 (37)');
  assert.deepEqual(Object.keys(settings.read('ECC Conversation').tagMap), []);
});

test('blank manual settings and safe additional dropdowns', () => {
  const settings = makeSettings();
  assert.equal(settings.read('SCC Success').typeValue, '');
  settings.rows[2][5] = '[{"name":"result","value":"answered"}]';
  assert.equal(settings.read('ECC Conversation').extraDropdowns[0].value, 'answered');
  settings.rows[2][5] = '[{"name":"studentNumber","value":"123"}]';
  assert.throws(() => settings.read('ECC Conversation'), /Additional dropdowns/);
});

test('invalid date fields and tag maps are rejected', () => {
  const settings = makeSettings();
  settings.rows[1][8] = 'bad field[]';
  assert.throws(() => settings.read('SCC Attempt'), /log date field/);
  settings.rows[1][8] = DATE_FIELDS[1];
  settings.rows[1][10] = '{"7":"Attempt 7"}';
  assert.throws(() => settings.read('SCC Attempt'), /Attempt tags/);
});

test('partial Type/Subtype settings cannot create a handoff', () => {
  const settings = makeSettings();
  settings.rows[1][1] = 'only_type';
  assert.throws(() => settings.read('SCC Attempt'), /Fill both Log Type value and Subtype value/);
});
