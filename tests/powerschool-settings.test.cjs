'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'PowerSchoolSettings.gs'), 'utf8');

function makeSettings() {
  const rows = [
    ['SCC Success', '', '', '', '', '', ''],
    ['SCC Attempt', '', '', '', '', '', ''],
    ['ECC Conversation', '1187', 'Student Contact', 'GE:ECC', 'ECC', '', ''],
    ['ECC Attempt', '1187', 'Student Contact', 'GE:ECC', 'ECC', '', '']
  ];
  const spreadsheet = {
    getSheetByName(name) {
      return name === 'Instructions and Settings'
        ? { getRange(range) {
            assert.equal(range, 'A35:G38');
            return { getDisplayValues: () => rows };
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

test('blank manual settings and safe additional dropdowns', () => {
  const settings = makeSettings();
  assert.equal(settings.read('SCC Success').typeValue, '');
  settings.rows[2][5] = '[{"name":"result","value":"answered"}]';
  assert.equal(settings.read('ECC Conversation').extraDropdowns[0].value, 'answered');
  settings.rows[2][5] = '[{"name":"studentNumber","value":"123"}]';
  assert.throws(() => settings.read('ECC Conversation'), /Additional dropdowns/);
  settings.rows[2][5] = '[{"name":"callDate","value":"9"}]';
  assert.throws(() => settings.read('ECC Conversation'), /Additional dropdowns/);
});

test('partial Type/Subtype settings cannot create a handoff', () => {
  const settings = makeSettings();
  settings.rows[1][1] = 'only_type';
  assert.throws(() => settings.read('SCC Attempt'), /Fill both Log Type value and Subtype value/);
});
