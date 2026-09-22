'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(
  path.join(__dirname, '..', 'PowerSchoolDemographics.gs'),
  'utf8'
);

function fixture({
  sheetName = 'Call Entry', studentId = 12345678,
  selectedValue = 87654321, rows = 1, columns = 1
} = {}) {
  const events = [];
  const toasts = [];
  const callSheet = {
    getName: () => 'Call Entry',
    getRange: () => ({ getValue: () => studentId })
  };
  const activeSheet = { getName: () => sheetName };
  const activeRange = {
    getSheet: () => activeSheet,
    getValue: () => selectedValue,
    getNumRows: () => rows,
    getNumColumns: () => columns
  };
  const ss = { toast: (...args) => toasts.push(args) };
  const context = vm.createContext({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getActiveRange: () => activeRange,
      flush() {}
    },
    Utilities: {
      Charset: { UTF_8: 'UTF-8' },
      base64EncodeWebSafe: data => Buffer.from(data).toString('base64url')
    },
    SCC_CONFIG: {
      sheets: { callEntry: 'Call Entry' },
      callEntry: { demographicsAction: 'B15', studentId: 'O2' }
    },
    getRequiredSheet_: (_ss, name) => {
      assert.equal(name, 'Call Entry');
      return callSheet;
    },
    normalizeId_: value => String(value || '').trim().replace(/\.0$/, ''),
    logAutomationEvent_: (...args) => events.push(args),
    getErrorDetails_: error => String(error)
  });
  vm.runInContext(source, context);
  return {
    context,
    events,
    toasts
  };
}

function decodeMarker(toasts) {
  const marker = toasts.find(args =>
    String(args[0]).startsWith('DEMOGRAPHICS_HANDOFF_V1:')
  )?.[0];
  if (!marker) return null;
  return JSON.parse(Buffer.from(
    marker.split('DEMOGRAPHICS_HANDOFF_V1:')[1],
    'base64url'
  ).toString());
}

test('menu action on Call Entry uses the form student regardless of selected cell', () => {
  const env = fixture();
  env.context.openStudentDemographics();
  assert.deepEqual(decodeMarker(env.toasts), {
    v: 1,
    studentNumber: '12345678'
  });
  assert.equal(env.events[0][1], 'Demographics Handoff');
});

test('menu action on another tab uses the selected student-number cell', () => {
  const env = fixture({ sheetName: 'SCC', selectedValue: 87654321 });
  env.context.openStudentDemographics();
  assert.deepEqual(decodeMarker(env.toasts), {
    v: 1,
    studentNumber: '87654321'
  });
});

test('Call Entry without a selected student shows a friendly message', () => {
  const env = fixture({ studentId: '' });
  env.context.openStudentDemographics();
  assert.equal(decodeMarker(env.toasts), null);
  assert.match(env.toasts[0][0], /Select a student/);
});

test('another tab requires one cell containing a valid student number', () => {
  const text = fixture({ sheetName: 'SCC', selectedValue: 'Completed' });
  text.context.openStudentDemographics();
  assert.equal(decodeMarker(text.toasts), null);
  assert.match(text.toasts[0][0], /valid Student Number/);

  const many = fixture({ sheetName: 'SCC', rows: 2 });
  many.context.openStudentDemographics();
  assert.equal(decodeMarker(many.toasts), null);
  assert.match(many.toasts[0][0], /Select one cell/);
});
