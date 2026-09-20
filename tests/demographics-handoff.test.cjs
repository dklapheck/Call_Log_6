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

function fixture({ sheetName = 'Call Entry', cell = 'B15', studentId = 12345678 } = {}) {
  const events = [];
  const toasts = [];
  const callSheet = {
    getName: () => sheetName,
    getRange: () => ({ getValue: () => studentId })
  };
  const ss = { toast: (...args) => toasts.push(args) };
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    Utilities: {
      Charset: { UTF_8: 'UTF-8' },
      base64EncodeWebSafe: data => Buffer.from(data).toString('base64url')
    },
    SCC_CONFIG: {
      sheets: { callEntry: 'Call Entry' },
      callEntry: { demographicsAction: 'B15', studentId: 'O2' }
    },
    getRequiredSheet_: () => callSheet,
    normalizeId_: value => String(value || ''),
    logAutomationEvent_: (...args) => events.push(args),
    getErrorDetails_: error => String(error)
  });
  vm.runInContext(source, context);
  return {
    context,
    event: {
      range: {
        getSheet: () => callSheet,
        getA1Notation: () => cell
      }
    },
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

test('selecting Call Entry B15 sends the selected student to PowerSchool', () => {
  const env = fixture();
  env.context.onSelectionChange(env.event);
  assert.deepEqual(decodeMarker(env.toasts), {
    v: 1,
    studentNumber: '12345678'
  });
  assert.equal(env.events[0][1], 'Demographics Handoff');
});

test('other cells and sheets do nothing', () => {
  const otherCell = fixture({ cell: 'B14' });
  otherCell.context.onSelectionChange(otherCell.event);
  assert.equal(otherCell.toasts.length, 0);

  const otherSheet = fixture({ sheetName: 'SCC' });
  otherSheet.context.onSelectionChange(otherSheet.event);
  assert.equal(otherSheet.toasts.length, 0);
});

test('no selected student shows a friendly message without a handoff', () => {
  const env = fixture({ studentId: '' });
  env.context.onSelectionChange(env.event);
  assert.equal(decodeMarker(env.toasts), null);
  assert.match(env.toasts[0][0], /Select a student/);
});
