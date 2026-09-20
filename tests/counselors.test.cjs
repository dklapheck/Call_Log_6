'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'Counselors.gs'), 'utf8');

function fixture() {
  const state = {
    roster: [
      ['Student Number', 'Counselor', 'Counselor Email'],
      [1001, 'Old Counselor', 'old@example.org'],
      [1002, 'Known Counselor', 'known@example.org'],
      [1003, '', '']
    ],
    upstream: [
      ['STUDENT_NUMBER', 'COUNSELOR'],
      [1001, 'Known Counselor'],
      [1002, 'Known Counselor'],
      [1003, 'Unknown Counselor']
    ]
  };
  const toasts = [];

  const makeRange = (sheetName, row, col, numRows, numCols) => ({
    getDisplayValues() {
      const table = state[sheetName];
      return table.slice(row - 1, row - 1 + numRows)
        .map(r => r.slice(col - 1, col - 1 + numCols).map(v => String(v ?? '')));
    },
    setValues(values) {
      values.forEach((valuesRow, r) => valuesRow.forEach((value, c) => {
        state[sheetName][row - 1 + r][col - 1 + c] = value;
      }));
      return this;
    }
  });
  const sheets = {
    roster: {
      getLastRow: () => state.roster.length,
      getRange: (row, col, numRows, numCols) => makeRange('roster', row, col, numRows, numCols)
    },
    upstream: {}
  };
  const ss = { toast: (...args) => toasts.push(args) };
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    APP_CONFIG: { source: { upstream: 'Upstream' } },
    SCC_CONFIG: {
      sheets: { roster: 'SCC' },
      roster: {
        studentIdHeader: 'Student Number',
        counselorHeader: 'Counselor',
        counselorEmailHeader: 'Counselor Email'
      },
      upstream: { studentIdHeader: 'STUDENT_NUMBER', counselorHeader: 'COUNSELOR' }
    },
    refreshStudentData: () => {},
    withRosterLock_: fn => fn(),
    getSourceSpreadsheet_: () => ({}),
    getRequiredSheet_: (_book, name) => name === 'Upstream' ? sheets.upstream : sheets.roster,
    readSheetTable_: () => ({
      headerMap: { STUDENT_NUMBER: 0, COUNSELOR: 1 },
      rows: state.upstream.slice(1)
    }),
    findHeaderColumn_: (_sheet, header) => ({
      'Student Number': 1,
      Counselor: 2,
      'Counselor Email': 3
    })[header],
    normalizeId_: value => String(value ?? '').trim()
  });
  vm.runInContext(source, context);
  return { context, state, toasts };
}

test('refresh syncs upstream names and reuses known counselor emails', () => {
  const env = fixture();
  env.context.refreshCounselorData();
  assert.deepEqual(env.state.roster.slice(1), [
    [1001, 'Known Counselor', 'known@example.org'],
    [1002, 'Known Counselor', 'known@example.org'],
    [1003, 'Unknown Counselor', '']
  ]);
  assert.match(env.toasts[0][0], /Names updated: 2/);
  assert.match(env.toasts[0][0], /emails updated: 1/);
});
