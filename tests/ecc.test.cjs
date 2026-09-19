'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const script = fs.readFileSync(path.join(__dirname, '..', 'LogECC.gs'), 'utf8');
const headers = [
  'Prefered Name', 'Student Number', 'SCC Notes', 'ECC Date',
  'Recent ECC Notes \nOverall', 'Recent ECC Notes \nClasses',
  'Recent ECC Notes \nGrades', 'Recent ECC Notes \nSocially',
  'Old ECC Dates and Notes', 'Attempt 1', 'Attempt 2',
  'Small Group', 'Student Email', 'LAST NAME', 'FIRST NAME',
  'GRADE', 'START DATE', 'SCHOOL', 'ES (Signature on ISMA)',
  'SPED', 'EL', 'MKV', 'Parent Language', 'Name', 'Subject Line',
  'OneNote link'
];

function environment(extraRows = []) {
  const data = [
    [...headers],
    [
      'Test Student', 123456, '', new Date(2026, 8, 11),
      'Needs help with homework', 'Geometry is going well', '', 'Feeling connected',
      '-- 8/20/2026: Prior conversation', '', '',
      ...Array(15).fill('')
    ],
    ...extraRows
  ];
  let maxColumns = 29;
  let activeRow = 2;
  const toasts = [];
  const events = [];
  const range = (row, col, rowCount = 1, colCount = 1) => ({
    getValue: () => data[row - 1]?.[col - 1],
    getDisplayValue: () => String(data[row - 1]?.[col - 1] ?? ''),
    getDisplayValues: () => Array.from({ length: rowCount },
      (_, i) => Array.from({ length: colCount },
        (_, j) => String(data[row - 1 + i]?.[col - 1 + j] ?? ''))),
    setValue(value) {
      while (data.length < row) data.push([]);
      data[row - 1][col - 1] = value;
      return this;
    },
    setWrap() { return this; },
    setVerticalAlignment() { return this; },
    setDataValidation() { return this; },
    insertCheckboxes() { return this; },
    setNote() { return this; }
  });
  const sheet = {
    getName: () => 'ECC',
    getLastColumn: () => data[0].length,
    getLastRow: () => data.length,
    getMaxColumns: () => maxColumns,
    getMaxRows: () => 1000,
    insertColumnsAfter() { throw new Error('Existing columns must not be shifted.'); },
    getRange: range,
    getActiveCell: () => ({ getRow: () => activeRow })
  };
  const ss = {
    getSheetByName: name => name === 'ECC' ? sheet : null,
    getActiveSheet: () => sheet,
    getSpreadsheetTimeZone: () => 'America/Los_Angeles',
    toast: message => toasts.push(message)
  };
  const context = vm.createContext({
    Date,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      flush: () => {},
      newDataValidation: () => ({
        requireValueInList() { return this; },
        requireCheckbox() { return this; },
        setAllowInvalid() { return this; },
        build() { return {}; }
      })
    },
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      Charset: { UTF_8: 'UTF_8' },
      formatDate: date => (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear(),
      base64EncodeWebSafe: text => Buffer.from(text).toString('base64url')
    },
    APP_CONFIG: { sheets: { ecc: 'ECC' } },
    withRosterLock_: work => work(),
    getRequiredSheet_: (_ss, name) => ss.getSheetByName(name),
    ensureSheetColumns_: (_sheet, count) => { maxColumns = Math.max(maxColumns, count); },
    normalizeId_: value => String(value ?? '').trim(),
    getPowerSchoolContactSettings_: (_ss, workflow) => ({
      workflow, typeValue: '1187', subtypeValue: 'GE:ECC', extraDropdowns: []
    }),
    logAutomationEvent_: (...args) => events.push(args),
    getErrorDetails_: error => String(error?.message || error)
  });
  vm.runInContext(script, context, { filename: 'LogECC.gs' });
  return {
    data, sheet, context, toasts, events,
    setActiveRow: row => { activeRow = row; }
  };
}

test('reads all four wrapped headers and preserves recent notes after archiving', () => {
  const env = environment();
  const cols = env.context.getEccBatchColumns_(env.sheet);
  assert.deepEqual(
    [cols.overall, cols.classes, cols.grades, cols.socially, cols.history],
    [5, 6, 7, 8, 9]
  );
  const original = env.data[1].slice(3, 9);
  const result = env.context.logEccRow_(env.sheet, 2, cols);
  assert.equal(result.duplicate, false);
  assert.match(result.note, /Overall: Needs help with homework/);
  assert.match(result.note, /Classes: Geometry is going well/);
  assert.match(result.note, /Socially: Feeling connected/);
  assert.match(env.data[1][8], /Prior conversation[\s\S]*Overall:/);
  assert.deepEqual(env.data[1].slice(3, 8), original.slice(0, 5));
});

test('sets up only two controls after existing data and logs multiple ready rows', () => {
  const second = Array(26).fill('');
  second[1] = 789012;
  second[3] = new Date(2026, 8, 12);
  second[6] = 'Current grade update';
  second[8] = 'Earlier history';
  const env = environment([second]);
  const before = env.data[0].slice();
  env.context.setupEccBatchLayout();
  assert.deepEqual(env.data[0].slice(0, 26), before);
  assert.deepEqual(env.data[0].slice(26), ['ECC Type', 'Ready to Log']);
  env.data[1][27] = true;
  env.data[2][27] = true;
  env.context.logReadyEccRows();
  assert.match(env.data[1][8], /Overall:/);
  assert.match(env.data[2][8], /Grades: Current grade update/);
  assert.equal(env.data[1][27], false);
  assert.equal(env.data[2][27], false);
  assert.equal(env.data[1][4], 'Needs help with homework');
});

test('repeated handoff reopens without appending another history entry', () => {
  const env = environment();
  env.context.logCurrentEccRowAndOpenPowerSchool();
  const history = env.data[1][8];
  env.context.logCurrentEccRowAndOpenPowerSchool();
  assert.equal(env.data[1][8], history);
  assert.equal(env.toasts.filter(t => t.startsWith('ECC_HANDOFF_V1:')).length, 2);
});

test('a full attempt pair leaves history and recent notes intact', () => {
  const env = environment();
  env.context.setupEccBatchLayout();
  env.data[1][9] = 'Old attempt one';
  env.data[1][10] = 'Old attempt two';
  env.data[1][26] = 'Attempt';
  const cols = env.context.getEccBatchColumns_(env.sheet);
  const history = env.data[1][8];
  assert.throws(() => env.context.logEccRow_(env.sheet, 2, cols),
    /Both ECC attempt columns are full/);
  assert.equal(env.data[1][8], history);
  assert.equal(env.data[1][4], 'Needs help with homework');
});

test('the old handoff menu name still opens PowerSchool once', () => {
  const env = environment();
  env.context.archiveAndOpenPowerSchoolECC();
  assert.equal(env.toasts.filter(t => t.startsWith('ECC_HANDOFF_V1:')).length, 1);
  assert.match(env.data[1][8], /Overall:/);
});

test('the old archive-only menu name never opens PowerSchool', () => {
  const env = environment();
  env.context.archiveCurrentECCNote();
  assert.match(env.data[1][8], /Overall:/);
  assert.equal(env.toasts.filter(t => t.startsWith('ECC_HANDOFF_V1:')).length, 0);
});
