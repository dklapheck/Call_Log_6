'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'EccEmailBatch.gs'), 'utf8');
const id = '12345678-1234-1234-1234-123456789abc';
let sequence = 0;
const nextId = () => id.slice(0, -1) + (sequence++ % 10);

function sheet(name, rows) {
  return {
    getName: () => name,
    getLastColumn: () => rows[0].length,
    getDataRange: () => ({
      getValues: () => rows,
      getDisplayValues: () => rows.map(row => row.map(value => String(value ?? '')))
    })
  };
}

function environment() {
  sequence = 0;
  const eccRows = [
    ['Prefered Name', 'Student Number', 'Attempts', 'Student Email',
      'Email?', 'Email Reminder', 'Email Reschedule'],
    ['Alex Example', 11, 'Old call', 'alex@example.edu', true,
      'Hello [Student Preferred First Name] and [LC]\nPlease reply.', 'Reschedule [LC]'],
    ['Jo Example', 22, '', 'jo@example.edu', true, 'Hi Jo', 'Reschedule Jo']
  ];
  const dataRows = [
    ['Student Number', 'Learning Coach', 'Learning Coach Email'],
    [11, 'Taylor', 'taylor@example.edu'],
    [22, '', '']
  ];
  const ecc = sheet('ECC', eccRows);
  ecc.getRange = (row, column) => ({
    setValue: value => { eccRows[row - 1][column - 1] = value; }
  });
  const student = sheet('StudentData', dataRows);
  const store = {};
  const ss = {
    getId: () => 'spreadsheet-one',
    getSpreadsheetTimeZone: () => 'America/Los_Angeles',
    getSheetByName: name => name === 'ECC' ? ecc : student
  };
  const context = vm.createContext({
    Date,
    APP_CONFIG: { sheets: { ecc: 'ECC', studentData: 'StudentData' } },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    PropertiesService: { getDocumentProperties: () => ({
      setProperty: (key, value) => { store[key] = value; },
      getProperty: key => store[key]
    }) },
    Utilities: {
      getUuid: nextId,
      formatDate: () => '9/24/2026 9:00 AM'
    },
    getRequiredSheet_: (spreadsheet, name) => spreadsheet.getSheetByName(name),
    findHeaderColumn_: (s, name) => {
      const rows = s.getDataRange().getValues();
      const index = rows[0].indexOf(name);
      if (index < 0) throw new Error('Missing ' + name);
      return index + 1;
    },
    normalizeId_: value => String(value ?? '').trim(),
    appendHistoryEntry_: (oldValue, entry) => oldValue ? oldValue + '\n' + entry : entry,
    withRosterLock_: fn => fn()
  });
  vm.runInContext(source, context, { filename: 'EccEmailBatch.gs' });
  return { ecc, student, eccRows, store, context };
}

test('prepares checked students, resolves placeholders and student-only fallback', () => {
  const env = environment();
  const batch = env.context.eccEmailBatchData_(env.ecc, env.student, 'reminder',
    'spreadsheet-one', nextId);
  assert.equal(batch.rows.length, 2);
  assert.deepEqual(Array.from(batch.rows[0].to), ['alex@example.edu', 'taylor@example.edu']);
  assert.equal(batch.rows[0].body, 'Hello Alex and Taylor\nPlease reply.');
  assert.deepEqual(Array.from(batch.rows[1].to), ['jo@example.edu']);
  assert.match(batch.rows[1].warnings[0], /student only/);
  assert.equal(env.eccRows[1][4], true);
});

test('a selected row with invalid email aborts the entire batch', () => {
  const env = environment();
  env.eccRows[2][3] = 'bad@example.edu;other@example.edu';
  assert.throws(() => env.context.eccEmailBatchData_(env.ecc, env.student,
    'reminder', 'spreadsheet-one', nextId), /ECC row 3.*Student Email/s);
});

test('receipt appends attempts once, verifies batch and leaves checkbox checked', () => {
  const env = environment();
  const batch = env.context.eccEmailBatchData_(env.ecc, env.student,
    'reschedule', 'spreadsheet-one', nextId);
  env.store['ecc_email_batch_' + batch.batchId] = JSON.stringify({
    kind: batch.kind, issuedAt: batch.issuedAt,
    rows: batch.rows.map(r => ({ requestId: r.requestId, studentNumber: r.studentNumber }))
  });
  const receipt = {
    version: 1, batchId: batch.batchId, spreadsheetId: 'spreadsheet-one',
    kind: 'reschedule', clicked: [{
      requestId: batch.rows[0].requestId, studentNumber: '11',
      clickedAt: '2026-09-24T16:00:00.000Z'
    }]
  };
  assert.match(env.context.recordEccEmailReceipts(JSON.stringify(receipt)), /1 attempt/);
  assert.match(env.eccRows[1][2], /Old call\n.*Email attempt 1 \(reschedule; Send clicked\)/);
  assert.equal(env.eccRows[1][4], true);
  assert.match(env.context.recordEccEmailReceipts(JSON.stringify(receipt)), /0 attempt/);
  assert.equal((env.eccRows[1][2].match(/Email attempt/g) || []).length, 1);
  receipt.clicked[0].studentNumber = '22';
  assert.throws(() => env.context.recordEccEmailReceipts(JSON.stringify(receipt)),
    /unknown, duplicate, or invalid/);
});
