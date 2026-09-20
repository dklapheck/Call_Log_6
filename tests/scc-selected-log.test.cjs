'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'SaveSCC.gs'), 'utf8');

function fixture({
  selectedHeader = 'SCC Notes',
  note = 'On 9/20/2026, spoke with parent.',
  studentId = 12345678,
  rows = 1,
  columns = 1,
  selectedRow = 2
} = {}) {
  const headers = [
    'Student Number', 'SCC Notes', 'Attempt 1', 'Attempt 2',
    'Attempt 3', 'Attempt 4', 'Attempt 5', 'SCC Completion'
  ];
  const selectedHeaderIndex = headers.indexOf(selectedHeader);
  const selectedColumn = selectedHeaderIndex === -1 ? 8 : selectedHeaderIndex + 1;
  const events = [];
  const toasts = [];
  const cell = value => ({
    getValue: () => value,
    getDisplayValue: () => String(value ?? '')
  });
  const rosterSheet = {
    getName: () => 'SCC',
    getRange(row, column) {
      if (row === 1) return cell(headers[column - 1]);
      if (row === selectedRow && column === 1) return cell(studentId);
      if (row === selectedRow && column === selectedColumn) return cell(note);
      return cell('');
    }
  };
  const activeRange = {
    getSheet: () => rosterSheet,
    getRow: () => selectedRow,
    getColumn: () => selectedColumn,
    getNumRows: () => rows,
    getNumColumns: () => columns,
    getDisplayValue: () => note
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
      sheets: { callEntry: 'Call Entry', roster: 'SCC' },
      callEntry: { studentId: 'O2', note: 'D1', toDo: 'I1' },
      roster: {
        headerRow: 1,
        studentIdHeader: 'Student Number',
        notesHeader: 'SCC Notes',
        attemptHeaders: ['Attempt 1', 'Attempt 2', 'Attempt 3', 'Attempt 4', 'Attempt 5'],
        toDoHeader: 'SCC To Do',
        completionHeader: 'SCC Completion',
        completedValue: 'Completed'
      }
    },
    findHeaderColumn_: (_sheet, header) => {
      const index = headers.indexOf(header);
      if (index === -1) throw new Error('Missing header: ' + header);
      return index + 1;
    },
    normalizeId_: value => String(value ?? '').trim(),
    getPowerSchoolContactSettings_: (_ss, workflow) => ({
      workflow,
      typeValue: workflow === 'SCC Success' ? 'SUCCESS_TYPE' : 'ATTEMPT_TYPE',
      subtypeValue: workflow === 'SCC Success' ? 'SUCCESS_SUBTYPE' : 'ATTEMPT_SUBTYPE',
      extraDropdowns: []
    }),
    logAutomationEvent_: (...args) => events.push(args),
    getErrorDetails_: error => String(error),
    withRosterLock_: fn => fn(),
    getRequiredSheet_: () => { throw new Error('Call Entry should not be read'); },
    findStudentRosterRow_: () => selectedRow,
    getStudentName_: () => 'Student',
    findQuestionRow_: () => 14,
    saveFailedSccAttempt_: () => ({ saved: true })
  });
  vm.runInContext(source, context);
  return { context, events, toasts };
}

function payload(events) {
  const marker = events.find(event => event[1] === 'SCC Handoff')?.[4]
    ?.split('SCC_HANDOFF_V1:')[1];
  return marker ? JSON.parse(Buffer.from(marker, 'base64url').toString()) : null;
}

test('selected SCC Notes cell uses SCC Success settings and its own note', () => {
  const env = fixture({ selectedHeader: 'SCC Notes', note: 'Successful family call.' });
  env.context.logSccInPowerSchool();
  const handoff = payload(env.events);
  assert.equal(handoff.note, 'Successful family call.');
  assert.equal(handoff.studentNumber, '12345678');
  assert.equal(handoff.outcome, 'SCC Success');
  assert.equal(handoff.settings.subtypeValue, 'SUCCESS_SUBTYPE');
});

for (let attempt = 1; attempt <= 5; attempt++) {
  test('selected Attempt ' + attempt + ' cell uses SCC Attempt settings', () => {
    const note = 'Attempt ' + attempt + ' note.';
    const env = fixture({ selectedHeader: 'Attempt ' + attempt, note });
    env.context.logSccInPowerSchool();
    const handoff = payload(env.events);
    assert.equal(handoff.note, note);
    assert.equal(handoff.outcome, 'SCC Attempt');
    assert.equal(handoff.settings.typeValue, 'ATTEMPT_TYPE');
  });
}

test('blank and unrelated SCC cells do not create a handoff', () => {
  const blank = fixture({ selectedHeader: 'Attempt 3', note: '' });
  blank.context.logSccInPowerSchool();
  assert.equal(payload(blank.events), null);
  assert.match(blank.toasts[0][0], /blank/);

  const unrelated = fixture({ selectedHeader: 'SCC Completion', note: 'Completed' });
  unrelated.context.logSccInPowerSchool();
  assert.equal(payload(unrelated.events), null);
  assert.match(unrelated.toasts[0][0], /SCC Notes or Attempt 1/);
});

test('multi-cell selections do not create an ambiguous handoff', () => {
  const env = fixture({ selectedHeader: 'Attempt 1', rows: 2 });
  env.context.logSccInPowerSchool();
  assert.equal(payload(env.events), null);
  assert.match(env.toasts[0][0], /Select one/);
});
