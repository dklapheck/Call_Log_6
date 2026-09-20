'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'SaveSCC.gs'), 'utf8');

function fixture({ yes = true, no = false, attemptSaved = true, duplicate = false } = {}) {
  const values = {
    studentId: 12345678,
    note: 'On 9/18/2026, spoke with parent.',
    toDo: 'Send resources.',
    yes, no,
    rosterNote: duplicate ? 'On 9/18/2026, spoke with parent.' : '',
    completion: duplicate ? 'Completed' : '',
    rosterToDo: ''
  };
  const events = [];
  const calls = [];
  const makeCell = (key) => ({
    getValue: () => values[key],
    getDisplayValue: () => String(values[key] ?? ''),
    setValue(value) { values[key] = value; return this; }
  });
  const callSheet = { getRange(a, col) {
    const key = typeof a === 'string'
      ? ({ O2: 'studentId', D1: 'note', I1: 'toDo' })[a]
      : a === 14 && col === 4 ? 'yes' : a === 14 && col === 6 ? 'no' : null;
    if (!key) throw new Error('Unexpected Call Entry cell');
    return makeCell(key);
  } };
  const rosterSheet = { getRange(_row, col) {
    return makeCell(({ 1: 'rosterNote', 2: 'rosterToDo', 3: 'completion' })[col]);
  } };
  const ss = { toast: (...args) => calls.push(args) };
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    Utilities: {
      Charset: { UTF_8: 'UTF-8' },
      base64EncodeWebSafe: data => Buffer.from(data).toString('base64url')
    },
    SCC_CONFIG: {
      sheets: { callEntry: 'Call Entry', roster: 'SCC' },
      callEntry: { studentId: 'O2', note: 'D1', toDo: 'I1' },
      roster: {
        notesHeader: 'SCC Notes',
        toDoHeader: 'SCC To Do',
        completionHeader: 'SCC Completion',
        completedValue: 'Completed'
      }
    },
    withRosterLock_: fn => fn(),
    getRequiredSheet_: (_ss, name) => name === 'Call Entry' ? callSheet : rosterSheet,
    findStudentRosterRow_: () => 2,
    getStudentName_: () => 'Student',
    findQuestionRow_: () => 14,
    findHeaderColumn_: (_sheet, name) => ({ 'SCC Notes': 1, 'SCC To Do': 2, 'SCC Completion': 3 })[name],
    saveFailedSccAttempt_: () => ({ saved: attemptSaved }),
    normalizeId_: id => String(id),
    getPowerSchoolContactSettings_: (_ss, workflow) => ({
      workflow, typeValue: workflow === 'SCC Success' ? 'SUCCESS_TYPE' : 'ATTEMPT_TYPE',
      subtypeValue: workflow === 'SCC Success' ? 'SUCCESS_SUBTYPE' : 'ATTEMPT_SUBTYPE',
      extraDropdowns: []
    }),
    logAutomationEvent_: (...args) => events.push(args),
    getErrorDetails_: e => String(e)
  });
  vm.runInContext(source, context);
  return { context, values, events, calls };
}

function payload(events) {
  const marker = events.find(event => event[1] === 'SCC Handoff')?.[4]?.split('SCC_HANDOFF_V1:')[1];
  return marker ? JSON.parse(Buffer.from(marker, 'base64url').toString()) : null;
}

test('logging a successful call creates a handoff without saving the roster', () => {
  const env = fixture();
  env.context.logSccInPowerSchool();
  assert.equal(env.values.rosterNote, '');
  assert.equal(env.values.rosterToDo, '');
  assert.equal(env.values.completion, '');
  assert.equal(payload(env.events).note, env.values.note);
  assert.equal(payload(env.events).studentNumber, '12345678');
  assert.equal(payload(env.events).outcome, 'SCC Success');
  assert.equal(payload(env.events).settings.subtypeValue, 'SUCCESS_SUBTYPE');
});

test('previously saved call can be logged again without changing the roster', () => {
  const env = fixture({ duplicate: true });
  env.context.logSccInPowerSchool();
  assert.equal(payload(env.events).note, env.values.note);
});

test('existing save-only action never opens PowerSchool', () => {
  const env = fixture();
  env.context.saveSccToRoster();
  assert.equal(env.values.rosterNote, env.values.note);
  assert.equal(payload(env.events), null);
});

test('invalid contact does not create a handoff, but logging does not depend on Attempt columns', () => {
  const invalid = fixture({ yes: false, no: false });
  invalid.context.logSccInPowerSchool();
  assert.equal(payload(invalid.events), null);
  const full = fixture({ yes: false, no: true, attemptSaved: false });
  full.context.logSccInPowerSchool();
  assert.equal(payload(full.events).outcome, 'SCC Attempt');
  const attempt = fixture({ yes: false, no: true });
  attempt.context.logSccInPowerSchool();
  assert.equal(payload(attempt.events).note, attempt.values.note);
  assert.equal(payload(attempt.events).outcome, 'SCC Attempt');
  assert.equal(payload(attempt.events).settings.typeValue, 'ATTEMPT_TYPE');
});

test('legacy combined-action name now logs only', () => {
  const env = fixture();
  env.context.saveSccAndOpenPowerSchool();
  assert.equal(env.values.rosterNote, '');
  assert.equal(payload(env.events).outcome, 'SCC Success');
});
