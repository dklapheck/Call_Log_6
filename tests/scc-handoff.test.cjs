'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'SaveSCC.gs'), 'utf8');

function fixture({ yes = true, no = false, attemptSaved = true, duplicate = false, attempts = [] } = {}) {
  const values = {
    studentId: 12345678,
    note: 'On 9/18/2026, spoke with parent.',
    toDo: 'Send resources.',
    yes, no,
    rosterNote: duplicate ? 'On 9/18/2026, spoke with parent.' : '',
    completion: duplicate ? 'Completed' : '',
    rosterToDo: '',
    attempt1: attempts[0] || '', attempt2: attempts[1] || '',
    attempt3: attempts[2] || '', attempt4: attempts[3] || '',
    attempt5: attempts[4] || ''
  };
  const events = [];
  const calls = [];
  const dialogs = [];
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
    return makeCell(({
      1: 'rosterNote', 2: 'rosterToDo', 3: 'completion',
      4: 'attempt1', 5: 'attempt2', 6: 'attempt3', 7: 'attempt4', 8: 'attempt5'
    })[col]);
  } };
  const ss = { toast: (...args) => calls.push(args) };
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    Utilities: {
      Charset: { UTF_8: 'UTF-8' },
      getUuid: () => 'request-123',
      base64EncodeWebSafe: data => Buffer.from(data).toString('base64url')
    },
    SCC_CONFIG: {
      sheets: { callEntry: 'Call Entry', roster: 'SCC' },
      callEntry: { studentId: 'O2', note: 'D1', toDo: 'I1' },
      roster: {
        notesHeader: 'SCC Notes',
        toDoHeader: 'SCC To Do',
        completionHeader: 'SCC Completion',
        attemptHeaders: ['Attempt 1', 'Attempt 2', 'Attempt 3', 'Attempt 4', 'Attempt 5'],
        completedValue: 'Completed'
      }
    },
    withRosterLock_: fn => fn(),
    getRequiredSheet_: (_ss, name) => name === 'Call Entry' ? callSheet : rosterSheet,
    findStudentRosterRow_: () => 2,
    getStudentName_: () => 'Student',
    findQuestionRow_: () => 14,
    findHeaderColumn_: (_sheet, name) => ({
      'SCC Notes': 1, 'SCC To Do': 2, 'SCC Completion': 3,
      'Attempt 1': 4, 'Attempt 2': 5, 'Attempt 3': 6,
      'Attempt 4': 7, 'Attempt 5': 8
    })[name],
    saveFailedSccAttempt_: () => ({ saved: attemptSaved }),
    normalizeId_: id => String(id),
    getPowerSchoolContactSettings_: (_ss, workflow) => ({
      workflow, typeValue: workflow === 'SCC Success' ? 'SUCCESS_TYPE' : 'ATTEMPT_TYPE',
      subtypeValue: workflow === 'SCC Success' ? 'SUCCESS_SUBTYPE' : 'ATTEMPT_SUBTYPE',
      extraDropdowns: [],
      dateFields: ['entryLogDate', 'UF-008019-1', 'UF-008041-1'],
      dateField: 'entryLogDate',
      tagMap: workflow === 'SCC Attempt'
        ? { 1: 'Attempt 1 (34)', 2: 'Attempt 2 (35)', 3: 'Attempt 3 (36)' } : {},
      tagLabel: ''
    }),
    showPowerSchoolHandoffDialog_: (type, handoffPayload, title) => {
      dialogs.push({ type, payload: handoffPayload, title });
    },
    logAutomationEvent_: (...args) => events.push(args),
    getErrorDetails_: e => String(e)
  });
  vm.runInContext(source, context);
  return { context, values, events, calls, dialogs };
}

function payload(env) {
  return env.dialogs[0]?.payload || null;
}

test('logging a successful call creates a handoff without saving the roster', () => {
  const env = fixture();
  env.context.logSccInPowerSchool();
  assert.equal(env.values.rosterNote, '');
  assert.equal(env.values.rosterToDo, '');
  assert.equal(env.values.completion, '');
  assert.equal(env.dialogs.length, 1);
  assert.equal(env.dialogs[0].type, 'scc');
  assert.equal(payload(env).note, env.values.note);
  assert.equal(payload(env).studentNumber, '12345678');
  assert.equal(payload(env).requestId, 'request-123');
  assert.equal(payload(env).outcome, 'SCC Success');
  assert.equal(payload(env).date, '9/18/2026');
  assert.equal(payload(env).settings.subtypeValue, 'SUCCESS_SUBTYPE');
  assert.deepEqual(payload(env).settings.dateFields,
    ['entryLogDate', 'UF-008019-1', 'UF-008041-1']);
  assert.equal(env.calls.some(args => /SCC_HANDOFF_V1:/.test(String(args[0]))), false);
});

test('previously saved call can be logged again without changing the roster', () => {
  const env = fixture({ duplicate: true });
  env.context.logSccInPowerSchool();
  assert.equal(payload(env).note, env.values.note);
});

test('existing save-only action never opens PowerSchool', () => {
  const env = fixture();
  env.context.saveSccToRoster();
  assert.equal(env.values.rosterNote, env.values.note);
  assert.equal(payload(env), null);
});

test('invalid contact does not create a handoff and a new attempt uses the first open number', () => {
  const invalid = fixture({ yes: false, no: false });
  invalid.context.logSccInPowerSchool();
  assert.equal(payload(invalid), null);
  const full = fixture({ yes: false, no: true, attemptSaved: false });
  full.context.logSccInPowerSchool();
  assert.equal(payload(full).outcome, 'SCC Attempt');
  assert.equal(payload(full).attemptNumber, 1);
  const attempt = fixture({ yes: false, no: true });
  attempt.context.logSccInPowerSchool();
  assert.equal(payload(attempt).note, attempt.values.note);
  assert.equal(payload(attempt).outcome, 'SCC Attempt');
  assert.equal(payload(attempt).settings.typeValue, 'ATTEMPT_TYPE');
});

test('an unsaved third attempt is tagged Attempt 3 without saving the roster', () => {
  const env = fixture({ yes: false, no: true, attempts: ['older note 1', 'older note 2'] });
  env.context.logSccInPowerSchool();
  assert.equal(payload(env).attemptNumber, 3);
  assert.equal(env.values.attempt3, '');
});

test('a previously saved attempt keeps its existing attempt number', () => {
  const note = 'On 9/18/2026, spoke with parent.';
  const env = fixture({ yes: false, no: true, attempts: ['older note', note] });
  env.context.logSccInPowerSchool();
  assert.equal(payload(env).attemptNumber, 2);
});

test('legacy combined-action name now logs only', () => {
  const env = fixture();
  env.context.saveSccAndOpenPowerSchool();
  assert.equal(env.values.rosterNote, '');
  assert.equal(payload(env).outcome, 'SCC Success');
});
