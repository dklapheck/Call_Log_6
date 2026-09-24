'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function table(headers, rows) {
  return { headers, rows, headerMap: Object.fromEntries(headers.map((h, i) => [h, i])) };
}
function sheet(headers, rows) {
  const cells = [headers.slice(), ...rows.map(r => r.slice())];
  return {
    cells, getLastColumn: () => headers.length, getLastRow: () => cells.length,
    getRange(row, col, height = 1, width = 1) {
      return {
        getDisplayValues: () => Array.from({length: height}, (_, r) =>
          Array.from({length: width}, (_, c) => String(cells[row + r - 1]?.[col + c - 1] ?? ''))),
        setValue(value) { while (cells.length < row) cells.push([]); cells[row - 1][col - 1] = value; },
        copyTo() {}
      };
    }
  };
}
function environment() {
  const context = vm.createContext({
    normalizeId_: value => String(value ?? '').trim().replace(/\.0$/, ''),
    getHeaderIndexMap_: headers => table(headers, []).headerMap,
    readSheetTable_: s => table(s.cells[0], s.cells.slice(1)),
    ensureSheetRows_() {},
    SpreadsheetApp: {CopyPasteType: {PASTE_FORMAT: 'format'}}
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'RefreshStudentData.gs'), 'utf8'), context);
  return context;
}
const headers = ['', 'Student Number', 'LAST NAME', 'FIRST NAME', 'GRADE', 'START DATE', 'SCHOOL', 'SPED'];
const rows = [
  ['Teacher', 12345678, 'Example', 'Morgan', 11, '2026-09-28', 'School A', ''],
  ['Teacher', 12345679, 'Sample', 'Alex', 11, '2026-09-28', 'School A', '']
];
const empty = table([], []);
const sccHeaders = ['Student Number', 'LAST NAME', 'FIRST NAME', 'GRADE', 'START DATE', 'SCHOOL',
  'SCC Notes', 'Attempt 1', 'SCC Completion', 'Student Email', 'Prefered Name'];
function build(context, roster, scc = sheet(sccHeaders, []), upstream = empty, addDrop = empty) {
  return context.refreshBuildStudents_(upstream, empty, addDrop, scc, {student: () => null}, roster);
}

test('roster columns are matched by header even when columns are reordered', () => {
  const c = environment();
  const normal = c.refreshClassRosterIndex_(table(headers, rows));
  const reordered = c.refreshClassRosterIndex_(table([...headers].reverse(), rows.map(r => [...r].reverse())));
  assert.equal(JSON.stringify(normal), JSON.stringify(reordered));
  assert.equal(normal['12345678']['FIRST NAME'], 'Morgan');
  assert.equal(normal['12345678']['SCHOOL'], 'School A');
});

test('mixed layouts, malformed IDs, missing headers, and conflicting duplicates stop the import', () => {
  const c = environment();
  assert.throws(() => c.refreshClassRosterIndex_(table(headers, [rows[0], ['Teacher', ...rows[1]]])), /columns/);
  assert.throws(() => c.refreshClassRosterIndex_(table(headers, rows.map(r => ['Teacher', ...r]))), /columns/);
  assert.throws(() => c.refreshClassRosterIndex_(table(headers, [[...rows[0].slice(0,1), 'not-an-id', ...rows[0].slice(2)]])), /columns/);
  assert.throws(() => c.refreshClassRosterIndex_(table(headers.filter(h => h !== 'Student Number'), [])), /header/);
  const duplicate = rows[0].slice(); duplicate[3] = 'Different';
  assert.throws(() => c.refreshClassRosterIndex_(table(headers, [rows[0], duplicate])), /conflicting duplicate/);
});

test('new Class Roster students reach SCC without Upstream and repeated refreshes preserve notes', () => {
  const c = environment();
  const roster = c.refreshClassRosterIndex_(table(headers, rows));
  const scc = sheet(sccHeaders, [[12345670, 'Prior', 'Student', 10, '', 'School A', 'Keep my note', 'Keep attempt', 'Completed']]);
  const data = build(c, roster, scc);
  assert.equal(data.ids.length, 3);
  assert.equal(data.byId['12345678']['Current Active Student'], 'Y');
  assert.equal(data.byId['12345678']['Upstream Match'], 'N');
  assert.match(data.byId['12345678']['Source Status'], /awaiting Upstream/);
  c.refreshSyncScc_(scc, data.byId);
  assert.equal(scc.cells.length, 4);
  assert.equal(scc.cells[1][6], 'Keep my note');
  assert.equal(scc.cells[1][7], 'Keep attempt');
  assert.equal(scc.cells[1][8], 'Completed');
  assert.equal(scc.cells[2][2], 'Morgan');
  assert.equal(scc.cells[2][6] ?? '', '');
  scc.cells[2][6] = 'New student call';
  c.refreshSyncScc_(scc, build(c, roster, scc).byId);
  assert.equal(scc.cells.length, 4);
  assert.equal(scc.cells[2][6], 'New student call');
});

test('Upstream enrichment later updates the same student without duplicating or erasing notes', () => {
  const c = environment();
  const roster = c.refreshClassRosterIndex_(table(headers, rows));
  const scc = sheet(sccHeaders, []);
  c.refreshSyncScc_(scc, build(c, roster, scc).byId);
  scc.cells[1][6] = 'Preserve contact history';
  const upstream = table(['STUDENT_NUMBER', 'CURRENT_ACTIVE_STUDENT', 'STUDENT_LAST_NAME', 'STUDENT_FIRST_NAME', 'EMAIL_O365'],
    [[12345678, 'Y', 'Example', 'Morgan', 'student@example.invalid']]);
  const data = build(c, roster, scc, upstream);
  c.refreshSyncScc_(scc, data.byId);
  assert.equal(scc.cells.length, 3);
  assert.equal(scc.cells[1][6], 'Preserve contact history');
  assert.equal(scc.cells[1][9], 'student@example.invalid');
  assert.equal(data.byId['12345678']['Source Status'], 'Active Upstream');
});

test('explicit inactive or dropped students are not reactivated by a stale Class Roster', () => {
  const c = environment();
  const roster = c.refreshClassRosterIndex_(table(headers, rows));
  const upstream = table(['STUDENT_NUMBER', 'CURRENT_ACTIVE_STUDENT'], [[12345678, 'N']]);
  const addDrop = table(['STUDENT NUMBER', 'ADD / DROP'], [[12345679, 'DROP']]);
  const scc = sheet(sccHeaders, []);
  const data = build(c, roster, scc, upstream, addDrop);
  assert.equal(data.byId['12345678']['Current Active Student'], 'N');
  assert.equal(data.byId['12345679']['Current Active Student'], 'N');
  c.refreshSyncScc_(scc, data.byId);
  assert.equal(scc.cells.length, 1);
});

test('identical duplicate roster rows and blank rows do not duplicate students', () => {
  const c = environment();
  const index = c.refreshClassRosterIndex_(table(headers, [rows[0], rows[0], []]));
  assert.deepEqual(Object.keys(index), ['12345678']);
});

module.exports = {environment, table, sheet};
