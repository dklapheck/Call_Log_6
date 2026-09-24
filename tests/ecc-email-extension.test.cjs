'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const script = fs.readFileSync(path.join(__dirname, '..', 'ecc-email-extension', 'mailer.js'), 'utf8');
const elements = Object.fromEntries(['status', 'review', 'send', 'stop', 'copy', 'batch',
  'preview', 'result', 'receipt', 'saved', 'clear'].map(id => [id, {
    textContent: '', value: '', hidden: false, disabled: false,
    addEventListener() {}, replaceChildren() {}
  }]));
const context = vm.createContext({
  document: { getElementById: id => elements[id] },
  chrome: { storage: { local: { get: async () => ({}) } } },
  URL, setTimeout, clearTimeout
});
vm.runInContext(script, context, { filename: 'mailer.js' });

test('Outlook compose deeplink percent-encodes spaces and two recipients', () => {
  const url = context.composeUrl('https://outlook.cloud.microsoft', {
    to: ['student@example.edu', 'coach@example.edu'],
    subject: 'ECC reminder — Alex', body: 'Hello Alex\nPlease reply'
  });
  assert.match(url, /to=student%40example.edu%3Bcoach%40example.edu/);
  assert.match(url, /subject=ECC%20reminder/);
  assert.match(url, /body=Hello%20Alex%0APlease%20reply/);
  assert.doesNotMatch(url, /\+/);
});

test('batch validation rejects duplicate To recipients and oversized batches', () => {
  const data = {
    version: 1, batchId: '12345678-1234-1234-1234-123456789abc',
    spreadsheetId: 'sheet', kind: 'reminder',
    rows: [{
      requestId: '22345678-1234-1234-1234-123456789abc',
      rowNumber: 2, studentNumber: '11', to: ['a@example.edu', 'A@example.edu'],
      subject: 'Hello', body: 'Body'
    }]
  };
  assert.throws(() => context.validateBatch(JSON.stringify(data)), /duplicate To/);
  data.rows[0].to.pop();
  assert.equal(context.validateBatch(JSON.stringify(data)).rows.length, 1);
  data.rows = Array(31).fill(data.rows[0]);
  assert.throws(() => context.validateBatch(JSON.stringify(data)), /message count/);
});
