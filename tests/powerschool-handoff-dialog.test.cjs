'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(
  path.join(__dirname, '..', 'PowerSchoolHandoffDialog.gs'),
  'utf8'
);

function fixture() {
  const dialogs = [];
  const context = vm.createContext({
    Utilities: {
      Charset: { UTF_8: 'UTF-8' },
      base64EncodeWebSafe: data => Buffer.from(data).toString('base64url')
    },
    HtmlService: {
      createHtmlOutput(html) {
        return {
          html,
          width: null,
          height: null,
          setWidth(value) { this.width = value; return this; },
          setHeight(value) { this.height = value; return this; }
        };
      }
    },
    SpreadsheetApp: {
      getUi: () => ({
        showModalDialog: (output, title) => dialogs.push({ output, title })
      })
    }
  });
  vm.runInContext(source, context);
  return { context, dialogs };
}

test('a supported handoff displays one safe PowerSchool dialog', () => {
  const env = fixture();
  const payload = { v: 1, requestId: 'request-1', studentNumber: '12345678' };
  const url = env.context.showPowerSchoolHandoffDialog_(
    'demographics', payload, 'Open PowerSchool Demographics'
  );
  assert.equal(env.dialogs.length, 1);
  assert.match(url, /^https:\/\/californiak12\.powerschool\.com\/teachers\/home\.html#demographics=/);
  assert.equal(new URL(url).origin, 'https://californiak12.powerschool.com');
  const encoded = url.split('#demographics=')[1];
  assert.deepEqual(JSON.parse(Buffer.from(encoded, 'base64url').toString()), payload);
  const html = env.dialogs[0].output.html;
  assert.match(html, /target="_blank"/);
  assert.match(html, /rel="noopener noreferrer"/);
  assert.match(html, />Open PowerSchool<\/a>/);
  assert.match(html, />Cancel<\/button>/);
  assert.match(html, /submit it manually/i);
  assert.equal(html.replace(/<[^>]+>/g, '').includes(encoded), false);
});

test('only scc, ecc, and demographics handoff types are accepted', () => {
  const env = fixture();
  const encoded = env.context.encodePowerSchoolHandoffPayload_({ requestId: 'request-1' });
  for (const type of ['scc', 'ecc', 'demographics']) {
    assert.match(env.context.buildPowerSchoolHandoffUrl_(type, encoded),
      new RegExp('#' + type + '='));
  }
  assert.throws(() => env.context.buildPowerSchoolHandoffUrl_('other', encoded),
    /Unsupported/);
  assert.throws(() => env.context.buildPowerSchoolHandoffUrl_('scc', 'bad payload'),
    /invalid/);
  assert.throws(() => env.context.buildPowerSchoolHandoffDialogHtml_(
    'https://example.com/teachers/home.html#scc=' + encoded), /invalid/);
  assert.equal(env.dialogs.length, 0);
});
