const HOSTS = new Set(['outlook.cloud.microsoft', 'outlook.office.com', 'outlook.office365.com']);
const STORE_KEY = 'eccMailRuns';
let batch = null;
let sending = false;
let stopRequested = false;

const $ = id => document.getElementById(id);
const status = message => { $('status').textContent = message; };

function validateBatch(input) {
  const data = JSON.parse(input);
  if (data.version !== 1 || !/^[0-9a-f-]{36}$/i.test(data.batchId) ||
      typeof data.spreadsheetId !== 'string' || !data.spreadsheetId ||
      !['reminder', 'reschedule'].includes(data.kind) ||
      !Array.isArray(data.rows) || data.rows.length < 1 || data.rows.length > 30) {
    throw new Error('Invalid ECC batch header or message count.');
  }
  const seen = new Set();
  for (const row of data.rows) {
    if (!/^[0-9a-f-]{36}$/i.test(row.requestId) || seen.has(row.requestId) ||
        !Number.isInteger(row.rowNumber) || row.rowNumber < 2 ||
        !String(row.studentNumber || '').trim() ||
        !Array.isArray(row.to) || row.to.length < 1 || row.to.length > 2 ||
        row.to.some(address => typeof address !== 'string' ||
          !/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(address)) ||
        typeof row.subject !== 'string' || !row.subject.trim() ||
        typeof row.body !== 'string' || !row.body.trim() ||
        row.subject.length > 500 || row.body.length > 20000) {
      throw new Error('A message has missing, duplicate, or invalid fields.');
    }
    if (new Set(row.to.map(x => x.toLowerCase())).size !== row.to.length) {
      throw new Error('A message contains a duplicate To recipient.');
    }
    seen.add(row.requestId);
  }
  return data;
}

// URLSearchParams uses '+' for spaces. Outlook displayed that '+' literally in
// the self-email test; encodeURIComponent supplies '%20' in each compose field.
function composeUrl(origin, row) {
  return origin + '/mail/deeplink/compose?to=' +
    encodeURIComponent(row.to.join(';')) +
    '&subject=' + encodeURIComponent(row.subject) +
    '&body=' + encodeURIComponent(row.body);
}

async function getRuns() {
  return (await chrome.storage.local.get(STORE_KEY))[STORE_KEY] || {};
}
async function saveRun(run) {
  const runs = await getRuns();
  runs[run.batchId] = run;
  await chrome.storage.local.set({ [STORE_KEY]: runs });
}

function showReceipt(run) {
  if (!run?.clicked?.length) return;
  $('result').hidden = false;
  $('receipt').value = JSON.stringify({
    version: 1, batchId: run.batchId, spreadsheetId: run.spreadsheetId,
    kind: run.kind, clicked: run.clicked
  }, null, 2);
}

async function refreshSavedReceipts(preferredId) {
  const runs = await getRuns();
  const saved = Object.values(runs).filter(run => run.clicked.length);
  const select = $('saved');
  select.replaceChildren();
  for (const run of saved) {
    const option = document.createElement('option');
    option.value = run.batchId;
    option.textContent = run.kind + ' — ' + run.clicked.length + ' Send click(s) — ' + run.batchId;
    select.append(option);
  }
  const chosen = runs[preferredId]?.clicked.length ? runs[preferredId] : saved.at(-1);
  $('result').hidden = !chosen;
  if (chosen) { select.value = chosen.batchId; showReceipt(chosen); }
}

function review(data) {
  const target = $('preview');
  target.replaceChildren();
  data.rows.forEach((row, index) => {
    const card = document.createElement('article');
    const heading = document.createElement('h3');
    heading.textContent = (index + 1) + '. ' + row.preferredName + ' (ECC row ' + row.rowNumber + ')';
    const to = document.createElement('p');
    to.textContent = 'To: ' + row.to.join('; ');
    const subject = document.createElement('p');
    subject.textContent = 'Subject: ' + row.subject;
    const body = document.createElement('pre');
    body.textContent = row.body;
    card.append(heading, to, subject, body);
    for (const warning of row.warnings || []) {
      const note = document.createElement('p');
      note.className = 'warning';
      note.textContent = warning;
      card.append(note);
    }
    target.append(card);
  });
}

async function outlookOrigin() {
  const tabs = await chrome.tabs.query({});
  const tab = tabs.find(item => {
    try { return item.url && HOSTS.has(new URL(item.url).hostname); }
    catch { return false; }
  });
  if (!tab) throw new Error('Open a signed-in Outlook on the web tab first.');
  return new URL(tab.url).origin;
}

async function waitUntilLoaded(tabId) {
  const current = await chrome.tabs.get(tabId);
  if (current.status === 'complete') return;
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => { cleanup(); reject(new Error('Outlook tab did not finish loading.')); }, 45000);
    function changed(id, info) { if (id === tabId && info.status === 'complete') { cleanup(); resolve(); } }
    function removed(id) { if (id === tabId) { cleanup(); reject(new Error('Outlook tab was closed.')); } }
    function cleanup() {
      clearTimeout(timer); chrome.tabs.onUpdated.removeListener(changed);
      chrome.tabs.onRemoved.removeListener(removed);
    }
    chrome.tabs.onUpdated.addListener(changed);
    chrome.tabs.onRemoved.addListener(removed);
  });
}

async function sendOne(origin, row) {
  const tab = await chrome.tabs.create({ url: composeUrl(origin, row), active: true });
  await waitUntilLoaded(tab.id);
  const current = await chrome.tabs.get(tab.id);
  if (!current.url || !HOSTS.has(new URL(current.url).hostname)) {
    throw new Error('Outlook redirected away from the mail page. Check sign-in and the draft.');
  }
  const result = await chrome.scripting.executeScript({
    target: { tabId: tab.id }, func: verifyAndSend, args: [{
      to: row.to, subject: row.subject, body: row.body
    }]
  });
  return result?.[0]?.result || { clicked: false, message: 'No draft result. Check Outlook.' };
}

$('review').addEventListener('click', async () => {
  try {
    if (sending) return;
    batch = validateBatch($('batch').value);
    review(batch);
    const run = (await getRuns())[batch.batchId];
    await refreshSavedReceipts(batch.batchId);
    $('send').disabled = Boolean(run);
    status(run
      ? 'This batch has already been started. Check its Outlook drafts and receipt before preparing a new batch.'
      : 'Reviewed ' + batch.rows.length + ' message(s). Check every To, Subject, and Body, then click Send.');
  } catch (error) {
    batch = null; $('send').disabled = true; $('preview').replaceChildren();
    status(error.message);
  }
});

$('send').addEventListener('click', async () => {
  if (!batch || sending) return;
  if (!confirm('Send ' + batch.rows.length + ' reviewed ECC emails through Outlook?')) return;
  if (!navigator.locks) { status('This browser does not support the batch send lock.'); return; }
  await navigator.locks.request('ecc-mail-send', { ifAvailable: true }, async lock => {
    if (!lock) { status('Another ECC mail batch is already sending in this browser.'); return; }
    sending = true; stopRequested = false;
    $('send').disabled = true; $('review').disabled = true; $('stop').disabled = false;
    const run = {
      batchId: batch.batchId, spreadsheetId: batch.spreadsheetId,
      kind: batch.kind, clicked: [], inFlight: null
    };
    try {
      if ((await getRuns())[batch.batchId]) throw new Error('This batch has already been started.');
      const origin = await outlookOrigin();
      await saveRun(run);
      for (let i = 0; i < batch.rows.length; i++) {
        if (stopRequested) break;
        const row = batch.rows[i];
        run.inFlight = row.requestId;
        await saveRun(run); // Crash/unknown result blocks replay of this batch.
        status('Sending ' + (i + 1) + ' of ' + batch.rows.length + ': ' + row.preferredName);
        const result = await sendOne(origin, row);
        if (!result.clicked) throw new Error('Stopped on ECC row ' + row.rowNumber + ': ' + result.message);
        run.clicked.push({
          requestId: row.requestId, rowNumber: row.rowNumber,
          studentNumber: row.studentNumber, clickedAt: new Date().toISOString()
        });
        run.inFlight = null;
        await saveRun(run);
        await refreshSavedReceipts(run.batchId);
      }
      status('Stopped after ' + run.clicked.length + ' Send click(s). Check Sent Items and import the receipt.');
    } catch (error) {
      status(error.message + '\nCheck Outlook drafts and Sent Items before preparing another batch. ' +
        'Import any Send-click receipt shown below.');
    } finally {
      sending = false; $('review').disabled = false; $('stop').disabled = true;
      await refreshSavedReceipts(run.batchId);
    }
  });
});

$('stop').addEventListener('click', () => {
  stopRequested = true;
  status('Stopping after the current message.');
});
$('copy').addEventListener('click', async () => {
  await navigator.clipboard.writeText($('receipt').value);
  status('Receipt copied. Check Sent Items, then import it in the ECC Sheet.');
});
$('saved').addEventListener('change', async () => {
  showReceipt((await getRuns())[$('saved').value]);
});
$('clear').addEventListener('click', async () => {
  if (sending || !confirm('Delete the selected saved receipt? Do this only after importing it into ECC.')) return;
  const runs = await getRuns();
  const selectedId = $('saved').value;
  // Keep a body-free tombstone so clearing a receipt cannot replay the batch.
  runs[selectedId] = { batchId: selectedId, clicked: [], cleared: true };
  await chrome.storage.local.set({ [STORE_KEY]: runs });
  await refreshSavedReceipts();
  status('Selected saved receipt deleted. Check ECC before preparing another batch.');
});

// A reloaded manager never automatically resumes a batch. The prior receipt
// survives browser restarts in extension-local storage; no message bodies do.
(async () => {
  const runs = await getRuns();
  await refreshSavedReceipts();
  if (Object.keys(runs).length) {
    status('A prior batch was started. Review saved receipts and Sent Items before importing.');
  }
})().catch(error => status(error.message));

// Serialized by chrome.scripting and executed in the Outlook tab.
async function verifyAndSend(expected) {
  const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
  const visible = element => {
    const rect = element.getBoundingClientRect(), style = getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden';
  };
  const label = element =>
    (element.getAttribute('aria-label') || element.getAttribute('title') || element.innerText || '').trim();
  const sendButtons = root => Array.from(root.querySelectorAll('button, [role="button"]'))
    .filter(element => visible(element) && !element.disabled &&
      /^send(?:\b|$)/i.test(label(element)) &&
      !/^send\s+(later|options|schedule)/i.test(label(element)));
  const normalized = value => String(value || '').replace(/\r\n/g, '\n').trim();
  let subjectFields = [], subject, editors = [];
  for (let i = 0; i < 60; i++) {
    subjectFields = Array.from(document.querySelectorAll('input, [role="textbox"]'))
      .filter(element => visible(element) && /subject/i.test([
        element.getAttribute('aria-label'), element.getAttribute('placeholder'),
        element.getAttribute('name')
      ].filter(Boolean).join(' ')));
    subject = subjectFields.find(element =>
      normalized(element.value ?? element.innerText) === expected.subject);
    editors = Array.from(document.querySelectorAll('[contenteditable="true"]'))
      .filter(element => visible(element) &&
        normalized(element.innerText) === normalized(expected.body));
    if (subject && editors.length === 1) break;
    await pause(500);
  }
  if (subjectFields.length !== 1 || !subject) {
    return { clicked: false, message: 'Exactly one matching subject field was not found. Nothing was clicked.' };
  }
  if (editors.length !== 1) {
    return { clicked: false, message: 'The draft body did not match exactly. Nothing was clicked.' };
  }
  let compose = null;
  for (let parent = subject.parentElement; parent && parent !== document.body; parent = parent.parentElement) {
    if (parent.contains(editors[0]) && sendButtons(parent).length === 1) {
      compose = parent; break;
    }
  }
  if (!compose) return { clicked: false, message: 'A unique Send button was not tied to this draft.' };
  const emailsIn = root => {
    const values = [root.innerText || ''];
    for (const element of root.querySelectorAll('input, [title], [aria-label], [data-email]')) {
      for (const attribute of ['value', 'title', 'aria-label', 'data-email']) {
        const value = attribute === 'value' ? element.value : element.getAttribute(attribute);
        if (value) values.push(String(value));
      }
    }
    return [...new Set(values.flatMap(value =>
      value.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || []))]
      .map(email => email.toLowerCase());
  };
  const wanted = expected.to.map(email => email.toLowerCase()).sort();
  let recipientVerified = false;
  for (let attempt = 0; attempt < 30 && !recipientVerified; attempt++) {
    const toLabels = Array.from(compose.querySelectorAll('label, span, div, [aria-label]'))
      .filter(element => visible(element) && (
        /^to(?:\s+recipients)?\s*:?$/i.test((element.getAttribute('aria-label') || '').trim()) ||
        (element.childElementCount === 0 && /^to:?$/i.test((element.textContent || '').trim()))
      ));
    for (const toLabel of toLabels) {
      let parent = toLabel;
      for (let depth = 0; depth < 6 && parent && parent !== compose; depth++, parent = parent.parentElement) {
        if (parent.contains(subject) || parent.contains(editors[0])) break;
        const found = emailsIn(parent).sort();
        if (found.length === wanted.length && found.every((email, i) => email === wanted[i])) {
          recipientVerified = true;
        }
      }
    }
    if (!recipientVerified) await pause(500);
  }
  if (!recipientVerified) {
    return { clicked: false, message: 'The To row did not match the intended recipients. Nothing was clicked.' };
  }
  if (window.__eccEmailSendClicked) {
    return { clicked: false, message: 'Send was already clicked in this draft.' };
  }
  window.__eccEmailSendClicked = true;
  sendButtons(compose)[0].click();
  return { clicked: true, message: 'Send clicked. Check Sent Items for delivery status.' };
}
