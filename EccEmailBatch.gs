/**
 * Manual, two-part ECC email handoff. Preparing a batch never sends mail or
 * changes the ECC tab. The extension returns clicked-Send receipts for import.
 */
const ECC_EMAIL_BATCH_VERSION = 1;
const ECC_EMAIL_BATCH_LIMIT = 30;
const ECC_EMAIL_BATCH_PREFIX = 'ecc_email_batch_';

function prepareEccReminderEmails() { prepareEccEmailBatch_('reminder'); }
function prepareEccRescheduleEmails() { prepareEccEmailBatch_('reschedule'); }

function eccEmailColumnMap_(sheet, names) {
  const result = {};
  names.forEach(function(name) { result[name] = findHeaderColumn_(sheet, name) - 1; });
  return result;
}

function eccEmailAddress_(value, label) {
  const address = String(value || '').trim();
  if (!/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(address)) {
    throw new Error(label + ' must contain one valid email address.');
  }
  return address;
}

function eccEmailBatchData_(ecc, studentData, kind, spreadsheetId, makeId) {
  if (kind !== 'reminder' && kind !== 'reschedule') throw new Error('Unknown ECC email type.');
  const ec = eccEmailColumnMap_(ecc, [
    'Prefered Name', 'Student Number', 'Student Email', 'Email?',
    kind === 'reminder' ? 'Email Reminder' : 'Email Reschedule'
  ]);
  const sc = eccEmailColumnMap_(studentData, [
    'Student Number', 'Learning Coach', 'Learning Coach Email'
  ]);
  const eccValues = ecc.getDataRange().getDisplayValues();
  const eccRaw = ecc.getDataRange().getValues();
  const studentValues = studentData.getDataRange().getDisplayValues();
  const lookup = {};
  studentValues.slice(1).forEach(function(row) {
    const id = normalizeId_(row[sc['Student Number']]);
    if (!id) return;
    if (lookup[id]) lookup[id].duplicate = true;
    else lookup[id] = { row: row, duplicate: false };
  });

  const errors = [];
  const rows = [];
  const seen = {};
  for (let i = 1; i < eccValues.length; i++) {
    const rawCheck = eccRaw[i][ec['Email?']];
    if (rawCheck !== true && String(rawCheck).toLowerCase() !== 'true' && rawCheck !== 1) continue;
    const row = eccValues[i], rowNumber = i + 1;
    try {
      const studentNumber = normalizeId_(row[ec['Student Number']]);
      if (!studentNumber) throw new Error('Student Number is blank.');
      if (seen[studentNumber]) throw new Error('Student Number is selected more than once.');
      seen[studentNumber] = true;
      const preferredName = String(row[ec['Prefered Name']] || '').trim();
      if (!preferredName) throw new Error('Prefered Name is blank.');
      const rawBody = String(row[ec[kind === 'reminder' ? 'Email Reminder' : 'Email Reschedule']] || '');
      if (!rawBody.trim()) throw new Error('Email message is blank.');
      if (rawBody.length > 20000) throw new Error('Email message exceeds 20,000 characters.');
      const studentAddress = eccEmailAddress_(row[ec['Student Email']], 'Student Email');
      const match = lookup[studentNumber];
      if (match && match.duplicate) throw new Error('StudentData has duplicate Student Number rows.');
      const lcName = match ? String(match.row[sc['Learning Coach']] || '').trim() : '';
      const lcRaw = match ? String(match.row[sc['Learning Coach Email']] || '').trim() : '';
      const recipients = [studentAddress];
      const warnings = [];
      if (lcRaw) {
        const coachAddress = eccEmailAddress_(lcRaw, 'Learning Coach Email');
        if (coachAddress.toLowerCase() !== studentAddress.toLowerCase()) recipients.push(coachAddress);
      } else {
        warnings.push('No Learning Coach email in StudentData; student only.');
      }
      const body = rawBody
        .replace(/\[Student Preferred First Name\]/g, preferredName.split(/\s+/)[0])
        .replace(/\[LC\]/g, lcName || 'Learning Coach');
      rows.push({
        requestId: makeId(), rowNumber: rowNumber, studentNumber: studentNumber,
        preferredName: preferredName, to: recipients,
        subject: kind === 'reminder'
          ? 'ECC check-in reminder — ' + preferredName
          : 'ECC reschedule request — ' + preferredName,
        body: body, warnings: warnings
      });
    } catch (error) {
      errors.push('ECC row ' + rowNumber + ': ' + error.message);
    }
  }
  if (errors.length) throw new Error('Fix the selected ECC rows before preparing mail:\n' + errors.join('\n'));
  if (!rows.length) throw new Error('No ECC rows have Email? checked.');
  if (rows.length > ECC_EMAIL_BATCH_LIMIT) {
    throw new Error('Select at most ' + ECC_EMAIL_BATCH_LIMIT + ' rows for one batch.');
  }
  return {
    version: ECC_EMAIL_BATCH_VERSION, batchId: makeId(), spreadsheetId: spreadsheetId,
    kind: kind, issuedAt: new Date().toISOString(), rows: rows
  };
}

function eccEmailEscapeHtml_(value) {
  return String(value).replace(/[&<>"']/g, function(c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function prepareEccEmailBatch_(kind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const batch = eccEmailBatchData_(
    getRequiredSheet_(ss, APP_CONFIG.sheets.ecc),
    getRequiredSheet_(ss, APP_CONFIG.sheets.studentData),
    kind, ss.getId(), function() { return Utilities.getUuid(); }
  );
  const manifest = {
    kind: batch.kind, issuedAt: batch.issuedAt,
    rows: batch.rows.map(function(r) {
      return { requestId: r.requestId, studentNumber: r.studentNumber };
    })
  };
  PropertiesService.getDocumentProperties().setProperty(
    ECC_EMAIL_BATCH_PREFIX + batch.batchId, JSON.stringify(manifest)
  );
  const json = eccEmailEscapeHtml_(JSON.stringify(batch, null, 2));
  const html = '<div style="font:14px Arial;padding:16px">' +
    '<h2>ECC ' + kind + ' batch: ' + batch.rows.length + ' messages</h2>' +
    '<p>Review the recipients in the extension before sending. Copy this batch into the extension. ' +
    'This preparation has sent no email and changed no ECC cells.</p>' +
    '<textarea id="batch" readonly style="width:100%;height:330px;box-sizing:border-box">' +
    json + '</textarea><p><button onclick="copyBatch()">Copy batch JSON</button> ' +
    '<span id="status"></span></p>' +
    '<script>function copyBatch(){var el=document.getElementById("batch");el.select();' +
    'var copied=document.execCommand("copy");document.getElementById("status").textContent=' +
    'copied?"Copied. Paste into the ECC Outlook Mailer extension.":' +
    '"Copy was blocked. Select the JSON and press Ctrl+C (or Cmd+C).";}</script></div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(650).setHeight(500), 'Prepare ECC email batch'
  );
}

function showEccEmailReceiptDialog() {
  const html = '<div style="font:14px Arial;padding:16px"><h2>Record ECC email attempts</h2>' +
    '<p>Paste the extension receipt after checking Outlook Sent Items. Only entries where Send was clicked are recorded.</p>' +
    '<textarea id="receipt" style="width:100%;height:250px;box-sizing:border-box"></textarea>' +
    '<p><button id="record" onclick="record()">Record attempts</button> <span id="status"></span></p>' +
    '<script>function record(){var b=document.getElementById("record");b.disabled=true;' +
    'var s=document.getElementById("status");s.textContent="Checking receipt…";' +
    'google.script.run.withSuccessHandler(function(x){s.textContent=x;})' +
    '.withFailureHandler(function(e){s.textContent=e.message;b.disabled=false;})' +
    '.recordEccEmailReceipts(document.getElementById("receipt").value);}</script></div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(650).setHeight(400), 'Record ECC email receipts'
  );
}

function eccEmailReceiptUpdates_(receipt, manifest, ecc) {
  if (!receipt || receipt.version !== ECC_EMAIL_BATCH_VERSION ||
      !Array.isArray(receipt.clicked) || !receipt.clicked.length ||
      receipt.kind !== manifest.kind) throw new Error('Invalid or empty ECC email receipt.');
  const ids = eccEmailColumnMap_(ecc, ['Student Number', 'Attempts']);
  const data = ecc.getDataRange().getDisplayValues();
  const rowsById = {};
  data.slice(1).forEach(function(row, i) {
    const id = normalizeId_(row[ids['Student Number']]);
    if (!id) return;
    if (rowsById[id]) rowsById[id].duplicate = true;
    else rowsById[id] = { rowNumber: i + 2, attempts: row[ids.Attempts] };
  });
  const allowed = {};
  manifest.rows.forEach(function(r) { allowed[r.requestId] = r.studentNumber; });
  const seen = {};
  const updates = [];
  receipt.clicked.forEach(function(item) {
    if (!item || typeof item.requestId !== 'string' || seen[item.requestId] ||
        allowed[item.requestId] !== normalizeId_(item.studentNumber) ||
        !/^\d{4}-\d{2}-\d{2}T/.test(String(item.clickedAt)) ||
        !isFinite(Date.parse(item.clickedAt))) {
      throw new Error('Receipt contains an unknown, duplicate, or invalid send record.');
    }
    seen[item.requestId] = true;
    const target = rowsById[normalizeId_(item.studentNumber)];
    if (!target || target.duplicate) throw new Error('Student Number is missing or duplicated in ECC.');
    const marker = '[Email ID: ' + item.requestId + ']';
    if (target.attempts.indexOf(marker) !== -1) return;
    const next = ((target.attempts.match(/-- Email attempt\b/gi) || []).length + 1);
    const when = Utilities.formatDate(new Date(item.clickedAt),
      SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'M/d/yyyy h:mm a');
    const entry = when + ' -- Email attempt ' + next + ' (' + receipt.kind +
      '; Send clicked) ' + marker;
    updates.push({
      rowNumber: target.rowNumber, column: ids.Attempts + 1,
      value: appendHistoryEntry_(target.attempts, entry)
    });
  });
  return updates;
}

function recordEccEmailReceipts(json) {
  return withRosterLock_(function() {
    const receipt = JSON.parse(json);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (receipt.spreadsheetId !== ss.getId() ||
        !/^[0-9a-f-]{36}$/i.test(String(receipt.batchId))) {
      throw new Error('Receipt does not belong to this spreadsheet or has an invalid batch ID.');
    }
    const raw = PropertiesService.getDocumentProperties().getProperty(
      ECC_EMAIL_BATCH_PREFIX + receipt.batchId
    );
    if (!raw) throw new Error('Batch was not prepared by this spreadsheet.');
    const manifest = JSON.parse(raw);
    const ecc = getRequiredSheet_(ss, APP_CONFIG.sheets.ecc);
    const updates = eccEmailReceiptUpdates_(receipt, manifest, ecc);
    updates.forEach(function(item) {
      ecc.getRange(item.rowNumber, item.column).setValue(item.value);
    });
    return updates.length + ' attempt(s) appended to ECC. ' +
      (receipt.clicked.length - updates.length) + ' duplicate receipt(s) skipped. Email? was unchanged.';
  });
}
