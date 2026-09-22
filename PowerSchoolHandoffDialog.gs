const POWERSCHOOL_HANDOFF_HOST_ = 'https://californiak12.powerschool.com';
const POWERSCHOOL_HANDOFF_PATH_ = '/teachers/home.html';
const POWERSCHOOL_HANDOFF_TYPES_ = Object.freeze({
  scc: true,
  ecc: true,
  demographics: true
});

function encodePowerSchoolHandoffPayload_(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    throw new Error('The PowerSchool handoff payload is invalid.');
  }
  return Utilities.base64EncodeWebSafe(
    JSON.stringify(payload),
    Utilities.Charset.UTF_8
  ).replace(/=+$/g, '');
}

function buildPowerSchoolHandoffUrl_(type, encodedPayload) {
  const handoffType = String(type || '').toLowerCase();
  if (!POWERSCHOOL_HANDOFF_TYPES_[handoffType]) {
    throw new Error('Unsupported PowerSchool handoff type.');
  }
  if (!/^[A-Za-z0-9_-]+$/.test(String(encodedPayload || ''))) {
    throw new Error('The PowerSchool handoff data is invalid.');
  }
  return POWERSCHOOL_HANDOFF_HOST_ + POWERSCHOOL_HANDOFF_PATH_ +
    '#' + handoffType + '=' + encodedPayload;
}

function escapePowerSchoolDialogAttribute_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function buildPowerSchoolHandoffDialogHtml_(url) {
  const allowedUrl = new RegExp(
    '^https://californiak12\\.powerschool\\.com/teachers/home\\.html#' +
    '(scc|ecc|demographics)=[A-Za-z0-9_-]+$'
  );
  if (!allowedUrl.test(String(url))) {
    throw new Error('The PowerSchool handoff URL is invalid.');
  }
  const safeUrl = escapePowerSchoolDialogAttribute_(url);
  return '<!doctype html><html><head><base target="_blank">' +
    '<meta charset="utf-8"><style>' +
    'body{font:14px Arial,sans-serif;color:#202124;margin:20px;line-height:1.45}' +
    '.actions{display:flex;gap:10px;align-items:center;margin-top:18px}' +
    '.open{background:#174ea6;color:#fff;text-decoration:none;padding:10px 16px;' +
    'border-radius:4px;font-weight:700}.close{padding:9px 14px;cursor:pointer}' +
    '</style></head><body>' +
    '<p>Open PowerSchool to prepare this entry. Review every field, then submit it manually.</p>' +
    '<p>This tool will never click <strong>Submit</strong> for you.</p>' +
    '<div class="actions"><a id="open-powerschool" class="open" href="' + safeUrl +
    '" target="_blank" rel="noopener noreferrer" onclick="return openOnce(event)">' +
    'Open PowerSchool</a>' +
    '<button class="close" type="button" onclick="google.script.host.close()">Cancel</button></div>' +
    '<script>let opened=false;function openOnce(event){if(opened){event.preventDefault();return false;}' +
    'opened=true;const link=event.currentTarget;link.textContent="Opening PowerSchool…";' +
    'link.setAttribute("aria-disabled","true");link.style.pointerEvents="none";' +
    'setTimeout(function(){google.script.host.close();},500);return true;}</script>' +
    '</body></html>';
}

function showPowerSchoolHandoffDialog_(type, payload, title) {
  const encodedPayload = encodePowerSchoolHandoffPayload_(payload);
  const url = buildPowerSchoolHandoffUrl_(type, encodedPayload);
  const output = HtmlService.createHtmlOutput(
    buildPowerSchoolHandoffDialogHtml_(url)
  ).setWidth(430).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(output, title || 'Open PowerSchool');
  return url;
}
