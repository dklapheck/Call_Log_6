function refreshStudentDataAndCounselors() {
  refreshStudentData();
  refreshCounselorData();
}

function refreshCounselorData() {
  return withRosterLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const source = getSourceSpreadsheet_();
    const upstreamSheet = getRequiredSheet_(source, APP_CONFIG.source.upstream);
    const rosterSheet = getRequiredSheet_(ss, SCC_CONFIG.sheets.roster);

    const upstream = readSheetTable_(upstreamSheet);
    const upstreamIdIndex = upstream.headerMap[SCC_CONFIG.upstream.studentIdHeader];
    const upstreamCounselorIndex = upstream.headerMap[SCC_CONFIG.upstream.counselorHeader];
    if (typeof upstreamIdIndex === 'undefined' || typeof upstreamCounselorIndex === 'undefined') {
      throw new Error('Upstream must include STUDENT_NUMBER and COUNSELOR columns.');
    }

    const counselorByStudentId = {};
    upstream.rows.forEach(function(row) {
      const studentId = normalizeId_(row[upstreamIdIndex]);
      const counselor = String(row[upstreamCounselorIndex] || '').trim();
      if (studentId && counselor) counselorByStudentId[studentId] = counselor;
    });

    const rosterIdColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.studentIdHeader);
    const counselorColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.counselorHeader);
    const emailColumn = findHeaderColumn_(rosterSheet, SCC_CONFIG.roster.counselorEmailHeader);
    const lastRow = rosterSheet.getLastRow();
    if (lastRow < 2) return;

    const rowCount = lastRow - 1;
    const studentIds = rosterSheet.getRange(2, rosterIdColumn, rowCount, 1).getDisplayValues();
    const existing = rosterSheet.getRange(2, counselorColumn, rowCount, 2).getDisplayValues();
    const emailByCounselor = {};
    existing.forEach(function(row) {
      const counselor = String(row[0] || '').trim();
      const email = String(row[1] || '').trim();
      if (counselor && email) emailByCounselor[counselor] = email;
    });

    let namesUpdated = 0;
    let emailsUpdated = 0;
    const output = existing.map(function(row, index) {
      const existingCounselor = String(row[0] || '').trim();
      const existingEmail = String(row[1] || '').trim();
      const upstreamCounselor = counselorByStudentId[normalizeId_(studentIds[index][0])] || '';
      if (!upstreamCounselor) return [existingCounselor, existingEmail];

      const counselorEmail = emailByCounselor[upstreamCounselor] ||
        (existingCounselor === upstreamCounselor ? existingEmail : '');
      if (upstreamCounselor !== existingCounselor) namesUpdated++;
      if (counselorEmail !== existingEmail) emailsUpdated++;
      return [upstreamCounselor, counselorEmail];
    });

    rosterSheet.getRange(2, counselorColumn, rowCount, 2).setValues(output);
    SpreadsheetApp.flush();
    ss.toast(
      'Counselor data refreshed. Names updated: ' + namesUpdated +
        '; emails updated: ' + emailsUpdated + '.',
      'Counselor Data',
      5
    );
  });
}
