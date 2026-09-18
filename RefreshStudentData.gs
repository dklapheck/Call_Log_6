function refreshStudentData() {
  return withRosterLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const source = getSourceSpreadsheet_();

    const upstream = readSheetTable_(getRequiredSheet_(source, APP_CONFIG.source.upstream));
    const proRoster = readSheetTable_(getRequiredSheet_(source, APP_CONFIG.source.proRoster));
    const courseGrades = readSheetTable_(getRequiredSheet_(source, APP_CONFIG.source.courseGrades));
    const addDrop = readSheetTable_(getRequiredSheet_(source, APP_CONFIG.source.addDrop));

    const studentDataSheet = getRequiredSheet_(ss, APP_CONFIG.sheets.studentData);
    const contactsSheet = getRequiredSheet_(ss, APP_CONFIG.sheets.contacts);
    const sccSheet = getRequiredSheet_(ss, APP_CONFIG.sheets.scc);
    const onboardingSheet = getRequiredSheet_(ss, APP_CONFIG.sheets.onboarding);
    const starSheet = getRequiredSheet_(ss, APP_CONFIG.sheets.star);
    const dashSheet = getRequiredSheet_(ss, APP_CONFIG.sheets.dash);
    const overridesSheet = getRequiredSheet_(ss, APP_CONFIG.sheets.overrides);
    const settingsSheet = getRequiredSheet_(ss, APP_CONFIG.sheets.settings);

    const overrides = refreshBuildOverrides_(overridesSheet);
    const data = refreshBuildStudents_(upstream, courseGrades, addDrop, sccSheet, overrides);
    writeRowsBelowHeader_(studentDataSheet, data.rows);

    const contacts = refreshBuildContacts_(proRoster, upstream, overrides);
    writeRowsBelowHeader_(contactsSheet, contacts);

    refreshSyncScc_(sccSheet, data.byId);
    refreshSyncOnboarding_(onboardingSheet, data, sccSheet);
    refreshSyncStar_(starSheet, data);
    refreshDash_(dashSheet, data, onboardingSheet, sccSheet, overridesSheet);

    const activeCol = findHeaderColumn_(overridesSheet, 'Active');
    const updatedCol = findHeaderColumn_(overridesSheet, 'Source Updated?');
    overridesSheet.getRange(2, activeCol, overridesSheet.getMaxRows() - 1, 1).insertCheckboxes();
    overridesSheet.getRange(2, updatedCol, overridesSheet.getMaxRows() - 1, 1).insertCheckboxes();

    settingsSheet.getRange(APP_CONFIG.source.lastRefreshCell)
      .setValue(new Date())
      .setNumberFormat('m/d/yyyy h:mm am/pm');

    SpreadsheetApp.flush();
    ss.toast('Refreshed ' + data.rows.length + ' students and ' + contacts.length + ' contacts.', 'Student Data Refreshed', 6);
  });
}

function refreshValue_(row, headerMap, header) {
  return typeof headerMap[header] === 'undefined' ? '' : row[headerMap[header]];
}

function refreshBuildOverrides_(sheet) {
  const table = readSheetTable_(sheet);
  const h = table.headerMap;
  const index = {};

  table.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, h, 'Student Number'));
    const person = String(refreshValue_(row, h, 'Person / Contact Key') || '').trim().toLowerCase();
    const field = String(refreshValue_(row, h, 'Field') || '').trim().toLowerCase();
    const activeValue = refreshValue_(row, h, 'Active');
    const updatedValue = refreshValue_(row, h, 'Source Updated?');
    const active = activeValue === true || String(activeValue).toUpperCase() === 'TRUE';
    const updated = updatedValue === true || String(updatedValue).toUpperCase() === 'TRUE';
    if (!id || !field || !active || updated) return;

    const received = refreshValue_(row, h, 'Date Received');
    const stamp = received instanceof Date ? received.getTime() : 0;
    const key = [id, person, field].join('|');
    if (!index[key] || stamp >= index[key].stamp) {
      index[key] = { value: refreshValue_(row, h, 'Override Value'), stamp: stamp };
    }
  });

  return {
    student: function(id, field) {
      const keys = [
        [id, 'student', field.toLowerCase()].join('|'),
        [id, '', field.toLowerCase()].join('|')
      ];
      for (let i = 0; i < keys.length; i++) if (index[keys[i]]) return index[keys[i]].value;
      return null;
    },
    contact: function(id, key, name, field) {
      const keys = [
        [id, String(key || '').toLowerCase(), field.toLowerCase()].join('|'),
        [id, String(name || '').toLowerCase(), field.toLowerCase()].join('|')
      ];
      for (let i = 0; i < keys.length; i++) if (index[keys[i]]) return index[keys[i]].value;
      return null;
    }
  };
}

function refreshBuildStudents_(upstream, courseGrades, addDrop, sccSheet, overrides) {
  const uh = upstream.headerMap;
  const ch = courseGrades.headerMap;
  const ah = addDrop.headerMap;
  const upstreamById = {};

  upstream.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, uh, 'STUDENT_NUMBER'));
    if (id) upstreamById[id] = row;
  });

  const cte = {};
  courseGrades.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, ch, 'STUDENT_IDENTITY_ID'));
    if (!id) return;
    const isCte = String(refreshValue_(row, ch, 'HS_COURSE_OFFERINGS_IS_CTE_COURSE')).toLowerCase() === 'yes' ||
      String(refreshValue_(row, ch, 'HS_COURSE_OFFERINGS_CONTENT_CATEGORY')) === 'CTE Courses';
    if (!isCte) return;
    if (!cte[id]) cte[id] = { any: false, fail: false };
    cte[id].any = true;
    const flag = refreshValue_(row, ch, 'CANVAS_PASSING_FLAG');
    if (flag === 0 || flag === false || String(flag).toLowerCase() === 'no') cte[id].fail = true;
  });

  const addLatest = {};
  addDrop.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, ah, 'STUDENT NUMBER'));
    if (!id) return;
    const when = refreshValue_(row, ah, 'DATE / TIME');
    const stamp = when instanceof Date ? when.getTime() : 0;
    if (!addLatest[id] || stamp >= addLatest[id].stamp) {
      addLatest[id] = { event: refreshValue_(row, ah, 'ADD / DROP'), date: when, stamp: stamp };
    }
  });

  const sccIndex = {};
  const sccTable = readSheetTable_(sccSheet);
  sccTable.rows.forEach(function(row, offset) {
    const id = normalizeId_(refreshValue_(row, sccTable.headerMap, 'Student Number'));
    if (id && !sccIndex[id]) sccIndex[id] = { row: row, rowNumber: offset + 2 };
  });

  const ids = [];
  upstream.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, uh, 'STUDENT_NUMBER'));
    const active = String(refreshValue_(row, uh, 'CURRENT_ACTIVE_STUDENT')).toUpperCase();
    if (id && ['Y','YES','TRUE'].indexOf(active) !== -1 && ids.indexOf(id) === -1) ids.push(id);
  });
  Object.keys(sccIndex).forEach(function(id) { if (ids.indexOf(id) === -1) ids.push(id); });

  const headers = [
    'Student Number','Source Last Name','Source First Name','Grade','Enroll Date','School',
    'Source Student Email','Learning Coach','Learning Coach Email','SPED/504','Attendance %','Chronically Absent',
    'STAR Math Date','STAR Reading Date','STAR Math Proficiency','STAR Reading Proficiency','STAR Math Score','STAR Reading Score',
    'Preferred Display Name','Effective Student Email','Source Status','Current Active Student','Upstream Match','SCC Match',
    'Math Status','ELA Status','CTE Status','Latest Add/Drop','Add/Drop Date','Data Review'
  ];

  const byId = {};
  const rows = [];

  function scc(id, header) {
    if (!sccIndex[id]) return '';
    return refreshValue_(sccIndex[id].row, sccTable.headerMap, header);
  }

  ids.forEach(function(id) {
    const row = upstreamById[id] || null;
    const activeText = row ? String(refreshValue_(row, uh, 'CURRENT_ACTIVE_STUDENT')).toUpperCase() : '';
    const active = ['Y','YES','TRUE'].indexOf(activeText) !== -1;
    const last = row ? refreshValue_(row, uh, 'STUDENT_LAST_NAME') : scc(id, 'LAST NAME');
    const first = row ? refreshValue_(row, uh, 'STUDENT_FIRST_NAME') : scc(id, 'FIRST NAME');
    const sourceEmail = row ? refreshValue_(row, uh, 'EMAIL_O365') : scc(id, 'Student Email');
    const sourcePreferred = scc(id, 'Prefered Name') || (String(first || '') + ' ' + String(last || '')).trim();
    const preferred = overrides.student(id, 'Preferred Name');
    const emailOverride = overrides.student(id, 'Student Email');
    const latest = addLatest[id] || {};

    function subject(prefix) {
      if (!row) return '';
      const failing = Number(refreshValue_(row, uh, prefix + '_FAILING') || 0);
      const classes = Number(refreshValue_(row, uh, prefix + '_CLASSES') || 0);
      if (failing > 0) return 'Failing';
      if (classes > 0) return 'Passing';
      return '';
    }

    let sourceStatus = active ? 'Active Upstream' : (row ? 'Inactive Upstream' : 'SCC only / source missing');
    if (!active && latest.event && String(latest.event).toUpperCase().indexOf('DROP') === 0) sourceStatus = 'Dropped / not active';

    const cteInfo = cte[id];
    const out = [
      Number(id) || id,
      last || '', first || '',
      row ? refreshValue_(row, uh, 'GRADE') : scc(id, 'GRADE'),
      row ? refreshValue_(row, uh, 'ENROLL_DATE_PS') : scc(id, 'START DATE'),
      row ? refreshValue_(row, uh, 'SCHOOL') : scc(id, 'SCHOOL'),
      sourceEmail || '',
      row ? refreshValue_(row, uh, 'LC_COACH_NAME') : '',
      row ? refreshValue_(row, uh, 'EMAIL_LEARNING_COACH') : '',
      row ? refreshValue_(row, uh, 'SPED504PLAN') : scc(id, 'SPED'),
      row ? refreshValue_(row, uh, 'PAR_ATTENDANCE_PCT') : '',
      row ? refreshValue_(row, uh, 'PAR_CHRONICALLY_ABSENT') : '',
      row ? refreshValue_(row, uh, 'STAR_2TO12_MATH_DATE') : '',
      row ? refreshValue_(row, uh, 'STAR_2TO12_READING_DATE') : '',
      row ? refreshValue_(row, uh, 'STAR_2TO12_MATH_PROFICIENCY') : '',
      row ? refreshValue_(row, uh, 'STAR_2TO12_READING_PROFICIENCY') : '',
      row ? refreshValue_(row, uh, 'STAR_2TO12_MATH_UNIFIED_SCORE') : '',
      row ? refreshValue_(row, uh, 'STAR_2TO12_READING_UNIFIED_SCORE') : '',
      preferred === null ? sourcePreferred : String(preferred || '') + (last ? ' ' + last : ''),
      emailOverride === null ? sourceEmail : emailOverride,
      sourceStatus, active ? 'Y' : 'N', row ? 'Y' : 'N', sccIndex[id] ? 'Y' : 'N',
      subject('MATH'), subject('ELA'),
      cteInfo ? (cteInfo.fail ? 'Failing' : 'Passing') : '',
      latest.event || '', latest.date || '',
      active ? '' : 'Review — not active in current source'
    ];

    rows.push(out);
    const record = {};
    headers.forEach(function(header, i) { record[header] = out[i]; });
    byId[id] = record;
  });

  return { headers: headers, rows: rows, ids: ids, byId: byId };
}

function refreshCleanName_(value) {
  return String(value || '').toLowerCase().replace(/\([^)]*\)/g, '').replace(/[^a-z0-9]/g, '');
}

function refreshBuildContacts_(proRoster, upstream, overrides) {
  const ph = proRoster.headerMap;
  const uh = upstream.headerMap;
  const upstreamById = {};
  upstream.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, uh, 'STUDENT_NUMBER'));
    if (id) upstreamById[id] = row;
  });

  let id = '', name = '', gender = '', studentPhone = '';
  const rows = [];
  const represented = {};

  proRoster.rows.forEach(function(row) {
    const nextId = normalizeId_(refreshValue_(row, ph, 'Student Number'));
    if (nextId) id = nextId;
    if (refreshValue_(row, ph, 'Name')) name = refreshValue_(row, ph, 'Name');
    if (refreshValue_(row, ph, 'Gender')) gender = refreshValue_(row, ph, 'Gender');
    if (refreshValue_(row, ph, 'Phone')) studentPhone = refreshValue_(row, ph, 'Phone');

    const contact = String(refreshValue_(row, ph, 'Contact Name') || '').trim();
    if (!id || !contact) return;

    const rel = contact.match(/\(([^)]+)\)/);
    const sourcePhone = refreshValue_(row, ph, 'Contact Phone') || '';
    const sourceEmail = refreshValue_(row, ph, 'Contact Email') || '';
    const key = id + '|' + contact.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    const u = upstreamById[id];
    const lc = u ? String(refreshValue_(u, uh, 'LC_COACH_NAME') || '') : '';
    const isLc = lc && (refreshCleanName_(lc).indexOf(refreshCleanName_(contact)) !== -1 ||
      refreshCleanName_(contact).indexOf(refreshCleanName_(lc)) !== -1) ? 'Y' : '';
    const phoneOverride = overrides.contact(id, key, contact, 'Phone');
    const emailOverride = overrides.contact(id, key, contact, 'Email');

    rows.push([
      Number(id) || id, name, gender, studentPhone, contact, rel ? rel[1] : '',
      sourcePhone, sourceEmail,
      phoneOverride === null ? sourcePhone : phoneOverride,
      emailOverride === null ? sourceEmail : emailOverride,
      isLc, key
    ]);
    represented[id + '|' + refreshCleanName_(contact)] = true;
  });

  upstream.rows.forEach(function(row) {
    const sid = normalizeId_(refreshValue_(row, uh, 'STUDENT_NUMBER'));
    const lc = String(refreshValue_(row, uh, 'LC_COACH_NAME') || '').trim();
    if (!sid || !lc || represented[sid + '|' + refreshCleanName_(lc)]) return;
    const key = sid + '|' + lc.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    const sourceEmail = refreshValue_(row, uh, 'EMAIL_LEARNING_COACH') || '';
    const emailOverride = overrides.contact(sid, key, lc, 'Email');
    const phoneOverride = overrides.contact(sid, key, lc, 'Phone');
    rows.push([
      Number(sid) || sid,
      String(refreshValue_(row, uh, 'STUDENT_LAST_NAME') || '') + ', ' + String(refreshValue_(row, uh, 'STUDENT_FIRST_NAME') || ''),
      '', '', lc, 'Learning Coach', '', sourceEmail,
      phoneOverride === null ? '' : phoneOverride,
      emailOverride === null ? sourceEmail : emailOverride,
      'Y', key
    ]);
  });

  return rows;
}

function refreshSyncScc_(sheet, byId) {
  const h = getHeaderIndexMap_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);
  const existing = {};
  const idCol = h['Student Number'];
  if (typeof idCol === 'undefined') throw new Error('SCC is missing Student Number.');
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, idCol + 1, sheet.getLastRow() - 1, 1).getDisplayValues().forEach(function(row, i) {
      const id = normalizeId_(row[0]); if (id) existing[id] = i + 2;
    });
  }

  Object.keys(byId).forEach(function(id) {
    const student = byId[id];
    let row = existing[id];
    if (!row && student['Current Active Student'] === 'Y') {
      row = sheet.getLastRow() + 1;
      ensureSheetRows_(sheet, row);
      if (row > 2) sheet.getRange(row - 1, 1, 1, sheet.getLastColumn()).copyTo(
        sheet.getRange(row, 1, 1, sheet.getLastColumn()), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.getRange(row, idCol + 1).setValue(Number(id) || id);
    }
    if (!row) return;

    const fields = {
      'Prefered Name':'Preferred Display Name','LAST NAME':'Source Last Name','FIRST NAME':'Source First Name',
      'GRADE':'Grade','START DATE':'Enroll Date','SCHOOL':'School','SPED':'SPED/504','Student Email':'Effective Student Email'
    };
    Object.keys(fields).forEach(function(target) {
      if (typeof h[target] !== 'undefined' && student[fields[target]] !== '') {
        sheet.getRange(row, h[target] + 1).setValue(student[fields[target]]);
      }
    });
  });
}

function refreshSyncOnboarding_(sheet, data, sccSheet) {
  const old = readSheetTable_(sheet);
  const oldById = {};
  old.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, old.headerMap, 'Student Number'));
    if (!id) return;
    const record = {};
    old.headers.forEach(function(header, i) { record[String(header || '').trim()] = row[i]; });
    oldById[id] = record;
  });

  const sccTable = readSheetTable_(sccSheet);
  const sccById = {};
  sccTable.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, sccTable.headerMap, 'Student Number'));
    if (id) sccById[id] = row;
  });
  function scc(id, header) {
    return sccById[id] ? refreshValue_(sccById[id], sccTable.headerMap, header) : '';
  }

  const rows = data.ids.map(function(id) {
    const student = data.byId[id], prior = oldById[id] || {};
    const reading = prior['BOY STAR Reading PR'] || student['STAR Reading Proficiency'] || '';
    const math = prior['BOY STAR Math PR'] || student['STAR Math Proficiency'] || '';
    return [
      Number(id) || id, student['Preferred Display Name'], student['Enroll Date'],
      prior['Orientation Day 1'] || '', prior['Orientation Day 2'] || '',
      scc(id,'Welcome Email') || prior['Welcome Email'] || '',
      scc(id,'ISMA') || prior['ISMA'] || '',
      reading, math, reading !== '' && math !== '' ? 'BOY scores recorded' : 'Review STAR',
      scc(id,'SCC Completion') || '', student['School'], student['Current Active Student'], student['Data Review']
    ];
  });
  writeRowsBelowHeader_(sheet, rows);
}

function refreshSyncStar_(sheet, data) {
  const rows = data.ids.map(function(id) {
    const s = data.byId[id];
    return [
      Number(id) || id, s['Preferred Display Name'], s['STAR Reading Proficiency'], s['STAR Math Proficiency'],
      s['STAR Reading Date'], s['STAR Math Date'], s['STAR Reading Score'], s['STAR Math Score'],
      'Latest source values from Upstream; preserve BOY values on Onboarding.'
    ];
  });
  writeRowsBelowHeader_(sheet, rows);
}

function refreshDash_(sheet, data, onboardingSheet, sccSheet, overridesSheet) {
  const sccTable = readSheetTable_(sccSheet);
  const sccById = {};
  sccTable.rows.forEach(function(row) {
    const id = normalizeId_(refreshValue_(row, sccTable.headerMap, 'Student Number'));
    if (id) sccById[id] = row;
  });

  const follow = [], chronic = [], sourceReview = [];
  let active = 0, failMath = 0, failEla = 0, failCte = 0;
  data.ids.forEach(function(id) {
    const student = data.byId[id];
    const status = sccById[id] ? String(refreshValue_(sccById[id], sccTable.headerMap, 'SCC Completion') || '') : '';
    if (status !== 'Completed') follow.push([Number(id) || id, student['Preferred Display Name'], status]);
    if (student['Current Active Student'] === 'Y') active++;
    if (String(student['Chronically Absent']).toUpperCase() === 'Y') chronic.push([Number(id) || id, student['Preferred Display Name']]);
    if (student['Data Review']) sourceReview.push([Number(id) || id, student['Preferred Display Name'], student['Source Status']]);
    if (student['Math Status'] === 'Failing') failMath++;
    if (student['ELA Status'] === 'Failing') failEla++;
    if (student['CTE Status'] === 'Failing') failCte++;
  });

  const onboarding = readSheetTable_(onboardingSheet);
  const star = [];
  onboarding.rows.forEach(function(row) {
    if (String(refreshValue_(row, onboarding.headerMap, 'STAR Status')) === 'Review STAR') {
      star.push([refreshValue_(row, onboarding.headerMap, 'Student Number'), refreshValue_(row, onboarding.headerMap, 'Preferred Display Name')]);
    }
  });

  let openOverrides = 0;
  const ov = readSheetTable_(overridesSheet);
  ov.rows.forEach(function(row) {
    const a = refreshValue_(row, ov.headerMap, 'Active');
    const u = refreshValue_(row, ov.headerMap, 'Source Updated?');
    const activeOv = a === true || String(a).toUpperCase() === 'TRUE';
    const updated = u === true || String(u).toUpperCase() === 'TRUE';
    if (activeOv && !updated) openOverrides++;
  });

  sheet.getRange('B3:B9').setValues([[data.rows.length],[active],[follow.length],[star.length],[chronic.length],[openOverrides],[sourceReview.length]]);
  ['A12:C100','E12:F100','H12:I100','K12:M100'].forEach(function(r) { sheet.getRange(r).clearContent(); });
  if (follow.length) sheet.getRange(12,1,follow.length,3).setValues(follow);
  if (star.length) sheet.getRange(12,5,star.length,2).setValues(star);
  if (chronic.length) sheet.getRange(12,8,chronic.length,2).setValues(chronic);
  if (sourceReview.length) sheet.getRange(12,11,sourceReview.length,3).setValues(sourceReview);

  const now = new Date();
  const tuesday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - ((now.getDay() - 2 + 7) % 7));
  sheet.getRange(APP_CONFIG.dash.wigRow, 1, 1, 7).setValues([[
    tuesday, active, chronic.length, chronic.map(function(r) { return r[1]; }).join(', '), failMath, failEla, failCte
  ]]);
  sheet.getRange(APP_CONFIG.dash.wigRow,1).setNumberFormat('m/d/yyyy');
}
