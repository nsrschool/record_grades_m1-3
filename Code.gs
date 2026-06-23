/*************************************************************************
 * ระบบรายงานผลการเรียน หลักสูตรการศึกษาประถมศึกษาตอนต้น พุทธศักราช 2568
 * พร้อมเทียบเคียงเป็นผลการเรียนตามหลักสูตรแกนกลางฯ พุทธศักราช 2551
 * --------------------------------------------------------------------
 * Backend: Google Apps Script (REST API) + Google Sheets (ฐานข้อมูล)
 * Frontend: GitHub Pages เรียกผ่าน fetch() (ไม่ใช้ google.script.run)
 * พัฒนาโดย: ครูรุ่งนิรันดร์
 * --------------------------------------------------------------------
 * หมายเหตุ: ฟีเจอร์แปลง PDF→รูป และนำเข้าเช็คชื่อจาก Excel ใช้ Drive REST API
 *   ผ่าน UrlFetch (ไม่ต้องเปิด Advanced Service) — แต่ครั้งแรกต้องอนุญาตสิทธิ์
 *   Drive + External requests ตอน Deploy เวอร์ชันใหม่
 * --------------------------------------------------------------------
 * Part 1/4 : Backend ทั้งหมด
 *************************************************************************/

/** ====================== ค่าคงที่ / โครงสร้างชีต ====================== */

/* ============================================================
   เมนู NSR (แสดงบนแถบเมนูของ Google Sheets)
   ============================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NSR')
    .addItem('🔧 ติดตั้ง/ตรวจสอบชีตทั้งหมด', 'menuSetupSheets')
    .addSeparator()
    .addItem('🗑️ ล้างรายชื่อนักเรียนทั้งหมด (เพื่ออัปใหม่)', 'menuClearStudents')
    .addItem('🧹 ล้างคะแนน/เกรด/ความสามารถทั้งหมด', 'menuClearScores')
    .addItem('🧨 ล้างข้อมูลทั้งระบบ (ยกเว้นตั้งค่า)', 'menuClearAll')
    .addSeparator()
    .addItem('📥 นำเข้าการเช็คชื่อจากไฟล์ Excel', 'menuImportAttendance')
    .addSeparator()
    .addItem('💾 สำรองข้อมูลลง Drive เดี๋ยวนี้', 'menuBackupNow')
    .addToUi();
}

function menuSetupSheets() {
  setupSheets();
  SpreadsheetApp.getUi().alert('NSR', 'ตรวจสอบ/สร้างชีตครบแล้ว ✅', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ล้างเฉพาะรายชื่อนักเรียน (และข้อมูลที่ผูกกับนักเรียน) เพื่ออัปรายชื่อใหม่
function menuClearStudents() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('ล้างรายชื่อนักเรียน',
    'จะลบ "รายชื่อนักเรียนเดิมทั้งหมด" พร้อมคะแนน/เกรด/ความสามารถ/การเช็คชื่อ ที่ผูกกับนักเรียน\n\nเพื่อให้พร้อมอัปรายชื่อใหม่ — ดำเนินการต่อหรือไม่?',
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  ['Students', 'Scores', 'Grades', 'Abilities68', 'AbilityDetail', 'Attendance', 'Locks'].forEach(function (name) {
    var s = ss().getSheetByName(name);
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });
  SpreadsheetApp.flush();
  ui.alert('NSR', 'ล้างรายชื่อนักเรียนและข้อมูลที่เกี่ยวข้องแล้ว ✅\nนำเข้ารายชื่อใหม่ได้เลย', ui.ButtonSet.OK);
}

function menuClearScores() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('ล้างผลการเรียน', 'จะลบคะแนน/เกรด/ความสามารถทั้งหมด (รายชื่อนักเรียนยังอยู่) ดำเนินการต่อ?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ['Scores', 'Grades', 'Abilities68', 'AbilityDetail', 'LearningUnits'].forEach(function (name) {
    var s = ss().getSheetByName(name);
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });
  SpreadsheetApp.flush();
  ui.alert('NSR', 'ล้างผลการเรียนแล้ว ✅', ui.ButtonSet.OK);
}

function menuClearAll() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ ล้างข้อมูลทั้งระบบ',
    'จะลบทุกข้อมูล (นักเรียน/รายวิชา/คะแนน/เกรด/เช็คชื่อ ฯลฯ) ยกเว้นค่าตั้งค่าโรงเรียนและตารางแปลงเกรด\n\nดำเนินการต่อหรือไม่?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  Object.keys(SHEETS).forEach(function (name) {
    if (name === 'Settings' || name === 'GradeMapping' || name === 'Teachers') return;
    var s = ss().getSheetByName(name);
    if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
  });
  SpreadsheetApp.flush();
  ui.alert('NSR', 'ล้างข้อมูลทั้งระบบแล้ว ✅', ui.ButtonSet.OK);
}

function menuBackupNow() {
  var r = runBackupToDrive();
  SpreadsheetApp.getUi().alert('NSR', 'สำรองข้อมูลลง Drive แล้ว ✅\nไฟล์: ' + r.name, SpreadsheetApp.getUi().ButtonSet.OK);
}

// นำเข้าการเช็คชื่อจากไฟล์ Excel (รูปแบบระบบเช็คชื่อเดิม: ชีต Students + Attendance)
function menuImportAttendance() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('นำเข้าการเช็คชื่อจาก Excel',
    '1) อัปโหลดไฟล์ .xlsx เข้า Google Drive ก่อน\n2) วางลิงก์ (URL) หรือ File ID ของไฟล์ที่นี่\n\n(ไฟล์ต้องมีชีต "Students" และ "Attendance" รูปแบบ uuid|ชั้น|เลขที่|คำนำหน้า|ชื่อ|นามสกุล และ วันที่|uuid|สถานะ)',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var idOrUrl = resp.getResponseText().trim();
  if (!idOrUrl) { ui.alert('ยังไม่ได้ใส่ลิงก์/ID'); return; }
  try {
    var r = importAttendanceFromFile_(idOrUrl);
    ui.alert('NSR — นำเข้าสำเร็จ ✅',
      'นำเข้า: ' + r.imported + ' รายการ\nข้ามซ้ำ: ' + r.skippedDup + '\nจับคู่นักเรียนไม่ได้: ' + r.unmatched + ' รายการ' +
      (r.unmatchedClasses.length ? '\n(ชั้นที่ไม่มีในระบบ: ' + r.unmatchedClasses.join(', ') + ')' : ''),
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('เกิดข้อผิดพลาด', String(e.message || e), ui.ButtonSet.OK);
  }
}

// คัดลอกไฟล์ Drive พร้อมแปลงชนิด (ผ่าน REST API + OAuth token) — ไม่ต้องเปิด Advanced Service
function driveCopyConvert_(fileId, name, targetMime) {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/copy?supportsAllDrives=true', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ name: name, mimeType: targetMime }),
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText() || '{}');
  if (!data.id) throw new Error('แปลงไฟล์ไม่สำเร็จ: ' + ((data.error && data.error.message) || res.getContentText()));
  return data.id;
}

function importAttendanceFromFile_(idOrUrl) {
  var m = String(idOrUrl).match(/[-\w]{25,}/);
  if (!m) throw new Error('ลิงก์/ID ไม่ถูกต้อง');
  var fileId = m[0];

  // แปลง xlsx → Google Sheet ชั่วคราว (ผ่าน Drive REST API — ไม่ต้องเปิด Advanced Service)
  var tmpId = driveCopyConvert_(fileId, '__att_import_tmp', 'application/vnd.google-apps.spreadsheet');
  try {
    var tmp = SpreadsheetApp.openById(tmpId);
    var sStu = tmp.getSheetByName('Students');
    var sAtt = tmp.getSheetByName('Attendance');
    if (!sStu || !sAtt) throw new Error('ไฟล์ต้องมีชีต "Students" และ "Attendance"');

    // uuid → ข้อมูลนักเรียนในไฟล์
    var fsv = sStu.getDataRange().getValues();
    var uuidInfo = {};
    fsv.forEach(function (r) {
      if (!r[0]) return;
      uuidInfo[String(r[0])] = { cls: String(r[1] || '').trim(), num: Number(r[2]) || 0,
        first: String(r[4] || '').trim(), last: String(r[5] || '').trim() };
    });

    // ระบบ: ดัชนีหลายแบบ — ชื่อ-นามสกุล (หลัก), ชั้น+เลขที่ (สำรอง), ชั้น+ชื่อ (แก้ชื่อซ้ำ)
    var nrm = function (x) { return String(x || '').replace(/\s+/g, '').trim(); };
    var sysByName = {}, sysByNum = {}, sysByClassName = {};
    getStudents().forEach(function (s) {
      var nk = nrm(s.firstName) + '|' + nrm(s.lastName);
      (sysByName[nk] = sysByName[nk] || []).push(s.ID);     // อาจมีชื่อซ้ำหลายคน
      sysByNum[s.classLevel + '|' + (Number(s.number) || 0)] = s.ID;
      sysByClassName[s.classLevel + '|' + nk] = s.ID;
    });

    // อ่าน attendance ในไฟล์
    var av = sAtt.getDataRange().getValues();
    var STMAP = { present: 'มา', late: 'สาย', leave: 'ลา', absent: 'ขาด' };

    // ชุดคีย์ที่มีอยู่แล้วในระบบ (กันซ้ำ): date|studentID|period0
    var sys = sheet('Attendance');
    var existing = {};
    var ev = sys.getDataRange().getValues();
    for (var i = 1; i < ev.length; i++) {
      if (String(ev[i][3]) === '0') existing[ymd(ev[i][1]) + '|' + ev[i][4]] = true;
    }

    var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
    var out = [], imported = 0, skippedDup = 0, unmatched = 0, unmatchedCls = {};
    av.forEach(function (r) {
      if (!r[0] || !r[1]) return;
      var info = uuidInfo[String(r[1])];
      if (!info) { unmatched++; return; }
      var nk = nrm(info.first) + '|' + nrm(info.last);
      var sid = '';
      var byName = sysByName[nk] || [];
      if (byName.length === 1) sid = byName[0];                                  // ชื่อไม่ซ้ำ → จับด้วยชื่อ
      else if (byName.length > 1) sid = sysByClassName[info.cls + '|' + nk] || ''; // ชื่อซ้ำ → ใช้ชั้นช่วย
      if (!sid) sid = sysByClassName[info.cls + '|' + nk] || sysByNum[info.cls + '|' + info.num] || ''; // สำรอง
      if (!sid) { unmatched++; unmatchedCls[info.cls] = true; return; }
      var date = ymd(r[0]);
      var key = date + '|' + sid;
      if (existing[key]) { skippedDup++; return; }
      existing[key] = true;
      var status = STMAP[String(r[2] || '').toLowerCase()] || 'มา';
      out.push([genId('ATT'), "'" + date, info.cls, '0', sid, status, 'นำเข้า', now]);
      imported++;
    });

    if (out.length) sys.getRange(sys.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    SpreadsheetApp.flush();

    return { imported: imported, skippedDup: skippedDup, unmatched: unmatched,
      unmatchedClasses: Object.keys(unmatchedCls).sort() };
  } finally {
    try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) {}
  }
}


var SHEETS = {
  Settings:     ['Key', 'Value'],
  Students:     ['ID', 'studentCode', 'number', 'fullName', 'classLevel', 'dateAdded', 'prefix', 'firstName', 'lastName'],
  Subjects:     ['ID', 'code', 'name', 'learningArea', 'hours', 'type', 'classLevel', 'sortOrder'],
  SubjectDocs:  ['ID', 'subjectID', 'fileName', 'fileId', 'type', 'updatedAt'],
  LearningUnits:['ID', 'subjectID', 'unitName', 'maxScore', 'sortOrder', 'semester'],
  Scores:       ['ID', 'studentID', 'subjectID', 'unitID', 'score', 'semester'],
  Grades:       ['ID', 'studentID', 'subjectID', 'mode', 'percent', 'level68', 'grade51', 'letter51', 'note', 'semester'],
  Abilities68:  ['ID', 'studentID', 'category', 'itemKey', 'result', 'semester'],
  AbilityDetail:['ID', 'studentID', 'group', 'item', 'value', 'semester'],
  GradeMapping: ['minPercent', 'maxPercent', 'grade51', 'letter51'],
  Teachers:     ['ID', 'name', 'pin', 'role', 'homeroomClass'],
  Schedule:     ['ID', 'day', 'period', 'classLevel', 'subject', 'teacher'],
  Attendance:   ['ID', 'date', 'classLevel', 'period', 'studentID', 'status', 'checkedBy', 'updatedAt', 'semester'],
  Calendar:     ['ID', 'date', 'type', 'title', 'note'],
  Locks:        ['classLevel', 'status', 'updatedAt', 'by']
};

// ภาคเรียนปัจจุบัน (default ถ้า frontend ยังไม่ส่ง semester มา)
function curSem_(p) {
  if (p && p.semester) return String(p.semester);
  var cs = getSettings().currentSemester;
  return String(cs || '1');
}

// ลบแถวที่ตรงสองเงื่อนไข (ใช้ header-name หาคอลัมน์ — ปลอดภัยจากคอลัมน์เลื่อน)
function clearRowsWhere2(name, f1, v1, f2, v2) {
  var s = sheet(name);
  var data = s.getDataRange().getValues();
  if (data.length < 2) return;
  var c1 = data[0].indexOf(f1), c2 = data[0].indexOf(f2);
  if (c1 < 0) return;
  var keep = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i].join('') === '') continue;
    var match = String(data[i][c1]) === String(v1) && (c2 < 0 || String(data[i][c2]) === String(v2));
    if (!match) keep.push(data[i]);
  }
  var cols = data[0].length;
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, cols).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, cols).setValues(keep);
}

// ค่าตั้งต้นข้อมูลสถานศึกษา (แก้ไขได้ในเมนูตั้งค่า)
var DEFAULT_SETTINGS = {
  schoolName:      'โรงเรียนของฉัน',
  slogan:          'ระบบรายงานผลการเรียน หลักสูตรแกนกลางฯ พุทธศักราช 2551',
  academicYear:    '2569',
  area:            '',                 // สำนักงานเขตพื้นที่
  directorName:    '',                 // ชื่อผู้อำนวยการ
  directorSignURL: '',                 // URL ลายเซ็น ผอ.
  deputyName:      '',                 // ชื่อรองผู้อำนวยการ (ฝ่ายวิชาการ)
  deputySignURL:   '',                 // URL ลายเซ็นรอง ผอ.
  logoURL:         '',                 // URL ตราสัญลักษณ์
  remark:          '',
  hoursPerYear:    '200',              // มาตรฐานเวลาเรียน ชม./ปี (ประถม)
  currentSemester: '1',                // ภาคเรียนปัจจุบัน (1/2)
  sem1Start:       '',                 // ช่วงวันที่ภาคเรียนที่ 1 (YYYY-MM-DD) — ใช้กรองเวลาเรียน/มส.
  sem1End:         '',
  sem2Start:       '',
  sem2End:         ''
};

// ตารางแปลงเกรดเริ่มต้น (อิงตารางที่ ๕ ของแนวปฏิบัติฯ) — แก้ไขได้ในเมนูตั้งค่า
var DEFAULT_MAPPING = [
  { minPercent: 80, maxPercent: 100, grade51: 4,   letter51: 'ดีเยี่ยม' },
  { minPercent: 75, maxPercent: 79,  grade51: 3.5, letter51: 'ดีมาก'   },
  { minPercent: 70, maxPercent: 74,  grade51: 3,   letter51: 'ดี'      },
  { minPercent: 65, maxPercent: 69,  grade51: 2.5, letter51: 'ค่อนข้างดี' },
  { minPercent: 60, maxPercent: 64,  grade51: 2,   letter51: 'ปานกลาง'  },
  { minPercent: 55, maxPercent: 59,  grade51: 1.5, letter51: 'พอใช้'    },
  { minPercent: 50, maxPercent: 54,  grade51: 1,   letter51: 'ผ่านเกณฑ์ขั้นต่ำ' },
  { minPercent: 0,  maxPercent: 49,  grade51: 0,   letter51: 'ต่ำกว่าเกณฑ์' }
];

// โฟลเดอร์เก็บรูปบน Drive
var DRIVE_FOLDER_NAME = 'KPS_GradeReport_Images';


/** ====================== ROUTER (REST) ====================== */

function doGet(e)  { return handleRequest(e, 'GET');  }
function doPost(e) { return handleRequest(e, 'POST'); }

function handleRequest(e, method) {
  var action  = '';
  var payload = {};
  try {
    // อ่าน action + payload จาก POST (text/plain) หรือ GET (?action=)
    if (method === 'POST' && e && e.postData && e.postData.contents) {
      var body = JSON.parse(e.postData.contents);
      action  = body.action || '';
      payload = body.payload || {};
    } else if (e && e.parameter) {
      action  = e.parameter.action || '';
      if (e.parameter.payload) payload = JSON.parse(e.parameter.payload);
    }

    setupSheets(); // สร้างชีตอัตโนมัติถ้ายังไม่มี

    var result = dispatch(action, payload);
    return jsonOut({ ok: true, data: result });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function dispatch(action, p) {
  switch (action) {
    // โหลดข้อมูลเริ่มต้น
    case 'getBootstrap':      return { settings: getSettings(), mapping: getGradeMapping() };

    // ตั้งค่าสถานศึกษา
    case 'getSettings':       return getSettings();
    case 'saveSettings':      return saveSettings(p);

    // ตารางแปลงเกรด
    case 'getGradeMapping':   return getGradeMapping();
    case 'saveGradeMapping':  return saveGradeMapping(p.mapping);

    // นักเรียน
    case 'getStudents':       return getStudents();
    case 'addStudent':        return addStudent(p);
    case 'updateStudent':     return updateStudent(p.id, p);
    case 'deleteStudent':     return deleteStudent(p.id);
    case 'bulkImportStudents':return bulkImportStudents(p.rows);

    // รายวิชา
    case 'getSubjects':       return getSubjects();
    case 'addSubject':        return addSubject(p);
    case 'updateSubject':     return updateSubject(p.id, p);
    case 'deleteSubject':     return deleteSubject(p.id);

    // หน่วย/ผลลัพธ์การเรียนรู้
    case 'getUnits':          return getUnits(p.subjectID, p.semester);
    case 'saveUnits':         return saveUnits(p.subjectID, p.units, p.semester);

    // คะแนน (โหมด A)
    case 'getScores':         return getScores(p.subjectID, p.semester);
    case 'saveScores':        return saveScores(p.subjectID, p.scores, p.semester);

    // เกรด/ผลการเรียน
    case 'getGrades':         return getGrades(p.semester);
    case 'saveGrade':         return saveGrade(p);
    case 'saveGradesBatch':   return saveGradesBatch(p.grades, p.semester);
    case 'recomputeGrades':   return recomputeGrades(p.subjectID, p.semester);

    // ความสามารถ / คุณลักษณะ / กิจกรรม / อ่านคิดวิเคราะห์เขียน
    case 'getAbilities':      return getAbilities(p.studentID, p.semester);
    case 'saveAbilities':     return saveAbilities(p.studentID, p.items, p.semester);
    case 'getAbilityDetailClass': return getAbilityDetailClass(p.classLevel, p.semester);
    case 'saveAbilityDetailBatch': return saveAbilityDetailBatch(p);

    // รายงาน
    case 'getReportData':     return getReportData(p.studentID, p.semester);
    case 'getReportBatch':    return getReportBatch(p.classLevel);
    case 'getPP5Cover':       return getPP5Cover(p.classLevel);
    case 'getPP5Book':        return getPP5Book(p);
    case 'getCalendar':       return getCalendar();
    case 'getLocks':          return getLocks();
    case 'getRiskStudents':   return getRiskStudents(p);
    case 'getDirectorOverview': return getDirectorOverview(p);
    case 'setLock':           return setLock(p);
    case 'addCalendarEvent':  return addCalendarEvent(p);
    case 'deleteCalendarEvent': return deleteCalendarEvent(p.id);
    case 'getStats':          return getStats();
    case 'getGradeStats':     return getGradeStats(p);
    case 'getClassGradeSummary': return getClassGradeSummary(p.classLevel, p.semester);
    case 'exportSchoolMIS':   return exportSchoolMIS(p.classLevel);
    case 'exportSchoolMISWide': return exportSchoolMISWide(p.classLevel);

    // จัดการข้อมูล
    case 'backupAllData':     return backupAllData();
    case 'runBackupToDrive':  return runBackupToDrive();
    case 'setAutoBackup':     return setAutoBackup(p);
    case 'getBackupStatus':   return getBackupStatus();
    case 'restoreBackup':     return restoreBackup(p);
    case 'importBackupData':  return importBackupData(p.json);
    case 'clearAllData':      return clearAllData();

    // อัปโหลดรูป
    case 'uploadImage':       return uploadImageToDrive(p.base64, p.filename);
    case 'getSubjectDocs':    return getSubjectDocs(p.subjectID);
    case 'getDocBase64':      return getDocBase64(p.fileId);
    case 'uploadSubjectDoc':  return uploadSubjectDoc(p);
    case 'deleteSubjectDoc':  return deleteSubjectDoc(p.id);

    case 'ping':              return { pong: true, time: new Date().toString() };

    // ---- Part 5: ล็อกอิน / ครู / ตารางสอน / เช็คชื่อ ----
    case 'login':             return login(p.name, p.pin);
    case 'changeMyPassword':  return changeMyPassword(p);
    case 'getTeachers':       return getTeachers();
    case 'saveTeacher':       return saveTeacher(p);
    case 'saveTeachersBatch': return saveTeachersBatch(p.teachers);
    case 'deleteTeacher':     return deleteTeacher(p.id);

    case 'getSchedule':       return getSchedule(p.teacher);
    case 'importSchedule':    return importSchedule(p.rows);
    case 'syncSubjectsFromSchedule': return syncSubjectsFromSchedule();

    case 'bulkImportRoster':  return bulkImportRoster(p.rows);
    case 'renumberClass':     return renumberClass(p.classLevel);

    case 'getAttendance':     return getAttendance(p.date, p.classLevel, p.period);
    case 'getFlagAttendance': return getAttendance(p.date, p.classLevel, 0);
    case 'saveAttendance':    return saveAttendance(p);
    case 'getAttendanceSummary':   return getAttendanceSummary(p.classLevel, p.fromDate, p.toDate);
    case 'getAttendanceDashboard': return getAttendanceDashboard(p.date);
    case 'getAttendanceRegister':  return getAttendanceRegister(p.classLevel, p.fromDate, p.toDate);
    case 'getMonthlyGrid':         return getMonthlyGrid(p.classLevel, p.ym);

    default: throw new Error('ไม่พบ action: "' + action + '"');
  }
}


/** ====================== HELPERS ====================== */

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet(name) {
  var s = ss().getSheetByName(name);
  if (!s) { s = ss().insertSheet(name); s.appendRow(SHEETS[name]); }
  return s;
}

// สร้างชีตทั้งหมด + ใส่ค่าตั้งต้นเมื่อรันครั้งแรก
// แปลงหัวคอลัมน์ภาษาไทย/รูปแบบเดิม → คีย์อังกฤษที่ระบบใช้ (เพื่อให้ใช้ชีตจากระบบเดิมที่หัวเป็นไทยได้)
var HEADER_ALIASES = {
  Teachers: { 'ชื่อ': 'name', 'ชื่อ-สกุล': 'name', 'ชื่อ-นามสกุล': 'name', 'ชื่อครู': 'name', 'PIN': 'pin', 'รหัส': 'pin', 'รหัสผ่าน': 'pin', 'บทบาท': 'role', 'สิทธิ์': 'role', 'ห้องประจำชั้น': 'homeroomClass', 'ห้องที่ปรึกษา': 'homeroomClass', 'ครูประจำชั้น': 'homeroomClass' },
  Students: { 'รหัสนักเรียน': 'studentCode', 'เลขประจำตัว': 'studentCode', 'เลขประจำตัวนักเรียน': 'studentCode', 'เลขที่': 'number', 'ชื่อ-นามสกุล': 'fullName', 'ชื่อ-สกุล': 'fullName', 'ชื่อสกุล': 'fullName', 'ชื่อ - นามสกุล': 'fullName', 'ระดับชั้น/ห้อง': 'classLevel', 'ระดับชั้น': 'classLevel', 'ชั้น/ห้อง': 'classLevel', 'ชั้น': 'classLevel', 'ห้อง': 'classLevel', 'วันที่เพิ่ม': 'dateAdded', 'คำนำหน้า': 'prefix', 'ชื่อจริง': 'firstName', 'นามสกุล': 'lastName' },
  Settings: { 'คีย์': 'Key', 'ค่า': 'Value' }
};
function normalizeHeaders_() {
  Object.keys(HEADER_ALIASES).forEach(function (name) {
    var s = ss().getSheetByName(name); if (!s) return;
    var lastCol = s.getLastColumn(); if (lastCol < 1) return;
    var hdr = s.getRange(1, 1, 1, lastCol).getValues()[0];
    var alias = HEADER_ALIASES[name];
    var existing = {}; hdr.forEach(function (h) { existing[String(h).trim()] = true; });
    var changed = false;
    for (var c = 0; c < hdr.length; c++) {
      var key = String(hdr[c]).trim();
      var canon = alias[key];
      if (canon && key !== canon && !existing[canon]) { hdr[c] = canon; existing[canon] = true; changed = true; }
    }
    if (changed) s.getRange(1, 1, 1, hdr.length).setValues([hdr]);
  });
}

function setupSheets() {
  Object.keys(SHEETS).forEach(function (name) {
    var s = ss().getSheetByName(name);
    if (!s) { s = ss().insertSheet(name); s.appendRow(SHEETS[name]); }
  });
  normalizeHeaders_(); // แปลงหัวคอลัมน์ไทย→อังกฤษ ให้ระบบอ่านข้อมูลเดิมได้
  // เติมหัวคอลัมน์ที่ "ต่อท้ายใหม่" (เช่น semester) ให้ชีตเดิมที่ยังไม่มี — ปลอดภัยเพราะเป็นการต่อท้ายเท่านั้น
  ['LearningUnits', 'Scores', 'Grades', 'Abilities68', 'AbilityDetail', 'Attendance'].forEach(function (name) {
    var s = ss().getSheetByName(name);
    if (!s) return;
    var lastCol = s.getLastColumn();
    var hdr = s.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
    var schema = SHEETS[name];
    if (hdr.length < schema.length) {
      s.getRange(1, 1, 1, schema.length).setValues([schema]); // เขียนหัวให้ครบตาม schema
    }
  });
  // ใส่ค่าตั้งต้น Settings
  var st = sheet('Settings');
  if (st.getLastRow() < 2) {
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      st.appendRow([k, DEFAULT_SETTINGS[k]]);
    });
  }
  // ใส่ตารางแปลงเกรดตั้งต้น
  var gm = sheet('GradeMapping');
  if (gm.getLastRow() < 2) {
    DEFAULT_MAPPING.forEach(function (m) {
      gm.appendRow([m.minPercent, m.maxPercent, m.grade51, m.letter51]);
    });
  }
  // seed บัญชีแอดมินเริ่มต้น (PIN: 1234 — เปลี่ยนได้ในหน้าจัดการครู)
  var tc = sheet('Teachers');
  if (tc.getLastRow() < 2) {
    tc.appendRow([genId('TCH'), 'แอดมิน', '1234', 'admin', '']);
  }
  SpreadsheetApp.flush();
}

// อ่านทั้งชีตเป็น array ของ object (แปลง Date เป็น string เสมอ)
function readAll(name) {
  var s = sheet(name);
  var values = s.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue; // ข้ามแถวว่าง
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var v = row[c];
      if (v instanceof Date) v = Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
      obj[headers[c]] = v;
    }
    obj._rowIndex = i + 1; // เลขแถวจริงในชีต (1-based)
    rows.push(obj);
  }
  return rows;
}

function genId(prefix) {
  return (prefix || 'ID') + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

// แปลงค่าวันที่ (Date หรือ string) ให้เป็น 'yyyy-MM-dd' เสมอ (กันปัญหา Google Sheets แปลงเป็น Date)
function ymd(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'GMT+7', 'yyyy-MM-dd');
  return String(v == null ? '' : v).slice(0, 10);
}

function findRowById(name, id) {
  var s = sheet(name);
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based row
  }
  return -1;
}


/** ====================== SETTINGS ====================== */

function getSettings() {
  var rows = readAll('Settings');
  var obj = {};
  rows.forEach(function (r) { obj[r.Key] = r.Value; });
  // เติมคีย์ที่อาจขาดด้วยค่าตั้งต้น
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    if (typeof obj[k] === 'undefined') obj[k] = DEFAULT_SETTINGS[k];
  });
  return obj;
}

function saveSettings(data) {
  var s = sheet('Settings');
  s.clearContents();
  var rows = [SHEETS.Settings];
  Object.keys(data).forEach(function (k) {
    if (k === 'action' || k === 'payload') return;
    rows.push([k, String(data[k] == null ? '' : data[k])]);
  });
  // บังคับคอลัมน์ Value เป็นข้อความ ป้องกัน Google Sheets แปลงวันที่/ตัวเลขอัตโนมัติ
  s.getRange(1, 2, rows.length, 1).setNumberFormat('@');
  s.getRange(1, 1, rows.length, 2).setValues(rows);
  SpreadsheetApp.flush();
  return getSettings();
}


/** ====================== GRADE MAPPING + เครื่องมือแปลงเกรด ====================== */

function getGradeMapping() {
  var rows = readAll('GradeMapping');
  return rows.map(function (r) {
    return {
      minPercent: Number(r.minPercent),
      maxPercent: Number(r.maxPercent),
      grade51: r.grade51,
      letter51: r.letter51
    };
  });
}

function saveGradeMapping(mapping) {
  var s = sheet('GradeMapping');
  s.clearContents();
  s.appendRow(SHEETS.GradeMapping);
  (mapping || []).forEach(function (m) {
    s.appendRow([m.minPercent, m.maxPercent, m.grade51, m.letter51]);
  });
  SpreadsheetApp.flush();
  return getGradeMapping();
}

// แกนกลาง: ร้อยละ -> {grade51, letter51} (หลักสูตร 2551, ไม่มีระดับพัฒนาการ 2568)
function percentToGrade(percent, mapping) {
  mapping = mapping || getGradeMapping();
  var p = Number(percent);
  if (isNaN(p)) return { percent: null, level68: '', grade51: '', letter51: '' };
  for (var i = 0; i < mapping.length; i++) {
    if (p >= mapping[i].minPercent && p <= mapping[i].maxPercent) {
      return { percent: p, level68: '', grade51: mapping[i].grade51, letter51: mapping[i].letter51 };
    }
  }
  return { percent: p, level68: '', grade51: '', letter51: '' };
}


/** ====================== STUDENTS ====================== */

function getStudents() { ensureStudentCols_(); return readAll('Students'); }

// คำนำหน้าที่รู้จัก (เรียงยาว→สั้น เพื่อจับ "นางสาว" ก่อน "นาง")
var NAME_PREFIXES = ['เด็กชาย', 'เด็กหญิง', 'ด.ช.', 'ด.ญ.', 'นางสาว', 'นาง', 'นาย', 'ว่าที่ร้อยตรี', 'ว่าที่ร้อยตรีหญิง'];
function parseThaiName_(full) {
  full = String(full || '').trim();
  var prefix = '';
  for (var i = 0; i < NAME_PREFIXES.length; i++) {
    if (full.indexOf(NAME_PREFIXES[i]) === 0) { prefix = NAME_PREFIXES[i]; full = full.slice(NAME_PREFIXES[i].length).trim(); break; }
  }
  var parts = full.split(/\s+/).filter(String);
  var first = parts.shift() || '';
  return { prefix: prefix, first: first, last: parts.join(' ') };
}
function buildFullName_(prefix, first, last) {
  return ((prefix || '') + (first || '') + (last ? ' ' + last : '')).trim();
}
// เพิ่มคอลัมน์ใหม่ + แยกชื่อเดิม (ทำครั้งเดียว)
function ensureStudentCols_() {
  var s = sheet('Students');
  var lastCol = s.getLastColumn();
  var header = s.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
  if (header.indexOf('firstName') >= 0) return; // ย้ายข้อมูลแล้ว
  s.getRange(1, 1, 1, SHEETS.Students.length).setValues([SHEETS.Students]);
  var n = s.getLastRow();
  if (n < 2) return;
  var rng = s.getRange(2, 1, n - 1, SHEETS.Students.length);
  var vals = rng.getValues();
  vals.forEach(function (r) {
    if (!r[6] && !r[7] && !r[8]) {
      var p = parseThaiName_(r[3]);
      r[6] = p.prefix; r[7] = p.first; r[8] = p.last;
    }
  });
  rng.setValues(vals);
  SpreadsheetApp.flush();
}

function addStudent(p) {
  ensureStudentCols_();
  var s = sheet('Students');
  var id = genId('STD');
  var full = p.fullName || buildFullName_(p.prefix, p.firstName, p.lastName);
  var parts = (p.firstName || p.lastName || p.prefix) ? { prefix: p.prefix || '', first: p.firstName || '', last: p.lastName || '' } : parseThaiName_(full);
  s.appendRow([
    id, p.studentCode || '', p.number || '', full, p.classLevel || '',
    Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss'),
    parts.prefix, parts.first, parts.last
  ]);
  SpreadsheetApp.flush();
  return { id: id };
}

function updateStudent(id, p) {
  ensureStudentCols_();
  var row = findRowById('Students', id);
  if (row < 0) throw new Error('ไม่พบนักเรียน ID: ' + id);
  var s = sheet('Students');
  var parts = (p.firstName || p.lastName || p.prefix) ? { prefix: p.prefix || '', first: p.firstName || '', last: p.lastName || '' } : parseThaiName_(p.fullName);
  var full = buildFullName_(parts.prefix, parts.first, parts.last) || (p.fullName || '');
  s.getRange(row, 2, 1, 4).setValues([[p.studentCode || '', p.number || '', full, p.classLevel || '']]);
  s.getRange(row, 7, 1, 3).setValues([[parts.prefix, parts.first, parts.last]]);
  SpreadsheetApp.flush();
  return { id: id };
}

function deleteStudent(id) {
  var row = findRowById('Students', id);
  if (row < 0) throw new Error('ไม่พบนักเรียน ID: ' + id);
  // จำชั้นของนักเรียนคนนี้ไว้ก่อนลบ เพื่อเลื่อนเลขที่
  var s = sheet('Students');
  var classLevel = s.getRange(row, 5).getValue(); // คอลัมน์ classLevel
  s.deleteRow(row);
  // ลบคะแนน/เกรด/ความสามารถที่ผูกกับนักเรียนคนนี้ (ใช้ clearContent ปลอดภัยกว่า)
  clearRowsWhere('Scores', 'studentID', id);
  clearRowsWhere('Grades', 'studentID', id);
  clearRowsWhere('Abilities68', 'studentID', id);
  clearRowsWhere('Attendance', 'studentID', id);
  SpreadsheetApp.flush();
  if (classLevel) renumberClass(classLevel); // เลื่อนเลขที่ด้านหลังขึ้นอัตโนมัติ
  return { id: id };
}

// เรียงเลขที่ของห้องให้ต่อเนื่อง 1..N (ตามลำดับเลขที่เดิม)
function renumberClass(classLevel) {
  var s = sheet('Students');
  var data = s.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]) === String(classLevel) && data[i].join('') !== '') {
      items.push({ row: i + 1, number: Number(data[i][2]) || 0 });
    }
  }
  items.sort(function (a, b) { return a.number - b.number; });
  items.forEach(function (it, idx) {
    s.getRange(it.row, 3).setValue(idx + 1); // คอลัมน์ number
  });
  SpreadsheetApp.flush();
  return { classLevel: classLevel, count: items.length };
}

function bulkImportStudents(rows) {
  if (!rows || !rows.length) return { count: 0 };
  var s = sheet('Students');
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var out = [];
  rows.forEach(function (r) {
    out.push([
      genId('STD'),
      r.studentCode || '',
      r.number || '',
      r.fullName || '',
      r.classLevel || '',
      now
    ]);
  });
  s.getRange(s.getLastRow() + 1, 1, out.length, SHEETS.Students.length).setValues(out);
  SpreadsheetApp.flush();
  return { count: out.length };
}


/** ====================== SUBJECTS ====================== */

function getSubjects() { return readAll('Subjects'); }

function addSubject(p) {
  var s = sheet('Subjects');
  var id = genId('SUB');
  s.appendRow([
    id,
    p.code || '',
    p.name || '',
    p.learningArea || '',
    p.hours || '',
    p.type || 'พื้นฐาน',
    p.classLevel || '',
    p.sortOrder || (s.getLastRow())
  ]);
  SpreadsheetApp.flush();
  return { id: id };
}

function updateSubject(id, p) {
  var row = findRowById('Subjects', id);
  if (row < 0) throw new Error('ไม่พบรายวิชา ID: ' + id);
  var s = sheet('Subjects');
  s.getRange(row, 2, 1, 7).setValues([[
    p.code || '', p.name || '', p.learningArea || '', p.hours || '',
    p.type || 'พื้นฐาน', p.classLevel || '', p.sortOrder || ''
  ]]);
  SpreadsheetApp.flush();
  return { id: id };
}

function deleteSubject(id) {
  var row = findRowById('Subjects', id);
  if (row < 0) throw new Error('ไม่พบรายวิชา ID: ' + id);
  sheet('Subjects').deleteRow(row);
  clearRowsWhere('LearningUnits', 'subjectID', id);
  clearRowsWhere('Scores', 'subjectID', id);
  clearRowsWhere('Grades', 'subjectID', id);
  SpreadsheetApp.flush();
  return { id: id };
}


/** ====================== LEARNING UNITS (หน่วย/ผลลัพธ์ ตั้งชื่ออิสระ) ====================== */

function getUnits(subjectID, semester) {
  var sem = semester ? String(semester) : '';
  return readAll('LearningUnits')
    .filter(function (u) {
      if (String(u.subjectID) !== String(subjectID)) return false;
      if (sem && u.semester && String(u.semester) !== sem) return false;
      return true;
    })
    .sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
}

// แทนที่หน่วยของรายวิชานี้ "เฉพาะภาคเรียนนั้น" ด้วยชุดใหม่
function saveUnits(subjectID, units, semester) {
  var sem = String(semester || '1');
  clearRowsWhere2('LearningUnits', 'subjectID', subjectID, 'semester', sem);
  var s = sheet('LearningUnits');
  (units || []).forEach(function (u, i) {
    s.appendRow([
      u.ID || genId('UNT'),
      subjectID,
      u.unitName || ('หน่วยที่ ' + (i + 1)),
      u.maxScore || 0,
      i + 1,
      sem
    ]);
  });
  SpreadsheetApp.flush();
  return getUnits(subjectID, sem);
}


/** ====================== SCORES (โหมด A) ====================== */

function getScores(subjectID, semester) {
  var sem = semester ? String(semester) : '';
  return readAll('Scores')
    .filter(function (r) {
      if (String(r.subjectID) !== String(subjectID)) return false;
      if (sem && r.semester && String(r.semester) !== sem) return false;
      return true;
    });
}

// บันทึกคะแนนรายหน่วยของรายวิชานี้ "เฉพาะภาคเรียนนั้น" (แทนที่) + คำนวณเกรดให้อัตโนมัติ
// scores: [{studentID, unitID, score}, ...]
function saveScores(subjectID, scores, semester) {
  var sem = String(semester || '1');
  assertNotLocked_((scores || []).map(function (r) { return r.studentID; }));
  clearRowsWhere2('Scores', 'subjectID', subjectID, 'semester', sem);
  var s = sheet('Scores');
  (scores || []).forEach(function (r) {
    s.appendRow([genId('SCR'), r.studentID, subjectID, r.unitID, r.score, sem]);
  });
  SpreadsheetApp.flush();
  return recomputeGrades(subjectID, sem);
}

// คำนวณเกรด (โหมด A) จากคะแนนรายหน่วยของรายวิชา+ภาคเรียน -> ร้อยละ -> เกรด 2551
function recomputeGrades(subjectID, semester) {
  var sem = String(semester || '1');
  var units = getUnits(subjectID, sem);
  var totalMax = units.reduce(function (sum, u) { return sum + Number(u.maxScore || 0); }, 0);
  var scores = getScores(subjectID, sem);
  var mapping = getGradeMapping();

  var byStudent = {};
  scores.forEach(function (sc) {
    var sid = sc.studentID;
    if (!byStudent[sid]) byStudent[sid] = 0;
    byStudent[sid] += Number(sc.score || 0);
  });

  var results = [];
  Object.keys(byStudent).forEach(function (sid) {
    var total = byStudent[sid];
    var percent = totalMax > 0 ? Math.round((total / totalMax) * 100) : 0;
    var g = percentToGrade(percent, mapping);
    writeGrade({
      studentID: sid, subjectID: subjectID, mode: 'A', semester: sem,
      percent: percent, level68: '', grade51: g.grade51, letter51: g.letter51, note: ''
    });
    results.push({ studentID: sid, total: total, totalMax: totalMax, percent: percent,
                   grade51: g.grade51, letter51: g.letter51, semester: sem });
  });
  SpreadsheetApp.flush();
  return results;
}


/** ====================== GRADES ====================== */

function getGrades(semester) {
  var rows = readAll('Grades');
  if (!semester) return rows;
  var sem = String(semester);
  return rows.filter(function (g) { return !g.semester || String(g.semester) === sem; });
}

// บันทึกเกรด 1 รายการ (โหมด B: กรอกร้อยละเอง / override เกรด)
function saveGrade(p) {
  var mapping = getGradeMapping();
  var sem      = String(p.semester || '1');
  var percent  = (p.percent !== '' && p.percent != null) ? Number(p.percent) : null;
  var grade51  = p.grade51;
  var letter51 = p.letter51;

  if (percent != null && !p.manualOverride) {
    var g = percentToGrade(percent, mapping);
    grade51 = g.grade51; letter51 = g.letter51;
  }
  writeGrade({
    studentID: p.studentID, subjectID: p.subjectID, mode: p.mode || 'B', semester: sem,
    percent: percent, level68: '', grade51: grade51, letter51: letter51, note: p.note || ''
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

// บันทึกเกรดหลายคนในครั้งเดียว (โหมด B ทั้งห้อง)
function saveGradesBatch(grades, semester) {
  assertNotLocked_((grades || []).map(function (p) { return p.studentID; }));
  var mapping = getGradeMapping();
  (grades || []).forEach(function (p) {
    var sem      = String(p.semester || semester || '1');
    var percent  = (p.percent !== '' && p.percent != null) ? Number(p.percent) : null;
    var grade51  = p.grade51;
    var letter51 = p.letter51;
    if (percent != null && !p.manualOverride) {
      var g = percentToGrade(percent, mapping);
      grade51 = g.grade51; letter51 = g.letter51;
    }
    writeGrade({
      studentID: p.studentID, subjectID: p.subjectID, mode: p.mode || 'B', semester: sem,
      percent: percent, level68: '', grade51: grade51, letter51: letter51, note: p.note || ''
    });
  });
  SpreadsheetApp.flush();
  return { count: (grades || []).length };
}

// upsert เกรดของ studentID+subjectID+semester
function writeGrade(g) {
  var s = sheet('Grades');
  var sem = String(g.semester || '1');
  var data = s.getDataRange().getValues();
  var found = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(g.studentID) &&
        String(data[i][2]) === String(g.subjectID) &&
        String(data[i][9] || '1') === sem) {
      found = i + 1; break;
    }
  }
  var rowVals = [
    found > 0 ? data[found - 1][0] : genId('GRD'),
    g.studentID, g.subjectID, g.mode,
    (g.percent == null ? '' : g.percent),
    '', g.grade51, g.letter51, g.note || '', sem
  ];
  if (found > 0) s.getRange(found, 1, 1, SHEETS.Grades.length).setValues([rowVals]);
  else           s.appendRow(rowVals);
}


/** ====================== ABILITIES (ความสามารถ/คุณลักษณะ/กิจกรรม/อ่านคิดวิเคราะห์เขียน) ====================== */

function getAbilities(studentID, semester) {
  return readAll('Abilities68')
    .filter(function (r) {
      if (String(r.studentID) !== String(studentID)) return false;
      if (semester && r.semester && String(r.semester) !== String(semester)) return false;
      return true;
    });
}

// items: [{category, itemKey, result}, ...]  (แทนที่ของนักเรียน "เฉพาะภาคเรียนนั้น")
function saveAbilities(studentID, items, semester) {
  var sem = String(semester || '1');
  clearRowsWhere2('Abilities68', 'studentID', studentID, 'semester', sem);
  var s = sheet('Abilities68');
  (items || []).forEach(function (it) {
    s.appendRow([genId('ABL'), studentID, it.category, it.itemKey, it.result, sem]);
  });
  SpreadsheetApp.flush();
  return getAbilities(studentID, sem);
}

/** ====================== WP1: การประเมินละเอียด (รายตัวชี้วัด) ====================== */
// scale: 0=ไม่ผ่าน, 1=ผ่าน, 2=ดี, 3=ดีเยี่ยม
function scoreToLevel(avg) {
  if (avg >= 2.5) return 'ดีเยี่ยม';
  if (avg >= 1.5) return 'ดี';
  if (avg >= 0.5) return 'ผ่าน';
  return 'ไม่ผ่าน';
}

// ดึงผลประเมินละเอียดของนักเรียนทั้งห้อง (สำหรับหน้ากรอกแบบตาราง)
function getAbilityDetailClass(classLevel, semester) {
  var ids = {};
  getStudents().forEach(function (s) { if (String(s.classLevel) === String(classLevel)) ids[s.ID] = true; });
  var sem = semester ? String(semester) : '';
  return readAll('AbilityDetail').filter(function (r) {
    if (!ids[r.studentID]) return false;
    if (sem && r.semester && String(r.semester) !== sem) return false;
    return true;
  });
}

// บันทึกผลประเมินละเอียดแบบกลุ่ม "เฉพาะภาคเรียนนั้น"
// p: { studentIDs:[...], rows:[{studentID, group, item, value}], semester }
function saveAbilityDetailBatch(p) {
  assertNotLocked_(p.studentIDs || []);
  var sem = String(p.semester || '1');
  var ids = {}; (p.studentIDs || []).forEach(function (id) { ids[id] = true; });
  var s = sheet('AbilityDetail');
  var data = s.getDataRange().getValues();
  var semCol = data[0].indexOf('semester');
  var keep = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i].join('') === '') continue;
    var rowSem = semCol >= 0 ? String(data[i][semCol] || '1') : '1';
    // คงไว้ ถ้าไม่ใช่ของนักเรียนชุดนี้ หรือเป็นคนละภาคเรียน
    if (!ids[data[i][1]] || rowSem !== sem) keep.push(data[i]);
  }
  (p.rows || []).forEach(function (r) {
    if (r.value === '' || r.value == null) return;
    keep.push([genId('ABD'), r.studentID, r.group, r.item, r.value, sem]);
  });
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, SHEETS.AbilityDetail.length).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, SHEETS.AbilityDetail.length).setValues(keep);
  SpreadsheetApp.flush();
  return { count: (p.rows || []).length };
}

// สรุปผลการประเมินจากรายตัวชี้วัด → object {itemKey: result} สำหรับใบรายงาน
function summarizeAbilityDetail(rows) {
  var ab = {};
  var cVals = [], rVals = [];
  (rows || []).forEach(function (r) {
    if (r.group === 'characteristic') cVals.push(String(r.value || ''));
    else if (r.group === 'readthink') rVals.push(String(r.value || ''));
    else if (r.group === 'activity' || r.group === 'competency') ab[r.item] = r.value;
  });
  // ผ่าน/ไม่ผ่าน: มีข้อใดไม่ผ่าน = ไม่ผ่าน, มีค่าครบและผ่านหมด = ผ่าน
  var passOf = function (vals) {
    var has = vals.some(function (v) { return v !== ''; });
    if (!has) return '';
    return vals.some(function (v) { return v === 'ไม่ผ่าน'; }) ? 'ไม่ผ่าน' : 'ผ่าน';
  };
  var c = passOf(cVals), r = passOf(rVals);
  if (c) ab['คุณลักษณะอันพึงประสงค์'] = c;
  if (r) ab['การอ่าน คิดวิเคราะห์ และเขียน'] = r;
  return ab;
}

/** ====================== แจ้งเตือนนักเรียนเสี่ยง ====================== */
/* ============================================================
   ภาพรวมรายห้อง (สำหรับผู้บริหาร) — สรุปความคืบหน้าแต่ละห้อง
   p: { semester? }
   ============================================================ */
function getDirectorOverview(p) {
  p = p || {};
  var sem = String(p.semester || getSettings().currentSemester || '1');
  var students = getStudents();
  var subjects = getSubjects();
  var grades = readAll('Grades').filter(function (g) {
    var gs = String(g.semester == null ? '' : g.semester);
    return gs === sem || gs === '';
  });
  var today = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
  var att = readAll('Attendance').filter(function (a) { return ymd(a.date) === today && String(a.period) === '0'; });

  var gradeKey = {};
  grades.forEach(function (g) { gradeKey[g.studentID + '|' + g.subjectID] = g; });

  var byClass = {};
  students.forEach(function (st) { (byClass[st.classLevel] = byClass[st.classLevel] || []).push(st); });
  var classes = Object.keys(byClass).filter(Boolean).sort();

  var rows = classes.map(function (cls) {
    var sts = byClass[cls];
    var subs = subjects.filter(function (su) { return !su.classLevel || classMatch(su.classLevel, cls); });
    var expected = sts.length * subs.length, filled = 0, risk = 0;
    sts.forEach(function (st) {
      var hasFail = false;
      subs.forEach(function (su) {
        var g = gradeKey[st.ID + '|' + su.ID];
        if (g && g.grade51 !== '' && g.grade51 != null) {
          filled++;
          if (String(g.grade51) === '0') hasFail = true;
        } else if (g && g.percent !== '' && g.percent != null && Number(g.percent) < 50) {
          hasFail = true;
        }
      });
      if (hasFail) risk++;
    });
    var attRows = att.filter(function (a) { return String(a.classLevel) === cls; });
    var present = attRows.filter(function (a) { return a.status === 'มา' || a.status === 'สาย'; }).length;
    var homeroom = readAll('Teachers').filter(function (t) { return String(t.homeroomClass) === cls; }).map(function (t) { return t.name; });
    return {
      classLevel: cls, students: sts.length, subjects: subs.length,
      gradePct: expected ? Math.round(filled / expected * 100) : 0,
      atRisk: risk, attChecked: attRows.length > 0, attPresent: present, attTotal: sts.length,
      homeroom: homeroom
    };
  });

  return {
    semester: sem, classes: rows,
    totals: {
      classes: rows.length, students: students.length,
      teachers: getTeachers().filter(function (t) { return t.role !== 'admin' && t.role !== 'director'; }).length,
      atRisk: rows.reduce(function (a, r) { return a + r.atRisk; }, 0)
    }
  };
}

/* ============================================================
   ระบบติดตามนักเรียนเสี่ยง (รายวิชา) — ณ วันปัจจุบัน
   1) เสี่ยงผลการเรียน: % สะสมของวิชา (คะแนนที่กรอกแล้ว / เต็มของหน่วยที่กรอกแล้ว) < เกณฑ์ (ค่าตั้งต้น 50)
   2) เสี่ยงติด มส.: เวลาเรียนรายวิชา (คาบที่มา/คาบที่ต้องเรียนตามตารางสอน) < เกณฑ์ (ค่าตั้งต้น 80)
   p: { classLevel?, semester?, gradeMin=50, attMin=80, critical=40 }
   ============================================================ */
function getRiskStudents(p) {
  p = p || {};
  var thAtt   = (p.attMin   != null) ? Number(p.attMin)   : 80; // % เวลาเรียนขั้นต่ำ (มส.)
  var thGrade = (p.gradeMin != null) ? Number(p.gradeMin) : 50; // % ผลการเรียนขั้นต่ำ
  var thCrit  = (p.critical != null) ? Number(p.critical) : 40; // % วิกฤต
  var semester = p.semester ? String(p.semester) : '';          // '' = ทุกภาคเรียน

  var students = getStudents();
  if (p.classLevel) students = students.filter(function (s) { return String(s.classLevel) === String(p.classLevel); });
  students.sort(function (a, b) { return String(a.classLevel).localeCompare(String(b.classLevel)) || Number(a.number) - Number(b.number); });

  // ---------- ผลการเรียนสะสม (รายคน-รายวิชา) ----------
  var subjects = getSubjects();
  var unitMax = {};
  readAll('LearningUnits').forEach(function (u) {
    if (semester && u.semester && String(u.semester) !== semester) return;
    unitMax[u.ID] = Number(u.maxScore || 0);
  });
  var scAgg = {}; // sid|subjectID -> {got, max}
  readAll('Scores').forEach(function (sc) {
    if (semester && sc.semester && String(sc.semester) !== semester) return;
    var mx = unitMax[sc.unitID]; if (mx == null) return;       // หน่วยไม่อยู่ในภาคนี้/ไม่มีคะแนนเต็ม
    var k = sc.studentID + '|' + sc.subjectID;
    var a = scAgg[k] || (scAgg[k] = { got: 0, max: 0 });
    a.got += Number(sc.score || 0);
    a.max += mx;                                                // นับเฉพาะหน่วยที่ "ทำแล้ว" = ณ วันนี้
  });

  // ---------- เวลาเรียนรายคาบ (จากตารางสอน + การเช็คชื่อ) ----------
  var att = readAll('Attendance');
  var range = semesterRange_(getSettings(), semester);
  if (range) att = att.filter(function (r) { var d = ymd(r.date); return d >= range.from && d <= range.to; });
  var look = {}, flag = {}, datesByClass = {};
  att.forEach(function (r) {
    var c = String(r.classLevel), d = ymd(r.date);
    look[c + '|' + d + '|' + r.period + '|' + r.studentID] = r.status;
    if (String(r.period) === '0') { flag[c + '|' + d + '|' + r.studentID] = r.status; (datesByClass[c] = datesByClass[c] || {})[d] = true; }
  });
  var schByClassDay = {}; // [class][weekday][period] = subjectName
  readAll('Schedule').forEach(function (s) {
    var c = String(s.classLevel);
    (schByClassDay[c] = schByClassDay[c] || {});
    (schByClassDay[c][s.day] = schByClassDay[c][s.day] || {})[Number(s.period)] = s.subject;
  });

  var out = [];
  students.forEach(function (st) {
    var c = String(st.classLevel);

    // (1) ผลการเรียนต่ำกว่าเกณฑ์ — รายวิชา
    var lowSubjects = [];
    subjects.forEach(function (su) {
      if (su.classLevel && !classMatch(su.classLevel, c)) return;
      var a = scAgg[st.ID + '|' + su.ID];
      if (!a || a.max <= 0) return;                              // ยังไม่มีคะแนน = ยังไม่ประเมิน
      var pct = Math.round(a.got / a.max * 100);
      if (pct < thGrade) lowSubjects.push({ subject: su.name || su.code || '-', percent: pct });
    });

    // (2) เวลาเรียนต่ำกว่าเกณฑ์ (เสี่ยง มส.) — รายวิชา ตามคาบในตารางสอน
    var msSubjects = [];
    var bySubj = {}; // subjectName -> {exp, pres}
    var sch = schByClassDay[c] || {};
    Object.keys(datesByClass[c] || {}).forEach(function (d) {
      var wd = WD_TH[new Date(d + 'T00:00:00').getDay()];
      var periods = sch[wd] || {};
      Object.keys(periods).forEach(function (per) {
        var nm = periods[per];
        var b = bySubj[nm] || (bySubj[nm] = { exp: 0, pres: 0 });
        b.exp++;
        var status = look[c + '|' + d + '|' + per + '|' + st.ID] || flag[c + '|' + d + '|' + st.ID];
        if (status === 'มา' || status === 'สาย') b.pres++;
      });
    });
    Object.keys(bySubj).forEach(function (nm) {
      var b = bySubj[nm];
      if (b.exp <= 0) return;
      var pct = Math.round(b.pres / b.exp * 100);
      if (pct < thAtt) {
        msSubjects.push({ subject: nm, percent: pct, present: b.pres, expected: b.exp });
      }
    });

    if (lowSubjects.length || msSubjects.length) {
      var minPct = lowSubjects.length ? lowSubjects.reduce(function (m, x) { return Math.min(m, x.percent); }, 100) : 100;
      var level = (lowSubjects.length && minPct < thCrit) ? 'วิกฤต' : 'เสี่ยง';
      out.push({
        studentID: st.ID, number: st.number, fullName: st.fullName, classLevel: st.classLevel,
        level: level, minPercent: lowSubjects.length ? minPct : null,
        lowSubjects: lowSubjects, msSubjects: msSubjects
      });
    }
  });

  // เรียง: วิกฤตก่อน แล้วตาม % ต่ำสุด
  out.sort(function (a, b) {
    if (a.level !== b.level) return a.level === 'วิกฤต' ? -1 : 1;
    return (a.minPercent == null ? 999 : a.minPercent) - (b.minPercent == null ? 999 : b.minPercent);
  });
  var counts = { grade: 0, ms: 0, critical: 0, total: out.length };
  out.forEach(function (r) { if (r.lowSubjects.length) counts.grade++; if (r.msSubjects.length) counts.ms++; if (r.level === 'วิกฤต') counts.critical++; });
  return { thresholds: { gradeMin: thGrade, attMin: thAtt, critical: thCrit }, semester: semester, counts: counts, total: out.length, students: out };
}

// ช่วงวันที่ของภาคเรียนจาก Settings (sem1Start/sem1End/sem2Start/sem2End) — คืน null ถ้าไม่ได้ตั้ง
function semesterRange_(set, semester) {
  if (!semester) return null;
  var s = set['sem' + semester + 'Start'], e = set['sem' + semester + 'End'];
  if (!s || !e) return null;
  return { from: String(s), to: String(e) };
}

/** ====================== ล็อก/อนุมัติผลการเรียน (ต่อห้อง) ====================== */
function getLocks() { return readAll('Locks'); }
function setLock(p) {
  var s = sheet('Locks');
  var data = s.getDataRange().getValues(); var row = -1;
  for (var i = 1; i < data.length; i++) { if (String(data[i][0]) === String(p.classLevel)) { row = i + 1; break; } }
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var rec = [p.classLevel, p.locked ? 'locked' : 'open', now, p.by || ''];
  if (row < 0) s.appendRow(rec); else s.getRange(row, 1, 1, 4).setValues([rec]);
  SpreadsheetApp.flush();
  return getLocks();
}
function lockedRoomsSet_() {
  var o = {}; readAll('Locks').forEach(function (r) { if (String(r.status) === 'locked') o[String(r.classLevel)] = true; });
  return o;
}
// ตรวจว่ามีนักเรียนอยู่ในห้องที่ถูกล็อกหรือไม่ → ถ้ามี โยน error
function assertNotLocked_(studentIDs) {
  var locked = lockedRoomsSet_();
  if (!Object.keys(locked).length) return;
  var room = {}; getStudents().forEach(function (s) { room[s.ID] = String(s.classLevel); });
  var hit = {};
  (studentIDs || []).forEach(function (id) { var r = room[id]; if (r && locked[r]) hit[r] = true; });
  var rooms = Object.keys(hit);
  if (rooms.length) throw new Error('ห้อง ' + rooms.join(', ') + ' ถูกล็อก (อนุมัติผลแล้ว) — ให้แอดมินปลดล็อกก่อนจึงจะแก้ไขได้');
}

/** ====================== ปฏิทินวิชาการ ====================== */
function getCalendar() {
  return readAll('Calendar').map(function (r) { r.date = ymd(r.date); return r; })
    .sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
}
function addCalendarEvent(p) {
  if (!p.date || !p.type) throw new Error('ระบุวันที่และประเภท');
  var s = sheet('Calendar');
  s.appendRow([genId('CAL'), '', p.type, p.title || '', p.note || '']);
  // เขียน date เป็นข้อความ (กัน Sheets แปลงเป็น Date)
  var row = s.getLastRow();
  s.getRange(row, 2).setNumberFormat('@').setValue(ymd(p.date));
  SpreadsheetApp.flush();
  return getCalendar();
}
function deleteCalendarEvent(id) {
  var row = findRowById('Calendar', id);
  if (row < 0) throw new Error('ไม่พบรายการ');
  sheet('Calendar').deleteRow(row);
  SpreadsheetApp.flush();
  return getCalendar();
}

/** ====================== เล่มสรุปผลการเรียน (รวมทุกหน้า) ====================== */
function getPP5Book(p) {
  p = p || {};
  var classLevel = p.classLevel;
  var subjectMode = p.scope === 'subject';
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });

  var subjects = getSubjects().filter(function (sub) { return !sub.classLevel || classMatch(sub.classLevel, classLevel); })
    .sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
  if (subjectMode && p.subjectID) subjects = subjects.filter(function (s) { return String(s.ID) === String(p.subjectID); });

  var gradeBy = {};
  getGrades().forEach(function (g) { (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });

  var docsBy = {};
  subjects.forEach(function (sub) { docsBy[sub.ID] = getSubjectDocs(sub.ID); });

  var from = p.from || '2000-01-01', to = p.to || '2999-12-31';
  var register = getAttendanceRegister(classLevel, from, to);
  var grids = register.months.map(function (m) { return getMonthlyGrid(classLevel, m.ym); });

  var detail = readAll('AbilityDetail');
  var byStu = {};
  detail.forEach(function (r) { (byStu[r.studentID] = byStu[r.studentID] || []).push(r); });
  var abilityRaw = students.map(function (st) { return { studentID: st.ID, rows: byStu[st.ID] || [] }; });

  return {
    settings: getSettings(),
    classLevel: classLevel,
    scope: subjectMode ? 'subject' : 'class',
    cover: getPP5Cover(classLevel),
    stats: getGradeStats({ classLevel: classLevel }),
    students: students.map(function (s) {
      return { ID: s.ID, studentCode: s.studentCode, number: s.number, fullName: s.fullName,
        prefix: s.prefix, firstName: s.firstName, lastName: s.lastName, classLevel: s.classLevel };
    }),
    subjects: subjects.map(function (s) { return { ID: s.ID, code: s.code, name: s.name, classLevel: s.classLevel }; }),
    docs: docsBy,
    grades: students.map(function (st) {
      return { studentID: st.ID, bysubject: subjects.map(function (sub) {
        var g = (gradeBy[st.ID] || {})[sub.ID] || {};
        return { subjectID: sub.ID, grade51: g.grade51, level68: g.level68, percent: g.percent };
      }) };
    }),
    register: register,
    grids: grids,
    abilityRaw: abilityRaw
  };
}

/** ====================== WP3: ปกเล่ม สรุปผลรายห้อง ====================== */

function getPP5Cover(classLevel) {
  var settings = getSettings();
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  var ids = {}; students.forEach(function (s) { ids[s.ID] = true; });

  // นับชาย/หญิง จากคำนำหน้า
  var male = 0, female = 0;
  students.forEach(function (s) {
    var n = String(s.fullName || '');
    if (/^(เด็กชาย|ด\.ช\.|นาย)/.test(n)) male++;
    else if (/^(เด็กหญิง|ด\.ญ\.|นางสาว|นาง)/.test(n)) female++;
  });

  // การกระจายผลการเรียน (เกรด 2551) รายวิชา
  var GR = ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0'];
  var subs = getSubjects().filter(function (sub) {
    return !sub.classLevel || classMatch(sub.classLevel, classLevel);
  }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
  var gradeBy = {};
  readAll('Grades').forEach(function (g) { if (ids[g.studentID]) (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });
  var subjects = subs.map(function (sub) {
    var counts = {}; GR.forEach(function (k) { counts[k] = 0; });
    students.forEach(function (st) {
      var g = (gradeBy[st.ID] || {})[sub.ID];
      if (g && g.grade51 !== '' && g.grade51 != null && counts[String(g.grade51)] !== undefined) counts[String(g.grade51)]++;
    });
    return { code: sub.code || '', name: sub.name, counts: GR.map(function (k) { return counts[k]; }) };
  });

  // สรุปผลประเมินจากรายตัวชี้วัด
  var detail = readAll('AbilityDetail').filter(function (r) { return ids[r.studentID]; });
  var byStu = {};
  detail.forEach(function (r) { (byStu[r.studentID] = byStu[r.studentID] || []).push(r); });
  var LV = ['ผ่าน', 'ไม่ผ่าน'];
  var charCounts = { 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
  var rtCounts = { 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
  var actPass = 0, actFail = 0;
  students.forEach(function (st) {
    var rows = byStu[st.ID] || [];
    var ab = summarizeAbilityDetail(rows);
    if (ab['คุณลักษณะอันพึงประสงค์']) charCounts[ab['คุณลักษณะอันพึงประสงค์']]++;
    if (ab['การอ่าน คิดวิเคราะห์ และเขียน']) rtCounts[ab['การอ่าน คิดวิเคราะห์ และเขียน']]++;
    var acts = rows.filter(function (r) { return r.group === 'activity' && r.item !== 'ชื่อชุมนุม'; });
    if (acts.length) {
      var fail = acts.some(function (r) { return r.value === 'ไม่ผ่าน'; });
      if (fail) actFail++; else actPass++;
    }
  });

  // ครูประจำชั้น (จากหน้าจัดการครู)
  var homeroom = readAll('Teachers')
    .filter(function (t) { return String(t.homeroomClass) === String(classLevel); })
    .map(function (t) { return t.name; });

  return {
    settings: settings, classLevel: classLevel,
    totals: { all: students.length, male: male, female: female },
    gradeLevels: GR, subjects: subjects,
    levels: LV,
    charCounts: LV.map(function (k) { return charCounts[k]; }),
    rtCounts: LV.map(function (k) { return rtCounts[k]; }),
    activity: { pass: actPass, fail: actFail },
    homeroomTeachers: homeroom
  };
}


/** ====================== REPORT DATA ====================== */

// รวมข้อมูลทั้งหมดของนักเรียน 1 คน สำหรับพิมพ์ใบรายงาน A4 (2 หลักสูตร)
function getReportData(studentID, semester) {
  var students = getStudents();
  var student = students.filter(function (s) { return String(s.ID) === String(studentID); })[0];
  if (!student) throw new Error('ไม่พบนักเรียน ID: ' + studentID);

  var subjects = getSubjects().filter(function (sub) {
    return !sub.classLevel || classMatch(sub.classLevel, student.classLevel);
  }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });

  var grades = readAll('Grades').filter(function (g) { return String(g.studentID) === String(studentID); });

  // บล็อกผลการเรียนของภาคเรียนหนึ่ง ๆ
  function semBlock(sem) {
    var gmap = {};
    grades.filter(function (g) { return String(g.semester || '1') === String(sem); })
      .forEach(function (g) { gmap[g.subjectID] = g; });
    var rows = subjects.map(function (sub) {
      var g = gmap[sub.ID] || {};
      return {
        subjectID: sub.ID, code: sub.code, name: sub.name,
        learningArea: sub.learningArea, hours: sub.hours, type: sub.type,
        percent: g.percent, grade51: g.grade51, letter51: g.letter51 || ''
      };
    });
    // GPA รายภาค: เฉลี่ยเกรด (ถ่วงน้ำหนักด้วยหน่วยกิตจากชั่วโมง/40 ถ้ามี ไม่งั้นเฉลี่ยปกติ)
    var sumGP = 0, sumW = 0;
    rows.forEach(function (r) {
      var v = Number(r.grade51);
      if (r.grade51 === '' || r.grade51 == null || isNaN(v)) return;
      var w = 1; // ประถม: ถ่วงเท่ากันทุกวิชา
      sumGP += v * w; sumW += w;
    });
    return { rows: rows, gpa: sumW ? Math.round((sumGP / sumW) * 100) / 100 : 0, countGraded: sumW };
  }

  var sem1 = semBlock('1');
  var sem2 = semBlock('2');
  var allGP = sem1.gpa * sem1.countGraded + sem2.gpa * sem2.countGraded;
  var allW = sem1.countGraded + sem2.countGraded;
  var gpaYear = allW ? Math.round((allGP / allW) * 100) / 100 : 0;

  // เกรดรวม (ใช้ตอนยังไม่แยกภาค — คงไว้เพื่อความเข้ากันได้)
  var gradeMapAll = {};
  grades.forEach(function (g) { gradeMapAll[g.subjectID] = g; });
  var subjectRows = subjects.map(function (sub) {
    var g = gradeMapAll[sub.ID] || {};
    return {
      subjectID: sub.ID, code: sub.code, name: sub.name,
      learningArea: sub.learningArea, hours: sub.hours, type: sub.type,
      percent: g.percent, grade51: g.grade51, letter51: g.letter51 || ''
    };
  });

  // คุณลักษณะ / อ่านคิดวิเคราะห์เขียน (สรุปจาก AbilityDetail แยกภาค) — แบบ 2551
  function assess(sem) {
    var rows = readAll('AbilityDetail').filter(function (r) {
      return String(r.studentID) === String(studentID) && String(r.semester || '1') === String(sem);
    });
    return summarizeAbilityDetail(rows);
  }

  return {
    settings: getSettings(),
    student: student,
    subjects: subjectRows,          // คงไว้ (เข้ากันได้กับของเดิม)
    abilities: getAbilities(studentID, semester),
    abilityGroups: [],              // ตัดส่วนความสามารถ 8 ด้าน/ระดับพัฒนาการ (หลักสูตร 2568) ออก
    semester1: sem1,
    semester2: sem2,
    summary: { gpa1: sem1.gpa, gpa2: sem2.gpa, gpaYear: gpaYear },
    assessment: { s1: assess('1'), s2: assess('2') }
  };
}

// ชั้นตรงกันไหม รองรับ "ป.3", "ป.3/1" ฯลฯ (จับคู่ตามตัวเลขชั้น)
function classMatch(subjectClass, studentClass) {
  if (!subjectClass || !studentClass) return false;
  var a = String(subjectClass).replace(/[^0-9]/g, '').charAt(0);
  var b = String(studentClass).replace(/[^0-9]/g, '').charAt(0);
  return a && b && a === b;
}

// รวมข้อมูลนักเรียนทั้งชั้น สำหรับพิมพ์ใบรายงานทีเดียว (เรียก API ครั้งเดียว)
function getReportBatch(classLevel) {
  var students = getStudents();
  if (classLevel) {
    if (String(classLevel).indexOf('/') >= 0) {
      // ระบุห้องเต็ม เช่น ป.6/1 → เฉพาะห้องนั้น
      students = students.filter(function (s) { return String(s.classLevel) === String(classLevel); });
    } else {
      students = students.filter(function (s) { return classMatch(classLevel, s.classLevel); });
    }
  }
  students.sort(function (a, b) {
    return String(a.classLevel).localeCompare(String(b.classLevel)) || Number(a.number) - Number(b.number);
  });

  var settings = getSettings();
  var allSubjects = getSubjects();
  var allGrades = getGrades();
  var allAbilDetail = readAll('AbilityDetail');
  var allTeachers = readAll('Teachers');

  var gradeBy = {};
  allGrades.forEach(function (g) { (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });
  var abilDetailBy = {};
  allAbilDetail.forEach(function (a) { (abilDetailBy[a.studentID] = abilDetailBy[a.studentID] || []).push(a); });
  // ครูประจำชั้นตามห้อง (ห้อง → รายชื่อครู)
  var homeroomBy = {};
  allTeachers.forEach(function (t) {
    if (t.homeroomClass) (homeroomBy[t.homeroomClass] = homeroomBy[t.homeroomClass] || []).push(t.name);
  });

  // สรุปเวลาเรียน (หน้าเสาธง period 0) ต่อนักเรียน
  var attRec = readAll('Attendance').filter(function (r) { return String(r.period) === '0'; });
  var schoolDaysByClass = {}, attLook = {};
  attRec.forEach(function (r) {
    var d = ymd(r.date), c = String(r.classLevel);
    (schoolDaysByClass[c] = schoolDaysByClass[c] || {})[d] = true;
    attLook[c + '|' + d + '|' + r.studentID] = r.status;
  });

  var out = students.map(function (student) {
    var subs = allSubjects.filter(function (sub) {
      return !sub.classLevel || classMatch(sub.classLevel, student.classLevel);
    }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
    var rows = subs.map(function (sub) {
      var gg = (gradeBy[student.ID] || {})[sub.ID] || {};
      return {
        code: sub.code, name: sub.name, learningArea: sub.learningArea, hours: sub.hours, type: sub.type,
        percent: gg.percent, grade51: gg.grade51, letter51: gg.letter51 || ''
      };
    });
    var abObj = summarizeAbilityDetail(abilDetailBy[student.ID] || []);
    var abilities = Object.keys(abObj).map(function (k) { return { itemKey: k, result: abObj[k] }; });
    // เวลาเรียนของนักเรียนคนนี้
    var c = String(student.classLevel);
    var days = Object.keys(schoolDaysByClass[c] || {});
    var present = 0, absent = 0, leave = 0, late = 0;
    days.forEach(function (d) {
      var s = attLook[c + '|' + d + '|' + student.ID];
      if (_attendedDay(s)) present++;
      if (s === 'ขาด') absent++; else if (s === 'ลา') leave++; else if (s === 'สาย') late++;
    });
    var attendance = { schoolDays: days.length, present: present, absent: absent, leave: leave, late: late,
      percent: days.length ? Math.round(present / days.length * 100) : 0 };
    return { student: student, subjects: rows, abilities: abilities, abilityGroups: [],
      homeroomTeachers: homeroomBy[student.classLevel] || [], attendance: attendance };
  });

  return { settings: settings, students: out };
}

// ชั้นตรงกันไหม (เดิม) — คงไว้สำหรับ getReportData
function classMatch_legacy(subjectClass, studentClass) {
  return classMatch(subjectClass, studentClass);
}


/** ====================== WP3: สรุปผลการเรียนรายชั้น (หน้าปก) ====================== */
function getClassGradeSummary(classLevel, semester) {
  var settings = getSettings();
  var sem = semester ? String(semester) : '';
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });
  var subjects = getSubjects().filter(function (sub) {
    return !sub.classLevel || classMatch(sub.classLevel, classLevel);
  }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });

  var gradeBy = {};
  readAll('Grades').filter(function (g) { return !sem || String(g.semester || '1') === sem; })
    .forEach(function (g) { (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });
  var abBy = {};
  readAll('AbilityDetail').filter(function (a) { return !sem || String(a.semester || '1') === sem; })
    .forEach(function (a) { (abBy[a.studentID] = abBy[a.studentID] || []).push(a); });

  var LEVELS = ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0'];
  var subjRows = subjects.map(function (sub) {
    var counts = {}; LEVELS.forEach(function (l) { counts[l] = 0; });
    students.forEach(function (st) {
      var g = (gradeBy[st.ID] || {})[sub.ID];
      if (g && g.grade51 !== '' && g.grade51 != null) {
        var k = String(g.grade51);
        if (counts[k] !== undefined) counts[k]++;
      }
    });
    return { code: sub.code, name: sub.name, counts: counts };
  });

  var charCount = { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
  var rtCount = { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
  var actPass = 0, actFail = 0;
  students.forEach(function (st) {
    var ab = summarizeAbilityDetail(abBy[st.ID] || []);
    var c = ab['คุณลักษณะอันพึงประสงค์']; if (c && charCount[c] !== undefined) charCount[c]++;
    var r = ab['การอ่าน คิดวิเคราะห์ และเขียน']; if (r && rtCount[r] !== undefined) rtCount[r]++;
    var acts = Object.keys(ab).filter(function (k) { return /กิจกรรม|ลูกเสือ|ชุมนุม|ชมรม|เนตรนารี/.test(k); });
    if (acts.length) { if (acts.every(function (k) { return ab[k] === 'ผ่าน'; })) actPass++; else actFail++; }
  });

  var male = 0, female = 0;
  students.forEach(function (st) {
    var n = String(st.fullName || '');
    if (/^(เด็กชาย|นาย|ด\.ช)/.test(n)) male++;
    else if (/^(เด็กหญิง|นางสาว|นาง|ด\.ญ)/.test(n)) female++;
  });

  var homeroom = readAll('Teachers').filter(function (t) {
    return t.homeroomClass && String(t.homeroomClass) === String(classLevel);
  }).map(function (t) { return t.name; });

  return {
    settings: settings, classLevel: classLevel, subjects: subjRows,
    characteristic: charCount, readthink: rtCount, activity: { pass: actPass, fail: actFail },
    total: students.length, male: male, female: female, homeroomTeachers: homeroom
  };
}


/** ====================== STATS ====================== */

/** ====================== แผนภูมิสรุปผลรายชั้น-รายวิชา ====================== */
function getGradeStats(p) {
  p = p || {};
  var GR = ['4', '3.5', '3', '2.5', '2', '1.5', '1', '0'];
  var students = getStudents();
  if (p.classLevel) {
    if (String(p.classLevel).indexOf('/') >= 0)
      students = students.filter(function (s) { return String(s.classLevel) === String(p.classLevel); });
    else
      students = students.filter(function (s) { return classMatch(p.classLevel, s.classLevel); });
  }
  var inScope = {}, roomOf = {};
  students.forEach(function (s) { inScope[s.ID] = true; roomOf[s.ID] = String(s.classLevel); });

  var subjects = getSubjects();
  var subMap = {}; subjects.forEach(function (s) { subMap[s.ID] = s; });
  var sem = p.semester ? String(p.semester) : '';
  var grades = readAll('Grades').filter(function (g) {
    return inScope[g.studentID] && (!sem || String(g.semester || '1') === sem);
  });

  // ต่อรายวิชา: นับเกรด + เฉลี่ย
  var bySub = {};
  // ต่อห้อง: เฉลี่ยรวม
  var byRoom = {};
  var overall = {}; GR.forEach(function (k) { overall[k] = 0; });
  grades.forEach(function (g) {
    var k = String(g.grade51 == null ? '' : g.grade51);
    if (GR.indexOf(k) < 0) return;
    var num = Number(k);
    var sub = subMap[g.subjectID]; if (!sub) return;
    var b = bySub[g.subjectID] = bySub[g.subjectID] || { code: sub.code, name: sub.name, classLevel: sub.classLevel, counts: {}, sum: 0, n: 0 };
    b.counts[k] = (b.counts[k] || 0) + 1; b.sum += num; b.n++;
    overall[k]++;
    var rm = roomOf[g.studentID];
    var r = byRoom[rm] = byRoom[rm] || { sum: 0, n: 0 };
    r.sum += num; r.n++;
  });

  var subjectsOut = Object.keys(bySub).map(function (id) {
    var b = bySub[id];
    return { code: b.code, name: b.name, classLevel: b.classLevel,
      counts: GR.map(function (k) { return b.counts[k] || 0; }),
      avg: b.n ? Number((b.sum / b.n).toFixed(2)) : 0, n: b.n };
  }).sort(function (a, b) { return String(a.classLevel).localeCompare(String(b.classLevel)) || String(a.code).localeCompare(String(b.code)); });

  var roomsOut = Object.keys(byRoom).sort().map(function (rm) {
    var r = byRoom[rm];
    return { room: rm, avg: r.n ? Number((r.sum / r.n).toFixed(2)) : 0, n: r.n };
  });

  return { gradeLevels: GR, subjects: subjectsOut, rooms: roomsOut, overall: GR.map(function (k) { return overall[k]; }) };
}

function getStats() {
  var students = getStudents();
  var grades = getGrades();
  var byLevel = {};
  grades.forEach(function (g) {
    var k = (g.grade51 === '' || g.grade51 == null) ? 'ยังไม่ประเมิน' : ('เกรด ' + g.grade51);
    byLevel[k] = (byLevel[k] || 0) + 1;
  });
  return {
    studentCount: students.length,
    subjectCount: getSubjects().length,
    gradeCount: grades.length,
    byLevel: byLevel
  };
}


/** ====================== EXPORT SCHOOL MIS ====================== */

// คืนข้อมูลเกรดสำหรับนำไปลง School MIS (frontend แปลงเป็น Excel/CSV)
// classLevel ว่าง = ทั้งหมด
// ส่งออกแบบ wide ตรงรูปแบบ SchoolMIS: 1 แถว/คน, คอลัมน์วิชา=เกรด 2551, กิจกรรม=ผ/มผ
function exportSchoolMISWide(classLevel) {
  var students = getStudents();
  if (classLevel) students = students.filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });

  // วิชาของชั้นนี้ (เรียงตาม sortOrder)
  var subjects = getSubjects().filter(function (sub) {
    return !sub.classLevel || classMatch(sub.classLevel, classLevel);
  }).sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });

  // กิจกรรมพัฒนาผู้เรียน (SchoolMIS ใช้ ผ/มผ)
  var ACTS = ['กิจกรรมแนะแนว', 'ลูกเสือ – เนตรนารี', 'ชมรม/ชุมนุม', 'กิจกรรมเพื่อสังคมและสาธารณประโยชน์'];
  var ACT_LABEL = { 'กิจกรรมแนะแนว': 'แนะแนว', 'ลูกเสือ – เนตรนารี': 'ลูกเสือ-เนตรนารี', 'ชมรม/ชุมนุม': 'ชุมนุม', 'กิจกรรมเพื่อสังคมและสาธารณประโยชน์': 'กิจกรรมเพื่อสังคมและสาธารณประโยชน์' };

  // เกรดต่อคนต่อวิชา
  var gradeBy = {};
  getGrades().forEach(function (g) { (gradeBy[g.studentID] = gradeBy[g.studentID] || {})[g.subjectID] = g; });

  // กิจกรรมต่อคน (จาก AbilityDetail group=activity)
  var actBy = {};
  readAll('AbilityDetail').forEach(function (r) {
    if (r.group === 'activity') (actBy[r.studentID] = actBy[r.studentID] || {})[r.item] = r.value;
  });

  // หัวตาราง
  var header = ['#', 'รหัสนักเรียน', 'ชื่อ-สกุล'];
  subjects.forEach(function (sub) { header.push(((sub.code || '') + ' ' + sub.name).trim()); });
  ACTS.forEach(function (a) { header.push(ACT_LABEL[a]); });

  var rows = students.map(function (st) {
    var row = [st.number, st.studentCode, st.fullName];
    subjects.forEach(function (sub) {
      var g = (gradeBy[st.ID] || {})[sub.ID];
      row.push(g && g.grade51 != null && g.grade51 !== '' ? g.grade51 : '');
    });
    ACTS.forEach(function (a) {
      var v = (actBy[st.ID] || {})[a];
      row.push(v === 'ผ่าน' ? 'ผ' : (v === 'ไม่ผ่าน' ? 'มผ' : ''));
    });
    return row;
  });

  return { header: header, rows: rows };
}

function exportSchoolMIS(classLevel) {
  var students = getStudents();
  if (classLevel) {
    students = students.filter(function (s) { return classMatch(classLevel, s.classLevel); });
  }
  var subjects = getSubjects();
  var subMap = {};
  subjects.forEach(function (s) { subMap[s.ID] = s; });
  var grades = getGrades();

  var out = [];
  grades.forEach(function (g) {
    var st = students.filter(function (s) { return String(s.ID) === String(g.studentID); })[0];
    if (!st) return;
    var sub = subMap[g.subjectID];
    if (!sub) return;
    out.push({
      studentCode: st.studentCode,
      number: st.number,
      fullName: st.fullName,
      classLevel: st.classLevel,
      subjectCode: sub.code,
      subjectName: sub.name,
      percent: g.percent,
      grade51: g.grade51,
      level68: g.level68
    });
  });
  return out;
}


/** ====================== BACKUP / IMPORT / CLEAR ====================== */

function backupAllData() {
  var dump = {};
  Object.keys(SHEETS).forEach(function (name) { dump[name] = readAll(name); });
  dump._meta = { exportedAt: new Date().toString(), version: 1 };
  return dump;
}

/** ====================== สำรองข้อมูลอัตโนมัติลง Drive ====================== */
var AUTO_BACKUP_FOLDER = 'สำรองข้อมูล - ระบบรายงานผลการเรียน';
var AUTO_BACKUP_KEEP = 30; // เก็บไฟล์ย้อนหลังสูงสุด

// สร้างไฟล์สำรอง (เรียกจากปุ่มหรือ trigger) → คืนข้อมูลไฟล์ล่าสุด
function runBackupToDrive() {
  var dump = backupAllData();
  var folder = getOrCreateFolder(AUTO_BACKUP_FOLDER);
  var stamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd_HHmm');
  var name = 'backup_' + stamp + '.json';
  var file = folder.createFile(Utilities.newBlob(JSON.stringify(dump), 'application/json', name));
  // ลบไฟล์เก่าเกินจำนวนที่กำหนด
  pruneOldBackups_(folder);
  PropertiesService.getScriptProperties().setProperty('LAST_BACKUP', new Date().toISOString());
  return { name: name, id: file.getId(), at: new Date().toString() };
}
function pruneOldBackups_(folder) {
  var files = [];
  var it = folder.getFilesByType('application/json');
  while (it.hasNext()) { var f = it.next(); files.push({ f: f, t: f.getDateCreated().getTime() }); }
  files.sort(function (a, b) { return b.t - a.t; });
  for (var i = AUTO_BACKUP_KEEP; i < files.length; i++) { try { files[i].f.setTrashed(true); } catch (e) {} }
}

// เปิด/ปิด trigger สำรองอัตโนมัติรายวัน
function setAutoBackup(p) {
  var enable = p && p.enable;
  var hour = (p && p.hour != null) ? Number(p.hour) : 22;
  // ลบ trigger เดิมของ runBackupToDrive ก่อนเสมอ
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runBackupToDrive') ScriptApp.deleteTrigger(t);
  });
  if (enable) {
    ScriptApp.newTrigger('runBackupToDrive').timeBased().everyDays(1).atHour(hour).create();
  }
  PropertiesService.getScriptProperties().setProperty('AUTO_BACKUP_ON', enable ? '1' : '0');
  PropertiesService.getScriptProperties().setProperty('AUTO_BACKUP_HOUR', String(hour));
  return getBackupStatus();
}

function getBackupStatus() {
  var pr = PropertiesService.getScriptProperties();
  var on = pr.getProperty('AUTO_BACKUP_ON') === '1';
  // ตรวจ trigger จริง
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'runBackupToDrive'; });
  var folder = null, list = [];
  try {
    var it = DriveApp.getFoldersByName(AUTO_BACKUP_FOLDER);
    if (it.hasNext()) {
      folder = it.next();
      var fit = folder.getFilesByType('application/json'), arr = [];
      while (fit.hasNext()) { var f = fit.next(); arr.push({ name: f.getName(), id: f.getId(), at: f.getDateCreated().getTime() }); }
      arr.sort(function (a, b) { return b.at - a.at; });
      list = arr.slice(0, 10).map(function (x) {
        return { name: x.name, id: x.id, at: Utilities.formatDate(new Date(x.at), 'GMT+7', 'yyyy-MM-dd HH:mm') };
      });
    }
  } catch (e) {}
  return {
    enabled: on && has,
    hour: Number(pr.getProperty('AUTO_BACKUP_HOUR') || 22),
    lastBackup: pr.getProperty('LAST_BACKUP') || '',
    files: list,
    folderId: folder ? folder.getId() : ''
  };
}

// กู้คืนจากไฟล์สำรองใน Drive
function restoreBackup(p) {
  var file = DriveApp.getFileById(p.id);
  var json = file.getBlob().getDataAsString();
  return importBackupData(json);
}

function importBackupData(json) {
  var data = (typeof json === 'string') ? JSON.parse(json) : json;
  Object.keys(SHEETS).forEach(function (name) {
    if (!data[name]) return;
    var s = sheet(name);
    s.clearContents();
    s.appendRow(SHEETS[name]);
    var rows = data[name].map(function (obj) {
      return SHEETS[name].map(function (h) { return typeof obj[h] === 'undefined' ? '' : obj[h]; });
    });
    if (rows.length) s.getRange(2, 1, rows.length, SHEETS[name].length).setValues(rows);
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

function clearAllData() {
  // ล้างทุกชีตยกเว้นหัวตาราง (ไม่แตะ Settings/GradeMapping ที่เป็นค่าตั้งค่า)
  ['Students', 'Subjects', 'LearningUnits', 'Scores', 'Grades', 'Abilities68'].forEach(function (name) {
    var s = sheet(name);
    var last = s.getLastRow();
    if (last > 1) s.getRange(2, 1, last - 1, s.getLastColumn()).clearContent();
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

// ล้างเฉพาะแถวที่ตรงเงื่อนไข (clearContent ปลอดภัยกว่า deleteRows)
function clearRowsWhere(name, field, value) {
  var s = sheet(name);
  var data = s.getDataRange().getValues();
  if (data.length < 2) return;
  var col = data[0].indexOf(field);
  if (col < 0) return;
  // เก็บแถวที่ไม่ตรงเงื่อนไขไว้ แล้วเขียนทับใหม่
  var keep = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][col]) !== String(value) && data[i].join('') !== '') keep.push(data[i]);
  }
  var cols = data[0].length;
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, cols).clearContent();
  if (keep.length) s.getRange(2, 1, keep.length, cols).setValues(keep);
}


/** ====================== UPLOAD IMAGE TO DRIVE ====================== */

function uploadImageToDrive(base64, filename) {
  if (!base64) throw new Error('ไม่มีข้อมูลรูปภาพ');
  // รองรับทั้งแบบมี prefix "data:image/...;base64," และไม่มี
  var matches = String(base64).match(/^data:(.+);base64,(.*)$/);
  var contentType = 'image/png';
  var b64 = base64;
  if (matches) { contentType = matches[1]; b64 = matches[2]; }

  var folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), contentType, filename || ('img_' + new Date().getTime()));
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var id = file.getId();
  // URL ที่ฝังใน <img> ได้
  return { url: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1000', fileId: id };
}

function getOrCreateFolder(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/** ====================== คำอธิบายรายวิชา (รูป/PDF) ====================== */
function getSubjectDocs(subjectID) {
  return readAll('SubjectDocs').filter(function (r) { return String(r.subjectID) === String(subjectID); });
}
// คืนไฟล์ Drive เป็น base64 เพื่อให้ frontend (pdf.js) เรนเดอร์ฝังในเล่ม (เลี่ยงปัญหา CORS)
function getDocBase64(fileId) {
  var f = DriveApp.getFileById(fileId);
  var blob = f.getBlob();
  return { base64: Utilities.base64Encode(blob.getBytes()), mimeType: blob.getContentType(), name: f.getName() };
}
function uploadSubjectDoc(p) {
  if (!p.base64) throw new Error('ไม่มีไฟล์');
  var m = String(p.base64).match(/^data:(.+);base64,(.*)$/);
  var ct = m ? m[1] : 'application/octet-stream';
  var b64 = m ? m[2] : p.base64;
  var folder = getOrCreateFolder('คำอธิบายรายวิชา - ' + (getSettings().schoolName || ''));
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), ct, p.fileName || ('doc_' + new Date().getTime()));
  var isPdf = (ct.indexOf('pdf') >= 0);

  if (isPdf) {
    // แปลง PDF → Google Slides → ส่ง PNG ทีละหน้า
    var pdfPages = convertPdfToImages_(blob, folder, p.fileName || 'doc');
    pdfPages.forEach(function (imgFile) {
      sheet('SubjectDocs').appendRow([
        genId('DOC'), p.subjectID, imgFile.getName(), imgFile.getId(), 'image',
        Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss')
      ]);
    });
  } else {
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    sheet('SubjectDocs').appendRow([genId('DOC'), p.subjectID, p.fileName || '', file.getId(), 'image',
      Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss')]);
  }
  SpreadsheetApp.flush();
  return getSubjectDocs(p.subjectID);
}

// แปลง PDF blob → PNG รายหน้า ผ่าน Google Slides API
// คืน array ของ DriveFile (image/png)
function convertPdfToImages_(pdfBlob, folder, baseName) {
  // บันทึก PDF ลง Drive ชั่วคราว แล้วคัดลอกแปลงเป็น Slides (ผ่าน REST — ไม่ต้องเปิด Advanced Service)
  var pdfFile = folder.createFile(pdfBlob.setName(baseName + '_src.pdf'));
  var slideFileId = '';
  try {
    slideFileId = driveCopyConvert_(pdfFile.getId(), baseName + '_tmp', 'application/vnd.google-apps.presentation');
  } catch (e) {
    try { pdfFile.setTrashed(true); } catch (e2) {}
    throw e;
  }

  Utilities.sleep(3000); // รอ conversion เสร็จ

  var pres = SlidesApp.openById(slideFileId);
  var slides = pres.getSlides();
  var imgs = [];
  slides.forEach(function (slide, i) {
    var url = 'https://docs.google.com/presentation/d/' + slideFileId +
      '/export/png?id=' + slideFileId + '&pageid=' + slide.getObjectId();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    var imgBlob = response.getBlob().setName(baseName + '_p' + (i + 1) + '.png');
    var imgFile = folder.createFile(imgBlob);
    imgFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    imgs.push(imgFile);
  });

  // ลบไฟล์ชั่วคราว (PDF ต้นฉบับ + Slides)
  try { DriveApp.getFileById(slideFileId).setTrashed(true); } catch (e) {}
  try { pdfFile.setTrashed(true); } catch (e) {}
  return imgs;
}
function deleteSubjectDoc(id) {
  var row = findRowById('SubjectDocs', id);
  if (row < 0) throw new Error('ไม่พบไฟล์');
  var fileId = sheet('SubjectDocs').getRange(row, 4).getValue();
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
  sheet('SubjectDocs').deleteRow(row);
  SpreadsheetApp.flush();
  return { id: id };
}


/** ====================== PART 5: ล็อกอิน / ครู ====================== */

function getTeachers() { return readAll('Teachers'); }

// ---- การจับคู่ชื่อครูแบบยืดหยุ่น (รองรับคำนำหน้า + ช่องว่างซ้อน) ----
function normName(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
function teacherCore(name) {
  var n = normName(name).replace(/^(นางสาว|นาง|นาย|ครู|เด็กชาย|เด็กหญิง|ด\.ช\.|ด\.ญ\.|ว่าที่ ?ร\.?ต\.?|ว่าที่)/, '').trim();
  return n.split(' ')[0];
}

function login(name, pin) {
  var key = normName(name);
  var t = getTeachers().filter(function (x) { return normName(x.name) === key; })[0];
  if (!t) throw new Error('ไม่พบชื่อครูในระบบ');
  if (String(t.pin).trim() !== String(pin).trim()) throw new Error('PIN ไม่ถูกต้อง');
  return { id: t.ID, name: t.name, role: t.role || 'teacher', homeroomClass: t.homeroomClass || '' };
}

// เปลี่ยนรหัสผ่าน (PIN) ของครูที่ล็อกอินอยู่ — ยืนยันด้วยรหัสเดิม
// p: { name, oldPin, newPin }
function changeMyPassword(p) {
  var key = normName(p.name);
  var rows = getTeachers();
  var t = null, rowIndex = -1;
  for (var i = 0; i < rows.length; i++) {
    if (normName(rows[i].name) === key) { t = rows[i]; rowIndex = rows[i]._rowIndex; break; }
  }
  if (!t) throw new Error('ไม่พบบัญชีผู้ใช้');
  if (String(t.pin).trim() !== String(p.oldPin).trim()) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  var newPin = String(p.newPin || '').trim();
  if (newPin.length < 4) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 4 ตัว');
  // คอลัมน์ pin = ลำดับที่ 3 ของชีต Teachers (ID,name,pin,role,homeroomClass)
  sheet('Teachers').getRange(rowIndex, 3).setValue(newPin);
  SpreadsheetApp.flush();
  return { ok: true };
}

function getSchedule(teacher) {
  var rows = readAll('Schedule');
  if (teacher) {
    var core = teacherCore(teacher);
    rows = rows.filter(function (r) { return teacherCore(r.teacher) === core || normName(r.teacher) === normName(teacher); });
  }
  return rows;
}

function saveTeacher(p) {
  var s = sheet('Teachers');
  if (p.id) {
    var row = findRowById('Teachers', p.id);
    if (row < 0) throw new Error('ไม่พบครู ID: ' + p.id);
    s.getRange(row, 2, 1, 4).setValues([[p.name || '', p.pin || '', p.role || 'teacher', p.homeroomClass || '']]);
  } else {
    s.appendRow([genId('TCH'), p.name || '', p.pin || '1234', p.role || 'teacher', p.homeroomClass || '']);
  }
  SpreadsheetApp.flush();
  return getTeachers();
}

// บันทึกครูทั้งหมดในครั้งเดียว (เขียนทับทั้งชีต โดยคงรหัส ID เดิม)
function saveTeachersBatch(teachers) {
  var s = sheet('Teachers');
  s.clearContents();
  s.appendRow(SHEETS.Teachers);
  var out = (teachers || []).map(function (t) {
    return [t.ID || genId('TCH'), t.name || '', t.pin || '1234', t.role || 'teacher', t.homeroomClass || ''];
  });
  if (out.length) s.getRange(2, 1, out.length, SHEETS.Teachers.length).setValues(out);
  SpreadsheetApp.flush();
  return getTeachers();
}

function deleteTeacher(id) {
  var row = findRowById('Teachers', id);
  if (row < 0) throw new Error('ไม่พบครู ID: ' + id);
  sheet('Teachers').deleteRow(row);
  SpreadsheetApp.flush();
  return { id: id };
}


/** ====================== PART 5: ตารางสอน ====================== */

// นำเข้าตารางสอน (แทนที่ทั้งหมด) + สร้างบัญชีครูอัตโนมัติจากชื่อที่พบ
// rows: [{day, period, classLevel, subject, teacher}, ...] (frontend กรอง ม.1-3 มาแล้ว)
function importSchedule(rows) {
  var sc = sheet('Schedule');
  sc.clearContents();
  sc.appendRow(SHEETS.Schedule);
  var out = [];
  (rows || []).forEach(function (r) {
    out.push([genId('SCH'), r.day, r.period, r.classLevel, r.subject, r.teacher]);
  });
  if (out.length) sc.getRange(2, 1, out.length, SHEETS.Schedule.length).setValues(out);

  // สร้างบัญชีครูอัตโนมัติ (role=teacher, PIN=1234) สำหรับชื่อที่ยังไม่มี (จับคู่แบบยืดหยุ่น)
  var existing = {};
  getTeachers().forEach(function (t) { existing[teacherCore(t.name)] = true; });
  var tcSheet = sheet('Teachers');
  var added = 0;
  (rows || []).forEach(function (r) {
    var nm = (r.teacher || '').trim();
    var core = teacherCore(nm);
    if (nm && core && !existing[core]) { tcSheet.appendRow([genId('TCH'), nm, '1234', 'teacher', '']); existing[core] = true; added++; }
  });
  SpreadsheetApp.flush();
  var subjectsAdded = createSubjectsFromScheduleRows(rows);
  return { count: out.length, teachersAdded: added, subjectsAdded: subjectsAdded };
}

// กิจกรรมที่ไม่ใช่รายวิชาที่ตัดเกรด (ไม่สร้างเป็นรายวิชา)
var NON_SUBJECT_RE = /(ชุมนุม|ชมรม|แนะแนว|ลูกเสือ|เนตรนารี|ยุวกาชาด|กิจกรรมพัฒนา)/;

// สร้างรายวิชาในระบบเกรดจากตารางสอน (1 วิชาต่อระดับชั้น เช่น "คณิตศาสตร์ ป.5") เฉพาะที่ยังไม่มี
function createSubjectsFromScheduleRows(rows) {
  var subjSheet = sheet('Subjects');
  var existing = {};
  getSubjects().forEach(function (su) { existing[su.name + '|' + su.classLevel] = true; });
  var added = 0, seen = {};
  (rows || []).forEach(function (r) {
    if (!r.subject || NON_SUBJECT_RE.test(r.subject)) return;
    var digit = String(r.classLevel).replace(/[^0-9]/g, '').charAt(0);
    if (!digit) return;
    var grade = 'ม.' + digit;
    var key = r.subject + '|' + grade;
    if (existing[key] || seen[key]) return;
    seen[key] = true;
    subjSheet.appendRow([genId('SUB'), '', r.subject, '', '', 'พื้นฐาน', grade, subjSheet.getLastRow()]);
    added++;
  });
  SpreadsheetApp.flush();
  return added;
}

// สร้างรายวิชาจากตารางสอนที่นำเข้าไว้แล้ว (ไม่ต้องนำเข้าไฟล์ใหม่)
function syncSubjectsFromSchedule() {
  var rows = readAll('Schedule').filter(function (s) { return /^ม\.[1-3]/.test(String(s.classLevel)); });
  return { added: createSubjectsFromScheduleRows(rows) };
}


/** ====================== PART 5: นำเข้ารายชื่อ (A-F) + เลื่อนเลขที่ ====================== */

// แทนที่รายชื่อทั้งหมดด้วยไฟล์นำเข้า แล้วเรียงเลขที่แต่ละห้องใหม่
// rows: [{classLevel, number, studentCode, prefix, firstName, lastName}, ...]
function bulkImportRoster(rows) {
  if (!rows || !rows.length) return { count: 0, added: 0, updated: 0 };
  ensureStudentCols_();
  var s = sheet('Students');
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var W = SHEETS.Students.length; // 9 คอลัมน์
  var nrm = function (x) { return String(x || '').replace(/\s+/g, '').trim(); };

  // อ่านข้อมูลเดิม (รวมหัวตาราง) แล้วทำดัชนีตามชื่อ
  var data = s.getDataRange().getValues();
  var body = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i].slice(0, W);
    while (r.length < W) r.push('');
    if (String(r[0] || '') === '' && String(r[3] || '') === '') continue; // ข้ามแถวว่าง
    body.push(r);
  }
  // คอลัมน์: 0=ID 1=studentCode 2=number 3=fullName 4=classLevel 5=dateAdded 6=prefix 7=firstName 8=lastName
  var byName = {}, byNameClass = {};
  body.forEach(function (r, idx) {
    var nk = nrm(r[7]) + '|' + nrm(r[8]);
    (byName[nk] = byName[nk] || []).push(idx);
    byNameClass[r[4] + '|' + nk] = idx;
  });

  var added = 0, updated = 0;
  rows.forEach(function (r) {
    var prefix = r.prefix || '', first = r.firstName || '', last = r.lastName || '';
    var cls = r.classLevel || '';
    var nk = nrm(first) + '|' + nrm(last);
    var idx = -1;
    if (byNameClass[cls + '|' + nk] != null) idx = byNameClass[cls + '|' + nk];      // ชื่อ+ชั้นตรง
    else if ((byName[nk] || []).length >= 1) idx = byName[nk][0];                     // ชื่อตรง (คงไว้)
    if (idx >= 0) {
      // ชื่อซ้ำ → คงข้อมูลเดิม อัปเดตเฉพาะเลขที่ + ชั้น (และเลขประจำตัวถ้ามีในไฟล์)
      body[idx][2] = (r.number != null && r.number !== '') ? r.number : body[idx][2];
      body[idx][4] = cls || body[idx][4];
      if (r.studentCode) body[idx][1] = r.studentCode;
      updated++;
    } else {
      // ชื่อใหม่ → เพิ่ม
      var full = buildFullName_(prefix, first, last);
      var nr = [genId('STD'), r.studentCode || '', r.number || '', full, cls, now, prefix, first, last];
      var ni = body.push(nr) - 1;
      (byName[nk] = byName[nk] || []).push(ni);
      byNameClass[cls + '|' + nk] = ni;
      added++;
    }
  });

  // เขียนกลับทั้งหมด (ไม่ renumber — ใช้เลขที่ตามไฟล์)
  s.clearContents();
  s.getRange(1, 1, 1, W).setValues([SHEETS.Students]);
  if (body.length) s.getRange(2, 1, body.length, W).setValues(body);
  SpreadsheetApp.flush();
  return { count: rows.length, added: added, updated: updated };
}


/** ====================== PART 5: เช็คชื่อ ====================== */

// period: 0 = หน้าเสาธง, 1-6 = รายคาบ
function getAttendance(date, classLevel, period) {
  return readAll('Attendance').filter(function (r) {
    return ymd(r.date) === ymd(date) &&
           String(r.classLevel) === String(classLevel) &&
           String(r.period) === String(period);
  });
}

// บันทึกการเช็คชื่อ (แทนที่ของ วันที่+ห้อง+คาบ นั้น)
// p: {date, classLevel, period, checkedBy, records:[{studentID, status}]}
function saveAttendance(p) {
  // ลบรายการเดิมของ วันที่+ห้อง+คาบ นี้
  var s = sheet('Attendance');
  var data = s.getDataRange().getValues();
  var keep = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i].join('') === '') continue;
    var same = ymd(data[i][1]) === ymd(p.date) &&
               String(data[i][2]) === String(p.classLevel) &&
               String(data[i][3]) === String(p.period);
    if (!same) { data[i][1] = ymd(data[i][1]); keep.push(data[i]); }
  }
  var now = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
  var sem = String(p.semester || getSettings().currentSemester || '1');
  (p.records || []).forEach(function (r) {
    keep.push([genId('ATT'), ymd(p.date), p.classLevel, p.period, r.studentID, r.status || 'มา', p.checkedBy || '', now, sem]);
  });

  // กฎ: ถ้าเช็คชื่อ "รายคาบ" (period ≥ 1) แล้วนักเรียน "มา" แต่หน้าเสาธง (period 0) วันนั้นบันทึกว่า "ขาด"
  //     → แก้หน้าเสาธงของนักเรียนคนนั้นเป็น "สาย" (มาสาย ไม่ใช่ขาด) ; ถ้ายังไม่มีบันทึกหน้าเสาธง ให้สร้างเป็น "สาย"
  if (String(p.period) !== '0') {
    var dymd = ymd(p.date), cls = String(p.classLevel);
    // index หน้าเสาธงที่ถูกเก็บไว้ (period 0) ของวัน+ห้องนี้ ตาม studentID
    var flagIdx = {};
    for (var k = 0; k < keep.length; k++) {
      if (String(keep[k][3]) === '0' && ymd(keep[k][1]) === dymd && String(keep[k][2]) === cls) {
        flagIdx[String(keep[k][4])] = k;
      }
    }
    (p.records || []).forEach(function (r) {
      if (r.status !== 'มา') return;             // เฉพาะคนที่ "มา" ในคาบนี้
      var sid = String(r.studentID);
      if (flagIdx[sid] != null) {
        var row = keep[flagIdx[sid]];
        if (row[5] === 'ขาด' || row[5] === '' || row[5] == null) { row[5] = 'สาย'; row[7] = now; }
      } else {
        // ยังไม่มีหน้าเสาธง → สร้างเป็น "สาย"
        keep.push([genId('ATT'), dymd, p.classLevel, 0, r.studentID, 'สาย', p.checkedBy || '', now, sem]);
        flagIdx[sid] = keep.length - 1;
      }
    });
  }
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, SHEETS.Attendance.length).clearContent();
  s.getRange(2, 2, Math.max(keep.length, 1), 1).setNumberFormat('@'); // คอลัมน์ date เป็นข้อความ
  if (keep.length) s.getRange(2, 1, keep.length, SHEETS.Attendance.length).setValues(keep);
  SpreadsheetApp.flush();
  return { count: (p.records || []).length };
}


/** ====================== PART 8: รายงานการมาเรียน ====================== */

var WD_TH = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];

// สรุปการมาเรียนรายคน (ช่วงวันที่) — รายวันจากหน้าเสาธง + เวลาเรียนรายคาบ (มีกฎ fallback ใช้ผลหน้าเสาธง)
function getAttendanceSummary(classLevel, fromDate, toDate) {
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });

  var att = readAll('Attendance').filter(function (r) {
    return String(r.classLevel) === String(classLevel) &&
           ymd(r.date) >= String(fromDate) && ymd(r.date) <= String(toDate);
  });

  // คาบที่แต่ละชั้นมีในแต่ละวัน (จากตารางสอน)
  var periodsByDay = {};
  readAll('Schedule').filter(function (s) { return String(s.classLevel) === String(classLevel); })
    .forEach(function (s) {
      var d = s.day; if (!periodsByDay[d]) periodsByDay[d] = {};
      periodsByDay[d][Number(s.period)] = true;
    });

  // วันที่ที่มีการเช็คหน้าเสาธง (ถือเป็นวันเรียนที่มีข้อมูล)
  var datesWithFlag = {};
  var look = {}; // look[date|period|studentID] = status
  att.forEach(function (r) {
    var d = ymd(r.date);
    look[d + '|' + r.period + '|' + r.studentID] = r.status;
    if (String(r.period) === '0') datesWithFlag[d] = true;
  });
  var dates = Object.keys(datesWithFlag).sort();

  var rows = students.map(function (st) {
    var ma = 0, khad = 0, la = 0, sai = 0, expP = 0, presP = 0;
    dates.forEach(function (date) {
      var flag = look[date + '|0|' + st.ID];
      if (flag === 'มา') ma++; else if (flag === 'ขาด') khad++; else if (flag === 'ลา') la++; else if (flag === 'สาย') sai++;
      var wd = WD_TH[new Date(date + 'T00:00:00').getDay()];
      var ps = periodsByDay[wd] ? Object.keys(periodsByDay[wd]) : [];
      ps.forEach(function (p) {
        expP++;
        var s = look[date + '|' + p + '|' + st.ID] || flag; // กฎ fallback: ไม่มีผลรายคาบ → ใช้หน้าเสาธง
        if (s === 'มา' || s === 'สาย') presP++;
      });
    });
    return {
      studentID: st.ID, number: st.number, fullName: st.fullName, classLevel: st.classLevel,
      ma: ma, khad: khad, la: la, sai: sai, schoolDays: dates.length,
      expectedPeriods: expP, presentPeriods: presP,
      percent: expP > 0 ? Math.round(presP / expP * 100) : 0
    };
  });
  return { classLevel: classLevel, from: fromDate, to: toDate, schoolDays: dates.length, students: rows };
}

// ภาพรวมการมาเรียนของทั้งโรงเรียนในวันหนึ่ง (จากหน้าเสาธง)
function getAttendanceDashboard(date) {
  var students = getStudents();
  var classCount = {};
  students.forEach(function (s) { if (s.classLevel) classCount[s.classLevel] = (classCount[s.classLevel] || 0) + 1; });

  var att = readAll('Attendance').filter(function (r) {
    return ymd(r.date) === ymd(date) && String(r.period) === '0';
  });
  var byClass = {};
  att.forEach(function (r) {
    var c = byClass[r.classLevel] = byClass[r.classLevel] || { checked: 0, present: 0, absent: 0, leave: 0, late: 0 };
    c.checked++;
    if (r.status === 'มา') c.present++; else if (r.status === 'ขาด') c.absent++;
    else if (r.status === 'ลา') c.leave++; else if (r.status === 'สาย') c.late++;
  });

  var rows = Object.keys(classCount).sort().map(function (c) {
    var s = byClass[c] || { checked: 0, present: 0, absent: 0, leave: 0, late: 0 };
    return { classLevel: c, total: classCount[c], taken: s.checked > 0,
      present: s.present, absent: s.absent, leave: s.leave, late: s.late };
  });
  var tot = { total: 0, present: 0, absent: 0, leave: 0, late: 0, classes: rows.length, classesTaken: 0 };
  rows.forEach(function (r) {
    tot.total += r.total; tot.present += r.present; tot.absent += r.absent; tot.leave += r.leave; tot.late += r.late;
    if (r.taken) tot.classesTaken++;
  });
  return { date: date, classes: rows, totals: tot };
}

/** ====================== WP2: บันทึกเวลาเรียน ====================== */
var TH_MONTH_ABBR = ['', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
function thaiMonthLabel(ym) {
  var p = String(ym).split('-'); var m = Number(p[1]); var by = (Number(p[0]) + 543) % 100;
  return (TH_MONTH_ABBR[m] || ym) + ' ' + (by < 10 ? '0' + by : by);
}

// นับเป็น "มาเรียน" = มา + สาย
function _attendedDay(s) { return s === 'มา' || s === 'สาย'; }

// สรุปเวลาเรียนรายเดือน → ทั้งปี (จากหน้าเสาธง period 0)
function getAttendanceRegister(classLevel, fromDate, toDate) {
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });
  var att = readAll('Attendance').filter(function (r) {
    return String(r.classLevel) === String(classLevel) && String(r.period) === '0' &&
           ymd(r.date) >= String(fromDate) && ymd(r.date) <= String(toDate);
  });
  var schoolDaysByYM = {}, look = {};
  att.forEach(function (r) {
    var d = ymd(r.date); var ym = d.slice(0, 7);
    (schoolDaysByYM[ym] = schoolDaysByYM[ym] || {})[d] = true;
    look[d + '|' + r.studentID] = r.status;
  });
  var yms = Object.keys(schoolDaysByYM).sort();
  var months = yms.map(function (ym) { return { ym: ym, label: thaiMonthLabel(ym), schoolDays: Object.keys(schoolDaysByYM[ym]).length }; });
  var totalSchool = months.reduce(function (a, m) { return a + m.schoolDays; }, 0);

  var rows = students.map(function (st) {
    var byMonth = {}, totalPresent = 0, ab = 0, la = 0, sa = 0;
    yms.forEach(function (ym) {
      var dates = Object.keys(schoolDaysByYM[ym]); var p = 0;
      dates.forEach(function (d) {
        var s = look[d + '|' + st.ID];
        if (_attendedDay(s)) p++;
        if (s === 'ขาด') ab++; else if (s === 'ลา') la++; else if (s === 'สาย') sa++;
      });
      byMonth[ym] = { present: p, total: dates.length };
      totalPresent += p;
    });
    return {
      studentID: st.ID, number: st.number, fullName: st.fullName, byMonth: byMonth,
      totalPresent: totalPresent, absent: ab, leave: la, late: sa,
      percent: totalSchool > 0 ? Math.round(totalPresent / totalSchool * 100) : 0
    };
  });
  return { classLevel: classLevel, months: months, totalSchoolDays: totalSchool, students: rows };
}

// กริดเช็คชื่อรายวันของ 1 เดือน (ym = 'YYYY-MM')
function getMonthlyGrid(classLevel, ym) {
  var students = getStudents().filter(function (s) { return String(s.classLevel) === String(classLevel); });
  students.sort(function (a, b) { return Number(a.number) - Number(b.number); });
  var att = readAll('Attendance').filter(function (r) {
    return String(r.classLevel) === String(classLevel) && String(r.period) === '0' && ymd(r.date).slice(0, 7) === String(ym);
  });
  var datesSet = {}, look = {};
  att.forEach(function (r) { var d = ymd(r.date); datesSet[d] = true; look[d + '|' + r.studentID] = r.status; });
  var dates = Object.keys(datesSet).sort();
  var rows = students.map(function (st) {
    var byDate = {}, p = 0, ab = 0, la = 0, sa = 0;
    dates.forEach(function (d) {
      var s = look[d + '|' + st.ID] || '';
      byDate[d] = s;
      if (_attendedDay(s)) p++;
      if (s === 'ขาด') ab++; else if (s === 'ลา') la++; else if (s === 'สาย') sa++;
    });
    return { studentID: st.ID, number: st.number, fullName: st.fullName, byDate: byDate, present: p, absent: ab, leave: la, late: sa };
  });
  return { classLevel: classLevel, ym: ym, label: thaiMonthLabel(ym), dates: dates, students: rows };
}



function INIT() {
  setupSheets();
  Logger.log('สร้างชีตเรียบร้อย: ' + Object.keys(SHEETS).join(', '));
}
