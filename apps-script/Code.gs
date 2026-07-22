const SPREADSHEET_ID = '1955ghvAH8XdYs2Q356-YQhFdt4_0LoS0jTbnfaXswF8';
const ANSWERS_SHEET = 'คำตอบนักเรียน';
const QUESTIONS_SHEET = 'คำถาม RIASEC';
const QUESTION_START_COLUMN = 8;
const QUESTION_COUNT = 216;
const SCORE_START_COLUMN = QUESTION_START_COLUMN + QUESTION_COUNT;
const ANSWERS_HEADER_ROW = 1;
const ANSWERS_FIRST_DATA_ROW = 3;
const STUDENT_META_START_COLUMN = 232;
const OCCUPATION_START_COLUMN = 235;
const OCCUPATION_HEADERS = ['อาชีพที่ 1 (ล่าสุด)', 'อาชีพที่ 2', 'อาชีพที่ 3', 'อาชีพที่ 4', 'อาชีพที่ 5 (เก่าที่สุด)'];
const MATCH_START_COLUMN = 240;
const MATCH_HEADERS = ['อาชีพที่ 1 สอดคล้องกับความสามารถ', 'อาชีพที่ 2 สอดคล้องกับความสามารถ', 'อาชีพที่ 3 สอดคล้องกับความสามารถ', 'อาชีพที่ 4 สอดคล้องกับความสามารถ', 'อาชีพที่ 5 สอดคล้องกับความสามารถ'];
const RIASEC_CODES = ['R', 'I', 'A', 'S', 'E', 'C'];
const TIE_BREAK_SECTIONS = ['occupations', 'competencies', 'activities', 'selfEstimates'];
const SUMMARY_HEADERS = ['R', 'I', 'A', 'S', 'E', 'C', 'Holland Code', 'บุคลิกภาพเด่น'];
const PERSONALITY_BY_CODE = {
  R: 'นักปฏิบัติ', I: 'นักคิดวิเคราะห์', A: 'นักสร้างสรรค์',
  S: 'ผู้เข้าใจผู้คน', E: 'ผู้นำการเปลี่ยนแปลง', C: 'ผู้จัดระบบ'
};

function doGet(e) {
  if (e.parameter.action !== 'questions') return json_({ ok: true, service: 'SDS RIASEC' });
  const rows = questionRows_();
  const questions = rows.map(function (r) { return { id: r[0], type: r[1], code: r[2], text: r[3], positive: r[4], negative: r[5] }; });
  return json_({ ok: true, questions: questions });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(30000);
    locked = true;
    const data = JSON.parse((e.parameter && e.parameter.payload) || e.postData.contents);
    const student = data.student || {};
    if (!student.firstName || !student.lastName || !student.gradeLevel || !student.room || !student.studentNumber || !student.consent) throw new Error('ข้อมูลนักเรียนไม่ครบ');
    if (!Array.isArray(student.occupations) || student.occupations.length !== 5 || student.occupations.some(function (value) { return !String(value || '').trim(); })) throw new Error('กรุณากรอกอาชีพให้ครบ 5 อันดับ');
    if (!Array.isArray(data.answers) || data.answers.length !== QUESTION_COUNT) throw new Error('คำตอบต้องครบ 216 ข้อ');
    const answers = data.answers.map(function (v) { return Math.max(0, Number(v) || 0); });
    const workbook = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = workbook.getSheetByName(ANSWERS_SHEET);
    const questionRows = questionRows_();
    const result = calculateRiasecResult_(questionRows, answers);
    // ใช้แถวเดิมเมื่อเป็นนักเรียนคนเดิม เพื่อไม่ให้ข้อมูลซ้ำและให้ผลล่าสุดแทนผลเก่า
    const row = findExistingStudentRow_(sheet, student) || nextStudentRow_(sheet);
    sheet.getRange(row, 1, 1, 7).setValues([[Utilities.getUuid(), new Date(), student.firstName + ' ' + student.lastName, student.nickName || '', student.gradeLevel, student.room, student.studentNumber]]);
    sheet.getRange(row, QUESTION_START_COLUMN, 1, QUESTION_COUNT).setValues([answers]);
    ensureScoringHeaders_(sheet);
    sheet.getRange(row, SCORE_START_COLUMN, 1, SUMMARY_HEADERS.length).setValues([[
      result.scores.R, result.scores.I, result.scores.A,
      result.scores.S, result.scores.E, result.scores.C,
      result.code, PERSONALITY_BY_CODE[result.code.charAt(0)]
    ]]);
    sheet.getRange(row, STUDENT_META_START_COLUMN, 1, 3).setValues([[student.firstName, student.lastName, 'ยินยอม']]);
    ensureOccupationHeaders_(sheet);
    sheet.getRange(row, OCCUPATION_START_COLUMN, 1, 5).setValues([student.occupations.map(function (value) { return String(value).trim(); })]);
    ensureMatchHeaders_(sheet);
    sheet.getRange(row, MATCH_START_COLUMN, 1, 5).setValues([occupationMatchValues_(data.occupationMatches)]);
    SpreadsheetApp.flush();
    return json_({ ok: true, row: row });
  } catch (error) { return json_({ ok: false, error: error.message }); }
  finally { if (locked) lock.releaseLock(); }
}

function ensureScoringHeaders_(sheet) {
  const requiredColumns = SCORE_START_COLUMN + SUMMARY_HEADERS.length - 1;
  if (sheet.getMaxColumns() < requiredColumns) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  const current = sheet.getRange(ANSWERS_HEADER_ROW, SCORE_START_COLUMN, 1, SUMMARY_HEADERS.length).getDisplayValues()[0];
  if (current.some(function (value) { return value === ''; })) {
    sheet.getRange(ANSWERS_HEADER_ROW, SCORE_START_COLUMN, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  }
}

function questionRows_() {
  const values = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(QUESTIONS_SHEET).getDataRange().getValues();
  const rows = values.filter(function (row) {
    const code = String(row[2] || '').trim().toUpperCase();
    return RIASEC_CODES.indexOf(code) !== -1 && normalizeQuestionSection_(row[1]) !== '';
  }).slice(0, QUESTION_COUNT);
  if (rows.length !== QUESTION_COUNT) throw new Error('ตารางคำถามต้องมีคำถาม RIASEC ครบ 216 ข้อ');
  return rows;
}

function calculateRiasecResult_(questionRows, answers) {
  const scores = {};
  const sectionScores = {};
  RIASEC_CODES.forEach(function (code) {
    scores[code] = 0;
  });
  TIE_BREAK_SECTIONS.forEach(function (section) {
    sectionScores[section] = {};
    RIASEC_CODES.forEach(function (code) { sectionScores[section][code] = 0; });
  });

  questionRows.forEach(function (question, index) {
    const code = String(question[2] || '').trim().toUpperCase();
    const section = normalizeQuestionSection_(question[1]);
    if (RIASEC_CODES.indexOf(code) === -1) return;
    const value = Math.max(0, Number(answers[index]) || 0);
    scores[code] += value;
    if (section) sectionScores[section][code] += value;
  });

  const ranking = RIASEC_CODES.map(function (code) {
    return { code: code, score: scores[code] };
  });
  ranking.sort(function (a, b) {
    if (b.score !== a.score) return b.score - a.score;
    for (let i = 0; i < TIE_BREAK_SECTIONS.length; i++) {
      const section = TIE_BREAK_SECTIONS[i];
      const difference = sectionScores[section][b.code] - sectionScores[section][a.code];
      if (difference !== 0) return difference;
    }
    return RIASEC_CODES.indexOf(a.code) - RIASEC_CODES.indexOf(b.code);
  });

  return {
    scores: scores,
    code: ranking.slice(0, 3).map(function (entry) { return entry.code; }).join('')
  };
}

function normalizeQuestionSection_(type) {
  const value = String(type || '').toLocaleLowerCase();
  if (/อาชีพ|occupation/.test(value)) return 'occupations';
  if (/ความสามารถ|competenc|ability|skill/.test(value)) return 'competencies';
  if (/ประเมินตนเอง|self.?estimate/.test(value)) return 'selfEstimates';
  if (/กิจกรรม|activit/.test(value)) return 'activities';
  return '';
}

function ensureOccupationHeaders_(sheet) {
  const requiredColumns = OCCUPATION_START_COLUMN + OCCUPATION_HEADERS.length - 1;
  if (sheet.getMaxColumns() < requiredColumns) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  const headers = sheet.getRange(ANSWERS_HEADER_ROW, OCCUPATION_START_COLUMN, 1, OCCUPATION_HEADERS.length).getDisplayValues()[0];
  if (headers.some(function (value) { return value === ''; })) sheet.getRange(ANSWERS_HEADER_ROW, OCCUPATION_START_COLUMN, 1, OCCUPATION_HEADERS.length).setValues([OCCUPATION_HEADERS]);
}

function ensureMatchHeaders_(sheet) {
  const requiredColumns = MATCH_START_COLUMN + MATCH_HEADERS.length - 1;
  if (sheet.getMaxColumns() < requiredColumns) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  const headers = sheet.getRange(ANSWERS_HEADER_ROW, MATCH_START_COLUMN, 1, MATCH_HEADERS.length).getDisplayValues()[0];
  if (headers.some(function (value, index) { return value !== MATCH_HEADERS[index]; })) sheet.getRange(ANSWERS_HEADER_ROW, MATCH_START_COLUMN, 1, MATCH_HEADERS.length).setValues([MATCH_HEADERS]);
}

function occupationMatchValues_(groups) {
  return MATCH_HEADERS.map(function (_, index) {
    const group = Array.isArray(groups) ? groups[index] : null;
    if (!group) return 'ยังไม่มีผลประเมิน';
    const label = String(group.label || 'ยังประเมินไม่ได้').trim();
    const detail = String(group.detail || '').trim();
    return detail ? label + ': ' + detail : label;
  });
}

function findExistingStudentRow_(sheet, student) {
  const startRow = ANSWERS_FIRST_DATA_ROW;
  const rowCount = sheet.getMaxRows() - startRow + 1;
  if (rowCount <= 0) return 0;

  const values = sheet.getRange(startRow, 3, rowCount, 5).getDisplayValues();
  const name = normalizeMatchValue_(student.firstName + ' ' + student.lastName);
  const grade = normalizeGradeValue_(student.gradeLevel);
  const room = normalizeMatchValue_(student.room);
  const studentNumber = normalizeMatchValue_(student.studentNumber);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (
      normalizeMatchValue_(row[0]) === name &&
      normalizeGradeValue_(row[2]) === grade &&
      normalizeMatchValue_(row[3]) === room &&
      normalizeMatchValue_(row[4]) === studentNumber
    ) {
      return startRow + i;
    }
  }
  return 0;
}

function normalizeMatchValue_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLocaleLowerCase();
}

function normalizeGradeValue_(value) {
  const text = normalizeMatchValue_(value);
  const match = text.match(/([1-6])$/);
  return match ? match[1] : text;
}

function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }

// ไม่ใช้ getLastRow() เพราะสูตร ARRAYFORMULA ในคอลัมน์คะแนนขยายลงมาทั้งชีต
function nextStudentRow_(sheet) {
  const startRow = ANSWERS_FIRST_DATA_ROW;
  const values = sheet.getRange(startRow, 3, sheet.getMaxRows() - startRow + 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '') return i + startRow;
  }
  return sheet.getMaxRows() + 1;
}
