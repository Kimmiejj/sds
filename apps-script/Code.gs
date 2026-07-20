const SPREADSHEET_ID = '1955ghvAH8XdYs2Q356-YQhFdt4_0LoS0jTbnfaXswF8';
const ANSWERS_SHEET = 'คำตอบนักเรียน';
const QUESTIONS_SHEET = 'คำถาม RIASEC';
const OCCUPATION_START_COLUMN = 235;
const OCCUPATION_HEADERS = ['อาชีพที่ 1 (ล่าสุด)', 'อาชีพที่ 2', 'อาชีพที่ 3', 'อาชีพที่ 4', 'อาชีพที่ 5 (เก่าที่สุด)'];

function doGet(e) {
  if (e.parameter.action !== 'questions') return json_({ ok: true, service: 'SDS RIASEC' });
  const rows = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(QUESTIONS_SHEET).getRange('A2:F217').getValues();
  const questions = rows.map(function (r) { return { id: r[0], type: r[1], code: r[2], text: r[3], positive: r[4], negative: r[5] }; });
  return json_({ ok: true, questions: questions });
}

function doPost(e) {
  try {
    const data = JSON.parse((e.parameter && e.parameter.payload) || e.postData.contents);
    const student = data.student || {};
    if (!student.firstName || !student.lastName || !student.gradeLevel || !student.room || !student.studentNumber || !student.consent) throw new Error('ข้อมูลนักเรียนไม่ครบ');
    if (!Array.isArray(student.occupations) || student.occupations.length !== 5 || student.occupations.some(function (value) { return !String(value || '').trim(); })) throw new Error('กรุณากรอกอาชีพให้ครบ 5 อันดับ');
    if (!Array.isArray(data.answers) || data.answers.length !== 216) throw new Error('คำตอบต้องครบ 216 ข้อ');
    const answers = data.answers.map(function (v) { return Number(v) === 1 ? 1 : 0; });
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ANSWERS_SHEET);
    // ใช้แถวเดิมเมื่อเป็นนักเรียนคนเดิม เพื่อไม่ให้ข้อมูลซ้ำและให้ผลล่าสุดแทนผลเก่า
    const row = findExistingStudentRow_(sheet, student) || nextStudentRow_(sheet);
    sheet.getRange(row, 1, 1, 7).setValues([[Utilities.getUuid(), new Date(), student.firstName + ' ' + student.lastName, student.nickName || '', student.gradeLevel, student.room, student.studentNumber]]);
    sheet.getRange(row, 8, 1, 216).setValues([answers]);
    sheet.getRange(row, 232, 1, 3).setValues([[student.firstName, student.lastName, 'ยินยอม']]);
    ensureOccupationHeaders_(sheet);
    sheet.getRange(row, OCCUPATION_START_COLUMN, 1, 5).setValues([student.occupations.map(function (value) { return String(value).trim(); })]);
    SpreadsheetApp.flush();
    return json_({ ok: true, row: row });
  } catch (error) { return json_({ ok: false, error: error.message }); }
}

function ensureOccupationHeaders_(sheet) {
  const requiredColumns = OCCUPATION_START_COLUMN + OCCUPATION_HEADERS.length - 1;
  if (sheet.getMaxColumns() < requiredColumns) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  const headers = sheet.getRange(2, OCCUPATION_START_COLUMN, 1, OCCUPATION_HEADERS.length).getDisplayValues()[0];
  if (headers.some(function (value) { return value === ''; })) sheet.getRange(2, OCCUPATION_START_COLUMN, 1, OCCUPATION_HEADERS.length).setValues([OCCUPATION_HEADERS]);
}

function findExistingStudentRow_(sheet, student) {
  const startRow = 3;
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
  const values = sheet.getRange(3, 3, sheet.getMaxRows() - 2, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '') return i + 3;
  }
  return sheet.getMaxRows() + 1;
}
