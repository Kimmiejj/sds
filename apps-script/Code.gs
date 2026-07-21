const SPREADSHEET_ID = '1955ghvAH8XdYs2Q356-YQhFdt4_0LoS0jTbnfaXswF8';
const ANSWERS_SHEET = 'คำตอบนักเรียน';
const QUESTIONS_SHEET = 'คำถาม RIASEC';
const QUESTION_START_COLUMN = 8;
const QUESTION_COUNT = 216;
const SCORE_START_COLUMN = QUESTION_START_COLUMN + QUESTION_COUNT;
const STUDENT_META_START_COLUMN = 232;
const OCCUPATION_START_COLUMN = 235;
const OCCUPATION_HEADERS = ['อาชีพที่ 1 (ล่าสุด)', 'อาชีพที่ 2', 'อาชีพที่ 3', 'อาชีพที่ 4', 'อาชีพที่ 5 (เก่าที่สุด)'];
const MATCH_START_COLUMN = 240;
const MATCH_HEADERS = ['อาชีพที่ 1 พบอาชีพในระบบ', 'อาชีพที่ 2 พบอาชีพในระบบ', 'อาชีพที่ 3 พบอาชีพในระบบ', 'อาชีพที่ 4 พบอาชีพในระบบ', 'อาชีพที่ 5 พบอาชีพในระบบ'];
const RIASEC_CODES = ['R', 'I', 'A', 'S', 'E', 'C'];
const TIE_BREAK_SECTIONS = ['occupations', 'competencies', 'activities', 'selfEstimates'];
const SECTION_SCORE_START_COLUMN = 245;
const RANK_KEY_START_COLUMN = SECTION_SCORE_START_COLUMN + RIASEC_CODES.length * TIE_BREAK_SECTIONS.length;
const SUMMARY_HEADERS = ['R', 'I', 'A', 'S', 'E', 'C', 'Holland Code', 'บุคลิกภาพเด่น'];
const SECTION_SCORE_HEADERS = TIE_BREAK_SECTIONS.flatMap(function (section) { return RIASEC_CODES.map(function (code) { return section + ' ' + code; }); });
const RANK_KEY_HEADERS = RIASEC_CODES.map(function (code) { return 'Tie key ' + code; });

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
    // ใช้แถวเดิมเมื่อเป็นนักเรียนคนเดิม เพื่อไม่ให้ข้อมูลซ้ำและให้ผลล่าสุดแทนผลเก่า
    const row = findExistingStudentRow_(sheet, student) || nextStudentRow_(sheet);
    sheet.getRange(row, 1, 1, 7).setValues([[Utilities.getUuid(), new Date(), student.firstName + ' ' + student.lastName, student.nickName || '', student.gradeLevel, student.room, student.studentNumber]]);
    sheet.getRange(row, QUESTION_START_COLUMN, 1, QUESTION_COUNT).setValues([answers]);
    ensureScoringHeaders_(sheet);
    writeScoringFormulas_(sheet, row, questionRows);
    sheet.getRange(row, STUDENT_META_START_COLUMN, 1, 3).setValues([[student.firstName, student.lastName, 'ยินยอม']]);
    ensureOccupationHeaders_(sheet);
    sheet.getRange(row, OCCUPATION_START_COLUMN, 1, 5).setValues([student.occupations.map(function (value) { return String(value).trim(); })]);
    ensureMatchHeaders_(sheet);
    sheet.getRange(row, MATCH_START_COLUMN, 1, 5).setValues([occupationMatchValues_(data.occupationMatches)]);
    sheet.getRange(row, RANK_KEY_START_COLUMN, 1, RIASEC_CODES.length).setNote('จัดอันดับจากคะแนนรวมก่อน หากเท่ากันใช้ อาชีพ > ความสามารถ > กิจกรรม > ประเมินตนเอง; หากยังเท่ากันใช้ลำดับ R-I-A-S-E-C เป็นลำดับคงที่');
    SpreadsheetApp.flush();
    return json_({ ok: true, row: row });
  } catch (error) { return json_({ ok: false, error: error.message }); }
  finally { if (locked) lock.releaseLock(); }
}

function ensureScoringHeaders_(sheet) {
  const requiredColumns = RANK_KEY_START_COLUMN + RANK_KEY_HEADERS.length - 1;
  if (sheet.getMaxColumns() < requiredColumns) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  const blocks = [
    [SCORE_START_COLUMN, SUMMARY_HEADERS],
    [SECTION_SCORE_START_COLUMN, SECTION_SCORE_HEADERS],
    [RANK_KEY_START_COLUMN, RANK_KEY_HEADERS]
  ];
  blocks.forEach(function (block) {
    const startColumn = block[0];
    const headers = block[1];
    const headerRow = headerRow_(sheet);
    const current = sheet.getRange(headerRow, startColumn, 1, headers.length).getDisplayValues()[0];
    if (current.some(function (value) { return value === ''; })) sheet.getRange(headerRow, startColumn, 1, headers.length).setValues([headers]);
  });
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

function writeScoringFormulas_(sheet, row, questionRows) {
  const refs = {};
  RIASEC_CODES.forEach(function (code) {
    refs[code] = { total: [], activities: [], competencies: [], occupations: [], selfEstimates: [] };
  });
  questionRows.forEach(function (question, index) {
    const code = String(question[2] || '').trim().toUpperCase();
    const section = normalizeQuestionSection_(question[1]);
    if (!refs[code]) return;
    const cell = columnLetter_(QUESTION_START_COLUMN + index) + row;
    refs[code].total.push(cell);
    if (section) refs[code][section].push(cell);
  });

  const blankGuard = 'IF($C' + row + '="","",';
  const totalFormulas = RIASEC_CODES.map(function (code) { return '=' + blankGuard + sumFormula_(refs[code].total) + ')'; });
  sheet.getRange(row, SCORE_START_COLUMN, 1, RIASEC_CODES.length).setFormulas([totalFormulas]);

  const sectionFormulas = [];
  TIE_BREAK_SECTIONS.forEach(function (section) {
    RIASEC_CODES.forEach(function (code) { sectionFormulas.push('=' + blankGuard + sumFormula_(refs[code][section]) + ')'); });
  });
  sheet.getRange(row, SECTION_SCORE_START_COLUMN, 1, SECTION_SCORE_HEADERS.length).setFormulas([sectionFormulas]);

  const scoreLetters = RIASEC_CODES.map(function (_, index) { return columnLetter_(SCORE_START_COLUMN + index); });
  const sectionLetters = {};
  TIE_BREAK_SECTIONS.forEach(function (section, sectionIndex) {
    sectionLetters[section] = RIASEC_CODES.map(function (_, codeIndex) { return columnLetter_(SECTION_SCORE_START_COLUMN + sectionIndex * RIASEC_CODES.length + codeIndex); });
  });
  const rankKeys = RIASEC_CODES.map(function (_, codeIndex) {
    const epsilon = (RIASEC_CODES.length - codeIndex) + '/1000';
    return '=' + blankGuard + scoreLetters[codeIndex] + row + '*1000000000+' + sectionLetters.occupations[codeIndex] + row + '*1000000+' + sectionLetters.competencies[codeIndex] + row + '*10000+' + sectionLetters.activities[codeIndex] + row + '*100+' + sectionLetters.selfEstimates[codeIndex] + row + '+' + epsilon + ')';
  });
  sheet.getRange(row, RANK_KEY_START_COLUMN, 1, RANK_KEY_HEADERS.length).setFormulas([rankKeys]);

  const keyRange = columnLetter_(RANK_KEY_START_COLUMN) + row + ':' + columnLetter_(RANK_KEY_START_COLUMN + RIASEC_CODES.length - 1) + row;
  const codeParts = RIASEC_CODES.map(function (code, index) {
    const keyCell = columnLetter_(RANK_KEY_START_COLUMN + index) + row;
    return 'IF(' + keyCell + '=LARGE(' + keyRange + ',';
  });
  const codeFormula = '=' + blankGuard + codeParts.map(function (part, index) { return part + (index + 1) + ',"' + RIASEC_CODES[index] + '","")'; }).join('&') + ')';
  sheet.getRange(row, SCORE_START_COLUMN + RIASEC_CODES.length).setFormula(codeFormula);
  const hollandCell = columnLetter_(SCORE_START_COLUMN + RIASEC_CODES.length) + row;
  const personality = ['นักปฏิบัติ', 'นักคิดวิเคราะห์', 'นักสร้างสรรค์', 'ผู้เข้าใจผู้คน', 'ผู้นำการเปลี่ยนแปลง', 'ผู้จัดระบบ'];
  const personalityFormula = '=' + blankGuard + 'IF(LEFT(' + hollandCell + ',1)="R","' + personality[0] + '",IF(LEFT(' + hollandCell + ',1)="I","' + personality[1] + '",IF(LEFT(' + hollandCell + ',1)="A","' + personality[2] + '",IF(LEFT(' + hollandCell + ',1)="S","' + personality[3] + '",IF(LEFT(' + hollandCell + ',1)="E","' + personality[4] + '","' + personality[5] + '")))))' + ')';
  sheet.getRange(row, SCORE_START_COLUMN + RIASEC_CODES.length + 1).setFormula(personalityFormula);
}

function normalizeQuestionSection_(type) {
  const value = String(type || '').toLocaleLowerCase();
  if (/อาชีพ|occupation/.test(value)) return 'occupations';
  if (/ความสามารถ|competenc|ability|skill/.test(value)) return 'competencies';
  if (/ประเมินตนเอง|self.?estimate/.test(value)) return 'selfEstimates';
  if (/กิจกรรม|activit/.test(value)) return 'activities';
  return '';
}

function sumFormula_(cells) { return cells.length ? 'SUM(' + cells.join(',') + ')' : '0'; }

function columnLetter_(column) {
  let value = '';
  while (column > 0) { const remainder = (column - 1) % 26; value = String.fromCharCode(65 + remainder) + value; column = Math.floor((column - 1) / 26); }
  return value;
}

function ensureOccupationHeaders_(sheet) {
  const requiredColumns = OCCUPATION_START_COLUMN + OCCUPATION_HEADERS.length - 1;
  if (sheet.getMaxColumns() < requiredColumns) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  const headerRow = headerRow_(sheet);
  const headers = sheet.getRange(headerRow, OCCUPATION_START_COLUMN, 1, OCCUPATION_HEADERS.length).getDisplayValues()[0];
  if (headers.some(function (value) { return value === ''; })) sheet.getRange(headerRow, OCCUPATION_START_COLUMN, 1, OCCUPATION_HEADERS.length).setValues([OCCUPATION_HEADERS]);
}

function ensureMatchHeaders_(sheet) {
  const requiredColumns = MATCH_START_COLUMN + MATCH_HEADERS.length - 1;
  if (sheet.getMaxColumns() < requiredColumns) sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  const headerRow = headerRow_(sheet);
  const headers = sheet.getRange(headerRow, MATCH_START_COLUMN, 1, MATCH_HEADERS.length).getDisplayValues()[0];
  if (headers.some(function (value) { return value === ''; })) sheet.getRange(headerRow, MATCH_START_COLUMN, 1, MATCH_HEADERS.length).setValues([MATCH_HEADERS]);
}

function occupationMatchValues_(groups) {
  return MATCH_HEADERS.map(function (_, index) {
    const group = Array.isArray(groups) ? groups[index] : null;
    const matches = group && Array.isArray(group.matches) ? group.matches : [];
    if (!matches.length) return 'ไม่พบอาชีพที่ตรงหรือใกล้เคียง';
    return matches.map(function (match) {
      const name = String((match && match.name) || '').trim();
      const type = String((match && match.type) || 'ใกล้เคียงกับ').trim();
      return name ? type + ': ' + name : '';
    }).filter(Boolean).join(' | ');
  });
}

function findExistingStudentRow_(sheet, student) {
  const startRow = headerRow_(sheet) + 1;
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
  const startRow = headerRow_(sheet) + 1;
  const values = sheet.getRange(startRow, 3, sheet.getMaxRows() - startRow + 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === '') return i + startRow;
  }
  return sheet.getMaxRows() + 1;
}

function headerRow_(sheet) {
  return sheet.getRange(4, 1).getDisplayValue() === 'รหัสรายการ' ? 4 : 2;
}
