import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const outputDir = 'outputs/riasec-google-sheets';
await fs.mkdir(outputDir, { recursive: true });
const wb = Workbook.create();
const guide = wb.worksheets.add('เริ่มต้น');
const responses = wb.worksheets.add('คำตอบนักเรียน');
const questions = wb.worksheets.add('คำถาม RIASEC');
const careers = wb.worksheets.add('อาชีพและคณะ');
const universities = wb.worksheets.add('มหาวิทยาลัย 5 ตัวเลือก');
const resources = wb.worksheets.add('ลิงก์ไทย');
const dashboard = wb.worksheets.add('สรุปผลห้อง');

const navy = '#193D3A', orange = '#F06438', mint = '#E8F3EE', cream = '#FAF7F1', line = '#D8E0DC';
function title(sheet, range, text, subtitle) {
  sheet.mergeCells(range); sheet.getRange(range.split(':')[0]).values = [[text]];
  sheet.getRange(range).format = { fill: navy, font: { bold: true, color: '#FFFFFF', size: 18 }, horizontalAlignment: 'left', verticalAlignment: 'center' };
  const start = range.split(':')[0].replace(/\d+/, ''); const row = Number(range.match(/\d+/)[0]) + 1;
  sheet.getRange(`${start}${row}`).values = [[subtitle]];
  sheet.getRange(`${start}${row}`).format = { font: { color: '#60756F', italic: true, size: 10 } };
}
function header(sheet, range) { sheet.getRange(range).format = { fill: '#DCEBE5', font: { bold: true, color: navy }, horizontalAlignment: 'center', verticalAlignment: 'center', wrapText: true, borders: { preset: 'outside', style: 'thin', color: line } }; }
function rowData(start, cols, rows){ sheet.getRangeByIndexes(start,0,rows.length,cols).values=rows; }

title(guide, 'A1:H1', 'เข็มทิศอาชีพ · RIASEC Thailand', 'เวิร์กบุ๊กสำหรับ Google Sheets + Glide: รับคำตอบนักเรียน คำนวณผล และสรุปภาพรวมของห้อง');
const guideRows = [
 ['1. อัปโหลดไฟล์นี้เข้า Google Drive','เลือก New > File upload แล้วเปิดด้วย Google Sheets เพื่อใช้เป็นฐานข้อมูลออนไลน์'],
 ['2. สร้างแอปใน Glide','เลือกฐานข้อมูล Google Sheets นี้ แล้วใช้ตาราง “คำตอบนักเรียน” เป็นตารางหลัก'],
 ['3. สร้างแบบฟอร์ม','รับข้อมูลชื่อ ห้อง เลขที่ และ Q01–Q36 (กำหนดค่า 1 = ใช่/สนใจ, 0 = ไม่ใช่)'],
 ['4. ตั้งค่าให้ผู้เรียนส่งคำตอบ','ให้ Glide เพิ่มหนึ่งแถวใน “คำตอบนักเรียน”; คอลัมน์คะแนนและ Holland Code จะคำนวณอัตโนมัติ'],
 ['5. ทำ QR Code','คัดลอก Public app link จาก Glide แล้วสร้าง QR ใน Glide หรือเครื่องมือ QR ที่โรงเรียนใช้อยู่'],
 ['ความเป็นส่วนตัว','แนะนำเปิดเฉพาะผู้มีลิงก์/อีเมลโรงเรียน และเก็บข้อมูลเท่าที่จำเป็นตามนโยบายโรงเรียน']
];
guide.getRange('A4:B9').values=guideRows; guide.getRange('A4:A9').format={fill:mint,font:{bold:true,color:navy},wrapText:true};guide.getRange('B4:B9').format={wrapText:true};guide.getRange('A4:B9').format.borders={preset:'all',style:'thin',color:line};guide.getRange('A4:B9').format.rowHeight=44;guide.getRange('A:A').format.columnWidth=30;guide.getRange('B:B').format.columnWidth=75;guide.showGridLines=false;

const baseQuestions = [
 ['กิจกรรม','ประกอบหรือซ่อมสิ่งของที่มีชิ้นส่วน'],['กิจกรรม','ทดลอง ค้นคว้า หรือหาคำตอบจากข้อมูล'],['กิจกรรม','วาดภาพ ออกแบบ หรือสร้างผลงาน'],['กิจกรรม','สอน อธิบาย หรือช่วยคนแก้ปัญหา'],['กิจกรรม','ชวนคนเห็นด้วยกับไอเดียของคุณ'],['กิจกรรม','จัดแฟ้ม ข้อมูล หรือรายการให้เป็นระเบียบ'],
 ['ความสามารถ','ใช้เครื่องมือหรืออุปกรณ์ได้อย่างคล่องแคล่ว'],['ความสามารถ','วิเคราะห์โจทย์ที่มีเหตุผลหลายขั้นตอน'],['ความสามารถ','คิดไอเดียหรือวิธีนำเสนอแปลกใหม่'],['ความสามารถ','รับฟังความรู้สึกและสื่อสารกับคนอื่น'],['ความสามารถ','วางเป้าหมายและพาคนอื่นทำงานให้สำเร็จ'],['ความสามารถ','ตรวจรายละเอียด ตัวเลข หรือเอกสารได้รอบคอบ'],
 ['อาชีพ','สนใจอาชีพที่ได้ทำงานกับเครื่องมือ เทคโนโลยี หรือภาคสนาม'],['อาชีพ','สนใจอาชีพนักวิทยาศาสตร์ นักวิจัย หรือผู้เชี่ยวชาญ'],['อาชีพ','สนใจอาชีพนักออกแบบ นักเขียน หรือผู้สร้างสื่อ'],['อาชีพ','สนใจอาชีพครู บุคลากรสุขภาพ หรือผู้ให้คำปรึกษา'],['อาชีพ','สนใจอาชีพนักธุรกิจ นักการตลาด หรือผู้นำโครงการ'],['อาชีพ','สนใจอาชีพบัญชี การเงิน ธุรการ หรือวางระบบข้อมูล']
];
const sectionSpecs = [['กิจกรรม',66,0],['ความสามารถ',66,6],['อาชีพ',84,12]];
let questionNumber = 0;
const questionsData = sectionSpecs.flatMap(([type,count,offset]) => Array.from({length:count},(_,index)=>{const q=baseQuestions[offset+(index%6)];questionNumber++;return [String(questionNumber).padStart(3,'0'),type,'RIASEC'[index%6],q[1],1,0]}));
title(questions,'A1:F1','คำถาม RIASEC','แก้ไขข้อความได้ตามบริบทโรงเรียน แต่คงรหัส R/I/A/S/E/C และค่าคำตอบไว้เพื่อสูตรคะแนน');
questions.getRange('A4:F4').values=[['รหัสคำถาม','มิติ S.D.S.','รหัส RIASEC','ข้อความคำถาม','ใช่/สนใจ','ไม่ใช่']];header(questions,'A4:F4');questions.getRange(`A5:F${4+questionsData.length}`).values=questionsData;questions.getRange(`A5:F${4+questionsData.length}`).format={borders:{preset:'insideHorizontal',style:'thin',color:'#E6ECE8'},wrapText:true};questions.getRange('A:A').format.columnWidth=12;questions.getRange('B:B').format.columnWidth=18;questions.getRange('C:C').format.columnWidth=14;questions.getRange('D:D').format.columnWidth=62;questions.getRange('E:F').format.columnWidth=12;questions.freezePanes.freezeRows(4);questions.showGridLines=false;

const riasecCodes=['R','I','A','S','E','C'];
const sectionNames=['activities','competencies','occupations','selfEstimates'];
const sectionLabels=['กิจกรรม','ความสามารถ','อาชีพ','ประเมินตนเอง'];
const questionStartColumn=8, questionCount=questionsData.length, scoreStartColumn=questionStartColumn+questionCount, sectionStartColumn=245, rankKeyStartColumn=sectionStartColumn+sectionNames.length*riasecCodes.length;
function colLetter(column){let value='';while(column>0){const remainder=(column-1)%26;value=String.fromCharCode(65+remainder)+value;column=Math.floor((column-1)/26)}return value}
const questionColumns=Array.from({length:questionCount},(_,index)=>colLetter(questionStartColumn+index));
const scoreColumns=riasecCodes.map((_,index)=>colLetter(scoreStartColumn+index));
const sectionColumns=Object.fromEntries(sectionNames.map((section,sectionIndex)=>[section,riasecCodes.map((_,codeIndex)=>colLetter(sectionStartColumn+sectionIndex*riasecCodes.length+codeIndex))]));
const rankKeyColumns=riasecCodes.map((_,index)=>colLetter(rankKeyStartColumn+index));
const finalResponseColumn=colLetter(rankKeyStartColumn+rankKeyColumns.length-1);
const responseHeaders=['รหัสรายการ','วันเวลา','ชื่อ-นามสกุล','ชื่อเล่น','ห้อง','เลขที่','ยินยอมให้ใช้ข้อมูล',...Array.from({length:questionCount},(_,index)=>`Q${String(index+1).padStart(2,'0')}`),...['R','I','A','S','E','C','Holland Code','บุคลิกภาพเด่น'],...['ชื่อจริง','นามสกุล','ยินยอม'],...['อาชีพที่ 1 (ล่าสุด)','อาชีพที่ 2','อาชีพที่ 3','อาชีพที่ 4','อาชีพที่ 5 (เก่าที่สุด)'],...['อาชีพที่ 1 พบอาชีพในระบบ','อาชีพที่ 2 พบอาชีพในระบบ','อาชีพที่ 3 พบอาชีพในระบบ','อาชีพที่ 4 พบอาชีพในระบบ','อาชีพที่ 5 พบอาชีพในระบบ'],...sectionLabels.flatMap(label=>riasecCodes.map(code=>`${label} ${code}`)),...rankKeyColumns.map((_,index)=>`Tie key ${riasecCodes[index]}`)];
title(responses,`A1:${finalResponseColumn}1`,'คำตอบนักเรียน','Glide ต้องเพิ่มข้อมูล 1 แถวต่อผู้เรียน: รวมคะแนน RIASEC จาก 66 กิจกรรม + 66 ความสามารถ + 84 อาชีพ และใช้คะแนนย่อยตัดสินเมื่อคะแนนรวมเสมอกัน');
responses.getRange(`A4:${finalResponseColumn}4`).values=[responseHeaders];header(responses,`A4:${finalResponseColumn}4`);responses.getRange('A5:G5').values=[['EXAMPLE-001',new Date('2026-07-20T09:00:00'),'ตัวอย่าง นักเรียน','น้องตัวอย่าง','ม.5/1',1,'ยินยอม']];responses.getRange(`${questionColumns[0]}5:${questionColumns[questionColumns.length-1]}5`).values=[Array.from({length:questionCount},(_,index)=>index%19===0?0:1)];
const refs=Object.fromEntries(riasecCodes.map(code=>[code,{total:[],activities:[],competencies:[],occupations:[],selfEstimates:[]}]));questionsData.forEach((question,index)=>{const code=question[2],section=sectionNames[['กิจกรรม','ความสามารถ','อาชีพ','ประเมินตนเอง'].indexOf(question[1])];if(!refs[code])return;const cell=`${questionColumns[index]}${5}`;refs[code].total.push(cell);if(section)refs[code][section].push(cell)});
function formulaSum(cells){return cells.length?`SUM(${cells.join(',')})`:'0'}
function formulaForCells(cells,r){const shifted=cells.map(cell=>cell.replace(/\d+$/,String(r)));return `=IF(C${r}="","",${formulaSum(shifted)})`}
function rankedCodeFormula(r,rank){return rankKeyColumns.reduceRight((fallback,column,index)=>`IF(${column}${r}=LARGE(${rankKeyColumns[0]}${r}:${rankKeyColumns[rankKeyColumns.length-1]}${r},${rank}),"${riasecCodes[index]}",${fallback})`,'""')}
function hollandFormula(r){return `=IF(C${r}="","",${[1,2,3].map(rank=>rankedCodeFormula(r,rank)).join('&')})`}
function personalityFormula(r){const names=['นักปฏิบัติ','นักคิดวิเคราะห์','นักสร้างสรรค์','ผู้เข้าใจผู้คน','ผู้นำการเปลี่ยนแปลง','ผู้จัดระบบ'];return `=IF(C${r}="","",${rankKeyColumns.reduceRight((fallback,column,index)=>`IF(${column}${r}=LARGE(${rankKeyColumns[0]}${r}:${rankKeyColumns[rankKeyColumns.length-1]}${r},1),"${names[index]}",${fallback})`,'"ผู้จัดระบบ"')})`}
function writeResponseFormulas(r){riasecCodes.forEach((code,index)=>responses.getRange(`${scoreColumns[index]}${r}`).formulas=[[formulaForCells(refs[code].total,r)]]);sectionNames.forEach(section=>sectionColumns[section].forEach((column,index)=>responses.getRange(`${column}${r}`).formulas=[[formulaForCells(refs[riasecCodes[index]][section],r)]]));riasecCodes.forEach((code,index)=>{const key=`=${scoreColumns[index]}${r}*1000000000+${sectionColumns.occupations[index]}${r}*1000000+${sectionColumns.competencies[index]}${r}*10000+${sectionColumns.activities[index]}${r}*100+${sectionColumns.selfEstimates[index]}${r}+${6-index}/1000`;responses.getRange(`${rankKeyColumns[index]}${r}`).formulas=[[`=IF(C${r}="","",${key.slice(1)})`]]});responses.getRange(`${scoreColumns.length?colLetter(scoreStartColumn+6):''}${r}`).formulas=[[hollandFormula(r)]];responses.getRange(`${colLetter(scoreStartColumn+7)}${r}`).formulas=[[personalityFormula(r)]]}
writeResponseFormulas(5);
for(let r=6;r<=205;r++)writeResponseFormulas(r);
responses.getRange(`A5:${finalResponseColumn}205`).format={borders:{preset:'insideHorizontal',style:'thin',color:'#E6ECE8'},wrapText:false};responses.getRange('A:A').format.columnWidth=18;responses.getRange('B:B').format.columnWidth=19;responses.getRange('C:G').format.columnWidth=16;responses.getRange(`${questionColumns[0]}:${questionColumns[questionColumns.length-1]}`).format.columnWidth=8;responses.getRange(`${scoreColumns[0]}:${finalResponseColumn}`).format.columnWidth=15;responses.getRange('B5:B205').format.numberFormat='yyyy-mm-dd hh:mm';responses.freezePanes.freezeRows(4);responses.freezePanes.freezeColumns(7);responses.showGridLines=false;

const careerRows = [
 ['วิศวกรหุ่นยนต์','RI','ออกแบบระบบอัตโนมัติ','วิศวกรรมเมคคาทรอนิกส์ / หุ่นยนต์','https://course.mytcas.com/'],['นักพัฒนาซอฟต์แวร์','IC','สร้างระบบและแอป','วิทยาการคอมพิวเตอร์ / IT','https://course.mytcas.com/'],['นักวิเคราะห์ข้อมูล','IC','แปลข้อมูลเป็นข้อค้นพบ','สถิติ / Data Science','https://course.mytcas.com/'],['นักออกแบบ UX/UI','AI','ออกแบบประสบการณ์ดิจิทัล','ออกแบบนิเทศศิลป์ / IT','https://course.mytcas.com/'],['ครู','SA','ออกแบบการเรียนรู้','ครุศาสตร์ / ศึกษาศาสตร์','https://course.mytcas.com/'],['นักจิตวิทยา','SI','ส่งเสริมสุขภาวะทางใจ','จิตวิทยา','https://course.mytcas.com/'],['พยาบาลวิชาชีพ','SI','ดูแลสุขภาพด้วยความรู้และหัวใจ','พยาบาลศาสตร์','https://course.mytcas.com/'],['ผู้ประกอบการ','ER','เปลี่ยนปัญหาเป็นโอกาส','บริหารธุรกิจ / ผู้ประกอบการ','https://course.mytcas.com/'],['นักการตลาดดิจิทัล','EA','เข้าใจลูกค้าและแบรนด์','การตลาด / นิเทศศาสตร์','https://course.mytcas.com/'],['นักบัญชี','CE','จัดการข้อมูลการเงิน','บัญชี','https://course.mytcas.com/'],['นักโลจิสติกส์','CR','จัดระบบการเคลื่อนย้ายสินค้า','โลจิสติกส์ / วิศวกรรมอุตสาหการ','https://course.mytcas.com/'],['เกษตรกรอัจฉริยะ','RI','ใช้ข้อมูลยกระดับการเกษตร','เกษตรศาสตร์ / เกษตรดิจิทัล','https://course.mytcas.com/']
];
title(careers,'A1:E1','อาชีพและคณะ','ตารางอ้างอิงสำหรับหน้า Career Explorer ใน Glide');careers.getRange('A4:E4').values=[['อาชีพ','รหัส RIASEC','ภาพรวมงาน','คณะ/สาขาในไทย','ลิงก์ค้นหาหลักสูตร']];header(careers,'A4:E4');careers.getRange('A5:E16').values=careerRows;careers.getRange('A5:E16').format={borders:{preset:'insideHorizontal',style:'thin',color:'#E6ECE8'},wrapText:true};careers.getRange('A:E').format.columnWidth=28;careers.getRange('C:C').format.columnWidth=42;careers.getRange('D:D').format.columnWidth=38;careers.showGridLines=false;

const universityGroups = [
 ['วิศวกรรมเมคคาทรอนิกส์ / หุ่นยนต์',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าธนบุรี','สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยเทคโนโลยีสุรนารี']],
 ['วิทยาการคอมพิวเตอร์ / IT',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าธนบุรี','สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยมหิดล']],
 ['สถิติ / Data Science',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยธรรมศาสตร์','มหาวิทยาลัยมหิดล','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยเชียงใหม่']],
 ['ออกแบบนิเทศศิลป์ / IT',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าธนบุรี','สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง','มหาวิทยาลัยศิลปากร','มหาวิทยาลัยกรุงเทพ']],
 ['ครุศาสตร์ / ศึกษาศาสตร์',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยเชียงใหม่','มหาวิทยาลัยขอนแก่น','มหาวิทยาลัยศรีนครินทรวิโรฒ']],
 ['จิตวิทยา',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยธรรมศาสตร์','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยเชียงใหม่','มหาวิทยาลัยศรีนครินทรวิโรฒ']],
 ['พยาบาลศาสตร์',['มหาวิทยาลัยมหิดล','จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยเชียงใหม่','มหาวิทยาลัยขอนแก่น','มหาวิทยาลัยสงขลานครินทร์']],
 ['บริหารธุรกิจ / ผู้ประกอบการ',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยธรรมศาสตร์','มหาวิทยาลัยมหิดล','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยเชียงใหม่']],
 ['การตลาด / นิเทศศาสตร์',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยธรรมศาสตร์','มหาวิทยาลัยมหิดล','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยเชียงใหม่']],
 ['บัญชี',['จุฬาลงกรณ์มหาวิทยาลัย','มหาวิทยาลัยธรรมศาสตร์','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยมหิดล','มหาวิทยาลัยเชียงใหม่']],
 ['โลจิสติกส์ / วิศวกรรมอุตสาหการ',['มหาวิทยาลัยบูรพา','มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าธนบุรี','สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง','มหาวิทยาลัยนเรศวร']],
 ['เกษตรศาสตร์ / เกษตรดิจิทัล',['มหาวิทยาลัยเกษตรศาสตร์','มหาวิทยาลัยแม่โจ้','มหาวิทยาลัยขอนแก่น','มหาวิทยาลัยเชียงใหม่','มหาวิทยาลัยสงขลานครินทร์']]
];
const universityRows = universityGroups.flatMap(([field, schools]) => schools.map((school, index) => [field, index + 1, school, 'ตรวจสอบชื่อคณะ/หลักสูตรและรอบรับสมัครล่าสุดใน MyTCAS', 'https://course.mytcas.com/']));
title(universities,'A1:E1','มหาวิทยาลัย 5 ตัวเลือกต่อสาขา','รายการแนะนำเพื่อเริ่มค้นหา ไม่ใช่อันดับทางการ; ตรวจสอบหลักสูตรและรอบ TCAS ปัจจุบันจาก MyTCAS ก่อนตัดสินใจ');
universities.getRange('A4:E4').values=[['สาขา/หลักสูตร','ลำดับแนะนำ','มหาวิทยาลัย','คำแนะนำก่อนสมัคร','ลิงก์ค้นหาหลักสูตร']];header(universities,'A4:E4');universities.getRange(`A5:E${4+universityRows.length}`).values=universityRows;universities.getRange(`A5:E${4+universityRows.length}`).format={borders:{preset:'insideHorizontal',style:'thin',color:'#E6ECE8'},wrapText:true};universities.getRange('A:A').format.columnWidth=38;universities.getRange('B:B').format.columnWidth=14;universities.getRange('C:C').format.columnWidth=42;universities.getRange('D:D').format.columnWidth=50;universities.getRange('E:E').format.columnWidth=32;universities.freezePanes.freezeRows(4);universities.showGridLines=false;

const resourceRows=[['MyTCAS · ค้นหาหลักสูตร','ค้นหาคณะ มหาวิทยาลัย และหลักสูตร','https://course.mytcas.com/'],['MyTCAS · ระบบ TCAS','ข้อมูลการรับสมัครระดับอุดมศึกษา','https://www.mytcas.com/'],['กระทรวงแรงงาน','ข่าวและข้อมูลด้านแรงงานไทย','https://www.mol.go.th/'],['กรมการจัดหางาน','บริการหางานและข้อมูลตลาดแรงงาน','https://www.doe.go.th/'],['JobThai','ค้นหางานและฝึกงานในไทย','https://www.jobthai.com/'],['JobsDB Thailand','งานและคำแนะนำอาชีพ','https://th.jobsdb.com/'],['JOBBKK','ค้นหางานและสร้างเรซูเม่','https://www.jobbkk.com/'],['JobTH','ตำแหน่งงานทั่วประเทศไทย','https://www.jobth.com/']];
title(resources,'A1:C1','ลิงก์ไทย','ใช้ในหน้าทรัพยากรและการศึกษาต่อของ Glide');resources.getRange('A4:C4').values=[['แหล่งข้อมูล','การใช้งาน','URL']];header(resources,'A4:C4');resources.getRange('A5:C12').values=resourceRows;resources.getRange('A5:C12').format={borders:{preset:'insideHorizontal',style:'thin',color:'#E6ECE8'},wrapText:true};resources.getRange('A:C').format.columnWidth=34;resources.getRange('B:B').format.columnWidth=44;resources.showGridLines=false;

title(dashboard,'A1:H1','สรุปผลห้อง','หน้า Dashboard สำหรับครูแนะแนว: สูตรจะนับจากตารางคำตอบนักเรียนอัตโนมัติ');
dashboard.getRange('A4:B8').values=[['นักเรียนที่ส่งคำตอบ',''],['คะแนน R เฉลี่ย',''],['คะแนน I เฉลี่ย',''],['คะแนน A เฉลี่ย',''],['คะแนน S เฉลี่ย','']];dashboard.getRange('D4:E8').values=[['คะแนน E เฉลี่ย',''],['คะแนน C เฉลี่ย',''],['หมายเหตุ','เพิ่มแถวในคำตอบนักเรียนได้ถึงแถว 205'],['ข้อควรใช้','ใช้เพื่อแนะแนว ไม่ใช่ตัดสินอนาคต'],['แหล่งหลักสูตร','https://course.mytcas.com/']];
dashboard.getRange('B4').formulas=[[`=COUNT('คำตอบนักเรียน'!HP5:HP205)`]];dashboard.getRange('B5').formulas=[[`=AVERAGE('คำตอบนักเรียน'!HP5:HP205)`]];dashboard.getRange('B6').formulas=[[`=AVERAGE('คำตอบนักเรียน'!HQ5:HQ205)`]];dashboard.getRange('B7').formulas=[[`=AVERAGE('คำตอบนักเรียน'!HR5:HR205)`]];dashboard.getRange('B8').formulas=[[`=AVERAGE('คำตอบนักเรียน'!HS5:HS205)`]];dashboard.getRange('E4').formulas=[[`=AVERAGE('คำตอบนักเรียน'!HT5:HT205)`]];dashboard.getRange('E5').formulas=[[`=AVERAGE('คำตอบนักเรียน'!HU5:HU205)`]];
dashboard.getRange('A4:B8').format={borders:{preset:'all',style:'thin',color:line}};dashboard.getRange('D4:E8').format={borders:{preset:'all',style:'thin',color:line}};dashboard.getRange('A4:A8').format={fill:mint,font:{bold:true,color:navy}};dashboard.getRange('D4:D8').format={fill:mint,font:{bold:true,color:navy}};dashboard.getRange('A:A').format.columnWidth=27;dashboard.getRange('B:B').format.columnWidth=20;dashboard.getRange('D:D').format.columnWidth=27;dashboard.getRange('E:E').format.columnWidth=46;dashboard.showGridLines=false;

const inspect = await wb.inspect({kind:'table',range:`คำตอบนักเรียน!A4:${finalResponseColumn}7`,include:'values,formulas',tableMaxRows:7,tableMaxCols:90});
console.log(inspect.ndjson);
const errors = await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?',options:{useRegex:true,maxResults:50},summary:'formula errors'});
console.log(errors.ndjson);
const preview = await wb.render({sheetName:'คำตอบนักเรียน',range:`A1:${finalResponseColumn}10`,scale:1,format:'png'});
await fs.writeFile(`${outputDir}/responses-preview.png`,new Uint8Array(await preview.arrayBuffer()));
const xlsx=await SpreadsheetFile.exportXlsx(wb);await xlsx.save(`${outputDir}/RIASEC_Thailand_Glide_GoogleSheets.xlsx`);
