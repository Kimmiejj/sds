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
const questionsData = Array.from({length:2},(_,cycle)=>baseQuestions.map((q,i)=>[String(cycle*18+i+1).padStart(2,'0'),q[0],'RIASEC'[i%6],q[1],1,0])).flat();
title(questions,'A1:F1','คำถาม RIASEC','แก้ไขข้อความได้ตามบริบทโรงเรียน แต่คงรหัส R/I/A/S/E/C และค่าคำตอบไว้เพื่อสูตรคะแนน');
questions.getRange('A4:F4').values=[['รหัสคำถาม','มิติ S.D.S.','รหัส RIASEC','ข้อความคำถาม','ใช่/สนใจ','ไม่ใช่']];header(questions,'A4:F4');questions.getRange('A5:F40').values=questionsData;questions.getRange('A5:F40').format={borders:{preset:'insideHorizontal',style:'thin',color:'#E6ECE8'},wrapText:true};questions.getRange('A:A').format.columnWidth=12;questions.getRange('B:B').format.columnWidth=18;questions.getRange('C:C').format.columnWidth=14;questions.getRange('D:D').format.columnWidth=62;questions.getRange('E:F').format.columnWidth=12;questions.freezePanes.freezeRows(4);questions.showGridLines=false;

const responseHeaders=['รหัสรายการ','วันเวลา','ชื่อ-นามสกุล','ชื่อเล่น','ห้อง','เลขที่','ยินยอมให้ใช้ข้อมูล','Q01','Q02','Q03','Q04','Q05','Q06','Q07','Q08','Q09','Q10','Q11','Q12','Q13','Q14','Q15','Q16','Q17','Q18','Q19','Q20','Q21','Q22','Q23','Q24','Q25','Q26','Q27','Q28','Q29','Q30','Q31','Q32','Q33','Q34','Q35','Q36','R','I','A','S','E','C','Holland Code','บุคลิกภาพเด่น'];
title(responses,'A1:AY1','คำตอบนักเรียน','Glide ต้องเพิ่มข้อมูล 1 แถวต่อผู้เรียน: Q01–Q36 ใช้ 1 เมื่อ ใช่/สนใจ และ 0 เมื่อ ไม่ใช่');
responses.getRange('A4:AY4').values=[responseHeaders];header(responses,'A4:AY4');responses.getRange('A5:G5').values=[['EXAMPLE-001',new Date('2026-07-20T09:00:00'),'ตัวอย่าง นักเรียน','น้องตัวอย่าง','ม.5/1',1,'ยินยอม']];responses.getRange('H5:AQ5').values=[[1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,0,1,1,1,1,1,1,1,1,0,1,1,1,1]];
const scoreCols = [['AR','H,N,T,Z,AF,AL'],['AS','I,O,U,AA,AG,AM'],['AT','J,P,V,AB,AH,AN'],['AU','K,Q,W,AC,AI,AO'],['AV','L,R,X,AD,AJ,AP'],['AW','M,S,Y,AE,AK,AQ']];
for (const [col,refs] of scoreCols) responses.getRange(`${col}5`).formulas=[[`=SUM(${refs.split(',').map(x=>x+'5').join(',')})`]];
function hollandFormula(r){ return `=CONCAT(IF(AR${r}>=LARGE(AR${r}:AW${r},3),"R",""),IF(AS${r}>=LARGE(AR${r}:AW${r},3),"I",""),IF(AT${r}>=LARGE(AR${r}:AW${r},3),"A",""),IF(AU${r}>=LARGE(AR${r}:AW${r},3),"S",""),IF(AV${r}>=LARGE(AR${r}:AW${r},3),"E",""),IF(AW${r}>=LARGE(AR${r}:AW${r},3),"C",""))`; }
function personalityFormula(r){ return `=IF(AR${r}=MAX(AR${r}:AW${r}),"นักปฏิบัติ",IF(AS${r}=MAX(AR${r}:AW${r}),"นักคิดวิเคราะห์",IF(AT${r}=MAX(AR${r}:AW${r}),"นักสร้างสรรค์",IF(AU${r}=MAX(AR${r}:AW${r}),"ผู้เข้าใจผู้คน",IF(AV${r}=MAX(AR${r}:AW${r}),"ผู้นำการเปลี่ยนแปลง","ผู้จัดระบบ")))))`; }
responses.getRange('AX5').formulas=[[hollandFormula(5)]];
responses.getRange('AY5').formulas=[[personalityFormula(5)]];
for (let r=6;r<=205;r++){ for(const [col,refs] of scoreCols) responses.getRange(`${col}${r}`).formulas=[[`=IF(C${r}="","",SUM(${refs.split(',').map(x=>x+r).join(',')}))`]]; responses.getRange(`AX${r}`).formulas=[[`=IF(C${r}="","",${hollandFormula(r).slice(1)})`]];responses.getRange(`AY${r}`).formulas=[[`=IF(C${r}="","",${personalityFormula(r).slice(1)})`]]; }
responses.getRange('A5:AY205').format={borders:{preset:'insideHorizontal',style:'thin',color:'#E6ECE8'},wrapText:false};responses.getRange('A:A').format.columnWidth=18;responses.getRange('B:B').format.columnWidth=19;responses.getRange('C:G').format.columnWidth=16;responses.getRange('H:AQ').format.columnWidth=8;responses.getRange('AR:AY').format.columnWidth=15;responses.getRange('B5:B205').format.numberFormat='yyyy-mm-dd hh:mm';responses.freezePanes.freezeRows(4);responses.freezePanes.freezeColumns(7);responses.showGridLines=false;

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
dashboard.getRange('B4').formulas=[[`=COUNTIF('คำตอบนักเรียน'!C5:C205,"<>")`]];dashboard.getRange('B5').formulas=[[`=AVERAGEIF('คำตอบนักเรียน'!AR5:AR205,">=0")`]];dashboard.getRange('B6').formulas=[[`=AVERAGEIF('คำตอบนักเรียน'!AS5:AS205,">=0")`]];dashboard.getRange('B7').formulas=[[`=AVERAGEIF('คำตอบนักเรียน'!AT5:AT205,">=0")`]];dashboard.getRange('B8').formulas=[[`=AVERAGEIF('คำตอบนักเรียน'!AU5:AU205,">=0")`]];dashboard.getRange('E4').formulas=[[`=AVERAGEIF('คำตอบนักเรียน'!AV5:AV205,">=0")`]];dashboard.getRange('E5').formulas=[[`=AVERAGEIF('คำตอบนักเรียน'!AW5:AW205,">=0")`]];
dashboard.getRange('A4:B8').format={borders:{preset:'all',style:'thin',color:line}};dashboard.getRange('D4:E8').format={borders:{preset:'all',style:'thin',color:line}};dashboard.getRange('A4:A8').format={fill:mint,font:{bold:true,color:navy}};dashboard.getRange('D4:D8').format={fill:mint,font:{bold:true,color:navy}};dashboard.getRange('A:A').format.columnWidth=27;dashboard.getRange('B:B').format.columnWidth=20;dashboard.getRange('D:D').format.columnWidth=27;dashboard.getRange('E:E').format.columnWidth=46;dashboard.showGridLines=false;

const inspect = await wb.inspect({kind:'table',range:"คำตอบนักเรียน!A4:AY7",include:'values,formulas',tableMaxRows:7,tableMaxCols:51});
console.log(inspect.ndjson);
const errors = await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?',options:{useRegex:true,maxResults:50},summary:'formula errors'});
console.log(errors.ndjson);
const preview = await wb.render({sheetName:'คำตอบนักเรียน',range:'A1:AY10',scale:1,format:'png'});
await fs.writeFile(`${outputDir}/responses-preview.png`,new Uint8Array(await preview.arrayBuffer()));
const xlsx=await SpreadsheetFile.exportXlsx(wb);await xlsx.save(`${outputDir}/RIASEC_Thailand_Glide_GoogleSheets.xlsx`);
