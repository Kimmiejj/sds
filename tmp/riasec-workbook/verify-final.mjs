import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const inputPath = 'C:/Users/ChisanuchaK/OneDrive/Desktop/sds/RIASEC_Thailand_Glide_GoogleSheets.xlsx';
const outputDir = 'C:/Users/ChisanuchaK/OneDrive/Desktop/sds/tmp/riasec-workbook/final-verify';
await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));

for (const [sheetName, range] of [['คำตอบนักเรียน', 'HP4:JN7'], ['คำถาม RIASEC', 'A4:F8'], ['สรุปผลห้อง', 'A1:H8']]) {
  const result = await workbook.inspect({ kind: 'table', sheetId: sheetName, range, include: 'values,formulas', tableMaxRows: 8, tableMaxCols: 40, maxChars: 12000 });
  console.log(`--- ${sheetName} ${range} ---`);
  console.log(result.ndjson);
}
const errors = await workbook.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A', options: { useRegex: true, maxResults: 100 }, summary: 'final formula error scan' });
console.log('--- ERRORS ---');
console.log(errors.ndjson);
for (const [sheetName, range, fileName] of [['คำตอบนักเรียน', 'HP1:JN10', 'responses-summary.png'], ['สรุปผลห้อง', 'A1:H8', 'dashboard.png']]) {
  const preview = await workbook.render({ sheetName, range, scale: 1.5, format: 'png' });
  await fs.writeFile(`${outputDir}/${fileName}`, new Uint8Array(await preview.arrayBuffer()));
}
