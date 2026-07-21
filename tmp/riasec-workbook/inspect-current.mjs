import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const inputPath = 'C:/Users/ChisanuchaK/OneDrive/Desktop/sds/RIASEC_Thailand_Glide_GoogleSheets.xlsx';
const outputDir = 'C:/Users/ChisanuchaK/OneDrive/Desktop/sds/tmp/riasec-workbook/current-inspect';
await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));

const summary = await workbook.inspect({
  kind: 'sheet',
  include: 'id,name',
  maxChars: 4000,
});
console.log(summary.ndjson);

for (const sheetName of ['คำตอบนักเรียน', 'สรุปผลห้อง']) {
  const check = await workbook.inspect({
    kind: 'table',
    sheetId: sheetName,
    range: sheetName === 'คำตอบนักเรียน' ? 'A1:JN8' : 'A1:H12',
    include: 'values,formulas',
    tableMaxRows: 12,
    tableMaxCols: 300,
    tableMaxCellChars: 100,
    maxChars: 16000,
  });
  console.log(`--- ${sheetName} ---`);
  console.log(check.ndjson);
  const preview = await workbook.render({ sheetName, autoCrop: 'all', scale: 1, format: 'png' });
  await fs.writeFile(`${outputDir}/${sheetName}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const errors = await workbook.inspect({
  kind: 'match',
  searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A',
  options: { useRegex: true, maxResults: 100 },
  summary: 'formula errors',
});
console.log('--- ERRORS ---');
console.log(errors.ndjson);
