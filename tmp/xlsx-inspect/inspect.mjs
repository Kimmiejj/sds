import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("C:/Users/ChisanuchaK/Downloads/SDS-data.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 12000,
  tableMaxRows: 4,
  tableMaxCols: 18,
  tableMaxCellChars: 80,
});
console.log(summary.ndjson);

for (const sheet of workbook.worksheets.items) {
  const used = sheet.getUsedRange(true);
  console.log(JSON.stringify({sheet: sheet.name, usedAddress: used?.address || null}));
  for (const range of ["A1:Z5", "HP1:IL5"]) {
    try {
      const check = await workbook.inspect({
        kind: "table",
        sheetId: sheet.name,
        range,
        include: "values,formulas",
        tableMaxRows: 5,
        tableMaxCols: 18,
        tableMaxCellChars: 80,
      });
      console.log(JSON.stringify({sheet: sheet.name, range, data: check.ndjson}));
    } catch (error) {
      console.log(JSON.stringify({sheet: sheet.name, range, error: String(error)}));
    }
  }
}
