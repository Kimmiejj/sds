import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("C:/Users/ChisanuchaK/Downloads/SDS-data (1).xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("คำตอบนักเรียน");
const headers = sheet.getRange("A1:IJ1").values[0];
const rows = sheet.getRange("A36:IJ39").values;
const columnLetters = (number) => {
  let result = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
};
const payload = {
  sheet: "คำตอบนักเรียน",
  firstRow: 36,
  headers: headers.map((value, index) => ({column: columnLetters(index + 1), header: value ?? ""})),
  rows: rows.map((values, rowIndex) => ({row: rowIndex + 36, values}))
};
await fs.writeFile("rows-36-39.json", JSON.stringify(payload), "utf8");
console.log(JSON.stringify({rows: payload.rows.length, columns: payload.headers.length}));
