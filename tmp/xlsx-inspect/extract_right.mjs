import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = await FileBlob.load("C:/Users/ChisanuchaK/Downloads/SDS-data.xlsx");
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("คำตอบนักเรียน");
const values = sheet.getRange("HX1:IJ1995").values;
let last = values.length;
while (last > 1 && values[last - 1].every((value) => value == null || value === "")) last -= 1;
const payload = { range: "HX1:IJ" + last, rows: values.slice(0, last) };
await fs.writeFile("target-right.json", JSON.stringify(payload), "utf8");
console.log(JSON.stringify({ range: payload.range, rows: payload.rows.length, columns: payload.rows[0].length }));
