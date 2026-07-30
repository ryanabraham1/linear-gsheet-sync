import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "/Users/ryanabraham/Downloads/2026-27 Hardware Resources.xlsx";
const imagePath = "tools/machining-tracker.png";

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 8000,
  tableMaxRows: 8,
  tableMaxCols: 16,
  tableMaxCellChars: 120,
});
console.log("SUMMARY");
console.log(summary.ndjson);

const target = await workbook.inspect({
  kind: "table",
  sheetId: "Machining Tracker",
  maxChars: 24000,
  tableMaxRows: 120,
  tableMaxCols: 30,
  tableMaxCellChars: 300,
});
console.log("TARGET");
console.log(target.ndjson);

const formulas = await workbook.inspect({
  kind: "formula",
  sheetId: "Machining Tracker",
  maxChars: 8000,
  options: { maxResults: 200 },
});
console.log("FORMULAS");
console.log(formulas.ndjson);

const preview = await workbook.render({
  sheetName: "Machining Tracker",
  autoCrop: "all",
  scale: 1.5,
  format: "png",
});
await fs.writeFile(imagePath, new Uint8Array(await preview.arrayBuffer()));
console.log(`PREVIEW ${imagePath}`);

