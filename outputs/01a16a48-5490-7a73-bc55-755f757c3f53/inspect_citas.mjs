import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const filePath = "C:/Users/IA ACADEMY 11/Downloads/rptCitDetalladoCitasporEstado.xlsx";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
const structure = await workbook.inspect({ kind: "workbook,sheet", maxChars: 5000 });
console.log("STRUCTURE\n" + structure.ndjson);

for (const sheet of workbook.worksheets.items) {
  const used = sheet.getUsedRange();
  if (!used) continue;
  console.log(`SHEET_META ${JSON.stringify({ name: sheet.name, address: used.address })}`);
  const headerRow = sheet.getRange("A4:AC4").values[0];
  console.log("HEADERS " + JSON.stringify(headerRow.map((value, idx) => value !== null && value !== "" ? { column: idx + 1, header: String(value) } : null).filter(Boolean)));
}
