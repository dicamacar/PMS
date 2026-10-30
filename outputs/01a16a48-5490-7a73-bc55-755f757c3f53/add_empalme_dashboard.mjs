import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const outDir = "C:/Users/IA ACADEMY 11/Documents/ChatGPT/TABLERO AGENDAMIENTO PMS/outputs/01a16a48-5490-7a73-bc55-755f757c3f53";
const inputPath = "C:/Users/IA ACADEMY 11/Documents/ChatGPT/TABLERO AGENDAMIENTO PMS/outputs/01a16a6f-c5b1-7330-83d9-b63a72597d6d/Tablero_Agendamiento_Enfermeria_Ejemplo.xlsx";
const outputPath = `${outDir}/Tablero_Agendamiento_Con_Ficha_de_Empalme.xlsx`;
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = workbook.worksheets.add("Empalme");

const navy = "#1F4E78";
const lightBlue = "#D9EAF7";
const paleBlue = "#EEF5FA";
const ink = "#1F2937";
const greenFill = "#E2F0D9";
const amberFill = "#FFF2CC";
const redFill = "#FCE4D6";

sheet.showGridLines = false;
sheet.tabColor = navy;

sheet.getRange("A2").values = [["Ficha de empalme"]];
sheet.getRange("A3").values = [["Datos de ejemplo simulados. Úsalos para ver la tarjeta y las gráficas; reemplázalos por información real antes de tomar decisiones."]];

const cardSpecs = [
  { header: "A5:C5", value: "A6:C7", label: "Fichas registradas", formula: "=COUNTA($A$28:$A$57)", fill: paleBlue, color: navy },
  { header: "E5:G5", value: "E6:G7", label: "Pendientes abiertos", formula: "=SUM($H$28:$H$57)", fill: amberFill, color: "#9A6700" },
  { header: "I5:K5", value: "I6:K7", label: "Fichas en proceso", formula: '=COUNTIF($J$28:$J$57,"En curso")', fill: lightBlue, color: navy },
  { header: "M5:O5", value: "M6:O7", label: "Fichas completadas", formula: '=COUNTIF($J$28:$J$57,"Completada")', fill: greenFill, color: "#006100" },
];
for (const card of cardSpecs) {
  sheet.mergeCells(card.header);
  sheet.mergeCells(card.value);
  const h = sheet.getRange(card.header);
  const v = sheet.getRange(card.value);
  h.values = [[card.label]];
  v.formulas = [[card.formula]];
  h.format = { fill: card.fill, font: { name: "Arial", size: 10, bold: true, color: card.color }, horizontalAlignment: "center", verticalAlignment: "center" };
  v.format = { fill: card.fill, font: { name: "Arial", size: 16, bold: true, color: card.color }, horizontalAlignment: "center", verticalAlignment: "center", borders: { bottom: { style: "thin", color: card.fill } } };
}

sheet.getRange("A10:B10").values = [["Estado de la ficha", "Cantidad"]];
sheet.getRange("A11:A13").values = [["Pendiente"], ["En curso"], ["Completada"]];
sheet.getRange("B11").formulas = [["=COUNTIF($J$28:$J$57,A11)"]];
sheet.getRange("B11:B13").fillDown();

sheet.getRange("A16:B16").values = [["Enfoque del empalme", "Pendientes abiertos"]];
const focusAreas = ["Agenda por confirmar", "Reprogramaciones", "Llamadas pendientes", "Novedades de turno", "Gestiones administrativas"];
sheet.getRange("A17:A21").values = focusAreas.map((x) => [x]);
sheet.getRange("B17").formulas = [["=SUMIFS($H$28:$H$57,$E$28:$E$57,A17)"]];
sheet.getRange("B17:B21").fillDown();

// Seeded mock data makes the example reproducible and safe to replace.
let seed = 482031;
const random = () => {
  seed = (seed * 1664525 + 1013904223) >>> 0;
  return seed / 4294967296;
};
const auxiliaries = [
  "Auxiliar de ejemplo 1",
  "Auxiliar de ejemplo 2",
  "Auxiliar de ejemplo 3",
  "Auxiliar de ejemplo 4",
];
const records = [];
for (let i = 1; i <= 30; i++) {
  const day = 1 + Math.floor(random() * 30);
  const fromIndex = Math.floor(random() * auxiliaries.length);
  const toIndex = (fromIndex + 1 + Math.floor(random() * (auxiliaries.length - 1))) % auxiliaries.length;
  const focus = focusAreas[Math.floor(random() * focusAreas.length)];
  const received = 2 + Math.floor(random() * 7);
  const resolved = Math.floor(random() * (received + 1));
  records.push([
    `EMP-${String(i).padStart(3, "0")}`,
    new Date(Date.UTC(2026, 9, day, 12)),
    auxiliaries[fromIndex],
    auxiliaries[toIndex],
    focus,
    received,
    resolved,
    null,
    null,
    null,
  ]);
}

sheet.getRange("A27:J27").values = [[
  "ID de ejemplo",
  "Fecha de empalme",
  "Auxiliar que entrega",
  "Auxiliar que recibe",
  "Enfoque del empalme",
  "Pendientes recibidos",
  "Pendientes resueltos",
  "Pendientes abiertos",
  "Avance",
  "Estado",
]];
sheet.getRange("A28:G57").values = records.map((row) => row.slice(0, 7));
sheet.getRange("H28").formulas = [["=MAX(0,F28-G28)"]];
sheet.getRange("H28:H57").fillDown();
sheet.getRange("I28").formulas = [["=IF(F28=0,0,G28/F28)"]];
sheet.getRange("I28:I57").fillDown();
sheet.getRange("J28").formulas = [['=IF(H28=0,"Completada",IF(G28=0,"Pendiente","En curso"))']];
sheet.getRange("J28:J57").fillDown();
const mockTable = sheet.tables.add("A27:J57", true, "EmpalmesDeEjemplo");
mockTable.style = "TableStyleMedium2";

const chartStatus = sheet.charts.add("bar", sheet.getRange("A10:B13"));
chartStatus.title = "Fichas por estado (registros)";
chartStatus.titleTextStyle.typeface = "Arial";
chartStatus.titleTextStyle.fontSize = 12;
chartStatus.hasLegend = false;
chartStatus.xAxis = { textStyle: { typeface: "Arial", fontSize: 10 } };
chartStatus.yAxis = { numberFormatCode: "0", numberFormatSourceLinked: false, textStyle: { typeface: "Arial", fontSize: 10 } };
chartStatus.series.items[0].fill = navy;
chartStatus.setPosition("D10", "J23");

const chartFocus = sheet.charts.add("bar", sheet.getRange("A16:B21"));
chartFocus.title = "Pendientes abiertos por enfoque (tareas)";
chartFocus.titleTextStyle.typeface = "Arial";
chartFocus.titleTextStyle.fontSize = 12;
chartFocus.hasLegend = false;
chartFocus.xAxis = { textStyle: { typeface: "Arial", fontSize: 10 } };
chartFocus.yAxis = { numberFormatCode: "0", numberFormatSourceLinked: false, textStyle: { typeface: "Arial", fontSize: 10 } };
chartFocus.series.items[0].fill = "#4A86A8";
chartFocus.setPosition("K10", "S23");

sheet.getRange("A2:S57").format.font = { name: "Arial", size: 10, color: ink };
sheet.getRange("A2").format.font = { name: "Arial", size: 15, bold: true, color: navy };
sheet.getRange("A2:S2").format.borders = { bottom: { style: "thin", color: "#9FBAD0" } };
sheet.getRange("A3:S3").format.font = { name: "Arial", size: 9, italic: true, color: "#595959" };
sheet.getRange("A10:B10").format = { fill: navy, font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center" };
sheet.getRange("A16:B16").format = { fill: navy, font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
sheet.getRange("A11:B13").format = { fill: "#FFFFFF", font: { name: "Arial", size: 10, color: ink }, borders: { insideHorizontal: { style: "thin", color: "#E3EAF0" } } };
sheet.getRange("A17:B21").format = { fill: "#FFFFFF", font: { name: "Arial", size: 10, color: ink }, borders: { insideHorizontal: { style: "thin", color: "#E3EAF0" } } };
sheet.getRange("B11:B13").format.horizontalAlignment = "center";
sheet.getRange("B17:B21").format.horizontalAlignment = "center";
sheet.getRange("B11:B13").format.numberFormat = "#,##0";
sheet.getRange("B17:B21").format.numberFormat = "#,##0";
sheet.getRange("A27:J27").format = { fill: navy, font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
sheet.getRange("B28:B57").format.numberFormat = "dd/mm/yyyy";
sheet.getRange("F28:H57").format.numberFormat = "#,##0";
sheet.getRange("I28:I57").format.numberFormat = "0%";
sheet.getRange("F28:I57").format.horizontalAlignment = "center";
sheet.getRange("J28:J57").format.horizontalAlignment = "center";
sheet.getRange("J28:J57").conditionalFormats.add("containsText", { text: "Pendiente", format: { fill: redFill, font: { color: "#9C0006", bold: true } } });
sheet.getRange("J28:J57").conditionalFormats.add("containsText", { text: "En curso", format: { fill: amberFill, font: { color: "#9A6700", bold: true } } });
sheet.getRange("J28:J57").conditionalFormats.add("containsText", { text: "Completada", format: { fill: greenFill, font: { color: "#006100", bold: true } } });

const widths = { A: 16, B: 17, C: 25, D: 25, E: 26, F: 19, G: 18, H: 18, I: 12, J: 16, K: 14, L: 14, M: 14, N: 14, O: 14, P: 12, Q: 12, R: 12, S: 12 };
for (const [col, width] of Object.entries(widths)) sheet.getRange(`${col}:${col}`).format.columnWidth = width;
sheet.getRange("A2:S2").format.rowHeight = 27;
sheet.getRange("A3:S3").format.rowHeight = 23;
sheet.getRange("A5:O7").format.rowHeight = 23;
sheet.getRange("A10:B21").format.rowHeight = 23;
sheet.getRange("A27:J27").format.rowHeight = 36;
sheet.getRange("A28:J57").format.rowHeight = 21;
sheet.freezePanes.freezeRows(27);

workbook.recalculate();

const summaryCheck = await workbook.inspect({
  kind: "table",
  range: "Empalme!A2:O21",
  include: "values,formulas",
  tableMaxRows: 22,
  tableMaxCols: 15,
  maxChars: 8000,
});
console.log("SUMMARY_CHECK\n" + summaryCheck.ndjson);
const detailCheck = await workbook.inspect({
  kind: "table",
  range: "Empalme!A27:J33",
  include: "values,formulas",
  tableMaxRows: 7,
  tableMaxCols: 10,
  maxChars: 4000,
});
console.log("DETAIL_CHECK\n" + detailCheck.ndjson);
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log("FORMULA_ERRORS\n" + errors.ndjson);
const chartCheck = await workbook.inspect({ kind: "drawing", sheetId: "Empalme", maxChars: 4000 });
console.log("CHART_CHECK\n" + chartCheck.ndjson);

const dashboardPreview = await workbook.render({ sheetName: "Empalme", range: "A1:S23", scale: 1, format: "png" });
await fs.writeFile(`${outDir}/empalme_dashboard_preview.png`, new Uint8Array(await dashboardPreview.arrayBuffer()));
const recordsPreview = await workbook.render({ sheetName: "Empalme", range: "A27:J43", scale: 1, format: "png" });
await fs.writeFile(`${outDir}/empalme_records_preview.png`, new Uint8Array(await recordsPreview.arrayBuffer()));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(`EXPORTED ${outputPath}`);

const saved = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const savedDrawings = await saved.inspect({ kind: "drawing", sheetId: "Empalme", maxChars: 4000 });
console.log("EXPORTED_CHART_CHECK\n" + savedDrawings.ndjson);
const savedPreview = await saved.render({ sheetName: "Empalme", range: "A1:S23", scale: 1, format: "png" });
await fs.writeFile(`${outDir}/empalme_exported_preview.png`, new Uint8Array(await savedPreview.arrayBuffer()));
