import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outDir = "C:/Users/IA ACADEMY 11/Documents/ChatGPT/TABLERO AGENDAMIENTO PMS/outputs/01a16a48-5490-7a73-bc55-755f757c3f53";
const workbook = Workbook.create();
const dash = workbook.worksheets.add("Tablero");
const reg = workbook.worksheets.add("Registro");
const navy = "#1F4E78";
const blue = "#D9EAF7";
const text = "#1F2937";
const amberInput = "#FFF2CC";

dash.showGridLines = false;
reg.showGridLines = false;
dash.tabColor = navy;
reg.tabColor = "#D6B656";

// Dashboard title, date selector and selected-day schedule.
dash.getRange("A2").values = [["Tablero de agendamiento diario"]];
dash.getRange("A3").values = [["Fecha consultada"]];
dash.getRange("B3").values = [[new Date(Date.UTC(2026, 9, 23, 12))]];
dash.getRange("D3").values = [["Día"]];
dash.getRange("E3").formulas = [["=CHOOSE(WEEKDAY(B3,2),\"Lunes\",\"Martes\",\"Miércoles\",\"Jueves\",\"Viernes\",\"Sábado\",\"Domingo\")"]];
dash.getRange("G3").values = [["Jornada"]];
dash.getRange("H3").formulas = [["=INDEX($B$45:$B$51,MATCH(E3,$A$45:$A$51,0))"]];

dash.getRange("A5").values = [["Citas agendadas"]];
dash.getRange("A6").formulas = [["=SUM(B10:B39)"]];
dash.getRange("C5").values = [["Meta del equipo"]];
dash.getRange("C6").formulas = [["=COUNTA(Registro!$F$5:$F$34)*IF(WEEKDAY($B$3,2)=7,0,Registro!$B$5)"]];
dash.getRange("E5").values = [["Cumplimiento del equipo"]];
dash.getRange("E6").formulas = [["=IF(C6=0,\"\",A6/C6)"]];
dash.getRange("G5").values = [["Estado del día"]];
dash.getRange("G6").formulas = [["=IF(WEEKDAY($B$3,2)=7,\"No laborable\",IF(C6=0,\"Agrega auxiliares\",IF(COUNTIF(F10:F39,\"Sin registro\")>0,\"Registro incompleto\",IF(E6<0.8,\"Bajo meta\",IF(E6<1,\"En proceso\",\"Meta cumplida\")))))"]];

dash.getRange("A9:F9").values = [["Auxiliar de enfermería", "Citas agendadas", "Meta diaria", "% de avance", "Citas faltantes", "Estado"]];
dash.getRange("A10").formulas = [["=IF(Registro!F5=\"\",\"\",Registro!F5)"]];
dash.getRange("A10:A39").fillDown();
dash.getRange("B10").formulas = [["=IF($A10=\"\",\"\",IF(COUNTIFS(Registro!$A$9:$A$5008,$B$3,Registro!$B$9:$B$5008,$A10)=0,\"\",SUMIFS(Registro!$C$9:$C$5008,Registro!$A$9:$A$5008,$B$3,Registro!$B$9:$B$5008,$A10)))"]];
dash.getRange("B10:B39").fillDown();
dash.getRange("C10").formulas = [["=IF($A10=\"\",\"\",IF(WEEKDAY($B$3,2)=7,0,Registro!$B$5))"]];
dash.getRange("C10:C39").fillDown();
dash.getRange("D10").formulas = [["=IF(OR($A10=\"\",B10=\"\",C10=0),\"\",B10/C10)"]];
dash.getRange("D10:D39").fillDown();
dash.getRange("E10").formulas = [["=IF(OR($A10=\"\",B10=\"\",C10=0),\"\",MAX(0,C10-B10))"]];
dash.getRange("E10:E39").fillDown();
dash.getRange("F10").formulas = [["=IF($A10=\"\",\"\",IF(WEEKDAY($B$3,2)=7,\"No laborable\",IF(B10=\"\",\"Sin registro\",IF(D10<0.8,\"Bajo meta\",IF(D10<1,\"En proceso\",\"Meta cumplida\")))))"]];
dash.getRange("F10:F39").fillDown();

// The schedule reflects the hours supplied by the user; Monday's start/end time was not specified.
dash.getRange("A43:C43").values = [["Horario laboral indicado", null, null]];
dash.getRange("A44:C51").values = [
  ["Día", "Horario", "Horas"],
  ["Lunes", "4 horas", 4],
  ["Martes", "7:00 a. m. a 5:00 p. m.", 10],
  ["Miércoles", "7:00 a. m. a 5:00 p. m.", 10],
  ["Jueves", "7:00 a. m. a 5:00 p. m.", 10],
  ["Viernes", "7:00 a. m. a 4:00 p. m.", 9],
  ["Sábado", "7:00 a. m. a 1:00 p. m.", 6],
  ["Domingo", "No laborable", 0],
];
dash.getRange("E43:F43").values = [["Semáforo de avance", null]];
dash.getRange("E44:F47").values = [
  ["Rojo", "Menos de 80%"],
  ["Amarillo", "80% a menos de 100%"],
  ["Verde", "100% o más"],
  ["Gris", "Sin registro o domingo"],
];

// Input and setup sheet.
reg.getRange("A2").values = [["Registro de agendamiento"]];
reg.getRange("A3").values = [["Una fila por fecha y auxiliar. Si hay varias cargas del mismo día, el tablero las suma."]];
reg.getRange("A5:B6").values = [
  ["Meta diaria por auxiliar", 45],
  ["Regla del semáforo", "Rojo <80%; amarillo 80% a <100%; verde >=100%"],
];
reg.getRange("F4").values = [["Lista de auxiliares"]];
reg.getRange("F5:F34").values = Array.from({ length: 30 }, () => [null]);
reg.getRange("F36").values = [["Completa esta lista antes de registrar citas."]];
reg.getRange("A8:C8").values = [["Fecha", "Auxiliar de enfermería", "Citas agendadas"]];

// A native table provides filter controls and expands when the user continues below the prepared entry area.
const inputTable = reg.tables.add("A8:C5008", true, "AgendaDiaria");
inputTable.style = "TableStyleMedium2";
reg.getRange("A9:C5008").format.fill = amberInput;
reg.getRange("A9:A5008").format.numberFormat = "dd/mm/yyyy";
reg.getRange("C9:C5008").format.numberFormat = "#,##0";
reg.getRange("B9:B5008").dataValidation = { rule: { type: "list", formula1: "=$F$5:$F$34" } };
reg.getRange("C9:C5008").dataValidation = { rule: { type: "whole", operator: "greaterThanOrEqual", formula1: 0 } };

// Shared typography and restrained styling.
dash.getRange("A2:H51").format.font = { name: "Arial", size: 10, color: text };
reg.getRange("A2:G36").format.font = { name: "Arial", size: 10, color: text };
dash.getRange("A2").format.font = { name: "Arial", size: 15, bold: true, color: navy };
reg.getRange("A2").format.font = { name: "Arial", size: 15, bold: true, color: navy };
dash.getRange("A2:H2").format.borders = { bottom: { style: "thin", color: "#9FBAD0" } };
reg.getRange("A2:G2").format.borders = { bottom: { style: "thin", color: "#9FBAD0" } };
dash.getRange("A3:H3").format.verticalAlignment = "center";
dash.getRange("B3").format.fill = amberInput;
dash.getRange("B3").format.numberFormat = "dd/mm/yyyy";
dash.getRange("A3").format.font = { bold: true };
dash.getRange("D3").format.font = { bold: true };
dash.getRange("G3").format.font = { bold: true };
dash.getRange("A5:H5").format.font = { bold: true, color: navy };
dash.getRange("A5:A6").format.fill = blue;
dash.getRange("C5:C6").format.fill = blue;
dash.getRange("E5:E6").format.fill = blue;
dash.getRange("G5:G6").format.fill = blue;
dash.getRange("A6").format.font = { name: "Arial", size: 14, bold: true, color: navy };
dash.getRange("C6").format.font = { name: "Arial", size: 14, bold: true, color: navy };
dash.getRange("E6").format.font = { name: "Arial", size: 14, bold: true, color: navy };
dash.getRange("G6").format.font = { name: "Arial", size: 10, bold: true, color: navy };
dash.getRange("E6").format.numberFormat = "0%";
dash.getRange("A9:F9").format = { fill: navy, font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, wrapText: true, horizontalAlignment: "center", verticalAlignment: "center" };
dash.getRange("A10:F39").format.font = { name: "Arial", size: 10, color: text };
dash.getRange("B10:C39").format.numberFormat = "#,##0";
dash.getRange("D10:D39").format.numberFormat = "0%";
dash.getRange("E10:E39").format.numberFormat = "#,##0";
dash.getRange("B10:E39").format.horizontalAlignment = "center";
dash.getRange("F10:F39").format.horizontalAlignment = "center";
dash.getRange("A9:F39").format.borders = { insideHorizontal: { style: "thin", color: "#E3EAF0" }, bottom: { style: "thin", color: "#D6E0E8" } };

// Dynamic traffic-light formatting: below 80% red, 80% to below 100% amber, and at least 100% green.
const progress = dash.getRange("D10:D39");
progress.conditionalFormats.addCustom('=AND($D10<>"",$D10<0.8)', { fill: "#FCE4D6", font: { color: "#9C0006", bold: true } });
progress.conditionalFormats.addCustom('=AND($D10>=0.8,$D10<1)', { fill: "#FFF2CC", font: { color: "#9A6700", bold: true } });
progress.conditionalFormats.addCustom('=AND($D10<>"",$D10>=1)', { fill: "#E2F0D9", font: { color: "#006100", bold: true } });
const state = dash.getRange("F10:F39");
state.conditionalFormats.add("containsText", { text: "Bajo meta", format: { fill: "#FCE4D6", font: { color: "#9C0006", bold: true } } });
state.conditionalFormats.add("containsText", { text: "En proceso", format: { fill: "#FFF2CC", font: { color: "#9A6700", bold: true } } });
state.conditionalFormats.add("containsText", { text: "Sin registro", format: { fill: "#E7E6E6", font: { color: "#595959" } } });
state.conditionalFormats.add("containsText", { text: "No laborable", format: { fill: "#E7E6E6", font: { color: "#595959" } } });
state.conditionalFormats.add("containsText", { text: "Meta cumplida", format: { fill: "#E2F0D9", font: { color: "#006100", bold: true } } });
dash.getRange("G6").conditionalFormats.add("containsText", { text: "Bajo meta", format: { fill: "#FCE4D6", font: { color: "#9C0006", bold: true } } });
dash.getRange("G6").conditionalFormats.add("containsText", { text: "En proceso", format: { fill: "#FFF2CC", font: { color: "#9A6700", bold: true } } });
dash.getRange("G6").conditionalFormats.add("containsText", { text: "Registro incompleto", format: { fill: "#FFF2CC", font: { color: "#9A6700", bold: true } } });
dash.getRange("G6").conditionalFormats.add("containsText", { text: "Meta cumplida", format: { fill: "#E2F0D9", font: { color: "#006100", bold: true } } });

dash.getRange("A43:C43").format = { fill: blue, font: { bold: true, color: navy } };
dash.getRange("E43:F43").format = { fill: blue, font: { bold: true, color: navy } };
dash.getRange("A44:C44").format = { fill: navy, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
dash.getRange("E44:F44").format = { fill: "#FCE4D6", font: { bold: true, color: "#9C0006" } };
dash.getRange("E45:F45").format = { fill: "#FFF2CC", font: { bold: true, color: "#9A6700" } };
dash.getRange("E46:F46").format = { fill: "#E2F0D9", font: { bold: true, color: "#006100" } };
dash.getRange("E47:F47").format = { fill: "#E7E6E6", font: { color: "#595959" } };
dash.getRange("C45:C51").format.horizontalAlignment = "center";
dash.getRange("A45:C51").format.borders = { insideHorizontal: { style: "thin", color: "#E3EAF0" } };

reg.getRange("A5:B5").format.font = { bold: true };
reg.getRange("B5").format = { fill: amberInput, font: { bold: true, color: text }, horizontalAlignment: "center", numberFormat: "#,##0" };
reg.getRange("A6").format.font = { bold: true };
reg.getRange("B6").format.font = { color: "#595959", italic: true };
reg.getRange("F4").format = { fill: blue, font: { bold: true, color: navy } };
reg.getRange("F5:F34").format.fill = amberInput;
reg.getRange("F5:F34").format.borders = { insideHorizontal: { style: "thin", color: "#E3EAF0" } };
reg.getRange("F36").format = { font: { italic: true, color: "#595959" }, wrapText: true };
reg.getRange("A8:C8").format = { fill: navy, font: { bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };

dash.getRange("A:A").format.columnWidth = 25;
dash.getRange("B:B").format.columnWidth = 17;
dash.getRange("C:C").format.columnWidth = 15;
dash.getRange("D:D").format.columnWidth = 16;
dash.getRange("E:E").format.columnWidth = 20;
dash.getRange("F:F").format.columnWidth = 20;
dash.getRange("G:G").format.columnWidth = 22;
dash.getRange("H:H").format.columnWidth = 38;
dash.getRange("A2:H2").format.rowHeight = 26;
dash.getRange("A3:H3").format.rowHeight = 24;
dash.getRange("A5:H6").format.rowHeight = 24;
dash.getRange("A9:F9").format.rowHeight = 34;
dash.getRange("A10:F39").format.rowHeight = 21;
dash.getRange("A44:C51").format.rowHeight = 21;
dash.getRange("A45:C45").format.rowHeight = 30;

reg.getRange("A:A").format.columnWidth = 19;
reg.getRange("B:B").format.columnWidth = 28;
reg.getRange("C:C").format.columnWidth = 20;
reg.getRange("D:E").format.columnWidth = 3;
reg.getRange("F:F").format.columnWidth = 34;
reg.getRange("A2:G2").format.rowHeight = 26;
reg.getRange("A3:G3").format.rowHeight = 24;
reg.getRange("A8:C8").format.rowHeight = 32;
reg.freezePanes.freezeRows(8);
dash.freezePanes.freezeRows(9);

workbook.recalculate();

// Verify the live entry workflow with temporary sample data, then restore a blank template.
reg.getRange("F5:F8").values = [["Auxiliar 1"], ["Auxiliar 2"], ["Auxiliar 3"], ["Auxiliar 4"]];
reg.getRange("A9:C11").values = [
  [new Date(Date.UTC(2026, 9, 23, 12)), "Auxiliar 1", 35],
  [new Date(Date.UTC(2026, 9, 23, 12)), "Auxiliar 2", 44],
  [new Date(Date.UTC(2026, 9, 23, 12)), "Auxiliar 3", 45],
];
workbook.recalculate();
const workflowCheck = await workbook.inspect({
  kind: "table",
  range: "Tablero!A10:F13",
  include: "values,formulas",
  tableMaxRows: 5,
  tableMaxCols: 6,
  maxChars: 3000,
});
console.log("WORKFLOW_CHECK\n" + workflowCheck.ndjson);
const workflowRows = JSON.parse(workflowCheck.ndjson.trim().split(/\r?\n/)[0]).values;
if (workflowRows[0][5] !== "Bajo meta" || workflowRows[1][5] !== "En proceso" || workflowRows[2][5] !== "Meta cumplida" || workflowRows[3][5] !== "Sin registro") {
  throw new Error("Daily traffic-light workflow did not return the expected statuses.");
}
const teamTargetCheck = dash.getRange("C6").values[0][0];
const teamStateCheck = dash.getRange("G6").values[0][0];
if (teamTargetCheck !== 180 || teamStateCheck !== "Registro incompleto") {
  throw new Error("Team target or incomplete-registration state did not calculate as expected.");
}
const trafficPreview = await workbook.render({ sheetName: "Tablero", range: "A8:F13", scale: 1, format: "png" });
await fs.writeFile(`${outDir}/semaforo_preview.png`, new Uint8Array(await trafficPreview.arrayBuffer()));

dash.getRange("B3").values = [[new Date(Date.UTC(2026, 9, 25, 12))]];
workbook.recalculate();
if (dash.getRange("E3").values[0][0] !== "Domingo" || dash.getRange("C6").values[0][0] !== 0 || dash.getRange("G6").values[0][0] !== "No laborable") {
  throw new Error("Sunday schedule and target did not calculate as expected.");
}

dash.getRange("B3").values = [[new Date(Date.UTC(2026, 9, 23, 12))]];
reg.getRange("F5:F8").clear({ applyTo: "contents" });
reg.getRange("A9:C11").clear({ applyTo: "contents" });
workbook.recalculate();

const dashboardCheck = await workbook.inspect({
  kind: "table",
  range: "Tablero!A2:H51",
  include: "values,formulas",
  tableMaxRows: 55,
  tableMaxCols: 8,
  maxChars: 9000,
});
console.log("DASHBOARD_CHECK\n" + dashboardCheck.ndjson);
const formulaErrors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log("FORMULA_ERRORS\n" + formulaErrors.ndjson);

for (const [sheetName, range, fileName] of [
  ["Tablero", "A1:H51", "tablero_preview.png"],
  ["Registro", "A1:G20", "registro_preview.png"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(`${outDir}/${fileName}`, new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
const outputPath = `${outDir}/Tablero_Agendamiento_Enfermeria.xlsx`;
await xlsx.save(outputPath);
console.log(`EXPORTED ${outputPath}`);

const savedWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const savedDashboardCheck = await savedWorkbook.inspect({
  kind: "table",
  range: "Tablero!A9:F13",
  include: "values,formulas",
  tableMaxRows: 5,
  tableMaxCols: 6,
  maxChars: 3500,
});
console.log("SAVED_WORKBOOK_CHECK\n" + savedDashboardCheck.ndjson);
for (const [sheetName, range, fileName] of [
  ["Tablero", "A1:H51", "tablero_preview.png"],
  ["Registro", "A1:G20", "registro_preview.png"],
]) {
  const preview = await savedWorkbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(`${outDir}/${fileName}`, new Uint8Array(await preview.arrayBuffer()));
}
