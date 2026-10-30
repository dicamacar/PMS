import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourcePath = "C:/Users/IA ACADEMY 11/Downloads/rptCitDetalladoCitasporEstado.xlsx";
const distDir = "C:/Users/IA ACADEMY 11/Documents/ChatGPT/TABLERO AGENDAMIENTO PMS/outputs/01a16a48-5490-7a73-bc55-755f757c3f53/interface/dist";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
const sheet = workbook.worksheets.getItemAt(0);
const table = sheet.getUsedRange();
const totalRows = Number(table.address.match(/:(?:[A-Z]+)(\d+)$/)?.[1]) - 4;
const startRow = 5;
const chunkSize = 4000;
const map = new Map();
const statusTotals = new Map();
const specialtyTotals = new Map();
const monthTotals = new Map();
const requestRouteTotals = new Map();
const attendanceTotals = new Map();
const missing = { assignedDate: 0, registrationDate: 0, registeringUser: 0, status: 0, specialty: 0 };
const records = [];
let minMonth = "9999-99";
let maxMonth = "0000-00";
let processed = 0;

const valueAt = (value, fallback = "Sin dato") => value == null || String(value).trim() === "" ? fallback : String(value).trim();
const monthAt = (value) => {
  if (value == null || value === "") return null;
  let date;
  if (value instanceof Date) date = value;
  else if (typeof value === "number") date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
  else date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
};
const dateAt = (value) => {
  if (value == null || String(value).trim() === "") return "Sin fecha";
  let date;
  if (value instanceof Date) date = value;
  else if (typeof value === "number") date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
  else {
    const raw = String(value).trim();
    const localDate = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (localDate) {
      date = new Date(Date.UTC(Number(localDate[3]), Number(localDate[2]) - 1, Number(localDate[1]), Number(localDate[4] ?? 0), Number(localDate[5] ?? 0), Number(localDate[6] ?? 0)));
    } else date = new Date(raw);
  }
  if (Number.isNaN(date.getTime())) return String(value).trim();
  const day = `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
  const time = [date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()];
  return time.some(Boolean) ? `${day} ${String(time[0]).padStart(2, "0")}:${String(time[1]).padStart(2, "0")}` : day;
};
const inc = (target, key, value = 1) => target.set(key, (target.get(key) ?? 0) + value);

for (let offset = 0; offset < totalRows; offset += chunkSize) {
  const count = Math.min(chunkSize, totalRows - offset);
  const rows = sheet.getRangeByIndexes(startRow - 1 + offset, 0, count, 29).values;
  for (const row of rows) {
    processed++;
    const month = monthAt(row[14]);
    const status = valueAt(row[27]);
    const specialty = valueAt(row[17]);
    const route = valueAt(row[25]);
    const attendance = valueAt(row[26]);
    const registrationDate = dateAt(row[13]);
    const registeringUser = valueAt(row[28]);
    const assignedDate = dateAt(row[14]);
    if (!month) missing.assignedDate++;
    if (registrationDate === "Sin fecha") missing.registrationDate++;
    if (registeringUser === "Sin dato") missing.registeringUser++;
    if (status === "Sin dato") missing.status++;
    if (specialty === "Sin dato") missing.specialty++;
    if (month) {
      minMonth = month < minMonth ? month : minMonth;
      maxMonth = month > maxMonth ? month : maxMonth;
      inc(monthTotals, month);
    }
    inc(statusTotals, status);
    inc(specialtyTotals, specialty);
    inc(requestRouteTotals, route);
    inc(attendanceTotals, attendance);
    records.push({ registrationDate, registeringUser, assignedDate, status, specialty, requestRoute: route });
    const key = JSON.stringify([month ?? "Sin fecha", status, specialty, route, attendance]);
    map.set(key, (map.get(key) ?? 0) + 1);
  }
}

const groups = [...map.entries()].map(([key, count]) => {
  const [month, status, specialty, requestRoute, attendance] = JSON.parse(key);
  return { month, status, specialty, requestRoute, attendance, count };
});
const payload = {
  source: "rptCitDetalladoCitasporEstado.xlsx",
  periodField: "Fecha Asignación Cita",
  totalRows: processed,
  dateRange: { from: minMonth === "9999-99" ? null : minMonth, to: maxMonth === "0000-00" ? null : maxMonth },
  missing,
  groups,
  records,
};

const sumStatus = [...statusTotals.values()].reduce((a, b) => a + b, 0);
if (processed !== totalRows || groups.reduce((a, row) => a + row.count, 0) !== processed || sumStatus !== processed || records.length !== processed) {
  throw new Error("Aggregate reconciliation failed.");
}

await fs.mkdir(distDir, { recursive: true });
await fs.writeFile(`${distDir}/data.js`, `window.CITAS_DATA = ${JSON.stringify(payload)};\n`, "utf8");
console.log(JSON.stringify({
  source: payload.source,
  rows: processed,
  dateRange: payload.dateRange,
  categories: {
    states: [...statusTotals.entries()].sort((a, b) => b[1] - a[1]),
    specialties: specialtyTotals.size,
    requestRoutes: [...requestRouteTotals.entries()].sort((a, b) => b[1] - a[1]),
    attendanceTypes: [...attendanceTotals.entries()].sort((a, b) => b[1] - a[1]),
  },
  missing: payload.missing,
  anonymizedAggregateGroups: groups.length,
  detailRows: records.length,
  reconciliation: { groupsTotal: groups.reduce((a, row) => a + row.count, 0), statesTotal: sumStatus },
}, null, 2));
