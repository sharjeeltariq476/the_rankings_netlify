import { getStore } from "@netlify/blobs";
import ExcelJS from "exceljs";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import path from "node:path";
import os from "node:os";

const FILE_NAME = "2027 THE Rankings analysis.xlsx";
const BLOB_KEY = "workbooks/2027-the-rankings-analysis.xlsx";
const STORE_NAME = "rankings-excel";

const json = (data, status = 200) => new Response(JSON.stringify(data), {
  status,
  headers: { "content-type": "application/json; charset=utf-8" }
});

function bundledWorkbookPath() {
  return path.resolve(process.cwd(), "data", FILE_NAME);
}

function isNetlifyRuntime() {
  return Boolean(process.env.NETLIFY || process.env.CONTEXT || process.env.SITE_ID);
}

async function getWorkbookBytes() {
  if (isNetlifyRuntime()) {
    const store = getStore({ name: STORE_NAME, consistency: "strong" });
    let bytes = await store.get(BLOB_KEY, { type: "arrayBuffer" });
    if (bytes === null) {
      const initial = await readFile(bundledWorkbookPath());
      const arrayBuffer = initial.buffer.slice(initial.byteOffset, initial.byteOffset + initial.byteLength);
      await store.set(BLOB_KEY, arrayBuffer, { metadata: { filename: FILE_NAME, initialized: new Date().toISOString() } });
      bytes = arrayBuffer;
    }
    return Buffer.from(bytes);
  }
  return readFile(bundledWorkbookPath());
}

async function saveWorkbookBytes(buffer) {
  if (isNetlifyRuntime()) {
    const store = getStore({ name: STORE_NAME, consistency: "strong" });
    const bytes = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    await store.set(BLOB_KEY, bytes, { metadata: { filename: FILE_NAME, updated: new Date().toISOString() } });
    return;
  }
  const dir = path.join(os.tmpdir(), "rankings-excel-local");
  await mkdir(dir, { recursive: true });
  await writeFile(path.join(dir, FILE_NAME), buffer);
}

async function loadWorkbook() {
  const wb = new ExcelJS.Workbook();
  const bytes = await getWorkbookBytes();
  await wb.xlsx.load(bytes);
  return wb;
}

function rawCellValue(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    if (Object.prototype.hasOwnProperty.call(v, "formula")) {
      return v.result ?? null;
    }
    if (Object.prototype.hasOwnProperty.call(v, "richText")) {
      return v.richText.map(x => x.text ?? "").join("");
    }
    if (v instanceof Date) return v.toISOString();
  }
  return v;
}

function formulaText(cell) {
  const v = cell.value;
  if (v && typeof v === "object" && Object.prototype.hasOwnProperty.call(v, "formula")) {
    return `=${v.formula}`;
  }
  return null;
}

function valueAt(ws, row, col) {
  return rawCellValue(ws.getRow(row).getCell(col));
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function lastUsedRow(ws) {
  let max = 1;
  ws.eachRow({ includeEmpty: false }, row => { if (row.number > max) max = row.number; });
  return max;
}

function lastUsedCol(ws) {
  let max = 1;
  ws.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, cell => { if (cell.col > max) max = cell.col; });
  });
  return max;
}

function formulaCount(ws) {
  let count = 0;
  ws.eachRow({ includeEmpty: true }, row => row.eachCell({ includeEmpty: true }, cell => {
    if (formulaText(cell)) count += 1;
  }));
  return count;
}

function parseAddress(address) {
  const m = /^([A-Z]+)([1-9][0-9]*)$/.exec(address || "");
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]), col };
}

async function workbookMeta() {
  const wb = await loadWorkbook();
  let formulas = 0;
  const sheets = wb.worksheets.map(ws => {
    const fc = formulaCount(ws);
    formulas += fc;
    return { name: ws.name, rows: lastUsedRow(ws), columns: lastUsedCol(ws), formula_count: fc };
  });
  return { file: FILE_NAME, sheets, formula_count: formulas, storage: isNetlifyRuntime() ? "Netlify Blobs" : "Bundled Excel" };
}

async function qsPayload() {
  const wb = await loadWorkbook();
  const ws = wb.getWorksheet("QS Ranking");
  if (!ws) throw new Error("QS Ranking sheet not found");
  const maxCol = lastUsedCol(ws);
  const maxRow = lastUsedRow(ws);
  const groups = [];
  for (let col = 3; col <= maxCol; col += 2) {
    const metric = valueAt(ws, 1, col);
    if (metric) groups.push({ metric, score_col: colLetter(col), rank_col: colLetter(col + 1) });
  }
  const records = [];
  for (let r = 3; r <= maxRow; r++) {
    if (valueAt(ws, r, 1) === null && valueAt(ws, r, 2) === null) continue;
    const rec = { excel_row: r, Rank: valueAt(ws, r, 1), Year: valueAt(ws, r, 2), metrics: {} };
    groups.forEach((g, i) => {
      const c = 3 + i * 2;
      rec.metrics[g.metric] = { Score: valueAt(ws, r, c), Rank: valueAt(ws, r, c + 1) };
    });
    records.push(rec);
  }
  return { groups, records };
}

async function timePayload() {
  const wb = await loadWorkbook();
  const ws = wb.getWorksheet("TIME Rankings");
  if (!ws) throw new Error("TIME Rankings sheet not found");
  const maxCol = lastUsedCol(ws);
  const maxRow = lastUsedRow(ws);
  const headers = Array.from({ length: maxCol }, (_, i) => valueAt(ws, 1, i + 1));
  const records = [];
  for (let r = 2; r <= maxRow; r++) {
    const values = headers.map((_, i) => valueAt(ws, r, i + 1));
    if (!values.some(v => v !== null && v !== "")) continue;
    const rec = { excel_row: r };
    headers.forEach((h, i) => { rec[h] = values[i]; });
    records.push(rec);
  }
  return { headers, records };
}

async function qualityPayload() {
  const wb = await loadWorkbook();
  const issues = [];
  const doubleDot = /^\d+\.\.\d+$/;
  const trailingEquals = /\d+=$/;
  wb.worksheets.forEach(ws => {
    ws.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, cell => {
      const v = rawCellValue(cell);
      if (typeof v !== "string") return;
      if (doubleDot.test(v)) issues.push({ sheet: ws.name, cell: cell.address, value: v, reason: "Possible repeated decimal point" });
      else if (trailingEquals.test(v)) issues.push({ sheet: ws.name, cell: cell.address, value: v, reason: "Possible trailing equals sign" });
    }));
  });
  return { issues };
}

async function sheetPayload(name) {
  const wb = await loadWorkbook();
  const ws = wb.getWorksheet(name);
  if (!ws) return null;
  const rows = [];
  const formulas = [];
  const maxRow = lastUsedRow(ws), maxCol = lastUsedCol(ws);
  for (let r = 1; r <= maxRow; r++) {
    const row = [];
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getRow(r).getCell(c);
      const formula = formulaText(cell);
      const item = {
        cell: `${colLetter(c)}${r}`,
        value: rawCellValue(cell),
        formula,
        cached_value: formula ? rawCellValue(cell) : null,
        is_formula: Boolean(formula),
        number_format: cell.numFmt || "General"
      };
      row.push(item);
      if (item.is_formula) formulas.push(item);
    }
    rows.push(row);
  }
  return { sheet: name, max_row: maxRow, max_column: maxCol, merged_ranges: Object.keys(ws._merges || {}), formula_count: formulas.length, formulas, rows };
}

async function updateCell(payload) {
  const { sheet, cell: address, value } = payload || {};
  const pos = parseAddress(address);
  if (!sheet || !pos) return { error: "Invalid sheet or cell address", status: 400 };
  const wb = await loadWorkbook();
  const ws = wb.getWorksheet(sheet);
  if (!ws) return { error: "Sheet not found", status: 404 };
  const cell = ws.getCell(address);
  if (formulaText(cell)) return { error: "Formula cells are read-only in the web editor", status: 409 };
  cell.value = value === "" ? null : value;
  wb.calcProperties.fullCalcOnLoad = true;
  const output = await wb.xlsx.writeBuffer();
  await saveWorkbookBytes(Buffer.from(output));
  return { ok: true, sheet, cell: address, value };
}

export default async (request) => {
  try {
    const url = new URL(request.url);
    const functionPrefix = "/.netlify/functions/api";
    let route = url.pathname;
    if (route.startsWith(functionPrefix)) route = route.slice(functionPrefix.length) || "/";

    if (request.method === "GET" && route === "/workbook") return json(await workbookMeta());
    if (request.method === "GET" && route === "/qs") return json(await qsPayload());
    if (request.method === "GET" && route === "/time") return json(await timePayload());
    if (request.method === "GET" && route === "/quality") return json(await qualityPayload());
    if (request.method === "GET" && route.startsWith("/sheet/")) {
      const result = await sheetPayload(decodeURIComponent(route.slice(7)));
      return result ? json(result) : json({ detail: "Sheet not found" }, 404);
    }
    if (request.method === "PUT" && route === "/cell") {
      const result = await updateCell(await request.json());
      if (result.error) return json({ detail: result.error }, result.status);
      return json(result);
    }
    return json({ detail: "Not found" }, 404);
  } catch (error) {
    console.error(error);
    return json({ detail: error?.message || "Server error" }, 500);
  }
};
