'use strict';
// Reads a workbook (.xlsx .xlsm .xls .ods) with SheetJS or a delimited text file (.csv .tsv) with the streaming
// parser into plain rows of strings, one array per sheet. Every reader is bounded: file size, zip inflate size,
// sheets, rows, columns and cells have caps, and a cut sheet carries a note that says so.
const fs = require('fs');
const path = require('path');
const { CsvParser, detectDelimiter } = require('./csv');
const { detectEncoding } = require('./encoding');
const { checkZip, isZip } = require('./zipguard');

const MB = 1024 * 1024;
const DEFAULTS = { maxBytes: 200 * MB, maxRows: 300000, maxCols: 1000, maxSheets: 200, maxCells: 12000000, maxMerges: 20000 };
const FORMATS = { xlsx: 'xlsx', xlsm: 'xlsx', xltx: 'xlsx', xls: 'xls', ods: 'ods', csv: 'csv', tsv: 'tsv', tab: 'tsv', txt: 'csv' };
const EXTENSIONS = ['xlsx', 'xlsm', 'xls', 'ods', 'csv', 'tsv'];
const ERRORS = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA' };

function limitError(msg) { const e = new Error(msg); e.code = 'LIMIT'; return e; }
function formatOf(name) { const m = /\.([A-Za-z0-9]+)$/.exec(String(name || '')); return m ? FORMATS[m[1].toLowerCase()] || null : null; }
const fmtMB = (n) => `${Math.round(n / MB)} MB`;

let XLSX = null;
function sheetjs() { if (!XLSX) XLSX = require('xlsx'); return XLSX; }

function numStr(n) {
  if (!Number.isFinite(n)) return String(n);
  if (Number.isInteger(n)) return String(n);
  return String(Number(n.toPrecision(15)));
}
const pad = (n, w = 2) => String(n).padStart(w, '0');
function dateStr(d) {
  if (!(d instanceof Date) || Number.isNaN(d.getTime())) return '';
  const day = `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const h = d.getUTCHours(), m = d.getUTCMinutes(), s = d.getUTCSeconds();
  return h || m || s ? `${day} ${pad(h)}:${pad(m)}:${pad(s)}` : day;
}

/** The compared value of a SheetJS cell, as text. */
function cellValue(c) {
  if (!c) return '';
  switch (c.t) {
    case 'n': return numStr(c.v);
    case 'd': return dateStr(c.v);
    case 'b': return c.v ? 'TRUE' : 'FALSE';
    case 'e': return c.w || ERRORS[c.v] || '#ERROR';
    case 's': return c.v == null ? '' : String(c.v);
    case 'z': return '';
    default: return c.v == null ? '' : String(c.v);
  }
}

function sheetFromWs(X, ws, name, L, budget) {
  const sheet = { name: String(name), rows: [], f: null, cols: 0, merges: [], notes: [], cutRows: 0, cutCols: 0 };
  if (!ws || !ws['!ref']) return sheet;
  const full = X.utils.decode_range(ws['!fullref'] || ws['!ref']);
  const ref = X.utils.decode_range(ws['!ref']);
  const data = ws['!data'] || [];
  let lastRow = Math.min(ref.e.r, data.length - 1);
  let lastCol = ref.e.c;
  if (lastCol >= L.maxCols) { sheet.cutCols = lastCol + 1; lastCol = L.maxCols - 1; }
  const width = lastCol + 1;
  let rowCap = Math.min(L.maxRows, Math.max(1, Math.floor(budget.left / Math.max(1, width))));
  if (full.e.r + 1 > rowCap || lastRow + 1 > rowCap) { sheet.cutRows = Math.max(full.e.r, lastRow) + 1; lastRow = Math.min(lastRow, rowCap - 1); }
  let anyF = false;
  const f = [];
  for (let r = 0; r <= lastRow; r++) {
    const src = data[r];
    const row = [];
    let fr = null;
    if (src) {
      const n = Math.min(src.length, width);
      for (let c = 0; c < n; c++) {
        const cell = src[c];
        if (!cell) { row.push(''); continue; }
        row.push(cellValue(cell));
        if (cell.f) { if (!fr) fr = []; fr[c] = '=' + String(cell.f); anyF = true; }
      }
      while (row.length && row[row.length - 1] === '') row.pop();
    }
    if (row.length > sheet.cols) sheet.cols = row.length;
    sheet.rows.push(row);
    f.push(fr);
  }
  budget.left -= sheet.rows.length * Math.max(1, sheet.cols);
  if (anyF) sheet.f = f;
  const merges = ws['!merges'] || [];
  for (let i = 0; i < merges.length && i < L.maxMerges; i++) sheet.merges.push(X.utils.encode_range(merges[i]));
  if (sheet.cutRows) sheet.notes.push(`Only the first ${sheet.rows.length.toLocaleString('en-US')} of ${sheet.cutRows.toLocaleString('en-US')} rows were read (row limit).`);
  if (sheet.cutCols) sheet.notes.push(`Only the first ${L.maxCols.toLocaleString('en-US')} of ${sheet.cutCols.toLocaleString('en-US')} columns were read (column limit).`);
  return sheet;
}

/** Workbook bytes -> book. */
function loadWorkbookBytes(u8, name, format, limits) {
  const L = { ...DEFAULTS, ...(limits || {}) };
  if (u8.length > L.maxBytes) throw limitError(`This file is ${fmtMB(u8.length)}; files are compared up to ${fmtMB(L.maxBytes)} (setting excelCompare.maxFileSizeMB).`);
  if (u8.length === 0) throw limitError('The file is empty.');
  if (isZip(u8)) checkZip(u8);
  else if (format === 'xlsx' || format === 'ods') {
    // .xlsx and .ods are zip archives; anything else with that name is not a workbook (or is encrypted)
    const cfb = u8.length >= 8 && u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0;
    if (!cfb) throw limitError(`This file has the .${path.extname(name).slice(1) || format} extension but is not a zip archive, so it is not a workbook (it may be cut short or renamed).`);
  } else if (format === 'xls' && !(u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0) && !(u8[0] === 0x09 && (u8[1] === 0x08 || u8[1] === 0x04 || u8[1] === 0x02))) {
    // An .xls that is not a BIFF file can still be an HTML or text export (SheetJS reads those); binary junk is not.
    const head = u8.subarray(0, 4096);
    let ctl = 0;
    for (const b of head) if (b < 9 || (b > 13 && b < 32)) ctl++;
    if (ctl > head.length * 0.05) throw limitError('This file has the .xls extension but is neither an Excel workbook nor a text export (it may be damaged or renamed).');
  }
  const X = sheetjs();
  let wb;
  try {
    wb = X.read(u8, { type: 'array', dense: true, cellDates: true, cellFormula: true, cellHTML: false, cellNF: false, cellStyles: false, cellText: false, UTC: true, sheetRows: L.maxRows + 1, WTF: false });
  } catch (e) {
    const msg = String((e && e.message) || e);
    if (/password|encrypt/i.test(msg)) throw limitError('This workbook is password protected. Save a copy without the password to compare it.');
    const err = new Error(`The workbook could not be read: ${msg.slice(0, 200)}`); err.code = 'READ'; throw err;
  }
  const book = { name, format, sheets: [], notes: [] };
  const names = wb.SheetNames || [];
  if (names.length > L.maxSheets) book.notes.push(`Only the first ${L.maxSheets} of ${names.length} sheets were read (sheet limit).`);
  const budget = { left: L.maxCells };
  for (let i = 0; i < names.length && i < L.maxSheets; i++) {
    if (budget.left <= 0) { book.notes.push(`Sheets after "${names[i - 1]}" were not read: the workbook has more than ${L.maxCells.toLocaleString('en-US')} cells (cell limit).`); break; }
    book.sheets.push(sheetFromWs(X, wb.Sheets[names[i]], names[i], L, budget));
  }
  return book;
}

/** Async chunks of bytes from a file (bounded by maxBytes) or from a buffer. */
async function* chunksOf(src, maxBytes, state) {
  if (src.bytes) {
    const u8 = src.bytes;
    const end = Math.min(u8.length, maxBytes);
    if (u8.length > maxBytes) state.cutBytes = u8.length;
    for (let o = 0; o < end; o += MB) yield u8.subarray(o, Math.min(end, o + MB));
    return;
  }
  const fh = await fs.promises.open(src.file, 'r');
  try {
    const st = await fh.stat();
    if (st.size > maxBytes) state.cutBytes = st.size;
    let pos = 0;
    const buf = Buffer.alloc(MB);
    while (pos < Math.min(st.size, maxBytes)) {
      const { bytesRead } = await fh.read(buf, 0, Math.min(MB, maxBytes - pos), pos);
      if (!bytesRead) break;
      pos += bytesRead;
      yield new Uint8Array(buf.subarray(0, bytesRead));
      if (state.stop) break;
    }
  } finally { await fh.close(); }
}

/** Delimited text -> book with one sheet. Streamed: a 1M-row file stops at the row cap without reading the rest. */
async function loadDelimited(src, name, format, limits) {
  const L = { ...DEFAULTS, ...(limits || {}) };
  const state = { cutBytes: 0, stop: false };
  const sheet = { name: path.basename(name).replace(/\.[^.]+$/, '') || 'Sheet1', rows: [], f: null, cols: 0, merges: [], notes: [], cutRows: 0, cutCols: 0 };
  const book = { name, format, sheets: [sheet], notes: [], encoding: null, delimiter: null };
  let decoder = null, parser = null, pending = '', extraRows = 0, cells = 0;
  const onRow = (fields) => {
    if (sheet.rows.length >= L.maxRows || cells >= L.maxCells) { extraRows++; state.stop = true; parser.stop(); return; }
    if (fields.length > L.maxCols) { sheet.cutCols = Math.max(sheet.cutCols, fields.length); fields.length = L.maxCols; }
    while (fields.length && fields[fields.length - 1] === '') fields.pop();
    // line breaks inside a quoted field: CRLF and CR become LF, as Excel stores them
    for (let i = 0; i < fields.length; i++) if (fields[i].indexOf('\r') >= 0) fields[i] = fields[i].replace(/\r\n?/g, '\n');
    if (fields.length > sheet.cols) sheet.cols = fields.length;
    cells += Math.max(1, fields.length);
    sheet.rows.push(fields);
  };
  let first = true;
  for await (const chunk of chunksOf(src, L.maxBytes, state)) {
    if (first) {
      first = false;
      book.encoding = detectEncoding(chunk.subarray(0, 65536));
      decoder = new TextDecoder(book.encoding);
    }
    const text = decoder.decode(chunk, { stream: true });
    if (!parser) {
      pending += text;
      if (pending.length < 65536) continue;
      book.delimiter = format === 'tsv' ? '\t' : detectDelimiter(pending.slice(0, 65536));
      parser = new CsvParser(book.delimiter, onRow);
      parser.push(pending); pending = '';
    } else parser.push(text);
    if (state.stop) break;
  }
  if (!parser) {
    if (decoder) pending += decoder.decode();
    book.delimiter = format === 'tsv' ? '\t' : detectDelimiter(pending.slice(0, 65536));
    parser = new CsvParser(book.delimiter, onRow);
    parser.push(pending);
  } else if (!state.stop && decoder) parser.push(decoder.decode());
  parser.end();
  if (state.stop) sheet.notes.push(`Only the first ${sheet.rows.length.toLocaleString('en-US')} rows were read (row limit); the file has more.`);
  if (state.cutBytes) sheet.notes.push(`Only the first ${fmtMB(L.maxBytes)} of this ${fmtMB(state.cutBytes)} file were read (setting excelCompare.maxFileSizeMB).`);
  if (sheet.cutCols) sheet.notes.push(`Only the first ${L.maxCols.toLocaleString('en-US')} of ${sheet.cutCols.toLocaleString('en-US')} columns were read (column limit).`);
  return book;
}

/**
 * src: { file } or { bytes: Uint8Array }, name: file name (decides the format). Returns a book:
 * { name, format, sheets: [{ name, rows: string[][], f: (string[]|null)[] | null, cols, merges, notes }], notes }
 */
async function loadBook(src, name, limits) {
  const L = { ...DEFAULTS, ...(limits || {}) };
  const format = formatOf(name);
  if (!format) throw limitError(`${path.basename(name)} is not a file type this extension compares (.xlsx, .xlsm, .xls, .ods, .csv, .tsv).`);
  if (format === 'csv' || format === 'tsv') return loadDelimited(src, name, format, L);
  let u8 = src.bytes;
  if (!u8) {
    const st = await fs.promises.stat(src.file);
    if (st.size > L.maxBytes) throw limitError(`This file is ${fmtMB(st.size)}; files are compared up to ${fmtMB(L.maxBytes)} (setting excelCompare.maxFileSizeMB).`);
    u8 = new Uint8Array(await fs.promises.readFile(src.file));
  }
  return loadWorkbookBytes(u8, name, format, L);
}

module.exports = { loadBook, loadWorkbookBytes, loadDelimited, formatOf, cellValue, numStr, dateStr, DEFAULTS, EXTENSIONS, FORMATS };
