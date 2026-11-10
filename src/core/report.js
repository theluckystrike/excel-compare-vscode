'use strict';
// Pro tool 1: the diff report. XLSX (a Summary sheet, a Changes list and one sheet per compared sheet with the
// changed, added and removed cells filled in colour), HTML (one self-contained page) or CSV (the Changes list).
const { cellText, colLetter, rowCells, KIND, CELL } = require('../cells');
const { writeXlsx, STYLE } = require('./xlsxwrite');

const MAX_CHANGES = 1000000;
const MAX_HTML_ROWS = 20000; // rows per sheet in the HTML page
const KIND_NAME = { 1: 'changed', 2: 'added', 3: 'removed', 4: 'moved', 5: 'moved and changed' };

function sheetLabel(d) { return d.status === 'both' ? (d.nameA === d.nameB ? d.nameB : `${d.nameA} / ${d.nameB}`) : d.nameA || d.nameB; }
function sheetStatus(d) { return d.status === 'onlyA' ? 'only in the first file' : d.status === 'onlyB' ? 'only in the second file' : ''; }

/** One line per changed cell, added row or removed row: { sheet, cell, rowA, rowB, column, kind, before, after }. */
function changeList(A, B, res) {
  const out = [];
  for (const d of res.sheets) {
    const sa = d.ia >= 0 ? A.sheets[d.ia] : null, sb = d.ib >= 0 ? B.sheets[d.ib] : null;
    for (const i of d.chg) {
      if (out.length >= MAX_CHANGES) return { list: out, cut: true };
      const k = d.k[i], a = d.ra[i], b = d.rb[i];
      const base = { sheet: sheetLabel(d), rowA: a >= 0 ? a + 1 : '', rowB: b >= 0 ? b + 1 : '' };
      if (k === KIND.ADDED || k === KIND.REMOVED) {
        const s = k === KIND.ADDED ? sb : sa, r = k === KIND.ADDED ? b : a;
        const text = (s && s.rows[r] ? d.cols.map((c) => cellText(s, r, k === KIND.ADDED ? c.b : c.a, res.mode)) : []).filter((v) => v !== '').join(' | ');
        out.push({ ...base, cell: `row ${r + 1}`, column: '', kind: KIND_NAME[k], before: k === KIND.REMOVED ? text : '', after: k === KIND.ADDED ? text : '' });
        continue;
      }
      if (k === KIND.MOVED) { out.push({ ...base, cell: `row ${b + 1}`, column: '', kind: 'moved', before: `row ${a + 1}`, after: `row ${b + 1}` }); continue; }
      const cells = rowCells(d, sa, sb, i, res.mode);
      cells.forEach((cell, j) => {
        if (cell.s !== CELL.CHANGED) return;
        const col = d.cols[j];
        out.push({ ...base, cell: colLetter(col.b) + (b + 1), column: header(sb, col.b) || colLetter(col.b), kind: k === KIND.MOVED_CHANGED ? 'moved and changed' : 'changed', before: cell.o, after: cell.t });
      });
    }
  }
  return { list: out, cut: false };
}

function header(sheet, c) { return sheet && sheet.rows[0] && c >= 0 ? String(sheet.rows[0][c] || '') : ''; }

function summaryRows(res, nameA, nameB) {
  const rows = [[{ v: 'Sheet', s: STYLE.HEADER }, { v: 'Status', s: STYLE.HEADER }, { v: 'Rows changed', s: STYLE.HEADER }, { v: 'Rows added', s: STYLE.HEADER }, { v: 'Rows removed', s: STYLE.HEADER }, { v: 'Rows moved', s: STYLE.HEADER }, { v: 'Cells changed', s: STYLE.HEADER }, { v: 'Columns added', s: STYLE.HEADER }, { v: 'Columns removed', s: STYLE.HEADER }, { v: 'Notes', s: STYLE.HEADER }]];
  for (const d of res.sheets) {
    const c = d.counts;
    rows.push([sheetLabel(d), sheetStatus(d) || (c.changed + c.added + c.removed + c.moved + c.colsAdded + c.colsRemoved ? 'different' : 'same'), c.changed, c.added, c.removed, c.moved, c.cells, c.colsAdded, c.colsRemoved, d.notes.join(' ')]);
  }
  rows.push([]);
  rows.push(['First file', nameA]);
  rows.push(['Second file', nameB]);
  rows.push(['Compared', res.mode === 'formulas' ? 'formulas (a cell without a formula: its value)' : 'values']);
  return rows;
}

const KIND_STYLE = { [KIND.CHANGED]: 0, [KIND.ADDED]: STYLE.ADDED, [KIND.REMOVED]: STYLE.REMOVED, [KIND.MOVED]: STYLE.MOVED, [KIND.MOVED_CHANGED]: STYLE.MOVED };
const CELL_STYLE = { [CELL.SAME]: 0, [CELL.CHANGED]: STYLE.CHANGED, [CELL.ADDED]: STYLE.ADDED, [CELL.REMOVED]: STYLE.REMOVED };
const KIND_LABEL = { 0: '', 1: 'changed', 2: 'added', 3: 'removed', 4: 'moved', 5: 'moved, changed' };

/** The diff report as an .xlsx Buffer. */
function toXlsx(A, B, res, nameA, nameB) {
  const sheets = [{ name: 'Summary', rows: summaryRows(res, nameA, nameB), widths: [24, 22, 13, 11, 13, 11, 13, 14, 16, 60], freeze: true }];
  const { list, cut } = changeList(A, B, res);
  const ch = [[{ v: 'Sheet', s: STYLE.HEADER }, { v: 'Cell', s: STYLE.HEADER }, { v: 'Row in first file', s: STYLE.HEADER }, { v: 'Row in second file', s: STYLE.HEADER }, { v: 'Column', s: STYLE.HEADER }, { v: 'Change', s: STYLE.HEADER }, { v: 'Before', s: STYLE.HEADER }, { v: 'After', s: STYLE.HEADER }]];
  for (const x of list) ch.push([x.sheet, x.cell, x.rowA, x.rowB, x.column, x.kind, { v: x.before, num: true }, { v: x.after, num: true }]);
  if (cut) ch.push([`The list stops at ${MAX_CHANGES.toLocaleString('en-US')} changes.`]);
  sheets.push({ name: 'Changes', rows: ch, widths: [20, 10, 16, 18, 18, 16, 40, 40], freeze: true, autoFilter: true });
  for (const d of res.sheets) {
    const sa = d.ia >= 0 ? A.sheets[d.ia] : null, sb = d.ib >= 0 ? B.sheets[d.ib] : null;
    const head = [{ v: 'Change', s: STYLE.HEADER }, { v: 'Row (first)', s: STYLE.HEADER }, { v: 'Row (second)', s: STYLE.HEADER }];
    for (const c of d.cols) head.push({ v: c.b >= 0 ? colLetter(c.b) + (c.a >= 0 && c.a !== c.b ? ` (was ${colLetter(c.a)})` : '') : `${colLetter(c.a)} (removed)`, s: STYLE.HEADER });
    const rows = [head];
    for (let i = 0; i < d.k.length; i++) {
      const k = d.k[i];
      const cells = rowCells(d, sa, sb, i, res.mode);
      const line = [{ v: KIND_LABEL[k], s: KIND_STYLE[k] || 0 }, d.ra[i] >= 0 ? d.ra[i] + 1 : '', d.rb[i] >= 0 ? d.rb[i] + 1 : ''];
      for (const cell of cells) line.push({ v: cell.s === CELL.CHANGED ? `${cell.t}` : cell.t, s: CELL_STYLE[cell.s] || (k === KIND.MOVED ? STYLE.MOVED : 0), num: true });
      rows.push(line);
    }
    sheets.push({ name: sheetLabel(d), rows, widths: [14, 11, 12, ...d.cols.map(() => 14)].slice(0, 400), freeze: true });
  }
  return writeXlsx(sheets);
}

function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }

/** The diff report as one HTML page (no scripts, no remote files). */
function toHtml(A, B, res, nameA, nameB) {
  const t = res.totals;
  const parts = [`<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'"><title>${esc(nameA)} vs ${esc(nameB)}</title><style>
body{font:13px/1.4 -apple-system,Segoe UI,sans-serif;margin:20px;color:#222}table{border-collapse:collapse;margin:8px 0 24px}td,th{border:1px solid #ccc;padding:2px 6px;vertical-align:top;white-space:pre-wrap;max-width:360px}
th{background:#d9e1f2;text-align:left}.c{background:#ffeb9c}.a{background:#c6efce}.r{background:#ffc7ce}.m{background:#ddebf7}.old{color:#9c0006;text-decoration:line-through;display:block}.muted{color:#666}</style></head><body>`];
  parts.push(`<h1>Excel Compare report</h1><p>First file: <b>${esc(nameA)}</b><br>Second file: <b>${esc(nameB)}</b><br>Compared: ${res.mode === 'formulas' ? 'formulas' : 'values'}</p>`);
  parts.push(`<p>${t.changed} rows changed, ${t.added} added, ${t.removed} removed${t.moved ? `, ${t.moved} moved` : ''}; ${t.cells} cells changed${t.colsAdded ? `; ${t.colsAdded} columns added` : ''}${t.colsRemoved ? `; ${t.colsRemoved} columns removed` : ''}.</p>`);
  parts.push('<table><tr><th>Sheet</th><th>Status</th><th>Rows changed</th><th>Added</th><th>Removed</th><th>Moved</th><th>Cells changed</th><th>Notes</th></tr>');
  for (const d of res.sheets) { const c = d.counts; parts.push(`<tr><td>${esc(sheetLabel(d))}</td><td>${esc(sheetStatus(d) || 'compared')}</td><td>${c.changed}</td><td>${c.added}</td><td>${c.removed}</td><td>${c.moved}</td><td>${c.cells}</td><td>${esc(d.notes.join(' '))}</td></tr>`); }
  parts.push('</table>');
  const cls = { [CELL.CHANGED]: 'c', [CELL.ADDED]: 'a', [CELL.REMOVED]: 'r' };
  for (const d of res.sheets) {
    if (!d.chg.length) continue;
    const sa = d.ia >= 0 ? A.sheets[d.ia] : null, sb = d.ib >= 0 ? B.sheets[d.ib] : null;
    parts.push(`<h2>${esc(sheetLabel(d))}</h2><table><tr><th>Change</th><th>Row (first)</th><th>Row (second)</th>${d.cols.map((c) => `<th>${esc(c.b >= 0 ? colLetter(c.b) : colLetter(c.a) + ' (removed)')}</th>`).join('')}</tr>`);
    let n = 0;
    for (const i of d.chg) {
      if (++n > MAX_HTML_ROWS) { parts.push(`<tr><td colspan="${d.cols.length + 3}" class="muted">The page stops at ${MAX_HTML_ROWS.toLocaleString('en-US')} changed rows for this sheet; the XLSX report has all of them.</td></tr>`); break; }
      const k = d.k[i];
      const cells = rowCells(d, sa, sb, i, res.mode);
      const rc = k === KIND.ADDED ? 'a' : k === KIND.REMOVED ? 'r' : k === KIND.MOVED || k === KIND.MOVED_CHANGED ? 'm' : '';
      parts.push(`<tr><td class="${rc}">${esc(KIND_LABEL[k])}</td><td>${d.ra[i] >= 0 ? d.ra[i] + 1 : ''}</td><td>${d.rb[i] >= 0 ? d.rb[i] + 1 : ''}</td>${cells.map((c) => `<td class="${cls[c.s] || ''}">${c.s === CELL.CHANGED ? `<span class="old">${esc(c.o)}</span>` : ''}${esc(c.t)}</td>`).join('')}</tr>`);
    }
    parts.push('</table>');
  }
  parts.push('<p class="muted">Made with Excel Compare & XLSX Diff for VS Code.</p></body></html>');
  return parts.join('\n');
}

const FORMULA_START = /^[=+@\t\r]|^-(?![\d.])/;
function csvField(v) { const s = String(v == null ? '' : v); const f = FORMULA_START.test(s); return f || /[",\r\n]/.test(s) ? `"${(f ? "'" : '') + s.replace(/"/g, '""')}"` : s; }

/** The Changes list as CSV (UTF-8 with BOM, so Excel reads it right). Cells that start like a formula get a quote. */
function toCsv(A, B, res) {
  const { list, cut } = changeList(A, B, res);
  const lines = [['Sheet', 'Cell', 'Row in first file', 'Row in second file', 'Column', 'Change', 'Before', 'After'].join(',')];
  for (const x of list) lines.push([x.sheet, x.cell, x.rowA, x.rowB, x.column, x.kind, x.before, x.after].map(csvField).join(','));
  if (cut) lines.push(csvField(`The list stops at ${MAX_CHANGES} changes.`));
  return '﻿' + lines.join('\r\n') + '\r\n';
}

module.exports = { toXlsx, toHtml, toCsv, changeList, sheetLabel, csvField, esc };
