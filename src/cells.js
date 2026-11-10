'use strict';
// Shared by the worker (diff, reports) and the extension host (the rows the grid shows): what a cell "is" in each
// compare mode, and the Excel column letters.

/** Text of a cell. mode 'formulas': the formula when the cell has one, else its value. */
function cellText(sheet, r, c, mode) {
  if (!sheet || r < 0 || c < 0) return '';
  if (mode === 'formulas' && sheet.f) {
    const fr = sheet.f[r];
    if (fr && fr[c]) return fr[c];
  }
  const row = sheet.rows[r];
  if (!row) return '';
  const v = row[c];
  return v == null ? '' : v;
}

function colLetter(c) {
  let s = '';
  let n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// Row kinds in a sheet comparison.
const KIND = { SAME: 0, CHANGED: 1, ADDED: 2, REMOVED: 3, MOVED: 4, MOVED_CHANGED: 5 };
// Cell states in a shown row.
const CELL = { SAME: 0, CHANGED: 1, ADDED: 2, REMOVED: 3 };

/**
 * The cells of one aligned row: [{ t: text, s: state, o: old text (changed only) }] for each merged column.
 * d: the sheet comparison, sa / sb: the two sheets (either may be null), i: aligned row index.
 */
function rowCells(d, sa, sb, i, mode) {
  const a = d.ra[i], b = d.rb[i], k = d.k[i];
  const out = new Array(d.cols.length);
  for (let j = 0; j < d.cols.length; j++) {
    const col = d.cols[j];
    const va = a >= 0 && col.a >= 0 ? cellText(sa, a, col.a, mode) : '';
    const vb = b >= 0 && col.b >= 0 ? cellText(sb, b, col.b, mode) : '';
    if (k === KIND.ADDED) out[j] = { t: vb, s: CELL.ADDED };
    else if (k === KIND.REMOVED) out[j] = { t: va, s: CELL.REMOVED };
    else if (col.a < 0) out[j] = { t: vb, s: vb === '' ? CELL.SAME : CELL.ADDED };
    else if (col.b < 0) out[j] = { t: va, s: va === '' ? CELL.SAME : CELL.REMOVED };
    else if (va === vb) out[j] = { t: vb, s: CELL.SAME };
    else out[j] = { t: vb, s: CELL.CHANGED, o: va };
  }
  return out;
}

module.exports = { cellText, colLetter, rowCells, KIND, CELL };
