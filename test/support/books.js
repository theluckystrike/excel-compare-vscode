'use strict';
// Workbooks and CSV files built in code for the tests (no fixture files). SheetJS comes from NODE_PATH
// (scripts/test.js points it at the table-viewer product's node_modules).
const fs = require('fs');
const os = require('os');
const path = require('path');
const X = require('xlsx');

/** sheets: [[name, rows, { formulas: { 'C2': 'A2+B2' }, merges: ['A1:B1'] }]] -> Uint8Array */
function book(sheets, bookType = 'xlsx') {
  const wb = X.utils.book_new();
  for (const [name, rows, o = {}] of sheets) {
    const ws = X.utils.aoa_to_sheet(rows, { cellDates: true });
    for (const [ref, f] of Object.entries(o.formulas || {})) {
      ws[ref] = { ...(ws[ref] || { t: 'n', v: 0 }), f };
      const r = X.utils.decode_range(ws['!ref'] || 'A1'), c = X.utils.decode_cell(ref);
      r.e.r = Math.max(r.e.r, c.r); r.e.c = Math.max(r.e.c, c.c); ws['!ref'] = X.utils.encode_range(r);
    }
    if (!ws['!ref'] && !rows.length) ws['!ref'] = 'A1';
    if (o.merges) ws['!merges'] = o.merges.map((m) => X.utils.decode_range(m));
    X.utils.book_append_sheet(wb, ws, name);
  }
  return new Uint8Array(X.write(wb, { type: 'array', bookType }));
}

function tmpdir(tag) { return fs.mkdtempSync(path.join(os.tmpdir(), `xlcmp-${tag}-`)); }

module.exports = { X, book, tmpdir };
