'use strict';
// Writes the two walkthrough samples (samples/budget-march.xlsx, samples/budget-april.xlsx) with SheetJS.
// April: a changed amount, a row inserted in the middle, a row removed, a changed formula, a new Owner column and
// a new Notes sheet. Run with NODE_PATH pointing at table-viewer-vscode/node_modules (npm run samples does it).
const path = require('path');
const os = require('os');
const mods = process.env.ZOVO_TV_MODULES || path.join(os.homedir(), 'plugin-market-intel-work', 'factory', 'products', 'table-viewer-vscode', 'node_modules');
const X = require(path.join(mods, 'xlsx'));
const out = path.join(__dirname, '..', 'samples');

function sheet(rows, formulaCol) {
  const ws = X.utils.aoa_to_sheet(rows);
  if (formulaCol !== undefined) {
    for (let r = 1; r < rows.length; r++) {
      const ref = X.utils.encode_cell({ r, c: formulaCol });
      const f = rows[r][formulaCol];
      // the cached value is what Excel shows before it recalculates: Q1 + Q2 (row 9 of April: Q1 * 1.1 + Q2)
      if (typeof f === 'string' && f.startsWith('=')) {
        const q1 = rows[r][formulaCol - 2], q2 = rows[r][formulaCol - 1];
        ws[ref] = { t: 'n', f: f.slice(1), v: /\*1\.1/.test(f) ? Math.round((q1 * 1.1 + q2) * 100) / 100 : q1 + q2 };
      }
    }
  }
  return ws;
}
function write(name, sheets) {
  const wb = X.utils.book_new();
  for (const [n, ws] of sheets) X.utils.book_append_sheet(wb, ws, n);
  X.writeFile(wb, path.join(out, name), { bookType: 'xlsx', compression: true });
}
const march = [
  ['ID', 'Department', 'Item', 'Q1', 'Q2', 'Total'],
  [101, 'Sales', 'Travel', 4200, 3900, '=D2+E2'],
  [102, 'Sales', 'Events', 12000, 8000, '=D3+E3'],
  [103, 'Marketing', 'Ads', 15500, 16000, '=D4+E4'],
  [104, 'Marketing', 'Print', 2100, 1800, '=D5+E5'],
  [105, 'Engineering', 'Cloud', 23000, 24500, '=D6+E6'],
  [106, 'Engineering', 'Laptops', 9600, 0, '=D7+E7'],
  [107, 'Support', 'Training', 3100, 2900, '=D8+E8'],
  [108, 'Support', 'Tools', 1200, 1200, '=D9+E9'],
  [109, 'Finance', 'Audit', 7800, 0, '=D10+E10'],
];
const april = [
  ['ID', 'Department', 'Owner', 'Item', 'Q1', 'Q2', 'Total'],
  [101, 'Sales', 'Dana', 'Travel', 4200, 4400, '=E2+F2'],
  [102, 'Sales', 'Dana', 'Events', 12000, 8000, '=E3+F3'],
  [103, 'Marketing', 'Lee', 'Ads', 15500, 16000, '=E4+F4'],
  [110, 'Marketing', 'Lee', 'Video', 0, 5200, '=E5+F5'],
  [105, 'Engineering', 'Sam', 'Cloud', 23000, 26800, '=E6+F6'],
  [106, 'Engineering', 'Sam', 'Laptops', 9600, 0, '=E7+F7'],
  [107, 'Support', 'Kim', 'Training', 3100, 2900, '=E8+F8'],
  [108, 'Support', 'Kim', 'Tools', 1200, 1200, '=E9*1.1+F9'],
  [109, 'Finance', 'Ana', 'Audit', 7800, 0, '=E10+F10'],
];
const rates = [['Currency', 'Rate'], ['EUR', 1.08], ['GBP', 1.27], ['JPY', 0.0066]];
write('budget-march.xlsx', [['Budget', sheet(march, 5)], ['Rates', sheet(rates)]]);
write('budget-april.xlsx', [['Budget', sheet(april, 6)], ['Rates', sheet(rates)], ['Notes', sheet([['Note'], ['Video budget approved in April.'], ['Print moved to the Ads line.']])]]);
console.log('wrote', out);
