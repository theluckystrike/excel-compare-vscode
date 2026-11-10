'use strict';
// Reading and comparing: workbooks built with SheetJS (formulas, dates, merged cells, empty sheets, .xls and .ods),
// CSV quirks, row alignment, column matching, sheet pairing, key matching with moved rows, 50,000-row timing.
const test = require('node:test');
const assert = require('node:assert');
const { book } = require('./support/books');
const { loadBook } = require('../src/core/load');
const { compareBooks, alignSeq, identical } = require('../src/core/diff');
const { rowCells, KIND, CELL } = require('../src/cells');

const load = (u8, name = 'x.xlsx', limits) => loadBook({ bytes: u8 instanceof Uint8Array ? u8 : new Uint8Array(Buffer.from(u8)) }, name, limits);
const kinds = (d) => Array.from(d.k);

test('values: numbers, text, booleans, dates, formulas and merged cells are read', async () => {
  const b = await load(book([['Data', [['Name', 'When', 'N', 'OK'], ['a', new Date(2026, 0, 31), 1.5, true], ['b', new Date(2026, 1, 1, 13, 30), 0.1 + 0.2, false]], { formulas: { E2: 'C2*2' }, merges: ['A5:B5'] }], ['Empty', []]]));
  const s = b.sheets[0];
  assert.deepStrictEqual(s.rows[1].slice(0, 4), ['a', '2026-01-31', '1.5', 'TRUE']);
  assert.deepStrictEqual(s.rows[2].slice(0, 4), ['b', '2026-02-01 13:30:00', '0.3', 'FALSE']);
  assert.strictEqual(s.f[1][4], '=C2*2');
  assert.deepStrictEqual(s.merges, ['A5:B5']);
  assert.strictEqual(b.sheets[1].name, 'Empty');
  assert.strictEqual(b.sheets[1].rows.length <= 1, true);
});

test('inserted and removed rows are aligned, so the rows after them are not marked', async () => {
  const base = [['ID', 'Item', 'Qty']];
  for (let i = 1; i <= 40; i++) base.push([i, 'item ' + i, i * 10]);
  const next = base.map((r) => r.slice());
  next.splice(5, 0, [99, 'new', 1]); // inserted after row 4
  next.splice(30, 1); // a row removed
  next[20][2] = 777; // a changed cell
  const A = await load(book([['S', base]])), B = await load(book([['S', next]]));
  const d = compareBooks(A, B).sheets[0];
  assert.deepStrictEqual([d.counts.added, d.counts.removed, d.counts.changed, d.counts.cells], [1, 1, 1, 1]);
  assert.strictEqual(d.counts.same, 39);
  const i = kinds(d).indexOf(KIND.CHANGED);
  assert.deepStrictEqual([d.ra[i] + 1, d.rb[i] + 1], [20, 21], 'row 20 of the first file is row 21 of the second');
  assert.deepStrictEqual(rowCells(d, A.sheets[0], B.sheets[0], i, 'values')[2], { t: '777', s: CELL.CHANGED, o: '190' });
});

test('changed cell: old and new value; added column found by the header row', async () => {
  const a = [['ID', 'Name', 'Qty'], [1, 'pear', 3], [2, 'fig', 5]];
  const b = [['ID', 'Name', 'Price', 'Qty'], [1, 'pear', 2, 3], [2, 'fig', 4, 6]];
  const A = await load(book([['S', a]])), B = await load(book([['S', b]]));
  const r = compareBooks(A, B);
  const d = r.sheets[0];
  assert.strictEqual(d.byHeader, true);
  assert.deepStrictEqual(d.cols, [{ a: 0, b: 0 }, { a: 1, b: 1 }, { a: -1, b: 2 }, { a: 2, b: 3 }]);
  assert.deepStrictEqual([d.counts.colsAdded, d.counts.changed, d.counts.cells], [1, 1, 1]);
  const i = kinds(d).indexOf(KIND.CHANGED);
  const c = rowCells(d, A.sheets[0], B.sheets[0], i, 'values');
  assert.deepStrictEqual(c[3], { t: '6', s: CELL.CHANGED, o: '5' });
  assert.strictEqual(c[2].s, CELL.ADDED);
});

test('values vs formulas: a changed formula with the same value shows only in formulas mode', async () => {
  const a = book([['S', [['x', 'y'], [2, 4]], { formulas: { B2: 'A2*2' } }]]);
  const b = book([['S', [['x', 'y'], [2, 4]], { formulas: { B2: 'A2+2' } }]]);
  const A = await load(a), B = await load(b);
  assert.ok(identical(compareBooks(A, B, { mode: 'values' })));
  const f = compareBooks(A, B, { mode: 'formulas' });
  assert.strictEqual(f.totals.cells, 1);
  const d = f.sheets[0];
  const c = rowCells(d, A.sheets[0], B.sheets[0], d.chg[0], 'formulas');
  assert.deepStrictEqual(c[1], { t: '=A2+2', s: CELL.CHANGED, o: '=A2*2' });
});

test('sheets: matched by name, then ignoring case, then by position; extra sheets only in one file', async () => {
  const A = await load(book([['Jan', [['a']]], ['Feb', [['b']]], ['Old', [['o']]]]));
  const B = await load(book([['feb', [['b']]], ['Jan', [['a2']]], ['Mar', [['m']]], ['Apr', [['p']]]]));
  const r = compareBooks(A, B);
  const pairs = r.sheets.map((d) => [d.nameA, d.nameB, d.status]);
  assert.deepStrictEqual(pairs.find((p) => p[0] === 'Jan'), ['Jan', 'Jan', 'both']);
  assert.deepStrictEqual(pairs.find((p) => p[0] === 'Feb'), ['Feb', 'feb', 'both']);
  assert.deepStrictEqual(pairs.find((p) => p[0] === 'Old'), ['Old', 'Mar', 'both']);
  assert.deepStrictEqual(pairs.find((p) => p[1] === 'Apr'), [null, 'Apr', 'onlyB']);
  assert.strictEqual(r.sheetsOnlyB, 1);
});

test('merged cells that differ get a note', async () => {
  const A = await load(book([['S', [['a', 'b'], [1, 2]], { merges: ['A1:B1'] }]]));
  const B = await load(book([['S', [['a', 'b'], [1, 2]], { merges: [] }]]));
  const d = compareBooks(A, B).sheets[0];
  assert.match(d.notes.join(' '), /Merged cells differ: no longer merged A1:B1/);
});

test('.xls and .ods workbooks are read like .xlsx', async () => {
  const rows = [['k', 'v'], ['a', 1], ['b', 2]];
  for (const t of ['xls', 'ods']) {
    const b = await load(book([['S', rows]], t === 'xls' ? 'biff8' : 'ods'), `x.${t}`);
    assert.deepStrictEqual(b.sheets[0].rows, [['k', 'v'], ['a', '1'], ['b', '2']], t);
  }
});

test('CSV: BOM, quoted fields with commas and newlines, CRLF, ; delimiter, TSV', async () => {
  const c1 = await load('﻿id,name,note\r\n1,"Smith, J","line1\r\nline2"\r\n2,Lee,"say ""hi"""\r\n', 'a.csv');
  assert.deepStrictEqual(c1.sheets[0].rows, [['id', 'name', 'note'], ['1', 'Smith, J', 'line1\nline2'], ['2', 'Lee', 'say "hi"']]);
  const c2 = await load('id;amount\n1;3,50\n2;4,00\n', 'b.csv');
  assert.strictEqual(c2.delimiter, ';');
  assert.deepStrictEqual(c2.sheets[0].rows[1], ['1', '3,50']);
  const c3 = await load('a\tb\n1\t2\n', 'c.tsv');
  assert.deepStrictEqual(c3.sheets[0].rows, [['a', 'b'], ['1', '2']]);
  // a CSV and an xlsx compare (one sheet each, matched by position)
  const x = await load(book([['Sheet1', [['id', 'name', 'note'], [1, 'Smith, J', 'line1\nline2'], [2, 'Lee', 'changed']]]]));
  const r = compareBooks(c1, x);
  assert.deepStrictEqual([r.totals.changed, r.totals.cells], [1, 1]);
});

test('Pro key match: reordered rows are matched by ID; moved rows are marked; repeated keys get a note', async () => {
  const a = [['ID', 'Name', 'Qty'], [1, 'a', 1], [2, 'b', 2], [3, 'c', 3], [4, 'd', 4], [5, 'e', 5]];
  const b = [['ID', 'Name', 'Qty'], [1, 'a', 1], [4, 'd', 40], [2, 'b', 2], [3, 'c', 3], [6, 'f', 6], [6, 'f2', 7]];
  const A = await load(book([['S', a]])), B = await load(book([['S', b]]));
  const byOrder = compareBooks(A, B);
  const byKey = compareBooks(A, B, { keys: { 0: [0] } });
  const d = byKey.sheets[0];
  assert.deepStrictEqual(d.keys, [0]);
  assert.strictEqual(d.counts.moved, 1);
  assert.strictEqual(d.counts.changed, 1, 'ID 4: moved and changed');
  assert.ok(kinds(d).includes(KIND.MOVED_CHANGED));
  assert.strictEqual(d.counts.removed, 1, 'ID 5');
  assert.strictEqual(d.counts.added, 2, 'ID 6 twice (the repeat is added)');
  assert.match(d.notes.join(' '), /repeat a key value/);
  assert.ok(byOrder.sheets[0].counts.moved === 0);
});

test('50,000 rows: read and compare in a few seconds; one change found', async () => {
  const rows = [['ID', 'Name', 'Value', 'Date']];
  for (let i = 1; i <= 50000; i++) rows.push([i, 'name ' + i, i * 1.25, new Date(Date.UTC(2020, 0, 1 + (i % 365)))]);
  const rows2 = rows.map((r) => r.slice());
  rows2[25000][2] = -1;
  rows2.splice(100, 0, [0, 'inserted', 0, new Date(Date.UTC(2021, 0, 1))]);
  const ua = book([['Big', rows]]), ub = book([['Big', rows2]]);
  const t0 = Date.now();
  const r = compareBooks(await load(ua), await load(ub));
  const ms = Date.now() - t0;
  assert.deepStrictEqual([r.totals.changed, r.totals.added, r.totals.removed, r.totals.cells], [1, 1, 0, 1]);
  assert.ok(ms < 15000, `took ${ms} ms`);
});

test('alignment stays bounded: 200,000 equal rows against a shuffled copy', () => {
  const n = 200000;
  const a = new Int32Array(n), b = new Int32Array(n);
  for (let i = 0; i < n; i++) { a[i] = i % 7; b[i] = (i * 7919) % n; }
  const t0 = Date.now();
  const { pa, pb } = alignSeq(a, b);
  assert.ok(pa.length >= n && pa.length === pb.length);
  assert.ok(Date.now() - t0 < 10000, `${Date.now() - t0} ms`);
});
