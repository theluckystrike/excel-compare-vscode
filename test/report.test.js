'use strict';
// Pro tool 1: the XLSX report reads back (Summary, Changes, one sheet per compared sheet) with fills on the changed,
// added and removed cells; the HTML and CSV reports list the same changes.
const test = require('node:test');
const assert = require('node:assert');
const { book, X } = require('./support/books');
const { loadBook } = require('../src/core/load');
const { compareBooks } = require('../src/core/diff');
const report = require('../src/core/report');
const { sheetNames, crc32 } = require('../src/core/xlsxwrite');
const zlib = require('zlib');

const load = (u8, name = 'x.xlsx') => loadBook({ bytes: u8 }, name);
const a = [['ID', 'Name', 'Qty'], [1, 'pear', 3], [2, 'fig', 5], [3, 'kiwi', 7]];
const b = [['ID', 'Name', 'Qty'], [1, 'pear', 4], [3, 'kiwi', 7], [4, 'plum', 9]];

function partOf(buf, name) {
  // read one member of our own zip (stored names, deflate)
  let p = buf.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04]));
  while (p >= 0) {
    const n = buf.readUInt16LE(p + 26), e = buf.readUInt16LE(p + 28), cs = buf.readUInt32LE(p + 18);
    const nm = buf.subarray(p + 30, p + 30 + n).toString('utf8');
    const data = buf.subarray(p + 30 + n + e, p + 30 + n + e + cs);
    if (nm === name) return zlib.inflateRawSync(data).toString('utf8');
    p = buf.indexOf(Buffer.from([0x50, 0x4b, 0x03, 0x04]), p + 30 + n + e + cs);
  }
  return null;
}

test('XLSX report: SheetJS reads it back; the changed cell is filled yellow, added row green, removed row red', async () => {
  const A = await load(book([['Stock', a]])), B = await load(book([['Stock', b], ['New', [['x']]]]));
  const r = compareBooks(A, B);
  const buf = report.toXlsx(A, B, r, 'old.xlsx', 'new.xlsx');
  const wb = X.read(buf, { type: 'buffer' });
  assert.deepStrictEqual(wb.SheetNames, ['Summary', 'Changes', 'Stock', 'New']);
  const changes = X.utils.sheet_to_json(wb.Sheets.Changes, { header: 1 });
  assert.deepStrictEqual(changes[0], ['Sheet', 'Cell', 'Row in first file', 'Row in second file', 'Column', 'Change', 'Before', 'After']);
  assert.ok(changes.some((row) => row[1] === 'C2' && row[5] === 'changed' && row[6] === 3 && row[7] === 4), JSON.stringify(changes));
  assert.ok(changes.some((row) => row[5] === 'removed' && /fig/.test(row[6])));
  assert.ok(changes.some((row) => row[5] === 'added' && /plum/.test(row[7])));
  const styles = partOf(buf, 'xl/styles.xml');
  assert.match(styles, /FFFFEB9C/); assert.match(styles, /FFC6EFCE/); assert.match(styles, /FFFFC7CE/);
  const sheet3 = partOf(buf, 'xl/worksheets/sheet3.xml');
  assert.match(sheet3, /<c r="F3" s="1"><v>4<\/v><\/c>/, 'changed Qty cell (column F = 3 info columns + C; row 3 = column header + the sheet header row + 1) has the yellow style');
  assert.match(sheet3, /s="3" t="inlineStr"><is><t xml:space="preserve">fig/, 'removed row is red');
  assert.match(sheet3, /s="2" t="inlineStr"><is><t xml:space="preserve">plum/, 'added row is green');
  const summary = X.utils.sheet_to_json(wb.Sheets.Summary, { header: 1 });
  assert.deepStrictEqual(summary[1].slice(0, 7), ['Stock', 'different', 1, 1, 1, 0, 1]);
  assert.deepStrictEqual(summary[2].slice(0, 2), ['New', 'only in the second file']);
});

test('HTML and CSV reports list the same changes', async () => {
  const A = await load(book([['Stock', a]])), B = await load(book([['Stock', b]]));
  const r = compareBooks(A, B);
  const html = report.toHtml(A, B, r, 'old.xlsx', 'new.xlsx');
  assert.match(html, /1 rows changed, 1 added, 1 removed; 1 cells changed/);
  assert.match(html, /<span class="old">3<\/span>4/);
  const csv = report.toCsv(A, B, r);
  assert.ok(csv.startsWith('﻿Sheet,Cell,'));
  assert.match(csv, /Stock,C2,2,2,Qty,changed,3,4/);
  assert.strictEqual(csv.trim().split('\r\n').length, 4);
});

test('sheet names in the report are valid and unique; control characters are dropped from cell XML', () => {
  assert.deepStrictEqual(sheetNames(['a/b', 'A/B', 'x'.repeat(40), 'Summary']), ['a b', 'A B (2)', 'x'.repeat(31), 'Summary']);
  assert.strictEqual(crc32(Buffer.from('123456789')), 0xcbf43926);
});
