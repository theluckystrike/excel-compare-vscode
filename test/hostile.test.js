'use strict';
// Hostile input: zip bombs, a cut or fake workbook, a 1M-row CSV, a huge column count, too many sheets, HTML in
// cells and sheet names, and the worker's memory cap and watchdog.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const { book, tmpdir, X } = require('./support/books');
const { loadBook } = require('../src/core/load');
const { checkZip } = require('../src/core/zipguard');
const { zip } = require('../src/core/xlsxwrite');
const { compareBooks } = require('../src/core/diff');
const report = require('../src/core/report');
const runner = require('../src/runner');

const dir = tmpdir('hostile');
test.after(() => fs.rmSync(dir, { recursive: true, force: true }));
const load = (u8, name = 'x.xlsx', limits) => loadBook({ bytes: u8 }, name, limits);

test('zip bomb: 64 MB of zeros in a 64 KB member is refused before inflating', async () => {
  const bomb = zip([{ name: '[Content_Types].xml', data: '<Types/>' }, { name: 'xl/worksheets/sheet1.xml', data: Buffer.alloc(64 * 1024 * 1024) }]);
  assert.ok(bomb.length < 200 * 1024, `bomb is ${bomb.length} bytes`);
  const t0 = Date.now();
  await assert.rejects(load(new Uint8Array(bomb)), (e) => e.code === 'LIMIT' && /zip bombs/.test(e.message));
  assert.ok(Date.now() - t0 < 2000);
});

test('zip bomb: a directory that claims 5 GB in total is refused', () => {
  const z = zip(Array.from({ length: 30 }, (_, i) => ({ name: `m${i}.xml`, data: 'x' })));
  // patch every central directory entry's uncompressed size to 200 MB (still under the per-member cap)
  let p = z.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  while (p >= 0) { z.writeUInt32LE(200 * 1024 * 1024, p + 24); z.writeUInt32LE(1024 * 1024, p + 20); p = z.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]), p + 4); }
  assert.throws(() => checkZip(new Uint8Array(z)), /inflates to more than 600 MB/);
});

test('cut, fake and empty workbooks give a message, not a crash', async () => {
  const good = book([['S', [['a', 1]]]]);
  await assert.rejects(load(good.subarray(0, Math.floor(good.length / 2))), /directory is missing|cut short|could not be read/);
  await assert.rejects(load(new Uint8Array(Buffer.from('id,name\n1,a\n'))), /not a zip archive/);
  await assert.rejects(load(new Uint8Array(0)), /empty/);
  const rnd = new Uint8Array(4096); for (let i = 0; i < rnd.length; i++) rnd[i] = (i * 2654435761) >>> 24;
  await assert.rejects(load(rnd, 'r.xls'), /neither an Excel workbook nor a text export/);
  await assert.rejects(loadBook({ bytes: new Uint8Array(4) }, 'x.docx'), /not a file type/);
});

test('1,000,000-row CSV: streamed, stops at the row cap with a note, quickly', async () => {
  const f = path.join(dir, 'million.csv');
  const fd = fs.openSync(f, 'w');
  const chunk = [];
  for (let i = 0; i < 1000000; i++) { chunk.push(`${i},row ${i},${i % 97}\n`); if (chunk.length === 50000) { fs.writeSync(fd, chunk.join('')); chunk.length = 0; } }
  fs.closeSync(fd);
  assert.ok(fs.statSync(f).size < 50 * 1024 * 1024);
  const t0 = Date.now();
  const b = await loadBook({ file: f }, f);
  const ms = Date.now() - t0;
  assert.strictEqual(b.sheets[0].rows.length, 300000);
  assert.match(b.sheets[0].notes.join(' '), /first 300,000 rows/);
  assert.ok(ms < 8000, `${ms} ms`);
  fs.rmSync(f);
});

test('huge column count: a cell in column XFD and a 5,000-column CSV are cut at 1,000 columns with a note', async () => {
  const wb = X.utils.book_new();
  const ws = { '!ref': 'A1:XFD3', A1: { t: 's', v: 'start' }, XFD1: { t: 's', v: 'far' }, B2: { t: 'n', v: 1 } };
  X.utils.book_append_sheet(wb, ws, 'Wide');
  const b = await load(new Uint8Array(X.write(wb, { type: 'array', bookType: 'xlsx' })));
  assert.ok(b.sheets[0].cols <= 1000);
  assert.match(b.sheets[0].notes.join(' '), /first 1,000 of 16,384 columns/);
  const wide = Array.from({ length: 5000 }, (_, i) => 'c' + i).join(',') + '\n' + Array.from({ length: 5000 }, (_, i) => i).join(',') + '\n';
  const c = await load(new Uint8Array(Buffer.from(wide)), 'w.csv');
  assert.strictEqual(c.sheets[0].cols, 1000);
  assert.match(c.sheets[0].notes.join(' '), /1,000 of 5,000 columns/);
});

test('too many sheets: the first 200 are read, with a note', async () => {
  const sheets = Array.from({ length: 230 }, (_, i) => [`S${i}`, [['v', i]]]);
  const b = await load(book(sheets));
  assert.strictEqual(b.sheets.length, 200);
  assert.match(b.notes.join(' '), /first 200 of 230 sheets/);
});

test('HTML in cells and sheet names stays text in the reports (escaped), CSV formula cells get a quote', async () => {
  const evil = '<img src=x onerror=alert(1)>';
  const A = await load(book([[evil.slice(0, 28).replace(/[[\]:*?/\\]/g, ''), [['h', 'v'], ['a', '1']]]]));
  const B = await load(book([[evil.slice(0, 28).replace(/[[\]:*?/\\]/g, ''), [['h', 'v'], ['a', evil]], [['=cmd|calc', '+1']]]]));
  const r = compareBooks(A, B);
  const html = report.toHtml(A, B, r, 'a<b>.xlsx', 'b.xlsx');
  assert.ok(!html.includes('<img'), 'no live tag');
  assert.ok(html.includes('&lt;img src=x onerror=alert(1)&gt;'));
  assert.ok(html.includes("default-src 'none'"));
  assert.strictEqual(report.csvField('=1+2'), `"'=1+2"`);
  assert.strictEqual(report.csvField('-5'), '-5');
  assert.strictEqual(report.csvField('@SUM(A1)'), `"'@SUM(A1)"`);
});

test('worker: memory cap and watchdog stop a job; the host goes on and the next job works', async () => {
  const f = path.join(dir, 'w.csv');
  fs.writeFileSync(f, 'a,b\n1,2\n');
  const bad = path.join(dir, 'spin.js');
  fs.writeFileSync(bad, "require('worker_threads').parentPort.on('message', () => { const a = []; for (;;) a.push(new Array(1e6).fill(1)); });");
  const hang = path.join(dir, 'hang.js');
  fs.writeFileSync(hang, "require('worker_threads').parentPort.on('message', () => { for (;;) {} });");
  const real = runner.WORKER;
  try {
    runner.setWorker(bad);
    await assert.rejects(runner.runJob({ op: 'x' }, { memoryMb: 64, idleMs: 20000 }), (e) => e.code === 'MEMORY' || e.code === 'WORKER');
    runner.setWorker(hang);
    await assert.rejects(runner.runJob({ op: 'x' }, { idleMs: 800 }), (e) => e.code === 'TIMEOUT');
  } finally { runner.setWorker(real); }
  const r = await runner.runJob({ op: 'compare', a: { file: f, name: 'w.csv' }, b: { file: f, name: 'w.csv' }, opts: {} });
  assert.strictEqual(r.res.totals.same, 2);
});
