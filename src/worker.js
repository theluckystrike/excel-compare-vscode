'use strict';
// Worker thread: reads both files and compares them, so a hostile or huge workbook can never block or crash the
// extension host. The host (runner.js) kills this worker when it takes too long or uses too much memory.
// Bundled with SheetJS into dist/worker.js by scripts/build.js.
const { parentPort } = require('worker_threads');
const { loadBook } = require('./core/load');
const { compareBooks, identical } = require('./core/diff');
const report = require('./core/report');

function src(s) { return s.bytes ? { bytes: new Uint8Array(s.bytes) } : { file: s.file }; }

async function both(job, progress) {
  progress({ step: 'read', which: 1 });
  const A = await loadBook(src(job.a), job.a.name, job.limits);
  progress({ step: 'read', which: 2 });
  const B = await loadBook(src(job.b), job.b.name, job.limits);
  progress({ step: 'compare' });
  const res = compareBooks(A, B, job.opts || {});
  return { A, B, res };
}

async function run(job, progress) {
  if (job.op === 'compare') return both(job, progress);
  if (job.op === 'summary') {
    const { res } = await both(job, progress);
    return { totals: res.totals, same: identical(res), sheets: res.sheets.length, sheetsOnlyA: res.sheetsOnlyA, sheetsOnlyB: res.sheetsOnlyB };
  }
  if (job.op === 'report') {
    const { A, B, res } = await both(job, progress);
    progress({ step: 'write' });
    const names = [job.a.label || job.a.name, job.b.label || job.b.name];
    if (job.kind === 'xlsx') { const buf = report.toXlsx(A, B, res, ...names); return { data: new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength), changes: res.totals }; }
    if (job.kind === 'html') return { text: report.toHtml(A, B, res, ...names), changes: res.totals };
    return { text: report.toCsv(A, B, res), changes: res.totals };
  }
  throw new Error(`unknown job ${job.op}`);
}

parentPort.on('message', async (msg) => {
  try {
    const result = await run(msg.job, (p) => parentPort.postMessage({ id: msg.id, progress: p }));
    parentPort.postMessage({ id: msg.id, ok: true, result });
  } catch (e) {
    parentPort.postMessage({ id: msg.id, ok: false, error: String((e && e.message) || e), code: e && e.code });
  }
});
