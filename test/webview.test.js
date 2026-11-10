'use strict';
// Runs media/main.js in jsdom against the real extension host (stub vscode): tabs with counts, the virtual grid
// (rows asked from the host), changed cells with old and new value, next change, cell detail, hostile names as
// text, the folder table. jsdom is not a dependency of this extension: the test uses JSDOM_PATH or the copy in a
// sibling product, and is skipped when there is none.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { book, tmpdir } = require('./support/books');
const { loadExtension } = require('./support/vscode-stub');

const candidates = [process.env.JSDOM_PATH, path.join(__dirname, '..', 'node_modules', 'jsdom'), path.join(__dirname, '..', '..', 'docx-viewer-vscode', 'node_modules', 'jsdom')].filter(Boolean);
let JSDOM = null;
for (const c of candidates) { try { ({ JSDOM } = require(c)); break; } catch { /* next */ } }
const script = fs.readFileSync(path.join(__dirname, '..', 'media', 'main.js'), 'utf8');
const dir = tmpdir('wv');
test.after(() => fs.rmSync(dir, { recursive: true, force: true }));
const tick = (ms = 30) => new Promise((r) => setTimeout(r, ms));

async function boot(a, b) {
  const s = loadExtension();
  s.ext.activate(s.context);
  const U = s.vscode.Uri;
  await s.vscode.commands.executeCommand('excelCompare.compareSelected', U.file(a), [U.file(a), U.file(b)]);
  const panel = s.panels.at(-1);
  const dom = new JSDOM('<!DOCTYPE html><body><div id="app"></div><template id="zovo-footer"><div id="zovo-support">Rate</div></template></body>', { runScripts: 'outside-only', pretendToBeVisual: true });
  const sent = [];
  let relayed = 0;
  const relay = () => { while (relayed < panel.posted.length) dom.window.dispatchEvent(new dom.window.MessageEvent('message', { data: JSON.parse(JSON.stringify(panel.posted[relayed++])) })); };
  dom.window.acquireVsCodeApi = () => ({ postMessage: (m) => { sent.push(m); panel.send(JSON.parse(JSON.stringify(m))).then(() => setTimeout(relay, 0)); } });
  dom.window.eval(script);
  for (let i = 0; i < 100 && !panel.posted.some((m) => m.type === 'result'); i++) await tick(20);
  relay(); await tick(60); relay(); await tick(30);
  return { s, dom, doc: dom.window.document, sent, panel, relay };
}

test('webview: tabs, counts, grid rows from the host, changed cell old/new, next change, detail, footer', { skip: !JSDOM && 'jsdom not found' }, async () => {
  const evil = '<img src=x onerror=alert(1)>';
  const a = path.join(dir, 'a.xlsx'), b = path.join(dir, 'b.xlsx');
  fs.writeFileSync(a, book([['Data', [['ID', 'Name', 'Qty'], [1, 'pear', 3], [2, 'fig', 5], [3, evil, 1]]], ['Other', [['x'], [1]]]]));
  fs.writeFileSync(b, book([['Data', [['ID', 'Name', 'Qty'], [1, 'pear', 4], [2, 'fig', 5], [3, evil, 2], [9, 'new', 1]]], ['Other', [['x'], [1]]]]));
  const { s, doc, sent, dom } = await boot(a, b);
  try {
    assert.ok(sent.some((m) => m.type === 'rendered'));
    const tabs = [...doc.querySelectorAll('.tab')].map((t) => t.textContent);
    assert.deepStrictEqual(tabs, ['Data3', 'Other=']);
    assert.match(doc.querySelector('.totals').textContent, /2 rows changed, 1 row added, 2 cells changed/);
    const rows = [...doc.querySelectorAll('.gbody .tr')];
    assert.strictEqual(rows.length, 5);
    assert.ok(rows[0].style.getPropertyValue('top') === '0px' && rows[1].style.getPropertyValue('top') === '24px', 'positions set through the CSSOM');
    const ch = doc.querySelector('.td.changed');
    assert.strictEqual(ch.querySelector('.old').textContent, '3');
    assert.strictEqual(ch.querySelector('.new').textContent, '4');
    assert.ok(doc.querySelector('.tr.added'));
    assert.strictEqual(doc.querySelectorAll('#app img').length, 0, 'HTML in a cell stays text');
    assert.ok([...doc.querySelectorAll('.td')].some((t) => t.textContent === evil));
    doc.querySelector('[data-nav=next]').click(); await tick();
    assert.strictEqual(doc.querySelector('.pos').textContent, 'Change 1 of 3');
    assert.ok(doc.querySelector('.tr.cur'));
    doc.querySelector('[data-nav=next]').click(); await tick();
    assert.strictEqual(doc.querySelector('.pos').textContent, 'Change 2 of 3');
    doc.querySelector('.td.changed').dispatchEvent(new dom.window.MouseEvent('click', { bubbles: true }));
    assert.match(doc.querySelector('.detail').textContent, /C2 \(Qty\) before: 3 after: 4/);
    assert.ok(doc.getElementById('zovo-support'));
    assert.strictEqual(doc.querySelectorAll('[data-tool]').length, 3);
    assert.ok(![...doc.querySelectorAll('*')].some((e) => e.children.length === 0 && /^\s*PRO\s*$/i.test(e.textContent)), 'no Pro badge');
    doc.querySelector('[data-sheet="1"]').click(); await tick();
    assert.match(doc.querySelector('.sheetinfo').textContent, /No difference in this sheet/);
  } finally { s.restore(); }
});

test('webview: folder table and error message', { skip: !JSDOM && 'jsdom not found' }, async () => {
  const dom = new JSDOM('<!DOCTYPE html><body><div id="app"></div></body>', { runScripts: 'outside-only', pretendToBeVisual: true });
  dom.window.acquireVsCodeApi = () => ({ postMessage: () => {} });
  dom.window.eval(script);
  const send = (m) => dom.window.dispatchEvent(new dom.window.MessageEvent('message', { data: m }));
  send({ type: 'folder', a: '/x/a', b: '/x/b', done: true, pro: { show: true, tries: {} }, rows: [{ rel: 'q1.xlsx', a: true, b: true, status: 'different', totals: { changed: 2, added: 1, removed: 0, cells: 3 } }, { rel: '<b>gone</b>.csv', a: true, b: false, status: 'onlyA' }] });
  const doc = dom.window.document;
  assert.match(doc.querySelector('.totals').textContent, /2 files: 1 different, 1 only in the first folder/);
  assert.strictEqual(doc.querySelectorAll('[data-open]').length, 1);
  assert.strictEqual(doc.querySelectorAll('#app b').length, 0);
  send({ type: 'error', message: 'This file has the .xlsx extension but is not a zip archive' });
});
