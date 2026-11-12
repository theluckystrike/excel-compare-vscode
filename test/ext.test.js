'use strict';
// The extension host with a stub vscode: activation, comparisons in a panel, rows on request, values/formulas,
// Git HEAD, hostile files, no network, the Pro tries and the panel, folders, the CSP.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const http = require('http');
const https = require('https');
const net = require('net');
const { execFileSync } = require('child_process');
const { book, tmpdir } = require('./support/books');
const { loadExtension } = require('./support/vscode-stub');
const { buildHtml } = require('../src/html');

const dir = tmpdir('ext');
test.after(() => fs.rmSync(dir, { recursive: true, force: true }));
const file = (name, data) => { const p = path.join(dir, name); fs.mkdirSync(path.dirname(p), { recursive: true }); fs.writeFileSync(p, data); return p; };
const A = file('old.xlsx', book([['Stock', [['ID', 'Name', 'Qty'], [1, 'pear', 3], [2, 'fig', 5], [3, 'kiwi', 7]], { formulas: { D2: 'C2*2' } }]]));
const B = file('new.xlsx', book([['Stock', [['ID', 'Name', 'Qty'], [1, 'pear', 4], [3, 'kiwi', 7], [4, 'plum', 9]], { formulas: { D3: 'C3*3' } }]]));
const C = file('sorted.xlsx', book([['Stock', [['ID', 'Name', 'Qty'], [3, 'kiwi', 7], [1, 'pear', 4], [4, 'plum', 9]]]]));
const FAKE = file('fake.xlsx', 'this is not a workbook');

function trapNetwork() {
  const calls = [];
  const saved = { fetch: global.fetch, hr: http.request, hg: http.get, sr: https.request, sg: https.get, nc: net.connect, ncc: net.createConnection };
  global.fetch = async (u) => { calls.push(String(u)); throw new Error('network blocked in test'); };
  http.request = http.get = https.request = https.get = (...a) => { calls.push(String(a[0] && (a[0].href || a[0].host || a[0]))); throw new Error('network blocked in test'); };
  net.connect = net.createConnection = (...a) => { calls.push('net ' + JSON.stringify(a[0])); throw new Error('network blocked in test'); };
  return { calls, restore: () => { global.fetch = saved.fetch; http.request = saved.hr; http.get = saved.hg; https.request = saved.sr; https.get = saved.sg; net.connect = saved.nc; net.createConnection = saved.ncc; } };
}

async function openCompare(s, a, b) {
  const U = s.vscode.Uri;
  await s.vscode.commands.executeCommand('excelCompare.compareSelected', U.file(a), [U.file(a), U.file(b)]);
  const panel = s.panels.at(-1);
  await panel.send({ type: 'ready' });
  return panel;
}
const last = (panel, type) => panel.posted.filter((m) => m.type === type).at(-1);
const tries = (s) => s.state.get('excelCompare.proTrial') || {};

test('activation and a comparison make no network call and show no message; rows come on request', async () => {
  const nw = trapNetwork();
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    for (const c of ['excelCompare.compareSelected', 'excelCompare.compareWith', 'excelCompare.compareFiles', 'excelCompare.compareWithHead', 'excelCompare.openSample', 'excelCompare.exportReport', 'excelCompare.matchByKey', 'excelCompare.compareFolders', 'excelCompare.getPro', 'excelCompare.payWithPayPal', 'excelCompare.enterLicenseKey']) assert.ok(s.commands.has(c), c);
    assert.strictEqual(s.panels.length, 0, 'nothing opens on activation');
    const p = await openCompare(s, A, B);
    const r = last(p, 'result');
    assert.ok(r, JSON.stringify(p.posted.map((m) => m.type)));
    assert.deepStrictEqual([r.totals.changed, r.totals.added, r.totals.removed], [2, 1, 1], 'pear Qty, and kiwi gets the D3 formula cell');
    assert.deepStrictEqual(r.pro.tries, { report: 3, key: 3, folders: 3 });
    assert.strictEqual(r.sheets[0].cols[2].h, 'Qty');
    await p.send({ type: 'rows', sheet: 0, start: 0, end: 50, seq: 1 });
    const rows = last(p, 'rows');
    assert.strictEqual(rows.rows.length, r.sheets[0].n);
    const changed = rows.rows.find((x) => x.k === 1);
    assert.deepStrictEqual(changed.c[2], ['4', 1, '3']);
    await p.send({ type: 'mode', mode: 'formulas' });
    assert.strictEqual(last(p, 'result').mode, 'formulas');
    await p.send({ type: 'rendered' });
    await p.send({ type: 'rows', sheet: 'x', start: 0, end: 1e9 }); // ignored
    assert.deepStrictEqual(nw.calls, []);
    assert.deepStrictEqual([s.log.info.length, s.log.warn.length, s.log.error.length], [0, 0, 0], 'no message');
    assert.deepStrictEqual(s.log.external, []);
  } finally { nw.restore(); s.restore(); }
});

test('the sample comparison opens from the extension folder; a hostile file gives a message in the panel', async () => {
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    await s.vscode.commands.executeCommand('excelCompare.openSample');
    const p = s.panels.at(-1);
    await p.send({ type: 'ready' });
    const r = last(p, 'result');
    assert.ok(r && r.totals.changed === 3 && r.totals.removed === 1 && r.totals.colsAdded === 1 && r.sheetsOnlyB === 1, JSON.stringify(r && r.totals));
    const h = await openCompare(s, FAKE, B);
    const e = last(h, 'error');
    assert.match(e.message, /not a zip archive/);
    assert.strictEqual(s.log.error.length, 0);
  } finally { s.restore(); }
});

test('Compare with Git HEAD: the committed version against the working copy; outside a repository only a status message', async () => {
  const repo = path.join(dir, 'repo');
  fs.mkdirSync(repo);
  const f = path.join(repo, 'plan.xlsx');
  fs.writeFileSync(f, book([['P', [['k', 'v'], ['a', 1], ['b', 2]]]]));
  const git = (...a) => execFileSync('git', ['-c', 'user.name=t', '-c', 'user.email=t@example.invalid', '-c', 'commit.gpgsign=false', ...a], { cwd: repo, stdio: 'pipe' });
  git('init', '-q'); git('add', '.'); git('commit', '-q', '-m', 'one');
  fs.writeFileSync(f, book([['P', [['k', 'v'], ['a', 1], ['b', 3], ['c', 4]]]]));
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    await s.vscode.commands.executeCommand('excelCompare.compareWithHead', s.vscode.Uri.file(f));
    const p = s.panels.at(-1);
    await p.send({ type: 'ready' });
    const r = last(p, 'result');
    assert.strictEqual(r.a.label, 'plan.xlsx (HEAD)');
    assert.deepStrictEqual([r.totals.changed, r.totals.added], [1, 1]);
    const n = s.panels.length;
    await s.vscode.commands.executeCommand('excelCompare.compareWithHead', s.vscode.Uri.file(A));
    assert.strictEqual(s.panels.length, n, 'no panel outside a repository');
    assert.strictEqual(s.log.info.length + s.log.error.length + s.log.warn.length, 0, 'quiet');
    assert.match(s.log.status.at(-1), /not in a Git repository/);
  } finally { s.restore(); }
});

test('Pro export: Escape keeps the try; 3 reports use the 3 tries; the 4th click shows the panel; Sponsors opens tier 679878', async () => {
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    const p = await openCompare(s, A, B);
    s.ui.quick.push(undefined);
    await p.send({ type: 'tool', tool: 'report' });
    s.ui.quick.push((items) => items[0]); s.ui.save.push(undefined);
    await p.send({ type: 'tool', tool: 'report' });
    assert.strictEqual(tries(s).report, undefined, 'Escape uses no try');
    for (const [i, kind] of [[0, 'xlsx'], [1, 'html'], [2, 'csv']]) {
      const out = path.join(dir, `report.${kind}`);
      s.ui.quick.push((items) => items[i]); s.ui.save.push(s.vscode.Uri.file(out));
      await p.send({ type: 'tool', tool: 'report' });
      assert.ok(fs.existsSync(out) && fs.statSync(out).size > 100, kind);
      assert.strictEqual(tries(s).report, i + 1);
    }
    assert.ok(fs.readFileSync(path.join(dir, 'report.xlsx')).subarray(0, 2).toString() === 'PK');
    assert.match(s.log.status.join('\n'), /free try 3 of 3 of "Export a diff report" used/);
    const before = s.log.info.length;
    s.ui.info.push('Get Zovo Pro with GitHub Sponsors');
    await p.send({ type: 'tool', tool: 'report' });
    const panelMsg = s.log.info.slice(before).find((m) => /is part of Zovo Pro/.test(m.text));
    assert.ok(panelMsg, JSON.stringify(s.log.info.slice(before)));
    assert.strictEqual(panelMsg.text, 'Export a diff report is part of Zovo Pro.');
    assert.strictEqual(panelMsg.opts.modal, true);
    assert.ok(panelMsg.buttons.includes('Get Zovo Pro with GitHub Sponsors') && panelMsg.buttons.includes('Pay with PayPal'), panelMsg.buttons.join(' | '));
    assert.match(panelMsg.opts.detail, /Match rows by key columns/);
    assert.ok(s.log.external.some((u) => u.includes('sponsorships?tier_id=679878')), s.log.external.join(' '));
    assert.ok(!fs.existsSync(path.join(dir, 'report4.xlsx')));
  } finally { s.restore(); }
});

test('Pro key match: Escape on the column pick keeps the try; a keyed comparison uses one and marks the moved row', async () => {
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    const p = await openCompare(s, A, C);
    assert.deepStrictEqual([last(p, 'result').totals.changed, last(p, 'result').totals.added], [0, 2], 'by order the reordered row is removed and added');
    s.ui.quick.push(undefined);
    await p.send({ type: 'tool', tool: 'key' });
    assert.strictEqual(tries(s).key, undefined);
    s.ui.quick.push((items) => [items.find((x) => x.description === 'ID')]);
    await p.send({ type: 'tool', tool: 'key', sheet: 0 }); // the sheet on screen: no sheet pick
    const r = last(p, 'result');
    assert.strictEqual(r.keyed, true);
    assert.deepStrictEqual(r.sheets[0].keys, [0]);
    assert.strictEqual(r.sheets[0].counts.moved, 1, JSON.stringify(r.sheets[0].counts));
    assert.strictEqual(tries(s).key, 1);
    await p.send({ type: 'byOrder' });
    assert.strictEqual(last(p, 'result').keyed, false);
  } finally { s.restore(); }
});

test('Pro folders: pairs by name, a summary row for each, any pair opens; one try after the summary', async () => {
  const fa = path.join(dir, 'fa'), fb = path.join(dir, 'fb');
  fs.mkdirSync(path.join(fa, 'sub'), { recursive: true }); fs.mkdirSync(path.join(fb, 'sub'), { recursive: true });
  fs.copyFileSync(A, path.join(fa, 'stock.xlsx')); fs.copyFileSync(B, path.join(fb, 'stock.xlsx'));
  fs.writeFileSync(path.join(fa, 'sub', 'same.csv'), 'a,b\n1,2\n'); fs.writeFileSync(path.join(fb, 'sub', 'same.csv'), 'a,b\n1,2\n');
  fs.writeFileSync(path.join(fa, 'gone.csv'), 'x\n'); fs.writeFileSync(path.join(fb, 'new.tsv'), 'y\n'); fs.writeFileSync(path.join(fb, '~$stock.xlsx'), 'lock');
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    const U = s.vscode.Uri;
    const run = s.vscode.commands.executeCommand('excelCompare.compareFolders', U.file(fa), [U.file(fa), U.file(fb)]);
    for (let i = 0; i < 50 && !s.panels.length; i++) await new Promise((r) => setTimeout(r, 20));
    const p = s.panels.at(-1);
    await p.send({ type: 'ready' });
    await run;
    const f = last(p, 'folder');
    assert.strictEqual(f.done, true);
    const st = Object.fromEntries(f.rows.map((r) => [r.rel, r.status]));
    assert.deepStrictEqual(st, { 'gone.csv': 'onlyA', 'new.tsv': 'onlyB', 'stock.xlsx': 'different', [path.join('sub', 'same.csv')]: 'same' });
    assert.strictEqual(tries(s).folders, 1);
    const i = f.rows.findIndex((r) => r.rel === 'stock.xlsx');
    await p.send({ type: 'openPair', index: i });
    assert.strictEqual(s.panels.at(-1).viewType, 'excelCompare.diff');
  } finally { s.restore(); }
});

test('Pro tools on comparisons with no change use no try: report (same file, blank workbooks, empty CSVs), key match, folders of identical pairs', async () => {
  const SAME = file('same-copy.xlsx', fs.readFileSync(A));
  const BL1 = file('blank1.xlsx', book([['Sheet', []]])), BL2 = file('blank2.xlsx', book([['Sheet', []]]));
  const E1 = file('empty1.csv', ''), E2 = file('empty2.csv', '');
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    for (const [a, b] of [[A, SAME], [BL1, BL2], [E1, E2]]) {
      const p = await openCompare(s, a, b);
      const out = path.join(dir, `none-${path.basename(a)}.csv`);
      s.ui.quick.push((items) => items[2]); s.ui.save.push(s.vscode.Uri.file(out));
      const n = s.log.info.length;
      await p.send({ type: 'tool', tool: 'report' });
      assert.ok(!fs.existsSync(out), `no report file for ${path.basename(a)}`);
      assert.ok(s.log.info.slice(n).some((m) => m.text === 'The two files are the same (values compared), so there is no change to report. No free try was used.'), JSON.stringify(s.log.info.slice(n)));
      assert.strictEqual(tries(s).report, undefined, 'no report try');
      s.ui.quick.length = 0; s.ui.save.length = 0;
    }
    // key match on two equal sheets: the comparison re-runs by key, nothing differs, no try
    const p = await openCompare(s, A, file('same-copy2.xlsx', fs.readFileSync(A)));
    s.ui.quick.push((items) => [items.find((x) => x.description === 'ID')]);
    const n = s.log.info.length;
    await p.send({ type: 'tool', tool: 'key', sheet: 0 });
    assert.strictEqual(last(p, 'result').keyed, true);
    assert.ok(s.log.info.slice(n).some((m) => /^Matched by key, the rows of "Stock" are the same in both files\. No free try was used\.$/.test(m.text)), JSON.stringify(s.log.info.slice(n)));
    assert.strictEqual(tries(s).key, undefined, 'no key try');
    // folders where every pair is the same and no file is on one side only
    const ga = path.join(dir, 'ga'), gb = path.join(dir, 'gb');
    fs.mkdirSync(ga, { recursive: true }); fs.mkdirSync(gb, { recursive: true });
    fs.copyFileSync(A, path.join(ga, 'stock.xlsx')); fs.copyFileSync(A, path.join(gb, 'stock.xlsx'));
    fs.writeFileSync(path.join(ga, 'x.csv'), 'a,b\n1,2\n'); fs.writeFileSync(path.join(gb, 'x.csv'), 'a,b\n1,2\n');
    const U = s.vscode.Uri;
    const before = s.panels.length;
    const run = s.vscode.commands.executeCommand('excelCompare.compareFolders', U.file(ga), [U.file(ga), U.file(gb)]);
    for (let i = 0; i < 50 && s.panels.length === before; i++) await new Promise((r) => setTimeout(r, 20));
    const fp = s.panels.at(-1);
    const m = s.log.info.length;
    await fp.send({ type: 'ready' });
    await run;
    assert.deepStrictEqual(last(fp, 'folder').rows.map((r) => r.status), ['same', 'same']);
    assert.ok(s.log.info.slice(m).some((x) => x.text === 'All 2 file pairs are the same (values compared). No free try was used.'), JSON.stringify(s.log.info.slice(m)));
    assert.strictEqual(tries(s).folders, undefined, 'no folders try');
  } finally { s.restore(); }
});

test('showProTools off: the panel gets show=false; the CSP has a nonce and no unsafe-inline or remote origin', async () => {
  const s = loadExtension();
  try {
    s.config.showProTools = false;
    s.ext.activate(s.context);
    const p = await openCompare(s, A, B);
    assert.strictEqual(last(p, 'result').pro.show, false);
    const html = p.webview.html;
    assert.match(html, /script-src 'nonce-[A-Za-z0-9_-]+'/);
    assert.ok(!/unsafe-inline|unsafe-eval|https?:\/\//.test(html.match(/content="([^"]+)"/)[1]));
  } finally { s.restore(); }
  const h = buildHtml({ cspSource: 'vscode-resource:', scriptUri: 'a"b', styleUri: 's', title: '<x>' });
  assert.ok(h.includes('a&quot;b') && h.includes('&lt;x&gt;'));
});

test('REACH1: the Pro tools row is in the comparison page head with the tries left, its links are allowed, a close hides it, toolbar clicks are tagged', async () => {
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    const p = await openCompare(s, A, B);
    const html = p.webview.html;
    assert.match(html, /<script nonce="[A-Za-z0-9_-]+">\(function \(\) \{\n  var CLOSED_KEY/);
    assert.ok(html.indexOf('zovo-reach1') < html.indexOf('main.js'), 'the row script runs before main.js (shared acquireVsCodeApi)');
    for (const t of ['3 free tries', 'Export a Diff Report', 'Match Rows by Key Columns', 'Compare Two Folders of Workbooks', 'Get Zovo Pro']) assert.ok(html.includes(t), t);
    assert.match(html, /"header\.top"/);
    const allowed = p.options.enableCommandUris;
    for (const c of ['excelCompare.exportReport', 'excelCompare.matchByKey', 'excelCompare.compareFolders', 'excelCompare.getPro']) assert.ok(allowed.includes(c), c);
    // a free try used elsewhere updates the row in open panels
    await s.state.set('excelCompare.proTrial', { report: 2 });
    s.ext._refreshPro();
    const row = p.posted.filter((m) => m.type === 'zovoReach1').at(-1);
    assert.ok(row && /1 try left|1 tries left/.test(row.html), row && row.html);
    // with no tries left, the toolbar's Pro button opens the panel; Pay by card carries <install id>_toolbar
    await s.state.set('excelCompare.proTrial', { report: 3 });
    s.ui.info.push('Pay by card (USD 19)');
    await p.send({ type: 'tool', tool: 'report' });
    assert.ok(s.log.external.some((u) => /license\.zovo\.one\/buy\/zovo-pro\?client_reference_id=[\w-]+_toolbar/.test(u)), s.log.external.join(' '));
    await p.send({ type: 'zovoReach1Closed' });
    assert.strictEqual(p.posted.filter((m) => m.type === 'zovoReach1').at(-1).html, '', 'closed for the session');
  } finally { s.restore(); }
});

test('REACH1 and listing: one Explorer entry for the folder tool, name, description and tags', () => {
  const pkg = require('../package.json');
  const ex = pkg.contributes.menus['explorer/context'];
  assert.deepStrictEqual(ex.filter((m) => m.command.startsWith('excelCompare.zovoPro.')), [{ command: 'excelCompare.zovoPro.folders', when: 'explorerResourceIsFolder && config.excelCompare.showProTools', group: 'zovoPro@1' }]);
  assert.ok(!ex.some((m) => m.command === 'excelCompare.compareFolders'), 'one Explorer entry per tool');
  assert.ok(pkg.displayName.length <= 30 && pkg.displayName.startsWith('Excel Compare') && pkg.description.toLowerCase().startsWith('excel compare'));
  assert.deepStrictEqual(pkg.keywords, ['excel compare', 'excel diff', 'xlsx diff', 'compare excel', 'csv diff']);
});
