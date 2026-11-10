'use strict';
// Excel Compare & XLSX Diff: compares two workbooks or CSV files in a webview grid. Reading and comparing run in a
// worker thread (runner.js, dist/worker.js) with a time and memory cap, so a hostile file cannot hang or crash the
// extension host. The host keeps the result and sends the grid the rows it shows. No network.
const vscode = require('vscode');
const fs = require('fs');
const path = require('path');
const { runJob } = require('./runner');
const { readFromHead } = require('./git');
const { rowCells, colLetter } = require('./cells');
const { buildHtml } = require('./html');
const { createPro } = require('./pro');
const { createRatingPrompt } = require('./rate');
const { registerSupport, commandUris, showLinks } = require('./support');

const EXT_ID = 'zovo.zovo-excel-compare';
const PREFIX = 'excelCompare';
const VIEW_TYPE = 'excelCompare.diff';
const EXTENSIONS = ['xlsx', 'xlsm', 'xls', 'ods', 'csv', 'tsv'];
const EXT_RE = /\.(xlsx|xlsm|xls|ods|csv|tsv)$/i;
const PICK_FILTERS = { 'Excel and CSV files': EXTENSIONS, 'All files': ['*'] };
const MAX_REMOTE_BYTES = 256 * 1024 * 1024;
const MAX_CELL_TEXT = 2000;
const MAX_WINDOW = 2000;
const SKIP_DIRS = new Set(['node_modules', '.git', '.svn', '.hg', '__pycache__', '.venv', '.cache']);

let rating = { success() {}, failure() {} };
let pro;
let ctx;
const sessions = new Map(); // key -> Compare
let active = null;

function setting(key, dflt) { try { return vscode.workspace.getConfiguration(PREFIX).get(key, dflt); } catch { return dflt; } }
function limits() {
  const mb = Math.max(1, Math.min(2000, Number(setting('maxFileSizeMB', 200)) || 200));
  const rows = Math.max(1000, Math.min(1000000, Number(setting('maxRows', 300000)) || 300000));
  return { maxBytes: mb * 1024 * 1024, maxRows: rows };
}
const baseName = (uri) => path.posix.basename(uri.path) || uri.toString();
function proState() { return { show: setting('showProTools', true) !== false, licensed: !!(pro && pro.licensed), tries: pro ? pro.tries() : {} }; }

/** side: { uri, name, label, bytes? } -> the job part the worker reads. */
async function jobSide(side) {
  if (side.bytes) { const copy = new Uint8Array(side.bytes); return { part: { bytes: copy.buffer, name: side.name, label: side.label }, transfer: [copy.buffer] }; }
  if (side.uri.scheme === 'file') return { part: { file: side.uri.fsPath, name: side.name, label: side.label }, transfer: [] };
  const st = await vscode.workspace.fs.stat(side.uri);
  if (st.size > MAX_REMOTE_BYTES) throw Object.assign(new Error(`${side.name} is ${Math.round(st.size / 1048576)} MB on a remote or virtual file system; files there are read up to 256 MB.`), { code: 'LIMIT' });
  const data = await vscode.workspace.fs.readFile(side.uri);
  const copy = new Uint8Array(data);
  return { part: { bytes: copy.buffer, name: side.name, label: side.label }, transfer: [copy.buffer] };
}

function fileSide(uri) { const n = baseName(uri); return { uri, name: n, label: n }; }
const sideKey = (s) => (s.bytes ? `HEAD:${s.uri.toString()}` : s.uri.toString());

class Compare {
  constructor(left, right) {
    this.left = left; this.right = right; this.mode = 'values'; this.keys = null; this.data = null; this.counted = false; this.seq = 0;
    this.key = `${sideKey(left)}|${sideKey(right)}`;
  }
  get title() { return `${this.left.label} ↔ ${this.right.label}`; }
  post(msg) { return this.panel ? this.panel.webview.postMessage(msg).then(undefined, () => false) : Promise.resolve(false); }
  open() {
    const media = vscode.Uri.joinPath(ctx.extensionUri, 'media');
    const panel = vscode.window.createWebviewPanel(VIEW_TYPE, this.title, vscode.ViewColumn.Active, {
      enableScripts: true, enableCommandUris: commandUris(PREFIX), localResourceRoots: [media], retainContextWhenHidden: true,
    });
    panel.iconPath = vscode.Uri.joinPath(ctx.extensionUri, 'icon.png');
    panel.webview.html = buildHtml({
      cspSource: panel.webview.cspSource,
      scriptUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'main.js')).toString(),
      styleUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'main.css')).toString(),
      title: this.title,
      showSupportLinks: showLinks(vscode, PREFIX),
    });
    this.panel = panel;
    active = this;
    panel.onDidChangeViewState(() => { if (panel.active) active = this; });
    panel.onDidDispose(() => { sessions.delete(this.key); if (active === this) active = null; this.panel = null; this.data = null; });
    panel.webview.onDidReceiveMessage((m) => this.onMessage(m).then(undefined, (e) => this.post({ type: 'notice', text: String((e && e.message) || e) })));
  }
  async run() {
    const seq = ++this.seq;
    this.post({ type: 'busy', text: 'Reading the files...' });
    try {
      const a = await jobSide(this.left), b = await jobSide(this.right);
      const job = { op: 'compare', a: a.part, b: b.part, opts: { mode: this.mode, keys: this.keys }, limits: limits() };
      const data = await runJob(job, {
        transfer: [...a.transfer, ...b.transfer],
        onProgress: (p) => this.post({ type: 'busy', text: p.step === 'read' ? `Reading the ${p.which === 1 ? 'first' : 'second'} file...` : 'Comparing...' }),
      });
      if (seq !== this.seq || !this.panel) return false;
      this.data = data;
      await this.post({ type: 'result', ...this.meta(), pro: proState() });
      return true;
    } catch (e) {
      if (seq !== this.seq) return false;
      this.data = null;
      if (!['LIMIT', 'KEY', 'ENOENT'].includes(e.code)) rating.failure();
      await this.post({ type: 'error', message: (e && e.message) || String(e), pro: proState() });
      return false;
    }
  }
  meta() {
    const { A, B, res } = this.data;
    const sheets = res.sheets.map((d, i) => {
      const sa = d.ia >= 0 ? A.sheets[d.ia] : null, sb = d.ib >= 0 ? B.sheets[d.ib] : null;
      const cols = d.cols.map((c) => {
        const h = c.b >= 0 && sb && sb.rows[0] ? sb.rows[0][c.b] : c.a >= 0 && sa && sa.rows[0] ? sa.rows[0][c.a] : '';
        const l = c.b >= 0 ? colLetter(c.b) + (c.a >= 0 && c.a !== c.b ? ` (${colLetter(c.a)})` : '') : colLetter(c.a);
        return { l, h: String(h == null ? '' : h).slice(0, 80), s: c.a < 0 ? 'added' : c.b < 0 ? 'removed' : '' };
      });
      return { i, nameA: d.nameA, nameB: d.nameB, status: d.status, counts: d.counts, notes: d.notes, cols, n: d.k.length, chg: Array.from(d.chg), keys: d.keys || null, byHeader: d.byHeader, formulas: !!((sa && sa.f) || (sb && sb.f)) };
    });
    return {
      a: { name: this.left.name, label: this.left.label, file: !this.left.bytes }, b: { name: this.right.name, label: this.right.label, file: !this.right.bytes },
      mode: res.mode, totals: res.totals, sheetsOnlyA: res.sheetsOnlyA, sheetsOnlyB: res.sheetsOnlyB, sheets,
      notes: [...A.notes.map((n) => `First file: ${n}`), ...B.notes.map((n) => `Second file: ${n}`)], keyed: !!this.keys,
    };
  }
  rows(sheet, start, end) {
    if (!this.data) return null;
    const { A, B, res } = this.data;
    const d = res.sheets[sheet];
    if (!d) return null;
    const sa = d.ia >= 0 ? A.sheets[d.ia] : null, sb = d.ib >= 0 ? B.sheets[d.ib] : null;
    start = Math.max(0, Math.min(d.k.length, start)); end = Math.max(start, Math.min(d.k.length, end, start + MAX_WINDOW));
    const out = [];
    const cut = (s) => (s && s.length > MAX_CELL_TEXT ? s.slice(0, MAX_CELL_TEXT) + '…' : s);
    for (let i = start; i < end; i++) {
      const cells = rowCells(d, sa, sb, i, res.mode).map((c) => (c.s === 1 ? [cut(c.t), 1, cut(c.o)] : [cut(c.t), c.s]));
      out.push({ i, k: d.k[i], a: d.ra[i], b: d.rb[i], c: cells });
    }
    return out;
  }
  async onMessage(msg) {
    if (!msg || typeof msg.type !== 'string') return undefined;
    switch (msg.type) {
      case 'ready': return this.run();
      case 'rows': {
        const sheet = Number(msg.sheet), start = Number(msg.start), end = Number(msg.end);
        if (![sheet, start, end].every(Number.isSafeInteger)) return undefined;
        const rows = this.rows(sheet, start, end);
        return rows ? this.post({ type: 'rows', sheet, start, rows, seq: msg.seq }) : undefined;
      }
      case 'rendered':
        if (!this.counted && this.data) { this.counted = true; rating.success(this.right.uri); }
        return undefined;
      case 'mode':
        if (msg.mode !== 'values' && msg.mode !== 'formulas') return undefined;
        this.mode = msg.mode; return this.run();
      case 'swap': { const l = this.left; this.left = this.right; this.right = l; if (this.panel) this.panel.title = this.title; this.keys = null; return this.run(); }
      case 'byOrder': this.keys = null; return this.run();
      case 'openFile': {
        const side = msg.which === 'a' ? this.left : msg.which === 'b' ? this.right : null;
        if (!side || side.bytes) return undefined;
        return vscode.commands.executeCommand('revealInExplorer', side.uri).then(undefined, () => undefined);
      }
      case 'tool': return runTool(String(msg.tool), this);
      case 'copy': return vscode.env.clipboard.writeText(String(msg.text || '').slice(0, 1 << 20)).then(() => vscode.window.setStatusBarMessage('Copied.', 3000));
      case 'settings': return vscode.commands.executeCommand('workbench.action.openSettings', `@ext:${EXT_ID}`);
      default: return undefined;
    }
  }
}

async function compare(left, right) {
  const c = new Compare(left, right);
  const old = sessions.get(c.key);
  if (old && old.panel) { old.panel.reveal(); active = old; return old.run(); }
  sessions.set(c.key, c);
  c.open();
  return c;
}

// ---------------- commands ----------------

async function pickFile(title, near) {
  const p = await vscode.window.showOpenDialog({ canSelectMany: false, openLabel: 'Compare', title, filters: PICK_FILTERS, defaultUri: near && near.scheme === 'file' ? vscode.Uri.file(path.dirname(near.fsPath)) : undefined });
  return p && p[0] ? p[0] : null;
}

function activeFileUri() {
  try {
    const tab = vscode.window.tabGroups.activeTabGroup.activeTab;
    const u = tab && tab.input && (tab.input.uri || (tab.input.modified && tab.input.modified));
    if (u && EXT_RE.test(u.path)) return u;
  } catch { /* older API */ }
  const ed = vscode.window.activeTextEditor;
  if (ed && EXT_RE.test(ed.document.uri.path)) return ed.document.uri;
  return null;
}

async function compareSelected(uri, uris) {
  const list = Array.isArray(uris) && uris.length ? uris : uri ? [uri] : [];
  if (list.length === 2) {
    let dirs = 0;
    for (const u of list) { try { if ((await vscode.workspace.fs.stat(u)).type & vscode.FileType.Directory) dirs++; } catch { /* missing */ } }
    if (dirs === 2) return compareFolders(list[0], list[1]);
    return compare(fileSide(list[0]), fileSide(list[1]));
  }
  if (list.length === 1) return compareWith(list[0]);
  if (list.length > 2) { vscode.window.showInformationMessage(`Select exactly two files to compare (${list.length} are selected).`); return undefined; }
  return compareFiles();
}

async function compareWith(uri) {
  const first = uri && uri.scheme ? uri : activeFileUri() || (await pickFile('Pick the first Excel or CSV file'));
  if (!first) return undefined;
  const second = await pickFile(`Compare ${baseName(first)} with...`, first);
  if (!second) return undefined;
  return compare(fileSide(first), fileSide(second));
}

async function compareFiles() {
  const first = await pickFile('Compare: pick the first (older) file');
  if (!first) return undefined;
  const second = await pickFile(`Compare ${baseName(first)} with... (the newer file)`, first);
  if (!second) return undefined;
  return compare(fileSide(first), fileSide(second));
}

async function compareWithHead(uri) {
  const target = uri && uri.scheme ? uri : activeFileUri() || (await pickFile('Compare with Git HEAD: pick a workbook in a Git repository'));
  if (!target) return undefined;
  if (target.scheme !== 'file') { vscode.window.setStatusBarMessage('Compare with Git HEAD works with files on this computer.', 6000); return undefined; }
  const r = await readFromHead(target.fsPath, limits().maxBytes);
  if (!r.ok) {
    if (r.quiet) vscode.window.setStatusBarMessage(`Excel Compare: ${r.message}`, 8000);
    else vscode.window.showInformationMessage(r.message);
    return undefined;
  }
  const name = baseName(target);
  return compare({ uri: target, name, label: `${name} (HEAD)`, bytes: new Uint8Array(r.bytes) }, { uri: target, name, label: `${name} (working copy)` });
}

// ---------------- Pro tools ----------------

async function sessionOrPick(title) {
  if (active && active.data) return active;
  const open = [...sessions.values()].filter((s) => s.data);
  if (open.length === 1) return open[0];
  if (open.length > 1) {
    const p = await vscode.window.showQuickPick(open.map((s) => ({ label: s.title, s })), { title });
    return p ? p.s : null;
  }
  vscode.window.showInformationMessage('Open a comparison first (Compare Selected Excel Files, or Compare with... on a file).');
  return null;
}

async function runTool(tool, from) {
  if (tool === 'report') return exportReport(from);
  if (tool === 'key') return matchByKey(from);
  if (tool === 'folders') return compareFolders();
  return undefined;
}

async function exportReport(from) {
  if (!(await pro.allowed('report'))) return;
  const s = from && from.data ? from : await sessionOrPick('Export a report of which comparison?');
  if (!s || !s.data) return;
  const kind = await vscode.window.showQuickPick([
    { label: 'XLSX workbook', description: 'changed, added and removed cells filled in colour, with Summary and Changes sheets', kind: 'xlsx' },
    { label: 'HTML page', description: 'one file you can open in a browser or attach', kind: 'html' },
    { label: 'CSV list', description: 'one line per change: sheet, cell, before, after', kind: 'csv' },
  ], { title: 'Export a diff report' });
  if (!kind) return; // Escape: no try used
  const stem = `${s.right.name.replace(/\.[^.]+$/, '')}-diff`;
  const dir = s.right.uri.scheme === 'file' ? path.dirname(s.right.uri.fsPath) : null;
  const target = await vscode.window.showSaveDialog({ defaultUri: dir ? vscode.Uri.file(path.join(dir, `${stem}.${kind.kind}`)) : undefined, filters: { [kind.label]: [kind.kind] }, title: `Save the diff report (${kind.label})` });
  if (!target) return; // cancelled: no try used
  const a = await jobSide(s.left), b = await jobSide(s.right);
  let r;
  try {
    r = await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: 'Writing the diff report' }, () => runJob({ op: 'report', kind: kind.kind, a: a.part, b: b.part, opts: { mode: s.mode, keys: s.keys }, limits: limits() }, { transfer: [...a.transfer, ...b.transfer] }));
  } catch (e) { vscode.window.showWarningMessage(`The report could not be written: ${e.message}`); return; }
  const bytes = r.data ? Buffer.from(r.data.buffer, r.data.byteOffset, r.data.byteLength) : Buffer.from(r.text, 'utf8');
  await vscode.workspace.fs.writeFile(target, bytes);
  await pro.produced('report');
  refreshPro();
  if (kind.kind === 'csv') await vscode.window.showTextDocument(target, { preview: false, viewColumn: vscode.ViewColumn.Beside });
  else vscode.window.setStatusBarMessage(`Excel Compare: saved ${path.basename(target.path)} (${bytes.length.toLocaleString('en-US')} bytes).`, 10000);
}

async function matchByKey(from) {
  if (!(await pro.allowed('key'))) return;
  const s = from && from.data ? from : await sessionOrPick('Match rows by key in which comparison?');
  if (!s || !s.data) return;
  const { A, res } = s.data;
  const both = res.sheets.filter((d) => d.status === 'both');
  if (!both.length) { vscode.window.showInformationMessage('No sheet is in both files, so there is nothing to match.'); return; }
  let d = both[0];
  if (both.length > 1) {
    const p = await vscode.window.showQuickPick(both.map((x) => ({ label: x.nameB, description: x.nameA !== x.nameB ? `first file: ${x.nameA}` : '', d: x })), { title: 'Match rows by key: which sheet?' });
    if (!p) return;
    d = p.d;
  }
  const sa = A.sheets[d.ia];
  const items = d.cols.filter((c) => c.a >= 0 && c.b >= 0).slice(0, 500).map((c) => ({ label: colLetter(c.a), description: String((sa.rows[0] && sa.rows[0][c.a]) || '').slice(0, 80), col: c.a, picked: !!(s.keys && s.keys[d.ia] && s.keys[d.ia].includes(c.a)) }));
  const picks = await vscode.window.showQuickPick(items, { canPickMany: true, title: `Key columns for "${d.nameB}" (for example an ID column)`, placeHolder: 'Rows with the same values in these columns are matched; the first row is the header' });
  if (!picks || !picks.length) return; // Escape: no try used
  s.keys = { ...(s.keys || {}), [d.ia]: picks.map((p) => p.col) };
  const ok = await s.run();
  if (ok) { await pro.produced('key'); refreshPro(); s.panel && s.panel.reveal(); }
}

/** Workbook and CSV files under a folder: { rel -> fsPath }, at most 2,000 files and 20,000 entries looked at. */
async function listBooks(root) {
  const out = new Map();
  let seen = 0;
  const stack = [[root, 0]];
  while (stack.length && seen < 20000 && out.size < 2000) {
    const [dir, depth] = stack.pop();
    let entries;
    try { entries = await fs.promises.readdir(dir, { withFileTypes: true }); } catch { continue; }
    for (const e of entries) {
      if (++seen > 20000) break;
      const p = path.join(dir, e.name);
      if (e.isDirectory()) { if (!SKIP_DIRS.has(e.name) && depth < 8) stack.push([p, depth + 1]); continue; }
      if (e.isFile() && EXT_RE.test(e.name) && !e.name.startsWith('~$')) out.set(path.relative(root, p), p);
    }
  }
  return out;
}

function pairFolders(mapA, mapB) {
  const pairs = [];
  const usedB = new Set();
  const lowerB = new Map([...mapB.keys()].map((k) => [k.toLowerCase(), k]));
  const leftA = [];
  for (const [rel, p] of mapA) {
    const kb = mapB.has(rel) ? rel : lowerB.get(rel.toLowerCase());
    if (kb && !usedB.has(kb)) { usedB.add(kb); pairs.push({ rel, a: p, b: mapB.get(kb) }); } else leftA.push([rel, p]);
  }
  // then by file name alone, when it is unique on both sides
  const byName = (entries) => { const m = new Map(); for (const [rel, p] of entries) { const n = path.basename(rel).toLowerCase(); m.set(n, m.has(n) ? null : [rel, p]); } return m; };
  const nb = byName([...mapB].filter(([k]) => !usedB.has(k)));
  for (const [rel, p] of leftA) {
    const hit = nb.get(path.basename(rel).toLowerCase());
    if (hit) { usedB.add(hit[0]); nb.delete(path.basename(rel).toLowerCase()); pairs.push({ rel, a: p, b: hit[1] }); } else pairs.push({ rel, a: p, b: null });
  }
  for (const [rel, p] of mapB) if (!usedB.has(rel)) pairs.push({ rel, a: null, b: p });
  pairs.sort((x, y) => x.rel.localeCompare(y.rel));
  return pairs;
}

let folderPanels = 0;
async function compareFolders(dirA, dirB) {
  if (!(await pro.allowed('folders'))) return;
  const pickDir = async (title) => { const p = await vscode.window.showOpenDialog({ canSelectFiles: false, canSelectFolders: true, canSelectMany: false, openLabel: 'Select folder', title }); return p && p[0] ? p[0] : null; };
  const a = dirA && dirA.fsPath ? dirA : await pickDir('Compare folders: the first (older) folder');
  if (!a) return;
  const b = dirB && dirB.fsPath ? dirB : await pickDir(`Compare ${baseName(a)} with... (the newer folder)`);
  if (!b) return; // cancelled: no try used
  const pairs = pairFolders(await listBooks(a.fsPath), await listBooks(b.fsPath));
  const media = vscode.Uri.joinPath(ctx.extensionUri, 'media');
  const panel = vscode.window.createWebviewPanel(VIEW_TYPE + '.folders', `${baseName(a)} ↔ ${baseName(b)} (folders)`, vscode.ViewColumn.Active, { enableScripts: true, enableCommandUris: commandUris(PREFIX), localResourceRoots: [media], retainContextWhenHidden: true });
  folderPanels++;
  panel.iconPath = vscode.Uri.joinPath(ctx.extensionUri, 'icon.png');
  panel.webview.html = buildHtml({
    cspSource: panel.webview.cspSource, title: panel.title, showSupportLinks: showLinks(vscode, PREFIX),
    scriptUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'main.js')).toString(),
    styleUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'main.css')).toString(),
  });
  let closed = false, counted = false;
  const rows = pairs.map((p) => ({ rel: p.rel, a: !!p.a, b: !!p.b, status: p.a && p.b ? 'waiting' : p.a ? 'onlyA' : 'onlyB' }));
  const send = (done) => panel.webview.postMessage({ type: 'folder', a: a.fsPath, b: b.fsPath, rows, done, pro: proState() }).then(undefined, () => false);
  const ready = new Promise((resolve) => {
    panel.webview.onDidReceiveMessage(async (m) => {
      if (!m) return;
      if (m.type === 'ready') resolve();
      if (m.type === 'openPair') { const p = pairs[Number(m.index)]; if (p && p.a && p.b) await compare(fileSide(vscode.Uri.file(p.a)), fileSide(vscode.Uri.file(p.b))); }
      if (m.type === 'tool') await runTool(String(m.tool), null);
    });
  });
  panel.onDidDispose(() => { closed = true; });
  await ready;
  await send(false);
  let compared = 0;
  await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: 'Comparing folders', cancellable: true }, async (progress, token) => {
    for (let i = 0; i < pairs.length && !closed && !token.isCancellationRequested; i++) {
      const p = pairs[i];
      if (!p.a || !p.b) continue;
      progress.report({ message: `${i + 1} of ${pairs.length}: ${p.rel}` });
      try {
        const r = await runJob({ op: 'summary', a: { file: p.a, name: path.basename(p.a) }, b: { file: p.b, name: path.basename(p.b) }, opts: {}, limits: limits() }, { idleMs: 60000 });
        Object.assign(rows[i], { status: r.same ? 'same' : 'different', totals: r.totals, sheetsOnlyA: r.sheetsOnlyA, sheetsOnlyB: r.sheetsOnlyB });
        compared++;
      } catch (e) { Object.assign(rows[i], { status: 'error', error: String(e.message || e).slice(0, 200) }); }
      if (i % 5 === 0) await send(false);
    }
  });
  if (closed) return;
  await send(true);
  if (compared && !counted) { counted = true; await pro.produced('folders'); refreshPro(); }
}

function refreshPro() { for (const s of sessions.values()) s.post({ type: 'pro', pro: proState() }); }

// ---------------- activation ----------------

function activate(context) {
  ctx = context;
  rating = createRatingPrompt({ vscode, context, extensionId: EXT_ID, section: PREFIX, name: 'Excel Compare' });
  registerSupport({ vscode, context, prefix: PREFIX, selfId: EXT_ID });
  pro = createPro({ vscode, context });
  context.subscriptions.push(pro.onDidChange(() => refreshPro()));
  const reg = (id, fn) => context.subscriptions.push(vscode.commands.registerCommand(id, fn));
  reg(`${PREFIX}.compareSelected`, compareSelected);
  reg(`${PREFIX}.compareWith`, compareWith);
  reg(`${PREFIX}.compareFiles`, compareFiles);
  reg(`${PREFIX}.compareWithHead`, compareWithHead);
  reg(`${PREFIX}.openSample`, () => {
    const s = (n) => vscode.Uri.joinPath(context.extensionUri, 'samples', n);
    return compare(fileSide(s('budget-march.xlsx')), fileSide(s('budget-april.xlsx')));
  });
  reg(`${PREFIX}.openWalkthrough`, () => vscode.commands.executeCommand('workbench.action.openWalkthrough', `${EXT_ID}#${PREFIX}.start`, false));
  reg(`${PREFIX}.exportReport`, () => exportReport(active));
  reg(`${PREFIX}.matchByKey`, () => matchByKey(active));
  reg(`${PREFIX}.compareFolders`, (uri, uris) => (Array.isArray(uris) && uris.length === 2 ? compareFolders(uris[0], uris[1]) : compareFolders()));
  context.subscriptions.push(vscode.window.registerUriHandler({ handleUri: (uri) => pro.handleUri(uri) }));
  context.subscriptions.push(vscode.workspace.onDidChangeConfiguration((e) => { if (e.affectsConfiguration(`${PREFIX}.showProTools`)) refreshPro(); }));
}

function deactivate() {}

module.exports = { activate, deactivate, pairFolders, listBooks, Compare, _sessions: sessions, _state: () => ({ active, folderPanels }) };
