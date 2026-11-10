'use strict';
// Zovo Pro for Excel Compare, on top of the shared license module (src/license, copied unchanged from
// factory/sprint/license). Three Pro tools, 3 free tries each on this computer. A try counts only after the tool
// produced its output (a saved report, a key-matched comparison on screen, a folder summary); Escape or Cancel never
// uses one. Nothing about Pro shows on open or activation: the panel appears only after a click on a Pro tool with
// no tries left.
const { createLicense } = require('./license');

const SKU = 'excel-compare-pro';
const TRIAL_USES = 3;
const TRIAL_KEY = 'excelCompare.proTrial';

const FEATURES = {
  report: {
    title: 'Export a diff report',
    detail: 'Saves the comparison as an XLSX workbook with the changed, added and removed cells filled in colour (plus a Summary and a Changes list), as one HTML page, or as a CSV list of every change.',
  },
  key: {
    title: 'Match rows by key columns',
    detail: 'Matches rows by the values of key columns such as an ID instead of by their order, so sorted or reordered sheets compare cleanly, and marks the rows that moved.',
  },
  folders: {
    title: 'Compare two folders of workbooks',
    detail: 'Pairs the Excel and CSV files of two folders by name, compares every pair and shows a summary table; any pair opens in the full comparison.',
  },
};
const TOOLS = Object.values(FEATURES).map((f) => f.title);

function createPro({ vscode, context, license, prefix = 'excelCompare', productName = 'Excel Compare' }) {
  const emitter = new vscode.EventEmitter();
  const lic = license || createLicense({ vscode, context, sku: SKU, name: 'Zovo', product: productName, commandPrefix: prefix, proTools: TOOLS });
  if (lic.onDidChange) lic.onDidChange(() => emitter.fire(lic.isPro(SKU)));

  const used = (f) => Number((context.globalState.get(TRIAL_KEY) || {})[f]) || 0;
  const left = (f) => (FEATURES[f] ? Math.max(0, TRIAL_USES - used(f)) : 0);
  const licensed = () => lic.isPro(SKU);

  /** Before a Pro tool runs: true when it may run (licensed, or a free try left). Otherwise the one panel. */
  async function allowed(f) {
    if (licensed()) { if (lic.refresh) lic.refresh(); return true; }
    if (left(f) > 0) return true;
    const feat = FEATURES[f] || { title: 'This tool', detail: '' };
    return !!(await lic.requirePro(feat.title, { detail: feat.detail, triesUsed: used(f), tools: TOOLS }));
  }

  /** After the tool produced its output: counts one free try (nothing when licensed). */
  async function produced(f) {
    if (licensed() || !FEATURES[f]) return;
    const m = { ...(context.globalState.get(TRIAL_KEY) || {}) };
    m[f] = used(f) + 1;
    await context.globalState.update(TRIAL_KEY, m);
    const l = left(f);
    vscode.window.setStatusBarMessage(`Zovo Pro: free try ${used(f)} of ${TRIAL_USES} of "${FEATURES[f].title}" used${l ? `, ${l} left` : ''} on this computer.`, 8000);
    emitter.fire(false);
  }

  context.subscriptions.push(emitter);
  return {
    SKU, FEATURES, TRIAL_USES,
    get licensed() { return licensed(); },
    ready: lic.ready,
    allowed, produced, left,
    tries: () => Object.fromEntries(Object.keys(FEATURES).map((f) => [f, left(f)])),
    handleUri: (uri) => (lic.handleUri ? lic.handleUri(uri) : undefined),
    onDidChange: emitter.event,
    license: lic,
  };
}

module.exports = { createPro, FEATURES, SKU, TRIAL_USES, TRIAL_KEY, TOOLS };
