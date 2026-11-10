'use strict';
// Compares two books. Sheets are matched by name (exact, then ignoring case), then by position. Columns are matched
// by the header row when both sheets have one, so an inserted column does not change every cell. Rows are aligned
// on the text of the shared columns: common start and end, then rows that occur once on each side as anchors
// (patience diff), then an LCS table for the small gaps that remain. Inserted rows therefore do not shift the rest.
// Removed and added rows next to each other that share most cells are shown as one changed row.
// Pro: rows matched by key columns instead of order, with moved rows.
const { cellText, KIND } = require('../cells');

const DP_MAX = 4000000; // cells of one LCS table (n x m), so one table is at most 8 MB
const MAX_DEPTH = 60;
const SIMILAR = 0.5;

/** Aligns two int sequences. Returns { pa, pb }: pairs in order, -1 for a gap. Equal ids are paired; a pair of
 *  different ids can come from the positional fallback (no anchor and too big for the LCS table). */
function alignSeq(a, b) {
  const pa = [], pb = [];
  const emit = (i, j) => { pa.push(i); pb.push(j); };
  function dp(a0, a1, b0, b1) {
    const n = a1 - a0, m = b1 - b0, W = m + 1;
    const t = new Uint16Array((n + 1) * W);
    for (let i = n - 1; i >= 0; i--) {
      const ai = a[a0 + i];
      for (let j = m - 1; j >= 0; j--) {
        const x = i * W + j;
        t[x] = ai === b[b0 + j] ? t[x + W + 1] + 1 : Math.max(t[x + W], t[x + 1]);
      }
    }
    let i = 0, j = 0;
    while (i < n && j < m) {
      if (a[a0 + i] === b[b0 + j]) { emit(a0 + i, b0 + j); i++; j++; }
      else if (t[(i + 1) * W + j] >= t[i * W + j + 1]) { emit(a0 + i, -1); i++; }
      else { emit(-1, b0 + j); j++; }
    }
    while (i < n) emit(a0 + i++, -1);
    while (j < m) emit(-1, b0 + j++);
  }
  function positional(a0, a1, b0, b1) {
    const n = a1 - a0, m = b1 - b0, k = Math.min(n, m);
    for (let i = 0; i < k; i++) emit(a0 + i, b0 + i);
    for (let i = k; i < n; i++) emit(a0 + i, -1);
    for (let j = k; j < m; j++) emit(-1, b0 + j);
  }
  function middle(a0, a1, b0, b1, depth) {
    if (a0 === a1) { for (let j = b0; j < b1; j++) emit(-1, j); return; }
    if (b0 === b1) { for (let i = a0; i < a1; i++) emit(i, -1); return; }
    if ((a1 - a0) * (b1 - b0) <= DP_MAX) { dp(a0, a1, b0, b1); return; }
    if (depth > MAX_DEPTH) { positional(a0, a1, b0, b1); return; }
    // ids that occur exactly once on each side
    const ca = new Map();
    for (let i = a0; i < a1; i++) { const v = a[i]; const e = ca.get(v); if (e === undefined) ca.set(v, i); else if (e >= 0) ca.set(v, -2); }
    const cb = new Map();
    for (let j = b0; j < b1; j++) { const v = b[j]; if (!ca.has(v) || ca.get(v) < 0) continue; const e = cb.get(v); if (e === undefined) cb.set(v, j); else if (e >= 0) cb.set(v, -2); }
    const cand = [];
    for (let i = a0; i < a1; i++) { const j = cb.get(a[i]); if (j !== undefined && j >= 0 && ca.get(a[i]) === i) cand.push([i, j]); }
    const anchors = lisPairs(cand);
    if (!anchors.length) { positional(a0, a1, b0, b1); return; }
    let pa0 = a0, pb0 = b0;
    for (const [i, j] of anchors) { rec(pa0, i, pb0, j, depth + 1); emit(i, j); pa0 = i + 1; pb0 = j + 1; }
    rec(pa0, a1, pb0, b1, depth + 1);
  }
  function rec(a0, a1, b0, b1, depth) {
    while (a0 < a1 && b0 < b1 && a[a0] === b[b0]) emit(a0++, b0++);
    let sa = a1, sb = b1;
    while (sa > a0 && sb > b0 && a[sa - 1] === b[sb - 1]) { sa--; sb--; }
    middle(a0, sa, b0, sb, depth);
    for (let k = 0; k < a1 - sa; k++) emit(sa + k, sb + k);
  }
  rec(0, a.length, 0, b.length, 0);
  return { pa, pb };
}

/** Longest chain of pairs increasing in both members (pairs come sorted by the first member). */
function lisPairs(pairs) {
  const n = pairs.length;
  if (!n) return [];
  const tails = [], prev = new Int32Array(n).fill(-1);
  for (let k = 0; k < n; k++) {
    const v = pairs[k][1];
    let lo = 0, hi = tails.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (pairs[tails[mid]][1] < v) lo = mid + 1; else hi = mid; }
    if (lo > 0) prev[k] = tails[lo - 1];
    tails[lo] = k;
  }
  const out = [];
  for (let k = tails[tails.length - 1]; k >= 0; k = prev[k]) out.push(pairs[k]);
  return out.reverse();
}

/** Indices (into seq) of a longest increasing subsequence. */
function lisIndex(seq) {
  const pairs = seq.map((v, i) => [i, v]);
  return new Set(lisPairs(pairs).map((p) => p[0]));
}

function looksLikeHeader(row) {
  if (!row || row.length < 2) return false;
  const vals = row.filter((v) => v !== '');
  if (vals.length < Math.max(2, row.length * 0.6)) return false;
  if (vals.some((v) => /^-?\d+([.,]\d+)?$/.test(v))) return false;
  return new Set(vals).size >= vals.length * 0.8;
}

/** Column pairs [{a, b}] in display order. */
function matchColumns(sa, sb) {
  const na = sa ? sa.cols : 0, nb = sb ? sb.cols : 0;
  const ha = sa && sa.rows[0], hb = sb && sb.rows[0];
  const cols = [];
  if (sa && sb && looksLikeHeader(ha) && looksLikeHeader(hb)) {
    const ids = new Map();
    const id = (v, side, c) => { const key = v === '' ? `\u0000${side}${c}` : v; let x = ids.get(key); if (x === undefined) { x = ids.size; ids.set(key, x); } return x; };
    const A = Int32Array.from({ length: na }, (_, c) => id(ha[c] || '', 'a', c));
    const B = Int32Array.from({ length: nb }, (_, c) => id(hb[c] || '', 'b', c));
    const { pa, pb } = alignSeq(A, B);
    // a removed and an added column next to each other: one column with a renamed header
    let k = 0;
    while (k < pa.length) {
      if (pa[k] >= 0 && pb[k] >= 0) { cols.push({ a: pa[k], b: pb[k] }); k++; continue; }
      const rem = [], add = [];
      while (k < pa.length && !(pa[k] >= 0 && pb[k] >= 0)) { if (pa[k] >= 0) rem.push(pa[k]); else add.push(pb[k]); k++; }
      const p = Math.min(rem.length, add.length);
      for (let i = 0; i < p; i++) cols.push({ a: rem[i], b: add[i] });
      for (let i = p; i < rem.length; i++) cols.push({ a: rem[i], b: -1 });
      for (let i = p; i < add.length; i++) cols.push({ a: -1, b: add[i] });
    }
    return { cols, byHeader: true };
  }
  for (let c = 0; c < Math.max(na, nb); c++) cols.push({ a: c < na ? c : -1, b: c < nb ? c : -1 });
  return { cols, byHeader: false };
}

function rowKeyFn(sheet, idxs, mode) {
  return (r) => { let s = ''; for (let k = 0; k < idxs.length; k++) { if (k) s += '\u0001'; s += cellText(sheet, r, idxs[k], mode); } return s; };
}

/** Share of non-empty shared cells that are equal (0..1). */
function similarity(sa, sb, ra, rb, shared, mode) {
  let eq = 0, n = 0;
  for (const c of shared) {
    const va = cellText(sa, ra, c.a, mode), vb = cellText(sb, rb, c.b, mode);
    if (va === '' && vb === '') continue;
    n++;
    if (va === vb) eq++;
  }
  return n ? eq / n : 1;
}

function emptyCounts() { return { same: 0, changed: 0, added: 0, removed: 0, moved: 0, cells: 0, colsAdded: 0, colsRemoved: 0 }; }

/** Fills kinds and counts for aligned pairs; pairs with different text become CHANGED. */
function finish(d, sa, sb, pa, pb, mode, keyA, keyB, moved) {
  const n = pa.length;
  d.ra = Int32Array.from(pa); d.rb = Int32Array.from(pb); d.k = new Uint8Array(n);
  const shared = d.cols.filter((c) => c.a >= 0 && c.b >= 0);
  const c = d.counts;
  const chg = [];
  for (let i = 0; i < n; i++) {
    const a = pa[i], b = pb[i];
    let k;
    if (a < 0) k = KIND.ADDED;
    else if (b < 0) k = KIND.REMOVED;
    else {
      let diff = 0;
      for (const col of shared) if (cellText(sa, a, col.a, mode) !== cellText(sb, b, col.b, mode)) diff++;
      // cells of added / removed columns are not "changed cells"; they count once, as columns
      c.cells += diff;
      k = diff ? KIND.CHANGED : KIND.SAME;
      if (moved && moved.has(i)) k = diff ? KIND.MOVED_CHANGED : KIND.MOVED;
    }
    d.k[i] = k;
    if (k === KIND.SAME) c.same++;
    else {
      chg.push(i);
      if (k === KIND.CHANGED) c.changed++;
      else if (k === KIND.ADDED) c.added++;
      else if (k === KIND.REMOVED) c.removed++;
      else { c.moved++; if (k === KIND.MOVED_CHANGED) c.changed++; }
    }
  }
  d.chg = Int32Array.from(chg);
  return d;
}

/** Order-based comparison of two sheets. */
function compareByOrder(d, sa, sb, mode) {
  const shared = d.cols.filter((c) => c.a >= 0 && c.b >= 0);
  const ids = new Map();
  const intern = (s) => { let x = ids.get(s); if (x === undefined) { x = ids.size; ids.set(s, x); } return x; };
  const na = sa ? sa.rows.length : 0, nb = sb ? sb.rows.length : 0;
  const keyA = rowKeyFn(sa, shared.map((c) => c.a), mode), keyB = rowKeyFn(sb, shared.map((c) => c.b), mode);
  const A = new Int32Array(na), B = new Int32Array(nb);
  for (let r = 0; r < na; r++) A[r] = intern(keyA(r));
  for (let r = 0; r < nb; r++) B[r] = intern(keyB(r));
  ids.clear();
  const { pa, pb } = alignSeq(A, B);
  // removed and added rows next to each other that share most cells: one changed row
  const oa = [], ob = [];
  let k = 0;
  while (k < pa.length) {
    if (pa[k] >= 0 && pb[k] >= 0) { oa.push(pa[k]); ob.push(pb[k]); k++; continue; }
    const rem = [], add = [];
    while (k < pa.length && !(pa[k] >= 0 && pb[k] >= 0)) { if (pa[k] >= 0) rem.push(pa[k]); else add.push(pb[k]); k++; }
    let i = 0, j = 0;
    while (i < rem.length && j < add.length) {
      if (similarity(sa, sb, rem[i], add[j], shared, mode) >= SIMILAR) { oa.push(rem[i]); ob.push(add[j]); i++; j++; }
      else if (j + 1 < add.length && similarity(sa, sb, rem[i], add[j + 1], shared, mode) >= SIMILAR) { oa.push(-1); ob.push(add[j]); j++; }
      else { oa.push(rem[i]); ob.push(-1); i++; }
    }
    for (; i < rem.length; i++) { oa.push(rem[i]); ob.push(-1); }
    for (; j < add.length; j++) { oa.push(-1); ob.push(add[j]); }
  }
  return finish(d, sa, sb, oa, ob, mode);
}

/** Pro: rows matched by the values of key columns (A column indices); moved rows found by a longest ordered chain. */
function compareByKey(d, sa, sb, mode, keyColsA) {
  const map = keyColsA.map((ca) => d.cols.find((c) => c.a === ca)).filter((c) => c && c.b >= 0);
  if (!map.length) throw Object.assign(new Error('The key columns are not in both sheets.'), { code: 'KEY' });
  const keyA = rowKeyFn(sa, map.map((c) => c.a), 'values'), keyB = rowKeyFn(sb, map.map((c) => c.b), 'values');
  const na = sa.rows.length, nb = sb.rows.length;
  const where = new Map();
  let dupA = 0, dupB = 0;
  for (let r = 1; r < na; r++) { const key = keyA(r); if (key.replace(/\u0001/g, '') === '') continue; if (where.has(key)) dupA++; else where.set(key, r); }
  const matchB = new Int32Array(nb).fill(-1);
  const usedA = new Uint8Array(na);
  const seenB = new Set();
  if (na && nb) { matchB[0] = 0; usedA[0] = 1; } // header row
  for (let r = 1; r < nb; r++) {
    const key = keyB(r);
    if (key.replace(/\u0001/g, '') === '') continue;
    if (seenB.has(key)) { dupB++; continue; }
    seenB.add(key);
    const a = where.get(key);
    if (a !== undefined) { matchB[r] = a; usedA[a] = 1; }
  }
  // B order; removed A rows go right after the nearest earlier matched A row
  const posOfA = new Int32Array(na).fill(-1);
  for (let r = 0; r < nb; r++) if (matchB[r] >= 0) posOfA[matchB[r]] = r;
  const after = new Map(); // b row -> removed a rows to show after it (-1: at the start)
  let lastB = -1;
  for (let a = 0; a < na; a++) {
    if (usedA[a]) { lastB = posOfA[a]; continue; }
    if (!after.has(lastB)) after.set(lastB, []);
    after.get(lastB).push(a);
  }
  const pa = [], pb = [];
  for (const a of after.get(-1) || []) { pa.push(a); pb.push(-1); }
  const seq = [], seqAt = [];
  for (let r = 0; r < nb; r++) {
    if (matchB[r] >= 0) { seq.push(matchB[r]); seqAt.push(pa.length); }
    pa.push(matchB[r]); pb.push(r);
    for (const a of after.get(r) || []) { pa.push(a); pb.push(-1); }
  }
  const inOrder = lisIndex(seq);
  const moved = new Set();
  for (let s = 0; s < seq.length; s++) if (!inOrder.has(s)) moved.add(seqAt[s]);
  finish(d, sa, sb, pa, pb, mode, null, null, moved);
  d.keys = map.map((c) => c.a);
  if (dupA || dupB) d.notes.push(`${(dupA + dupB).toLocaleString('en-US')} rows repeat a key value of an earlier row (${dupA} in the first file, ${dupB} in the second); each repeat is shown as removed or added.`);
  return d;
}

function pairSheets(A, B) {
  const pairs = [];
  const usedB = new Set();
  const left = [];
  A.sheets.forEach((s, i) => { const j = B.sheets.findIndex((t, jj) => !usedB.has(jj) && t.name === s.name); if (j >= 0) { usedB.add(j); pairs.push([i, j]); } else left.push(i); });
  const left2 = [];
  for (const i of left) { const nm = A.sheets[i].name.toLowerCase(); const j = B.sheets.findIndex((t, jj) => !usedB.has(jj) && t.name.toLowerCase() === nm); if (j >= 0) { usedB.add(j); pairs.push([i, j]); } else left2.push(i); }
  const restB = B.sheets.map((_, j) => j).filter((j) => !usedB.has(j));
  // by position among the sheets left over
  const p = Math.min(left2.length, restB.length);
  for (let k = 0; k < p; k++) { pairs.push([left2[k], restB[k]]); usedB.add(restB[k]); }
  for (let k = p; k < left2.length; k++) pairs.push([left2[k], -1]);
  for (let k = p; k < restB.length; k++) pairs.push([-1, restB[k]]);
  // workbook order of the second file; a sheet only in the first file goes after the sheet it followed there
  const keyOf = new Map();
  for (const pr of pairs) if (pr[1] >= 0) keyOf.set(pr, pr[1]);
  const bOfA = new Map(pairs.filter((pr) => pr[0] >= 0 && pr[1] >= 0).map((pr) => [pr[0], pr[1]]));
  for (const pr of pairs) {
    if (pr[1] >= 0) continue;
    let prevB = -1;
    for (let i = pr[0] - 1; i >= 0; i--) if (bOfA.has(i)) { prevB = bOfA.get(i); break; }
    keyOf.set(pr, prevB + 0.5 + pr[0] / 1e6);
  }
  pairs.sort((x, y) => keyOf.get(x) - keyOf.get(y));
  return pairs;
}

/**
 * opts: { mode: 'values' | 'formulas', keys: { [sheet index in A]: [A column indices] } }
 * Returns { sheets: [diff], totals, mode }. Each diff: { ia, ib, nameA, nameB, status, cols, byHeader, ra, rb, k, chg,
 * counts, notes, keys? }.
 */
function compareBooks(A, B, opts = {}) {
  const mode = opts.mode === 'formulas' ? 'formulas' : 'values';
  const out = { sheets: [], totals: emptyCounts(), mode, sheetsOnlyA: 0, sheetsOnlyB: 0 };
  for (const [ia, ib] of pairSheets(A, B)) {
    const sa = ia >= 0 ? A.sheets[ia] : null, sb = ib >= 0 ? B.sheets[ib] : null;
    const { cols, byHeader } = matchColumns(sa, sb);
    const d = { ia, ib, nameA: sa ? sa.name : null, nameB: sb ? sb.name : null, status: sa && sb ? 'both' : sa ? 'onlyA' : 'onlyB', cols, byHeader, counts: emptyCounts(), notes: [] };
    for (const s of [sa, sb]) if (s) for (const n of s.notes) d.notes.push(`${s === sa ? 'First' : 'Second'} file: ${n}`);
    if (sa && sb) {
      d.counts.colsAdded = cols.filter((c) => c.a < 0).length;
      d.counts.colsRemoved = cols.filter((c) => c.b < 0).length;
    }
    const keys = opts.keys && ia >= 0 ? opts.keys[ia] : null;
    if (sa && sb && keys && keys.length) compareByKey(d, sa, sb, mode, keys);
    else compareByOrder(d, sa, sb, mode);
    if (sa && sb) {
      const ma = sa.merges.join(' '), mb = sb.merges.join(' ');
      if (ma !== mb) {
        const setA = new Set(sa.merges), setB = new Set(sb.merges);
        const gone = sa.merges.filter((m) => !setB.has(m)), added = sb.merges.filter((m) => !setA.has(m));
        if (gone.length || added.length) d.notes.push(`Merged cells differ: ${added.length ? `merged in the second file ${added.slice(0, 6).join(', ')}${added.length > 6 ? ' and more' : ''}` : ''}${added.length && gone.length ? '; ' : ''}${gone.length ? `no longer merged ${gone.slice(0, 6).join(', ')}${gone.length > 6 ? ' and more' : ''}` : ''}.`);
      }
    }
    if (d.status === 'onlyA') out.sheetsOnlyA++;
    if (d.status === 'onlyB') out.sheetsOnlyB++;
    for (const key of Object.keys(out.totals)) out.totals[key] += d.counts[key];
    out.sheets.push(d);
  }
  return out;
}

/** True when nothing differs. */
function identical(res) {
  const t = res.totals;
  return !res.sheetsOnlyA && !res.sheetsOnlyB && !t.changed && !t.added && !t.removed && !t.moved && !t.colsAdded && !t.colsRemoved;
}

module.exports = { compareBooks, alignSeq, lisPairs, matchColumns, pairSheets, identical, looksLikeHeader };
