'use strict';
// Checks the central directory of a zip based workbook (.xlsx, .xlsm, .ods) before any byte is inflated.
// A zip bomb (a few KB that inflate to GBs) is refused here; the worker's memory cap and watchdog stay as a
// second guard for archives whose directory lies about the sizes.

const LIMITS = {
  maxEntries: 20000,
  maxTotal: 600 * 1024 * 1024, // all members together, uncompressed
  maxEntry: 400 * 1024 * 1024, // one member, uncompressed
  maxRatio: 400, // uncompressed / compressed, checked for members over 8 MB
};

function fail(msg) { const e = new Error(msg); e.code = 'LIMIT'; return e; }

/** True when the bytes start like a zip archive. */
function isZip(u8) { return u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4b && (u8[2] === 3 || u8[2] === 5) && (u8[3] === 4 || u8[3] === 6); }

/** Throws (code LIMIT) when the archive would inflate past the limits. Returns { entries, total }. */
function checkZip(u8, limits) {
  const L = { ...LIMITS, ...(limits || {}) };
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const n = u8.length;
  // End of central directory: the last 22 + up to 65535 comment bytes.
  let eocd = -1;
  for (let i = n - 22; i >= Math.max(0, n - 22 - 65535); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw fail('The file looks like a zip archive, but its directory is missing (the file may be cut short).');
  let count = dv.getUint16(eocd + 10, true);
  let cdSize = dv.getUint32(eocd + 12, true);
  let cdOff = dv.getUint32(eocd + 16, true);
  if (count === 0xffff || cdOff === 0xffffffff) {
    // zip64: locator 20 bytes before the EOCD
    const loc = eocd - 20;
    if (loc < 0 || dv.getUint32(loc, true) !== 0x07064b50) throw fail('The zip64 directory of this file is damaged.');
    const z = Number(dv.getBigUint64(loc + 8, true));
    if (z < 0 || z + 56 > n || dv.getUint32(z, true) !== 0x06064b50) throw fail('The zip64 directory of this file is damaged.');
    count = Number(dv.getBigUint64(z + 32, true));
    cdSize = Number(dv.getBigUint64(z + 40, true));
    cdOff = Number(dv.getBigUint64(z + 48, true));
  }
  if (count > L.maxEntries) throw fail(`The archive has ${count} members; workbooks are read up to ${L.maxEntries}.`);
  if (cdOff + cdSize > n) throw fail('The zip directory points past the end of the file (the file may be cut short).');
  let p = cdOff, total = 0;
  for (let k = 0; k < count; k++) {
    if (p + 46 > n || dv.getUint32(p, true) !== 0x02014b50) throw fail('The zip directory of this file is damaged.');
    let comp = dv.getUint32(p + 20, true);
    let size = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
    if (size === 0xffffffff || comp === 0xffffffff) {
      // zip64 extra field: sizes are 8 bytes each, in this order, only when the 32-bit value is 0xffffffff
      let e = p + 46 + nameLen; const end = e + extraLen;
      while (e + 4 <= end && e + 4 <= n) {
        const id = dv.getUint16(e, true), len = dv.getUint16(e + 2, true);
        if (id === 1) {
          let q = e + 4;
          if (size === 0xffffffff && q + 8 <= n) { size = Number(dv.getBigUint64(q, true)); q += 8; }
          if (comp === 0xffffffff && q + 8 <= n) { comp = Number(dv.getBigUint64(q, true)); }
          break;
        }
        e += 4 + len;
      }
    }
    if (size > L.maxEntry) throw fail(`One part of this workbook inflates to ${Math.round(size / 1048576)} MB; parts are read up to ${Math.round(L.maxEntry / 1048576)} MB.`);
    if (size > 8 * 1024 * 1024 && size / Math.max(1, comp) > L.maxRatio) throw fail(`One part of this workbook is compressed ${Math.round(size / Math.max(1, comp))} to 1, which is how zip bombs look. It was not opened.`);
    total += size;
    if (total > L.maxTotal) throw fail(`This workbook inflates to more than ${Math.round(L.maxTotal / 1048576)} MB, the limit for one file.`);
    p += 46 + nameLen + extraLen + commentLen;
  }
  return { entries: count, total };
}

module.exports = { checkZip, isZip, LIMITS };
