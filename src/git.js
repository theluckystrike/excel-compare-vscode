'use strict';
// Reads the committed (HEAD) bytes of a file with `git show HEAD:./<name>`, run without a shell, with a timeout and
// with hooks and fsmonitor switched off. Nothing leaves the machine.
const path = require('path');
const { execFile } = require('child_process');

function runGit(args, cwd, maxBuffer, exec) {
  const run = exec || execFile;
  return new Promise((resolve) => {
    run('git', ['-c', 'core.fsmonitor=false', '-c', 'core.hooksPath=/dev/null', '--no-pager', ...args], { cwd, encoding: 'buffer', maxBuffer, timeout: 20000, windowsHide: true }, (err, stdout, stderr) => {
      if (err) {
        const msg = (stderr && stderr.length ? stderr.toString('utf8') : String(err.message || err)).trim();
        return resolve({ ok: false, code: err.code, killed: !!err.killed, error: msg.slice(0, 400) });
      }
      resolve({ ok: true, stdout });
    });
  });
}

/**
 * { ok: true, bytes } or { ok: false, quiet, message }. quiet: not a repository or no git (the caller only shows
 * a status bar message then).
 */
async function readFromHead(fsPath, maxBytes, exec) {
  const dir = path.dirname(fsPath);
  const base = path.basename(fsPath);
  const r = await runGit(['show', `HEAD:./${base}`], dir, Math.max(maxBytes, 1024 * 1024), exec);
  if (r.ok) {
    const head = r.stdout.subarray(0, 64).toString('utf8');
    if (head.startsWith('version https://git-lfs.github.com/spec/')) return { ok: false, message: 'The committed version is a Git LFS pointer; run git lfs pull to get the workbook.' };
    return { ok: true, bytes: r.stdout };
  }
  const e = r.error || '';
  if (r.code === 'ENOENT') return { ok: false, quiet: true, message: 'Git is not installed or not on the PATH.' };
  if (/not a git repository/i.test(e)) return { ok: false, quiet: true, message: `${base} is not in a Git repository.` };
  if (/exists on disk, but not in|does not exist in|path .* does not exist|not a valid object name|invalid object name|bad revision|unknown revision|ambiguous argument 'HEAD/i.test(e)) return { ok: false, message: `${base} has no committed version yet (it is new, or the repository has no commits).` };
  if (/maxBuffer/i.test(e) || r.code === 'ERR_CHILD_PROCESS_STDIO_MAXBUFFER') return { ok: false, message: 'The committed version is larger than the file size limit (setting excelCompare.maxFileSizeMB).' };
  if (r.killed) return { ok: false, message: 'Git took longer than 20 seconds and was stopped.' };
  return { ok: false, message: `Git could not read the committed version: ${e}` };
}

module.exports = { readFromHead, runGit };
