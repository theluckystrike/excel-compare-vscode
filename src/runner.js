'use strict';
// Starts a worker thread per job, with a memory cap and a watchdog. A job that sends nothing for `idleMs` is killed
// and rejected with code 'TIMEOUT'; a worker that runs out of memory is rejected with code 'MEMORY'. The extension
// host keeps running either way.
const path = require('path');
const { Worker } = require('worker_threads');

let WORKER = path.join(__dirname, '..', 'dist', 'worker.js');
const MEMORY_MB = 2048;

function setWorker(p) { WORKER = p; }

function runJob(job, { idleMs = 90000, onProgress, transfer, memoryMb = MEMORY_MB } = {}) {
  return new Promise((resolve, reject) => {
    let w;
    try { w = new Worker(WORKER, { resourceLimits: { maxOldGenerationSizeMb: memoryMb, maxYoungGenerationSizeMb: 64 } }); } catch (e) { reject(e); return; }
    let done = false, timer;
    const finish = (fn, v) => { if (done) return; done = true; clearTimeout(timer); w.terminate().catch(() => {}); fn(v); };
    const arm = () => { clearTimeout(timer); timer = setTimeout(() => { const e = new Error(`Reading took longer than ${Math.round(idleMs / 1000)} seconds and was stopped.`); e.code = 'TIMEOUT'; finish(reject, e); }, idleMs); };
    arm();
    w.on('message', (m) => {
      if (m.progress) { arm(); try { if (onProgress) onProgress(m.progress); } catch { /* ignore */ } return; }
      if (m.ok) finish(resolve, m.result);
      else { const e = new Error(m.error); e.code = m.code; finish(reject, e); }
    });
    w.on('error', (err) => {
      const oom = err && err.code === 'ERR_WORKER_OUT_OF_MEMORY';
      const e = new Error(oom ? `Reading needed more than ${memoryMb} MB of memory and was stopped.` : String((err && err.message) || err));
      e.code = oom ? 'MEMORY' : 'WORKER';
      finish(reject, e);
    });
    w.on('exit', (code) => { if (!done) { const e = new Error(`The reader stopped (exit code ${code}).`); e.code = 'WORKER'; finish(reject, e); } });
    w.postMessage({ id: 1, job }, transfer || []);
  });
}

module.exports = { runJob, setWorker, get WORKER() { return WORKER; } };
