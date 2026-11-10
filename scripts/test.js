'use strict';
// npm test: builds the worker bundle, then runs every test/*.test.js with the Node test runner. SheetJS (for the
// workbooks the tests build) comes from the table-viewer product's node_modules, like the bundle.
const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const root = path.join(__dirname, '..');
const mods = process.env.ZOVO_TV_MODULES || path.join(os.homedir(), 'plugin-market-intel-work', 'factory', 'products', 'table-viewer-vscode', 'node_modules');
const b = spawnSync(process.execPath, [path.join(__dirname, 'build.js')], { stdio: 'inherit', env: { ...process.env, ZOVO_TV_MODULES: mods } });
if (b.status !== 0) process.exit(1);
const dir = path.join(root, 'test');
const only = process.argv.slice(2);
const files = fs.readdirSync(dir).filter((f) => f.endsWith('.test.js') && (!only.length || only.some((o) => f.includes(o)))).map((f) => path.join(dir, f));
const env = { ...process.env, NODE_PATH: mods, ZOVO_SPONSORS_MODULE: path.join(root, 'src', 'license', 'sponsors.js') };
const r = spawnSync(process.execPath, ['--test', ...files], { stdio: 'inherit', env });
process.exit(r.status === null ? 1 : r.status);
