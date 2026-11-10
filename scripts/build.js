'use strict';
// Bundles the worker (src/worker.js + src/core + SheetJS) into dist/worker.js with esbuild. esbuild and xlsx come from
// the table-viewer product's node_modules (ZOVO_TV_MODULES), so this repo has no node_modules tree of its own.
const path = require('path');
const os = require('os');
const root = path.join(__dirname, '..');
const mods = process.env.ZOVO_TV_MODULES || path.join(os.homedir(), 'plugin-market-intel-work', 'factory', 'products', 'table-viewer-vscode', 'node_modules');
const esbuild = require(path.join(mods, 'esbuild'));
esbuild.build({
  absWorkingDir: root, entryPoints: ['src/worker.js'], outfile: 'dist/worker.js', bundle: true, minify: true, legalComments: 'none',
  platform: 'node', format: 'cjs', target: 'node18', nodePaths: [mods], logLevel: 'info',
  banner: { js: '// Excel Compare worker bundle: src/worker.js, src/core and SheetJS xlsx 0.20.3 (Apache-2.0, see THIRD-PARTY-NOTICES.txt).' },
}).catch(() => process.exit(1));
