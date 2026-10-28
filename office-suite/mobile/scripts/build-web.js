'use strict';
// Copies the app's pages into www/ for Capacitor, with just the library files they load
// (the pages refer to them as ../../node_modules/..., so the same layout is kept).
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '../..');      // office-suite/
const OUT = path.resolve(__dirname, '../www');
const SRC = path.join(ROOT, 'src/renderer');

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });
fs.cpSync(SRC, path.join(OUT, 'src/renderer'), { recursive: true, filter: p => !/\.(map|md)$/.test(p) });

// Library files the pages and workers use.
const wanted = new Set();
for (const f of fs.readdirSync(SRC, { recursive: true })) {
  if (!/\.(html|js)$/.test(f)) continue;
  const text = fs.readFileSync(path.join(SRC, f), 'utf8');
  for (const m of text.matchAll(/node_modules\/((?:@[\w.-]+\/)?[\w.-]+\/[\w./-]+\.(?:js|css|wasm))/g)) wanted.add(m[1]);
}
for (const rel of wanted) {
  const from = path.join(ROOT, 'node_modules', rel);
  if (!fs.existsSync(from)) { console.error('Missing library file: node_modules/' + rel + ' (run npm ci in office-suite/ first)'); process.exit(1); }
  const to = path.join(OUT, 'node_modules', rel);
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
}
fs.copyFileSync(path.join(ROOT, 'build/icon.png'), path.join(OUT, 'icon.png'));

// The start page. A meta refresh, because the pages' CSP allows no inline script.
fs.writeFileSync(path.join(OUT, 'index.html'), `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta http-equiv="refresh" content="0; url=src/renderer/index.html">
<title>Atlas Office</title></head>
<body style="background:#f3f2f1"></body></html>
`);
console.log(`www/ ready: pages + ${wanted.size} library files`);
