'use strict';
// End-to-end tests: drives the real pages in Chromium and the desktop app in Electron.
//   npm run e2e                 all suites (Electron suites need a display; use xvfb-run on Linux CI)
//   npm run e2e -- sheets cf    only suites whose names contain one of the words
// Each suite prints PASS/FAIL lines; any FAIL, crash or timeout fails the run.
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const SUITES = path.join(__dirname, 'suites');
const TIMEOUT = 6 * 60 * 1000;
const PARALLEL = Math.max(1, Math.min(4, os.cpus().length - 1));

// A static server for the pages (the suites load them over http, as the web version runs).
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.zip': 'application/zip', '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml', '.json': 'application/json', '.wasm': 'application/wasm', '.woff2': 'font/woff2', '.onnx': 'application/octet-stream' };
const server = http.createServer((req, res) => {
  const p = path.normalize(path.join(ROOT, decodeURIComponent(new URL(req.url, 'http://x').pathname)));
  if (!p.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
  fs.stat(p, (err, st) => {
    if (err || !st.isFile()) { res.writeHead(404); return res.end('not found'); }
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(p)] || 'application/octet-stream' });
    fs.createReadStream(p).pipe(res);
  });
});

function runSuite(file, base) {
  return new Promise(resolve => {
    const work = fs.mkdtempSync(path.join(os.tmpdir(), 'isadek-e2e-'));
    fs.cpSync(path.join(__dirname, 'fixtures'), path.join(work, 'files'), { recursive: true });
    fs.mkdirSync(path.join(work, 'shots'), { recursive: true });
    const t0 = Date.now();
    const child = spawn(process.execPath, [path.join(SUITES, file)], { cwd: ROOT, env: { ...process.env, E2E_BASE: base, E2E_WORK: work, NODE_OPTIONS: '' } });
    let out = '';
    child.stdout.on('data', d => { out += d; });
    child.stderr.on('data', d => { out += d; });
    const timer = setTimeout(() => { out += '\nFAIL suite timed out\n'; child.kill('SIGKILL'); }, TIMEOUT);
    child.on('close', code => {
      clearTimeout(timer);
      const pass = (out.match(/^PASS /gm) || []).length;
      const fail = (out.match(/^(FAIL|CRASH) /gm) || []).length + (code && !/^FAIL /m.test(out) ? 1 : 0);
      fs.rmSync(work, { recursive: true, force: true });
      resolve({ file, pass, fail, out, secs: ((Date.now() - t0) / 1000).toFixed(0) });
    });
  });
}

(async () => {
  const words = process.argv.slice(2);
  const all = fs.readdirSync(SUITES).filter(f => f.endsWith('.js')).sort();
  const chosen = words.length ? all.filter(f => words.some(w => f.includes(w))) : all;
  // Electron suites start the desktop app: one at a time. Browser suites run side by side.
  const isElectron = f => /_electron/.test(fs.readFileSync(path.join(SUITES, f), 'utf8'));
  // Electron fetches its binary the first time it is asked for (a fresh npm ci doesn't).
  if (chosen.some(isElectron)) require('electron');
  await new Promise(r => server.listen(0, '127.0.0.1', r));
  const base = `http://127.0.0.1:${server.address().port}`;
  const results = [];
  const report = r => {
    results.push(r);
    console.log(`${r.fail ? '✗' : '✓'} ${r.file.padEnd(22)} ${String(r.pass).padStart(3)} passed${r.fail ? `, ${r.fail} failed` : ''} (${r.secs}s)`);
    if (r.fail) console.log(r.out.split('\n').filter(l => /^(FAIL|CRASH)|Error|at /.test(l)).slice(0, 30).map(l => '    ' + l).join('\n'));
  };
  const web = chosen.filter(f => !isElectron(f)), desktop = chosen.filter(isElectron);
  const queue = [...web];
  await Promise.all(Array.from({ length: PARALLEL }, async () => { while (queue.length) report(await runSuite(queue.shift(), base)); }));
  for (const f of desktop) report(await runSuite(f, base));
  server.close();
  const failed = results.filter(r => r.fail);
  const passed = results.reduce((n, r) => n + r.pass, 0);
  console.log(`\n${results.length} suites, ${passed} checks passed, ${failed.length} suite${failed.length === 1 ? '' : 's'} failed`);
  process.exit(failed.length ? 1 : 0);
})();
