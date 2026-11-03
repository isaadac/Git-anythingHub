'use strict';
// SheetJS 0.20.3 is published by SheetJS on cdn.sheetjs.com, not on the npm registry, so the app
// installs an npm republish of it (@e965/xlsx, built with provenance from that tarball). This
// checks the installed files are exactly SheetJS's own: every file in the official tarball must
// be in node_modules/xlsx with the same contents (package.json aside, which names the republish).
//   node scripts/verify-sheetjs.js            (skips if cdn.sheetjs.com can't be reached)
//   node scripts/verify-sheetjs.js --required (CI: fails instead)
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const https = require('https');
const { execFileSync } = require('child_process');

const VERSION = '0.20.3';
const URL_ = `https://cdn.sheetjs.com/xlsx-${VERSION}/xlsx-${VERSION}.tgz`;
const ROOT = path.join(__dirname, '..');
const required = process.argv.includes('--required');

function get(url, redirects = 3) {
  return new Promise((resolve, reject) => {
    https.get(url, { timeout: 60000 }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location && redirects) { res.resume(); return resolve(get(new URL(res.headers.location, url).href, redirects - 1)); }
      if (res.statusCode !== 200) { res.resume(); return reject(new Error(`HTTP ${res.statusCode}`)); }
      const parts = [];
      res.on('data', d => parts.push(d));
      res.on('end', () => resolve(Buffer.concat(parts)));
    }).on('error', reject).on('timeout', function () { this.destroy(new Error('timed out')); });
  });
}
const sha = f => crypto.createHash('sha256').update(fs.readFileSync(f)).digest('hex');
function files(dir, base = dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(e => (e.isDirectory() ? files(path.join(dir, e.name), base) : [path.relative(base, path.join(dir, e.name))]));
}

(async () => {
  const installed = path.join(ROOT, 'node_modules/xlsx');
  const pkg = JSON.parse(fs.readFileSync(path.join(installed, 'package.json'), 'utf8'));
  if (pkg.version !== VERSION) throw new Error(`node_modules/xlsx is ${pkg.version}, expected ${VERSION}`);
  let tgz;
  try { tgz = await get(URL_); } catch (err) {
    const msg = `Couldn’t download ${URL_} (${err.message})`;
    if (required) throw new Error(msg);
    console.log(msg + ': check skipped');
    return;
  }
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheetjs-'));
  fs.writeFileSync(path.join(dir, 'x.tgz'), tgz);
  execFileSync('tar', ['xzf', 'x.tgz'], { cwd: dir });
  const official = path.join(dir, 'package');
  const list = files(official).filter(f => f !== 'package.json');
  const bad = [];
  for (const f of list) {
    const mine = path.join(installed, f);
    if (!fs.existsSync(mine)) bad.push('missing: ' + f);
    else if (sha(mine) !== sha(path.join(official, f))) bad.push('different: ' + f);
  }
  const extra = files(installed).filter(f => f !== 'package.json' && !list.includes(f) && !/^(README\.md|LICENSE)$/i.test(f));
  fs.rmSync(dir, { recursive: true, force: true });
  if (bad.length) throw new Error('node_modules/xlsx isn’t SheetJS ' + VERSION + ':\n  ' + bad.join('\n  '));
  console.log(`SheetJS ${VERSION}: ${list.length} files match the official tarball (sha256 ${crypto.createHash('sha256').update(tgz).digest('hex').slice(0, 16)}…)${extra.length ? `; extra files in the install: ${extra.join(', ')}` : ''}`);
})().catch(err => { console.error(err.message); process.exit(1); });
