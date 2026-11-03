// Desktop: an add-in runs in the app (its worker loads from the app's files) and reads the
// document; a file added to the Assistant is read there too.
const { _electron: electron } = require('playwright');
const fs = require('fs');
const path = require('path');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-addins-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow(); await sleep(1200);
  await home.evaluate(() => window.atlas.newWindow('write'));
  let w; for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes('write.html')); }
  await w.waitForFunction(() => window.AtlasWrite && AtlasWrite.addins);
  await w.click('#doc');
  await w.keyboard.type('Alpha beta gamma.');
  const text = fs.readFileSync(path.join(ROOT, 'addins/word-goal.isadek-addin'), 'utf8');
  await w.evaluate(t => { UI.choose = async () => 0; return AtlasWrite.addins.install(t); }, text);
  await w.click('[data-tab="insert"]');
  await w.click('.addin-cmd[data-cmd="show"]');
  const ok = await w.waitForFunction(() => /3 of 500 words/.test(document.querySelector('.addin-pane .body')?.textContent || ''), null, { timeout: 15000 }).then(() => true, () => false);
  check('desktop: an add-in runs and reads the document', ok, await w.evaluate(() => document.querySelector('.addin-pane .body')?.textContent || document.querySelector('.toast')?.textContent || ''));
  const md = await w.evaluate(async () => (await FileText.extract('a.md', new TextEncoder().encode('hello'))).text);
  check('desktop: files are read for the Assistant', md === 'hello');
  const pdf = await w.evaluate(async () => {
    await new Promise(r => { const s = document.createElement('script'); s.src = '../../node_modules/pdf-lib/dist/pdf-lib.min.js'; s.onload = r; document.head.appendChild(s); });
    const d = await PDFLib.PDFDocument.create(); d.addPage([300, 200]).drawText('Quarterly report', { x: 20, y: 100, size: 18 });
    try { return (await FileText.extract('r.pdf', await d.save())).text; } catch (e) { return 'ERR ' + e.message; }
  });
  check('desktop: …PDFs too (pdf.js from the app’s files)', pdf === '## Page 1\nQuarterly report', pdf);
  // Nothing an add-in tries gets out of the app (here: through the prototypes, import(), a worker, a socket).
  const hits = [];
  const trap = require('http').createServer((q, r) => { hits.push(q.url); r.writeHead(200, { 'access-control-allow-origin': '*', 'content-type': 'text/javascript' }); r.end('// nothing'); });
  trap.on('upgrade', (q, sock) => { hits.push(q.url); sock.destroy(); });
  await new Promise(r => trap.listen(0, '127.0.0.1', r));
  await w.evaluate(t => AtlasWrite.addins.install(t), require('../escape-addin')(`http://127.0.0.1:${trap.address().port}`));
  await w.click('.addin-cmd[data-cmd="escape"]');
  await w.waitForFunction(() => /"indexedDB"/.test(document.querySelector('.addin-pane .body')?.textContent || ''), null, { timeout: 15000 }).catch(() => {});
  const esc = await w.evaluate(() => document.querySelector('.addin-pane .body p')?.textContent || '');
  check('desktop: an add-in can’t get out of its worker', hits.length === 0 && /"fetch":"gone"/.test(esc), `${esc} :: ${hits.join(' ')}`);
  trap.close();
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
