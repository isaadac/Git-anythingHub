// Desktop: reading pictures and handwriting works in the app, from its own files (no download):
// the text reader's worker, engine and English model load from the app's folder.
const { _electron: electron } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-ocr-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const web = [];
  app.context().on('request', r => { if (/^https?:/.test(r.url())) web.push(r.url()); });
  const home = await app.firstWindow(); await sleep(1200);
  await home.click('.app-card[data-app="notes"]');
  let n; for (let i = 0; i < 40 && !n; i++) { await sleep(250); n = app.windows().find(w => w.url().includes('notes.html')); }
  await n.waitForFunction(() => window.AtlasNotes && AtlasNotes.page && window.Ocr);
  const r = await n.evaluate(async () => {
    const c = document.createElement('canvas'); c.width = 600; c.height = 90;
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, 600, 90); g.fillStyle = '#111'; g.font = '30px serif'; g.fillText('Invoice total due Friday', 14, 56);
    try { return (await Ocr.recognize(c)).text; } catch (e) { return 'ERROR ' + e.message; }
  });
  check('desktop: a picture’s words are read on this device', /Invoice total due Friday/.test(r), r);
  check('…with nothing fetched from the web', !web.length, web.join(' '));
  await n.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
