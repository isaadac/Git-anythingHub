// Desktop: Python in cells runs in the app (its worker loads Pyodide from the app's own files).
const { _electron: electron } = require('playwright');
const path = require('path');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${process.env.E2E_WORK}/ud-py`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow();
  await sleep(1000);
  await home.evaluate(() => window.atlas.newWindow('sheets'));
  let s; for (let i = 0; i < 40 && !s; i++) { await sleep(250); s = app.windows().find(w => w.url().includes('sheets.html')); }
  await s.waitForFunction(() => window.AtlasSheets);
  await s.evaluate(() => { AtlasSheets.setCell('A1', '20'); AtlasSheets.setCell('A2', '22'); AtlasSheets.setCell('B1', '=PY("sum(r[0] for r in xl(""A1:A2""))")'); });
  await s.waitForFunction(() => AtlasSheets.valueOf('B1') === 42, null, { timeout: 90000 }).catch(() => {});
  const v = await s.evaluate(() => { const x = AtlasSheets.valueOf('B1'); return x && x.error || x; });
  check('desktop: a Python formula runs in the app', v === 42, String(v));
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
