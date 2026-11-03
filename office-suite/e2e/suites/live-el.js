const { _electron: electron } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-live-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow();
  await sleep(1500);
  await home.evaluate(() => { localStorage.setItem('atlas-user-name', 'Desk User'); window.atlas.newWindow('sheets'); });
  let a; for (let i = 0; i < 40 && !a; i++) { await sleep(250); a = app.windows().find(w => w.url().includes('sheets.html')); }
  await a.waitForFunction(() => window.AtlasSheets && document.querySelector('.live-btn'));
  await a.evaluate(() => AtlasSheets.setCell('A1', 'desktop host'));
  await a.click('.live-btn');
  await a.click('.modal .actions button:has-text("Start on this computer")');
  await sleep(3000); console.log('dbg', await a.evaluate(async () => [JSON.stringify(await window.atlas.collabHost()), AtlasSheets.live.status, document.querySelector('.toast')?.textContent, document.querySelector('.modal')?.textContent]));
  await a.waitForSelector('.modal p:has-text("room=")', { timeout: 10000 });
  const link = (await a.textContent('.modal p')).match(/ws:\/\/\S+/)[0];
  check('session hosted on this computer with a LAN address', /^ws:\/\/(\d+\.){3}\d+:\d+\/\?room=/.test(link), link);
  await a.click('.modal .actions button:has-text("Close")');
  await a.evaluate(l => window.atlas.collabJoin('sheets', l.replace(/ws:\/\/[^/]+/, 'ws://127.0.0.1:' + l.match(/:(\d+)\//)[1])), link);
  let b; for (let i = 0; i < 40 && !b; i++) { await sleep(250); b = app.windows().filter(w => w.url().includes('sheets.html'))[1]; }
  check('a second window joins from the link', await b.waitForFunction(() => window.AtlasSheets && AtlasSheets.valueOf('A1') === 'desktop host', null, { timeout: 15000 }).then(() => true, () => false));
  await b.evaluate(() => AtlasSheets.setCell('B2', 7));
  check('edits flow between windows', await a.waitForFunction(() => AtlasSheets.valueOf('B2') === 7, null, { timeout: 8000 }).then(() => true, () => false));
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
