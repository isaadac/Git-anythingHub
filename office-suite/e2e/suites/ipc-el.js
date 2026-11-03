// Desktop: only the app's own pages can use the main process (files, network, accounts).
// A page from anywhere else, even with the app's preload, is refused.
const { _electron: electron } = require('playwright');
const path = require('path');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${process.env.E2E_WORK}/ud-ipc`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow();
  await sleep(800);
  check('desktop: the app\'s own pages can use it', Array.isArray(await home.evaluate(() => window.atlas.recent())));
  // A foreign page in a window that has the app's preload.
  await app.evaluate(async ({ BrowserWindow }, preload) => {
    const w = new BrowserWindow({ show: false, webPreferences: { preload, contextIsolation: true, sandbox: true } });
    await w.loadURL('data:text/html,<title>foreign</title><p>not the app');
  }, path.join(process.cwd(), 'src/main/preload.js'));
  let foreign; for (let i = 0; i < 20 && !foreign; i++) { await sleep(200); foreign = app.windows().find(w => w.url().startsWith('data:')); }
  const res = await foreign.evaluate(async () => {
    const out = {};
    for (const [k, f] of Object.entries({ recent: () => window.atlas.recent(), notes: () => window.atlas.notesList(), cloud: () => window.atlas.cloudAccountsGet(), web: () => window.atlas.fetchData('https://example.com') })) {
      try { await f(); out[k] = 'allowed'; } catch (e) { out[k] = /Not allowed/.test(e.message) ? 'refused' : e.message; }
    }
    out.frame = window.atlas.frame; // a synchronous message: no answer for a stranger
    return out;
  });
  check('desktop: a page that isn\'t the app\'s is refused', res.recent === 'refused' && res.notes === 'refused' && res.cloud === 'refused' && res.web === 'refused' && res.frame === null, JSON.stringify(res));
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
