// Isadek Forms in the desktop app: publishing to "This computer" starts a small server on the
// local network that serves only the answer page; answers are kept in the app's data folder and
// arrive in the workbook.
const { _electron: electron, chromium } = require('playwright');
const path = require('path'), fs = require('fs');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-forms-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow(); await sleep(1200);
  await home.click('.app-card[data-app="sheets"]');
  let s; for (let i = 0; i < 40 && !s; i++) { await sleep(250); s = app.windows().find(w => w.url().includes('sheets.html')); }
  await s.waitForFunction(() => window.AtlasSheets);
  const errors = [];
  s.on('pageerror', e => errors.push(e.message));
  s.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await s.click('[data-tab="insert"]');
  await s.click('#formBtn');
  await s.waitForSelector('.form-pane .fm-q');
  await s.fill('.form-pane [data-k="title"]', 'Sign-up');
  await s.click('.form-pane [data-a="publish"]');
  // Where to collect: this computer.
  await s.waitForSelector('.modal');
  const choices = await s.$$eval('.modal button', bs => bs.map(b => b.textContent));
  check('the desktop app offers this computer', choices.includes('This computer'), choices.join(','));
  await s.click('.modal button:has-text("This computer")');
  await s.waitForFunction(() => AtlasSheets.book.forms[0].id);
  const f = await s.evaluate(() => AtlasSheets.book.forms[0]);
  const link = await s.inputValue('.form-pane .fm-link');
  check('published to this computer, with a network link', f.server === 'local' && /^http:\/\/[\d.]+:\d+\/f\//.test(link), link);
  check('kept in the app’s data folder, without the key', fs.existsSync(path.join(ud, 'forms', f.id + '.json')) && !fs.readFileSync(path.join(ud, 'forms', f.id + '.json'), 'utf8').includes(f.key));

  // Someone on the network answers (here: a browser on this machine, by the local address).
  const local = link.replace(/^http:\/\/[\d.]+:/, 'http://127.0.0.1:');
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto(local);
  await p.waitForSelector('.fv-q');
  await p.fill('.fv-q[data-q="q1"] input', 'Grace');
  await p.click('.fv-submit');
  await p.waitForSelector('.fv-done');
  check('the answer page works from the network', /recorded/.test(await p.textContent('.fv-done')));
  const other = await p.goto(local.replace(/\/f\/.*/, '/src/renderer/sheets.html'));
  check('nothing else of the app is served', other.status() === 404);
  await b.close();

  await s.click('.form-pane [data-a="refresh"]');
  await s.waitForFunction(() => { const i = AtlasSheets.book.sheets.findIndex(x => x.name === AtlasSheets.book.forms[0].sheet); if (i < 0) return false; AtlasSheets.setSheet(i); return AtlasSheets.valueOf('B2') === 'Grace'; }, null, { timeout: 8000 });
  check('the answer arrives in the workbook', true);
  check('no page errors', errors.length === 0, errors.join(' | '));
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
