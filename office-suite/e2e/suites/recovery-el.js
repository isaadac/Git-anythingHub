const { _electron: electron } = require('playwright');
const path = require('path'), fs = require('fs');
const SP = process.env.E2E_WORK;
const XLSX = require(path.join(process.cwd(), 'node_modules/xlsx'));
const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
const ud = path.join(SP, 'ud-' + Date.now());
const launch = args => electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.', ...args], env: { ...process.env, NODE_OPTIONS: '' } });
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const book = path.join(SP, 'autosave.xlsx');
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['start']]), 'Sheet1'); XLSX.writeFile(wb, book);
  let app = await launch([book]);
  await app.firstWindow(); await sleep(2500);
  const sheets = app.windows().find(w => w.url().includes('sheets.html'));
  const sw = await sheets.evaluate(() => { const i = document.querySelector('.autosave input'); return { disabled: i.disabled, title: document.querySelector('.autosave').title }; });
  check('AutoSave available for an opened .xlsx', !sw.disabled, JSON.stringify(sw));
  await sheets.click('.autosave input');
  await sheets.evaluate(() => AtlasSheets.setCell('A1', 'autosaved!'));
  await sleep(5000);
  const onDisk = XLSX.readFile(book).Sheets.Sheet1.A1.v;
  check('AutoSave wrote the change to the file', onDisk === 'autosaved!', onDisk);
  check('window is clean after AutoSave', (await sheets.evaluate(() => window.__atlasDirty)) === false);

  // Unsaved work in Write, then the app is killed
  await sheets.evaluate(() => window.atlas.newWindow('write'));
  await sleep(2000);
  const write = app.windows().find(w => w.url().includes('write.html'));
  await write.click('#doc'); await write.keyboard.type('Words that must survive a crash');
  await write.evaluate(() => window.dispatchEvent(new Event('pagehide')));
  await sleep(800);
  const drafts = fs.readdirSync(path.join(ud, 'recovery')).filter(n => n.endsWith('.json'));
  check('draft written to the recovery folder', drafts.length === 1, drafts.join());
  app.process().kill('SIGKILL');
  await sleep(1000);

  // Relaunch: Home offers the draft
  app = await launch([]);
  await app.firstWindow(); await sleep(2500);
  const home = app.windows().find(w => w.url().endsWith('index.html'));
  const listed = await home.$$eval('#recovered .row .n', ns => ns.map(n => n.textContent));
  check('Home lists the recovered document after the crash', listed.length === 1, JSON.stringify(listed));
  await home.click('#recovered .row button.primary');
  await sleep(2500);
  const restored = app.windows().find(w => w.url().includes('write.html'));
  const text = restored && await restored.evaluate(() => document.getElementById('doc').textContent);
  check('recovered document opens with the lost words', /survive a crash/.test(text || ''), text);
  const relisted = await home.evaluate(async () => (await Platform.drafts.list()).length);
  check('an open draft is not offered twice', relisted === 0, String(relisted));
  // Closing with "Don't Save" equivalent: clean close removes the draft
  await restored.evaluate(() => window.atlas.closeNow());
  await sleep(800);
  const left = fs.readdirSync(path.join(ud, 'recovery')).filter(n => n.endsWith('.json'));
  check('closing on purpose removes the draft', left.length === 0, left.join());
  for (const w of app.windows()) await w.evaluate(() => { window.__atlasDirty = false; }).catch(() => {});
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e); process.exit(1); });
