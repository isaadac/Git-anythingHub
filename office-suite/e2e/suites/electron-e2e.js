// The desktop app: opening a file from the command line, and the limits on what a page may touch.
const { _electron: electron } = require('playwright');
const fs = require('fs');
const path = require('path');
const SP = process.env.E2E_WORK;
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const sample = path.join(SP, 'sample.csv');
  fs.writeFileSync(sample, 'Name,Score\nAda,90\nLin,85\n');
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${path.join(SP, 'ud')}`, '.', sample], cwd: process.cwd(), env: { ...process.env, NODE_OPTIONS: '' } });
  await app.firstWindow();
  let w;
  for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes('sheets.html')); }
  check('a file given on the command line opens in Sheets', !!w);
  await w.waitForFunction(() => window.AtlasSheets && AtlasSheets.valueOf('B3') === 85);
  check('desktop bridge present', await w.evaluate(() => Platform.desktop === true));
  const target = path.join(SP, 'saved.xlsx');
  const rogue = await w.evaluate(async t => { try { await window.atlas.save({ data: new Uint8Array([1]), filePath: t }); return 'written'; } catch { return 'refused'; } }, target);
  check('saving to a path the user never chose is refused', rogue === 'refused' && !fs.existsSync(target), rogue);
  await w.evaluate(() => AtlasSheets.setCell('B4', '=AVERAGE(B2:B3)'));
  const saved = await w.evaluate(async t => window.atlas.save({ data: await AtlasSheets.serialize('csv'), filePath: t }), sample);
  check('saving back to the opened file works', !!saved && fs.readFileSync(sample, 'utf8').trim().split(/\r?\n/).length === 4);
  check('opening a path the user never chose is refused', !(await w.evaluate(() => window.atlas.openPath('/etc/passwd.csv'))));
  check('no general file-reading bridge', (await w.evaluate(() => typeof window.atlas.readFile)) === 'undefined');
  check('pages have no Node.js', (await w.evaluate(() => typeof require + typeof process)) === 'undefinedundefined');
  await w.evaluate(() => { window.atlas.newWindow('write'); window.atlas.newWindow('slides'); window.atlas.newWindow('notes'); });
  await sleep(2500);
  const titles = await Promise.all(app.windows().map(x => x.title()));
  check('every app opens in its own window', ['Isadek Write', 'Isadek Slides', 'Isadek Notes'].every(t => titles.some(x => x.includes(t))), titles.join(' | '));
  check('recent list has the opened file', (await w.evaluate(() => window.atlas.recent())).some(r => r.name === 'sample.csv'));
  const menu = await app.evaluate(({ Menu }) => Menu.getApplicationMenu().items.map(i => i.label));
  check('native menu', menu.includes('File') && menu.includes('Edit'), menu.join(','));
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('CRASH', e); process.exit(1); });
