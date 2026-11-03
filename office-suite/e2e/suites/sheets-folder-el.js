// From Folder in the desktop app: read a folder the person picks, read it again with Refresh
// All, and ask before reading a folder a workbook names that wasn't picked here.
const { _electron: electron } = require('playwright');
const path = require('path'), fs = require('fs');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const dir = path.join(process.env.E2E_WORK, 'sales-' + Date.now());
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, 'a.csv'), 'Item,Qty\nPens,3\n');
  fs.writeFileSync(path.join(dir, 'b.csv'), 'Item,Qty\nInk,5\n');
  const ud = path.join(process.env.E2E_WORK, 'ud-folder-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  // The native dialogs answer as a person would.
  await app.evaluate(({ dialog }, d) => {
    global.__asked = [];
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [d] });
    dialog.showMessageBox = async (w, o) => { global.__asked.push((o || w).detail); return { response: 1 }; };
  }, dir);
  const home = await app.firstWindow(); await sleep(1200);
  await home.click('.app-card[data-app="sheets"]');
  let s; for (let i = 0; i < 40 && !s; i++) { await sleep(250); s = app.windows().find(w => w.url().includes('sheets.html')); }
  await s.waitForFunction(() => window.AtlasSheets);
  await s.click('[data-tab="data"]');
  await s.click('#fromFolder');
  await s.waitForFunction(() => AtlasSheets.book.sheets.some(x => x.folder), null, { timeout: 15000 });
  const rows = () => s.evaluate(() => { const out = []; for (let r = 2; r <= 4; r++) out.push(AtlasSheets.valueOf('A' + r) + ':' + AtlasSheets.valueOf('B' + r)); return out.join(' '); });
  check('the folder’s files become one table', (await rows()) === 'a.csv:Pens b.csv:Ink :', await rows());
  fs.writeFileSync(path.join(dir, 'c.csv'), 'Item,Qty\nPaper,9\n');
  await s.click('#refreshAll');
  await s.waitForFunction(() => AtlasSheets.valueOf('A4') === 'c.csv', null, { timeout: 15000 });
  check('Refresh All reads the folder again (no question: it was picked here)', (await app.evaluate(() => global.__asked.length)) === 0);
  // A workbook naming another folder: the app asks first (and the answer here is Cancel).
  await s.evaluate(d => { AtlasSheets.book.sheets.find(x => x.folder).folder.path = d; AtlasSheets.setCell('A2', 'kept'); }, path.join(process.env.E2E_WORK));
  await s.click('#refreshAll');
  await sleep(1500);
  const asked = await app.evaluate(() => global.__asked);
  check('a folder the workbook names is read only after asking', asked.length === 1 && asked[0].startsWith(process.env.E2E_WORK) && (await s.evaluate(() => AtlasSheets.valueOf('A2'))) === 'kept', JSON.stringify(asked));
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
