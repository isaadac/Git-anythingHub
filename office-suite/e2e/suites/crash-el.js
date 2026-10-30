// Desktop: crash reports — off by default and kept on the computer, sent once turned on; the Help window.
const { _electron: electron } = require('playwright');
const fs = require('fs');
const http = require('http');
const os = require('os');
const path = require('path');
const { createCrashCollector } = require('../../src/server/crash-collector');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const until = async (f, ms = 15000) => { const t = Date.now(); while (Date.now() - t < ms) { const v = await f(); if (v) return v; await sleep(250); } return null; };

(async () => {
  const inbox = path.join(process.env.E2E_WORK, 'crash-inbox');
  const collector = createCrashCollector({ dir: inbox });
  const server = http.createServer((req, res) => collector.handle(req, res));
  await new Promise(r => server.listen(0, '127.0.0.1', r));
  const ud = path.join(process.env.E2E_WORK, 'ud-crash-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '', ISADEK_CRASH_URL: `http://127.0.0.1:${server.address().port}/api/crash` } });
  const home = await app.firstWindow(); await sleep(1200);
  await home.waitForFunction(() => window.Platform);
  const log = () => { try { return fs.readFileSync(path.join(ud, 'crash-reports', 'errors.jsonl'), 'utf8').split('\n').filter(Boolean).map(l => JSON.parse(l)); } catch { return []; } };
  const received = () => fs.readdirSync(inbox);
  const menuItem = label => app.evaluate(({ Menu }, l) => { const find = items => { for (const i of items) { if (i.label === l) return { checked: i.checked }; if (i.submenu) { const f = find(i.submenu.items); if (f) return f; } } return null; }; return find(Menu.getApplicationMenu().items); }, label);

  // ---- Off by default: the report is kept, not sent
  check('crash reports are off to begin with', !(await home.evaluate(() => window.atlas.crashStatus())).enabled && !(await menuItem('Send Crash Reports')).checked);
  await home.evaluate(h => setTimeout(() => { throw new Error('Boom in ' + h + '/Documents/secret.docx'); }), os.homedir());
  const first = await until(() => log().find(e => /Boom/.test(e.message)));
  check('a script error in a window is kept in the local crash log', first && first.where === 'index' && first.kind === 'error', JSON.stringify(first));
  check('…without the home folder (and so the user name) in it', first && !first.message.includes(os.homedir()) && first.message.includes('~/Documents/secret.docx'), first && first.message);
  await sleep(800);
  check('…and nothing is sent', received().length === 0, received().join());

  // ---- Turned on: errors go to the report address
  await home.evaluate(() => window.atlas.setCrashReports(true));
  check('turning reports on ticks the Help menu item', (await menuItem('Send Crash Reports')).checked);
  await home.evaluate(() => { Promise.reject(new Error('Second problem')); });
  const sent = await until(() => received().find(f => /^errors-.*\.jsonl$/.test(f)));
  const rep = sent && fs.readFileSync(path.join(inbox, sent), 'utf8').split('\n').filter(Boolean).map(l => JSON.parse(l));
  check('once on, an error is sent as a report', rep && rep.some(r => r.message === 'Second problem' && r.kind === 'rejection' && r.version && r.platform === process.platform), JSON.stringify(rep));

  // ---- A window that stops working: noted, and its minidump sent
  await app.evaluate(({ dialog }) => { dialog.showMessageBox = async () => ({ response: 1 }); });
  await home.evaluate(() => window.atlas.newWindow('write'));
  let w; for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes('write.html')); }
  await w.waitForFunction(() => window.AtlasWrite);
  await app.evaluate(({ BrowserWindow }) => { const win = BrowserWindow.getAllWindows().find(x => x.webContents.getURL().includes('write.html')); win.webContents.forcefullyCrashRenderer(); });
  const gone = await until(() => log().find(e => e.kind === 'window-gone'));
  check('a window that stops working is noted', gone && gone.where === 'write' && gone.reason, JSON.stringify(gone));
  const dump = await until(() => received().find(f => f.endsWith('.dmp')), 30000);
  const meta = dump && JSON.parse(fs.readFileSync(path.join(inbox, dump.replace('.dmp', '.json')), 'utf8'));
  check('…and its minidump reaches the collector, with the app version', dump && fs.readFileSync(path.join(inbox, dump)).slice(0, 4).toString() === 'MDMP' && meta && meta._version, JSON.stringify(meta || {}).slice(0, 200));

  // ---- Help
  await app.evaluate(({ Menu }) => { const find = items => { for (const i of items) { if (i.label === 'Privacy') return i; if (i.submenu) { const f = find(i.submenu.items); if (f) return f; } } }; find(Menu.getApplicationMenu().items).click(); });
  const help = await until(() => app.windows().find(x => x.url().includes('help.html')));
  await until(async () => (await help.evaluate(() => location.hash)) === '#privacy');
  const h = await help.evaluate(() => ({ hash: location.hash, toc: document.querySelectorAll('#toc a').length, on: document.querySelector('#toc a.on') && document.querySelector('#toc a.on').hash, setting: !document.getElementById('crashSetting').hidden && document.getElementById('crashToggle').checked }));
  check('Help → Privacy opens the help window at the privacy policy', h.hash === '#privacy' && h.toc >= 10 && h.on === '#privacy', JSON.stringify(h));
  check('…which shows the crash-report setting as it is', h.setting);
  await help.fill('#q', 'pivottable');
  const shown = await help.$$eval('main section:not(.hidden)', s => s.map(x => x.id));
  check('searching Help shows the matching topics only', shown.join() === 'sheets', shown.join());
  await help.fill('#q', '');
  await help.click('#crashToggle');
  await sleep(300);
  check('turning reports off in Help turns them off', !(await home.evaluate(() => window.atlas.crashStatus())).enabled && !(await menuItem('Send Crash Reports')).checked);

  await require('../close-app')(app);
  collector.close();
  server.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
