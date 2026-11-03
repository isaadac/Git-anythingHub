// Offline-first cloud documents: a save without a connection waits on the device and goes up when
// the connection is back; a file changed in the cloud meanwhile is never overwritten (a
// "conflicted copy" goes up next to it); recent cloud files open offline, the web app included.
const { chromium } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const XLSX = require(path.join(ROOT, 'node_modules/xlsx'));
const start = require('../cloud-mock');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const xlsx = rows => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sheet1'); return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }); };
const cellA1 = buf => { const wb = XLSX.read(buf, { type: 'buffer' }); return wb.Sheets[wb.SheetNames[0]].A1 && wb.Sheets[wb.SheetNames[0]].A1.v; };

(async () => {
  const { port, S, server } = await start(ROOT, { 'dav/Budget.xlsx': xlsx([['first']]) });
  const base = `http://127.0.0.1:${port}/src/renderer/`;
  const FILE = '/remote.php/dav/files/u/Budget.xlsx';
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  // The web app's offline files (sw.js) are switched on for this test.
  await ctx.addInitScript(() => localStorage.setItem('atlas-sw', '1'));
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.goto(base + 'index.html');
  await page.waitForFunction(() => window.Cloud && window.Offline);
  await page.evaluate(async () => { await navigator.serviceWorker.ready; });
  // Connect the WebDAV account and open the workbook from it.
  const acct = await page.evaluate(async url => (await Cloud.connect('webdav', { url, username: 'u', password: 'secret' })).id, `http://127.0.0.1:${port}/remote.php/dav/files/u/`);
  await page.evaluate(async ([a, id]) => {
    const f = await Cloud.fetchFile(a, { id, name: 'Budget.xlsx' }, '/remote.php/dav/files/u/');
    await Platform.openData('sheets', f.name, f.data, true, f.path);
  }, [acct, FILE]);
  await page.waitForURL(/sheets\.html/);
  await page.waitForFunction(() => window.AtlasSheets && AtlasSheets.valueOf('A1') === 'first');
  check('a cloud workbook opens', true);

  // ---- Save while offline
  await ctx.setOffline(true);
  await page.evaluate(() => AtlasSheets.setCell('A1', 'edited offline'));
  await page.evaluate(() => Platform.emit('save'));
  await page.waitForFunction(() => Offline.state.waiting === 1, null, { timeout: 5000 }).catch(() => {});
  const chip = await page.evaluate(() => { const c = document.querySelector('.sync-chip'); return c && !c.hidden ? c.textContent : ''; });
  check('saving offline keeps the save on this device', (await page.evaluate(() => Offline.state.waiting)) === 1 && cellA1(S.dav[FILE]) === 'first', String(await page.evaluate(() => Offline.state.waiting)));
  check('…and the title bar says so', chip === 'Saved on this device · uploads when online', chip);
  check('…and the toast too', /Saved on this device/.test(await page.evaluate(() => new Promise(r => setTimeout(() => r(document.querySelector('.toast').textContent), 200)))));
  // ---- Back online: it goes up by itself.
  await ctx.setOffline(false);
  await page.waitForFunction(() => Offline.state.waiting === 0, null, { timeout: 10000 }).catch(() => {});
  check('when the connection is back, the save goes up by itself', cellA1(S.dav[FILE]) === 'edited offline', cellA1(S.dav[FILE]));
  check('…and the title bar clears', await page.evaluate(() => document.querySelector('.sync-chip').hidden));

  // ---- A conflict: changed in the cloud while this device was offline
  await ctx.setOffline(true);
  await page.evaluate(() => AtlasSheets.setCell('A1', 'mine'));
  await page.evaluate(() => Platform.emit('save'));
  await page.waitForFunction(() => Offline.state.waiting === 1, null, { timeout: 5000 }).catch(() => {});
  await sleep(300);
  S.dav[FILE] = xlsx([['theirs']]); // someone else saved it meanwhile
  await ctx.setOffline(false);
  await page.waitForFunction(() => Offline.state.waiting === 0, null, { timeout: 10000 }).catch(() => {});
  const copies = Object.keys(S.dav).filter(k => /conflicted%20copy/.test(k));
  check('a file changed in the cloud meanwhile is not overwritten', cellA1(S.dav[FILE]) === 'theirs', cellA1(S.dav[FILE]));
  check('…this device\'s version goes up next to it as a conflicted copy', copies.length === 1 && /^\/remote\.php\/dav\/files\/u\/Budget%20\(conflicted%20copy%20\d{4}-\d\d-\d\d%20\d\d\.\d\d\)\.xlsx$/.test(copies[0]) && cellA1(S.dav[copies[0]]) === 'mine', copies.join());
  check('…and the person is told', /changed in the cloud while you were offline/.test(await page.textContent('.toast')), await page.textContent('.toast'));

  // ---- A save that can't reach the server (the connection drops) also waits
  await page.route('**/remote.php/**', r => r.abort('internetdisconnected'));
  await page.evaluate(() => AtlasSheets.setCell('A2', 'dropped'));
  await page.evaluate(() => Platform.emit('save'));
  await page.waitForFunction(() => Offline.state.waiting === 1, null, { timeout: 5000 }).catch(() => {});
  check('a save that can\'t reach the cloud waits too, rather than failing', (await page.evaluate(() => Offline.state.waiting)) === 1);
  await page.unroute('**/remote.php/**');
  await page.evaluate(() => Offline.flush());
  check('…and goes up on the next try', (await page.evaluate(() => Offline.state.waiting)) === 0);

  // ---- Offline, the web app opens, and so does the recent cloud file (this device's copy)
  await page.evaluate(() => { window.__atlasDirty = false; });
  await page.goto(base + 'index.html');
  await page.waitForFunction(() => document.querySelector('#recent .file-card'));
  await ctx.setOffline(true);
  await page.reload();
  const home = await page.evaluate(() => !!document.querySelector('#recent .file-card') && !!window.Platform);
  check('offline, the web app still opens (its files are kept on the device)', home);
  await page.click('#recent .file-card:has-text("Budget.xlsx")');
  await page.waitForURL(/sheets\.html/, { timeout: 10000 }).catch(() => {});
  await page.waitForFunction(() => window.AtlasSheets && AtlasSheets.valueOf('A2') === 'dropped', null, { timeout: 10000 }).catch(() => {});
  check('…and a recent cloud workbook opens from the copy on this device', await page.evaluate(() => window.AtlasSheets && AtlasSheets.valueOf('A2') === 'dropped' && AtlasSheets.valueOf('A1') === 'mine'));
  check('…saying so', /You're offline: this is the copy on this device/.test(await page.evaluate(() => new Promise(r => setTimeout(() => r((document.querySelector('.toast') || {}).textContent || ''), 1200)))));
  check('…with the Offline marker in the title bar', await page.evaluate(() => { const c = document.querySelector('.sync-chip'); return !!c && !c.hidden && c.textContent === 'Offline'; }));
  await ctx.setOffline(false);

  await page.evaluate(() => { window.__atlasDirty = false; });
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
  server.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
