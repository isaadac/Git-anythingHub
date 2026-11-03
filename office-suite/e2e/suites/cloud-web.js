const { chromium } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const XLSX = require(path.join(ROOT, 'node_modules/xlsx'));
const D = require(path.join(ROOT, 'node_modules/docx'));
const JSZip = require(path.join(ROOT, 'node_modules/jszip'));
const start = require('../cloud-mock');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const xlsx = rows => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sheet1'); return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }); };
const docText = async buf => (await (await JSZip.loadAsync(buf)).file('word/document.xml').async('string')).replace(/<[^>]+>/g, '');

(async () => {
  const report = await D.Packer.toBuffer(new D.Document({ sections: [{ children: [new D.Paragraph('Quarterly report from Drive')] }] }));
  const { port, S, server } = await start(ROOT, {
    'google/Report.docx': report, 'google/notes.zip': Buffer.from('x'), __nativeXlsx: xlsx([['from google sheets']]),
    'dropbox/Plan.docx': report, 'ms/Deck data.xlsx': xlsx([['onedrive']]), 'dav/Shared Budget.xlsx': xlsx([['dav', 42]])
  });
  const base = `http://127.0.0.1:${port}/src/renderer/`;
  const mk = p => `http://127.0.0.1:${port}/mock/${p}`;
  const cfg = { google: { clientId: 'gid' }, microsoft: { clientId: 'mid' }, dropbox: { clientId: 'did' },
    endpoints: { google: { auth: mk('google/auth'), token: mk('google/token'), api: mk('google/api'), upload: mk('google/upload') },
      microsoft: { auth: mk('ms/auth'), token: mk('ms/token'), api: mk('ms/api') },
      dropbox: { auth: mk('dropbox/auth'), token: mk('dropbox/token'), api: mk('dropbox/api'), content: mk('dropbox/content') } } };
  const browser = await chromium.launch();
  const ctx = await browser.newContext();
  await ctx.addInitScript(c => { if (!localStorage.getItem('atlas-cloud-config')) localStorage.setItem('atlas-cloud-config', c); }, JSON.stringify(cfg));
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.goto(base + 'index.html');

  // ---- Connect Google Drive through the popup
  await page.click('#cloudAdd');
  await page.click('.modal .actions button:has-text("Google Drive")');
  await page.waitForSelector('#cloud .row', { timeout: 10000 });
  const rows = await page.$$eval('#cloud .row .p', ns => ns.map(n => n.textContent));
  check('Google account connected via the popup (PKCE checked by the server)', rows[0] === 'ada@gmail.test', rows.join());
  const stored = await page.evaluate(() => JSON.parse(localStorage.getItem('atlas-cloud-accounts')).accounts[0]);
  check('tokens stored', stored.access && stored.refresh === 'refresh-google', JSON.stringify(stored));

  // ---- Browse and open
  await page.click('#cloud .row button:has-text("Browse")');
  await page.waitForSelector('.cloud-row');
  const names = await page.$$eval('.cloud-row .nm', ns => ns.map(n => n.textContent));
  check('browser lists folders and documents only', names.join('|') === 'Report.docx|Projects|Native Sheet.xlsx' || (names.includes('Projects') && names.includes('Report.docx') && names.includes('Native Sheet.xlsx') && !names.includes('notes.zip')), names.join('|'));
  await page.dblclick('.cloud-row:has-text("Report.docx")');
  await page.waitForURL(/write\.html/);
  await page.waitForFunction(() => document.querySelector('#doc') && document.querySelector('#doc').innerText.includes('Quarterly'), null, { timeout: 10000 });
  check('docx opened from Drive in Write', true);
  const p1 = await page.evaluate(() => document.querySelector('.titlebar .docname input').value);
  check('document name kept', p1 === 'Report.docx', p1);
  const autosave = await page.evaluate(() => ({ dis: document.querySelector('.autosave input').disabled, t: document.querySelector('.autosave').title }));
  check('AutoSave is offered for a cloud document on the web', !autosave.dis, JSON.stringify(autosave));

  // ---- Edit and Ctrl+S writes back
  await page.click('#doc');
  await page.keyboard.press('End');
  await page.keyboard.type(' — edited in Atlas');
  await page.keyboard.press('Control+s');
  await page.waitForFunction(() => window.__atlasDirty === false, null, { timeout: 10000 });
  const g0 = Object.values(S.google.files).find(f => f.name === 'Report.docx');
  check('Save wrote back to the same Drive file', g0.saves === 1 && (await docText(g0.data)).includes('edited in Atlas'), g0.saves);

  // ---- Share with link
  await page.click('.titlebar [data-cmd=share]');
  await page.selectOption('.modal select', 'writer');
  await page.click('.modal .actions .primary');
  await page.waitForSelector('.modal p:has-text("drive.test")');
  const link = await page.textContent('.modal p');
  check('share link shown', link.includes(`https://drive.test/file/d/${g0.id}/view`), link);
  check('anyone-with-link edit permission created', S.perms.some(p => p.id === g0.id && p.type === 'anyone' && p.role === 'writer' && p.notify === 'false'), JSON.stringify(S.perms));
  await page.click('.modal .actions button:has-text("Close")');
  // invite
  await page.click('.titlebar [data-cmd=share]');
  await page.fill('.modal input[type=email]', 'bob@example.test');
  await page.click('.modal .actions .primary');
  await page.waitForSelector('.modal p:has-text("Invitation sent")');
  check('email invitation sent', S.perms.some(p => p.type === 'user' && p.emailAddress === 'bob@example.test' && p.role === 'reader' && p.notify === 'true'));
  await page.click('.modal .actions button:has-text("Close")');

  // ---- Token refresh after a 401
  S.tokens.clear();
  await page.evaluate(() => document.querySelector('#doc').append('more'));
  await page.evaluate(() => { window.__atlasDirty = true; });
  await page.keyboard.press('Control+s');
  await sleep(1500);
  check('expired token refreshed and save retried', S.refreshes === 1 && g0.saves === 2, `${S.refreshes} ${g0.saves}`);

  // ---- Sheets: Save to Cloud into a folder
  await page.goto(base + 'sheets.html');
  await page.waitForFunction(() => window.AtlasSheets);
  await page.evaluate(() => AtlasSheets.setCell('A1', 'saved to drive'));
  await page.evaluate(() => { Platform.emit('save-cloud'); });
  await page.waitForSelector('.cloud-row:has-text("Projects")');
  await page.dblclick('.cloud-row:has-text("Projects")');
  await page.waitForSelector('.cloud-list .empty:has-text("empty")');
  await page.fill('.cloud-name input', 'Budget 2027');
  await page.click('.modal .actions .primary');
  await page.waitForFunction(() => window.__atlasDirty === false, null, { timeout: 10000 });
  const nf = Object.values(S.google.files).find(f => f.name === 'Budget 2027.xlsx');
  check('new workbook uploaded into the folder with an extension', nf && nf.parents[0] === 'gf' && /spreadsheetml/.test(nf.mimeType), nf && JSON.stringify({ n: nf.name, p: nf.parents, m: nf.mimeType }));
  check('uploaded workbook is a real xlsx', nf && XLSX.read(nf.data).Sheets.Sheet1.A1.v === 'saved to drive');
  const title = await page.evaluate(() => document.querySelector('.titlebar .docname input').value);
  check('workbook now named after the cloud file', title === 'Budget 2027.xlsx', title);
  await page.evaluate(() => AtlasSheets.setCell('A2', 'second save'));
  await page.keyboard.press('Control+s');
  await page.waitForFunction(() => window.__atlasDirty === false, null, { timeout: 10000 });
  check('second save updates the uploaded file', XLSX.read(nf.data).Sheets.Sheet1.A2.v === 'second save' && nf.saves === 1);

  // ---- Native Google Sheet opens as xlsx; save creates an Office copy
  await page.goto(base + 'sheets.html');
  await page.waitForFunction(() => window.AtlasSheets);
  await page.evaluate(() => { Platform.emit('open-cloud'); });
  await page.dblclick('.cloud-row:has-text("Native Sheet.xlsx")');
  await page.waitForFunction(() => String(AtlasSheets.valueOf('A1')).includes('from google'), null, { timeout: 10000 }).catch(() => {});
  const a1 = await page.evaluate(() => AtlasSheets.valueOf('A1'));
  check('Google Sheets file opened (exported as xlsx)', String(a1).includes('from google sheets'), a1);
  await page.evaluate(() => AtlasSheets.setCell('B1', 'x'));
  await page.keyboard.press('Control+s');
  await page.waitForFunction(() => window.__atlasDirty === false, null, { timeout: 10000 });
  const copy = Object.values(S.google.files).filter(f => f.name === 'Native Sheet.xlsx' && f.id !== 'gn');
  check('saving a Google-native file creates an xlsx beside it', copy.length === 1 && copy[0].parents[0] === 'root', copy.length);

  // ---- Dropbox
  await page.goto(base + 'index.html');
  await page.click('#cloudAdd');
  await page.click('.modal .actions button:has-text("Dropbox")');
  await page.waitForFunction(() => document.querySelectorAll('#cloud .row').length === 2, null, { timeout: 10000 });
  await page.evaluate(() => { Cloud.openDialog({}); });
  await page.selectOption('.cloud-top select', { label: 'Dropbox — ada@dropbox.test' });
  await page.waitForSelector('.cloud-row:has-text("Plan.docx")');
  await page.dblclick('.cloud-row:has-text("Plan.docx")');
  await page.waitForURL(/write\.html/);
  await page.waitForFunction(() => document.querySelector('#doc').innerText.includes('Quarterly'));
  await page.click('#doc'); await page.keyboard.type('Dropbox edit ');
  await page.keyboard.press('Control+s');
  await page.waitForFunction(() => window.__atlasDirty === false, null, { timeout: 10000 });
  check('Dropbox save overwrites in place', S.lastDropboxArg.path === '/Plan.docx' && S.lastDropboxArg.mode === 'overwrite' && (await docText(S.dropbox['/Plan.docx'])).includes('Dropbox edit'), JSON.stringify(S.lastDropboxArg));
  await page.click('.titlebar [data-cmd=share]');
  await page.click('.modal .actions .primary');
  await page.waitForSelector('.modal p:has-text("dropbox.test")');
  check('Dropbox share falls back to the existing link', (await page.textContent('.modal p')).includes('https://dropbox.test/s/%2FPlan.docx'));
  await page.click('.modal .actions button:has-text("Close")');

  // ---- OneDrive
  await page.goto(base + 'index.html');
  await page.click('#cloudAdd');
  await page.click('.modal .actions button:has-text("OneDrive")');
  await page.waitForFunction(() => document.querySelectorAll('#cloud .row').length === 3, null, { timeout: 10000 });
  await page.evaluate(() => { Cloud.openDialog({}); });
  await page.selectOption('.cloud-top select', { label: 'OneDrive — ada@outlook.test' });
  await page.dblclick('.cloud-row:has-text("Deck data.xlsx")');
  await page.waitForURL(/sheets\.html/);
  await page.waitForFunction(() => window.AtlasSheets && AtlasSheets.valueOf('A1') === 'onedrive', null, { timeout: 10000 });
  check('OneDrive file opened (content redirect followed)', true);
  await page.evaluate(() => AtlasSheets.setCell('A1', 'onedrive edited'));
  await page.keyboard.press('Control+s');
  await page.waitForFunction(() => window.__atlasDirty === false, null, { timeout: 10000 });
  const md = Object.values(S.ms).find(f => f.name === 'Deck data.xlsx');
  check('OneDrive save updates the item', XLSX.read(md.data).Sheets.Sheet1.A1.v === 'onedrive edited');
  await page.click('.titlebar [data-cmd=share]');
  await page.click('.modal .actions .primary');
  await page.waitForSelector('.modal p:has-text("1drv.test")');
  check('OneDrive anonymous view link', S.shares.some(s => s.ms === md.id && s.type === 'view' && s.scope === 'anonymous'));
  await page.click('.modal .actions button:has-text("Close")');

  // ---- WebDAV (Nextcloud)
  await page.goto(base + 'index.html');
  await page.click('#cloudAdd');
  await page.click('.modal .actions button:has-text("WebDAV")');
  await page.fill('.modal input >> nth=0', `http://127.0.0.1:${port}/remote.php/dav/files/u/`);
  await page.fill('.modal input >> nth=1', 'u');
  await page.fill('.modal input >> nth=2', 'wrong');
  await page.click('.modal .actions .primary');
  await sleep(800);
  check('wrong WebDAV password is refused and not kept', (await page.$$('#cloud .row')).length === 3);
  await page.click('#cloudAdd');
  await page.click('.modal .actions button:has-text("WebDAV")');
  await page.fill('.modal input >> nth=0', `http://127.0.0.1:${port}/remote.php/dav/files/u/`);
  await page.fill('.modal input >> nth=1', 'u');
  await page.fill('.modal input >> nth=2', 'secret');
  await page.click('.modal .actions .primary');
  await page.waitForFunction(() => document.querySelectorAll('#cloud .row').length === 4, null, { timeout: 10000 });
  await page.evaluate(() => { Cloud.openDialog({}); });
  await page.selectOption('.cloud-top select', { index: 3 });
  await page.dblclick('.cloud-row:has-text("Shared Budget.xlsx")');
  await page.waitForURL(/sheets\.html/);
  await page.waitForFunction(() => window.AtlasSheets && AtlasSheets.valueOf('B1') === 42, null, { timeout: 10000 });
  check('WebDAV file with a space in its name opened', true);
  await page.evaluate(() => AtlasSheets.setCell('B1', 43));
  await page.keyboard.press('Control+s');
  await page.waitForFunction(() => window.__atlasDirty === false, null, { timeout: 10000 });
  check('WebDAV save PUTs back', XLSX.read(S.dav['/remote.php/dav/files/u/Shared%20Budget.xlsx']).Sheets.Sheet1.B1.v === 43);
  // Save to Cloud with an existing name -> no silent overwrite on create
  await page.evaluate(() => { Platform.emit('save-cloud'); });
  await page.selectOption('.cloud-top select', { index: 3 });
  await page.waitForSelector('.cloud-row:has-text("Shared Budget.xlsx")');
  await page.fill('.cloud-name input', 'Shared Budget.xlsx');
  await page.click('.modal.cloud-browser .actions .primary');
  await page.waitForSelector('.modal h2:has-text("Replace file")');
  await page.click('.modal .actions .cancel');
  await page.fill('.cloud-name input', 'Copy');
  await page.click('.modal.cloud-browser .actions .primary');
  await page.waitForFunction(() => window.__atlasDirty === false && document.querySelector('.titlebar .docname input').value === 'Copy.xlsx', null, { timeout: 10000 });
  check('Save to Cloud asks before replacing, then saves a new file', !!S.dav['/remote.php/dav/files/u/Copy.xlsx']);
  await page.click('.titlebar [data-cmd=share]');
  await page.click('.modal .actions .primary');
  await page.waitForSelector('.modal p:has-text("cloud.test")');
  const sh = S.shares.find(s => s.dav);
  check('Nextcloud public share via OCS', sh && sh.dav === '/Copy.xlsx' && sh.type === '3' && sh.perms === '1' && sh.api === 'true', JSON.stringify(sh));
  await page.click('.modal .actions button:has-text("Close")');

  // ---- Share on an unsaved local document asks to save to cloud first
  await page.goto(base + 'slides.html');
  await page.waitForFunction(() => window.Cloud);
  await page.click('.titlebar [data-cmd=share]');
  const ask = await page.textContent('.modal p');
  check('sharing a local document offers Save to cloud', ask.includes('save it to a cloud account first'), ask);
  await page.click('.modal .actions button:has-text("Cancel")');

  // ---- Disconnect
  await page.goto(base + 'index.html');
  page.once('dialog', d => d.accept());
  await page.click('#cloud .row:nth-child(1) button:has-text("Disconnect")');
  await page.click('.modal .ok');
  await page.waitForFunction(() => document.querySelectorAll('#cloud .row').length === 3);
  check('account disconnected', true);
  // ---- Built-in app registrations (cloud-apps.json): nothing to set up
  const ctx2 = await browser.newContext();
  await ctx2.addInitScript(c => { localStorage.setItem('atlas-cloud-config', c); }, JSON.stringify({ endpoints: cfg.endpoints }));
  await ctx2.route('**/src/renderer/cloud-apps.json', r => r.fulfill({ contentType: 'application/json', body: JSON.stringify({ dropbox: { clientId: 'did' } }) }));
  const p2 = await ctx2.newPage();
  p2.on('pageerror', e => errors.push(e.message));
  await p2.goto(base + 'index.html');
  await p2.click('#cloudAdd');
  await p2.click('.modal .actions button:has-text("Dropbox")');
  await p2.waitForSelector('#cloud .row', { timeout: 10000 });
  check('built-in app: Dropbox connects without asking for an app key', (await p2.$$eval('#cloud .row .p', ns => ns.map(n => n.textContent))).includes('ada@dropbox.test'));
  await p2.evaluate(() => { Cloud.settingsDialog('dropbox'); });
  await p2.waitForSelector('.modal input');
  const ph = await p2.$eval('.modal input', i => ({ v: i.value, p: i.placeholder }));
  check('built-in app: settings show it is built in, and keep your own empty', ph.v === '' && /Built in/.test(ph.p), JSON.stringify(ph));
  await p2.click('.modal .actions .primary');
  check('built-in app: saving settings does not copy it into yours', await p2.evaluate(() => !JSON.parse(localStorage.getItem('atlas-cloud-config')).dropbox.clientId));
  await ctx2.close();
  check('no page errors', errors.length === 0, errors.join(' | '));
  await browser.close();
  server.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
