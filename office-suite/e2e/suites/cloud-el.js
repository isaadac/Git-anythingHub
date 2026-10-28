const { _electron: electron } = require('playwright');
const path = require('path'), fs = require('fs');
const ROOT = process.cwd();
const SP = process.env.E2E_WORK;
const XLSX = require(path.join(ROOT, 'node_modules/xlsx'));
const D = require(path.join(ROOT, 'node_modules/docx'));
const JSZip = require(path.join(ROOT, 'node_modules/jszip'));
const start = require('../cloud-mock');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const xlsx = rows => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sheet1'); return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }); };

(async () => {
  const report = await D.Packer.toBuffer(new D.Document({ sections: [{ children: [new D.Paragraph('Desktop drive doc')] }] }));
  const { port, S, server } = await start(ROOT, { 'google/Report.docx': report, __nativeXlsx: xlsx([['n']]), 'dav/Budget.xlsx': xlsx([['dav', 1]]) });
  const mk = p => `http://127.0.0.1:${port}/mock/${p}`;
  const ud = path.join(SP, 'ud-cloud-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow();
  await sleep(2000);
  // The "browser": follow the sign-in page's redirect to the loopback address.
  await app.evaluate(({ shell }) => { global.__opened = []; shell.openExternal = async url => { global.__opened.push(url); fetch(url).catch(() => {}); }; });
  const blocked = await home.evaluate(() => window.atlas.cloudFetch({ url: 'https://evil.example/steal' }));
  check('fetch proxy refuses hosts it does not know', blocked.status === 0 && /evil\.example/.test(blocked.error), JSON.stringify(blocked));
  const blocked2 = await home.evaluate(p => window.atlas.cloudFetch({ url: `http://127.0.0.1:${p}/mock/google/api/about` }), port);
  check('test server refused before it is configured', blocked2.status === 0);
  await home.evaluate(c => window.atlas.cloudConfigSet(c), { google: { clientId: 'gid' },
    endpoints: { google: { auth: mk('google/auth'), token: mk('google/token'), api: mk('google/api'), upload: mk('google/upload') }, evil: { a: 'https://evil.example' } } });
  const cfg = await home.evaluate(() => window.atlas.cloudConfigGet());
  check('config cleaned (unknown providers dropped)', !cfg.endpoints.evil && cfg.google.clientId === 'gid', JSON.stringify(cfg.endpoints));

  await home.click('#cloudAdd');
  await home.click('.modal .actions button:has-text("Google Drive")');
  await home.waitForSelector('#cloud .row', { timeout: 15000 });
  const opened = await app.evaluate(() => global.__opened);
  const au = new URL(opened[0]);
  check('sign-in opened in the system browser with a loopback redirect', /^http:\/\/127\.0\.0\.1:\d+\/callback$/.test(au.searchParams.get('redirect_uri')) && au.searchParams.get('code_challenge_method') === 'S256', opened[0]);
  check('account connected', (await home.textContent('#cloud .row .p')) === 'ada@gmail.test');
  const tf = path.join(ud, 'cloud-accounts.bin');
  const raw = fs.readFileSync(tf);
  const enc = await app.evaluate(({ safeStorage }) => safeStorage.isEncryptionAvailable());
  check('tokens file is private' + (enc ? ' and encrypted' : ''), (fs.statSync(tf).mode & 0o077) === 0 && (!enc || !raw.toString().includes('refresh-google')), (fs.statSync(tf).mode & 0o777).toString(8));
  const used = await app.evaluate(async () => {
    const http = process.mainModule.require('http');
    return await new Promise(r => { http.get(global.__opened[0].replace(/^.*redirect_uri=([^&]+).*$/, (m, x) => decodeURIComponent(x)) + '?code=x&state=y', res => r(res.statusCode)).on('error', () => r('closed')); });
  });
  check('loopback listener closes after one sign-in', used === 'closed', used);

  // Open from cloud: a new Write window
  await home.evaluate(() => { Cloud.openDialog({}); });
  await home.dblclick('.cloud-row:has-text("Report.docx")');
  let write;
  for (let i = 0; i < 40 && !write; i++) { await sleep(250); write = app.windows().find(w => w.url().includes('write.html')); }
  await write.waitForFunction(() => document.querySelector('#doc').innerText.includes('Desktop drive doc'), null, { timeout: 15000 });
  check('cloud file opened in a new Write window', true);
  const sw = await write.evaluate(() => ({ d: document.querySelector('.autosave input').disabled, path: null }));
  check('AutoSave available', !sw.d);
  await write.click('#doc'); await write.keyboard.type('Saved from desktop ');
  await write.evaluate(() => { Platform.emit('save'); });
  await write.waitForFunction(() => window.__atlasDirty === false, null, { timeout: 10000 });
  const g0 = Object.values(S.google.files).find(f => f.name === 'Report.docx');
  const txt = (await (await JSZip.loadAsync(g0.data)).file('word/document.xml').async('string')).replace(/<[^>]+>/g, '');
  check('Ctrl+S saved back to Drive from the desktop', g0.saves === 1 && txt.includes('Saved from desktop'), g0.saves);
  check('no local file written for the cloud doc', !fs.existsSync(path.join(ROOT, 'Report.docx')));

  // WebDAV on desktop (host allowed once the account is stored)
  await home.click('#cloudAdd');
  await home.click('.modal .actions button:has-text("WebDAV")');
  await home.fill('.modal input >> nth=0', `http://127.0.0.1:${port}/remote.php/dav/files/u/`);
  await home.fill('.modal input >> nth=1', 'u');
  await home.fill('.modal input >> nth=2', 'secret');
  await home.click('.modal .actions .primary');
  await home.waitForFunction(() => document.querySelectorAll('#cloud .row').length === 2, null, { timeout: 10000 });
  check('WebDAV account connected on desktop', true);
  const listed = await home.evaluate(async () => { const a = (await Cloud.accounts()).find(x => x.provider === 'webdav'); return (await Cloud.list(a.id)).map(i => i.name); });
  check('WebDAV listing through the main process', listed.join() === 'Budget.xlsx', listed.join());

  await require('../close-app')(app);
  server.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
