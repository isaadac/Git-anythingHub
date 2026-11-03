// Password-protected documents: protecting and saving in Write (an encrypted .docx), opening it
// again with a wrong and then the right password, and opening an .xlsx protected by other software
// in Sheets. Nothing of a protected document goes into thumbnails.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const ROOT = process.cwd();
const OfficeCrypto = require(path.join(ROOT, 'src/renderer/js/office-crypto.js'));
const JSZip = require(path.join(ROOT, 'node_modules/jszip'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ acceptDownloads: true });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.waitForFunction(() => window.AtlasWrite && window.OfficeCrypto);
  await p.click('#doc');
  await p.keyboard.type('Quarterly salaries: confidential');

  // ---- Protect, then save
  await p.evaluate(() => Platform.emit('protect'));
  await p.waitForSelector('.modal');
  const pw = await p.$$('.modal input[type=password]');
  await pw[0].fill('Hunter-2025'); await pw[1].fill('Hunter-2025');
  await p.click('.modal .actions .primary');
  await sleep(200);
  check('File → Protect with Password marks the document as protected', await p.isVisible('.protect-badge'));
  const [download] = await Promise.all([
    p.waitForEvent('download', { timeout: 30000 }),
    (async () => { await p.evaluate(() => Platform.emit('save')); await p.waitForSelector('.modal input'); await p.fill('.modal input', 'Salaries.docx'); await p.click('.modal .actions .primary'); })()
  ]);
  const file = path.join(process.env.E2E_WORK, 'Salaries.docx');
  await download.saveAs(file);
  const saved = new Uint8Array(fs.readFileSync(file));
  check('saving writes an encrypted file', OfficeCrypto.isEncrypted(saved) && !Buffer.from(saved).toString('latin1').includes('confidential'));
  const plain = await OfficeCrypto.decrypt(saved, 'Hunter-2025');
  const xml = await (await JSZip.loadAsync(plain)).file('word/document.xml').async('string');
  check('…which, with the password, is the document', /Quarterly salaries: confidential/.test(xml));
  check('…with no thumbnail of its content inside', !(await JSZip.loadAsync(plain)).file('docProps/thumbnail.jpeg'));
  check('no thumbnail of it is kept for Home either', await p.evaluate(() => !Object.values(JSON.parse(localStorage.getItem('atlas-thumbs') || '{}')).some(t => JSON.stringify(t).includes('confidential'))));

  // ---- Open it again: wrong password, then the right one
  await p.evaluate(() => { document.getElementById('doc').innerHTML = '<p>other</p>'; Platform.protection.password = null; window.__atlasDirty = false; });
  const opening = p.evaluate(b => AtlasWrite.loadFile({ name: 'Salaries.docx', data: new Uint8Array(b) }), [...saved]);
  await p.waitForSelector('.modal input[type=password]');
  check('opening a protected file asks for its password', /protected with a password/.test(await p.textContent('.modal')));
  await p.fill('.modal input[type=password]', 'hunter-2025');
  await p.click('.modal .actions .primary');
  await p.waitForFunction(() => /isn’t right/.test((document.querySelector('.modal') || {}).textContent || ''), null, { timeout: 15000 }).catch(() => {});
  check('a wrong password is refused, and it asks again', /isn’t right/.test(await p.textContent('.modal')));
  await p.fill('.modal input[type=password]', 'Hunter-2025');
  await p.click('.modal .actions .primary');
  await opening;
  check('…the right one opens it, still protected', /Quarterly salaries: confidential/.test(await p.textContent('#doc')) && await p.isVisible('.protect-badge'));
  // Cancelling leaves things as they were.
  const cancel = p.evaluate(b => AtlasWrite.loadFile({ name: 'Salaries.docx', data: new Uint8Array(b) }), [...saved]);
  await p.waitForSelector('.modal input[type=password]');
  await p.click('.modal .actions button:not(.primary)');
  await cancel;
  check('cancelling the password leaves the open document as it was', /Quarterly salaries/.test(await p.textContent('#doc')));
  // Removing protection.
  await p.evaluate(() => Platform.emit('protect'));
  await p.waitForSelector('.modal');
  await p.click('.modal .actions .primary');
  await sleep(200);
  check('protection can be removed', !(await p.isVisible('.protect-badge')) && await p.evaluate(() => Platform.protection.password === null));
  await p.evaluate(() => { window.__atlasDirty = false; });

  // ---- Sheets: a workbook protected by other software
  const s = await ctx.newPage();
  s.on('pageerror', e => errors.push(e.message));
  await s.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await s.waitForFunction(() => window.AtlasSheets && window.OfficeCrypto);
  const theirs = [...fs.readFileSync(path.join(ROOT, 'test/fixtures/protected-msoffcrypto.xlsx'))];
  const openS = s.evaluate(b => AtlasSheets.loadFile({ name: 'Budget.xlsx', data: new Uint8Array(b) }), theirs);
  await s.waitForSelector('.modal input[type=password]');
  await s.fill('.modal input[type=password]', 'Secret-42');
  await s.click('.modal .actions .primary');
  await openS;
  const a1 = await s.evaluate(() => AtlasSheets.rawOf('A1'));
  check('Sheets opens a workbook protected by other software', a1 !== '' && await s.isVisible('.protect-badge'), JSON.stringify(a1));
  await s.evaluate(() => { window.__atlasDirty = false; });

  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
