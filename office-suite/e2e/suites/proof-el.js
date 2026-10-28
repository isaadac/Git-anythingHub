const { _electron: electron } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-proof-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow(); await sleep(1200);
  const win = async name => { let w; for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes(name)); } return w; };
  await home.evaluate(() => window.atlas.newWindow('write'));
  const w = await win('write.html');
  await w.waitForFunction(() => window.Proofing);
  const langs = await w.evaluate(() => window.atlas.spellLanguages());
  console.log('info languages', langs.current.join(','), langs.available.length);
  // wait for the dictionary
  let dict = false;
  for (let i = 0; i < 20 && !dict; i++) { dict = (await w.evaluate(() => window.atlas.spellCheck(['recieve', 'house']))).join() === 'true,false'; if (!dict) await sleep(500); }
  console.log('info dictionary loaded:', dict);
  await w.click('#doc');
  await w.keyboard.type('The the cat sat on a apple.  It was hapy. we left');
  await w.keyboard.press('F7');
  await w.waitForSelector('.proof-card');
  const sum = await w.textContent('.proof-sum');
  check('pane lists issues', /grammar/.test(sum), sum);
  // walk: fix everything with the first suggestion
  for (let i = 0; i < 12; i++) {
    if (!(await w.$('.proof-card'))) break;
    const b = await w.$('.proof-card .sugg button');
    if (b) await b.click(); else await w.click('.proof-card .acts button:has-text("Ignore once")');
    await sleep(100);
  }
  const text = await w.$eval('#doc', d => d.innerText.trim());
  const expect = dict ? 'The cat sat on an apple. It was happy. We left' : 'The cat sat on an apple. It was hapy. We left';
  check('fixes applied in the document', text.replace(/\s+/g, ' ') === expect, text);
  check('ends with no suggestions', !!(await w.$('.proof-done')));
  await w.keyboard.press('Control+z');
  check('undo works after a fix', (await w.$eval('#doc', d => d.innerText)).replace(/\s+/g, ' ').trim() !== expect);
  // language
  await w.click('#stLang');
  await w.waitForSelector('.modal select');
  const opts = await w.$$eval('.modal select option', os => os.map(o => o.value));
  check('language list offered', opts.length > 3, opts.slice(0, 6).join(','));
  const pick = opts.find(v => /^fr/.test(v)) || opts[1];
  await w.selectOption('.modal select', pick);
  await w.click('.modal .actions .primary');
  await sleep(500);
  const l = await w.evaluate(() => [document.getElementById('doc').lang, document.getElementById('stLang').textContent]);
  check('document language set and shown', l[0] === pick && /fran|French/i.test(l[1]), l.join(' / '));
  const html = await w.evaluate(() => AtlasWrite.serialize('html'));
  check('language saved with the document', html.includes('&quot;lang&quot;:&quot;' + pick));
  // context menu suggestions exist (main process handler registered)
  check('context menu handler registered', await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().every(x => x.webContents.listenerCount('context-menu') >= 1)));
  // Sheets
  await home.evaluate(() => window.atlas.newWindow('sheets'));
  const s = await win('sheets.html');
  await s.waitForFunction(() => window.AtlasSheets);
  await s.evaluate(() => { AtlasSheets.setCell('A1', 'Total  sales'); AtlasSheets.setCell('B2', 'an banana'); AtlasSheets.setCell('C3', '=A1'); });
  await s.click('[data-tab="review"]');
  await s.click('#proofEditor');
  await s.waitForSelector('.proof-card');
  for (let i = 0; i < 5 && await s.$('.proof-card .sugg button'); i++) { await s.click('.proof-card .sugg button'); await sleep(100); }
  check('sheets: cells corrected', (await s.evaluate(() => [AtlasSheets.valueOf('A1'), AtlasSheets.valueOf('B2')].join('|'))) === 'Total sales|a banana');
  await s.evaluate(() => Platform.emit('undo'));
  check('sheets: undo restores', (await s.evaluate(() => AtlasSheets.valueOf('B2'))) === 'an banana');
  // Slides
  await home.evaluate(() => window.atlas.newWindow('slides'));
  const p = await win('slides.html');
  await p.waitForFunction(() => window.AtlasSlides);
  await p.evaluate(() => { const s = AtlasSlides.pres.slides[0]; s.elements[0].html = 'Our  <b>plan</b> for for 2027'; s.notes = 'i will explain.'; });
  await p.click('[data-tab="review"]');
  await p.click('#proofEditor');
  await p.waitForSelector('.proof-card');
  for (let i = 0; i < 6 && await p.$('.proof-card .sugg button'); i++) { await p.click('.proof-card .sugg button'); await sleep(100); }
  const sl = await p.evaluate(() => [AtlasSlides.pres.slides[0].elements[0].html, AtlasSlides.pres.slides[0].notes]);
  check('slides: text box fixed, formatting kept', sl[0] === 'Our <b>plan</b> for 2027', sl[0]);
  check('slides: notes fixed', sl[1] === 'I will explain.', sl[1]);
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
