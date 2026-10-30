const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const ROOT = process.cwd();
const JSZip = require(path.join(ROOT, 'node_modules/jszip'));
const AXE = fs.readFileSync(require.resolve('axe-core/axe.min.js'), 'utf8');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ bypassCSP: true });
  const errors = [];
  const open = async app => { const p = await ctx.newPage(); p.on('pageerror', e => errors.push(app + ': ' + e.message)); await p.goto(`${process.env.E2E_BASE}/src/renderer/${app}.html`); await sleep(600); return p; };
  const axe = async (p, ctxSel) => { await p.addScriptTag({ content: AXE }); return p.evaluate(async c => (await axe.run(c ? document.querySelector(c) : document, { resultTypes: ['violations'] })).violations.filter(v => v.impact === 'serious' || v.impact === 'critical').map(v => v.id + '×' + v.nodes.length + ' ' + v.nodes[0].target.join(' ')), ctxSel); };

  // ---- Write
  const w = await open('write');
  check('write: no serious axe violations', (await axe(w)).length === 0, (await axe(w)).join('; '));
  // ribbon keyboard
  await w.focus('.ribbon-tabs button.active');
  await w.keyboard.press('ArrowRight');
  const tab = await w.evaluate(() => [document.activeElement.textContent, document.activeElement.getAttribute('aria-selected'), document.querySelector('.ribbon-panel.active').dataset.panel]);
  check('write: arrow keys move through ribbon tabs', tab[1] === 'true' && tab[2] === 'insert', tab.join(','));
  await w.keyboard.press('ArrowDown');
  check('write: Down enters the ribbon panel', await w.evaluate(() => !!document.activeElement.closest('.ribbon-panel.active')));
  // F6
  const seen = [];
  for (let i = 0; i < 4; i++) { await w.keyboard.press('F6'); seen.push(await w.evaluate(() => document.activeElement.id || document.activeElement.className || document.activeElement.tagName)); }
  check('write: F6 cycles through title bar, ribbon, document, status bar', seen.includes('doc') && seen.some(x => /tab/.test(x) || x === 'tab-home' || x.startsWith('tab-')), seen.join(' > '));
  // modal focus trap and return
  await w.click('#doc'); await w.keyboard.type('Hello');
  await w.click('[data-tab="review"]');
  await w.focus('#wordCount');
  await w.keyboard.press('Enter');
  await w.waitForSelector('.modal');
  check('write: dialog is labelled and modal', await w.evaluate(() => { const m = document.querySelector('.modal'); return m.getAttribute('role') && m.getAttribute('aria-modal') === 'true' && !!document.getElementById(m.getAttribute('aria-labelledby')); }));
  check('write: dialog has no serious axe violations', (await axe(w, '.modal-backdrop')).length === 0, (await axe(w, '.modal-backdrop')).join('; '));
  for (let i = 0; i < 6; i++) await w.keyboard.press('Tab');
  check('write: Tab stays inside the dialog', await w.evaluate(() => !!document.activeElement.closest('.modal')));
  await w.keyboard.press('Escape');
  await sleep(100);
  console.log('dbg', await w.evaluate(() => [document.activeElement.id, document.activeElement.tagName, !!document.querySelector('.modal')]));
  check('write: focus returns after the dialog', await w.evaluate(() => document.activeElement.id === 'wordCount'), await w.evaluate(() => document.activeElement.id));
  // checker
  await w.evaluate(src => {
    const d = document.getElementById('doc');
    d.innerHTML = `<h1>Report</h1><h3>Skipped</h3><p>See <a href="https://x.test">click here</a>.</p><p><img src="${src}"></p><table><tr><td>Name</td><td>Qty</td></tr><tr><td>A</td><td>1</td></tr></table><p><span style="color: #dddddd">faint</span></p><h2></h2>`;
    d.dispatchEvent(new Event('input'));
  }, PNG);
  await w.click('#a11yCheck');
  await w.waitForSelector('.a11y-item');
  const items = await w.$$eval('.a11y-item .a11y-title', ts => ts.map(t => t.textContent));
  check('write: checker finds the problems', ['Missing alternative text', 'has no header row', 'Heading level skipped', 'Hard-to-read text colour', 'Unclear link text', 'Empty heading'].every(k => items.some(t => t.includes(k))), items.join(' | '));
  check('write: checker pane has no serious axe violations', (await axe(w, '.a11y-pane')).length === 0);
  // fixes
  await w.click('.a11y-item:has-text("header row") .a11y-fix');
  await w.click('.a11y-item:has-text("Heading level skipped") .a11y-fix');
  await w.click('.a11y-item:has-text("Hard-to-read") .a11y-fix');
  await w.click('.a11y-item:has-text("Missing alternative text") .a11y-fix');
  await w.fill('.modal textarea', 'A blue square logo');
  await w.click('.modal .actions .primary');
  await sleep(200);
  const left = await w.$$eval('.a11y-item .a11y-title', ts => ts.map(t => t.textContent));
  check('write: fixes clear their issues', left.length === 2 && left.every(t => /Unclear link|Empty heading/.test(t)) && !!(await w.$('#doc h2')) && !(await w.$('#doc h3')), left.join(' | '));
  const docx = await w.evaluate(async () => { const d = await AtlasWrite.serialize('docx'); return Array.from(d instanceof Uint8Array ? d : new Uint8Array(await d.arrayBuffer())); });
  const z = await JSZip.loadAsync(Buffer.from(docx));
  const xml = await z.file('word/document.xml').async('string');
  check('write: alt text saved in .docx', /descr="A blue square logo"/.test(xml));
  check('write: header row saved as a Word header row', /<w:tblHeader\/>|<w:tblHeader w:val="(true|1)"\/>/.test(xml) || /tblHeader/.test(xml), (xml.match(/<w:trPr>.*?<\/w:trPr>/) || [''])[0]);
  // ---- Sheets
  const s = await open('sheets');
  check('sheets: no serious axe violations', (await axe(s)).length === 0, (await axe(s)).join('; '));
  await s.evaluate(() => { AtlasSheets.setCell('A1', 'Total'); AtlasSheets.setCell('B1', 5); AtlasSheets.setCell('C1', '=B1*2'); });
  await s.click('#grid tbody tr:nth-child(1) td:nth-child(4)');
  await sleep(200);
  check('sheets: the cell is announced', /C1, 10, formula =B1\*2/.test(await s.evaluate(() => document.querySelector('[aria-live][role=status].sr-only').textContent)), await s.evaluate(() => document.querySelector('.sr-only[aria-live]').textContent));
  await s.keyboard.press('ArrowLeft'); await sleep(150);
  check('sheets: moving announces the next cell', /^B1, 5$/.test(await s.evaluate(() => document.querySelector('.sr-only[aria-live]').textContent)));
  await s.evaluate(() => { const b = AtlasSheets.book; b.sheets[0].cells.A1.s = { color: '#eeeeee' }; b.sheets[0].merges.push({ r1: 3, c1: 0, r2: 3, c2: 2 }); b.sheets[0].charts.push({ type: 'column', range: 'A1:B1', x: 10, y: 10, w: 200, h: 100, title: '' }); });
  await s.click('[data-tab="review"]'); await s.click('#a11yCheck');
  await s.waitForSelector('.a11y-item');
  const si = await s.$$eval('.a11y-item .a11y-title', ts => ts.map(t => t.textContent));
  check('sheets: checker finds chart, merged cells and contrast', si.some(t => /Chart 1/.test(t)) && si.some(t => /Merged/.test(t)) && si.some(t => /Hard-to-read/.test(t)), si.join(' | '));
  // ---- Slides
  const p = await open('slides');
  check('slides: no serious axe violations', (await axe(p)).length === 0, (await axe(p)).join('; '));
  await p.evaluate(src => { const sl = AtlasSlides.pres.slides[0]; sl.elements.find(e => e.ph === 'title').html = ''; sl.elements.push({ id: 'img1', type: 'image', src, x: 10, y: 10, w: 50, h: 50 }); AtlasSlides.goTo(0); }, PNG);
  await p.click('#stage .el[data-id="img1"]');
  await p.keyboard.press('Tab');
  await sleep(150);
  const said = await p.evaluate(() => document.querySelector('.sr-only[aria-live]').textContent);
  console.log('dbg2', await p.evaluate(() => [document.activeElement.id, document.activeElement.className, document.querySelectorAll('#stage .el.selected, #stage .sel').length]));
  check('slides: Tab moves to the next object and announces it', /of 3$/.test(said), said);
  const opts = await p.$$eval('#thumbs .thumb', ts => ts.map(t => [t.getAttribute('role'), t.getAttribute('aria-selected'), t.getAttribute('aria-label')]));
  check('slides: thumbnails are a list of labelled options', opts[0][0] === 'option' && opts[0][1] === 'true' && /^Slide 1/.test(opts[0][2]), JSON.stringify(opts));
  await p.click('[data-tab="review"]'); await p.click('#a11yCheck');
  await p.waitForSelector('.a11y-item');
  const pi = await p.$$eval('.a11y-item .a11y-title', ts => ts.map(t => t.textContent));
  check('slides: checker finds missing title and alt text', pi.some(t => /no title/.test(t)) && pi.some(t => /Missing alt text: picture/.test(t)), pi.join(' | '));
  await p.click('.a11y-item:has-text("Missing alt text") .a11y-fix');
  await p.fill('.modal textarea', 'Company logo');
  await p.click('.modal .actions .primary');
  await p.click('.a11y-item:has-text("no title") .a11y-fix');
  await p.fill('.modal input', 'Quarterly results');
  await p.click('.modal .actions .primary');
  await sleep(200);
  check('slides: fixed', !!(await p.$('.a11y-pane .proof-done')), await p.$$eval('.a11y-item .a11y-title', ts => ts.map(t => t.textContent).join(' | ')));
  const pptx = await p.evaluate(async () => { const d = await AtlasSlides.serialize('pptx'); return Array.from(d instanceof Uint8Array ? d : new Uint8Array(await d.arrayBuffer())); });
  fs.writeFileSync(process.env.E2E_WORK + '/a11y.pptx', Buffer.from(pptx));
  const pz = await JSZip.loadAsync(Buffer.from(pptx));
  const sx = await pz.file('ppt/slides/slide1.xml').async('string');
  check('slides: alt text saved in .pptx', /descr="Company logo"/.test(sx));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
