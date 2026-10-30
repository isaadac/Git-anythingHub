// Suite-wide tools: version history, command search (Alt+Q) and the interface languages.
const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.clock.install();
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.waitForFunction(() => window.AtlasWrite && window.I18n);
  const submit = async fill => { await p.waitForSelector('.modal'); if (fill) await fill(); await p.click('.modal .actions .primary'); await p.waitForSelector('.modal', { state: 'detached' }); };
  const setText = t => p.evaluate(t => { const d = document.getElementById('doc'); d.innerHTML = `<p>${t}</p>`; d.dispatchEvent(new Event('input', { bubbles: true })); }, t);
  const docText = () => p.evaluate(() => document.getElementById('doc').innerText.trim());
  const save = async () => {
    const dl = p.waitForEvent('download');
    await p.evaluate(() => Platform.emit('save'));
    await submit(() => p.fill('.modal input', 'Report.docx'));
    await dl;
    await p.waitForFunction(async () => (await Platform.versions.list('Report.docx')).length > 0);
  };

  // ---- Version history: one version per 10 minutes, restore an older one
  await setText('First draft');
  await save();
  await setText('First draft, fixed');
  await save();
  let list = await p.evaluate(() => Platform.versions.list('Report.docx'));
  check('saves within 10 minutes keep one version (the latest)', list.length === 1, list.length);
  await p.clock.fastForward('11:00');
  await setText('Second draft');
  await save();
  await p.clock.fastForward('11:00');
  await setText('Third draft');
  await save();
  list = await p.evaluate(() => Platform.versions.list('Report.docx'));
  check('a save more than 10 minutes later adds a version', list.length === 3, list.length);
  await p.evaluate(() => Platform.emit('version-history'));
  await p.waitForSelector('.modal.versions .v-row');
  const rows = await p.$$eval('.modal.versions .v-row .when', r => r.map(x => x.textContent));
  check('Version History lists them, newest first', rows.length === 3 && /latest/.test(rows[0]), rows.join(' | '));
  await p.click('.modal.versions .v-row:nth-child(3) button[data-a="restore"]');
  await submit();
  await p.waitForFunction(() => /First draft, fixed/.test(document.getElementById('doc').innerText));
  check('restoring brings the older text back, unsaved', (await docText()) === 'First draft, fixed' && (await p.evaluate(() => window.__atlasDirty)), await docText());

  // ---- Command search
  await setText('Hello');
  await p.evaluate(() => { const d = document.getElementById('doc'); d.focus(); const r = document.createRange(); r.selectNodeContents(d); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await p.keyboard.press('Alt+q');
  check('Alt+Q goes to the search box', await p.evaluate(() => document.activeElement && document.activeElement.closest('.tell-me') !== null));
  await p.keyboard.type('page break');
  await p.waitForSelector('.tell-me .tm-item');
  const first = await p.$eval('.tell-me .tm-item', r => [r.querySelector('.l').textContent, r.querySelector('.w').textContent]);
  check('it finds the command and says where it is', first[0] === 'Page Break' && /^Insert/.test(first[1]), first.join(' / '));
  await p.keyboard.press('Enter');
  await sleep(200);
  check('Enter runs it (a page break in the document)', await p.evaluate(() => !!document.querySelector('#doc .page-break')));
  check('…on the Insert tab', await p.evaluate(() => document.querySelector('.ribbon-tabs button.active').dataset.tab === 'insert'));
  await p.keyboard.press('Alt+q');
  await p.keyboard.type('version');
  await p.waitForSelector('.tell-me .tm-item');
  check('File commands are found too', await p.$eval('.tell-me .tm-item', r => r.querySelector('.l').textContent === 'Version History' && r.querySelector('.w').textContent === 'File'));
  await p.keyboard.press('Escape');

  // ---- Languages
  await setText('Save the Table of Contents');
  await p.evaluate(() => I18n.set('es'));
  await p.waitForFunction(() => document.querySelector('.ribbon-tabs button[data-tab="insert"]').textContent.trim() === 'Insertar');
  const es = await p.evaluate(() => ({
    home: document.querySelector('.ribbon-tabs button[data-tab="home"]').textContent.trim(),
    pb: [...document.getElementById('insPageBreak').childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join(' '),
    br: !!document.querySelector('#insPageBreak br'),
    search: document.querySelector('.tell-me input').placeholder,
    lang: document.documentElement.lang,
    doc: document.getElementById('doc').innerText.trim()
  }));
  check('Spanish: tabs, big buttons (still on two lines) and placeholders', es.home === 'Inicio' && es.pb === 'Salto de página' && es.br && es.search === 'Buscar (Alt+Q)' && es.lang === 'es', JSON.stringify(es));
  check('…but never the document', es.doc === 'Save the Table of Contents', es.doc);
  await p.evaluate(() => UI.toast('Saved'));
  await sleep(100);
  check('messages that appear later are translated too', (await p.$eval('.toast', t => t.textContent)) === 'Guardado');
  await p.evaluate(() => Platform.emit('language'));
  await p.waitForSelector('.modal');
  const dlg = await p.$eval('.modal', m => m.textContent);
  check('the Language dialog (in Spanish, with each language in its own words)', /Idioma/.test(dlg) && /Deutsch/.test(dlg) && /Français/.test(dlg), dlg.slice(0, 120));
  await p.click('.modal button:has-text("Deutsch")');
  await p.waitForFunction(() => document.querySelector('.ribbon-tabs button[data-tab="insert"]').textContent.trim() === 'Einfügen');
  check('switching live to German', await p.evaluate(() => document.querySelector('.ribbon-tabs button[data-tab="review"]').textContent.trim() === 'Überprüfen' && document.documentElement.lang === 'de'));

  // Other apps open in the chosen language from the start; their documents stay as they are.
  const s = await ctx.newPage();
  s.on('pageerror', e => errors.push(e.message));
  await s.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await s.waitForFunction(() => window.AtlasSheets);
  await s.evaluate(() => AtlasSheets.setCell('A1', 'Delete'));
  await sleep(200);
  const sh = await s.evaluate(() => ({ data: document.querySelector('.ribbon-tabs button[data-tab="data"]').textContent.trim(), sort: document.querySelector('.ribbon').textContent.includes('Sortieren und Filtern'), lang: document.documentElement.lang, cell: document.getElementById('grid').textContent.includes('Delete') && !document.getElementById('grid').textContent.includes('Löschen') }));
  check('a new Sheets window opens in German (cells untouched)', sh.data === 'Daten' && sh.sort && sh.lang === 'de' && sh.cell, JSON.stringify(sh));
  // Back to English in one window: the other follows.
  await p.evaluate(() => I18n.set('en'));
  await s.waitForFunction(() => document.querySelector('.ribbon-tabs button[data-tab="data"]').textContent.trim() === 'Data', null, { timeout: 5000 }).catch(() => {});
  check('back to English, in every window', await s.evaluate(() => document.querySelector('.ribbon-tabs button[data-tab="data"]').textContent.trim() === 'Data') && await p.evaluate(() => [...document.getElementById('insPageBreak').childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join(' ') === 'Page Break'));

  check('no page errors', !errors.length, errors.join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
