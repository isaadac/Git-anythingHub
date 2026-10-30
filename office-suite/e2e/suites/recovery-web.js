const { chromium } = require('playwright');
const BASE = process.env.E2E_BASE + '/src/renderer/';
(async () => {
  const b = await chromium.launch(); const p = await b.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message));
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x ? ' :: ' + x : ''));
  for (const [app, type] of [['write', async () => { await p.click('#doc'); await p.keyboard.type('Unsaved words here'); }],
                             ['sheets', async () => { await p.click('#grid tr[data-r="0"] td[data-c="0"]'); await p.keyboard.type('42'); await p.keyboard.press('Enter'); }],
                             ['slides', async () => { await p.click('#stage .el.text >> nth=0'); await p.keyboard.type('Draft title'); await p.keyboard.press('Escape'); }]]) {
    await p.goto(BASE + app + '.html');
    await type();
    await p.evaluate(() => window.dispatchEvent(new Event('pagehide')));
    await p.waitForTimeout(300);
    await p.evaluate(() => { window.__atlasDirty = false; });
    await p.goto(BASE + 'index.html');
    await p.waitForTimeout(300);
    const rows = await p.$$eval('#recovered .row .n', ns => ns.map(n => n.textContent));
    check(`${app}: draft listed on Home after leaving without saving`, rows.length === 1, JSON.stringify(rows));
    await p.click('#recovered .row button.primary');
    await p.waitForLoadState('load'); await p.waitForTimeout(500);
    const content = await p.evaluate(app => app === 'write' ? document.getElementById('doc').textContent
      : app === 'sheets' ? String(AtlasSheets.valueOf('A1')) : AtlasSlides.pres.slides[0].elements[0].html, app);
    const dirty = await p.evaluate(() => window.__atlasDirty);
    check(`${app}: draft restored and marked unsaved`, /Unsaved words here|^42$|Draft title/.test(content) && dirty === true, content.slice(0, 60));
    // Saving clears the draft (web save downloads the file)
    p.once('download', d => d.cancel().catch(() => {}));
    await p.evaluate(async () => { UI.prompt = async (t, l, v) => v; Platform.emit('save'); });
    await p.waitForTimeout(800);
    await p.evaluate(() => { window.__atlasDirty = false; });
    await p.goto(BASE + 'index.html');
    await p.waitForTimeout(300);
    const after = await p.$$eval('#recovered .row', r => r.length);
    check(`${app}: saving removes the draft`, after === 0, String(after));
  }
  // AutoSave switch present and explains itself in the browser
  await p.goto(BASE + 'write.html');
  const sw = await p.evaluate(() => { const l = document.querySelector('.titlebar .autosave'); return { disabled: l.querySelector('input').disabled, title: l.title }; });
  check('AutoSave switch shows why it is unavailable on the web', sw.disabled && /desktop/.test(sw.title), JSON.stringify(sw));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e); process.exit(1); });
