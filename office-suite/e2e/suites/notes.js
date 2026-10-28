const { chromium } = require('playwright');
const fs = require('fs');
const AXE = fs.readFileSync(require.resolve('axe-core/axe.min.js'), 'utf8');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const URL = process.env.E2E_BASE + '/src/renderer/notes.html';
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ acceptDownloads: true, bypassCSP: true });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(URL);
  await p.waitForFunction(() => window.AtlasNotes && AtlasNotes.page);
  check('a first notebook is created', (await p.evaluate(() => [AtlasNotes.notebook.name, AtlasNotes.notebook.sections[0].name].join('/'))) === 'My Notebook/Quick Notes');
  // type a page
  await p.fill('#pageTitle', 'Shopping');
  await p.click('#body');
  await p.keyboard.type('Milk');
  await p.keyboard.press('Control+1');
  await p.keyboard.press('Enter');
  await p.keyboard.type('Bread');
  await p.keyboard.press('Enter');
  await p.keyboard.press('Control+1');
  await p.keyboard.type('Call the bakery');
  await p.keyboard.press('Control+2');
  const html = await p.$eval('#body', d => d.innerHTML);
  check('to-do and Important tags', (html.match(/class="todo"/g) || []).length === 2 && /tag-important/.test(html), html);
  // tick Milk
  const box = await (await p.$('#body .todo')).boundingBox();
  await p.mouse.click(box.x + 8, box.y + 10);
  check('ticking a to-do', /class="todo done"/.test(await p.$eval('#body', d => d.innerHTML)) && /1\/2 to-dos/.test(await p.textContent('#stCount')), await p.textContent('#stCount'));
  check('page title shows in the list', (await p.textContent('#pages li.on .t')) === 'Shopping');
  // second page + section
  await p.click('#pageAdd');
  await p.fill('#pageTitle', 'Ideas');
  await p.click('#body'); await p.keyboard.type('A notes app with ink and search');
  await p.click('#secAdd'); await p.fill('.modal input', 'Work'); await p.click('.modal .actions .primary');
  await p.fill('#pageTitle', 'Meeting'); await p.click('#body'); await p.keyboard.type('Budget review with Ada on Friday');
  // ink
  await p.click('[data-tab="draw"]');
  await p.click('.ribbon .pen[data-pen="#c62828"]');
  const pg = await (await p.$('#page')).boundingBox();
  await p.mouse.move(pg.x + 100, pg.y + 200); await p.mouse.down();
  for (let i = 1; i < 8; i++) await p.mouse.move(pg.x + 100 + i * 15, pg.y + 200 + i * 6);
  await p.mouse.up();
  await p.mouse.move(pg.x + 100, pg.y + 300); await p.mouse.down(); await p.mouse.move(pg.x + 200, pg.y + 300, { steps: 5 }); await p.mouse.up();
  check('drawing with a pen', (await p.evaluate(() => AtlasNotes.page.ink.length)) === 2 && (await p.$$('#ink path')).length === 2);
  await p.click('#eraser');
  await p.mouse.move(pg.x + 150, pg.y + 300); await p.mouse.down(); await p.mouse.up();
  check('the eraser removes a stroke', (await p.evaluate(() => AtlasNotes.page.ink.length)) === 1);
  await p.click('#drawOff');
  await sleep(900);
  // search
  await p.fill('#search', 'bakery');
  await sleep(300);
  const res = await p.$$eval('#results li', ls => ls.map(l => l.textContent));
  check('search finds the page', res.length === 1 && /Shopping/.test(res[0]) && /Quick Notes/.test(res[0]), res.join(' | '));
  await p.click('#results li');
  check('opening a result selects the match', (await p.evaluate(() => getSelection().toString())) === 'bakery' && (await p.inputValue('#pageTitle')) === 'Shopping');
  await p.fill('#search', '');
  await p.dispatchEvent('#search', 'input'); await sleep(200);
  // persistence
  await p.reload();
  await p.waitForFunction(() => window.AtlasNotes && AtlasNotes.page);
  const nb = await p.evaluate(() => AtlasNotes.notebook);
  check('everything is kept after reloading', nb.sections.length === 2 && nb.sections[0].pages.length === 2 && nb.sections[1].pages[0].ink.length === 1 && /Budget review/.test(nb.sections[1].pages[0].html), JSON.stringify(nb.sections.map(s => s.pages.map(p => p.title))));
  check('reopens where you left off', (await p.inputValue('#pageTitle')) === 'Shopping');
  // markdown
  const md = await p.evaluate(() => AtlasNotes.toMarkdown(AtlasNotes.page.html));
  check('Markdown export with task list', /- \[x\] Milk/.test(md) && /- \[ \] Bread/.test(md) && /★ Call the bakery/.test(md), md);
  // export .anote and re-import as a copy
  const dl = p.waitForEvent('download');
  await p.evaluate(() => Platform.emit('save-as'));
  await p.waitForSelector('.modal input'); await p.click('.modal .actions .primary');
  const file = await (await dl).path();
  const anote = JSON.parse(fs.readFileSync(file, 'utf8'));
  check('notebook exported as .anote', anote.format === 'atlas-notes' && anote.sections.length === 2);
  // import with a malicious page
  anote.sections[0].pages[0].html += '<img src=x onerror="window.__pwned=1"><script>window.__pwned=2</script>';
  fs.writeFileSync(process.env.E2E_WORK + '/test.anote', JSON.stringify(anote));
  await p.click('.file-tab');
  const [chooser] = await Promise.all([p.waitForEvent('filechooser'), p.click('.backstage [data-cmd="open"]')]);
  await chooser.setFiles(process.env.E2E_WORK + '/test.anote');
  await p.waitForSelector('.modal p:has-text("already")');
  await p.click('.modal .actions button:has-text("Keep both")');
  await sleep(600);
  check('imported as a copy', (await p.$$eval('#nbSelect option', os => os.map(o => o.textContent))).some(t => /copy/.test(t)));
  check('scripts in an imported notebook are removed', await p.evaluate(() => !window.__pwned && !document.querySelector('#body [onerror], #body script')));
  // open in Write
  await p.evaluate(() => Platform.emit('open-in-write'));
  await p.waitForURL(/write\.html/);
  await p.waitForFunction(() => document.getElementById('doc').innerText.includes('Milk'));
  check('page opens in Write with its to-dos as text', /☑ Milk/.test(await p.$eval('#doc', d => d.innerText)));
  await p.goto(URL); await p.waitForFunction(() => window.AtlasNotes && AtlasNotes.page);
  await p.addScriptTag({ content: AXE });
  const v = await p.evaluate(async () => (await axe.run(document, { resultTypes: ['violations'] })).violations.map(x => x.id + '×' + x.nodes.length + ' ' + x.nodes[0].target.join(' ')));
  check('no axe violations', v.length === 0, v.join('; '));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
