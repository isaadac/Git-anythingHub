const { chromium } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const { createCollabServer } = require(path.join(ROOT, 'src/main/collab-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const ONLY = process.env.E2E_ONLY || 'all';
const BASE = process.env.E2E_BASE + '/src/renderer/';

(async () => {
  const srv = createCollabServer({ limits: { idleMs: 3000 } });
  await new Promise(r => srv.listen(0, '127.0.0.1', r));
  const WS = `ws://127.0.0.1:${srv.address().port}/`;
  const browser = await chromium.launch();
  const errors = [];
  async function open(app, name) {
    const ctx = await browser.newContext();
    await ctx.addInitScript(n => { localStorage.setItem('atlas-user-name', n); }, name);
    await ctx.grantPermissions(['clipboard-read', 'clipboard-write']);
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(name + ': ' + e.message));
    await p.goto(BASE + app + '.html');
    await p.waitForFunction(() => window.Collab && document.querySelector('.live-btn'));
    return p;
  }
  async function host(p) {
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Start on a server")');
    await p.fill('.modal input', WS);
    await p.click('.modal .actions .primary');
    await p.waitForSelector('.modal p:has-text("room=")', { timeout: 10000 });
    const text = await p.textContent('.modal p');
    await p.click('.modal .actions button:has-text("Close")');
    return text.match(/ws:\/\/\S+/)[0];
  }
  async function join(p, link) {
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Join")');
    await p.fill('.modal input', link);
    await p.click('.modal .actions .primary');
    await p.waitForFunction(() => document.querySelector('.live-btn').textContent.startsWith('Live ·'), null, { timeout: 10000 });
  }
  const until = (p, fn, arg) => p.waitForFunction(fn, arg, { timeout: 8000 }).then(() => true, () => false);

  if (ONLY === 'all' || ONLY === 'sheets') {
    const a = await open('sheets', 'Ada Lovelace'), b = await open('sheets', 'Bob');
    await a.evaluate(() => { AtlasSheets.setCell('A1', 'hello'); AtlasSheets.setCell('A2', 5); AtlasSheets.setCell('A3', '=A2*2'); });
    const link = await host(a);
    check('invite link made', /room=[0-9a-f]{32}&app=sheets/.test(link), link);
    await join(b, link);
    check('joiner gets the document', await until(b, () => AtlasSheets.valueOf('A1') === 'hello' && AtlasSheets.valueOf('A3') === 10));
    check('both see two people', (await a.textContent('.live-btn')) === 'Live · 2' && (await b.textContent('.live-btn')) === 'Live · 2');
    check('avatar shows initials', (await a.textContent('.live-people')) === 'B' && (await b.textContent('.live-people')) === 'AL');
    await b.evaluate(() => AtlasSheets.setCell('A2', 21));
    check('edit flows to the other person and formulas recalc', await until(a, () => AtlasSheets.valueOf('A3') === 42));
    // simultaneous edits to different cells
    await Promise.all([a.evaluate(() => AtlasSheets.setCell('B1', 'from a')), b.evaluate(() => AtlasSheets.setCell('C1', 'from b'))]);
    check('concurrent edits to different cells both land', await until(a, () => AtlasSheets.valueOf('C1') === 'from b') && await until(b, () => AtlasSheets.valueOf('B1') === 'from a'));
    // same cell at once: converge
    await Promise.all([a.evaluate(() => AtlasSheets.setCell('D1', 'A wins?')), b.evaluate(() => AtlasSheets.setCell('D1', 'B wins?'))]);
    await sleep(1200);
    const d1 = [await a.evaluate(() => AtlasSheets.valueOf('D1')), await b.evaluate(() => AtlasSheets.valueOf('D1'))];
    check('same-cell conflict converges', d1[0] === d1[1] && /wins/.test(d1[0]), d1.join(' / '));
    // presence: selection box of the other person
    await b.evaluate(() => AtlasSheets.select(4, 2));
    check('peer selection drawn with name', await until(a, () => [...document.querySelectorAll('.live-flag .tag')].some(t => t.textContent === 'Bob')));
    // sheets added/renamed
    await b.evaluate(() => { document.querySelector('#addSheet, .sheet-tabs .add, [data-cmd="add-sheet"]')?.click(); });
    await sleep(300);
    const nb = await b.evaluate(() => AtlasSheets.book.sheets.length);
    check('new sheet replicated', nb === 2 && await until(a, () => AtlasSheets.book.sheets.length === 2), nb);
    const tabsA = await a.$$eval('.tab', ts => ts.map(t => t.textContent));
    check('sheet tab appears for the other person (they stay on their sheet)', tabsA.length === 2 && await a.evaluate(() => AtlasSheets.book.active) === 0, tabsA.join());
    await b.click('.tab >> nth=0');
    // undo takes back only my change
    await a.click('#grid tbody tr:nth-child(1) td:nth-child(6)');
    await a.keyboard.type('mine'); await a.keyboard.press('Enter');
    await sleep(500);
    await b.evaluate(() => AtlasSheets.setCell('E2', 'theirs'));
    await until(a, () => AtlasSheets.valueOf('E2') === 'theirs');
    console.log('dbg-before-undo', await a.evaluate(() => [AtlasSheets.valueOf('E1'), AtlasSheets.valueOf('E2'), AtlasSheets.live.mark()]));
    await a.evaluate(() => Platform.emit('undo'));
    await sleep(500);
    console.log('dbg-after-undo', await a.evaluate(() => [AtlasSheets.valueOf('E1'), AtlasSheets.valueOf('E2'), AtlasSheets.live.mark()]));
    check('undo removes my edit everywhere', await until(b, () => AtlasSheets.valueOf('E1') === '' || AtlasSheets.valueOf('E1') == null));
    check('undo keeps the other person\'s edit', await a.evaluate(() => AtlasSheets.valueOf('E2')) === 'theirs');
    // reconnect: server drops, edits made offline sync after
    const conns = [...srv.rooms.values()][0].clients;
    for (const c of conns) if (c.name === 'Bob') c.socket.destroy();
    await sleep(100);
    await b.evaluate(() => AtlasSheets.setCell('F1', 'offline edit'));
    await sleep(1500);
    console.log('dbg', await b.evaluate(() => [AtlasSheets.valueOf('F1'), document.querySelector('.live-btn').textContent]), await a.evaluate(() => AtlasSheets.valueOf('F1')));
    check('offline edit arrives after reconnect', await until(a, () => AtlasSheets.valueOf('F1') === 'offline edit'));
    check('still two people after reconnect', await until(a, () => document.querySelector('.live-btn').textContent === 'Live · 2'));
    // leave
    await b.click('.live-btn');
    await b.click('.modal .actions button:has-text("Leave")');
    check('leaving updates the count', await until(a, () => document.querySelector('.live-btn').textContent === 'Live · 1'));
    await b.evaluate(() => AtlasSheets.setCell('G1', 'after leaving'));
    await sleep(600);
    check('no sync after leaving', await a.evaluate(() => AtlasSheets.valueOf('G1')) !== 'after leaving');
    // a link for the wrong app
    if (ONLY !== 'sheets') {
    const w = await open('write', 'Wu');
    await w.click('.live-btn');
    await w.click('.modal .actions button:has-text("Join")');
    await w.fill('.modal input', link);
    await w.click('.modal .actions .primary');
    await w.waitForURL(/sheets\.html/);
    check('a Sheets link opened from Write switches to Sheets and joins', await until(w, () => window.AtlasSheets && AtlasSheets.valueOf('A1') === 'hello'));
    }
  }

  if (ONLY === 'all' || ONLY === 'write') {
    const a = await open('write', 'Ada'), b = await open('write', 'Bob');
    const paras = p => p.$$eval('#doc > *', ns => ns.map(n => n.textContent));
    await a.click('#doc');
    await a.keyboard.type('First paragraph');
    await a.keyboard.press('Enter');
    await a.keyboard.type('Second paragraph');
    await a.keyboard.press('Enter');
    await a.keyboard.type('Third paragraph');
    const link = await host(a);
    await join(b, link);
    check('write: joiner gets the paragraphs', await until(b, () => [...document.querySelectorAll('#doc > *')].map(n => n.textContent).join('|') === 'First paragraph|Second paragraph|Third paragraph'), (await paras(b)).join('|'));
    // A's caret at the end of paragraph 1; B types in paragraph 3
    await a.click('#doc > p:nth-child(1)');
    await a.keyboard.press('End');
    await b.click('#doc > p:nth-child(3)');
    await b.keyboard.press('End');
    await b.keyboard.type(' (Bob)');
    check('write: typing reaches the other person', await until(a, () => document.querySelector('#doc > p:nth-child(3)').textContent === 'Third paragraph (Bob)'));
    await a.keyboard.type(' (Ada)');
    check('write: my caret stays put while others type', (await paras(a))[0] === 'First paragraph (Ada)', (await paras(a)).join('|'));
    check('write: and my typing reaches them', await until(b, () => document.querySelector('#doc > p:nth-child(1)').textContent === 'First paragraph (Ada)'));
    // simultaneous typing in different paragraphs
    await b.click('#doc > p:nth-child(2)'); await b.keyboard.press('End');
    await Promise.all([a.keyboard.type(' one two three'), b.keyboard.type(' four five six')]);
    await sleep(1200);
    const pa = await paras(a), pb = await paras(b);
    check('write: concurrent typing in two paragraphs converges', pa.join('|') === pb.join('|') && pa[0].endsWith('one two three') && pa[1].endsWith('four five six'), pa.join('|') + ' // ' + pb.join('|'));
    // new paragraph in the middle
    await b.keyboard.press('Enter');
    await b.keyboard.type('Inserted by Bob');
    check('write: a new paragraph lands in the right place', await until(a, () => [...document.querySelectorAll('#doc > *')].map(n => n.textContent)[2] === 'Inserted by Bob'), (await paras(a)).join('|'));
    // formatting
    await b.keyboard.down('Shift'); await b.keyboard.press('Home'); await b.keyboard.up('Shift');
    await b.keyboard.press('Control+b');
    console.log('dbg', await b.evaluate(() => document.querySelector('#doc > :nth-child(3)').outerHTML));
    check('write: bold reaches the other person', await until(a, () => /<b>|<strong|font-weight: ?(bold|700)/.test(document.querySelector('#doc > :nth-child(3)').innerHTML)));
    // remote caret
    await b.click('#doc > p:nth-child(1)');
    check('write: other person\'s caret shown with their name', await until(a, () => [...document.querySelectorAll('.live-flag .tag')].some(t => t.textContent === 'Bob')));
    // delete a paragraph
    await a.evaluate(() => { document.querySelector('#doc > :nth-child(4)').remove(); document.getElementById('doc').dispatchEvent(new Event('input')); });
    check('write: deleting a paragraph syncs', await until(b, () => document.querySelectorAll('#doc > *').length === 3));
    // untrusted html from a peer is sanitized
    await b.evaluate(() => { const p = document.createElement('p'); p.innerHTML = 'x<img src=x onerror="window.__pwned=1">'; document.getElementById('doc').appendChild(p); document.getElementById('doc').dispatchEvent(new Event('input')); });
    await sleep(800);
    check('write: scripts in a peer\'s paragraph are removed', await a.evaluate(() => !window.__pwned && !document.querySelector('#doc [onerror]')));
    const input = () => document.getElementById('doc').dispatchEvent(new Event('input'));
    // citation sources travel with the document
    await a.evaluate(input => {
      AtlasWrite.notes.loadSources([{ id: 'slive01', type: 'book', title: 'Live Book', authors: 'Smith, Jane', year: '2021', publisher: 'Press' }]);
      const p = document.createElement('p'); p.innerHTML = 'Cited <cite class="citation" data-src="slive01">?</cite>'; document.getElementById('doc').appendChild(p);
      AtlasWrite.notes.refreshCitations(); (0, eval)(input)();
    }, input.toString());
    check('write: citation sources reach the other person', await until(b, () => AtlasWrite.notes.sources.length === 1 && /Smith/.test(document.querySelector('#doc cite.citation')?.textContent || '')), await b.evaluate(() => document.querySelector('#doc cite.citation')?.textContent));
    // endnotes: the generated list is not sent, each side builds its own
    await b.evaluate(() => document.getElementById('showComments').click());
    await a.evaluate(input => {
      const p = document.createElement('p'); p.innerHTML = 'Noted<sup class="note" data-kind="end" title="An endnote" contenteditable="false">0</sup> and <span class="comment" title="Live comment" data-author="Ada">this</span>';
      document.getElementById('doc').appendChild(p); (0, eval)(input)();
    }, input.toString());
    check('write: endnotes listed once, at the end', await until(b, () => document.querySelectorAll('.endnotes-gen').length === 1 && document.getElementById('doc').lastElementChild.matches('.endnotes-gen') && /An endnote/.test(document.querySelector('.endnotes-gen').textContent)));
    check('write: comments pane shows a comment someone else added', await until(b, () => /Live comment/.test(document.querySelector('#commentsPane, .comments-pane')?.textContent || '')));
    // undo takes back only your own typing
    await b.click('#doc > p:nth-child(1)'); await b.keyboard.press('End'); await b.keyboard.type(' BOBX');
    await until(a, () => document.querySelector('#doc > p:nth-child(1)').textContent.endsWith('BOBX'));
    await sleep(1200);
    await a.click('#doc > p:nth-child(2)'); await a.keyboard.press('End'); await a.keyboard.type(' ADAX');
    await until(b, () => document.querySelector('#doc > p:nth-child(2)').textContent.endsWith('ADAX'));
    await sleep(300);
    await a.keyboard.press('Control+z');
    check('write: undo takes back my typing', await until(a, () => !document.querySelector('#doc > p:nth-child(2)').textContent.includes('ADAX')), (await paras(a)).join('|'));
    check('write: undo leaves the other person\'s typing', (await paras(a))[0].endsWith('BOBX'));
    check('write: the undo reaches them', await until(b, () => !document.querySelector('#doc > p:nth-child(2)').textContent.includes('ADAX') && document.querySelector('#doc > p:nth-child(1)').textContent.endsWith('BOBX')));
    await a.keyboard.press('Control+y');
    check('write: redo puts it back', await until(b, () => document.querySelector('#doc > p:nth-child(2)').textContent.endsWith('ADAX')));
    // page setup
    await b.evaluate(() => { document.getElementById('doc').dispatchEvent(new Event('input')); });
    const same = (await paras(a)).join('|') === (await paras(b)).join('|');
    check('write: documents identical at the end', same);
  }

  if (ONLY === 'all' || ONLY === 'slides') {
    const a = await open('slides', 'Ada'), b = await open('slides', 'Bob');
    await a.evaluate(() => { AtlasSlides.addSlide('titleContent'); });
    const link = await host(a);
    await join(b, link);
    check('slides: joiner gets the deck', await until(b, () => AtlasSlides.pres.slides.length === 2));
    const ids = await a.evaluate(() => AtlasSlides.pres.slides.map(s => s.id));
    check('slides: same slide ids', JSON.stringify(ids) === JSON.stringify(await b.evaluate(() => AtlasSlides.pres.slides.map(s => s.id))));
    // B types into the title of slide 1
    await b.evaluate(() => AtlasSlides.goTo(0));
    await b.dblclick('#stage .el >> nth=0');
    await b.keyboard.type('Hello from Bob');
    check('slides: typing in a text box reaches the other person live', await until(a, () => (AtlasSlides.pres.slides[0].elements[0].html || '').includes('Hello from Bob')));
    // A moves the second element on slide 1 while B is still typing in the first
    await a.evaluate(() => AtlasSlides.goTo(0));
    const box = await a.$('#stage .el >> nth=1');
    const bb = await box.boundingBox();
    await a.mouse.move(bb.x + bb.width / 2, bb.y + 10); await a.mouse.down(); await a.mouse.move(bb.x + bb.width / 2 + 60, bb.y + 40, { steps: 5 }); await a.mouse.up();
    await sleep(300);
    const ax = await a.evaluate(() => AtlasSlides.pres.slides[0].elements[1].x);
    await b.keyboard.type(' again');
    check('slides: a move by one and typing by the other both land', await until(b, x => AtlasSlides.pres.slides[0].elements[1].x === x, ax) && await until(a, () => (AtlasSlides.pres.slides[0].elements[0].html || '').includes('Hello from Bob again')), ax);
    const bt = await b.evaluate(() => document.querySelector('#stage .el.editing .content').textContent);
    check('slides: still editing with the caret at the end after the remote move', bt.endsWith('Hello from Bob again') && await b.evaluate(() => document.activeElement.closest('.el.editing') !== null), bt);
    await b.keyboard.press('Escape');
    // B adds a slide; A sees it; presence dot
    await b.evaluate(() => AtlasSlides.addSlide('title'));
    check('slides: new slide replicated', await until(a, () => AtlasSlides.pres.slides.length === 3));
    check('slides: the other person shown on their slide', await until(a, i => document.querySelectorAll('.thumb .live-dot').length === 1 && document.querySelectorAll('.thumb')[i].querySelector('.live-dot') !== null, await b.evaluate(() => [...document.querySelectorAll('.thumb')].findIndex(t => t.classList.contains('active')))));
    // reorder: A moves slide 3 to the front
    await a.evaluate(() => { const p = AtlasSlides.pres; p.slides.unshift(p.slides.pop()); Platform.emit('noop'); });
    await a.evaluate(() => { window.dispatchEvent(new Event('resize')); });
    await a.evaluate(() => AtlasSlides.goTo(0));
    // goTo doesn't mark dirty; nudge a real change
    await a.evaluate(() => { document.getElementById('notes').value = 'speaker note'; document.getElementById('notes').dispatchEvent(new Event('input')); });
    check('slides: reordering replicated', await until(b, first => AtlasSlides.pres.slides[0].id === first, (await a.evaluate(() => AtlasSlides.pres.slides[0].id))));
    check('slides: notes replicated', await until(b, () => AtlasSlides.pres.slides[0].notes === 'speaker note'));
    // theme (deck) change
    await b.evaluate(() => { document.querySelector('[data-theme]:not(.on), .theme-btn')?.click(); });
    await sleep(800);
    const th = [await a.evaluate(() => AtlasSlides.pres.theme), await b.evaluate(() => AtlasSlides.pres.theme)];
    check('slides: theme converges', th[0] === th[1], th.join('/'));
    const canon = () => { const c = v => Array.isArray(v) ? v.map(c) : v && typeof v === 'object' ? Object.fromEntries(Object.keys(v).sort().map(k => [k, c(v[k])])) : v; return JSON.stringify(c(AtlasSlides.pres)); };
    const sa = await a.evaluate(canon), sb = await b.evaluate(canon);
    check('slides: decks identical at the end', sa === sb, sa.length + ' vs ' + sb.length);
  }
  check('no page errors', errors.length === 0, errors.join(' | '));
  await browser.close();
  srv.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
