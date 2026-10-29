// Live editing with a hostile participant: someone who joins a session with their own client and
// sends malformed or malicious changes. Everyone else's copy stays clean and keeps working — the
// changes go through the same checks as a file.
const { chromium } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const { createCollabServer } = require(path.join(ROOT, 'src/main/collab-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.E2E_BASE + '/src/renderer/';

(async () => {
  const srv = createCollabServer();
  await new Promise(res => srv.listen(0, '127.0.0.1', res));
  const WS = `ws://127.0.0.1:${srv.address().port}/`;
  const browser = await chromium.launch();
  const errors = [];
  async function hostApp(app) {
    const ctx = await browser.newContext({ viewport: { width: 1300, height: 860 } });
    await ctx.addInitScript(() => { localStorage.setItem('atlas-user-name', 'Ada'); localStorage.setItem('atlas-author', 'Ada'); });
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(app + ': ' + e.message));
    await p.goto(BASE + app + '.html');
    await p.waitForFunction(() => window.Collab && document.querySelector('.live-btn'));
    await p.click('.live-btn');
    await p.click('.modal .actions button:has-text("Start on a server")');
    await p.fill('.modal input', WS);
    await p.click('.modal .actions .primary');
    await p.waitForSelector('.modal p:has-text("room=")', { timeout: 10000 });
    const room = (await p.textContent('.modal p')).match(/room=([\w-]+)/)[1];
    await p.click('.modal .actions button:has-text("Close")');
    return { p, room };
  }
  // The hostile participant: a bare WebSocket client.
  async function intruder(room, app) {
    const ws = new WebSocket(WS);
    await new Promise(r => { ws.onopen = r; });
    const got = [];
    ws.onmessage = e => got.push(JSON.parse(e.data));
    ws.send(JSON.stringify({ t: 'hello', room, app, name: 'Mallory', color: '#ff0000' }));
    for (let i = 0; i < 40 && !got.some(m => m.t === 'welcome'); i++) await sleep(50);
    return { ws, got, send: ops => ws.send(JSON.stringify({ t: 'op', n: 1, seen: (got.find(m => m.t === 'welcome') || {}).seq || 0, ops })) };
  }

  // ---- Slides
  const s = await hostApp('slides');
  const sid = await s.p.evaluate(() => AtlasSlides.pres.slides[0].id);
  const m1 = await intruder(s.room, 'slides');
  m1.send([
    { k: `e:${sid}:evil1`, v: { id: 'evil1', type: 'text', x: 10, y: 10, w: 300, h: 80, html: '<img src=x onerror="window.__pwned=1"><b>hello</b><script>window.__pwned=2</script><a href="javascript:window.__pwned=3">link</a>', z: 9 } },
    { k: `e:${sid}:evil2`, v: { id: 'evil2', type: 'connector', kind: 'straight', x: 0, y: 0, w: 10, h: 10, color: 'red"><img src=x onerror=1>', x1: 0, y1: 0, x2: 50, y2: 50 } },
    { k: `e:${sid}:evil3`, v: { id: 'evil3', type: 'ink', x: 0, y: 0, w: 10, h: 10, strokes: 'not a list' } },
    { k: `e:${sid}:evil4`, v: { id: 'x"><svg onload=alert(1)>', type: 'shape', shape: 'rect', x: 0, y: 0, w: 10, h: 10 } },
    { k: `e:${sid}:evil5`, v: { id: 'evil5', type: 'whatever' } },
    { k: `e:${sid}:evil6`, v: 'a string' },
    { k: `s:${sid}`, v: { pos: 0, layout: 'nonsense', bg: 'url(javascript:1)', transition: 'explode', notes: { not: 'text' } } },
    { k: 'deck', v: { theme: '<script>', master: 'oops', layouts: { 'bad key!': {} } } }
  ]);
  await sleep(800);
  const sl = await s.p.evaluate(() => {
    const els = AtlasSlides.pres.slides[0].elements;
    const evil1 = els.find(e => e.id === 'evil1');
    return {
      ids: els.map(e => e.id), html: evil1 && evil1.html, theme: AtlasSlides.pres.theme, layout: AtlasSlides.pres.slides[0].layout, bg: AtlasSlides.pres.slides[0].bg, notes: AtlasSlides.pres.slides[0].notes,
      onerror: !!document.querySelector('#stage [onerror], #stage script, #stage [onload]'), js: !!document.querySelector('#stage a[href^="javascript"]'), pwned: window.__pwned || 0, slides: AtlasSlides.pres.slides.length
    };
  });
  check('slides: a hostile participant\'s text is cleaned (no scripts, handlers or javascript: links)', sl.html && /<b>hello<\/b>/.test(sl.html) && !/onerror|<script|javascript:/i.test(sl.html) && !sl.onerror && !sl.js && !sl.pwned, JSON.stringify(sl.html));
  check('slides: malformed objects are dropped, the rest is kept', sl.ids.includes('evil1') && sl.ids.includes('evil2') && !sl.ids.includes('evil3') && !sl.ids.includes('evil5') && !sl.ids.includes('evil6') && !sl.ids.some(i => /[<>"]/.test(i)), sl.ids.join());
  check('slides: bad slide and design settings fall back to safe values', sl.theme === 'office' && sl.layout === 'blank' && sl.bg === null && sl.notes === '' && sl.slides === 1, JSON.stringify(sl));
  // Still works: a normal change from the host goes through.
  await s.p.evaluate(() => AtlasSlides.addSlide('content'));
  await sleep(400);
  check('slides: the session keeps working', m1.got.some(m => m.t === 'op' && m.ops.some(o => o.k.startsWith('s:') && o.k !== `s:${sid}`)));
  await s.p.evaluate(() => { window.__atlasDirty = false; });

  // ---- Sheets
  const h = await hostApp('sheets');
  const uid = await h.p.evaluate(() => AtlasSheets.book.sheets[0].uid || null);
  const m2 = await intruder(h.room, 'sheets');
  const sheetKey = uid ? `sh:${uid}` : (m2.got.find(m => m.t === 'welcome').state && Object.keys(m2.got.find(m => m.t === 'welcome').state).find(k => k.startsWith('sh:')));
  const u = sheetKey.slice(3);
  m2.send([
    { k: `c:${u}:A1`, v: { v: 'ok', s: { bg: 'url(javascript:1)', color: '#ff0000', evil: '<script>' }, f: { __proto__: { polluted: 1 } } } },
    { k: `c:${u}:A2`, v: { v: { not: 'a value' } } },
    { k: `c:${u}:B1`, v: 'a bare string' },
    { k: `c:${u}:ZZZZ99999999`, v: { v: 'out of range' } },
    { k: sheetKey, v: { pos: 0, name: '../evil/[x]*', charts: [{ type: 'pie3d<script>', range: 'A1:B2' }, { type: 'bar', range: 'A1:B3', x: 10, y: 10, w: 400, h: 300 }], merges: [{ r1: 5, c1: 5, r2: 1, c2: 1 }, { r1: 0, c1: 0, r2: 0, c2: 1 }], colWidths: { 0: 1e9, 1: 80 }, cf: 'nope', freezeRow: 'yes' } }
  ]);
  await sleep(800);
  const sh = await h.p.evaluate(() => {
    const s = AtlasSheets.book.sheets[0];
    return { name: s.name, a1: s.cells.A1, a2: s.cells.A2, b1: s.cells.B1, charts: s.charts.map(c => c.type), merges: s.merges, widths: s.colWidths, cf: Array.isArray(s.cf), freeze: s.freezeRow, polluted: ({}).polluted };
  });
  check('sheets: hostile cell data is reduced to plain values and known styles', sh.a1 && sh.a1.v === 'ok' && sh.a1.s && sh.a1.s.color === '#ff0000' && !sh.a1.s.bg && !sh.a1.s.evil && !('f' in sh.a1) && sh.a2 && sh.a2.v === '' && !sh.b1 && !sh.polluted, JSON.stringify([sh.a1, sh.a2, sh.b1]));
  check('sheets: bad sheet settings are checked like a file\'s', sh.name === '.._evil__x__' && sh.charts.join() === 'bar' && sh.merges.length === 1 && !(0 in sh.widths) && sh.widths[1] === 80 && sh.cf && sh.freeze !== 'yes', JSON.stringify(sh));
  await h.p.evaluate(() => AtlasSheets.setCell('C3', '=1+1'));
  await sleep(300);
  check('sheets: the session keeps working', (await h.p.evaluate(() => AtlasSheets.valueOf('C3'))) === 2 && m2.got.some(m => m.t === 'op' && m.ops.some(o => /C3$/.test(o.k))));
  await h.p.evaluate(() => { window.__atlasDirty = false; });

  m1.ws.close(); m2.ws.close();
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
  await new Promise(res => { srv.close(res); for (const rm of srv.rooms.values()) for (const c of rm.clients) c.socket.destroy(); });
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
