// Pen and touch: drawing in Slides, Write and Notes (pressure from a pen, palm rejection, the
// eraser, saving), and Touch mode.
const { chromium } = require('playwright');
const JSZip = require('jszip');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const BASE = process.env.E2E_BASE + '/src/renderer/';

(async () => {
  const browser = await chromium.launch();
  const errors = [];
  async function open(app) {
    const ctx = await browser.newContext({ viewport: { width: 1300, height: 860 }, hasTouch: true });
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(app + ': ' + e.message));
    await p.goto(BASE + app + '.html');
    const cdp = await ctx.newCDPSession(p);
    // A stroke from (x1,y1) to (x2,y2) with a mouse, a pen (with pressure) or a finger.
    p.stroke = async (x1, y1, x2, y2, kind = 'mouse') => {
      const steps = 12;
      const pt = i => ({ x: x1 + (x2 - x1) * i / steps, y: y1 + (y2 - y1) * i / steps });
      if (kind === 'mouse') { await p.mouse.move(x1, y1); await p.mouse.down(); for (let i = 1; i <= steps; i++) await p.mouse.move(pt(i).x, pt(i).y); await p.mouse.up(); return; }
      if (kind === 'pen') {
        await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: x1, y: y1, pointerType: 'pen' });
        await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: x1, y: y1, button: 'left', clickCount: 1, pointerType: 'pen', force: 0.2 });
        for (let i = 1; i <= steps; i++) await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', ...pt(i), button: 'left', buttons: 1, pointerType: 'pen', force: Math.min(1, Math.round((0.2 + 0.8 * i / steps) * 100) / 100) });
        await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: x2, y: y2, button: 'left', clickCount: 1, pointerType: 'pen' });
        return;
      }
      await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: x1, y: y1 }] });
      for (let i = 1; i <= steps; i++) await cdp.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [pt(i)] });
      await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
    };
    return p;
  }

  // ---- Slides
  const s = await open('slides');
  await s.waitForFunction(() => window.AtlasSlides && window.Ink);
  await s.click('.ribbon-tabs button[data-tab="draw"]');
  await s.click('[data-panel="draw"] .pen[title="Black pen"]');
  const box = await s.evaluate(() => document.getElementById('stage').getBoundingClientRect().toJSON());
  const X = f => box.x + box.width * f, Y = f => box.y + box.height * f;
  await s.stroke(X(0.2), Y(0.5), X(0.4), Y(0.6));
  await s.stroke(X(0.2), Y(0.7), X(0.4), Y(0.7));
  let inks = await s.evaluate(() => AtlasSlides.pres.slides[0].elements.filter(e => e.type === 'ink').map(e => ({ n: e.strokes.length, x: e.x, w: e.w, f: e.strokes.map(st => st.f || 0) })));
  check('Slides: strokes drawn close together make one drawing on the slide', inks.length === 1 && inks[0].n === 2 && Math.abs(inks[0].x - 0.2 * 960) < 12 && Math.abs(inks[0].w - 0.2 * 960) < 16, JSON.stringify(inks));
  check('…drawn on the slide as it is shown', (await s.$$('#stage .el.ink svg path')).length === 2);
  await s.stroke(X(0.6), Y(0.3), X(0.8), Y(0.35), 'pen');
  inks = await s.evaluate(() => AtlasSlides.pres.slides[0].elements.filter(e => e.type === 'ink').flatMap(e => e.strokes.map(st => st.f || 0)));
  check('…a pen with pressure draws a shaped stroke', inks.length === 3 && inks[2] === 1, JSON.stringify(inks));
  // Palm rejection: once a pen is used, a finger doesn't draw.
  await s.stroke(X(0.1), Y(0.2), X(0.3), Y(0.2), 'touch');
  check('…once a pen is used, a finger on the screen leaves no mark', (await s.evaluate(() => AtlasSlides.pres.slides[0].elements.filter(e => e.type === 'ink').flatMap(e => e.strokes).length)) === 3);
  await s.click('#inkTouch');
  await s.stroke(X(0.1), Y(0.2), X(0.3), Y(0.2), 'touch');
  check('…unless Draw with Touch is on', (await s.evaluate(() => AtlasSlides.pres.slides[0].elements.filter(e => e.type === 'ink').flatMap(e => e.strokes).length)) === 4);
  await s.click('#inkEraser');
  await s.stroke(X(0.3), Y(0.68), X(0.3), Y(0.72));
  check('…the eraser removes the stroke it touches', (await s.evaluate(() => AtlasSlides.pres.slides[0].elements.filter(e => e.type === 'ink').flatMap(e => e.strokes).length)) === 3);
  await s.click('#inkSelect');
  await s.keyboard.press('Control+z');
  check('…and undo brings it back', (await s.evaluate(() => AtlasSlides.pres.slides[0].elements.filter(e => e.type === 'ink').flatMap(e => e.strokes).length)) === 4);
  const before = await s.evaluate(() => JSON.stringify(AtlasSlides.pres.slides[0].elements.filter(e => e.type === 'ink').map(e => [e.x, e.y, e.w, e.h, e.strokes.length])));
  const bytes = await s.evaluate(async () => Array.from(await AtlasSlides.serialize('pptx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  check('…saved in .pptx as pictures PowerPoint shows', Object.keys(zip.files).filter(n => /ppt\/media\/.*\.png$/.test(n)).length >= 1);
  await s.evaluate(async b => AtlasSlides.loadFile({ name: 'ink.pptx', data: new Uint8Array(b) }), bytes);
  const after = await s.evaluate(() => JSON.stringify(AtlasSlides.pres.slides[0].elements.filter(e => e.type === 'ink').map(e => [e.x, e.y, e.w, e.h, e.strokes.length])));
  check('…and opened again as ink', after === before, before + ' / ' + after);
  await s.evaluate(() => { window.__atlasDirty = false; });

  // ---- Write
  const w = await open('write');
  await w.waitForFunction(() => window.AtlasWrite && window.Ink);
  await w.click('#doc');
  await w.keyboard.type('A sketch:');
  await w.click('.ribbon-tabs button[data-tab="draw"]');
  await w.click('#insDrawing');
  await w.waitForSelector('.ink-dialog');
  const pb = await w.evaluate(() => document.querySelector('.ink-dialog .ink-svg').getBoundingClientRect().toJSON());
  await w.stroke(pb.x + 40, pb.y + 40, pb.x + 200, pb.y + 120, 'pen');
  await w.stroke(pb.x + 60, pb.y + 150, pb.x + 260, pb.y + 150);
  await w.click('.ink-dialog .ok');
  const img = await w.evaluate(() => { const i = document.querySelector('#doc img.ink-drawing'); return i && { src: i.src.slice(0, 22), n: JSON.parse(i.dataset.ink).s.length, w: i.width }; });
  check('Write: a drawing goes in at the cursor as a picture with its strokes', img && img.src === 'data:image/png;base64,' && img.n === 2 && img.w > 200 && img.w < 300, JSON.stringify(img));
  await w.dblclick('#doc img.ink-drawing');
  await w.waitForSelector('.ink-dialog');
  check('…double-click opens it to draw on again', (await w.$$('.ink-dialog .ink-svg path')).length === 2);
  const pb2 = await w.evaluate(() => document.querySelector('.ink-dialog .ink-svg').getBoundingClientRect().toJSON());
  await w.stroke(pb2.x + 300, pb2.y + 40, pb2.x + 360, pb2.y + 60);
  await w.click('.ink-dialog .ok');
  check('…and Update replaces it', await w.evaluate(() => document.querySelectorAll('#doc img.ink-drawing').length === 1 && JSON.parse(document.querySelector('#doc img.ink-drawing').dataset.ink).s.length === 3));
  const html = await w.evaluate(() => AtlasWrite.serialize('html'));
  await w.evaluate(async h => AtlasWrite.loadFile({ name: 'ink.html', data: new TextEncoder().encode(h) }), html);
  check('…and the strokes survive saving and opening', await w.evaluate(() => { const i = document.querySelector('#doc img.ink-drawing'); return !!i && JSON.parse(i.dataset.ink).s.length === 3; }));
  await w.evaluate(() => { window.__atlasDirty = false; });

  // ---- Notes: a pen's pressure
  const n = await open('notes');
  await n.waitForFunction(() => window.AtlasNotes && window.Ink);
  await n.click('[data-tab="draw"]');
  await n.click('.ribbon .pen[title="Blue pen"]');
  const nb = await n.evaluate(() => document.getElementById('page').getBoundingClientRect().toJSON());
  await n.stroke(nb.x + 100, nb.y + 200, nb.x + 300, nb.y + 260, 'pen');
  check('Notes: a pen draws a shaped stroke', await n.evaluate(() => AtlasNotes.page.ink.length === 1 && AtlasNotes.page.ink[0].f === 1 && !!document.querySelector('#ink path[fill="#1565c0"]')));

  // ---- Touch mode (on by default on a touch screen, as this one is)
  check('Touch mode is on by default where touch is the main input', await n.evaluate(() => Platform.touchMode.get()));
  await n.evaluate(() => Platform.touchMode.set(false));
  await n.click('.ribbon-mode');
  await n.click('.ribbon-menu .touch-toggle');
  const big = await n.evaluate(() => ({ on: document.documentElement.classList.contains('touch-mode'), h: document.querySelector('.ribbon-tabs button[data-tab]').getBoundingClientRect().height }));
  check('Touch mode gives commands more room', big.on && big.h >= 40, JSON.stringify(big));
  await n.reload();
  await n.waitForFunction(() => window.AtlasNotes);
  check('…and is remembered', await n.evaluate(() => document.documentElement.classList.contains('touch-mode')));
  await n.evaluate(() => { Platform.touchMode.set(false); window.__atlasDirty = false; });

  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
