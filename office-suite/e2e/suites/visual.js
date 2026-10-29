// Visual regression: each app's main screen in light and dark, at three window sizes, compared
// with the reference pictures in e2e/visual/. A change that moves, recolours or hides something
// shows up here even when every behaviour test still passes.
//
//   npm run e2e -- visual                 compare
//   UPDATE_VISUAL=1 npm run e2e -- visual  accept the current look as the new reference
//
// The comparison runs in Chromium (no image libraries needed) and forgives anti-aliasing: a pixel
// counts as changed only if no pixel next to it in the other picture is close in colour. When
// something does differ, a picture of the differences (changed pixels in red) is written to
// e2e/visual/diff/.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const BASE = process.env.E2E_BASE + '/src/renderer/';
const DIR = path.join(process.cwd(), 'e2e/visual');
const DIFF = path.join(DIR, 'diff');
const UPDATE = process.env.UPDATE_VISUAL === '1';
const SIZES = [[1024, 768], [1366, 820], [1680, 1050]];
const THEMES = ['light', 'dark'];
const MAX_CHANGED = 0.002; // share of pixels that may differ (text rendering varies a little)

// Each screen with some content in it, the same every time.
const SCREENS = {
  home: { page: 'index.html', ready: () => !!document.querySelector('#templates .tpl') },
  write: {
    page: 'write.html', ready: () => !!window.AtlasWrite,
    fill: () => { document.getElementById('doc').innerHTML = '<h1>Quarterly report</h1><p>Revenue grew <b>14%</b> on the year, led by the north region and two new partners.</p><h2>Highlights</h2><ul><li>Three new markets</li><li>Costs held flat</li></ul><p>Next quarter we focus on retention.</p>'; }
  },
  sheets: {
    page: 'sheets.html', ready: () => !!window.AtlasSheets,
    fill: () => {
      [['Region', 'Q1', 'Q2', 'Q3'], ['North', 120, 135, 150], ['South', 95, 102, 118], ['West', 80, 91, 99]].forEach((r, i) => r.forEach((v, c) => AtlasSheets.setCell('ABCD'[c] + (i + 1), v)));
      AtlasSheets.setCell('E1', 'Total'); [2, 3, 4].forEach(r => AtlasSheets.setCell('E' + r, `=SUM(B${r}:D${r})`));
      AtlasSheets.book.sheets[0].charts.push({ type: 'bar', range: 'A1:D4', x: 470, y: 10, w: 380, h: 240, title: 'Sales', style: 'modern' });
      AtlasSheets.setCell('A6', '');
    }
  },
  slides: {
    page: 'slides.html', ready: () => !!window.AtlasSlides,
    fill: () => { const s = AtlasSlides.pres.slides[0]; s.elements[0].html = 'Launch plan'; s.elements[1].html = 'Product team'; AtlasSlides.goTo(0); }
  },
  notes: { page: 'notes.html', ready: () => !!document.querySelector('.nb-list li') }
};

/** Compares two PNGs (as base64) in the page; returns the changed share and, if any, a diff PNG. */
function compare(page, a, b) {
  return page.evaluate(async ([a, b]) => {
    const load = async b64 => { const bmp = await createImageBitmap(await (await fetch('data:image/png;base64,' + b64)).blob()); const c = new OffscreenCanvas(bmp.width, bmp.height); const g = c.getContext('2d'); g.drawImage(bmp, 0, 0); return g.getImageData(0, 0, bmp.width, bmp.height); };
    const A = await load(a), B = await load(b);
    if (A.width !== B.width || A.height !== B.height) return { size: `${A.width}×${A.height} vs ${B.width}×${B.height}`, share: 1 };
    const W = A.width, H = A.height, da = A.data, db = B.data, T = 40;
    const near = (d1, i, d2, x, y) => {
      for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
        const xx = x + dx, yy = y + dy;
        if (xx < 0 || yy < 0 || xx >= W || yy >= H) continue;
        const j = (yy * W + xx) * 4;
        if (Math.abs(d1[i] - d2[j]) <= T && Math.abs(d1[i + 1] - d2[j + 1]) <= T && Math.abs(d1[i + 2] - d2[j + 2]) <= T) return true;
      }
      return false;
    };
    const out = new ImageData(W, H);
    let changed = 0;
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const i = (y * W + x) * 4;
      const same = da[i] === db[i] && da[i + 1] === db[i + 1] && da[i + 2] === db[i + 2];
      const bad = !same && !(near(da, i, db, x, y) && near(db, i, da, x, y));
      if (bad) { changed++; out.data.set([255, 0, 0, 255], i); }
      else { const v = 255 - (255 - (da[i] + da[i + 1] + da[i + 2]) / 3) * 0.25; out.data.set([v, v, v, 255], i); }
    }
    let diff = null;
    if (changed) { const c = new OffscreenCanvas(W, H); c.getContext('2d').putImageData(out, 0, 0); const blob = await c.convertToBlob({ type: 'image/png' }); diff = await new Promise(res => { const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(',')[1]); fr.readAsDataURL(blob); }); }
    return { share: changed / (W * H), changed, diff };
  }, [a.toString('base64'), b.toString('base64')]);
}

(async () => {
  fs.mkdirSync(DIR, { recursive: true });
  fs.rmSync(DIFF, { recursive: true, force: true });
  // Text drawn the same way on every computer: no hinting or subpixel colour.
  const browser = await chromium.launch({ args: ['--font-render-hinting=none', '--disable-lcd-text', '--force-color-profile=srgb', '--disable-skia-runtime-opts'] });
  const errors = [];
  const cmp = await (await browser.newContext()).newPage();
  let made = 0;
  for (const theme of THEMES) for (const [w, h] of SIZES) {
    const ctx = await browser.newContext({ viewport: { width: w, height: h }, deviceScaleFactor: 1, reducedMotion: 'reduce', colorScheme: theme });
    await ctx.addInitScript(t => {
      try { localStorage.setItem('atlas-theme', t); localStorage.removeItem('atlas-user-name'); } catch { /* ignore */ }
      // No blinking caret in the pictures.
      document.addEventListener('DOMContentLoaded', () => { const s = document.createElement('style'); s.textContent = '*{caret-color:transparent!important}'; document.head.appendChild(s); });
    }, theme);
    for (const [name, sc] of Object.entries(SCREENS)) {
      const page = await ctx.newPage();
      await page.clock.setFixedTime(new Date('2026-03-10T10:00:00'));
      page.on('pageerror', e => errors.push(`${name}: ${e.message}`));
      await page.goto(BASE + sc.page);
      await page.waitForFunction(sc.ready);
      if (sc.fill) await page.evaluate(sc.fill);
      await page.evaluate(async () => { await document.fonts.ready; document.activeElement && document.activeElement.blur && document.activeElement.blur(); });
      await page.mouse.move(0, h - 1);
      await page.waitForTimeout(350);
      const shot = await page.screenshot({ animations: 'disabled' });
      const file = path.join(DIR, `${name}-${theme}-${w}x${h}.png`);
      const label = `${name}, ${theme}, ${w}×${h}`;
      if (UPDATE || !fs.existsSync(file)) {
        if (!UPDATE) { check(`${label}: has a reference picture`, false, 'run with UPDATE_VISUAL=1 to make it'); }
        else { fs.writeFileSync(file, shot); made++; }
      } else {
        const r = await compare(cmp, fs.readFileSync(file), shot);
        if (r.diff) { fs.mkdirSync(DIFF, { recursive: true }); fs.writeFileSync(path.join(DIFF, path.basename(file)), Buffer.from(r.diff, 'base64')); fs.writeFileSync(path.join(DIFF, path.basename(file, '.png') + '-now.png'), shot); }
        check(`${label} looks as it did`, !r.size && r.share <= MAX_CHANGED, r.size || `${(r.share * 100).toFixed(3)}% of pixels changed`);
      }
      await page.evaluate(() => { window.__atlasDirty = false; });
      await page.close();
    }
    await ctx.close();
  }
  if (UPDATE) check(`reference pictures written (${made})`, made === THEMES.length * SIZES.length * Object.keys(SCREENS).length);
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
