// Slides: Morph, Designer, SmartArt, handouts and recording narration — in the editor, the show
// and the .pptx PowerPoint reads.
const { chromium } = require('playwright');
const JSZip = require('jszip');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAEklEQVR4nGP4z8CAB+GTG8HSALfKY52fTcuYAAAAAElFTkSuQmCC';

(async () => {
  // A fake microphone (a tone) for recording.
  const browser = await chromium.launch({ args: ['--use-fake-ui-for-media-stream', '--use-fake-device-for-media-stream'] });
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 }, permissions: ['microphone'] });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/slides.html');
  await p.waitForFunction(() => window.AtlasSlides && window.SlidesExtras);
  const click = id => p.evaluate(i => document.getElementById(i).click(), id);
  const submit = async fill => { await p.waitForSelector('.modal'); if (fill) await fill(); await p.click('.modal .actions .primary'); await p.waitForSelector('.modal', { state: 'detached' }); };
  const pres = () => p.evaluate(() => JSON.parse(JSON.stringify(AtlasSlides.pres)));

  // ---- Morph: a title and a shape, duplicated and moved
  await p.evaluate(() => {
    const s = AtlasSlides.pres.slides[0];
    s.elements[0].html = 'Morph demo';
    s.elements.push({ id: 'circle1', type: 'shape', shape: 'ellipse', x: 80, y: 360, w: 120, h: 120, fill: '#4472c4', line: null, html: '', size: 18, align: 'center', valign: 'middle' });
    AtlasSlides.goTo(0);
  });
  await click('dupSlide');
  await p.evaluate(() => {
    const s = AtlasSlides.pres.slides[1];
    const c = s.elements.find(e => e.type === 'shape');
    Object.assign(c, { x: 700, y: 120, w: 200, h: 200, fill: '#ed7d31' });
    s.transition = 'morph'; s.transDur = 1.2;
    AtlasSlides.goTo(1);
  });
  let pr = await pres();
  check('duplicated objects share morph keys', pr.slides[0].elements.every((e, i) => e.mk && e.mk === pr.slides[1].elements[i].mk));
  check('objects pair up between the slides', await p.evaluate(() => SlidesExtras.pairs(AtlasSlides.pres.slides[0], AtlasSlides.pres.slides[1]).size === AtlasSlides.pres.slides[1].elements.length));
  // In the show: midway through the transition the circle is between its two places.
  await p.evaluate(() => AtlasSlides.present(0));
  await sleep(400);
  await p.keyboard.press('ArrowRight');
  await sleep(550);
  const mid = await p.evaluate(() => { const n = [...document.querySelectorAll('#presentSlide .el.shape')].pop(); const cs = getComputedStyle(n); return [parseFloat(cs.left), parseFloat(cs.width), n.getAnimations().length]; });
  check('Morph moves and grows the shape during the transition', mid[0] > 90 && mid[0] < 690 && mid[1] > 120 && mid[1] < 200 && mid[2] > 0, JSON.stringify(mid));
  await sleep(1000);
  const end = await p.evaluate(() => { const n = [...document.querySelectorAll('#presentSlide .el.shape')].pop(); return [parseFloat(getComputedStyle(n).left), n.getAnimations().length]; });
  check('…and ends where the second slide has it', end[0] === 700 && end[1] === 0, JSON.stringify(end));
  await p.keyboard.press('Escape');
  await sleep(200);

  // ---- SmartArt
  await click('insSmartArt');
  await submit(async () => { await p.selectOption('.modal select', 'hierarchy'); await p.fill('.modal textarea', 'CEO\n  CTO\n    Dev\n    QA\n  CFO'); });
  pr = await pres();
  let g = pr.slides[1].elements.find(e => e.smart);
  check('SmartArt hierarchy: a group of boxes and lines from the outline', g && g.smart.type === 'hierarchy' && g.children.filter(c => c.html).length === 5 && g.children.some(c => c.decorative), g && g.children.length);
  check('…drawn on the slide', await p.evaluate(() => [...document.querySelectorAll('#stage .el.group .el.shape .content')].map(c => c.textContent).filter(Boolean).join(',') === 'CEO,CTO,Dev,QA,CFO'));
  await p.dblclick('#stage .el.group');
  await submit(async () => { await p.selectOption('.modal select', 'process'); await p.fill('.modal textarea', 'Plan\nBuild\nShip'); });
  pr = await pres();
  g = pr.slides[1].elements.find(e => e.smart);
  check('double-click edits it: now a three-step process with arrows', g.smart.type === 'process' && g.children.filter(c => c.shape === 'roundRect').length === 3 && g.children.filter(c => c.shape === 'rightArrow').length === 2);

  // ---- Designer: a slide with a title, text and a picture
  await p.evaluate(src => {
    AtlasSlides.addSlide('content');
    const s = AtlasSlides.pres.slides[AtlasSlides.pres.slides.length - 1];
    s.elements[0].html = 'Our results';
    s.elements[1].html = '<ul><li>Up 20%</li><li>New customers</li></ul>';
    s.elements.push({ id: 'pic1', type: 'image', src, x: 600, y: 300, w: 200, h: 150 });
    AtlasSlides.goTo(AtlasSlides.pres.slides.length - 1);
  }, PNG);
  await click('designIdeas');
  await p.waitForSelector('.designer-pane .idea');
  const names = await p.$$eval('.designer-pane .idea span', s => s.map(x => x.textContent));
  check('Designer offers layouts for title, text and picture', names.includes('Picture right') && names.includes('Full-slide picture') && names.includes('Title band'), names.join(', '));
  await p.click('.designer-pane .idea:has-text("Picture right")');
  pr = await pres();
  const last = pr.slides[pr.slides.length - 1];
  const pic = last.elements.find(e => e.type === 'image');
  const ttl = last.elements.find(e => e.ph === 'title');
  check('applying an idea rearranges the slide (picture right, title left)', pic.x >= 500 && ttl.x < 100 && ttl.html.includes('Our results') && Math.abs(pic.w / pic.h - 200 / 150) < 0.02, JSON.stringify({ pic, x: ttl.x }));
  await p.evaluate(() => Platform.emit('undo'));
  pr = await pres();
  check('undo puts the slide back', pr.slides[pr.slides.length - 1].elements.find(e => e.type === 'image').x === 600);
  await p.click('.designer-pane .x');

  // ---- Handouts: 3 per page has lines; notes pages have the notes
  await p.evaluate(() => { AtlasSlides.pres.slides[0].notes = 'Say hello first.'; });
  const hand = await p.evaluate(() => {
    const box = document.createElement('div');
    document.body.appendChild(box);
    const r = n => { SlidesExtras.handouts(box, AtlasSlides.pres.slides, n, (s, el) => { el.dataset.rendered = s.id; }, { title: 'Deck' }); return { pages: box.querySelectorAll('.handout-page').length, slides: box.querySelectorAll('.handout-slide').length, lines: box.querySelectorAll('.handout-lines').length, notes: [...box.querySelectorAll('.handout-notes')].map(n => n.textContent) }; };
    const out = { three: r(3), six: r(6), notes: r('notes') };
    box.remove();
    return out;
  });
  check('handouts: 3 per page with lines, 6 per page, notes pages', hand.three.pages === 1 && hand.three.lines === 3 && hand.six.pages === 1 && hand.six.slides === 3 && hand.notes.pages === 3 && hand.notes.notes[0] === 'Say hello first.', JSON.stringify(hand));

  // ---- Recording narration and timings
  await p.evaluate(() => AtlasSlides.goTo(0));
  await click('recordShow');
  await p.waitForFunction(() => !document.getElementById('presenter').hidden);
  await sleep(1300);
  await p.keyboard.press('ArrowRight');
  await sleep(900);
  await p.keyboard.press('Escape');
  await p.waitForFunction(() => AtlasSlides.pres.slides[0].advance && AtlasSlides.pres.slides[0].elements.some(e => e.narration), null, { timeout: 15000 }).catch(() => {});
  pr = await pres();
  const narr = pr.slides[0].elements.find(e => e.narration);
  check('recording: each slide gets its time', pr.slides[0].advance >= 1 && pr.slides[0].advance < 4 && pr.slides[1].advance >= 0.5, `${pr.slides[0].advance} ${pr.slides[1].advance}`);
  check('recording: narration as a WAV clip that plays by itself, icon hidden', narr && /^data:audio\/wav;base64,UklGR/.test(narr.src) && narr.autoplay && narr.hideIcon, narr ? narr.src.slice(0, 40) : 'none');

  // ---- .pptx: Morph for PowerPoint, and everything comes back
  const bytes = Buffer.from(await p.evaluate(async () => Array.from(await AtlasSlides.serialize('pptx'))));
  const zip = await JSZip.loadAsync(bytes);
  const s1 = await zip.file('ppt/slides/slide1.xml').async('string'), s2 = await zip.file('ppt/slides/slide2.xml').async('string');
  check('pptx: Morph transition with a fallback', /<mc:Choice [^>]*Requires="p159"><p:transition [^>]*p14:dur="1200"[^>]*><p159:morph option="byObject"\/>/.test(s2) && /<mc:Fallback><p:transition[^>]*><p:fade\/>/.test(s2));
  const names1 = [...s1.matchAll(/name="!!([^"]+)"/g)].map(m => m[1]), names2 = [...s2.matchAll(/name="!!([^"]+)"/g)].map(m => m[1]);
  check('pptx: matched objects share "!!" names on both slides', names1.length >= 2 && names1.every(n => names2.includes(n)), names1.join(',') + ' / ' + names2.join(','));
  check('pptx: narration saved as audio', Object.keys(zip.files).some(f => /^ppt\/media\/.*\.wav$/.test(f)));
  await p.evaluate(async b => { await AtlasSlides.loadFile({ name: 'x.pptx', data: new Uint8Array(b) }); }, Array.from(bytes));
  pr = await pres();
  check('reopened: Morph, its keys, SmartArt and narration', pr.slides[1].transition === 'morph' && pr.slides[1].elements.filter(e => e.mk).length >= 2 && pr.slides.some(s => s.elements.some(e => e.smart && e.smart.type === 'process')) && pr.slides[0].elements.some(e => e.type === 'media' && e.autoplay), JSON.stringify(pr.slides[1].elements.map(e => [e.type, e.mk || ''])));
  check('reopened: objects still pair up for Morph', await p.evaluate(() => SlidesExtras.pairs(AtlasSlides.pres.slides[0], AtlasSlides.pres.slides[1]).size >= 2));

  check('no page errors', !errors.length, errors.join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
