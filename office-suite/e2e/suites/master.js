const { chromium } = require('playwright');
const JSZip = require('jszip');
const fs = require('fs');
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/slides.html'); await p.waitForTimeout(300);
  // Master view: add a logo shape to the master, edit a layout, add a custom layout
  await p.click('[data-tab="viewtab"]'); await p.click('#masterView');
  const mv = await p.evaluate(() => ({ tab: !document.querySelector('[data-tab="master"]').hidden, thumbs: [...document.querySelectorAll('#thumbs .thumb-name')].map(x => x.textContent).slice(0, 3).join('|'), st: document.getElementById('stSlide').textContent }));
  check('Slide Master view lists the master and layouts', mv.tab && mv.thumbs === 'Slide Master|Title Slide|Title and Content' && mv.st === 'Slide Master', JSON.stringify(mv));
  await p.click('[data-tab="insert"]'); await p.click('[data-shape="star5"]');
  const masterEls = await p.evaluate(() => AtlasSlides.pres.master.elements.map(e => e.ph || e.shape).join());
  check('shapes inserted in Master view go on the master', masterEls === 'dt,footer,sldNum,star5', masterEls);
  // Edit the Title and Content layout: move the title
  await p.click('#thumbs .layout-thumb:nth-child(3)');
  await p.evaluate(() => { const L = AtlasSlides.pres.layouts.content; L.elements[0].y = 300; });
  await p.click('[data-tab="master"]');
  await p.evaluate(() => { window.__a = [{ name: 'Big Quote', from: 'title' }]; UI.form = async () => window.__a.shift(); });
  await p.click('#insLayout');
  await p.click('#closeMaster');
  // New slides use the edited layout, and the master star shows on every slide
  await p.click('[data-tab="home"]');
  await p.click('#newSlide');
  const s2 = await p.evaluate(() => ({ titleY: AtlasSlides.pres.slides[1].elements.find(e => e.ph === 'title').y, star: document.querySelectorAll('#stage .master-el.shape').length, opts: [...document.getElementById('layout').options].map(o => o.textContent).slice(-1)[0] }));
  check('new slides use edited layouts; master shapes show; custom layout offered', s2.titleY === 300 && s2.star === 1 && s2.opts === 'Big Quote', JSON.stringify(s2));
  // Header & footer
  await p.click('[data-tab="insert"]');
  await p.evaluate(() => { window.__a = [{ date: '', number: '1', text: 'Atlas Inc.', skip: '1' }]; UI.form = async () => window.__a.shift(); });
  await p.click('#headerFooter');
  const foot = await p.evaluate(() => ({ s2: [...document.querySelectorAll('#stage .master-el')].map(x => x.textContent).filter(Boolean).join('|'), thumb1: [...document.querySelectorAll('#thumbs .thumb')[0].querySelectorAll('.master-el')].map(x => x.textContent).filter(Boolean).join('|') }));
  check('footer and slide number on slides, not on the title slide', foot.s2 === 'Atlas Inc.|2' && foot.thumb1 === '', JSON.stringify(foot));
  // Hide background graphics on slide 2
  await p.click('[data-tab="design"]'); await p.click('#hideMaster');
  check('Hide background graphics', (await p.evaluate(() => document.querySelectorAll('#stage .master-el').length)) === 0);
  await p.click('#hideMaster');
  // pptx round trip keeps master/layouts exactly and doesn't duplicate master shapes
  const bytes = await p.evaluate(async () => Array.from(await AtlasSlides.serialize('pptx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const sx = await zip.file('ppt/slides/slide2.xml').async('string');
  check('pptx: master shapes and footers drawn on the slide', /name="atlas-master-/.test(sx) && /Atlas Inc\./.test(sx) && !!zip.file('docProps/custom.xml'));
  const before = await p.evaluate(() => JSON.stringify({ m: AtlasSlides.pres.master.elements.map(e => [e.ph || e.shape, e.x, e.y]), f: AtlasSlides.pres.master.footer, l: Object.keys(AtlasSlides.pres.layouts), n: AtlasSlides.pres.slides.map(s => s.elements.filter(e => e.html || e.type !== 'text').length) }));
  const after = await p.evaluate(async b => { await AtlasSlides.loadFile({ name: 'm.pptx', path: null, data: new Uint8Array(b) }); return JSON.stringify({ m: AtlasSlides.pres.master.elements.map(e => [e.ph || e.shape, e.x, e.y]), f: AtlasSlides.pres.master.footer, l: Object.keys(AtlasSlides.pres.layouts), n: AtlasSlides.pres.slides.map(s => s.elements.filter(e => e.html || e.type !== 'text').length) }); }, bytes);
  check('reopened .pptx: same master, footers, layouts and slide contents', before === after, before + '\n   vs ' + after);
  // A PowerPoint-style deck (python-pptx): inheritance
  await p.evaluate(async b => { await AtlasSlides.loadFile({ name: 'i.pptx', path: null, data: new Uint8Array(b) }); }, Array.from(fs.readFileSync(SP + '/files/deck-inherit2.pptx')));
  const inh = await p.evaluate(() => {
    const P = AtlasSlides.pres;
    const s1 = P.slides[0].elements.map(e => `${e.ph || e.type}@${e.x},${e.y},${e.w},${e.h}:${e.size}`);
    const g = P.slides[1].elements.find(e => e.type === 'group'); const s2 = g ? g.children.map(e => `${e.shape}@${g.x + e.x},${g.y + e.y},${e.w},${e.h}`) : ['no group'];
    return { s1, s2, master: P.master.elements.map(e => e.ph || e.shape + '@' + e.y + ',' + e.h), layouts: Object.values(P.layouts).map(l => l.name).join('|'), slideLayouts: P.slides.map(s => P.layouts[s.layout] && P.layouts[s.layout].name).join('|'), theme: P.theme, font: P.customTheme && P.customTheme.titleFont };
  });
  check('placeholders without a position take their layout\'s position and size', inh.s1[0] === 'title@72,168,816,116:44' && inh.s1[1] === 'subtitle@144,306,672,138:32', JSON.stringify(inh.s1));
  check('group transforms scale children', inh.s2.join(' ') === 'ellipse@0,0,192,144 rect@384,0,192,72', inh.s2.join(' '));
  check('master shapes and footer placeholders, layouts and theme fonts imported', inh.master.join() === 'dt,footer,sldNum,rect@0,31' && inh.slideLayouts === 'Title Slide|Title and Content' && inh.theme === 'custom' && inh.font === 'Calibri', JSON.stringify(inh));
  const shown = await p.evaluate(() => document.querySelectorAll('#stage .master-el.shape').length);
  check('the master bar shows on imported slides', shown === 1);
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
