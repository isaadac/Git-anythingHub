const { chromium } = require('playwright');
const JSZip = require('jszip');
const fs = require('fs');
const { execFileSync } = require('child_process');
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/slides.html');
  await p.waitForTimeout(300);
  // Slide 1: title (fade on click), body (fly from left, after previous); slide 2 with a wipe transition
  await p.evaluate(() => {
    const s = AtlasSlides.pres.slides[0];
    s.elements = [
      { id: 't1', type: 'text', html: '<div>Title</div>', x: 80, y: 60, w: 800, h: 100, size: 40, align: 'left', valign: 'top', ph: 'title' },
      { id: 'b1', type: 'text', html: '<div>Body</div>', x: 80, y: 200, w: 800, h: 200, size: 24, align: 'left', valign: 'top', ph: 'body' },
      { id: 'x1', type: 'shape', shape: 'rect', fill: '#ff0000', html: '', x: 700, y: 400, w: 100, h: 60, size: 18, align: 'left', valign: 'top' }
    ];
    AtlasSlides.goTo(0);
  });
  await p.click('[data-tab="animations"]');
  const pick = async (id, fx) => { await p.click(`#stage [data-id="${id}"]`, { position: { x: 5, y: 5 } }); await p.selectOption('#animEffect', fx); };
  await pick('t1', 'fade');
  await pick('b1', 'fly'); await p.selectOption('#animDir', 'left'); await p.selectOption('#animStart', 'after'); await p.fill('#animDur', '0.3'); await p.dispatchEvent('#animDur', 'change');
  await pick('x1', 'fadeOut');
  const badges = await p.evaluate(() => [...document.querySelectorAll('#stage .anim-badge')].map(x => x.textContent).join());
  const anims = await p.evaluate(() => AtlasSlides.pres.slides[0].elements.map(e => e.anim && `${e.anim.effect}/${e.anim.dir}/${e.anim.start}/${e.anim.order}`).join());
  check('animations set from the ribbon with order badges', anims === 'fade/bottom/click/0,fly/left/after/1,fadeOut/bottom/click/2' && badges === '1,1,2', anims + ' | ' + badges);
  await p.evaluate(() => { AtlasSlides.addSlide('content'); const s = AtlasSlides.pres.slides[1]; s.transition = 'wipe'; s.transDur = 0.5; s.advance = 0.2; });
  // Slide show
  const vis = () => p.evaluate(() => ['t1', 'b1', 'x1'].map(id => { const n = document.querySelector(`#presentSlide [data-id="${id}"]`); return n ? (getComputedStyle(n).visibility === 'hidden' ? 0 : 1) : '-'; }).join(''));
  await p.evaluate(() => AtlasSlides.present(0)); await p.waitForTimeout(400);
  const v0 = await vis();
  await p.keyboard.press('ArrowRight'); await p.waitForTimeout(1200);
  const v1 = await vis();
  await p.keyboard.press('ArrowRight'); await p.waitForTimeout(700);
  const v2 = await vis();
  await p.keyboard.press('ArrowLeft'); await p.waitForTimeout(100);
  const vBack = await vis();
  check('builds: hidden entrances, click shows title then body after it, exit hides the box, Left undoes a build', v0 === '001' && v1 === '111' && v2 === '110' && vBack === '111', [v0, v1, v2, vBack].join(' '));
  await p.keyboard.press('ArrowRight'); await p.waitForTimeout(700);
  await p.keyboard.press('ArrowRight'); await p.waitForTimeout(150);
  const onSlide2 = await p.evaluate(() => document.querySelector('#presentSlide').getAnimations().length > 0);
  await p.waitForTimeout(1200);
  const ended = await p.evaluate(() => document.getElementById('presenter').hidden);
  check('slide 2 comes in with its transition and advances by itself', onSlide2 && ended, JSON.stringify({ onSlide2, ended }));
  // PPTX export: transitions and timing XML
  const bytes = await p.evaluate(async () => Array.from(await AtlasSlides.serialize('pptx')));
  fs.writeFileSync(SP + '/files/anim.pptx', Buffer.from(bytes));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const s1 = await zip.file('ppt/slides/slide1.xml').async('string'), s2 = await zip.file('ppt/slides/slide2.xml').async('string');
  const ids = Object.fromEntries([...s1.matchAll(/<p:cNvPr id="(\d+)" name="atlas-([^"]+)"/g)].map(m => [m[2], m[1]]));
  check('pptx: slide 2 transition with timing', /<\/p:clrMapOvr><p:transition spd="fast" advTm="200"><p:wipe dir="r"\/><\/p:transition>/.test(s2), (s2.match(/<p:transition[\s\S]*?<\/p:transition>/) || [''])[0]);
  check('pptx: effects target the right shapes', /presetID="10" presetClass="entr"[^>]*nodeType="clickEffect"[\s\S]*?<p:spTgt spid="(\d+)"/.exec(s1)[1] === ids.t1 && /presetID="2" presetClass="entr" presetSubtype="8"[^>]*nodeType="afterEffect"[\s\S]*?<p:spTgt spid="(\d+)"/.exec(s1)[1] === ids.b1 && /presetClass="exit"[\s\S]*?<p:spTgt spid="(\d+)"/.exec(s1)[1] === ids.x1, JSON.stringify(ids));
  const py = execFileSync('python3', ['-I', '-c', `
import sys
from pptx import Presentation
p = Presentation(sys.argv[1])
print(len(p.slides), [sh.name for sh in p.slides[0].shapes])`, SP + '/files/anim.pptx']).toString().trim();
  check('python-pptx opens the file', /^2 \['atlas-/.test(py), py);
  // Re-import
  const back = await p.evaluate(async b => { await AtlasSlides.loadFile({ name: 'a.pptx', path: null, data: new Uint8Array(b) }); const s = AtlasSlides.pres.slides; return { a: s[0].elements.map(e => e.anim && `${e.anim.effect}/${e.anim.dir}/${e.anim.start}/${e.anim.dur}`).join(), t: [s[1].transition, s[1].transDur, s[1].advance].join() }; }, bytes);
  check('pptx import restores animations and transitions', back.a === 'fade/bottom/click/0.5,fly/left/after/0.3,fadeOut/bottom/click/0.5' && back.t === 'wipe,0.5,0.2', JSON.stringify(back));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
