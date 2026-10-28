// Write: section breaks, wrapped pictures, equations, fields, editable styles and Compare,
// in the editor and through a .docx round trip.
const { chromium } = require('playwright');
const JSZip = require('jszip');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAIAAAACUFjqAAAAEklEQVR4nGP4z8CAB+GTG8HSALfKY52fTcuYAAAAAElFTkSuQmCC';

(async () => {
  const browser = await chromium.launch();
  const p = await browser.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.waitForFunction(() => window.AtlasWrite && window.WriteExtras);
  const form = async (btn, fill) => {
    await p.evaluate(id => document.getElementById(id).click(), btn);
    await p.waitForSelector('.modal');
    await fill();
    await p.click('.modal .actions .primary');
    await p.waitForSelector('.modal', { state: 'detached' });
  };
  await p.click('#doc');
  await p.keyboard.type('Intro paragraph with ');

  // Equation
  await form('insEquation', async () => { await p.fill('.modal textarea', 'x = (-b \\pm sqrt(b^2 - 4ac))/(2a)'); });
  check('equation drawn with MathML (fraction, root, ±)', await p.evaluate(() => { const m = document.querySelector('#doc span.equation math'); return !!m && !!m.querySelector('mfrac') && !!m.querySelector('msqrt') && m.textContent.includes('±'); }));
  // Field
  await p.keyboard.press('End');
  await p.keyboard.type(' dated ');
  await form('insField', async () => { await p.selectOption('.modal select', 'date'); });
  const year = String(new Date().getFullYear());
  check('date field shows today', await p.evaluate(y => (document.querySelector('#doc span.fld[data-fld=date]') || {}).textContent?.includes(y), year));
  await p.keyboard.press('End');
  await p.keyboard.press('Enter');
  await p.keyboard.type('Heading text');
  await p.evaluate(() => document.querySelector('.style-btn.s-h1').click());
  await p.keyboard.press('End');
  await p.keyboard.press('Enter');
  await p.evaluate(() => document.querySelector('.style-btn[data-block=P]').click());
  await p.keyboard.type('Before the picture. ');
  await p.evaluate(src => { document.execCommand('insertHTML', false, `<img src="${src}" alt="Square" style="width:25%">`); }, PNG);
  await p.keyboard.type('Text that wraps around the picture on its left side and keeps going.');
  // Wrap the picture
  await p.$eval('#doc img', i => i.click());
  await p.evaluate(() => document.querySelector('#imgTools [data-wrap="float-right"]').click());
  check('picture floats right with text around it', await p.evaluate(() => { const i = document.querySelector('#doc img'); return i.classList.contains('float-right') && getComputedStyle(i).float === 'right'; }));
  // Section break: landscape, two columns
  await p.keyboard.press('Control+End');
  await form('insSectionBreak', async () => { await p.selectOption('.modal select >> nth=0', 'landscape'); await p.selectOption('.modal select >> nth=1', '2'); });
  await p.keyboard.type('Second section text.');
  check('section break in the document', await p.evaluate(() => { const h = document.querySelector('#doc hr.section-break'); return !!h && h.dataset.orient === 'landscape' && h.dataset.cols === '2'; }));
  // Modify Heading 1
  await p.click('#doc h1');
  await form('modifyStyle', async () => {
    await p.selectOption('.modal select >> nth=0', 'h1');
    await p.fill('.modal input[type=color]', '#c00000');
    await p.fill('.modal input[type=number] >> nth=0', '20');
  });
  check('modified style applies to every Heading 1', await p.evaluate(() => { const cs = getComputedStyle(document.querySelector('#doc h1')); return cs.color === 'rgb(192, 0, 0)' && Math.round(parseFloat(cs.fontSize)) === Math.round(20 * 96 / 72); }));

  // .docx
  const bytes = await p.evaluate(async () => Array.from(await AtlasWrite.serialize('docx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const xml = await zip.file('word/document.xml').async('string');
  const styles = await zip.file('word/styles.xml').async('string');
  check('docx: equation as Office Math', /<m:oMath>[\s\S]*<m:f>[\s\S]*<m:rad>/.test(xml) && /xmlns:m=/.test(xml));
  check('docx: date as a Word field', /<w:fldSimple w:instr=" DATE \\@ &quot;d MMMM yyyy&quot; ">/.test(xml));
  const sects = xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) || [];
  check('docx: two sections, the second landscape with two columns', sects.length === 2 && /w:orient="landscape"/.test(sects[1]) && /w:num="2"/.test(sects[1]) && !/w:orient="landscape"/.test(sects[0]), sects.length);
  check('docx: wrapped picture anchored with square wrapping on the right', /<wp:anchor\b[\s\S]*<wp:align>right<\/wp:align>[\s\S]*<wp:wrapSquare wrapText="bothSides"\/>[\s\S]*<\/wp:anchor>/.test(xml) && !/<wp:inline/.test(xml));
  check('docx: Heading 1 style carries the change', /w:styleId="Heading1"[\s\S]*?<w:color w:val="C00000"\/>/i.test(styles) && /w:styleId="Heading1"[\s\S]*?<w:sz w:val="40"\/>/.test(styles));
  // Reopen it.
  await p.evaluate(async b => { await AtlasWrite.loadFile({ name: 'objects.docx', data: new Uint8Array(b) }); }, bytes);
  await sleep(300);
  const back = await p.evaluate(() => ({
    eq: document.querySelector('#doc span.equation')?.dataset.eq,
    eqOk: (() => { const m = document.querySelector('#doc span.equation math'); return !!m && !!m.querySelector('mfrac') && !!m.querySelector('msqrt'); })(),
    fld: document.querySelector('#doc span.fld')?.dataset.fld,
    sec: (() => { const h = document.querySelector('#doc hr.section-break'); return h ? h.dataset.orient + '/' + h.dataset.cols : null; })(),
    img: document.querySelector('#doc img')?.className,
    h1: getComputedStyle(document.querySelector('#doc h1')).color
  }));
  check('reopened: the equation is an equation again', back.eqOk, back.eq);
  check('reopened: the field is a live date field', back.fld === 'date');
  check('reopened: the section break with its orientation and columns', back.sec === 'landscape/2', back.sec);
  check('reopened: the picture still floats right', /float-right/.test(back.img || ''), back.img);
  check('reopened: Heading 1 keeps its style', back.h1 === 'rgb(192, 0, 0)', back.h1);

  // Markdown and HTML
  const md = await p.evaluate(() => AtlasWrite.serialize('md'));
  check('markdown: the equation in $…$', /\$x\s*=.*sqrt.*\$/.test(md), md.slice(0, 120));
  const html = await p.evaluate(() => AtlasWrite.serialize('html'));
  check('web page: MathML and the document styles', /<math/.test(html) && /main\.atlas-doc h1:not\(\.title\) \{[^}]*color: #c00000/.test(html));

  // Compare an earlier version with this one.
  await p.evaluate(() => {
    const earlier = '<p>The quick brown fox jumps over the lazy dog.</p><p>A paragraph that was removed later.</p><p>Closing words stay.</p>';
    document.getElementById('doc').innerHTML = '<p>The quick red fox leaps over the lazy dog.</p><p>Closing words stay.</p><p>A brand new paragraph.</p>';
    Platform.open = async () => ({ name: 'earlier.html', data: new TextEncoder().encode(earlier) });
    Platform.openData = async (app, name, b) => { window.__compared = { app, name, html: new TextDecoder().decode(b) }; return true; };
  });
  await p.evaluate(() => document.getElementById('compareDocs').click());
  await p.waitForFunction(() => window.__compared, null, { timeout: 5000 });
  const cmp = await p.evaluate(() => { const d = new DOMParser().parseFromString(window.__compared.html, 'text/html'); const t = sel => [...d.querySelectorAll(sel)].map(n => n.textContent.trim()).filter(Boolean); return { ins: t('ins.tc'), del: t('del.tc'), author: d.querySelector('ins.tc')?.dataset.author, app: window.__compared.app }; });
  check('compare: changed words marked as insertions and deletions', cmp.ins.includes('red') && cmp.ins.includes('leaps') && cmp.del.includes('brown') && cmp.del.includes('jumps'), JSON.stringify(cmp));
  check('compare: removed and added paragraphs marked', cmp.del.some(t => /removed later/.test(t)) && cmp.ins.some(t => /brand new paragraph/.test(t)), JSON.stringify(cmp));
  check('compare: result opens as a new Write document, changes by "Compare"', cmp.app === 'write' && cmp.author === 'Compare');
  check('compare: this document is left as it was', await p.evaluate(() => !document.querySelector('#doc ins, #doc del')));

  check('no page errors', !errors.length, errors.join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
