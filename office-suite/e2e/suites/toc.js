const { chromium } = require('playwright');
const JSZip = require('/home/user/Git-anythingHub/office-suite/node_modules/jszip');
const fs = require('fs');
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.evaluate(() => {
    const filler = n => Array.from({ length: n }, (_, i) => `<p>Paragraph ${i + 1}. ${'Text that fills the line nicely. '.repeat(5)}</p>`).join('');
    document.getElementById('doc').innerHTML = '<p>Intro line</p><h1>Getting started</h1>' + filler(25) + '<h2>Installing &amp; setup</h2>' + filler(20) + '<h1>Using it</h1>' + filler(10) + '<h3>Deep detail</h3><p>x</p>';
    document.getElementById('doc').dispatchEvent(new Event('input'));
    const el = document.querySelector('#doc p'); const r = document.createRange(); r.selectNodeContents(el); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); document.getElementById('doc').focus();
  });
  await p.click('[data-tab="insert"]');
  await p.click('#insToc');
  await p.waitForTimeout(400);
  const t = await p.evaluate(() => {
    const toc = document.querySelector('#doc .toc');
    const pageOf = h => { const z = 1; return Math.floor((h.getBoundingClientRect().top - document.getElementById('doc').getBoundingClientRect().top) / (document.querySelector('.page-bg').offsetHeight + 24)) + 1; };
    return {
      after: toc.previousElementSibling && toc.previousElementSibling.textContent,
      entries: [...toc.querySelectorAll('p a')].map(a => a.textContent + '|' + a.closest('p').className + '|' + a.closest('p').querySelector('.toc-pg').textContent),
      real: [...document.querySelectorAll('#doc h1,#doc h2,#doc h3')].map(pageOf),
      ce: toc.contentEditable, pages: document.getElementById('stPages').textContent
    };
  });
  check('TOC inserted at the cursor with one entry per heading and level', t.after === 'Intro line' && t.entries.length === 4 && /toc-l2/.test(t.entries[1]) && /toc-l3/.test(t.entries[3]) && t.ce === 'false', JSON.stringify(t));
  check('TOC page numbers match where the headings are', t.entries.map(e => e.split('|')[2]).join() === t.real.join() && t.real[2] > 1, t.entries.map(e => e.split('|')[2]).join() + ' vs ' + t.real.join());
  // Click an entry: jumps to the heading, URL unchanged
  const url0 = p.url();
  await p.click('#doc .toc p:nth-child(4) a');
  await p.waitForTimeout(700);
  const jump = await p.evaluate(() => { const h = [...document.querySelectorAll('#doc h1')][1]; const r = h.getBoundingClientRect(); return { visible: r.top >= 0 && r.top < innerHeight, caret: getSelection().anchorNode && (getSelection().anchorNode === h || h.contains(getSelection().anchorNode)) }; });
  check('clicking an entry jumps to its heading', jump.visible && jump.caret && p.url() === url0, JSON.stringify(jump));
  // Add a heading and update
  await p.evaluate(() => { document.querySelector('#doc h3').insertAdjacentHTML('beforebegin', '<h2>New section</h2>'); document.getElementById('doc').dispatchEvent(new Event('input')); });
  await p.click('#updToc'); await p.waitForTimeout(300);
  const up = await p.evaluate(() => [...document.querySelectorAll('#doc .toc p a')].map(a => a.textContent));
  check('Update Table picks up new headings', up.join('|') === 'Getting started|Installing & setup|Using it|New section|Deep detail', up.join('|'));
  // Pagination grows and page numbers follow
  await p.evaluate(() => { document.querySelector('#doc h1').insertAdjacentHTML('afterend', Array.from({ length: 30 }, () => '<p>More text more text more text more text more text more text more text.</p>').join('')); document.getElementById('doc').dispatchEvent(new Event('input')); });
  await p.waitForTimeout(600);
  const moved = await p.evaluate(() => [...document.querySelectorAll('#doc .toc .toc-pg')].map(s => s.textContent).join());
  check('page numbers update with the layout', moved !== t.entries.map(e => e.split('|')[2]).join() && !p.dirty, moved);
  // HTML round trip
  const html = await p.evaluate(() => AtlasWrite.serialize('html'));
  check('html keeps the contents', /class="toc"/.test(html) && /toc-l2/.test(html) && /href="#toc_/.test(html) && !/contenteditable/.test(html.match(/<div class="toc"[^>]*>/)[0]));
  // Markdown
  const md = await p.evaluate(() => AtlasWrite.serialize('md'));
  check('markdown contents are links without page numbers', /\[Getting started\]\(#toc_[a-z0-9]+\)\n/.test(md), md.slice(0, 160));
  // docx: Word TOC field with hyperlinks + PAGEREFs to bookmarks on the headings
  const bytes = await p.evaluate(async () => Array.from(await AtlasWrite.serialize('docx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes)); const d = await zip.file('word/document.xml').async('string');
  const anchors = [...d.matchAll(/<w:hyperlink w:anchor="([^"]+)"/g)].map(m => m[1]);
  const bms = [...d.matchAll(/<w:bookmarkStart[^>]*w:name="([^"]+)"/g)].map(m => m[1]);
  check('docx: Word table of contents field', /<w:docPartGallery w:val="Table of Contents"\/>/.test(d) && /TOC \\o (&quot;|")1-3(&quot;|") \\h \\z \\u/.test(d) && (d.match(/PAGEREF /g) || []).length === 5 && /<w:t xml:space="preserve">Contents<\/w:t>/.test(d), (d.match(/<w:sdt>[\s\S]{0,600}/) || [''])[0]);
  check('docx: every entry links to a bookmark on its heading', anchors.length === 5 && anchors.every(a => bms.includes(a)), JSON.stringify({ anchors, bms }));
  check('docx: no markers left', !/[\uE050-\uE052\uE0FF]/.test(d));
  const begins = (d.match(/fldCharType="begin"/g) || []).length, ends = (d.match(/fldCharType="end"/g) || []).length;
  check('docx: fields balanced', begins === ends && begins === 6, begins + '/' + ends);
  fs.writeFileSync(SP + '/files/toc-out.docx', Buffer.from(bytes));
  // Re-open our own docx
  await p.evaluate(async b => { await AtlasWrite.loadFile({ name: 't.docx', path: null, data: new Uint8Array(b) }); }, bytes);
  await p.waitForTimeout(600);
  const back = await p.evaluate(() => { const toc = document.querySelector('#doc .toc'); return toc && { n: toc.querySelectorAll('p a').length, pg: [...toc.querySelectorAll('.toc-pg')].map(s => s.textContent).join(), ok: [...toc.querySelectorAll('p a')].every(a => document.getElementById(a.getAttribute('href').slice(1))?.matches('h1,h2,h3')), dup: document.querySelectorAll('#doc p').length && [...document.querySelectorAll('#doc p')].filter(p => !p.closest('.toc') && /^Getting started/.test(p.textContent)).length }; });
  check('our .docx reopens with a live contents', back && back.n === 5 && back.ok && back.dup === 0 && /\d/.test(back.pg), JSON.stringify(back));
  // A Word document's contents
  await p.evaluate(async b => { await AtlasWrite.loadFile({ name: 'w.docx', path: null, data: new Uint8Array(b) }); }, Array.from(fs.readFileSync(SP + '/files/wordtoc.docx')));
  await p.waitForTimeout(600);
  const w = await p.evaluate(() => { const toc = document.querySelector('#doc .toc'); return toc && { title: toc.querySelector('.toc-title').textContent, e: [...toc.querySelectorAll('p a')].map(a => a.textContent + '>' + (document.getElementById(a.getAttribute('href').slice(1))?.tagName || '?') + '@' + a.closest('p').querySelector('.toc-pg').textContent), first: document.querySelector('#doc').firstElementChild.className }; });
  check("Word's table of contents opens as a live contents", w && w.title === 'Contents' && w.first === 'toc' && w.e.map(x => x.replace(/@\d+$/, '@n')).join() === 'Alpha chapter>H1@n,Alpha detail>H2@n,Beta chapter>H1@n', JSON.stringify(w));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
