const { chromium } = require('playwright');
const JSZip = require('jszip');
const SHOTS = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  // 150 paragraphs of varying length, a heading, a list and a page break
  await p.evaluate(() => {
    const parts = [];
    for (let i = 1; i <= 150; i++) {
      if (i === 40) parts.push('<h1>Chapter Two</h1>');
      if (i === 60) parts.push('<ul>' + Array.from({ length: 50 }, (_, k) => `<li>List item ${k + 1}</li>`).join('') + '</ul>');
      parts.push(`<p>Paragraph ${i}. ` + 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. '.repeat(1 + (i % 5)) + '</p>');
      if (i === 100) parts.push('<hr class="page-break">');
    }
    document.getElementById('doc').innerHTML = parts.join('');
    document.getElementById('doc').dispatchEvent(new Event('input'));
  });
  await p.waitForTimeout(400);
  const geo = await p.evaluate(() => {
    const doc = document.getElementById('doc');
    const H = 11 * 96, M = 96, pitch = H + 24, body = H - 2 * M;
    const o = doc.getBoundingClientRect().top;
    const bad = [];
    doc.querySelectorAll('p, h1, li').forEach(el => {
      const r = el.getBoundingClientRect();
      const top = r.top - o, bottom = r.bottom - o;
      const pg = Math.floor((top + 0.5) / pitch);
      if (top < pg * pitch - 0.5 || bottom > pg * pitch + body + 1) bad.push(el.textContent.slice(0, 20) + ` [${top.toFixed(0)}-${bottom.toFixed(0)}]`);
    });
    const after = [...doc.querySelectorAll('p')].find(e => e.textContent.startsWith('Paragraph 101.'));
    const t = after.getBoundingClientRect().top - o;
    return { bad, pages: document.querySelectorAll('.page-bg').length, status: document.getElementById('stPages').textContent, afterBreakTop: Math.min(t % pitch, pitch - (t % pitch)), pageHeight: document.getElementById('page').style.height };
  });
  check('long document is split into pages', geo.pages > 5, geo.pages + ' pages, ' + geo.status);
  check('no paragraph or list item crosses a page edge', geo.bad.length === 0, geo.bad.slice(0, 3).join(' | '));
  check('page break starts the next paragraph at the top of a page', Math.abs(geo.afterBreakTop) < 1, geo.afterBreakTop.toFixed(2));
  // Page numbers in the footer, different first page
  await p.click('[data-tab="insert"]');
  await p.selectOption('#insPageNum', 'footer:center:nofm');
  await p.evaluate(() => { const h = document.getElementById('pageHeader'); h.innerHTML = '<p>Running head</p>'; h.dispatchEvent(new Event('input')); });
  await p.check('#differentFirst');
  await p.evaluate(() => { const h = document.getElementById('firstHeader'); h.innerHTML = '<p>Cover page</p>'; h.dispatchEvent(new Event('input')); });
  await p.waitForTimeout(400);
  const hf = await p.evaluate(() => {
    const bgs = [...document.querySelectorAll('.page-bg')];
    const text = i => [...bgs[i].querySelectorAll('.hf-clone')].map(c => c.textContent.trim()).join(' | ');
    const ed = id => { const e = document.getElementById(id); return e.hidden ? null : e.textContent.trim(); };
    return { n: bgs.length, p3: text(2), last: text(bgs.length - 1), firstHeader: ed('firstHeader'), mainHeaderOnP2: ed('pageHeader'), footerP2: ed('pageFooter') };
  });
  check('page 3 shows the running head and "Page 3 of N"', hf.p3 === `Running head | Page 3 of ${hf.n}`, hf.p3);
  check('last page footer numbers correctly', hf.last.endsWith(`Page ${hf.n} of ${hf.n}`), hf.last);
  check('first page has its own header', hf.firstHeader === 'Cover page', String(hf.firstHeader));
  check('editable running head sits on page 2', hf.mainHeaderOnP2 === 'Running head' && hf.footerP2 === `Page 2 of ${hf.n}`, JSON.stringify(hf));
  await p.evaluate(() => document.getElementById('canvas').scrollTo(0, 0));
  await p.screenshot({ path: SHOTS + '/write-pages.png' });
  await p.evaluate(() => document.getElementById('canvas').scrollTo(0, 1000));
  await p.screenshot({ path: SHOTS + '/write-pages-2.png' });
  // DOCX export has real fields and a title page
  const bytes = await p.evaluate(async () => Array.from(await AtlasWrite.serialize('docx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const docXml = await zip.file('word/document.xml').async('string');
  const parts = await Promise.all(Object.keys(zip.files).filter(f => /word\/(header|footer)\d+\.xml$/.test(f)).map(f => zip.file(f).async('string')));
  const all = parts.join('');
  check('docx: different first page flag', /<w:titlePg\/>/.test(docXml));
  check('docx: PAGE and NUMPAGES fields in footer', /PAGE<\/w:instrText>/.test(all) && /NUMPAGES<\/w:instrText>/.test(all));
  check('docx: header text exported', /Running head/.test(all) && /Cover page/.test(all));
  check('docx: no layout margins leak into the file', !/data-pg/.test(docXml));
  // Re-import: everything comes back
  await p.evaluate(async b => { await AtlasWrite.loadFile({ name: 'pages.docx', path: null, data: new Uint8Array(b) }); }, bytes);
  await p.waitForTimeout(400);
  const back = await p.evaluate(() => ({
    first: document.getElementById('firstHeader').textContent.trim(), header: document.getElementById('pageHeader').innerHTML,
    footer: document.getElementById('pageFooter').innerHTML, diff: document.getElementById('differentFirst').checked, pages: document.querySelectorAll('.page-bg').length
  }));
  check('docx round trip keeps headers, page fields and first-page setting', back.first === 'Cover page' && /Running head/.test(back.header) && /data-fld="numpages"/.test(back.footer) && /Page <span/.test(back.footer) && /<\/span> of <span/.test(back.footer) && back.diff, JSON.stringify(back).slice(0, 300));
  // HTML round trip
  const html = await p.evaluate(() => AtlasWrite.serialize('html'));
  check('html export has no layout margins', !/data-pg/.test(html));
  await p.evaluate(async h => { await AtlasWrite.loadFile({ name: 'x.html', path: null, data: new TextEncoder().encode(h) }); }, html);
  const back2 = await p.evaluate(() => ({ first: document.getElementById('firstHeader').textContent.trim(), fld: document.querySelectorAll('#pageFooter .fld').length, diff: document.getElementById('differentFirst').checked }));
  check('html round trip keeps first page header and fields', back2.first === 'Cover page' && back2.fld === 2 && back2.diff, JSON.stringify(back2));
  // Print layout: no gaps, sheet height exact
  const pr = await p.evaluate(() => { window.dispatchEvent(new Event('beforeprint')); const h = document.getElementById('page').style.height; const n = document.querySelectorAll('.page-bg').length; window.dispatchEvent(new Event('afterprint')); return { h, n, after: document.getElementById('page').style.height }; });
  check('print layout is exactly N sheets tall', parseFloat(pr.h) === pr.n * 11 * 96, JSON.stringify(pr));
  // Typing at the end keeps working and layout follows
  await p.evaluate(() => { const d = document.getElementById('doc'); const r = document.createRange(); r.selectNodeContents(d.lastElementChild); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); d.focus(); });
  for (let i = 0; i < 3; i++) { await p.keyboard.press('Enter'); await p.keyboard.type('New text typed at the end of the document. '.repeat(8)); }
  await p.waitForTimeout(400);
  const end = await p.evaluate(() => document.getElementById('stPages').textContent);
  check('status bar shows the caret page', /^Page \d+ of \d+$/.test(end), end);
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e); process.exit(1); });
