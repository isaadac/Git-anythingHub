// Write, pro features: the Sources manager (search, BibTeX/RIS/DOI import, cite, BibTeX export,
// IEEE and Harvard styles) and Compare side by side.
const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

const BIB = `@article{lovelace1843,
  author = {Lovelace, Ada and Babbage, Charles},
  title = {Notes on the {Analytical Engine}},
  journal = {Scientific Memoirs}, year = {1843}, volume = {3}, pages = {666--731}
}
@book{knuth1968, author = {Knuth, Donald E.}, title = {The Art of Computer Programming}, publisher = {Addison-Wesley}, year = {1968}}`;
const CROSSREF = { message: { DOI: '10.1038/171737a0', type: 'journal-article', title: ['Molecular Structure of Nucleic Acids'], author: [{ family: 'Watson', given: 'James D.' }, { family: 'Crick', given: 'Francis' }], 'container-title': ['Nature'], volume: '171', issue: '4356', page: '737-738', issued: { 'date-parts': [[1953]] } } };

(async () => {
  const browser = await chromium.launch();
  const p = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  let lookups = 0;
  await p.route('https://api.crossref.org/works/**', route => { lookups++; route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(CROSSREF), headers: { 'Access-Control-Allow-Origin': '*' } }); });
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.waitForFunction(() => window.AtlasWrite && window.WriteRefs);
  await p.evaluate(() => localStorage.removeItem('atlas-sources'));
  await p.click('#doc');
  await p.keyboard.type('Ada wrote the first program');
  await p.evaluate(() => document.querySelector('.ribbon-tabs button[data-tab="references"], .ribbon-tabs button[data-tab="insert"]').click());

  // ---- The Sources pane
  await p.evaluate(() => document.getElementById('manageSources').click());
  await p.waitForSelector('.refs-pane');
  check('the Sources pane opens, empty', (await p.$$eval('.refs-pane .refs-group', g => g.map(x => x.textContent))).join() === 'In this document (0),Library (0)');
  // Paste BibTeX.
  await p.click('.refs-pane .paste');
  await p.fill('.modal textarea', BIB);
  await p.click('.modal .actions .primary');
  await p.waitForFunction(() => document.querySelectorAll('.refs-pane .ref').length === 2);
  check('pasted BibTeX becomes sources', (await p.$$eval('.refs-pane .ref .who', w => w.map(x => x.textContent))).join() === 'Lovelace & Babbage (1843),Knuth (1968)');
  // Add by DOI (Crossref).
  await p.click('.refs-pane .doi');
  await p.fill('.modal input', 'https://doi.org/10.1038/171737a0');
  await p.click('.modal .actions .primary');
  await p.waitForFunction(() => document.querySelectorAll('.refs-pane .ref').length === 3, null, { timeout: 5000 }).catch(() => {});
  check('a DOI is looked up and added', lookups === 1 && (await p.$$eval('.refs-pane .ref .title', t => t.map(x => x.textContent))).includes('Molecular Structure of Nucleic Acids'));
  // Pasting the same again adds nothing.
  await p.click('.refs-pane .paste');
  await p.fill('.modal textarea', BIB);
  await p.click('.modal .actions .primary');
  await sleep(200);
  check('sources already there are not added twice', (await p.$$('.refs-pane .ref')).length === 3);
  // Search.
  await p.fill('.refs-pane .refs-search', 'nucleic');
  check('search narrows the list', (await p.$$eval('.refs-pane .ref .who', w => w.map(x => x.textContent))).join() === 'Watson & Crick (1953)');
  await p.fill('.refs-pane .refs-search', '');

  // Cite from the pane, where the cursor was.
  await p.click('.refs-pane .ref:has-text("Lovelace") .acts button:has-text("Cite")');
  check('Cite puts the citation at the cursor', /^Ada wrote the first program \(Lovelace & Babbage, 1843\)$/.test((await p.textContent('#doc p')).replace(/ /g, ' ')), await p.textContent('#doc p'));
  check('…and the pane counts it', /cited 1×/.test(await p.textContent('.refs-pane .ref:has-text("Lovelace") .meta')));
  await p.evaluate(() => { const d = document.getElementById('doc'); const r = document.createRange(); r.selectNodeContents(d.lastElementChild); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r); });
  await p.keyboard.type(', and DNA was described');
  await p.click('.refs-pane .ref:has-text("Watson") .acts button:has-text("Cite")');
  await p.click('.refs-pane .ref:has-text("Lovelace") .acts button:has-text("Cite")');
  await p.evaluate(() => document.getElementById('insBibliography').click());

  // IEEE: numbers by first citation, a numbered list.
  await p.selectOption('#citeStyle', 'ieee');
  const ieee = await p.evaluate(() => ({ cites: [...document.querySelectorAll('#doc cite.citation')].map(c => c.textContent), bib: [...document.querySelectorAll('#doc section.bibliography .bib-entry')].map(e => e.textContent) }));
  check('IEEE: citations are numbered in order of first use', ieee.cites.join() === '[1],[2],[1]', ieee.cites.join());
  check('…and the reference list follows that order', ieee.bib.length === 2 && /^\[1\] A\. Lovelace and C\. Babbage, “Notes on the Analytical Engine,” Scientific Memoirs, vol\. 3, pp\. 666–731, 1843\.$/.test(ieee.bib[0]) && /^\[2\] J\. D\. Watson and F\. Crick, “Molecular Structure of Nucleic Acids,” Nature, vol\. 171, no\. 4356, pp\. 737–738, 1953, doi: 10\.1038\/171737a0\.$/.test(ieee.bib[1]), ieee.bib.join(' | '));
  await p.selectOption('#citeStyle', 'harvard');
  const harv = await p.evaluate(() => ({ cite: document.querySelector('#doc cite.citation').textContent, first: document.querySelector('#doc .bib-entry').textContent, head: document.querySelector('#doc section.bibliography h2').textContent }));
  check('Harvard: author-date citations and entries', harv.cite === '(Lovelace and Babbage 1843)' && /^Lovelace, A\. and Babbage, C\. \(1843\) ‘Notes on the Analytical Engine’, Scientific Memoirs, 3, pp\. 666–731\.$/.test(harv.first) && harv.head === 'References', JSON.stringify(harv));

  // Export BibTeX.
  const saved = await p.evaluate(async () => {
    let out = null;
    Platform.save = async o => { out = { name: o.name, text: new TextDecoder().decode(o.data) }; return { name: o.name }; };
    document.querySelector('.refs-pane .export').click();
    await new Promise(r => setTimeout(r, 100));
    return out;
  });
  check('BibTeX export has every source', saved && saved.name === 'references.bib' && (saved.text.match(/^@/gm) || []).length === 3 && /@article\{watson1953molecular,/.test(saved.text), saved && saved.text.slice(0, 200));
  // RIS import from a file.
  await p.evaluate(() => { Platform.open = async () => ({ name: 'scholar.ris', data: new TextEncoder().encode('TY  - RPRT\nAU  - Turing, Alan\nTI  - Proposed Electronic Calculator\nPB  - National Physical Laboratory\nPY  - 1945\nER  - \n') }); });
  await p.click('.refs-pane .import');
  await p.waitForFunction(() => document.querySelectorAll('.refs-pane .ref').length === 4);
  check('RIS files import', (await p.$$eval('.refs-pane .ref .who', w => w.map(x => x.textContent))).includes('Turing (1945)'));
  check('an uncited source can be removed', !!(await p.$('.refs-pane .ref:has-text("Turing") .acts button:has-text("Remove")')) && !(await p.$('.refs-pane .ref:has-text("Lovelace") .acts button:has-text("Remove")')));
  // The library keeps sources for other documents.
  await p.click('.refs-pane .ref:has-text("Turing") .acts button:has-text("Remove")');
  check('…and stays in the library for other documents', (await p.$$eval('.refs-pane .refs-group', g => g.map(x => x.textContent))).join() === 'In this document (3),Library (1)');
  // The sources go with the file.
  const html = await p.evaluate(() => AtlasWrite.serialize ? AtlasWrite.serialize('html') : '');
  if (html) check('sources (with DOIs) are saved in the document', /atlas-sources/.test(html) && /171737a0/.test(html));
  await p.click('.refs-pane .x');

  // ---- Compare side by side
  await p.evaluate(() => {
    const earlier = '<p>The quick brown fox jumps over the lazy dog.</p><p>A paragraph that was removed later.</p><p>Closing words stay.</p>';
    document.getElementById('doc').innerHTML = '<p>The quick red fox leaps over the lazy dog.</p><p>Closing words stay.</p><p>A brand new paragraph.</p>';
    Platform.open = async () => ({ name: 'earlier.html', data: new TextEncoder().encode(earlier) });
    Platform.openData = async (app, name, b) => { window.__compared = { app, name, html: new TextDecoder().decode(b) }; return true; };
  });
  await p.evaluate(() => document.getElementById('compareDocs').click());
  await p.waitForSelector('.compare-view');
  const cv = await p.evaluate(() => {
    const rows = [...document.querySelectorAll('.compare-view .cmp-row')];
    return {
      heads: [...document.querySelectorAll('.compare-view .cmp-heads > div')].map(d => d.textContent),
      rows: rows.map(r => [r.querySelector('.l').textContent.replace(/\s+/g, ' ').trim(), r.querySelector('.r').textContent.replace(/\s+/g, ' ').trim(), r.classList.contains('changed')]),
      leftDel: [...document.querySelectorAll('.compare-view .cmp-cell.l del.tc')].map(n => n.textContent.trim()).filter(Boolean),
      rightIns: [...document.querySelectorAll('.compare-view .cmp-cell.r ins.tc')].map(n => n.textContent.trim()).filter(Boolean),
      noInsLeft: !document.querySelector('.compare-view .cmp-cell.l ins'), noDelRight: !document.querySelector('.compare-view .cmp-cell.r del'),
      gaps: [...document.querySelectorAll('.compare-view .cmp-cell.gap')].map(g => g.classList.contains('l') ? 'l' : 'r'),
      sum: document.querySelector('.compare-view .sum').textContent
    };
  });
  check('compare opens side by side: earlier on the left, this document on the right', cv.heads[0] === 'earlier.html' && cv.rows.length === 4 && cv.noInsLeft && cv.noDelRight, JSON.stringify(cv.rows));
  check('…removed words marked on the left, added words on the right', cv.leftDel.includes('brown') && cv.leftDel.includes('jumps') && cv.rightIns.includes('red') && cv.rightIns.includes('leaps'), JSON.stringify([cv.leftDel, cv.rightIns]));
  check('…a removed or added paragraph faces a gap, so the two sides line up', cv.gaps.join() === 'r,l' && cv.rows[1][0] === 'A paragraph that was removed later.' && cv.rows[3][1] === 'A brand new paragraph.', JSON.stringify(cv.rows));
  check('…with a summary of the changes', cv.sum === '3 changed paragraphs · 3 insertions · 3 deletions', cv.sum);
  const rowsTop = await p.evaluate(() => [...document.querySelectorAll('.compare-view .cmp-row')].map(r => { const [a, b] = r.querySelectorAll('.cmp-cell'); return Math.abs(a.getBoundingClientRect().top - b.getBoundingClientRect().top) + Math.abs(a.offsetHeight - b.offsetHeight); }));
  check('…each paragraph sits level with its counterpart', rowsTop.every(d => d < 1), rowsTop.join());
  const cur = () => p.evaluate(() => [...document.querySelectorAll('.compare-view .cmp-row')].findIndex(r => r.classList.contains('current')));
  const first = await cur();
  await p.click('.compare-view .next');
  const second = await cur();
  await p.keyboard.press('Alt+ArrowDown');
  const third = await cur();
  await p.keyboard.press('Alt+ArrowUp');
  check('Next and Previous step through the changes', first === 0 && second === 1 && third === 3 && (await cur()) === 1, [first, second, third].join());
  await p.click('.compare-view .tracked');
  await p.waitForFunction(() => window.__compared, null, { timeout: 5000 });
  check('…and they open as tracked changes on request', !(await p.$('.compare-view')) && (await p.evaluate(() => /<ins class="tc"/.test(window.__compared.html))));
  await p.evaluate(() => document.getElementById('compareDocs').click());
  await p.waitForSelector('.compare-view');
  await p.keyboard.press('Escape');
  check('Escape closes the comparison', !(await p.$('.compare-view')));

  await p.evaluate(() => { window.__atlasDirty = false; });
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
