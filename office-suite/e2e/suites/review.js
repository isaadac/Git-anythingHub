const { chromium } = require('playwright');
const JSZip = require('jszip');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.evaluate(() => {
    localStorage.setItem('atlas-author', 'Ada Lovelace');
    document.getElementById('doc').innerHTML = '<p>The quick brown fox jumps.</p><p>Second paragraph here.</p>';
  });
  const caretAfter = async (text, offsetInText) => p.evaluate(([t, o]) => {
    const w = document.createTreeWalker(document.getElementById('doc'), NodeFilter.SHOW_TEXT);
    while (w.nextNode()) { const i = w.currentNode.nodeValue.replace(/\u00a0/g, ' ').indexOf(t); if (i >= 0) { const r = document.createRange(); r.setStart(w.currentNode, i + o); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); document.getElementById('doc').focus(); return true; } }
    return false;
  }, [text, offsetInText]);
  await p.click('[data-tab="review"]');
  await p.click('#trackChanges');
  check('track changes switch is on', await p.evaluate(() => document.getElementById('trackChanges').classList.contains('on')));
  // Type an insertion after "brown"
  await caretAfter('brown', 5);
  await p.keyboard.type(' red');
  let html = await p.$eval('#doc', d => d.innerHTML.replace(/&nbsp;/g, ' '));
  check('typing is recorded as one tracked insertion', /<ins class="tc" data-author="Ada Lovelace" datetime="[^"]+"> red<\/ins>/.test(html), html);
  // Select "quick " and delete
  await p.evaluate(() => { const w = document.createTreeWalker(document.getElementById('doc'), NodeFilter.SHOW_TEXT); w.nextNode(); const n = w.currentNode; const r = document.createRange(); r.setStart(n, 4); r.setEnd(n, 10); getSelection().removeAllRanges(); getSelection().addRange(r); });
  await p.keyboard.press('Backspace');
  html = await p.$eval('#doc', d => d.innerHTML.replace(/&nbsp;/g, ' '));
  check('deleting a selection marks it deleted instead of removing it', /<del class="tc"[^>]*>quick <\/del>/.test(html) && /The /.test(html), html);
  // Backspace twice at end of "jumps" deletes characters as tracked
  await caretAfter('jumps', 5);
  await p.keyboard.press('Backspace');
  await p.keyboard.press('Backspace');
  html = await p.$eval('#doc', d => d.innerHTML.replace(/&nbsp;/g, ' '));
  const delText = await p.$$eval('#doc del.tc', ds => ds.map(d => d.textContent).join('|'));
  check('backspace marks characters deleted (no nesting)', delText.replace(/\u00a0/g, ' ') === 'quick |ps' && (await p.$$eval('#doc del.tc del.tc', x => x.length)) === 0 && /jumps/.test(await p.$eval('#doc', d => d.textContent)), delText);
  // Deleting inside your own insertion really removes it
  await caretAfter(' red', 4);
  await p.keyboard.press('Backspace');
  check('deleting your own insertion removes it', (await p.$eval('#doc ins.tc', i => i.textContent.replace(/\u00a0/g, ' '))) === ' re');
  // Markup views
  await p.selectOption('#markupMode', 'final');
  const finalText = await p.evaluate(() => document.getElementById('doc').innerText.split('\n')[0].replace(/\u00a0/g, ' '));
  check('No Markup shows the result', finalText === 'The brown re fox jum.', JSON.stringify(finalText));
  await p.selectOption('#markupMode', 'original');
  const origText = await p.evaluate(() => document.getElementById('doc').innerText.split('\n')[0].replace(/\u00a0/g, ' '));
  check('Original shows the text before changes', origText === 'The quick brown fox jumps.', JSON.stringify(origText));
  await p.selectOption('#markupMode', 'all');
  // Next change navigation
  await caretAfter('The', 0);
  await p.click('#nextChange');
  const sel1 = await p.evaluate(() => getSelection().toString().replace(/\u00a0/g, ' '));
  await p.click('#nextChange');
  const sel2 = await p.evaluate(() => getSelection().toString().replace(/\u00a0/g, ' '));
  check('Next goes through changes in order', sel1 === 'quick ' && sel2 === ' re', sel1 + ' / ' + sel2);
  // Comment
  await p.evaluate(() => { const w = document.createTreeWalker(document.getElementById('doc'), NodeFilter.SHOW_TEXT); while (w.nextNode()) { const i = w.currentNode.nodeValue.indexOf('Second'); if (i >= 0) { const r = document.createRange(); r.setStart(w.currentNode, i); r.setEnd(w.currentNode, i + 6); getSelection().removeAllRanges(); getSelection().addRange(r); break; } } UI.form = async () => ({ t: 'Is this the right word?' }); });
  await p.click('#addComment');
  await p.waitForTimeout(200);
  const pane = await p.evaluate(() => ({ hidden: document.getElementById('commentsPane').hidden, cards: [...document.querySelectorAll('#commentsPane .card')].map(c => c.querySelector('b').textContent + ': ' + c.querySelector('.text').textContent) }));
  check('comment appears in the comments pane', !pane.hidden && pane.cards[0] === 'Ada Lovelace: Is this the right word?', JSON.stringify(pane));
  // DOCX export
  const bytes = await p.evaluate(async () => Array.from(await AtlasWrite.serialize('docx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const docXml = await zip.file('word/document.xml').async('string');
  const commentsXml = zip.file('word/comments.xml') ? await zip.file('word/comments.xml').async('string') : '';
  const settingsXml = zip.file('word/settings.xml') ? await zip.file('word/settings.xml').async('string') : '';
  check('docx: insertion is a Word revision', /<w:ins [^>]*w:author="Ada Lovelace"[^>]*>[\s\S]*? re<\/w:t>/.test(docXml), (docXml.match(/<w:ins [\s\S]{0,200}/) || [''])[0]);
  check('docx: deletion is a Word revision with deleted text', /<w:del [^>]*>[\s\S]*?<w:delText[^>]*>quick <\/w:delText>/.test(docXml));
  check('docx: comment range and reference', /<w:commentRangeStart w:id="0"\/>/.test(docXml) && /<w:commentRangeEnd w:id="0"\/>/.test(docXml) && /<w:commentReference w:id="0"\/>/.test(docXml));
  check('docx: comment text and author', /w:author="Ada Lovelace"/.test(commentsXml) && /Is this the right word\?/.test(commentsXml));
  check('docx: Track Changes stays on in Word', /<w:trackRevisions\/>/.test(settingsXml));
  check('docx: no private markers left', !/[-]/.test(docXml));
  // DOCX import round trip
  await p.evaluate(async b => { await AtlasWrite.loadFile({ name: 'rev.docx', path: null, data: new Uint8Array(b) }); }, bytes);
  const back = await p.evaluate(() => ({
    ins: [...document.querySelectorAll('#doc ins.tc')].map(e => e.textContent.replace(/\u00a0/g, ' ') + '@' + e.dataset.author),
    del: [...document.querySelectorAll('#doc del.tc')].map(e => e.textContent.replace(/\u00a0/g, ' ')),
    cmt: [...document.querySelectorAll('#doc span.comment')].map(e => e.textContent + '=' + e.title + '@' + e.dataset.author),
    tracking: document.getElementById('trackChanges').classList.contains('on')
  }));
  check('docx import restores insertions with author', back.ins.join() === ' re@Ada Lovelace', JSON.stringify(back.ins));
  check('docx import restores deletions', back.del.join('|').includes('quick ') && back.del.length >= 2, JSON.stringify(back.del));
  check('docx import restores the comment', back.cmt[0] === 'Second=Is this the right word?@Ada Lovelace', JSON.stringify(back.cmt));
  check('docx import restores Track Changes on', back.tracking);
  // Accept all / reject all
  await p.click('[data-tab="review"]');
  await p.click('#acceptAll');
  const accepted = await p.evaluate(() => document.getElementById('doc').innerText.split('\n')[0].replace(/\u00a0/g, ' '));
  check('Accept All applies every change', accepted === 'The brown re fox jum.', JSON.stringify(accepted));
  await p.evaluate(async b => { await AtlasWrite.loadFile({ name: 'rev.docx', path: null, data: new Uint8Array(b) }); }, bytes);
  await p.click('#rejectAll');
  const rejected = await p.evaluate(() => document.getElementById('doc').innerText.split('\n')[0].replace(/\u00a0/g, ' '));
  check('Reject All restores the original', rejected === 'The quick brown fox jumps.', JSON.stringify(rejected));
  // HTML round trip keeps markup
  await p.evaluate(async b => { await AtlasWrite.loadFile({ name: 'rev.docx', path: null, data: new Uint8Array(b) }); }, bytes);
  const h = await p.evaluate(() => AtlasWrite.serialize('html'));
  await p.evaluate(async h => { await AtlasWrite.loadFile({ name: 'x.html', path: null, data: new TextEncoder().encode(h) }); }, h);
  const hb = await p.evaluate(() => ({ ins: document.querySelectorAll('#doc ins.tc[data-author]').length, del: document.querySelectorAll('#doc del.tc[datetime]').length, c: document.querySelectorAll('#doc span.comment[data-author]').length }));
  check('html round trip keeps changes and comments', hb.ins === 1 && hb.del >= 2 && hb.c === 1, JSON.stringify(hb));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e); process.exit(1); });
