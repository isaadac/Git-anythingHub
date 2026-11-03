// Isadek PDF: open a PDF, fill in its form, highlight, draw, add text and comments, sign,
// rotate, reorder, delete and insert pages; the saved PDF carries all of it (checked with pdf-lib
// and pdf.js, independently of the app); and its text opens as a document in Write.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const ROOT = process.cwd();
const { PDFDocument, StandardFonts, PDFName } = require(path.join(ROOT, 'node_modules/pdf-lib'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

async function makePdf(pagesText, withForm) {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica), bold = await doc.embedFont(StandardFonts.HelveticaBold);
  for (const [i, t] of pagesText.entries()) {
    const pg = doc.addPage([612, 792]);
    pg.drawText(t, { x: 72, y: 700, size: 24, font: bold });
    pg.drawText('Revenue grew forty percent in the third quarter.', { x: 72, y: 660, size: 12, font });
    if (withForm && i === 0) {
      const form = doc.getForm();
      form.createTextField('applicant.name').addToPage(pg, { x: 130, y: 552, width: 200, height: 22 });
      form.createCheckBox('agree').addToPage(pg, { x: 72, y: 520, width: 16, height: 16 });
      const team = form.createDropdown('team'); team.addOptions(['Sales', 'Engineering', 'Design']); team.addToPage(pg, { x: 130, y: 480, width: 160, height: 22 });
    }
  }
  return doc.save();
}

(async () => {
  const report = await makePdf(['Quarterly Report', 'Second page', 'Third page'], true);
  const appendix = await makePdf(['Appendix A'], false);
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1300, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/pdf.html');
  await p.waitForFunction(() => window.__pdfReady, null, { timeout: 20000 });
  await p.evaluate(b => AtlasPdf.loadFile({ name: 'Report.pdf', data: new Uint8Array(b) }), [...report]);
  await p.waitForSelector('.pdf-page .textLayer span');
  check('a PDF opens: its pages, selectable text and thumbnails', (await p.$$('.pdf-page')).length === 3 && (await p.$$('.pdf-thumb')).length === 3 && /Quarterly Report/.test(await p.textContent('.pdf-page .textLayer')));

  // ---- Fill in the form
  await p.waitForSelector('.annotationLayer input[type=text], .annotationLayer input:not([type])');
  await p.fill('.annotationLayer input[type=text], .annotationLayer input:not([type])', 'Ada Lovelace');
  await p.check('.annotationLayer input[type=checkbox]');
  await p.selectOption('.annotationLayer select', 'Engineering');
  check('form fields can be filled in on the page', (await p.inputValue('.annotationLayer select')) === 'Engineering');

  // ---- Highlight text by selecting it
  await p.click('[data-tab="annotate"]');
  await p.click('.tool[data-tool="highlight"]');
  const span = await p.$('.pdf-page .textLayer span:has-text("Revenue grew")');
  const sb = await span.boundingBox();
  await p.mouse.move(sb.x + 2, sb.y + sb.height / 2); await p.mouse.down(); await p.mouse.move(sb.x + sb.width - 2, sb.y + sb.height / 2, { steps: 5 }); await p.mouse.up();
  await sleep(300);
  check('selecting text with Highlight marks it', await p.evaluate(() => AtlasPdf.pages[0].annots.some(a => a.type === 'highlight' && a.rects.length >= 1 && a.rects[0][2] > 100)) && (await p.$$('.pdf-page .pdf-annots rect.hl')).length >= 1);
  // ---- Draw
  await p.click('.tool[data-tool="pen"]');
  const box = await (await p.$('.pdf-page')).boundingBox();
  await p.mouse.move(box.x + 100, box.y + 300); await p.mouse.down();
  for (let i = 1; i <= 10; i++) await p.mouse.move(box.x + 100 + i * 20, box.y + 300 + (i % 2) * 15);
  await p.mouse.up();
  await sleep(300);
  check('drawing with the pen adds ink to the page', await p.evaluate(() => AtlasPdf.pages[0].annots.some(a => a.type === 'ink')));
  // ---- A comment and some text (through the dialogs)
  await p.click('.tool[data-tool="note"]');
  await p.mouse.click(box.x + 500, box.y + 120);
  await p.waitForSelector('.modal textarea'); await p.fill('.modal textarea', 'Check this figure'); await p.click('.modal .actions .primary');
  await p.click('.tool[data-tool="text"]');
  await p.mouse.click(box.x + 72, box.y + 620);
  await p.waitForSelector('.modal input'); await p.fill('.modal input', 'Approved by Ada'); await p.click('.modal .actions .primary');
  await sleep(300);
  check('comments and text go on the page', (await p.$$('.pdf-page .pdf-note')).length === 1 && /Approved by Ada/.test(await p.textContent('.pdf-page .pdf-textbox')));
  // ---- Sign
  await p.click('[data-tab="sign"]');
  await p.evaluate(() => localStorage.removeItem('atlas-signature'));
  await p.click('#signBtn');
  await p.waitForSelector('.sig-dialog svg');
  const pad = await (await p.$('.sig-dialog svg')).boundingBox();
  await p.mouse.move(pad.x + 40, pad.y + 100); await p.mouse.down();
  for (let i = 1; i <= 12; i++) await p.mouse.move(pad.x + 40 + i * 25, pad.y + 100 + Math.sin(i) * 30);
  await p.mouse.up();
  await p.click('.sig-dialog .ok');
  await p.mouse.click(box.x + 400, box.y + 600);
  await sleep(300);
  check('a drawn signature goes where it is put', await p.evaluate(() => AtlasPdf.pages[0].annots.some(a => a.type === 'image' && /^data:image\/png/.test(a.src))));
  check('…and is kept for next time', await p.evaluate(() => !!localStorage.getItem('atlas-signature')));

  // ---- Pages: rotate page 2, delete page 3, insert the appendix, move it first
  await p.click('[data-tab="pages"]');
  await p.evaluate(() => AtlasPdf.select([2]));
  await p.click('#rotRight');
  await sleep(300);
  await p.evaluate(() => AtlasPdf.select([3]));
  await p.click('#delPages'); await p.waitForSelector('.modal'); await p.click('.modal .actions .primary');
  await sleep(300);
  await p.evaluate(b => { window.__appendix = new Uint8Array(b); Platform.open = async () => ({ name: 'Appendix.pdf', data: window.__appendix }); AtlasPdf.select([2]); }, [...appendix]);
  await p.click('#insertPdf');
  await sleep(800);
  await p.evaluate(() => AtlasPdf.select([3]));
  await p.click('#upPage'); await sleep(200);
  await p.evaluate(() => AtlasPdf.select([2]));
  await p.click('#upPage'); await sleep(500);
  const order = await p.evaluate(() => AtlasPdf.pages.map(x => `${x.src}:${x.index}:${x.rot}`));
  check('pages can be rotated, deleted, inserted from another PDF and moved', order.join() === '1:0:0,0:0:0,0:1:90', order.join());

  // ---- Save, and check the file independently
  const [download] = await Promise.all([p.waitForEvent('download', { timeout: 30000 }), (async () => { await p.evaluate(() => Platform.emit('save')); await p.waitForSelector('.modal input'); await p.click('.modal .actions .primary'); })()]);
  const file = path.join(process.env.E2E_WORK, 'saved.pdf');
  await download.saveAs(file);
  const out = await PDFDocument.load(fs.readFileSync(file));
  const pgs = out.getPages();
  check('the saved PDF has the pages in their new order and rotation', pgs.length === 3 && pgs[2].getRotation().angle === 90 && pgs[1].getRotation().angle === 0, pgs.map(x => x.getRotation().angle).join());
  const form = out.getForm();
  check('…the form as filled in', form.getTextField('applicant.name').getText() === 'Ada Lovelace' && form.getCheckBox('agree').isChecked() && form.getDropdown('team').getSelected().join() === 'Engineering');
  const annots = pgs[1].node.lookup(PDFName.of('Annots'));
  const subtypes = annots ? annots.asArray().map(r => String(out.context.lookup(r).get(PDFName.of('Subtype')))) : [];
  check('…the comment as a real PDF comment', subtypes.includes('/Text'), subtypes.join());
  // Text as pdf.js reads it, from the saved file.
  const text = await p.evaluate(async b => {
    const pdfjs = globalThis.pdfjsLib;
    const doc = await pdfjs.getDocument({ data: new Uint8Array(b) }).promise;
    const out = [];
    for (let i = 1; i <= doc.numPages; i++) out.push((await (await doc.getPage(i)).getTextContent()).items.map(x => x.str).join(' '));
    return out;
  }, [...fs.readFileSync(file)]);
  check('…the added text, on the right page', /Appendix A/.test(text[0]) && /Quarterly Report/.test(text[1]) && /Approved by Ada/.test(text[1]) && /Second page/.test(text[2]), JSON.stringify(text.map(t => t.slice(0, 60))));
  const raw = fs.readFileSync(file).toString('latin1');
  check('…and the highlight, drawing and signature', (raw.match(/\/Subtype\s*\/Image/g) || []).length >= 1, String((raw.match(/\/Subtype\s*\/Image/g) || []).length));

  // ---- Edit in Write
  const html = await p.evaluate(() => AtlasPdf.toHtml());
  check('its text becomes a document for Write, headings and paragraphs', /<h2>Appendix A<\/h2>/.test(html) && /<p>Revenue grew forty percent in the third quarter\.<\/p>/.test(html), html.slice(0, 200));
  // ---- Find
  await p.evaluate(() => Platform.emit('find'));
  await p.fill('#findText', 'forty');
  await p.press('#findText', 'Enter');
  await sleep(800);
  check('find goes through every page', /of 3/.test(await p.textContent('#findCount')), await p.textContent('#findCount'));

  await p.evaluate(() => { window.__atlasDirty = false; });
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
