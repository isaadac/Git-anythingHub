// Desktop: .docx, .xlsx and .pptx files saved by the app carry a thumbnail of their first page,
// cells or slide (docProps/thumbnail.jpeg), which Finder, Quick Look and File Explorer show.
const { _electron: electron } = require('playwright');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${process.env.E2E_WORK}/ud-thumbs`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow();
  await sleep(1000);
  const open = async page => {
    await home.evaluate(p => window.atlas.newWindow(p), page);
    let w; for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes(page + '.html')); }
    await w.waitForLoadState();
    await sleep(800);
    return w;
  };
  async function saved(w, file) {
    await app.evaluate(({ dialog }, f) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: f }); }, file);
    await w.evaluate(() => Platform.emit('save'));
    for (let i = 0; i < 40 && !fs.existsSync(file); i++) await sleep(250);
    await sleep(300);
    const zip = await JSZip.loadAsync(fs.readFileSync(file));
    const jpg = zip.file('docProps/thumbnail.jpeg');
    const bytes = jpg ? await jpg.async('nodebuffer') : null;
    if (bytes && process.env.THUMB_OUT) fs.writeFileSync(path.join(process.env.THUMB_OUT, path.basename(file) + '.jpg'), bytes);
    const rels = await zip.file('_rels/.rels').async('string'), ct = await zip.file('[Content_Types].xml').async('string');
    return { jpeg: !!bytes && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes.length > 500, size: bytes ? bytes.length : 0, rel: /metadata\/thumbnail" Target="docProps\/thumbnail\.jpeg"/.test(rels), ct: /Extension="jpeg"/.test(ct) };
  }
  const w = await open('write');
  await w.click('#doc');
  await w.keyboard.type('Quarterly report');
  const d = await saved(w, path.join(process.env.E2E_WORK, 'thumb.docx'));
  check('desktop: a saved .docx carries a thumbnail of its first page', d.jpeg && d.rel && d.ct, JSON.stringify(d));
  const s = await open('sheets');
  await s.evaluate(() => { AtlasSheets.setCell('A1', 'Sales'); AtlasSheets.setCell('A2', '42'); });
  const x = await saved(s, path.join(process.env.E2E_WORK, 'thumb.xlsx'));
  check('desktop: …and a .xlsx, of its cells', x.jpeg && x.rel && x.ct, JSON.stringify(x));
  const sl = await open('slides');
  const p = await saved(sl, path.join(process.env.E2E_WORK, 'thumb.pptx'));
  check('desktop: …and a .pptx, of its first slide', p.jpeg && p.rel && p.ct, JSON.stringify(p));
  // Still opens.
  const back = await sl.evaluate(async b => { await AtlasSlides.loadFile({ name: 'thumb.pptx', data: new Uint8Array(b) }); return AtlasSlides.pres.slides.length; }, [...fs.readFileSync(path.join(process.env.E2E_WORK, 'thumb.pptx'))]);
  check('desktop: files with thumbnails open as before', back >= 1);
  for (const win of app.windows()) await win.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
