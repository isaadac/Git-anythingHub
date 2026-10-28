const { _electron: electron } = require('playwright');
const path = require('path'), fs = require('fs');
const SP = process.env.E2E_WORK;
(async () => {
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${SP}/ud-pdftoc`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow(); await new Promise(r => setTimeout(r, 1500));
  await home.evaluate(() => window.atlas.newWindow('write')); await new Promise(r => setTimeout(r, 2000));
  const w = app.windows().find(x => x.url().includes('write.html'));
  const out = SP + '/toc.pdf';
  try { fs.unlinkSync(out); } catch {}
  await app.evaluate(({ dialog, shell }, out) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: out }); shell.openPath = async () => ''; }, out);
  await w.evaluate(async () => {
    const filler = n => Array.from({ length: n }, (_, i) => `<p>Paragraph ${i + 1}. ${'Some words to fill the line. '.repeat(6)}</p>`).join('');
    const d = document.getElementById('doc');
    d.innerHTML = '<p>Start</p><h1>Chapter One</h1>' + filler(30) + '<h2>Section A</h2>' + filler(30) + '<h1>Chapter Two</h1>' + filler(10);
    d.dispatchEvent(new Event('input'));
    const r = document.createRange(); r.selectNodeContents(d.firstChild); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); d.focus();
    await new Promise(r => setTimeout(r, 300));
    document.getElementById('insToc').click();
    await new Promise(r => setTimeout(r, 400));
    document.querySelector('[data-cmd="export-pdf"]').click();
  });
  for (let i = 0; i < 40 && !fs.existsSync(out); i++) await new Promise(r => setTimeout(r, 250));
  await new Promise(r => setTimeout(r, 500));
  const pdf = fs.readFileSync(out).toString('latin1');
  const pages = (pdf.match(/\/Type\s*\/Page[^s]/g) || []).length;
  const links = (pdf.match(/\/Subtype\s*\/Link/g) || []).length;
  const internal = (pdf.match(/\/Dest\s*[\[(\/]/g) || []).length + (pdf.match(/\/S\s*\/GoTo/g) || []).length;
  const outline = /\/Outlines\s+\d+ 0 R/.test(pdf);
  const toc = await w.evaluate(() => [...document.querySelectorAll('#doc .toc .toc-pg')].map(s => s.textContent).join());
  check('PDF exported', pages >= 4, 'pages=' + pages + ' toc pages=' + toc);
  check('PDF contents entries are clickable links inside the document', links >= 3 && internal >= 3, `links=${links} internal=${internal}`);
  check('PDF has bookmarks (document outline)', outline);
  const z = require('zlib');
  // Outline titles may be in compressed object streams; look for them in inflated streams too
  let titles = pdf;
  for (const m of pdf.matchAll(/stream\r?\n([\s\S]*?)endstream/g)) { try { titles += z.inflateSync(Buffer.from(m[1], 'latin1')).toString('latin1'); } catch {} }
  check('bookmarks name the headings', /Chapter One/.test(titles.replace(/\\/g, '')) || /C\x00h\x00a\x00p/.test(titles) || /FEFF0043/.test(titles), '');
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
