const { _electron: electron } = require('playwright');
const path = require('path'), fs = require('fs');
(async () => {
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${process.env.E2E_WORK}/ud-pdf`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow(); await new Promise(r => setTimeout(r, 1500));
  await home.evaluate(() => window.atlas.newWindow('write')); await new Promise(r => setTimeout(r, 2000));
  const w = app.windows().find(x => x.url().includes('write.html'));
  const n = await w.evaluate(async () => {
    document.getElementById('doc').innerHTML = Array.from({ length: 90 }, (_, i) => `<p>Paragraph ${i + 1}. ${'Some words to fill the line. '.repeat(6)}</p>`).join('');
    document.getElementById('doc').dispatchEvent(new Event('input'));
    await new Promise(r => setTimeout(r, 300));
    const f = document.getElementById('pageFooter'); f.innerHTML = '<p style="text-align:center">Page <span class="fld" data-fld="page">1</span></p>'; f.dispatchEvent(new Event('input'));
    await new Promise(r => setTimeout(r, 300));
    window.dispatchEvent(new Event('beforeprint'));
    return document.querySelectorAll('.page-bg').length;
  });
  const pdf = await app.evaluate(async ({ BrowserWindow }) => {
    const win = BrowserWindow.getAllWindows().find(b => b.webContents.getURL().includes('write.html'));
    return (await win.webContents.printToPDF({ printBackground: true, pageSize: 'Letter', preferCSSPageSize: true })).toString('latin1');
  });
  fs.writeFileSync(process.env.E2E_WORK + '/write.pdf', Buffer.from(pdf, 'latin1'));
  const pdfPages = (pdf.match(/\/Type\s*\/Page[^s]/g) || []).length;
  console.log((pdfPages === n ? 'PASS' : 'FAIL') + ` PDF has one page per sheet :: sheets=${n} pdfPages=${pdfPages}`);
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
