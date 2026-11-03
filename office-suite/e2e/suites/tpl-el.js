const { _electron: electron } = require('playwright');
const path = require('path');
(async () => {
  const app = await electron.launch({ executablePath: path.join(process.cwd(), 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${process.env.E2E_WORK}/ud-tpl`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  const home = await app.firstWindow(); await new Promise(r => setTimeout(r, 1500));
  await home.click('#templates .tpl >> text=Cover letter');
  await new Promise(r => setTimeout(r, 2500));
  const w = app.windows().find(x => x.url().includes('write.html'));
  const info = w && await w.evaluate(() => ({ name: document.querySelector('.titlebar .docname input').value, first: document.querySelector('#doc p').textContent }));
  console.log((info && info.name === 'Cover letter.docx' && /\[Your Name\]/.test(info.first) ? 'PASS' : 'FAIL') + ' desktop: template opens in a new Write window :: ' + JSON.stringify(info));
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
