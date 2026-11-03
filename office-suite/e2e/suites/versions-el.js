// Desktop: versions kept by the main process on each save, and the native menus in the chosen language.
const { _electron: electron } = require('playwright');
const fs = require('fs');
const path = require('path');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-versions-' + Date.now());
  const file = path.join(process.env.E2E_WORK, 'notes.txt');
  const launch = () => electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '' } });
  let app = await launch();
  const home = await app.firstWindow(); await sleep(1200);
  await home.evaluate(() => window.atlas.newWindow('write'));
  let w; for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes('write.html')); }
  await w.waitForFunction(() => window.AtlasWrite && window.Platform);

  // ---- Versions (the file is chosen once with Save As, as a person would)
  await app.evaluate(({ dialog }, f) => { dialog.showSaveDialog = async () => ({ canceled: false, filePath: f }); }, file);
  await w.evaluate(() => Platform.chooseSavePath('notes.txt', [{ name: 'Text', extensions: ['txt'] }]));
  const save = text => w.evaluate(([f, t]) => Platform.save({ data: new TextEncoder().encode(t), name: 'notes.txt', path: f }), [file, text]);
  await save('one');
  await save('one');
  await save('two');
  await sleep(300);
  let list = await w.evaluate(f => Platform.versions.list(f), file);
  check('saves write the file and keep a version (saves close together share one)', fs.readFileSync(file, 'utf8') === 'two' && list.length === 1, list.length);
  // An older version, as if saved an hour ago.
  const dir = fs.readdirSync(path.join(ud, 'versions')).map(d => path.join(ud, 'versions', d))[0];
  const idx = JSON.parse(fs.readFileSync(path.join(dir, 'index.json'), 'utf8'));
  idx.versions[0].created -= 3600e3;
  fs.writeFileSync(path.join(dir, 'index.json'), JSON.stringify(idx));
  await save('three');
  list = await w.evaluate(f => Platform.versions.list(f), file);
  const texts = await w.evaluate(async ([f, l]) => { const out = []; for (const v of l) out.push(new TextDecoder().decode(await Platform.versions.read(f, v.id))); return out; }, [file, list]);
  check('a later save adds one; each reads back, newest first', texts.join() === 'three,two', texts.join());
  check('bad requests are refused', await w.evaluate(async f => (await window.atlas.versionsRead(f, '../../x')) === null && (await window.atlas.versionsList('')).length === 0, file));

  // ---- Menus in the chosen language
  const menu = () => app.evaluate(({ Menu }) => { const m = Menu.getApplicationMenu(); return m.items.map(i => i.label).concat(m.items.find(i => i.submenu && i.submenu.items.some(s => s.role === 'cut'))?.submenu.items.filter(s => s.role === 'cut').map(s => s.label) || []); });
  check('English menus to begin with', (await menu()).includes('File'));
  await w.evaluate(() => I18n.set('fr'));
  await sleep(300);
  let m = await menu();
  check('choosing French translates the native menus', m.includes('Fichier') && m.includes('Édition') && m.includes('Couper'), m.join(', '));
  check('…and the other windows', await home.evaluate(() => document.documentElement.lang === 'fr' && document.body.textContent.includes('Récent')));
  await require('../close-app')(app);
  app = await launch();
  const h2 = await app.firstWindow(); await sleep(1200);
  m = await menu();
  check('remembered after a restart (menus and pages)', m.includes('Fichier') && await h2.evaluate(() => document.documentElement.lang === 'fr'), m.join(', '));
  await h2.evaluate(() => I18n.set('en'));
  await sleep(300);
  check('back to English', (await menu()).includes('File'));
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
