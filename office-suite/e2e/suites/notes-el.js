const { _electron: electron } = require('playwright');
const path = require('path'), fs = require('fs');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const ud = path.join(process.env.E2E_WORK, 'ud-notes-' + Date.now());
  const launch = args => electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.', ...args], env: { ...process.env, NODE_OPTIONS: '' } });
  let app = await launch([]);
  const home = await app.firstWindow(); await sleep(1200);
  await home.click('.app-card[data-app="notes"]');
  let n; for (let i = 0; i < 40 && !n; i++) { await sleep(250); n = app.windows().find(w => w.url().includes('notes.html')); }
  await n.waitForFunction(() => window.AtlasNotes && AtlasNotes.page);
  await n.fill('#pageTitle', 'Desktop page');
  await n.click('#body'); await n.keyboard.type('Saved without asking');
  await sleep(1200);
  const files = fs.readdirSync(path.join(ud, 'notebooks')).filter(f => f.endsWith('.json'));
  const saved = JSON.parse(fs.readFileSync(path.join(ud, 'notebooks', files[0]), 'utf8'));
  check('notebook saved to the data folder as you type', files.length === 1 && saved.sections[0].pages[0].title === 'Desktop page' && /Saved without asking/.test(saved.sections[0].pages[0].html));
  check('closing needs no Save prompt', (await n.evaluate(() => window.__atlasDirty)) === false);
  // an .anote file opened from the system
  const anote = { ...saved, id: 'nShared1', name: 'Shared notebook', sections: [{ id: 's1', name: 'Plans', color: '#217346', pages: [{ id: 'p1', title: 'Trip', html: '<p class="todo">Book flights</p>', ink: [], level: 0, created: 1, modified: 1 }] }] };
  const file = path.join(process.env.E2E_WORK, 'shared.anote');
  fs.writeFileSync(file, JSON.stringify(anote));
  await require('../close-app')(app);
  app = await launch([file]);
  await app.firstWindow(); await sleep(1500);
  n = app.windows().find(w => w.url().includes('notes.html'));
  await n.waitForFunction(() => window.AtlasNotes && AtlasNotes.notebook && AtlasNotes.notebook.name === 'Shared notebook', null, { timeout: 8000 });
  check('an .anote file opened from the system becomes a notebook', (await n.inputValue('#pageTitle')) === 'Trip');
  const opts = await n.$$eval('#nbSelect option', os => os.map(o => o.textContent));
  check('both notebooks listed', opts.includes('My Notebook') && opts.includes('Shared notebook'), opts.join(','));
  // delete goes to trash
  n.on('dialog', d => d.accept());
  await n.evaluate(() => Platform.emit('delete-notebook'));
  await n.click('.modal .ok');
  await sleep(800);
  check('deleted notebook goes to the trash folder', fs.readdirSync(path.join(ud, 'notebooks', '.trash')).length === 1 && (await n.evaluate(() => AtlasNotes.notebook.name)) === 'My Notebook');
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
