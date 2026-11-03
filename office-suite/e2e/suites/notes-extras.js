// Notes: page templates, file attachments, audio notes, and sync between two devices through a
// cloud folder (here a WebDAV server), including edits on both sides and deletions.
const { chromium } = require('playwright');
const path = require('path');
const start = require('../cloud-mock');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const ROOT = process.cwd();

(async () => {
  const { port, S, server } = await start(ROOT, {});
  const URL = `http://127.0.0.1:${port}/src/renderer/notes.html`;
  const browser = await chromium.launch({ args: ['--use-fake-ui-for-media-stream', '--use-fake-device-for-media-stream'] });
  const errors = [];
  async function device(name) {
    const ctx = await browser.newContext({ permissions: ['microphone'] });
    const p = await ctx.newPage();
    p.on('pageerror', e => errors.push(name + ': ' + e.message));
    await p.goto(URL);
    await p.waitForFunction(() => window.AtlasNotes && AtlasNotes.page);
    await p.evaluate(async port => { await Cloud.connect('webdav', { url: `http://127.0.0.1:${port}/remote.php/dav/files/u/`, username: 'u', password: 'secret' }); }, port);
    return p;
  }
  const nbOf = p => p.evaluate(() => JSON.parse(JSON.stringify(AtlasNotes.notebook)));
  const pageTitles = nb => nb.sections.flatMap(s => s.pages.map(x => x.title));
  const A = await device('A');

  // ---- Template
  await A.evaluate(() => document.getElementById('pageTemplate').click());
  await A.waitForSelector('.modal');
  await A.selectOption('.modal select', 'meeting');
  await A.click('.modal .actions .primary');
  await A.waitForSelector('.modal', { state: 'detached' });
  await A.fill('#pageTitle', 'Team meeting');
  await A.evaluate(() => document.getElementById('pageTitle').dispatchEvent(new Event('input')));
  check('a page from the Meeting notes template', await A.evaluate(() => /Agenda/.test(document.getElementById('body').innerHTML) && !!document.querySelector('#body .todo')));

  // ---- Attachment: kept in the notebook, saved back exactly
  await A.evaluate(() => {
    window.__file = new Uint8Array([37, 80, 68, 70, 45, 1, 2, 3, 250, 251]);
    Platform.open = async () => ({ name: 'report.pdf', data: window.__file });
    Platform.save = async o => { window.__saved = o; return { path: 'x' }; };
  });
  await A.click('#body');
  await A.keyboard.press('Control+End');
  await A.evaluate(() => document.getElementById('insFile').click());
  await A.waitForSelector('#body .attachment');
  let nb = await nbOf(A);
  const fid = Object.keys(nb.files)[0];
  check('file attached: a chip on the page and the file in the notebook', fid && nb.files[fid].name === 'report.pdf' && nb.files[fid].type === 'application/pdf' && nb.files[fid].size === 10);
  await A.click('#body .attachment');
  await A.waitForFunction(() => window.__saved);
  check('clicking the chip saves an identical copy', await A.evaluate(() => window.__saved.name === 'report.pdf' && window.__saved.data.join() === window.__file.join()));

  // ---- Audio note
  await A.evaluate(() => document.getElementById('recAudio').click());
  await sleep(1300);
  await A.evaluate(() => document.getElementById('recAudio').click());
  await A.waitForSelector('#body .audio-note', { timeout: 8000 });
  nb = await nbOf(A);
  const audio = Object.values(nb.files).find(f => /^audio\//.test(f.type));
  check('audio recorded into the page', audio && audio.size > 500 && audio.duration >= 1, audio && JSON.stringify({ type: audio.type, size: audio.size, d: audio.duration }));
  await A.click('#body .audio-note');
  await sleep(300);
  check('clicking it plays it', await A.evaluate(() => document.querySelector('#body .audio-note').classList.contains('playing')));
  await A.click('#body .audio-note');
  await A.evaluate(() => AtlasNotes.flush());

  // ---- Sync: A keeps the notebook in the cloud
  await A.evaluate(() => document.getElementById('syncBtn').click());
  await A.click('.modal .actions button:has-text("Keep this notebook in the cloud")');
  await A.waitForSelector('.modal input');
  await A.click('.modal .actions .primary');
  await A.waitForFunction(() => /Synced|Sync on/.test(document.getElementById('syncState').textContent), null, { timeout: 10000 });
  const key = Object.keys(S.dav).find(k => k.endsWith('.anote'));
  check('the notebook is a file in the cloud folder', !!key && JSON.parse(S.dav[key].toString()).sections[0].pages.some(x => x.title === 'Team meeting'), key);

  // ---- B opens it from the cloud
  const B = await device('B');
  await B.evaluate(() => document.getElementById('syncBtn').click());
  await B.click('.modal .actions button:has-text("Open a notebook from the cloud")');
  await B.waitForSelector('.cloud-row');
  await B.dblclick('.cloud-row:has-text(".anote")');
  await B.waitForFunction(() => AtlasNotes.notebook.sections[0].pages.some(x => x.title === 'Team meeting'), null, { timeout: 10000 });
  nb = await nbOf(B);
  check('B has the notebook with its pages and attachments', pageTitles(nb).includes('Team meeting') && Object.keys(nb.files).length === 2, JSON.stringify(Object.values(nb.files).map(f => [f.name, f.type, f.size])) + ' ' + pageTitles(nb).join());

  // Both edit: B adds a page and changes the meeting page; A changes the first page.
  await B.evaluate(() => { const s = AtlasNotes.notebook.sections[0]; const m = s.pages.find(x => x.title === 'Team meeting'); m.html += '<p>Added on B</p>'; m.modified = Date.now(); s.pages.push({ id: 'pFromB1', title: 'From B', html: '<p>hello from B</p>', ink: [], level: 0, created: Date.now(), modified: Date.now() }); });
  await A.evaluate(() => { const first = AtlasNotes.notebook.sections[0].pages[0]; first.title = 'Edited on A'; first.modified = Date.now(); });
  await B.evaluate(() => AtlasNotes.syncNow());
  await A.evaluate(() => AtlasNotes.syncNow());
  await B.evaluate(() => AtlasNotes.syncNow());
  const na = await nbOf(A), nbB = await nbOf(B);
  check('A gets B\'s new page and B\'s edit, keeping its own', pageTitles(na).includes('From B') && pageTitles(na).includes('Edited on A') && na.sections[0].pages.find(x => x.title === 'Team meeting').html.includes('Added on B'), pageTitles(na).join(', '));
  check('B gets A\'s edit', pageTitles(nbB).includes('Edited on A'), pageTitles(nbB).join(', '));
  // A deletes B's page; the deletion reaches B.
  await A.evaluate(() => { const nbk = AtlasNotes.notebook; const s = nbk.sections[0]; const i = s.pages.findIndex(x => x.id === 'pFromB1'); s.pages.splice(i, 1); nbk.deleted.pFromB1 = Date.now(); });
  await A.evaluate(() => AtlasNotes.syncNow());
  await B.evaluate(() => AtlasNotes.syncNow());
  check('a page deleted on A disappears on B', !pageTitles(await nbOf(B)).includes('From B'));

  check('no page errors', !errors.length, errors.join(' | '));
  await browser.close();
  server.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
