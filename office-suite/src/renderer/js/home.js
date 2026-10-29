'use strict';
// Start screen: launch an app, start from a template, or reopen a pinned or recent file (each shown
// as a miniature of its first page, sheet or slide).
(function () {
  const APP_INFO = {
    write: { label: 'W', cls: 't-write', page: 'write.html' },
    sheets: { label: 'S', cls: 't-sheets', page: 'sheets.html' },
    slides: { label: 'P', cls: 't-slides', page: 'slides.html' },
    notes: { label: 'N', cls: 't-notes', page: 'notes.html' },
    pdf: { label: 'PDF', cls: 't-pdf', page: 'pdf.html' }
  };
  const EXT_APP = { docx: 'write', html: 'write', htm: 'write', txt: 'write', md: 'write',
    xlsx: 'sheets', xls: 'sheets', csv: 'sheets', ods: 'sheets', tsv: 'sheets', pptx: 'slides', aslides: 'slides', anote: 'notes', pdf: 'pdf' };

  // Greeting: the time of day, and the person's first name once they have given it (Live, comments).
  const now = new Date(), h = now.getHours();
  let who = '';
  try { who = (localStorage.getItem('atlas-user-name') || '').trim().split(/\s+/)[0].slice(0, 40); } catch { /* none */ }
  const part = h < 5 ? 'Good evening' : h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening';
  // The phrase in its own element, so it is translated like the rest of the page.
  const greet = document.getElementById('greeting'), phrase = document.createElement('span');
  phrase.textContent = part;
  greet.replaceChildren(phrase, who ? ', ' + who : '');
  document.getElementById('today').textContent = now.toLocaleDateString(document.documentElement.lang || undefined, { weekday: 'long', day: 'numeric', month: 'long' });

  function launch(app) {
    if (Platform.desktop) Platform.newWindow(app);
    else location.href = APP_INFO[app].page;
  }

  document.querySelectorAll('.app-card').forEach(c => c.addEventListener('click', () => launch(c.dataset.app)));

  document.getElementById('open').addEventListener('click', async () => {
    if (Platform.desktop) {
      const f = await Platform.open([{ name: 'Office files', extensions: Object.keys(EXT_APP) }]);
      if (f) Platform.openPath(f.path);
    } else {
      // In the browser, hand the picked file to the right app through sessionStorage.
      const f = await Platform.open([{ name: 'Office files', extensions: Object.keys(EXT_APP) }]);
      if (!f) return;
      const app = EXT_APP[Platform.extOf(f.name)];
      if (!app) return UI.toast('Unsupported file type');
      const dataUrl = await Platform.bytesToDataUrl(f.data, 'application/octet-stream');
      try { sessionStorage.setItem('atlas-handoff', JSON.stringify({ name: f.name, dataUrl })); } catch { return UI.toast('File too large to hand off; open it from inside the app'); }
      location.href = APP_INFO[app].page;
    }
  });

  document.getElementById('clear').addEventListener('click', async () => { await Platform.clearRecent(); render(); });

  function ago(ts) {
    const s = (Date.now() - ts) / 1000;
    if (s < 60) return 'Just now';
    if (s < 3600) return Math.floor(s / 60) + ' min ago';
    if (s < 86400) return Math.floor(s / 3600) + ' h ago';
    return new Date(ts).toLocaleDateString();
  }

  const keyOf = r => r.path || r.name;
  function openEntry(r) {
    if (Platform.isCloudPath(r.path) && window.Cloud) openCloud(r);
    else if (Platform.desktop && r.path) Platform.openPath(r.path);
    else UI.toast('Open it with “Open a file…”');
  }
  /**
   * A recent cloud document: fetched from the cloud, or — offline — the copy kept on this device
   * (with any changes still waiting to go up).
   */
  async function openCloud(r) {
    let data, note = '';
    try {
      if (!navigator.onLine) throw new TypeError('offline');
      const p = Cloud.parsePath(r.path);
      data = await Cloud.download(p.acct, { id: p.id, name: p.name });
      Cloud.noteRev(r.path);
      if (window.Offline) {
        const mine = await Offline.copy(r.path);
        // A save made offline that hasn't gone up yet is newer than the cloud's copy.
        if (mine && mine.waiting) { data = mine.data; note = 'Your changes saved on this device are still waiting to go up.'; } else Offline.keep(r.path, r.name, data);
      }
    } catch (err) {
      if (!window.Offline || !Offline.unreachable(err)) return UI.toast(`Couldn't open “${r.name}”: ${err.message}`, 5000);
      const c = await Offline.copy(r.path);
      if (!c) return UI.toast(`You're offline, and “${r.name}” hasn't been opened on this device before.`, 5000);
      data = c.data;
      note = 'You\'re offline: this is the copy on this device. Changes you save go up when you\'re back online.';
    }
    if (note) { try { sessionStorage.setItem('atlas-open-note', note); } catch { /* the note is optional */ } }
    await Platform.openData(r.app, r.name, data, !Platform.desktop, r.path);
  }
  // icons: pin pin-off
  /** A file as a card: its miniature, name, when it was opened, and a pin. */
  function fileCard(r, onChange) {
    const info = APP_INFO[r.app] || APP_INFO.write;
    const card = document.createElement('div');
    card.className = 'file-card';
    card.tabIndex = 0;
    card.setAttribute('role', 'button');
    card.appendChild(UI.thumbnail(Platform.thumbs.get(keyOf(r)), r.app));
    const row = document.createElement('div');
    row.className = 'info';
    row.innerHTML = `<div class="badge ${info.cls}">${info.label}</div><div class="meta"><div class="n"></div><div class="when"></div></div>`;
    row.querySelector('.n').textContent = r.name;
    row.querySelector('.n').title = r.path || r.name;
    row.querySelector('.when').textContent = r.opened ? ago(r.opened) : (r.path || '');
    card.appendChild(row);
    card.setAttribute('aria-label', r.name);
    const pinned = Platform.pins.has(keyOf(r));
    const pin = document.createElement('button');
    pin.className = 'pin' + (pinned ? ' on' : '');
    pin.title = pinned ? 'Unpin' : 'Pin to the top';
    pin.setAttribute('aria-label', pin.title + ': ' + r.name);
    pin.setAttribute('aria-pressed', String(pinned));
    pin.innerHTML = `<span class="ico" data-icon="${pinned ? 'pin-off' : 'pin'}"></span>`;
    pin.addEventListener('click', e => { e.stopPropagation(); Platform.pins.toggle({ key: keyOf(r), name: r.name, app: r.app, path: r.path }); onChange(); });
    card.appendChild(pin);
    card.addEventListener('click', () => openEntry(r));
    card.addEventListener('keydown', e => { if ((e.key === 'Enter' || e.key === ' ') && e.target === card) { e.preventDefault(); openEntry(r); } });
    return card;
  }

  const findBox = document.getElementById('findRecent');
  async function render() {
    const list = await Platform.recent();
    const byKey = new Map(list.map(r => [keyOf(r), r]));
    // Pinned files: the latest details if the file is still in the recent list.
    const pins = Platform.pins.list().map(p => byKey.get(p.key) || { name: p.name, app: p.app, path: p.path });
    const pinBox = document.getElementById('pinned');
    document.getElementById('pinnedBox').hidden = !pins.length;
    pinBox.innerHTML = '';
    pins.forEach(r => pinBox.appendChild(fileCard(r, render)));
    const q = findBox.value.trim().toLowerCase();
    const shown = list.filter(r => !Platform.pins.has(keyOf(r)) && (!q || (r.name + ' ' + (r.path || '')).toLowerCase().includes(q)));
    const box = document.getElementById('recent');
    box.innerHTML = '';
    if (!shown.length) {
      const e = document.createElement('div');
      e.className = 'empty';
      e.textContent = q ? 'No recent files match.' : list.length ? 'Every recent file is pinned.' : 'Files you open will show up here.';
      box.appendChild(e);
      return;
    }
    shown.forEach(r => box.appendChild(fileCard(r, render)));
  }
  findBox.addEventListener('input', render);
  async function renderRecovered() {
    const list = await Platform.drafts.list();
    const box = document.getElementById('recovered');
    document.getElementById('recoveredBox').hidden = !list.length;
    box.innerHTML = '';
    list.forEach(d => {
      const info = APP_INFO[d.app] || APP_INFO.write;
      const row = document.createElement('div');
      row.className = 'row';
      row.innerHTML = `<div class="badge ${info.cls}">${info.label}</div><div class="meta"><div class="n"></div><div class="p"></div></div><div class="acts"><button class="primary">Open</button><button>Discard</button></div>`;
      row.querySelector('.n').textContent = d.name;
      row.querySelector('.p').textContent = 'Last changed ' + ago(d.savedAt);
      const [open, discard] = row.querySelectorAll('button');
      open.addEventListener('click', async () => { await Platform.drafts.open(d.id, d.app); if (Platform.desktop) renderRecovered(); });
      discard.addEventListener('click', async () => {
        if (!(await UI.confirm('Discard recovered document', `Throw away the unsaved changes to "${d.name}"?`))) return;
        await Platform.drafts.discard(d.id);
        renderRecovered();
      });
      box.appendChild(row);
    });
  }

  // Templates for every app, each shown as a miniature of its first page, sheet or slide.
  const tplBox = document.getElementById('templates');
  const writeCard = t => WriteTemplates.card(t, () => WriteTemplates.open(t, !Platform.desktop));
  const suiteCard = app => t => SuiteTemplates.card(app, t, () => SuiteTemplates.open(app, t, !Platform.desktop));
  const docs = WriteTemplates.list.filter(t => t.id !== 'blank');
  const SHOW = {
    all: () => [...docs.filter(t => ['resume', 'cover-letter'].includes(t.id)).map(writeCard), ...SuiteTemplates.sheets.slice(0, 2).map(suiteCard('sheets')), suiteCard('slides')(SuiteTemplates.slides[0])],
    write: () => docs.slice(0, 5).map(writeCard),
    sheets: () => SuiteTemplates.sheets.map(suiteCard('sheets')),
    slides: () => SuiteTemplates.slides.map(suiteCard('slides'))
  };
  const chips = [...document.querySelectorAll('#tplFilter button')];
  function showTemplates(f) {
    chips.forEach(c => c.setAttribute('aria-checked', String(c.dataset.f === f)));
    tplBox.replaceChildren(...SHOW[f]());
  }
  chips.forEach(c => c.addEventListener('click', () => showTemplates(c.dataset.f)));
  document.getElementById('tplFilter').addEventListener('keydown', e => {
    const i = chips.findIndex(c => c.getAttribute('aria-checked') === 'true');
    const d = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
    if (!d) return;
    e.preventDefault();
    const next = chips[(i + d + chips.length) % chips.length];
    showTemplates(next.dataset.f); next.focus();
  });
  showTemplates('all');
  document.getElementById('moreTemplates').addEventListener('click', async () => {
    const t = await WriteTemplates.pick('Write templates');
    if (t) WriteTemplates.open(t, !Platform.desktop);
  });

  // Cloud accounts
  async function renderCloud() {
    const list = await Cloud.accounts();
    const box = document.getElementById('cloud');
    if (!list.length) return;
    box.innerHTML = '';
    list.forEach(a => {
      const p = Cloud.PROVIDERS[a.provider];
      const row = document.createElement('div');
      row.className = 'row';
      row.innerHTML = '<div class="badge"></div><div class="meta"><div class="n"></div><div class="p"></div></div><div class="acts"><button class="primary">Browse</button><button>Disconnect</button></div>';
      row.querySelector('.badge').textContent = p.badge;
      row.querySelector('.badge').style.background = p.color;
      row.querySelector('.n').textContent = p.label;
      row.querySelector('.p').textContent = a.label;
      const [browse, remove] = row.querySelectorAll('button');
      browse.addEventListener('click', () => Cloud.openDialog({}));
      remove.addEventListener('click', async () => {
        if (!(await UI.confirm('Disconnect', `Disconnect ${p.label} (${a.label})? Your files stay in the cloud.`))) return;
        await Cloud.disconnect(a.id);
        location.reload();
      });
      box.appendChild(row);
    });
  }
  document.getElementById('cloudAdd').addEventListener('click', async () => { if (await Cloud.addAccountDialog()) renderCloud(); });
  document.getElementById('cloudSettings').addEventListener('click', () => Cloud.settingsDialog());

  document.getElementById('themeBtn').addEventListener('click', () => Theme.choose());
  document.getElementById('langBtn').addEventListener('click', () => I18n.choose());
  document.getElementById('helpBtn').addEventListener('click', () => Platform.help());

  render();
  renderRecovered();
  renderCloud();
  window.addEventListener('focus', () => { render(); renderRecovered(); });
})();
