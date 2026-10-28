'use strict';
// Start screen: launch an app or reopen a recent file.
(function () {
  const APP_INFO = {
    write: { label: 'W', cls: 't-write', page: 'write.html' },
    sheets: { label: 'S', cls: 't-sheets', page: 'sheets.html' },
    slides: { label: 'P', cls: 't-slides', page: 'slides.html' }
  };
  const EXT_APP = { docx: 'write', html: 'write', htm: 'write', txt: 'write', md: 'write',
    xlsx: 'sheets', xls: 'sheets', csv: 'sheets', ods: 'sheets', tsv: 'sheets', pptx: 'slides', aslides: 'slides' };

  const h = new Date().getHours();
  document.getElementById('greeting').textContent = h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening';

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

  async function render() {
    const list = await Platform.recent();
    const box = document.getElementById('recent');
    box.innerHTML = '';
    if (!list.length) { box.innerHTML = '<div class="empty">Files you open will show up here.</div>'; return; }
    list.forEach(r => {
      const info = APP_INFO[r.app] || APP_INFO.write;
      const row = document.createElement('div');
      row.className = 'row';
      row.innerHTML = `<div class="badge ${info.cls}">${info.label}</div><div class="meta"><div class="n"></div><div class="p"></div></div><div class="when"></div>`;
      row.querySelector('.n').textContent = r.name;
      row.querySelector('.p').textContent = r.path || '';
      row.querySelector('.when').textContent = ago(r.opened);
      row.addEventListener('click', () => {
        if (Platform.desktop && r.path) Platform.openPath(r.path);
        else UI.toast('Open it with “Open a file…”');
      });
      box.appendChild(row);
    });
  }
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

  // Write templates
  const tplBox = document.getElementById('templates');
  WriteTemplates.list.filter(t => t.id !== 'blank').slice(0, 6).forEach(t => {
    const b = document.createElement('button');
    b.className = 'tpl';
    b.innerHTML = '<span class="tpl-icon"></span><b></b><small></small>';
    b.querySelector('.tpl-icon').textContent = t.icon;
    b.querySelector('b').textContent = t.name;
    b.querySelector('small').textContent = t.desc;
    b.addEventListener('click', () => WriteTemplates.open(t, !Platform.desktop));
    tplBox.appendChild(b);
  });
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

  render();
  renderRecovered();
  renderCloud();
  window.addEventListener('focus', () => { render(); renderRecovered(); });
})();
