'use strict';
// Shared runtime for all Atlas Office apps.
// Wraps the desktop bridge (window.atlas, from preload.js) and falls back to plain
// browser APIs so the same apps also run as a web page on phones, tablets and ChromeOS.
(function () {
  const desktop = !!(window.atlas && window.atlas.isDesktop);
  const menuHandlers = [];

  function pickFileWeb(filters) {
    return new Promise(resolve => {
      const input = document.createElement('input');
      input.type = 'file';
      const exts = (filters || []).flatMap(f => f.extensions).filter(e => e !== '*');
      if (exts.length) input.accept = exts.map(e => '.' + e).join(',');
      input.onchange = async () => {
        const file = input.files[0];
        if (!file) return resolve(null);
        resolve({ name: file.name, path: null, data: new Uint8Array(await file.arrayBuffer()) });
      };
      input.click();
    });
  }

  function downloadWeb(data, name) {
    const blob = data instanceof Blob ? data : new Blob([data]);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  const Platform = {
    desktop,
    async open(filters) {
      return desktop ? window.atlas.openDialog(filters) : pickFileWeb(filters);
    },
    async readPath(p) { return desktop ? window.atlas.readFile(p) : null; },
    /** Saves bytes or text. Returns {name, path} or null if cancelled. */
    async save({ data, name, path, filters, saveAs }) {
      if (desktop) return window.atlas.save({ data, filePath: path, defaultName: name, filters, saveAs });
      let fileName = name;
      if (saveAs || !fileName) {
        fileName = await UI.prompt('Save As', 'File name', name || 'Untitled');
        if (!fileName) return null;
      }
      downloadWeb(data, fileName);
      return { name: fileName, path: null };
    },
    async pending() {
      if (desktop) return window.atlas.pendingFile();
      // Web: the start page hands a picked file over through sessionStorage.
      let raw = null;
      try { raw = sessionStorage.getItem('atlas-handoff'); sessionStorage.removeItem('atlas-handoff'); } catch { /* ignore */ }
      if (!raw) return null;
      const { name, dataUrl } = JSON.parse(raw);
      const bytes = new Uint8Array(await (await fetch(dataUrl)).arrayBuffer());
      return { name, path: null, data: bytes };
    },
    /** Desktop: asks where to save without writing anything yet. Web: asks for a file name. */
    async chooseSavePath(name, filters) {
      if (desktop) return window.atlas.chooseSavePath({ defaultName: name, filters });
      const fileName = await UI.prompt('Save As', 'File name', name);
      return fileName ? { name: fileName, path: null } : null;
    },
    /** Opens a picked file in place when the current doc is empty, else in a new window (desktop). */
    shouldOpenInPlace(isEmpty) { return !desktop || isEmpty; },
    async exportPdf(name, landscape) {
      if (desktop) return window.atlas.exportPdf({ defaultName: name, landscape });
      UI.toast('Choose "Save as PDF" as the printer');
      window.print();
      return null;
    },
    async print() { return desktop ? window.atlas.print() : window.print(); },
    async recent() { return desktop ? window.atlas.recent() : JSON.parse(localStorageGet('atlas-recent') || '[]'); },
    async clearRecent() { return desktop ? window.atlas.clearRecent() : localStorageSet('atlas-recent', '[]'); },
    newWindow(app) {
      if (desktop) return window.atlas.newWindow(app);
      window.open(app === 'home' ? 'index.html' : app + '.html', '_blank');
    },
    openPath(p) { return desktop && window.atlas.openPath(p); },
    setTitle(t) { document.title = t; if (desktop) window.atlas.setTitle(t); },
    closeNow() { if (desktop) window.atlas.closeNow(); },
    fullscreen(on) {
      if (desktop) return window.atlas.fullscreen(on);
      if (on) document.documentElement.requestFullscreen?.(); else if (document.fullscreenElement) document.exitFullscreen?.();
    },
    onMenu(cb) { menuHandlers.push(cb); },
    emit(cmd) { menuHandlers.forEach(h => h(cmd)); }
  };

  function localStorageGet(k) { try { return localStorage.getItem(k); } catch { return null; } }
  function localStorageSet(k, v) { try { localStorage.setItem(k, v); } catch { /* ignore */ } }

  if (desktop) {
    window.atlas.onMenu(cmd => Platform.emit(cmd));
  } else {
    // In the browser there is no native menu, so map the usual shortcuts ourselves.
    document.addEventListener('keydown', e => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      const k = e.key.toLowerCase();
      const map = { s: e.shiftKey ? 'save-as' : 'save', o: 'open', p: 'print', f: 'find' };
      if (map[k]) { e.preventDefault(); Platform.emit(map[k]); }
    });
    window.addEventListener('beforeunload', e => { if (window.__atlasDirty) { e.preventDefault(); e.returnValue = ''; } });
  }

  // ---------- UI helpers ----------
  const UI = {
    toast(msg, ms = 2200) {
      let el = document.querySelector('.toast');
      if (!el) { el = document.createElement('div'); el.className = 'toast'; document.body.appendChild(el); }
      el.textContent = msg;
      el.classList.add('show');
      clearTimeout(el._t);
      el._t = setTimeout(() => el.classList.remove('show'), ms);
    },
    /** Generic modal form. fields: [{name,label,type,value,options}] -> resolves to values object or null. */
    form(title, fields, okLabel = 'OK') {
      return new Promise(resolve => {
        const back = document.createElement('div');
        back.className = 'modal-backdrop';
        const modal = document.createElement('div');
        modal.className = 'modal';
        modal.innerHTML = `<h2></h2><div class="fields"></div><div class="actions"><button class="cancel">Cancel</button><button class="primary ok"></button></div>`;
        modal.querySelector('h2').textContent = title;
        modal.querySelector('.ok').textContent = okLabel;
        const box = modal.querySelector('.fields');
        const inputs = {};
        fields.forEach(f => {
          const label = document.createElement('label');
          label.textContent = f.label;
          let input;
          if (f.type === 'select') {
            input = document.createElement('select');
            f.options.forEach(o => { const opt = document.createElement('option'); opt.value = o.value ?? o; opt.textContent = o.label ?? o; input.appendChild(opt); });
          } else if (f.type === 'textarea') {
            input = document.createElement('textarea');
            input.rows = 4;
          } else {
            input = document.createElement('input');
            input.type = f.type || 'text';
            if (f.min !== undefined) input.min = f.min;
            if (f.max !== undefined) input.max = f.max;
          }
          if (f.value !== undefined) input.value = f.value;
          label.appendChild(input);
          box.appendChild(label);
          inputs[f.name] = input;
        });
        back.appendChild(modal);
        document.body.appendChild(back);
        const first = Object.values(inputs)[0];
        if (first) { first.focus(); first.select?.(); }
        const done = ok => {
          back.remove();
          if (!ok) return resolve(null);
          const out = {};
          for (const [k, el] of Object.entries(inputs)) out[k] = el.value;
          resolve(out);
        };
        modal.querySelector('.ok').onclick = () => done(true);
        modal.querySelector('.cancel').onclick = () => done(false);
        back.addEventListener('mousedown', e => { if (e.target === back) done(false); });
        modal.addEventListener('keydown', e => {
          if (e.key === 'Escape') done(false);
          if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA') { e.preventDefault(); done(true); }
        });
      });
    },
    async prompt(title, label, value = '') {
      const r = await UI.form(title, [{ name: 'v', label, value }]);
      return r ? r.v : null;
    },
    confirm(title, message) {
      return new Promise(resolve => {
        const back = document.createElement('div');
        back.className = 'modal-backdrop';
        back.innerHTML = `<div class="modal"><h2></h2><p></p><div class="actions"><button class="cancel">Cancel</button><button class="primary ok">OK</button></div></div>`;
        back.querySelector('h2').textContent = title;
        back.querySelector('p').textContent = message;
        document.body.appendChild(back);
        back.querySelector('.ok').onclick = () => { back.remove(); resolve(true); };
        back.querySelector('.cancel').onclick = () => { back.remove(); resolve(false); };
        back.querySelector('.ok').focus();
      });
    },
    /** Wires ribbon tab buttons to their panels and the File tab to the backstage view. */
    initRibbon() {
      const tabs = document.querySelectorAll('.ribbon-tabs button[data-tab]');
      tabs.forEach(btn => btn.addEventListener('click', () => {
        tabs.forEach(b => b.classList.toggle('active', b === btn));
        document.querySelectorAll('.ribbon-panel').forEach(p => p.classList.toggle('active', p.dataset.panel === btn.dataset.tab));
      }));
      const backstage = document.querySelector('.backstage');
      const fileTab = document.querySelector('.ribbon-tabs .file-tab');
      if (backstage && fileTab) {
        fileTab.addEventListener('click', () => backstage.classList.add('open'));
        backstage.querySelector('.back').addEventListener('click', () => backstage.classList.remove('open'));
        backstage.querySelectorAll('nav button[data-cmd]').forEach(b => b.addEventListener('click', () => {
          backstage.classList.remove('open');
          Platform.emit(b.dataset.cmd);
        }));
        document.addEventListener('keydown', e => { if (e.key === 'Escape') backstage.classList.remove('open'); });
      }
      // Buttons with data-cmd anywhere in the chrome emit commands.
      document.querySelectorAll('.titlebar [data-cmd], .ribbon [data-cmd]').forEach(b =>
        b.addEventListener('click', () => Platform.emit(b.dataset.cmd)));
      // Keep focus in the editor when pressing ribbon buttons.
      document.querySelectorAll('.ribbon .btn, .titlebar .qa button').forEach(b => b.addEventListener('mousedown', e => e.preventDefault()));
    },
    /** Tracks the document name, path and unsaved state, and keeps the window title in sync. */
    docState(appLabel, defaultName) {
      const state = {
        name: defaultName, path: null, dirty: false,
        set(props) { Object.assign(state, props); state.render(); },
        markDirty() { if (!state.dirty) { state.dirty = true; state.render(); } },
        render() {
          window.__atlasDirty = state.dirty;
          const input = document.querySelector('.titlebar .docname input');
          if (input && document.activeElement !== input) input.value = state.name;
          const saved = document.querySelector('.status-saved');
          if (saved) saved.textContent = state.dirty ? 'Unsaved changes' : (state.path ? 'Saved' : '');
          Platform.setTitle(`${state.dirty ? '• ' : ''}${state.name} — ${appLabel}`);
        }
      };
      const input = document.querySelector('.titlebar .docname input');
      if (input) input.addEventListener('change', () => { state.set({ name: input.value.trim() || defaultName }); });
      state.render();
      return state;
    }
  };

  /** Records a file in the browser's recent list (desktop keeps its own list in the main process). */
  Platform.rememberRecent = function (name, app) {
    if (desktop) return;
    const list = JSON.parse(localStorageGet('atlas-recent') || '[]').filter(r => r.name !== name);
    list.unshift({ name, app, opened: Date.now() });
    localStorageSet('atlas-recent', JSON.stringify(list.slice(0, 20)));
  };

  Platform.extOf = name => (name.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
  Platform.baseName = name => name.replace(/\.[^.]+$/, '');
  Platform.bytesToText = bytes => new TextDecoder().decode(bytes);
  Platform.bytesToDataUrl = (bytes, mime) => new Promise(res => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.readAsDataURL(new Blob([bytes], { type: mime }));
  });

  window.Platform = Platform;
  window.UI = UI;
})();
