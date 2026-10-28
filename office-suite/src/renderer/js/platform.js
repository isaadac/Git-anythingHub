'use strict';
// Shared runtime for all Isadek Office apps.
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
    /** Saves bytes or text. Returns {name, path} or null if cancelled. */
    async save({ data, name, path, filters, saveAs }) {
      // A document opened from (or saved to) a cloud account goes back there.
      if (Platform.isCloudPath(path) && window.Cloud) return window.Cloud.saveTo(path, name, data);
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
      // Web: a draft to restore (?recover=id), or a file handed over by the start page.
      const rid = new URLSearchParams(location.search).get('recover');
      if (rid) {
        const d = await Drafts.get(rid);
        history.replaceState(null, '', location.pathname);
        if (d) { sessionStorage.setItem('atlas-draft-id', rid); return { recovered: { draftId: rid, name: d.name, path: null, data: d.data } }; }
      }
      let raw = null;
      try { raw = sessionStorage.getItem('atlas-handoff'); sessionStorage.removeItem('atlas-handoff'); } catch { /* ignore */ }
      if (!raw) return null;
      const { name, dataUrl, cloudPath } = JSON.parse(raw);
      // Decode directly: fetch() of a data: URL is (rightly) blocked by the pages' CSP.
      const bin = atob(String(dataUrl).slice(String(dataUrl).indexOf(',') + 1));
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return { name, path: Platform.isCloudPath(cloudPath) ? cloudPath : null, data: bytes };
    },
    /** Desktop: asks where to save without writing anything yet. Web: asks for a file name. */
    async chooseSavePath(name, filters) {
      if (desktop) return window.atlas.chooseSavePath({ defaultName: name, filters });
      let fileName = await UI.prompt('Save As', 'File name', name);
      if (!fileName) return null;
      // Keep a supported extension so the file is written in a format we can produce.
      const exts = (filters || []).flatMap(f => f.extensions);
      if (exts.length && !exts.includes(Platform.extOf(fileName))) fileName += '.' + (Platform.extOf(name) || exts[0]);
      return { name: fileName, path: null };
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
    /** Opens new, unsaved content (bytes) in an app window: a template, a merge result… */
    async openData(app, name, bytes, sameTab, cloudPath) {
      if (desktop) return window.atlas.openData(app, name, bytes, cloudPath);
      const dataUrl = await Platform.bytesToDataUrl(bytes, 'application/octet-stream');
      try { sessionStorage.setItem('atlas-handoff', JSON.stringify({ name, dataUrl, cloudPath })); } catch { UI.toast('Too large to open in a new tab'); return false; }
      if (sameTab) location.href = app + '.html'; else window.open(app + '.html', '_blank');
      return true;
    },
    setTitle(t) { document.title = t; if (desktop) window.atlas.setTitle(t); },
    closeNow() { if (desktop) window.atlas.closeNow(); },
    fullscreen(on) {
      if (desktop) return window.atlas.fullscreen(on);
      if (on) document.documentElement.requestFullscreen?.(); else if (document.fullscreenElement) document.exitFullscreen?.();
    },
    onMenu(cb) { menuHandlers.push(cb); },
    emit(cmd) {
      // Undo/redo inside an ordinary text box (find bar, notes, dialogs) belongs to that box.
      const a = document.activeElement;
      if ((cmd === 'undo' || cmd === 'redo') && a && (a.tagName === 'INPUT' || (a.tagName === 'TEXTAREA' && !a.dataset.appEditor))) {
        document.execCommand(cmd);
        return;
      }
      menuHandlers.forEach(h => h(cmd));
    }
  };

  function localStorageGet(k) { try { return localStorage.getItem(k); } catch { return null; } }
  function localStorageSet(k, v) { try { localStorage.setItem(k, v); } catch { /* ignore */ } }

  // File → Office Theme in every app.
  Platform.onMenu(cmd => { if (cmd === 'theme' && window.Theme) Theme.choose(); });
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
      if (!el) { el = document.createElement('div'); el.className = 'toast'; el.setAttribute('role', 'status'); el.setAttribute('aria-live', 'polite'); document.body.appendChild(el); }
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
          if (f.placeholder) input.placeholder = f.placeholder;
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
    /** A message with a row of buttons; resolves to the index of the one pressed, or -1 (Escape). */
    choose(title, message, buttons, icon = '') {
      return new Promise(resolve => {
        const back = document.createElement('div');
        back.className = 'modal-backdrop';
        back.innerHTML = '<div class="modal" role="alertdialog"><h2></h2><p style="white-space:pre-wrap"></p><div class="actions"></div></div>';
        back.querySelector('h2').textContent = (icon ? icon + ' ' : '') + title;
        back.querySelector('p').textContent = message;
        const row = back.querySelector('.actions');
        const done = i => { back.remove(); resolve(i); };
        buttons.forEach((label, i) => { const b = document.createElement('button'); b.textContent = label; if (i === 0) b.className = 'primary'; b.onclick = () => done(i); row.appendChild(b); });
        back.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); done(-1); } });
        document.body.appendChild(back);
        row.querySelector('button').focus();
      });
    },
    confirm(title, message) {
      return new Promise(resolve => {
        const back = document.createElement('div');
        back.className = 'modal-backdrop';
        back.innerHTML = `<div class="modal"><h2></h2><p></p><div class="actions"><button class="cancel">Cancel</button><button class="primary ok">OK</button></div></div>`;
        back.querySelector('h2').textContent = title;
        back.querySelector('p').textContent = message;
        document.body.appendChild(back);
        const done = ok => { back.remove(); resolve(ok); };
        back.querySelector('.ok').onclick = () => done(true);
        back.querySelector('.cancel').onclick = () => done(false);
        back.addEventListener('keydown', e => {
          if (e.key === 'Escape') { e.preventDefault(); done(false); }
          if (e.key === 'Enter') { e.preventDefault(); done(true); }
        });
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
      if (input) input.addEventListener('change', () => {
        const name = input.value.trim() || defaultName;
        if (name === state.name) return;
        // A renamed document is a new file: the next Save asks where to put it instead of
        // overwriting the old one under the old name.
        state.set({ name, path: null, dirty: true });
      });
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

  /**
   * Cleans untrusted HTML (opened files, pasted content) with an allow-list: no scripts,
   * event handlers, frames or forms; links only to http(s)/mailto/#anchors; images only as
   * data: URLs; ids and classes that can't collide with the app's own elements.
   */
  const DROP_TAGS = 'script,style,iframe,object,embed,link,meta,form,input,button,textarea,select,option,base,frame,frameset,noscript,template,svg,math,canvas,audio,video,source,track,portal,applet,dialog,slot,title,head';
  const ALLOWED_ATTRS = new Set(['href', 'src', 'alt', 'title', 'style', 'class', 'id', 'colspan', 'rowspan', 'width', 'height', 'align', 'valign', 'color', 'face', 'size', 'start', 'type', 'dir', 'lang', 'reversed', 'value', 'border', 'cellpadding', 'cellspacing', 'bgcolor', 'span', 'headers', 'scope']);
  const ALLOWED_CLASSES = new Set(['title', 'page-break', 'textbox', 'comment', 'toc', 'toc-title', 'toc-l1', 'toc-l2', 'toc-l3', 'toc-pg', 'fld', 'pgnum', 'tc', 'note', 'citation', 'bibliography', 'bib-entry', 'mergefield', 'decorative', 'todo', 'done', 'tag-important', 'tag-question', 'tag-idea']);
  Platform.sanitizeHtml = function (html) {
    const d = new DOMParser().parseFromString(String(html || ''), 'text/html');
    d.querySelectorAll(DROP_TAGS).forEach(n => n.remove());
    // The app's own ids; ids in the document being edited may come back (e.g. reopening a file).
    const reserved = new Set([...document.querySelectorAll('[id]')].filter(e => !e.parentElement || !e.parentElement.closest('[contenteditable="true"]')).map(e => e.id));
    const clean = v => v.replace(/[\u0000-\u0020\u007f-\u009f]/g, '');
    d.body.querySelectorAll('*').forEach(el => {
      [...el.attributes].forEach(a => {
        const n = a.name.toLowerCase();
        if (n === 'data-fld' && ['page', 'numpages'].includes(a.value)) return; // page-number fields
        if (n === 'data-author' && /^[\w .@'-]{1,64}$/.test(a.value)) return; // review: who made a change/comment
        if ((n === 'data-date' || n === 'datetime') && !isNaN(Date.parse(a.value))) return;
        if (n === 'data-kind' && ['foot', 'end'].includes(a.value)) return;   // footnote / endnote
        if (n === 'data-src' && /^s[a-z0-9]{4,24}$/.test(a.value)) return;    // citation source id
        if (n === 'data-pages' && /^[\w .,–-]{1,40}$/.test(a.value)) return;
        if (n === 'data-field' && /^[\w .'-]{1,64}$/.test(a.value)) return; // mail-merge field name
        if (!ALLOWED_ATTRS.has(n)) { el.removeAttribute(a.name); return; }
        if (n === 'href' && !/^(https?:|mailto:|#)/i.test(clean(a.value))) el.removeAttribute(a.name);
        if (n === 'src' && !/^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml)[;,]/i.test(clean(a.value))) el.removeAttribute(a.name);
        if (n === 'id' && (reserved.has(a.value) || !/^[A-Za-z_][\w-]{0,80}$/.test(a.value))) el.removeAttribute(a.name);
      });
      if (el.classList.length) {
        const keep = [...el.classList].filter(c => ALLOWED_CLASSES.has(c));
        if (keep.length) el.className = keep.join(' '); else el.removeAttribute('class');
      }
      // Content may not float over the app's own interface.
      if (el.style && el.style.position) el.style.removeProperty('position');
    });
    d.body.querySelectorAll('img:not([src])').forEach(n => n.remove());
    return d.body.innerHTML;
  };
  /**
   * Inserts inline HTML (a field, a note marker…) exactly at the caret. Chromium's insertHTML
   * moves non-editable inline elements out of the paragraph at a line end and drops the space
   * before them, so this uses the DOM directly and then reports an input to the editor.
   */
  Platform.insertInline = function (html) {
    const sel = getSelection();
    if (!sel.rangeCount) return null;
    const range = sel.getRangeAt(0);
    const host = (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement).closest('[contenteditable="true"]');
    if (!host) return null;
    range.deleteContents();
    const tpl = document.createElement('template');
    tpl.innerHTML = html;
    const nodes = [...tpl.content.childNodes];
    range.insertNode(tpl.content);
    const last = nodes[nodes.length - 1];
    const after = document.createRange();
    after.setStartAfter(last);
    after.collapse(true);
    sel.removeAllRanges();
    sel.addRange(after);
    host.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertFromPaste' }));
    return nodes;
  };
  Platform.isColor = v => typeof v === 'string' && /^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.test(v);

  // ---------- Drafts: crash recovery storage ----------
  // Desktop keeps drafts in the app's data folder (main process); the web uses IndexedDB.
  const idb = () => new Promise((resolve, reject) => {
    const req = indexedDB.open('atlas-recovery', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('drafts');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  const idbDo = async (mode, fn) => {
    const db = await idb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction('drafts', mode);
      const r = fn(tx.objectStore('drafts'));
      tx.oncomplete = () => resolve(r && r.result);
      tx.onerror = () => reject(tx.error);
    });
  };
  const Drafts = {
    async save(d) {
      if (desktop) return window.atlas.recoverySave(d);
      try { await idbDo('readwrite', st => st.put({ app: d.app, name: d.name, data: d.data, savedAt: Date.now() }, d.draftId)); } catch { /* storage full or blocked */ }
    },
    async clear(id) {
      if (desktop) return window.atlas.recoveryClear();
      try { await idbDo('readwrite', st => st.delete(id)); } catch { /* ignore */ }
    },
    async list() {
      if (desktop) return window.atlas.recoveryList();
      try {
        const db = await idb();
        return await new Promise(resolve => {
          const out = [];
          const req = db.transaction('drafts').objectStore('drafts').openCursor();
          req.onsuccess = () => {
            const c = req.result;
            if (!c) return resolve(out.sort((a, b) => b.savedAt - a.savedAt));
            out.push({ id: c.key, app: c.value.app, name: c.value.name, savedAt: c.value.savedAt });
            c.continue();
          };
          req.onerror = () => resolve([]);
        });
      } catch { return []; }
    },
    async get(id) { try { return await idbDo('readonly', st => st.get(id)); } catch { return null; } },
    async open(id, app) {
      if (desktop) return window.atlas.recoveryOpen(id);
      location.href = `${app}.html?recover=${encodeURIComponent(id)}`;
    },
    async discard(id) {
      if (desktop) return window.atlas.recoveryDiscard(id);
      try { await idbDo('readwrite', st => st.delete(id)); } catch { /* ignore */ }
    }
  };
  Platform.drafts = Drafts;

  /**
   * Turns on crash-recovery drafts and the AutoSave switch for an app.
   * opts: { app, state, snapshot() -> string, autosave() -> Promise, canAutoSave() -> bool|string }
   * canAutoSave returns true, or a reason why AutoSave can't be used for the current file.
   */
  Platform.enableRecovery = function (opts) {
    const { app, state } = opts;
    let draftId = null;
    try { draftId = sessionStorage.getItem('atlas-draft-id'); } catch { /* ignore */ }
    if (!draftId) draftId = 'd' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
    try { sessionStorage.setItem('atlas-draft-id', draftId); } catch { /* ignore */ }
    let lastSaved = 0, changedSince = 0;

    async function writeDraft() {
      if (!state.dirty || changedSince <= lastSaved) return;
      lastSaved = Date.now();
      try { await Drafts.save({ draftId, app, name: state.name, path: state.path, data: opts.snapshot() }); } catch { /* best effort */ }
    }
    setInterval(writeDraft, 30000);
    document.addEventListener('visibilitychange', () => { if (document.hidden) writeDraft(); });
    window.addEventListener('pagehide', writeDraft);

    // AutoSave switch in the title bar.
    const bar = document.querySelector('.titlebar');
    const label = document.createElement('label');
    label.className = 'autosave';
    label.innerHTML = '<input type="checkbox" role="switch"> <span>AutoSave</span>';
    bar.insertBefore(label, bar.querySelector('.qa'));
    const box = label.querySelector('input');
    const key = 'atlas-autosave-' + app;
    let wanted = false;
    try { wanted = localStorage.getItem(key) === '1'; } catch { /* ignore */ }
    function refreshSwitch() {
      const ok = opts.canAutoSave();
      box.disabled = ok !== true;
      box.checked = wanted && ok === true;
      label.title = ok === true ? 'Save changes to the file automatically' : ok;
      label.classList.toggle('off', ok !== true);
    }
    box.addEventListener('change', () => {
      wanted = box.checked;
      try { localStorage.setItem(key, wanted ? '1' : '0'); } catch { /* ignore */ }
      if (wanted && state.dirty) scheduleAutoSave(500);
    });

    let autoTimer = null, saving = false;
    function scheduleAutoSave(delay = 3000) {
      clearTimeout(autoTimer);
      autoTimer = setTimeout(async () => {
        if (!wanted || opts.canAutoSave() !== true || !state.dirty) return;
        if (saving) return scheduleAutoSave(1000);
        saving = true;
        try { await opts.autosave(); } finally { saving = false; }
      }, delay);
    }

    // Hook the document state: changes schedule drafts/AutoSave; a clean state removes the draft.
    const markDirty = state.markDirty.bind(state);
    state.markDirty = () => { changedSince = Date.now(); markDirty(); scheduleAutoSave(); };
    const set = state.set.bind(state);
    state.set = props => {
      set(props);
      if (props.dirty === false) { lastSaved = changedSince = 0; Drafts.clear(draftId); }
      if (props.dirty === true) { changedSince = Date.now(); scheduleAutoSave(); }
      if ('path' in props || 'name' in props) refreshSwitch();
    };
    refreshSwitch();
    return {
      draftId,
      /** Continue the draft that was restored, and write it again right away. */
      adopt(id) {
        draftId = id;
        try { sessionStorage.setItem('atlas-draft-id', id); } catch { /* ignore */ }
        lastSaved = 0; changedSince = Date.now();
        writeDraft();
      },
      refresh: refreshSwitch
    };
  };

  Platform.isCloudPath = p => typeof p === 'string' && /^cloud:\/\/[\w-]{1,40}\//.test(p);
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
