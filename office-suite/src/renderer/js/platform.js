'use strict';
// Shared runtime for all Isadek Office apps.
// Wraps the desktop bridge (window.atlas, from preload.js) and falls back to plain
// browser APIs so the same apps also run as a web page on phones, tablets and ChromeOS.
(function () {
  const desktop = !!(window.atlas && window.atlas.isDesktop);
  const menuHandlers = [];

  // Script errors go to the crash log (desktop; kept on the computer unless crash reports are on).
  if (desktop && window.atlas.reportError) {
    const report = (kind, err, source) => { try { window.atlas.reportError({ kind, message: String(err && err.message || err || ''), stack: String(err && err.stack || ''), source: String(source || '') }).catch(() => {}); } catch { /* ignore */ } };
    window.addEventListener('error', e => { if (e.error || e.message) report('error', e.error || e.message, e.filename ? `${e.filename.replace(/^.*\/(?=[^/]+$)/, '')}:${e.lineno}:${e.colno}` : ''); });
    window.addEventListener('unhandledrejection', e => report('rejection', e.reason));
  }

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
    async save(opts) {
      const r = await saveNow(opts);
      // Each save also keeps a version of the document (File → Version History).
      if (r && opts.data && !opts.noVersion) Versions.add(r.path || opts.path || r.name || opts.name, r.name || opts.name, opts.data).catch(() => {});
      return r;
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
    /** The Help window (desktop) or page (web), at a section. */
    help(section = 'start') {
      if (desktop && window.atlas.help) return window.atlas.help(section);
      window.open('help.html#' + encodeURIComponent(section), 'isadek-help');
    },
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

  // File → Office Theme and File → Language in every app.
  Platform.onMenu(cmd => {
    if (cmd === 'theme' && window.Theme) Theme.choose();
    if (cmd === 'language' && window.I18n) I18n.choose();
    if (cmd === 'help') Platform.help();
  });
  if (desktop) {
    window.atlas.onMenu(cmd => Platform.emit(cmd));
  } else {
    // In the browser there is no native menu, so map the usual shortcuts ourselves.
    document.addEventListener('keydown', e => { if (e.key === 'F1') { e.preventDefault(); Platform.help(); } });
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
      // Panes that float beside the document start just below the ribbon, wherever that is.
      const ribbon = document.querySelector('.ribbon');
      if (ribbon && window.ResizeObserver) {
        const sync = () => document.documentElement.style.setProperty('--chrome-h', Math.round(ribbon.getBoundingClientRect().bottom) + 'px');
        new ResizeObserver(sync).observe(ribbon);
        sync();
      }
      const tabs = document.querySelectorAll('.ribbon-tabs button[data-tab]');
      tabs.forEach(btn => btn.addEventListener('click', () => {
        // The new panel slides in from the side of the tab you moved to.
        const was = [...tabs].findIndex(b => b.classList.contains('active')), now = [...tabs].indexOf(btn);
        tabs.forEach(b => b.classList.toggle('active', b === btn));
        document.querySelectorAll('.ribbon-panel').forEach(p => {
          p.classList.toggle('active', p.dataset.panel === btn.dataset.tab);
          p.classList.toggle('from-right', now > was);
        });
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
      UI.commandSearch();
      UI.ribbonModes();
      // A ribbon too wide for the window gets more compact step by step, as Office's does: small
      // buttons drop their words (still read out and shown as tooltips), then big buttons shrink.
      if (ribbon && window.ResizeObserver) {
        const fit = () => {
          ribbon.classList.remove('compact-1', 'compact-2');
          const panel = ribbon.querySelector('.ribbon-panel.active');
          if (!panel || document.body.classList.contains('ribbon-tabs') && !document.body.classList.contains('ribbon-peek')) return;
          for (const step of ['compact-1', 'compact-2']) {
            if (ribbon.scrollWidth <= ribbon.clientWidth + 1) break;
            ribbon.classList.add(step);
          }
        };
        new ResizeObserver(fit).observe(ribbon);
        document.querySelectorAll('.ribbon-tabs button').forEach(t => t.addEventListener('click', () => requestAnimationFrame(fit)));
        window.addEventListener('resize', fit);
        // Widths change again once the icons are drawn and the fonts have loaded.
        requestAnimationFrame(fit);
        window.addEventListener('load', fit);
        if (document.fonts) document.fonts.ready.then(fit);
      }
      UI.miniToolbar({ 'app-write': '#doc', 'app-notes': '#body', 'app-slides': '#stage' }[[...document.body.classList].find(c => c.startsWith('app-'))]);
      // Keep focus in the editor when pressing ribbon buttons.
      document.querySelectorAll('.ribbon .btn, .titlebar .qa button').forEach(b => b.addEventListener('mousedown', e => e.preventDefault()));
    },
    /**
     * How much of the ribbon shows: the full ribbon, a single line of commands, or only the tabs
     * (a tab then opens the ribbon over the document until a command is used). Remembered;
     * Ctrl+F1 and double-clicking a tab switch between full and tabs only, as in Office.
     */
    ribbonModes() {
      const tabsBar = document.querySelector('.ribbon-tabs'), ribbon = document.querySelector('.ribbon');
      if (!tabsBar || !ribbon || tabsBar.querySelector('.ribbon-mode')) return;
      // icons: panel-top rows-2 panel-top-close chevron-up chevron-down check
      const MODES = [['full', 'Full ribbon', 'panel-top'], ['line', 'Single line', 'rows-2'], ['tabs', 'Tabs only', 'panel-top-close']];
      const read = () => { const v = localStorageGet('atlas-ribbon'); return MODES.some(m => m[0] === v) ? v : 'full'; };
      const btn = document.createElement('button');
      btn.className = 'ribbon-mode';
      btn.title = 'Ribbon display options';
      btn.setAttribute('aria-haspopup', 'menu');
      btn.innerHTML = '<span class="ico" data-icon="chevron-down"></span>';
      tabsBar.appendChild(btn);
      let mode = read(), openTab = false;
      const apply = m => {
        mode = m;
        localStorageSet('atlas-ribbon', m);
        document.body.classList.remove('ribbon-full', 'ribbon-line', 'ribbon-tabs');
        document.body.classList.add('ribbon-' + m);
        document.body.classList.remove('ribbon-peek');
        Icons.set(btn.querySelector('.ico'), m === 'tabs' ? 'chevron-down' : 'chevron-up');
        window.dispatchEvent(new Event('resize'));
      };
      apply(mode);
      // Tabs only: a tab shows the ribbon over the document until a command is used or you click away.
      tabsBar.querySelectorAll('button[data-tab]').forEach(t => {
        t.addEventListener('click', () => { if (mode === 'tabs') { document.body.classList.add('ribbon-peek'); openTab = true; } });
        t.addEventListener('dblclick', () => apply(mode === 'tabs' ? 'full' : 'tabs'));
      });
      const closePeek = () => { if (openTab) { document.body.classList.remove('ribbon-peek'); openTab = false; } };
      ribbon.addEventListener('click', e => { if (mode === 'tabs' && e.target.closest('button:not(.ribbon-mode)') && !e.target.closest('select')) setTimeout(closePeek, 0); });
      document.addEventListener('mousedown', e => { if (openTab && !e.target.closest('.ribbon, .ribbon-tabs')) closePeek(); });
      document.addEventListener('keydown', e => {
        if (e.key === 'Escape') closePeek();
        if (e.key === 'F1' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); apply(mode === 'tabs' ? 'full' : 'tabs'); }
      });
      btn.addEventListener('click', () => {
        const old = document.querySelector('.ribbon-menu');
        if (old) { old.remove(); return; }
        const menu = document.createElement('div');
        menu.className = 'ribbon-menu';
        menu.setAttribute('role', 'menu');
        MODES.forEach(([m, label, icon]) => {
          const b = document.createElement('button');
          b.setAttribute('role', 'menuitemradio');
          b.setAttribute('aria-checked', String(m === mode));
          b.innerHTML = `<span class="ico" data-icon="${icon}"></span><span></span><span class="ico check" data-icon="check"></span>`;
          b.children[1].textContent = label;
          b.addEventListener('click', () => { apply(m); menu.remove(); btn.focus(); });
          menu.appendChild(b);
        });
        const r = btn.getBoundingClientRect();
        menu.style.top = Math.round(r.bottom + 6) + 'px';
        menu.style.right = Math.round(window.innerWidth - r.right) + 'px';
        document.body.appendChild(menu);
        menu.querySelector('[aria-checked="true"]').focus();
        menu.addEventListener('keydown', e => {
          const items = [...menu.querySelectorAll('button')], i = items.indexOf(document.activeElement);
          if (e.key === 'ArrowDown') { e.preventDefault(); items[(i + 1) % items.length].focus(); }
          if (e.key === 'ArrowUp') { e.preventDefault(); items[(i - 1 + items.length) % items.length].focus(); }
          if (e.key === 'Escape') { menu.remove(); btn.focus(); }
        });
        setTimeout(() => document.addEventListener('mousedown', function off(e) { if (!menu.contains(e.target) && e.target !== btn && !btn.contains(e.target)) { menu.remove(); } document.removeEventListener('mousedown', off); }), 0);
      });
      Platform.ribbonMode = { get: () => mode, set: apply };
    },
    /**
     * The mini toolbar: the formatting commands most used on a selection, floating just above it.
     * Each button presses the ribbon's own button, so it behaves exactly the same.
     */
    miniToolbar(rootSel) {
      if (!rootSel || document.querySelector('.minibar')) return;
      // icons: bold italic underline strikethrough a-arrow-up a-arrow-down highlighter baseline list list-ordered link message-square-plus
      const ITEMS = [
        ['.ribbon [data-exec="bold"]', 'bold', 'Bold'], ['.ribbon [data-exec="italic"]', 'italic', 'Italic'], ['.ribbon [data-exec="underline"]', 'underline', 'Underline'],
        ['.ribbon [data-exec="strikeThrough"]', 'strikethrough', 'Strikethrough'], null,
        ['.ribbon #growFont', 'a-arrow-up', 'Increase font size'], ['.ribbon #shrinkFont', 'a-arrow-down', 'Decrease font size'], null,
        ['.ribbon .color-btn[title="Text highlight colour"] input', 'highlighter', 'Highlight'], ['.ribbon .color-btn[title="Font colour"] input', 'baseline', 'Font colour'], null,
        ['.ribbon [data-exec="insertUnorderedList"]', 'list', 'Bullets'], ['.ribbon [data-exec="insertOrderedList"]', 'list-ordered', 'Numbering'], null,
        ['.ribbon #insLink', 'link', 'Link'], ['.ribbon #addComment', 'message-square-plus', 'New Comment']
      ];
      const bar = document.createElement('div');
      bar.className = 'minibar';
      bar.setAttribute('role', 'toolbar');
      bar.setAttribute('aria-label', 'Formatting');
      let lastSep = true;
      ITEMS.forEach(it => {
        if (!it) { if (!lastSep) { const s = document.createElement('span'); s.className = 'sep'; bar.appendChild(s); lastSep = true; } return; }
        const target = document.querySelector(it[0]);
        if (!target) return;
        const b = document.createElement('button');
        b.className = 'btn';
        b.title = it[2];
        b.setAttribute('aria-label', it[2]);
        b.innerHTML = `<span class="ico" data-icon="${it[1]}"></span>`;
        b.addEventListener('mousedown', e => e.preventDefault());
        b.addEventListener('click', () => { target.click(); setTimeout(sync, 0); });
        b._target = target;
        bar.appendChild(b);
        lastSep = false;
      });
      if (bar.lastChild && bar.lastChild.className === 'sep') bar.lastChild.remove();
      if (!bar.children.length) return;
      bar.hidden = true;
      document.body.appendChild(bar);
      const sync = () => bar.querySelectorAll('.btn').forEach(b => {
        const t = b._target.closest('.btn') || b._target;
        b.classList.toggle('on', t.classList.contains('on') || t.getAttribute('aria-pressed') === 'true');
      });
      const hide = () => { bar.classList.remove('show'); bar.hidden = true; };
      const selectionRect = () => {
        const sel = getSelection();
        if (!sel.rangeCount || sel.isCollapsed) return null;
        const root = sel.anchorNode && (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement);
        if (!root || !root.closest(rootSel) || !root.closest('[contenteditable="true"]')) return null;
        const r = sel.getRangeAt(0).getBoundingClientRect();
        return r.width || r.height ? r : null;
      };
      const show = () => {
        const r = selectionRect();
        if (!r || !String(getSelection()).trim()) return hide();
        bar.hidden = false;
        sync();
        const w = bar.offsetWidth, h = bar.offsetHeight, gap = 10;
        let top = r.top - h - gap;
        if (top < 8 + (parseFloat(getComputedStyle(document.documentElement).getPropertyValue('--chrome-h')) || 0)) top = r.bottom + gap;
        const left = Math.max(8, Math.min(window.innerWidth - w - 8, r.left + r.width / 2 - w / 2));
        bar.style.top = Math.round(top) + 'px';
        bar.style.left = Math.round(left) + 'px';
        requestAnimationFrame(() => bar.classList.add('show'));
      };
      // After selecting with the mouse (as Office does), not while typing or selecting with keys.
      document.addEventListener('mouseup', e => { if (!bar.contains(e.target)) setTimeout(show, 10); });
      document.addEventListener('dblclick', () => setTimeout(show, 10));
      document.addEventListener('keydown', e => { if (!bar.contains(e.target) && !['Shift', 'Control', 'Meta', 'Alt'].includes(e.key)) hide(); });
      document.addEventListener('mousedown', e => { if (!bar.contains(e.target)) hide(); });
      document.addEventListener('selectionchange', () => { if (!bar.hidden && !selectionRect()) hide(); });
      window.addEventListener('scroll', hide, true);
      window.addEventListener('resize', hide);
      UI.miniBar = { show, hide, el: bar };
    },
    /**
     * Search box in the title bar (Alt+Q): finds ribbon and File commands by name or tooltip
     * and runs them — switching to the right tab, then pressing the button (or focusing a list).
     */
    commandSearch() {
      const bar = document.querySelector('.titlebar');
      if (!bar || bar.querySelector('.tell-me')) return;
      const box = document.createElement('div');
      box.className = 'tell-me';
      box.innerHTML = '<input type="search" placeholder="Search (Alt+Q)" aria-label="Search for a command" autocomplete="off" spellcheck="false"><div class="tm-list" role="listbox" hidden></div>';
      bar.insertBefore(box, bar.querySelector('.status-saved') || bar.querySelector('.share-btn'));
      const input = box.querySelector('input'), list = box.querySelector('.tm-list');
      const clean = t => String(t || '').replace(/[^\p{L}\p{N}&%'’ .,:/+-]/gu, ' ').replace(/\s+/g, ' ').trim();
      // A button's words, with a big button's two lines as one (and without its icon).
      const words = el => { const c = el.cloneNode(true); c.querySelectorAll('.ico').forEach(n => n.remove()); c.querySelectorAll('br').forEach(n => n.replaceWith(' ')); return c.textContent; };
      let index = null, hits = [], at = 0;
      function build() {
        index = [];
        document.querySelectorAll('.ribbon-panel').forEach(panel => {
          const tabBtn = document.querySelector(`.ribbon-tabs button[data-tab="${panel.dataset.panel}"]`);
          if (!tabBtn) return;
          const tab = clean(tabBtn.textContent);
          panel.querySelectorAll('button, select, input[type=checkbox], input[type=color]').forEach(el => {
            if (el.closest('[hidden]')) return;
            const group = clean(el.closest('.group') && el.closest('.group').querySelector('.label') ? el.closest('.group').querySelector('.label').textContent : '');
            const label = clean(el.tagName === 'SELECT' ? (el.title || el.options[0] && el.options[0].text) : el.tagName === 'INPUT' ? (el.closest('label') ? el.closest('label').textContent : el.title) : words(el)) || clean(el.title);
            if (!label) return;
            index.push({ el, tabBtn, label, extra: clean(el.title) + ' ' + group, where: tab + (group ? ' › ' + group : '') });
          });
        });
        document.querySelectorAll('.backstage nav button[data-cmd]').forEach(el => index.push({ cmd: el.dataset.cmd, label: clean(el.textContent), extra: clean(el.title), where: 'File' }));
      }
      function search(q) {
        const words = q.toLowerCase().split(/\s+/).filter(Boolean);
        if (!words.length) return [];
        return index.map(x => {
          const hay = (x.label + ' ' + x.extra + ' ' + x.where).toLowerCase(), lab = x.label.toLowerCase();
          if (!words.every(w => hay.includes(w))) return null;
          const score = (lab.startsWith(words[0]) ? 0 : lab.includes(words[0]) ? 1 : 2) + (x.label.length / 100);
          return { x, score };
        }).filter(Boolean).sort((a, b) => a.score - b.score).slice(0, 8).map(h => h.x);
      }
      function draw() {
        list.innerHTML = '';
        hits.forEach((h, i) => {
          const row = document.createElement('div');
          row.className = 'tm-item' + (i === at ? ' on' : '');
          row.setAttribute('role', 'option');
          row.setAttribute('aria-selected', String(i === at));
          row.innerHTML = '<span class="l"></span><span class="w"></span>';
          row.querySelector('.l').textContent = h.label;
          row.querySelector('.w').textContent = h.where;
          row.addEventListener('mousedown', e => { e.preventDefault(); run(h); });
          list.appendChild(row);
        });
        list.hidden = !hits.length;
      }
      function run(h) {
        input.value = ''; hits = []; draw(); input.blur();
        if (h.cmd) { Platform.emit(h.cmd); return; }
        h.tabBtn.click();
        if (h.el.tagName === 'SELECT' || h.el.tagName === 'INPUT') { h.el.focus(); if (h.el.showPicker) try { h.el.showPicker(); } catch { /* not allowed */ } }
        else h.el.click();
        h.el.classList.add('tm-flash');
        setTimeout(() => h.el.classList.remove('tm-flash'), 1200);
      }
      input.addEventListener('focus', () => { build(); });
      input.addEventListener('input', () => { at = 0; hits = search(input.value); draw(); });
      input.addEventListener('keydown', e => {
        e.stopPropagation();
        if (e.key === 'ArrowDown' && hits.length) { e.preventDefault(); at = (at + 1) % hits.length; draw(); }
        else if (e.key === 'ArrowUp' && hits.length) { e.preventDefault(); at = (at - 1 + hits.length) % hits.length; draw(); }
        else if (e.key === 'Enter' && hits[at]) { e.preventDefault(); run(hits[at]); }
        else if (e.key === 'Escape') { input.value = ''; hits = []; draw(); input.blur(); }
      });
      input.addEventListener('blur', () => setTimeout(() => { list.hidden = true; }, 150));
      document.addEventListener('keydown', e => { if (e.altKey && !e.ctrlKey && !e.metaKey && (e.key === 'q' || e.key === 'Q')) { e.preventDefault(); input.focus(); input.select(); } });
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

  async function saveNow({ data, name, path, filters, saveAs }) {
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
  }

  // ---------- Version history ----------
  // Desktop: the main process keeps copies in the app's data folder (src/main/versions.js).
  // Web: IndexedDB, with the same rules (one copy per 10 minutes, the latest 30 per document).
  const Versions = (() => {
    if (desktop) return { add: (k, n, d) => window.atlas.versionsAdd(String(k), n, toU8(d)), list: k => window.atlas.versionsList(String(k)), read: (k, id) => window.atlas.versionsRead(String(k), id) };
    let dbp = null;
    const db = () => dbp || (dbp = new Promise((res, rej) => {
      const r = indexedDB.open('atlas-versions', 1);
      r.onupgradeneeded = () => { const st = r.result.createObjectStore('v', { keyPath: ['key', 'id'] }); st.createIndex('key', 'key'); };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    }));
    const all = async key => { const d = await db(); return new Promise((res, rej) => { const q = d.transaction('v').objectStore('v').index('key').getAll(key); q.onsuccess = () => res(q.result.sort((a, b) => a.saved - b.saved)); q.onerror = () => rej(q.error); }); };
    const hashOf = async bytes => { try { return [...new Uint8Array(await crypto.subtle.digest('SHA-1', bytes))].map(b => b.toString(16).padStart(2, '0')).join(''); } catch { return String(bytes.length); } };
    return {
      async add(key, name, data) {
        key = String(key);
        const bytes = toU8(data);
        if (bytes.length > 60e6) return false;
        const list = await all(key), last = list[list.length - 1], hash = await hashOf(bytes), now = Date.now();
        if (last && last.hash === hash) return true;
        const d = await db();
        const t = d.transaction('v', 'readwrite'), st = t.objectStore('v');
        let created = now;
        if (last && now - last.created < 600000) { st.delete([key, last.id]); list.pop(); created = last.created; }
        st.put({ key, id: String(now), name: String(name || '').slice(0, 255), size: bytes.length, hash, created, saved: now, data: bytes });
        list.slice(0, Math.max(0, list.length + 1 - 30)).forEach(v => st.delete([key, v.id]));
        return new Promise(res => { t.oncomplete = () => res(true); t.onerror = () => res(false); });
      },
      async list(key) { return (await all(String(key))).map(v => ({ id: v.id, name: v.name, size: v.size, saved: v.saved })).reverse(); },
      async read(key, id) { const d = await db(); return new Promise(res => { const q = d.transaction('v').objectStore('v').get([String(key), String(id)]); q.onsuccess = () => res(q.result ? q.result.data : null); q.onerror = () => res(null); }); }
    };
  })();
  function toU8(d) { return d instanceof Uint8Array ? d : d instanceof ArrayBuffer ? new Uint8Array(d) : new TextEncoder().encode(String(d)); }
  Platform.versions = Versions;
  /**
   * File → Version History: the saved versions of the open document, each to open as a copy
   * or to restore. opts: { state, app, load(file) } (load is the app's own file loader).
   */
  Platform.versionHistory = async function ({ state, app, load }) {
    const key = state.path || state.name;
    let list = [];
    try { list = await Versions.list(key); } catch { list = []; }
    if (!list.length) return UI.toast('No earlier versions yet. Isadek Office keeps a version each time you save (at most one every 10 minutes).', 6000);
    const back = document.createElement('div');
    back.className = 'modal-backdrop';
    back.innerHTML = '<div class="modal versions" role="dialog" aria-label="Version History"><h2>Version History</h2><p class="muted"></p><div class="v-list" role="list"></div><div class="actions"><button class="cancel">Close</button></div></div>';
    back.querySelector('.muted').textContent = `${state.name} — ${list.length} version${list.length === 1 ? '' : 's'}`;
    const box = back.querySelector('.v-list');
    const fmtSize = n => (n < 1024 ? n + ' B' : n < 1048576 ? (n / 1024).toFixed(0) + ' KB' : (n / 1048576).toFixed(1) + ' MB');
    const close = () => back.remove();
    list.forEach((v, i) => {
      const row = document.createElement('div');
      row.className = 'v-row';
      row.setAttribute('role', 'listitem');
      const when = new Date(+v.saved);
      row.innerHTML = '<span class="when"></span><span class="size"></span><button data-a="open">Open a copy</button><button data-a="restore">Restore</button>';
      row.querySelector('.when').textContent = when.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' }) + (i === 0 ? ' (latest)' : '');
      row.querySelector('.size').textContent = fmtSize(v.size);
      row.querySelectorAll('button').forEach(b => b.addEventListener('click', async () => {
        const data = await Versions.read(key, v.id);
        if (!data) return UI.toast('That version could not be read');
        const ext = Platform.extOf(state.name) || Platform.extOf(v.name);
        if (b.dataset.a === 'open') {
          close();
          const stamp = when.toLocaleString(undefined, { dateStyle: 'short', timeStyle: 'short' }).replace(/[\\/:*?"<>|]/g, '-');
          await Platform.openData(app, `${Platform.baseName(state.name)} (version ${stamp})${ext ? '.' + ext : ''}`, new Uint8Array(data));
          return;
        }
        if (!(await UI.confirm('Restore Version', `Replace what is open with the version from ${when.toLocaleString()}? You can still undo by restoring the latest version.`))) return;
        close();
        await load({ name: state.name, path: state.path, data: new Uint8Array(data), recovered: true });
        state.set({ dirty: true });
        UI.toast('Version restored — save to keep it');
      }));
      box.appendChild(row);
    });
    back.querySelector('.cancel').addEventListener('click', close);
    back.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
    document.body.appendChild(back);
    back.querySelector('.cancel').focus();
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
  const ALLOWED_CLASSES = new Set(['title', 'page-break', 'textbox', 'comment', 'toc', 'toc-title', 'toc-l1', 'toc-l2', 'toc-l3', 'toc-pg', 'fld', 'pgnum', 'tc', 'note', 'citation', 'bibliography', 'bib-entry', 'mergefield', 'decorative', 'todo', 'done', 'tag-important', 'tag-question', 'tag-idea', 'section-break', 'equation', 'eq-block', 'float-left', 'float-right', 'wrap-block', 'attachment', 'audio-note']);
  Platform.sanitizeHtml = function (html) {
    const d = new DOMParser().parseFromString(String(html || ''), 'text/html');
    d.querySelectorAll(DROP_TAGS).forEach(n => n.remove());
    // The app's own ids; ids in the document being edited may come back (e.g. reopening a file).
    const reserved = new Set([...document.querySelectorAll('[id]')].filter(e => !e.parentElement || !e.parentElement.closest('[contenteditable="true"]')).map(e => e.id));
    const clean = v => v.replace(/[\u0000-\u0020\u007f-\u009f]/g, '');
    d.body.querySelectorAll('*').forEach(el => {
      [...el.attributes].forEach(a => {
        const n = a.name.toLowerCase();
        if (n === 'data-fld' && ['page', 'numpages', 'date', 'time', 'author', 'filename', 'numwords'].includes(a.value)) return; // fields
        if (n === 'data-instr' && el.classList.contains('fld') && a.value.length <= 500) return; // a Word field's instruction
        if (n === 'data-eq' && a.value.length <= 2000) return;                 // equation (drawn by the app from this text)
        if (n === 'data-orient' && ['portrait', 'landscape'].includes(a.value)) return; // section break
        if (n === 'data-cols' && ['1', '2', '3'].includes(a.value)) return;
        if (n === 'data-att' && /^f[\w]{4,40}$/.test(a.value)) return;        // Notes: an attachment
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
