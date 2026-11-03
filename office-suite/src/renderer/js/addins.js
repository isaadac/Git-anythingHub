'use strict';
// Isadek — add-ins: small programs that add buttons and a pane to Write, Sheets or Slides.
//
// An add-in is one JSON file (.isadek-addin):
//   { "id": "com.example.word-goal", "name": "Word Goal", "version": "1.0.0", "author": "…",
//     "description": "…", "apps": ["write"], "permissions": ["read", "write"],
//     "commands": [{ "id": "show", "label": "Word Goal", "icon": "target" }],
//     "script": "isadek.onCommand('show', async () => { … })" }
// It runs in its own worker (addin-worker.js) with no network, storage or page access, and works
// on the document only through the `isadek` API, which checks its permissions: "read" (text,
// selection, cells, slides) and "write" (insert text, set cells, add slides). Installing one shows
// what it asks for. Add-ins are kept on this device; see ADDINS.md for the API.
// Icons an add-in's buttons may use (others show a puzzle piece):
// icons: puzzle target calculator wand-2 list-checks chart-bar languages timer sparkles hash book-open clipboard-list calendar-days quote smile
(function () {
  const KEY = 'atlas-addins';
  const MAX_CODE = 1e6;
  const PERMS = { read: 'Read this document', write: 'Change this document' };
  const READS = new Set(['getText', 'getSelection', 'getValues', 'getSelectedRange', 'getSlides']);
  const WRITES = new Set(['insertText', 'setValues', 'addSlide']);
  const UI_CALLS = new Set(['toast', 'showPane', 'closePane', 'form']);
  const load = () => { try { const a = JSON.parse(localStorage.getItem(KEY) || '[]'); return Array.isArray(a) ? a : []; } catch { return []; } };
  const save = list => { try { localStorage.setItem(KEY, JSON.stringify(list)); return true; } catch { UI.toast('There is no room to keep that add-in', 5000); return false; } };

  /** Checks an add-in file: known fields, sensible sizes. Returns the add-in or throws. */
  function check(text) {
    let m;
    try { m = JSON.parse(text); } catch { throw new Error('This isn’t an add-in file (it isn’t JSON)'); }
    const str = (v, max) => (typeof v === 'string' ? v.trim().slice(0, max) : '');
    if (!m || typeof m !== 'object') throw new Error('This isn’t an add-in file');
    const id = str(m.id, 100);
    if (!/^[a-z0-9][a-z0-9.-]{2,99}$/i.test(id)) throw new Error('The add-in has no valid id (letters, digits, dots and dashes)');
    const name = str(m.name, 60);
    if (!name) throw new Error('The add-in has no name');
    if (typeof m.script !== 'string' || !m.script.trim()) throw new Error('The add-in has no script');
    if (m.script.length > MAX_CODE) throw new Error('The add-in is too large');
    const apps = (Array.isArray(m.apps) ? m.apps : []).filter(a => ['write', 'sheets', 'slides'].includes(a));
    if (!apps.length) throw new Error('The add-in doesn’t say which apps it works in (write, sheets, slides)');
    const permissions = (Array.isArray(m.permissions) ? m.permissions : []).filter(p => PERMS[p]);
    const commands = (Array.isArray(m.commands) ? m.commands : []).slice(0, 8).map(c => ({
      id: str(c && c.id, 60), label: str(c && c.label, 40), icon: /^[a-z0-9-]{1,40}$/.test(c && c.icon) ? c.icon : 'puzzle',
      apps: (Array.isArray(c && c.apps) ? c.apps : apps).filter(a => apps.includes(a))
    })).filter(c => c.id && c.label);
    if (!commands.length) throw new Error('The add-in has no commands (buttons)');
    return { id, name, version: str(m.version, 20) || '1.0.0', author: str(m.author, 80), description: str(m.description, 300), apps, permissions, commands, script: m.script };
  }

  function attach(adapter) {
    const app = adapter.app;
    const running = new Map(); // id -> { worker, addin, runs }
    let paneFor = null;

    // ---- The pane an add-in shows
    const pane = document.createElement('aside');
    pane.className = 'addin-pane';
    pane.hidden = true;
    pane.innerHTML = '<header><b class="ttl"></b><span class="sp"></span><button class="x" title="Close"><span class="ico" data-icon="x"></span></button></header><div class="body"></div><footer class="by"></footer>';
    document.body.appendChild(pane);
    pane.querySelector('.x').addEventListener('click', () => { pane.hidden = true; paneFor = null; });
    function showPane(addin, title, blocks) {
      paneFor = addin.id;
      pane.querySelector('.ttl').textContent = title || addin.name;
      pane.querySelector('.by').textContent = `Add-in: ${addin.name}${addin.author ? ' · ' + addin.author : ''}`;
      const body = pane.querySelector('.body');
      body.innerHTML = '';
      const inputs = {};
      // Everything is built as text and plain controls: nothing an add-in sends becomes HTML.
      for (const b of blocks.slice(0, 200)) {
        if (!b || typeof b !== 'object') continue;
        const text = v => String(v === undefined || v === null ? '' : v).slice(0, 5000);
        let el;
        if (b.type === 'heading') { el = document.createElement('h3'); el.textContent = text(b.text); }
        else if (b.type === 'text') { el = document.createElement('p'); el.textContent = text(b.text); }
        else if (b.type === 'list') { el = document.createElement('ul'); (Array.isArray(b.items) ? b.items : []).slice(0, 200).forEach(i => { const li = document.createElement('li'); li.textContent = text(i); el.appendChild(li); }); }
        else if (b.type === 'progress') { el = document.createElement('progress'); el.max = 1; el.value = Math.max(0, Math.min(1, +b.value || 0)); if (b.label) el.setAttribute('aria-label', text(b.label)); }
        else if (b.type === 'input') {
          el = document.createElement('label');
          el.textContent = text(b.label);
          const inp = document.createElement('input');
          inp.type = b.inputType === 'number' ? 'number' : 'text';
          inp.value = text(b.value);
          el.appendChild(inp);
          if (b.id) inputs[String(b.id)] = inp;
        } else if (b.type === 'button') {
          el = document.createElement('button');
          el.textContent = text(b.label) || 'OK';
          if (b.primary) el.className = 'primary';
          const id = String(b.id || '');
          el.addEventListener('click', () => {
            const values = Object.fromEntries(Object.entries(inputs).map(([k, v]) => [k, v.value]));
            run(addin.id, 'action', id, values);
          });
        } else continue;
        body.appendChild(el);
      }
      pane.hidden = false;
    }

    // ---- Calls from add-ins
    async function handle(addin, method, args) {
      if (READS.has(method) && !addin.permissions.includes('read')) throw new Error('This add-in may not read the document');
      if (WRITES.has(method) && !addin.permissions.includes('write')) throw new Error('This add-in may not change the document');
      if (UI_CALLS.has(method)) {
        if (method === 'toast') return void UI.toast(`${addin.name}: ${String(args[0]).slice(0, 300)}`, 4000);
        if (method === 'showPane') return void showPane(addin, args[0], args[1]);
        if (method === 'closePane') { if (paneFor === addin.id) pane.hidden = true; return; }
        if (method === 'form') {
          const fields = args[1].slice(0, 20).filter(f => f && typeof f.name === 'string').map(f => ({ name: f.name.slice(0, 40), label: String(f.label || f.name).slice(0, 200), type: ['text', 'number', 'textarea', 'select'].includes(f.type) ? f.type : 'text', value: f.value === undefined ? undefined : String(f.value).slice(0, 10000), options: f.type === 'select' && Array.isArray(f.options) ? f.options.slice(0, 100).map(o => (typeof o === 'object' && o ? { value: String(o.value), label: String(o.label ?? o.value) } : String(o))) : undefined }));
          return UI.form(`${String(args[0]).slice(0, 80)} — ${addin.name}`, fields, args[2]);
        }
      }
      if (!READS.has(method) && !WRITES.has(method)) throw new Error(`Unknown call: ${method}`);
      const fn = adapter[method];
      if (typeof fn !== 'function') throw new Error(`${method} isn’t available in Isadek ${app[0].toUpperCase() + app.slice(1)}`);
      // What comes back is copied as data (structured clone): no functions or page objects.
      return fn(...args);
    }
    function start(addin) {
      const w = new Worker('js/addin-worker.js');
      const r = { worker: w, addin, runs: new Map(), ready: null };
      r.ready = new Promise((resolve, reject) => {
        w.onmessage = async e => {
          const m = e.data || {};
          if (m.t === 'ready') return resolve();
          if (m.t === 'failed') { reject(new Error(m.error)); return; }
          if (m.t === 'call') {
            let value, error;
            try { value = await handle(addin, String(m.method), Array.isArray(m.args) ? m.args : []); } catch (err) { error = String(err && err.message || err); }
            try { w.postMessage({ t: 'result', id: m.id, value, error }); } catch (err) { w.postMessage({ t: 'result', id: m.id, error: 'The result couldn’t be sent: ' + err.message }); }
            return;
          }
          if (m.t === 'done' && r.runs.has(m.run)) { const d = r.runs.get(m.run); r.runs.delete(m.run); clearTimeout(d.timer); m.error ? d.reject(new Error(m.error)) : d.resolve(); }
        };
        w.onerror = e => { e.preventDefault(); reject(new Error(e.message || 'The add-in stopped')); };
      });
      w.postMessage({ t: 'load', code: addin.script, manifest: { id: addin.id, name: addin.name, version: addin.version, app } });
      running.set(addin.id, r);
      r.ready.catch(err => { UI.toast(`${addin.name} couldn’t start: ${err.message}`, 6000); stop(addin.id); });
      return r;
    }
    function stop(id) {
      const r = running.get(id);
      if (!r) return;
      r.worker.terminate();
      r.runs.forEach(d => { clearTimeout(d.timer); d.reject(new Error('The add-in was stopped')); });
      running.delete(id);
      if (paneFor === id) { pane.hidden = true; paneFor = null; }
    }
    let runSeq = 0;
    /** Runs a command (or a pane button) of an add-in, starting it if needed. A minute at most. */
    async function run(id, kind, cmd, values) {
      const addin = load().find(a => a.id === id && a.enabled !== false);
      if (!addin) return;
      const r = running.get(id) || start(addin);
      try {
        await r.ready;
        const runId = ++runSeq;
        await new Promise((resolve, reject) => {
          const timer = setTimeout(() => { reject(new Error('it took too long and was stopped')); stop(id); }, 60000);
          r.runs.set(runId, { resolve, reject, timer });
          r.worker.postMessage({ t: kind, id: cmd, values, run: runId });
        });
      } catch (err) { UI.toast(`${addin.name}: ${err.message}`, 6000); }
    }

    // ---- Ribbon: the Add-ins group on the Insert tab
    const panel = document.querySelector('.ribbon-panel[data-panel="insert"]');
    const group = document.createElement('div');
    group.className = 'group addins-group';
    group.innerHTML = '<div class="controls"><button class="btn big" id="addinsBtn" title="Install, turn on or remove add-ins"><span class="ico" data-icon="puzzle"></span>Add-ins</button></div><div class="label">Add-ins</div>';
    if (panel) panel.appendChild(group);
    const controls = group.querySelector('.controls');
    function renderButtons() {
      controls.querySelectorAll('.addin-cmd').forEach(b => b.remove());
      for (const a of load()) {
        if (a.enabled === false || !a.apps.includes(app)) continue;
        for (const c of a.commands) {
          if (!c.apps.includes(app)) continue;
          const b = document.createElement('button');
          b.className = 'btn big addin-cmd';
          b.dataset.addin = a.id; b.dataset.cmd = c.id;
          b.title = `${c.label} (${a.name})`;
          const ico = document.createElement('span'); ico.className = 'ico'; ico.dataset.icon = window.Icons && Icons.has(c.icon) ? c.icon : 'puzzle';
          b.append(ico, document.createTextNode(c.label));
          b.addEventListener('mousedown', e => e.preventDefault());
          b.addEventListener('click', () => run(a.id, 'command', c.id));
          controls.appendChild(b);
        }
      }
    }
    renderButtons();
    window.addEventListener('storage', e => { if (e.key === KEY) { renderButtons(); for (const id of [...running.keys()]) if (!load().some(a => a.id === id && a.enabled !== false)) stop(id); } });

    async function install(text) {
      let a;
      try { a = check(text); } catch (err) { UI.toast(err.message, 6000); return null; }
      const list = load(), old = list.find(x => x.id === a.id);
      const perms = a.permissions.length ? a.permissions.map(p => '• ' + PERMS[p]).join('\n') : '• Nothing in your documents (it can only show its own pane)';
      const i = await UI.choose(old ? 'Update Add-in' : 'Install Add-in',
        `${a.name} ${a.version}${a.author ? ' by ' + a.author : ''}\n${a.description}\n\nWorks in: ${a.apps.map(x => 'Isadek ' + x[0].toUpperCase() + x.slice(1)).join(', ')}\nIt asks to:\n${perms}\n\nAdd-ins can’t reach the internet or your files. Install add-ins only from people you trust.`,
        [old ? 'Update' : 'Install', 'Cancel'], 'puzzle');
      if (i !== 0) return null;
      stop(a.id);
      const next = list.filter(x => x.id !== a.id);
      next.push({ ...a, enabled: true, installed: Date.now() });
      if (!save(next)) return null;
      renderButtons();
      UI.toast(`${a.name} installed${a.apps.includes(app) ? ' — its buttons are on the Insert tab' : ''}`, 4000);
      return a;
    }
    async function manage() {
      const list = load();
      const lines = list.length ? list.map(a => `${a.enabled === false ? '○' : '●'} ${a.name} ${a.version} — ${a.apps.join(', ')}${a.permissions.length ? ' — may ' + a.permissions.map(p => PERMS[p].toLowerCase()).join(', ') : ''}`).join('\n') : 'No add-ins are installed.';
      const acts = ['Install from File…', ...(list.length ? ['Turn On or Off…', 'Remove…'] : []), 'Close'];
      const i = await UI.choose('Add-ins', lines, acts, 'puzzle');
      const act = acts[i];
      if (act === 'Install from File…') {
        const f = await Platform.open([{ name: 'Isadek add-in', extensions: ['isadek-addin', 'json'] }]);
        if (f) await install(new TextDecoder().decode(f.data));
      } else if (act === 'Turn On or Off…' || act === 'Remove…') {
        const v = await UI.form(act.replace('…', ''), [{ name: 'id', label: 'Add-in', type: 'select', options: list.map(a => ({ value: a.id, label: `${a.name}${a.enabled === false ? ' (off)' : ''}` })) }], act === 'Remove…' ? 'Remove' : 'Switch');
        if (!v) return;
        const a = list.find(x => x.id === v.id);
        if (!a) return;
        stop(a.id);
        if (act === 'Remove…') save(list.filter(x => x.id !== a.id));
        else { a.enabled = a.enabled === false; save(list); }
        renderButtons();
        UI.toast(act === 'Remove…' ? `${a.name} removed` : `${a.name} turned ${a.enabled === false ? 'off' : 'on'}`);
      }
    }
    group.querySelector('#addinsBtn').addEventListener('click', manage);
    return { install, manage, run, stop, get installed() { return load(); }, get running() { return [...running.keys()]; } };
  }

  window.Addins = { attach, check };
})();
