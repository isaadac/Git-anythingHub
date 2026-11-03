'use strict';
// Isadek Sheets — Forms: build a form in the workbook, publish it (to the Isadek Office server
// the web app runs on, a server of your own, or this computer for people on the same network),
// and the answers arrive as rows in a "Form Responses" sheet, ready for charts and PivotTables.
//
// A workbook keeps its forms in book.forms: { uid, def (FormsCore.clean), server ('local' or an
// address), id, key (the owner key: whoever has the file can read the answers), got (responses
// fetched so far), sheet (where they go), cols (question ids, one column each, in order) }.
(function () {
  const C = () => window.FormsCore;
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const { isMine, addMine, SERVER_OK, call, hostUrls, linkOf } = window.FormsClient;
  const uid = () => 'f' + Math.random().toString(36).slice(2, 10);
  // icons: send eye square-pen copy-plus copy refresh-cw trash-2 plus arrow-up arrow-down x clipboard-list

  /** A form from a file: known fields, a clean definition and a server address that can be called. */
  function cleanEntry(f) {
    if (!f || typeof f !== 'object') return null;
    const def = C().clean(f.def);
    if (!def) return null;
    const out = { uid: /^f[a-z0-9]{1,12}$/.test(f.uid) ? f.uid : uid(), def, got: Math.max(0, Math.floor(+f.got) || 0), sheet: typeof f.sheet === 'string' ? f.sheet.slice(0, 31) : '' };
    out.cols = (Array.isArray(f.cols) ? f.cols : []).filter(id => typeof id === 'string' && C().ID.test(id)).slice(0, 500);
    if ((f.server === 'local' || (typeof f.server === 'string' && SERVER_OK.test(f.server))) && /^[A-Za-z0-9_-]{22}$/.test(f.id) && /^[A-Za-z0-9_-]{20,100}$/.test(f.key)) Object.assign(out, { server: f.server, id: f.id, key: f.key });
    return out;
  }
  const cleanList = list => (Array.isArray(list) ? list : []).map(cleanEntry).filter(Boolean).slice(0, 50);

  function SheetsForms(ctx) {
    const F = window.Formula;
    const book = () => ctx.book();
    const list = () => (book().forms = book().forms || []);
    let pane = null, cur = 0, poll = null;

    // ---------- The responses sheet ----------
    const sheetOf = f => (f.sheet && ctx.sheetByName(f.sheet)) || null;
    function ensureSheet(f) {
      let s = sheetOf(f);
      if (!s) {
        s = ctx.newSheet(ctx.uniqueName(f.sheet || 'Form Responses ' + (list().indexOf(f) + 1)));
        book().sheets.push(s);
        f.sheet = s.name;
        s.freezeRow = true;
        s.colWidths = { 0: 140 };
      }
      // A column for every question, in the order they were added; titles follow the form.
      f.def.questions.forEach(q => { if (!f.cols.includes(q.id)) f.cols.push(q.id); });
      ctx.setCell(0, 0, { v: 'Timestamp', s: { b: true } }, s);
      f.cols.forEach((id, i) => {
        const q = f.def.questions.find(x => x.id === id);
        const old = String((s.cells[F.addr(0, i + 1)] || {}).v ?? '');
        if (q) ctx.setCell(0, i + 1, { v: C().toCells({ t: q.title }, ['t'])[0], s: { b: true } }, s);
        else if (!old) ctx.setCell(0, i + 1, { v: '(deleted question)', s: { b: true } }, s);
        if (!(s.colWidths || {})[i + 1]) (s.colWidths = s.colWidths || {})[i + 1] = q && q.type === 'paragraph' ? 260 : 160;
      });
      return s;
    }
    const nextRow = s => { let r = 0; for (const k of Object.keys(s.cells)) { const p = F.parseRef(k); if (p && p.row > r) r = p.row; } return r + 1; };
    const serial = d => F.dateToSerial(d);
    function addRows(f, responses) {
      if (!responses.length) return;
      const s = ensureSheet(f);
      let r = nextRow(s);
      for (const resp of responses) {
        const at = new Date(resp.at);
        ctx.setCell(r, 0, { v: isNaN(at) ? '' : String(serial(at)), s: { fmt: 'yyyy-mm-dd hh:mm' } }, s);
        const cells = C().toCells(resp.values, f.cols);
        f.cols.forEach((id, i) => {
          const q = f.def.questions.find(x => x.id === id);
          let v = cells[i], st = null;
          const raw = resp.values ? resp.values[id] : undefined;
          if (q && q.type === 'date' && typeof raw === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(raw)) { const [y, m, d] = raw.split('-').map(Number); v = String(Math.floor(serial(new Date(y, m - 1, d)))); st = { fmt: 'yyyy-mm-dd' }; }
          else if (typeof v === 'number') v = String(v);
          if (v !== '') ctx.setCell(r, i + 1, st ? { v, s: st } : { v }, s);
        });
        r++;
      }
    }

    // ---------- Getting responses ----------
    async function refresh(f, { quiet = false } = {}) {
      if (!f.id) return 0;
      if (!isMine(f)) {
        if (quiet) return 0;
        const where = f.server === 'local' ? 'this computer' : new URL(f.server).host;
        if (!(await UI.confirm('Get Form Responses', `This workbook collects answers to “${f.def.title}” at ${where}. Get the responses from there?`))) return 0;
        addMine(f);
      }
      let n = 0, data;
      do {
        data = await call(f, 'responses', { after: f.got, limit: 1000 });
        const fresh = (data.responses || []).filter(x => x && Number.isInteger(x.n) && x.n > f.got).sort((a, b) => a.n - b.n);
        if (!fresh.length) break;
        if (!n && !quiet) ctx.snapshot();
        addRows(f, fresh);
        f.got = fresh[fresh.length - 1].n;
        n += fresh.length;
      } while ((data.responses || []).length >= 1000);
      f.open = data.open;
      if (n) { ctx.state.markDirty(); ctx.buildGrid(); }
      if (pane && !pane.hidden) draw();
      return n;
    }
    async function refreshAll(quiet) {
      let n = 0;
      for (const f of list()) { try { n += await refresh(f, { quiet }); } catch (err) { if (!quiet) UI.toast(`“${f.def.title}”: ${err.message}`, 4000); } }
      return n;
    }
    // Forms you published keep coming in while the workbook is open.
    function startPolling() {
      if (poll) return;
      poll = setInterval(async () => {
        if (document.hidden || !list().some(f => f.id && isMine(f))) return;
        const n = await refreshAll(true);
        if (n) UI.toast(`${n} new form response${n === 1 ? '' : 's'}`);
      }, 60000);
    }
    async function opened() {
      book().forms = cleanList(book().forms);
      if (list().some(f => f.server === 'local' && isMine(f)) && window.atlas && window.atlas.formsHost) hostUrls().catch(() => {});
      if (list().some(f => f.id && isMine(f))) { startPolling(); refreshAll(true).then(n => { if (n) UI.toast(`${n} new form response${n === 1 ? '' : 's'}`); }); }
    }

    // ---------- Publishing ----------
    async function publish(f) {
      if (!f.def.questions.length) return UI.toast('Add a question first');
      try {
        if (f.id) {
          const r = await call(f, 'update', { form: f.def });
          f.open = r.open;
          ensureSheet(f); ctx.buildGrid();
          UI.toast('Form updated');
        } else {
          const server = await FormsClient.chooseServer('Publish Form', 'answer');
          if (!server) return;
          f.server = server;
          const r = await call(f, 'create', { form: f.def });
          Object.assign(f, { id: r.id, key: r.key, got: 0, open: true });
          addMine(f);
          ctx.snapshot();
          ensureSheet(f);
          ctx.state.markDirty();
          ctx.buildGrid();
          startPolling();
          UI.toast('Published. Share the link to collect answers.', 3500);
        }
      } catch (err) {
        if (!f.id) delete f.server;
        UI.toast('Couldn’t publish: ' + err.message, 5000);
      }
      draw();
    }
    async function setOpen(f, open) {
      try { const r = await call(f, 'update', { open }); f.open = r.open; UI.toast(open ? 'Accepting responses' : 'No longer accepting responses'); } catch (err) { UI.toast(err.message, 4000); }
      draw();
    }
    async function remove(f) {
      const ok = await UI.confirm('Delete Form', f.id ? `Delete “${f.def.title}”? Its link stops working and responses not yet in the workbook are lost. Responses already in “${f.sheet}” stay.` : `Delete “${f.def.title}”?`);
      if (!ok) return;
      if (f.id) { try { await call(f, 'remove'); } catch (err) { if (err.status !== 404 && !(await UI.confirm('Delete Form', `${err.message}. Remove it from this workbook anyway?`))) return; } }
      ctx.snapshot();
      list().splice(list().indexOf(f), 1);
      cur = 0;
      ctx.state.markDirty();
      draw();
    }

    // ---------- Answering here: a preview, or a response typed in on this device ----------
    function answerDialog(f, preview) {
      const back = document.createElement('div');
      back.className = 'modal-backdrop';
      const modal = document.createElement('div');
      modal.className = 'modal form-modal';
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-label', preview ? 'Form preview' : 'Fill in the form');
      const bar = document.createElement('div');
      bar.className = 'form-modal-bar';
      bar.innerHTML = `<b>${preview ? 'Preview' : 'Fill In'}</b><span class="sp"></span><button class="x" title="Close"><span class="ico" data-icon="x"></span></button>`;
      const box = document.createElement('div');
      modal.append(bar, box);
      back.appendChild(modal);
      document.body.appendChild(back);
      const close = () => back.remove();
      bar.querySelector('.x').addEventListener('click', close);
      back.addEventListener('mousedown', e => { if (e.target === back) close(); });
      modal.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Escape') close(); });
      FormsView.render(box, f.def, {
        submitLabel: preview ? 'Submit (preview)' : 'Submit',
        onSubmit: async values => {
          if (preview) return;
          if (f.id) { await call(f, 'respond', { answers: values }); await refresh(f, { quiet: true }); }
          else { ctx.snapshot(); addRows(f, [{ n: 0, at: new Date().toISOString(), values }]); ctx.state.markDirty(); ctx.buildGrid(); }
        }
      });
      const first = box.querySelector('input, textarea, select');
      if (first) first.focus();
    }

    // ---------- Summary: counts for choices and ratings, averages for numbers ----------
    function summary(f) {
      const s = sheetOf(f);
      if (!s) return '<p class="hint">No responses yet.</p>';
      const rows = nextRow(s) - 1;
      const val = (r, c) => { const v = ctx.engine.value(s.name, r, c); return v === undefined || v === null ? '' : v; };
      let html = `<p class="fm-count"><b>${rows}</b> response${rows === 1 ? '' : 's'}</p>`;
      f.def.questions.forEach(q => {
        const c = f.cols.indexOf(q.id) + 1;
        if (c < 1) return;
        const vals = []; for (let r = 1; r <= rows; r++) { const v = val(r, c); if (v !== '') vals.push(v); }
        if (['choice', 'dropdown', 'checkboxes', 'rating'].includes(q.type)) {
          const opts = q.type === 'rating' ? Array.from({ length: q.max }, (_, i) => String(i + 1)) : q.options;
          const counts = opts.map(o => vals.filter(v => (q.type === 'checkboxes' ? String(v).split(', ').includes(o) : String(v) === o)).length);
          const max = Math.max(1, ...counts);
          html += `<div class="fm-sum"><b>${esc(q.title)}</b>${opts.map((o, i) => `<div class="fm-bar"><span class="fm-lbl" title="${esc(o)}">${esc(o)}</span><span class="fm-track"><i style="width:${(100 * counts[i] / max).toFixed(1)}%"></i></span><span class="fm-n">${counts[i]}</span></div>`).join('')}</div>`;
        } else if (q.type === 'number') {
          const nums = vals.map(Number).filter(isFinite);
          if (nums.length) html += `<div class="fm-sum"><b>${esc(q.title)}</b><p class="hint">Average ${F.formatValue(nums.reduce((a, b) => a + b, 0) / nums.length, '#,##0.##')} · lowest ${nums.reduce((a, b) => Math.min(a, b))} · highest ${nums.reduce((a, b) => Math.max(a, b))}</p></div>`;
        } else html += `<div class="fm-sum"><b>${esc(q.title)}</b><p class="hint">${vals.length} answer${vals.length === 1 ? '' : 's'}</p></div>`;
      });
      return html;
    }

    // ---------- The Forms pane ----------
    function newForm() {
      const f = { uid: uid(), def: C().clean({ title: 'Untitled form', questions: [{ id: 'q1', type: 'text', title: 'Name', required: true }] }), got: 0, sheet: '', cols: [] };
      ctx.snapshot();
      list().push(f);
      cur = list().length - 1;
      ctx.state.markDirty();
      formsPane(true);
    }
    // Edits change the definition through the same checks as everything else.
    function edit(f, fn) { const d = JSON.parse(JSON.stringify(f.def)); fn(d); f.def = C().clean(d); f.dirty = !!f.id; ctx.state.markDirty(); }
    const qid = f => { let i = f.def.questions.length + 1; while (f.def.questions.some(q => q.id === 'q' + i) || f.cols.includes('q' + i)) i++; return 'q' + i; };

    function formsPane(show = !pane || pane.hidden) {
      if (!pane) {
        pane = document.createElement('aside');
        pane.className = 'script-pane form-pane';
        pane.setAttribute('aria-label', 'Forms');
        document.body.appendChild(pane);
        pane.addEventListener('keydown', e => e.stopPropagation());
        pane.addEventListener('input', onInput);
        pane.addEventListener('change', onInput);
        pane.addEventListener('click', onClick);
      }
      pane.hidden = !show;
      document.body.classList.toggle('script-open', show);
      if (show) { if (!list().length) newForm(); else draw(); }
    }
    async function draw() {
      if (!pane || pane.hidden) return;
      cur = Math.min(cur, Math.max(0, list().length - 1));
      const f = list()[cur];
      if (!f) { pane.innerHTML = ''; return; }
      const TYPES = C().TYPES;
      const link = f.id ? await linkOf(f).catch(() => '') : '';
      const scrollTop = pane.querySelector('.fm-scroll') ? pane.querySelector('.fm-scroll').scrollTop : 0;
      pane.innerHTML = `<header><b>Forms</b><span class="sp"></span><button class="x" data-a="close" title="Close"><span class="ico" data-icon="x"></span></button></header>
        <div class="row"><select data-k="pick" aria-label="Form">${list().map((x, i) => `<option value="${i}"${i === cur ? ' selected' : ''}>${esc(x.def.title)}</option>`).join('')}</select><button data-a="new" title="New form"><span class="ico" data-icon="plus"></span></button><button data-a="delete" title="Delete form"><span class="ico" data-icon="trash-2"></span></button></div>
        <div class="fm-scroll">
          <section class="fm-status">${f.id ? `
            <p><span class="fm-dot${f.open === false ? ' off' : ''}"></span>${f.open === false ? 'Not accepting responses' : 'Accepting responses'}${f.server === 'local' ? ' on this computer’s network' : ''}</p>
            <div class="row"><input readonly class="fm-link" value="${esc(link)}" aria-label="Link to the form"><button data-a="copy" title="Copy the link"><span class="ico" data-icon="copy"></span></button></div>
            <div class="row"><button data-a="refresh"><span class="ico" data-icon="refresh-cw"></span>Get Responses</button><button data-a="toggle">${f.open === false ? 'Accept Responses' : 'Stop Accepting'}</button></div>` : '<p class="hint">Not published yet: only you can fill it in, here.</p>'}
          </section>
          <label class="fm-field">Title<input data-k="title" maxlength="300"></label>
          <label class="fm-field">Description<textarea data-k="description" rows="2" maxlength="5000"></textarea></label>
          <div class="row fm-look"><label>Colour <input type="color" data-k="color" value="${f.def.color || '#13834a'}"></label><label class="sp">After submitting <input data-k="confirm" placeholder="Your response has been recorded." maxlength="2000"></label></div>
          <div class="fm-qs">${f.def.questions.map((q, i) => `
            <div class="fm-q" data-i="${i}">
              <div class="row"><input data-k="q.title" aria-label="Question ${i + 1}" maxlength="300"><select data-k="q.type" aria-label="Question type">${Object.entries(TYPES).map(([k, l]) => `<option value="${k}"${k === q.type ? ' selected' : ''}>${l}</option>`).join('')}</select></div>
              ${C().WITH_OPTIONS.has(q.type) ? `<textarea data-k="q.options" rows="${Math.min(8, q.options.length + 1)}" aria-label="Options, one per line" placeholder="One option per line"></textarea>` : ''}
              ${q.type === 'number' ? `<div class="row"><label>Min <input type="number" data-k="q.min" value="${q.min ?? ''}"></label><label>Max <input type="number" data-k="q.max" value="${q.max ?? ''}"></label><label><input type="checkbox" data-k="q.whole"${q.whole ? ' checked' : ''}> Whole</label></div>` : ''}
              ${q.type === 'rating' ? `<div class="row"><label>Highest <select data-k="q.max">${[3, 4, 5, 6, 7, 8, 9, 10].map(n => `<option${n === q.max ? ' selected' : ''}>${n}</option>`).join('')}</select></label></div>` : ''}
              <div class="row"><label><input type="checkbox" data-k="q.required"${q.required ? ' checked' : ''}> Required</label><span class="sp"></span>
                <button data-a="up" title="Move up"${i ? '' : ' disabled'}><span class="ico" data-icon="arrow-up"></span></button><button data-a="down" title="Move down"${i < f.def.questions.length - 1 ? '' : ' disabled'}><span class="ico" data-icon="arrow-down"></span></button><button data-a="dup" title="Duplicate"><span class="ico" data-icon="copy-plus"></span></button><button data-a="del" title="Delete question"><span class="ico" data-icon="trash-2"></span></button></div>
            </div>`).join('')}</div>
          <button data-a="add" class="fm-add"><span class="ico" data-icon="plus"></span>Add question</button>
          ${f.id || sheetOf(f) ? `<section class="fm-summary"><h3>Summary</h3>${summary(f)}</section>` : ''}
        </div>
        <div class="row fm-foot"><button data-a="preview"><span class="ico" data-icon="eye"></span>Preview</button><button data-a="fill" title="Type in a response on this device"><span class="ico" data-icon="square-pen"></span>Fill In</button><span class="sp"></span><button class="primary" data-a="publish">${f.id ? (f.dirty ? 'Update Form' : 'Updated') : 'Publish'}</button></div>`;
      // Values from the form are set as properties, never as markup.
      pane.querySelector('[data-k="title"]').value = f.def.title;
      pane.querySelector('[data-k="description"]').value = f.def.description;
      pane.querySelector('[data-k="confirm"]').value = f.def.confirm || '';
      pane.querySelectorAll('.fm-q').forEach(box => {
        const q = f.def.questions[+box.dataset.i];
        box.querySelector('[data-k="q.title"]').value = q.title;
        const o = box.querySelector('[data-k="q.options"]');
        if (o) o.value = q.options.join('\n');
      });
      const pub = pane.querySelector('[data-a="publish"]');
      if (f.id && !f.dirty) pub.disabled = true;
      pane.querySelector('.fm-scroll').scrollTop = scrollTop;
    }
    function onInput(e) {
      const k = e.target.dataset && e.target.dataset.k;
      const f = list()[cur];
      if (!k || !f) return;
      if (k === 'pick') { if (e.type === 'change') { cur = +e.target.value; draw(); } return; }
      const v = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
      const box = e.target.closest('.fm-q');
      const i = box ? +box.dataset.i : -1;
      // Typing updates the form quietly; choices that change the layout redraw it.
      if (k.startsWith('q.') && i >= 0) {
        const field = k.slice(2);
        edit(f, d => {
          const q = d.questions[i];
          if (field === 'options') q.options = v.split('\n');
          else if (field === 'min' || field === 'max') q[field] = v === '' ? undefined : +v;
          else q[field] = v;
          if (field === 'type' && C().WITH_OPTIONS.has(v) && !q.options) q.options = ['Option 1', 'Option 2'];
        });
        if (e.type === 'change' && field === 'type') draw();
      } else if (['title', 'description', 'color', 'confirm'].includes(k)) {
        edit(f, d => { d[k] = v; });
        // (Not a redraw: leaving the box by clicking a button must not lose that click.)
        if (k === 'title') { const o = pane.querySelector(`[data-k="pick"] option[value="${cur}"]`); if (o) o.textContent = f.def.title; }
      }
      const pub = pane.querySelector('[data-a="publish"]');
      if (pub && f.id) { pub.disabled = false; pub.textContent = 'Update Form'; }
    }
    async function onClick(e) {
      const b = e.target.closest('button[data-a]');
      const f = list()[cur];
      if (!b) return;
      const a = b.dataset.a;
      if (a === 'close') return formsPane(false);
      if (a === 'new') return newForm();
      if (!f) return;
      const box = b.closest('.fm-q');
      const i = box ? +box.dataset.i : -1;
      if (a === 'add') { edit(f, d => d.questions.push({ id: qid(f), type: 'choice', title: 'Question', options: ['Option 1', 'Option 2'] })); draw().then(() => { const qs = pane.querySelectorAll('.fm-q [data-k="q.title"]'); const last = qs[qs.length - 1]; if (last) { last.focus(); last.select(); } }); return; }
      if (a === 'del') { edit(f, d => d.questions.splice(i, 1)); return draw(); }
      if (a === 'dup') { const id = qid(f); edit(f, d => d.questions.splice(i + 1, 0, { ...JSON.parse(JSON.stringify(d.questions[i])), id })); return draw(); }
      if (a === 'up' || a === 'down') { const j = a === 'up' ? i - 1 : i + 1; edit(f, d => { const [q] = d.questions.splice(i, 1); d.questions.splice(j, 0, q); }); return draw(); }
      if (a === 'publish') { await publish(f); f.dirty = false; return draw(); }
      if (a === 'preview') return answerDialog(f, true);
      if (a === 'fill') return answerDialog(f, false);
      if (a === 'delete') return remove(f);
      if (a === 'toggle') return setOpen(f, f.open === false);
      if (a === 'copy') { const link = await linkOf(f); try { await navigator.clipboard.writeText(link); UI.toast('Link copied'); } catch { pane.querySelector('.fm-link').select(); } return; }
      if (a === 'refresh') { try { const n = await refresh(f); UI.toast(n ? `${n} new response${n === 1 ? '' : 's'}` : 'No new responses'); } catch (err) { UI.toast(err.message, 4000); } }
    }

    return { formsPane, newForm, refresh, refreshAll, opened, cleanList, publish, answerDialog, get current() { return list()[cur]; } };
  }

  SheetsForms.cleanList = cleanList;
  window.SheetsForms = SheetsForms;
})();
