'use strict';
// Isadek Slides — the audience and the stage: live polls on slides, audience Q&A, a phone remote
// for the show, and Zoom (a slide that jumps to another and comes back). Polls and questions are
// Isadek Forms (forms-client.js): people answer from a link or a QR code, on any device, and the
// results update on the slide while it is shown.
//
//   poll element: { type: 'poll', question, options: [..], form: { server, id, key } }
//   zoom element: { type: 'zoom', target: slideId }
//   pres.qa:      { server, id, key } (the Q&A form), when audience Q&A is on
(function () {
  const FC = () => window.FormsClient;
  const ID_OK = /^[\w-]{1,60}$/;
  const str = (v, n) => (typeof v === 'string' ? v.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').slice(0, n) : '');
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; };
  // icons: chart-bar-big message-circle-question smartphone zoom-in qr-code send x check copy

  /** A QR code as an SVG element (dark modules on white, with the quiet zone). */
  function qrSvg(text, px) {
    if (typeof qrcode !== 'function') return el('span');
    const q = qrcode(0, 'M');
    q.addData(text);
    q.make();
    const n = q.getModuleCount(), m = n + 8;
    let d = '';
    for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) if (q.isDark(r, c)) d += `M${c + 4},${r + 4}h1v1h-1z`;
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('viewBox', `0 0 ${m} ${m}`);
    svg.setAttribute('class', 'qr');
    svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', 'QR code for ' + text);
    if (px) { svg.setAttribute('width', px); svg.setAttribute('height', px); }
    svg.setAttribute('shape-rendering', 'crispEdges');
    const bg = document.createElementNS(svg.namespaceURI, 'rect');
    bg.setAttribute('width', m); bg.setAttribute('height', m); bg.setAttribute('fill', '#fff');
    const path = document.createElementNS(svg.namespaceURI, 'path');
    path.setAttribute('d', d); path.setAttribute('fill', '#000');
    svg.append(bg, path);
    return svg;
  }

  /** Cleans a poll or zoom element's own fields onto `out` (the common ones are done already). */
  function clean(type, src, out) {
    if (type === 'poll') {
      out.question = str(src.question, 300) || 'Question';
      const seen = new Set();
      out.options = (Array.isArray(src.options) ? src.options : []).map(o => str(o, 120).trim()).filter(o => o && !seen.has(o) && seen.add(o)).slice(0, 10);
      if (out.options.length < 2) out.options = ['Yes', 'No'];
      const f = FC() && FC().cleanRef(src.form);
      if (f) out.form = f;
      return out;
    }
    if (type === 'zoom') {
      if (typeof src.target === 'string' && ID_OK.test(src.target)) out.target = src.target;
      return out;
    }
    return null;
  }
  const pollForm = p => ({ title: p.question, questions: [{ id: 'a', type: 'choice', title: p.question, options: p.options, required: true }] });
  const QA_FORM = { title: 'Questions', description: 'Ask the speaker a question.', confirm: 'Thanks! Your question has been sent.', questions: [{ id: 'q', type: 'paragraph', title: 'Your question', required: true }, { id: 'name', type: 'text', title: 'Your name (optional)' }] };

  function SlidesLive(ctx) {
    const links = new Map();   // form id -> link to answer at
    const results = new Map(); // form id -> { got, counts: Map(option -> n), total, timer }
    let qa = null;             // { got, items: [{ n, q, name, at }], done: Set, timer, panel }
    let remote = null;         // { ref, token, link, after, alive, stateTimer }

    async function linkFor(ref, path = 'f') {
      const k = path + ref.id;
      if (!links.has(k)) links.set(k, (await FC().base(ref.server)) + `/${path}/` + ref.id);
      return links.get(k);
    }
    const hostOf = link => { try { const u = new URL(link); return u.host + u.pathname; } catch { return link; } };

    // ---------- Polls ----------
    /** The DOM inside a poll's box: the question, a bar per option and how to answer. */
    function buildPoll(e, div) {
      div.style.setProperty('--pk', String(Math.max(0.4, Math.min(3, e.w / 640))));
      const card = el('div', 'poll-card');
      card.appendChild(el('div', 'poll-q', e.question));
      const res = e.form && results.get(e.form.id);
      const total = res ? res.total : 0;
      const list = el('div', 'poll-opts');
      e.options.forEach(o => {
        const n = res ? res.counts.get(o) || 0 : 0;
        const row = el('div', 'poll-row');
        row.dataset.option = o;
        const pct = total ? Math.round(100 * n / total) : 0;
        const bar = el('span', 'poll-bar'); const fill = el('i'); fill.style.width = pct + '%'; bar.appendChild(fill);
        row.append(el('span', 'poll-label', o), bar, el('span', 'poll-n', total ? `${pct}%` : ''));
        list.appendChild(row);
      });
      card.appendChild(list);
      const foot = el('div', 'poll-foot');
      if (e.form) {
        const where = el('span', 'poll-link', '');
        const qr = el('span', 'poll-qr');
        linkFor(e.form).then(l => { where.textContent = hostOf(l); qr.replaceChildren(qrSvg(l)); }).catch(() => { where.textContent = 'Not reachable'; });
        foot.append(qr, el('span', 'poll-join', 'Answer at '), where, el('span', 'poll-total', total ? ` · ${total} vote${total === 1 ? '' : 's'}` : ''));
      } else foot.append(el('span', 'poll-join', 'Not published: double-click to publish'));
      card.appendChild(foot);
      div.appendChild(card);
    }
    function redrawPolls(formId) {
      for (const node of document.querySelectorAll(`.el.poll[data-form="${formId}"]`)) {
        const s = ctx.pres().slides.find(x => x.elements.some(e => e.id === node.dataset.id));
        const e = s && s.elements.find(x => x.id === node.dataset.id);
        if (!e) continue;
        node.querySelector('.poll-card')?.remove();
        buildPoll(e, node);
      }
    }
    async function fetchPoll(ref) {
      const r = results.get(ref.id) || { got: 0, counts: new Map(), total: 0 };
      results.set(ref.id, r);
      const data = await FC().call(ref, 'responses', { after: r.got, limit: 5000 });
      const fresh = (data.responses || []).filter(x => x && x.n > r.got);
      fresh.forEach(x => { const v = x.values && x.values.a; if (typeof v === 'string') { r.counts.set(v, (r.counts.get(v) || 0) + 1); r.total++; } });
      if (fresh.length) { r.got = fresh[fresh.length - 1].n; redrawPolls(ref.id); }
    }
    function watchPolls(s) {
      const want = new Set(s ? s.elements.filter(e => e.type === 'poll' && e.form).map(e => e.form.id) : []);
      for (const [id, r] of results) if (r.timer && !want.has(id)) { clearInterval(r.timer); r.timer = null; }
      if (!s) return;
      // Results come from polls published here; a deck from someone else asks first (on double-click).
      s.elements.filter(e => e.type === 'poll' && e.form && FC().isMine(e.form)).forEach(e => {
        const r = results.get(e.form.id) || { got: 0, counts: new Map(), total: 0 };
        results.set(e.form.id, r);
        if (r.timer) return;
        const tick = () => fetchPoll(e.form).catch(() => {});
        tick();
        r.timer = setInterval(tick, 2000);
      });
    }
    async function publishPoll(e) {
      const server = await FC().chooseServer('Publish Poll', 'answer');
      if (!server) return false;
      try {
        const r = await FC().call({ server }, 'create', { form: pollForm(e) });
        ctx.snapshot();
        e.form = { server, id: r.id, key: r.key };
        FC().addMine(e.form);
        ctx.changed();
        return true;
      } catch (err) { UI.toast('Couldn’t publish the poll: ' + err.message, 5000); return false; }
    }
    async function pollDialog(e, isNew) {
      const v = await UI.form(isNew ? 'Insert Poll' : 'Poll', [
        { name: 'question', label: 'Question', value: e.question },
        { name: 'options', label: 'Answers (one per line, 2 to 10)', type: 'textarea', value: e.options.join('\n') }
      ], isNew ? 'Insert' : 'Apply');
      if (!v) return false;
      const next = clean('poll', { question: v.question, options: v.options.split('\n'), form: e.form }, { ...e });
      if (!isNew) ctx.snapshot();
      Object.assign(e, next);
      if (e.form) {
        try { await FC().call(e.form, 'update', { form: pollForm(e) }); } catch (err) { UI.toast('The poll changed here, but not online: ' + err.message, 5000); }
        results.delete(e.form.id);
      }
      return true;
    }
    async function insertPoll() {
      const e = { id: ctx.newId(), type: 'poll', question: 'What do you think?', options: ['Great', 'Good', 'Not sure'], x: Math.round(ctx.W * 0.15), y: Math.round(ctx.H * 0.18), w: Math.round(ctx.W * 0.7), h: Math.round(ctx.H * 0.66) };
      if (!(await pollDialog(e, true))) return;
      ctx.snapshot();
      ctx.slide().elements.push(e);
      ctx.select(e.id);
      ctx.changed();
      if (await publishPoll(e)) UI.toast('Poll published. Results appear on the slide as people answer.', 4000);
    }
    async function editPoll(e) {
      if (e.form && !FC().isMine(e.form)) {
        const where = e.form.server === 'local' ? 'the computer that made it' : new URL(e.form.server).host;
        const i = await UI.choose('Poll', `This poll collects answers at ${where}. Show its results here during the show?`, ['Show Results', 'Edit Poll', 'Cancel'], 'chart-bar-big');
        if (i === 0) { FC().addMine(e.form); return UI.toast('Results will show during the show'); }
        if (i !== 1) return;
      }
      if (!(await pollDialog(e, false))) return;
      ctx.changed();
      if (!e.form) await publishPoll(e);
    }
    async function clearPoll(e) {
      if (!e.form || !(await UI.confirm('Clear Results', 'Start this poll again with no votes? The old poll’s link stops working.'))) return;
      const old = e.form;
      try { await FC().call(old, 'remove'); } catch { /* gone already */ }
      results.delete(old.id);
      ctx.snapshot();
      delete e.form;
      ctx.changed();
      await publishPoll(e);
    }

    // ---------- Audience Q&A ----------
    async function toggleQa() {
      const p = ctx.pres();
      if (p.qa) {
        const i = await UI.choose('Audience Q&A', 'People can send questions during the show from the link or QR code (press Q in the show to see them).', ['Show Link', 'Turn Off', 'Close'], 'message-circle-question');
        if (i === 0) return shareDialog('Audience Q&A', await linkFor(p.qa), 'People open this to send questions.');
        if (i !== 1) return;
        try { await FC().call(p.qa, 'remove'); } catch { /* gone already */ }
        ctx.snapshot(); delete p.qa; qaStop(); ctx.changed(); UI.toast('Audience Q&A is off');
        return;
      }
      const server = await FC().chooseServer('Audience Q&A', 'answer');
      if (!server) return;
      try {
        const r = await FC().call({ server }, 'create', { form: QA_FORM });
        ctx.snapshot();
        p.qa = { server, id: r.id, key: r.key };
        FC().addMine(p.qa);
        ctx.changed();
        await shareDialog('Audience Q&A is on', await linkFor(p.qa), 'During the show, the link and QR code appear in a corner; press Q to see the questions.');
      } catch (err) { UI.toast('Couldn’t start Q&A: ' + err.message, 5000); }
    }
    function qaStart(container) {
      const p = ctx.pres();
      if (!p.qa || !FC().isMine(p.qa)) return;
      if (!qa) qa = { got: 0, items: [], done: new Set(), timer: null, panel: null, pill: null };
      const host = container.parentElement;
      if (!qa.pill || !qa.pill.isConnected) {
        qa.pill = el('div', 'qa-pill');
        const qr = el('span', 'qa-qr');
        const t = el('span', 'qa-text');
        t.append(el('b', '', 'Questions? '), el('span', 'qa-link', ''));
        qa.pill.append(qr, t, el('span', 'qa-count', ''));
        linkFor(p.qa).then(l => { qr.replaceChildren(qrSvg(l)); t.querySelector('.qa-link').textContent = hostOf(l); }).catch(() => {});
        host.appendChild(qa.pill);
      }
      if (!qa.timer) {
        const tick = async () => {
          try {
            const data = await FC().call(p.qa, 'responses', { after: qa.got, limit: 1000 });
            const fresh = (data.responses || []).filter(x => x && x.n > qa.got);
            fresh.forEach(x => qa.items.push({ n: x.n, q: String(x.values.q || ''), name: String(x.values.name || ''), at: x.at }));
            if (fresh.length) { qa.got = fresh[fresh.length - 1].n; drawQa(); }
          } catch { /* next time */ }
        };
        tick();
        qa.timer = setInterval(tick, 3000);
      }
      drawQa();
    }
    function drawQa() {
      if (!qa) return;
      const open = qa.items.filter(x => !qa.done.has(x.n));
      if (qa.pill) qa.pill.querySelector('.qa-count').textContent = open.length ? `${open.length} question${open.length === 1 ? '' : 's'} · Q` : '';
      if (!qa.panel || qa.panel.hidden) return;
      const list = qa.panel.querySelector('.qa-list');
      list.textContent = '';
      if (!open.length) list.appendChild(el('p', 'qa-empty', 'No questions yet.'));
      [...open].reverse().forEach(x => {
        const item = el('div', 'qa-item');
        item.append(el('p', 'qa-q', x.q), el('span', 'qa-who', x.name || 'Anonymous'));
        const b = el('button', 'qa-done', '✓'); b.title = 'Answered'; b.setAttribute('aria-label', 'Mark as answered');
        b.addEventListener('click', ev => { ev.stopPropagation(); qa.done.add(x.n); drawQa(); });
        item.appendChild(b);
        list.appendChild(item);
      });
    }
    function qaPanel(container) {
      if (!ctx.pres().qa) return UI.toast('Turn on Audience Q&A (Slide Show tab) first');
      if (!qa) qaStart(container);
      if (!qa) return;
      if (!qa.panel || !qa.panel.isConnected) {
        qa.panel = el('aside', 'qa-panel');
        qa.panel.setAttribute('aria-label', 'Audience questions');
        qa.panel.append(el('h2', '', 'Audience questions'), el('div', 'qa-list'));
        qa.panel.addEventListener('click', ev => ev.stopPropagation());
        qa.panel.addEventListener('contextmenu', ev => ev.stopPropagation());
        container.parentElement.appendChild(qa.panel);
        qa.panel.hidden = true;
      }
      qa.panel.hidden = !qa.panel.hidden;
      drawQa();
    }
    function qaStop() {
      if (!qa) return;
      clearInterval(qa.timer); qa.timer = null;
      qa.pill?.remove(); qa.panel?.remove();
      qa.pill = qa.panel = null;
    }

    // ---------- Phone remote ----------
    const plain = html => { const d = document.createElement('div'); d.innerHTML = Platform.sanitizeHtml(html || ''); return d.textContent.replace(/\s+/g, ' ').trim(); };
    const titleOf = s => { if (!s) return ''; const t = s.elements.find(e => e.ph === 'title') || s.elements.find(e => e.type === 'text'); return t ? plain(t.html).slice(0, 200) : ''; };
    function remoteState() {
      if (!remote) return;
      clearTimeout(remote.stateTimer);
      // Many moves in a row (builds, holding a key) send one update.
      remote.stateTimer = setTimeout(() => {
        const sh = ctx.show(), slides = ctx.pres().slides;
        const s = slides[sh.index];
        FC().remote(remote.ref, 'state', { state: { slide: sh.index, total: slides.length, title: titleOf(s), next: titleOf(slides[sh.index + 1]), notes: s ? s.notes || '' : '', black: sh.black, showing: sh.showing, started: sh.started } }).catch(() => {});
      }, 80);
    }
    async function remoteLoop() {
      const me = remote;
      while (remote === me && me.alive) {
        try {
          const r = await FC().remote(me.ref, 'commands', { after: me.after });
          if (remote !== me) return;
          for (const c of r.commands || []) if (c.n > me.after) { me.after = c.n; if (Date.now() - c.at < 30000) ctx.command(c.cmd); }
          me.after = Math.max(me.after, r.n || 0);
        } catch (err) {
          if (err.status === 404 || err.status === 403) { stopRemote(true); return; }
          await new Promise(res => setTimeout(res, 2000));
        }
      }
    }
    async function startRemote() {
      if (remote) {
        const i = await UI.choose('Phone Remote', 'Your phone is controlling this show.', ['Show Link', 'Stop Remote', 'Close'], 'smartphone');
        if (i === 0) await shareDialog('Phone Remote', remote.link, 'Open this on your phone to move the slides, see your notes and the time.');
        if (i === 1) stopRemote();
        return;
      }
      const server = await FC().chooseServer('Phone Remote', 'remote');
      if (!server) return;
      try {
        const r = await FC().remote({ server }, 'create');
        const ref = { server, id: r.id, key: r.key };
        remote = { ref, link: (await linkFor(ref, 'r')) + '#' + r.token, after: 0, alive: true };
        remoteState();
        remoteLoop();
        document.body.classList.add('remote-on');
        await shareDialog('Phone Remote', remote.link, 'Scan with your phone (on the same network) to move the slides, see your notes and the time. The link controls the show: keep it to yourself.');
      } catch (err) { UI.toast('Couldn’t start the remote: ' + err.message, 5000); }
    }
    function stopRemote(quiet) {
      if (!remote) return;
      const r = remote;
      remote = null;
      r.alive = false;
      document.body.classList.remove('remote-on');
      FC().remote(r.ref, 'remove').catch(() => {});
      if (!quiet) UI.toast('Phone remote stopped');
    }
    window.addEventListener('pagehide', () => { if (remote) stopRemote(true); });

    /** A link to share: QR code, the address, and Copy. */
    function shareDialog(title, link, note) {
      return new Promise(resolve => {
        const back = el('div', 'modal-backdrop');
        const m = el('div', 'modal share-qr');
        m.setAttribute('role', 'dialog');
        m.setAttribute('aria-label', title);
        const input = el('input'); input.readOnly = true; input.value = link; input.setAttribute('aria-label', 'Link');
        const copy = el('button', '', 'Copy Link');
        const close = el('button', 'primary', 'Done');
        const actions = el('div', 'actions'); actions.append(copy, close);
        m.append(el('h2', '', title), qrSvg(link, 240), el('p', '', note), input, actions);
        back.appendChild(m);
        document.body.appendChild(back);
        const done = () => { back.remove(); resolve(); };
        copy.addEventListener('click', async () => { try { await navigator.clipboard.writeText(link); UI.toast('Link copied'); } catch { input.select(); } });
        close.addEventListener('click', done);
        back.addEventListener('mousedown', ev => { if (ev.target === back) done(); });
        m.addEventListener('keydown', ev => { ev.stopPropagation(); if (ev.key === 'Escape' || ev.key === 'Enter') done(); });
        close.focus();
      });
    }

    // ---------- Zoom ----------
    function buildZoom(e, div) {
      const slides = ctx.pres().slides;
      const t = slides.find(s => s.id === e.target);
      div.classList.toggle('zoom-missing', !t);
      if (!t) { div.appendChild(el('div', 'zoom-ph', 'Zoom: double-click to choose a slide')); return; }
      const frame = el('div', 'zoom-frame');
      const inner = el('div', 'slide zoom-inner');
      const k = Math.min(e.w / ctx.W, e.h / ctx.H);
      Object.assign(inner.style, { width: ctx.W + 'px', height: ctx.H + 'px', transform: `scale(${k})` });
      frame.style.width = ctx.W * k + 'px'; frame.style.height = ctx.H * k + 'px';
      // A zoom inside the target shows as a plain box, so zooms can't draw each other forever.
      ctx.renderSlide({ ...t, elements: t.elements.filter(x => x.type !== 'zoom') }, inner, false);
      frame.appendChild(inner);
      frame.appendChild(el('span', 'zoom-num', String(slides.indexOf(t) + 1)));
      div.appendChild(frame);
    }
    async function zoomDialog(e, isNew) {
      const slides = ctx.pres().slides;
      const here = ctx.slide();
      const opts = slides.map((s, i) => ({ value: s.id, label: `${i + 1}. ${titleOf(s) || 'Slide ' + (i + 1)}` })).filter(o => o.value !== here.id);
      if (!opts.length) { UI.toast('Add another slide to zoom to'); return false; }
      const v = await UI.form(isNew ? 'Insert Slide Zoom' : 'Slide Zoom', [{ name: 'target', label: 'In the show, clicking this goes to slide… and comes back after it', type: 'select', options: opts, value: e.target && opts.some(o => o.value === e.target) ? e.target : opts[0].value }], isNew ? 'Insert' : 'Apply');
      if (!v) return false;
      if (!isNew) ctx.snapshot();
      e.target = v.target;
      return true;
    }
    async function insertZoom() {
      const w = Math.round(ctx.W * 0.3);
      const e = { id: ctx.newId(), type: 'zoom', x: Math.round(ctx.W * 0.35), y: Math.round(ctx.H * 0.3), w, h: Math.round(w * ctx.H / ctx.W) };
      if (!(await zoomDialog(e, true))) return;
      ctx.snapshot();
      ctx.slide().elements.push(e);
      ctx.select(e.id);
      ctx.changed();
    }

    // ---------- The show ----------
    /** A slide appeared in the show (container: where it is drawn). */
    function shown(s, container) {
      watchPolls(s);
      if (ctx.pres().qa) qaStart(container);
      container.querySelectorAll('.el.zoom').forEach(node => {
        const e = s.elements.find(x => x.id === node.dataset.id);
        if (!e || !e.target) return;
        node.addEventListener('click', ev => { ev.stopPropagation(); ctx.zoomTo(e.target, node); });
      });
      remoteState();
    }
    function ended() { watchPolls(null); qaStop(); remoteState(); }

    return {
      build(e, div) { if (e.type === 'poll') { if (e.form) div.dataset.form = e.form.id; buildPoll(e, div); } else if (e.type === 'zoom') buildZoom(e, div); },
      dblclick(e) { if (e.type === 'poll') editPoll(e); else if (e.type === 'zoom') zoomDialog(e, false).then(ok => { if (ok) ctx.changed(); }); },
      insertPoll, insertZoom, clearPoll, toggleQa, startRemote, stopRemote, qaPanel,
      shown, ended, moved: remoteState,
      get remote() { return remote; },
      get qa() { return qa; },
      titleOf, qrSvg, results
    };
  }
  SlidesLive.clean = clean;
  SlidesLive.qrSvg = qrSvg;
  window.SlidesLive = SlidesLive;
})();
