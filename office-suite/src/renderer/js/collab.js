'use strict';
// Live editing: several people in one document at once, through an Isadek Office relay server
// (src/main/collab-server.js — hosted by the desktop app on the local network, or your own).
//
// Each app describes its document as a map of keys to JSON values (a paragraph, a cell, a slide,
// a shape) through an adapter. After each local change the map is compared with the last
// version everyone has, and only the keys that changed are sent. The server orders all changes,
// so every copy ends up the same. When someone else's change to a key arrives while yours to it is
// still on its way, the two are merged (merge3.js: by field, and by words for text) and the merged
// value is sent, so two people typing in one paragraph both keep their words.
(function () {
  const COLORS = ['#d13438', '#0078d4', '#107c10', '#8764b8', '#ca5010', '#038387', '#c239b3', '#986f0b'];
  const enc = s => JSON.stringify(s === undefined ? null : s);

  function userName() {
    try { return localStorage.getItem('atlas-user-name') || ''; } catch { return ''; }
  }
  function userColor() {
    let c = null;
    try { c = localStorage.getItem('atlas-user-color'); } catch { /* ignore */ }
    if (!c) { c = COLORS[Math.floor(Math.random() * COLORS.length)]; try { localStorage.setItem('atlas-user-color', c); } catch { /* ignore */ } }
    return c;
  }
  const newRoom = () => [...crypto.getRandomValues(new Uint8Array(16))].map(b => b.toString(16).padStart(2, '0')).join('');

  /** Invite links: <server address>?room=<id>&app=<app>. */
  function parseInvite(text) {
    let u;
    try { u = new URL(String(text).trim()); } catch { return null; }
    if (u.protocol === 'http:') u.protocol = 'ws:';
    if (u.protocol === 'https:') u.protocol = 'wss:';
    if (!/^wss?:$/.test(u.protocol)) return null;
    const room = u.searchParams.get('room');
    if (!room || !/^[\w-]{16,64}$/.test(room)) return null;
    const app = u.searchParams.get('app') || '';
    u.search = '';
    return { server: u.toString(), room, app };
  }
  const inviteFor = (server, room, app) => { const u = new URL(server); u.searchParams.set('room', room); u.searchParams.set('app', app); return u.toString(); };

  function Session(adapter) {
    const S = {
      ws: null, room: null, server: null, id: null, peers: [], status: 'off',
      known: new Map(),       // key -> JSON everyone has (as far as we know)
      pending: new Map(),     // key -> number of our latest unacknowledged change
      base: new Map(),        // key -> JSON the server has for it, as far as we know (to merge from)
      sent: new Map(),        // key -> [[n, JSON before, JSON after]] of our changes not yet acknowledged
      hist: new Map(),        // key -> [{ seq, before, after }]: our latest acknowledged changes to it
      seq: 0,                 // the last change from the server we have
      held: new Map(),        // key -> JSON the server holds after the changes handled so far
      lastFrom: new Map(),    // person -> key -> { seq, j }: their latest change to it
      activity: [],           // who changed what: { who, color, own, keys, at }
      remoteLog: [], remoteCount: 0, // remote changes, for undo
      undo: [], redo: [], record: true, // your own changes, for undo that leaves others' edits alone
      queue: [], n: 0, timer: null, retry: 0, wantOpen: false, first: true, creating: false
    };
    const listeners = new Set(), activityListeners = new Set();
    const emit = () => listeners.forEach(f => { try { f(S); } catch { /* ignore */ } });

    function send(m) { if (S.ws && S.ws.readyState === 1) { S.ws.send(JSON.stringify(m)); return true; } return false; }

    /** Compares the document with what everyone has and sends the difference. */
    function flush() {
      clearTimeout(S.timer); S.timer = null;
      if (S.status !== 'live') return;
      if (adapter.busy && adapter.busy()) { S.timer = setTimeout(flush, 300); return; }
      const cur = adapter.state();
      const ops = [], log = [];
      const keep = k => { if (!S.base.has(k)) S.base.set(k, S.known.has(k) ? S.known.get(k) : null); };
      for (const [k, v] of Object.entries(cur)) { const j = enc(v); if (S.known.get(k) !== j) { keep(k); ops.push({ k, v }); log.push({ k, before: S.known.get(k), after: j }); S.known.set(k, j); } }
      for (const k of [...S.known.keys()]) if (!(k in cur)) { keep(k); ops.push({ k, v: null }); log.push({ k, before: S.known.get(k), after: undefined }); S.known.delete(k); }
      if (!ops.length) return;
      note({ who: S.name || 'You', color: S.color, own: true }, ops.map(o => o.k));
      if (S.record) remember(log);
      const n = ++S.n;
      ops.forEach((o, i) => { S.pending.set(o.k, n); if (!S.sent.has(o.k)) S.sent.set(o.k, []); S.sent.get(o.k).push([n, log[i].before === undefined ? null : log[i].before, enc(o.v)]); });
      // Big changes (pasting a large table) go in several messages.
      let batch = [], size = 0;
      const out = () => { if (batch.length) send({ t: 'op', n, seen: S.seq, ops: batch }); batch = []; size = 0; };
      for (const o of ops) {
        const s = (o.v === null ? 4 : enc(o.v).length) + o.k.length;
        if (size + s > 8e6) out();
        batch.push(o); size += s;
      }
      out();
    }
    /** Keeps your changes for undo; quick successive changes (typing a word) are one step. */
    function remember(log) {
      const now = Date.now(), last = S.undo[S.undo.length - 1];
      S.redo = [];
      if (last && now - last.at < 1000) {
        for (const e of log) { const o = last.ops.find(x => x.k === e.k); if (o) o.after = e.after; else last.ops.push(e); }
        last.at = now;
      } else S.undo.push({ at: now, ops: log });
      if (S.undo.length > 200) S.undo.shift();
    }
    /** Takes back your last change (or puts it back), leaving keys others have changed since as they are. */
    function step(from, to) {
      flush();
      const entry = from.pop();
      if (!entry) return false;
      const changes = {}, inverse = [];
      for (const o of entry.ops) {
        if (S.known.get(o.k) !== o.after) continue; // someone else changed it since
        changes[o.k] = o.before === undefined ? null : JSON.parse(o.before);
        inverse.push({ k: o.k, before: o.after, after: o.before });
      }
      if (inverse.length) to.push({ at: 0, ops: inverse });
      if (!Object.keys(changes).length) return from.length ? step(from, to) : false;
      adapter.apply(changes);
      S.record = false;
      try { flush(); } finally { S.record = true; }
      inverse.forEach(o => { o.after = S.known.get(o.k); });
      return true;
    }
    const schedule = (ms = 150) => { if (!S.timer) S.timer = setTimeout(flush, ms); };

    function applyRemote(changes) {
      if (!Object.keys(changes).length) return;
      for (const [k, v] of Object.entries(changes)) { if (v === null) S.known.delete(k); else S.known.set(k, enc(v)); }
      adapter.apply(changes);
    }
    // Remote changes wait while you are in the middle of something (typing in a cell, dragging).
    function deliver() {
      if (!S.queue.length) return;
      if (adapter.busy && adapter.busy()) { setTimeout(deliver, 250); return; }
      flush();
      const changes = {}, merged = {};
      let cur = null;
      for (const m of S.queue.splice(0)) {
        if (m.from === S.id) { acknowledged(m); continue; }
        S.seq = Math.max(S.seq, m.seq || 0); // "seen" counts only changes applied here
        const p = S.peers.find(x => x.id === m.from);
        note({ who: p ? p.name : 'Someone', color: p ? p.color : '#888' }, m.ops.map(o => o.k));
        for (const o of m.ops) {
          S.remoteLog.push({ i: ++S.remoteCount, k: o.k, v: o.v });
          const track = () => {
            S.held.set(o.k, enc(o.v));
            if (!S.lastFrom.has(m.from)) S.lastFrom.set(m.from, new Map());
            S.lastFrom.get(m.from).set(o.k, { seq: m.seq, j: enc(o.v) });
          };
          // Our changes to this key that the sender hadn't seen: acknowledged ones (put before
          // theirs) and any still on the way (to be put after).
          const unseen = (S.hist.get(o.k) || []).filter(h => h.seq > (Number.isInteger(m.seen) ? m.seen : Infinity));
          const pend = S.pending.has(o.k);
          if (!unseen.length && !pend) { changes[o.k] = o.v; delete merged[o.k]; track(); continue; }
          if (!cur) cur = adapter.state();
          const parse = j => (j === null || j === undefined ? undefined : JSON.parse(j));
          // What they started from: their own previous version of this key if it came after our
          // first change they hadn't seen, else what the server had just before that change.
          const prev = (S.lastFrom.get(m.from) || new Map()).get(o.k);
          const base = unseen.length ? parse(prev && prev.seq > unseen[0].seq ? prev.j : unseen[0].before) : parse(S.base.get(o.k));
          const mine = o.k in merged ? merged[o.k] : cur[o.k];
          // Both people's apps work this out the same way: whose change the server put first decides.
          const m3 = Merge3.merge3(base, mine, o.v, adapter.textFields || [], '', !pend);
          if (pend) S.base.set(o.k, o.v === null ? null : enc(o.v));
          if (enc(m3) === enc(o.v)) { changes[o.k] = o.v; delete merged[o.k]; track(); continue; } // theirs has it all
          if (pend && enc(m3) === enc(mine)) { track(); continue; } // ours lands after theirs with the same result
          // Otherwise the merge is shown here and sent to everyone.
          if (o.v === null) S.known.delete(o.k); else S.known.set(o.k, enc(o.v));
          track();
          merged[o.k] = m3 === undefined ? null : m3;
        }
      }
      if (S.remoteLog.length > 20000) S.remoteLog.splice(0, S.remoteLog.length - 20000);
      applyRemote(changes);
      // The merged values are local changes: shown here and sent to everyone.
      const local = {};
      for (const [k, v] of Object.entries(merged)) if (!cur || enc(cur[k]) !== enc(v)) local[k] = v;
      if (Object.keys(local).length) adapter.apply(local);
      if (Object.keys(merged).length) flush();
    }
    /** The server has our changes up to m.n. */
    function acknowledged(m) {
      // They are now the base to merge from, and are remembered with where the server put them,
      // for changes others made without seeing them.
      S.seq = Math.max(S.seq, m.seq || 0);
      for (const [k, list] of [...S.sent]) {
        let done = null;
        while (list.length && list[0][0] <= m.n) { const [, , after] = list.shift(); done = done ? { ...done, after } : { seq: m.seq, before: S.held.has(k) ? S.held.get(k) : null, after }; }
        if (done) S.held.set(k, done.after);
        if (done) {
          S.base.set(k, done.after);
          const h = S.hist.get(k) || [];
          h.push(done);
          if (h.length > 20) h.shift();
          S.hist.set(k, h);
        }
        if (!list.length) S.sent.delete(k);
      }
      for (const [k, n] of [...S.pending]) if (n <= m.n) { S.pending.delete(k); S.base.delete(k); }
    }
    /** Who changed what, for the Activity list (a person's quick successive changes are one entry). */
    function note(who, keys) {
      const now = Date.now(), last = S.activity[S.activity.length - 1];
      if (last && last.who === who.who && !!last.own === !!who.own && now - last.at < 60000) { keys.forEach(k => last.keys.add(k)); last.at = now; }
      else S.activity.push({ ...who, keys: new Set(keys), at: now });
      if (S.activity.length > 300) S.activity.shift();
      activityListeners.forEach(f => { try { f(); } catch { /* ignore */ } });
    }

    function onMessage(e) {
      let m;
      try { m = JSON.parse(e.data); } catch { return; }
      if (m.t === 'error') {
        S.error = m.error;
        UI.toast(m.error, 5000);
        if (S.status !== 'live') stop();
        return;
      }
      if (m.t === 'welcome') {
        S.id = m.id; S.peers = m.peers; S.status = 'live'; S.retry = 0; S.error = null;
        const local = adapter.state();
        const server = m.state;
        const apply = {}, resend = [];
        const keys = new Set([...Object.keys(local), ...Object.keys(server)]);
        for (const k of keys) {
          const lj = k in local ? enc(local[k]) : null, sj = k in server ? enc(server[k]) : null;
          if (lj === sj) continue;
          // Joining: the session's document replaces yours. Reconnecting: your edits made
          // while offline are sent, everything else is brought up to date.
          if (!S.first && S.pending.has(k)) resend.push(k);
          else apply[k] = k in server ? server[k] : null;
        }
        S.known = new Map(Object.entries(server).map(([k, v]) => [k, enc(v)]));
        S.held = new Map(S.known); S.lastFrom.clear();
        S.seq = m.seq || 0;
        resend.forEach(k => S.known.delete(k));
        S.pending.clear(); S.base.clear(); S.sent.clear(); S.hist.clear();
        S.first = false;
        applyRemote(apply);
        if (resend.length) flush();
        emit();
        adapter.joined && adapter.joined(S);
        sendPresence(true);
        return;
      }
      if (m.t === 'op') {
        // Ours (acknowledged) and others' changes are handled in the server's order.
        S.queue.push(m);
        deliver();
        return;
      }
      if (m.t === 'peers') {
        const old = new Map(S.peers.map(p => [p.id, p]));
        S.peers = m.peers.map(p => ({ ...p, presence: p.presence || old.get(p.id)?.presence || null }));
        emit();
        return;
      }
      if (m.t === 'presence') {
        const p = S.peers.find(x => x.id === m.from);
        if (p) { p.presence = m.data; emit(); }
      }
    }

    function connect() {
      S.status = 'connecting';
      emit();
      let ws;
      try { ws = new WebSocket(S.server); } catch (err) { S.error = err.message; return fail(); }
      S.ws = ws;
      ws.onopen = () => send({ t: 'hello', room: S.room, create: S.creating, app: adapter.app, name: S.name, color: S.color, init: S.creating ? initState() : undefined });
      ws.onmessage = onMessage;
      ws.onclose = () => { if (S.ws === ws) fail(); };
      ws.onerror = () => {};
    }
    function initState() {
      const st = adapter.state();
      S.known = new Map(Object.entries(st).map(([k, v]) => [k, enc(v)]));
      return st;
    }
    function fail() {
      S.ws = null;
      if (!S.wantOpen) return;
      if (S.status === 'live') S.creating = false; // the session exists now
      S.status = 'reconnecting';
      emit();
      const wait = Math.min(15000, 500 * 2 ** S.retry++);
      if (S.retry > 12 || (S.creating && S.retry > 2)) { UI.toast('Lost the live session' + (S.error ? ': ' + S.error : ''), 5000); return stop(); }
      setTimeout(() => { if (S.wantOpen) connect(); }, wait);
    }
    function stop() {
      S.wantOpen = false;
      const ws = S.ws; S.ws = null;
      if (ws) try { ws.close(); } catch { /* ignore */ }
      S.status = 'off'; S.peers = []; S.room = null;
      emit();
      adapter.left && adapter.left();
    }

    let lastPresence = '', presenceTimer = null;
    function sendPresence(force) {
      if (S.status !== 'live' || !adapter.presence) return;
      clearTimeout(presenceTimer);
      presenceTimer = setTimeout(() => {
        const p = adapter.presence();
        const j = enc(p);
        if (!force && j === lastPresence) return;
        lastPresence = j;
        send({ t: 'presence', data: p });
      }, 120);
    }

    return {
      get status() { return S.status; },
      get peers() { return S.peers; },
      get id() { return S.id; },
      get invite() { return S.room ? inviteFor(S.inviteServer || S.server, S.room, adapter.app) : null; },
      onChange(f) { listeners.add(f); },
      onActivity(f) { activityListeners.add(f); },
      get activity() { return S.activity; },
      /** Starts a session for this document on a server, or joins one. */
      open({ server, inviteServer, room, create, name, color }) {
        if (S.wantOpen) stop();
        Object.assign(S, { server, inviteServer: inviteServer || null, room: room || newRoom(), creating: !!create, name, color, wantOpen: true, first: !create, retry: 0, error: null, n: 0 });
        S.pending.clear(); S.base.clear(); S.sent.clear(); S.hist.clear(); S.held.clear(); S.lastFrom.clear(); S.seq = 0; S.activity = []; S.known.clear(); S.queue = []; S.remoteLog = []; S.remoteCount = 0; S.undo = []; S.redo = [];
        connect();
      },
      close: stop,
      changed() { if (S.status === 'live') schedule(); else if (S.wantOpen) { /* sent on reconnect */ markOffline(); } sendPresence(); },
      moved() { sendPresence(); },
      /** Undo and redo of only your own changes, for apps whose undo is the browser's. */
      undoLocal() { return step(S.undo, S.redo); },
      redoLocal() { return step(S.redo, S.undo); },
      /** Undo support: a marker for "now", and replay of remote changes made after a marker. */
      mark() { return S.remoteCount; },
      replaySince(mark) {
        const changes = {};
        for (const r of S.remoteLog) if (r.i > mark) changes[r.k] = r.v;
        for (const k of Object.keys(changes)) if (S.pending.has(k)) delete changes[k];
        if (Object.keys(changes).length) adapter.apply(changes);
        schedule(50);
      }
    };
    function markOffline() {
      // Offline edits: remember which keys changed so they win on reconnect.
      const cur = adapter.state();
      for (const [k, v] of Object.entries(cur)) if (S.known.get(k) !== enc(v)) S.pending.set(k, ++S.n);
      for (const k of S.known.keys()) if (!(k in cur)) S.pending.set(k, ++S.n);
    }
  }

  // ---------- UI ----------
  function initials(n) { return (String(n).trim().split(/\s+/).map(w => w[0]).join('').slice(0, 2) || '?').toUpperCase(); }

  /**
   * attach(adapter): adds the Live button and people to the title bar of an app.
   * adapter: { app, doc (the app's docState), state() -> {key: value}, apply(changes), busy?(), presence?(), showPeers?(peers), joined?(), left?(),
   *            describe?(key) -> label for the Activity list, reveal?(key), follow?(presence, peer) }
   */
  function attach(adapter) {
    const session = Session(adapter);
    // Every change to the document (the apps all call markDirty) is a reason to sync.
    const ds = adapter.doc;
    if (ds) {
      const markDirty = ds.markDirty.bind(ds);
      ds.markDirty = () => { markDirty(); session.changed(); };
      const set = ds.set.bind(ds);
      ds.set = props => { set(props); if (props.dirty === true) session.changed(); };
    }
    const bar = document.querySelector('.titlebar');
    const people = document.createElement('div');
    people.className = 'live-people';
    people.setAttribute('aria-live', 'polite');
    const btn = document.createElement('button');
    btn.className = 'share-btn live-btn';
    btn.innerHTML = '<span class="ico" data-icon="users"></span><span class="lbl">Live</span>';
    const label = btn.querySelector('.lbl');
    btn.title = 'Edit together in real time';
    const share = bar.querySelector('.share-btn');
    bar.insertBefore(people, share || bar.querySelector('.qa'));
    bar.insertBefore(btn, share || bar.querySelector('.qa'));
    btn.addEventListener('mousedown', e => e.preventDefault());
    btn.addEventListener('click', () => dialog());

    // Following someone: your view goes where they are (their slide, their paragraph).
    let following = null;
    const follow = p => {
      following = following === p.id ? null : p.id;
      UI.toast(following ? `Following ${p.name}` : `Stopped following ${p.name}`);
      session.moved();
      paint(session);
      if (following && p.presence) adapter.follow(p.presence, p);
    };
    session.onChange(S => {
      if (following) {
        const p = S.peers.find(x => x.id === following);
        if (!p || S.status !== 'live') following = null;
        else if (p.presence) adapter.follow(p.presence, p);
      }
      paint(S);
    });
    function paint(S) {
      btn.classList.toggle('on', S.status !== 'off');
      label.textContent = S.status === 'live' ? `Live · ${S.peers.length}` : S.status === 'off' ? 'Live' : 'Live…';
      btn.title = S.status === 'live' ? 'In a live session — click for the invite link' : S.status === 'off' ? 'Edit together in real time' : 'Connecting…';
      people.innerHTML = '';
      if (S.status !== 'off') S.peers.filter(p => p.id !== S.id).forEach(p => {
        const a = document.createElement('span');
        a.className = 'live-avatar' + (following === p.id ? ' following' : '');
        a.style.background = p.color;
        a.textContent = initials(p.name);
        a.title = p.name;
        if (adapter.follow) {
          a.setAttribute('role', 'button');
          a.tabIndex = 0;
          a.title = following === p.id ? `${p.name} — following (click to stop)` : `${p.name} — click to follow`;
          a.setAttribute('aria-pressed', String(following === p.id));
          a.addEventListener('click', () => follow(p));
          a.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); follow(p); } });
        }
        people.appendChild(a);
      });
      adapter.showPeers && adapter.showPeers(S.status === 'live' ? S.peers.filter(p => p.id !== S.id) : []);
    }

    // ---- Activity: who changed what in this session
    let pane = null;
    const ago = t => { const s = Math.round((Date.now() - t) / 1000); return s < 10 ? 'just now' : s < 60 ? `${s} s ago` : s < 3600 ? `${Math.floor(s / 60)} min ago` : new Date(t).toLocaleTimeString(); };
    function renderActivity() {
      if (!pane || pane.hidden) return;
      const list = pane.querySelector('.list');
      list.innerHTML = '';
      const items = session.activity.slice().reverse();
      if (!items.length) { const e = document.createElement('p'); e.className = 'empty'; e.textContent = 'Changes made in this session appear here.'; list.appendChild(e); return; }
      for (const a of items) {
        const row = document.createElement('div');
        row.className = 'act';
        const dot = document.createElement('span'); dot.className = 'dot'; dot.style.background = a.color || '#888';
        const body = document.createElement('div'); body.className = 'what';
        const who = document.createElement('b'); who.textContent = a.own ? `${a.who} (you)` : a.who;
        const when = document.createElement('span'); when.className = 'when'; when.textContent = ago(a.at);
        const keys = [...a.keys];
        const labels = [...new Set(keys.map(k => (adapter.describe ? adapter.describe(k) : null)).filter(Boolean))];
        const what = document.createElement('div');
        what.textContent = labels.length ? labels.slice(0, 3).join(' · ') + (labels.length > 3 ? ` and ${labels.length - 3} more` : '') : `${keys.length} change${keys.length === 1 ? '' : 's'}`;
        body.append(who, ' ', when, what);
        row.append(dot, body);
        if (adapter.reveal) { row.tabIndex = 0; row.addEventListener('click', () => adapter.reveal(keys[keys.length - 1])); }
        list.appendChild(row);
      }
    }
    function showActivity() {
      if (!pane) {
        pane = document.createElement('aside');
        pane.className = 'activity-pane';
        pane.setAttribute('aria-label', 'Activity');
        pane.innerHTML = '<header><b>Activity</b><span class="sp"></span><button class="x" title="Close"><span class="ico" data-icon="x"></span></button></header><div class="list" aria-live="polite"></div>';
        pane.querySelector('.x').addEventListener('click', () => { pane.hidden = true; });
        document.body.appendChild(pane);
        session.onActivity(() => renderActivity());
        setInterval(renderActivity, 30000);
      }
      pane.hidden = false;
      renderActivity();
    }

    async function askName() {
      let name = userName();
      if (!name) {
        name = (await UI.prompt('Your name', 'Name others will see', '') || '').trim().slice(0, 60);
        if (!name) return null;
        try { localStorage.setItem('atlas-user-name', name); } catch { /* ignore */ }
      }
      return name;
    }
    async function dialog() {
      if (session.status !== 'off') {
        const i = await UI.choose('Live session', `Send this link to people you want to edit with:\n\n${session.invite}\n\n${session.peers.length} ${session.peers.length === 1 ? 'person' : 'people'} connected.` + (adapter.follow ? ' Click someone\'s picture to follow them.' : ''), ['Copy link', 'Activity', 'Leave session', 'Close']);
        if (i === 0) { try { await navigator.clipboard.writeText(session.invite); UI.toast('Link copied'); } catch { UI.toast('Copy the link from the dialog'); } }
        if (i === 1) showActivity();
        if (i === 2) { session.close(); UI.toast('Left the live session'); }
        return;
      }
      const opts = Platform.desktop ? ['Start on this computer', 'Start on a server…', 'Join…', 'Cancel'] : ['Start on a server…', 'Join…', 'Cancel'];
      const i = await UI.choose('Edit together', 'Start a live session for this document and invite others, or join one with the link you were sent.' +
        (Platform.desktop ? '\n\n“This computer” lets people on the same network join.' : ''), opts);
      const pick = opts[i];
      if (!pick || pick === 'Cancel') return;
      const name = await askName();
      if (!name) return;
      if (pick === 'Join…') {
        const link = await UI.prompt('Join a live session', 'Invite link', '');
        if (!link) return;
        return join(link, name);
      }
      let server, inviteServer = null;
      if (pick === 'Start on this computer') {
        const h = await window.atlas.collabHost();
        if (!h || !h.urls) return UI.toast('Could not start a session: ' + (h && h.error || 'unknown error'), 5000);
        // This computer connects to itself directly; others use its network address.
        server = h.urls[h.urls.length - 1];
        inviteServer = h.urls[0];
      } else {
        let last = '';
        try { last = localStorage.getItem('atlas-collab-server') || ''; } catch { /* ignore */ }
        const v = await UI.prompt('Start a live session', 'Server address (ws:// or wss://)', last || 'wss://');
        if (!v) return;
        const u = parseInvite(v.replace(/[?#].*$/, '') + '?room=' + 'x'.repeat(16));
        if (!u) return UI.toast('Enter a ws:// or wss:// address', 4000);
        server = u.server;
        try { localStorage.setItem('atlas-collab-server', server); } catch { /* ignore */ }
      }
      session.open({ server, inviteServer, create: true, name, color: userColor() });
      waitLive(() => dialog());
    }
    function waitLive(then) {
      const t0 = Date.now();
      const tick = () => {
        if (session.status === 'live') return then && then();
        if (session.status === 'off' || Date.now() - t0 > 15000) return;
        setTimeout(tick, 150);
      };
      tick();
    }
    async function join(link, name) {
      const inv = parseInvite(link);
      if (!inv) return UI.toast('That is not an Isadek Office live link', 4000);
      if (inv.app && inv.app !== adapter.app) {
        // A link for another app: open that app and join there.
        if (Platform.desktop) return window.atlas.collabJoin(inv.app, link);
        location.href = `${inv.app}.html?live=${encodeURIComponent(link)}`;
        return;
      }
      if (adapter.dirty && adapter.dirty() && !(await UI.confirm('Join live session', 'Joining replaces this document with the shared one. Continue?'))) return;
      session.open({ server: inv.server, room: inv.room, create: false, name: name || await askName(), color: userColor() });
      waitLive(() => UI.toast('Joined the live session'));
    }
    // ?live=<invite> (web) or a link handed over by the desktop app.
    (async () => {
      let link = new URLSearchParams(location.search).get('live');
      if (!link && Platform.desktop && window.atlas.collabPending) link = await window.atlas.collabPending();
      if (!link) return;
      if (!Platform.desktop) history.replaceState(null, '', location.pathname);
      const name = await askName();
      if (name) join(link, name);
    })();
    window.addEventListener('pagehide', () => session.close());
    session.showActivity = showActivity;
    // Who made the next saved version: you, and everyone who changed something since the last one.
    let since = 0;
    Platform.contributors = () => {
      const me = userName() || (() => { try { return localStorage.getItem('atlas-author') || ''; } catch { return ''; } })();
      const names = [me, ...session.activity.filter(a => a.at >= since && !a.own).map(a => a.who)].filter(Boolean);
      since = Date.now();
      return [...new Set(names)];
    };
    session.follow = id => { const p = session.peers.find(x => x.id === id || x.name === id); if (p) follow(p); };
    return session;
  }

  // ---------- Helpers for adapters ----------
  /** Keys that sort between a and b ('' = open end); used to order paragraphs. */
  const DIGITS = '0123456789abcdefghijklmnopqrstuvwxyz';
  function between(a, b) {
    a = a || ''; b = b || '';
    // Appending (the usual case when typing) counts up one digit at a time, so keys stay short.
    if (!b && a) {
      for (let i = 0; i < a.length; i++) if (a[i] !== 'z') return a.slice(0, i) + DIGITS[DIGITS.indexOf(a[i]) + 1];
      return a + '1';
    }
    let out = '';
    for (let i = 0; ; i++) {
      const da = i < a.length ? DIGITS.indexOf(a[i]) : 0;
      const db = i < b.length ? DIGITS.indexOf(b[i]) : DIGITS.length;
      if (db - da > 1) return out + DIGITS[Math.floor((da + db) / 2)];
      out += DIGITS[da];
      if (da < db) {
        // b is now effectively unbounded above for the remaining digits
        b = '';
        for (let j = i + 1; ; j++) {
          const d = j < a.length ? DIGITS.indexOf(a[j]) : 0;
          if (d < DIGITS.length - 1) return out + DIGITS[Math.floor((d + DIGITS.length) / 2)];
          out += DIGITS[d];
        }
      }
    }
  }

  window.Collab = { attach, Session, parseInvite, inviteFor, between, COLORS };
})();
