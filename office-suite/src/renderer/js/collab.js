'use strict';
// Live editing: several people in one document at once, through an Isadek Office relay server
// (src/main/collab-server.js — hosted by the desktop app on the local network, or your own).
//
// Each app describes its document as a map of keys to JSON values (a paragraph, a cell, a slide,
// a shape) through an adapter. After each local change the map is compared with the last
// version everyone has, and only the keys that changed are sent. The server orders all changes,
// so every copy ends up the same: the latest change to a key wins. While your own change to a key
// is on its way, changes from others to that key are ignored (yours comes after them).
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
      remoteLog: [], remoteCount: 0, // remote changes, for undo
      queue: [], n: 0, timer: null, retry: 0, wantOpen: false, first: true, creating: false
    };
    const listeners = new Set();
    const emit = () => listeners.forEach(f => { try { f(S); } catch { /* ignore */ } });

    function send(m) { if (S.ws && S.ws.readyState === 1) { S.ws.send(JSON.stringify(m)); return true; } return false; }

    /** Compares the document with what everyone has and sends the difference. */
    function flush() {
      clearTimeout(S.timer); S.timer = null;
      if (S.status !== 'live') return;
      if (adapter.busy && adapter.busy()) { S.timer = setTimeout(flush, 300); return; }
      const cur = adapter.state();
      const ops = [];
      for (const [k, v] of Object.entries(cur)) { const j = enc(v); if (S.known.get(k) !== j) { ops.push({ k, v }); S.known.set(k, j); } }
      for (const k of [...S.known.keys()]) if (!(k in cur)) { ops.push({ k, v: null }); S.known.delete(k); }
      if (!ops.length) return;
      const n = ++S.n;
      ops.forEach(o => S.pending.set(o.k, n));
      // Big changes (pasting a large table) go in several messages.
      let batch = [], size = 0;
      const out = () => { if (batch.length) send({ t: 'op', n, ops: batch }); batch = []; size = 0; };
      for (const o of ops) {
        const s = (o.v === null ? 4 : enc(o.v).length) + o.k.length;
        if (size + s > 8e6) out();
        batch.push(o); size += s;
      }
      out();
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
      const changes = {};
      for (const m of S.queue.splice(0)) {
        for (const o of m.ops) {
          if (S.pending.has(o.k)) continue;
          changes[o.k] = o.v;
          S.remoteLog.push({ i: ++S.remoteCount, k: o.k, v: o.v });
        }
      }
      if (S.remoteLog.length > 20000) S.remoteLog.splice(0, S.remoteLog.length - 20000);
      applyRemote(changes);
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
        resend.forEach(k => S.known.delete(k));
        S.pending.clear();
        S.first = false;
        applyRemote(apply);
        if (resend.length) flush();
        emit();
        adapter.joined && adapter.joined(S);
        sendPresence(true);
        return;
      }
      if (m.t === 'op') {
        if (m.from === S.id) {
          for (const [k, n] of [...S.pending]) if (n <= m.n) S.pending.delete(k);
          return;
        }
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
      /** Starts a session for this document on a server, or joins one. */
      open({ server, inviteServer, room, create, name, color }) {
        if (S.wantOpen) stop();
        Object.assign(S, { server, inviteServer: inviteServer || null, room: room || newRoom(), creating: !!create, name, color, wantOpen: true, first: !create, retry: 0, error: null, n: 0 });
        S.pending.clear(); S.known.clear(); S.queue = []; S.remoteLog = []; S.remoteCount = 0;
        connect();
      },
      close: stop,
      changed() { if (S.status === 'live') schedule(); else if (S.wantOpen) { /* sent on reconnect */ markOffline(); } sendPresence(); },
      moved() { sendPresence(); },
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
   * adapter: { app, doc (the app's docState), state() -> {key: value}, apply(changes), busy?(), presence?(), showPeers?(peers), joined?(), left?() }
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
    btn.textContent = 'Live';
    btn.title = 'Edit together in real time';
    const share = bar.querySelector('.share-btn');
    bar.insertBefore(people, share || bar.querySelector('.qa'));
    bar.insertBefore(btn, share || bar.querySelector('.qa'));
    btn.addEventListener('mousedown', e => e.preventDefault());
    btn.addEventListener('click', () => dialog());

    session.onChange(S => {
      btn.classList.toggle('on', S.status !== 'off');
      btn.textContent = S.status === 'live' ? `Live · ${S.peers.length}` : S.status === 'off' ? 'Live' : 'Live…';
      btn.title = S.status === 'live' ? 'In a live session — click for the invite link' : S.status === 'off' ? 'Edit together in real time' : 'Connecting…';
      people.innerHTML = '';
      if (S.status !== 'off') S.peers.filter(p => p.id !== S.id).forEach(p => {
        const a = document.createElement('span');
        a.className = 'live-avatar';
        a.style.background = p.color;
        a.textContent = initials(p.name);
        a.title = p.name;
        people.appendChild(a);
      });
      adapter.showPeers && adapter.showPeers(S.status === 'live' ? S.peers.filter(p => p.id !== S.id) : []);
    });

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
        const i = await UI.choose('Live session', `Send this link to people you want to edit with:\n\n${session.invite}\n\n${session.peers.length} ${session.peers.length === 1 ? 'person' : 'people'} connected.`, ['Copy link', 'Leave session', 'Close'], '👥');
        if (i === 0) { try { await navigator.clipboard.writeText(session.invite); UI.toast('Link copied'); } catch { UI.toast('Copy the link from the dialog'); } }
        if (i === 1) { session.close(); UI.toast('Left the live session'); }
        return;
      }
      const opts = Platform.desktop ? ['Start on this computer', 'Start on a server…', 'Join…', 'Cancel'] : ['Start on a server…', 'Join…', 'Cancel'];
      const i = await UI.choose('Edit together', 'Start a live session for this document and invite others, or join one with the link you were sent.' +
        (Platform.desktop ? '\n\n“This computer” lets people on the same network join.' : ''), opts, '👥');
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
