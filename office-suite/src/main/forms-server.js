'use strict';
// Isadek Forms — collecting answers. A form is published with its definition and gets an id
// (the link people answer at, unguessable) and an owner key (needed to change it, read the
// answers or delete it; only its hash is kept). Answers are checked against the form as they
// arrive and kept in order, numbered, so Sheets can fetch the ones it hasn't got yet.
//
// createFormsStore({ dir }) keeps forms in a folder (one .json and one .ndjson per form), or in
// memory without one. store.handle(req, res, url) answers /api/forms/… and /f/<id>; the web
// server uses it, and the desktop app runs createFormsServer() to collect answers on the local
// network: just the answer page and the API.
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const FormsCore = require('../renderer/js/forms-core');

const LIMITS = { remotes: 2000, remoteIdleMs: 12 * 3600e3, commandsPerMinute: 240, waitMs: 20000, forms: 10000, responses: 50000, storedBytes: 100e6, formBytes: 512e3, answerBytes: 256e3, perMinute: 30, createPerHour: 60 };
const ID_RE = /^[A-Za-z0-9_-]{22}$/;
const token = n => crypto.randomBytes(n).toString('base64url');
const hash = k => crypto.createHash('sha256').update(String(k)).digest('hex');
const same = (a, b) => { const x = Buffer.from(a), y = Buffer.from(b); return x.length === y.length && crypto.timingSafeEqual(x, y); };

function createFormsStore(opts = {}) {
  const limits = { ...LIMITS, ...(opts.limits || {}) };
  const dir = opts.dir ? path.resolve(opts.dir) : null;
  if (dir) fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  const forms = new Map(); // id -> { form, keyHash, open, created, count, responses (in memory only) }

  // ---------- Keeping ----------
  const file = (id, ext) => path.join(dir, id + ext);
  function saveMeta(id, f) {
    if (!dir) return;
    const tmp = file(id, '.json.tmp');
    fs.writeFileSync(tmp, JSON.stringify({ form: f.form, keyHash: f.keyHash, open: f.open, created: f.created }), { mode: 0o600 });
    fs.renameSync(tmp, file(id, '.json'));
  }
  function load(id) {
    if (forms.has(id)) return forms.get(id);
    if (!dir || !ID_RE.test(id)) return null;
    let meta;
    try { meta = JSON.parse(fs.readFileSync(file(id, '.json'), 'utf8')); } catch { return null; }
    const form = FormsCore.clean(meta.form);
    if (!form || typeof meta.keyHash !== 'string') return null;
    let count = 0, bytes = 0;
    try { const text = fs.readFileSync(file(id, '.ndjson'), 'utf8'); count = text ? text.split('\n').filter(Boolean).length : 0; bytes = Buffer.byteLength(text); } catch { /* none yet */ }
    const f = { form, keyHash: meta.keyHash, open: meta.open !== false, created: meta.created || 0, count, bytes };
    forms.set(id, f);
    return f;
  }
  const total = () => (dir ? fs.readdirSync(dir).filter(n => n.endsWith('.json')).length : forms.size);

  // ---------- The operations (also called directly by the desktop app) ----------
  function create(def) {
    const form = FormsCore.clean(def);
    if (!form || !form.questions.length) throw Object.assign(new Error('A form needs at least one question'), { status: 400 });
    if (total() >= limits.forms) throw Object.assign(new Error('This server holds as many forms as it can'), { status: 507 });
    const id = token(16), key = token(32);
    const f = { form, keyHash: hash(key), open: true, created: Date.now(), count: 0, bytes: 0, responses: dir ? null : [] };
    forms.set(id, f);
    saveMeta(id, f);
    return { id, key };
  }
  function owned(id, key) {
    const f = ID_RE.test(String(id)) ? load(id) : null;
    if (!f) throw Object.assign(new Error('There is no such form'), { status: 404 });
    if (typeof key !== 'string' || !same(hash(key), f.keyHash)) throw Object.assign(new Error('That key doesn’t open this form'), { status: 403 });
    return f;
  }
  function update(id, key, { form: def, open } = {}) {
    const f = owned(id, key);
    if (def !== undefined) {
      const form = FormsCore.clean(def);
      if (!form || !form.questions.length) throw Object.assign(new Error('A form needs at least one question'), { status: 400 });
      f.form = form;
    }
    if (typeof open === 'boolean') f.open = open;
    saveMeta(id, f);
    return { open: f.open, count: f.count };
  }
  function remove(id, key) {
    owned(id, key);
    forms.delete(id);
    if (dir) for (const ext of ['.json', '.ndjson']) try { fs.unlinkSync(file(id, ext)); } catch { /* gone */ }
    return { deleted: true };
  }
  /** The public part: what people answering see. */
  function publicForm(id) {
    const f = ID_RE.test(String(id)) ? load(id) : null;
    if (!f) throw Object.assign(new Error('This form doesn’t exist, or was deleted'), { status: 404 });
    return { form: f.form, open: f.open };
  }
  function respond(id, answers) {
    const f = ID_RE.test(String(id)) ? load(id) : null;
    if (!f) throw Object.assign(new Error('This form doesn’t exist, or was deleted'), { status: 404 });
    if (!f.open) throw Object.assign(new Error('This form is no longer accepting responses'), { status: 403 });
    if (f.count >= limits.responses || f.bytes >= limits.storedBytes) throw Object.assign(new Error('This form has as many responses as it can take'), { status: 507 });
    const r = FormsCore.check(f.form, answers);
    if (!r.ok) throw Object.assign(new Error('Some answers need fixing'), { status: 422, errors: r.errors });
    const entry = { n: f.count + 1, at: new Date().toISOString(), values: r.values };
    const line = JSON.stringify(entry) + '\n';
    if (dir) fs.appendFileSync(file(id, '.ndjson'), line, { mode: 0o600 });
    else f.responses.push(entry);
    f.count++;
    f.bytes += Buffer.byteLength(line);
    return { n: entry.n };
  }
  /** Responses after number `after` (at most `limit`), for the owner. */
  function responses(id, key, after = 0, limit = 1000) {
    const f = owned(id, key);
    after = Math.max(0, Math.floor(+after) || 0);
    limit = Math.min(5000, Math.max(1, Math.floor(+limit) || 1000));
    let list;
    if (dir) {
      let text = '';
      try { text = fs.readFileSync(file(id, '.ndjson'), 'utf8'); } catch { /* none yet */ }
      list = text.split('\n').filter(Boolean).slice(after, after + limit).map(l => { try { return JSON.parse(l); } catch { return null; } }).filter(Boolean);
    } else list = f.responses.slice(after, after + limit);
    return { form: f.form, open: f.open, count: f.count, responses: list };
  }

  // ---------- The phone remote: a presenter's show and a phone, through this server ----------
  // Kept in memory only. The presenter holds the key (state out, commands in); the phone holds
  // the token (state in, commands out). Waiting requests are answered as soon as there is news.
  const remotes = new Map();
  const COMMANDS = ['next', 'prev', 'first', 'last', 'black'];
  const wake = set => { for (const fn of [...set]) fn(); set.clear(); };
  function createRemote() {
    if (remotes.size >= limits.remotes) throw Object.assign(new Error('This server has as many remotes as it can take'), { status: 507 });
    const id = token(16), key = token(32), tok = token(32);
    remotes.set(id, { keyHash: hash(key), tokenHash: hash(tok), state: { v: 0 }, cmds: [], n: 0, stateWait: new Set(), cmdWait: new Set(), touched: Date.now(), hits: [] });
    return { id, key, token: tok };
  }
  function remoteFor(id, secret, which) {
    const r = ID_RE.test(String(id)) ? remotes.get(id) : null;
    if (!r) throw Object.assign(new Error('This remote has ended'), { status: 404 });
    if (typeof secret !== 'string' || !same(hash(secret), r[which])) throw Object.assign(new Error('That link doesn’t control this show'), { status: 403 });
    r.touched = Date.now();
    return r;
  }
  const text = (v, n) => (typeof v === 'string' ? v.slice(0, n) : '');
  function setRemoteState(id, key, st = {}) {
    const r = remoteFor(id, key, 'keyHash');
    const int = (v, hi) => (Number.isInteger(v) && v >= 0 && v <= hi ? v : 0);
    r.state = { v: r.state.v + 1, slide: int(st.slide, 10000), total: int(st.total, 10000), title: text(st.title, 300), next: text(st.next, 300), notes: text(st.notes, 20000), black: st.black === true, showing: st.showing === true, started: int(st.started, 1e13) };
    wake(r.stateWait);
    return { v: r.state.v };
  }
  /** Waits (up to ms) for something newer than `since`; resolves to its value, or the current one. */
  function waitFor(r, set, ready, ms, close) {
    if (ready()) return Promise.resolve();
    return new Promise(resolve => {
      const done = () => { clearTimeout(t); set.delete(done); resolve(); };
      const t = setTimeout(done, ms);
      set.add(done);
      if (close) close(done);
    });
  }
  async function remoteState(id, tok, since = 0, ms = 0, close) {
    const r = remoteFor(id, tok, 'tokenHash');
    since = Math.floor(+since) || 0;
    await waitFor(r, r.stateWait, () => r.state.v > since, Math.min(ms, limits.waitMs), close);
    return r.state;
  }
  function remoteCommand(id, tok, cmd) {
    const r = remoteFor(id, tok, 'tokenHash');
    if (!COMMANDS.includes(cmd)) throw Object.assign(new Error('Unknown command'), { status: 400 });
    const cut = Date.now() - 60000;
    r.hits = r.hits.filter(x => x > cut);
    if (r.hits.length >= limits.commandsPerMinute) throw Object.assign(new Error('Too many taps in the last minute'), { status: 429 });
    r.hits.push(Date.now());
    r.cmds.push({ n: ++r.n, cmd, at: Date.now() });
    r.cmds = r.cmds.slice(-100);
    wake(r.cmdWait);
    return { n: r.n };
  }
  async function remoteCommands(id, key, after = 0, ms = 0, close) {
    const r = remoteFor(id, key, 'keyHash');
    after = Math.floor(+after) || 0;
    await waitFor(r, r.cmdWait, () => r.n > after, Math.min(ms, limits.waitMs), close);
    return { n: r.n, commands: r.cmds.filter(c => c.n > after) };
  }
  function removeRemote(id, key) {
    const r = remoteFor(id, key, 'keyHash');
    wake(r.stateWait); wake(r.cmdWait);
    remotes.delete(id);
    return { deleted: true };
  }

  // ---------- HTTP ----------
  // Answers per address per minute, and forms created per address per hour.
  const hits = new Map();
  const timer = setInterval(() => {
    const cut = Date.now() - 3600e3;
    for (const [k, t] of hits) { const keep = t.filter(x => x > cut); if (keep.length) hits.set(k, keep); else hits.delete(k); }
    for (const [id, r] of remotes) if (Date.now() - r.touched > limits.remoteIdleMs) { wake(r.stateWait); wake(r.cmdWait); remotes.delete(id); }
  }, 60000);
  timer.unref();
  function allowed(kind, who, max, ms) {
    const k = kind + ' ' + who, cut = Date.now() - ms;
    const t = (hits.get(k) || []).filter(x => x > cut);
    if (t.length >= max) return false;
    t.push(Date.now()); hits.set(k, t.slice(-max));
    return true;
  }
  function readBody(req, max) {
    return new Promise((resolve, reject) => {
      const parts = []; let n = 0;
      req.on('data', d => { n += d.length; if (n > max) { reject(Object.assign(new Error('That is too much to send'), { status: 413 })); req.destroy(); } else parts.push(d); });
      req.on('end', () => { try { resolve(JSON.parse(Buffer.concat(parts).toString('utf8') || '{}')); } catch { reject(Object.assign(new Error('Bad request'), { status: 400 })); } });
      req.on('error', reject);
    });
  }
  const json = (res, code, body) => { res.writeHead(code, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }); res.end(JSON.stringify(body)); };
  const keyOf = req => (String(req.headers.authorization || '').match(/^Bearer ([A-Za-z0-9_-]{1,100})$/) || [])[1] || '';

  /** Handles /api/forms… and /f/<id>; returns false for anything else. */
  function handle(req, res, url) {
    if (url.pathname.startsWith('/f/')) {
      const id = url.pathname.slice(3);
      if (!ID_RE.test(id)) { res.writeHead(404); res.end('Not found'); return true; }
      res.writeHead(302, { Location: '/src/renderer/form.html?id=' + id, 'Cache-Control': 'no-store' });
      res.end();
      return true;
    }
    if (url.pathname.startsWith('/r/')) {
      const id = url.pathname.slice(3);
      if (!ID_RE.test(id)) { res.writeHead(404); res.end('Not found'); return true; }
      // The phone's token stays in the link's #fragment, which browsers keep across the redirect.
      res.writeHead(302, { Location: '/src/renderer/remote.html?id=' + id, 'Cache-Control': 'no-store' });
      res.end();
      return true;
    }
    if (!url.pathname.startsWith('/api/forms') && !url.pathname.startsWith('/api/remote')) return false;
    // The owner's calls carry their key, never cookies, so any page (the desktop app's, a phone's)
    // may make them.
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
    res.setHeader('Access-Control-Max-Age', '600');
    if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return true; }
    const who = req.socket.remoteAddress || '?';
    const parts = url.pathname.split('/').filter(Boolean); // api, forms|remote, id?, sub?
    const id = parts[2], sub = parts[3];
    // A waiting request ends early when its connection closes.
    const onClose = done => res.on('close', done);
    const fail = err => { if (!res.headersSent) json(res, err.status || 500, { error: err.status ? err.message : 'Something went wrong', ...(err.errors ? { errors: err.errors } : {}) }); else res.end(); };
    (async () => {
      if (parts[1] === 'remote') {
        if (parts.length === 2 && req.method === 'POST') {
          if (!allowed('create', who, limits.createPerHour, 3600e3)) return json(res, 429, { error: 'Too many new remotes from here in the last hour' });
          return json(res, 201, createRemote());
        }
        if (parts.length === 3 && req.method === 'DELETE') return json(res, 200, removeRemote(id, keyOf(req)));
        if (parts.length === 4 && sub === 'state' && req.method === 'PUT') return json(res, 200, setRemoteState(id, keyOf(req), await readBody(req, 64e3)));
        if (parts.length === 4 && sub === 'state' && req.method === 'GET') return json(res, 200, await remoteState(id, keyOf(req), url.searchParams.get('since'), limits.waitMs, onClose));
        if (parts.length === 4 && sub === 'command' && req.method === 'POST') return json(res, 201, remoteCommand(id, keyOf(req), (await readBody(req, 1e3)).cmd));
        if (parts.length === 4 && sub === 'commands' && req.method === 'GET') return json(res, 200, await remoteCommands(id, keyOf(req), url.searchParams.get('after'), limits.waitMs, onClose));
        return json(res, 404, { error: 'Not found' });
      }
      if (parts.length === 2 && req.method === 'GET') return json(res, 200, { available: true });
      if (parts.length === 2 && req.method === 'POST') {
        if (!allowed('create', who, limits.createPerHour, 3600e3)) return json(res, 429, { error: 'Too many new forms from here in the last hour' });
        const body = await readBody(req, limits.formBytes);
        return json(res, 201, create(body.form));
      }
      if (parts.length === 3 && req.method === 'GET') return json(res, 200, publicForm(id));
      if (parts.length === 3 && req.method === 'PUT') return json(res, 200, update(id, keyOf(req), await readBody(req, limits.formBytes)));
      if (parts.length === 3 && req.method === 'DELETE') return json(res, 200, remove(id, keyOf(req)));
      if (parts.length === 4 && sub === 'responses' && req.method === 'POST') {
        if (!allowed('answer', who, limits.perMinute, 60000)) return json(res, 429, { error: 'Too many responses from here in the last minute. Try again shortly.' });
        const body = await readBody(req, limits.answerBytes);
        return json(res, 201, respond(id, body.answers));
      }
      if (parts.length === 4 && sub === 'responses' && req.method === 'GET') return json(res, 200, responses(id, keyOf(req), url.searchParams.get('after'), url.searchParams.get('limit')));
      json(res, 404, { error: 'Not found' });
    })().catch(fail);
    return true;
  }

  const close = () => { clearInterval(timer); for (const r of remotes.values()) { wake(r.stateWait); wake(r.cmdWait); } };
  return { create, update, remove, publicForm, respond, responses, createRemote, setRemoteState, remoteState, remoteCommand, remoteCommands, removeRemote, handle, close };
}

// The files the answer page needs, and nothing else of the app.
const PAGE_FILES = ['src/renderer/form.html', 'src/renderer/css/form.css', 'src/renderer/js/form.js', 'src/renderer/js/forms-core.js', 'src/renderer/js/forms-view.js', 'src/renderer/remote.html', 'src/renderer/css/remote.css', 'src/renderer/js/remote.js'];
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8' };
const ROOT = path.resolve(__dirname, '../..');

/** A server for the local network: the answer page and the forms API (the desktop app's "This computer"). */
function createFormsServer(opts = {}) {
  const store = opts.store || createFormsStore(opts);
  const server = http.createServer((req, res) => {
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Frame-Options', 'DENY');
    res.setHeader('Referrer-Policy', 'no-referrer');
    let url;
    try { url = new URL(req.url, 'http://x'); } catch { res.writeHead(400); return res.end(); }
    try { if (store.handle(req, res, url)) return; } catch { if (!res.headersSent) { res.writeHead(500); res.end(); } return; }
    const rel = url.pathname.replace(/^\/+/, '');
    if ((req.method === 'GET' || req.method === 'HEAD') && PAGE_FILES.includes(rel)) {
      fs.readFile(path.join(ROOT, rel), (err, data) => {
        if (err) { res.writeHead(404); return res.end('Not found'); }
        res.writeHead(200, { 'Content-Type': TYPES[path.extname(rel)], 'Content-Length': data.length, 'Cache-Control': 'no-cache' });
        res.end(req.method === 'HEAD' ? undefined : data);
      });
      return;
    }
    res.writeHead(404); res.end('Not found');
  });
  server.headersTimeout = 20000;
  server.requestTimeout = 60000;
  server.keepAliveTimeout = 10000;
  server.store = store;
  server.on('close', () => store.close());
  return server;
}

module.exports = { createFormsStore, createFormsServer, LIMITS };
