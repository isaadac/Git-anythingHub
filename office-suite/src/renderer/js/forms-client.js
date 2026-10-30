'use strict';
// Isadek Forms — talking to where answers are collected: the Isadek Office web server the app
// runs on (or one the person names), or, in the desktop app, this computer ("local": the main
// process's store, reached over IPC; people on the network reach its small server). Used by
// Sheets (forms) and Slides (polls, audience Q&A and the phone remote).
(function () {
  const MINE = 'atlas-forms-mine';
  const SERVER_OK = /^https:\/\/[^\s/?#]+(:\d+)?$|^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/;
  const getMine = () => { try { return JSON.parse(localStorage.getItem(MINE) || '[]'); } catch { return []; } };
  // Forms published from this device: their answers are fetched without asking.
  const isMine = f => getMine().includes(f.server + '|' + f.id);
  const addMine = f => { try { localStorage.setItem(MINE, JSON.stringify([...new Set([...getMine(), f.server + '|' + f.id])].slice(-500))); } catch { /* private mode */ } };
  const desktop = () => !!(window.atlas && window.atlas.formsHost);

  async function http(server, method, path, key, body, signal) {
    let res;
    try {
      res = await fetch(server + path, {
        method, credentials: 'omit', cache: 'no-store', signal,
        headers: { ...(key ? { Authorization: 'Bearer ' + key } : {}), ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}) },
        body: body !== undefined ? JSON.stringify(body) : undefined
      });
    } catch (err) {
      if (signal && signal.aborted) throw err;
      throw new Error(`${new URL(server).host} can’t be reached`);
    }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw Object.assign(new Error(data.error || `The server answered ${res.status}`), { status: res.status, errors: data.errors });
    return data;
  }
  async function local(op, args) {
    if (!desktop()) throw new Error('This collects answers in the Isadek Office desktop app');
    const r = await window.atlas.formsCall(op, args);
    if (r && r.error) throw Object.assign(new Error(r.error), { status: r.status, errors: r.errors });
    return r;
  }

  /** A form's operations: create, update, remove, respond, responses. f: { server, id, key }. */
  function call(f, op, args = {}) {
    if (f.server === 'local') return local(op, { id: f.id, key: f.key, ...args });
    const base = '/api/forms' + (f.id ? '/' + f.id : '');
    switch (op) {
      case 'create': return http(f.server, 'POST', '/api/forms', null, { form: args.form });
      case 'update': return http(f.server, 'PUT', base, f.key, { form: args.form, open: args.open });
      case 'remove': return http(f.server, 'DELETE', base, f.key);
      case 'respond': return http(f.server, 'POST', base + '/responses', null, { answers: args.answers });
      case 'responses': return http(f.server, 'GET', `${base}/responses?after=${args.after || 0}&limit=${args.limit || 1000}`, f.key);
    }
    return Promise.reject(new Error('Unknown request'));
  }

  /** The phone remote's operations. r: { server, id, key }. */
  function remote(r, op, args = {}) {
    if (r.server === 'local') return local('remote-' + op, { id: r.id, key: r.key, ...args });
    const base = '/api/remote' + (r.id ? '/' + r.id : '');
    switch (op) {
      case 'create': return http(r.server, 'POST', '/api/remote', null, {});
      case 'state': return http(r.server, 'PUT', base + '/state', r.key, args.state);
      case 'commands': return http(r.server, 'GET', `${base}/commands?after=${args.after || 0}`, r.key, undefined, args.signal);
      case 'remove': return http(r.server, 'DELETE', base, r.key);
    }
    return Promise.reject(new Error('Unknown request'));
  }

  /** The server the page came from, when it collects forms. */
  async function serverHere() {
    if (!/^https?:$/.test(location.protocol) || (window.Platform && Platform.native)) return null;
    try { const r = await fetch('/api/forms', { cache: 'no-store', credentials: 'omit' }); return r.ok && (await r.json()).available ? location.origin : null; } catch { return null; }
  }
  /** Starts this computer's server for the local network (desktop): its addresses. */
  async function hostUrls() {
    const r = await window.atlas.formsHost();
    if (r.error) throw new Error('Couldn’t start collecting answers on this computer: ' + r.error);
    return r.urls;
  }
  async function base(server) { return server === 'local' ? (await hostUrls())[0] : server; }
  async function linkOf(f) { return f.id ? (await base(f.server)) + '/f/' + f.id : ''; }

  /**
   * Where to collect: the page's own server if it can, otherwise the person picks this computer
   * (desktop) or a server. what: 'answer' (forms, polls, questions) or 'remote'.
   */
  async function chooseServer(title, what = 'answer') {
    const here = await serverHere();
    if (here) return here;
    const doing = what === 'remote' ? 'your phone controls the show' : 'people answer';
    const buttons = [...(desktop() ? ['This computer'] : []), 'A server…', 'Cancel'];
    const i = await UI.choose(title, (desktop() ? `This computer: ${doing} over the same network (Wi-Fi) while Isadek Office is open here.\n\n` : '') + `A server: an Isadek Office web server${what === 'remote' ? '' : ' (anyone with the link can answer, any time)'}.`, buttons, 'send');
    if (buttons[i] === 'This computer') return 'local';
    if (buttons[i] !== 'A server…') return null;
    const url = await UI.prompt('Isadek Office Server', 'Address of the Isadek Office server (https://…)', 'https://');
    if (!url) return null;
    let u;
    try { u = new URL(url.trim()); } catch { UI.toast('Enter an address such as https://office.example.com', 4000); return null; }
    const server = u.origin;
    if (!SERVER_OK.test(server)) { UI.toast('The server must use https', 4000); return null; }
    try { const r = await fetch(server + '/api/forms', { cache: 'no-store', credentials: 'omit' }); if (!r.ok) throw 0; } catch { UI.toast(`${u.host} doesn’t collect answers (or can’t be reached)`, 4000); return null; }
    return server;
  }

  /** A saved { server, id, key } that can be called, or null. */
  function cleanRef(v) {
    if (!v || typeof v !== 'object') return null;
    if (!(v.server === 'local' || (typeof v.server === 'string' && SERVER_OK.test(v.server)))) return null;
    if (!/^[A-Za-z0-9_-]{22}$/.test(v.id) || !/^[A-Za-z0-9_-]{20,100}$/.test(v.key)) return null;
    return { server: v.server, id: v.id, key: v.key };
  }

  window.FormsClient = { SERVER_OK, getMine, isMine, addMine, call, remote, serverHere, hostUrls, base, linkOf, chooseServer, cleanRef };
})();
