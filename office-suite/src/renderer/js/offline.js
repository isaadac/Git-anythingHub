'use strict';
// Isadek — working offline with cloud documents.
//
// A save to the cloud that can't reach it (no connection, or the connection drops) is kept on this
// device in an outbox and uploaded as soon as the connection is back — when the app next runs, if
// need be. If someone else saved the file in the meantime, nothing of theirs is overwritten: this
// device's version goes up next to it as a "conflicted copy". Cloud files are also kept on the
// device when opened, so recent ones open without a connection. Both live in IndexedDB.
(function () {
  const DB = 'isadek-offline', KEEP_MAX = 30, KEEP_BYTES = 25e6;
  let dbp = null;
  const db = () => dbp || (dbp = new Promise((resolve, reject) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => { r.result.createObjectStore('outbox', { keyPath: 'path' }); r.result.createObjectStore('kept', { keyPath: 'path' }); };
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  }));
  async function tx(store, mode, fn) {
    const d = await db();
    return new Promise((resolve, reject) => {
      const t = d.transaction(store, mode), st = t.objectStore(store);
      let out;
      Promise.resolve(fn(st, v => { out = v; })).catch(reject);
      t.oncomplete = () => resolve(out);
      t.onerror = () => reject(t.error);
    });
  }
  const req = r => new Promise((resolve, reject) => { r.onsuccess = () => resolve(r.result); r.onerror = () => reject(r.error); });
  const all = store => tx(store, 'readonly', async (st, done) => done(await req(st.getAll())));

  /** True for a failure to reach the service at all (not a refusal from it). */
  const unreachable = err => !navigator.onLine || err instanceof TypeError || /net::|network|failed to fetch|load failed|ENOTFOUND|ECONN|ETIMEDOUT|EAI_AGAIN|offline/i.test(String(err && err.message || err));

  const listeners = new Set();
  let state = { waiting: 0, busy: false, offline: !navigator.onLine };
  const tell = () => listeners.forEach(f => { try { f(state); } catch { /* a listener's problem */ } });
  async function refresh() { state = { ...state, waiting: (await all('outbox')).length, offline: !navigator.onLine }; tell(); }

  const Offline = {
    unreachable,
    onChange(f) { listeners.add(f); f(state); return () => listeners.delete(f); },
    get state() { return state; },
    /** Keeps a save for the cloud on this device until it can go up; returns what Platform.save returns. */
    async queue(path, name, data) {
      const p = Cloud.parsePath(path);
      await tx('outbox', 'readwrite', st => st.put({ path, name, data: new Uint8Array(data), savedAt: Date.now(), base: Cloud.revisionSeen(path), service: p.acct }));
      await Offline.keep(path, name, data);
      await refresh();
      return { name, path, queued: true };
    },
    /** Keeps a copy of a cloud file on this device (the latest few). */
    async keep(path, name, data) {
      if (!data || data.length > KEEP_BYTES) return;
      try {
        await tx('kept', 'readwrite', async st => {
          st.put({ path, name, data: new Uint8Array(data), at: Date.now() });
          const list = (await req(st.getAll())).sort((a, b) => b.at - a.at);
          list.slice(KEEP_MAX).forEach(x => st.delete(x.path));
        });
      } catch { /* storage full or unavailable: only the offline copy is lost */ }
    },
    /** This device's newest copy of a cloud file: { name, data, waiting } or null. */
    async copy(path) {
      const [w] = await tx('outbox', 'readonly', async (st, done) => done([await req(st.get(path))]));
      if (w) return { name: w.name, data: w.data, waiting: true };
      const k = await tx('kept', 'readonly', async (st, done) => done(await req(st.get(path))));
      return k ? { name: k.name, data: k.data, waiting: false } : null;
    },
    pending: () => all('outbox'),
    /** Forgets a waiting save (a newer one went straight up). */
    async drop(path) { try { await tx('outbox', 'readwrite', st => st.delete(path)); await refresh(); } catch { /* nothing waiting */ } },
    /** Uploads what is waiting. Returns how many went up. */
    async flush() {
      if (state.busy || !navigator.onLine) return 0;
      const list = await all('outbox');
      if (!list.length) return 0;
      state = { ...state, busy: true }; tell();
      let sent = 0, conflicts = 0;
      try {
        for (const w of list.sort((a, b) => a.savedAt - b.savedAt)) {
          try {
            let current = '';
            try { current = await Cloud.stat(w.path); } catch (err) { if (unreachable(err)) break; }
            if (w.base && current && current !== w.base) {
              // Changed in the cloud since this device last had it: keep both.
              const p = Cloud.parsePath(w.path), ext = Platform.extOf(w.name), stem = Platform.baseName(w.name);
              const when = new Date(w.savedAt).toISOString().slice(0, 16).replace('T', ' ').replace(':', '.');
              const copyName = `${stem} (conflicted copy ${when})${ext ? '.' + ext : ''}`;
              await Cloud.upload(p.acct, p.folder, copyName, w.data);
              UI.toast(`“${w.name}” was changed in the cloud while you were offline. Your version was saved next to it as “${copyName}”.`, 8000);
              conflicts++;
            } else {
              await Cloud.saveTo(w.path, w.name, w.data);
            }
            await tx('outbox', 'readwrite', st => st.delete(w.path));
            sent++;
          } catch (err) {
            if (unreachable(err)) break;
            // The service refused it (signed out, no space…): say so, and keep it for later.
            UI.toast(`Couldn't upload “${w.name}”: ${err.message}. It is still saved on this device.`, 6000);
            break;
          }
        }
      } finally {
        state = { ...state, busy: false };
        await refresh();
      }
      if (sent && !conflicts) UI.toast(`${sent} document${sent === 1 ? '' : 's'} saved on this device went up to the cloud`, 4000);
      return sent;
    }
  };

  // Try again whenever the connection comes back, now and then while something waits, and when the
  // app starts.
  addEventListener('online', () => { refresh(); Offline.flush(); });
  addEventListener('offline', refresh);
  setInterval(() => { if (state.waiting && navigator.onLine) Offline.flush(); }, 60000);
  if (window.indexedDB) setTimeout(() => refresh().then(() => Offline.flush()).catch(() => {}), 1500);

  window.Offline = Offline;
})();
