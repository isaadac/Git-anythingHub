'use strict';
// Isadek on the web, offline: keeps the app's own files (pages, scripts, styles, fonts, icons) so
// the apps open without a connection. Online, every request goes to the server first, so updates
// show at once; offline, the kept copy answers. Documents themselves are never kept here (see
// offline.js), and nothing from other sites or the server's API passes through.
const CACHE = 'isadek-app-v1';
const PAGES = ['index.html', 'write.html', 'sheets.html', 'slides.html', 'notes.html', 'help.html'];
const SCOPE = new URL('./', self.location).href;

/** Every page, and the files they and their style sheets load, fetched once when installed. */
async function precache() {
  const cache = await caches.open(CACHE);
  const seen = new Set();
  const take = async url => {
    if (seen.has(url) || new URL(url).origin !== self.location.origin) return null;
    seen.add(url);
    try {
      const r = await fetch(url, { cache: 'no-cache' });
      if (!r.ok) return null;
      await cache.put(url, r.clone());
      return r;
    } catch { return null; }
  };
  for (const page of PAGES) {
    const url = new URL(page, SCOPE).href;
    const r = await take(url);
    if (!r) continue;
    const html = await r.text();
    const refs = [...html.matchAll(/(?:src|href)="([^"#:?]+\.(?:js|css|png|svg|ico|woff2|json))"/g)].map(m => new URL(m[1], url).href);
    await Promise.all(refs.map(async ref => {
      const res = await take(ref);
      // Fonts and pictures named in style sheets.
      if (res && ref.endsWith('.css')) {
        const css = await res.text();
        await Promise.all([...css.matchAll(/url\(["']?([^"')#?:]+)["']?\)/g)].map(m => take(new URL(m[1], ref).href)));
      }
    }));
  }
}

self.addEventListener('install', e => { self.skipWaiting(); e.waitUntil(precache()); });
self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
    await self.clients.claim();
  })());
});
self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin || url.pathname.startsWith('/api/') || url.pathname.startsWith('/mock/')) return;
  if (!/\/(src\/renderer|node_modules)\//.test(url.pathname)) return;
  e.respondWith((async () => {
    try {
      const r = await fetch(req);
      if (r.ok && r.type === 'basic') { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {}); }
      return r;
    } catch (err) {
      const kept = await caches.match(req, { ignoreSearch: true });
      if (kept) return kept;
      if (req.mode === 'navigate') { const home = await caches.match(new URL('index.html', SCOPE).href); if (home) return home; }
      throw err;
    }
  })());
});
