'use strict';
// Shared by the workers that run code from documents or add-ins: lockDown() removes everything
// that reaches outside the worker (network, storage, other pages), from the worker's global object
// and from the prototypes behind it. The browser's own APIs live on those prototypes too
// (WorkerGlobalScope.prototype), so removing them from `self` alone is not enough: their getters
// could still be called with `self`.
self.lockDown = function lockDown() {
  const OUTSIDE = ['fetch', 'XMLHttpRequest', 'WebSocket', 'WebSocketStream', 'WebTransport', 'EventSource', 'importScripts', 'indexedDB', 'caches',
    'Worker', 'SharedWorker', 'BroadcastChannel', 'Request', 'navigator', 'RTCPeerConnection', 'RTCDataChannel', 'FileSystemHandle', 'storageFoundation'];
  for (let o = self; o && o !== Object.prototype; o = Object.getPrototypeOf(o)) {
    for (const k of OUTSIDE) {
      if (o !== self && !Object.prototype.hasOwnProperty.call(o, k)) continue;
      try { Object.defineProperty(o, k, { value: undefined, configurable: false, writable: false }); } catch { try { delete o[k]; } catch { /* nothing more to do */ } }
    }
  }
  delete self.lockDown;
};
