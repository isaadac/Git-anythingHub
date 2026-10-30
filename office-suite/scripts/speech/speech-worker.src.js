// Isadek Office — on-device speech recognition (bundled by scripts/build-speech.js into
// src/renderer/vendor/speech/speech-worker.js).
//
// Whisper (OpenAI's open speech model) runs here with Transformers.js and ONNX Runtime on
// WebAssembly: audio never leaves the computer. The model is downloaded from Hugging Face once,
// when the person turns on-device speech on, and kept in IndexedDB for good (works offline after).
//
// Messages in:  { t: 'load', model }            → { t: 'progress', file, loaded, total } … { t: 'ready' } | { t: 'error', error }
//               { t: 'run', id, audio, lang }    → { t: 'text', id, text } | { t: 'error', id, error }
//               { t: 'cached', model }           → { t: 'cached', yes }
import { pipeline, env } from '@huggingface/transformers';

// ---------- Where the model is kept: IndexedDB, under its download address ----------
const DB = 'isadek-speech';
const db = () => new Promise((resolve, reject) => {
  const r = indexedDB.open(DB, 1);
  r.onupgradeneeded = () => r.result.createObjectStore('files');
  r.onsuccess = () => resolve(r.result);
  r.onerror = () => reject(r.error);
});
const idb = async (mode, fn) => { const d = await db(); return new Promise((resolve, reject) => { const t = d.transaction('files', mode); const q = fn(t.objectStore('files')); t.oncomplete = () => resolve(q && q.result); t.onerror = () => reject(t.error); }); };
const keyOf = req => (typeof req === 'string' ? req : req.url);
env.useBrowserCache = false;
env.allowLocalModels = false;
env.useCustomCache = true;
env.customCache = {
  async match(req) {
    const hit = await idb('readonly', st => st.get(keyOf(req)));
    return hit ? new Response(hit.body, { headers: hit.headers }) : undefined;
  },
  async put(req, res) {
    const body = await res.clone().arrayBuffer();
    const headers = {};
    res.headers.forEach((v, k) => { headers[k] = v; });
    await idb('readwrite', st => st.put({ body, headers }, keyOf(req)));
  }
};
// ONNX Runtime's WebAssembly comes with the app (never from a CDN), single-threaded (no cross-origin isolation needed).
env.backends.onnx.wasm.wasmPaths = { wasm: new URL('./ort-wasm-simd-threaded.jsep.wasm', import.meta.url).href, mjs: new URL('./ort-wasm-simd-threaded.jsep.mjs', import.meta.url).href };
env.backends.onnx.wasm.numThreads = 1;
env.backends.onnx.wasm.proxy = false;

let asr = null, loading = null;
const LANGS = { en: 'english', es: 'spanish', fr: 'french', de: 'german', pt: 'portuguese', it: 'italian', nl: 'dutch' };

self.onmessage = async e => {
  const m = e.data || {};
  if (m.t === 'load') {
    try {
      if (!loading) loading = pipeline('automatic-speech-recognition', m.model, {
        dtype: 'q8', device: 'wasm',
        progress_callback: p => { if (p.status === 'progress') self.postMessage({ t: 'progress', file: p.file, loaded: p.loaded, total: p.total }); }
      });
      asr = await loading;
      self.postMessage({ t: 'ready' });
    } catch (err) { loading = null; self.postMessage({ t: 'error', error: String(err && err.message || err) }); }
    return;
  }
  if (m.t === 'cached') {
    const keys = await idb('readonly', st => st.getAllKeys()).catch(() => []);
    self.postMessage({ t: 'cached', yes: (keys || []).some(k => String(k).includes('/' + m.model + '/') && /encoder_model/.test(k)) });
    return;
  }
  if (m.t === 'run') {
    try {
      if (!asr) throw new Error('not ready');
      const lang = LANGS[String(m.lang || 'en').slice(0, 2)] || 'english';
      const out = await asr(m.audio, { language: lang, task: 'transcribe', ...(m.opts || {}) });
      self.postMessage({ t: 'text', id: m.id, text: String(out && out.text || '').trim() });
    } catch (err) { self.postMessage({ t: 'error', id: m.id, error: String(err && err.message || err) }); }
  }
};
