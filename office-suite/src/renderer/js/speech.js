'use strict';
// Isadek — speech to text for dictation, live captions and Rehearse with Coach.
//
// Two engines:
//   on this device — Whisper running in a worker (vendor/speech/speech-worker.js): the model is
//     downloaded from Hugging Face once, with the person's say-so, and kept; after that it works
//     offline, in the desktop app, and nothing that is said leaves the computer;
//   the browser's own — Chrome, Edge and Safari send the audio to their provider.
// The choice is remembered ('atlas-speech'). Speech.listen(onText, onError, lang) → stop().
(function () {
  const PREF = 'atlas-speech';
  const MODELS = { fast: { id: 'Xenova/whisper-tiny', size: '40 MB', label: 'Fast' }, accurate: { id: 'Xenova/whisper-base', size: '80 MB', label: 'More accurate' } };
  const get = k => { try { return localStorage.getItem(k); } catch { return null; } };
  const set = (k, v) => { try { if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch { /* ignore */ } };
  const browserSR = () => window.SpeechRecognition || window.webkitSpeechRecognition;
  const modelKey = () => (get(PREF + '-model') === 'accurate' ? 'accurate' : 'fast');

  // ---------- The on-device engine ----------
  let worker = null, ready = null, seq = 0;
  const waiting = new Map();
  function engine() {
    if (ready) return ready;
    worker = new Worker('vendor/speech/speech-worker.js', { type: 'module' });
    worker.onmessage = e => {
      const m = e.data || {};
      if (m.t === 'progress') Speech.onProgress(m);
      else if ((m.t === 'text' || m.t === 'error') && waiting.has(m.id)) { const w = waiting.get(m.id); waiting.delete(m.id); m.t === 'text' ? w.resolve(m.text) : w.reject(new Error(m.error)); }
    };
    ready = new Promise((resolve, reject) => {
      const on = e => {
        if (e.data.t === 'ready') { worker.removeEventListener('message', on); resolve(); }
        if (e.data.t === 'error' && e.data.id === undefined) { worker.removeEventListener('message', on); reject(new Error(e.data.error)); }
      };
      worker.addEventListener('message', on);
      worker.postMessage({ t: 'load', model: MODELS[modelKey()].id });
    });
    ready.catch(() => { ready = null; if (worker) worker.terminate(); worker = null; });
    return ready;
  }
  const transcribe = (audio, lang) => new Promise((resolve, reject) => {
    const id = ++seq;
    waiting.set(id, { resolve, reject });
    worker.postMessage({ t: 'run', id, audio, lang, opts: { return_timestamps: true } }, [audio.buffer]);
  });

  /**
   * Listens with the on-device engine: the microphone at 16 kHz, cut into phrases at pauses (or
   * every 12 seconds); each finished phrase is final, and while someone speaks the phrase so far
   * comes as interim text every couple of seconds.
   */
  async function deviceListen(onText, onError, lang) {
    let stopped = false, stream = null, ctx = null, node = null, feeder = null;
    const RATE = 16000;
    let phrase = [], phraseLen = 0, silent = 0, speaking = false, floor = 0.004, busy = false, lastInterim = 0;
    const stop = () => {
      stopped = true;
      if (feeder) clearInterval(feeder);
      try { node && node.disconnect(); } catch { /* gone */ }
      try { ctx && ctx.close(); } catch { /* gone */ }
      if (stream) stream.getTracks().forEach(t => t.stop());
    };
    const take = () => { const a = new Float32Array(phraseLen); let at = 0; for (const p of phrase) { a.set(p, at); at += p.length; } return a; };
    async function finish() {
      if (!phraseLen) return;
      const audio = take();
      phrase = []; phraseLen = 0; speaking = false; silent = 0;
      if (audio.length < RATE * 0.3) return;
      try { const t = await transcribe(audio, lang); if (t && !stopped) onText(t, true); } catch (err) { if (!stopped) onError && onError('Speech recognition failed: ' + err.message); }
    }
    async function interim() {
      if (busy || !speaking || phraseLen < RATE) return;
      busy = true;
      try { const t = await transcribe(take(), lang); if (t && !stopped && speaking) onText(t, false); } catch { /* the final one will say */ }
      busy = false;
    }
    function samples(buf) {
      if (stopped) return;
      let sum = 0;
      for (let i = 0; i < buf.length; i++) sum += buf[i] * buf[i];
      const rms = Math.sqrt(sum / buf.length);
      // The room's own noise is learned while nobody speaks.
      if (!speaking) floor = floor * 0.95 + rms * 0.05;
      const loud = rms > Math.max(0.01, floor * 3);
      if (loud) { speaking = true; silent = 0; } else if (speaking) silent += buf.length;
      if (speaking) { phrase.push(new Float32Array(buf)); phraseLen += buf.length; }
      if (speaking && (silent > RATE * 0.7 || phraseLen > RATE * 12)) finish();
      else if (speaking && Date.now() - lastInterim > 2000) { lastInterim = Date.now(); interim(); }
    }
    try { await engine(); } catch (err) { onError && onError('On-device speech couldn’t start: ' + err.message); return null; }
    if (stopped) return stop;
    // Tests (and recordings) can feed audio instead of the microphone.
    if (Speech.testAudio) {
      const a = Speech.testAudio, chunk = RATE / 10;
      let at = 0;
      feeder = setInterval(() => {
        if (at >= a.length + RATE) { clearInterval(feeder); return; }
        samples(at < a.length ? a.subarray(at, at + chunk) : new Float32Array(chunk));
        at += chunk;
      }, 100);
      return stop;
    }
    try {
      stream = await navigator.mediaDevices.getUserMedia({ audio: { channelCount: 1, echoCancellation: true, noiseSuppression: true } });
      ctx = new AudioContext({ sampleRate: RATE });
      const src = ctx.createMediaStreamSource(stream);
      node = ctx.createScriptProcessor(4096, 1, 1);
      node.onaudioprocess = e => samples(e.inputBuffer.getChannelData(0));
      src.connect(node); node.connect(ctx.destination);
    } catch (err) {
      stop();
      onError && onError(err && err.name === 'NotAllowedError' ? 'Listening stopped: the microphone isn’t allowed' : 'Listening stopped: no microphone was found');
      return null;
    }
    return stop;
  }

  // ---------- The browser's engine ----------
  function browserListen(onText, onError, lang) {
    const SR = browserSR();
    let stopped = false, rec = null, fails = 0;
    const start = () => {
      rec = new SR();
      rec.continuous = true;
      rec.interimResults = true;
      rec.lang = lang;
      rec.onresult = e => { fails = 0; for (let i = e.resultIndex; i < e.results.length; i++) onText(e.results[i][0].transcript, e.results[i].isFinal); };
      rec.onerror = e => {
        if (e.error === 'no-speech' || e.error === 'aborted') return;
        fails++;
        const why = { 'not-allowed': 'the microphone isn’t allowed', 'service-not-allowed': 'the microphone isn’t allowed', network: 'the speech service can’t be reached', 'audio-capture': 'no microphone was found' }[e.error] || e.error;
        if (['not-allowed', 'service-not-allowed', 'audio-capture', 'network'].includes(e.error) || fails > 3) {
          stopped = true;
          onError && onError('Listening stopped: ' + why + (e.error === 'network' ? '. Speech recognition on this device works without it (File → Speech).' : ''));
        }
      };
      // A recognizer ends itself after a pause; carry on until told to stop.
      rec.onend = () => { if (!stopped) setTimeout(() => { if (!stopped) { try { start(); } catch { /* next end */ } } }, 250); };
      try { rec.start(); } catch (err) { stopped = true; onError && onError('Listening could not start: ' + err.message); }
    };
    start();
    return () => { stopped = true; try { rec && rec.stop(); } catch { /* already stopped */ } };
  }

  // ---------- Choosing ----------
  async function hasModel() {
    if (!window.indexedDB) return false;
    try {
      const keys = await new Promise((resolve, reject) => {
        const r = indexedDB.open('isadek-speech', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('files');
        r.onsuccess = () => { const q = r.result.transaction('files').objectStore('files').getAllKeys(); q.onsuccess = () => resolve(q.result); q.onerror = () => reject(q.error); };
        r.onerror = () => reject(r.error);
      });
      return keys.some(k => String(k).includes('/' + MODELS[modelKey()].id + '/') && /encoder_model/.test(k));
    } catch { return false; }
  }
  /** The speech settings: which engine, and the on-device model. Resolves to the engine chosen (or null). */
  async function settings(reason) {
    const have = await hasModel();
    const m = MODELS[modelKey()];
    const lines = [
      reason || 'How should Isadek Office turn speech into text (dictation, live captions, Rehearse with Coach)?',
      `On this device: private and works offline. ${have ? 'The speech model is downloaded.' : `Downloads a speech model (${m.size}) from Hugging Face once.`}`,
      browserSR() ? 'Browser’s service: no download, but what you say is sent to the browser’s provider (Google for Chrome, Apple for Safari, Microsoft for Edge).' : ''
    ].filter(Boolean).join('\n\n');
    const buttons = ['On this device', ...(browserSR() ? ['Browser’s service'] : []), 'Model…', 'Cancel'];
    const i = await UI.choose('Speech recognition', lines, buttons, 'mic');
    const pick = buttons[i];
    if (pick === 'On this device') { set(PREF, 'device'); return 'device'; }
    if (pick === 'Browser’s service') { set(PREF, 'browser'); return 'browser'; }
    if (pick === 'Model…') {
      const r = await UI.form('Speech Model', [{ name: 'm', label: 'Model', type: 'select', value: modelKey(), options: Object.entries(MODELS).map(([value, x]) => ({ value, label: `${x.label} (${x.size})` })) }], 'OK');
      if (r && r.m !== modelKey()) { set(PREF + '-model', r.m); if (worker) worker.terminate(); worker = null; ready = null; }
      return settings(reason);
    }
    return null;
  }

  const Speech = {
    MODELS,
    testAudio: null,
    onProgress: () => {},
    /** Listens: onText(text, final) for what is heard. Returns stop(), or null if it couldn't start. */
    async listen(onText, onError, lang) {
      lang = lang || ({ en: 'en-US', es: 'es-ES', fr: 'fr-FR', de: 'de-DE', pt: 'pt-BR' }[(window.I18n && I18n.lang) || ''] || navigator.language || 'en-US');
      let mode = get(PREF);
      // The desktop app has no browser service to use; elsewhere it is the default until a choice is made.
      if (mode === 'browser' && !browserSR()) mode = null;
      if (!mode) mode = browserSR() && !(window.atlas && window.atlas.isDesktop) ? 'browser' : await settings('Speech recognition runs on this device here. It needs a speech model, downloaded once.');
      if (!mode) return null;
      if (mode === 'browser') return browserListen(onText, onError, lang);
      if (!(await hasModel())) {
        // Download progress, as a toast that updates.
        let last = 0;
        Speech.onProgress = p => { if (Date.now() - last > 400 && p.total) { last = Date.now(); UI.toast(`Downloading the speech model… ${Math.round(100 * p.loaded / p.total)}% (${p.file.split('/').pop()})`, 2000); } };
      }
      return deviceListen(onText, onError, lang.slice(0, 2));
    },
    /** Like listen, but returns stop() at once (for callers that can't wait). onEnd: when listening ends by itself. */
    start(onText, onError, lang, onEnd) {
      let stop = null, cancelled = false, ended = false;
      const end = () => { if (!ended) { ended = true; onEnd && onEnd(); } };
      Speech.listen(onText, msg => { onError && onError(msg); end(); }, lang).then(s => { if (!s) end(); else if (cancelled) s(); else stop = s; }, () => end());
      return () => { cancelled = true; if (stop) stop(); end(); };
    },
    settings,
    hasModel,
    /** Forgets the downloaded model (it is downloaded again when next needed). */
    async removeModel() { if (worker) worker.terminate(); worker = null; ready = null; await new Promise(r => { const q = indexedDB.deleteDatabase('isadek-speech'); q.onsuccess = q.onerror = q.onblocked = () => r(); }); }
  };
  window.Speech = Speech;
  if (window.Platform && Platform.onMenu) Platform.onMenu(cmd => { if (cmd === 'speech-settings') settings(); });
})();
