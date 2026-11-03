'use strict';
// Runs one add-in away from the page (addins.js starts one worker per add-in). The add-in gets
// the `isadek` API below and nothing else: no network, storage, page or other add-ins. Every
// call to the document goes to the page, which checks it against the add-in's permissions.
importScripts('worker-lockdown.js');
self.lockDown();
const post = self.postMessage.bind(self);
let seq = 0;
const waiting = new Map();
const commands = new Map(), actions = new Map();
let manifest = null;
const call = (method, ...args) => new Promise((resolve, reject) => {
  const id = ++seq;
  waiting.set(id, { resolve, reject });
  post({ t: 'call', id, method, args });
});
/** The API an add-in sees. */
const isadek = Object.freeze({
  get app() { return manifest.app; },
  get addin() { return { id: manifest.id, name: manifest.name, version: manifest.version }; },
  /** Runs fn when the ribbon button with this id is clicked. */
  onCommand(id, fn) { if (typeof fn === 'function') commands.set(String(id), fn); },
  /** Runs fn when a button in the add-in's pane is clicked: fn(values of its inputs). */
  onAction(id, fn) { if (typeof fn === 'function') actions.set(String(id), fn); },
  document: Object.freeze({
    getText: () => call('getText'),
    getSelection: () => call('getSelection'),
    insertText: text => call('insertText', String(text)),
    // Sheets
    getValues: address => call('getValues', String(address)),
    setValues: (address, values) => call('setValues', String(address), values),
    getSelectedRange: () => call('getSelectedRange'),
    // Slides
    getSlides: () => call('getSlides'),
    addSlide: (title, bullets) => call('addSlide', String(title || ''), Array.isArray(bullets) ? bullets.map(String) : [])
  }),
  ui: Object.freeze({
    toast: text => call('toast', String(text)),
    /** A pane beside the document: blocks of { type: 'heading'|'text'|'list'|'input'|'button'|'progress', … }. */
    showPane: (title, blocks) => call('showPane', String(title || ''), Array.isArray(blocks) ? blocks : []),
    closePane: () => call('closePane'),
    /** A dialog: fields { name, label, type: 'text'|'number'|'textarea'|'select', value, options } → values or null. */
    form: (title, fields, ok) => call('form', String(title || ''), Array.isArray(fields) ? fields : [], ok ? String(ok) : 'OK')
  })
});

self.onmessage = async e => {
  const m = e.data || {};
  if (m.t === 'load') {
    manifest = m.manifest;
    try {
      // The add-in's code sees `isadek` and the usual JavaScript built-ins.
      const run = new Function('isadek', '"use strict";\n' + m.code);
      await run(isadek);
      post({ t: 'ready' });
    } catch (err) { post({ t: 'failed', error: String(err && err.message || err) }); }
    return;
  }
  if (m.t === 'result' && waiting.has(m.id)) {
    const w = waiting.get(m.id); waiting.delete(m.id);
    if (m.error) w.reject(new Error(m.error)); else w.resolve(m.value);
    return;
  }
  if (m.t === 'command' || m.t === 'action') {
    const fn = (m.t === 'command' ? commands : actions).get(String(m.id));
    try {
      if (!fn) throw new Error(`Nothing handles “${m.id}”`);
      await fn(m.values || {});
      post({ t: 'done', run: m.run });
    } catch (err) { post({ t: 'done', run: m.run, error: String(err && err.message || err) }); }
  }
};
