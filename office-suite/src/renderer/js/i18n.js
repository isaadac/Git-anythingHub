'use strict';
// Interface languages. Loaded in <head> (after theme.js) on every page. The dictionaries in
// i18n/<lang>.js map the English text of the interface to the language; this file swaps the text
// of buttons, tabs, labels, menus and dialogs, and keeps the English so the language can change
// live. Documents are never touched: anything inside the document (and any typed text) is skipped.
(function () {
  const KEY = 'atlas-lang';
  const LANGS = { en: 'English', es: 'Español', fr: 'Français', de: 'Deutsch', pt: 'Português' };
  const ATTRS = ['title', 'aria-label', 'placeholder'];
  // The document and anything typed: never translated (a text box's placeholder and tooltip are).
  const SKIP = '[data-no-i18n],[contenteditable="true"],[contenteditable=""],textarea,script,style,#doc,#canvas,#grid,#sheetTabs,#stage,#thumbs,#presentSlide,#notes,#pages,#sections,.el,.handout-page';
  const read = () => { try { const v = localStorage.getItem(KEY); return LANGS[v] ? v : 'en'; } catch { return 'en'; } };
  const dicts = (window.ISADEK_I18N = window.ISADEK_I18N || {});
  let lang = read();

  const textOrig = new WeakMap(); // text node → its English
  const attrOrig = new WeakMap(); // element → { attr: English }
  const bigOrig = new WeakMap(); // big ribbon button → [English label, had a line break]
  const dict = () => (lang === 'en' ? null : dicts[lang] || null);
  /** The language's word for an English string (or the string itself). */
  const t = s => { const d = dict(); if (!d || typeof s !== 'string') return s; const k = s.trim(); const v = d[k]; return v ? s.replace(k, v) : s; };

  // Load the dictionary before the page is parsed, so it shows in the language from the start.
  function load(l, then, early) {
    if (l === 'en' || dicts[l]) { then(); return; }
    if (early && document.readyState === 'loading') {
      document.write(`<script src="i18n/${l}.js"><\/script>`);
      then();
      return;
    }
    const s = document.createElement('script');
    s.src = `i18n/${l}.js`;
    s.onload = then;
    s.onerror = then;
    document.head.appendChild(s);
  }

  const skipped = el => !!(el && el.closest && el.closest(SKIP));
  function bigButton(btn) {
    let o = bigOrig.get(btn);
    if (!o) {
      const parts = [];
      let br = false;
      for (const n of btn.childNodes) {
        if (n.nodeType === 3 && n.nodeValue.trim()) parts.push(n.nodeValue.trim());
        else if (n.nodeName === 'BR') br = true;
      }
      if (!parts.length) return;
      o = [parts.join(' '), br];
      bigOrig.set(btn, o);
    }
    const d = dict();
    const text = (d && d[o[0]]) || o[0];
    // Same words as before? Leave the button alone (keeps the original line break).
    const cur = [...btn.childNodes].filter(n => n.nodeType === 3 && n.nodeValue.trim()).map(n => n.nodeValue.trim()).join(' ');
    if (cur === text) return;
    [...btn.childNodes].forEach(n => { if (n.nodeType === 3 || n.nodeName === 'BR') n.remove(); });
    const words = text.split(' ');
    if (o[1] && words.length > 1) {
      // Two lines, split where they come out most even.
      let best = 1, diff = Infinity;
      for (let i = 1; i < words.length; i++) { const a = words.slice(0, i).join(' ').length, b = words.slice(i).join(' ').length; if (Math.abs(a - b) < diff) { diff = Math.abs(a - b); best = i; } }
      btn.append(words.slice(0, best).join(' '), document.createElement('br'), words.slice(best).join(' '));
    } else btn.append(text);
  }
  function textNode(n) {
    let en = textOrig.get(n);
    if (en === undefined) {
      if (!n.nodeValue.trim()) return;
      en = n.nodeValue;
      textOrig.set(n, en);
    } else if (n.nodeValue !== n._i18nSet) {
      // The page changed the text since: that is the new English.
      en = n.nodeValue;
      textOrig.set(n, en);
    }
    const v = t(en);
    if (n.nodeValue !== v) n.nodeValue = v;
    n._i18nSet = v;
  }
  function attrs(el) {
    let o = attrOrig.get(el);
    for (const a of ATTRS) {
      if (!el.hasAttribute(a)) continue;
      const cur = el.getAttribute(a);
      if (!o) { o = {}; attrOrig.set(el, o); }
      if (o[a] === undefined || (cur !== o[a] && cur !== t(o[a]))) o[a] = cur;
      const v = t(o[a]);
      if (cur !== v) el.setAttribute(a, v);
    }
  }
  /** Translate a part of the page (all of it by default). */
  function translate(root = document.body) {
    if (!root) return;
    busy = true;
    try {
      const el = root.nodeType === 1 ? root : root.parentElement;
      if (root.nodeType === 3) { if (!skipped(el)) textNode(root); return; }
      if (skipped(el)) { if (el.tagName === 'TEXTAREA' && el === root) attrs(el); return; }
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
        acceptNode: n => {
          if (n.nodeType !== 1 || !n.matches(SKIP)) return NodeFilter.FILTER_ACCEPT;
          if (n.tagName === 'TEXTAREA') attrs(n);
          return NodeFilter.FILTER_REJECT;
        }
      });
      for (let n = walker.currentNode; n; n = walker.nextNode()) {
        if (n.nodeType === 1) {
          attrs(n);
          if (n.classList.contains('big') && n.classList.contains('btn')) bigButton(n);
        } else if (!(n.parentElement && n.parentElement.classList.contains('big') && n.parentElement.classList.contains('btn'))) textNode(n);
      }
    } finally { busy = false; }
  }

  // Content added later (menus, dialogs, panes) is translated as it appears.
  let busy = false, queue = new Set(), timer = 0;
  const observer = new MutationObserver(list => {
    if (busy || lang === 'en' && !document.documentElement.dataset.i18nWas) return;
    for (const m of list) {
      if (m.type === 'childList') m.addedNodes.forEach(n => queue.add(n));
      else queue.add(m.target);
    }
    // Translated before the next paint.
    if (!timer) timer = requestAnimationFrame(() => { timer = 0; const q = [...queue]; queue = new Set(); q.forEach(n => n.isConnected && translate(n)); });
  });
  function start() {
    translate();
    translate(document.querySelector('title'));
    observer.observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ATTRS });
  }

  function apply(l) {
    lang = LANGS[l] ? l : 'en';
    document.documentElement.lang = lang;
    if (lang !== 'en') document.documentElement.dataset.i18nWas = '1';
    load(lang, () => { if (document.body) { translate(); translate(document.querySelector('title')); } });
  }

  const I18n = {
    LANGS,
    get lang() { return lang; },
    t,
    translate,
    set(l) {
      if (!LANGS[l]) return;
      try { localStorage.setItem(KEY, l); } catch { /* ignore */ }
      apply(l);
      // The native menus follow too (desktop).
      if (window.atlas && window.atlas.setLanguage) window.atlas.setLanguage(l);
    },
    async choose() {
      const codes = Object.keys(LANGS);
      const i = await UI.choose(t('Language'), 'Isadek Office: ' + LANGS[lang], codes.map(c => LANGS[c]), '🌐');
      if (i >= 0) I18n.set(codes[i]);
    }
  };
  window.I18n = I18n;

  document.documentElement.lang = lang;
  load(lang, () => {}, true);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
  // Another window changed it.
  window.addEventListener('storage', e => { if (e.key === KEY) apply(read()); });
})();
