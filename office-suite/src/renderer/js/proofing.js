'use strict';
// Proofing: the Spelling & Grammar pane (F7) and the document language.
// Spelling uses the desktop app's spellchecker (the same dictionaries that draw the red
// squiggles); grammar is a set of rules that catch common slips. Each app hands over the
// text it holds as DOM roots, so replacements land in the right place.
(function () {
  const BLOCK = /^(P|H[1-6]|LI|TD|TH|DIV|BLOCKQUOTE|PRE|DT|DD|FIGCAPTION|CAPTION|SECTION|ARTICLE|HEADER|FOOTER)$/;
  const LANGS = ['en-US', 'en-GB', 'en-AU', 'en-CA', 'fr-FR', 'de-DE', 'es-ES', 'es-MX', 'it-IT', 'pt-BR', 'pt-PT', 'nl-NL', 'sv-SE', 'da-DK', 'nb-NO', 'fi-FI', 'pl-PL', 'cs-CZ', 'ru-RU', 'uk-UA', 'tr-TR', 'el-GR', 'hu-HU', 'ro-RO'];
  let names = null;
  function langName(code) {
    if (!code) return 'Default language';
    try { names = names || new Intl.DisplayNames([navigator.language || 'en'], { type: 'language' }); return names.of(code); } catch { return code; }
  }

  // ---------- Grammar ----------
  const AN_EXCEPT = /^(uni(?!nt|mp|nh|nd)|use|usu|ur[aeiou]|eu|one|once|ewe|u[bcfgklmnpqrstvz][aeiou])/i; // "a university", "a one-off"
  const A_EXCEPT = /^(hour|honest|honou?r|heir|herb\b|mba|fbi|faq|html|http|sql|svg|mp3|x-?ray|nda|llc|rsvp|f\b|h\b|l\b|m\b|n\b|r\b|s\b|x\b)/i; // "an hour", "an FBI agent"
  const ABBREV = /\b(e\.g|i\.e|etc|vs|mr|mrs|ms|dr|prof|st|no|fig|approx|inc|ltd|jr|sr|al|cf|p|pp|vol|ch|sec|min|max)\.$/i;
  const CONFUSED = [
    [/\b(could|would|should|must|might) of\b/gi, m => m.replace(/ of$/i, ' have'), 'Use “have” after could, would, should.'],
    [/\balot\b/gi, () => 'a lot', '“A lot” is two words.'],
    [/\birregardless\b/gi, () => 'regardless', 'Use “regardless”.'],
    [/\b(more|less|better|worse|bigger|smaller|greater|fewer|higher|lower|rather|other) then\b/gi, m => m.replace(/then$/i, 'than'), 'Use “than” for comparisons.'],
    [/\byour welcome\b/gi, () => 'you’re welcome', 'Did you mean “you’re” (you are)?'],
    [/\bits a\b/gi, m => m[0] + 't’s a', 'Did you mean “it’s” (it is)?'],
    [/\bshould of\b/gi, () => 'should have', 'Use “have”.'],
    [/\bloose (?=(weight|the game|my|your|his|her|their|our))/gi, () => 'lose ', 'Did you mean “lose”?'],
    [/\bdefinately\b/gi, () => 'definitely', 'Spelling: “definitely”.'],
    [/\bseperate\b/gi, () => 'separate', 'Spelling: “separate”.'],
    [/\brecieve\b/gi, () => 'receive', 'Spelling: “receive”.'],
    [/\bteh\b/gi, m => (m[0] === 'T' ? 'The' : 'the'), 'Spelling: “the”.']
  ];

  /** Grammar and style issues in a piece of text: [{start, end, message, suggestions, kind}]. */
  function grammar(text, lang) {
    const out = [];
    const add = (start, end, message, suggestions, rule) => out.push({ start, end, message, suggestions: suggestions.filter(s => s !== null && s !== undefined), kind: 'grammar', rule, word: text.slice(start, end) });
    const en = !lang || /^en/i.test(lang);
    let m;
    // Repeated words ("the the"), except the few that may repeat.
    const rep = /\b([\p{L}']+)(\s+)\1\b/giu;
    while ((m = rep.exec(text))) if (!/^(had|that|very|no|bye|ha|so)$/i.test(m[1])) add(m.index, m.index + m[0].length, `“${m[1]}” is repeated.`, [m[1]], 'repeat');
    // Spaces
    const sp = /(\S)([ \u00a0]{2,})(?=\S)/g; // typed double spaces arrive as space + no-break space
    while ((m = sp.exec(text))) add(m.index + 1, m.index + 1 + m[2].length, 'More than one space between words.', [' '], 'spaces');
    const before = /(\S)(\s+)([,.;:!?])(?![\w.])/g;
    while ((m = before.exec(text))) if (!(m[3] === '.' && /\d$/.test(m[1])) && !(lang && /^fr/i.test(lang) && /[;:!?]/.test(m[3]))) add(m.index + 1, m.index + 1 + m[2].length, 'No space before punctuation.', [''], 'space-before');
    const after = /([a-zß-ÿ]{2,})([,;])(?=[A-Za-zÀ-ÿ])/g;
    while ((m = after.exec(text))) add(m.index + m[1].length, m.index + m[1].length + 1, 'Add a space after the punctuation.', [m[2] + ' '], 'space-after');
    // A sentence should start with a capital letter.
    const cap = /([.!?])(\s+)([a-zß-ÿ])/g;
    while ((m = cap.exec(text))) {
      const lead = text.slice(Math.max(0, m.index - 8), m.index + 1);
      if (m[1] === '.' && (ABBREV.test(lead) || /\.\.$/.test(lead) || /\d\.$/.test(lead) || /\b\w\.$/.test(lead))) continue;
      const i = m.index + 1 + m[2].length;
      add(i, i + 1, 'Start a sentence with a capital letter.', [m[3].toUpperCase()], 'capital');
    }
    const first = /^\s*([a-zß-ÿ])/.exec(text);
    if (first && text.trim().length > 25 && /[.!?]/.test(text) && !/^\s*(e\.g|i\.e|etc)\b/i.test(text)) add(first[0].length - 1, first[0].length, 'Start a sentence with a capital letter.', [first[1].toUpperCase()], 'capital');
    if (en) {
      const pron = /(^|[^\p{L}'’-])(i)(?=[\s,.!?;:'’]|$)/gmu;
      while ((m = pron.exec(text))) { const i = m.index + m[1].length; if (text[i + 1] !== '.' && text[i - 1] !== '.') add(i, i + 1, 'Write “I” as a capital letter.', ['I'], 'pronoun-i'); }
      const a = /\b([Aa])\s+([aeiouAEIOU][\p{L}-]*)/gu;
      while ((m = a.exec(text))) if (!AN_EXCEPT.test(m[2])) add(m.index, m.index + 1, `Use “an” before “${m[2]}”.`, [m[1] === 'A' ? 'An' : 'an'], 'a-an');
      const an = /\b([Aa]n)\s+([b-df-hj-np-tv-zB-DF-HJ-NP-TV-Z][\p{L}-]*)/gu;
      while ((m = an.exec(text))) if (!A_EXCEPT.test(m[2])) add(m.index, m.index + 2, `Use “a” before “${m[2]}”.`, [m[1] === 'An' ? 'A' : 'a'], 'a-an');
      for (const [re, fix, msg] of CONFUSED) {
        re.lastIndex = 0;
        while ((m = re.exec(text))) add(m.index, m.index + m[0].length, msg, [fix(m[0])], 'confused');
      }
    }
    // Very long sentences
    for (const s of text.matchAll(/[^.!?]+[.!?]*/g)) {
      const words = s[0].match(/\S+/g) || [];
      if (words.length > 45) { const lead = s[0].search(/\S/); add(s.index + lead, s.index + s[0].trimEnd().length, `This sentence has ${words.length} words. Consider splitting it.`, [], 'long'); }
    }
    // Overlaps: keep the first issue at any position.
    out.sort((x, y) => x.start - y.start || y.end - x.end);
    return out.filter((x, i) => !out.slice(0, i).some(y => y.start < x.end && x.start < y.end && y.rule !== 'long' && x.rule !== 'long'));
  }

  // ---------- Text from DOM roots ----------
  /** Splits a root into blocks of text, each with the text nodes it came from. */
  function segments(root) {
    const segs = [];
    let cur = null, curBlock;
    const blockOf = n => { let e = n.parentElement; while (e && e !== root && !BLOCK.test(e.tagName)) e = e.parentElement; return e || root; };
    const walk = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: n => (n.parentElement && n.parentElement.closest('code, pre, .toc, .mergefield, .citation, .bibliography, sup.fn-ref, script, style, [contenteditable="false"]') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT)
    });
    let n;
    while ((n = walk.nextNode())) {
      const b = blockOf(n);
      if (!cur || b !== curBlock) { cur = { text: '', parts: [] }; segs.push(cur); curBlock = b; }
      cur.parts.push({ node: n, start: cur.text.length });
      cur.text += n.data;
    }
    return segs.filter(s => s.text.trim());
  }
  function point(seg, off) {
    let p = seg.parts[0];
    for (const q of seg.parts) { if (q.start <= off) p = q; else break; }
    return [p.node, Math.min(off - p.start, p.node.length)];
  }
  function rangeFor(seg, start, end) {
    const r = document.createRange();
    const a = point(seg, start), b = point(seg, end);
    r.setStart(a[0], a[1]); r.setEnd(b[0], b[1]);
    return r;
  }

  // ---------- Spelling ----------
  const WORD = /[\p{L}\p{M}][\p{L}\p{M}'’-]*[\p{L}\p{M}]|[\p{L}\p{M}]/gu;
  const canSpell = () => !!(window.atlas && window.atlas.spellCheck);
  function spelling(seg, ignore) {
    const words = [];
    for (const m of seg.text.matchAll(WORD)) {
      const w = m[0].replace(/['’]s$/i, '');
      if (w.length < 2 || /\d/.test(w) || (w.length > 1 && w === w.toUpperCase()) || ignore.has(w.toLowerCase())) continue;
      // Skip words in addresses and file names.
      const around = seg.text.slice(Math.max(0, m.index - 1), m.index + m[0].length + 1);
      if (/[@/\\_]|\.\w/.test(around)) continue;
      words.push({ w, start: m.index, end: m.index + w.length });
    }
    if (!words.length) return [];
    const bad = window.atlas.spellCheck(words.map(x => x.w));
    return words.filter((_, i) => bad[i]).map(x => ({ start: x.start, end: x.end, kind: 'spelling', word: x.w, message: `“${x.w}” isn’t in the dictionary.`, suggestions: null }));
  }

  // ---------- Pane ----------
  /**
   * attach({ app, roots() -> [{ el, live?, commit?(), focus?() }], lang() -> code, setLang(code) })
   * live roots are edited in place in the page (Write); others are copies the app writes back.
   */
  function attach(o) {
    const pane = document.createElement('aside');
    pane.className = 'ai-pane proof-pane';
    pane.hidden = true;
    pane.setAttribute('aria-label', 'Spelling and grammar');
    pane.innerHTML = `<header><b>Editor</b><span class="sp"></span><button class="lang" title="Proofing language"></button><button class="x" title="Close"><span class="ico" data-icon="x"></span></button></header>
      <div class="proof-body"></div>`;
    document.body.appendChild(pane);
    const body = pane.querySelector('.proof-body');
    const langBtn = pane.querySelector('.lang');
    pane.querySelector('.x').onclick = () => close();
    langBtn.onclick = () => languageDialog();
    let issues = [], at = 0, ignoreAll = new Set(), ignoreOnce = new Set(), segs = [];

    function scan() {
      segs = [];
      issues = [];
      const lang = o.lang();
      for (const root of o.roots()) {
        for (const seg of segments(root.el)) {
          seg.root = root;
          const idx = segs.push(seg) - 1;
          const found = [...(canSpell() ? spelling(seg, ignoreAll) : []), ...grammar(seg.text, lang)]
            .filter(x => !ignoreOnce.has(onceKey(seg, x)) && !(x.kind === 'grammar' && ignoreAll.has('rule:' + x.rule + ':' + x.word.toLowerCase())));
          found.sort((a, b) => a.start - b.start);
          // Spelling wins where a grammar rule covers the same word.
          found.filter((x, i) => !found.slice(0, i).some(y => y.start < x.end && x.start < y.end && y.kind === 'spelling' && x.rule !== 'long'))
            .forEach(x => issues.push({ ...x, seg: idx }));
        }
      }
    }
    const onceKey = (seg, x) => `${x.word}|${seg.text.slice(Math.max(0, x.start - 20), x.end + 20)}`;

    function open() {
      document.querySelectorAll('.ai-pane').forEach(p => { if (p !== pane) p.hidden = true; });
      pane.hidden = false;
      document.body.classList.add('ai-open');
      langBtn.innerHTML = '<span class="ico" data-icon="languages"></span>'; langBtn.append(langName(o.lang()));
      ignoreOnce = new Set();
      scan();
      at = 0;
      show();
    }
    function close() { pane.hidden = true; document.body.classList.remove('ai-open'); }

    function show() {
      body.innerHTML = '';
      const sum = document.createElement('div');
      sum.className = 'proof-sum';
      const ns = issues.filter(x => x.kind === 'spelling').length, ng = issues.length - ns;
      sum.textContent = issues.length ? `${ns} spelling · ${ng} grammar${canSpell() ? '' : ' (spelling is underlined by your browser)'}` : '';
      body.appendChild(sum);
      if (!issues.length) {
        const d = document.createElement('div');
        d.className = 'proof-done';
        d.innerHTML = '<div class="big">✓</div><b>No suggestions</b><p></p>';
        d.querySelector('p').textContent = canSpell() ? 'Spelling and grammar look good.' : 'Grammar looks good. In the browser, misspelled words are underlined as you type; the desktop app checks spelling here too.';
        body.appendChild(d);
        return;
      }
      at = Math.max(0, Math.min(at, issues.length - 1));
      const x = issues[at], seg = segs[x.seg];
      if (x.kind === 'spelling' && !x.suggestions) x.suggestions = window.atlas.spellSuggest(x.word);
      const card = document.createElement('div');
      card.className = 'proof-card ' + x.kind;
      card.innerHTML = `<div class="kind"></div><div class="ctx"></div><div class="msg"></div><div class="sugg"></div><div class="acts"></div><div class="nav"></div>`;
      card.querySelector('.kind').textContent = x.kind === 'spelling' ? 'Spelling' : x.rule === 'long' ? 'Clarity' : 'Grammar';
      const ctx = card.querySelector('.ctx');
      const s0 = Math.max(0, x.start - 50), e0 = Math.min(seg.text.length, x.end + 50);
      ctx.append((s0 > 0 ? '…' : '') + seg.text.slice(s0, x.start));
      const mark = document.createElement('mark');
      mark.textContent = seg.text.slice(x.start, x.end) || '␣';
      ctx.append(mark, seg.text.slice(x.end, e0) + (e0 < seg.text.length ? '…' : ''));
      card.querySelector('.msg').textContent = x.message;
      const sugg = card.querySelector('.sugg');
      (x.suggestions || []).slice(0, 6).forEach(s => {
        const b = document.createElement('button');
        b.className = 'primary';
        b.textContent = s === '' ? 'Remove' : s === ' ' ? 'One space' : s;
        b.onclick = () => change(x, s, false);
        sugg.appendChild(b);
      });
      if (x.kind === 'spelling' && !(x.suggestions || []).length) sugg.textContent = 'No suggestions';
      const acts = card.querySelector('.acts');
      const btn = (label, fn) => { const b = document.createElement('button'); b.textContent = label; b.onclick = fn; acts.appendChild(b); };
      if (x.kind === 'spelling' && (x.suggestions || []).length) btn('Change all', () => change(x, x.suggestions[0], true));
      btn('Ignore once', () => { ignoreOnce.add(onceKey(seg, x)); rescan(x); });
      btn('Ignore all', () => { ignoreAll.add(x.kind === 'spelling' ? x.word.toLowerCase() : 'rule:' + x.rule + ':' + x.word.toLowerCase()); rescan(x); });
      if (x.kind === 'spelling' && window.atlas && window.atlas.spellAddWord) btn('Add to dictionary', async () => { await window.atlas.spellAddWord(x.word); ignoreAll.add(x.word.toLowerCase()); rescan(x); });
      const nav = card.querySelector('.nav');
      const prev = document.createElement('button'), next = document.createElement('button'), pos = document.createElement('span');
      prev.textContent = '‹'; next.textContent = '›'; pos.textContent = `${at + 1} of ${issues.length}`;
      prev.disabled = at === 0; next.disabled = at === issues.length - 1;
      prev.onclick = () => { at--; show(); }; next.onclick = () => { at++; show(); };
      nav.append(prev, pos, next);
      body.appendChild(card);
      reveal(x);
    }
    function reveal(x) {
      const seg = segs[x.seg];
      if (seg.root.focus) seg.root.focus();
      if (seg.root.live) {
        try {
          const r = rangeFor(seg, x.start, x.end);
          const sel = getSelection();
          sel.removeAllRanges(); sel.addRange(r);
          const el = r.startContainer.parentElement;
          if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' });
        } catch { /* the text changed */ }
      }
    }
    function replaceAt(seg, x, s) {
      const r = rangeFor(seg, x.start, x.end);
      if (seg.root.live) {
        const host = seg.root.el;
        host.focus();
        const sel = getSelection();
        sel.removeAllRanges(); sel.addRange(r);
        // insertText keeps Undo working; an empty replacement deletes.
        if (s === '') document.execCommand('delete'); else document.execCommand('insertText', false, s);
      } else {
        r.deleteContents();
        if (s) r.insertNode(document.createTextNode(s));
        seg.root.el.normalize();
      }
    }
    function change(x, s, all) {
      const touched = new Set();
      if (all) {
        // Same word everywhere, last first so earlier offsets stay right.
        const same = issues.filter(y => y.kind === 'spelling' && y.word === x.word).reverse();
        for (const y of same) { replaceAt(segs[y.seg], y, s); touched.add(segs[y.seg].root); }
      } else { replaceAt(segs[x.seg], x, s); touched.add(segs[x.seg].root); }
      touched.forEach(r => r.commit && r.commit());
      rescan(x);
    }
    function rescan(x) {
      // Carry on from the same place.
      const segIdx = x.seg, pos = x.start;
      scan();
      at = Math.max(0, issues.findIndex(y => y.seg > segIdx || (y.seg === segIdx && y.start >= pos)));
      show();
    }

    async function languageDialog() {
      let available = [];
      if (window.atlas && window.atlas.spellLanguages) available = (await window.atlas.spellLanguages()).available || [];
      const codes = [...new Set([...LANGS, ...available])].filter(c => !available.length || available.includes(c) || available.some(a => a.split('-')[0] === c.split('-')[0]));
      codes.sort((a, b) => langName(a).localeCompare(langName(b)));
      const v = await UI.form('Language', [{ name: 'lang', label: 'Check spelling and grammar in', type: 'select', value: o.lang() || '', options: [{ value: '', label: 'Default (' + langName(navigator.language) + ')' }, ...codes.map(c => ({ value: c, label: langName(c) }))] }], 'OK');
      if (!v) return;
      await setLanguage(v.lang);
      if (!pane.hidden) open();
    }
    async function setLanguage(code) {
      o.setLang(code);
      if (window.atlas && window.atlas.spellSetLanguages) {
        const { available = [] } = await window.atlas.spellLanguages();
        const pick = code && (available.find(a => a === code) || available.find(a => a.split('-')[0] === code.split('-')[0]));
        await window.atlas.spellSetLanguages(pick ? [pick] : []);
      }
      langBtn.innerHTML = '<span class="ico" data-icon="languages"></span>'; langBtn.append(langName(code));
      UI.toast('Proofing language: ' + langName(code || navigator.language));
    }

    document.addEventListener('keydown', e => { if (e.key === 'F7' && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); if (pane.hidden) open(); else close(); } });
    return { open, close, languageDialog, setLanguage, scan: () => { scan(); return issues; } };
  }

  window.Proofing = { attach, grammar, segments, langName, LANGS };
})();
