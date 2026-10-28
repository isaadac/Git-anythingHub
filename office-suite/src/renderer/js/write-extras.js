'use strict';
// Isadek Write — section breaks (orientation and columns per section), pictures that text wraps
// around, equations, fields (date, time, author, file name, page, pages, words), editable
// styles, and comparing two documents. Each goes to and comes back from .docx as Word's own
// construct: sections, anchored drawings, Office Math, Word fields and the styles part.
//
// In the document:
//   <hr class="section-break" data-orient="landscape" data-cols="2">   starts a new section
//   <img class="float-left|float-right|wrap-block">                      wrapping
//   <span class="equation [eq-block]" data-eq="a^2+b^2=c^2">             (MathML drawn from data-eq)
//   <span class="fld" data-fld="date|time|author|filename|page|numpages|numwords">
(function () {
  const MK = { eqS: '', eqE: '', fS: '', fE: '', secS: '', secE: '' };
  const MK_RE = /(\d+)|(\d+|[a-z]+)|([a-z]+)\|(\d)/g;
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  const FIELDS = {
    date: { label: 'Date', instr: ' DATE \\@ "d MMMM yyyy" ' },
    time: { label: 'Time', instr: ' TIME \\@ "HH:mm" ' },
    author: { label: 'Author', instr: ' AUTHOR ' },
    filename: { label: 'File name', instr: ' FILENAME ' },
    page: { label: 'Page number', instr: ' PAGE ' },
    numpages: { label: 'Number of pages', instr: ' NUMPAGES ' },
    numwords: { label: 'Number of words', instr: ' NUMWORDS ' }
  };
  const FIELD_OF_INSTR = instr => {
    const m = String(instr).trim().toUpperCase().match(/^(DATE|TIME|AUTHOR|FILENAME|PAGE|NUMPAGES|NUMWORDS|CREATEDATE|SAVEDATE|PRINTDATE)\b/);
    if (!m) return null;
    return { DATE: 'date', CREATEDATE: 'date', SAVEDATE: 'date', PRINTDATE: 'date', TIME: 'time' }[m[1]] || m[1].toLowerCase();
  };

  const STYLE_KEYS = { normal: 'Normal', title: 'Title', h1: 'Heading 1', h2: 'Heading 2', h3: 'Heading 3', quote: 'Quote' };
  const STYLE_SEL = { normal: 'p, li', title: 'h1.title', h1: 'h1:not(.title)', h2: 'h2', h3: 'h3', quote: 'blockquote' };
  const DOCX_STYLE = { normal: 'document', title: 'title', h1: 'heading1', h2: 'heading2', h3: 'heading3' };
  const WORD_STYLE_ID = { Normal: 'normal', Title: 'title', Heading1: 'h1', Heading2: 'h2', Heading3: 'h3', Quote: 'quote' };
  const FONT_RE = /^[\w .'-]{1,60}$/;

  /** Styles from a file or a peer: only fonts, sizes, colours, bold/italic and spacing. */
  function cleanStyles(v) {
    const out = {};
    if (!v || typeof v !== 'object') return out;
    for (const k of Object.keys(STYLE_KEYS)) {
      const s = v[k];
      if (!s || typeof s !== 'object') continue;
      const o = {};
      if (typeof s.font === 'string' && FONT_RE.test(s.font)) o.font = s.font;
      if (+s.size >= 4 && +s.size <= 200) o.size = Math.round(+s.size * 2) / 2;
      if (typeof s.color === 'string' && /^#[0-9a-f]{6}$/i.test(s.color)) o.color = s.color.toLowerCase();
      if (typeof s.bold === 'boolean') o.bold = s.bold;
      if (typeof s.italic === 'boolean') o.italic = s.italic;
      if (+s.before >= 0 && +s.before <= 200 && s.before !== '' && s.before !== null) o.before = +s.before;
      if (+s.after >= 0 && +s.after <= 200 && s.after !== '' && s.after !== null) o.after = +s.after;
      if (Object.keys(o).length) out[k] = o;
    }
    return out;
  }
  function styleCss(styles, scope) {
    return Object.entries(cleanStyles(styles)).map(([k, s]) => {
      const d = [];
      if (s.font) d.push(`font-family: "${s.font}"`);
      if (s.size) d.push(`font-size: ${s.size}pt`);
      if (s.color) d.push(`color: ${s.color}`);
      if (s.bold !== undefined) d.push(`font-weight: ${s.bold ? 'bold' : 'normal'}`);
      if (s.italic !== undefined) d.push(`font-style: ${s.italic ? 'italic' : 'normal'}`);
      if (s.before !== undefined) d.push(`margin-top: ${s.before}pt`);
      if (s.after !== undefined) d.push(`margin-bottom: ${s.after}pt`);
      return STYLE_SEL[k].split(', ').map(sel => (scope ? `${scope} ${sel}` : sel)).join(', ') + ` { ${d.join('; ')} }`;
    }).join('\n');
  }

  // ---------------- Compare ----------------
  const tokens = s => s.match(/\s+|[^\s]+/g) || [];
  /** Longest-common-subsequence alignment of two lists: [['=', i, j] | ['-', i] | ['+', j]]. */
  function align(a, b, same = (x, y) => x === y) {
    const n = a.length, m = b.length;
    // Long lists: trim the common start and end first; very long middles are compared coarsely.
    let s = 0;
    while (s < n && s < m && same(a[s], b[s])) s++;
    let e = 0;
    while (e < n - s && e < m - s && same(a[n - 1 - e], b[m - 1 - e])) e++;
    const A = a.slice(s, n - e), B = b.slice(s, m - e);
    const out = [];
    for (let i = 0; i < s; i++) out.push(['=', i, i]);
    if (A.length * B.length > 4e6) {
      A.forEach((_, i) => out.push(['-', s + i]));
      B.forEach((_, j) => out.push(['+', s + j]));
    } else {
      const L = Array.from({ length: A.length + 1 }, () => new Uint32Array(B.length + 1));
      for (let i = A.length - 1; i >= 0; i--) for (let j = B.length - 1; j >= 0; j--) L[i][j] = same(A[i], B[j]) ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
      let i = 0, j = 0;
      while (i < A.length && j < B.length) {
        if (same(A[i], B[j])) { out.push(['=', s + i, s + j]); i++; j++; }
        else if (L[i + 1][j] >= L[i][j + 1]) { out.push(['-', s + i]); i++; }
        else { out.push(['+', s + j]); j++; }
      }
      for (; i < A.length; i++) out.push(['-', s + i]);
      for (; j < B.length; j++) out.push(['+', s + j]);
    }
    for (let k = 0; k < e; k++) out.push(['=', n - e + k, m - e + k]);
    return out;
  }
  const norm = el => el.textContent.replace(/\s+/g, ' ').trim();
  function similarity(x, y) {
    const a = new Set(x.toLowerCase().split(/\W+/).filter(Boolean)), b = y.toLowerCase().split(/\W+/).filter(Boolean);
    if (!a.size || !b.length) return 0;
    const common = b.filter(w => a.has(w)).length;
    return (2 * common) / (a.size + b.length);
  }

  /**
   * Marks how `revised` (an element whose children are the blocks of the newer document)
   * differs from `original` as tracked changes by `author`. Formatting in the revised text is kept.
   */
  function compareInto(original, revised, author) {
    const date = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const mark = tag => { const el = document.createElement(tag); el.className = 'tc'; el.dataset.author = author; el.setAttribute('datetime', date); return el; };
    const wrapText = (root, tag) => {
      const tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      const list = [];
      while (tw.nextNode()) if (tw.currentNode.nodeValue.trim()) list.push(tw.currentNode);
      list.forEach(t => { const m = mark(tag); t.replaceWith(m); m.appendChild(t); });
      if (!list.length && root.querySelector('img')) root.querySelectorAll('img').forEach(img => { const m = mark(tag); img.replaceWith(m); m.appendChild(img); });
      return root;
    };
    /** Word-level changes inside one block, keeping its formatting. */
    function diffBlock(oldEl, newEl) {
      const a = tokens(oldEl.textContent), b = tokens(newEl.textContent);
      const ops = align(a, b);
      // Character offsets in the revised text: insertions to wrap, deletions to put back.
      const wraps = [], dels = [];
      let pos = 0, pendingDel = '', delAt = 0;
      const flushDel = () => { if (pendingDel) dels.push([delAt, pendingDel]); pendingDel = ''; };
      for (const op of ops) {
        if (op[0] === '=') { flushDel(); pos += b[op[2]].length; }
        else if (op[0] === '-') { if (!pendingDel) delAt = pos; pendingDel += a[op[1]]; }
        else {
          flushDel();
          const len = b[op[1]].length;
          const last = wraps[wraps.length - 1];
          if (last && last[1] === pos) last[1] += len; else wraps.push([pos, pos + len]);
          pos += len;
        }
      }
      flushDel();
      // Whitespace-only changes are not worth marking.
      const realWraps = wraps.filter(([s, e]) => newEl.textContent.slice(s, e).trim());
      const realDels = dels.filter(([, t]) => t.trim());
      const textNodes = () => {
        const out = [];
        const tw = document.createTreeWalker(newEl, NodeFilter.SHOW_TEXT, { acceptNode: n => (n.parentElement.closest('del.tc') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
        let at = 0;
        while (tw.nextNode()) { out.push([tw.currentNode, at]); at += tw.currentNode.nodeValue.length; }
        return out;
      };
      const ops2 = [...realWraps.map(w => ({ k: 'w', s: w[0], e: w[1] })), ...realDels.map(d => ({ k: 'd', s: d[0], t: d[1] }))]
        .sort((x, y) => y.s - x.s || (x.k === 'w' ? -1 : 1));
      for (const op of ops2) {
        if (op.k === 'w') {
          for (const [node, at] of textNodes().reverse()) {
            const end = at + node.nodeValue.length;
            if (end <= op.s || at >= op.e) continue;
            let t = node;
            if (op.e < end) t.splitText(op.e - at);
            if (op.s > at) t = t.splitText(op.s - at);
            const m = mark('ins'); t.replaceWith(m); m.appendChild(t);
          }
        } else {
          const list = textNodes();
          const hit = list.find(([n, at]) => op.s >= at && op.s <= at + n.nodeValue.length) || list[list.length - 1];
          const d = mark('del'); d.textContent = op.t;
          if (!hit) { newEl.appendChild(d); continue; }
          const [node, at] = hit;
          if (op.s - at >= node.nodeValue.length) node.after(d);
          else if (op.s - at <= 0) node.before(d);
          else { const rest = node.splitText(op.s - at); rest.before(d); }
        }
      }
    }
    const A = [...original.children], B = [...revised.children];
    const ops = align(A.map(norm), B.map(norm));
    // Pair up removed and added blocks that are mostly the same text: those are edits.
    const out = [];
    let dels = [], adds = [];
    const flush = () => {
      // In each changed stretch: removed blocks first (as Word shows them), then the new
      // blocks in their order, edited ones marked word by word.
      const pairOf = new Map(), paired = new Set();
      for (const d of dels) {
        const k = adds.findIndex((x, i) => !pairOf.has(i) && similarity(norm(A[d]), norm(B[x])) >= 0.5);
        if (k >= 0) { pairOf.set(k, d); paired.add(d); }
      }
      dels.forEach(d => { if (!paired.has(d)) out.push({ k: 'del', a: A[d] }); });
      adds.forEach((x, i) => out.push(pairOf.has(i) ? { k: 'mod', a: A[pairOf.get(i)], b: B[x] } : { k: 'add', b: B[x] }));
      dels = []; adds = [];
    };
    for (const op of ops) {
      if (op[0] === '=') { flush(); out.push({ k: 'same', b: B[op[2]] }); }
      else if (op[0] === '-') dels.push(op[1]);
      else adds.push(op[1]);
    }
    flush();
    // Rebuild the revised blocks in order, with deleted blocks put back where they were.
    const result = document.createDocumentFragment();
    let changes = 0;
    for (const r of out) {
      if (r.k === 'same') result.appendChild(r.b);
      else if (r.k === 'add') { result.appendChild(wrapText(r.b, 'ins')); changes++; }
      else if (r.k === 'del') { result.appendChild(wrapText(r.a.cloneNode(true), 'del')); changes++; }
      else { diffBlock(r.a, r.b); result.appendChild(r.b); if (r.b.querySelector('ins.tc, del.tc')) changes++; }
    }
    revised.innerHTML = '';
    revised.appendChild(result);
    return changes;
  }

  function WriteExtras(o) {
    const { doc } = o;

    // ---------------- Equations ----------------
    const drawn = new WeakMap();
    function renderEquations(root = doc) {
      root.querySelectorAll('span.equation').forEach(s => {
        const src = s.dataset.eq || '';
        const block = s.classList.contains('eq-block');
        if (drawn.get(s) === src + block && s.querySelector('math')) return;
        try { s.innerHTML = WriteMath.toMathML(src, block); } catch { s.textContent = src; }
        drawn.set(s, src + block);
        s.contentEditable = 'false';
        s.title = 'Equation — click to edit';
      });
    }
    async function editEquation(el) {
      if (!el) o.captureSelection();
      const r = await UI.form(el ? 'Equation' : 'Insert Equation', [
        { name: 'eq', label: 'Equation, e.g.  x = (-b \\pm sqrt(b^2-4ac))/(2a)   a_i^2   \\sum_(i=1)^n i   sqrt[3](x)   \\alpha  (empty to remove)', type: 'textarea', value: el ? el.dataset.eq : '' },
        { name: 'display', label: 'Show', type: 'select', options: [{ value: 'inline', label: 'In the line of text' }, { value: 'block', label: 'On its own line' }], value: el && el.classList.contains('eq-block') ? 'block' : 'inline' }
      ], el ? 'Save' : 'Insert');
      if (!r) return;
      const src = r.eq.trim().slice(0, 2000);
      if (el) {
        if (!src) el.remove();
        else { el.dataset.eq = src; el.classList.toggle('eq-block', r.display === 'block'); }
      } else if (src) {
        o.restoreSelection();
        Platform.insertInline(`<span class="equation${r.display === 'block' ? ' eq-block' : ''}" data-eq="${esc(src)}" contenteditable="false"></span>`);
      }
      renderEquations();
      o.changed();
    }

    // ---------------- Fields ----------------
    function fieldText(kind, el) {
      switch (kind) {
        case 'date': return new Date().toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' });
        case 'time': return new Date().toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
        case 'author': return o.author();
        case 'filename': return o.fileName();
        case 'page': return String(o.pageOf(el));
        case 'numpages': return String(o.pages());
        case 'numwords': return String(o.wordCount());
      }
      return '';
    }
    /**
     * Brings fields up to date. As in Word, page numbers, the date and the time follow by
     * themselves (opening, layout, saving, printing); the others (author, file name, a document's
     * creation date…) keep what they show until F9 (all = true).
     */
    function updateFields(all) {
      doc.querySelectorAll('span.fld').forEach(f => {
        if (!FIELDS[f.dataset.fld]) return;
        const instr = (f.dataset.instr || '').trim().toUpperCase();
        // A creation, save or print date isn't today's date: Isadek Write keeps Word's.
        if (/^(CREATEDATE|SAVEDATE|PRINTDATE)\b/.test(instr)) return;
        const auto = ['page', 'numpages', 'date', 'time'].includes(f.dataset.fld);
        if (!all && !auto && f.textContent.trim()) return;
        const t = fieldText(f.dataset.fld, f) || ' ';
        if (f.textContent !== t) f.textContent = t;
        f.contentEditable = 'false';
      });
    }
    async function insertField() {
      o.captureSelection();
      const r = await UI.form('Insert Field', [
        { name: 'kind', label: 'Field (it updates by itself; F9 updates all fields)', type: 'select', options: Object.entries(FIELDS).map(([value, f]) => ({ value, label: f.label })) }
      ], 'Insert');
      if (!r || !FIELDS[r.kind]) return;
      o.restoreSelection();
      Platform.insertInline(`<span class="fld" data-fld="${r.kind}" contenteditable="false">${esc(fieldText(r.kind, doc) || ' ')}</span>`);
      updateFields();
      o.changed();
    }

    // ---------------- Section breaks ----------------
    async function insertSectionBreak(existing) {
      if (!existing) o.captureSelection();
      const r = await UI.form(existing ? 'Section' : 'Insert Section Break', [
        { name: 'orient', label: 'The new section (from the next page) is', type: 'select', options: [{ value: 'portrait', label: 'Portrait' }, { value: 'landscape', label: 'Landscape' }], value: existing ? existing.dataset.orient : o.setup().orientation === 'landscape' ? 'portrait' : 'landscape' },
        { name: 'cols', label: 'Columns', type: 'select', options: ['1', '2', '3'], value: existing ? existing.dataset.cols || '1' : '1' }
      ], existing ? 'Save' : 'Insert');
      if (!r) return;
      const orient = r.orient === 'landscape' ? 'landscape' : 'portrait', cols = ['1', '2', '3'].includes(r.cols) ? r.cols : '1';
      if (existing) { existing.dataset.orient = orient; existing.dataset.cols = cols; }
      else { o.restoreSelection(); document.execCommand('insertHTML', false, `<hr class="section-break" data-orient="${orient}" data-cols="${cols}" contenteditable="false"><p><br></p>`); }
      o.changed();
    }
    /** The document's top-level blocks split at section breaks: [{ html, orient, cols }]. */
    function splitSections(body) {
      const first = o.setup();
      const out = [{ nodes: [], orient: first.orientation, cols: first.columns }];
      for (const el of [...body.childNodes]) {
        if (el.nodeType === 1 && el.matches('hr.section-break')) out.push({ nodes: [], orient: el.dataset.orient === 'landscape' ? 'landscape' : 'portrait', cols: Math.min(3, Math.max(1, +el.dataset.cols || 1)) });
        else out[out.length - 1].nodes.push(el);
      }
      return out.map(s => { const d = document.createElement('div'); s.nodes.forEach(n => d.appendChild(n)); return { html: d.innerHTML || '<p></p>', orient: s.orient, cols: s.cols }; });
    }

    // ---------------- Pictures: wrapping ----------------
    const WRAPS = ['float-left', 'float-right', 'wrap-block'];
    function setWrap(img, wrap) {
      img.classList.remove(...WRAPS);
      if (WRAPS.includes(wrap)) {
        img.classList.add(wrap);
        // Alignment set earlier (display:block with auto margins) would fight the float.
        img.style.removeProperty('display'); img.style.removeProperty('margin-left'); img.style.removeProperty('margin-right');
      }
    }
    const wrapOf = img => WRAPS.find(c => img.classList.contains(c)) || null;

    // ---------------- Styles ----------------
    let styleEl = null;
    function applyStyles() {
      if (!styleEl) { styleEl = document.createElement('style'); styleEl.id = 'docUserStyles'; document.head.appendChild(styleEl); }
      styleEl.textContent = styleCss(o.setup().styles, '#doc');
    }
    function styleKeyOf(block) {
      if (!block) return 'normal';
      if (block.tagName === 'H1') return block.classList.contains('title') ? 'title' : 'h1';
      return { H2: 'h2', H3: 'h3', BLOCKQUOTE: 'quote' }[block.tagName] || 'normal';
    }
    async function modifyStyle(block) {
      const key0 = styleKeyOf(block);
      const cur = { ...((o.setup().styles || {})[key0] || {}) };
      // Start from what the paragraph looks like now.
      const cs = block ? getComputedStyle(block) : null;
      const hex = c => { const m = String(c).match(/\d+/g); return m ? '#' + m.slice(0, 3).map(x => (+x).toString(16).padStart(2, '0')).join('') : '#000000'; };
      const fonts = [...document.querySelectorAll('#fontName option')].map(op => op.value);
      const font0 = cur.font || (cs ? cs.fontFamily.split(',')[0].replace(/["']/g, '').trim() : 'Calibri');
      const r = await UI.form('Modify Style', [
        { name: 'key', label: 'Style', type: 'select', options: Object.entries(STYLE_KEYS).map(([value, label]) => ({ value, label })), value: key0 },
        { name: 'font', label: 'Font', type: 'select', options: [...new Set([font0, ...fonts])], value: font0 },
        { name: 'size', label: 'Size (pt)', type: 'number', min: 4, max: 200, value: cur.size || (cs ? Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2 : 11) },
        { name: 'color', label: 'Colour', type: 'color', value: cur.color || (cs ? hex(cs.color) : '#000000') },
        { name: 'bold', label: 'Bold', type: 'select', options: [{ value: 'yes', label: 'Bold' }, { value: 'no', label: 'Not bold' }], value: (cur.bold ?? (cs && +cs.fontWeight >= 600)) ? 'yes' : 'no' },
        { name: 'italic', label: 'Italic', type: 'select', options: [{ value: 'yes', label: 'Italic' }, { value: 'no', label: 'Not italic' }], value: (cur.italic ?? (cs && cs.fontStyle === 'italic')) ? 'yes' : 'no' },
        { name: 'before', label: 'Space before (pt)', type: 'number', min: 0, max: 200, value: cur.before ?? (cs ? Math.round(parseFloat(cs.marginTop) * 0.75) : 0) },
        { name: 'after', label: 'Space after (pt)', type: 'number', min: 0, max: 200, value: cur.after ?? (cs ? Math.round(parseFloat(cs.marginBottom) * 0.75) : 8) }
      ], 'Apply to all');
      if (!r) return;
      const st = { ...(o.setup().styles || {}) };
      st[r.key] = { font: r.font, size: +r.size, color: r.color, bold: r.bold === 'yes', italic: r.italic === 'yes', before: +r.before, after: +r.after };
      o.setStyles(cleanStyles(st));
      applyStyles();
      o.changed();
      UI.toast(`${STYLE_KEYS[r.key]} updated everywhere it is used`);
    }
    /** Word style definitions (docx library form) and CSS for its HTML conversion. */
    function docxStyles() {
      const st = cleanStyles(o.setup().styles);
      const out = {};
      for (const [k, s] of Object.entries(st)) {
        if (!DOCX_STYLE[k]) continue;
        const run = {};
        if (s.font) run.font = s.font;
        if (s.size) run.size = Math.round(s.size * 2);
        if (s.color) run.color = s.color.slice(1);
        if (s.bold !== undefined) run.bold = s.bold;
        if (s.italic !== undefined) run.italics = s.italic;
        const spacing = {};
        if (s.before !== undefined) spacing.before = Math.round(s.before * 20);
        if (s.after !== undefined) spacing.after = Math.round(s.after * 20);
        out[DOCX_STYLE[k]] = { run, ...(Object.keys(spacing).length ? { paragraph: { spacing } } : {}) };
      }
      return out;
    }

    // ---------------- .docx: saving ----------------
    /** Turns equations and fields into markers before the HTML-to-Word conversion. */
    function prepareForDocx(body) {
      const st = { eqs: [], fields: [], wraps: [] };
      body.querySelectorAll('span.equation').forEach(s => {
        st.eqs.push({ src: s.dataset.eq || '', block: s.classList.contains('eq-block') });
        s.replaceWith(document.createTextNode(MK.eqS + (st.eqs.length - 1) + MK.eqE));
      });
      body.querySelectorAll('span.fld').forEach(f => {
        if (!FIELDS[f.dataset.fld]) return;
        st.fields.push({ kind: f.dataset.fld, instr: f.dataset.instr || FIELDS[f.dataset.fld].instr, text: f.textContent });
        f.replaceWith(document.createTextNode(MK.fS + (st.fields.length - 1) + MK.fE));
      });
      st.wraps = [...body.querySelectorAll('img')].map(wrapOf);
      return st;
    }
    /** Replaces the markers with Office Math and Word fields, and anchors wrapped pictures. */
    async function patchDocx(bytes, st) {
      if (!st.eqs.length && !st.fields.length && !st.wraps.some(Boolean)) return bytes;
      const zip = await JSZip.loadAsync(bytes);
      const f = zip.file('word/document.xml');
      let xml = await f.async('string');
      if (st.eqs.length || st.fields.length) {
        xml = xml.replace(/<w:r>((?:(?!<\/w:r>)[\s\S])*?)<\/w:r>/g, (all, inner) => {
          if (!/[]/.test(inner)) return all;
          const rPr = (inner.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
          const text = [...inner.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
          const runOf = t => (t ? `<w:r>${rPr}<w:t xml:space="preserve">${t}</w:t></w:r>` : '');
          let out = '', last = 0;
          text.replace(/(\d+)|(\d+)/g, (m, eq, fl, at) => {
            out += runOf(text.slice(last, at));
            last = at + m.length;
            if (eq !== undefined) {
              const e = st.eqs[+eq];
              const math = e ? WriteMath.toOMML(e.src) : '';
              out += e && e.block ? `<m:oMathPara>${math}</m:oMathPara>` : math;
            } else {
              const fd = st.fields[+fl];
              if (fd) out += `<w:fldSimple w:instr="${xmlEsc(' ' + fd.instr.trim() + ' ')}"><w:r>${rPr}<w:t xml:space="preserve">${xmlEsc(fd.text)}</w:t></w:r></w:fldSimple>`;
            }
            return m;
          });
          return out + runOf(text.slice(last));
        });
        if (!/xmlns:m=/.test(xml.slice(0, 3000))) xml = xml.replace(/<w:document\b/, `<w:document xmlns:m="${M_NS}"`);
      }
      // Wrapped pictures: Word "anchors" them instead of placing them in the line.
      const inlines = xml.match(/<wp:inline\b[\s\S]*?<\/wp:inline>/g) || [];
      if (st.wraps.some(Boolean) && inlines.length === st.wraps.length) {
        let i = 0, z = 251659264;
        xml = xml.replace(/<wp:inline\b([^>]*)>([\s\S]*?)<\/wp:inline>/g, (all, attrs, inner) => {
          const w = st.wraps[i++];
          if (!w) return all;
          const at = inner.indexOf('<wp:docPr');
          const size = inner.slice(0, at), rest = inner.slice(at);
          const ext = size.match(/<wp:extent\b[^>]*\/>/);
          const eff = size.match(/<wp:effectExtent\b[^>]*\/>/);
          const pos = w === 'wrap-block' ? '<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>' : `<wp:positionH relativeFrom="column"><wp:align>${w === 'float-right' ? 'right' : 'left'}</wp:align></wp:positionH>`;
          const wrap = w === 'wrap-block' ? '<wp:wrapTopAndBottom/>' : '<wp:wrapSquare wrapText="bothSides"/>';
          return `<wp:anchor distT="57150" distB="57150" distL="114300" distR="114300" simplePos="0" relativeHeight="${z++}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
            `<wp:simplePos x="0" y="0"/>${pos}<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
            `${ext ? ext[0] : ''}${eff ? eff[0] : '<wp:effectExtent l="0" t="0" r="0" b="0"/>'}${wrap}${rest}</wp:anchor>`;
        });
      }
      zip.file('word/document.xml', xml);
      return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
    }

    // ---------------- .docx: opening ----------------
    /** Keeps what the HTML converter would drop: equations, fields, section breaks, wrapping, styles. */
    async function preprocessDocx(bytes) {
      const zip = await JSZip.loadAsync(bytes);
      const f = zip.file('word/document.xml');
      const meta = { eqs: [], fields: [], wraps: null, first: null, styles: {} };
      if (!f) return { bytes, meta };
      let xml = await f.async('string');
      const mark = s => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
      // Equations → the linear form.
      const linear = frag => {
        try {
          const d = new DOMParser().parseFromString(`<r xmlns:m="${M_NS}" xmlns:w="${W_NS}">${frag}</r>`, 'application/xml');
          if (d.getElementsByTagName('parsererror').length) return '';
          return [...d.getElementsByTagNameNS(M_NS, 'oMath')].map(WriteMath.fromOMML).join('  ');
        } catch { return ''; }
      };
      const takeMath = block => all => {
        const src = linear(all);
        if (!src) return all;
        meta.eqs.push({ src, block });
        return mark(MK.eqS + (meta.eqs.length - 1) + MK.eqE);
      };
      xml = xml.replace(/<m:oMathPara\b[\s\S]*?<\/m:oMathPara>/g, takeMath(true));
      xml = xml.replace(/<m:oMath\b[\s\S]*?<\/m:oMath>/g, takeMath(false));
      // Fields Isadek Write keeps live; the others (TOC, links…) are left alone.
      // A field keeps what Word last showed in it (its result) and its instruction, so it reads
      // the same as in Word and is saved back as it was.
      const unXml = t => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
      const resultOf = x => [...x.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(m => unXml(m[1])).join('');
      const field = (instr, result) => {
        const k = FIELD_OF_INSTR(instr);
        if (!k) return null;
        meta.fields.push({ kind: k, instr: instr.trim(), text: result });
        return mark(MK.fS + (meta.fields.length - 1) + MK.fE);
      };
      xml = xml.replace(/<w:fldSimple\b[^>]*w:instr="([^"]*)"[^>]*(?:\/>|>([\s\S]*?)<\/w:fldSimple>)/g, (all, instr, inner) => field(unXml(instr), resultOf(inner || '')) || all);
      xml = xml.replace(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:fldChar w:fldCharType="begin"[^>]*\/>[\s\S]*?<w:fldChar w:fldCharType="end"[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g, all => {
        if ((all.match(/fldCharType="begin"/g) || []).length > 1) return all;
        const instr = unXml([...all.matchAll(/<w:instrText\b[^>]*>([^<]*)<\/w:instrText>/g)].map(m => m[1]).join(''));
        const sep = all.search(/<w:fldChar w:fldCharType="separate"/);
        return field(instr, sep >= 0 ? resultOf(all.slice(sep)) : '') || all;
      });
      // Section breaks: a paragraph that ends a section carries that section's properties; the
      // break we put after it gets the properties of the section that follows.
      const sects = [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)];
      const props = s => {
        const sz = s.match(/<w:pgSz\b[^>]*>/);
        const w = sz && +((sz[0].match(/w:w="(\d+)"/) || [])[1]), h = sz && +((sz[0].match(/w:h="(\d+)"/) || [])[1]);
        const orient = sz && (/w:orient="landscape"/.test(sz[0]) || (w && h && w > h)) ? 'landscape' : 'portrait';
        const cols = Math.min(3, Math.max(1, +((s.match(/<w:cols\b[^>]*w:num="(\d+)"/) || [])[1]) || 1));
        return { orient, cols };
      };
      if (sects.length > 1) {
        meta.first = props(sects[0][0]);
        for (let i = sects.length - 2; i >= 0; i--) {
          const m = sects[i];
          const before = xml.slice(0, m.index);
          if (before.lastIndexOf('<w:pPr') <= before.lastIndexOf('</w:pPr>')) continue; // not in a paragraph
          const end = xml.indexOf('</w:p>', m.index + m[0].length);
          if (end < 0) continue;
          const nx = props(sects[i + 1][0]);
          xml = xml.slice(0, end) + mark(`${MK.secS}${nx.orient}|${nx.cols}${MK.secE}`) + xml.slice(end);
        }
      }
      // How each picture sits in the text, in document order.
      meta.wraps = (xml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/g) || []).map(d => {
        if (!/<wp:anchor\b/.test(d) || /<wp:wrapNone\/>/.test(d)) return null;
        if (/<wp:wrapTopAndBottom\b/.test(d)) return 'wrap-block';
        const h = (d.match(/<wp:positionH\b[\s\S]*?<\/wp:positionH>/) || [''])[0];
        const al = (h.match(/<wp:align>(\w+)<\/wp:align>/) || [])[1];
        const off = +((h.match(/<wp:posOffset>(-?\d+)<\/wp:posOffset>/) || [])[1] || 0);
        return al === 'right' || al === 'outside' || (!al && off > 2500000) ? 'float-right' : 'float-left';
      });
      // Styles (Normal, Title, Headings 1–3, Quote).
      try { meta.styles = await readStyles(zip); } catch { meta.styles = {}; }
      zip.file('word/document.xml', xml);
      return { bytes: new Uint8Array(await zip.generateAsync({ type: 'uint8array' })), meta };
    }
    async function readStyles(zip) {
      const sf = zip.file('word/styles.xml');
      if (!sf) return {};
      const d = new DOMParser().parseFromString(await sf.async('string'), 'application/xml');
      const tf = zip.file('word/theme/theme1.xml');
      const theme = tf ? new DOMParser().parseFromString(await tf.async('string'), 'application/xml') : null;
      const themeFont = kind => {
        if (!theme) return '';
        const f = theme.getElementsByTagNameNS('*', kind === 'major' ? 'majorFont' : 'minorFont')[0];
        const l = f && f.getElementsByTagNameNS('*', 'latin')[0];
        return l ? l.getAttribute('typeface') : '';
      };
      const a = (el, n) => el && (el.getAttributeNS(W_NS, n) || el.getAttribute('w:' + n));
      const one = (el, n) => el && [...el.children].find(c => c.localName === n);
      const out = {};
      for (const s of d.getElementsByTagNameNS(W_NS, 'style')) {
        const key = WORD_STYLE_ID[a(s, 'styleId')];
        if (!key || a(s, 'type') !== 'paragraph') continue;
        const r = one(s, 'rPr'), p = one(s, 'pPr');
        const st = {};
        const fonts = one(r, 'rFonts');
        if (fonts) {
          const t = a(fonts, 'asciiTheme');
          st.font = a(fonts, 'ascii') || (t ? themeFont(/major/.test(t) ? 'major' : 'minor') : '');
          if (!st.font) delete st.font;
        }
        const sz = one(r, 'sz');
        if (sz && +a(sz, 'val')) st.size = +a(sz, 'val') / 2;
        const color = one(r, 'color');
        if (color && /^[0-9a-f]{6}$/i.test(a(color, 'val'))) st.color = '#' + a(color, 'val');
        const b = one(r, 'b'), i = one(r, 'i');
        if (b) st.bold = !['0', 'false'].includes(a(b, 'val'));
        if (i) st.italic = !['0', 'false'].includes(a(i, 'val'));
        const sp = one(p, 'spacing');
        if (sp && a(sp, 'before') !== null && a(sp, 'before') !== '') st.before = +a(sp, 'before') / 20;
        if (sp && a(sp, 'after') !== null && a(sp, 'after') !== '') st.after = +a(sp, 'after') / 20;
        if (Object.keys(st).length) out[key] = st;
      }
      return cleanStyles(out);
    }
    function postprocessHtml(html, meta) {
      const d = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
      const body = d.body;
      const found = [];
      const tw = d.createTreeWalker(body, NodeFilter.SHOW_TEXT);
      while (tw.nextNode()) if (/[]/.test(tw.currentNode.nodeValue)) found.push(tw.currentNode);
      for (const t of found) {
        const frag = d.createDocumentFragment();
        let last = 0, section = null;
        const v = t.nodeValue;
        v.replace(MK_RE, (m, eq, fld, orient, cols, at) => {
          if (at > last) frag.appendChild(d.createTextNode(v.slice(last, at)));
          last = at + m.length;
          if (eq !== undefined) {
            const e = meta.eqs[+eq];
            if (!e) return m;
            const s = d.createElement('span');
            s.className = 'equation' + (e.block ? ' eq-block' : '');
            s.dataset.eq = e.src;
            frag.appendChild(s);
          } else if (fld !== undefined) {
            const s = d.createElement('span');
            s.className = 'fld';
            const fd = /^\d+$/.test(fld) ? (meta.fields || [])[+fld] : { kind: fld, text: '' };
            if (!fd) return m;
            s.dataset.fld = fd.kind;
            if (fd.instr && fd.instr !== FIELDS[fd.kind].instr.trim()) s.dataset.instr = fd.instr.slice(0, 500);
            s.textContent = fd.text || ' ';
            frag.appendChild(s);
          } else section = { orient, cols };
          return m;
        });
        if (last < v.length) frag.appendChild(d.createTextNode(v.slice(last)));
        const parent = t.parentNode;
        parent.replaceChild(frag, t);
        if (section) {
          let block = parent;
          while (block.parentNode && block.parentNode !== body) block = block.parentNode;
          const hr = d.createElement('hr');
          hr.className = 'section-break';
          hr.dataset.orient = section.orient;
          hr.dataset.cols = section.cols;
          if (block === body) body.appendChild(hr);
          else if (!block.textContent.trim() && !block.querySelector('img, span.equation, span.fld')) block.replaceWith(hr);
          else block.after(hr);
        }
      }
      if (meta.wraps) {
        const imgs = [...body.querySelectorAll('img')];
        if (imgs.length === meta.wraps.length) imgs.forEach((img, i) => { if (meta.wraps[i]) img.classList.add(meta.wraps[i]); });
      }
      return body.innerHTML;
    }

    /** Equations and fields as plain text, for Markdown and text files. */
    function prepareForText(root) {
      root.querySelectorAll('span.equation').forEach(s => s.replaceWith(document.createTextNode(`$${s.dataset.eq || ''}$`)));
      root.querySelectorAll('hr.section-break').forEach(h => { const p = document.createElement('hr'); h.replaceWith(p); });
    }

    // Clicking an equation or a section break edits it.
    doc.addEventListener('click', e => {
      const eq = e.target.closest && e.target.closest('span.equation');
      if (eq && doc.contains(eq)) { e.preventDefault(); editEquation(eq); return; }
      const sb = e.target.closest && e.target.closest('hr.section-break');
      if (sb && doc.contains(sb)) { e.preventDefault(); insertSectionBreak(sb); }
    });

    return {
      renderEquations, editEquation, updateFields, insertField, insertSectionBreak, splitSections,
      setWrap, wrapOf, applyStyles, modifyStyle, docxStyles, styleCss: scope => styleCss(o.setup().styles, scope),
      prepareForDocx, patchDocx, preprocessDocx, postprocessHtml, prepareForText,
      compare: (original, revised, author) => compareInto(original, revised, author)
    };
  }

  WriteExtras.cleanStyles = cleanStyles;
  WriteExtras.FIELDS = FIELDS;
  window.WriteExtras = WriteExtras;
})();
