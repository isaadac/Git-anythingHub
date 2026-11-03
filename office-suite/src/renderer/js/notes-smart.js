'use strict';
// Isadek Notes — Ink to Text, text from pictures, math help and the web clipper.
//
//   Ink to Text (Draw): the handwriting on the page is read on this device (ocr.js), shown for
//   correcting, and put in the page where it was written; the drawing can go or stay. With the
//   Assistant set up, it can read the writing instead (a picture of the strokes is sent).
//   Picture to Text (Insert): the words in a picture, to copy or put below it. Pictures added to a
//   page are read in the background too, so searching finds the words in them (img data-ocr).
//   Math (Draw): the selected (or handwritten, or typed) expression or equation is worked out with
//   notes-math.js: the answer, the steps, a graph to insert, or the Assistant to explain it.
//   Web Clip (Insert): a web page's article (or just a bookmark of it) as a new page, with its
//   pictures kept in the notebook and a link back to where it came from.
(function () {
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // ---------------- Clipping (no page needed: kept testable) ----------------
  const JUNK = /(^|[\s_-])(comment|share|social|promo|related|footer|nav|menu|sidebar|advert|ads?|cookie|newsletter|subscribe|banner|popup|modal|breadcrumb)s?([\s_-]|$)/i;
  /** The article in a web page: { title, html (untrusted: sanitize it), images: [absolute src], excerpt, site }. */
  function extractArticle(html, url) {
    const doc = new DOMParser().parseFromString(String(html).slice(0, 10e6), 'text/html');
    const base = (() => { try { return new URL(doc.querySelector('base[href]')?.getAttribute('href') || url, url).href; } catch { return url; } })();
    const meta = n => (doc.querySelector(`meta[property="${n}"], meta[name="${n}"]`) || {}).content || '';
    const title = (meta('og:title') || doc.title || (doc.querySelector('h1') || {}).textContent || url).trim().replace(/\s+/g, ' ').slice(0, 200);
    const site = meta('og:site_name') || (() => { try { return new URL(url).hostname.replace(/^www\./, ''); } catch { return ''; } })();
    const excerpt = (meta('og:description') || meta('description')).trim().slice(0, 500);
    doc.querySelectorAll('script, style, noscript, iframe, object, embed, form, button, input, select, textarea, svg, canvas, nav, footer, aside, template, link, meta').forEach(n => n.remove());
    doc.querySelectorAll('[hidden], [aria-hidden="true"]').forEach(n => n.remove());
    doc.querySelectorAll('[class], [id]').forEach(n => { if (n !== doc.body && !/^(ARTICLE|MAIN)$/.test(n.tagName) && JUNK.test(`${n.className} ${n.id}`)) n.remove(); });
    // The container with the most paragraph text (paragraphs of a sentence or more).
    const textLen = n => [...n.querySelectorAll('p, li, pre, blockquote')].reduce((t, p) => t + (p.textContent.trim().length >= 40 ? p.textContent.trim().length : 0), 0);
    let best = doc.querySelector('article') || doc.querySelector('main, [role="main"]');
    if (!best || textLen(best) < 200) {
      let score = -1;
      for (const n of doc.body ? doc.body.querySelectorAll('div, section, article, main, td') : []) {
        const direct = [...n.children].filter(c => /^(P|PRE|BLOCKQUOTE|UL|OL|H[2-6]|FIGURE)$/.test(c.tagName)).reduce((t, c) => t + c.textContent.trim().length, 0);
        if (direct > score) { score = direct; best = n; }
      }
      if (!best || score < 100) best = doc.body;
    }
    if (!best) return { title, html: '', images: [], excerpt, site };
    const abs = v => { try { const u = new URL(v, base); return /^https?:$/.test(u.protocol) ? u.href : null; } catch { return null; } };
    const images = [];
    best.querySelectorAll('img').forEach(img => {
      const src = abs(img.getAttribute('src') || img.getAttribute('data-src') || '');
      if (!src || images.length >= 20) { img.remove(); return; }
      img.setAttribute('src', src);
      ['srcset', 'sizes', 'loading', 'width', 'height', 'style'].forEach(a => img.removeAttribute(a));
      images.push(src);
    });
    best.querySelectorAll('a[href]').forEach(a => { const h = abs(a.getAttribute('href')); if (h) a.setAttribute('href', h); else a.removeAttribute('href'); });
    // The title is the page's title: a matching first heading isn't repeated.
    const h1 = best.querySelector('h1');
    if (h1 && h1.textContent.trim().replace(/\s+/g, ' ') === title) h1.remove();
    return { title, html: best.innerHTML, images, excerpt, site };
  }
  /** Text from a web page's bytes, in its character set. */
  function decode(bytes, contentType) {
    const head = new TextDecoder('latin1').decode(bytes.slice(0, 4096));
    const cs = (/charset=["']?([\w-]+)/i.exec(contentType || '') || /<meta[^>]+charset=["']?([\w-]+)/i.exec(head) || [])[1] || 'utf-8';
    try { return new TextDecoder(cs).decode(bytes); } catch { return new TextDecoder('utf-8').decode(bytes); }
  }

  function NotesSmart(ctx) {
    const { body, pageEl } = ctx;

    // ---------------- Ink to Text ----------------
    /** Where text read from strokes goes: after the last block that starts above them. */
    function insertAtInk(strokes, html) {
      const b = Ink.bounds(strokes), top = pageEl.getBoundingClientRect().top;
      const blocks = [...body.children].filter(n => n.getBoundingClientRect().height > 0);
      let after = null;
      for (const n of blocks) if (n.getBoundingClientRect().top - top < b.y) after = n;
      const holder = document.createElement('div');
      holder.innerHTML = ctx.sanitize(html);
      const nodes = [...holder.childNodes];
      if (!after) body.prepend(...nodes);
      else if (after === blocks[blocks.length - 1] && !after.textContent.trim() && !after.querySelector('img, table')) after.replaceWith(...nodes);
      else after.after(...nodes);
      ctx.markChanged();
    }
    const paragraphs = text => String(text).split(/\n+/).map(l => l.trim()).filter(Boolean).map(l => `<p>${esc(l)}</p>`).join('');
    async function inkToText() {
      const strokes = ctx.strokes().filter(s => s.o >= 0.9);
      if (!strokes.length) return UI.toast('Write something on the page with a pen first');
      UI.toast('Reading your handwriting…', 20000);
      let r;
      try { r = await Ocr.readInk(strokes); } catch (err) { return UI.toast('Couldn’t read the handwriting: ' + err.message, 5000); }
      UI.toast(r.text ? 'Check the text, then insert it' : 'No words were recognised — type them, or try the Assistant', 2500);
      const assistant = await Assistant.ready();
      const fields = [
        { name: 'text', label: 'Text (correct anything that was misread)', type: 'textarea', value: r.text },
        { name: 'ink', label: 'The drawing', type: 'select', options: [{ value: 'remove', label: 'Replace it with the text' }, { value: 'keep', label: 'Keep it too' }] }
      ];
      if (assistant) fields.push({ name: 'reader', label: 'Reader', type: 'select', options: [{ value: 'device', label: 'This text (read on this device)' }, { value: 'assistant', label: 'Ask the Assistant to read it (sends a picture of the writing)' }] });
      const v = await UI.form('Ink to Text', fields, 'Insert');
      if (!v) return;
      let text = v.text;
      if (v.reader === 'assistant') {
        UI.toast('The Assistant is reading the writing…', 30000);
        try {
          text = await Assistant.readImage(Ocr.inkCanvas(strokes, 60 * Math.max(1, Ocr.inkLines(strokes).length)).toDataURL('image/png'), 'Transcribe the handwriting in this picture exactly, keeping its line breaks. Reply with only the text. If a word can’t be read, give your best guess.');
        } catch (err) { return UI.toast('The Assistant couldn’t read it: ' + err.message, 5000); }
        const again = await UI.form('Ink to Text', [{ name: 'text', label: 'Text from the Assistant (correct anything)', type: 'textarea', value: text }], 'Insert');
        if (!again) return;
        text = again.text;
      }
      if (!text.trim()) return;
      ctx.snapshot && ctx.snapshot();
      insertAtInk(strokes, paragraphs(text));
      if (v.ink === 'remove') ctx.removeStrokes(strokes);
      UI.toast('Inserted');
    }

    // ---------------- Text from pictures ----------------
    let lastImg = null;
    body.addEventListener('click', e => { if (e.target.tagName === 'IMG' && body.contains(e.target)) lastImg = e.target; });
    function targetPicture() {
      const s = getSelection();
      if (s.rangeCount && body.contains(s.anchorNode)) {
        const r = s.getRangeAt(0), frag = r.cloneContents();
        if (frag.querySelector && frag.querySelector('img')) { const all = [...body.querySelectorAll('img')]; const hit = all.find(i => r.intersectsNode(i)); if (hit) return hit; }
      }
      if (lastImg && body.contains(lastImg)) return lastImg;
      const all = body.querySelectorAll('img');
      return all.length === 1 ? all[0] : null;
    }
    async function readPicture(img) {
      if (!img.complete) await new Promise(r => { img.onload = img.onerror = r; });
      const c = document.createElement('canvas');
      const k = Math.min(1, 4000 / Math.max(img.naturalWidth || 1, img.naturalHeight || 1));
      c.width = Math.max(1, Math.round(img.naturalWidth * k)); c.height = Math.max(1, Math.round(img.naturalHeight * k));
      const g = c.getContext('2d');
      g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
      g.drawImage(img, 0, 0, c.width, c.height);
      return Ocr.recognize(c);
    }
    async function pictureToText() {
      const img = targetPicture();
      if (!img) return UI.toast(body.querySelector('img') ? 'Click the picture to read first' : 'Insert or paste a picture first');
      UI.toast('Reading the picture…', 20000);
      let r;
      try { r = await readPicture(img); } catch (err) { return UI.toast('Couldn’t read the picture: ' + err.message, 5000); }
      if (r.text) img.setAttribute('data-ocr', r.text.slice(0, 20000));
      UI.toast(r.text ? 'Text found' : 'No text was found in the picture', 2000);
      if (!r.text) return;
      const v = await UI.form('Picture to Text', [{ name: 'text', label: 'Text in the picture (copy it, or insert it below the picture)', type: 'textarea', value: r.text }], 'Insert Below Picture');
      if (!v || !v.text.trim()) { ctx.markChanged(); return; }
      const holder = document.createElement('div');
      holder.innerHTML = ctx.sanitize(paragraphs(v.text));
      const block = img.closest('p, div, li') && img.closest('p, div, li') !== body ? img.closest('p, div, li') : img;
      block.after(...holder.childNodes);
      ctx.markChanged();
    }
    /** Pictures without their words read yet, read quietly one at a time so search finds them. */
    let indexing = false;
    const tried = new WeakSet(); // pictures already read (kept off the page: an attribute would change it)
    async function indexPictures() {
      if (indexing) return;
      indexing = true;
      try {
        for (let i = 0; i < 20; i++) {
          const img = [...body.querySelectorAll('img:not([data-ocr])')].find(x => x.naturalWidth >= 60 && x.naturalHeight >= 20 && !tried.has(x));
          if (!img) break;
          tried.add(img);
          const r = await readPicture(img).catch(() => null);
          if (r && body.contains(img)) { img.setAttribute('data-ocr', r.confidence >= 50 ? r.text.slice(0, 20000) : ''); ctx.markChanged(true); }
        }
      } finally { indexing = false; }
    }
    let indexTimer = null;
    const scheduleIndex = () => { clearTimeout(indexTimer); indexTimer = setTimeout(() => { (window.requestIdleCallback || setTimeout)(() => indexPictures()); }, 1500); };
    new MutationObserver(muts => { if (muts.some(m => [...m.addedNodes].some(n => n.nodeType === 1 && (n.tagName === 'IMG' || n.querySelector?.('img'))))) scheduleIndex(); }).observe(body, { childList: true, subtree: true });

    // ---------------- Math ----------------
    const pane = document.getElementById('mathPane');
    const mInput = document.getElementById('mathInput'), mOut = document.getElementById('mathOut');
    let current = null;
    function render() {
      mOut.innerHTML = '';
      current = null;
      const src = mInput.value.trim();
      if (!src) { mOut.innerHTML = '<p class="math-hint">Type or select an expression or an equation, for example <code>2x + 3 = 11</code>, <code>x² − 5x + 6 = 0</code>, <code>12 × (3 + 4)</code> or <code>y = sin(x)</code>.</p>'; return; }
      let r;
      try { r = NotesMath.analyze(src); } catch (err) { const p = document.createElement('p'); p.className = 'math-error'; p.textContent = err.message; mOut.appendChild(p); return; }
      current = r;
      const ans = document.createElement('div');
      ans.className = 'math-answer';
      ans.textContent = r.text;
      mOut.appendChild(ans);
      if (r.kind === 'graph') {
        const img = document.createElement('img');
        img.className = 'math-graph';
        img.alt = r.text;
        img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(NotesMath.graphSvg(r.fn, r.label, { width: 320, height: 220, dark: document.documentElement.dataset.theme === 'dark' }));
        mOut.appendChild(img);
        try { const z = r.roots(); if (z.solutions.length) { const p = document.createElement('p'); p.className = 'math-roots'; p.textContent = `Crosses zero at ${z.text.replace(/ or /g, ', ')}`; mOut.appendChild(p); } } catch { /* no roots to show */ }
      }
      if (r.steps && r.steps.length > 1) {
        const h = document.createElement('div'); h.className = 'math-sub'; h.textContent = 'Steps';
        const ol = document.createElement('ol'); ol.className = 'math-steps';
        r.steps.forEach(s => { const li = document.createElement('li'); li.textContent = s; ol.appendChild(li); });
        mOut.append(h, ol);
      }
      const acts = document.createElement('div');
      acts.className = 'math-actions';
      const btn = (label, run, primary) => { const b = document.createElement('button'); b.textContent = label; if (primary) b.className = 'primary'; b.addEventListener('click', run); acts.appendChild(b); };
      if (r.kind !== 'graph') btn('Insert Answer', () => insertMath(`<p>${esc(src)} → <b>${esc(r.text)}</b></p>`), true);
      if (r.steps && r.steps.length > 1) btn('Insert Steps', () => insertMath(`<ol class="math-steps">${r.steps.map(s => `<li>${esc(s)}</li>`).join('')}</ol>`));
      if (r.kind === 'graph') btn('Insert Graph', insertGraph, true);
      btn('Explain with Assistant', explain);
      mOut.appendChild(acts);
    }
    function insertMath(html) { ctx.restoreSel(); ctx.insertHtml(html); UI.toast('Inserted'); }
    async function insertGraph() {
      if (!current || current.kind !== 'graph') return;
      // As a PNG picture (the page keeps pictures as PNG, JPEG, GIF or WebP).
      const svg = NotesMath.graphSvg(current.fn, current.label);
      const img = new Image();
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
      await new Promise((res, rej) => { img.onload = res; img.onerror = rej; });
      const c = document.createElement('canvas');
      c.width = img.width; c.height = img.height;
      c.getContext('2d').drawImage(img, 0, 0);
      insertMath(`<p><img src="${c.toDataURL('image/png')}" alt="${esc(current.text)}" data-ocr=""></p>`);
    }
    async function explain() {
      const st = await Assistant.ready();
      if (!st) return UI.toast('Set up the Assistant (✨ in the title bar) to have it explain this');
      ctx.ask(`Explain step by step, for a student, how to work this out: ${mInput.value.trim()}${current && current.text ? ` (the answer is ${current.text})` : ''}. Use short numbered steps.`);
    }
    mInput.addEventListener('input', render);
    mInput.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Escape') closeMath(); });
    document.getElementById('mathClose').addEventListener('click', () => closeMath());
    function closeMath() { pane.hidden = true; document.getElementById('mathBtn').classList.remove('on'); }
    async function openMath() {
      const sel = ctx.selectionText().trim();
      let src = sel;
      if (!src && !pane.hidden) { closeMath(); return; }
      if (!src) {
        const strokes = ctx.strokes().filter(s => s.o >= 0.9);
        if (strokes.length && await UI.choose('Math', 'Read the handwriting on this page as math?', ['Read Handwriting', 'Type It Instead']) === 0) {
          UI.toast('Reading your handwriting…', 20000);
          try { src = (await Ocr.readInk(strokes)).text.replace(/\n/g, ' '); } catch { src = ''; }
          UI.toast(src ? 'Check what was read, then correct it if needed' : 'Couldn’t read it — type it instead', 3000);
        }
      }
      pane.hidden = false;
      document.getElementById('mathBtn').classList.add('on');
      if (src || !mInput.value) mInput.value = src;
      render();
      mInput.focus();
    }

    // ---------------- Web Clip ----------------
    async function download(url) {
      if (window.atlas && window.atlas.fetchData) {
        const r = await window.atlas.fetchData(url);
        if (r.error) throw new Error(r.error);
        return { bytes: new Uint8Array(r.bytes), contentType: r.contentType || '' };
      }
      let res;
      try { res = await fetch(url, { credentials: 'omit', cache: 'no-store' }); } catch { throw new Error('This site doesn’t let web pages read it. Clip it from the desktop app, or copy the article and paste it.'); }
      if (!res.ok) throw new Error(`The site answered ${res.status}`);
      return { bytes: new Uint8Array(await res.arrayBuffer()), contentType: res.headers.get('content-type') || '' };
    }
    async function pictureData(src) {
      try {
        const r = await download(src);
        const type = (r.contentType.split(';')[0] || '').trim().toLowerCase();
        if (!/^image\/(png|jpeg|gif|webp)$/.test(type) || r.bytes.length > 3e6) return null;
        return await Platform.bytesToDataUrl(r.bytes, type);
      } catch { return null; }
    }
    async function webClip() {
      const v = await UI.form('Web Clip', [
        { name: 'url', label: 'Web page address', value: 'https://' },
        { name: 'mode', label: 'Clip', type: 'select', options: [{ value: 'article', label: 'The article (text and pictures)' }, { value: 'bookmark', label: 'A bookmark (title, summary and link)' }] }
      ], 'Clip');
      if (!v) return;
      let url;
      try { url = new URL(v.url.trim()); } catch { return UI.toast('Enter a web address such as https://example.com/article'); }
      if (!/^https?:$/.test(url.protocol)) return UI.toast('Only http and https addresses can be clipped');
      UI.toast('Clipping…', 30000);
      let page;
      try {
        const r = await download(url.href);
        if (r.bytes.length > 10e6) throw new Error('That page is too large to clip');
        if (r.contentType && !/html|xml|text\/plain/i.test(r.contentType)) throw new Error('That address isn’t a web page');
        page = extractArticle(decode(r.bytes, r.contentType), url.href);
      } catch (err) { return UI.toast('Couldn’t clip that page: ' + err.message, 6000); }
      const when = new Date().toLocaleString();
      const source = `<p class="clip-src">Clipped from <a href="${esc(url.href)}">${esc(page.site || url.hostname)}</a> · ${esc(when)}</p>`;
      let html;
      if (v.mode === 'bookmark') {
        html = source + `<h2><a href="${esc(url.href)}">${esc(page.title)}</a></h2>${page.excerpt ? `<p>${esc(page.excerpt)}</p>` : ''}`;
      } else {
        // Pictures are kept in the notebook (the page never loads anything from the web).
        // (In an inert document: nothing in it loads or runs; the result is sanitized after.)
        const doc = document.implementation.createHTMLDocument('');
        doc.body.innerHTML = page.html;
        let kept = 0;
        for (const img of [...doc.body.querySelectorAll('img')]) {
          const data = kept < 12 ? await pictureData(img.getAttribute('src') || '') : null;
          if (data) { img.setAttribute('src', data); kept++; } else img.remove();
        }
        const holder = document.createElement('div');
        holder.innerHTML = ctx.sanitize(doc.body.innerHTML);
        holder.querySelectorAll('p, div, span, li').forEach(n => { if (!n.textContent.trim() && !n.querySelector('img, br, table')) n.remove(); });
        if (!holder.textContent.trim() && !kept) return UI.toast('No article was found on that page — try a bookmark instead', 5000);
        html = source + holder.innerHTML;
      }
      ctx.newPage(page.title, ctx.sanitize(html));
      UI.toast(v.mode === 'bookmark' ? 'Bookmark clipped' : 'Page clipped');
    }

    document.getElementById('inkToText').addEventListener('click', inkToText);
    document.getElementById('mathBtn').addEventListener('click', openMath);
    document.getElementById('picToText').addEventListener('click', pictureToText);
    document.getElementById('webClip').addEventListener('click', webClip);
    scheduleIndex();
    return { inkToText, pictureToText, openMath, webClip, indexPictures, scheduleIndex };
  }
  NotesSmart.extractArticle = extractArticle;
  NotesSmart.decode = decode;
  window.NotesSmart = NotesSmart;
})();
