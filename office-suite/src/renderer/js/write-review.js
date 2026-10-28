'use strict';
// Isadek Write — Review: tracked changes and comments, including their round trip through .docx.
//
// In the document a tracked insertion is <ins class="tc">, a deletion <del class="tc">, both
// with data-author and datetime; a comment is <span class="comment" title="text"
// data-author data-date>. Edits go through execCommand where possible so Ctrl+Z still works.
(function () {
  const MARK = {
    insS: '\uE010', insE: '\uE011', delS: '\uE012', delE: '\uE013',
    cmtS: '\uE020', cmtE: '\uE021', sep: '\uE0FF'
  };
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const cleanAuthor = a => String(a || '').replace(/[^\w .@'-]/g, '').slice(0, 64) || 'Author';
  const validDate = d => (d && !isNaN(Date.parse(d)) ? new Date(d).toISOString() : new Date().toISOString());

  function WriteReview(o) {
    const { doc } = o;
    let tracking = false;

    // ---------------- Tracked changes ----------------
    const attrs = () => `class="tc" data-author="${esc(o.author())}" datetime="${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}"`;
    const inTc = (node, tag) => {
      let n = node && (node.nodeType === 3 ? node.parentElement : node);
      while (n && n !== doc) { if (n.tagName === tag && n.classList.contains('tc')) return n; n = n.parentElement; }
      return null;
    };
    const mine = el => el && el.dataset.author === cleanAuthor(o.author());

    /** Marks the selection as deleted (or removes it if it is the author's own tracked insertion). */
    function deleteSelection(backward) {
      const sel = getSelection();
      if (!sel.rangeCount || sel.isCollapsed) return;
      const range = sel.getRangeAt(0);
      const ownIns = inTc(range.commonAncestorContainer, 'INS');
      if (ownIns && mine(ownIns)) { document.execCommand('delete'); return; }
      // Already-deleted text stays as it is.
      if (inTc(range.commonAncestorContainer, 'DEL')) { collapseOutside(inTc(range.commonAncestorContainer, 'DEL'), backward); return; }
      const frag = range.cloneContents();
      // Text inside the deletion that was already deleted stays a single deletion.
      frag.querySelectorAll && frag.querySelectorAll('del.tc').forEach(d => d.replaceWith(...d.childNodes));
      const blocks = frag.querySelectorAll ? frag.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,div,blockquote,pre,table') : [];
      if (blocks.length) {
        // Across paragraphs: wrap each text piece in place (paragraphs themselves are kept).
        wrapTextInRange(range, 'del');
        sel.collapse(backward ? range.startContainer : range.endContainer, backward ? range.startOffset : range.endOffset);
        o.changed();
        return;
      }
      const holder = document.createElement('div');
      holder.appendChild(frag);
      if (!holder.textContent && !holder.querySelector('img')) { document.execCommand('delete'); return; }
      const id = 'tc' + Math.random().toString(36).slice(2, 9);
      document.execCommand('insertHTML', false, `<del ${attrs()} data-tmp="${id}">${holder.innerHTML}</del>`);
      let del = doc.querySelector(`del[data-tmp="${id}"]`);
      if (del) {
        del.removeAttribute('data-tmp');
        // Typing Backspace next to an existing deletion can nest the new one inside it; flatten.
        const outer = del.parentElement && del.parentElement.closest('del.tc');
        if (outer) { del.replaceWith(...del.childNodes); del = outer; }
        del.querySelectorAll('del.tc').forEach(d => d.replaceWith(...d.childNodes));
        collapseOutside(del, backward);
      }
      o.changed();
    }
    function collapseOutside(el, before) {
      const r = document.createRange();
      if (before) r.setStartBefore(el); else r.setStartAfter(el);
      r.collapse(true);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
    }
    /** Wraps every text node (part) inside a range in a tracked element. */
    function wrapTextInRange(range, tag) {
      const walker = document.createTreeWalker(range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer, NodeFilter.SHOW_TEXT);
      const nodes = [];
      while (walker.nextNode()) if (range.intersectsNode(walker.currentNode) && walker.currentNode.nodeValue) nodes.push(walker.currentNode);
      for (let n of nodes) {
        if (inTc(n, 'DEL')) continue;
        const own = inTc(n, 'INS');
        let start = n === range.startContainer ? range.startOffset : 0;
        let end = n === range.endContainer ? range.endOffset : n.nodeValue.length;
        if (start >= end) continue;
        if (own && mine(own) && tag === 'del') { n.nodeValue = n.nodeValue.slice(0, start) + n.nodeValue.slice(end); continue; }
        if (end < n.nodeValue.length) n.splitText(end);
        if (start > 0) n = n.splitText(start);
        const w = document.createElement(tag);
        w.className = 'tc';
        w.dataset.author = cleanAuthor(o.author());
        w.setAttribute('datetime', new Date().toISOString().replace(/\.\d+Z$/, 'Z'));
        n.replaceWith(w);
        w.appendChild(n);
      }
      doc.querySelectorAll('ins.tc:empty').forEach(e => e.remove());
    }
    /** Inserts typed text as a tracked insertion. */
    function insertText(text) {
      if (!text) return;
      const sel = getSelection();
      if (!sel.rangeCount) return;
      if (!sel.isCollapsed) deleteSelection(false);
      const del = inTc(sel.anchorNode, 'DEL');
      if (del) collapseOutside(del, false);
      const ins = inTc(sel.anchorNode, 'INS');
      if (ins && mine(ins)) { document.execCommand('insertText', false, text); return; }
      document.execCommand('insertHTML', false, `<ins ${attrs()}>${esc(text).replace(/\n/g, '<br>')}</ins>`);
      o.changed();
    }
    doc.addEventListener('beforeinput', e => {
      if (!tracking) return;
      const t = e.inputType;
      if (t === 'insertText' || t === 'insertReplacementText') {
        e.preventDefault();
        insertText(e.data != null ? e.data : (e.dataTransfer ? e.dataTransfer.getData('text/plain') : ''));
      } else if (t === 'insertFromDrop' || t === 'deleteByDrag') {
        e.preventDefault(); // moving text by drag isn't tracked; keep the document unchanged
      } else if (t.startsWith('delete')) {
        e.preventDefault();
        const sel = getSelection();
        const backward = /Backward/.test(t);
        if (sel.isCollapsed) sel.modify('extend', backward ? 'backward' : 'forward', /Word/.test(t) ? 'word' : /Line|Soft|Hard/.test(t) ? 'lineboundary' : 'character');
        deleteSelection(backward);
      }
    });
    function changes() { return [...doc.querySelectorAll('ins.tc, del.tc')]; }
    function accept(el) { if (el.tagName === 'INS') el.replaceWith(...el.childNodes); else el.remove(); }
    function reject(el) { if (el.tagName === 'INS') el.remove(); else el.replaceWith(...el.childNodes); }
    /** The change at the caret, or all changes in the selection. */
    function changesAtSelection() {
      const sel = getSelection();
      if (!sel.rangeCount) return [];
      const r = sel.getRangeAt(0);
      const at = inTc(r.startContainer, 'INS') || inTc(r.startContainer, 'DEL');
      const list = changes().filter(c => r.intersectsNode(c));
      return list.length ? list : at ? [at] : [];
    }
    function act(fn, all) {
      const list = all ? changes() : changesAtSelection();
      if (!list.length) { UI.toast(all ? 'There are no tracked changes' : 'Place the cursor in a tracked change'); return; }
      list.forEach(fn);
      doc.normalize();
      o.changed();
      if (!all) next();
    }
    /** Selects the next (or previous) tracked change after the current selection, wrapping around. */
    function next(back) {
      const list = changes();
      if (!list.length) { UI.toast('There are no tracked changes'); return; }
      const sel = getSelection();
      const cur = sel.rangeCount && doc.contains(sel.anchorNode) ? sel.getRangeAt(0) : null;
      let target;
      if (!cur) target = back ? list[list.length - 1] : list[0];
      else if (back) target = [...list].reverse().find(c => cur.comparePoint(c, c.childNodes.length) === -1) || list[list.length - 1];
      else target = list.find(c => cur.comparePoint(c, 0) === 1) || list[0];
      const r = document.createRange();
      r.selectNodeContents(target);
      sel.removeAllRanges();
      sel.addRange(r);
      target.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }

    // ---------------- Comments ----------------
    const pane = o.pane;
    function comments() { return [...doc.querySelectorAll('span.comment')]; }
    async function addComment() {
      const sel = getSelection();
      if (!sel.rangeCount || sel.isCollapsed || !doc.contains(sel.anchorNode)) { UI.toast('Select some text to comment on'); return; }
      const range = sel.getRangeAt(0).cloneRange();
      const text = await UI.form('New Comment', [{ name: 't', label: `Comment by ${cleanAuthor(o.author())}`, type: 'textarea' }], 'Post');
      if (!text || !text.t.trim()) return;
      const span = document.createElement('span');
      span.className = 'comment';
      span.title = text.t.trim();
      span.dataset.author = cleanAuthor(o.author());
      span.dataset.date = new Date().toISOString();
      try { range.surroundContents(span); } catch { span.appendChild(range.extractContents()); range.insertNode(span); }
      o.changed();
      showPane(true);
    }
    function renderPane() {
      const list = comments();
      pane.querySelector('.list').innerHTML = '';
      if (!list.length) { pane.querySelector('.list').innerHTML = '<p class="empty">No comments. Select text and choose New Comment.</p>'; return; }
      list.forEach(c => {
        const card = document.createElement('div');
        card.className = 'card';
        card.innerHTML = '<div class="who"><b></b> <span class="when"></span></div><div class="quote"></div><div class="text"></div><div class="acts"><button data-a="edit">Edit</button><button data-a="resolve">Resolve</button></div>';
        card.querySelector('b').textContent = c.dataset.author || 'Author';
        card.querySelector('.when').textContent = c.dataset.date ? new Date(c.dataset.date).toLocaleString() : '';
        card.querySelector('.quote').textContent = '“' + c.textContent.slice(0, 80) + (c.textContent.length > 80 ? '…' : '') + '”';
        card.querySelector('.text').textContent = c.title;
        card.addEventListener('click', e => {
          const a = e.target.dataset.a;
          if (a === 'edit') { editComment(c); return; }
          if (a === 'resolve') { c.replaceWith(...c.childNodes); o.changed(); renderPane(); return; }
          c.scrollIntoView({ block: 'center', behavior: 'smooth' });
          const r = document.createRange(); r.selectNodeContents(c);
          getSelection().removeAllRanges(); getSelection().addRange(r);
        });
        pane.querySelector('.list').appendChild(card);
      });
    }
    async function editComment(c) {
      const r = await UI.form('Edit Comment', [{ name: 't', label: 'Comment', type: 'textarea', value: c.title }], 'Save');
      if (!r) return;
      if (!r.t.trim()) c.replaceWith(...c.childNodes); else c.title = r.t.trim();
      o.changed();
      renderPane();
    }
    function showPane(on) { pane.hidden = !on; if (on) renderPane(); }

    // ---------------- .docx export ----------------
    /** Replaces comment spans with private-use markers that survive HTML-to-Word conversion. */
    function markCommentsForDocx(root) {
      const list = [];
      root.querySelectorAll('span.comment').forEach((c, i) => {
        list.push({ id: i, author: cleanAuthor(c.dataset.author), date: validDate(c.dataset.date), text: c.title || '' });
        c.before(document.createTextNode(MARK.cmtS + i + MARK.sep));
        c.after(document.createTextNode(MARK.cmtE + i + MARK.sep));
        c.replaceWith(...c.childNodes);
      });
      return list;
    }
    /** Turns the markers in word/document.xml into comment ranges and references. */
    async function patchDocx(bytes) {
      const zip = await JSZip.loadAsync(bytes);
      const f = zip.file('word/document.xml');
      let xml = await f.async('string');
      if (!/[\uE020\uE021\uE030\uE040]/.test(xml)) return bytes;
      // Mail-merge fields: a marked run becomes a MERGEFIELD showing «Name».
      xml = xml.replace(/<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g, run => {
        if (!/\uE040/.test(run)) return run;
        const rPr = (run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
        const t = run.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/);
        if (!t) return run;
        const re = /\uE040([\w .'-]{1,64})\uE0FF([\s\S]*?)\uE041\uE0FF/g;
        let out = '', last = 0, m;
        const piece = s => (s ? `<w:r>${rPr}<w:t xml:space="preserve">${s}</w:t></w:r>` : '');
        while ((m = re.exec(t[1]))) {
          out += piece(t[1].slice(last, m.index));
          const name = /\s/.test(m[1]) ? `&quot;${m[1]}&quot;` : m[1];
          out += `<w:fldSimple w:instr=" MERGEFIELD ${name} ">${piece(m[2])}</w:fldSimple>`;
          last = re.lastIndex;
        }
        out += piece(t[1].slice(last).replace(/[\uE040\uE041][^\uE0FF]*\uE0FF/g, ''));
        return out;
      });
      // Citations: a marked run becomes a CITATION field showing the same text.
      xml = xml.replace(/<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g, run => {
        if (!/\uE030/.test(run)) return run;
        const rPr = (run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
        const t = run.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/);
        if (!t) return run;
        const re = /\uE030(s[a-z0-9]{4,24})\|([\w.,–-]*)\uE0FF([\s\S]*?)\uE031\uE0FF/g;
        let out = '', last = 0, m;
        const piece = s => (s ? `<w:r>${rPr}<w:t xml:space="preserve">${s}</w:t></w:r>` : '');
        while ((m = re.exec(t[1]))) {
          out += piece(t[1].slice(last, m.index));
          out += `<w:fldSimple w:instr=" CITATION ${m[1]}${m[2] ? ' \\p ' + m[2] : ''} ">${piece(m[3])}</w:fldSimple>`;
          last = re.lastIndex;
        }
        out += piece(t[1].slice(last).replace(/[\uE030\uE031][^\uE0FF]*\uE0FF/g, ''));
        return out;
      });
      xml = xml.replace(/<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g, run => {
        if (!/[\uE020\uE021]/.test(run)) return run;
        const rPr = (run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
        const tMatch = run.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/);
        if (!tMatch) return run;
        const out = [];
        const re = /([\uE020\uE021])(\d+)\uE0FF/g;
        let last = 0, m;
        const text = tMatch[1];
        const piece = s => { if (s) out.push(`<w:r>${rPr}<w:t xml:space="preserve">${s}</w:t></w:r>`); };
        while ((m = re.exec(text))) {
          piece(text.slice(last, m.index));
          out.push(m[1] === MARK.cmtS ? `<w:commentRangeStart w:id="${m[2]}"/>` : `<w:commentRangeEnd w:id="${m[2]}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${m[2]}"/></w:r>`);
          last = re.lastIndex;
        }
        piece(text.slice(last));
        return out.join('');
      });
      zip.file('word/document.xml', xml);
      return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
    }

    // ---------------- .docx import ----------------
    /** Rewrites a .docx so the HTML converter keeps revisions and comment anchors as markers. */
    async function preprocessDocx(bytes) {
      const zip = await JSZip.loadAsync(bytes);
      const f = zip.file('word/document.xml');
      if (!f) return { bytes, meta: null };
      let xml = await f.async('string');
      const meta = { changes: {}, comments: {} };
      let n = 0;
      const mk = (s, id) => `<w:r><w:t xml:space="preserve">${s}${id}${MARK.sep}</w:t></w:r>`;
      const attr = (tag, name) => { const m = tag.match(new RegExp(`w:${name}="([^"]*)"`)); return m ? m[1] : ''; };
      // Deletions: show the deleted text (as normal runs) between markers.
      xml = xml.replace(/<w:del\b([^>]*[^/])>([\s\S]*?)<\/w:del>/g, (all, a, inner) => {
        const id = ++n;
        meta.changes[id] = { type: 'del', author: attr(a, 'author'), date: attr(a, 'date') };
        return mk(MARK.delS, id) + inner.replace(/<w:delText\b/g, '<w:t').replace(/<\/w:delText>/g, '</w:t>') + mk(MARK.delE, id);
      });
      xml = xml.replace(/<w:ins\b([^>]*[^/])>([\s\S]*?)<\/w:ins>/g, (all, a, inner) => {
        const id = ++n;
        meta.changes[id] = { type: 'ins', author: attr(a, 'author'), date: attr(a, 'date') };
        return mk(MARK.insS, id) + inner + mk(MARK.insE, id);
      });
      const cf = zip.file('word/comments.xml');
      if (cf) {
        const cx = new DOMParser().parseFromString(await cf.async('string'), 'application/xml');
        const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        [...cx.getElementsByTagNameNS(W, 'comment')].forEach(c => {
          const g = name => c.getAttributeNS(W, name) || c.getAttribute('w:' + name) || '';
          meta.comments[g('id')] = { author: g('author'), date: g('date'), text: [...c.getElementsByTagNameNS(W, 'p')].map(p => p.textContent).join('\n') };
        });
        xml = xml.replace(/<w:commentRangeStart\b[^>]*w:id="(\d+)"[^>]*\/>/g, (all, id) => mk(MARK.cmtS, id))
          .replace(/<w:commentRangeEnd\b[^>]*w:id="(\d+)"[^>]*\/>/g, (all, id) => mk(MARK.cmtE, id))
          .replace(/<w:r\b[^>]*>(?:(?!<\/w:r>)[\s\S])*?<w:commentReference\b[^>]*\/>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g, '');
      }
      // Citation fields written by Isadek (CITATION <source> \p <pages>).
      meta.cites = {};
      xml = xml.replace(/<w:fldSimple\b[^>]*w:instr="\s*CITATION\s+(\S+)(?:\s+\\p\s+(\S+))?[^"]*"[^>]*>([\s\S]*?)<\/w:fldSimple>/g, (all, src, pages, inner) => {
        const id = ++n;
        meta.cites[id] = { src, pages: pages || '' };
        return mk('\uE030', id) + inner + mk('\uE031', id);
      });
      meta.fields = {};
      xml = xml.replace(/<w:fldSimple\b[^>]*w:instr="\s*MERGEFIELD\s+(?:&quot;([^&]+)&quot;|(\S+))[^"]*"[^>]*>([\s\S]*?)<\/w:fldSimple>/g, (all, q, plain, inner) => {
        const id = ++n;
        meta.fields[id] = (q || plain || '').replace(/[^\w .'-]/g, '').slice(0, 64);
        return mk('\uE040', id) + inner + mk('\uE041', id);
      });
      if (!n && !Object.keys(meta.comments).length) return { bytes, meta: null };
      zip.file('word/document.xml', xml);
      return { bytes: new Uint8Array(await zip.generateAsync({ type: 'uint8array' })), meta };
    }
    /** Turns the markers in the converted HTML back into <ins>, <del> and comment spans. */
    function postprocessHtml(html, meta) {
      if (!meta) return html;
      const d = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
      // 1. Replace each marker text with an empty boundary element.
      const walker = d.createTreeWalker(d.body, NodeFilter.SHOW_TEXT);
      const texts = [];
      while (walker.nextNode()) if (/[\uE010-\uE041]/.test(walker.currentNode.nodeValue)) texts.push(walker.currentNode);
      texts.forEach(t => {
        const parts = t.nodeValue.split(/([\uE010\uE011\uE012\uE013\uE020\uE021\uE030\uE031\uE040\uE041]\d+\uE0FF)/);
        const frag = d.createDocumentFragment();
        parts.forEach(p => {
          const m = /^([\uE010-\uE041])(\d+)\uE0FF$/.exec(p);
          if (m) { const b = d.createElement('x-mark'); b.setAttribute('k', m[1]); b.setAttribute('i', m[2]); frag.appendChild(b); }
          else if (p) frag.appendChild(d.createTextNode(p));
        });
        t.replaceWith(frag);
      });
      // 2. Wrap the text between each start/end pair.
      const pairs = [[MARK.insS, MARK.insE], [MARK.delS, MARK.delE], [MARK.cmtS, MARK.cmtE], ['\uE030', '\uE031'], ['\uE040', '\uE041']];
      for (const [s, e] of pairs) {
        d.querySelectorAll(`x-mark[k="${s}"]`).forEach(start => {
          const id = start.getAttribute('i');
          const end = d.querySelector(`x-mark[k="${e}"][i="${id}"]`);
          if (!end) return;
          const range = d.createRange();
          range.setStartAfter(start);
          range.setEndBefore(end);
          const w = d.createTreeWalker(range.commonAncestorContainer, NodeFilter.SHOW_TEXT);
          const nodes = [];
          while (w.nextNode()) if (range.intersectsNode(w.currentNode) && w.currentNode.nodeValue) nodes.push(w.currentNode);
          nodes.forEach(n => {
            let el;
            if (s === '\uE040') {
              // A merge field is one element showing «Name»; drop the other text pieces.
              if (n !== nodes[0]) { n.remove(); return; }
              el = d.createElement('span');
              el.className = 'mergefield';
              el.dataset.field = (meta.fields || {})[id] || 'Field';
              n.nodeValue = `«${el.dataset.field}»`;
            } else if (s === '\uE030') {
              const c = (meta.cites || {})[id] || {};
              el = d.createElement('cite');
              el.className = 'citation';
              if (/^s[a-z0-9]{4,24}$/.test(c.src || '')) el.dataset.src = c.src;
              if (c.pages && /^[\w.,–-]{1,40}$/.test(c.pages)) el.dataset.pages = c.pages;
            } else if (s === MARK.cmtS) {
              const c = meta.comments[id] || {};
              el = d.createElement('span');
              el.className = 'comment';
              el.title = c.text || '';
              el.dataset.author = cleanAuthor(c.author);
              el.dataset.date = validDate(c.date);
            } else {
              const c = meta.changes[id] || {};
              el = d.createElement(s === MARK.insS ? 'ins' : 'del');
              el.className = 'tc';
              el.dataset.author = cleanAuthor(c.author);
              el.setAttribute('datetime', validDate(c.date).replace(/\.\d+Z$/, 'Z'));
            }
            n.replaceWith(el);
            el.appendChild(n);
          });
        });
      }
      d.querySelectorAll('x-mark').forEach(m => m.remove());
      return d.body.innerHTML;
    }

    return {
      get tracking() { return tracking; },
      setTracking(on) { tracking = !!on; },
      insertText, deleteSelection,
      acceptAtCursor: () => act(accept, false), rejectAtCursor: () => act(reject, false),
      acceptAll: () => act(accept, true), rejectAll: () => act(reject, true),
      next: () => next(false), previous: () => next(true),
      count: () => changes().length,
      addComment, renderPane, showPane, comments,
      markCommentsForDocx, patchDocx, preprocessDocx, postprocessHtml
    };
  }

  window.WriteReview = WriteReview;
})();
