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
      const text = await UI.form('New Comment', [{ name: 't', label: `Comment by ${cleanAuthor(o.author())} (type @ to mention someone)`, type: 'textarea', mentions: people }], 'Post');
      if (!text || !text.t.trim()) return;
      const span = document.createElement('span');
      span.className = 'comment';
      span.title = text.t.trim();
      span.dataset.author = cleanAuthor(o.author());
      span.dataset.date = new Date().toISOString();
      rememberPeople(UI.mentioned(span.title, people()));
      try { range.surroundContents(span); } catch { span.appendChild(range.extractContents()); range.insertNode(span); }
      o.changed();
      showPane(true);
    }
    // ---------------- @mentions ----------------
    // People to suggest: whoever is in the live session, everyone who has commented, and names
    // mentioned before. A comment that mentions you is marked, and arriving in a live session it
    // tells you so.
    const store = (k, v) => { try { if (v === undefined) return JSON.parse(localStorage.getItem(k) || '[]'); localStorage.setItem(k, JSON.stringify(v)); } catch { return []; } return v; };
    const myName = () => { let n = ''; try { n = localStorage.getItem('atlas-user-name') || ''; } catch { /* none */ } return n || cleanAuthor(o.author()); };
    function people() {
      const names = [...(o.peers ? o.peers() : []), ...comments().map(c => c.dataset.author), ...store('atlas-people')];
      return [...new Set(names.filter(n => n && n !== 'Author' && n !== myName()))];
    }
    function rememberPeople(names) { if (names.length) store('atlas-people', [...new Set([...names, ...store('atlas-people')])].slice(0, 200)); }
    const mentionsMe = c => { const me = myName(); return !!me && me !== 'Author' && (' ' + (c.title || '').toLowerCase()).includes('@' + me.toLowerCase()); };
    function showMentions(el, text) {
      const names = [...UI.mentioned(text, [...people(), myName()])].sort((a, b) => b.length - a.length);
      el.textContent = '';
      if (!names.length) { el.textContent = text; return; }
      const re = new RegExp('(@(?:' + names.map(n => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + '))', 'gi');
      String(text).split(re).forEach((part, i) => {
        if (i % 2) { const m = document.createElement('span'); m.className = 'mention'; m.textContent = part; el.appendChild(m); } else if (part) el.appendChild(document.createTextNode(part));
      });
    }
    /** Tells you about comments that mention you which you haven't been told about. */
    function checkMentions() {
      const seen = new Set(store('atlas-seen-mentions'));
      const fresh = comments().filter(c => mentionsMe(c) && !seen.has((c.dataset.date || '') + '|' + c.dataset.author));
      if (!fresh.length) return [];
      fresh.forEach(c => seen.add((c.dataset.date || '') + '|' + c.dataset.author));
      store('atlas-seen-mentions', [...seen].slice(-500));
      const c = fresh[fresh.length - 1];
      UI.toast(`${c.dataset.author || 'Someone'} mentioned you: “${c.title.slice(0, 80)}”`, 5000);
      showPane(true);
      return fresh;
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
        showMentions(card.querySelector('.text'), c.title);
        if (mentionsMe(c)) card.classList.add('mentions-me');
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
      const r = await UI.form('Edit Comment', [{ name: 't', label: 'Comment', type: 'textarea', value: c.title, mentions: people }], 'Save');
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

    // ---------------- A rewrite as tracked changes ----------------
    // The Assistant's rewrite of a passage goes in as the edits a person would have made: the
    // paragraphs are lined up, unchanged ones are left alone, changed ones get word-level
    // insertions and deletions, and added or removed paragraphs are tracked whole. One Ctrl+Z
    // takes it all back.
    const words = t => t.match(/\s+|[^\s]+/g) || [];
    const norm = t => t.replace(/\s+/g, ' ').trim();
    /** Word-level differences: [['eq'|'del'|'ins', text], …]. */
    function diffWords(a, b) {
      const A = words(a), B = words(b), n = A.length, m = B.length;
      if (n * m > 4e6) return [['del', a], ['ins', b]];
      const L = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
      for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = A[i] === B[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
      const out = [];
      const add = (k, t) => { const last = out[out.length - 1]; if (last && last[0] === k) last[1] += t; else out.push([k, t]); };
      let i = 0, j = 0;
      while (i < n && j < m) {
        if (A[i] === B[j]) { add('eq', A[i]); i++; j++; }
        else if (L[i + 1][j] >= L[i][j + 1]) add('del', A[i++]);
        else add('ins', B[j++]);
      }
      while (i < n) add('del', A[i++]);
      while (j < m) add('ins', B[j++]);
      // Changes separated only by spaces read as one change: "a lot of money" → "about $40,000",
      // not five alternating pieces.
      const res = [];
      for (let k = 0; k < out.length;) {
        if (out[k][0] === 'eq') { res.push(out[k++]); continue; }
        let del = '', ins = '', end = k;
        for (let x = k; x < out.length; x++) {
          const [kind, t] = out[x];
          if (kind === 'eq' && !(/^\s+$/.test(t) && out[x + 1] && out[x + 1][0] !== 'eq')) break;
          if (kind !== 'ins') del += t;
          if (kind !== 'del') ins += t;
          end = x + 1;
        }
        if (del) res.push(['del', del]);
        if (ins) res.push(['ins', ins]);
        k = end;
      }
      return res;
    }
    const opsHtml = ops => ops.map(([k, t]) => (k === 'eq' ? esc(t) : `<${k} ${attrs()}>${esc(t)}</${k}>`)).join('');
    const SIMPLE = /^(P|H[1-6]|BLOCKQUOTE|LI)$/;
    /** An element with all its text marked as deleted (structure kept). */
    function deletedCopy(el) {
      const c = el.cloneNode(true);
      const w = document.createTreeWalker(c, NodeFilter.SHOW_TEXT);
      const nodes = [];
      while (w.nextNode()) if (w.currentNode.nodeValue.trim() && !w.currentNode.parentElement.closest('del.tc')) nodes.push(w.currentNode);
      nodes.forEach(n => { const d = document.createElement('del'); d.className = 'tc'; d.dataset.author = cleanAuthor(o.author()); d.setAttribute('datetime', new Date().toISOString().replace(/\.\d+Z$/, 'Z')); n.replaceWith(d); d.appendChild(n); });
      return c.outerHTML;
    }
    /** A new element with its content marked as inserted. */
    function insertedCopy(el) {
      const c = el.cloneNode(true);
      if (/^(UL|OL)$/.test(c.tagName)) { [...c.children].forEach(li => { li.innerHTML = `<ins ${attrs()}>${li.innerHTML}</ins>`; }); return c.outerHTML; }
      if (c.tagName === 'TABLE') { c.querySelectorAll('td,th').forEach(td => { if (td.innerHTML.trim()) td.innerHTML = `<ins ${attrs()}>${td.innerHTML}</ins>`; }); return c.outerHTML; }
      c.innerHTML = `<ins ${attrs()}>${c.innerHTML}</ins>`;
      return c.outerHTML;
    }
    /**
     * Replaces whole blocks (top-level elements, in order) with new ones (from the reply's HTML)
     * as tracked changes. Returns false if there was nothing to change.
     */
    function replaceBlocksTracked(oldBlocks, newHtml) {
      const t = document.createElement('template');
      t.innerHTML = newHtml;
      const fresh = [...t.content.children];
      const A = oldBlocks.map(el => norm(el.textContent)), B = fresh.map(el => norm(el.textContent));
      // Line the paragraphs up: the longest run of unchanged ones stays as it is.
      const n = A.length, m = B.length;
      const L = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
      for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = A[i] === B[j] && A[i] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
      const parts = [];
      let i = 0, j = 0, changes = 0;
      const flush = (olds, news) => {
        const k = Math.min(olds.length, news.length);
        for (let x = 0; x < k; x++) {
          const a = olds[x], b = news[x];
          if (SIMPLE.test(a.tagName) && SIMPLE.test(b.tagName)) {
            // The paragraph keeps its kind (a heading stays a heading); its words are tracked.
            const el = a.cloneNode(false);
            el.innerHTML = opsHtml(diffWords(a.textContent, b.textContent));
            parts.push(el.outerHTML);
          } else { parts.push(deletedCopy(a), insertedCopy(b)); }
          changes++;
        }
        olds.slice(k).forEach(a => { parts.push(deletedCopy(a)); changes++; });
        news.slice(k).forEach(b => { parts.push(insertedCopy(b)); changes++; });
      };
      let olds = [], news = [];
      while (i < n || j < m) {
        if (i < n && j < m && A[i] === B[j] && A[i]) { flush(olds, news); olds = []; news = []; parts.push(oldBlocks[i].outerHTML); i++; j++; }
        else if (j < m && (i >= n || L[i][j + 1] >= L[i + 1][j])) news.push(fresh[j++]);
        else olds.push(oldBlocks[i++]);
      }
      flush(olds, news);
      if (!changes) return false;
      const was = tracking; tracking = false;
      try { replaceBlocks(oldBlocks, parts.join('')); } finally { tracking = was; }
      o.changed();
      return true;
    }
    /**
     * Puts HTML in place of whole blocks. Done on the document directly (the browser's insertHTML
     * merges and restyles neighbouring paragraphs), through o.edit so that one Ctrl+Z takes it back.
     */
    function replaceBlocks(blocks, html) {
      const run = () => {
        const t = document.createElement('template');
        t.innerHTML = html;
        const nodes = [...t.content.childNodes];
        blocks[0].before(...nodes);
        blocks.forEach(b => b.remove());
        const last = nodes.filter(n => n.nodeType === 1).pop();
        if (last) { const r = document.createRange(); r.selectNodeContents(last); r.collapse(false); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r); }
      };
      if (o.edit) o.edit(run); else run();
    }
    /** Replaces the text of a range inside one paragraph with new text, as tracked word changes. */
    function replaceTextTracked(range, text) {
      const ops = diffWords(range.toString(), text);
      if (ops.every(x => x[0] === 'eq')) return false;
      const sel = getSelection(); sel.removeAllRanges(); sel.addRange(range);
      const was = tracking; tracking = false;
      try { document.execCommand('insertHTML', false, opsHtml(ops)); } finally { tracking = was; }
      o.changed();
      return true;
    }

    return {
      replaceBlocksTracked, replaceTextTracked, replaceBlocks, diffWords,
      get tracking() { return tracking; },
      setTracking(on) { tracking = !!on; },
      insertText, deleteSelection,
      acceptAtCursor: () => act(accept, false), rejectAtCursor: () => act(reject, false),
      acceptAll: () => act(accept, true), rejectAll: () => act(reject, true),
      next: () => next(false), previous: () => next(true),
      count: () => changes().length,
      addComment, renderPane, showPane, comments, checkMentions, people,
      markCommentsForDocx, patchDocx, preprocessDocx, postprocessHtml
    };
  }

  window.WriteReview = WriteReview;
})();
