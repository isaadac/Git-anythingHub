'use strict';
// Atlas Write — table of contents.
// <div class="toc"> holds a title and one entry per heading:
//   <p class="toc-l2"><a href="#toc_x">Heading</a><span class="toc-pg">3</span></p>
// Page numbers follow the layout; Update Table rebuilds the entries from the headings.
// Entries are links, so the PDF's contents are clickable; .docx gets a real Word TOC field.
(function () {
  const HEADINGS = 'h1:not(.title),h2,h3';
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const newId = i => 'toc_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + i;
  const TOC_TITLE = /^(table of )?contents$/i;

  /** Entries for the headings of `root`; headings without an id get one. */
  function entriesHtml(root) {
    const heads = [...root.querySelectorAll(HEADINGS)].filter(h => !h.closest('.toc') && h.textContent.trim());
    heads.forEach((h, i) => { if (!h.id) h.id = newId(i); });
    return heads.map(h => `<p class="toc-l${h.tagName[1]}"><a href="#${h.id}">${esc(h.textContent.trim())}</a><span class="toc-pg"></span></p>`).join('');
  }
  const tocHtml = (root, title = 'Contents') =>
    `<div class="toc" contenteditable="false"><p class="toc-title">${esc(title)}</p>${entriesHtml(root) || '<p class="toc-l1">No headings yet: use Heading 1–3 (Home → Styles), then Update Table.</p>'}</div>`;

  function WriteToc(o) {
    const { doc } = o;

    /** Inserts a table of contents at the paragraph with the cursor (before it when the cursor is at its start), or updates the one there is. */
    function insert() {
      if (doc.querySelector('.toc')) { update(); UI.toast('Table of contents updated'); return; }
      if (!doc.querySelector(HEADINGS)) { UI.toast('Add some headings first (Home → Styles)'); return; }
      o.restoreSelection();
      const sel = getSelection();
      let block = sel.rangeCount && doc.contains(sel.anchorNode) ? sel.anchorNode : null;
      while (block && block.parentNode !== doc) block = block.parentNode;
      const wrap = document.createElement('div');
      wrap.innerHTML = tocHtml(doc);
      const toc = wrap.firstChild;
      if (!block) doc.prepend(toc);
      else {
        const before = document.createRange();
        before.setStart(block, 0); before.setEnd(sel.anchorNode, sel.anchorOffset);
        if (before.toString().trim()) block.after(toc); else block.before(toc);
      }
      o.changed();
      o.relayout();
    }
    /** Rebuilds the entries from the headings (keeping the title). */
    function update() {
      doc.querySelectorAll('.toc').forEach(t => {
        const title = t.querySelector('.toc-title, p b')?.textContent.trim() || 'Contents';
        const wrap = document.createElement('div');
        wrap.innerHTML = tocHtml(doc, title);
        t.replaceWith(wrap.firstChild);
      });
      o.changed();
      o.relayout();
    }
    /** Fills in page numbers from the current layout; doesn't count as an edit. */
    function refreshPages() {
      const tocs = doc.querySelectorAll('.toc');
      if (!tocs.length) return;
      const z = o.zoom(), top = doc.getBoundingClientRect().top;
      tocs.forEach(t => {
        if (t.contentEditable !== 'false') t.contentEditable = 'false';
        t.querySelectorAll('p').forEach(p => {
          const a = p.querySelector('a[href^="#"]'), pg = p.querySelector('.toc-pg');
          if (!a || !pg) return;
          const target = document.getElementById(decodeURIComponent(a.getAttribute('href').slice(1)));
          const n = target && doc.contains(target) ? String(o.pageAt((target.getBoundingClientRect().top - top) / z)) : '';
          if (pg.textContent !== n) pg.textContent = n;
        });
      });
    }
    // The contents isn't editable text, so a plain click follows the entry (Ctrl+click elsewhere).
    doc.addEventListener('click', e => {
      const a = e.target.closest('.toc a[href^="#"]');
      if (!a) return;
      e.preventDefault();
      const target = document.getElementById(decodeURIComponent(a.getAttribute('href').slice(1)));
      if (!target) return;
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      const r = document.createRange(); r.setStart(target, 0); r.collapse(true);
      const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
    });

    /** Marks the contents so patchDocx can replace it with a Word TOC field. */
    function prepareForDocx(root) {
      root.querySelectorAll('.toc').forEach(t => {
        const title = t.querySelector('.toc-title, p b')?.textContent.trim() || 'Contents';
        const rows = [...t.querySelectorAll('p')].flatMap(p => {
          const a = p.querySelector('a[href^="#"]');
          if (!a) return [];
          const level = Math.min(3, Math.max(1, +((p.className.match(/toc-l(\d)/) || [])[1] || (parseFloat(p.style.marginLeft) / 18 + 1) || 1)));
          const page = (p.querySelector('.toc-pg')?.textContent || '').replace(/\D/g, '');
          return [`<p><a href="${esc(a.getAttribute('href'))}">\uE052${level}|${page}\uE0FF${esc(a.textContent)}</a></p>`];
        });
        const wrap = document.createElement('div');
        wrap.innerHTML = `<p>\uE050${esc(title)}\uE0FF</p>${rows.join('')}<p>\uE051\uE0FF</p>`;
        t.replaceWith(...wrap.childNodes);
      });
    }

    return { insert, update, refreshPages, prepareForDocx };
  }

  /** Replaces the marked paragraphs in word/document.xml with a Word table of contents. */
  WriteToc.patchDocx = async function (bytes, { tabPos = 9360 } = {}) {
    const zip = await JSZip.loadAsync(bytes);
    const f = zip.file('word/document.xml');
    let xml = await f.async('string');
    if (!xml.includes('\uE050')) return bytes;
    const text = p => (p.match(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join('');
    const run = (t, rPr = '') => `<w:r>${rPr}<w:t xml:space="preserve">${t}</w:t></w:r>`;
    const fld = type => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
    const instr = s => `<w:r><w:instrText xml:space="preserve">${s}</w:instrText></w:r>`;
    xml = xml.replace(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?\uE050[\s\S]*?\uE051\uE0FF[\s\S]*?<\/w:p>/g, region => {
      const paras = region.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
      const title = text(paras[0]).replace(/[\uE050\uE051][^\uE0FF]*?\uE0FF|\uE050|\uE0FF/g, '').trim() || 'Contents';
      const entries = paras.slice(1).flatMap(p => {
        const m = text(p).match(/\uE052(\d)\|(\d*)\uE0FF([\s\S]*)/);
        const anchor = (p.match(/<w:hyperlink\b[^>]*w:anchor="([^"]+)"/) || [])[1];
        return m && anchor ? [{ level: +m[1], page: m[2], title: m[3], anchor }] : [];
      });
      const start = fld('begin') + instr(' TOC \\o "1-3" \\h \\z \\u ') + fld('separate');
      const pPr = level => `<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${tabPos}"/></w:tabs><w:spacing w:after="80"/><w:ind w:left="${(level - 1) * 240}"/></w:pPr>`;
      const body = entries.length ? entries.map((e, i) => `<w:p>${pPr(e.level)}${i === 0 ? start : ''}<w:hyperlink w:anchor="${e.anchor}" w:history="1">${run(e.title)}<w:r><w:tab/></w:r>`
        + fld('begin') + instr(` PAGEREF ${e.anchor} \\h `) + fld('separate') + run(e.page || '1') + fld('end') + '</w:hyperlink></w:p>').join('')
        : `<w:p>${pPr(1)}${start}${run('No table of contents entries found.')}</w:p>`;
      return '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'
        + `<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr>${run(title, '<w:rPr><w:b/><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr>')}</w:p>`
        + body + `<w:p>${fld('end')}</w:p></w:sdtContent></w:sdt>`;
    });
    zip.file('word/document.xml', xml);
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  };

  /** Word's contents (links to _Toc bookmarks) opened as HTML becomes an Atlas table of contents. */
  WriteToc.importHtml = function (html) {
    const d = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
    const b = d.body;
    // Word bookmarks its headings with empty <a id="_Toc…">: move the id to the heading.
    b.querySelectorAll('h1,h2,h3,h4,h5,h6').forEach(h => {
      h.querySelectorAll('a[id]:empty').forEach(a => { if (!h.id) h.id = a.id; a.remove(); });
    });
    // An entry is a paragraph that is only a link to a heading.
    const headingIds = new Set([...b.querySelectorAll('h1[id],h2[id],h3[id],h4[id],h5[id],h6[id]')].map(h => h.id));
    const isEntry = p => {
      if (!p || p.tagName !== 'P' || p.parentNode !== b) return false;
      const links = p.querySelectorAll('a[href^="#"]');
      return links.length === 1 && headingIds.has(decodeURIComponent(links[0].getAttribute('href').slice(1))) && p.textContent.trim() === links[0].textContent.trim();
    };
    const first = [...b.children].find(isEntry);
    if (!first) return html;
    const group = [];
    for (let p = first; isEntry(p); p = p.nextElementSibling) group.push(p);
    const prev = first.previousElementSibling;
    const title = prev && TOC_TITLE.test(prev.textContent.trim()) ? prev.textContent.trim() : 'Contents';
    if (prev && TOC_TITLE.test(prev.textContent.trim())) prev.remove();
    const holder = d.createElement('div');
    holder.innerHTML = tocHtml(b, title);
    first.before(holder.firstChild);
    group.forEach(p => p.remove());
    return b.innerHTML;
  };

  window.WriteToc = WriteToc;
})();
