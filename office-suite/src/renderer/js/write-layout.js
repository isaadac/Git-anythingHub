'use strict';
// Isadek Write — page layout.
// The document is one editable column. After each change this pass measures the blocks
// and gives any block that would cross the bottom of a page a top margin that moves it to
// the next page; page sheets, headers and footers are drawn behind and around the text.
// The extra margins are marked with data-pg and stripped from everything that is saved.
(function () {
  const PX = 96; // CSS px per inch
  const CONTAINERS = new Set(['UL', 'OL', 'BLOCKQUOTE', 'DIV', 'SECTION']);

  function WriteLayout(o) {
    const { page, doc, bgs, hf } = o;
    let pages = 1, timer = null, printing = false, showHF = { header: false, footer: false };

    function geometry() {
      const s = o.setup();
      const noGap = printing || matchMedia('print').matches;
      return { W: s.width * PX, H: s.height * PX, M: s.margin * PX, first: !!s.differentFirst, gap: noGap ? 0 : 24, columns: s.columns };
    }
    const zoom = () => parseFloat(page.style.zoom) || 1;

    function clearPushes(root = doc) {
      root.querySelectorAll('[data-pg]').forEach(el => {
        el.style.marginTop = el.dataset.pgMt || '';
        if (!el.getAttribute('style')) el.removeAttribute('style');
        delete el.dataset.pg;
        delete el.dataset.pgMt;
      });
    }

    /** Blocks in reading order with their extent; big lists and quotes are split into their items. */
    function collectUnits(bodyH) {
      const units = [];
      const z = zoom();
      const origin = doc.getBoundingClientRect().top;
      const box = el => { const r = el.getBoundingClientRect(); return { top: (r.top - origin) / z, bottom: (r.bottom - origin) / z, h: r.height / z }; };
      const walk = (container, skipFirst) => {
        const kids = [...container.children].filter(el => el.offsetParent !== null || el.tagName === 'HR');
        let prev = null;
        kids.forEach((el, i) => {
          const b = box(el);
          const prevBottom = prev ? prev.bottom : (container === doc ? 0 : null);
          prev = b;
          if (skipFirst && i === 0) return;
          if (el.tagName === 'HR' && (el.classList.contains('page-break') || el.classList.contains('section-break'))) { units.push({ isBreak: true, top: b.top }); return; }
          const pushable = prevBottom !== null;
          if (CONTAINERS.has(el.tagName) && el.children.length > 1 && b.h > bodyH * 0.4 && !el.classList.contains('textbox') && !el.classList.contains('toc')) {
            // Keep the container's first item with the container, then treat each item separately.
            const first = box(el.children[0]);
            units.push({ el, top: b.top, bottom: first.bottom, prevBottom, pushable, isContainer: true });
            walk(el, true);
          } else {
            units.push({ el, top: b.top, bottom: b.bottom, prevBottom, pushable });
          }
        });
      };
      walk(doc, false);
      return units;
    }

    function run() {
      clearTimeout(timer);
      const g = geometry();
      // While printing the window is only as wide as the paper, so never fall back to the phone layout then.
      const paged = printing || matchMedia('print').matches || o.isPaged();
      document.body.classList.toggle('paged', paged);
      clearPushes();
      if (!paged) {
        bgs.innerHTML = '';
        page.style.height = '';
        Object.values(hf).forEach(el => { el.hidden = true; });
        pages = Math.max(1, Math.ceil(doc.scrollHeight / (g.H - 2 * g.M)));
        o.onLayout && o.onLayout(pages);
        return pages;
      }
      const bodyH = g.H - 2 * g.M;
      const pitch = g.H + g.gap;
      const pushes = [];
      let shift = 0, breakAt = null, lastBottom = 0;
      // Footnotes take space at the bottom of the page their reference is on.
      const notes = (o.footnotes ? o.footnotes() : []).map(f => ({ ...f, h: measureNote(f, g) }));
      const reserve = new Map();   // page -> height used by footnotes (incl. separator)
      const notesOn = new Map();   // page -> footnotes
      const NOTE_SEP = 14;
      // Multi-column text flows across the whole document; it isn't split per page.
      const units = g.columns > 1 ? [] : collectUnits(bodyH);
      for (const u of units) {
        const top = u.top + shift, bottom = u.isBreak ? top : u.bottom + shift;
        if (u.isBreak) { breakAt = top; continue; }
        const own = notes.filter(n => u.el.contains(n.el) && !(u.isContainer && !u.el.children[0].contains(n.el)));
        const ownH = own.reduce((a, n) => a + n.h, 0);
        const pageIdx = Math.floor((top + 0.5) / pitch);
        const res = p => (reserve.get(p) || 0);
        const needOn = p => res(p) + ownH + (ownH && !res(p) ? NOTE_SEP : 0);
        const pageBottom = pageIdx * pitch + bodyH - needOn(pageIdx);
        let target = null;
        if (breakAt !== null) { target = (Math.floor((breakAt + 0.5) / pitch) + 1) * pitch; breakAt = null; }
        else if (top > pageBottom - 0.5) target = (pageIdx + 1) * pitch;                                   // starts below the text area
        else if (bottom > pageBottom + 0.5 && bottom - top <= bodyH - ownH - (ownH ? NOTE_SEP : 0)) target = (pageIdx + 1) * pitch; // would cross the page edge
        let landed = pageIdx;
        if (target !== null && target > top + 0.5 && u.pushable) {
          const delta = target - top;
          // New margin = current gap above the block (margins collapse) + the distance to move.
          pushes.push([u.el, Math.round((u.top - u.prevBottom + delta) * 100) / 100]);
          shift += delta;
          landed = Math.floor((target + 0.5) / pitch);
        }
        if (own.length) {
          reserve.set(landed, needOn(landed));
          notesOn.set(landed, [...(notesOn.get(landed) || []), ...own]);
        }
        lastBottom = Math.max(lastBottom, u.bottom + shift);
      }
      if (g.columns > 1) lastBottom = doc.scrollHeight;
      pages = Math.max(1, Math.floor(Math.max(0, lastBottom - 1) / pitch) + 1);
      if (breakAt !== null) pages = Math.max(pages, Math.floor((breakAt + 0.5) / pitch) + 2);
      for (const [el, mt] of pushes) {
        el.dataset.pgMt = el.style.marginTop || '';
        el.dataset.pg = '1';
        el.style.marginTop = mt + 'px';
      }
      page.style.height = (pages * pitch - g.gap) + 'px';
      drawPages(g, pitch, notesOn, reserve);
      o.onLayout && o.onLayout(pages);
      return pages;
    }

    const fieldText = (kind, pageNo) => (kind === 'numpages' ? String(pages) : String(pageNo));
    function fillFields(root, pageNo) {
      root.querySelectorAll('.fld').forEach(f => { const t = fieldText(f.dataset.fld, pageNo); if (f.textContent !== t) f.textContent = t; });
    }
    const hasContent = el => el.textContent.trim() !== '' || !!el.querySelector('img');

    let noteMeasurer = null;
    function measureNote(f, g) {
      if (!noteMeasurer) {
        noteMeasurer = document.createElement('div');
        noteMeasurer.className = 'fn-note';
        Object.assign(noteMeasurer.style, { position: 'absolute', visibility: 'hidden', left: '-10000px', top: '0' });
        page.appendChild(noteMeasurer);
      }
      noteMeasurer.style.width = (g.W - 2 * g.M) + 'px';
      noteMeasurer.innerHTML = `<sup>${f.n}</sup> `;
      noteMeasurer.appendChild(document.createTextNode(f.text));
      return noteMeasurer.offsetHeight;
    }
    function drawPages(g, pitch, notesOn = new Map(), reserve = new Map()) {
      const inset = Math.max(19, g.M * 0.4);
      // Which element supplies the header/footer of each page, and which page hosts the editable one.
      const src = (kind, i) => (g.first && i === 0 ? hf['first' + kind] : hf[kind.toLowerCase()]);
      const host = { header: g.first ? 1 : 0, footer: g.first ? 1 : 0, firstHeader: 0, firstFooter: 0 };
      const frag = document.createDocumentFragment();
      for (let i = 0; i < pages; i++) {
        const bg = document.createElement('div');
        bg.className = 'page-bg';
        bg.style.top = (i * pitch) + 'px';
        for (const kind of ['Header', 'Footer']) {
          const el = src(kind, i);
          const editableHere = host[el === hf.header ? 'header' : el === hf.footer ? 'footer' : el === hf.firstHeader ? 'firstHeader' : 'firstFooter'] === i;
          if (editableHere || !hasContent(el)) continue;
          const c = document.createElement('div');
          c.className = 'hf-clone';
          c.innerHTML = el.innerHTML;
          fillFields(c, i + 1);
          if (kind === 'Header') c.style.top = inset + 'px'; else c.style.bottom = inset + 'px';
          bg.appendChild(c);
        }
        const fn = notesOn.get(i);
        if (fn && fn.length) {
          const area = document.createElement('div');
          area.className = 'fn-area';
          area.style.top = (g.M + (g.H - 2 * g.M) - reserve.get(i)) + 'px';
          fn.forEach(f => {
            const p = document.createElement('div');
            p.className = 'fn-note';
            p.innerHTML = `<sup>${f.n}</sup> `;
            p.appendChild(document.createTextNode(f.text));
            area.appendChild(p);
          });
          bg.appendChild(area);
        }
        frag.appendChild(bg);
      }
      bgs.replaceChildren(frag);
      // Place the editable headers/footers on their host pages.
      const place = (el, key, kind) => {
        const pageIdx = host[key];
        const used = key.startsWith('first') ? g.first : true;
        const visible = used && pageIdx < pages && (hasContent(el) || document.activeElement === el || showHF[kind]);
        el.hidden = !visible;
        if (!visible) return;
        fillFields(el, pageIdx + 1);
        if (kind === 'header') el.style.top = (pageIdx * pitch + inset) + 'px';
        else el.style.top = (pageIdx * pitch + g.H - inset - el.offsetHeight) + 'px';
      };
      place(hf.header, 'header', 'header');
      place(hf.footer, 'footer', 'footer');
      place(hf.firstHeader, 'firstHeader', 'header');
      place(hf.firstFooter, 'firstFooter', 'footer');
    }

    function schedule() { clearTimeout(timer); timer = setTimeout(run, 120); }

    // Pictures change height when they finish loading.
    doc.addEventListener('load', schedule, true);

    return {
      run,
      schedule,
      get pages() { return pages; },
      /** Which page (1-based) a document-relative y falls on. */
      pageAt(y) { const g = geometry(); return Math.min(pages, Math.floor(Math.max(0, y) / (g.H + g.gap)) + 1); },
      /** Lay out without gaps so each sheet is exactly one printed page; returns a function that restores. */
      forPrint() { printing = true; run(); return () => { printing = false; run(); }; },
      showHeaderFooter(kind) { showHF[kind] = true; run(); },
      hideEmptyHeaderFooter() { showHF = { header: false, footer: false }; run(); },
      clearPushes
    };
  }

  window.WriteLayout = WriteLayout;
})();
