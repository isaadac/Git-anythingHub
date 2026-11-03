'use strict';
// Isadek Write — footnotes, endnotes, citations and bibliography.
//
// A note is <sup class="note" data-kind="foot|end" title="text">n</sup>; its number is
// recalculated from its position. A citation is <cite class="citation" data-src="id"
// data-pages="12">(Smith, 2020)</cite>; the sources live in the document's metadata.
// A bibliography is a <section class="bibliography"> rebuilt from the cited sources.
(function () {
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const roman = n => { const t = [[10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; for (const [v, r] of t) while (n >= v) { s += r; n -= v; } return s; };
  const ID = /^s[a-z0-9]{4,24}$/;
  const STYLES = { apa: 'APA 7th', mla: 'MLA 9th', chicago: 'Chicago (author-date)', harvard: 'Harvard', ieee: 'IEEE' };
  const TYPES = { book: 'Book', article: 'Journal article', website: 'Web page', report: 'Report' };

  /** Accepts only well-formed source records (they come from files). */
  function cleanSource(s) {
    if (!s || typeof s !== 'object' || !ID.test(s.id || '')) return null;
    const str = (v, n = 300) => (typeof v === 'string' ? v.slice(0, n) : '');
    return {
      id: s.id, type: TYPES[s.type] ? s.type : 'book', authors: str(s.authors), title: str(s.title, 500), year: str(s.year, 20),
      publisher: str(s.publisher), journal: str(s.journal), volume: str(s.volume, 40), issue: str(s.issue, 40), pages: str(s.pages, 40),
      url: /^https?:\/\//i.test(s.url || '') ? str(s.url, 1000) : '', accessed: str(s.accessed, 40),
      doi: /^10\.\d{4,9}\/\S+$/.test(s.doi || '') ? str(s.doi, 200) : ''
    };
  }
  /** "Smith, John; Jones, Ann" → [{last, first}] (also accepts "John Smith"). */
  function people(authors) {
    return String(authors || '').split(/;|\band\b|&/).map(a => a.trim()).filter(Boolean).map(a => {
      if (a.includes(',')) { const [last, first] = a.split(',').map(x => x.trim()); return { last, first: first || '' }; }
      const parts = a.split(/\s+/);
      return { last: parts.pop(), first: parts.join(' ') };
    });
  }
  const initials = first => first.split(/[\s-]+/).filter(Boolean).map(w => w[0].toUpperCase() + '.').join(' ');

  /** The citation in the text; num is the source's number (IEEE numbers sources by first citation). */
  function inText(src, pages, style, num) {
    const ps = people(src.authors);
    const year = src.year || 'n.d.';
    const names = sep => !ps.length ? `“${src.title}”` : ps.length === 1 ? ps[0].last : ps.length === 2 ? `${ps[0].last} ${sep} ${ps[1].last}` : `${ps[0].last} et al.`;
    if (style === 'ieee') return `[${num || '?'}${pages ? `, ${/[-–]/.test(pages) ? 'pp.' : 'p.'} ${pages}` : ''}]`;
    if (style === 'harvard') return `(${names('and')} ${year}${pages ? `, ${/[-–]/.test(pages) ? 'pp.' : 'p.'} ${pages}` : ''})`;
    if (style === 'mla') return `(${names('and')}${pages ? ' ' + pages : ''})`;
    if (style === 'chicago') return `(${names('and')} ${year}${pages ? ', ' + pages : ''})`;
    return `(${names('&')}, ${year}${pages ? `, ${/[-–]/.test(pages) ? 'pp.' : 'p.'} ${pages}` : ''})`;
  }
  function entry(src, style, num) {
    const ps = people(src.authors);
    const t = esc(src.title), year = esc(src.year || 'n.d.');
    const url = src.url ? ` <a href="${esc(src.url)}">${esc(src.url)}</a>` : '';
    if (style === 'ieee') {
      const nm = p => `${esc(initials(p.first))}${p.first ? ' ' : ''}${esc(p.last)}`;
      const a = !ps.length ? '' : (ps.length > 6 ? nm(ps[0]) + ' <i>et al.</i>' : ps.length === 1 ? nm(ps[0]) : ps.slice(0, -1).map(nm).join(', ') + (ps.length > 2 ? ',' : '') + ' and ' + nm(ps[ps.length - 1])) + ', ';
      const n = `<span class="bib-num">[${num}]</span> `;
      if (src.type === 'article') return `${n}${a}“${t},” <i>${esc(src.journal)}</i>${src.volume ? `, vol. ${esc(src.volume)}` : ''}${src.issue ? `, no. ${esc(src.issue)}` : ''}${src.pages ? `, pp. ${esc(src.pages)}` : ''}, ${year}${src.doi ? `, doi: ${esc(src.doi)}` : ''}.`;
      if (src.type === 'website') return `${n}${a}“${t}.” ${src.publisher ? esc(src.publisher) + '. ' : ''}${src.url ? `[Online]. Available:${url}` : ''}`;
      if (src.type === 'report') return `${n}${a}“${t},” ${src.publisher ? esc(src.publisher) + ', ' : ''}${year}.`;
      return `${n}${a}<i>${t}</i>. ${src.publisher ? esc(src.publisher) + ', ' : ''}${year}.`;
    }
    if (style === 'harvard') {
      const nm = p => `${esc(p.last)}${p.first ? ', ' + esc(initials(p.first)) : ''}`;
      const a = !ps.length ? '' : (ps.length === 1 ? nm(ps[0]) : ps.slice(0, -1).map(nm).join(', ') + ' and ' + nm(ps[ps.length - 1])) + ' ';
      if (src.type === 'article') return `${a}(${year}) ‘${t}’, <i>${esc(src.journal)}</i>${src.volume ? `, ${esc(src.volume)}` : ''}${src.issue ? `(${esc(src.issue)})` : ''}${src.pages ? `, pp. ${esc(src.pages)}` : ''}.${src.doi ? ` doi:${esc(src.doi)}.` : ''}`;
      if (src.type === 'website') return `${a}(${year}) <i>${t}</i>.${src.url ? ` Available at:${url}` : ''}${src.accessed ? ` (Accessed: ${esc(src.accessed)})` : ''}.`;
      return `${a}(${year}) <i>${t}</i>. ${src.publisher ? esc(src.publisher) + '.' : ''}`;
    }
    if (style === 'mla') {
      const a = !ps.length ? '' : ps.length === 1 ? `${esc(ps[0].last)}, ${esc(ps[0].first)}. ` : ps.length === 2 ? `${esc(ps[0].last)}, ${esc(ps[0].first)}, and ${esc(ps[1].first)} ${esc(ps[1].last)}. ` : `${esc(ps[0].last)}, ${esc(ps[0].first)}, et al. `;
      if (src.type === 'article') return `${a}“${t}.” <i>${esc(src.journal)}</i>${src.volume ? `, vol. ${esc(src.volume)}` : ''}${src.issue ? `, no. ${esc(src.issue)}` : ''}, ${year}${src.pages ? `, pp. ${esc(src.pages)}` : ''}.${url}`;
      if (src.type === 'website') return `${a}“${t}.” <i>${esc(src.publisher)}</i>, ${year}${src.url ? `, ${esc(src.url)}` : ''}.`;
      return `${a}<i>${t}</i>. ${src.publisher ? esc(src.publisher) + ', ' : ''}${year}.`;
    }
    if (style === 'chicago') {
      const a = !ps.length ? '' : ps.map((p, i) => (i === 0 ? `${esc(p.last)}, ${esc(p.first)}` : `${esc(p.first)} ${esc(p.last)}`)).join(ps.length === 2 ? ' and ' : ', ').replace(/, ([^,]+)$/, ps.length > 2 ? ', and $1' : ', $1') + '. ';
      if (src.type === 'article') return `${a}${year}. “${t}.” <i>${esc(src.journal)}</i> ${esc(src.volume)}${src.issue ? ` (${esc(src.issue)})` : ''}${src.pages ? `: ${esc(src.pages)}` : ''}.${url}`;
      if (src.type === 'website') return `${a}${year}. “${t}.” ${esc(src.publisher)}.${url}`;
      return `${a}${year}. <i>${t}</i>. ${esc(src.publisher)}.`;
    }
    // APA 7
    const names = ps.map(p => `${esc(p.last)}, ${esc(initials(p.first))}`);
    const a = !names.length ? '' : names.length === 1 ? names[0] + ' ' : names.slice(0, -1).join(', ') + ', &amp; ' + names[names.length - 1] + ' ';
    if (src.type === 'article') return `${a}(${year}). ${t}. <i>${esc(src.journal)}</i>${src.volume ? `, <i>${esc(src.volume)}</i>` : ''}${src.issue ? `(${esc(src.issue)})` : ''}${src.pages ? `, ${esc(src.pages)}` : ''}.${url}`;
    if (src.type === 'website') return `${a}(${year}). <i>${t}</i>. ${src.publisher ? esc(src.publisher) + '.' : ''}${url}`;
    return `${a}(${year}). <i>${t}</i>. ${esc(src.publisher)}.${url}`;
  }

  function WriteNotes(o) {
    const { doc } = o;
    let sources = [];
    const style = () => (STYLES[o.setup().citeStyle] ? o.setup().citeStyle : 'apa');

    // ---------------- Notes ----------------
    const notes = kind => [...doc.querySelectorAll(`sup.note[data-kind="${kind}"]`)];
    function renumber() {
      // Only touch the DOM when something changed, so typing elsewhere isn't disturbed.
      notes('foot').forEach((n, i) => { const t = String(i + 1); if (n.textContent !== t) n.textContent = t; if (n.contentEditable !== 'false') n.contentEditable = 'false'; });
      notes('end').forEach((n, i) => { const t = roman(i + 1); if (n.textContent !== t) n.textContent = t; if (n.contentEditable !== 'false') n.contentEditable = 'false'; });
      renderEndnotes();
    }
    /** Endnotes are listed in a generated block at the end of the document (never saved). */
    function renderEndnotes() {
      let box = doc.querySelector('.endnotes-gen');
      const list = notes('end');
      if (!list.length) { if (box) box.remove(); return; }
      if (!box) { box = document.createElement('section'); box.className = 'endnotes-gen'; box.contentEditable = 'false'; }
      if (box !== doc.lastElementChild) doc.appendChild(box);
      const html = '<h3>Endnotes</h3>' + list.map((n, i) => `<p data-i="${i}"><sup>${roman(i + 1)}</sup> ${esc(n.title)}</p>`).join('');
      if (box.innerHTML !== html) box.innerHTML = html;
    }
    async function insertNote(kind) {
      const r = await UI.form(kind === 'foot' ? 'Insert Footnote' : 'Insert Endnote', [{ name: 't', label: 'Note text', type: 'textarea' }], 'Insert');
      if (!r || !r.t.trim()) return;
      o.restoreSelection();
      Platform.insertInline(`<sup class="note" data-kind="${kind}" title="${esc(r.t.trim())}" contenteditable="false">0</sup>`);
      renumber();
      o.changed();
    }
    async function editNote(el) {
      const r = await UI.form(el.dataset.kind === 'foot' ? 'Footnote' : 'Endnote', [{ name: 't', label: 'Note text (empty to delete the note)', type: 'textarea', value: el.title }], 'Save');
      if (!r) return;
      if (!r.t.trim()) el.remove(); else el.title = r.t.trim();
      renumber();
      o.changed();
    }
    doc.addEventListener('click', e => {
      const n = e.target.closest('sup.note');
      if (n) { e.preventDefault(); editNote(n); return; }
      const g = e.target.closest('.endnotes-gen p[data-i]');
      if (g) { const ref = notes('end')[+g.dataset.i]; if (ref) editNote(ref); return; }
      const c = e.target.closest('cite.citation');
      if (c) { e.preventDefault(); editCitation(c); }
    });
    /** Footnote texts with their reference element, for the page layout. */
    const footnotes = () => notes('foot').map((el, i) => ({ el, n: i + 1, text: el.title }));

    // ---------------- Citations ----------------
    function loadSources(list) { sources = (Array.isArray(list) ? list : []).map(cleanSource).filter(Boolean); }
    const sourceById = id => sources.find(s => s.id === id);
    function library() { try { return (JSON.parse(localStorage.getItem('atlas-sources') || '[]') || []).map(cleanSource).filter(Boolean); } catch { return []; } }
    function remember(src) {
      const lib = library().filter(s => s.id !== src.id);
      lib.unshift(src);
      try { localStorage.setItem('atlas-sources', JSON.stringify(lib.slice(0, 200))); } catch { /* ignore */ }
    }
    async function editSource(src) {
      const s = src || { id: 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), type: 'book' };
      const r = await UI.form(src ? 'Edit Source' : 'New Source', [
        { name: 'type', label: 'Type', type: 'select', options: Object.entries(TYPES).map(([value, label]) => ({ value, label })), value: s.type },
        { name: 'authors', label: 'Authors (Last, First; Last, First)', value: s.authors || '' },
        { name: 'title', label: 'Title', value: s.title || '' },
        { name: 'year', label: 'Year', value: s.year || '' },
        { name: 'publisher', label: 'Publisher / website name', value: s.publisher || '' },
        { name: 'journal', label: 'Journal (articles)', value: s.journal || '' },
        { name: 'volume', label: 'Volume', value: s.volume || '' },
        { name: 'issue', label: 'Issue', value: s.issue || '' },
        { name: 'pages', label: 'Pages (articles)', value: s.pages || '' },
        { name: 'url', label: 'URL (https://…)', value: s.url || '' }
      ], 'Save');
      if (!r) return null;
      const next = cleanSource({ ...s, ...r });
      if (!next || !next.title.trim()) { UI.toast('A source needs at least a title'); return null; }
      const i = sources.findIndex(x => x.id === next.id);
      if (i >= 0) sources[i] = next; else sources.push(next);
      remember(next);
      refreshCitations();
      o.changed();
      drawPane();
      return next;
    }
    async function insertCitation() {
      // Sources already in this document first, then ones used in other documents.
      const pool = [...sources, ...library().filter(l => !sources.some(s => s.id === l.id))];
      const opts = [{ value: '__new', label: '+ Add a new source…' }, ...pool.map(s => ({ value: s.id, label: `${people(s.authors)[0]?.last || '—'} (${s.year || 'n.d.'}) — ${s.title.slice(0, 60)}` }))];
      const r = await UI.form('Insert Citation', [
        { name: 'src', label: 'Source', type: 'select', options: opts },
        { name: 'pages', label: 'Page(s), optional', value: '' }
      ], 'Insert');
      if (!r) return;
      let src = r.src === '__new' ? await editSource(null) : sourceById(r.src) || pool.find(s => s.id === r.src);
      if (!src) return;
      if (!sourceById(src.id)) sources.push(src);
      o.restoreSelection();
      const pages = r.pages.trim().slice(0, 40);
      Platform.insertInline(` <cite class="citation" data-src="${src.id}"${pages ? ` data-pages="${esc(pages)}"` : ''} contenteditable="false">${esc(inText(src, pages, style()))}</cite>`);
      refreshCitations();
      o.changed();
    }
    async function editCitation(c) {
      const src = sourceById(c.dataset.src);
      const r = await UI.form('Citation', [
        { name: 'pages', label: 'Page(s)', value: c.dataset.pages || '' },
        { name: 'action', label: 'Action', type: 'select', options: [{ value: 'save', label: 'Save' }, { value: 'source', label: 'Edit the source' }, { value: 'remove', label: 'Remove this citation' }] }
      ], 'OK');
      if (!r) return;
      if (r.action === 'remove') c.remove();
      else if (r.action === 'source' && src) await editSource(src);
      else { if (r.pages.trim()) c.dataset.pages = r.pages.trim().slice(0, 40); else delete c.dataset.pages; }
      refreshCitations();
      o.changed();
    }
    // ---------------- Sources manager ----------------
    // A pane listing this document's sources and the library of sources used before (kept on this
    // device), searchable; sources come in by hand, by DOI, or from BibTeX/RIS (Zotero, Mendeley,
    // EndNote, Google Scholar), and go out as BibTeX.
    let pane = null;
    const label = s => `${people(s.authors).map(p => p.last).slice(0, 2).join(' & ') || '—'}${people(s.authors).length > 2 ? ' et al.' : ''} (${s.year || 'n.d.'})`;
    const newId = () => 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    /** Adds sources (plain fields) to the document, skipping ones already there; returns how many. */
    function addSources(list) {
      let n = 0;
      for (const raw of list) {
        const dup = sources.find(x => (raw.doi && x.doi === raw.doi) || (x.title.toLowerCase() === String(raw.title || '').toLowerCase() && x.year === raw.year));
        if (dup) continue;
        const src = cleanSource({ ...raw, id: newId() });
        if (!src || !src.title.trim()) continue;
        sources.push(src);
        remember(src);
        n++;
      }
      if (n) { refreshCitations(); o.changed(); }
      return n;
    }
    /** Reads pasted or opened text: BibTeX, RIS, or DOIs. */
    async function importText(text) {
      const t = String(text || '');
      let list = /@\s*\w+\s*[{(]/.test(t) ? WriteRefs.parseBibTeX(t) : /^TY {2}- /m.test(t) ? WriteRefs.parseRIS(t) : [];
      if (!list.length) {
        const dois = [...new Set((t.match(/\b10\.\d{4,9}\/[^\s"<>]+/g) || []).map(d => WriteRefs.doiIn(d)))].slice(0, 25);
        if (!dois.length) { UI.toast('No BibTeX, RIS or DOI found there', 3500); return 0; }
        list = (await Promise.all(dois.map(lookupDoi))).filter(Boolean);
      }
      const n = addSources(list);
      UI.toast(n ? `${n} source${n === 1 ? '' : 's'} added` : 'Those sources are already in this document', 3000);
      drawPane();
      return n;
    }
    /** A source's details from Crossref, by DOI (null when it can't be found). */
    async function lookupDoi(doi) {
      try {
        const r = await fetch('https://api.crossref.org/works/' + encodeURIComponent(doi), { headers: { Accept: 'application/json' } });
        if (!r.ok) throw new Error(r.status);
        return WriteRefs.fromCrossref((await r.json()).message);
      } catch { UI.toast(`Couldn't look up ${doi}`, 3500); return null; }
    }
    async function manageSources() {
      if (pane) { pane.remove(); pane = null; return; }
      o.captureSelection && o.captureSelection();
      pane = document.createElement('aside');
      pane.className = 'activity-pane refs-pane';
      pane.setAttribute('aria-label', 'Sources');
      pane.innerHTML = `<header><b>Sources</b><span class="sp"></span><button class="x" title="Close" aria-label="Close"><span class="ico" data-icon="x"></span></button></header>
        <div class="refs-tools">
          <input type="search" class="refs-search" placeholder="Search authors, titles, years" aria-label="Search sources">
          <div class="refs-btns">
            <button class="btn new"><span class="ico" data-icon="plus"></span>New</button>
            <button class="btn doi"><span class="ico" data-icon="link"></span>DOI</button>
            <button class="btn paste"><span class="ico" data-icon="clipboard-paste"></span>Paste</button>
            <button class="btn import"><span class="ico" data-icon="file-up"></span>Import</button>
            <button class="btn export"><span class="ico" data-icon="file-down"></span>BibTeX</button>
          </div>
        </div>
        <div class="list" role="list"></div>`;
      document.body.appendChild(pane);
      const q = sel => pane.querySelector(sel);
      q('.x').addEventListener('click', () => { pane.remove(); pane = null; });
      q('.refs-search').addEventListener('input', drawPane);
      q('.new').addEventListener('click', async () => { if (await editSource(null)) drawPane(); });
      q('.doi').addEventListener('click', async () => {
        const doi = WriteRefs.doiIn(await UI.prompt('Add by DOI', 'DOI (10.xxxx/…) or a doi.org link'));
        if (!doi) return;
        const src = await lookupDoi(doi);
        if (src) { const n = addSources([src]); UI.toast(n ? `Added “${src.title.slice(0, 60)}”` : 'That source is already in this document', 3000); drawPane(); }
      });
      q('.paste').addEventListener('click', async () => {
        const r = await UI.form('Paste Sources', [{ name: 't', label: 'BibTeX, RIS, or DOIs (from Zotero, Mendeley, EndNote, Google Scholar…)', type: 'textarea' }], 'Add');
        if (r && r.t.trim()) await importText(r.t);
      });
      q('.import').addEventListener('click', async () => {
        const f = await Platform.open([{ name: 'References', extensions: ['bib', 'ris', 'txt'] }]);
        if (f) await importText(new TextDecoder().decode(f.data));
      });
      q('.export').addEventListener('click', async () => {
        if (!sources.length) return UI.toast('No sources in this document yet');
        await Platform.save({ data: new TextEncoder().encode(WriteRefs.toBibTeX(sources)), name: 'references.bib', filters: [{ name: 'BibTeX', extensions: ['bib'] }], noVersion: true });
      });
      drawPane();
      if (window.Icons) Icons.fill(pane);
      q('.refs-search').focus();
    }
    function drawPane() {
      if (!pane) return;
      const list = pane.querySelector('.list');
      const term = pane.querySelector('.refs-search').value.trim().toLowerCase();
      const hit = s => !term || [s.authors, s.title, s.year, s.journal, s.publisher, s.doi].join(' ').toLowerCase().includes(term);
      const counts = new Map();
      doc.querySelectorAll('cite.citation').forEach(c => counts.set(c.dataset.src, (counts.get(c.dataset.src) || 0) + 1));
      const lib = library().filter(l => !sources.some(s => s.id === l.id));
      list.innerHTML = '';
      const group = (title, items, inDoc) => {
        const h = document.createElement('div');
        h.className = 'refs-group';
        h.textContent = `${title} (${items.length})`;
        list.appendChild(h);
        if (!items.length) { const e = document.createElement('p'); e.className = 'refs-empty'; e.textContent = term ? 'Nothing matches.' : inDoc ? 'Add sources with New, DOI, Paste or Import.' : 'Sources you use are kept here for other documents.'; list.appendChild(e); return; }
        for (const s of items) {
          const row = document.createElement('div');
          row.className = 'ref';
          row.setAttribute('role', 'listitem');
          const n = counts.get(s.id) || 0;
          row.innerHTML = '<div class="who"></div><div class="title"></div><div class="meta"></div><div class="acts"></div>';
          row.querySelector('.who').textContent = label(s);
          row.querySelector('.title').textContent = s.title;
          row.querySelector('.meta').textContent = [TYPES[s.type], s.journal || s.publisher, inDoc ? (n ? `cited ${n}×` : 'not cited yet') : ''].filter(Boolean).join(' · ');
          const acts = row.querySelector('.acts');
          const act = (text, icon, fn) => { const b = document.createElement('button'); b.className = 'btn'; b.innerHTML = `<span class="ico" data-icon="${icon}"></span>`; b.append(text); b.addEventListener('click', fn); acts.appendChild(b); return b; };
          act('Cite', 'quote', () => cite(s));
          if (inDoc) {
            act('Edit', 'pencil', async () => { if (await editSource(s)) drawPane(); });
            if (!n) act('Remove', 'trash-2', () => { sources = sources.filter(x => x !== s); o.changed(); drawPane(); });
          }
          list.appendChild(row);
        }
      };
      group('In this document', sources.filter(hit), true);
      group('Library', lib.filter(hit), false);
      if (window.Icons) Icons.fill(list);
    }
    /** Cites a source where the cursor was. */
    function cite(src) {
      if (!sourceById(src.id)) sources.push(src);
      o.restoreSelection();
      Platform.insertInline(` <cite class="citation" data-src="${src.id}" contenteditable="false">${esc(inText(src, '', style()))}</cite>`);
      refreshCitations();
      o.changed();
      o.captureSelection && o.captureSelection();
      drawPane();
    }
    /** Updates citation text and every bibliography for the current sources and style. */
    function refreshCitations() {
      const nums = numbers();
      doc.querySelectorAll('cite.citation').forEach(c => {
        const src = sourceById(c.dataset.src);
        c.contentEditable = 'false';
        const t = src ? inText(src, c.dataset.pages || '', style(), nums.get(src.id)) : '(missing source)';
        if (c.textContent !== t) c.textContent = t;
      });
      doc.querySelectorAll('section.bibliography').forEach(fillBibliography);
    }
    /** Each cited source's number, in order of first citation (for IEEE). */
    function numbers() {
      const nums = new Map();
      doc.querySelectorAll('cite.citation').forEach(c => { if (sourceById(c.dataset.src) && !nums.has(c.dataset.src)) nums.set(c.dataset.src, nums.size + 1); });
      return nums;
    }
    function citedSources() {
      const ids = new Set([...doc.querySelectorAll('cite.citation')].map(c => c.dataset.src));
      if (style() === 'ieee') { const nums = numbers(); return sources.filter(s => ids.has(s.id)).sort((a, b) => nums.get(a.id) - nums.get(b.id)); }
      return sources.filter(s => ids.has(s.id)).sort((a, b) => (people(a.authors)[0]?.last || a.title).localeCompare(people(b.authors)[0]?.last || b.title));
    }
    function fillBibliography(sec) {
      const st = style();
      const heading = st === 'mla' ? 'Works Cited' : st === 'chicago' ? 'Bibliography' : 'References';
      const list = citedSources(), nums = numbers();
      sec.contentEditable = 'false';
      sec.classList.toggle('numbered', st === 'ieee');
      sec.innerHTML = `<h2>${heading}</h2>` + (list.length ? list.map(s => `<p class="bib-entry">${entry(s, st, nums.get(s.id))}</p>`).join('') : '<p><i>No citations yet. Use Insert Citation.</i></p>');
    }
    function insertBibliography() {
      o.restoreSelection();
      if (doc.querySelector('section.bibliography')) { refreshCitations(); UI.toast('The bibliography is up to date'); return; }
      document.execCommand('insertHTML', false, '<section class="bibliography" contenteditable="false"></section><p><br></p>');
      refreshCitations();
      o.changed();
    }

    // ---------------- Save / export ----------------
    /** Removes generated content and turns notes into Word's HTML note markup for .docx. */
    function prepareForDocx(root) {
      root.querySelectorAll('.endnotes-gen').forEach(n => n.remove());
      // Citations become Word CITATION fields (markers are turned into fields after conversion).
      root.querySelectorAll('cite.citation').forEach(c => {
        const pages = (c.dataset.pages || '').replace(/[^\w.,–-]/g, '');
        c.replaceWith(document.createTextNode(`\uE030${c.dataset.src}|${pages}\uE0FF${c.textContent}\uE031\uE0FF`));
      });
      const foot = [], end = [];
      root.querySelectorAll('sup.note').forEach(n => {
        const kind = n.dataset.kind === 'end' ? 'end' : 'foot';
        const list = kind === 'end' ? end : foot;
        list.push(n.title || '');
        const i = list.length;
        const k = kind === 'end' ? 'edn' : 'ftn';
        const cls = kind === 'end' ? 'MsoEndnoteReference' : 'MsoFootnoteReference';
        const a = document.createElement('a');
        a.setAttribute('style', `mso-${kind === 'end' ? 'endnote' : 'footnote'}-id:${k}${i}`);
        a.setAttribute('href', `#_${k}${i}`);
        a.setAttribute('name', `_${k}ref${i}`);
        a.innerHTML = `<span class="${cls}">[${i}]</span>`;
        n.replaceWith(a);
      });
      const listHtml = (arr, kind) => arr.length ? `<div style="mso-element:${kind}note-list">` + arr.map((t, j) => {
        const k = kind === 'end' ? 'edn' : 'ftn', i = j + 1;
        return `<div style="mso-element:${kind}note" id="${k}${i}"><p class="Mso${kind === 'end' ? 'End' : 'Foot'}noteText"><a style="mso-${kind}note-id:${k}${i}" href="#_${k}ref${i}" name="_${k}${i}"><span class="Mso${kind === 'end' ? 'End' : 'Foot'}noteReference">[${i}]</span></a> ${esc(t)}</p></div>`;
      }).join('') + '</div>' : '';
      root.insertAdjacentHTML('beforeend', listHtml(foot, 'foot') + listHtml(end, 'end'));
    }
    function clean(root) {
      root.querySelectorAll('.endnotes-gen').forEach(n => n.remove());
    }
    /** mammoth turns Word notes into links to a list at the end; turn them back into notes. */
    function importMammothNotes(html) {
      if (!/(foot|end)note-ref-/.test(html)) return html;
      const d = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
      // Read every reference first: footnotes and endnotes share one list at the end.
      for (const kind of ['foot', 'end']) {
        d.querySelectorAll(`a[id^="${kind}note-ref-"]`).forEach(a => {
          const target = d.getElementById(a.getAttribute('href').slice(1));
          const text = target ? target.textContent.replace(/\s*↑\s*$/, '').trim() : '';
          const sup = d.createElement('sup');
          sup.className = 'note';
          sup.dataset.kind = kind;
          sup.title = text;
          sup.textContent = '0';
          (a.parentElement && a.parentElement.tagName === 'SUP' ? a.parentElement : a).replaceWith(sup);
        });
      }
      d.querySelectorAll('li[id^="footnote-"], li[id^="endnote-"]').forEach(li => { const ol = li.closest('ol'); (ol || li).remove(); });
      return d.body.innerHTML;
    }
    function toMarkdownNotes(root) {
      // Footnotes as [^n] with definitions at the end.
      const defs = [];
      root.querySelectorAll('sup.note').forEach((n, i) => { defs.push(`[^${i + 1}]: ${n.title}`); n.replaceWith(document.createTextNode(`[^${i + 1}]`)); });
      root.querySelectorAll('.endnotes-gen').forEach(n => n.remove());
      return defs;
    }

    return {
      renumber, insertNote, footnotes, insertCitation, manageSources, insertBibliography, refreshCitations, importText, cite,
      loadSources, get sources() { return sources; },
      prepareForDocx, clean, importMammothNotes, toMarkdownNotes,
      STYLES
    };
  }

  window.WriteNotes = WriteNotes;
})();
