'use strict';
// Isadek Write — reference formats for the Sources manager: BibTeX and RIS (what Zotero, Mendeley,
// EndNote and Google Scholar export) in and out, and Crossref's record for a DOI. Records come out
// as plain source fields for write-notes.js (which checks them again).
(function (root) {
  const TYPE_FROM_BIB = { article: 'article', book: 'book', inbook: 'book', incollection: 'book', booklet: 'book', online: 'website', misc: 'website', webpage: 'website', techreport: 'report', report: 'report', phdthesis: 'report', mastersthesis: 'report', inproceedings: 'article', conference: 'article', proceedings: 'book', manual: 'report', unpublished: 'report' };
  const TYPE_FROM_RIS = { JOUR: 'article', JFULL: 'article', MGZN: 'article', NEWS: 'article', CONF: 'article', CPAPER: 'article', BOOK: 'book', CHAP: 'book', EBOOK: 'book', ECHAP: 'book', EDBOOK: 'book', ELEC: 'website', WEB: 'website', BLOG: 'website', RPRT: 'report', THES: 'report', GOVDOC: 'report', STAND: 'report' };
  const TYPE_FROM_CSL = { 'journal-article': 'article', 'proceedings-article': 'article', 'article-journal': 'article', 'posted-content': 'article', book: 'book', 'book-chapter': 'book', monograph: 'book', 'edited-book': 'book', 'reference-book': 'book', report: 'report', dissertation: 'report', standard: 'report', dataset: 'report', webpage: 'website' };

  // LaTeX accents and escapes that turn up in exported BibTeX.
  const ACCENTS = { '`': '\u0300', "'": '\u0301', '^': '\u0302', '~': '\u0303', '=': '\u0304', '"': '\u0308', '.': '\u0307', u: '\u0306', v: '\u030C', H: '\u030B', c: '\u0327', k: '\u0328' };
  const LETTERS = { ss: 'ß', o: 'ø', O: 'Ø', ae: 'æ', AE: 'Æ', oe: 'œ', OE: 'Œ', aa: 'å', AA: 'Å', l: 'ł', L: 'Ł', i: 'ı' };
  function latex(s) {
    return String(s)
      .replace(/\\([`'^~="."]|[uvHck](?=[\s{]))\s*\{?\\?([A-Za-z])\}?/g, (m, a, ch) => (ACCENTS[a] ? (ch + ACCENTS[a]).normalize('NFC') : m))
      .replace(/\{?\\(ss|o|O|ae|AE|oe|OE|aa|AA|l|L|i)\b\}?/g, (m, k) => LETTERS[k])
      .replace(/\\([&%$#_{}])/g, '$1')
      .replace(/--/g, '–')
      .replace(/\\textendash\b\s*/g, '–')
      .replace(/[{}]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /** BibTeX (or BibLaTeX) entries as source fields. */
  function parseBibTeX(text) {
    const out = [];
    const s = String(text || '');
    let i = 0;
    while ((i = s.indexOf('@', i)) >= 0) {
      const m = /^@\s*(\w+)\s*[{(]/.exec(s.slice(i, i + 40));
      if (!m) { i++; continue; }
      const kind = m[1].toLowerCase();
      i += m[0].length;
      if (kind === 'comment' || kind === 'preamble' || kind === 'string') { i = skipGroup(s, i - 1); continue; }
      const comma = s.indexOf(',', i);
      if (comma < 0) break;
      const fields = {};
      i = comma + 1;
      // name = {value} | "value" | number, … until the entry's closing brace
      for (;;) {
        while (i < s.length && /[\s,]/.test(s[i])) i++;
        if (i >= s.length || s[i] === '}' || s[i] === ')') { i++; break; }
        const f = /^([\w:-]+)\s*=\s*/.exec(s.slice(i, i + 80));
        if (!f) { i = skipGroup(s, i); break; }
        i += f[0].length;
        let v = '';
        // Values may be joined with #.
        for (;;) {
          if (s[i] === '{') { const e = skipGroup(s, i); v += s.slice(i + 1, e - 1); i = e; }
          else if (s[i] === '"') { let e = i + 1, depth = 0; while (e < s.length && !(s[e] === '"' && depth === 0)) { if (s[e] === '{') depth++; else if (s[e] === '}') depth--; e++; } v += s.slice(i + 1, e); i = e + 1; }
          else { const w = /^[^,}\s#]+/.exec(s.slice(i)); if (w) { v += w[0]; i += w[0].length; } }
          while (/\s/.test(s[i] || '')) i++;
          if (s[i] === '#') { i++; while (/\s/.test(s[i] || '')) i++; continue; }
          break;
        }
        fields[f[1].toLowerCase()] = latex(v);
      }
      const authors = (fields.author || fields.editor || '').split(/\s+and\s+/i).map(a => a.trim()).filter(Boolean).join('; ');
      const date = fields.year || (fields.date || '').slice(0, 4);
      out.push({
        type: TYPE_FROM_BIB[kind] || 'book',
        authors, title: fields.title || fields.booktitle || '', year: date,
        publisher: fields.publisher || fields.institution || fields.school || fields.organization || fields.howpublished || '',
        journal: fields.journal || fields.journaltitle || (kind === 'inproceedings' || kind === 'conference' ? fields.booktitle || '' : ''),
        volume: fields.volume || '', issue: fields.number || fields.issue || '', pages: (fields.pages || '').replace(/-+/g, '–'),
        url: fields.url || (fields.doi ? 'https://doi.org/' + fields.doi.replace(/^https?:\/\/(dx\.)?doi\.org\//i, '') : ''),
        doi: (fields.doi || '').replace(/^https?:\/\/(dx\.)?doi\.org\//i, ''),
        accessed: fields.urldate || ''
      });
    }
    return out;
  }
  /** The index just past the bracketed group that starts at i. */
  function skipGroup(s, i) {
    const open = s[i], close = open === '(' ? ')' : '}';
    let depth = 0;
    for (; i < s.length; i++) {
      if (s[i] === open) depth++;
      else if (s[i] === close && --depth === 0) return i + 1;
    }
    return s.length;
  }

  /** RIS records as source fields. */
  function parseRIS(text) {
    const out = [];
    let cur = null;
    for (const line of String(text || '').split(/\r?\n/)) {
      const m = /^([A-Z][A-Z0-9])  -\s?(.*)$/.exec(line);
      if (!m) continue;
      const [, tag, raw] = m, v = raw.trim();
      if (tag === 'TY') { cur = { type: TYPE_FROM_RIS[v] || 'book', authors: [], title: '', year: '', publisher: '', journal: '', volume: '', issue: '', pages: '', url: '', doi: '', sp: '', ep: '' }; continue; }
      if (!cur) continue;
      if (tag === 'ER') { finishRis(cur); out.push(cur); cur = null; continue; }
      if (['AU', 'A1', 'A2', 'ED'].includes(tag)) { if (tag === 'AU' || tag === 'A1' || !cur.authors.length) cur.authors.push(v); }
      else if ((tag === 'TI' || tag === 'T1') && !cur.title) cur.title = v;
      else if ((tag === 'PY' || tag === 'Y1' || tag === 'DA') && !cur.year) cur.year = (v.match(/\d{4}/) || [''])[0];
      else if ((tag === 'JO' || tag === 'JF' || tag === 'T2' || tag === 'JA') && !cur.journal) cur.journal = v;
      else if (tag === 'PB') cur.publisher = v;
      else if (tag === 'VL') cur.volume = v;
      else if (tag === 'IS') cur.issue = v;
      else if (tag === 'SP') cur.sp = v;
      else if (tag === 'EP') cur.ep = v;
      else if (tag === 'UR' && !cur.url) cur.url = v;
      else if (tag === 'DO') cur.doi = v.replace(/^https?:\/\/(dx\.)?doi\.org\//i, '');
    }
    if (cur) { finishRis(cur); out.push(cur); }
    return out;
  }
  function finishRis(r) {
    r.authors = r.authors.join('; ');
    r.pages = r.sp && r.ep ? `${r.sp}–${r.ep}` : r.sp;
    if (!r.url && r.doi) r.url = 'https://doi.org/' + r.doi;
    if (r.type !== 'article' && r.journal && !r.publisher && r.type === 'website') r.publisher = r.journal;
    delete r.sp; delete r.ep;
  }

  /** Crossref's /works/{doi} message as source fields. */
  function fromCrossref(m) {
    if (!m || typeof m !== 'object') return null;
    const first = v => (Array.isArray(v) ? v[0] || '' : v || '');
    const date = (m.issued || m.published || m['published-print'] || m['published-online'] || {})['date-parts'];
    const year = Array.isArray(date) && Array.isArray(date[0]) && date[0][0] ? String(date[0][0]) : '';
    const people = (m.author || m.editor || []).map(a => (a.family ? `${a.family}, ${a.given || ''}`.replace(/, $/, '') : a.name || '')).filter(Boolean);
    const type = TYPE_FROM_CSL[m.type] || 'article';
    return {
      type, authors: people.join('; '), title: first(m.title).replace(/<[^>]+>/g, ''), year,
      publisher: m.publisher || '', journal: type === 'article' ? first(m['container-title']) : '',
      volume: m.volume || '', issue: m.issue || '', pages: (m.page || '').replace(/-+/g, '–'),
      url: m.DOI ? 'https://doi.org/' + m.DOI : m.URL || '', doi: m.DOI || ''
    };
  }
  /** A DOI in text (a bare DOI, doi:…, or a doi.org link), or ''. */
  const doiIn = text => ((String(text || '').match(/\b(10\.\d{4,9}\/[^\s"<>]+)/) || [''])[0]).replace(/[.,;)\]]+$/, '');

  /** A citation key like smith2020quick. */
  function citeKey(src, taken = new Set()) {
    const last = (String(src.authors || '').split(';')[0].split(',')[0].trim().split(/\s+/).pop() || 'anon').normalize('NFD').replace(/[^\w]/g, '').toLowerCase() || 'anon';
    const word = (String(src.title || '').toLowerCase().match(/[a-z]{4,}/) || [''])[0];
    let key = last + (src.year || '') + word, n = 0;
    while (taken.has(key)) key = last + (src.year || '') + word + String.fromCharCode(97 + n++);
    taken.add(key);
    return key;
  }
  const bibEsc = v => String(v).replace(/([&%$#_])/g, '\\$1').replace(/–/g, '--');
  /** Sources as BibTeX. */
  function toBibTeX(list) {
    const taken = new Set();
    const KIND = { article: 'article', book: 'book', website: 'online', report: 'techreport' };
    return list.map(s => {
      const f = [];
      const add = (k, v) => { if (v) f.push(`  ${k} = {${bibEsc(v)}}`); };
      add('author', String(s.authors || '').split(';').map(a => a.trim()).filter(Boolean).join(' and '));
      add('title', s.title);
      add('year', s.year);
      if (s.type === 'article') add('journal', s.journal);
      add(s.type === 'report' ? 'institution' : 'publisher', s.publisher);
      add('volume', s.volume); add('number', s.issue); add('pages', s.pages);
      add('doi', s.doi);
      if (s.url && !(s.doi && /doi\.org/.test(s.url))) add('url', s.url);
      return `@${KIND[s.type] || 'misc'}{${citeKey(s, taken)},\n${f.join(',\n')}\n}`;
    }).join('\n\n') + '\n';
  }

  const api = { parseBibTeX, parseRIS, fromCrossref, doiIn, citeKey, toBibTeX, latex };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.WriteRefs = api;
})(typeof window !== 'undefined' ? window : globalThis);
