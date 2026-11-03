'use strict';
// Isadek — the text in a file, for the Assistant to read alongside the open document.
//   FileText.extract(name, bytes) → { kind, text }   (text: plain, with simple structure kept)
// Word (.docx), Excel (.xlsx, .csv, .tsv), PowerPoint (.pptx), OpenDocument text (.odt), PDF,
// Isadek notebooks (.anote), Markdown, HTML and plain text. Office files are read with JSZip
// (loaded on demand), PDFs with pdf.js; nothing runs from the file.
//
// library files: node_modules/jszip/dist/jszip.min.js node_modules/pdfjs-dist/build/pdf.min.mjs node_modules/pdfjs-dist/build/pdf.worker.min.mjs
(function () {
  const LIB = '../../node_modules/';
  const MAX = 400000; // characters kept from one file
  const loadScript = (src, test) => new Promise((resolve, reject) => {
    if (test()) return resolve();
    const s = document.createElement('script');
    s.src = src;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('A reader couldn’t be loaded'));
    document.head.appendChild(s);
  });
  const zip = async bytes => { await loadScript(LIB + 'jszip/dist/jszip.min.js', () => !!window.JSZip); return window.JSZip.loadAsync(bytes); };
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const decode = b => { const t = new TextDecoder('utf-8', { fatal: false }).decode(b); return t.charCodeAt(0) === 0xfeff ? t.slice(1) : t; };
  const byNumber = (a, b) => (+(a.match(/(\d+)\.xml$/) || [0, 0])[1]) - (+(b.match(/(\d+)\.xml$/) || [0, 0])[1]);
  /** Text of the elements named `tag` in each paragraph-like element `para`. */
  function paras(doc, para, tag) {
    const out = [];
    for (const p of doc.getElementsByTagName(para)) {
      let t = '';
      for (const n of p.getElementsByTagName('*')) {
        if (n.localName === tag) t += n.textContent;
        else if (n.localName === 'tab') t += '\t';
        else if (n.localName === 'br') t += '\n';
      }
      out.push(t);
    }
    return out;
  }

  async function docx(bytes) {
    const z = await zip(bytes);
    const f = z.file('word/document.xml');
    if (!f) throw new Error('This isn’t a Word document');
    const d = xml(await f.async('string'));
    // Headings keep a # so the structure shows.
    const out = [];
    for (const p of d.getElementsByTagName('w:p')) {
      const style = (p.getElementsByTagName('w:pStyle')[0] || { getAttribute: () => '' }).getAttribute('w:val') || '';
      const lvl = /^Heading(\d)$/i.exec(style) || /^Title$/i.test(style) && [0, '1'];
      let t = '';
      for (const n of p.getElementsByTagName('*')) { if (n.localName === 't') t += n.textContent; else if (n.localName === 'tab') t += '\t'; }
      out.push(lvl && t ? '#'.repeat(+lvl[1]) + ' ' + t : t);
    }
    return out.join('\n');
  }
  async function xlsx(bytes) {
    const z = await zip(bytes);
    const wb = z.file('xl/workbook.xml');
    if (!wb) throw new Error('This isn’t an Excel workbook');
    const ss = z.file('xl/sharedStrings.xml') ? [...xml(await z.file('xl/sharedStrings.xml').async('string')).getElementsByTagName('si')].map(si => [...si.getElementsByTagName('t')].map(t => t.textContent).join('')) : [];
    const names = [...xml(await wb.async('string')).getElementsByTagName('sheet')].map(s => s.getAttribute('name'));
    const files = Object.keys(z.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort(byNumber);
    const out = [];
    for (let i = 0; i < files.length; i++) {
      const d = xml(await z.file(files[i]).async('string'));
      out.push(`## Sheet: ${names[i] || 'Sheet' + (i + 1)}`);
      for (const row of d.getElementsByTagName('row')) {
        const cells = [];
        for (const c of row.getElementsByTagName('c')) {
          const v = c.getElementsByTagName('v')[0], is = c.getElementsByTagName('is')[0], t = c.getAttribute('t');
          const val = t === 's' && v ? ss[+v.textContent] ?? '' : t === 'inlineStr' && is ? is.textContent : v ? v.textContent : '';
          cells.push(`${c.getAttribute('r')}=${val}`);
        }
        if (cells.length) out.push(cells.join('\t'));
        if (out.length > 20000) break;
      }
    }
    return out.join('\n');
  }
  async function pptx(bytes) {
    const z = await zip(bytes);
    const files = Object.keys(z.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort(byNumber);
    if (!files.length) throw new Error('This isn’t a PowerPoint presentation');
    const out = [];
    for (let i = 0; i < files.length; i++) {
      out.push(`## Slide ${i + 1}`);
      out.push(...paras(xml(await z.file(files[i]).async('string')), 'a:p', 't').filter(Boolean));
      const notes = z.file(files[i].replace('slides/slide', 'notesSlides/notesSlide'));
      if (notes) { const n = paras(xml(await notes.async('string')), 'a:p', 't').filter(t => t && !/^\d+$/.test(t)); if (n.length) out.push('Notes: ' + n.join(' ')); }
    }
    return out.join('\n');
  }
  async function odt(bytes) {
    const z = await zip(bytes);
    const f = z.file('content.xml');
    if (!f) throw new Error('This isn’t an OpenDocument file');
    const d = xml(await f.async('string'));
    return [...d.getElementsByTagName('*')].filter(n => n.localName === 'p' || n.localName === 'h').map(n => n.textContent).join('\n');
  }
  async function pdf(bytes) {
    const pdfjs = await import(LIB.replace('../../', new URL('../../', location.href).href) + 'pdfjs-dist/build/pdf.min.mjs');
    if (!pdfjs.GlobalWorkerOptions.workerSrc) pdfjs.GlobalWorkerOptions.workerSrc = new URL(LIB + 'pdfjs-dist/build/pdf.worker.min.mjs', location.href).href;
    const doc = await pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false, enableXfa: false }).promise;
    const out = [];
    let n = 0;
    for (let i = 1; i <= doc.numPages && n < MAX; i++) {
      const tc = await (await doc.getPage(i)).getTextContent();
      const t = tc.items.map(x => x.str + (x.hasEOL ? '\n' : '')).join('').replace(/[ \t]+\n/g, '\n');
      out.push(`## Page ${i}\n${t}`);
      n += t.length;
    }
    (doc.loadingTask || doc).destroy();
    return out.join('\n');
  }
  function anote(bytes) {
    const nb = JSON.parse(decode(bytes));
    const text = html => { const d = new DOMParser().parseFromString(String(html || ''), 'text/html'); d.querySelectorAll('p, li, h1, h2, h3, div, tr, br').forEach(n => n.after('\n')); return d.body.textContent.replace(/\n{3,}/g, '\n\n').trim(); };
    const out = [`# Notebook: ${nb.name || ''}`];
    for (const s of nb.sections || []) { out.push(`## Section: ${s.name || ''}`); for (const p of s.pages || []) out.push(`### ${p.title || 'Untitled page'}\n${text(p.html)}`); }
    return out.join('\n\n');
  }
  function html(bytes) {
    const d = new DOMParser().parseFromString(decode(bytes), 'text/html');
    d.querySelectorAll('script, style, noscript, template').forEach(n => n.remove());
    d.querySelectorAll('p, li, h1, h2, h3, h4, div, tr, br').forEach(n => n.after('\n'));
    return (d.body ? d.body.textContent : '').replace(/\n{3,}/g, '\n\n').trim();
  }

  const KINDS = [
    [/\.docx$/i, 'Word document', docx], [/\.(xlsx|xlsm)$/i, 'Excel workbook', xlsx], [/\.pptx$/i, 'PowerPoint presentation', pptx],
    [/\.(odt|odp|ods)$/i, 'OpenDocument file', odt], [/\.pdf$/i, 'PDF', pdf], [/\.anote$/i, 'Isadek notebook', anote],
    [/\.(html?|xhtml)$/i, 'web page', html], [/\.(csv|tsv)$/i, 'table', b => decode(b)], [/\.(md|markdown|txt|text|json|xml|log|rtf)$/i, 'text', b => decode(b)]
  ];
  const EXTENSIONS = ['docx', 'xlsx', 'xlsm', 'pptx', 'odt', 'odp', 'ods', 'pdf', 'anote', 'html', 'htm', 'csv', 'tsv', 'md', 'markdown', 'txt', 'json', 'xml'];
  async function extract(name, bytes) {
    const k = KINDS.find(([re]) => re.test(name));
    if (!k) throw new Error('This kind of file can’t be read here');
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (b.length > 100e6) throw new Error('This file is too large');
    let text = String(await k[2](b) || '').replace(/\r\n?/g, '\n').replace(/\n{4,}/g, '\n\n\n').trim();
    const cut = text.length > MAX;
    if (cut) text = text.slice(0, MAX) + '\n[… the rest of the file is left out]';
    return { kind: k[1], text, cut };
  }
  window.FileText = { extract, EXTENSIONS };
})();
