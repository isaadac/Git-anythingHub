'use strict';
// Atlas Write — mail merge.
// A merge field is <span class="mergefield" data-field="First Name">«First Name»</span>.
// Recipients come from a CSV or spreadsheet whose first row holds the field names.
(function () {
  const FIELD = /^[\w .'-]{1,64}$/;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function parseCsv(text) {
    const delim = (() => { const first = text.slice(0, text.search(/\r?\n|$/)); return first.includes('\t') ? '\t' : first.includes(';') && !first.includes(',') ? ';' : ','; })();
    const rows = [];
    let row = [], cur = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
      else if (ch === '"' && cur === '') q = true;
      else if (ch === delim) { row.push(cur); cur = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
      else cur += ch;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows.filter(r => r.some(v => v.trim() !== ''));
  }
  /** Spreadsheets are parsed in the same isolated worker Sheets uses. */
  function parseSheet(bytes) {
    return new Promise((resolve, reject) => {
      const w = new Worker('js/xlsx-worker.js');
      const timer = setTimeout(() => { w.terminate(); reject(new Error('the file took too long to read')); }, 30000);
      w.onmessage = e => {
        clearTimeout(timer); w.terminate();
        if (!e.data.ok) return reject(new Error(e.data.error));
        const ws = e.data.Sheets[0] || {};
        const rows = [];
        for (const k of Object.keys(ws)) {
          if (k[0] === '!') continue;
          const m = /^([A-Z]+)(\d+)$/.exec(k);
          if (!m) continue;
          const c = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, r = +m[2] - 1;
          if (r > 100000 || c > 200) continue;
          (rows[r] = rows[r] || [])[c] = ws[k].w != null ? String(ws[k].w) : ws[k].v == null ? '' : String(ws[k].v);
        }
        resolve(rows.filter(Boolean).map(r => Array.from(r, v => v || '')));
      };
      w.onerror = err => { clearTimeout(timer); w.terminate(); reject(new Error(err.message || 'unreadable file')); };
      const copy = bytes.slice();
      w.postMessage(copy, [copy.buffer]);
    });
  }

  function WriteMerge(o) {
    const { doc } = o;
    let fields = [], records = [], previewing = false, index = 0, sourceName = '';
    const fieldEls = () => [...doc.querySelectorAll('span.mergefield')].concat(...o.extraRoots().map(r => [...r.querySelectorAll('span.mergefield')]));

    async function selectRecipients() {
      const f = await Platform.open([{ name: 'Recipient lists', extensions: ['csv', 'tsv', 'txt', 'xlsx', 'xls', 'ods'] }]);
      if (!f) return;
      let rows;
      try {
        rows = /\.(xlsx|xls|ods)$/i.test(f.name) ? await parseSheet(f.data) : parseCsv(Platform.bytesToText(f.data).replace(/^﻿/, ''));
      } catch (err) { UI.toast('Could not read the list: ' + err.message, 4000); return; }
      if (rows.length < 2) { UI.toast('The list needs a header row and at least one recipient', 4000); return; }
      const head = rows[0].map((h, i) => String(h).trim().replace(/[^\w .'-]/g, '').slice(0, 64) || 'Column ' + (i + 1));
      fields = head;
      records = rows.slice(1).map(r => Object.fromEntries(head.map((h, i) => [h, r[i] == null ? '' : String(r[i])])));
      sourceName = f.name;
      index = 0;
      o.onChange();
      UI.toast(`${records.length} recipient${records.length === 1 ? '' : 's'} from ${f.name}`);
    }
    function insertField(name) {
      if (!FIELD.test(name)) return;
      o.restoreSelection();
      Platform.insertInline(`<span class="mergefield" data-field="${esc(name)}" contenteditable="false">«${esc(name)}»</span>`);
      if (previewing) showRecord();
      o.changed();
    }
    function showRecord() {
      const rec = records[index] || {};
      fieldEls().forEach(el => {
        el.contentEditable = 'false';
        const name = el.dataset.field;
        const t = previewing && records.length ? (rec[name] != null ? rec[name] : '') : `«${name}»`;
        if (el.textContent !== t) el.textContent = t;
        el.classList.toggle('preview', previewing);
      });
      o.onChange();
    }
    function setPreview(on) {
      if (on && !records.length) { UI.toast('Select recipients first'); return false; }
      previewing = !!on;
      showRecord();
      return previewing;
    }
    function go(delta) {
      if (!records.length) return;
      index = Math.max(0, Math.min(records.length - 1, index + delta));
      showRecord();
    }
    /** One copy of the document per recipient, separated by page breaks. */
    function mergedBody(bodyHtml) {
      const parts = records.map(rec => {
        const d = new DOMParser().parseFromString(`<body>${bodyHtml}</body>`, 'text/html');
        d.querySelectorAll('span.mergefield').forEach(el => el.replaceWith(d.createTextNode(rec[el.dataset.field] != null ? rec[el.dataset.field] : '')));
        return d.body.innerHTML;
      });
      return parts.join('<hr class="page-break">');
    }
    async function finish() {
      if (!records.length) { UI.toast('Select recipients first'); return; }
      if (!fieldEls().length) { UI.toast('Insert at least one merge field first'); return; }
      const was = previewing;
      setPreview(false);
      const html = o.buildHtml(mergedBody);
      setPreview(was);
      await Platform.openData('write', `Merged ${o.docName()}`.replace(/\.(docx|html?|md|txt)$/i, '') + '.docx', new TextEncoder().encode(html));
      UI.toast(`Merged ${records.length} document${records.length === 1 ? '' : 's'}`);
    }
    /** For .docx: merge fields become MERGEFIELD markers turned into Word fields after conversion. */
    function prepareForDocx(root) {
      root.querySelectorAll('span.mergefield').forEach(el => {
        const name = (el.dataset.field || '').replace(/[^\w .'-]/g, '');
        el.replaceWith(document.createTextNode(`\uE040${name}\uE0FF«${name}»\uE041\uE0FF`));
      });
    }

    return {
      selectRecipients, insertField, setPreview, finish, prepareForDocx, refresh: showRecord,
      next: () => go(1), previous: () => go(-1), first: () => go(-Infinity), last: () => go(Infinity),
      get fields() { return fields; }, get count() { return records.length; }, get index() { return index; },
      get previewing() { return previewing; }, get source() { return sourceName; },
      fieldsInDocument: () => [...new Set(fieldEls().map(e => e.dataset.field))]
    };
  }

  window.WriteMerge = WriteMerge;
})();
