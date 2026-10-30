'use strict';
// Isadek Write — form fields: text boxes, check boxes, drop-down lists and date pickers that
// people fill in, in the document. "Restrict to filling in forms" leaves only the fields
// editable. In .docx they are Word content controls (and Word's are read as fields here).
//
// A field is <span class="ff" data-ff="text|check|drop|date" data-v="value" title="label"
// data-ph="placeholder" data-opts='["A","B"]'>shown text</span>, made non-editable on screen.
(function () {
  const TYPES = { text: 'Text', check: 'Check box', drop: 'Drop-down list', date: 'Date' };
  const BOX = { on: '☒', off: '☐' };
  // Markers that carry fields through the .docx converters (private-use characters).
  const MK = { s: '', e: '' };
  const MK_RE = /(\d+)/g;
  const xmlEsc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
  const unXml = t => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const optsOf = el => { try { const o = JSON.parse(el.dataset.opts || '[]'); return Array.isArray(o) ? o.filter(x => typeof x === 'string').map(x => x.slice(0, 200)).slice(0, 100) : []; } catch { return []; } };
  const ISO = /^\d{4}-\d{2}-\d{2}$/;
  const showDate = v => { if (!ISO.test(v)) return v; const [y, m, d] = v.split('-').map(Number); return new Date(y, m - 1, d).toLocaleDateString(); };

  /** Draws a field's text from its value (or placeholder), and makes it a single unit. */
  function paint(el) {
    const kind = el.dataset.ff, v = el.dataset.v || '';
    el.setAttribute('contenteditable', 'false');
    el.tabIndex = 0;
    el.setAttribute('role', kind === 'check' ? 'checkbox' : 'button');
    if (kind === 'check') { el.textContent = v === '1' ? BOX.on : BOX.off; el.setAttribute('aria-checked', String(v === '1')); }
    else el.textContent = v ? (kind === 'date' ? showDate(v) : v) : (el.dataset.ph || (kind === 'drop' ? 'Choose an item' : kind === 'date' ? 'Enter a date' : 'Click to enter text'));
    el.classList.toggle('ff-empty', !v && kind !== 'check');
    el.setAttribute('aria-label', (el.title || TYPES[kind]) + (kind === 'check' ? '' : ': ' + (v ? el.textContent : 'empty')));
  }
  const render = (root) => root.querySelectorAll('span.ff').forEach(paint);

  function WriteForms(o) {
    const { doc } = o;
    let pop = null;
    function closePop() { if (pop) { pop.remove(); pop = null; } }

    /** Fills in a field: check boxes toggle; the others open a small editor beside the field. */
    function fill(el) {
      const kind = el.dataset.ff;
      if (kind === 'check') {
        o.snapshot();
        el.dataset.v = el.dataset.v === '1' ? '' : '1';
        paint(el);
        o.changed();
        return;
      }
      closePop();
      pop = document.createElement('div');
      pop.className = 'ff-pop';
      let input;
      if (kind === 'drop') {
        input = document.createElement('select');
        const none = document.createElement('option'); none.value = ''; none.textContent = el.dataset.ph || 'Choose an item'; input.appendChild(none);
        optsOf(el).forEach(x => { const op = document.createElement('option'); op.value = op.textContent = x; input.appendChild(op); });
      } else {
        input = document.createElement('input');
        input.type = kind === 'date' ? 'date' : 'text';
        input.maxLength = 1000;
      }
      input.value = el.dataset.v || '';
      input.setAttribute('aria-label', el.title || TYPES[kind]);
      pop.appendChild(input);
      document.body.appendChild(pop);
      const r = el.getBoundingClientRect();
      Object.assign(pop.style, { left: Math.min(r.left, innerWidth - 280) + 'px', top: (r.bottom + 4) + 'px' });
      const commit = () => {
        if (!pop) return;
        const v = String(input.value).slice(0, 1000);
        if (kind === 'date' && v && !ISO.test(v)) return closePop();
        if (v !== (el.dataset.v || '')) { o.snapshot(); if (v) el.dataset.v = v; else delete el.dataset.v; paint(el); o.changed(); }
        closePop();
        el.focus();
      };
      input.addEventListener('keydown', e => { e.stopPropagation(); if (e.key === 'Enter') { e.preventDefault(); commit(); } if (e.key === 'Escape') { closePop(); el.focus(); } });
      input.addEventListener('change', () => { if (kind !== 'text') commit(); });
      input.addEventListener('blur', () => setTimeout(commit, 0));
      input.focus();
      if (input.select) input.select();
    }
    doc.addEventListener('click', e => { const el = e.target.closest && e.target.closest('span.ff'); if (el && doc.contains(el)) { e.preventDefault(); fill(el); } });
    doc.addEventListener('keydown', e => {
      const el = e.target.closest && e.target.closest('span.ff');
      if (el && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); fill(el); }
      // Filling in a restricted form: Tab goes from field to field.
      if (e.key === 'Tab' && document.body.classList.contains('forms-only')) {
        const all = [...doc.querySelectorAll('span.ff')];
        if (!all.length) return;
        e.preventDefault();
        const at = all.indexOf(el);
        all[(at + (e.shiftKey ? -1 : 1) + all.length) % all.length].focus();
      }
    }, true);

    async function insert() {
      const v = await UI.form('Insert Form Field', [
        { name: 'kind', label: 'Kind', type: 'select', options: Object.entries(TYPES).map(([value, label]) => ({ value, label })), value: 'text' },
        { name: 'title', label: 'Label (read out by screen readers, shown on hover)', placeholder: 'e.g. Full name' },
        { name: 'ph', label: 'Placeholder text (empty fields)', placeholder: 'Click to enter text' },
        { name: 'opts', label: 'Drop-down items, one per line', type: 'textarea' }
      ], 'Insert');
      if (!v) return;
      const el = document.createElement('span');
      el.className = 'ff';
      el.dataset.ff = v.kind;
      if (v.title.trim()) el.title = v.title.trim().slice(0, 200);
      if (v.ph.trim() && v.kind !== 'check') el.dataset.ph = v.ph.trim().slice(0, 200);
      if (v.kind === 'drop') {
        const opts = [...new Set(v.opts.split('\n').map(x => x.trim()).filter(Boolean))].slice(0, 100);
        el.dataset.opts = JSON.stringify(opts.length ? opts : ['Option 1', 'Option 2']);
      }
      paint(el);
      o.insertInline(el.outerHTML);
      render(doc);
      o.changed();
    }

    /** Restrict editing: only the fields can be changed (and Tab moves between them). */
    function setFormsOnly(on) {
      document.body.classList.toggle('forms-only', !!on);
      doc.setAttribute('contenteditable', on ? 'false' : 'true');
      doc.querySelectorAll('span.ff').forEach(f => f.setAttribute('contenteditable', 'false'));
      const b = document.getElementById('formsOnly');
      if (b) { b.classList.toggle('on', !!on); b.setAttribute('aria-pressed', String(!!on)); }
    }

    // ---------- .docx: content controls ----------
    function prepareForDocx(body) {
      const list = [];
      body.querySelectorAll('span.ff').forEach(el => {
        list.push({ kind: el.dataset.ff, v: el.dataset.v || '', title: el.title || '', ph: el.dataset.ph || '', opts: optsOf(el) });
        el.replaceWith(document.createTextNode(MK.s + (list.length - 1) + MK.e));
      });
      return list;
    }
    function sdtXml(f, rPr, id) {
      const alias = f.title ? `<w:alias w:val="${xmlEsc(f.title)}"/><w:tag w:val="${xmlEsc(f.title)}"/>` : '';
      const run = t => `<w:r>${rPr}<w:t xml:space="preserve">${xmlEsc(t)}</w:t></w:r>`;
      const empty = !f.v && f.kind !== 'check';
      const shown = f.kind === 'check' ? (f.v === '1' ? BOX.on : BOX.off) : f.v || f.ph || (f.kind === 'drop' ? 'Choose an item.' : f.kind === 'date' ? 'Click or tap to enter a date.' : 'Click or tap here to enter text.');
      let pr = `<w:sdtPr>${alias}<w:id w:val="${id}"/>${empty ? '<w:showingPlcHdr/>' : ''}`;
      if (f.kind === 'check') pr += `<w14:checkbox><w14:checked w14:val="${f.v === '1' ? 1 : 0}"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>`;
      else if (f.kind === 'drop') pr += `<w:dropDownList>${f.opts.map(x => `<w:listItem w:displayText="${xmlEsc(x)}" w:value="${xmlEsc(x)}"/>`).join('')}</w:dropDownList>`;
      else if (f.kind === 'date') pr += `<w:date${f.v ? ` w:fullDate="${f.v}T00:00:00Z"` : ''}><w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>`;
      else pr += '<w:text/>';
      pr += '</w:sdtPr>';
      const text = f.kind === 'date' && f.v ? (() => { const [y, m, d] = f.v.split('-').map(Number); return `${m}/${d}/${y}`; })() : shown;
      return `<w:sdt>${pr}<w:sdtContent>${run(text)}</w:sdtContent></w:sdt>`;
    }
    /** Replaces the markers in a .docx with Word content controls; "forms only" protects the file. */
    async function patchDocx(bytes, list, formsOnly) {
      if (!list.length && !formsOnly) return bytes;
      const zip = await JSZip.loadAsync(bytes);
      if (list.length) {
        const f = zip.file('word/document.xml');
        let xml = await f.async('string');
        let id = 900000;
        xml = xml.replace(/<w:r>((?:(?!<\/w:r>)[\s\S])*?)<\/w:r>/g, (all, inner) => {
          if (!//.test(inner)) return all;
          const rPr = (inner.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
          const text = [...inner.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
          const runOf = t => (t ? `<w:r>${rPr}<w:t xml:space="preserve">${t}</w:t></w:r>` : '');
          let out = '', last = 0;
          text.replace(MK_RE, (m, n, at) => { out += runOf(text.slice(last, at)); last = at + m.length; const fd = list[+n]; if (fd) out += sdtXml(fd, rPr, id++); return m; });
          return out + runOf(text.slice(last));
        });
        if (!/xmlns:w14=/.test(xml.slice(0, 4000))) xml = xml.replace(/<w:document\b/, '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"');
        zip.file('word/document.xml', xml);
      }
      if (formsOnly) {
        const sf = zip.file('word/settings.xml');
        if (sf) {
          let s = await sf.async('string');
          s = s.replace(/<w:documentProtection\b[^>]*\/>/g, '');
          s = s.replace(/(<w:settings\b[^>]*>)/, '$1<w:documentProtection w:edit="forms" w:enforcement="1"/>');
          zip.file('word/settings.xml', s);
        }
      }
      return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
    }
    /** Word's content controls (and Isadek's) become markers for the HTML converter; returns what they were. */
    async function preprocessDocx(bytes) {
      const zip = await JSZip.loadAsync(bytes);
      const f = zip.file('word/document.xml');
      const meta = { fields: [], formsOnly: false };
      if (!f) return { bytes, meta };
      let xml = await f.async('string');
      const resultOf = x => [...x.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(m => unXml(m[1])).join('');
      const attr = (x, tag, name) => { const m = x.match(new RegExp(`<${tag}\\b[^>]*\\b${name}="([^"]*)"`)); return m ? unXml(m[1]) : null; };
      // Inline content controls only (a control around whole paragraphs stays as its content).
      xml = xml.replace(/<w:sdt>(?:(?!<w:sdt>)[\s\S])*?<\/w:sdt>/g, all => {
        const pr = (all.match(/<w:sdtPr>[\s\S]*?<\/w:sdtPr>/) || [''])[0];
        const content = (all.match(/<w:sdtContent>([\s\S]*)<\/w:sdtContent>/) || [])[1] || '';
        if (/<w:p\b/.test(content) || /<w:tc\b|<w:tr\b/.test(content)) return all;
        const title = attr(pr, 'w:alias', 'w:val') || attr(pr, 'w:tag', 'w:val') || '';
        const placeholder = /<w:showingPlcHdr\b/.test(pr);
        let fd = null;
        if (/<w14:checkbox\b/.test(pr)) fd = { kind: 'check', v: attr(pr, 'w14:checked', 'w14:val') === '1' || attr(pr, 'w14:checked', 'w14:val') === 'true' ? '1' : '' };
        else if (/<w:(dropDownList|comboBox)\b/.test(pr)) {
          const opts = (pr.match(/<w:listItem\b[^>]*\/>/g) || []).map(t => unXml((t.match(/w:displayText="([^"]*)"/) || t.match(/w:value="([^"]*)"/) || [, ''])[1])).filter(Boolean).slice(0, 100);
          const shown = resultOf(content);
          fd = { kind: 'drop', opts, v: placeholder ? '' : shown, ph: placeholder ? shown : '' };
        } else if (/<w:date\b/.test(pr)) {
          const full = attr(pr, 'w:date', 'w:fullDate');
          fd = { kind: 'date', v: !placeholder && full && /^\d{4}-\d{2}-\d{2}/.test(full) ? full.slice(0, 10) : '', ph: placeholder ? resultOf(content) : '' };
        } else if (/<w:text\b/.test(pr) || /<w:richText\b/.test(pr) || !/<w:(picture|docPartObj|citation|bibliography|group|equation)\b/.test(pr)) {
          const shown = resultOf(content);
          fd = { kind: 'text', v: placeholder ? '' : shown, ph: placeholder ? shown : '' };
        }
        if (!fd) return all;
        fd.title = title.slice(0, 200);
        meta.fields.push(fd);
        const rPr = (content.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
        return `<w:r>${rPr}<w:t xml:space="preserve">${MK.s}${meta.fields.length - 1}${MK.e}</w:t></w:r>`;
      });
      zip.file('word/document.xml', xml);
      const sf = zip.file('word/settings.xml');
      if (sf) meta.formsOnly = /<w:documentProtection\b[^>]*w:edit="forms"[^>]*w:enforcement="(1|true|on)"/.test(await sf.async('string'));
      return { bytes: new Uint8Array(await zip.generateAsync({ type: 'uint8array' })), meta };
    }
    function postprocessHtml(html, meta) {
      if (!meta || !meta.fields.length) return html;
      return html.replace(MK_RE, (m, n) => {
        const fd = meta.fields[+n];
        if (!fd) return '';
        const el = document.createElement('span');
        el.className = 'ff';
        el.dataset.ff = fd.kind;
        if (fd.v) el.dataset.v = fd.v.slice(0, 1000);
        if (fd.title) el.title = fd.title;
        if (fd.ph) el.dataset.ph = fd.ph.slice(0, 200);
        if (fd.kind === 'drop') el.dataset.opts = JSON.stringify(fd.opts || []);
        paint(el);
        return el.outerHTML;
      });
    }
    /** Plain text and Markdown: a field is what it shows. */
    function prepareForText(root) { root.querySelectorAll('span.ff').forEach(el => el.replaceWith(document.createTextNode(el.dataset.ff === 'check' ? (el.dataset.v === '1' ? '[x]' : '[ ]') : el.dataset.v || ''))); }

    return { insert, render: () => render(doc), fill, setFormsOnly, prepareForDocx, patchDocx, preprocessDocx, postprocessHtml, prepareForText, closePop };
  }
  WriteForms.paint = paint;
  window.WriteForms = WriteForms;
})();
