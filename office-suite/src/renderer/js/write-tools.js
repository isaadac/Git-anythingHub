'use strict';
// Isadek Write — page and writing tools: watermarks, line numbers, drop caps, the outline,
// envelopes and labels, readability, and translation.
//
// Watermark and line numbers are page settings (setup.watermark, setup.lineNumbers), drawn on
// the page sheets and saved as Word's own (a header shape, the section's line numbering). A drop
// cap is class "dropcap" on a paragraph (Word's dropped frame in .docx).
(function () {
  const PX = 96;
  const xmlEsc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
  const unXml = t => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const MK = { wm: '', dcS: '', dcE: '' };
  const WM_PRESETS = ['CONFIDENTIAL', 'DRAFT', 'DO NOT COPY', 'SAMPLE', 'URGENT'];

  /** A watermark from a file: short text, a colour, diagonal or straight. */
  function cleanWatermark(v) {
    if (!v || typeof v !== 'object' || typeof v.text !== 'string' || !v.text.trim()) return null;
    return { text: v.text.replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, 60), color: /^#[0-9a-f]{6}$/i.test(v.color) ? v.color.toLowerCase() : '#c0c0c0', diagonal: v.diagonal !== false };
  }
  const LINE_MODES = ['none', 'continuous', 'page'];

  function WriteTools(o) {
    const { doc, page } = o;
    const setup = () => o.setup;

    // ================= Watermark =================
    async function watermarkDialog() {
      const cur = setup().watermark;
      const v = await UI.form('Watermark', [
        { name: 'preset', label: 'Watermark', type: 'select', options: [{ value: '', label: 'No watermark' }, ...WM_PRESETS.map(t => ({ value: t, label: t })), { value: '*', label: 'Custom text…' }], value: cur ? (WM_PRESETS.includes(cur.text) ? cur.text : '*') : '' },
        { name: 'text', label: 'Custom text', value: cur && !WM_PRESETS.includes(cur.text) ? cur.text : '' },
        { name: 'color', label: 'Colour', type: 'color', value: cur ? cur.color : '#c0c0c0' },
        { name: 'layout', label: 'Layout', type: 'select', options: [{ value: 'd', label: 'Diagonal' }, { value: 'h', label: 'Horizontal' }], value: cur && !cur.diagonal ? 'h' : 'd' }
      ], 'Apply');
      if (!v) return;
      const text = v.preset === '*' ? v.text : v.preset;
      o.snapshotSetup();
      setup().watermark = cleanWatermark({ text, color: v.color, diagonal: v.layout !== 'h' });
      o.applySetup();
      o.changed();
    }
    /** Draws the watermark and line numbers on the page sheets (called after each layout). */
    function decoratePages(bgs, g, pitch) {
      const wm = setup().watermark;
      if (wm) for (const bg of bgs.children) {
        const el = document.createElement('div');
        el.className = 'wm';
        el.textContent = wm.text;
        const room = wm.diagonal ? Math.hypot(g.W, g.H) * 0.78 : g.W * 0.82;
        el.style.fontSize = Math.max(18, Math.min(g.W * 0.2, room / Math.max(3, wm.text.length) * 1.6)) + 'px';
        el.style.color = wm.color;
        el.style.transform = `translate(-50%, -50%) rotate(${wm.diagonal ? -45 : 0}deg)`;
        bg.appendChild(el);
      }
      if (setup().lineNumbers && setup().lineNumbers !== 'none') drawLineNumbers(bgs, g, pitch);
    }

    // ================= Line numbers =================
    /** The lines of the text body: each block's line boxes, top to bottom (tables and boxes aren't numbered, as in Word). */
    function lineTops() {
      const z = parseFloat(page.style.zoom) || 1;
      const origin = page.getBoundingClientRect();
      const tops = [];
      const blocks = doc.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, pre, blockquote:not(:has(p, li))');
      const range = document.createRange();
      for (const b of blocks) {
        if (b.closest('table, .textbox, .toc, .note-area, [data-noln]') || b.querySelector('p, li, h1, h2, h3, h4, h5, h6')) continue;
        if (!b.textContent.trim() && !b.querySelector('img, span.equation, span.ff')) continue;
        range.selectNodeContents(b);
        // Line boxes: a drop cap's tall letter spans several lines and isn't one of them.
        const rects = [...range.getClientRects()].filter(r => r.width >= 1 && r.height >= 4).map(r => ({ top: (r.top - origin.top) / z, h: r.height / z }));
        if (!rects.length) continue;
        const hs = rects.map(r => r.h).sort((a, c) => a - c), med = hs[Math.floor(hs.length / 2)];
        let last = -1e9;
        for (const r of rects.filter(x => x.h <= med * 1.6).sort((a, c) => a.top - c.top)) {
          if (r.top - last < med * 0.5) continue; // the same line
          tops.push(r);
          last = r.top;
        }
      }
      return tops.sort((a, b) => a.top - b.top);
    }
    function drawLineNumbers(bgs, g, pitch) {
      const mode = setup().lineNumbers;
      const pagesEls = [...bgs.children];
      let n = 0, pageAt = -1;
      for (const line of lineTops()) {
        const i = Math.min(pagesEls.length - 1, Math.max(0, Math.floor(line.top / pitch)));
        if (mode === 'page' && i !== pageAt) n = 0;
        pageAt = i;
        n++;
        const el = document.createElement('span');
        el.className = 'ln';
        el.textContent = n;
        el.style.top = (line.top - i * pitch) + 'px';
        el.style.height = line.h + 'px';
        el.style.lineHeight = line.h + 'px';
        el.style.right = (g.W - g.M + 12) + 'px';
        pagesEls[i].appendChild(el);
      }
    }

    // ================= Drop cap =================
    function toggleDropCap() {
      const block = o.currentBlock();
      if (!block || block.tagName !== 'P' || !doc.contains(block)) return UI.toast('Put the cursor in a paragraph');
      if (!block.textContent.trim()) return UI.toast('A drop cap needs text in the paragraph');
      o.snapshot();
      block.classList.toggle('dropcap');
      document.getElementById('dropCapBtn')?.classList.toggle('on', block.classList.contains('dropcap'));
      o.changed();
    }
    /** .docx: each drop cap's first letter is marked, to become Word's dropped frame. */
    function prepareForDocx(body) {
      let n = 0;
      body.querySelectorAll('p.dropcap').forEach(p => {
        const tw = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        let t;
        while ((t = tw.nextNode()) && !t.nodeValue.trim()) { /* skip empty text */ }
        if (!t) return;
        const i = t.nodeValue.search(/\S/);
        const ch = [...t.nodeValue.slice(i)][0];
        t.nodeValue = t.nodeValue.slice(0, i) + MK.dcS + ch + MK.dcE + t.nodeValue.slice(i + ch.length);
        p.classList.remove('dropcap');
        n++;
      });
      return n;
    }
    async function patchDocx(bytes, { dropCaps, watermark, lineNumbers }) {
      if (!dropCaps && !watermark) return bytes;
      const zip = await JSZip.loadAsync(bytes);
      if (dropCaps) {
        const f = zip.file('word/document.xml');
        let xml = await f.async('string');
        xml = xml.replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*?[\s\S]*?<\/w:p>/g, para => {
          const m = para.match(/<w:r>((?:(?!<\/w:r>)[\s\S])*?)([^]*)/);
          if (!m) return para;
          const rPr = (m[1].match(/<w:rPr>([\s\S]*?)<\/w:rPr>/) || [, ''])[1].replace(/<w:(sz|szCs|position)\b[^>]*\/>/g, '');
          const letter = m[2];
          const drop = `<w:p><w:pPr><w:keepNext/><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/><w:spacing w:before="0" w:after="0" w:line="827" w:lineRule="exact"/><w:textAlignment w:val="baseline"/></w:pPr><w:r><w:rPr>${rPr}<w:position w:val="-8"/><w:sz w:val="103"/><w:szCs w:val="103"/></w:rPr><w:t>${letter}</w:t></w:r></w:p>`;
          return drop + para.replace(/[^]*/, '');
        });
        zip.file('word/document.xml', xml);
      }
      if (watermark) {
        // Word's watermark: a WordArt shape behind the text, in the header of every page.
        const rot = watermark.diagonal ? 'rotation:315;' : '';
        const len = Math.max(4, watermark.text.length);
        const shape = `<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:path textpathok="t" o:connecttype="custom"/><v:textpath on="t" fitshape="t"/><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype><v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:${Math.min(468, len * 52)}pt;height:${Math.min(117, Math.max(60, 468 / len * 1.6))}pt;${rot}z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="${watermark.color}" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="${xmlEsc(watermark.text)}"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>`;
        for (const name of Object.keys(zip.files).filter(n => /^word\/header\d*\.xml$/.test(n))) {
          let xml = await zip.file(name).async('string');
          if (!xml.includes(MK.wm)) continue;
          xml = xml.replace(/<w:r>(?:(?!<\/w:r>)[\s\S])*?(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/, shape);
          for (const [p, ns] of [['v', 'urn:schemas-microsoft-com:vml'], ['o', 'urn:schemas-microsoft-com:office:office'], ['w10', 'urn:schemas-microsoft-com:office:word']]) if (!new RegExp(`xmlns:${p}=`).test(xml.slice(0, 4000))) xml = xml.replace(/<w:hdr\b/, `<w:hdr xmlns:${p}="${ns}"`);
          zip.file(name, xml);
        }
      }
      void lineNumbers;
      return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
    }
    /** Opening a .docx: Word's drop caps join their paragraph again; the watermark and line numbers become settings. */
    async function preprocessDocx(bytes) {
      const zip = await JSZip.loadAsync(bytes);
      const meta = { watermark: null, lineNumbers: 'none', dropCaps: 0 };
      const f = zip.file('word/document.xml');
      if (!f) return { bytes, meta };
      let xml = await f.async('string');
      xml = xml.replace(/<w:p\b[^>]*>(?:(?!<\/w:p>)[\s\S])*?<w:framePr\b[^>]*w:dropCap="(?:drop|margin)"[\s\S]*?<\/w:p>(\s*<w:p\b[^>]*>)/g, (all, nextOpen) => {
        const para = all.slice(0, all.length - nextOpen.length);
        const letter = [...para.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
        if (!letter) return all;
        meta.dropCaps++;
        return nextOpen + `<w:r><w:t xml:space="preserve">${MK.dcS}${letter}${MK.dcE}</w:t></w:r>`;
      });
      zip.file('word/document.xml', xml);
      const sect = (xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>(?![\s\S]*<w:sectPr\b)/) || [''])[0];
      const ln = sect.match(/<w:lnNumType\b[^>]*>/);
      if (ln) meta.lineNumbers = /w:restart="continuous"/.test(ln[0]) ? 'continuous' : 'page';
      for (const name of Object.keys(zip.files).filter(n => /^word\/header\d*\.xml$/.test(n))) {
        const h = await zip.file(name).async('string');
        const shape = h.match(/<v:shape\b[^>]*(?:PowerPlusWaterMarkObject|WaterMark)[^>]*>[\s\S]*?<\/v:shape>/i);
        if (!shape) continue;
        const text = (shape[0].match(/<v:textpath\b[^>]*\bstring="([^"]*)"/) || [])[1];
        if (!text) continue;
        const color = (shape[0].match(/\bfillcolor="(#[0-9a-f]{6})"/i) || [])[1] || (/fillcolor="silver"/.test(shape[0]) ? '#c0c0c0' : '#c0c0c0');
        meta.watermark = cleanWatermark({ text: unXml(text), color, diagonal: /rotation:\s*3[01]\d/.test(shape[0]) });
        break;
      }
      return { bytes: new Uint8Array(await zip.generateAsync({ type: 'uint8array' })), meta };
    }
    function postprocessHtml(html) {
      if (!html.includes(MK.dcS)) return html;
      const d = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
      const tw = d.createTreeWalker(d.body, NodeFilter.SHOW_TEXT);
      const found = [];
      while (tw.nextNode()) if (tw.currentNode.nodeValue.includes(MK.dcS)) found.push(tw.currentNode);
      for (const t of found) {
        t.nodeValue = t.nodeValue.replace(/([^]*)/g, '$1');
        const p = t.parentNode.closest ? t.parentNode.closest('p') : null;
        if (p) p.classList.add('dropcap');
      }
      return d.body.innerHTML;
    }

    // ================= Outline =================
    const LEVEL = el => (/^H[1-6]$/.test(el.tagName) ? +el.tagName[1] : 0);
    /** A heading and the blocks under it (until the next heading of its level or higher). */
    function sectionOf(h) {
      const lv = LEVEL(h), out = [h];
      for (let n = h.nextElementSibling; n && !(LEVEL(n) && LEVEL(n) <= lv); n = n.nextElementSibling) out.push(n);
      return out;
    }
    const headingAt = () => { let b = o.currentBlock(); while (b && b.parentElement !== doc) b = b.parentElement; for (let n = b; n; n = n.previousElementSibling) if (LEVEL(n)) return n; return null; };
    function moveSection(dir, h = headingAt()) {
      if (!h) return UI.toast('Put the cursor in a heading (or under one)');
      const lv = LEVEL(h), mine = sectionOf(h);
      if (dir < 0) {
        let prev = h.previousElementSibling;
        while (prev && !(LEVEL(prev) && LEVEL(prev) <= lv)) prev = prev.previousElementSibling;
        if (!prev || LEVEL(prev) < lv) return UI.toast('This is the first section at its level');
        o.snapshot();
        mine.forEach(n => prev.before(n));
      } else {
        const after = mine[mine.length - 1].nextElementSibling;
        if (!after || LEVEL(after) < lv || !LEVEL(after)) return UI.toast('This is the last section at its level');
        const theirs = sectionOf(after);
        o.snapshot();
        theirs[theirs.length - 1].after(...mine);
      }
      o.changed();
      h.scrollIntoView({ block: 'nearest' });
      drawOutline();
    }
    /** Promote or demote a heading (and the headings under it). */
    function shiftLevel(d, h = headingAt()) {
      if (!h) return UI.toast('Put the cursor in a heading');
      const lv = LEVEL(h);
      if ((d < 0 && lv === 1) || (d > 0 && lv === 6)) return;
      o.snapshot();
      sectionOf(h).filter(n => LEVEL(n)).forEach(n => {
        const nl = Math.max(1, Math.min(6, LEVEL(n) + d));
        const r = document.createElement('h' + nl);
        [...n.attributes].forEach(a => r.setAttribute(a.name, a.value));
        r.append(...n.childNodes);
        n.replaceWith(r);
        if (n === h) h = r;
      });
      o.changed();
      drawOutline();
      o.placeCaret(h);
    }
    /** Outline view: headings with their levels; deeper levels and body text can be hidden. */
    function setOutline(on) {
      document.body.classList.toggle('view-outline', on);
      document.getElementById('viewOutline')?.classList.toggle('on', on);
      document.getElementById('outlineTools').hidden = !on;
      drawOutline();
      o.relayout();
    }
    function drawOutline() {
      const on = document.body.classList.contains('view-outline');
      const show = +(document.getElementById('outlineLevel')?.value || 9);
      for (const n of doc.children) {
        const lv = LEVEL(n);
        n.classList.toggle('ol-hide', on && (lv ? lv > show : show < 9));
        if (lv) n.dataset.ol = on ? String(lv) : ''; if (!on || !lv) delete n.dataset.ol;
      }
    }

    // ================= Envelopes and labels =================
    const RET = 'atlas-return-address';
    const remembered = () => { try { return localStorage.getItem(RET) || ''; } catch { return ''; } };
    const lines = t => String(t).split(/\r?\n/).map(l => esc(l.trim())).filter(Boolean).join('<br>');
    async function openNew(name, html, page) {
      const meta = `<meta name="atlas-setup" content="${esc(JSON.stringify({ ...page }))}">`;
      const full = `<!doctype html><html><head><meta charset="utf-8">${meta}<title>${esc(name)}</title></head><body><main class="atlas-doc">${html}</main></body></html>`;
      await Platform.openData('write', name + '.html', new TextEncoder().encode(full));
    }
    async function envelopes() {
      const sel = getSelection();
      const picked = sel.rangeCount && doc.contains(sel.anchorNode) ? sel.toString().trim() : '';
      const v = await UI.form('Envelopes', [
        { name: 'to', label: 'Delivery address', type: 'textarea', value: picked },
        { name: 'from', label: 'Return address', type: 'textarea', value: remembered() },
        { name: 'size', label: 'Envelope size', type: 'select', options: [{ value: 'env10', label: 'Size 10 (4⅛ × 9½ in)' }, { value: 'dl', label: 'DL (110 × 220 mm)' }, { value: 'c5', label: 'C5 (162 × 229 mm)' }], value: 'env10' }
      ], 'Create Envelope');
      if (!v) return;
      if (!v.to.trim()) return UI.toast('Type the delivery address');
      try { localStorage.setItem(RET, v.from); } catch { /* private mode */ }
      const big = v.size === 'c5';
      const html = (v.from.trim() ? `<p style="font-size:10pt;line-height:1.25;margin:0">${lines(v.from)}</p>` : '') +
        `<p style="margin:${big ? 1.6 : 1.1}in 0 0 ${big ? 3.4 : 3.6}in;font-size:12pt;line-height:1.3">${lines(v.to)}</p>`;
      await openNew('Envelope', html, { size: v.size, orientation: 'landscape', margin: '0.5in' });
    }
    const LABELS = {
      'avery5160': { name: 'Avery 5160 / 8160 — Address (30 per sheet)', size: 'letter', cols: 3, rows: 10, w: 2.625, h: 1, top: 0.5, left: 0.1875, gapX: 0.125, gapY: 0 },
      'avery5163': { name: 'Avery 5163 / 8163 — Shipping (10 per sheet)', size: 'letter', cols: 2, rows: 5, w: 4, h: 2, top: 0.5, left: 0.15625, gapX: 0.1875, gapY: 0 },
      'avery5167': { name: 'Avery 5167 — Return address (80 per sheet)', size: 'letter', cols: 4, rows: 20, w: 1.75, h: 0.5, top: 0.5, left: 0.28125, gapX: 0.3125, gapY: 0 },
      'l7160': { name: 'Avery L7160 — A4 (21 per sheet)', size: 'a4', cols: 3, rows: 7, w: 2.5, h: 1.5, top: 0.6, left: 0.28, gapX: 0.1, gapY: 0 }
    };
    async function labels() {
      const recips = o.mergeRecords();
      const v = await UI.form('Labels', [
        { name: 'kind', label: 'Labels', type: 'select', options: Object.entries(LABELS).map(([value, l]) => ({ value, label: l.name })), value: 'avery5160' },
        { name: 'text', label: recips.length ? `Address (merge fields like «${o.mergeFields()[0] || 'Name'}» are filled from the ${recips.length} recipients)` : 'Address (the same on every label)', type: 'textarea', value: recips.length ? o.mergeFields().slice(0, 4).map(f => `«${f}»`).join('\n') : remembered() },
        ...(recips.length ? [{ name: 'each', label: 'Make', type: 'select', options: [{ value: '1', label: 'One label per recipient' }, { value: '0', label: 'A full sheet of the same label' }], value: '1' }] : [])
      ], 'Create Labels');
      if (!v) return;
      if (!v.text.trim()) return UI.toast('Type what goes on the labels');
      const L = LABELS[v.kind] || LABELS.avery5160;
      const fill = rec => v.text.replace(/«([^»]{1,64})»/g, (m, k) => (rec && rec[k] != null ? String(rec[k]) : ''));
      const items = recips.length && v.each === '1' ? recips.map(fill) : Array.from({ length: L.cols * L.rows }, () => fill(recips[0] || null));
      const per = L.cols * L.rows, sheets = [];
      for (let i = 0; i < items.length; i += per) sheets.push(items.slice(i, i + per));
      const cell = t => `<td style="width:${L.w}in;height:${L.h}in;max-height:${L.h}in;padding:0 0.1in;vertical-align:middle;overflow:hidden;font-size:${L.h < 0.8 ? 7 : 10}pt;line-height:1.2;border:0">${lines(t)}</td>`;
      const gap = `<td style="width:${L.gapX}in;border:0;padding:0"></td>`;
      const html = sheets.map((sheet, si) => {
        const rows = [];
        for (let r = 0; r < L.rows; r++) rows.push(`<tr style="height:${L.h}in">${Array.from({ length: L.cols }, (_, c) => cell(sheet[r * L.cols + c] || '')).join(gap)}</tr>`);
        return `${si ? '<hr class="page-break">' : ''}<table data-labels="1" style="border-collapse:collapse;table-layout:fixed;margin:${L.top}in 0 0 ${L.left}in;border:0">${rows.join('')}</table>`;
      }).join('');
      await openNew('Labels', html, { size: L.size, orientation: 'portrait', margin: '0in' });
    }

    // ================= Readability =================
    const syllables = w => {
      w = w.toLowerCase().replace(/[^a-z]/g, '');
      if (!w) return 0;
      if (w.length <= 3) return 1;
      w = w.replace(/(?:[^laeiouy]es|[^laeiouy]ed|[^laeiouy]e)$/, '').replace(/^y/, '');
      return Math.max(1, (w.match(/[aeiouy]{1,2}/g) || []).length);
    };
    /** Sentences with their text node ranges, block by block. */
    function sentences() {
      const out = [];
      for (const b of doc.querySelectorAll('p, li, h1, h2, h3, h4, h5, h6, blockquote, td')) {
        if (b.querySelector('p, li, h1, h2, h3, h4, h5, h6') || b.closest('.toc')) continue;
        const nodes = [], tw = document.createTreeWalker(b, NodeFilter.SHOW_TEXT);
        let text = '';
        while (tw.nextNode()) { nodes.push({ n: tw.currentNode, at: text.length }); text += tw.currentNode.nodeValue; }
        const re = /[^.!?…]+(?:[.!?…]+["'”’)\]]*|$)/g;
        let m;
        while ((m = re.exec(text))) {
          const s = m[0];
          if (!s.trim()) { if (re.lastIndex === m.index) re.lastIndex++; continue; }
          const start = m.index + s.search(/\S/), end = m.index + s.trimEnd().length;
          const pos = i => { const k = nodes.findLast(x => x.at <= i) || nodes[0]; return [k.n, Math.min(k.n.nodeValue.length, i - k.at)]; };
          const words = (s.match(/[A-Za-zÀ-ÿ0-9'’-]+/g) || []);
          out.push({ text: s.trim(), words, range: () => { const r = document.createRange(); r.setStart(...pos(start)); r.setEnd(...pos(end)); return r; } });
          if (re.lastIndex === m.index) re.lastIndex++;
        }
      }
      return out;
    }
    function stats() {
      const ss = sentences().filter(s => s.words.length);
      const words = ss.flatMap(s => s.words);
      const syl = words.reduce((n, w) => n + syllables(w), 0);
      const W = Math.max(1, words.length), S = Math.max(1, ss.length);
      const passive = ss.filter(s => /\b(am|is|are|was|were|be|been|being)\s+(\w+ly\s+)?\w+(ed|en)\b/i.test(s.text)).length;
      return {
        words: words.length, sentences: ss.length, syllables: syl,
        wordsPerSentence: W / S, charsPerWord: words.join('').length / W,
        ease: 206.835 - 1.015 * (W / S) - 84.6 * (syl / W),
        grade: 0.39 * (W / S) + 11.8 * (syl / W) - 15.59,
        passive: passive / S,
        long: ss.filter(s => s.words.length > 25).length, veryLong: ss.filter(s => s.words.length > 35).length
      };
    }
    let marking = false;
    function markHard(on) {
      marking = on;
      if (!window.CSS || !CSS.highlights) return UI.toast('This browser can’t highlight sentences');
      CSS.highlights.delete('rd-long'); CSS.highlights.delete('rd-verylong');
      if (!on) return;
      const ss = sentences();
      CSS.highlights.set('rd-long', new Highlight(...ss.filter(s => s.words.length > 25 && s.words.length <= 35).map(s => s.range())));
      CSS.highlights.set('rd-verylong', new Highlight(...ss.filter(s => s.words.length > 35).map(s => s.range())));
    }
    async function readability() {
      const st = stats();
      if (!st.words) return UI.toast('There are no words to measure yet');
      const ease = Math.max(0, Math.min(100, st.ease)), grade = Math.max(0, st.grade);
      const verdict = ease >= 70 ? 'easy to read' : ease >= 50 ? 'fairly difficult' : ease >= 30 ? 'difficult' : 'very difficult';
      const msg = [
        `Flesch Reading Ease: ${ease.toFixed(1)} (${verdict})`,
        `Flesch-Kincaid Grade Level: ${grade.toFixed(1)}`,
        '',
        `Words: ${st.words.toLocaleString()} · Sentences: ${st.sentences.toLocaleString()}`,
        `Words per sentence: ${st.wordsPerSentence.toFixed(1)} · Characters per word: ${st.charsPerWord.toFixed(1)}`,
        `Passive sentences: ${Math.round(st.passive * 100)}%`,
        `Long sentences (over 25 words): ${st.long}${st.veryLong ? `, of which ${st.veryLong} very long (over 35)` : ''}`,
        '',
        'The scores are made for English text.'
      ].join('\n');
      const i = await UI.choose('Readability Statistics', msg, [marking ? 'Stop Highlighting' : 'Highlight Long Sentences', 'Close'], 'book-open-check');
      if (i === 0) markHard(!marking);
    }
    doc.addEventListener('input', () => { if (marking) { clearTimeout(markHard.t); markHard.t = setTimeout(() => markHard(true), 600); } });

    // ================= Translation =================
    const LANGS = ['English', 'Spanish', 'French', 'German', 'Portuguese', 'Italian', 'Dutch', 'Polish', 'Turkish', 'Russian', 'Ukrainian', 'Arabic', 'Hebrew', 'Hindi', 'Bengali', 'Chinese (Simplified)', 'Chinese (Traditional)', 'Japanese', 'Korean', 'Indonesian', 'Vietnamese', 'Thai', 'Swahili', 'Yoruba', 'Igbo', 'Hausa', 'Greek', 'Swedish'];
    async function translate() {
      const sel = getSelection();
      const hasSel = sel.rangeCount && !sel.isCollapsed && doc.contains(sel.anchorNode);
      const v = await UI.form('Translate', [
        { name: 'lang', label: 'Into', type: 'select', options: LANGS, value: (() => { try { return localStorage.getItem('atlas-translate-to') || 'Spanish'; } catch { return 'Spanish'; } })() },
        { name: 'what', label: 'What', type: 'select', options: [...(hasSel ? [{ value: 'sel', label: 'The selected text (replaced here)' }] : []), { value: 'doc', label: 'The whole document (in a new window)' }], value: hasSel ? 'sel' : 'doc' }
      ], 'Translate');
      if (!v) return;
      try { localStorage.setItem('atlas-translate-to', v.lang); } catch { /* private mode */ }
      if (v.what === 'sel') {
        const range = sel.getRangeAt(0).cloneRange();
        const text = range.toString();
        UI.toast('Translating…', 10000);
        let out;
        try { out = await Assistant.translate(text, v.lang); } catch (err) { return UI.toast('Couldn’t translate: ' + err.message, 6000); }
        o.snapshot();
        sel.removeAllRanges(); sel.addRange(range);
        document.execCommand('insertText', false, out);
        o.changed();
        UI.toast('Translated into ' + v.lang);
        return;
      }
      // The whole document, block by block, keeping the formatting.
      const holder = document.createElement('div');
      holder.innerHTML = o.cleanHtml();
      const blocks = [...holder.children].map(n => n.outerHTML);
      const chunks = [];
      let cur = [], size = 0;
      for (const b of blocks) { if (size + b.length > 6000 && cur.length) { chunks.push(cur); cur = []; size = 0; } cur.push(b); size += b.length; }
      if (cur.length) chunks.push(cur);
      const out = [];
      for (let i = 0; i < chunks.length; i++) {
        UI.toast(`Translating… ${Math.round(100 * i / chunks.length)}%`, 30000);
        try { out.push(...await Assistant.translateHtml(chunks[i], v.lang)); } catch (err) { return UI.toast('Couldn’t translate: ' + err.message, 6000); }
      }
      const html = Platform.sanitizeHtml(out.join('\n'));
      await openNew(`${Platform.baseName(o.name())} (${v.lang})`, html, { ...setup(), styles: setup().styles });
      UI.toast('Translated into ' + v.lang + ': opened in a new window');
    }

    return {
      watermarkDialog, decoratePages, toggleDropCap, prepareForDocx, patchDocx, preprocessDocx, postprocessHtml,
      moveSection, shiftLevel, setOutline, drawOutline, envelopes, labels, readability, stats, markHard, translate,
      sentences, LABELS
    };
  }
  WriteTools.cleanWatermark = cleanWatermark;
  WriteTools.LINE_MODES = LINE_MODES;
  WriteTools.MK = MK;
  window.WriteTools = WriteTools;
})();
