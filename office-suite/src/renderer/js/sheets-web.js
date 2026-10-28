'use strict';
// Isadek Sheets — Get Data from the Web: a CSV/TSV file, JSON (a list of records or rows) or
// the tables on a web page, as rows of values. The data is only parsed (as text, JSON or inert
// HTML), never run.
(function () {
  const MAX_ROWS = 100000, MAX_COLS = 500;
  const num = s => { const t = String(s).trim(); if (/^-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(t)) return parseFloat(t.replace(/,/g, '')); return t; };

  function delimited(text, delim) {
    const rows = [];
    let row = [], cur = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
      if (ch === '"' && cur === '') { q = true; continue; }
      if (ch === delim) { row.push(cur); cur = ''; continue; }
      if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; if (rows.length >= MAX_ROWS) break; continue; }
      cur += ch;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows.map(r => r.slice(0, MAX_COLS).map(num));
  }

  function fromJson(data) {
    // { results: [...] } and the like: the first list inside.
    if (!Array.isArray(data) && data && typeof data === 'object') {
      const list = Object.values(data).find(Array.isArray);
      data = list || [data];
    }
    if (!Array.isArray(data)) return [[String(data)]];
    data = data.slice(0, MAX_ROWS);
    if (data.every(Array.isArray)) return data.map(r => r.slice(0, MAX_COLS).map(cell));
    const keys = [];
    for (const o of data) if (o && typeof o === 'object') for (const k of Object.keys(o)) if (!keys.includes(k) && keys.length < MAX_COLS) keys.push(k);
    if (!keys.length) return data.map(v => [cell(v)]);
    return [keys, ...data.map(o => keys.map(k => cell(o && typeof o === 'object' ? o[k] : '')))];
  }
  const cell = v => (v === null || v === undefined ? '' : typeof v === 'object' ? JSON.stringify(v) : v);

  function htmlTables(text) {
    const d = new DOMParser().parseFromString(text, 'text/html'); // inert: nothing runs or loads
    return [...d.querySelectorAll('table')].map((t, i) => {
      const rows = [...t.rows].slice(0, MAX_ROWS).map(tr => {
        const out = [];
        for (const td of tr.cells) { out.push(num(td.textContent.replace(/\s+/g, ' ').trim())); for (let k = 1; k < Math.min(td.colSpan || 1, 50); k++) out.push(''); }
        return out.slice(0, MAX_COLS);
      }).filter(r => r.some(v => v !== ''));
      const cap = t.caption ? t.caption.textContent.trim() : '';
      return { name: cap || `Table ${i + 1}`, rows };
    }).filter(t => t.rows.length);
  }

  /** The data in a download: { tables: [{ name, rows }] } (a CSV or JSON file is one table). */
  function parse(bytes, contentType, url) {
    const text = new TextDecoder().decode(bytes).replace(/^﻿/, '');
    const ct = String(contentType || '').toLowerCase();
    const ext = (String(url || '').split(/[?#]/)[0].match(/\.(\w+)$/) || [])[1] || '';
    const head = text.trimStart().slice(0, 200).toLowerCase();
    if (ct.includes('json') || ext === 'json' || /^[[{]/.test(head)) {
      try { return { tables: [{ name: 'Data', rows: fromJson(JSON.parse(text)) }] }; } catch { /* not JSON after all */ }
    }
    if (ct.includes('html') || /^<(!doctype html|html|table|body|head)/.test(head) || ext === 'html' || ext === 'htm') {
      const tables = htmlTables(text);
      if (!tables.length) throw new Error('The page has no tables');
      return { tables };
    }
    const first = text.slice(0, text.search(/\r?\n|$/));
    const delim = ct.includes('tab-separated') || ext === 'tsv' || (first.includes('\t') && !first.includes(',')) ? '\t' : first.includes(';') && !first.includes(',') ? ';' : ',';
    return { tables: [{ name: 'Data', rows: delimited(text, delim) }] };
  }

  /** Downloads a URL: through the desktop app (no browser limits), else with fetch (the site must allow it). */
  async function download(url) {
    let u;
    try { u = new URL(url); } catch { throw new Error('Enter a web address such as https://example.com/data.csv'); }
    if (!/^https?:$/.test(u.protocol)) throw new Error('Only http and https addresses can be used');
    if (window.atlas && window.atlas.fetchData) {
      const r = await window.atlas.fetchData(u.toString());
      if (r.error) throw new Error(r.error);
      return { bytes: new Uint8Array(r.bytes), contentType: r.contentType };
    }
    let res;
    try { res = await fetch(u.toString(), { credentials: 'omit', cache: 'no-store' }); } catch { throw new Error('That site doesn’t let web pages read its data (or is unreachable). The desktop app can read it.'); }
    if (!res.ok) throw new Error(`The site answered ${res.status}`);
    const len = +res.headers.get('content-length') || 0;
    if (len > 50e6) throw new Error('That file is too large (over 50 MB)');
    return { bytes: new Uint8Array(await res.arrayBuffer()), contentType: res.headers.get('content-type') || '' };
  }

  window.SheetsWeb = { parse, download, delimited, fromJson, htmlTables };
})();
