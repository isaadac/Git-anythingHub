'use strict';
// Isadek Sheets — sparklines: a tiny line, column or win/loss chart inside one cell.
// On a sheet: sparklines = [{ type: 'line'|'column'|'winloss', src: 'A2:E2' (or 'Sheet1!A2:E2'), cell: 'F2', color }]
// In .xlsx they are Excel 2010's sparkline groups (x14:sparklineGroups in the sheet's extLst).
(function () {
  const TYPES = ['line', 'column', 'winloss'];
  const RANGE_RE = /^(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/;
  const CELL_RE = /^[A-Z]{1,3}[1-9]\d{0,6}$/;
  const NEG = '#c00000';

  function clean(x) {
    if (!x || typeof x !== 'object' || !TYPES.includes(x.type) || typeof x.src !== 'string' || !RANGE_RE.test(x.src) || !CELL_RE.test(x.cell || '')) return null;
    return { type: x.type, src: x.src, cell: x.cell, color: /^#[0-9a-f]{6}$/i.test(x.color || '') ? x.color : '#376092' };
  }

  /** The chart as SVG markup for a w×h cell. values: numbers (non-numbers are gaps). */
  function svg(sp, values, w, h) {
    const nums = values.map(v => (typeof v === 'number' && isFinite(v) ? v : null));
    const real = nums.filter(v => v !== null);
    const pad = 3, W = Math.max(10, w - pad * 2), H = Math.max(6, h - pad * 2);
    if (!real.length) return '';
    const out = [`<svg xmlns="http://www.w3.org/2000/svg" class="spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" aria-hidden="true">`];
    const n = nums.length;
    if (sp.type === 'line') {
      const lo = Math.min(...real), hi = Math.max(...real), span = hi - lo || 1;
      const x = i => pad + (n === 1 ? W / 2 : (i * W) / (n - 1)), y = v => pad + H - ((v - lo) / span) * H;
      let d = '', pen = false;
      nums.forEach((v, i) => { if (v === null) { pen = false; return; } d += `${pen ? 'L' : 'M'}${x(i).toFixed(1)},${y(v).toFixed(1)}`; pen = true; });
      out.push(`<path d="${d}" fill="none" stroke="${sp.color}" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>`);
      // High and low points
      const hiI = nums.indexOf(hi), loI = nums.indexOf(lo);
      out.push(`<circle cx="${x(hiI).toFixed(1)}" cy="${y(hi).toFixed(1)}" r="1.8" fill="${sp.color}"/><circle cx="${x(loI).toFixed(1)}" cy="${y(lo).toFixed(1)}" r="1.8" fill="${NEG}"/>`);
    } else {
      const bw = W / n;
      if (sp.type === 'winloss') {
        nums.forEach((v, i) => { if (v === null || v === 0) return; const top = v > 0 ? pad : pad + H / 2 + 1; out.push(`<rect x="${(pad + i * bw + bw * 0.15).toFixed(1)}" y="${top.toFixed(1)}" width="${(bw * 0.7).toFixed(1)}" height="${(H / 2 - 1).toFixed(1)}" fill="${v > 0 ? sp.color : NEG}"/>`); });
      } else {
        const lo = Math.min(0, ...real), hi = Math.max(0, ...real), span = hi - lo || 1;
        const y = v => pad + H - ((v - lo) / span) * H;
        nums.forEach((v, i) => { if (v === null) return; const a = y(Math.max(0, v)), b = y(Math.min(0, v)); out.push(`<rect x="${(pad + i * bw + bw * 0.15).toFixed(1)}" y="${a.toFixed(1)}" width="${(bw * 0.7).toFixed(1)}" height="${Math.max(0.8, b - a).toFixed(1)}" fill="${v < 0 ? NEG : sp.color}"/>`); });
      }
    }
    out.push('</svg>');
    return out.join('');
  }

  // ---------------- Excel ----------------
  const argb = c => 'FF' + c.slice(1).toUpperCase();
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  /** The sheet's extLst with its sparkline groups (one group per type and colour). */
  function toXlsx(list, sheetName) {
    if (!list || !list.length) return '';
    const groups = new Map();
    for (const sp of list) { const k = sp.type + sp.color; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(sp); }
    const q = /^[A-Za-z_][\w.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
    const body = [...groups.values()].map(g => {
      const t = g[0].type;
      const colors = `<x14:colorSeries rgb="${argb(g[0].color)}"/><x14:colorNegative rgb="FFC00000"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="${argb(g[0].color)}"/>` +
        '<x14:colorFirst rgb="FFD00000"/><x14:colorLast rgb="FFD00000"/><x14:colorHigh rgb="FFD00000"/><x14:colorLow rgb="FFD00000"/>';
      const items = g.map(sp => `<x14:sparkline><xm:f>${esc(sp.src.includes('!') ? sp.src : `${q}!${sp.src}`)}</xm:f><xm:sqref>${sp.cell}</xm:sqref></x14:sparkline>`).join('');
      return `<x14:sparklineGroup displayEmptyCellsAs="gap"${t === 'line' ? '' : ` type="${t === 'winloss' ? 'stacked' : 'column'}"`}${t === 'line' ? ' high="1" low="1"' : ' negative="1"'}>${colors}<x14:sparklines>${items}</x14:sparklines></x14:sparklineGroup>`;
    }).join('');
    return '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
      `<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${body}</x14:sparklineGroups></ext></extLst>`;
  }
  /** Sparklines from a sheet's <x14:sparklineGroups> element. */
  function fromXlsx(xml, sheetName) {
    // The fragment's namespaces were declared on elements around it.
    let frag = String(xml);
    const head = () => frag.slice(0, frag.indexOf('>') + 1);
    if (!/xmlns:x14=/.test(head())) frag = frag.replace('<x14:sparklineGroups', '<x14:sparklineGroups xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"');
    if (!/xmlns:xm=/.test(head())) frag = frag.replace('<x14:sparklineGroups', '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"');
    const d = new DOMParser().parseFromString(frag, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) return [];
    const out = [];
    for (const g of d.getElementsByTagNameNS('*', 'sparklineGroup')) {
      const tp = g.getAttribute('type');
      const type = tp === 'column' ? 'column' : tp === 'stacked' ? 'winloss' : 'line';
      const cs = g.getElementsByTagNameNS('*', 'colorSeries')[0];
      const rgb = cs && cs.getAttribute('rgb');
      const color = rgb && /^[0-9A-F]{8}$/i.test(rgb) ? '#' + rgb.slice(2).toLowerCase() : '#376092';
      for (const s of g.getElementsByTagNameNS('*', 'sparkline')) {
        const f = (s.getElementsByTagNameNS('*', 'f')[0] || {}).textContent || '';
        const cell = ((s.getElementsByTagNameNS('*', 'sqref')[0] || {}).textContent || '').trim();
        // A source on the same sheet is kept without the sheet name.
        const m = /^(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!(.+)$/.exec(f.trim());
        const own = m && (m[1] ? m[1].replace(/''/g, "'") : m[2]) === sheetName;
        const sp = clean({ type, src: (own ? m[3] : f.trim()).replace(/\$/g, ''), cell, color });
        if (sp) out.push(sp);
      }
    }
    return out.slice(0, 10000);
  }

  window.SheetsSpark = { TYPES, clean, svg, toXlsx, fromXlsx };
})();
