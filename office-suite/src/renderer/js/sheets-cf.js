'use strict';
// Isadek Sheets — conditional formatting.
// A sheet's rules live in sheet.cf, highest priority first:
//   { ref: 'B2:B20', type, ...options, style: { bg, color, b, i } }
// Types follow Excel's: cellIs, containsText, notContainsText, beginsWith, endsWith, top10,
// aboveAverage, duplicateValues, uniqueValues, expression, colorScale, dataBar, iconSet.
// They are saved as real Excel conditional formats in .xlsx, and read back from Excel files.
(function () {
  const F = window.Formula;
  const REF = /^[A-Z]{1,3}[1-9]\d{0,6}(:[A-Z]{1,3}[1-9]\d{0,6})?$/;
  const OPS = ['greaterThan', 'lessThan', 'between', 'notBetween', 'equal', 'notEqual', 'greaterThanOrEqual', 'lessThanOrEqual'];
  const TEXT_TYPES = ['containsText', 'notContainsText', 'beginsWith', 'endsWith'];
  const STYLE_TYPES = ['cellIs', ...TEXT_TYPES, 'top10', 'aboveAverage', 'duplicateValues', 'uniqueValues', 'expression'];
  const ICON_SETS = {
    '3Arrows': [['▼', '#c00000'], ['►', '#ffc000'], ['▲', '#00b050']],
    '3TrafficLights1': [['●', '#e02020'], ['●', '#f0c000'], ['●', '#20a040']],
    '3Symbols': [['✖', '#e02020'], ['!', '#f0c000'], ['✔', '#20a040']],
    '3Flags': [['⚑', '#e02020'], ['⚑', '#f0c000'], ['⚑', '#20a040']]
  };
  const PRESETS = {
    red: { bg: '#ffc7ce', color: '#9c0006' }, yellow: { bg: '#ffeb9c', color: '#9c5700' }, green: { bg: '#c6efce', color: '#006100' },
    redFill: { bg: '#ffc7ce' }, redText: { color: '#ff0000' }, bold: { b: true }
  };
  const isColor = c => typeof c === 'string' && /^#[0-9a-f]{6}$/i.test(c);
  const area = ref => { const [a, b] = ref.split(':'); const p = F.parseRef(a), q = F.parseRef(b || a); return { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) }; };
  const refOf = g => F.addr(g.r1, g.c1) + (g.r1 === g.r2 && g.c1 === g.c2 ? '' : ':' + F.addr(g.r2, g.c2));
  const str = (v, n = 255) => (typeof v === 'string' ? v.slice(0, n) : typeof v === 'number' && isFinite(v) ? String(v) : '');

  /** A rule from a file or the user, with only known fields and sane values; null if unusable. */
  function clean(r) {
    if (!r || typeof r !== 'object' || typeof r.ref !== 'string' || !REF.test(r.ref)) return null;
    const g = area(r.ref);
    if (g.r2 >= 1048576 || g.c2 >= 16384) return null;
    const out = { ref: r.ref, type: r.type };
    if (STYLE_TYPES.includes(r.type)) {
      const s = r.style && typeof r.style === 'object' ? r.style : {};
      out.style = {};
      if (isColor(s.bg)) out.style.bg = s.bg.toLowerCase();
      if (isColor(s.color)) out.style.color = s.color.toLowerCase();
      if (s.b === true) out.style.b = true;
      if (s.i === true) out.style.i = true;
    }
    switch (r.type) {
      case 'cellIs':
        if (!OPS.includes(r.op)) return null;
        out.op = r.op; out.v1 = str(r.v1); if (/etween/.test(r.op)) out.v2 = str(r.v2);
        break;
      case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith':
        out.text = str(r.text); break;
      case 'top10':
        out.rank = Math.max(1, Math.min(1000, Math.round(+r.rank) || 10)); out.percent = r.percent === true; out.bottom = r.bottom === true; break;
      case 'aboveAverage': out.below = r.below === true; break;
      case 'duplicateValues': case 'uniqueValues': break;
      case 'expression':
        out.formula = str(r.formula, 1000).replace(/^=/, '');
        if (!out.formula) return null;
        break;
      case 'colorScale': {
        const c = Array.isArray(r.colors) ? r.colors.filter(isColor).map(x => x.toLowerCase()) : [];
        if (c.length < 2 || c.length > 3) return null;
        out.colors = c; break;
      }
      case 'dataBar': out.color = isColor(r.color) ? r.color.toLowerCase() : '#638ec6'; break;
      case 'iconSet': out.set = Object.hasOwn(ICON_SETS, r.set) ? r.set : '3TrafficLights1'; out.reverse = r.reverse === true; break;
      default: return null;
    }
    return out;
  }

  // ---------- Evaluation ----------
  function SheetsCF(o) {
    // o: { engine, sheet: () => current sheet }
    let stats = new Map();
    const numeric = v => typeof v === 'number' && isFinite(v);
    const valuesIn = (s, g) => {
      const out = [];
      const n = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
      const keys = Object.keys(s.cells);
      if (n <= keys.length) {
        for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (s.cells[F.addr(r, c)]) out.push(o.engine.value(s.name, r, c));
      } else {
        for (const k of keys) { const p = F.parseRef(k); if (p.row >= g.r1 && p.row <= g.r2 && p.col >= g.c1 && p.col <= g.c2) out.push(o.engine.value(s.name, p.row, p.col)); }
      }
      return out.filter(v => v !== '' && !(v instanceof F.FormulaError));
    };
    const key = v => (typeof v === 'string' ? 's' + v.toLowerCase() : typeof v === 'boolean' ? 'b' + v : 'n' + v);
    function statsFor(s, rule, i) {
      const k = s.name + '\u0000' + i;
      if (stats.has(k)) return stats.get(k);
      const vals = valuesIn(s, area(rule.ref));
      const nums = vals.filter(numeric).sort((a, b) => a - b);
      const st = { nums, min: nums[0], max: nums[nums.length - 1] };
      if (rule.type === 'aboveAverage') st.avg = nums.reduce((a, b) => a + b, 0) / (nums.length || 1);
      if (rule.type === 'top10') {
        const n = rule.percent ? Math.max(1, Math.floor(nums.length * rule.rank / 100)) : rule.rank;
        st.threshold = rule.bottom ? nums[Math.min(nums.length, n) - 1] : nums[Math.max(0, nums.length - n)];
      }
      if (rule.type === 'duplicateValues' || rule.type === 'uniqueValues') { st.count = new Map(); vals.forEach(v => st.count.set(key(v), (st.count.get(key(v)) || 0) + 1)); }
      if (rule.type === 'colorScale' && nums.length) st.mid = nums[Math.floor((nums.length - 1) / 2)] + (nums.length % 2 ? 0 : (nums[nums.length / 2] - nums[nums.length / 2 - 1]) / 2);
      stats.set(k, st);
      return st;
    }
    /** An operand: a number, a formula (relative to the rule's first cell), or text. */
    function operand(s, raw, dr, dc) {
      if (raw === undefined || raw === '') return '';
      if (raw[0] === '=') return o.engine.compute(s.name, F.shiftFormula(raw, dr, dc));
      const n = Number(raw);
      return raw.trim() !== '' && isFinite(n) ? n : raw;
    }
    // Excel orders numbers before text; text compares case-insensitively.
    const cmp = (a, b) => {
      if (typeof a === 'boolean') a = a ? 1 : 0; if (typeof b === 'boolean') b = b ? 1 : 0;
      if (a === '') a = 0; if (b === '') b = 0;
      if (typeof a === 'number' && typeof b === 'number') return a - b;
      if (typeof a === 'number') return -1;
      if (typeof b === 'number') return 1;
      return String(a).toLowerCase().localeCompare(String(b).toLowerCase());
    };
    function matches(s, rule, i, r, c, v) {
      const g = area(rule.ref);
      switch (rule.type) {
        case 'cellIs': {
          const a = operand(s, rule.v1, r - g.r1, c - g.c1), b = operand(s, rule.v2, r - g.r1, c - g.c1);
          if (a instanceof F.FormulaError || b instanceof F.FormulaError) return false;
          const x = cmp(v, a);
          switch (rule.op) {
            case 'greaterThan': return x > 0;
            case 'lessThan': return x < 0;
            case 'greaterThanOrEqual': return x >= 0;
            case 'lessThanOrEqual': return x <= 0;
            case 'equal': return x === 0;
            case 'notEqual': return x !== 0;
            case 'between': case 'notBetween': {
              const lo = cmp(a, b) <= 0 ? a : b, hi = lo === a ? b : a;
              const inside = cmp(v, lo) >= 0 && cmp(v, hi) <= 0;
              return rule.op === 'between' ? inside : !inside;
            }
          }
          return false;
        }
        case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
          const t = String(v === '' ? '' : v).toLowerCase(), q = rule.text.toLowerCase();
          if (rule.type === 'containsText') return t.includes(q);
          if (rule.type === 'notContainsText') return !t.includes(q);
          if (rule.type === 'beginsWith') return t.startsWith(q);
          return t.endsWith(q);
        }
        case 'top10': { if (!numeric(v)) return false; const st = statsFor(s, rule, i); return st.threshold !== undefined && (rule.bottom ? v <= st.threshold : v >= st.threshold); }
        case 'aboveAverage': { if (!numeric(v)) return false; const st = statsFor(s, rule, i); return rule.below ? v < st.avg : v > st.avg; }
        case 'duplicateValues': case 'uniqueValues': {
          if (v === '') return false;
          const n = statsFor(s, rule, i).count.get(key(v)) || 0;
          return rule.type === 'duplicateValues' ? n > 1 : n === 1;
        }
        case 'expression': {
          const x = o.engine.compute(s.name, F.shiftFormula('=' + rule.formula, r - g.r1, c - g.c1));
          return x === true || (typeof x === 'number' && x !== 0);
        }
      }
      return false;
    }
    const mix = (a, b, t) => '#' + [1, 3, 5].map(i => Math.round(parseInt(a.substr(i, 2), 16) * (1 - t) + parseInt(b.substr(i, 2), 16) * t).toString(16).padStart(2, '0')).join('');
    /**
     * How (r, c) on the current sheet is formatted: { bg, color, b, i, bar: { pct, color }, icon: [char, color] }
     * or null. Higher-priority rules win where they set the same property, as in Excel.
     */
    function format(r, c, v) {
      const s = o.sheet();
      if (!s.cf || !s.cf.length || v === undefined) return null;
      let out = null;
      const set = (k, val) => { out = out || {}; if (out[k] === undefined) out[k] = val; };
      s.cf.forEach((rule, i) => {
        const g = area(rule.ref);
        if (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2) return;
        if (rule.style) {
          if (!matches(s, rule, i, r, c, v)) return;
          Object.entries(rule.style).forEach(([k, val]) => set(k, val));
          return;
        }
        if (!numeric(v)) return;
        const st = statsFor(s, rule, i);
        if (st.min === undefined) return;
        const span = st.max - st.min;
        if (rule.type === 'colorScale') {
          const [c1, c2, c3] = rule.colors;
          let col;
          if (!span) col = c3 ? c2 : c1;
          else if (c3) col = v <= st.mid ? mix(c1, c2, st.mid === st.min ? 1 : (v - st.min) / (st.mid - st.min)) : mix(c2, c3, st.max === st.mid ? 1 : (v - st.mid) / (st.max - st.mid));
          else col = mix(c1, c2, (v - st.min) / span);
          set('bg', col);
        } else if (rule.type === 'dataBar') {
          const lo = Math.min(0, st.min), hi = Math.max(0, st.max);
          set('bar', { pct: hi === lo ? 100 : Math.round((v - lo) / (hi - lo) * 100), color: rule.color });
        } else if (rule.type === 'iconSet') {
          const p = span ? (v - st.min) / span : 1;
          let idx = p >= 0.67 ? 2 : p >= 0.33 ? 1 : 0;
          if (rule.reverse) idx = 2 - idx;
          set('icon', ICON_SETS[rule.set][idx]);
        }
      });
      return out;
    }
    return { begin: () => { stats = new Map(); }, format };
  }

  /** Moves and resizes rules when rows/columns are inserted (count > 0) or deleted (count < 0). */
  function shift(rules, axis, index, count, sheetName) {
    return (rules || []).map(rule => {
      const g = area(rule.ref);
      const k1 = axis === 'row' ? 'r1' : 'c1', k2 = axis === 'row' ? 'r2' : 'c2';
      if (count < 0 && g[k1] >= index && g[k2] < index - count) return null;
      const n = { ...g };
      if (count < 0) {
        const lo = index, hi = index - count - 1;
        n[k1] = g[k1] > hi ? g[k1] + count : Math.min(g[k1], lo);
        n[k2] = g[k2] > hi ? g[k2] + count : g[k2] >= lo ? lo - 1 : g[k2];
        if (n[k2] < n[k1]) return null;
      } else {
        if (g[k1] >= index) n[k1] += count;
        if (g[k2] >= index) n[k2] += count;
      }
      const out = { ...rule, ref: refOf(n) };
      // rewriteRefs works on formulas written with their leading '='.
      const adj = f => (f && f[0] === '=' ? F.adjustForInsert(f, axis, index, count, sheetName, sheetName) : f);
      if (out.formula) out.formula = adj('=' + out.formula).slice(1);
      if (out.v1) out.v1 = adj(out.v1);
      if (out.v2) out.v2 = adj(out.v2);
      return out;
    }).filter(Boolean);
  }

  /** Words for a rule, for the Manage Rules list. */
  function describe(r) {
    const opWords = { greaterThan: '>', lessThan: '<', greaterThanOrEqual: '≥', lessThanOrEqual: '≤', equal: '=', notEqual: '≠', between: 'between', notBetween: 'not between' };
    switch (r.type) {
      case 'cellIs': return `Cell value ${opWords[r.op]} ${r.v1}${/etween/.test(r.op) ? ' and ' + r.v2 : ''}`;
      case 'containsText': return `Text contains “${r.text}”`;
      case 'notContainsText': return `Text does not contain “${r.text}”`;
      case 'beginsWith': return `Text begins with “${r.text}”`;
      case 'endsWith': return `Text ends with “${r.text}”`;
      case 'top10': return `${r.bottom ? 'Bottom' : 'Top'} ${r.rank}${r.percent ? '%' : ''}`;
      case 'aboveAverage': return r.below ? 'Below average' : 'Above average';
      case 'duplicateValues': return 'Duplicate values';
      case 'uniqueValues': return 'Unique values';
      case 'expression': return `Formula: =${r.formula}`;
      case 'colorScale': return `${r.colors.length}-colour scale`;
      case 'dataBar': return 'Data bar';
      case 'iconSet': return 'Icon set ' + ICON_SETS[r.set].map(x => x[0]).join(' ');
    }
    return r.type;
  }

  // ---------- Excel ----------
  const argb = c => 'FF' + c.slice(1).toUpperCase();
  const xmlEsc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const quote = t => '"' + String(t).replace(/"/g, '""') + '"';
  const formulaOf = raw => (raw[0] === '=' ? raw.slice(1) : raw.trim() !== '' && isFinite(Number(raw)) ? String(Number(raw)) : quote(raw));

  /** The <conditionalFormatting> XML for a sheet's rules; dxfs collects their differential styles. */
  function toXlsx(rules, dxfs) {
    let prio = 0;
    return (rules || []).map(r => {
      const g = area(r.ref), first = F.addr(g.r1, g.c1);
      const p = ++prio;
      let dxf = '';
      if (r.style) {
        const s = r.style;
        const font = (s.b ? '<b/>' : '') + (s.i ? '<i/>' : '') + (s.color ? `<color rgb="${argb(s.color)}"/>` : '');
        dxfs.push('<dxf>' + (font ? `<font>${font}</font>` : '') + (s.bg ? `<fill><patternFill><bgColor rgb="${argb(s.bg)}"/></patternFill></fill>` : '') + '</dxf>');
        dxf = ` dxfId="${dxfs.length - 1}"`;
      }
      const f = x => `<formula>${xmlEsc(x)}</formula>`;
      let rule;
      switch (r.type) {
        case 'cellIs': rule = `<cfRule type="cellIs"${dxf} priority="${p}" operator="${r.op}">${f(formulaOf(r.v1))}${/etween/.test(r.op) ? f(formulaOf(r.v2 || '')) : ''}</cfRule>`; break;
        case 'containsText': rule = `<cfRule type="containsText"${dxf} priority="${p}" operator="containsText" text="${xmlEsc(r.text)}">${f(`NOT(ISERROR(SEARCH(${quote(r.text)},${first})))`)}</cfRule>`; break;
        case 'notContainsText': rule = `<cfRule type="notContainsText"${dxf} priority="${p}" operator="notContains" text="${xmlEsc(r.text)}">${f(`ISERROR(SEARCH(${quote(r.text)},${first}))`)}</cfRule>`; break;
        case 'beginsWith': rule = `<cfRule type="beginsWith"${dxf} priority="${p}" operator="beginsWith" text="${xmlEsc(r.text)}">${f(`LEFT(${first},LEN(${quote(r.text)}))=${quote(r.text)}`)}</cfRule>`; break;
        case 'endsWith': rule = `<cfRule type="endsWith"${dxf} priority="${p}" operator="endsWith" text="${xmlEsc(r.text)}">${f(`RIGHT(${first},LEN(${quote(r.text)}))=${quote(r.text)}`)}</cfRule>`; break;
        case 'top10': rule = `<cfRule type="top10"${dxf} priority="${p}" rank="${r.rank}"${r.percent ? ' percent="1"' : ''}${r.bottom ? ' bottom="1"' : ''}/>`; break;
        case 'aboveAverage': rule = `<cfRule type="aboveAverage"${dxf} priority="${p}"${r.below ? ' aboveAverage="0"' : ''}/>`; break;
        case 'duplicateValues': case 'uniqueValues': rule = `<cfRule type="${r.type}"${dxf} priority="${p}"/>`; break;
        case 'expression': rule = `<cfRule type="expression"${dxf} priority="${p}">${f(r.formula)}</cfRule>`; break;
        case 'colorScale': rule = `<cfRule type="colorScale" priority="${p}"><colorScale><cfvo type="min"/>${r.colors.length === 3 ? '<cfvo type="percentile" val="50"/>' : ''}<cfvo type="max"/>${r.colors.map(c => `<color rgb="${argb(c)}"/>`).join('')}</colorScale></cfRule>`; break;
        case 'dataBar': rule = `<cfRule type="dataBar" priority="${p}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(r.color)}"/></dataBar></cfRule>`; break;
        case 'iconSet': rule = `<cfRule type="iconSet" priority="${p}"><iconSet iconSet="${r.set}"${r.reverse ? ' reverse="1"' : ''}><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>`; break;
        default: return '';
      }
      return `<conditionalFormatting sqref="${r.ref}">${rule}</conditionalFormatting>`;
    }).join('');
  }

  /** Rules from an Excel sheet's <conditionalFormatting> XML and the workbook's <dxfs>. */
  function fromXlsx(cfXml, dxfsXml) {
    if (!cfXml) return [];
    const parse = x => new DOMParser().parseFromString(x, 'application/xml');
    const kids = (el, name) => [...el.children].filter(e => e.localName === name);
    const colorOf = el => { const v = el && el.getAttribute('rgb'); return v && /^[0-9a-f]{8}$/i.test(v) ? '#' + v.slice(2).toLowerCase() : v && /^[0-9a-f]{6}$/i.test(v) ? '#' + v.toLowerCase() : null; };
    const dxfs = [];
    if (dxfsXml) {
      const d = parse(dxfsXml).documentElement;
      if (d && d.localName === 'dxfs') kids(d, 'dxf').forEach(x => {
        const st = {};
        const font = kids(x, 'font')[0];
        if (font) { if (kids(font, 'b')[0] && kids(font, 'b')[0].getAttribute('val') !== '0') st.b = true; if (kids(font, 'i')[0] && kids(font, 'i')[0].getAttribute('val') !== '0') st.i = true; const c = colorOf(kids(font, 'color')[0]); if (c) st.color = c; }
        const pf = kids(x, 'fill')[0] && kids(kids(x, 'fill')[0], 'patternFill')[0];
        if (pf) { const c = colorOf(kids(pf, 'bgColor')[0]) || colorOf(kids(pf, 'fgColor')[0]); if (c) st.bg = c; }
        dxfs.push(st);
      });
    }
    const doc = parse('<r xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + cfXml + '</r>').documentElement;
    if (!doc || doc.querySelector('parsererror')) return [];
    const unquote = t => (/^".*"$/.test(t) ? t.slice(1, -1).replace(/""/g, '"') : /^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(t) ? t : '=' + t);
    const rules = [];
    kids(doc, 'conditionalFormatting').forEach(cf => {
      const refs = (cf.getAttribute('sqref') || '').split(/\s+/).filter(x => REF.test(x)).slice(0, 20);
      kids(cf, 'cfRule').sort((a, b) => (+a.getAttribute('priority') || 0) - (+b.getAttribute('priority') || 0)).forEach(el => {
        const type = el.getAttribute('type'), forms = kids(el, 'formula').map(x => x.textContent);
        const base = { type, style: dxfs[+el.getAttribute('dxfId')] || {}, priority: +el.getAttribute('priority') || 0 };
        let r = null;
        if (type === 'cellIs') r = { ...base, op: el.getAttribute('operator'), v1: unquote(forms[0] || ''), v2: unquote(forms[1] || '') };
        else if (TEXT_TYPES.includes(type)) r = { ...base, text: el.getAttribute('text') || '' };
        else if (type === 'top10') r = { ...base, rank: +el.getAttribute('rank') || 10, percent: el.getAttribute('percent') === '1', bottom: el.getAttribute('bottom') === '1' };
        else if (type === 'aboveAverage') r = { ...base, below: el.getAttribute('aboveAverage') === '0' };
        else if (type === 'duplicateValues' || type === 'uniqueValues') r = base;
        else if (type === 'expression') r = { ...base, formula: forms[0] || '' };
        else if (type === 'colorScale') { const cs = kids(el, 'colorScale')[0]; r = cs && { type, colors: kids(cs, 'color').map(colorOf).filter(Boolean), priority: base.priority }; }
        else if (type === 'dataBar') { const db = kids(el, 'dataBar')[0]; r = db && { type, color: colorOf(kids(db, 'color')[0]) || '#638ec6', priority: base.priority }; }
        else if (type === 'iconSet') { const is = kids(el, 'iconSet')[0]; r = { type, set: (is && is.getAttribute('iconSet')) || '3TrafficLights1', reverse: !!is && is.getAttribute('reverse') === '1', priority: base.priority }; }
        if (r) refs.forEach(ref => rules.push({ ...r, ref }));
      });
    });
    return rules.sort((a, b) => a.priority - b.priority).map(clean).filter(Boolean).slice(0, 500);
  }

  window.SheetsCF = Object.assign(SheetsCF, { clean, shift, describe, toXlsx, fromXlsx, area, refOf, PRESETS, ICON_SETS });
})();
