'use strict';
// Atlas Sheets — data validation.
// A sheet's rules live in sheet.dv:
//   { ref, type: 'list'|'whole'|'decimal'|'date'|'time'|'textLength'|'custom', op, v1, v2,
//     allowBlank, dropdown, prompt: { title, text }, error: { style: 'stop'|'warning'|'information', title, text } }
// v1/v2 are numbers, text, or formulas starting with '=' (relative to the rule's first cell).
// A list's v1 is either items separated by commas or a range formula such as =$E$1:$E$9.
(function () {
  const F = window.Formula;
  const REF = /^[A-Z]{1,3}[1-9]\d{0,6}(:[A-Z]{1,3}[1-9]\d{0,6})?$/;
  const TYPES = ['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'];
  const OPS = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];
  const STYLES = ['stop', 'warning', 'information'];
  const str = (v, n) => (typeof v === 'string' ? v.slice(0, n) : typeof v === 'number' && isFinite(v) ? String(v) : '');
  const area = ref => window.SheetsCF.area(ref);

  function clean(r) {
    if (!r || typeof r !== 'object' || typeof r.ref !== 'string' || !REF.test(r.ref) || !TYPES.includes(r.type)) return null;
    const out = { ref: r.ref, type: r.type, allowBlank: r.allowBlank !== false };
    if (r.type === 'list') { out.v1 = str(r.v1, 2000); out.dropdown = r.dropdown !== false; if (!out.v1) return null; }
    else if (r.type === 'custom') { out.v1 = str(r.v1, 1000); if (!/^=./.test(out.v1)) out.v1 = out.v1 ? '=' + out.v1 : ''; if (!out.v1) return null; }
    else {
      out.op = OPS.includes(r.op) ? r.op : 'between';
      out.v1 = str(r.v1, 255);
      if (/etween/.test(out.op)) out.v2 = str(r.v2, 255);
      if (out.v1 === '' || (/etween/.test(out.op) && out.v2 === '')) return null;
    }
    const p = r.prompt && typeof r.prompt === 'object' ? r.prompt : null;
    if (p && (p.title || p.text)) out.prompt = { title: str(p.title, 32), text: str(p.text, 255) };
    const e = r.error && typeof r.error === 'object' ? r.error : {};
    out.error = { style: STYLES.includes(e.style) ? e.style : 'stop', title: str(e.title, 32), text: str(e.text, 255) };
    if (e.show === false) out.error.show = false; // Excel's "Show error alert" turned off: any value is accepted
    return out;
  }

  /**
   * o: { engine, sheet: () => sheet, sheetByName }. check(rule, r, c, value) → true when the
   * value is allowed. listItems(rule, r, c) → the dropdown's items.
   */
  function SheetsDV(o) {
    const at = (s, raw, dr, dc) => {
      if (raw === undefined || raw === '') return '';
      if (raw[0] === '=') return o.engine.compute(s.name, F.shiftFormula(raw, dr, dc));
      const t = raw.trim();
      // Dates typed as text (2024-05-01, 5/1/2024) are compared as dates.
      const d = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t) || /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
      if (d) return F.dateToSerial(t.includes('-') ? new Date(+d[1], d[2] - 1, +d[3]) : new Date(+d[3], d[1] - 1, +d[2]));
      const tm = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
      if (tm) return (+tm[1] * 3600 + +tm[2] * 60 + +(tm[3] || 0)) / 86400;
      const n = F.parseLiteral(t);
      return n;
    };
    function listItems(rule, r, c) {
      const s = o.sheet();
      if (rule.v1[0] !== '=') return rule.v1.split(',').map(x => x.trim()).filter(Boolean).slice(0, 1000);
      // A range (possibly on another sheet): its values, in order.
      const m = /^=\s*(?:((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?\s*$/i.exec(rule.v1);
      if (!m) return [];
      const src = m[1] ? o.sheetByName(m[1].replace(/^'|'$/g, '').replace(/''/g, "'")) : s;
      if (!src) return [];
      const g = area((m[2] + m[3] + (m[4] ? ':' + m[4] + m[5] : '')).toUpperCase());
      const out = [];
      for (let rr = g.r1; rr <= Math.min(g.r2, g.r1 + 999); rr++) for (let cc = g.c1; cc <= Math.min(g.c2, g.c1 + 49); cc++) {
        const v = o.engine.value(src.name, rr, cc);
        if (v !== '' && !(v instanceof F.FormulaError)) out.push(typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
      }
      return out;
    }
    const compare = (op, x, a, b) => {
      switch (op) {
        case 'between': return x >= Math.min(a, b) && x <= Math.max(a, b);
        case 'notBetween': return x < Math.min(a, b) || x > Math.max(a, b);
        case 'equal': return x === a;
        case 'notEqual': return x !== a;
        case 'greaterThan': return x > a;
        case 'lessThan': return x < a;
        case 'greaterThanOrEqual': return x >= a;
        case 'lessThanOrEqual': return x <= a;
      }
      return true;
    };
    function check(rule, r, c, v) {
      if (v === '' || v === null || v === undefined) return rule.allowBlank;
      const s = o.sheet(), g = area(rule.ref), dr = r - g.r1, dc = c - g.c1;
      if (v instanceof F.FormulaError) return false;
      switch (rule.type) {
        case 'list': {
          const t = String(typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : v).toLowerCase();
          return listItems(rule, r, c).some(x => x.toLowerCase() === t);
        }
        case 'custom': { const x = o.engine.compute(s.name, F.shiftFormula(rule.v1, dr, dc)); return x === true || (typeof x === 'number' && x !== 0); }
        case 'textLength': {
          const n = String(typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : v).length;
          const a = at(s, rule.v1, dr, dc), b = at(s, rule.v2, dr, dc);
          return typeof a === 'number' && compare(rule.op, n, a, typeof b === 'number' ? b : a);
        }
        default: {
          if (typeof v !== 'number') return false;
          if (rule.type === 'whole' && !Number.isInteger(v)) return false;
          if (rule.type === 'time' && (v < 0 || v >= 1)) return false;
          const a = at(s, rule.v1, dr, dc), b = at(s, rule.v2, dr, dc);
          if (typeof a !== 'number' || (/etween/.test(rule.op) && typeof b !== 'number')) return false;
          return compare(rule.op, v, a, b);
        }
      }
    }
    /** The rule that applies to (r, c) on the current sheet (the first that covers it). */
    function ruleAt(r, c) {
      return (o.sheet().dv || []).find(rule => { const g = area(rule.ref); return r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2; }) || null;
    }
    return { check, listItems, ruleAt };
  }

  const OP_WORDS = { between: 'between', notBetween: 'not between', equal: 'equal to', notEqual: 'not equal to', greaterThan: 'greater than', lessThan: 'less than', greaterThanOrEqual: 'greater than or equal to', lessThanOrEqual: 'less than or equal to' };
  const TYPE_WORDS = { whole: 'a whole number', decimal: 'a number', date: 'a date', time: 'a time', textLength: 'text with a length' };
  /** The message shown when a value is refused and the rule has no message of its own. */
  function defaultError(rule) {
    if (rule.type === 'list') return 'The value doesn\'t match the list of allowed values.';
    if (rule.type === 'custom') return 'The value doesn\'t meet the rule for this cell.';
    return `The value must be ${TYPE_WORDS[rule.type]} ${OP_WORDS[rule.op]} ${rule.v1}${/etween/.test(rule.op) ? ' and ' + rule.v2 : ''}.`;
  }

  // ---------- Excel ----------
  const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const dateRaw = t => { const d = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t) || /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t); return d ? String(F.dateToSerial(t.includes('-') ? new Date(+d[1], d[2] - 1, +d[3]) : new Date(+d[3], d[1] - 1, +d[2]))) : null; };
  function formula(raw, type) {
    if (raw[0] === '=') return raw.slice(1);
    const t = raw.trim();
    if (type === 'date' && dateRaw(t)) return dateRaw(t);
    if (type === 'time') { const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t); if (m) return String((+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400); }
    return isFinite(Number(t)) && t !== '' ? String(Number(t)) : '"' + t.replace(/"/g, '""') + '"';
  }
  function toXlsx(rules) {
    if (!rules || !rules.length) return '';
    const items = rules.map(r => {
      const a = [`type="${r.type}"`];
      if (r.op && r.type !== 'list' && r.type !== 'custom' && r.op !== 'between') a.push(`operator="${r.op}"`);
      if (r.allowBlank) a.push('allowBlank="1"');
      if (r.type === 'list' && !r.dropdown) a.push('showDropDown="1"'); // Excel's flag means "hide the arrow"
      a.push('showInputMessage="1"');
      if (r.error.show !== false) a.push('showErrorMessage="1"');
      if (r.error.style !== 'stop') a.push(`errorStyle="${r.error.style}"`);
      if (r.error.title) a.push(`errorTitle="${esc(r.error.title)}"`);
      if (r.error.text) a.push(`error="${esc(r.error.text)}"`);
      if (r.prompt && r.prompt.title) a.push(`promptTitle="${esc(r.prompt.title)}"`);
      if (r.prompt && r.prompt.text) a.push(`prompt="${esc(r.prompt.text)}"`);
      a.push(`sqref="${r.ref}"`);
      let f1, f2 = '';
      if (r.type === 'list') f1 = r.v1[0] === '=' ? r.v1.slice(1) : '"' + r.v1.split(',').map(x => x.trim()).join(',').replace(/"/g, '""') + '"';
      else if (r.type === 'custom') f1 = r.v1.slice(1);
      else { f1 = formula(r.v1, r.type); if (/etween/.test(r.op)) f2 = formula(r.v2, r.type); }
      return `<dataValidation ${a.join(' ')}><formula1>${esc(f1)}</formula1>${f2 ? `<formula2>${esc(f2)}</formula2>` : ''}</dataValidation>`;
    });
    return `<dataValidations count="${items.length}">${items.join('')}</dataValidations>`;
  }
  /** Rules from Excel's <dataValidations> and the newer <x14:dataValidations> (lists on other sheets). */
  function fromXlsx(xml) {
    if (!xml) return [];
    const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';
    const doc = new DOMParser().parseFromString(`<r ${ns}>${xml}</r>`, 'application/xml');
    if (doc.querySelector('parsererror')) return [];
    const unq = t => (/^".*"$/.test(t) ? t.slice(1, -1).replace(/""/g, '"') : null);
    const fromSerial = (t, type) => {
      if (type === 'date' && /^\d+(\.\d+)?$/.test(t)) { const d = F.serialToDate(+t); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
      if (type === 'time' && /^0?\.\d+$/.test(t)) { const s = Math.round(+t * 86400); return `${Math.floor(s / 3600)}:${String(Math.floor(s / 60) % 60).padStart(2, '0')}`; }
      return t;
    };
    const val = (t, type) => (t === undefined || t === null ? '' : unq(t) !== null ? unq(t) : /^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(t) ? fromSerial(t, type) : '=' + t);
    const on = a => a === '1' || a === 'true'; // both flags are off unless the file says otherwise
    const rules = [];
    [...doc.documentElement.getElementsByTagName('*')].filter(e => e.localName === 'dataValidation').slice(0, 1000).forEach(e => {
      const kid = name => [...e.children].find(k => k.localName === name);
      const fText = name => { const k = kid(name); if (!k) return undefined; const f = [...k.children].find(x => x.localName === 'f'); return (f || k).textContent; };
      const type = e.getAttribute('type') || 'none';
      if (!TYPES.includes(type)) return;
      const sq = (e.getAttribute('sqref') || (kid('sqref') && kid('sqref').textContent) || '').split(/\s+/).filter(x => REF.test(x)).slice(0, 20);
      const f1 = fText('formula1'), f2 = fText('formula2');
      const base = {
        type, op: e.getAttribute('operator') || 'between', allowBlank: e.getAttribute('allowBlank') === '1' || e.getAttribute('allowBlank') === 'true',
        dropdown: !(e.getAttribute('showDropDown') === '1' || e.getAttribute('showDropDown') === 'true'),
        prompt: on(e.getAttribute('showInputMessage')) ? { title: e.getAttribute('promptTitle') || '', text: e.getAttribute('prompt') || '' } : null,
        error: { style: e.getAttribute('errorStyle') || 'stop', title: e.getAttribute('errorTitle') || '', text: e.getAttribute('error') || '', show: on(e.getAttribute('showErrorMessage')) }
      };
      if (type === 'list') base.v1 = f1 && unq(f1) !== null ? unq(f1) : f1 ? '=' + f1 : '';
      else if (type === 'custom') base.v1 = f1 ? '=' + f1 : '';
      else { base.v1 = val(f1, type); base.v2 = val(f2, type); }
      sq.forEach(ref => rules.push({ ...base, ref }));
    });
    return rules.map(clean).filter(Boolean).slice(0, 500);
  }

  window.SheetsDV = Object.assign(SheetsDV, { clean, toXlsx, fromXlsx, defaultError, TYPES, OPS, OP_WORDS });
})();
