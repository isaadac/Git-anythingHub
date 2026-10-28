'use strict';
// Cell formatting in Excel's own terms: the fonts, fills, borders, alignment, number formats
// and protection of styles.xml. Saving gives every styled cell a real Excel style, so Excel,
// LibreOffice and Google Sheets show the colours and fonts; opening reads them back.
(function () {
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const DEFAULT_FONT = { name: 'Calibri', size: 11 };
  // Number formats Excel knows by number.
  const BUILTIN_FMT = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '@': 49 };
  const argb = c => 'FF' + String(c).replace('#', '').toUpperCase();

  /** The parts of a cell style that Excel stores (in a fixed key order, for de-duplication). */
  function normalize(st) {
    st = st || {};
    const out = {};
    if (st.font && st.font !== DEFAULT_FONT.name) out.font = String(st.font);
    if (st.size && +st.size !== DEFAULT_FONT.size) out.size = +st.size;
    if (st.b) out.b = 1;
    if (st.i) out.i = 1;
    if (st.u) out.u = 1;
    if (/^#[0-9a-f]{6}$/i.test(st.color || '')) out.color = st.color.toLowerCase();
    if (/^#[0-9a-f]{6}$/i.test(st.bg || '')) out.bg = st.bg.toLowerCase();
    if (st.border && (st.border.t || st.border.l || st.border.r || st.border.b)) out.border = { t: st.border.t ? 1 : 0, l: st.border.l ? 1 : 0, r: st.border.r ? 1 : 0, b: st.border.b ? 1 : 0 };
    if (['left', 'center', 'right'].includes(st.align)) out.align = st.align;
    if (['top', 'middle', 'bottom'].includes(st.valign)) out.valign = st.valign;
    if (st.wrap) out.wrap = 1;
    if (st.fmt && st.fmt !== 'General') out.fmt = String(st.fmt);
    if (st.unlocked) out.unlocked = 1;
    return out;
  }

  /** A table of distinct styles; index 0 is the plain default. */
  function table() {
    const list = [{}], index = new Map([['{}', 0]]);
    return {
      list,
      add(st) {
        const n = normalize(st);
        const key = JSON.stringify(n);
        if (!index.has(key)) { index.set(key, list.length); list.push(n); }
        return index.get(key);
      }
    };
  }

  /** A complete styles.xml for a list of styles from table(). */
  function toXml(list) {
    const fonts = [], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    const numFmts = [];
    const fontIdx = new Map(), fillIdx = new Map(), borderIdx = new Map(), fmtIdx = new Map();
    const fontXml = s => `<font>${s.b ? '<b/>' : ''}${s.i ? '<i/>' : ''}${s.u ? '<u/>' : ''}<sz val="${s.size || DEFAULT_FONT.size}"/>${s.color ? `<color rgb="${argb(s.color)}"/>` : '<color theme="1"/>'}<name val="${esc(s.font || DEFAULT_FONT.name)}"/><family val="2"/>${!s.font ? '<scheme val="minor"/>' : ''}</font>`;
    const id = (map, arr, key, make) => { if (!map.has(key)) { map.set(key, arr.length); arr.push(make()); } return map.get(key); };
    id(fontIdx, fonts, fontXml({}), () => fontXml({}));
    const xfs = list.map(s => {
      const f = fontXml(s);
      const fontId = id(fontIdx, fonts, f, () => f);
      const fillId = s.bg ? id(fillIdx, fills, s.bg, () => `<fill><patternFill patternType="solid"><fgColor rgb="${argb(s.bg)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
      const bkey = s.border ? JSON.stringify(s.border) : '';
      const side = (n, on) => (on ? `<${n} style="thin"><color indexed="64"/></${n}>` : `<${n}/>`);
      const borderId = s.border ? id(borderIdx, borders, bkey, () => `<border>${side('left', s.border.l)}${side('right', s.border.r)}${side('top', s.border.t)}${side('bottom', s.border.b)}<diagonal/></border>`) : 0;
      let numFmtId = 0;
      if (s.fmt) numFmtId = s.fmt in BUILTIN_FMT ? BUILTIN_FMT[s.fmt] : id(fmtIdx, numFmts, s.fmt, () => s.fmt) + 164;
      const align = s.align || s.valign || s.wrap ? `<alignment${s.align ? ` horizontal="${s.align}"` : ''}${s.valign ? ` vertical="${s.valign === 'middle' ? 'center' : s.valign}"` : ''}${s.wrap ? ' wrapText="1"' : ''}/>` : '';
      const prot = s.unlocked ? '<protection locked="0"/>' : '';
      const attrs = `numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${align ? ' applyAlignment="1"' : ''}${prot ? ' applyProtection="1"' : ''}`;
      return align || prot ? `<xf ${attrs}>${align}${prot}</xf>` : `<xf ${attrs}/>`;
    });
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(f)}"/>`).join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
  }

  // ---------- Reading ----------
  // Excel's legacy palette (indexed colours 0–63).
  const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
  // Office theme colours, in Excel's theme index order (light 1, dark 1, light 2, dark 2, accents).
  const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
  function themeColors(themeXml) {
    if (!themeXml) return OFFICE_THEME;
    try {
      const d = new DOMParser().parseFromString(themeXml, 'application/xml');
      const scheme = d.getElementsByTagNameNS('*', 'clrScheme')[0];
      if (!scheme) return OFFICE_THEME;
      const get = n => { const el = scheme.getElementsByTagNameNS('*', n)[0]; const c = el && el.firstElementChild; return c ? (c.getAttribute('lastClr') || c.getAttribute('val') || '') : ''; };
      // Excel swaps the first two pairs: index 0 is lt1, 1 is dk1, 2 is lt2, 3 is dk2.
      const out = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map((n, i) => (/^[0-9A-F]{6}$/i.test(get(n)) ? get(n).toUpperCase() : OFFICE_THEME[i]));
      return out;
    } catch { return OFFICE_THEME; }
  }
  function tint(hex, t) {
    if (!t) return hex;
    const ch = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16));
    const out = ch.map(v => Math.round(t < 0 ? v * (1 + t) : v + (255 - v) * t));
    return out.map(v => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function colorOf(el, theme) {
    if (!el || el.getAttribute('auto') === '1') return null;
    let hex = null;
    const rgb = el.getAttribute('rgb'), th = el.getAttribute('theme'), ix = el.getAttribute('indexed');
    if (rgb && /^[0-9A-F]{6,8}$/i.test(rgb)) hex = rgb.slice(-6).toUpperCase();
    else if (th !== null && theme[+th]) hex = theme[+th];
    else if (ix !== null) { if (+ix === 64 || +ix === 65) return null; hex = INDEXED[+ix] || null; }
    if (!hex) return null;
    return '#' + tint(hex, parseFloat(el.getAttribute('tint') || '0')).toLowerCase();
  }

  /**
   * The style of each cellXfs entry of a styles.xml, in Sheets' terms (without the number
   * format, which the cell reader already has).
   */
  function fromXml(stylesXml, themeXml) {
    if (!stylesXml) return [];
    const d = new DOMParser().parseFromString(stylesXml, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) return [];
    const theme = themeColors(themeXml);
    const kids = (el, name) => (el ? [...el.children].filter(c => c.localName === name) : []);
    const first = (el, name) => kids(el, name)[0] || null;
    const root = d.documentElement;
    const fonts = kids(first(root, 'fonts'), 'font').map(f => {
      const o = {};
      const v = (n, a = 'val') => { const e = first(f, n); return e ? (e.getAttribute(a) ?? '') : null; };
      const on = n => { const x = v(n); return x !== null && x !== '0' && x !== 'false' && x !== 'none'; };
      if (on('b')) o.b = 1;
      if (on('i')) o.i = 1;
      if (first(f, 'u') && v('u') !== 'none') o.u = 1;
      const sz = parseFloat(v('sz'));
      if (sz) o.size = sz;
      const name = v('name');
      if (name) o.font = name;
      const c = colorOf(first(f, 'color'), theme);
      if (c) o.color = c;
      return o;
    });
    const base = fonts[0] || {};
    const fills = kids(first(root, 'fills'), 'fill').map(f => {
      const p = first(f, 'patternFill');
      if (!p || !p.getAttribute('patternType') || p.getAttribute('patternType') === 'none') return null;
      return colorOf(first(p, 'fgColor'), theme) || colorOf(first(p, 'bgColor'), theme);
    });
    const borders = kids(first(root, 'borders'), 'border').map(b => {
      const has = n => { const e = first(b, n) || first(b, n === 'left' ? 'start' : n === 'right' ? 'end' : n); return !!(e && e.getAttribute('style') && e.getAttribute('style') !== 'none'); };
      const o = { t: has('top') ? 1 : 0, l: has('left') ? 1 : 0, r: has('right') ? 1 : 0, b: has('bottom') ? 1 : 0 };
      return o.t || o.l || o.r || o.b ? o : null;
    });
    return kids(first(root, 'cellXfs'), 'xf').map(x => {
      const st = {};
      const font = fonts[+x.getAttribute('fontId') || 0] || {};
      // Only what differs from the workbook's normal font is a style.
      if (font.font && font.font !== base.font && font.font !== DEFAULT_FONT.name) st.font = font.font;
      if (font.size && font.size !== (base.size || DEFAULT_FONT.size)) st.size = font.size;
      if (font.b) st.b = true;
      if (font.i) st.i = true;
      if (font.u) st.u = true;
      if (font.color && font.color !== base.color && font.color !== '#000000') st.color = font.color;
      const fill = fills[+x.getAttribute('fillId') || 0];
      if (fill) st.bg = fill;
      const border = borders[+x.getAttribute('borderId') || 0];
      if (border) st.border = border;
      const al = first(x, 'alignment');
      if (al) {
        const h = al.getAttribute('horizontal');
        if (h === 'left' || h === 'right') st.align = h;
        if (h === 'center' || h === 'centerContinuous') st.align = 'center';
        const v = al.getAttribute('vertical');
        if (v === 'top') st.valign = 'top';
        if (v === 'center') st.valign = 'middle';
        if (al.getAttribute('wrapText') === '1' || al.getAttribute('wrapText') === 'true') st.wrap = true;
      }
      const pr = first(x, 'protection');
      if (pr && /^(0|false)$/.test(pr.getAttribute('locked') || '')) st.unlocked = true;
      return st;
    });
  }

  window.SheetsStyles = { normalize, table, toXml, fromXml, themeColors, colorOf };
})();
