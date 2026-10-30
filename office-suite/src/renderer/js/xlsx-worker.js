'use strict';
// Parses spreadsheet files away from the app. Whatever a malicious file does to SheetJS here
// (prototype pollution, catastrophic regex backtracking) stays in this worker, which the app
// throws away after every file and kills if it runs too long.
importScripts('../../../node_modules/xlsx/dist/xlsx.full.min.js');

self.onmessage = e => {
  try {
    const wb = XLSX.read(e.data, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true, bookFiles: true });
    const sheets = []; // an array, so a sheet called __proto__ can't do anything odd
    // Workbooks in the 1904 date system (old Mac Excel): dates become the usual 1900 serials, as
    // Excel does when the option is turned off, so dates and date functions agree.
    const d1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
    const hasDate = z => typeof z === 'string' && /[dy]/i.test(z.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, ''));
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const out = {};
      for (const k of Object.keys(ws)) {
        const c = ws[k];
        if (k[0] === '!') {
          if (k === '!cols' || k === '!rows' || k === '!merges' || k === '!ref') out[k] = JSON.parse(JSON.stringify(c));
          continue;
        }
        // Plain values only: no prototypes, functions or getters cross back to the app.
        if (d1904 && c.t === 'n' && typeof c.v === 'number' && !c.f && hasDate(c.z)) c.v += 1462;
        out[k] = { t: c.t, v: c.v instanceof Date ? c.v.getTime() : c.v, f: c.f, F: typeof c.F === 'string' ? c.F : undefined, D: !!c.D, z: c.z, w: c.w, isDate: c.v instanceof Date,
          c: Array.isArray(c.c) ? c.c.slice(0, 20).map(x => ({ a: String(x.a || '').slice(0, 100), t: String(x.t || '').slice(0, 32767) })) : undefined };
      }
      sheets.push(out);
    }
    const extra = excelExtras(wb);
    const custom = wb.Custprops && typeof wb.Custprops.AtlasMeta === 'string' ? wb.Custprops.AtlasMeta : null;
    // Workbook names (Name Manager), as plain values.
    const names = (wb.Workbook && Array.isArray(wb.Workbook.Names) ? wb.Workbook.Names : []).slice(0, 5000)
      .map(n => ({ Name: String(n.Name || ''), Ref: String(n.Ref || ''), Sheet: Number.isInteger(n.Sheet) ? n.Sheet : null, Hidden: !!n.Hidden }));
    self.postMessage({ ok: true, SheetNames: wb.SheetNames.map(String), Sheets: sheets, AtlasMeta: custom, Extras: extra, Names: names });
  } catch (err) {
    self.postMessage({ ok: false, error: String(err && err.message || err) });
  }
};

/**
 * Raw XML that SheetJS doesn't read, for the app to parse as XML (no code): each sheet's
 * <conditionalFormatting> and <dataValidations> elements and the workbook's differential styles (<dxfs>).
 */
function excelExtras(wb) {
  const out = { sheets: [], dxfs: '' };
  const files = wb.files;
  if (!files || !wb.keys) return out;
  const text = p => {
    const f = files[p] || files['/' + p];
    if (!f || !f.content) return '';
    return typeof f.content === 'string' ? f.content : new TextDecoder().decode(f.content);
  };
  const LIMIT = 2e6;
  const between = (xml, open, close, all) => {
    const parts = [];
    let i = 0, total = 0;
    while ((i = xml.indexOf(open, i)) !== -1) {
      const j = xml.indexOf(close, i);
      if (j === -1) break;
      const part = xml.slice(i, j + close.length);
      total += part.length;
      if (total > LIMIT) break;
      parts.push(part);
      i = j + close.length;
      if (!all) break;
    }
    return parts.join('');
  };
  const rels = text('xl/_rels/workbook.xml.rels');
  const target = Object.create(null);
  for (const m of rels.match(/<Relationship\b[^>]{0,2000}>/g) || []) {
    const id = (m.match(/\bId="([^"]{1,200})"/) || [])[1], t = (m.match(/\bTarget="([^"]{1,500})"/) || [])[1];
    if (id && t) target[id] = t[0] === '/' ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
  }
  // Cell styles marked <protection locked="0"/>, and the cells that use them.
  const xfBlock = (text('xl/styles.xml').match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/) || ['', ''])[1];
  const unlockedXfs = new Set();
  (xfBlock.match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) || []).forEach((x, i) => { if (/<protection\b[^>]*\blocked="(0|false)"/.test(x)) unlockedXfs.add(String(i)); });
  const unlockedIn = xml => {
    if (!unlockedXfs.size) return [];
    const out = [];
    for (const m of xml.matchAll(/<c r="([A-Z]{1,3}\d{1,7})"[^>]{0,300}?\ss="(\d+)"/g)) { if (unlockedXfs.has(m[2])) out.push(m[1]); if (out.length >= 100000) break; }
    return out;
  };
  /** Which style each cell uses (cells with the default style are left out). */
  function styledCells(xml) {
    const map = Object.create(null);
    let n = 0;
    for (const m of xml.matchAll(/<c r="([A-Z]{1,3}\d{1,7})"[^>]{0,300}?\ss="(\d{1,6})"/g)) {
      if (m[2] !== '0') { map[m[1]] = +m[2]; if (++n >= 500000) break; }
    }
    return map;
  }
  /** Frozen panes: rows and columns kept in view. */
  function frozen(xml) {
    const p = /<pane\b[^>]{0,500}>/.exec(xml);
    if (!p || !/state="frozen(Split)?"/.test(p[0])) return null;
    return { rows: +((p[0].match(/ySplit="(\d+(?:\.\d+)?)"/) || [])[1] || 0), cols: +((p[0].match(/xSplit="(\d+(?:\.\d+)?)"/) || [])[1] || 0) };
  }
  const sheetTags = text('xl/workbook.xml').match(/<sheet\b[^>]{0,2000}>/g) || [];
  for (const tag of sheetTags) {
    const id = (tag.match(/\br:id="([^"]{1,200})"/) || [])[1];
    const xml = id && target[id] ? text(target[id]) : '';
    const sheetPath = id && target[id] ? target[id] : '';
    out.sheets.push({
      ...chartsOf(sheetPath),
      cf: between(xml, '<conditionalFormatting', '</conditionalFormatting>', true),
      dv: between(xml, '<dataValidations', '</dataValidations>', false) + between(xml, '<x14:dataValidations', '</x14:dataValidations>', false),
      protection: (xml.match(/<sheetProtection\b[^>]{0,2000}>/) || [''])[0],
      unlocked: unlockedIn(xml),
      xf: styledCells(xml),
      freeze: frozen(xml),
      spark: between(xml, '<x14:sparklineGroups', '</x14:sparklineGroups>', false),
      tables: tablesOf(sheetPath, xml)
    });
  }
  out.dxfs = between(text('xl/styles.xml'), '<dxfs', '</dxfs>', false);
  // The style table and theme colours, for the app to read as XML.
  const st = text('xl/styles.xml');
  out.styles = st.length < 5e6 ? st : '';
  const themePath = Object.values(target).find(t => /theme\/theme\d*\.xml$/.test(t));
  out.theme = themePath ? text(themePath).slice(0, 2e6) : '';
  return out;

  /** The sheet's drawing part and the chart parts it uses, as text. */
  function relsOf(path) {
    const map = Object.create(null);
    if (!path) return map;
    const dir = path.replace(/[^/]*$/, ''), rels = text(dir + '_rels/' + path.slice(dir.length) + '.rels');
    for (const m of rels.match(/<Relationship\b[^>]{0,2000}>/g) || []) {
      const id = (m.match(/\bId="([^"]{1,200})"/) || [])[1], t = (m.match(/\bTarget="([^"]{1,500})"/) || [])[1];
      if (!id || !t) continue;
      const parts = (t[0] === '/' ? t.slice(1) : dir + t).split('/'), stack = [];
      for (const p of parts) { if (p === '..') stack.pop(); else if (p !== '.') stack.push(p); }
      map[id] = stack.join('/');
    }
    return map;
  }
  /** The sheet's table parts (Format as Table), as text. */
  function tablesOf(sheetPath, xml) {
    const ids = [...xml.matchAll(/<tablePart\b[^>]{0,300}?r:id="([^"]{1,200})"/g)].map(m => m[1]).slice(0, 200);
    if (!ids.length) return [];
    const rels = relsOf(sheetPath);
    return ids.map(id => (rels[id] ? text(rels[id]).slice(0, LIMIT) : '')).filter(Boolean);
  }
  function chartsOf(sheetPath) {
    const xml = sheetPath ? text(sheetPath) : '';
    const m = /<drawing\b[^>]{0,500}?r:id="([^"]{1,200})"/.exec(xml);
    if (!m) return {};
    const dPath = relsOf(sheetPath)[m[1]];
    const drawing = dPath ? text(dPath) : '';
    if (!drawing || drawing.length > LIMIT) return {};
    const dRels = relsOf(dPath);
    const charts = Object.keys(dRels).filter(id => /charts\/chart\d+\.xml$/.test(dRels[id])).slice(0, 50).map(rid => ({ rid, xml: text(dRels[rid]).slice(0, LIMIT) }));
    return { drawing, charts };
  }
}
