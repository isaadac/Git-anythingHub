'use strict';
// Parses spreadsheet files away from the app. Whatever a malicious file does to SheetJS here
// (prototype pollution, catastrophic regex backtracking) stays in this worker, which the app
// throws away after every file and kills if it runs too long.
importScripts('../../../node_modules/xlsx/dist/xlsx.full.min.js');

self.onmessage = e => {
  try {
    const wb = XLSX.read(e.data, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true, bookFiles: true });
    const sheets = []; // an array, so a sheet called __proto__ can't do anything odd
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
        out[k] = { t: c.t, v: c.v instanceof Date ? c.v.getTime() : c.v, f: c.f, z: c.z, w: c.w, isDate: c.v instanceof Date };
      }
      sheets.push(out);
    }
    const extra = excelExtras(wb);
    const custom = wb.Custprops && typeof wb.Custprops.AtlasMeta === 'string' ? wb.Custprops.AtlasMeta : null;
    self.postMessage({ ok: true, SheetNames: wb.SheetNames.map(String), Sheets: sheets, AtlasMeta: custom, Extras: extra });
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
  const sheetTags = text('xl/workbook.xml').match(/<sheet\b[^>]{0,2000}>/g) || [];
  for (const tag of sheetTags) {
    const id = (tag.match(/\br:id="([^"]{1,200})"/) || [])[1];
    const xml = id && target[id] ? text(target[id]) : '';
    out.sheets.push({
      cf: between(xml, '<conditionalFormatting', '</conditionalFormatting>', true),
      dv: between(xml, '<dataValidations', '</dataValidations>', false) + between(xml, '<x14:dataValidations', '</x14:dataValidations>', false)
    });
  }
  out.dxfs = between(text('xl/styles.xml'), '<dxfs', '</dxfs>', false);
  return out;
}
