'use strict';
// Parses spreadsheet files away from the app. Whatever a malicious file does to SheetJS here
// (prototype pollution, catastrophic regex backtracking) stays in this worker, which the app
// throws away after every file and kills if it runs too long.
importScripts('../../../node_modules/xlsx/dist/xlsx.full.min.js');

self.onmessage = e => {
  try {
    const wb = XLSX.read(e.data, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true });
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
    const custom = wb.Custprops && typeof wb.Custprops.AtlasMeta === 'string' ? wb.Custprops.AtlasMeta : null;
    self.postMessage({ ok: true, SheetNames: wb.SheetNames.map(String), Sheets: sheets, AtlasMeta: custom });
  } catch (err) {
    self.postMessage({ ok: false, error: String(err && err.message || err) });
  }
};
