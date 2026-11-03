'use strict';
// Isadek Sheets — getting data from a folder of files and from SQLite databases.
//
// From Folder combines every spreadsheet, CSV and JSON file in a folder into one table (their
// columns matched by name, with the file each row came from); in the desktop app it refreshes
// from the same folder. From Database opens a SQLite file and imports a table or the result of a
// query; the database is read in a worker (sql.js), never changed.
(function () {
  const FILES = /\.(csv|tsv|txt|json|xlsx|xls|xlsm|ods)$/i;
  const MAX_ROWS = 200000;

  function SheetsConnect(ctx) {
    const F = window.Formula;

    // ---------------- Reading files ----------------
    /** The rows of a file's first sheet (or its CSV/JSON table). */
    async function rowsOf(file) {
      if (/\.(xlsx|xls|xlsm|ods)$/i.test(file.name)) {
        const wb = await ctx.parseSpreadsheet(file.data);
        const ws = wb.Sheets[0] || {};
        const rows = [];
        for (const k of Object.keys(ws)) {
          if (k[0] === '!') continue;
          const p = F.parseRef(k);
          if (!p || p.row >= MAX_ROWS || p.col >= 500) continue;
          const c = ws[k];
          (rows[p.row] = rows[p.row] || [])[p.col] = c.isDate ? c.w || '' : c.v;
        }
        return Array.from({ length: rows.length }, (_, i) => rows[i] || []);
      }
      return SheetsWeb.parse(file.data, '', file.name).tables[0].rows;
    }
    /** Every file's rows under one header: columns matched by name, plus the file's name first. */
    async function combine(files) {
      const header = ['Source.Name'], colOf = new Map(), out = [];
      const skipped = [];
      for (const f of files) {
        let rows;
        try { rows = await rowsOf(f); } catch (err) { skipped.push(f.name); continue; }
        if (!rows.length) continue;
        const head = rows[0].map((h, i) => (h === undefined || h === null || h === '' ? `Column${i + 1}` : String(h)));
        const map = head.map(h => { const k = h.trim().toLowerCase(); if (!colOf.has(k)) { colOf.set(k, header.length); header.push(h.trim()); } return colOf.get(k); });
        for (const r of rows.slice(1)) {
          if (!r.some(v => v !== undefined && v !== null && v !== '')) continue;
          const row = [f.name];
          r.forEach((v, i) => { if (map[i] !== undefined) row[map[i]] = v; });
          out.push(row);
          if (out.length >= MAX_ROWS) break;
        }
        if (out.length >= MAX_ROWS) break;
      }
      return { header, rows: out.map(r => Array.from({ length: header.length }, (_, i) => r[i] ?? '')), skipped };
    }
    async function pickFolderInBrowser() {
      return new Promise(resolve => {
        const inp = document.createElement('input');
        inp.type = 'file';
        inp.webkitdirectory = true;
        inp.multiple = true;
        inp.onchange = async () => {
          const list = [...inp.files].filter(f => FILES.test(f.name) && f.size <= 50e6).slice(0, 500);
          const files = [];
          for (const f of list) files.push({ name: f.webkitRelativePath.split('/').slice(1).join('/') || f.name, data: new Uint8Array(await f.arrayBuffer()) });
          resolve({ files, path: null, name: (list[0] && list[0].webkitRelativePath.split('/')[0]) || 'Folder' });
        };
        inp.click();
      });
    }
    async function readFolder(path) {
      if (window.atlas && window.atlas.readFolder) {
        const r = await window.atlas.readFolder(path || null);
        if (!r) return null;
        if (r.error) throw new Error(r.error);
        return { files: r.files.map(f => ({ name: f.name, data: new Uint8Array(f.data) })), path: r.path, name: r.path.split(/[\\/]/).pop() };
      }
      if (path) throw new Error('Refreshing a folder works in the desktop app; here, choose the folder again');
      return pickFolderInBrowser();
    }
    /** Writes the combined table into sheet s from A1 (clearing what an earlier refresh wrote). */
    function writeTable(s, t, previous) {
      if (previous) for (let r = 0; r <= previous.h; r++) for (let c = 0; c <= previous.w; c++) ctx.setCell(r, c, { v: '' }, s);
      t.header.forEach((h, c) => ctx.setCell(0, c, { v: ctx.rawFor(h), s: { b: true } }, s));
      t.rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '' && v !== null && v !== undefined) ctx.setCell(r + 1, c, { v: ctx.rawFor(v) }, s); }));
      s.freezeRow = true;
      return { h: t.rows.length, w: t.header.length - 1 };
    }
    async function fromFolder() {
      let got;
      try { got = await readFolder(null); } catch (err) { return UI.toast('Couldn’t read the folder: ' + err.message, 5000); }
      if (!got) return;
      if (!got.files.length) return UI.toast('That folder has no spreadsheets, CSV or JSON files');
      UI.toast(`Combining ${got.files.length} file${got.files.length === 1 ? '' : 's'}…`, 10000);
      const t = await combine(got.files);
      if (!t.rows.length) return UI.toast('The files have no rows of data');
      ctx.snapshot();
      const s = ctx.addSheet(got.name || 'Folder');
      const size = writeTable(s, t);
      if (got.path) s.folder = { path: got.path, h: size.h, w: size.w };
      ctx.done();
      UI.toast(`${t.rows.length.toLocaleString()} rows from ${got.files.length} files${t.skipped.length ? ` (${t.skipped.length} couldn’t be read)` : ''}${got.path ? '. Refresh All reads the folder again.' : ''}`, 5000);
    }
    /** Refresh All: folders read again (the desktop app asks before reading a folder a workbook names). */
    async function refreshFolders() {
      let n = 0;
      for (const s of ctx.book().sheets) {
        if (!s.folder || !s.folder.path) continue;
        let got;
        try { got = await readFolder(s.folder.path); } catch (err) { UI.toast(`${s.name}: ${err.message}`, 5000); continue; }
        if (!got) continue;
        const t = await combine(got.files);
        ctx.snapshot();
        const size = writeTable(s, t, s.folder);
        s.folder = { path: got.path, h: size.h, w: size.w };
        n++;
      }
      if (n) ctx.done();
      return n;
    }

    // ---------------- SQLite ----------------
    function sqlWorker() {
      const w = new Worker('js/sheets-sql-worker.js');
      const call = msg => new Promise((resolve, reject) => {
        const timer = setTimeout(() => { w.terminate(); reject(new Error('the database took too long')); }, 60000);
        w.onmessage = e => { clearTimeout(timer); if (e.data && e.data.ok) resolve(e.data); else reject(new Error((e.data && e.data.error) || 'the database could not be read')); };
        w.postMessage(msg);
      });
      return { call, close: () => w.terminate() };
    }
    async function fromDatabase() {
      const f = await Platform.open([{ name: 'SQLite database', extensions: ['sqlite', 'sqlite3', 'db', 'db3'] }]);
      if (!f) return;
      if (f.data.length > 500e6) return UI.toast('That database is too large to import here (over 500 MB)');
      const db = sqlWorker();
      try {
        const { tables } = await db.call({ t: 'open', bytes: f.data });
        if (!tables.length) { UI.toast('That database has no tables'); return; }
        const v = await UI.form('Import from Database', [
          { name: 'table', label: 'Table or view', type: 'select', options: [...tables.map(t => ({ value: t.name, label: `${t.name}${t.type === 'view' ? ' (view)' : ''}${t.rows !== '' && t.rows !== null ? ` — ${Number(t.rows).toLocaleString()} rows` : ''}` })), { value: '\u0000sql', label: 'A SQL query…' }] },
          { name: 'sql', label: 'SQL query (when “A SQL query…” is chosen)', type: 'textarea', placeholder: 'SELECT region, SUM(amount) AS total FROM sales GROUP BY region' }
        ], 'Import');
        if (!v) return;
        const custom = v.table === '\u0000sql';
        if (custom && !v.sql.trim()) { UI.toast('Type a SQL query'); return; }
        const sql = custom ? v.sql : `SELECT * FROM "${v.table.replace(/"/g, '""')}"`;
        const res = await db.call({ t: 'query', sql });
        ctx.snapshot();
        const s = ctx.addSheet(custom ? 'Query' : v.table);
        writeTable(s, { header: res.columns, rows: res.rows });
        ctx.done();
        UI.toast(`${res.rows.length.toLocaleString()} rows imported${res.more ? ' (the first 200,000)' : ''}`, 4000);
      } catch (err) {
        UI.toast('Couldn’t read the database: ' + err.message, 6000);
      } finally { db.close(); }
    }

    return { fromFolder, refreshFolders, fromDatabase, combine };
  }
  window.SheetsConnect = SheetsConnect;
})();
