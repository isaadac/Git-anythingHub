'use strict';
// Reads a SQLite database file (Data → From Database) away from the page: SQLite compiled to
// WebAssembly (sql.js), on a copy of the file in memory, with no network or storage.
// library files: node_modules/sql.js/dist/sql-wasm.wasm
importScripts('../../../node_modules/sql.js/dist/sql-wasm.js', 'worker-lockdown.js');
// The engine loads its WebAssembly first; then nothing here can reach the network.
const ready = initSqlJs({ locateFile: f => '../../../node_modules/sql.js/dist/' + f }).then(SQL => {
  self.lockDown();
  return SQL;
});
let db = null;
const MAX_ROWS = 200000, MAX_COLS = 500;
const cell = v => (v === null || v === undefined ? '' : v instanceof Uint8Array ? `(${v.length} bytes)` : typeof v === 'bigint' ? Number(v) : v);
self.onmessage = async e => {
  const m = e.data || {};
  try {
    const SQL = await ready;
    if (m.t === 'open') {
      if (db) db.close();
      db = new SQL.Database(new Uint8Array(m.bytes));
      const tables = [];
      const list = db.exec("SELECT name, type FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY type, name");
      for (const [name, type] of (list[0] ? list[0].values : [])) {
        let rows = null;
        try { rows = db.exec(`SELECT COUNT(*) FROM "${String(name).replace(/"/g, '""')}"`)[0].values[0][0]; } catch { /* a view that fails */ }
        tables.push({ name, type, rows: cell(rows) });
      }
      self.postMessage({ ok: true, tables });
    } else if (m.t === 'query') {
      if (!db) throw new Error('No database is open');
      const stmt = db.prepare(String(m.sql));
      const columns = stmt.getColumnNames().slice(0, MAX_COLS);
      const rows = [];
      let more = false;
      while (stmt.step()) { if (rows.length >= MAX_ROWS) { more = true; break; } rows.push(stmt.get().slice(0, MAX_COLS).map(cell)); }
      stmt.free();
      self.postMessage({ ok: true, columns, rows, more });
    }
  } catch (err) { self.postMessage({ ok: false, error: String((err && err.message) || err) }); }
};
