// Sheets: trendlines (and in Excel), FORECAST.ETS and Forecast Sheet, Analyze Data, the data
// model (relationships in PivotTables), recording actions as a script, From Database (SQLite)
// and From Folder.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const JSZip = require(path.join(process.cwd(), 'node_modules/jszip'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const W = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);
  const answer = list => p.evaluate(list => { window.__a = list.slice(); UI.form = async () => window.__a.shift(); UI.confirm = async () => true; }, list);
  const put = rows => p.evaluate(rows => rows.forEach((row, r) => row.forEach((v, c) => { if (v !== null) AtlasSheets.setCell(Formula.addr(r, c), String(v)); })), rows);
  const sel = (r1, c1, r2, c2) => p.evaluate(([a, b, c, d]) => { AtlasSheets.select(a, b); AtlasSheets.select(c, d, true); }, [r1, c1, r2, c2]);

  // ---- A season and a trend: FORECAST.ETS, then Forecast Sheet
  const q = Array.from({ length: 24 }, (_, i) => [i + 1, 100 + 2 * (i + 1) + [12, -6, 0, 6][i % 4]]);
  await put([['Quarter', 'Sales'], ...q]);
  await p.evaluate(() => AtlasSheets.setCell('D1', '=FORECAST.ETS(25,B2:B25,A2:A25)'));
  const f25 = await p.evaluate(() => AtlasSheets.valueOf('D1'));
  check('FORECAST.ETS continues the trend and season', Math.abs(f25 - 162) < 0.1, f25);
  await sel(0, 0, 24, 1);
  await answer([{ n: '8', conf: '95', season: '1' }]);
  await p.click('[data-tab="data"]');
  await p.click('#forecastBtn');
  await p.waitForFunction(() => AtlasSheets.activeSheetName() === 'Forecast');
  const fc = await p.evaluate(() => ({ c26: AtlasSheets.valueOf('C26'), lo: AtlasSheets.valueOf('D26'), hi: AtlasSheets.valueOf('E26'), a33: AtlasSheets.valueOf('A33'), f: AtlasSheets.rawOf('C27'), chart: AtlasSheets.book.sheets.find(s => s.name === 'Forecast').charts[0] }));
  check('Forecast Sheet: a forecast, a confidence band and a chart', Math.abs(fc.c26 - 162) < 0.1 && fc.lo < fc.c26 && fc.hi > fc.c26 && fc.a33 === 32 && /^=FORECAST\.ETS\(A27,/.test(fc.f) && fc.chart && fc.chart.type === 'line' && fc.chart.range === 'A1:E33', JSON.stringify(fc));

  // ---- Trendline on a chart, saved as Excel's own
  await p.evaluate(() => AtlasSheets.setSheet(0));
  await sel(0, 0, 24, 1);
  await p.click('[data-tab="insert"]');
  await p.click('[data-chart="scatter"]');
  await p.evaluate(() => { const ch = AtlasSheets.book.sheets[0].charts[0]; ch.trend = 'linear'; ch.trendEq = true; });
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const chartXml = await zip.file(Object.keys(zip.files).find(n => /xl\/charts\/chart\d+\.xml$/.test(n))).async('string');
  check('the trendline is saved as Excel’s (linear, with equation and R²)', /<c:trendline>[\s\S]*<c:trendlineType val="linear"\/>[\s\S]*<c:dispEq val="1"\/>/.test(chartXml));
  fs.writeFileSync(path.join(W, 'trend.xlsx'), Buffer.from(bytes));
  const opened = execFileSync('python3', ['-I', '-c', 'import sys, openpyxl; wb = openpyxl.load_workbook(sys.argv[1]); print(len(wb.sheetnames))', path.join(W, 'trend.xlsx')]).toString().trim();
  check('openpyxl opens the workbook', +opened >= 2, opened);
  const p2 = await b.newPage();
  await p2.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await p2.waitForFunction(() => window.AtlasSheets);
  await p2.evaluate(async b => AtlasSheets.loadFile({ name: 't.xlsx', path: null, data: new Uint8Array(b) }), bytes);
  const back = await p2.evaluate(() => AtlasSheets.book.sheets[0].charts[0]);
  check('and comes back when the file is opened', back && back.trend === 'linear' && back.trendEq === true, JSON.stringify(back));
  await p2.close();

  // ---- Analyze Data
  await p.evaluate(() => { const s = AtlasSheets.book; s.sheets.push({ ...JSON.parse(JSON.stringify(s.sheets[0])), name: 'Ads', cells: {}, charts: [], pivots: [], tables: [] }); AtlasSheets.setSheet(s.sheets.length - 1); });
  const regions = ['North', 'South', 'East', 'West'];
  const rows = Array.from({ length: 20 }, (_, i) => [regions[i % 4], i + 1, 50 + 10 * i + (i % 4 === 0 ? 30 : 0), 5 + 2 * i]);
  rows[13][2] = 2000; // one unusual value
  await put([['Region', 'Month', 'Sales', 'Ads'], ...rows]);
  await p.evaluate(() => AtlasSheets.select(1, 1));
  await p.click('[data-tab="data"]');
  await p.click('#analyzeBtn');
  await p.waitForSelector('.insights-pane .ins-card');
  const cards = await p.$$eval('.insights-pane .ins-card p', ps => ps.map(x => x.textContent));
  check('Analyze Data finds the unusual value', cards.some(t => /Row 15: Sales is 2,000, unusually high/.test(t)), cards.join(' / '));
  check('a trend, a top category and related columns', cards.some(t => /Ads rises steadily/.test(t)) && cards.some(t => /has the most Sales by Region/.test(t)) && cards.some(t => /Ads rises as Month rises/.test(t)), cards.join(' / '));
  const before = await p.evaluate(() => AtlasSheets.book.sheets.length);
  await p.click('.insights-pane .ins-category button');
  await p.waitForFunction(n => AtlasSheets.book.sheets.length === n + 1, before);
  const pv = await p.evaluate(() => { const s = AtlasSheets.book.sheets[AtlasSheets.book.active]; return { rows: s.pivots[0].rows, vals: s.pivots[0].vals, a4: AtlasSheets.valueOf('A4') }; });
  check('its PivotTable button makes the PivotTable', pv.rows[0] === 'Region' && pv.vals[0].field === 'Sales' && regions.includes(pv.a4), JSON.stringify(pv));
  await p.click('.insights-pane [data-a="close"]');

  // ---- The data model: a PivotTable on Sales uses the Products table's columns
  await p.evaluate(() => { const bk = AtlasSheets.book; bk.sheets.push({ ...JSON.parse(JSON.stringify(bk.sheets[0])), name: 'Model', cells: {}, charts: [], pivots: [], tables: [] }); AtlasSheets.setSheet(bk.sheets.length - 1); });
  await put([['ProductID', 'Qty', null, 'ID', 'Name'], [1, 5, null, 1, 'Apples'], [2, 3, null, 2, 'Pears'], [1, 4, null, 3, 'Plums'], [3, 7]]);
  await p.evaluate(() => { const s = AtlasSheets.book.sheets[AtlasSheets.book.active]; s.tables.push({ name: 'Sales', r1: 0, c1: 0, r2: 4, c2: 1, totals: false, banded: true, style: 'blue', filter: {} }, { name: 'Products', r1: 0, c1: 3, r2: 3, c2: 4, totals: false, banded: true, style: 'blue', filter: {} }); });
  await answer([{ ft: 'Sales', fc: 'ProductID', tt: 'Products', tc: 'ID' }]);
  await p.click('#relBtn');
  await p.waitForFunction(() => (AtlasSheets.book.relationships || []).length === 1);
  await p.evaluate(() => AtlasSheets.createPivot('Model', 'A1:B5', ['Products.Name'], [{ field: 'Qty', agg: 'sum' }]));
  const piv = await p.evaluate(() => [3, 4, 5, 6].map(r => [AtlasSheets.valueOf('A' + r), AtlasSheets.valueOf('B' + r)]));
  check('Relationships: a PivotTable on Sales by Products.Name', JSON.stringify(piv) === JSON.stringify([['Products.Name', 'Sum of Qty'], ['Apples', 9], ['Pears', 3], ['Plums', 7]]), JSON.stringify(piv));

  // ---- Record actions, then run them again
  await p.evaluate(() => { const bk = AtlasSheets.book; bk.sheets.push({ ...JSON.parse(JSON.stringify(bk.sheets[0])), name: 'Rec', cells: {}, charts: [], pivots: [], tables: [] }); AtlasSheets.setSheet(bk.sheets.length - 1); AtlasSheets.select(0, 0); });
  await p.click('[data-tab="automate"]');
  await p.click('#recordBtn');
  await p.focus('#gridWrap');
  await p.keyboard.type('Total');
  await p.keyboard.press('Enter');
  await p.evaluate(() => AtlasSheets.select(0, 0));
  await p.click('[data-tab="home"]');
  await p.click('[data-style="b"]');
  await p.evaluate(() => AtlasSheets.select(1, 0));
  await p.focus('#gridWrap');
  await p.keyboard.type('=1+2');
  await p.keyboard.press('Enter');
  await p.click('[data-tab="automate"]');
  await p.click('#recordBtn');
  await p.waitForSelector('.script-pane textarea');
  const code = await p.inputValue('.script-pane textarea');
  check('the recording is a script of what was done', /getRange\("A1"\)\.setValue\("Total"\)/.test(code) && /getRange\("A1"\)\.getFormat\(\)\.setBold\(true\)/.test(code) && /getRange\("A2"\)\.setFormula\("=1\+2"\)/.test(code), code);
  await p.evaluate(() => { const s = AtlasSheets.book.sheets[AtlasSheets.book.active]; s.cells = {}; AtlasSheets.setSheet(AtlasSheets.book.active); });
  await p.click('.script-pane [data-a="run"]');
  await p.waitForFunction(() => AtlasSheets.valueOf('A2') === 3, null, { timeout: 10000 });
  check('running it does it again', await p.evaluate(() => AtlasSheets.valueOf('A1') === 'Total' && AtlasSheets.book.sheets[AtlasSheets.book.active].cells.A1.s.b === true));
  await p.click('.script-pane:not(.insights-pane) .x');

  // ---- From Database (SQLite)
  const dbFile = path.join(W, 'shop.sqlite');
  execFileSync('python3', ['-I', '-c', `
import sqlite3, sys
c = sqlite3.connect(sys.argv[1])
c.execute('create table people (id integer, name text, city text)')
c.executemany('insert into people values (?,?,?)', [(1,'Ada','London'),(2,'Grace','New York'),(3,'Alan','Manchester')])
c.execute('create view londoners as select name from people where city = "London"')
c.commit()`, dbFile]);
  await p.evaluate(b => { Platform.open = async () => ({ name: 'shop.sqlite', path: null, data: new Uint8Array(b) }); }, Array.from(fs.readFileSync(dbFile)));
  await answer([{ table: 'people', sql: '' }]);
  await p.click('[data-tab="data"]');
  await p.click('#fromDb');
  await p.waitForFunction(() => AtlasSheets.activeSheetName() === 'people', null, { timeout: 15000 });
  const db = await p.evaluate(() => ['A1', 'B1', 'C1', 'A2', 'B3', 'C4'].map(AtlasSheets.valueOf));
  check('From Database imports a SQLite table', db.join('|') === 'id|name|city|1|Grace|Manchester', db.join('|'));
  await answer([{ table: '\u0000sql', sql: 'SELECT city, COUNT(*) AS n FROM people GROUP BY city ORDER BY city' }]);
  await p.click('#fromDb');
  await p.waitForFunction(() => AtlasSheets.activeSheetName() === 'Query', null, { timeout: 15000 });
  check('or the result of a SQL query', (await p.evaluate(() => ['A1', 'B1', 'A2', 'B2'].map(AtlasSheets.valueOf).join('|'))) === 'city|n|London|1');

  // ---- From Folder (the browser asks for a folder)
  const dir = path.join(W, 'monthly');
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, 'jan.csv'), 'Region,Amount\nNorth,10\nSouth,20\n');
  fs.writeFileSync(path.join(dir, 'feb.csv'), 'Amount,Region,Note\n30,North,late\n');
  execFileSync('python3', ['-I', '-c', 'import sys, openpyxl; wb = openpyxl.Workbook(); ws = wb.active; ws.append(["Region", "Amount"]); ws.append(["East", 40]); wb.save(sys.argv[1])', path.join(dir, 'mar.xlsx')]);
  fs.writeFileSync(path.join(dir, 'readme.md'), 'not data');
  p.once('filechooser', fc => fc.setFiles(dir));
  await p.click('#fromFolder');
  await p.waitForFunction(() => AtlasSheets.activeSheetName() === 'monthly', null, { timeout: 15000 });
  const folder = await p.evaluate(() => { const out = []; for (let r = 1; r <= 5; r++) out.push(['A', 'B', 'C', 'D'].map(c => AtlasSheets.valueOf(c + r)).join(',')); return out; });
  const head = folder[0].split(','), recs = folder.slice(1).map(l => Object.fromEntries(l.split(',').map((v, i) => [head[i], v]))).map(o => `${o['Source.Name']}:${o.Region}:${o.Amount}:${o.Note}`).sort();
  check('From Folder combines the files, matching columns by name', head[0] === 'Source.Name' && head.length === 4 && recs.join(' ') === 'feb.csv:North:30:late jan.csv:North:10: jan.csv:South:20: mar.xlsx:East:40:', folder.join(' / '));

  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().catch(e => { console.log('FAIL crashed :: ' + e.message); process.exit(1); });
