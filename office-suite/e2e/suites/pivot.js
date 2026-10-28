const { chromium } = require('playwright');
const JSZip = require('/home/user/Git-anythingHub/office-suite/node_modules/jszip');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  const data = [['Region', 'Rep', 'Year', 'Sales'], ['East', 'Ann', '2023', '100'], ['East', 'Bob', '2023', '50'], ['West', 'Cy', '2024', '70'], ['East', 'Ann', '2024', '30'], ['North', 'Di', '2024', '=5*2']];
  await p.evaluate(data => { data.forEach((row, r) => row.forEach((v, c) => AtlasSheets.setCell(String.fromCharCode(65 + c) + (r + 1), v))); AtlasSheets.select(0, 0); AtlasSheets.select(5, 3, true); }, data);
  const answer = list => p.evaluate(list => { window.__answers = list.slice(); UI.form = async () => window.__answers.shift(); }, list);
  const grid = () => p.evaluate(() => { const s = AtlasSheets.book.sheets[AtlasSheets.book.active]; const out = []; for (let r = 2; r < 12; r++) { const row = []; for (let c = 0; c < 5; c++) { const v = AtlasSheets.valueOf(String.fromCharCode(65 + c) + (r + 1)); row.push(v); } if (row.some(v => v !== '')) out.push(row.join('|')); } return out; });
  await p.click('[data-tab="insert"]');
  await answer([{ src: 'Sheet1!A1:D6', where: 'new', at: '' }]);
  await p.click('#insPivot');
  const st = await p.evaluate(() => ({ sheets: AtlasSheets.book.sheets.map(s => s.name), active: AtlasSheets.book.active, pane: !document.getElementById('pivotPane').hidden, fields: [...document.querySelectorAll('#pivotPane .pv-field span')].map(x => x.textContent), note: AtlasSheets.valueOf('A3') }));
  check('PivotTable on a new sheet with the field pane open', st.sheets.join() === 'Sheet1,PivotTable' && st.active === 1 && st.pane && st.fields.join() === 'Region,Rep,Year,Sales' && /choose fields/.test(st.note), JSON.stringify(st));
  // Tick Region (text -> Rows) and Sales (numbers -> Values)
  await p.click('#pivotPane .pv-field:nth-child(1) input');
  await p.click('#pivotPane .pv-field:nth-child(4) input');
  const g1 = await grid();
  check('ticking fields builds the table', g1.join(' / ') === 'Region|Sum of Sales||| / East|180||| / North|10||| / West|70||| / Grand Total|260|||', g1.join(' / '));
  const sty = await p.evaluate(() => { const c = AtlasSheets.book.sheets[1].cells; return [c.A3.s.b, c.A3.s.bg, c.A7.s.b, c.A7.s.border && c.A7.s.border.t]; });
  check('header and total rows styled', sty[0] && sty[1] === '#ddebf7' && sty[2] && sty[3] === 1, JSON.stringify(sty));
  // Move Year to columns via the Rows/Columns move menu: tick Year (numbers -> Values), then move to Columns
  await p.click('#pivotPane .pv-field:nth-child(3) input');
  const valsMove = await p.$$('#pivotPane .pv-area:nth-child(4) .pv-item .pv-move');
  await valsMove[1].selectOption('cols');
  const g2 = await grid();
  check('Year moved to Columns', g2.join(' / ') === 'Sum of Sales|Column Labels||| / Region|2023|2024|Grand Total| / East|150|30|180| / North||10|10| / West||70|70| / Grand Total|150|110|260|', g2.join(' / '));
  // Change aggregation to Average
  await p.selectOption('#pivotPane .pv-area:nth-child(4) .pv-item select.ctl:not(.pv-move)', 'average');
  const g3 = await grid();
  check('Summarize by Average', g3[0].startsWith('Average of Sales') && g3[2] === 'East|75|30|60|', g3.join(' / '));
  // Filter by Rep = Ann
  const repMove = await p.$('#pivotPane .pv-field:nth-child(2) input'); await repMove.click(); // Rep -> Rows
  const rowsMoves = await p.$$('#pivotPane .pv-area:nth-child(3) .pv-item .pv-move');
  await rowsMoves[1].selectOption('filters');
  await p.selectOption('#pivotPane .pv-area:nth-child(1) .pv-item select.ctl:not(.pv-move)', 'Ann');
  const g4 = await grid();
  check('Filter shows only Ann', g4.join(' / ') === 'Average of Sales|Column Labels||| / Region|2023|2024|Grand Total| / East|100|30|65| / Grand Total|100|30|65|', g4.join(' / '));
  await p.selectOption('#pivotPane .pv-area:nth-child(1) .pv-item select.ctl:not(.pv-move)', '\u0000all');
  // Change the source, refresh
  await p.evaluate(() => { AtlasSheets.book.active = 0; });
  await p.evaluate(() => { const s = AtlasSheets.book.sheets[0]; s.cells.D2.v = '1000'; AtlasSheets.refreshPivots(true); AtlasSheets.book.active = 1; });
  const g5 = await p.evaluate(() => { const s = AtlasSheets.book.sheets[1]; return s.cells.B5.v; });
  check('Refresh picks up source changes', g5 === '525', g5);
  // Clicking outside hides the pane, inside shows it
  await p.evaluate(() => { document.querySelector('.sheet-tabs .tabs .tab:nth-child(2)')?.click(); });
  await p.evaluate(() => AtlasSheets.select(20, 8));
  const hidden = await p.evaluate(() => document.getElementById('pivotPane').hidden);
  await p.evaluate(() => AtlasSheets.select(4, 1));
  const shown = await p.evaluate(() => !document.getElementById('pivotPane').hidden);
  check('pane follows the selection', hidden && shown);
  // Save and reopen: pivot definition kept and still live
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  await p.evaluate(async b => { await AtlasSheets.loadFile({ name: 'pv.xlsx', path: null, data: new Uint8Array(b) }); }, bytes);
  const re = await p.evaluate(() => { const s = AtlasSheets.book.sheets[1]; return { n: s.pivots.length, rows: s.pivots[0].rows, cols: s.pivots[0].cols, vals: s.pivots[0].vals, b5: AtlasSheets.book.sheets[1].cells.B5.v }; });
  check('xlsx keeps the PivotTable (values for other apps, definition for Atlas)', re.n === 1 && re.cols.join() === 'Year' && re.vals[0].agg === 'average' && re.b5 === '525', JSON.stringify(re));
  // Insert a row above the source data: source range follows
  await p.evaluate(() => { AtlasSheets.book.active = 0; });
  await p.evaluate(() => { AtlasSheets.select(0, 0); document.getElementById('insRow').click(); });
  const src = await p.evaluate(() => AtlasSheets.book.sheets[1].pivots[0].src.ref);
  check('source range moves with inserted rows', src === 'A2:D7', src);
  // Renaming the source sheet keeps the link
  await p.evaluate(() => { UI.prompt = async () => 'Data'; });
  await p.evaluate(() => document.querySelector('#sheetTabs .tab')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })));
  await p.waitForTimeout(200);
  const ren = await p.evaluate(() => [AtlasSheets.book.sheets[0].name, AtlasSheets.book.sheets[1].pivots[0].src.sheet]);
  check('renaming the source sheet updates the PivotTable', ren[0] === ren[1], JSON.stringify(ren));
  // Undo restores
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
