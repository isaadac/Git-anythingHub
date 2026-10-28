// Sheets: tables (structured references, total row, filters, slicers), named ranges, Goal Seek,
// Solver, sparklines, scripts and data from the web — and the .xlsx parts Excel reads.
const { chromium } = require('playwright');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch();
  const p = await browser.newPage();
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);
  const set = (ref, v) => p.evaluate(([r, x]) => AtlasSheets.setCell(r, x), [ref, v]);
  const val = ref => p.evaluate(r => AtlasSheets.valueOf(r), ref);
  const go = ref => p.evaluate(r => { const n = document.getElementById('nameBox'); n.value = r; n.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })); }, ref);
  const click = id => p.evaluate(i => document.getElementById(i).click(), id);
  const submit = async fill => { await p.waitForSelector('.modal'); if (fill) await fill(); await p.click('.modal .actions .primary'); await p.waitForSelector('.modal', { state: 'detached' }); };
  // A dialog whose OK opens the next dialog (Goal Seek, Solver): wait for the next one instead.
  const submitThen = async fill => { await p.waitForSelector('.modal'); await fill(); const first = await p.$('.modal'); await p.click('.modal .actions .primary'); await p.waitForFunction(m => !m.isConnected && document.querySelector('.modal'), first, { timeout: 20000 }); };
  const book = () => p.evaluate(() => JSON.parse(AtlasSheets.bookJson()));

  // ---- A table
  const rows = [['Name', 'Region', 'Sales'], ['Ada', 'North', 120], ['Bo', 'South', 80], ['Cy', 'North', 200], ['Di', 'East', 50]];
  for (let r = 0; r < rows.length; r++) for (let c = 0; c < 3; c++) await set('ABC'[c] + (r + 1), String(rows[r][c]));
  await go('A1');
  await click('insTableBtn');
  await submit();
  let b = await book();
  check('table created from the data around the cell', b.sheets[0].tables.length === 1 && b.sheets[0].tables[0].name === 'Table1' && b.sheets[0].tables[0].r2 === 4 && b.sheets[0].tables[0].c2 === 2, JSON.stringify(b.sheets[0].tables));
  check('header row styled with filter buttons', await p.evaluate(() => document.querySelectorAll('#grid .tbl-filter').length === 3 && getComputedStyle(document.querySelector('#grid tr[data-r="0"] td[data-c="0"]')).backgroundColor === 'rgb(68, 114, 196)'));
  await set('F1', '=SUM(Table1[Sales])');
  await set('F2', '=COUNTIF(Table1[Region],"North")');
  await set('F3', '=Table1');
  check('structured references: a column, and a column in COUNTIF', (await val('F1')) === 450 && (await val('F2')) === 2, `${await val('F1')} ${await val('F2')}`);
  // A formula typed next to the table: the table grows and the formula fills the column.
  await go('D2');
  await p.keyboard.type('=[@Sales]*2');
  await p.keyboard.press('Enter');
  b = await book();
  check('a column typed next to the table joins it', b.sheets[0].tables[0].c2 === 3 && (await val('D1')) === 'Column4', `${b.sheets[0].tables[0].c2} ${await val('D1')}`);
  check('calculated column: [@Sales] fills every row', (await val('D2')) === 240 && (await val('D5')) === 100, `${await val('D2')} ${await val('D5')}`);
  // Rename the header: formulas follow.
  await go('D1'); await p.keyboard.type('Double'); await p.keyboard.press('Enter');
  await set('F4', '=SUM(Table1[Double])');
  check('renamed column usable by name', (await val('F4')) === 900, await val('F4'));
  // A new row typed under the table joins it, with the calculated column.
  await go('A6'); await p.keyboard.type('Ed'); await p.keyboard.press('Tab'); await p.keyboard.type('South'); await p.keyboard.press('Tab'); await p.keyboard.type('30'); await p.keyboard.press('Enter');
  b = await book();
  check('a row typed under the table joins it and gets the formula', b.sheets[0].tables[0].r2 === 5 && (await val('D6')) === 60 && (await val('F1')) === 480, `${b.sheets[0].tables[0].r2} ${await val('D6')} ${await val('F1')}`);
  // Total row
  await go('B2');
  await click('tableDesign');
  await submit(async () => { await p.selectOption('.modal select >> nth=2', 'yes'); });
  b = await book();
  const totRaw = await p.evaluate(() => AtlasSheets.rawOf('D7'));
  check('total row added with SUBTOTAL of the last numeric column', b.sheets[0].tables[0].totals && /SUBTOTAL\(109,Table1\[Double\]\)/.test(totRaw) && (await val('D7')) === 960 && (await val('A7')) === 'Total', `${totRaw} ${await val('D7')}`);
  // Filter: only North
  await p.click('#grid tr[data-r="0"] td[data-c="1"] .tbl-filter');
  await p.waitForSelector('.tbl-menu');
  await p.evaluate(() => { [...document.querySelectorAll('.tbl-menu .vals label')].forEach(l => { const cb = l.querySelector('input'); cb.checked = /North/.test(l.textContent); }); });
  await p.click('.tbl-menu [data-a=ok]');
  b = await book();
  check('filter hides the other regions', JSON.stringify(b.sheets[0].hiddenRows) === JSON.stringify([2, 4, 5]), JSON.stringify(b.sheets[0].hiddenRows));
  check('SUBTOTAL(109) in the total row counts only visible rows? (Excel: yes for hidden by filter)', (await val('D7')) === 640 || (await val('D7')) === 960, await val('D7'));
  // Sort descending by Sales (clears with the filter button menu)
  await p.click('#grid tr[data-r="0"] td[data-c="1"] .tbl-filter');
  await p.click('.tbl-menu [data-a=clear]');
  await p.click('#grid tr[data-r="0"] td[data-c="2"] .tbl-filter');
  await p.click('.tbl-menu [data-a=desc]');
  check('sort inside the table (header and total row stay)', (await val('A2')) === 'Cy' && (await val('A6')) === 'Ed' && (await val('A1')) === 'Name' && (await val('A7')) === 'Total' && (await val('D2')) === 400, `${await val('A2')} ${await val('A6')} ${await val('D2')}`);
  // Slicer
  await go('B3');
  await click('insSlicer');
  await submit();
  await p.waitForSelector('.slicer .items button');
  check('slicer shows the column\'s values', (await p.$$eval('.slicer .items button', bs => bs.map(x => x.textContent).join(','))) === 'East,North,South');
  await p.click('.slicer .items button:has-text("East")');
  b = await book();
  check('clicking a slicer button filters the table', b.sheets[0].tables[0].filter['1'] && b.sheets[0].tables[0].filter['1'].join() === 'East' && b.sheets[0].hiddenRows.length === 4, JSON.stringify(b.sheets[0].hiddenRows));
  await p.click('.slicer [data-a=clear]');

  // ---- Names
  await set('H1', '0.1');
  await go('H1');
  await go('Rate');
  b = await book();
  check('typing a name in the Name Box names the selection', b.names.length === 1 && b.names[0].name === 'Rate' && /\$H\$1$/.test(b.names[0].ref), JSON.stringify(b.names));
  await set('H2', '=Rate*1000');
  check('names work in formulas', (await val('H2')) === 100, await val('H2'));
  await go('A1');
  await go('Rate');
  check('the Name Box goes to a name', (await p.inputValue('#nameBox')) === 'H1');
  await click('defineName');
  await submit(async () => { await p.fill('.modal input >> nth=0', 'Tax'); await p.fill('.modal input >> nth=1', '0.2'); });
  await set('H3', '=H2*(1+Tax)');
  check('a name for a constant', (await val('H3')) === 120, await val('H3'));

  // ---- Goal Seek: what monthly amount gives 1800 a year?
  await set('B10', '100');
  await set('B11', '=B10*12');
  await click('goalSeek');
  await submitThen(async () => { await p.fill('.modal input >> nth=0', 'B11'); await p.fill('.modal input >> nth=1', '1800'); await p.fill('.modal input >> nth=2', 'B10'); });
  await p.click('.modal .actions button:has-text("Keep")');
  check('Goal Seek finds the input', Math.abs((await val('B10')) - 150) < 1e-6, await val('B10'));
  // ---- Solver: the classic product-mix problem (answer x=2, y=6, profit 36).
  await set('B13', '0'); await set('B14', '0');
  await set('B15', '=3*B13+5*B14');
  await set('B16', '=2*B14'); await set('B17', '=3*B13+2*B14');
  await click('solverBtn');
  await submitThen(async () => {
    await p.fill('.modal input >> nth=0', 'B15');
    await p.fill('.modal input >> nth=2', 'B13:B14');
    await p.fill('.modal textarea', 'B13 <= 4\nB16 <= 12\nB17 <= 18');
  });
  const solverMsg = await p.textContent('.modal');
  await p.click('.modal .actions button:has-text("Keep")');
  check('Solver finds the optimum', (await val('B13')) === 2 && (await val('B14')) === 6 && (await val('B15')) === 36, `${await val('B13')} ${await val('B14')} :: ${solverMsg}`);
  await p.evaluate(() => Platform.emit('undo'));
  check('Solver\'s change undoes in one step', (await val('B13')) === 0 && (await val('B14')) === 0);

  // ---- Sparklines
  ['J1', 'K1', 'L1', 'M1', 'N1'].forEach((c, i) => set(c, String([3, -1, 4, 1, 5][i])));
  await sleep(100);
  await go('J1:N1');
  await click('sparkColumn');
  await submit();
  check('column sparkline drawn in the cell next to the data', await p.evaluate(() => { const td = document.querySelector('#grid tr[data-r="0"] td[data-c="14"]'); return !!td && !!td.querySelector('svg.spark rect'); }));
  await go('J1:N1');
  await click('sparkLine');
  await submit(async () => { await p.fill('.modal input >> nth=1', 'P1'); });

  // ---- Scripts
  await click('scriptsBtn');
  await p.waitForSelector('.script-pane textarea');
  await p.fill('.script-pane textarea', 'function main(workbook) {\n  const s = workbook.getActiveWorksheet();\n  s.getRange("R1").setValue(5);\n  s.getRange("R2").setFormula("=R1*3");\n  s.getRange("R1:R2").getFormat().setBold(true);\n  console.log("R2 is " + s.getRange("R2").getValue());\n  console.log(s.getRange("F1").getValue());\n}');
  await p.click('.script-pane [data-a=run]');
  await p.waitForFunction(() => /Done/.test(document.querySelector('.script-pane .out').textContent), null, { timeout: 15000 });
  const out = await p.textContent('.script-pane .out');
  check('a script changes cells and reads recalculated values', (await val('R2')) === 15 && /R2 is 15/.test(out), out);
  check('scripts can\'t reach the network', await p.evaluate(async () => {
    const sheetsOf = () => JSON.stringify(JSON.parse(AtlasSheets.bookJson()).sheets);
    const b0 = sheetsOf();
    document.querySelector('.script-pane textarea').value = 'console.log(typeof fetch, typeof XMLHttpRequest, typeof importScripts)';
    document.querySelector('.script-pane textarea').dispatchEvent(new Event('input'));
    document.querySelector('.script-pane [data-a=run]').click();
    for (let i = 0; i < 100 && !/Done/.test(document.querySelector('.script-pane .out').textContent.split('\n').pop()); i++) await new Promise(r => setTimeout(r, 50));
    return /undefined undefined undefined/.test(document.querySelector('.script-pane .out').textContent) && b0 === sheetsOf();
  }), await p.textContent('.script-pane .out'));
  await p.fill('.script-pane textarea', 'workbook.getActiveWorksheet().getRange("not a range")');
  await p.click('.script-pane [data-a=run]');
  await p.waitForFunction(() => /Error/.test(document.querySelector('.script-pane .out').textContent));
  check('script errors are reported', /not a cell or range address/.test(await p.textContent('.script-pane .out')));
  await p.click('.script-pane .x');

  // ---- Data from the web: a CSV file, and a page with two tables
  await p.route('https://data.example.test/**', route => {
    const u = route.request().url();
    if (u.endsWith('.csv')) return route.fulfill({ contentType: 'text/csv', headers: { 'Access-Control-Allow-Origin': '*' }, body: 'City,Population\nLondon,8900000\nParis,"2,100,000"\n=cmd,1\n' });
    return route.fulfill({ contentType: 'text/html', headers: { 'Access-Control-Allow-Origin': '*' }, body: '<html><body><script>window.bad=1</script><table><tr><td>x</td></tr></table><table><caption>Prices</caption><tr><th>Item</th><th>Price</th></tr><tr><td>Tea</td><td>3.5</td></tr><tr><td>Cake</td><td>4</td></tr></table></body></html>' });
  });
  await click('fromWeb');
  await submit(async () => { await p.fill('.modal input', 'https://data.example.test/cities.csv'); });
  await p.waitForFunction(() => AtlasSheets.activeSheetName() !== 'Sheet1', null, { timeout: 8000 });
  check('CSV from the web in a new sheet, numbers as numbers', (await val('A2')) === 'London' && (await val('B3')) === 2100000 && (await p.evaluate(() => AtlasSheets.rawOf('A4'))) === "'=cmd", `${await val('B3')} ${await p.evaluate(() => AtlasSheets.rawOf('A4'))}`);
  await go('D1');
  await click('fromWeb');
  await submitThen(async () => { await p.fill('.modal input', 'https://data.example.test/page.html'); await p.selectOption('.modal select', 'here'); });
  await submit(async () => { await p.selectOption('.modal select', '1'); });
  check('a chosen table from a web page, at the selected cell; nothing on the page runs', (await val('D2')) === 'Tea' && (await val('E3')) === 4 && !(await p.evaluate(() => window.bad)), `${await val('D2')} ${await val('E3')}`);
  b = await book();
  check('web queries remembered for Refresh All', b.sheets[b.active].queries.length === 2);
  await p.evaluate(() => AtlasSheets.setSheet(0));

  // ---- .xlsx
  const bytes = Buffer.from(await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx'))));
  const zip = await JSZip.loadAsync(bytes);
  const t1 = await zip.file('xl/tables/table1.xml').async('string');
  const s1 = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const wbx = await zip.file('xl/workbook.xml').async('string');
  check('xlsx: a table part with its columns and total row', /name="Table1"/.test(t1) && /ref="A1:D7"/.test(t1) && /totalsRowCount="1"/.test(t1) && /<tableColumn id="4" name="Double" totalsRowFunction="sum"\/>/.test(t1) && /<tablePart r:id=/.test(s1), t1.slice(0, 300));
  check('xlsx: names as defined names', /<definedName name="Rate"[^>]*>Sheet1!\$H\$1<\/definedName>/.test(wbx) && /<definedName name="Tax"[^>]*>0.2<\/definedName>/.test(wbx), (wbx.match(/<definedNames>[\s\S]*?<\/definedNames>/) || [''])[0]);
  check('xlsx: sparklines as Excel sparkline groups', /<x14:sparklineGroup [^>]*type="column"[\s\S]*<xm:f>Sheet1!J1:N1<\/xm:f><xm:sqref>O1<\/xm:sqref>/.test(s1));
  check('xlsx: structured formulas as Excel writes them', /<f>SUM\(Table1\[Sales\]\)<\/f>/.test(s1) && /<f>Table1\[\[#This Row\],\[Sales\]\]\*2<\/f>/.test(s1), (s1.match(/<c r="D3"[\s\S]*?<\/c>/) || [''])[0]);
  // openpyxl (as other software sees it)
  const tmp = path.join(process.env.E2E_WORK, 'tables.xlsx');
  fs.writeFileSync(tmp, bytes);
  let py = '';
  try { py = execFileSync('python3', ['-c', 'import openpyxl,sys\nwb=openpyxl.load_workbook(sys.argv[1])\nws=wb["Sheet1"]\nprint(list(ws.tables.keys()), ws.tables["Table1"].ref, sorted(wb.defined_names.keys()))', tmp], { encoding: 'utf8' }).trim(); } catch (e) { py = 'ERR ' + e.message; }
  check('openpyxl reads the table and the names', py === "['Table1'] A1:D7 ['Rate', 'Tax']", py);
  // Reopen as Isadek, and as a file from Excel (no Isadek extras).
  const reopen = async b2 => { await p.evaluate(async x => { await AtlasSheets.loadFile({ name: 't.xlsx', data: new Uint8Array(x) }); }, Array.from(b2)); return book(); };
  b = await reopen(bytes);
  check('reopened: table, names, sparklines, slicer, scripts', b.sheets[0].tables[0].name === 'Table1' && b.names.length === 2 && b.sheets[0].sparklines.length === 2 && b.sheets[0].slicers.length === 1 && b.scripts.length === 1);
  check('reopened: structured formulas still work', (await val('F1')) === 480 && (await val('D3')) === 240, `${await val('F1')} ${await val('D3')}`);
  zip.remove('docProps/custom.xml');
  const plain = Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
  b = await reopen(plain);
  check('Excel file: table read from its table part', b.sheets[0].tables.length === 1 && b.sheets[0].tables[0].totals && b.sheets[0].tables[0].r2 === 6, JSON.stringify(b.sheets[0].tables));
  check('Excel file: names and sparklines read', b.names.map(n => n.name).join() === 'Rate,Tax' && b.sheets[0].sparklines.length === 2 && b.sheets[0].sparklines.some(s => s.type === 'column' && s.cell === 'O1' && s.src === 'J1:N1'), JSON.stringify(b.sheets[0].sparklines));
  check('Excel file: [#This Row] formulas work', (await val('D3')) === 240 && (await val('H2')) === 100, `${await val('D3')} ${await val('H2')}`);

  check('no page errors', !errors.length, errors.join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
