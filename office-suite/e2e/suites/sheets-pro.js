// Sheets, pro features: Clean & Transform (steps on a copy of the data, run again by Refresh All),
// waterfall, funnel and map charts, and Python in cells (=PY), which a workbook from a file only
// runs once allowed.
const { chromium } = require('playwright');
const fs = require('fs');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
  const browser = await chromium.launch();
  const p = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);
  const fill = rows => p.evaluate(rows => rows.forEach((row, r) => row.forEach((v, c) => AtlasSheets.setCell(String.fromCharCode(65 + c) + (r + 1), String(v)))), rows);
  const val = ref => p.evaluate(r => { const v = AtlasSheets.valueOf(r); return v && v.error ? v.error : v; }, ref);
  // Fills a UI.form dialog: values in the order of its fields.
  const form = async values => {
    await p.waitForSelector('.modal-backdrop:last-child .modal .fields');
    const els = await p.$$('.modal-backdrop:last-child .modal .fields label > :is(input, select, textarea)');
    for (let i = 0; i < values.length; i++) {
      if ((await els[i].evaluate(e => e.tagName)) === 'SELECT') await els[i].selectOption(values[i]); else await els[i].fill(values[i]);
    }
    await p.click('.modal-backdrop:last-child .actions .primary');
  };
  const tool = async (label, values) => { await p.click(`.clean-editor .ce-tools button:has-text("${label}")`); if (values) await form(values); await sleep(100); };

  // ---- Clean & Transform
  await fill([['Name', 'City', 'Sales'], ['  ada   lovelace ', 'london', '1,200'], ['BOB SMITH', ' Paris', '300'], ['ada lovelace', 'London ', '1200'], ['', '', ''], ['carol jones', 'berlin', '50']]);
  await p.evaluate(() => { AtlasSheets.select(0, 0); AtlasSheets.select(5, 2, true); });
  await p.click('[data-tab="data"]');
  await p.click('#cleanData');
  await p.waitForSelector('.clean-editor');
  await tool('Trim spaces', ['*']);
  await tool('Change case', ['*', 'proper']);
  await tool('Change type', ['Sales', 'number']);
  await tool('Remove blank rows');
  await tool('Remove duplicates', ['*']);
  await tool('Sort', ['Sales', '1']);
  const preview = await p.$$eval('.clean-editor .ce-preview tbody tr', trs => trs.map(t => [...t.cells].map(c => c.textContent)));
  check('the preview shows the cleaned data', JSON.stringify(preview) === JSON.stringify([['Ada Lovelace', 'London', '1200'], ['Bob Smith', 'Paris', '300'], ['Carol Jones', 'Berlin', '50']]), JSON.stringify(preview));
  const steps = await p.$$eval('.clean-editor .ce-steps li', l => l.map(x => x.textContent));
  check('…with the applied steps listed', steps.length === 8 && steps[0] === 'Source (5 rows)' && /Removed duplicate rows/.test(steps.join('|')), steps.join(' | '));
  // A step can be taken out again.
  await p.click('.clean-editor .ce-steps li:last-child button');
  check('removing a step undoes just that step', (await p.$$eval('.clean-editor .ce-preview tbody tr td:first-child', t => t.map(x => x.textContent))).join() === 'Ada Lovelace,Bob Smith,Carol Jones');
  await tool('Sort', ['Sales', '1']);
  await p.click('.clean-editor .load');
  await sleep(300);
  let b = await p.evaluate(() => JSON.parse(AtlasSheets.bookJson()));
  check('Load puts the result on a new sheet', b.sheets.length === 2 && b.sheets[1].name === 'Cleaned Sheet1' && (await p.evaluate(() => AtlasSheets.activeSheetName())) === 'Cleaned Sheet1' && (await val('A2')) === 'Ada Lovelace' && (await val('C2')) === 1200 && (await val('A4')) === 'Carol Jones', `${await val('A2')} ${await val('C2')}`);
  // New data in the source, then Refresh All.
  await p.evaluate(() => { AtlasSheets.setSheet(0); AtlasSheets.setCell('A7', 'dee  dee'); AtlasSheets.setCell('B7', 'oslo'); AtlasSheets.setCell('C7', '5,000'); AtlasSheets.setSheet(1); });
  await p.evaluate(() => document.getElementById('refreshAll').click());
  await sleep(400);
  check('Refresh All runs the steps again on the latest data', (await val('A2')) === 'Dee Dee' && (await val('C2')) === 5000 && (await val('A5')) === 'Carol Jones', `${await val('A2')} ${await val('C2')}`);

  // The steps are kept in the file.
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  const file = process.env.E2E_WORK + '/files/pro.xlsx';
  fs.writeFileSync(file, Buffer.from(bytes));
  await p.evaluate(async b => AtlasSheets.loadFile({ name: 'pro.xlsx', data: new Uint8Array(b) }), bytes);
  b = await p.evaluate(() => JSON.parse(AtlasSheets.bookJson()));
  check('the steps are saved with the workbook', b.sheets[1].cleans && b.sheets[1].cleans[0].steps.length === 7 && b.sheets[1].cleans[0].src.ref === 'A1:C7', JSON.stringify(b.sheets[1].cleans));

  // ---- Waterfall, funnel and map charts
  await p.evaluate(() => { AtlasSheets.book.sheets.push({ ...JSON.parse(JSON.stringify(AtlasSheets.book.sheets[0])), name: 'Charts', cells: {}, charts: [], cleans: [] }); AtlasSheets.setSheet(2); });
  await fill([['Step', 'Amount'], ['Start', '100'], ['Sales', '40'], ['Costs', '-30'], ['Total', '110']]);
  await p.evaluate(() => { AtlasSheets.select(0, 0); AtlasSheets.select(4, 1, true); });
  await p.click('[data-tab="insert"]');
  await p.click('[data-chart="waterfall"]');
  await p.click('[data-chart="funnel"]');
  await fill([['Country', 'Users'], ['United States', '300'], ['Brazil', '120'], ['India', '200'], ['Germany', '80'], ['Atlantis', '5']]);
  await p.evaluate(() => { AtlasSheets.select(0, 0); AtlasSheets.select(5, 1, true); });
  await p.click('[data-chart="map"]');
  await sleep(1500);
  const shots = await p.evaluate(() => [...document.querySelectorAll('#chartsLayer canvas')].map(c => {
    const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
    let ink = 0;
    for (let i = 0; i < d.length; i += 4) if (d[i] < 235 || d[i + 1] < 235 || d[i + 2] < 235) ink++;
    return ink;
  }));
  check('waterfall, funnel and map charts draw', shots.length === 3 && shots.every(n => n > 2000), shots.join());
  const charts = await p.evaluate(() => AtlasSheets.book.sheets[2].charts.map(c => c.type));
  check('…as their own chart types', charts.join() === 'waterfall,funnel,map', charts.join());
  check('the map loads its outlines only when needed', await p.evaluate(() => !!window.WorldMap && [...document.scripts].filter(s => /world-map\.js/.test(s.src)).length === 1));

  // ---- Python in cells
  await p.evaluate(() => { AtlasSheets.setCell('E1', '=PY("sum(r[0] for r in xl(""B2:B5""))")'); AtlasSheets.setCell('E3', '=PY("[[n, n * n] for n in range(1, 4)]")'); AtlasSheets.setCell('E7', '=PY("args[0] * 2", B2)'); AtlasSheets.setCell('E8', '=PY("1/0")'); AtlasSheets.setCell('E9', '=E1+1'); });
  check('a Python cell shows #BUSY! while its code runs', (await val('E1')) === '#BUSY!');
  await p.waitForFunction(() => AtlasSheets.valueOf('E1') === 700, null, { timeout: 90000 }).catch(() => {});
  check('=PY runs Python and reads cells with xl()', (await val('E1')) === 700, String(await val('E1')));
  check('…a list of rows spills', (await val('E3')) === 1 && (await val('F5')) === 9 && (await val('E5')) === 3);
  check('…values after the code arrive as args', (await val('E7')) === 600);
  check('…and other formulas use the result', (await val('E9')) === 701);
  await p.evaluate(() => AtlasSheets.select(7, 4));
  check('a Python error is #CALC!, with the reason on the formula bar', (await val('E8')) === '#CALC!' && /ZeroDivisionError/.test(await p.getAttribute('#fx', 'title')), await p.getAttribute('#fx', 'title'));
  await p.evaluate(() => AtlasSheets.setCell('B2', '200'));
  await p.waitForFunction(() => AtlasSheets.valueOf('E1') === 600, null, { timeout: 30000 }).catch(() => {});
  check('changing a cell the code reads runs it again', (await val('E1')) === 600 && (await val('E7')) === 400);
  await p.evaluate(() => AtlasSheets.setCell('E11', '=PY("import js\\njs.fetch")'.replace('\\n', '"&CHAR(10)&"')));
  await p.waitForFunction(() => AtlasSheets.valueOf('E11') !== Formula.ERR.BUSY, null, { timeout: 30000 }).catch(() => {});
  check('Python has no way to the network', /^(|#CALC!|None)$/.test(String(await val('E11'))), String(await val('E11')));

  // A workbook from a file with Python in it.
  const pyBytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  await p.evaluate(async b => AtlasSheets.loadFile({ name: 'python.xlsx', data: new Uint8Array(b) }), pyBytes);
  await p.evaluate(() => AtlasSheets.setSheet(2));
  check('Python from a file waits until allowed', (await val('E1')) === '#BLOCKED!' && await p.isVisible('#pyBar'), String(await val('E1')));
  await p.click('#pyAllow');
  await p.waitForFunction(() => AtlasSheets.valueOf('E1') === 600, null, { timeout: 30000 }).catch(() => {});
  check('…and runs once it is', (await val('E1')) === 600 && !(await p.isVisible('#pyBar')), String(await val('E1')));

  await p.evaluate(() => { window.__atlasDirty = false; });
  check('no page errors', !errors.length, errors.slice(0, 3).join(' | '));
  await browser.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
