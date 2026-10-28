const { chromium } = require('playwright');
const JSZip = require('jszip');
const fs = require('fs');
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await p.evaluate(() => {
    [['Month', 'North', 'South'], ['Jan', '10', '30'], ['Feb', '40', '60'], ['Mar', '50', '70']].forEach((row, r) => row.forEach((v, c) => AtlasSheets.setCell(String.fromCharCode(65 + c) + (r + 1), v)));
    AtlasSheets.select(0, 0); AtlasSheets.select(3, 2, true);
  });
  await p.click('[data-tab="insert"]');
  for (const t of ['bar', 'area', 'scatter', 'doughnut', 'combo']) await p.click(`[data-chart="${t}"]`);
  const drawn = await p.evaluate(() => [...document.querySelectorAll('#chartsLayer canvas')].map(c => { const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let n = 0; for (let i = 0; i < d.length; i += 4) if (d[i] < 240 || d[i + 1] < 240 || d[i + 2] < 240) n++; return n; }));
  check('all five new charts draw', drawn.length === 5 && drawn.every(n => n > 500), drawn.join());
  // Rename sheet with a space to exercise quoting
  await p.evaluate(() => { AtlasSheets.book.sheets[0].name = 'My Data'; });
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  fs.writeFileSync(SP + '/files/charts-out.xlsx', Buffer.from(bytes));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const names = Object.keys(zip.files);
  const c1 = await zip.file('xl/charts/chart1.xml').async('string'), c3 = await zip.file('xl/charts/chart3.xml').async('string'), c5 = await zip.file('xl/charts/chart5.xml').async('string');
  const sx = await zip.file('xl/worksheets/sheet1.xml').async('string'), ct = await zip.file('[Content_Types].xml').async('string');
  const rels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');
  check('xlsx: five chart parts, a drawing and relationships', names.filter(n => /^xl\/charts\/chart\d+\.xml$/.test(n)).length === 5 && names.includes('xl/drawings/drawing1.xml') && /<drawing r:id="rIdAtlasDrawing"\/>/.test(sx) && /drawings\/drawing1\.xml/.test(rels) && (ct.match(/drawingml\.chart\+xml/g) || []).length === 5);
  check('chart 1: columns with Excel references and cached values', /<c:barChart><c:barDir val="col"\/>/.test(c1) && /<c:f>'My Data'!\$B\$2:\$B\$4<\/c:f>/.test(c1) && /<c:f>'My Data'!\$A\$2:\$A\$4<\/c:f>/.test(c1) && /<c:v>40<\/c:v>/.test(c1), (c1.match(/<c:ser>[\s\S]{0,300}/) || [''])[0]);
  check('chart 3: scatter with x and y values', /<c:scatterChart>/.test(c3) && /<c:xVal>/.test(c3));
  check('chart 5: combo is columns plus a line', /<c:barChart>[\s\S]*<\/c:barChart><c:lineChart>/.test(c5));
  const { execSync } = require('child_process');
  let wellFormed = true;
  for (const n of names.filter(n => /\.(xml|rels)$/.test(n))) { const x = await zip.file(n).async('string'); fs.writeFileSync(SP + '/files/_wf.xml', x); try { execSync(`python3 -I -c "import xml.dom.minidom,sys;xml.dom.minidom.parse(sys.argv[1])" ${SP}/files/_wf.xml`); } catch { wellFormed = false; console.log('bad xml', n); } }
  check('every part is well-formed XML', wellFormed);
  // Reopen without Atlas metadata: charts come from the Excel parts
  const z2 = await JSZip.loadAsync(Buffer.from(bytes));
  z2.remove('docProps/custom.xml');
  z2.file('[Content_Types].xml', (await z2.file('[Content_Types].xml').async('string')).replace(/<Override[^>]*custom\.xml[^>]*\/>/, ''));
  z2.file('_rels/.rels', (await z2.file('_rels/.rels').async('string')).replace(/<Relationship[^>]*custom\.xml[^>]*\/>/, ''));
  const plain = await z2.generateAsync({ type: 'nodebuffer' });
  const back = await p.evaluate(async b => { const before = AtlasSheets.book.sheets[0].charts.map(c => [c.type, c.range, c.x, c.y, c.w, c.h]); await AtlasSheets.loadFile({ name: 'c.xlsx', path: null, data: new Uint8Array(b) }); return { before, after: AtlasSheets.book.sheets[0].charts.map(c => [c.type, c.range, c.x, c.y, c.w, c.h]) }; }, Array.from(plain));
  const near = (a, b) => a[0] === b[0] && a[1] === b[1] && a.slice(2).every((v, i) => Math.abs(v - b[i + 2]) <= 1);
  check('Excel charts read back with type, data and position', back.after.length === 5 && back.after.every((c, i) => near(c, c[0] === 'doughnut' ? [...back.before[i].slice(0, 1), 'A1:B4', ...back.before[i].slice(2)] : back.before[i])), JSON.stringify(back));
  // openpyxl workbook
  await p.evaluate(async b => { await AtlasSheets.loadFile({ name: 'o.xlsx', path: null, data: new Uint8Array(b) }); }, Array.from(fs.readFileSync(SP + '/files/chart-openpyxl.xlsx')));
  const o = await p.evaluate(() => AtlasSheets.book.sheets.map(s => s.charts.map(c => `${c.type}:${c.range}:${c.title}`).join(' | ')));
  check('openpyxl charts open as Atlas charts', o[0] === 'bar:A1:C5:Sales by month | pie:A1:B5:North share' && o[1] === 'scatter:A2:B5:XY', JSON.stringify(o));
  const shown = await p.evaluate(() => document.querySelectorAll('#chartsLayer .chart').length);
  check('and are drawn', shown === 2);
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
