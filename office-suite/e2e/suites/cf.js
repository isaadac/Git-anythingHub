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
    for (let i = 1; i <= 10; i++) {
      AtlasSheets.setCell('A' + i, String(i)); AtlasSheets.setCell('B' + i, String(i * 10));
      AtlasSheets.setCell('C' + i, ['apple', 'pear', 'apple', 'fig', 'kiwi', 'pear', 'plum', 'lime', 'date', 'banana'][i - 1]);
      AtlasSheets.setCell('D' + i, String(100 - i * 7));
    }
  });
  const sel = (a, z) => p.evaluate(([a, z]) => { const P = s => { const m = /^([A-Z]+)(\d+)$/.exec(s); return [+m[2] - 1, m[1].charCodeAt(0) - 65]; }; const [r1, c1] = P(a), [r2, c2] = P(z); AtlasSheets.select(r1, c1); AtlasSheets.select(r2, c2, true); }, [a, z]);
  const answer = list => p.evaluate(list => { window.__answers = list.slice(); UI.form = async () => window.__answers.shift(); }, list);
  const cmd = v => p.selectOption('#condFmt', v);
  const td = ref => p.evaluate(ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); const t = document.querySelector(`#grid tr[data-r="${+m[2] - 1}"] td[data-c="${m[1].charCodeAt(0) - 65}"]`); return t ? { bg: t.style.backgroundColor, color: t.style.color, fw: t.style.fontWeight, img: t.style.backgroundImage, icon: t.querySelector('.cf-icon')?.textContent || '', iconColor: t.querySelector('.cf-icon')?.style.color || '', text: t.textContent } : null; }, ref);
  // Greater than 7 on A
  await sel('A1', 'A10'); await answer([{ v: '7', style: 'red' }]); await cmd('gt');
  const a = await Promise.all(['A7', 'A8', 'A10'].map(td));
  check('Greater Than highlights only matching cells', a[0].bg === '' && a[1].bg === 'rgb(255, 199, 206)' && a[1].color === 'rgb(156, 0, 6)' && a[2].bg === 'rgb(255, 199, 206)', JSON.stringify(a.map(x => x.bg)));
  // Recalculation: change A1 to 50 -> highlighted
  await p.evaluate(() => AtlasSheets.setCell('A1', '50'));
  check('rules follow edits', (await td('A1')).bg === 'rgb(255, 199, 206)');
  await p.evaluate(() => AtlasSheets.setCell('A1', '1'));
  // Between on D with formula operand
  await sel('D1', 'D10'); await answer([{ a: '30', b: '=B3*2', style: 'green' }]); await cmd('between');
  const d = await Promise.all(['D1', 'D6', 'D10'].map(td)); // 93, 58, 30
  check('Between with a formula bound', d[0].bg === '' && d[1].bg === 'rgb(198, 239, 206)' && d[2].bg === 'rgb(198, 239, 206)', JSON.stringify(d.map(x => x.bg)));
  // Duplicates and text
  await sel('C1', 'C10'); await answer([{ k: 'duplicateValues', style: 'yellow' }]); await cmd('dup');
  await answer([{ op: 'beginsWith', t: 'b', style: 'bold' }]); await cmd('text');
  const c = await Promise.all(['C1', 'C2', 'C4', 'C10'].map(td));
  check('Duplicate values and Text begins with', c[0].bg === 'rgb(255, 235, 156)' && c[1].bg === 'rgb(255, 235, 156)' && c[2].bg === '' && c[3].fw === '700' && c[3].bg === '', JSON.stringify(c.map(x => x.bg + '/' + x.fw)));
  // Top 3 and above average on B
  await sel('B1', 'B10'); await answer([{ n: '3', style: 'redText' }]); await cmd('top');
  await answer([{ style: 'bold' }]); await cmd('below');
  const bb = await Promise.all(['B1', 'B7', 'B8', 'B10'].map(td));
  check('Top 3 and Below Average', bb[0].fw === '700' && bb[1].color === '' && bb[2].color === 'rgb(255, 0, 0)' && bb[3].color === 'rgb(255, 0, 0)' && bb[3].fw === '', JSON.stringify(bb.map(x => x.color + '/' + x.fw)));
  // Data bars, colour scale, icons on new columns
  await p.evaluate(() => { for (let i = 1; i <= 5; i++) { AtlasSheets.setCell('E' + i, String(i * 20)); AtlasSheets.setCell('F' + i, String(i)); AtlasSheets.setCell('G' + i, String(i)); } });
  await sel('E1', 'E5'); await cmd('bar:#638ec6');
  await sel('F1', 'F5'); await cmd('scale:#63be7b,#ffeb84,#f8696b');
  await sel('G1', 'G5'); await cmd('icons:3Arrows');
  const e = await Promise.all(['E1', 'E5', 'F1', 'F3', 'F5', 'G1', 'G3', 'G5'].map(td));
  check('Data bars scale with the value', /20%/.test(e[0].img) && /100%/.test(e[1].img), e[0].img.slice(0, 90));
  check('Color scale: min green, middle yellow, max red', e[2].bg === 'rgb(99, 190, 123)' && e[3].bg === 'rgb(255, 235, 132)' && e[4].bg === 'rgb(248, 105, 107)', [e[2].bg, e[3].bg, e[4].bg].join());
  check('Icon set arrows', e[5].icon === '▼' && e[6].icon === '►' && e[7].icon === '▲' && e[7].text === '▲5', e.slice(5).map(x => x.icon).join());
  // Formula rule across rows
  await sel('H1', 'H10'); await p.evaluate(() => { for (let i = 1; i <= 10; i++) AtlasSheets.setCell('H' + i, 'row ' + i); });
  await answer([{ f: '=$A1>8', style: 'green' }]); await cmd('formula');
  const h = await Promise.all(['H8', 'H9', 'H10'].map(td));
  check('Formula rule is relative to each row', h[0].bg === '' && h[1].bg === 'rgb(198, 239, 206)' && h[2].bg === 'rgb(198, 239, 206)', h.map(x => x.bg).join());
  const n0 = await p.evaluate(() => AtlasSheets.book.sheets[0].cf.length);
  // xlsx export
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  fs.writeFileSync(SP + '/files/cf-out.xlsx', Buffer.from(bytes));
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const sx = await zip.file('xl/worksheets/sheet1.xml').async('string'), stx = await zip.file('xl/styles.xml').async('string');
  const nCf = (sx.match(/<cfRule /g) || []).length;
  check('xlsx: every rule written as an Excel conditional format', nCf === n0 && /<cfRule type="cellIs" dxfId="\d+" priority="\d+" operator="greaterThan"><formula>7<\/formula>/.test(sx) && /<cfRule type="expression"[^>]*><formula>\$A1&gt;8<\/formula>/.test(sx) && /<dataBar>/.test(sx) && /<iconSet iconSet="3Arrows">/.test(sx), nCf + '/' + n0);
  check('xlsx: differential styles in styles.xml', /<dxfs count="\d+">/.test(stx) && /<dxf><font><b\/><\/font><\/dxf>/.test(stx) && /<bgColor rgb="FFFFC7CE"\/>/.test(stx));
  check('xlsx: conditional formatting placed before pageMargins/ignoredErrors', sx.indexOf('<conditionalFormatting') > sx.indexOf('</sheetData>') && (sx.indexOf('<pageMargins') === -1 || sx.indexOf('<conditionalFormatting') < sx.indexOf('<pageMargins')) && (sx.indexOf('<ignoredErrors') === -1 || sx.indexOf('</conditionalFormatting>') < sx.indexOf('<ignoredErrors')));
  // Reopen (Atlas metadata path)
  const back = await p.evaluate(async b => { const before = JSON.stringify(AtlasSheets.book.sheets[0].cf); await AtlasSheets.loadFile({ name: 'cf.xlsx', path: null, data: new Uint8Array(b) }); return { same: JSON.stringify(AtlasSheets.book.sheets[0].cf) === before, n: AtlasSheets.book.sheets[0].cf.length }; }, bytes);
  check('reopened .xlsx keeps every rule exactly', back.same, JSON.stringify(back));
  // Excel path: drop the Atlas metadata
  const z2 = await JSZip.loadAsync(Buffer.from(bytes));
  z2.remove('docProps/custom.xml');
  const ct = await z2.file('[Content_Types].xml').async('string'); z2.file('[Content_Types].xml', ct.replace(/<Override[^>]*custom\.xml[^>]*\/>/, ''));
  const rels = await z2.file('_rels/.rels').async('string'); z2.file('_rels/.rels', rels.replace(/<Relationship[^>]*custom\.xml[^>]*\/>/, ''));
  const noMeta = await z2.generateAsync({ type: 'nodebuffer' });
  const ex = await p.evaluate(async ([b, before]) => { await AtlasSheets.loadFile({ name: 'plain.xlsx', path: null, data: new Uint8Array(b) }); const now = AtlasSheets.book.sheets[0].cf; return { n: now.length, same: JSON.stringify(now) === before, now: JSON.stringify(now).slice(0, 300) }; }, [Array.from(noMeta), JSON.stringify((await p.evaluate(() => null)) || null)]);
  const cmpRules = await p.evaluate(async b => {
    const plain = AtlasSheets.book.sheets[0].cf.map(r => JSON.stringify(r));
    return plain;
  }, null);
  check('Excel-only .xlsx: rules read from the conditional formats', ex.n === n0, JSON.stringify(ex));
  const h2 = await Promise.all(['A8', 'H9', 'G5', 'C1'].map(td));
  check('Excel-only .xlsx renders the same', h2[0].bg === 'rgb(255, 199, 206)' && h2[1].bg === 'rgb(198, 239, 206)' && h2[2].icon === '▲' && h2[3].bg === 'rgb(255, 235, 156)', JSON.stringify(h2.map(x => x.bg + x.icon)));
  // openpyxl-written workbook
  await p.evaluate(async b => { await AtlasSheets.loadFile({ name: 'openpyxl.xlsx', path: null, data: new Uint8Array(b) }); }, Array.from(fs.readFileSync(SP + '/files/cf-openpyxl.xlsx')));
  const o = await p.evaluate(() => AtlasSheets.book.sheets[0].cf.map(r => r.type + '@' + r.ref).join());
  const ot = await Promise.all(['A8', 'A7', 'B10', 'B1', 'C1', 'C4', 'D9', 'D8'].map(td));
  check('openpyxl workbook: rules in priority order', o === 'cellIs@A1:A10,colorScale@B1:B10,dataBar@A1:A10,iconSet@B1:B10,duplicateValues@C1:C10,expression@D1:D10', o);
  check('openpyxl workbook renders', ot[0].bg === 'rgb(255, 199, 206)' && ot[0].fw === '700' && ot[1].bg === '' && /linear-gradient/.test(ot[1].img) && ot[2].bg === 'rgb(99, 190, 123)' && ot[2].icon === '▲' && ot[3].bg === 'rgb(248, 105, 107)' && ot[4].bg === 'rgb(255, 235, 156)' && ot[5].bg === '' && ot[6].bg === 'rgb(198, 239, 206)' && ot[7].bg === '', JSON.stringify(ot.map(x => x.bg + '|' + x.icon)));
  // Insert a row above: ranges and formulas move
  await sel('A1', 'A1'); await p.click('#insRow');
  const shifted = await p.evaluate(() => AtlasSheets.book.sheets[0].cf.map(r => r.ref + (r.formula ? ' ' + r.formula : '')).join());
  check('inserting a row moves the rules', /^A2:A11,B2:B11,A2:A11,B2:B11,C2:C11,D2:D11 \$A2>8$/.test(shifted), shifted);
  check('formula rule still right after the move', (await td('D10')).bg === 'rgb(198, 239, 206)' && (await td('D9')).bg === '');
  await p.evaluate(() => document.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', ctrlKey: true, bubbles: true })));
  // Manage: delete the top rule, move one
  await p.click('#cfManage');
  const rows = await p.evaluate(() => [...document.querySelectorAll('.cf-manage .cf-row .cf-desc')].map(x => x.textContent));
  check('Manage Rules lists the rules', rows.length === 6 && rows[0] === 'Cell value > 7', rows.join(' | '));
  await p.click('.cf-manage .cf-row:first-child button[data-a="del"]');
  await p.fill('.cf-manage .cf-row:first-child input', 'B1:B5'); await p.dispatchEvent('.cf-manage .cf-row:first-child input', 'change');
  await p.click('.cf-manage .ok');
  const after = await p.evaluate(() => AtlasSheets.book.sheets[0].cf.map(r => r.type + '@' + r.ref).join());
  check('Manage Rules edits apply', /^colorScale@B1:B5,dataBar@A/.test(after), after);
  // Clear from selection
  await sel('C1', 'C3'); await cmd('clear-sel');
  check('Clear Rules from Selected Cells', !(await p.evaluate(() => AtlasSheets.book.sheets[0].cf.some(r => r.type === 'duplicateValues'))));
  // Hostile rules are dropped
  const hostile = await p.evaluate(() => [SheetsCF.clean({ ref: 'A1:B2', type: 'iconSet', set: '__proto__' }).set, SheetsCF.clean({ ref: 'A1;alert(1)', type: 'dataBar' }), SheetsCF.clean({ ref: 'A1', type: 'cellIs', op: 'constructor' }), SheetsCF.clean({ ref: 'A1', type: 'cellIs', op: 'equal', v1: 'x', style: { bg: 'red;background:url(x)' } }).style.bg]);
  check('invalid rules and styles are rejected', hostile[0] === '3TrafficLights1' && hostile[1] === null && hostile[2] === null && hostile[3] === undefined, JSON.stringify(hostile));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
