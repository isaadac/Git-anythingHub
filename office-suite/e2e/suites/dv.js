const { chromium } = require('playwright');
const JSZip = require('/home/user/Git-anythingHub/office-suite/node_modules/jszip');
const fs = require('fs');
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  const check = (n, c, x = '') => console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : ''));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  const sel = (r1, c1, r2 = r1, c2 = c1) => p.evaluate(([r1, c1, r2, c2]) => { AtlasSheets.select(r1, c1); if (r2 !== r1 || c2 !== c1) AtlasSheets.select(r2, c2, true); }, [r1, c1, r2, c2]);
  const type = async (text) => { await p.focus('#gridWrap'); await p.keyboard.type(text); await p.keyboard.press('Enter'); await p.waitForTimeout(120); };
  const val = ref => p.evaluate(ref => AtlasSheets.valueOf(ref), ref);
  const dialog = () => p.evaluate(() => { const m = document.querySelector('.modal[role=alertdialog]'); return m ? m.querySelector('h2').textContent + ' | ' + m.querySelector('p').textContent : null; });
  const press = label => p.evaluate(label => [...document.querySelectorAll('.modal[role=alertdialog] button')].find(x => x.textContent === label).click(), label);
  // Rule 1: list on A1:A5 with a prompt (via the dialog)
  await sel(0, 0, 4, 0);
  await p.click('[data-tab="data"]'); await p.click('#dataValidation');
  await p.selectOption('.dv-dialog [name=type]', 'list');
  await p.fill('.dv-dialog [name=src]', 'Yes, No, Maybe');
  await p.click('.dv-tabs button[data-p="1"]');
  await p.check('.dv-dialog [name=showPrompt]'); await p.fill('.dv-dialog [name=ptitle]', 'Answer'); await p.fill('.dv-dialog [name=ptext]', 'Choose Yes, No or Maybe');
  await p.click('.dv-dialog .ok');
  await sel(1, 0);
  const ui = await p.evaluate(() => ({ arrow: !document.getElementById('dvArrow').hidden, prompt: !document.getElementById('dvPrompt').hidden && document.getElementById('dvPrompt').textContent }));
  check('list cell shows the dropdown arrow and input message', ui.arrow && ui.prompt === 'AnswerChoose Yes, No or Maybe', JSON.stringify(ui));
  await p.click('#dvArrow');
  const items = await p.evaluate(() => [...document.querySelectorAll('#dvList div')].map(d => d.textContent));
  check('dropdown lists the items', items.join() === 'Yes,No,Maybe', items.join());
  await p.evaluate(() => [...document.querySelectorAll('#dvList div')][2].dispatchEvent(new MouseEvent('mousedown', { bubbles: true })));
  check('picking an item fills the cell', (await val('A2')) === 'Maybe');
  // Keyboard: Alt+Down, ArrowDown, Enter
  await sel(2, 0); await p.focus('#gridWrap'); await p.keyboard.press('Alt+ArrowDown'); await p.keyboard.press('ArrowDown'); await p.keyboard.press('ArrowDown'); await p.keyboard.press('Enter');
  check('Alt+Down opens the list and Enter picks', (await val('A3')) === 'No', String(await val('A3')));
  // Typing a value not in the list: Stop -> Retry reopens the editor with the text
  await sel(0, 0); await type('Perhaps');
  const d1 = await dialog();
  check('Stop alert refuses a value outside the list', /The value doesn't match the list/.test(d1 || '') && (await val('A1')) === '', d1);
  await press('Retry');
  const reopened = await p.evaluate(() => ({ editing: !document.getElementById('cellEditor').hidden, text: document.getElementById('cellEditor').value }));
  check('Retry reopens the cell with what was typed', reopened.editing && reopened.text === 'Perhaps', JSON.stringify(reopened));
  await p.keyboard.press('Escape');
  await sel(0, 0); await type('yes');
  check('list match ignores case', (await val('A1')) === 'yes' && !(await dialog()));
  // Rule 2: whole number between 1 and =C1 (formula bound), warning style
  await p.evaluate(() => AtlasSheets.setCell('C1', '10'));
  await sel(1, 1, 5, 1); await p.click('#dataValidation');
  await p.selectOption('.dv-dialog [name=type]', 'whole'); await p.selectOption('.dv-dialog [name=op]', 'between');
  await p.fill('.dv-dialog [name=v1]', '1'); await p.fill('.dv-dialog [name=v2]', '=$C$1');
  await p.click('.dv-tabs button[data-p="2"]'); await p.selectOption('.dv-dialog [name=estyle]', 'warning'); await p.fill('.dv-dialog [name=etext]', 'Use 1 to C1');
  await p.click('.dv-dialog .ok');
  await sel(1, 1); await type('12');
  const d2 = await dialog();
  check('Warning asks to continue', /Use 1 to C1[\s\S]*Continue\?/.test(d2 || ''), d2);
  await press('Yes');
  check('Yes keeps the value', (await val('B2')) === 12);
  await sel(2, 1); await type('5.5'); await press('No'); await p.keyboard.press('Escape');
  check('No refuses it', (await val('B3')) === '');
  await sel(3, 1); await type('7');
  check('valid values need no alert', (await val('B4')) === 7 && !(await dialog()));
  // Undo after a refused value doesn't skip the user's earlier edit
  // Circle invalid data
  await p.evaluate(() => AtlasSheets.setCell('B5', '99'));
  await p.click('#circleInvalid');
  const bad = await p.evaluate(() => [...document.querySelectorAll('#grid td.dv-bad')].map(t => t.closest('tr').dataset.r + ',' + t.dataset.c));
  check('Circle Invalid marks the cells that break the rules', bad.sort().join(' ') === '1,1 4,1', bad.join(' '));
  await p.click('#circleInvalid');
  // xlsx export
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  const zip = await JSZip.loadAsync(Buffer.from(bytes)); const sx = await zip.file('xl/worksheets/sheet1.xml').async('string');
  check('xlsx: Excel data validations', /<dataValidations count="2">/.test(sx) && /<dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorStyle="warning" error="Use 1 to C1" sqref="B2:B6"><formula1>1<\/formula1><formula2>\$C\$1<\/formula2>/.test(sx) && /type="list"[^>]*promptTitle="Answer" prompt="Choose Yes, No or Maybe" sqref="A1:A5"><formula1>&quot;Yes,No,Maybe&quot;<\/formula1>/.test(sx), (sx.match(/<dataValidations[\s\S]*<\/dataValidations>/) || [''])[0].slice(0, 300));
  check('xlsx: dataValidations after conditional formatting, before pageMargins', sx.indexOf('<dataValidations') > sx.indexOf('</sheetData>') && (sx.indexOf('<pageMargins') < 0 || sx.indexOf('</dataValidations>') < sx.indexOf('<pageMargins')));
  fs.writeFileSync(SP + '/files/dv-out.xlsx', Buffer.from(bytes));
  const back = await p.evaluate(async b => { const before = JSON.stringify(AtlasSheets.book.sheets[0].dv); await AtlasSheets.loadFile({ name: 'dv.xlsx', path: null, data: new Uint8Array(b) }); return JSON.stringify(AtlasSheets.book.sheets[0].dv) === before; }, bytes);
  check('reopened .xlsx keeps the rules exactly', back);
  // Excel-written (openpyxl) workbook
  await p.evaluate(async b => { await AtlasSheets.loadFile({ name: 'x.xlsx', path: null, data: new Uint8Array(b) }); }, Array.from(fs.readFileSync(SP + '/files/dv-openpyxl.xlsx')));
  const o = await p.evaluate(() => AtlasSheets.book.sheets[0].dv.map(r => `${r.type}@${r.ref}:${r.v1}${r.v2 ? '..' + r.v2 : ''}/${r.error.style}`).join(' '));
  check('openpyxl workbook: rules read', o === 'list@A1:A10:Yes,No,Maybe/stop list@B1:B10:=Lists!$A$1:$A$3/warning whole@C1:C10:1..10/stop date@D1:D10:2024-01-01/stop custom@E1:E10:=LEN(E1)<=5/information', o);
  await sel(0, 1); await p.click('#dvArrow');
  const colours = await p.evaluate(() => [...document.querySelectorAll('#dvList div')].map(d => d.textContent).join());
  check('a list from another sheet fills the dropdown', colours === 'Red,Green,Blue', colours);
  await p.keyboard.press('Escape');
  await sel(0, 3); await type('2023-12-01');
  const d3 = await dialog();
  check('date rule refuses an earlier date', /must be a date greater than 2024-01-01/.test(d3 || ''), d3);
  await press('Cancel');
  await sel(0, 4); await type('toolongtext');
  const d4 = await dialog();
  check('custom formula rule (Information) relative to the cell', /doesn't meet the rule/.test(d4 || ''), d4);
  await press('OK');
  check('Information OK keeps the value', (await val('E1')) === 'toolongtext');
  // Excel's x14 list (lists on other sheets in newer files)
  const x14 = await p.evaluate(() => SheetsDV.fromXlsx('<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>F2:F20 H2:H5</xm:sqref></x14:dataValidation></x14:dataValidations>'));
  check('Excel x14 data validations are read', x14.length === 2 && x14[0].ref === 'F2:F20' && x14[1].ref === 'H2:H5' && x14[0].v1 === '=Lists!$A$1:$A$3', JSON.stringify(x14));
  // Inserting a row moves the rule
  await sel(0, 0); await p.click('[data-tab="home"]'); await p.click('#insRow');
  const moved = await p.evaluate(() => AtlasSheets.book.sheets[0].dv.map(r => r.ref + (r.type === 'custom' ? ' ' + r.v1 : '')).join());
  check('rules move with inserted rows', moved === 'A2:A11,B2:B11,C2:C11,D2:D11,E2:E11 =LEN(E2)<=5', moved);
  const flags = await p.evaluate(() => SheetsDV.fromXlsx('<dataValidation type="whole" operator="greaterThan" sqref="A1" promptTitle="T" prompt="P"><formula1>5</formula1></dataValidation>')[0]);
  check('Excel flags: no alert and no input message unless turned on', flags.error.show === false && !flags.prompt, JSON.stringify(flags));
  const hostile = await p.evaluate(() => [SheetsDV.clean({ ref: 'A1', type: 'eval', v1: 'x' }), SheetsDV.clean({ ref: 'A1', type: 'list', v1: 'a', error: { style: '<b>' } }).error.style, SheetsDV.fromXlsx('<dataValidation type="list" sqref="A1"><formula1>"x"</formula1></dataValidation><!ENTITY x "y">').length]);
  check('bad rules rejected', hostile[0] === null && hostile[1] === 'stop', JSON.stringify(hostile));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})();
