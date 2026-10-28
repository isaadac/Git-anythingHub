const { chromium } = require('playwright');
const fs = require('fs'), path = require('path'), { execFileSync } = require('child_process');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const SP = process.env.E2E_WORK;
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);
  await p.evaluate(() => {
    const S = AtlasSheets.book.sheets[0].cells;
    AtlasSheets.setCell('A1', 'Header'); AtlasSheets.setCell('B1', 1234.5); AtlasSheets.setCell('C1', 0.25); AtlasSheets.setCell('D1', 'wrapped text here');
    S.A1.s = { b: true, i: true, u: true, font: 'Arial', size: 14, color: '#c00000', bg: '#ffff00', align: 'center', border: { t: 1, b: 1, l: 1, r: 1 } };
    S.B1.s = { fmt: '#,##0.00', bg: '#d9e1f2' };
    S.C1.s = { fmt: '0%', align: 'right', valign: 'top' };
    S.D1.s = { wrap: true, unlocked: true };
    S.E5 = { v: '', s: { bg: '#00b050' } }; // formatting only, no value
    AtlasSheets.book.sheets[0].freezeRow = true;
  });
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  const file = path.join(SP, 'styled.xlsx');
  fs.writeFileSync(file, Buffer.from(bytes));
  const py = `
import openpyxl, json
wb = openpyxl.load_workbook(r'${file}')
ws = wb.active
def c(ref):
    x = ws[ref]
    f = x.font; fl = x.fill; bd = x.border; al = x.alignment
    rgb = lambda col: (col.rgb if col is not None and isinstance(col.rgb, str) else None)
    return dict(b=f.b, i=f.i, u=f.u, name=f.name, sz=f.sz, color=rgb(f.color), fill=rgb(fl.fgColor) if fl.fill_type == 'solid' else None,
                bt=bool(bd.top.style), bb=bool(bd.bottom.style), bl=bool(bd.left.style), br=bool(bd.right.style), h=al.horizontal, v=al.vertical, wrap=al.wrap_text,
                fmt=x.number_format, locked=x.protection.locked)
print(json.dumps(dict(A1=c('A1'), B1=c('B1'), C1=c('C1'), D1=c('D1'), E5=c('E5'), freeze=ws.freeze_panes)))
`;
  const r = JSON.parse(execFileSync('python3', ['-c', py]).toString());
  check('Excel sees bold/italic/underline, font and size', r.A1.b && r.A1.i && r.A1.u === 'single' && r.A1.name === 'Arial' && r.A1.sz === 14, JSON.stringify(r.A1));
  check('Excel sees font colour and fill', r.A1.color === 'FFC00000' && r.A1.fill === 'FFFFFF00');
  check('Excel sees borders and alignment', r.A1.bt && r.A1.bb && r.A1.bl && r.A1.br && r.A1.h === 'center');
  check('number formats with fill', r.B1.fmt === '#,##0.00' && r.B1.fill === 'FFD9E1F2' && r.C1.fmt === '0%', JSON.stringify([r.B1.fmt, r.C1.fmt]));
  check('right/top alignment', r.C1.h === 'right' && r.C1.v === 'top');
  check('wrap and unlocked', r.D1.wrap === true && r.D1.locked === false && r.A1.locked === true);
  check('formatting-only cell kept', r.E5.fill === 'FF00B050');
  check('frozen header row', r.freeze === 'A2', r.freeze);
  // An Excel-made workbook (openpyxl), opened in Sheets
  const src = path.join(SP, 'excel-styled.xlsx');
  execFileSync('python3', ['-c', `
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
wb = openpyxl.Workbook(); ws = wb.active
ws['A1'] = 'Title'; ws['A1'].font = Font(bold=True, size=16, name='Georgia', color='1F4E79')
ws['B2'] = 5; ws['B2'].fill = PatternFill('solid', fgColor='FCE4D6'); ws['B2'].number_format = '0.0'
thin = Side(style='thin'); ws['C3'] = 'box'; ws['C3'].border = Border(left=thin, right=thin, top=thin, bottom=thin)
ws['D4'] = 'mid'; ws['D4'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
ws['E5'].fill = PatternFill('solid', fgColor='5B9BD5')
from openpyxl.styles.colors import Color
ws['F6'] = 'theme'; ws['F6'].font = Font(color=Color(theme=4, tint=-0.25))
ws.freeze_panes = 'B2'
wb.save(r'${src}')
`]);
  await p.evaluate(async b64 => { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); await AtlasSheets.loadFile({ name: 'excel-styled.xlsx', path: null, data: u }); }, fs.readFileSync(src).toString('base64'));
  const got = await p.evaluate(() => { const s = AtlasSheets.book.sheets[0]; const c = k => (s.cells[k] || {}).s || null; return { A1: c('A1'), B2: c('B2'), C3: c('C3'), D4: c('D4'), E5: c('E5'), F6: c('F6'), fr: s.freezeRow, fc: s.freezeCol }; });
  check('Excel font read', got.A1 && got.A1.b && got.A1.size === 16 && got.A1.font === 'Georgia' && got.A1.color === '#1f4e79', JSON.stringify(got.A1));
  check('Excel fill and number format read', got.B2 && got.B2.bg === '#fce4d6' && got.B2.fmt === '0.0', JSON.stringify(got.B2));
  check('Excel borders read', got.C3 && JSON.stringify(got.C3.border) === '{"t":1,"l":1,"r":1,"b":1}', JSON.stringify(got.C3));
  check('Excel alignment and wrap read', got.D4 && got.D4.align === 'center' && got.D4.valign === 'middle' && got.D4.wrap === true, JSON.stringify(got.D4));
  check('Excel formatting-only cell read', got.E5 && got.E5.bg === '#5b9bd5', JSON.stringify(got.E5));
  check('Excel theme colour with tint read', got.F6 && got.F6.color === '#3355' + '93' || (got.F6 && /^#3[0-9a-f]{5}$/.test(got.F6.color)), JSON.stringify(got.F6));
  check('Excel frozen panes read', got.fr && got.fc);
  const shown = await p.evaluate(() => { const td = document.querySelector('#grid tr[data-r="0"] td[data-c="0"]'); return [getComputedStyle(td).fontWeight, getComputedStyle(td).fontFamily]; });
  check('formatting drawn in the grid', shown[0] === '700' && /Georgia/.test(shown[1]), shown.join());
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
