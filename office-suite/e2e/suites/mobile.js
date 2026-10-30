const { chromium, devices } = require('playwright');
const path = require('path');
const ROOT = process.cwd();
const D = require(path.join(ROOT, 'node_modules/docx'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const B = process.env.E2E_BASE + '/src/renderer/';
(async () => {
  const b = await chromium.launch();
  const errors = [];
  // ---- Phone browser
  const ctx = await b.newContext({ ...devices['Pixel 7'] });
  const p = await ctx.newPage();
  p.on('pageerror', e => errors.push(e.message));
  for (const app of ['index', 'write', 'sheets', 'slides']) {
    await p.goto(B + app + '.html'); await sleep(700);
    const over = await p.evaluate(() => document.documentElement.scrollWidth - innerWidth);
    check(`${app}: fits the phone screen (no sideways page scroll)`, over <= 1, over);
    await p.screenshot({ path: `${process.env.E2E_WORK}/shots/phone-${app}.png` });
  }
  await p.goto(B + 'write.html'); await sleep(800);
  const pw = await p.evaluate(() => document.getElementById('page').getBoundingClientRect().width);
  check('write: text reflows to the phone width', pw <= 412, pw);
  // slides: drag a shape with a finger
  await p.goto(B + 'slides.html'); await sleep(700);
  await p.evaluate(() => { const s = AtlasSlides.pres.slides[0]; s.elements.push({ id: 'sh1', type: 'shape', shape: 'rect', fill: '#4472c4', line: '#2f5496', html: '', x: 100, y: 100, w: 200, h: 120 }); AtlasSlides.goTo(0); });
  const box = await (await p.$('#stage .el[data-id="sh1"]')).boundingBox();
  const cdp = await ctx.newCDPSession(p);
  const tp = (type, x, y) => cdp.send('Input.dispatchTouchEvent', { type, touchPoints: type === 'touchEnd' ? [] : [{ x, y }] });
  const x0 = box.x + box.width / 2, y0 = box.y + box.height / 2;
  await tp('touchStart', x0, y0);
  for (let i = 1; i <= 6; i++) { await tp('touchMove', x0 + i * 10, y0 + i * 5); await sleep(16); }
  await tp('touchEnd');
  await sleep(200);
  const moved = await p.evaluate(() => AtlasSlides.pres.slides[0].elements.find(e => e.id === 'sh1'));
  check('slides: a shape can be dragged by touch', moved.x > 100 && moved.y > 100, `${moved.x},${moved.y}`);
  // sheets: tap selects a cell, swipe scrolls the grid
  await p.goto(B + 'sheets.html'); await sleep(700);
  await p.tap('#grid tbody tr:nth-child(3) td:nth-child(3)');
  check('sheets: tapping a cell selects it', (await p.inputValue('.namebox')) === 'B3', await p.inputValue('.namebox'));

  // ---- The Android app (a stand-in for the Capacitor bridge)
  const actx = await b.newContext({ ...devices['Pixel 7'] });
  const report = await D.Packer.toBuffer(new D.Document({ sections: [{ children: [new D.Paragraph('Opened from the Files app')] }] }));
  await actx.addInitScript(b64 => {
    const calls = window.__calls = [];
    const listeners = {};
    window.Capacitor = {
      isNativePlatform: () => true,
      getPlatform: () => sessionStorage.getItem('plat') || 'android',
      Plugins: {
        IsadekPrint: {
          print: async o => { calls.push(['print', o.name, !!o.landscape]); },
          pdf: async o => { calls.push(['pdf', !!o.landscape]); return { data: btoa('%PDF-1.4 test') }; }
        },
        Filesystem: {
          writeFile: async o => { calls.push(['write', o.path, o.directory, o.data.length]); sessionStorage.setItem('fs:' + o.path, o.data); return { uri: 'file:///docs/' + o.path }; },
          readFile: async o => { calls.push(['read', o.path]); return { data: o.path.startsWith('content://') ? b64 : sessionStorage.getItem('fs:' + o.path) }; },
          getUri: async o => ({ uri: 'file:///docs/' + o.path })
        },
        Share: { share: async o => { calls.push(['share', o.files[0]]); } },
        App: {
          addListener: (ev, fn) => { listeners[ev] = fn; },
          getLaunchUrl: async () => (sessionStorage.getItem('launched') ? null : (sessionStorage.setItem('launched', '1'), { url: 'content://com.android.providers.downloads.documents/document/msf%3A1234' })),
          exitApp: () => calls.push(['exit'])
        }
      }
    };
    window.__fire = (ev, data) => listeners[ev] && listeners[ev](data);
  }, report.toString('base64'));
  const a = await actx.newPage();
  a.on('pageerror', e => errors.push('app: ' + e.message));
  await a.goto(B + 'index.html');
  await a.waitForURL(/write\.html/, { timeout: 8000 });
  await a.waitForFunction(() => document.getElementById('doc').innerText.includes('Opened from the Files app'));
  check('app: a file opened with "Open with" lands in the right app (type sniffed from a content:// address)', (await a.evaluate(() => document.querySelector('.titlebar .docname input').value)) === 'msf1234.docx' || /\.docx$/.test(await a.evaluate(() => document.querySelector('.titlebar .docname input').value)), await a.evaluate(() => document.querySelector('.titlebar .docname input').value));
  // Save: first time asks for a name, writes to Documents/Isadek Office, offers Share
  await a.evaluate(() => { document.querySelector('.titlebar .docname input').value = 'Notes'; document.querySelector('.titlebar .docname input').dispatchEvent(new Event('change')); });
  await a.evaluate(() => Platform.emit('save'));
  await a.waitForSelector('.modal input');
  await a.click('.modal .actions .primary');
  await a.waitForSelector('.modal p:has-text("Documents")');
  check('app: saved into Documents/Isadek Office', (await a.evaluate(() => JSON.stringify(window.__calls.filter(c => c[0] === 'write')))).includes('Isadek Office/Notes.docx'));
  await a.click('.modal .actions button:has-text("Share")');
  await sleep(200);
  check('app: Share sends the saved file', (await a.evaluate(() => window.__calls.some(c => c[0] === 'share' && /Notes\.docx$/.test(c[1])))));
  check('app: AutoSave available for device files', !(await a.evaluate(() => document.querySelector('.autosave input').disabled)));
  await a.click('#doc'); await a.keyboard.type(' more');
  await a.evaluate(() => Platform.emit('save'));
  await sleep(300);
  const writes = await a.evaluate(() => window.__calls.filter(c => c[0] === 'write').length);
  check('app: the next Save writes the same file without asking', writes === 2 && !(await a.$('.modal')), writes);
  // printing and PDF through the app's plugin
  await a.evaluate(() => Platform.emit('print'));
  await sleep(200);
  check('app: Print opens the system print dialog', await a.evaluate(() => window.__calls.some(c => c[0] === 'print')));
  await a.evaluate(() => Platform.emit('export-pdf'));
  await sleep(300);
  check('app: Export PDF on Android uses the print dialog (Save as PDF)', await a.evaluate(() => window.__calls.some(c => c[0] === 'print' && /Notes\.pdf$/.test(c[1]))), await a.evaluate(() => JSON.stringify(window.__calls.filter(c => c[0] === 'print'))));
  // back button: closes a pane, then goes Home
  await a.evaluate(() => Platform.emit('theme'));
  await a.waitForSelector('.modal');
  await a.evaluate(() => window.__fire('backButton'));
  await sleep(150);
  check('app: back closes the open dialog first', !(await a.$('.modal')));
  await a.evaluate(() => window.__fire('backButton'));
  await a.waitForURL(/index\.html/, { timeout: 5000 });
  check('app: back again goes to Home', true);
  await sleep(500);
  await a.evaluate(() => window.__fire('backButton'));
  check('app: back on Home leaves the app', await a.evaluate(() => window.__calls.some(c => c[0] === 'exit')));
  await a.evaluate(() => sessionStorage.setItem('plat', 'ios'));
  await a.goto(B + 'write.html');
  await a.waitForFunction(() => window.Platform && Platform.native);
  await a.evaluate(() => { document.querySelector('.titlebar .docname input').value = 'Notes'; document.querySelector('.titlebar .docname input').dispatchEvent(new Event('change')); });
  await a.evaluate(() => Platform.emit('export-pdf'));
  await sleep(400);
  check('app: Export PDF on iOS saves the PDF and offers to share it', await a.evaluate(() => window.__calls.some(c => c[0] === 'pdf') && window.__calls.some(c => c[0] === 'write' && /Isadek Office\/Notes\.pdf$/.test(c[1])) && window.__calls.some(c => c[0] === 'share' && /Notes\.pdf$/.test(c[1]))), await a.evaluate(() => JSON.stringify(window.__calls.slice(-4))));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
