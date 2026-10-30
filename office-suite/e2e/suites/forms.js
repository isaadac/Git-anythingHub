// Isadek Forms: build a form in Sheets, publish it to the web server the app runs on, answer it
// on the public page (as someone else would, on a phone), and get the answers as rows.
const { chromium } = require('playwright');
const path = require('path');
const { createWebServer } = require(path.join(process.cwd(), 'src/server/web-server.js'));
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
(async () => {
  const srv = createWebServer({ live: false, apiKey: '' });
  await new Promise(r => srv.listen(0, '127.0.0.1', r));
  const base = `http://127.0.0.1:${srv.address().port}`;
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  p.on('pageerror', e => errors.push(e.message));
  p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  await p.goto(base + '/src/renderer/sheets.html');
  await p.waitForFunction(() => window.AtlasSheets);

  // Build: Insert → Form starts a form with one question.
  await p.click('[data-tab="insert"]');
  await p.click('#formBtn');
  await p.waitForSelector('.form-pane .fm-q');
  await p.fill('.form-pane [data-k="title"]', 'Team lunch');
  await p.fill('.form-pane [data-k="description"]', 'Tell us what you’d like.');
  await p.click('.form-pane [data-a="add"]');
  let qs = p.locator('.form-pane .fm-q');
  await qs.nth(1).locator('[data-k="q.title"]').fill('Day');
  await qs.nth(1).locator('[data-k="q.options"]').fill('Mon\nTue\nWed');
  await qs.nth(1).locator('[data-k="q.required"]').check();
  await p.click('.form-pane [data-a="add"]');
  qs = p.locator('.form-pane .fm-q');
  await qs.nth(2).locator('[data-k="q.title"]').fill('Guests');
  await qs.nth(2).locator('[data-k="q.type"]').selectOption('number');
  await p.click('.form-pane [data-a="add"]');
  qs = p.locator('.form-pane .fm-q');
  await qs.nth(3).locator('[data-k="q.title"]').fill('Food');
  await qs.nth(3).locator('[data-k="q.type"]').selectOption('checkboxes');
  await qs.nth(3).locator('[data-k="q.options"]').fill('Pizza\nSalad\nSoup');
  const def = await p.evaluate(() => AtlasSheets.book.forms[0].def);
  check('the builder makes the form', def.title === 'Team lunch' && def.questions.map(q => q.type).join() === 'text,choice,number,checkboxes' && def.questions[1].required && def.questions[3].options.join() === 'Pizza,Salad,Soup', JSON.stringify(def));

  // Preview shows the form as people will see it, and checks required answers.
  await p.click('.form-pane [data-a="preview"]');
  await p.waitForSelector('.form-modal .fv-q');
  await p.click('.form-modal .fv-submit');
  const err = await p.textContent('.form-modal .fv-q[data-q="q1"] .fv-err');
  check('preview checks required questions', /needs an answer/.test(err), err);
  await p.click('.form-modal .x');

  // Publish: the server the app is on collects the answers.
  await p.click('.form-pane [data-a="publish"]');
  await p.waitForFunction(() => AtlasSheets.book.forms[0].id);
  const link = await p.inputValue('.form-pane .fm-link');
  check('publishing gives a link', link.startsWith(base + '/f/'), link);
  const sheetNames = await p.evaluate(() => AtlasSheets.book.sheets.map(s => s.name));
  check('a responses sheet with a header row', sheetNames.includes('Form Responses 1') && (await p.evaluate(() => { AtlasSheets.setSheet(1); return ['A1', 'B1', 'C1', 'D1', 'E1'].map(AtlasSheets.valueOf).join('|'); })) === 'Timestamp|Name|Day|Guests|Food', JSON.stringify(sheetNames));

  // Someone answers on the public page, on a phone.
  const phone = await b.newPage({ viewport: { width: 390, height: 800 } });
  phone.on('pageerror', e => errors.push('form page: ' + e.message));
  await phone.goto(link);
  await phone.waitForSelector('.fv-q');
  check('the link opens the form', (await phone.textContent('.fv-form-title')) === 'Team lunch' && (await phone.title()).startsWith('Team lunch'));
  const wide = await phone.evaluate(() => document.documentElement.scrollWidth <= innerWidth);
  check('the form fits a phone screen', wide);
  await phone.fill('.fv-q[data-q="q1"] input', '=HYPERLINK("http://evil.example","click")');
  await phone.click('.fv-submit');
  check('the page asks for required answers', /needs an answer/.test(await phone.textContent('.fv-q[data-q="q2"] .fv-err')));
  await phone.check('.fv-q[data-q="q2"] input[value="Tue"]');
  await phone.fill('.fv-q[data-q="q3"] input', '4');
  await phone.check('.fv-q[data-q="q4"] input[value="Soup"]');
  await phone.check('.fv-q[data-q="q4"] input[value="Pizza"]');
  await phone.click('.fv-submit');
  await phone.waitForSelector('.fv-done');
  check('the answer is sent', /recorded/.test(await phone.textContent('.fv-done')));
  await phone.click('.fv-done .fv-clear');
  await phone.fill('.fv-q[data-q="q1"] input', 'Bo');
  await phone.check('.fv-q[data-q="q2"] input[value="Mon"]');
  await phone.click('.fv-submit');
  await phone.waitForSelector('.fv-done');

  // The owner gets them as rows.
  await p.click('.form-pane [data-a="refresh"]');
  await p.waitForFunction(() => AtlasSheets.valueOf('B3') === 'Bo');
  const rows = await p.evaluate(() => ({
    a2: AtlasSheets.valueOf('A2'), b2: AtlasSheets.valueOf('B2'), raw: AtlasSheets.rawOf('B2'), c2: AtlasSheets.valueOf('C2'), d2: AtlasSheets.valueOf('D2'), e2: AtlasSheets.valueOf('E2'),
    shown: document.querySelector('#grid tr[data-r="1"] td[data-c="0"]') && document.querySelector('#grid tr[data-r="1"] td[data-c="0"]').textContent
  }));
  check('responses arrive as rows', typeof rows.a2 === 'number' && /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(rows.shown) && rows.c2 === 'Tue' && rows.d2 === 4 && rows.e2 === 'Pizza, Soup', JSON.stringify(rows));
  check('an answer that looks like a formula stays text', rows.b2 === '=HYPERLINK("http://evil.example","click")' && rows.raw.startsWith("'"), JSON.stringify(rows));
  const sum = await p.textContent('.form-pane .fm-summary');
  check('the summary counts the answers', /2\s*responses/.test(sum) && /Tue/.test(sum), sum.replace(/\s+/g, ' ').slice(0, 200));
  await p.click('.form-pane [data-a="refresh"]');
  await p.waitForTimeout(300);
  check('getting responses again adds nothing twice', (await p.evaluate(() => AtlasSheets.valueOf('B4'))) === '' || (await p.evaluate(() => AtlasSheets.valueOf('B4'))) === undefined);

  // Fill In: a response typed in here goes to the server too.
  await p.click('.form-pane [data-a="fill"]');
  await p.fill('.form-modal .fv-q[data-q="q1"] input', 'Cy');
  await p.check('.form-modal .fv-q[data-q="q2"] input[value="Wed"]');
  await p.click('.form-modal .fv-submit');
  await p.waitForFunction(() => AtlasSheets.valueOf('B4') === 'Cy');
  check('Fill In adds a response', true);
  await p.click('.form-modal .x');

  // Changing the form: a new question gets a new column; published forms update.
  await p.click('.form-pane [data-a="add"]');
  qs = p.locator('.form-pane .fm-q');
  await qs.nth(4).locator('[data-k="q.title"]').fill('Rating');
  await qs.nth(4).locator('[data-k="q.type"]').selectOption('rating');
  await p.click('.form-pane [data-a="publish"]');
  await p.waitForFunction(() => AtlasSheets.valueOf('F1') === 'Rating');
  await phone.goto(link);
  await phone.waitForSelector('.fv-q[data-q="q5"] .fv-rating');
  check('the update reaches the answer page', (await phone.locator('.fv-q[data-q="q5"] .fv-star').count()) === 5);

  // Stop accepting responses.
  await p.click('.form-pane [data-a="toggle"]');
  await p.waitForFunction(() => AtlasSheets.book.forms[0].open === false);
  await phone.goto(link);
  await phone.waitForSelector('.fv-done');
  check('a closed form says so', /no longer accepting/.test(await phone.textContent('.fv-done')));
  await p.click('.form-pane [data-a="toggle"]');

  // Saved in the workbook (with its key), and back.
  const bytes = await p.evaluate(async () => Array.from(await AtlasSheets.serialize('xlsx')));
  const p2 = await b.newPage();
  p2.on('pageerror', e => errors.push('reopened: ' + e.message));
  await p2.goto(base + '/src/renderer/sheets.html');
  await p2.waitForFunction(() => window.AtlasSheets);
  await p2.evaluate(() => { localStorage.removeItem('atlas-forms-mine'); });
  await p2.evaluate(async b => AtlasSheets.loadFile({ name: 'lunch.xlsx', path: null, data: new Uint8Array(b) }), bytes);
  const back = await p2.evaluate(() => { const f = AtlasSheets.book.forms[0]; return { title: f.def.title, id: f.id, got: f.got, n: f.def.questions.length }; });
  check('the form is saved in the workbook', back.title === 'Team lunch' && back.got === 3 && back.n === 5 && !!back.id, JSON.stringify(back));
  // Someone else's workbook asks before contacting its server.
  await phone.goto(link);
  await phone.waitForSelector('.fv-q');
  await phone.fill('.fv-q[data-q="q1"] input', 'Di');
  await phone.check('.fv-q[data-q="q2"] input[value="Mon"]');
  await phone.click('.fv-submit');
  await phone.waitForSelector('.fv-done');
  await p2.evaluate(() => { window.__asked = 0; UI.confirm = async () => { window.__asked++; return true; }; });
  await p2.evaluate(() => AtlasSheets.forms.refreshAll(false));
  const after = await p2.evaluate(() => ({ asked: window.__asked, b5: (AtlasSheets.setSheet(1), AtlasSheets.valueOf('B5')) }));
  check('a workbook from elsewhere asks first, then gets new answers', after.asked === 1 && after.b5 === 'Di', JSON.stringify(after));

  // A form that isn't published collects answers typed in here.
  await p2.evaluate(() => AtlasSheets.forms.newForm());
  await p2.waitForSelector('.form-pane .fm-q');
  await p2.click('.form-pane [data-a="fill"]');
  await p2.fill('.form-modal .fv-q[data-q="q1"] input', 'Local');
  await p2.click('.form-modal .fv-submit');
  await p2.waitForSelector('.form-modal .fv-done');
  const local = await p2.evaluate(() => { const f = AtlasSheets.book.forms[1]; const i = AtlasSheets.book.sheets.findIndex(s => s.name === f.sheet); AtlasSheets.setSheet(i); return AtlasSheets.valueOf('B2'); });
  check('an unpublished form fills in locally', local === 'Local', local);

  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
  srv.close();
})().catch(e => { console.log('FAIL crashed :: ' + e.message); process.exit(1); });
