const { chromium } = require('playwright');
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext();
  const errors = [];
  // Without a desktop dictionary: grammar only
  let p = await ctx.newPage();
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.click('#doc'); await p.keyboard.type('Teh dog is is here.');
  await p.keyboard.press('F7');
  await p.waitForSelector('.proof-card, .proof-done');
  check('web: grammar issues without a dictionary', /0 spelling · 2 grammar/.test(await p.textContent('.proof-sum')), await p.textContent('.proof-sum'));
  // With a stub dictionary standing in for the desktop spellchecker
  await ctx.addInitScript(() => {
    const bad = { hapy: ['happy', 'hap'], wrold: ['world'] };
    window.atlas = { spellCheck: ws => ws.map(w => w.toLowerCase() in bad), spellSuggest: w => bad[w.toLowerCase()] || [], spellAddWord: async () => true };
  });
  p = await ctx.newPage();
  p.on('pageerror', e => errors.push(e.message));
  await p.goto(process.env.E2E_BASE + '/src/renderer/write.html');
  await p.waitForFunction(() => window.Proofing);
  await p.click('#doc'); await p.keyboard.type('A hapy wrold, a hapy day, and a wrold of e-mail at me@hapy.org.');
  await p.keyboard.press('F7');
  await p.waitForSelector('.proof-card');
  check('spelling issues found (addresses skipped)', /4 spelling/.test(await p.textContent('.proof-sum')), await p.textContent('.proof-sum'));
  check('suggestions shown', (await p.$$eval('.proof-card .sugg button', bs => bs.map(x => x.textContent))).join() === 'happy,hap');
  await p.click('.proof-card .acts button:has-text("Change all")');
  await sleep(100);
  check('Change all fixes every copy', (await p.$eval('#doc', d => d.innerText)).match(/happy/g).length === 2 && !(await p.$eval('#doc', d => d.innerText)).includes(' hapy '));
  check('moves on to the next word', (await p.textContent('.proof-card mark')) === 'wrold');
  await p.click('.proof-card .acts button:has-text("Ignore all")');
  await sleep(100);
  check('Ignore all skips the word everywhere', !!(await p.$('.proof-done')), await p.textContent('.ai-pane.proof-pane'));
  check('address untouched', (await p.$eval('#doc', d => d.innerText)).includes('me@hapy.org'));
  check('no page errors', errors.length === 0, errors.join(' | '));
  await b.close();
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.log('FAIL exception', e); process.exit(1); });
