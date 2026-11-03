// Desktop: an organisation's policies (ENTERPRISE.md) turn features off, in the pages and in the
// main process; wrong values are ignored; the default language applies until someone picks one.
const { _electron: electron } = require('playwright');
const fs = require('fs');
const path = require('path');
const ROOT = process.cwd();
const check = (n, c, x = '') => { console.log((c ? 'PASS ' : 'FAIL ') + n + (x !== '' ? ' :: ' + x : '')); if (!c) process.exitCode = 1; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const W = process.env.E2E_WORK;
  const file = path.join(W, 'policies.json');
  fs.writeFileSync(file, JSON.stringify({
    DisableAssistant: true, DisableAddins: 1, DisableWebData: true, DisableCloudStorage: true, DisableCrashReports: true, DisableAutoUpdate: true,
    LiveServers: ['wss://live.example.org/', 'javascript:alert(1)', 'not a url'], DefaultLanguage: 'fr',
    DisableForms: 'yes please', Unknown: true, DisableLiveEditing: false
  }));
  const ud = path.join(W, 'ud-policy-' + Date.now());
  const app = await electron.launch({ executablePath: path.join(ROOT, 'node_modules/electron/dist/electron'), args: ['--no-sandbox', `--user-data-dir=${ud}`, '.'], env: { ...process.env, NODE_OPTIONS: '', ISADEK_POLICY_FILE: file } });
  const home = await app.firstWindow(); await sleep(1200);
  await home.evaluate(() => window.atlas.newWindow('write'));
  let w; for (let i = 0; i < 40 && !w; i++) { await sleep(250); w = app.windows().find(x => x.url().includes('write.html')); }
  await w.waitForFunction(() => window.AtlasWrite && window.I18n);
  const pol = await w.evaluate(() => window.atlas.policy);
  check('the policies reach the pages, wrong values and unknown names left out', JSON.stringify(pol) === JSON.stringify({ DisableAssistant: true, DisableCloudStorage: true, DisableLiveEditing: false, LiveServers: ['wss://live.example.org/'], DisableAddins: true, DisableWebData: true, DisableAutoUpdate: true, DisableCrashReports: true, DefaultLanguage: 'fr' }), JSON.stringify(pol));
  check('…and a page can’t change them', await w.evaluate(() => { try { window.atlas.policy.DisableAssistant = false; } catch { /* frozen */ } return window.atlas.policy.DisableAssistant === true && Platform.policy.DisableAssistant === true; }));
  check('the Assistant has no button', (await w.locator('.ai-btn').count()) === 0);
  check('…and the main process refuses it anyway', /turned off by your organisation/.test(JSON.stringify(await w.evaluate(() => window.atlas.aiAsk({ id: 'x', messages: [{ role: 'user', content: 'hi' }] })))));
  check('add-ins have no button', (await w.locator('#addinsBtn').count()) === 0);
  check('web downloads are refused in the main process', /turned off by your organisation/.test((await w.evaluate(() => window.atlas.fetchData('https://example.com/'))).error || ''));
  check('the organisation’s default language applies', (await w.evaluate(() => [I18n.lang, document.querySelector('.ribbon-tabs button[data-tab="insert"]').textContent].join())) === 'fr,Insertion');
  // Live editing: only the listed server; not this computer.
  check('hosting on this computer isn’t allowed (only listed servers)', /turned off/.test((await w.evaluate(() => window.atlas.collabHost())).error || ''));
  const toasts = [];
  await w.exposeFunction('__toast', t => toasts.push(t));
  await w.evaluate(() => { const t = UI.toast; UI.toast = (m, ms) => { window.__toast(String(m)); return t.call(UI, m, ms); }; });
  await w.evaluate(() => { UI.prompt = async () => 'wss://other.example.com/?room=' + 'a'.repeat(32) + '&app=write'; UI.choose = async (t, m, opts) => opts.indexOf('Join…'); });
  await w.click('.live-btn');
  await sleep(500);
  check('…joining a server that isn’t listed is refused', toasts.some(t => /doesn’t allow that live-editing server/.test(t)), toasts.join(' | '));
  // The native menu: crash reports and automatic updates can't be turned on.
  const menu = await app.evaluate(({ Menu }) => { const items = []; const walk = m => m.items.forEach(i => { items.push([i.label, i.enabled, i.checked]); if (i.submenu) walk(i.submenu); }); walk(Menu.getApplicationMenu()); return items; });
  const item = l => menu.find(i => i[0] === l) || [];
  check('crash reports and automatic updates are off and locked in the menu', item('Envoyer les rapports de plantage')[1] === false && item('Envoyer les rapports de plantage')[2] === false && item('Mettre à jour automatiquement')[1] === false, JSON.stringify([item('Envoyer les rapports de plantage'), item('Mettre à jour automatiquement')]));
  for (const x of app.windows()) await x.evaluate(() => { window.__atlasDirty = false; });
  await require('../close-app')(app);
})().then(() => process.exit(process.exitCode || 0)).catch(e => { console.error('CRASH', e.message); process.exit(1); });
