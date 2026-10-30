// Closes an Electron app under test without waiting on the unsaved-changes prompt.
module.exports = async app => {
  const close = (async () => {
    for (const w of app.windows()) await w.evaluate(() => { window.__atlasDirty = false; }).catch(() => {});
    await app.close();
  })().catch(() => {});
  await Promise.race([close, new Promise(r => setTimeout(r, 8000))]);
  try { app.process().kill(); } catch {}
};
