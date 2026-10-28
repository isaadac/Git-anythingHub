'use strict';
// Phones and tablets: touch dragging everywhere, and — inside the Android/iOS app (Capacitor) —
// saving to the device's Documents folder, the share sheet, files opened from other apps
// ("Open with"), and the Android back button.
(function () {
  const C = window.Capacitor;
  const native = !!(C && C.isNativePlatform && C.isNativePlatform());
  const touch = 'ontouchstart' in window || navigator.maxTouchPoints > 0;
  document.documentElement.classList.toggle('native-app', native);
  document.documentElement.classList.toggle('touch', touch);
  Platform.native = native;

  // ---------- Touch: drag objects, handles and charts like a mouse ----------
  // Taps already arrive as clicks; dragging needs mouse events, which touch doesn't send.
  const DRAG = '#stage, .handle, .fill-handle, .chart .bar, .chart .rs, .resizer, .img-tools, .dv-arrow';
  let lastTap = 0, lastTarget = null;
  const mouse = (type, t, target) => target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window, clientX: t.clientX, clientY: t.clientY, screenX: t.screenX, screenY: t.screenY, button: 0, buttons: type === 'mouseup' ? 0 : 1 }));
  document.addEventListener('touchstart', e => {
    if (e.touches.length !== 1) return;
    const target = e.target.closest && e.target.closest(DRAG) && e.target;
    if (!target || target.isContentEditable) return;
    e.preventDefault();
    const t = e.touches[0];
    let moved = false;
    mouse('mousedown', t, target);
    const move = ev => { if (ev.touches.length !== 1) return; moved = true; ev.preventDefault(); mouse('mousemove', ev.touches[0], document.elementFromPoint(ev.touches[0].clientX, ev.touches[0].clientY) || target); };
    // Touch events stay with the element first touched, even if a redraw removes it from the page.
    const src = e.target;
    const end = ev => {
      src.removeEventListener('touchmove', move);
      src.removeEventListener('touchend', end);
      src.removeEventListener('touchcancel', end);
      const p = ev.changedTouches[0];
      mouse('mouseup', p, document.elementFromPoint(p.clientX, p.clientY) || target);
      if (!moved) {
        mouse('click', p, target);
        // Double tap: edit the text, like a double click.
        if (Date.now() - lastTap < 350 && lastTarget === target) mouse('dblclick', p, target);
        lastTap = Date.now(); lastTarget = target;
      }
    };
    src.addEventListener('touchmove', move, { passive: false });
    src.addEventListener('touchend', end);
    src.addEventListener('touchcancel', end);
  }, { passive: false });

  if (!native) return;

  // ---------- Native app ----------
  const { Filesystem: FS, Share, App } = C.Plugins;
  const FOLDER = 'Isadek Office';
  const toB64 = bytes => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s); };
  const fromB64 = b64 => { const bin = atob(b64); const out = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i); return out; };
  const isDevicePath = p => typeof p === 'string' && p.startsWith('device:');
  Platform.isDevicePath = isDevicePath;

  async function writeDoc(name, data) {
    const bytes = data instanceof Uint8Array ? data : data instanceof ArrayBuffer ? new Uint8Array(data) : data instanceof Blob ? new Uint8Array(await data.arrayBuffer()) : new TextEncoder().encode(String(data));
    await FS.writeFile({ path: `${FOLDER}/${name}`, data: toB64(bytes), directory: 'DOCUMENTS', recursive: true });
  }
  async function shareDoc(name) {
    const { uri } = await FS.getUri({ path: `${FOLDER}/${name}`, directory: 'DOCUMENTS' });
    await Share.share({ title: name, files: [uri], dialogTitle: 'Share ' + name });
  }

  const announced = new Set();
  const webSave = Platform.save.bind(Platform);
  Platform.save = async opts => {
    const { data, name, path } = opts;
    if (Platform.isCloudPath(path)) return webSave(opts);
    const fileName = isDevicePath(path) ? path.slice(7) : name;
    try {
      await writeDoc(fileName, data);
    } catch (err) {
      throw new Error('Could not save to the device: ' + (err.message || err));
    }
    if (!announced.has(fileName)) {
      announced.add(fileName);
      // First save: say where it went, and offer to send it.
      UI.choose('Saved', `Saved to Documents › ${FOLDER} › ${fileName}`, ['OK', 'Share…']).then(i => { if (i === 1) shareDoc(fileName).catch(() => {}); });
    }
    return { name: fileName, path: 'device:' + fileName };
  };
  const webChoose = Platform.chooseSavePath.bind(Platform);
  Platform.chooseSavePath = async (name, filters) => {
    const t = await webChoose(name, filters);
    if (!t) return null;
    const nm = t.name.replace(/[\\/:*?"<>|]/g, '_');
    return { name: nm, path: 'device:' + nm };
  };
  Platform.canWriteBack = p => isDevicePath(p) || Platform.isCloudPath(p);
  Platform.exportPdf = async () => { UI.toast('PDF export isn’t available in the mobile app yet. Save the file and share it instead.', 4000); return null; };
  Platform.print = async () => UI.toast('Printing isn’t available in the mobile app. Share the file to a printing app instead.', 4000);
  Platform.shareCurrent = async (state, save) => {
    if (!isDevicePath(state.path)) { if (!(await save())) return; }
    else if (state.dirty) await save();
    if (isDevicePath(state.path)) await shareDoc(state.path.slice(7));
  };

  // Files opened from other apps ("Open with Isadek Office").
  const KIND = [['word/', 'docx'], ['xl/', 'xlsx'], ['ppt/', 'pptx']];
  const APP = { docx: 'write', html: 'write', htm: 'write', md: 'write', txt: 'write', xlsx: 'sheets', xls: 'sheets', csv: 'sheets', tsv: 'sheets', ods: 'sheets', pptx: 'slides', aslides: 'slides', anote: 'notes' };
  function guessName(url, bytes) {
    let name = decodeURIComponent(url.split(/[?#]/)[0].split('/').pop() || 'Document');
    name = name.split(':').pop().split('/').pop();
    if (!APP[Platform.extOf(name)]) {
      // No usable extension in a content:// address: look inside the zip.
      const head = new TextDecoder('latin1').decode(bytes.subarray(0, Math.min(bytes.length, 400000)));
      const k = bytes[0] === 0x50 && bytes[1] === 0x4b ? KIND.find(([dir]) => head.includes(dir)) : null;
      name = Platform.baseName(name) + '.' + (k ? k[1] : /^\s*</.test(head) ? 'html' : 'txt');
    }
    return name;
  }
  async function openUrl(url) {
    if (!url || /^https?:/.test(url)) return;
    try {
      const { data } = await FS.readFile({ path: url });
      const bytes = typeof data === 'string' ? fromB64(data) : new Uint8Array(await data.arrayBuffer());
      const name = guessName(url, bytes);
      const app = APP[Platform.extOf(name)];
      if (!app) return UI.toast('Isadek Office can’t open this kind of file', 4000);
      await Platform.openData(app, name, bytes, true);
    } catch (err) {
      UI.toast('Could not open the file: ' + (err.message || err), 5000);
    }
  }
  App.addListener('appUrlOpen', e => openUrl(e.url));
  if (/index\.html$/.test(location.pathname)) App.getLaunchUrl().then(r => r && r.url && openUrl(r.url)).catch(() => {});

  // Android back button: close what is open, then go back to Home, then leave.
  App.addListener('backButton', () => {
    const modal = document.querySelector('.modal-backdrop');
    if (modal) { modal.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); return; }
    const pane = document.querySelector('.ai-pane:not([hidden]) .x');
    if (pane) { pane.click(); return; }
    const back = document.querySelector('.backstage.open .back');
    if (back) { back.click(); return; }
    if (/index\.html$/.test(location.pathname)) { App.exitApp(); return; }
    if (window.__atlasDirty && !confirm('Leave without saving your changes?')) return;
    window.__atlasDirty = false;
    location.href = 'index.html';
  });
})();
