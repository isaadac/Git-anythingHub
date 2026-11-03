'use strict';
// Light / dark theme for every page. Loaded in <head> so the right colours apply before the page
// paints. The choice (System, Light or Dark) is shared by all windows and remembered, and so is
// "Reduce motion" (also on whenever the system asks for less motion).
(function () {
  const KEY = 'atlas-theme', MOTION = 'atlas-motion';
  const calmMedia = window.matchMedia ? matchMedia('(prefers-reduced-motion: reduce)') : null;
  const readMotion = () => { try { return localStorage.getItem(MOTION) === 'reduce'; } catch { return false; } };
  const media = window.matchMedia ? matchMedia('(prefers-color-scheme: dark)') : null;
  const read = () => { try { return localStorage.getItem(KEY) || 'system'; } catch { return 'system'; } };
  const resolve = mode => (mode === 'dark' || (mode === 'system' && media && media.matches) ? 'dark' : 'light');
  function apply(mode = read()) {
    document.documentElement.dataset.theme = resolve(mode);
    document.documentElement.dataset.themeMode = mode;
    if (readMotion()) document.documentElement.dataset.motion = 'reduce'; else delete document.documentElement.dataset.motion;
  }
  apply();
  // Desktop: how the window is framed (the title bar leaves room for the system's buttons).
  if (window.atlas && window.atlas.frame) document.documentElement.dataset.frame = window.atlas.frame;
  if (media && media.addEventListener) media.addEventListener('change', () => apply());
  // Another window changed it.
  window.addEventListener('storage', e => { if (e.key === KEY || e.key === MOTION) apply(); });

  const Theme = {
    get mode() { return read(); },
    get dark() { return document.documentElement.dataset.theme === 'dark'; },
    /** True when the interface should not animate (the user's setting or the system's). */
    get calm() { return readMotion() || !!(calmMedia && calmMedia.matches); },
    setMotion(reduce) {
      try { if (reduce) localStorage.setItem(MOTION, 'reduce'); else localStorage.removeItem(MOTION); } catch { /* ignore */ }
      apply();
    },
    set(mode) {
      if (!['system', 'light', 'dark'].includes(mode)) return;
      try { localStorage.setItem(KEY, mode); } catch { /* ignore */ }
      apply(mode);
      // Native dialogs, menus and the window background follow too (desktop).
      if (window.atlas && window.atlas.setTheme) window.atlas.setTheme(mode);
    },
    /** File → Office Theme: light, dark or the system's, and whether the interface animates. */
    choose() {
      const modes = ['system', 'light', 'dark'];
      const labels = ['Use system setting', 'Light', 'Dark'];
      const back = document.createElement('div');
      back.className = 'modal-backdrop';
      back.innerHTML = '<div class="modal" role="dialog" aria-modal="true"><h2>Office Theme</h2><p></p>' +
        '<label class="check-row"><input type="checkbox"> <span>Reduce motion</span></label><div class="actions"></div></div>';
      back.querySelector('h2').id = 'theme-title';
      back.firstChild.setAttribute('aria-labelledby', 'theme-title');
      back.querySelector('p').textContent = `Currently: ${labels[modes.indexOf(read())]}. Documents keep their own colours; the app around them changes.`;
      const box = back.querySelector('input');
      box.checked = readMotion();
      box.addEventListener('change', () => Theme.setMotion(box.checked));
      const row = back.querySelector('.actions');
      labels.forEach((label, i) => {
        const b = document.createElement('button');
        b.textContent = label;
        if (modes[i] === read()) b.className = 'primary';
        b.onclick = () => { back.remove(); Theme.set(modes[i]); };
        row.appendChild(b);
      });
      back.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); back.remove(); } });
      document.body.appendChild(back);
      (row.querySelector('.primary') || row.firstChild).focus();
    }
  };
  window.Theme = Theme;
})();
