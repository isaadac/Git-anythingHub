'use strict';
// Light / dark theme for every page. Loaded in <head> so the right colours apply before the page
// paints. The choice (System, Light or Dark) is shared by all windows and remembered.
(function () {
  const KEY = 'atlas-theme';
  const media = window.matchMedia ? matchMedia('(prefers-color-scheme: dark)') : null;
  const read = () => { try { return localStorage.getItem(KEY) || 'system'; } catch { return 'system'; } };
  const resolve = mode => (mode === 'dark' || (mode === 'system' && media && media.matches) ? 'dark' : 'light');
  function apply(mode = read()) {
    document.documentElement.dataset.theme = resolve(mode);
    document.documentElement.dataset.themeMode = mode;
  }
  apply();
  if (media && media.addEventListener) media.addEventListener('change', () => apply());
  // Another window changed it.
  window.addEventListener('storage', e => { if (e.key === KEY) apply(); });

  const Theme = {
    get mode() { return read(); },
    get dark() { return document.documentElement.dataset.theme === 'dark'; },
    set(mode) {
      if (!['system', 'light', 'dark'].includes(mode)) return;
      try { localStorage.setItem(KEY, mode); } catch { /* ignore */ }
      apply(mode);
      // Native dialogs, menus and the window background follow too (desktop).
      if (window.atlas && window.atlas.setTheme) window.atlas.setTheme(mode);
    },
    async choose() {
      const modes = ['system', 'light', 'dark'];
      const labels = ['Use system setting', 'Light', 'Dark'];
      const i = await UI.choose('Office Theme', `Currently: ${labels[modes.indexOf(read())]}. Documents keep their own colours; the app around them changes.`, labels, '🎨');
      if (i >= 0) Theme.set(modes[i]);
    }
  };
  window.Theme = Theme;
})();
