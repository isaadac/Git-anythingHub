'use strict';
// Accessibility for every page: landmarks and names for the chrome, a keyboard-friendly ribbon,
// F6 to move between areas, dialogs that keep and return focus, spoken announcements, and the
// Accessibility Checker pane.
(function () {
  const $ = s => document.querySelector(s);

  // ---------- Spoken announcements ----------
  const live = document.createElement('div');
  live.className = 'sr-only';
  live.setAttribute('aria-live', 'polite');
  live.setAttribute('role', 'status');
  let liveTimer = null;
  function announce(text) {
    clearTimeout(liveTimer);
    // Clear first so the same text is read again.
    live.textContent = '';
    liveTimer = setTimeout(() => { live.textContent = text; }, 60);
  }

  // ---------- Names for controls ----------
  function nameFor(el) {
    return el.getAttribute('aria-label') || el.title || el.placeholder || (el.id || '').replace(/([a-z])([A-Z])/g, '$1 $2').replace(/[-_]/g, ' ');
  }
  function labelControls(root = document) {
    root.querySelectorAll('input:not([type=hidden]), select, textarea').forEach(el => {
      if (el.getAttribute('aria-label') || el.getAttribute('aria-labelledby') || el.closest('label') || (el.id && document.querySelector(`label[for="${CSS.escape(el.id)}"]`))) return;
      const n = nameFor(el);
      if (n) el.setAttribute('aria-label', n);
    });
    root.querySelectorAll('button').forEach(b => {
      // Only symbol-only buttons get a name from their tooltip; visible words stay the name.
      if (b.getAttribute('aria-label') || b.textContent.trim().replace(/[^\p{L}\p{N}]/gu, '').length > 0) return;
      const n = b.title || nameFor(b);
      if (n) b.setAttribute('aria-label', n);
    });
  }

  // ---------- Landmarks and ribbon ----------
  function landmarks() {
    const app = (document.title.split('—').pop() || document.title).trim();
    const bar = $('.titlebar');
    if (bar) { bar.setAttribute('role', 'banner'); }
    if (!$('h1.sr-only') && ![...document.querySelectorAll('h1')].some(h => h.offsetParent !== null)) {
      const h = document.createElement('h1');
      h.className = 'sr-only';
      h.textContent = app;
      (bar || document.body).prepend(h);
    }
    const main = $('.workspace') || $('main') || $('.canvas');
    if (main && !main.closest('[role=main]') && main.tagName !== 'MAIN') { main.setAttribute('role', 'main'); main.setAttribute('aria-label', 'Document'); }
    const st = $('.statusbar');
    if (st) { st.setAttribute('role', 'contentinfo'); st.setAttribute('aria-label', 'Status bar'); }
    const ribbon = $('.ribbon');
    const tabsEl = $('.ribbon-tabs');
    if (ribbon && tabsEl && tabsEl.nextElementSibling === ribbon && !tabsEl.parentElement.classList.contains('ribbon-region')) {
      // Tabs and ribbon in one landmark; display: contents keeps the page layout as it was.
      const wrap = document.createElement('div');
      wrap.className = 'ribbon-region';
      wrap.setAttribute('role', 'region');
      wrap.setAttribute('aria-label', 'Ribbon');
      tabsEl.before(wrap);
      wrap.append(tabsEl, ribbon);
    } else if (ribbon) { ribbon.setAttribute('role', 'region'); ribbon.setAttribute('aria-label', 'Ribbon'); }
    const fx = $('.formula-bar');
    if (fx) { fx.setAttribute('role', 'region'); fx.setAttribute('aria-label', 'Formula bar'); }
    const st2 = $('#sheetTabs');
    if (st2 && !st2.closest('[role]')) { st2.setAttribute('role', 'navigation'); st2.setAttribute('aria-label', 'Sheets'); }
    const backstage = $('.backstage');
    if (backstage) { backstage.setAttribute('role', 'dialog'); backstage.setAttribute('aria-label', 'File'); }
    const tabs = document.querySelector('.ribbon-tabs');
    if (tabs) {
      tabs.setAttribute('role', 'tablist');
      tabs.setAttribute('aria-label', 'Ribbon tabs');
      const list = [...tabs.querySelectorAll('button')];
      const sync = () => list.forEach(b => {
        const on = b.classList.contains('active');
        b.setAttribute('aria-selected', on ? 'true' : 'false');
        b.tabIndex = on ? 0 : -1;
      });
      list.forEach(b => {
        b.setAttribute('role', 'tab');
        const p = b.dataset.tab && document.querySelector(`.ribbon-panel[data-panel="${b.dataset.tab}"]`);
        if (p) {
          if (!b.id) b.id = 'tab-' + b.dataset.tab;
          p.id = p.id || 'panel-' + b.dataset.tab;
          b.setAttribute('aria-controls', p.id);
          p.setAttribute('role', 'tabpanel');
          p.setAttribute('aria-labelledby', b.id);
        }
        b.addEventListener('click', () => setTimeout(sync));
      });
      new MutationObserver(sync).observe(tabs, { subtree: true, attributes: true, attributeFilter: ['class', 'hidden'] });
      sync();
      if (!list.some(b => b.classList.contains('active')) && list[0]) list[0].tabIndex = 0;
      // Arrow keys move between tabs (and select them), Down goes into the ribbon.
      tabs.addEventListener('keydown', e => {
        const vis = list.filter(b => !b.hidden && b.offsetParent !== null);
        const i = vis.indexOf(document.activeElement);
        if (i < 0) return;
        let n = null;
        if (e.key === 'ArrowRight') n = vis[(i + 1) % vis.length];
        if (e.key === 'ArrowLeft') n = vis[(i - 1 + vis.length) % vis.length];
        if (e.key === 'Home') n = vis[0];
        if (e.key === 'End') n = vis[vis.length - 1];
        if (n) { e.preventDefault(); n.focus(); if (!n.classList.contains('file-tab')) n.click(); return; }
        if (e.key === 'ArrowDown') {
          const p = document.querySelector('.ribbon-panel.active');
          const first = p && [...p.querySelectorAll('button, input, select')].find(x => !x.disabled && x.offsetParent !== null);
          if (first) { e.preventDefault(); first.focus(); }
        }
      });
    }
    // Ribbon buttons keep the editor's selection on mouse clicks, but must still be reachable by keyboard.
    document.querySelectorAll('.ribbon .btn, .titlebar button').forEach(b => { if (b.tabIndex < 0 && !b.closest('.ribbon-tabs')) b.tabIndex = 0; });
  }

  // ---------- F6: move between the main areas ----------
  function regions() {
    const docArea = $('#doc') || $('#gridWrap') || $('#stage') || $('main');
    return [
      $('.titlebar [data-cmd="save"]') || $('.titlebar button'),
      $('.ribbon-tabs button.active') || $('.ribbon-tabs button'),
      docArea,
      $('.ai-pane:not([hidden]) textarea, .ai-pane:not([hidden]) button'),
      $('.statusbar button, .statusbar input')
    ].filter(x => x && x.offsetParent !== null);
  }
  document.addEventListener('keydown', e => {
    if (e.key !== 'F6' || document.querySelector('.modal-backdrop')) return;
    const rs = regions();
    if (!rs.length) return;
    e.preventDefault();
    const at = rs.findIndex(r => r === document.activeElement || r.contains(document.activeElement) || (r.closest('.ribbon-tabs') && document.activeElement.closest('.ribbon-tabs')) || (r.closest('.ai-pane') && document.activeElement.closest('.ai-pane')));
    const n = rs[(at + (e.shiftKey ? -1 : 1) + rs.length) % rs.length];
    if (!n.hasAttribute('tabindex') && !/^(BUTTON|INPUT|SELECT|TEXTAREA|A)$/.test(n.tagName) && !n.isContentEditable) n.tabIndex = -1;
    n.focus();
  });

  // ---------- Dialogs ----------
  // Every modal: a labelled dialog that keeps Tab inside it and gives focus back when it closes.
  const opened = new WeakMap();
  // Where focus was before a dialog opened (dialogs focus their first field straight away).
  let lastOutside = null;
  document.addEventListener('focusin', e => { if (!e.target.closest('.modal-backdrop')) lastOutside = e.target; });
  const focusables = m => [...m.querySelectorAll('button, input, select, textarea, a[href], [tabindex]:not([tabindex="-1"])')].filter(x => !x.disabled && x.offsetParent !== null);
  new MutationObserver(muts => {
    for (const mu of muts) {
      mu.addedNodes.forEach(n => {
        if (n.nodeType !== 1 || !n.classList.contains('modal-backdrop')) return;
        const m = n.querySelector('.modal');
        if (!m) return;
        m.setAttribute('role', m.getAttribute('role') || 'dialog');
        m.setAttribute('aria-modal', 'true');
        const h = m.querySelector('h2');
        if (h) { h.id = h.id || 'dlg-' + Math.random().toString(36).slice(2, 8); m.setAttribute('aria-labelledby', h.id); }
        const p = m.querySelector('p');
        if (p) { p.id = p.id || 'dlgd-' + Math.random().toString(36).slice(2, 8); m.setAttribute('aria-describedby', p.id); }
        labelControls(m);
        opened.set(n, lastOutside);
        setTimeout(() => { if (!m.contains(document.activeElement)) { const f = focusables(m)[0]; if (f) f.focus(); } });
      });
      mu.removedNodes.forEach(n => {
        if (n.nodeType !== 1 || !n.classList || !n.classList.contains('modal-backdrop')) return;
        const back = opened.get(n);
        if (back && back.isConnected && !document.querySelector('.modal-backdrop') && (document.activeElement === document.body || !document.activeElement)) back.focus({ preventScroll: true });
      });
    }
  }).observe(document.body, { childList: true });
  document.addEventListener('keydown', e => {
    if (e.key !== 'Tab') return;
    const backs = document.querySelectorAll('.modal-backdrop');
    const m = backs.length && backs[backs.length - 1].querySelector('.modal');
    if (!m) return;
    const f = focusables(m);
    if (!f.length) return;
    const i = f.indexOf(document.activeElement);
    if (e.shiftKey && (i <= 0)) { e.preventDefault(); f[f.length - 1].focus(); }
    else if (!e.shiftKey && (i === f.length - 1 || i < 0)) { e.preventDefault(); f[0].focus(); }
  }, true);

  // ---------- Accessibility Checker ----------
  /**
   * checker({ check() -> [{ kind: 'error'|'warning'|'tip', title, detail, fix?, go?() }] })
   * Office-style categories: errors make content hard or impossible to use, warnings make it
   * harder, tips are good practice.
   */
  function checker(o) {
    const pane = document.createElement('aside');
    pane.className = 'ai-pane a11y-pane';
    pane.hidden = true;
    pane.setAttribute('aria-label', 'Accessibility');
    pane.innerHTML = '<header><b>Accessibility</b><span class="sp"></span><button class="re" title="Check again">↻</button><button class="x" title="Close">✕</button></header><div class="a11y-body"></div>';
    document.body.appendChild(pane);
    const body = pane.querySelector('.a11y-body');
    pane.querySelector('.x').onclick = () => { pane.hidden = true; document.body.classList.remove('ai-open'); };
    pane.querySelector('.re').onclick = () => run();
    function run() {
      const issues = o.check();
      body.innerHTML = '';
      if (!issues.length) {
        body.innerHTML = '<div class="proof-done"><div class="big">✓</div><b>No accessibility issues found</b><p>People with disabilities should not have trouble reading this.</p></div>';
        announce('No accessibility issues found');
        return;
      }
      const groups = [['error', 'Errors'], ['warning', 'Warnings'], ['tip', 'Tips']];
      for (const [k, label] of groups) {
        const list = issues.filter(x => x.kind === k);
        if (!list.length) continue;
        const h = document.createElement('h3');
        h.className = 'a11y-h ' + k;
        h.textContent = `${label} (${list.length})`;
        body.appendChild(h);
        list.forEach(x => {
          const b = document.createElement('div');
          b.className = 'a11y-item';
          const t = document.createElement('button');
          t.className = 'a11y-title';
          t.textContent = x.title;
          t.onclick = () => { if (x.go) x.go(); };
          const d = document.createElement('div');
          d.className = 'a11y-detail';
          d.textContent = x.detail || '';
          b.append(t, d);
          if (x.fix) {
            const f = document.createElement('button');
            f.className = 'a11y-fix';
            f.textContent = x.fixLabel || 'Fix';
            f.onclick = async () => { await x.fix(); run(); };
            b.appendChild(f);
          }
          body.appendChild(b);
        });
      }
      announce(`${issues.length} accessibility ${issues.length === 1 ? 'issue' : 'issues'} found`);
    }
    return {
      open() {
        document.querySelectorAll('.ai-pane').forEach(p => { if (p !== pane) p.hidden = true; });
        pane.hidden = false;
        document.body.classList.add('ai-open');
        run();
        pane.querySelector('.re').focus();
      },
      run
    };
  }

  // ---------- Colour contrast (WCAG) ----------
  function rgb(c) {
    if (!c) return null;
    const m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(c.trim());
    if (m) { let h = m[1]; if (h.length === 3) h = [...h].map(x => x + x).join(''); return [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16)); }
    const r = /rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/i.exec(c);
    return r ? [+r[1], +r[2], +r[3]] : null;
  }
  function luminance([r, g, b]) {
    const f = v => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b);
  }
  /** Contrast ratio of two CSS colours (1–21), or null if either can't be read. */
  function contrast(a, b) {
    const x = rgb(a), y = rgb(b);
    if (!x || !y) return null;
    const l1 = luminance(x), l2 = luminance(y);
    return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
  }

  function init() {
    landmarks();
    labelControls();
    const t = document.querySelector('.toast');
    if (t) t.setAttribute('role', 'status');
    // Labels for controls added later (dialogs, panes, the AutoSave switch…).
    let pending = false;
    new MutationObserver(() => { if (pending) return; pending = true; requestAnimationFrame(() => { pending = false; labelControls(); }); }).observe(document.body, { childList: true, subtree: true });
    document.body.appendChild(live);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else setTimeout(init);

  window.A11y = { announce, checker, contrast, labelControls };
})();
