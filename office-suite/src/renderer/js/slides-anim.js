'use strict';
// Isadek Slides — animations and transitions.
// An element's animation is el.anim = { effect, dir, start: 'click'|'with'|'after', dur, delay, order }.
// A slide's transition is s.transition (+ s.transDur seconds, s.advance seconds or null).
// Both play in the slide show and are saved as PowerPoint <p:timing> / <p:transition> XML.
(function (root) {
  const DIRS = ['bottom', 'left', 'right', 'top'];
  /** Effects: kind entr (appears), exit (disappears) or emph; PowerPoint preset ids; keyframes for the show. */
  const EFFECTS = {
    appear: { label: 'Appear', kind: 'entr', preset: 1, frames: () => [{ opacity: 1 }, { opacity: 1 }] },
    fade: { label: 'Fade', kind: 'entr', preset: 10, frames: () => [{ opacity: 0 }, { opacity: 1 }] },
    fly: { label: 'Fly In', kind: 'entr', preset: 2, dirs: true, frames: d => [{ translate: off(d, 700) }, { translate: '0 0' }] },
    float: { label: 'Float In', kind: 'entr', preset: 42, frames: () => [{ opacity: 0, translate: '0 40px' }, { opacity: 1, translate: '0 0' }] },
    wipe: { label: 'Wipe', kind: 'entr', preset: 22, dirs: true, frames: d => [{ clipPath: wipeFrom(d) }, { clipPath: 'inset(0 0 0 0)' }] },
    zoom: { label: 'Zoom', kind: 'entr', preset: 53, frames: () => [{ opacity: 0, scale: '0.3' }, { opacity: 1, scale: '1' }] },
    pulse: { label: 'Pulse', kind: 'emph', preset: 6, frames: () => [{ scale: '1' }, { scale: '1.15' }, { scale: '1' }] },
    spin: { label: 'Spin', kind: 'emph', preset: 8, frames: () => [{ rotate: '0deg' }, { rotate: '360deg' }] },
    fadeOut: { label: 'Fade Out', kind: 'exit', preset: 10, frames: () => [{ opacity: 1 }, { opacity: 0 }] },
    flyOut: { label: 'Fly Out', kind: 'exit', preset: 2, dirs: true, frames: d => [{ translate: '0 0' }, { translate: off(d, 700) }] },
    zoomOut: { label: 'Zoom Out', kind: 'exit', preset: 53, frames: () => [{ opacity: 1, scale: '1' }, { opacity: 0, scale: '0.3' }] }
  };
  function off(d, n) { return { bottom: `0 ${n}px`, top: `0 -${n}px`, left: `-${n * 1.8}px 0`, right: `${n * 1.8}px 0` }[d] || `0 ${n}px`; }
  function wipeFrom(d) { return { bottom: 'inset(100% 0 0 0)', top: 'inset(0 0 100% 0)', left: 'inset(0 100% 0 0)', right: 'inset(0 0 0 100%)' }[d] || 'inset(100% 0 0 0)'; }
  const TRANSITIONS = ['none', 'fade', 'push', 'wipe', 'split', 'cover', 'uncover', 'zoom'];

  function clean(a) {
    if (!a || typeof a !== 'object' || !Object.hasOwn(EFFECTS, a.effect)) return null;
    const n = (v, d, lo, hi) => (typeof v === 'number' && isFinite(v) ? Math.max(lo, Math.min(hi, v)) : d);
    return {
      effect: a.effect, dir: DIRS.includes(a.dir) ? a.dir : 'bottom', start: ['click', 'with', 'after'].includes(a.start) ? a.start : 'click',
      dur: n(a.dur, a.effect === 'appear' ? 0 : 0.5, 0, 60), delay: n(a.delay, 0, 0, 600), order: n(a.order, 0, 0, 10000)
    };
  }
  /**
   * The slide's builds: groups started by a click (the first group may start on its own when
   * the first effect is "with"/"after previous"). Each item: { el, anim, at } (ms from group start).
   */
  function sequence(slide) {
    const items = slide.elements.filter(e => e.anim).sort((a, b) => a.anim.order - b.anim.order);
    const groups = [];
    let g = null, prevEnd = 0, prevStart = 0;
    items.forEach((el, i) => {
      const a = el.anim;
      if (!g || a.start === 'click') { g = { auto: i === 0 && a.start !== 'click', items: [] }; groups.push(g); prevEnd = 0; prevStart = 0; }
      const at = (a.start === 'after' ? prevEnd : a.start === 'with' ? prevStart : 0) + a.delay * 1000;
      g.items.push({ el, anim: a, at });
      prevStart = at;
      prevEnd = Math.max(prevEnd, at + a.dur * 1000);
    });
    return groups;
  }
  /** Before any build: entrance elements hidden. After group k: its entrances shown, exits hidden. */
  function visibilityAfter(slide, groups, k) {
    const vis = new Map();
    slide.elements.forEach(e => { if (e.anim) { const k = EFFECTS[e.anim.effect].kind; if (k !== 'emph') vis.set(e.id, k === 'exit'); } });
    groups.slice(0, k).forEach(g => g.items.forEach(({ el, anim }) => {
      const kind = EFFECTS[anim.effect].kind;
      if (kind === 'entr') vis.set(el.id, true);
      if (kind === 'exit') vis.set(el.id, false);
    }));
    return vis;
  }
  /** Plays a group on the slide's DOM; returns the running animations. */
  function play(group, nodeOf) {
    const running = [];
    group.items.forEach(({ el, anim, at }) => {
      const node = nodeOf(el.id);
      if (!node) return;
      const fx = EFFECTS[anim.effect];
      if (fx.kind === 'entr') node.style.visibility = 'hidden';
      const a = node.animate(fx.frames(anim.dir), { duration: Math.max(1, anim.dur * 1000), delay: at, easing: 'ease-out', fill: 'backwards' });
      // Visibility flips when the effect starts (entrances) or ends (exits).
      if (fx.kind === 'entr') { const t = setTimeout(() => { node.style.visibility = ''; }, at); a.addEventListener('cancel', () => clearTimeout(t)); running.push({ a, done: () => { clearTimeout(t); node.style.visibility = ''; } }); }
      else if (fx.kind === 'exit') { a.addEventListener('finish', () => { if (!a.dropped) node.style.visibility = 'hidden'; }); running.push({ a, done: () => { node.style.visibility = 'hidden'; } }); }
      else running.push({ a, done: () => {} });
    });
    return running;
  }

  // ---------- PowerPoint XML ----------
  const PPT_DIR = { bottom: 4, left: 8, right: 2, top: 1 };
  const DIR_OF = { 4: 'bottom', 8: 'left', 2: 'right', 1: 'top' };
  /** <p:timing> for a slide; spidOf(el) gives the shape id PowerPoint uses. */
  function timingXml(slide, spidOf) {
    const groups = sequence(slide);
    if (!groups.length) return '';
    let id = 2;
    const nid = () => ++id;
    const tgt = sp => `<p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl>`;
    const setVis = (sp, v, delay) => `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${tgt(sp)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${v}"/></p:to></p:set>`;
    const prop = (sp, dur, name, from, to) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt(sp)}<p:attrNameLst><p:attrName>${name}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
    function behaviours(a, sp) {
      const dur = Math.max(1, Math.round(a.dur * 1000));
      const flyFrom = { bottom: ['#ppt_x', '1+#ppt_h/2'], top: ['#ppt_x', '0-#ppt_h/2'], left: ['0-#ppt_w/2', '#ppt_y'], right: ['1+#ppt_w/2', '#ppt_y'] }[a.dir];
      const wipe = { bottom: 'up', top: 'down', left: 'right', right: 'left' }[a.dir];
      switch (a.effect) {
        case 'appear': return setVis(sp, 'visible', 0);
        case 'fade': return setVis(sp, 'visible', 0) + `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt(sp)}</p:cBhvr></p:animEffect>`;
        case 'fly': return setVis(sp, 'visible', 0) + prop(sp, dur, 'ppt_x', flyFrom[0], '#ppt_x') + prop(sp, dur, 'ppt_y', flyFrom[1], '#ppt_y');
        case 'float': return setVis(sp, 'visible', 0) + `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt(sp)}</p:cBhvr></p:animEffect>` + prop(sp, dur, 'ppt_y', '#ppt_y+0.1', '#ppt_y');
        case 'wipe': return setVis(sp, 'visible', 0) + `<p:animEffect transition="in" filter="wipe(${wipe})"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt(sp)}</p:cBhvr></p:animEffect>`;
        case 'zoom': return setVis(sp, 'visible', 0) + prop(sp, dur, 'ppt_w', '0', '#ppt_w') + prop(sp, dur, 'ppt_h', '0', '#ppt_h') + `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt(sp)}</p:cBhvr></p:animEffect>`;
        case 'pulse': return `<p:animScale><p:cBhvr><p:cTn id="${nid()}" dur="${Math.round(dur / 2)}" autoRev="1" fill="hold"/>${tgt(sp)}</p:cBhvr><p:by x="115000" y="115000"/></p:animScale>`;
        case 'spin': return `<p:animRot by="21600000"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}" fill="hold"/>${tgt(sp)}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
        case 'fadeOut': return `<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt(sp)}</p:cBhvr></p:animEffect>` + setVis(sp, 'hidden', dur - 1);
        case 'flyOut': return prop(sp, dur, 'ppt_x', '#ppt_x', flyFrom[0]) + prop(sp, dur, 'ppt_y', '#ppt_y', flyFrom[1]) + setVis(sp, 'hidden', dur - 1);
        case 'zoomOut': return prop(sp, dur, 'ppt_w', '#ppt_w', '0') + prop(sp, dur, 'ppt_h', '#ppt_h', '0') + `<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="${nid()}" dur="${dur}"/>${tgt(sp)}</p:cBhvr></p:animEffect>` + setVis(sp, 'hidden', dur - 1);
      }
      return '';
    }
    const groupXml = groups.map(g => {
      // Within a click: sub-groups by start time ("after previous" starts a new one).
      const starts = [...new Set(g.items.map(x => x.at - x.anim.delay * 1000))].sort((a, b) => a - b);
      const outerId = nid();
      const inner = starts.map(st => {
        const pid = nid();
        const effects = g.items.filter(x => x.at - x.anim.delay * 1000 === st).map((x, k) => {
          const sp = spidOf(x.el);
          if (!sp) return '';
          const fx = EFFECTS[x.anim.effect];
          const node = k === 0 && st === starts[0] && !g.auto ? 'clickEffect' : x.anim.start === 'after' ? 'afterEffect' : 'withEffect';
          const sub = fx.dirs ? PPT_DIR[x.anim.dir] : 0;
          const eid = nid();
          return `<p:par><p:cTn id="${eid}" presetID="${fx.preset}" presetClass="${fx.kind}" presetSubtype="${sub}" fill="hold" nodeType="${node}"><p:stCondLst><p:cond delay="${Math.round(x.anim.delay * 1000)}"/></p:stCondLst><p:childTnLst>${behaviours(x.anim, sp)}</p:childTnLst></p:cTn></p:par>`;
        }).join('');
        return `<p:par><p:cTn id="${pid}" fill="hold"><p:stCondLst><p:cond delay="${Math.round(st)}"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>`;
      }).join('');
      const start = g.auto ? '<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' : '<p:cond delay="indefinite"/>';
      return `<p:par><p:cTn id="${outerId}" fill="hold"><p:stCondLst>${start}</p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
    }).join('');
    return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groupXml}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
  }
  /** Effects from a slide's <p:timing>: [{ spid, anim }] in order. */
  function fromTiming(timingEl) {
    if (!timingEl) return [];
    const all = [...timingEl.getElementsByTagName('*')];
    const out = [];
    all.filter(e => e.localName === 'cTn' && e.getAttribute('presetClass')).forEach(c => {
      const cls = c.getAttribute('presetClass'), preset = +c.getAttribute('presetID'), sub = +c.getAttribute('presetSubtype') || 0;
      const node = c.getAttribute('nodeType');
      const sp = [...c.getElementsByTagName('*')].find(e => e.localName === 'spTgt');
      if (!sp) return;
      const effect = Object.keys(EFFECTS).find(k => EFFECTS[k].preset === preset && EFFECTS[k].kind === cls)
        || (cls === 'entr' ? 'fade' : cls === 'exit' ? 'fadeOut' : cls === 'emph' ? 'pulse' : null);
      if (!effect) return;
      const durs = [...c.getElementsByTagName('*')].filter(e => e.localName === 'cTn' && e.getAttribute('dur') && e.getAttribute('dur') !== 'indefinite').map(e => +e.getAttribute('dur')).filter(n => n > 1);
      const delayCond = [...c.children].find(e => e.localName === 'stCondLst');
      const delay = delayCond ? +((delayCond.firstElementChild || {}).getAttribute?.('delay')) || 0 : 0;
      out.push({ spid: sp.getAttribute('spid'), anim: clean({ effect, dir: DIR_OF[sub] || 'bottom', start: node === 'withEffect' ? 'with' : node === 'afterEffect' ? 'after' : 'click', dur: durs.length ? Math.max(...durs) / 1000 : 0.5, delay: delay / 1000, order: out.length }) });
    });
    return out.filter(x => x.anim);
  }
  /** <p:transition> for a slide (after <p:clrMapOvr>). */
  function transitionXml(s) {
    const t = s.transition;
    const adv = typeof s.advance === 'number' ? ` advTm="${Math.round(s.advance * 1000)}"` : '';
    if ((!t || t === 'none') && !adv) return '';
    const spd = (s.transDur || 1) <= 0.5 ? 'fast' : (s.transDur || 1) <= 1 ? 'med' : 'slow';
    const inner = { fade: '<p:fade/>', push: '<p:push dir="l"/>', wipe: '<p:wipe dir="r"/>', split: '<p:split orient="vert" dir="out"/>', cover: '<p:cover dir="l"/>', uncover: '<p:pull dir="l"/>', zoom: '<p:zoom/>' }[t] || '';
    return `<p:transition spd="${spd}"${adv}>${inner}</p:transition>`;
  }
  function fromTransition(el) {
    if (!el) return null;
    const kidsAll = [...el.getElementsByTagName('*')];
    const map = { fade: 'fade', push: 'push', wipe: 'wipe', split: 'split', cover: 'cover', pull: 'uncover', zoom: 'zoom', reveal: 'uncover', dissolve: 'fade', cut: 'none' };
    const k = kidsAll.find(e => map[e.localName]);
    const spd = el.getAttribute('spd'), dur = el.getAttribute('p14:dur') || el.getAttribute('dur');
    return {
      transition: k ? map[k.localName] : 'none',
      transDur: dur ? +dur / 1000 : spd === 'fast' ? 0.5 : spd === 'slow' ? 1.5 : 1,
      advance: el.getAttribute('advTm') ? +el.getAttribute('advTm') / 1000 : null
    };
  }

  const api = { EFFECTS, DIRS, TRANSITIONS, clean, sequence, visibilityAfter, play, timingXml, fromTiming, transitionXml, fromTransition };
  root.SlidesAnim = api;
})(typeof window !== 'undefined' ? window : globalThis);
