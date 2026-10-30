'use strict';
// Isadek Notes — page templates, attachments and audio notes, and merging two copies of a
// notebook for sync.
//   Attachments (files and recordings) live in the notebook: nb.files = { id: { name, type, size, data } }
//   (data in base64); a page refers to one with <span class="attachment|audio-note" data-att="id">.
//   Sync keeps a notebook as an .anote file in a cloud folder. Each device merges the cloud copy
//   with its own: page by page the newer one wins, and deletions (nb.deleted = { id: time }) carry over.
(function () {
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const today = () => new Date().toLocaleDateString(undefined, { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  const todo = t => `<p class="todo">${t}</p>`;

  // ---------------- Templates ----------------
  const TEMPLATES = {
    meeting: { label: 'Meeting notes', title: () => 'Meeting — ' + new Date().toLocaleDateString(), html: () => `<p><b>Date:</b> ${esc(today())}</p><p><b>Attendees:</b> </p><h2>Agenda</h2><ol><li><br></li></ol><h2>Notes</h2><p><br></p><h2>Decisions</h2><ul><li><br></li></ul><h2>Action items</h2>${todo('<br>')}` },
    todo: { label: 'To-do list', title: () => 'To do', html: () => `<h2>Today</h2>${todo('<br>')}${todo('<br>')}<h2>This week</h2>${todo('<br>')}<h2>Later</h2>${todo('<br>')}` },
    lecture: { label: 'Lecture notes (Cornell)', title: () => 'Lecture — ', html: () => `<p><b>Course:</b> &nbsp; <b>Date:</b> ${esc(new Date().toLocaleDateString())}</p><table><tbody><tr><th style="width:30%">Cues and questions</th><th>Notes</th></tr><tr><td><p><br></p></td><td><p><br></p></td></tr></tbody></table><h2>Summary</h2><p><br></p>` },
    journal: { label: 'Daily journal', title: () => today(), html: () => '<h2>Grateful for</h2><ul><li><br></li></ul><h2>Today\'s focus</h2><p><br></p><h2>What happened</h2><p><br></p><h2>Tomorrow</h2>' + todo('<br>') },
    week: { label: 'Weekly planner', title: () => 'Week of ' + new Date().toLocaleDateString(), html: () => `<table><tbody>${['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'].map(d => `<tr><th style="width:22%">${d}</th><td><p><br></p></td></tr>`).join('')}</tbody></table><h2>Goals</h2>${todo('<br>')}` },
    project: { label: 'Project overview', title: () => 'Project: ', html: () => '<h2>Goal</h2><p><br></p><h2>Scope</h2><ul><li><br></li></ul><h2>Milestones</h2><table><tbody><tr><th>Milestone</th><th>Due</th><th>Owner</th><th>Status</th></tr><tr><td><br></td><td><br></td><td><br></td><td><br></td></tr></tbody></table><h2>Risks</h2><ul><li><br></li></ul><h2>Tasks</h2>' + todo('<br>') },
    decision: { label: 'Decision log', title: () => 'Decision: ', html: () => `<p><b>Date:</b> ${esc(new Date().toLocaleDateString())} &nbsp; <b>Decided by:</b> </p><h2>Context</h2><p><br></p><h2>Options</h2><ol><li><br></li></ol><h2>Decision</h2><p class="tag-important"><br></p><h2>Consequences</h2><p><br></p>` }
  };

  // ---------------- Attachments ----------------
  const fmtSize = n => (n < 1024 ? n + ' B' : n < 1048576 ? (n / 1024).toFixed(0) + ' KB' : (n / 1048576).toFixed(1) + ' MB');
  const ICONS = [[/^audio\//, '🎙'], [/^image\//, '🖼'], [/pdf$/, '📕'], [/(word|document|text)/, '📄'], [/(sheet|excel|csv)/, '📊'], [/(presentation|powerpoint)/, '📽'], [/zip|compressed/, '🗜']];
  const iconFor = type => (ICONS.find(([re]) => re.test(type || '')) || [null, '📎'])[1];
  /** The chip on the page for an attachment (or an audio note). */
  function chipHtml(id, f) {
    if (/^audio\//.test(f.type)) return `<span class="audio-note" data-att="${id}" contenteditable="false" title="Play">▶ ${esc(f.name)}${f.duration ? ' · ' + fmtDuration(f.duration) : ''}</span>&nbsp;`;
    return `<span class="attachment" data-att="${id}" contenteditable="false" title="Save a copy">${iconFor(f.type)} ${esc(f.name)} <small>${fmtSize(f.size)}</small></span>&nbsp;`;
  }
  const fmtDuration = s => `${Math.floor(s / 60)}:${String(Math.round(s % 60)).padStart(2, '0')}`;
  const ATT_ID = /^f[\w]{4,40}$/;
  function cleanFiles(files) {
    const out = {};
    if (!files || typeof files !== 'object') return out;
    for (const [id, f] of Object.entries(files)) {
      // (Base64 checked at its start and end: the whole of a large file would take a while.)
      if (!ATT_ID.test(id) || !f || typeof f !== 'object' || typeof f.data !== 'string' || f.data.length > 70e6) continue;
      const bare = f.data.replace(/={1,2}$/, '');
      if (!/^[A-Za-z0-9+/]*$/.test(bare.slice(0, 100)) || !/^[A-Za-z0-9+/]*$/.test(bare.slice(-100))) continue;
      out[id] = { name: String(f.name || 'file').slice(0, 255), type: /^[\w.+-]+\/[\w.+-]+$/.test(f.type || '') ? f.type : 'application/octet-stream', size: Math.max(0, Math.floor(+f.size || 0)), data: f.data, ...(Number.isFinite(+f.duration) && +f.duration > 0 ? { duration: +f.duration } : {}) };
    }
    return out;
  }
  /** Ids of attachments the pages still use. */
  function usedFiles(nb) {
    const used = new Set();
    for (const s of nb.sections) for (const p of s.pages) for (const m of (p.html || '').matchAll(/data-att="(f[\w]{4,40})"/g)) used.add(m[1]);
    return used;
  }

  // ---------------- Merging (sync) ----------------
  /** Two copies of one notebook → one, keeping the newer of every page and section. */
  function merge(a, b) {
    const deleted = { ...(a.deleted || {}) };
    for (const [k, t] of Object.entries(b.deleted || {})) deleted[k] = Math.max(deleted[k] || 0, t);
    const newer = (x, y) => ((y && (y.modified || 0) > (x ? x.modified || 0 : -1)) ? y : x);
    const head = newer(a, b);
    const out = { ...a, name: head.name, color: head.color, modified: Math.max(a.modified || 0, b.modified || 0), deleted, files: { ...(b.files || {}), ...(a.files || {}) } };
    // Pages: the newer copy of each, wherever (in whichever section) that copy is.
    const best = new Map();
    for (const nbk of [a, b]) for (const s of nbk.sections) for (const p of s.pages) { const cur = best.get(p.id); if (!cur || (p.modified || 0) > (cur.p.modified || 0)) best.set(p.id, { p, sid: s.id }); }
    const gone = (id, t) => deleted[id] && deleted[id] >= (t || 0);
    // Items only one side has go next to their neighbour there (else at the end).
    const weave = (into, list) => list.forEach((x, i) => {
      if (into.includes(x)) return;
      const prev = list.slice(0, i).reverse().find(y => into.includes(y));
      const next = list.slice(i + 1).find(y => into.includes(y));
      if (prev !== undefined) into.splice(into.indexOf(prev) + 1, 0, x);
      else if (next !== undefined) into.splice(into.indexOf(next), 0, x);
      else into.push(x);
    });
    const order = [];
    weave(order, a.sections.map(s => s.id)); weave(order, b.sections.map(s => s.id));
    const secA = new Map(a.sections.map(s => [s.id, s])), secB = new Map(b.sections.map(s => [s.id, s]));
    out.sections = order.map(id => {
      const sa = secA.get(id), sb = secB.get(id);
      const meta = newer(sa, sb) || sa || sb;
      if (gone(id, Math.max(sa ? sa.modified || 0 : 0, sb ? sb.modified || 0 : 0))) return null;
      const ids = [];
      weave(ids, (sa ? sa.pages : []).map(p => p.id)); weave(ids, (sb ? sb.pages : []).map(p => p.id));
      const pages = ids.map(pid => best.get(pid)).filter(x => x && x.sid === id && !gone(x.p.id, x.p.modified)).map(x => x.p);
      return { ...meta, pages };
    }).filter(Boolean);
    // A section left without pages gets an empty one (a notebook always shows a page).
    out.sections.forEach(s => { if (!s.pages.length) s.pages.push({ id: 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7), title: '', html: '<p><br></p>', ink: [], level: 0, created: Date.now(), modified: 0 }); });
    if (!out.sections.length) out.sections = a.sections.slice(0, 1);
    return out;
  }

  window.NotesExtras = { TEMPLATES, chipHtml, cleanFiles, usedFiles, merge, fmtSize, fmtDuration };
})();
