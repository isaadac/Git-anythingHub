'use strict';
// Isadek Slides — speech while presenting: live captions, and Rehearse with Coach, which listens to
// a practice run and reports on pace, filler words, reading the slides aloud and timing.
//
// listen() wraps the platform's speech recognizer (restarting it after pauses, which end a
// recognizer's session); analyze() is the coach's report, worked out from what was heard.
(function (root) {
  // Words and phrases that fill silence. ("Like" and "so" are left out: they're usually real words.)
  const FILLERS = ['um', 'umm', 'uh', 'uhh', 'erm', 'er', 'ah', 'hmm', 'basically', 'actually', 'literally', 'you know', 'i mean', 'kind of', 'sort of', 'okay so'];
  const PACE = { slow: 100, fast: 165 }; // words a minute that most audiences follow comfortably
  const words = t => (String(t || '').toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g, '').match(/[a-z0-9']+/g) || []);

  /** Filler words in a text: Map(filler -> count). */
  function fillersIn(text) {
    const w = words(text), out = new Map();
    for (let i = 0; i < w.length; i++) {
      const two = w[i] + ' ' + (w[i + 1] || '');
      const hit = FILLERS.find(f => (f.includes(' ') ? two === f : w[i] === f));
      if (!hit) continue;
      out.set(hit, (out.get(hit) || 0) + 1);
      if (hit.includes(' ')) i++;
    }
    return out;
  }

  /** How much of a slide's text was read out word for word (0..1): its 4-word runs that were said. */
  function readAloud(slideText, spoken) {
    const s = words(slideText), sp = words(spoken);
    if (s.length < 8 || sp.length < 4) return 0;
    const runs = new Set();
    for (let i = 0; i + 4 <= sp.length; i++) runs.add(sp.slice(i, i + 4).join(' '));
    let n = 0, total = 0;
    for (let i = 0; i + 4 <= s.length; i++) { total++; if (runs.has(s.slice(i, i + 4).join(' '))) n++; }
    return total ? n / total : 0;
  }

  /**
   * The coach's report. input: { slides: [{ text }], heard: [{ slide, text }], times: [seconds on
   * each slide], seconds: the whole run }.
   */
  function analyze({ slides = [], heard = [], times = [], seconds = 0 }) {
    const all = heard.map(h => h.text).join(' ');
    const count = words(all).length;
    const talking = Math.max(1, seconds);
    const wpm = Math.round(count / (talking / 60));
    const pace = count < 20 ? 'unknown' : wpm < PACE.slow ? 'slow' : wpm > PACE.fast ? 'fast' : 'good';
    const fill = fillersIn(all);
    const fillers = { total: [...fill.values()].reduce((a, b) => a + b, 0), top: [...fill.entries()].sort((a, b) => b[1] - a[1]).slice(0, 5) };
    const perSlide = slides.map((s, i) => heard.filter(h => h.slide === i).map(h => h.text).join(' '));
    const reading = slides.map((s, i) => ({ slide: i, share: readAloud(s.text, perSlide[i]) })).filter(r => r.share >= 0.5).map(r => r.slide);
    const visited = times.map((t, i) => ({ slide: i, seconds: Math.round(t || 0) })).filter(t => times[t.slide] > 0);
    const longest = visited.reduce((a, b) => (!a || b.seconds > a.seconds ? b : a), null);
    const quiet = slides.map((s, i) => i).filter(i => times[i] >= 10 && words(perSlide[i]).length < 3);
    const tips = [];
    if (pace === 'fast') tips.push(`You spoke at about ${wpm} words a minute. Slowing to around ${PACE.fast - 25} gives people time to take in each point.`);
    if (pace === 'slow') tips.push(`You spoke at about ${wpm} words a minute. A little faster (about ${PACE.slow + 20}) keeps the energy up.`);
    if (fillers.total >= 3) tips.push(`You said ${fillers.top.map(([w, n]) => `“${w}” ${n}×`).join(', ')}. A short pause works better than a filler word.`);
    if (reading.length) tips.push(`You read slide${reading.length === 1 ? '' : 's'} ${reading.map(i => i + 1).join(', ')} word for word. Talk around the points instead; the audience can read them.`);
    if (longest && visited.length > 2 && longest.seconds > 2.5 * (seconds / visited.length)) tips.push(`Slide ${longest.slide + 1} took ${fmt(longest.seconds)}, much longer than the others. Consider splitting it.`);
    if (quiet.length) tips.push(`You said little on slide${quiet.length === 1 ? '' : 's'} ${quiet.map(i => i + 1).join(', ')}.`);
    if (!tips.length && pace !== 'unknown') tips.push('Well paced, few fillers, and you talked to your slides rather than reading them. Nice work.');
    return { words: count, wpm, pace, fillers, reading, times: visited, longest, seconds: Math.round(seconds), tips };
  }
  const fmt = s => `${Math.floor(s / 60)}:${String(Math.round(s % 60)).padStart(2, '0')}`;

  /**
   * Listens until stopped: onText(text, final) for every phrase heard. Returns stop(), or null when
   * there's no recognizer here (onError says why).
   */
  function listen(onText, onError, lang) {
    const SR = root.SpeechRecognition || root.webkitSpeechRecognition;
    if (!SR) { onError && onError('Speech recognition isn\'t available here. Try Chrome, Edge or Safari.'); return null; }
    let stopped = false, rec = null, fails = 0;
    const start = () => {
      rec = new SR();
      rec.continuous = true;
      rec.interimResults = true;
      rec.lang = lang || (root.navigator && navigator.language) || 'en-US';
      rec.onresult = e => { fails = 0; for (let i = e.resultIndex; i < e.results.length; i++) onText(e.results[i][0].transcript, e.results[i].isFinal); };
      rec.onerror = e => {
        if (e.error === 'no-speech' || e.error === 'aborted') return;
        fails++;
        const why = { 'not-allowed': 'the microphone isn\'t allowed', 'service-not-allowed': 'the microphone isn\'t allowed', network: 'the speech service can\'t be reached', 'audio-capture': 'no microphone was found' }[e.error] || e.error;
        if (['not-allowed', 'service-not-allowed', 'audio-capture'].includes(e.error) || fails > 3) { stopped = true; onError && onError('Listening stopped: ' + why); }
      };
      // A recognizer ends itself after a pause; carry on until told to stop.
      rec.onend = () => { if (!stopped) setTimeout(() => { if (!stopped) { try { start(); } catch { /* try again on the next end */ } } }, 250); };
      try { rec.start(); } catch (err) { stopped = true; onError && onError('Listening could not start: ' + err.message); }
    };
    start();
    return () => { stopped = true; try { rec && rec.stop(); } catch { /* already stopped */ } };
  }

  const api = { analyze, fillersIn, readAloud, listen, words, fmt, FILLERS, PACE };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.SlidesCoach = api;
})(typeof window !== 'undefined' ? window : globalThis);
