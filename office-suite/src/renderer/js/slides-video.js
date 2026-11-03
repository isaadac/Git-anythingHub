'use strict';
// Isadek Slides — Export to Video: the slide show plays by itself (recorded timings and
// narration where there are any, otherwise a few seconds a slide, builds one after another) and
// is recorded as it plays, so transitions, animations, Morph and video clips look as they do in
// the show. The desktop app records its own window, frame by frame (never the screen); in a
// browser the person shares this tab. MP4 where the browser can make it, otherwise WebM.
(function () {
  const TYPES = ['video/mp4;codecs=avc1.640028', 'video/mp4;codecs=avc1', 'video/mp4', 'video/webm;codecs=vp9', 'video/webm'];
  const wait = ms => new Promise(r => setTimeout(r, ms));

  function SlidesVideo(ctx) {
    let job = null; // { recorder, chunks, canvas, g, audio, dest, stopFrames, done, cancelled, size }

    /** The slide's rectangle on screen, in CSS pixels. */
    const slideRect = () => ctx.slideNode().getBoundingClientRect();

    // Frames from the desktop app's own window.
    async function desktopFrames(j) {
      while (!j.ended) {
        const t0 = performance.now();
        const r = slideRect();
        try {
          const jpg = await window.atlas.captureFrame({ x: r.left, y: r.top, width: r.width, height: r.height });
          if (jpg && !j.ended) {
            const bmp = await createImageBitmap(new Blob([jpg], { type: 'image/jpeg' }));
            j.g.drawImage(bmp, 0, 0, j.canvas.width, j.canvas.height);
            bmp.close();
          }
        } catch { /* a frame missed */ }
        await wait(Math.max(0, 1000 / 30 - (performance.now() - t0)));
      }
    }
    // Frames from this tab, shared by the person (the browser asks), drawn on a steady clock.
    async function tabFrames(j) {
      const v = document.createElement('video');
      v.muted = true; v.playsInline = true; v.srcObject = j.display;
      await v.play();
      const draw = () => {
        if (!v.videoWidth) return;
        const r = slideRect(), kx = v.videoWidth / innerWidth, ky = v.videoHeight / innerHeight;
        j.g.drawImage(v, r.left * kx, r.top * ky, r.width * kx, r.height * ky, 0, 0, j.canvas.width, j.canvas.height);
      };
      // The recording starts on a picture of the slide, not black.
      for (let i = 0; i < 60 && !v.videoWidth; i++) await wait(50);
      draw();
      j.timer = setInterval(() => { if (j.ended) clearInterval(j.timer); else draw(); }, 1000 / 30);
    }

    async function start() {
      if (job) return;
      if (!window.MediaRecorder) return UI.toast('This browser can’t record video', 4000);
      const desktop = !!(window.atlas && window.atlas.captureFrame);
      const anyTimings = ctx.pres().slides.some(s => typeof s.advance === 'number');
      const v = await UI.form('Export to Video', [
        { name: 'secs', label: anyTimings ? 'Seconds on slides without recorded timings' : 'Seconds on each slide', type: 'number', value: 5, min: 1, max: 600 },
        { name: 'size', label: 'Size', type: 'select', options: [{ value: '1080', label: 'Full HD (1920 × 1080)' }, { value: '720', label: 'HD (1280 × 720)' }], value: '1080' }
      ], 'Export');
      if (!v) return;
      const secs = Math.max(1, Math.min(600, +v.secs || 5));
      const h = v.size === '720' ? 720 : 1080, w = Math.round(h * 16 / 9);
      const mimeType = TYPES.find(t => MediaRecorder.isTypeSupported(t));
      if (!mimeType) return UI.toast('This browser can’t record video', 4000);
      const j = { secs, chunks: [], ended: false, cancelled: false, mimeType };
      if (!desktop) {
        UI.toast('Choose this tab to share: the show is recorded from it', 4000);
        try {
          j.display = await navigator.mediaDevices.getDisplayMedia({ video: { frameRate: 30 }, audio: false, preferCurrentTab: true, selfBrowserSurface: 'include', surfaceSwitching: 'exclude', monitorTypeSurfaces: 'exclude' });
        } catch { return UI.toast('Recording needs this tab to be shared', 4000); }
        // If the person stops sharing, the export ends.
        j.display.getVideoTracks()[0].addEventListener('ended', () => { if (job === j && !j.ended) { j.cancelled = true; ctx.endShow(); } });
      }
      // Pictures go through a canvas the size of the video: the desktop app draws its window's
      // frames on it; in a browser, the shared tab's, cut to the slide.
      j.canvas = Object.assign(document.createElement('canvas'), { width: w, height: h });
      j.g = j.canvas.getContext('2d');
      j.g.fillStyle = '#000'; j.g.fillRect(0, 0, w, h);
      // Sound: clips and narration in the show are mixed into the recording, from the first slide
      // on. Constant silence underneath keeps the sound track going between clips (a recorder
      // holds the pictures back while it waits for sound). A deck without clips has no sound track.
      const sound = ctx.pres().slides.some(s => s.elements.some(function has(e) { return e.type === 'media' || (e.children || []).some(has); }));
      if (sound) {
        j.audio = new AudioContext();
        try { await j.audio.resume(); } catch { /* resumes with the first clip */ }
        j.dest = j.audio.createMediaStreamDestination();
        const silence = j.audio.createConstantSource();
        silence.offset.value = 0;
        silence.connect(j.dest);
        silence.start();
      }
      job = j;
      document.body.classList.add('exporting-video');
      ctx.present(0);
      await wait(700); // full screen settles before the first frame
      if (j.ended) return;
      if (desktop) { desktopFrames(j); await wait(150); } else await tabFrames(j);
      if (j.ended) return;
      // The recorder starts once real frames are on the canvas, so the video starts on the slide.
      const track = j.canvas.captureStream(30).getVideoTracks()[0];
      const stream = new MediaStream([track, ...(sound ? j.dest.stream.getAudioTracks() : [])]);
      j.recorder = new MediaRecorder(stream, { mimeType, videoBitsPerSecond: h >= 1080 ? 10e6 : 6e6 });
      j.recorder.ondataavailable = e => { if (e.data && e.data.size) j.chunks.push(e.data); };
      j.done = new Promise(r => { j.recorder.onstop = r; });
      // The first slide's time starts again as recording begins, so it gets all of it.
      j.recorder.start(1000);
      ctx.restart();
    }

    /** The show ended: the end of the deck (save the video) or Esc (discard it). */
    async function showEnded(finished) {
      const j = job;
      if (!j) return;
      job = null;
      j.ended = true;
      document.body.classList.remove('exporting-video');
      if (j.recorder && j.recorder.state !== 'inactive') { j.recorder.stop(); await j.done; }
      clearInterval(j.timer);
      if (j.display) j.display.getTracks().forEach(t => t.stop());
      if (j.audio) j.audio.close().catch(() => {});
      if (!finished || j.cancelled || !j.chunks.length) return UI.toast('Video export cancelled');
      const blob = new Blob(j.chunks, { type: j.mimeType.split(';')[0] });
      const ext = j.mimeType.startsWith('video/mp4') ? 'mp4' : 'webm';
      const data = new Uint8Array(await blob.arrayBuffer());
      await Platform.save({ data, name: `${Platform.baseName(ctx.name())}.${ext}`, filters: [{ name: ext === 'mp4' ? 'MPEG-4 Video' : 'WebM Video', extensions: [ext] }], saveAs: true, noVersion: true });
      if (ext === 'webm') UI.toast('Saved as WebM (this browser can’t make MP4)', 4000);
    }
    /** A clip that plays in the show is heard in the recording too (and still in the room). */
    function addSound(media) {
      if (!job || !job.audio || media.__mixed) return;
      try {
        const src = job.audio.createMediaElementSource(media);
        src.connect(job.dest); src.connect(job.audio.destination);
        media.__mixed = true;
      } catch { /* already routed */ }
    }

    return {
      start, showEnded, addSound,
      get active() { return !!job; },
      /** Seconds a slide stays when exporting (null: not exporting). */
      get secs() { return job ? job.secs : null; }
    };
  }
  window.SlidesVideo = SlidesVideo;
})();
