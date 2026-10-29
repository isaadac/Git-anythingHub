'use strict';
// Builds the on-device speech worker: Transformers.js (Whisper) and ONNX Runtime Web bundled into
// one module, with ONNX Runtime's WebAssembly next to it, in src/renderer/vendor/speech/.
//   node scripts/build-speech.js
// (Transformers.js is a development dependency: its Node-only parts never ship with the app.)
const path = require('path');
const fs = require('fs');
// After a production-only install (no development dependencies) there is nothing to build with:
// with --if-possible that is fine, the app just has no on-device speech.
let esbuild;
try { esbuild = require('esbuild'); require.resolve('@huggingface/transformers'); } catch {
  if (process.argv.includes('--if-possible')) { console.log('build-speech: skipped (development dependencies not installed)'); process.exit(0); }
  throw new Error('Run npm install first (esbuild and @huggingface/transformers are development dependencies)');
}

const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'src/renderer/vendor/speech');
fs.mkdirSync(OUT, { recursive: true });
esbuild.buildSync({
  entryPoints: [path.join(__dirname, 'speech/speech-worker.src.js')],
  bundle: true, format: 'esm', platform: 'browser', target: 'es2022', minify: true, legalComments: 'eof',
  outfile: path.join(OUT, 'speech-worker.js'),
  // The web build only: no Node modules, no sharp.
  mainFields: ['browser', 'module', 'main'], conditions: ['browser'],
  external: ['sharp', 'onnxruntime-node', 'fs', 'path', 'url']
});
const ORT = path.join(ROOT, 'node_modules/onnxruntime-web/dist');
for (const f of ['ort-wasm-simd-threaded.jsep.wasm', 'ort-wasm-simd-threaded.jsep.mjs']) fs.copyFileSync(path.join(ORT, f), path.join(OUT, f));
for (const f of fs.readdirSync(OUT)) console.log(`${f}: ${(fs.statSync(path.join(OUT, f)).size / 1e6).toFixed(1)} MB`);
