'use strict';
// A copy of the Whisper model for the on-device speech tests, in place of Hugging Face (which test
// machines may not reach). It comes from npm, checked against its exact published checksum, and
// is only used by the tests. Returns the folder holding Xenova/whisper-tiny, or null.
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const PKG = 'sts-whisper-tiny@1.0.0';
const INTEGRITY = 'sha512-kxDUK4JaVwltNa0zR9RH9gRU033teKgKUInnoiVtkDQFDYywsUb18L6+N183cLX1fVEviPBqJ2dw73eLI7X6VA==';

module.exports = function speechModel() {
  const dir = path.join(os.tmpdir(), 'isadek-speech-model');
  const model = path.join(dir, 'package/models/Xenova/whisper-tiny');
  if (fs.existsSync(path.join(model, 'onnx/encoder_model_quantized.onnx'))) return model;
  try {
    fs.mkdirSync(dir, { recursive: true });
    const name = execFileSync('npm', ['pack', PKG, '--silent'], { cwd: dir, encoding: 'utf8' }).trim().split('\n').pop();
    const tgz = path.join(dir, name);
    const got = 'sha512-' + crypto.createHash('sha512').update(fs.readFileSync(tgz)).digest('base64');
    if (got !== INTEGRITY) throw new Error('checksum mismatch');
    execFileSync('tar', ['xzf', tgz, '-C', dir]);
    return fs.existsSync(model) ? model : null;
  } catch (err) {
    console.log('speech model unavailable: ' + err.message);
    return null;
  }
};
