'use strict';
// Small persistent settings store (userData/settings.json) shared by the main process features.
const { app } = require('electron');
const fs = require('fs');
const path = require('path');
const { writeAtomic } = require('./safe-io');

const DEFAULTS = {
  autoUpdate: true
};

let cache = null;
const file = () => path.join(app.getPath('userData'), 'settings.json');

function load() {
  if (cache) return cache;
  try {
    const saved = JSON.parse(fs.readFileSync(file(), 'utf8'));
    cache = { ...DEFAULTS, ...(saved && typeof saved === 'object' && !Array.isArray(saved) ? saved : {}) };
  } catch { cache = { ...DEFAULTS }; }
  return cache;
}

function get(key) { return load()[key]; }

function set(key, value) {
  load()[key] = value;
  try {
    fs.mkdirSync(path.dirname(file()), { recursive: true });
    writeAtomic(file(), JSON.stringify(cache, null, 2));
  } catch { /* non-fatal: settings just won't persist */ }
}

module.exports = { get, set, DEFAULTS };
