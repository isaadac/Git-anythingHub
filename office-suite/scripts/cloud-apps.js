'use strict';
// Builds the publisher's cloud app registrations into this copy of Isadek Office
// (src/renderer/cloud-apps.json), so people can connect Google Drive, OneDrive and Dropbox
// without registering apps of their own. Values come from the environment:
//
//   ISADEK_GOOGLE_CLIENT_ID, ISADEK_GOOGLE_CLIENT_SECRET (desktop OAuth clients only; Google
//   does not treat it as secret for installed apps), ISADEK_MICROSOFT_CLIENT_ID,
//   ISADEK_MICROSOFT_TENANT (default "common"), ISADEK_DROPBOX_APP_KEY
//
//   node scripts/cloud-apps.js        writes the file (an empty {} when nothing is set)
const fs = require('fs');
const path = require('path');

function fromEnv(env = process.env) {
  const v = k => String(env[k] || '').trim();
  const out = {};
  if (v('ISADEK_GOOGLE_CLIENT_ID')) out.google = { clientId: v('ISADEK_GOOGLE_CLIENT_ID'), clientSecret: v('ISADEK_GOOGLE_CLIENT_SECRET') };
  if (v('ISADEK_MICROSOFT_CLIENT_ID')) out.microsoft = { clientId: v('ISADEK_MICROSOFT_CLIENT_ID'), tenant: v('ISADEK_MICROSOFT_TENANT') || 'common' };
  if (v('ISADEK_DROPBOX_APP_KEY')) out.dropbox = { clientId: v('ISADEK_DROPBOX_APP_KEY') };
  return out;
}

module.exports = { fromEnv };

if (require.main === module) {
  const apps = fromEnv();
  const file = path.resolve(__dirname, '../src/renderer/cloud-apps.json');
  fs.writeFileSync(file, JSON.stringify(apps, null, 2) + '\n');
  console.log('Built-in cloud apps: ' + (Object.keys(apps).join(', ') || 'none'));
}
