// In-memory mock of Google Drive, OneDrive, Dropbox and a Nextcloud WebDAV server, plus static files.
const http = require('http'), fs = require('fs'), path = require('path'), crypto = require('crypto');
module.exports = function start(root, files) {
  const S = {
    challenges: {}, codes: {}, tokens: new Set(), refreshes: 0, log: [],
    google: { root: [], files: {} }, dropbox: {}, ms: {}, dav: {}, perms: [], shares: [], expiresIn: 3600
  };
  // seed
  for (const [p, bytes] of Object.entries(files)) {
    if (p.startsWith('google/')) { const id = 'g' + Object.keys(S.google.files).length; S.google.files[id] = { id, name: p.slice(7), parents: ['root'], mimeType: 'application/octet-stream', data: bytes }; }
    if (p.startsWith('dropbox/')) S.dropbox['/' + p.slice(8)] = bytes;
    if (p.startsWith('ms/')) { const id = 'm' + Object.keys(S.ms).length; S.ms[id] = { id, name: p.slice(3), parent: 'root', data: bytes }; }
    if (p.startsWith('dav/')) S.dav['/remote.php/dav/files/u/' + p.slice(4).split('/').map(encodeURIComponent).join('/')] = bytes;
  }
  S.google.files.gf = { id: 'gf', name: 'Projects', parents: ['root'], mimeType: 'application/vnd.google-apps.folder' };
  S.google.files.gn = { id: 'gn', name: 'Native Sheet', parents: ['root'], mimeType: 'application/vnd.google-apps.spreadsheet', data: files.__nativeXlsx };
  const json = (res, code, v) => { res.writeHead(code, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }); res.end(JSON.stringify(v)); };
  const authed = req => S.tokens.has((req.headers.authorization || '').replace(/^Bearer /, ''));
  const issue = () => { const t = crypto.randomBytes(8).toString('hex'); S.tokens.add(t); return t; };
  const server = http.createServer(async (req, res) => {
    const u = new URL(req.url, 'http://x');
    const chunks = []; for await (const c of req) chunks.push(c);
    const body = Buffer.concat(chunks);
    S.log.push(req.method + ' ' + u.pathname);
    const p = u.pathname;
    // ---- OAuth (shared by all three)
    let m;
    if ((m = /^\/mock\/(google|ms|dropbox)\/auth$/.exec(p))) {
      const code = 'c' + crypto.randomBytes(6).toString('hex');
      S.codes[code] = { challenge: u.searchParams.get('code_challenge'), redirect: u.searchParams.get('redirect_uri'), client: u.searchParams.get('client_id'), provider: m[1], scope: u.searchParams.get('scope') };
      const r = new URL(u.searchParams.get('redirect_uri'));
      r.searchParams.set('code', code); r.searchParams.set('state', u.searchParams.get('state'));
      res.writeHead(302, { Location: r.toString() }); return res.end();
    }
    if ((m = /^\/mock\/(google|ms|dropbox)\/token$/.exec(p)) && req.method === 'POST') {
      const f = new URLSearchParams(body.toString());
      if (f.get('grant_type') === 'refresh_token') {
        if (f.get('refresh_token') !== 'refresh-' + m[1]) return json(res, 400, { error: 'invalid_grant' });
        S.refreshes++;
        return json(res, 200, { access_token: issue(), expires_in: 3600 });
      }
      const c = S.codes[f.get('code')];
      if (!c) return json(res, 400, { error: 'invalid_grant', error_description: 'bad code' });
      const ch = crypto.createHash('sha256').update(f.get('code_verifier') || '').digest('base64url');
      if (ch !== c.challenge || f.get('redirect_uri') !== c.redirect || f.get('client_id') !== c.client) return json(res, 400, { error: 'invalid_grant', error_description: 'pkce mismatch' });
      delete S.codes[f.get('code')];
      return json(res, 200, { access_token: issue(), refresh_token: 'refresh-' + m[1], expires_in: S.expiresIn });
    }
    // ---- Google Drive
    if (p.startsWith('/mock/google/')) {
      if (!authed(req)) return json(res, 401, { error: { message: 'Invalid Credentials' } });
      const G = S.google.files;
      if (p === '/mock/google/api/about') return json(res, 200, { user: { emailAddress: 'ada@gmail.test' } });
      if (p === '/mock/google/api/files' && req.method === 'GET') {
        const parent = /'([^']+)' in parents/.exec(u.searchParams.get('q'))[1];
        return json(res, 200, { files: Object.values(G).filter(f => f.parents.includes(parent)).map(f => ({ id: f.id, name: f.name, mimeType: f.mimeType, size: f.data ? String(f.data.length) : undefined, modifiedTime: '2026-10-01T10:00:00Z' })) });
      }
      if ((m = /^\/mock\/google\/api\/files\/([^/]+)\/export$/.exec(p))) return res.end(G[m[1]].data);
      if ((m = /^\/mock\/google\/api\/files\/([^/]+)\/permissions$/.exec(p))) { S.perms.push({ id: m[1], ...JSON.parse(body), notify: u.searchParams.get('sendNotificationEmail') }); return json(res, 200, { id: 'p1' }); }
      if ((m = /^\/mock\/google\/api\/files\/([^/]+)$/.exec(p))) {
        if (u.searchParams.get('alt') === 'media') { res.writeHead(200); return res.end(G[m[1]].data); }
        return json(res, 200, { webViewLink: `https://drive.test/file/d/${m[1]}/view` });
      }
      if (p === '/mock/google/upload/files' && req.method === 'POST') {
        const boundary = /boundary=(\S+)/.exec(req.headers['content-type'])[1];
        const s = body.toString('latin1');
        const parts = s.split('--' + boundary);
        const meta = JSON.parse(parts[1].split('\r\n\r\n')[1].trim());
        const media = parts[2].slice(parts[2].indexOf('\r\n\r\n') + 4, -2);
        const id = 'g' + crypto.randomBytes(3).toString('hex');
        G[id] = { id, name: meta.name, parents: meta.parents, mimeType: /Content-Type: (\S+)/.exec(parts[2])[1], data: Buffer.from(media, 'latin1') };
        return json(res, 200, { id, name: meta.name });
      }
      if ((m = /^\/mock\/google\/upload\/files\/([^/]+)$/.exec(p)) && req.method === 'PATCH') { G[m[1]].data = body; G[m[1]].saves = (G[m[1]].saves || 0) + 1; return json(res, 200, { id: m[1], name: G[m[1]].name }); }
    }
    // ---- OneDrive
    if (p.startsWith('/mock/ms/api/')) {
      if (!authed(req)) return json(res, 401, { error: { code: 'InvalidAuthenticationToken', message: 'expired' } });
      if (p === '/mock/ms/api/me') return json(res, 200, { userPrincipalName: 'ada@outlook.test' });
      if ((m = /^\/mock\/ms\/api\/me\/drive\/items\/([^/:]+)\/children$/.exec(p))) return json(res, 200, { value: Object.values(S.ms).filter(f => f.parent === m[1]).map(f => ({ id: f.id, name: f.name, size: f.data.length, file: {} })) });
      if ((m = /^\/mock\/ms\/api\/me\/drive\/items\/([^/:]+):\/([^:]+):\/content$/.exec(p))) { const id = 'm' + crypto.randomBytes(3).toString('hex'); S.ms[id] = { id, name: decodeURIComponent(m[2]), parent: m[1], data: body }; return json(res, 201, { id, name: S.ms[id].name }); }
      if ((m = /^\/mock\/ms\/api\/me\/drive\/items\/([^/:]+)\/content$/.exec(p))) {
        if (req.method === 'PUT') { S.ms[m[1]].data = body; return json(res, 200, { id: m[1], name: S.ms[m[1]].name }); }
        res.writeHead(302, { Location: `/mock/ms/blob/${m[1]}` }); return res.end();
      }
      if ((m = /^\/mock\/ms\/api\/me\/drive\/items\/([^/:]+)\/createLink$/.exec(p))) { S.shares.push({ ms: m[1], ...JSON.parse(body) }); return json(res, 200, { link: { webUrl: `https://1drv.test/${m[1]}` } }); }
    }
    if ((m = /^\/mock\/ms\/blob\/(.+)$/.exec(p))) { res.writeHead(200); return res.end(S.ms[m[1]].data); }
    // ---- Dropbox
    if (p.startsWith('/mock/dropbox/')) {
      if (!authed(req)) return json(res, 401, { error_summary: 'expired_access_token/' });
      const arg = req.headers['dropbox-api-arg'] ? JSON.parse(req.headers['dropbox-api-arg']) : null;
      if (p === '/mock/dropbox/api/users/get_current_account') return body.length ? json(res, 400, { error_summary: 'body not allowed' }) : json(res, 200, { email: 'ada@dropbox.test' });
      if (p === '/mock/dropbox/api/files/list_folder') { const { path: dir } = JSON.parse(body); return json(res, 200, { has_more: false, entries: Object.entries(S.dropbox).filter(([k]) => k.slice(0, k.lastIndexOf('/')) === dir).map(([k, v]) => ({ '.tag': 'file', name: k.split('/').pop(), path_display: k, size: v.length })) }); }
      if (p === '/mock/dropbox/content/files/download') { res.writeHead(200); return res.end(S.dropbox[arg.path]); }
      if (p === '/mock/dropbox/content/files/upload') { S.dropbox[arg.path] = body; S.lastDropboxArg = arg; return json(res, 200, { path_display: arg.path, name: arg.path.split('/').pop() }); }
      if (p === '/mock/dropbox/api/sharing/create_shared_link_with_settings') return json(res, 409, { error_summary: 'shared_link_already_exists/', error: { '.tag': 'shared_link_already_exists' } });
      if (p === '/mock/dropbox/api/sharing/list_shared_links') return json(res, 200, { links: [{ url: 'https://dropbox.test/s/' + encodeURIComponent(JSON.parse(body).path) }] });
    }
    // ---- Nextcloud WebDAV + OCS
    if (p.startsWith('/remote.php/dav/files/u/') || p.startsWith('/ocs/')) {
      const ok = req.headers.authorization === 'Basic ' + Buffer.from('u:secret').toString('base64');
      if (!ok) { res.writeHead(401, { 'WWW-Authenticate': 'Basic' }); return res.end(); }
      if (p.startsWith('/ocs/')) { const f = new URLSearchParams(body.toString()); S.shares.push({ dav: f.get('path'), type: f.get('shareType'), perms: f.get('permissions'), api: req.headers['ocs-apirequest'] }); return json(res, 200, { ocs: { data: { url: 'https://cloud.test/s/abc' } } }); }
      if (req.method === 'PROPFIND') {
        const dir = p.endsWith('/') ? p : p + '/';
        const kids = Object.keys(S.dav).filter(k => k.startsWith(dir) && !k.slice(dir.length).includes('/'));
        const resp = (href, folder, size) => `<d:response><d:href>${href}</d:href><d:propstat><d:prop><d:resourcetype>${folder ? '<d:collection/>' : ''}</d:resourcetype>${folder ? '' : `<d:getcontentlength>${size}</d:getcontentlength>`}<d:getlastmodified>Thu, 01 Oct 2026 10:00:00 GMT</d:getlastmodified></d:prop></d:propstat></d:response>`;
        res.writeHead(207, { 'Content-Type': 'application/xml' });
        return res.end(`<?xml version="1.0"?><d:multistatus xmlns:d="DAV:">${resp(dir, true)}${kids.map(k => resp(k, false, S.dav[k].length)).join('')}</d:multistatus>`);
      }
      if (req.method === 'GET') { if (!S.dav[p]) { res.writeHead(404); return res.end(); } res.writeHead(200); return res.end(S.dav[p]); }
      if (req.method === 'PUT') {
        if (req.headers['if-none-match'] === '*' && S.dav[p]) { res.writeHead(412); return res.end(); }
        S.dav[p] = body; res.writeHead(201); return res.end();
      }
    }
    // ---- static
    const file = path.join(root, decodeURIComponent(p));
    if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); return res.end('nf'); }
    const type = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml', '.wasm': 'application/wasm' }[path.extname(file)] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': type });
    fs.createReadStream(file).pipe(res);
  });
  return new Promise(r => server.listen(0, '127.0.0.1', () => r({ port: server.address().port, S, server })));
};
