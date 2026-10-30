'use strict';
// Cloud storage: Google Drive, OneDrive, Dropbox and WebDAV (Nextcloud, ownCloud, …) accounts.
// Sign-in (OAuth 2 with PKCE), browsing, opening, saving back, and share links/invitations.
// A document from the cloud carries a path like cloud://<account>/<file id>/<folder id>/<name>,
// so the apps' ordinary Save writes back to it (see Platform.save).
(function () {
  const desktop = Platform.desktop;
  const enc = new TextEncoder();
  const PROVIDERS = {
    google: { label: 'Google Drive', badge: 'G', color: '#1a73e8' },
    microsoft: { label: 'OneDrive', badge: 'O', color: '#0364b8' },
    dropbox: { label: 'Dropbox', badge: 'D', color: '#0061fe' },
    webdav: { label: 'WebDAV', badge: 'W', color: '#5c5c5c' }
  };
  const DEFAULTS = {
    google: { auth: 'https://accounts.google.com/o/oauth2/v2/auth', token: 'https://oauth2.googleapis.com/token', api: 'https://www.googleapis.com/drive/v3', upload: 'https://www.googleapis.com/upload/drive/v3' },
    microsoft: { auth: 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize', token: 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token', api: 'https://graph.microsoft.com/v1.0' },
    dropbox: { auth: 'https://www.dropbox.com/oauth2/authorize', token: 'https://api.dropboxapi.com/oauth2/token', api: 'https://api.dropboxapi.com/2', content: 'https://content.dropboxapi.com/2' }
  };
  const EXT_APP = { docx: 'write', html: 'write', htm: 'write', txt: 'write', md: 'write',
    xlsx: 'sheets', xls: 'sheets', xlsm: 'sheets', csv: 'sheets', ods: 'sheets', tsv: 'sheets', pptx: 'slides', aslides: 'slides' };
  const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', ods: 'application/vnd.oasis.opendocument.spreadsheet',
    csv: 'text/csv', tsv: 'text/tab-separated-values', html: 'text/html', htm: 'text/html', md: 'text/markdown', txt: 'text/plain', xls: 'application/vnd.ms-excel' };
  // Google's own formats, opened as their Office equivalents.
  const GOOGLE_NATIVE = {
    'application/vnd.google-apps.document': { ext: 'docx', mime: MIME.docx },
    'application/vnd.google-apps.spreadsheet': { ext: 'xlsx', mime: MIME.xlsx },
    'application/vnd.google-apps.presentation': { ext: 'pptx', mime: MIME.pptx }
  };
  const GFOLDER = 'application/vnd.google-apps.folder';
  const mimeOf = name => MIME[Platform.extOf(name)] || 'application/octet-stream';
  const appFor = name => EXT_APP[Platform.extOf(name)] || null;

  const b64url = bytes => btoa(String.fromCharCode(...bytes)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const randomId = n => b64url(crypto.getRandomValues(new Uint8Array(n)));
  const toBytes = d => (d instanceof Uint8Array ? d : d instanceof ArrayBuffer ? new Uint8Array(d) : enc.encode(String(d)));
  const concat = parts => {
    const bytes = parts.map(toBytes);
    const out = new Uint8Array(bytes.reduce((n, b) => n + b.length, 0));
    let at = 0;
    for (const b of bytes) { out.set(b, at); at += b.length; }
    return out;
  };
  // Dropbox-API-Arg is a header, so its JSON must be ASCII.
  const headerJson = v => JSON.stringify(v).replace(/[\u007f-\uffff]/g, c => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));

  // ---------- Settings and accounts ----------
  function cleanConfig(c) {
    const str = (v, n = 300) => (typeof v === 'string' ? v.trim().slice(0, n) : '');
    const out = {
      google: { clientId: str(c?.google?.clientId), clientSecret: str(c?.google?.clientSecret) },
      microsoft: { clientId: str(c?.microsoft?.clientId), tenant: /^[\w.-]{1,100}$/.test(str(c?.microsoft?.tenant)) ? str(c.microsoft.tenant) : 'common' },
      dropbox: { clientId: str(c?.dropbox?.clientId) },
      endpoints: {}
    };
    for (const [p, e] of Object.entries(c?.endpoints || {})) {
      if (!DEFAULTS[p] || !e || typeof e !== 'object') continue;
      out.endpoints[p] = {};
      for (const [k, v] of Object.entries(e)) if (/^\w{1,40}$/.test(k) && /^https?:\/\//.test(String(v))) out.endpoints[p][k] = String(v);
    }
    return out;
  }
  /** Your own app registrations (Cloud settings). */
  async function userConfig() {
    if (desktop) return cleanConfig(await window.atlas.cloudConfigGet());
    try { return cleanConfig(JSON.parse(localStorage.getItem('atlas-cloud-config') || '{}')); } catch { return cleanConfig({}); }
  }
  /** The publisher's registrations built into this copy (cloud-apps.json), used where you have none. */
  let builtInCache = null;
  async function builtInConfig() {
    if (builtInCache) return builtInCache;
    try {
      if (desktop) builtInCache = cleanConfig(await window.atlas.cloudBuiltIn());
      else { const r = await fetch('cloud-apps.json', { cache: 'no-store' }); builtInCache = cleanConfig(r.ok ? await r.json() : {}); }
    } catch { builtInCache = cleanConfig({}); }
    return builtInCache;
  }
  /** What is used: each service's own registration if you set one, else the built-in one. */
  async function getConfig() {
    const [mine, built] = await Promise.all([userConfig(), builtInConfig()]);
    const out = { ...mine, endpoints: mine.endpoints };
    if (!mine.google.clientId) out.google = built.google;
    if (!mine.microsoft.clientId) out.microsoft = built.microsoft;
    if (!mine.dropbox.clientId) out.dropbox = built.dropbox;
    return out;
  }
  async function setConfig(c) {
    if (desktop) return window.atlas.cloudConfigSet(c);
    try { localStorage.setItem('atlas-cloud-config', JSON.stringify(cleanConfig(c))); } catch { /* ignore */ }
  }
  function endpoint(cfg, p, key) {
    const v = (cfg.endpoints[p] && cfg.endpoints[p][key]) || DEFAULTS[p][key];
    return v.replace('{tenant}', encodeURIComponent(cfg.microsoft.tenant || 'common')).replace(/\/+$/, '');
  }

  let store = null;
  async function loadStore() {
    if (store) return store;
    let s = null;
    if (desktop) s = await window.atlas.cloudAccountsGet();
    else { try { s = JSON.parse(localStorage.getItem('atlas-cloud-accounts') || 'null'); } catch { /* ignore */ } }
    store = s && Array.isArray(s.accounts) ? s : { accounts: [] };
    return store;
  }
  async function saveStore() {
    if (desktop) return window.atlas.cloudAccountsSet(store);
    try { localStorage.setItem('atlas-cloud-accounts', JSON.stringify(store)); } catch { /* ignore */ }
  }
  async function account(id) {
    const a = (await loadStore()).accounts.find(x => x.id === id);
    if (!a) throw new Error('That cloud account is no longer connected');
    return a;
  }

  // ---------- HTTP ----------
  async function request(url, { method = 'GET', headers = {}, body } = {}) {
    let r;
    if (desktop) {
      r = await window.atlas.cloudFetch({ url, method, headers, body });
      if (!r.status) throw new Error(r.error || 'Network error');
    } else {
      const resp = await fetch(url, { method, headers, body });
      r = { status: resp.status, headers: Object.fromEntries(resp.headers.entries()), body: new Uint8Array(await resp.arrayBuffer()) };
    }
    r.text = () => new TextDecoder().decode(r.body);
    r.json = () => { try { return JSON.parse(r.text() || 'null'); } catch { return null; } };
    return r;
  }
  function errorText(r) {
    const j = r.json();
    const m = j && (j.error?.message || j.error_description || j.error_summary || (typeof j.error === 'string' && j.error) || j.ocs?.meta?.message);
    return m || `The server answered ${r.status}`;
  }
  async function tokenRequest(url, form) {
    const r = await request(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' }, body: new URLSearchParams(form).toString() });
    const j = r.json();
    if (r.status >= 400 || !j || !j.access_token) throw new Error('Sign-in failed: ' + errorText(r));
    return { access: j.access_token, refresh: j.refresh_token || form.refresh_token || null, expires: Date.now() + (+j.expires_in || 3600) * 1000 };
  }
  async function refresh(a) {
    if (!a.refresh) throw new Error(`Sign in to ${PROVIDERS[a.provider].label} again`);
    const cfg = await getConfig();
    const form = { grant_type: 'refresh_token', refresh_token: a.refresh, client_id: cfg[a.provider].clientId };
    if (a.provider === 'google' && cfg.google.clientSecret) form.client_secret = cfg.google.clientSecret;
    Object.assign(a, await tokenRequest(endpoint(cfg, a.provider, 'token'), form));
    await saveStore();
  }
  /** A request on an account's behalf: adds its credentials, refreshes an expired token once. */
  async function call(a, url, opts = {}) {
    const auth = () => (a.provider === 'webdav'
      ? 'Basic ' + btoa(String.fromCharCode(...enc.encode(`${a.username}:${a.password || ''}`)))
      : 'Bearer ' + a.access);
    if (a.provider !== 'webdav' && a.expires && a.expires < Date.now() + 60000) await refresh(a);
    let r = await request(url, { ...opts, headers: { ...opts.headers, Authorization: auth() } });
    if (r.status === 401 && a.provider !== 'webdav' && a.refresh) {
      await refresh(a);
      r = await request(url, { ...opts, headers: { ...opts.headers, Authorization: auth() } });
    }
    if (r.status >= 400 && !opts.allow?.includes(r.status)) throw new Error(r.status === 401 ? 'The account needs to sign in again' : errorText(r));
    return r;
  }
  const jsonBody = v => ({ headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(v) });

  // ---------- Sign-in ----------
  function popup(url) {
    return new Promise(resolve => {
      try { localStorage.removeItem('atlas-oauth-result'); } catch { /* ignore */ }
      const w = window.open(url, 'atlas-oauth', 'width=520,height=680');
      if (!w) return resolve({ error: 'The sign-in window was blocked. Allow pop-ups for this page and try again.' });
      let done = false;
      const finish = r => {
        if (done) return;
        done = true;
        clearInterval(timer);
        window.removeEventListener('message', onMsg);
        window.removeEventListener('storage', onStore);
        try { localStorage.removeItem('atlas-oauth-result'); } catch { /* ignore */ }
        resolve(r);
      };
      const onMsg = e => { if (e.origin === location.origin && e.data && e.data.atlasOauth) finish(e.data.atlasOauth); };
      // Some sign-in pages cut the link to the opener; the callback page also leaves the result here.
      const onStore = e => { if (e.key === 'atlas-oauth-result' && e.newValue) { try { finish(JSON.parse(e.newValue)); } catch { /* ignore */ } } };
      window.addEventListener('message', onMsg);
      window.addEventListener('storage', onStore);
      const timer = setInterval(() => { if (w.closed) setTimeout(() => finish({ error: 'access_denied' }), 800); }, 500);
    });
  }
  async function signIn(provider) {
    const cfg = await getConfig();
    const clientId = cfg[provider].clientId;
    if (!clientId) throw new Error(`Add your ${PROVIDERS[provider].label} app's client ID in Cloud settings first`);
    const verifier = randomId(48);
    const challenge = b64url(new Uint8Array(await crypto.subtle.digest('SHA-256', enc.encode(verifier))));
    const state = randomId(16);
    let begin = null, redirectUri;
    if (desktop) { begin = await window.atlas.cloudOauthBegin(); redirectUri = begin.redirectUri; }
    else redirectUri = new URL('oauth.html', location.href).href.replace(/[?#].*$/, '');
    const q = new URLSearchParams({ client_id: clientId, redirect_uri: redirectUri, response_type: 'code', state, code_challenge: challenge, code_challenge_method: 'S256' });
    if (provider === 'google') { q.set('scope', 'https://www.googleapis.com/auth/drive openid email'); q.set('access_type', 'offline'); q.set('prompt', 'consent'); }
    if (provider === 'microsoft') q.set('scope', 'offline_access Files.ReadWrite.All User.Read');
    if (provider === 'dropbox') q.set('token_access_type', 'offline');
    const authUrl = endpoint(cfg, provider, 'auth') + '?' + q;
    const res = desktop ? await window.atlas.cloudOauthWait({ id: begin.id, authUrl }) : await popup(authUrl);
    if (!res || res.error) throw new Error(!res || res.error === 'access_denied' ? 'Sign-in was cancelled' : res.error === 'timeout' ? 'Sign-in timed out' : 'Sign-in failed: ' + res.error);
    if (res.state !== state || !res.code) throw new Error('Sign-in failed: the response did not match the request');
    const form = { grant_type: 'authorization_code', code: res.code, redirect_uri: redirectUri, client_id: clientId, code_verifier: verifier };
    if (provider === 'google' && cfg.google.clientSecret) form.client_secret = cfg.google.clientSecret;
    return tokenRequest(endpoint(cfg, provider, 'token'), form);
  }

  // ---------- Services ----------
  // Each: who(a) -> label, root(a), list(a, folder) -> items, download(a, item) -> bytes,
  // create(a, folder, name, bytes) -> item, update(a, id, name, bytes) -> item, share(a, id, {role, email}) -> link|null,
  // stat(a, id) -> the file's revision (changes whenever anyone saves it), for offline saves.
  // Items: { id, name, folder, size, modified, native }.
  const SERVICES = {
    google: {
      async who(a, cfg) { return (await call(a, endpoint(cfg, 'google', 'api') + '/about?fields=user')).json().user.emailAddress; },
      root: () => 'root',
      async list(a, folder, cfg) {
        const api = endpoint(cfg, 'google', 'api');
        const q = `'${String(folder).replace(/['\\]/g, '\\$&')}' in parents and trashed = false`;
        let items = [], token = '';
        for (let page = 0; page < 10; page++) {
          const j = (await call(a, `${api}/files?` + new URLSearchParams({ q, pageSize: '200', orderBy: 'folder,name', fields: 'nextPageToken,files(id,name,mimeType,size,modifiedTime)', ...(token ? { pageToken: token } : {}) }))).json();
          items = items.concat(j.files.map(f => {
            const nat = GOOGLE_NATIVE[f.mimeType];
            return { id: f.id, name: nat ? f.name + '.' + nat.ext : f.name, folder: f.mimeType === GFOLDER, size: +f.size || 0, modified: f.modifiedTime, native: nat ? f.mimeType : null };
          }));
          if (!(token = j.nextPageToken)) break;
        }
        return items;
      },
      async download(a, item, cfg) {
        const api = endpoint(cfg, 'google', 'api');
        if (item.native) return (await call(a, `${api}/files/${encodeURIComponent(item.id)}/export?mimeType=${encodeURIComponent(GOOGLE_NATIVE[item.native].mime)}`)).body;
        return (await call(a, `${api}/files/${encodeURIComponent(item.id)}?alt=media`)).body;
      },
      async create(a, folder, name, bytes, cfg) {
        const b = 'atlas' + randomId(12);
        const meta = JSON.stringify({ name, parents: [folder] });
        const body = concat([`--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n--${b}\r\nContent-Type: ${mimeOf(name)}\r\n\r\n`, bytes, `\r\n--${b}--`]);
        const j = (await call(a, endpoint(cfg, 'google', 'upload') + '/files?uploadType=multipart&fields=id,name', { method: 'POST', headers: { 'Content-Type': `multipart/related; boundary=${b}` }, body })).json();
        return { id: j.id, name: j.name };
      },
      async update(a, id, name, bytes, cfg) {
        const j = (await call(a, endpoint(cfg, 'google', 'upload') + `/files/${encodeURIComponent(id)}?uploadType=media&fields=id,name`, { method: 'PATCH', headers: { 'Content-Type': mimeOf(name) }, body: bytes })).json();
        return { id: j.id, name: j.name };
      },
      async stat(a, id, cfg) { const j = (await call(a, `${endpoint(cfg, 'google', 'api')}/files/${encodeURIComponent(id)}?fields=version,modifiedTime`)).json(); return String(j.version || j.modifiedTime || ''); },
      async share(a, id, { role, email }, cfg) {
        const api = endpoint(cfg, 'google', 'api');
        const perm = email ? { type: 'user', role, emailAddress: email } : { type: 'anyone', role };
        await call(a, `${api}/files/${encodeURIComponent(id)}/permissions?sendNotificationEmail=${email ? 'true' : 'false'}`, { method: 'POST', ...jsonBody(perm) });
        return (await call(a, `${api}/files/${encodeURIComponent(id)}?fields=webViewLink`)).json().webViewLink;
      }
    },
    microsoft: {
      async who(a, cfg) { const j = (await call(a, endpoint(cfg, 'microsoft', 'api') + '/me')).json(); return j.mail || j.userPrincipalName || j.displayName; },
      root: () => 'root',
      async list(a, folder, cfg) {
        let url = `${endpoint(cfg, 'microsoft', 'api')}/me/drive/items/${encodeURIComponent(folder)}/children?$top=200&$select=id,name,size,lastModifiedDateTime,folder,file`;
        let items = [];
        for (let page = 0; page < 10 && url; page++) {
          const j = (await call(a, url)).json();
          items = items.concat(j.value.map(x => ({ id: x.id, name: x.name, folder: !!x.folder, size: x.size || 0, modified: x.lastModifiedDateTime })));
          url = j['@odata.nextLink'];
        }
        return items.sort((x, y) => (y.folder - x.folder) || x.name.localeCompare(y.name));
      },
      async download(a, item, cfg) { return (await call(a, `${endpoint(cfg, 'microsoft', 'api')}/me/drive/items/${encodeURIComponent(item.id)}/content`)).body; },
      async create(a, folder, name, bytes, cfg) {
        const j = (await call(a, `${endpoint(cfg, 'microsoft', 'api')}/me/drive/items/${encodeURIComponent(folder)}:/${encodeURIComponent(name)}:/content?@microsoft.graph.conflictBehavior=rename`,
          { method: 'PUT', headers: { 'Content-Type': mimeOf(name) }, body: bytes })).json();
        return { id: j.id, name: j.name };
      },
      async update(a, id, name, bytes, cfg) {
        const j = (await call(a, `${endpoint(cfg, 'microsoft', 'api')}/me/drive/items/${encodeURIComponent(id)}/content`, { method: 'PUT', headers: { 'Content-Type': mimeOf(name) }, body: bytes })).json();
        return { id: j.id, name: j.name };
      },
      async stat(a, id, cfg) { const j = (await call(a, `${endpoint(cfg, 'microsoft', 'api')}/me/drive/items/${encodeURIComponent(id)}?$select=eTag,lastModifiedDateTime`)).json(); return String(j.eTag || j.lastModifiedDateTime || ''); },
      async share(a, id, { role, email }, cfg) {
        const item = `${endpoint(cfg, 'microsoft', 'api')}/me/drive/items/${encodeURIComponent(id)}`;
        if (email) {
          await call(a, item + '/invite', { method: 'POST', ...jsonBody({ recipients: [{ email }], requireSignIn: true, sendInvitation: true, roles: [role === 'writer' ? 'write' : 'read'] }) });
          return (await call(a, item + '?$select=webUrl')).json().webUrl;
        }
        return (await call(a, item + '/createLink', { method: 'POST', ...jsonBody({ type: role === 'writer' ? 'edit' : 'view', scope: 'anonymous' }) })).json().link.webUrl;
      }
    },
    dropbox: {
      async who(a, cfg) { return (await call(a, endpoint(cfg, 'dropbox', 'api') + '/users/get_current_account', { method: 'POST' })).json().email; },
      root: () => '',
      async list(a, folder, cfg) {
        const api = endpoint(cfg, 'dropbox', 'api');
        let j = (await call(a, api + '/files/list_folder', { method: 'POST', ...jsonBody({ path: folder, limit: 2000 }) })).json();
        let entries = j.entries;
        for (let page = 0; page < 10 && j.has_more; page++) {
          j = (await call(a, api + '/files/list_folder/continue', { method: 'POST', ...jsonBody({ cursor: j.cursor }) })).json();
          entries = entries.concat(j.entries);
        }
        return entries.filter(e => e['.tag'] !== 'deleted')
          .map(e => ({ id: e.path_display, name: e.name, folder: e['.tag'] === 'folder', size: e.size || 0, modified: e.server_modified }))
          .sort((x, y) => (y.folder - x.folder) || x.name.localeCompare(y.name));
      },
      async download(a, item, cfg) {
        return (await call(a, endpoint(cfg, 'dropbox', 'content') + '/files/download', { method: 'POST', headers: { 'Dropbox-API-Arg': headerJson({ path: item.id }) } })).body;
      },
      async create(a, folder, name, bytes, cfg) {
        const j = (await call(a, endpoint(cfg, 'dropbox', 'content') + '/files/upload', { method: 'POST', body: bytes,
          headers: { 'Content-Type': 'application/octet-stream', 'Dropbox-API-Arg': headerJson({ path: `${folder}/${name}`, mode: 'add', autorename: true, mute: true }) } })).json();
        return { id: j.path_display, name: j.name };
      },
      async update(a, id, name, bytes, cfg) {
        const j = (await call(a, endpoint(cfg, 'dropbox', 'content') + '/files/upload', { method: 'POST', body: bytes,
          headers: { 'Content-Type': 'application/octet-stream', 'Dropbox-API-Arg': headerJson({ path: id, mode: 'overwrite', mute: true }) } })).json();
        return { id: j.path_display, name: j.name };
      },
      async stat(a, id, cfg) { const j = (await call(a, endpoint(cfg, 'dropbox', 'api') + '/files/get_metadata', { method: 'POST', ...jsonBody({ path: id }) })).json(); return String(j.rev || j.server_modified || ''); },
      async share(a, id, { role, email }, cfg) {
        const api = endpoint(cfg, 'dropbox', 'api');
        if (email) {
          await call(a, api + '/sharing/add_file_member', { method: 'POST', ...jsonBody({ file: id, members: [{ '.tag': 'email', email }], access_level: { '.tag': role === 'writer' ? 'editor' : 'viewer' }, quiet: false }) });
        }
        const r = await call(a, api + '/sharing/create_shared_link_with_settings', { method: 'POST', allow: [409], ...jsonBody({ path: id }) });
        const j = r.json();
        if (r.status !== 409) return j.url;
        if (j?.error?.shared_link_already_exists?.metadata?.url) return j.error.shared_link_already_exists.metadata.url;
        const l = (await call(a, api + '/sharing/list_shared_links', { method: 'POST', ...jsonBody({ path: id, direct_only: true }) })).json();
        if (l.links && l.links[0]) return l.links[0].url;
        throw new Error(errorText(r));
      }
    },
    webdav: {
      async who(a) { await this.list(a, this.root(a)); return `${a.username}@${new URL(a.url).host}`; },
      root: a => new URL(a.url).pathname.replace(/\/?$/, '/'),
      async list(a, folder) {
        const base = new URL(a.url);
        const body = '<?xml version="1.0" encoding="utf-8"?><d:propfind xmlns:d="DAV:"><d:prop><d:displayname/><d:getcontentlength/><d:getlastmodified/><d:resourcetype/></d:prop></d:propfind>';
        const r = await call(a, base.origin + folder, { method: 'PROPFIND', headers: { Depth: '1', 'Content-Type': 'application/xml; charset=utf-8' }, body });
        const xml = new DOMParser().parseFromString(r.text(), 'application/xml');
        const D = 'DAV:';
        const self = decodeURIComponent(folder).replace(/\/$/, '');
        return [...xml.getElementsByTagNameNS(D, 'response')].map(res => {
          const href = new URL(res.getElementsByTagNameNS(D, 'href')[0]?.textContent || '', base).pathname;
          const prop = res.getElementsByTagNameNS(D, 'prop')[0];
          const get = n => prop?.getElementsByTagNameNS(D, n)[0];
          const isFolder = !!get('resourcetype')?.getElementsByTagNameNS(D, 'collection')[0];
          const name = decodeURIComponent(href.replace(/\/$/, '').split('/').pop());
          return { id: isFolder ? href.replace(/\/?$/, '/') : href, name: get('displayname')?.textContent || name, folder: isFolder, size: +(get('getcontentlength')?.textContent || 0), modified: get('getlastmodified')?.textContent || '' };
        }).filter(x => decodeURIComponent(x.id).replace(/\/$/, '') !== self)
          .sort((x, y) => (y.folder - x.folder) || x.name.localeCompare(y.name));
      },
      async download(a, item) { return (await call(a, new URL(a.url).origin + item.id)).body; },
      async create(a, folder, name, bytes) {
        const origin = new URL(a.url).origin;
        const base = Platform.baseName(name), ext = Platform.extOf(name);
        for (let n = 1; n < 50; n++) {
          const nm = n === 1 ? name : `${base} (${n})${ext ? '.' + ext : ''}`;
          const id = folder.replace(/\/?$/, '/') + encodeURIComponent(nm);
          // If-None-Match: * — never overwrite a file that is already there.
          const r = await call(a, origin + id, { method: 'PUT', allow: [412], headers: { 'Content-Type': mimeOf(name), 'If-None-Match': '*' }, body: bytes });
          if (r.status !== 412) return { id, name: nm };
        }
        throw new Error('Could not find a free file name');
      },
      async update(a, id, name, bytes) {
        await call(a, new URL(a.url).origin + id, { method: 'PUT', headers: { 'Content-Type': mimeOf(name) }, body: bytes });
        return { id, name };
      },
      async stat(a, id) {
        const body = '<?xml version="1.0" encoding="utf-8"?><d:propfind xmlns:d="DAV:"><d:prop><d:getetag/><d:getlastmodified/><d:getcontentlength/></d:prop></d:propfind>';
        const r = await call(a, new URL(a.url).origin + id, { method: 'PROPFIND', headers: { Depth: '0', 'Content-Type': 'application/xml; charset=utf-8' }, body });
        const xml = new DOMParser().parseFromString(r.text(), 'application/xml');
        const get = n => (xml.getElementsByTagNameNS('DAV:', n)[0] || {}).textContent || '';
        return get('getetag') || (get('getlastmodified') + '/' + get('getcontentlength'));
      },
      async share(a, id, { role, email }) {
        // Nextcloud and ownCloud share through their OCS API next to the WebDAV endpoint.
        const m = /^(.*)\/remote\.php\/(?:dav\/files\/[^/]+|webdav)(\/.*)$/.exec(id);
        if (!m) throw new Error('This WebDAV server does not support share links');
        const base = new URL(a.url).origin + m[1];
        const form = { path: decodeURIComponent(m[2]), shareType: email ? '4' : '3', permissions: role === 'writer' ? '3' : '1' };
        if (email) form.shareWith = email;
        const r = await call(a, base + '/ocs/v2.php/apps/files_sharing/api/v1/shares?format=json', { method: 'POST',
          headers: { 'OCS-APIRequest': 'true', 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' }, body: new URLSearchParams(form).toString() });
        return r.json()?.ocs?.data?.url || null;
      }
    }
  };

  // ---------- Cloud paths ----------
  const makePath = (acct, id, folder, name) => `cloud://${acct}/${encodeURIComponent(id || '')}/${encodeURIComponent(folder || '')}/${encodeURIComponent(name)}`;
  function parsePath(p) {
    const m = /^cloud:\/\/([\w-]+)\/([^/]*)\/([^/]*)\/([^/]+)$/.exec(p || '');
    if (!m) throw new Error('Not a cloud document');
    return { acct: m[1], id: decodeURIComponent(m[2]), folder: decodeURIComponent(m[3]), name: decodeURIComponent(m[4]) };
  }

  // ---------- Revisions ----------
  // The revision each cloud file had when it was last opened or saved here: a save waiting to go
  // up (offline.js) is a conflict if the file has another revision by then.
  const REVS = 'atlas-cloud-revs';
  const revs = () => { try { return JSON.parse(localStorage.getItem(REVS) || '{}') || {}; } catch { return {}; } };
  function setRev(path, rev) {
    const r = revs();
    if (rev) r[path] = rev; else delete r[path];
    const keys = Object.keys(r);
    if (keys.length > 500) keys.slice(0, keys.length - 500).forEach(k => delete r[k]);
    try { localStorage.setItem(REVS, JSON.stringify(r)); } catch { /* ignore */ }
  }
  /** The file's revision in the cloud now ('' when the service can't say). */
  async function stat(path) {
    const p = parsePath(path);
    if (!p.id) return '';
    const a = await account(p.acct);
    return SERVICES[a.provider].stat ? SERVICES[a.provider].stat(a, p.id, await getConfig()) : '';
  }
  const noteRev = path => stat(path).then(r => setRev(path, r)).catch(() => {});

  // ---------- Public API ----------
  const Cloud = {
    PROVIDERS, appFor, makePath, parsePath, getConfig, setConfig,
    async accounts() { return (await loadStore()).accounts.map(a => ({ id: a.id, provider: a.provider, label: a.label })); },
    /** Signs in to a service and remembers the account. */
    async connect(provider, opts = {}) {
      const st = await loadStore();
      const cfg = await getConfig();
      let a;
      if (provider === 'webdav') {
        const url = new URL(opts.url);
        if (!/^https?:$/.test(url.protocol)) throw new Error('Enter an http(s) address');
        if (url.protocol === 'http:' && !/^(localhost|127\.0\.0\.1|\[::1\])$/.test(url.hostname)) throw new Error('Use an https:// address so your password is not sent in the clear');
        a = { id: 'w' + randomId(9).replace(/[^\w]/g, ''), provider, url: url.href.replace(/\/?$/, '/'), username: opts.username || '', password: opts.password || '' };
      } else {
        a = { id: provider[0] + randomId(9).replace(/[^\w]/g, ''), provider, ...(await signIn(provider)) };
      }
      // Remember it first: the desktop app only connects to WebDAV servers it knows about.
      st.accounts.push(a);
      await saveStore();
      try {
        a.label = await SERVICES[provider].who(a, cfg) || PROVIDERS[provider].label;
      } catch (err) {
        st.accounts = st.accounts.filter(x => x !== a);
        await saveStore();
        throw err;
      }
      const dup = st.accounts.find(x => x !== a && x.provider === a.provider && x.label === a.label);
      if (dup) { Object.assign(dup, a, { id: dup.id }); st.accounts = st.accounts.filter(x => x !== a); a = dup; }
      await saveStore();
      return { id: a.id, provider, label: a.label };
    },
    async disconnect(id) {
      const st = await loadStore();
      st.accounts = st.accounts.filter(a => a.id !== id);
      await saveStore();
    },
    async root(id) { const a = await account(id); return SERVICES[a.provider].root(a); },
    async list(id, folder) { const a = await account(id); return SERVICES[a.provider].list(a, folder ?? SERVICES[a.provider].root(a), await getConfig()); },
    async download(id, item) { const a = await account(id); return toBytes(await SERVICES[a.provider].download(a, item, await getConfig())); },
    /** Uploads a new file into a folder; returns { name, path }. */
    async upload(id, folder, name, data) {
      const a = await account(id);
      const it = await SERVICES[a.provider].create(a, folder, name, toBytes(data), await getConfig());
      return { name: it.name || name, path: makePath(id, it.id, folder, it.name || name) };
    },
    /** Platform.save for a cloud path: writes over the file (or creates it, for an exported Google file). */
    async saveTo(path, name, data) {
      const p = parsePath(path);
      if (!p.id) return Cloud.upload(p.acct, p.folder, name || p.name, data);
      const a = await account(p.acct);
      const it = await SERVICES[a.provider].update(a, p.id, name || p.name, toBytes(data), await getConfig());
      const out = { name: name || p.name, path: makePath(p.acct, it.id || p.id, p.folder, name || p.name) };
      await noteRev(out.path);
      return out;
    },
    stat, revisionSeen: path => revs()[path] || '', noteRev,
    async share(path, opts) {
      const p = parsePath(path);
      if (!p.id) throw new Error('Save the document first');
      const a = await account(p.acct);
      return SERVICES[a.provider].share(a, p.id, { role: opts.role === 'writer' ? 'writer' : 'reader', email: opts.email || '' }, await getConfig());
    },
    /** Downloads a picked file and returns what the apps' loadFile takes: { name, path, data }. */
    async fetchFile(id, item, folder) {
      const data = await Cloud.download(id, item);
      // An exported Google Docs/Sheets/Slides file is saved as a new Office file next to it.
      const f = { name: item.name, data, path: makePath(id, item.native ? '' : item.id, folder, item.name) };
      if (!item.native) await noteRev(f.path);
      if (window.Offline) Offline.keep(f.path, f.name, data);
      return f;
    }
  };

  // ---------- Dialogs ----------
  function el(tag, cls, text) { const e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
  function fmtSize(n) { return !n ? '' : n < 1024 ? n + ' B' : n < 1048576 ? (n / 1024).toFixed(0) + ' KB' : (n / 1048576).toFixed(1) + ' MB'; }

  /**
   * The cloud file browser. mode 'open' resolves to { account, item, folder } (a file);
   * mode 'save' resolves to { account, folder, name }. exts limits the files shown and saved.
   */
  Cloud.browse = async function ({ mode = 'open', name = '', exts = null, title } = {}) {
    let accts = await Cloud.accounts();
    if (!accts.length) {
      const added = await Cloud.addAccountDialog();
      if (!added) return null;
      accts = await Cloud.accounts();
    }
    return new Promise(resolve => {
      const back = el('div', 'modal-backdrop');
      const modal = el('div', 'modal cloud-browser');
      modal.setAttribute('role', 'dialog');
      modal.setAttribute('aria-label', title || (mode === 'save' ? 'Save to cloud' : 'Open from cloud'));
      modal.appendChild(el('h2', '', title || (mode === 'save' ? 'Save to cloud' : 'Open from cloud')));
      const top = el('div', 'cloud-top');
      const sel = el('select');
      sel.setAttribute('aria-label', 'Account');
      accts.forEach(a => { const o = el('option', '', `${PROVIDERS[a.provider].label} — ${a.label}`); o.value = a.id; sel.appendChild(o); });
      const addBtn = el('button', '', 'Add a place…');
      top.append(sel, addBtn);
      const crumbs = el('div', 'cloud-crumbs');
      const list = el('div', 'cloud-list');
      list.setAttribute('role', 'listbox');
      list.tabIndex = 0;
      const nameRow = el('label', 'cloud-name', 'File name');
      const nameInput = el('input');
      nameInput.value = name;
      nameRow.appendChild(nameInput);
      const actions = el('div', 'actions');
      const cancel = el('button', '', 'Cancel'), ok = el('button', 'primary', mode === 'save' ? 'Save' : 'Open');
      actions.append(cancel, ok);
      modal.append(top, crumbs, list);
      if (mode === 'save') modal.appendChild(nameRow);
      modal.appendChild(actions);
      back.appendChild(modal);
      document.body.appendChild(back);

      let path = [], picked = null, items = [];
      const done = v => { back.remove(); resolve(v); };
      const folder = () => path.length ? path[path.length - 1].id : null;
      async function show() {
        picked = null;
        list.innerHTML = '';
        list.appendChild(el('div', 'empty', 'Loading…'));
        crumbs.innerHTML = '';
        const rootBtn = el('button', '', PROVIDERS[accts.find(a => a.id === sel.value).provider].label);
        rootBtn.onclick = () => { path = []; show(); };
        crumbs.appendChild(rootBtn);
        path.forEach((p, i) => { crumbs.appendChild(el('span', '', '›')); const b = el('button', '', p.name); b.onclick = () => { path = path.slice(0, i + 1); show(); }; crumbs.appendChild(b); });
        try {
          const f = folder() ?? await Cloud.root(sel.value);
          items = (await Cloud.list(sel.value, f)).filter(it => it.folder || (exts ? exts.includes(Platform.extOf(it.name)) : !!appFor(it.name)));
          list.innerHTML = '';
          if (!items.length) list.appendChild(el('div', 'empty', mode === 'save' ? 'This folder is empty' : 'No documents here'));
          items.forEach(it => {
            const row = el('div', 'cloud-row' + (it.folder ? ' folder' : ''));
            row.setAttribute('role', 'option');
            row.append(el('span', 'ic', it.folder ? '📁' : ({ write: '📄', sheets: '📊', slides: '📽' }[appFor(it.name)] || '📄')), el('span', 'nm', it.name),
              el('span', 'sz', fmtSize(it.size)), el('span', 'dt', it.modified ? new Date(it.modified).toLocaleDateString() : ''));
            row.onclick = () => {
              list.querySelectorAll('.sel').forEach(r => r.classList.remove('sel'));
              row.classList.add('sel');
              picked = it;
              if (!it.folder && mode === 'save') nameInput.value = it.name;
            };
            row.ondblclick = () => { if (it.folder) { path.push({ id: it.id, name: it.name }); show(); } else go(); };
            list.appendChild(row);
          });
        } catch (err) {
          list.innerHTML = '';
          list.appendChild(el('div', 'empty', 'Could not list this folder: ' + err.message));
        }
      }
      async function go() {
        if (picked && picked.folder) { path.push({ id: picked.id, name: picked.name }); return show(); }
        const f = folder() ?? await Cloud.root(sel.value);
        if (mode === 'open') { if (picked) done({ account: sel.value, item: picked, folder: f }); return; }
        let nm = nameInput.value.trim().replace(/[\\/:*?"<>|]/g, '_');
        if (!nm) return nameInput.focus();
        if (exts && !exts.includes(Platform.extOf(nm))) nm += '.' + (exts.includes(Platform.extOf(name)) ? Platform.extOf(name) : exts[0]);
        const existing = items.find(it => !it.folder && it.name.toLowerCase() === nm.toLowerCase());
        if (existing && !(await UI.confirm('Replace file', `"${nm}" already exists in this folder. Replace it?`))) return;
        done({ account: sel.value, folder: f, name: nm, existing: existing || null });
      }
      sel.onchange = () => { path = []; show(); };
      addBtn.onclick = async () => {
        const a = await Cloud.addAccountDialog();
        if (!a) return;
        accts = await Cloud.accounts();
        if (![...sel.options].some(o => o.value === a.id)) { const o = el('option', '', `${PROVIDERS[a.provider].label} — ${a.label}`); o.value = a.id; sel.appendChild(o); }
        sel.value = a.id;
        path = [];
        show();
      };
      ok.onclick = go;
      cancel.onclick = () => done(null);
      back.addEventListener('mousedown', e => { if (e.target === back) done(null); });
      modal.addEventListener('keydown', e => {
        if (e.key === 'Escape') { e.preventDefault(); done(null); }
        if (e.key === 'Enter' && e.target.tagName !== 'BUTTON' && e.target.tagName !== 'SELECT') { e.preventDefault(); go(); }
        if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && e.target === list) {
          e.preventDefault();
          const rows = [...list.querySelectorAll('.cloud-row')];
          const i = rows.findIndex(r => r.classList.contains('sel'));
          const n = rows[Math.max(0, Math.min(rows.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1)))];
          if (n) { n.click(); n.scrollIntoView({ block: 'nearest' }); }
        }
      });
      (mode === 'save' ? nameInput : list).focus();
      show();
    });
  };

  /** Asks which service to add and signs in. Resolves to the new account or null. */
  Cloud.addAccountDialog = async function () {
    const keys = Object.keys(PROVIDERS);
    const i = await UI.choose('Add a place', 'Where are your documents?', [...keys.map(k => PROVIDERS[k].label), 'Cancel']);
    if (i < 0 || i >= keys.length) return null;
    const provider = keys[i];
    try {
      if (provider === 'webdav') {
        const v = await UI.form('Connect to a WebDAV server', [
          { name: 'url', label: 'Server address (for Nextcloud: https://host/remote.php/dav/files/USER/)', value: 'https://' },
          { name: 'username', label: 'User name' },
          { name: 'password', label: 'Password or app password', type: 'password' }
        ], 'Connect');
        if (!v) return null;
        UI.toast('Connecting…');
        const a = await Cloud.connect('webdav', v);
        UI.toast('Connected to ' + a.label);
        return a;
      }
      const cfg = await getConfig();
      if (!cfg[provider].clientId && !(await Cloud.settingsDialog(provider))) return null;
      UI.toast('Finish signing in in your browser…', 5000);
      const a = await Cloud.connect(provider);
      UI.toast('Connected to ' + a.label);
      return a;
    } catch (err) {
      UI.toast(err.message, 5000);
      return null;
    }
  };

  /** App registrations for the OAuth services (each person or organisation brings their own). */
  Cloud.settingsDialog = async function (only) {
    const cfg = await userConfig(), built = await builtInConfig();
    const fields = [];
    const redirect = desktop ? 'http://127.0.0.1 (any port)' : new URL('oauth.html', location.href).href.replace(/[?#].*$/, '');
    const hint = p => (built[p].clientId ? 'Built in — leave empty to use it' : '');
    if (!only || only === 'google') fields.push({ name: 'gId', label: 'Google Drive — OAuth client ID', value: cfg.google.clientId, placeholder: hint('google') },
      { name: 'gSecret', label: 'Google Drive — client secret (desktop clients only)', value: cfg.google.clientSecret });
    if (!only || only === 'microsoft') fields.push({ name: 'mId', label: 'OneDrive — application (client) ID', value: cfg.microsoft.clientId, placeholder: hint('microsoft') },
      { name: 'mTenant', label: 'OneDrive — tenant (common, consumers, organizations or a tenant ID)', value: cfg.microsoft.tenant });
    if (!only || only === 'dropbox') fields.push({ name: 'dId', label: 'Dropbox — app key', value: cfg.dropbox.clientId, placeholder: hint('dropbox') });
    fields.push({ name: 'note', label: 'If you use your own app, register this redirect address with the service', value: redirect });
    const v = await UI.form(only ? `${PROVIDERS[only].label} app` : 'Cloud settings', fields, 'Save');
    if (!v) return false;
    const next = { ...cfg, google: { ...cfg.google }, microsoft: { ...cfg.microsoft }, dropbox: { ...cfg.dropbox } };
    if ('gId' in v) { next.google.clientId = v.gId; next.google.clientSecret = v.gSecret; }
    if ('mId' in v) { next.microsoft.clientId = v.mId; next.microsoft.tenant = v.mTenant || 'common'; }
    if ('dId' in v) next.dropbox.clientId = v.dId;
    await setConfig(next);
    return only ? !!(await getConfig())[only].clientId : true;
  };

  /**
   * App helpers. Each app passes: state (its docState), serialize(ext), exts (formats it saves to
   * the cloud) and load(file) to open a file in place.
   */
  Cloud.saveAsDialog = async function ({ state, serialize, exts, before }) {
    if (before) before();
    const name = exts.includes(Platform.extOf(state.name)) ? state.name : Platform.baseName(state.name) + '.' + exts[0];
    const t = await Cloud.browse({ mode: 'save', name, exts });
    if (!t) return false;
    try {
      UI.toast('Saving…', 10000);
      const data = await serialize(Platform.extOf(t.name));
      const r = t.existing ? await Cloud.saveTo(makePath(t.account, t.existing.id, t.folder, t.existing.name), t.name, data) : await Cloud.upload(t.account, t.folder, t.name, data);
      state.set({ name: r.name, path: r.path, dirty: false });
      UI.toast('Saved to the cloud');
      return true;
    } catch (err) {
      UI.toast('Save failed: ' + err.message, 5000);
      return false;
    }
  };
  Cloud.openDialog = async function ({ app, exts, isEmpty, load }) {
    const t = await Cloud.browse({ mode: 'open', exts });
    if (!t) return;
    try {
      UI.toast('Opening…', 10000);
      const f = await Cloud.fetchFile(t.account, t.item, t.folder);
      const target = appFor(f.name);
      if (!target) return UI.toast('Isadek Office cannot open this kind of file');
      if (target === app && isEmpty && await isEmpty()) { await load(f); UI.toast(t.item.native ? 'Opened a copy; Save creates an Office file next to it' : 'Opened', 3000); }
      else await Platform.openData(target, f.name, f.data, !Platform.desktop && !app, f.path);
    } catch (err) {
      UI.toast('Could not open: ' + err.message, 5000);
    }
  };
  Cloud.shareDialog = async function (opts) {
    const { state, save } = opts;
    // In the phone app, a document on the device goes out through the share sheet.
    if (Platform.native && !Platform.isCloudPath(state.path)) {
      const i = await UI.choose('Share', 'Send a copy with another app, or save it to a cloud account to share a link.', ['Send a copy…', 'Save to cloud…', 'Cancel']);
      if (i === 0) return Platform.shareCurrent(state, save);
      if (i !== 1) return;
      if (!(await Cloud.saveAsDialog(opts))) return;
    }
    if (!Platform.isCloudPath(state.path) || !parsePath(state.path).id) {
      const i = await UI.choose('Share', 'To share this document, save it to a cloud account first.', ['Save to cloud…', 'Cancel']);
      if (i !== 0 || !(await Cloud.saveAsDialog(opts))) return;
    } else if (state.dirty && save) await save();
    const p = parsePath(state.path);
    const a = await account(p.acct).catch(() => null);
    if (!a) return UI.toast('That cloud account is no longer connected', 4000);
    const v = await UI.form(`Share "${state.name}"`, [
      { name: 'email', label: 'Invite people (email address — leave empty for a link anyone can use)', type: 'email' },
      { name: 'role', label: 'They can', type: 'select', options: [{ value: 'reader', label: 'View' }, { value: 'writer', label: 'Edit' }] }
    ], 'Share');
    if (!v) return;
    try {
      UI.toast('Sharing…', 10000);
      const link = await Cloud.share(state.path, v);
      if (!link) return UI.toast(v.email ? 'Invitation sent to ' + v.email : 'Shared', 4000);
      const msg = (v.email ? `Invitation sent to ${v.email}.\n\n` : '') + link;
      const i = await UI.choose('Share link', msg, ['Copy link', 'Close'], '🔗');
      if (i === 0) {
        try { await navigator.clipboard.writeText(link); UI.toast('Link copied'); } catch { UI.toast('Copy the link from the dialog'); }
      }
    } catch (err) {
      UI.toast('Could not share: ' + err.message, 5000);
    }
  };

  window.Cloud = Cloud;
})();
