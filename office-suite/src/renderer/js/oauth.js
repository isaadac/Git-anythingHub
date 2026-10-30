'use strict';
// OAuth redirect page for the web version: hands the authorization code to the page that opened it.
(function () {
  const q = new URLSearchParams(location.search);
  const result = { code: q.get('code'), state: q.get('state'), error: q.get('error') };
  history.replaceState(null, '', location.pathname);
  try { if (window.opener) window.opener.postMessage({ atlasOauth: result }, location.origin); } catch { /* opener gone */ }
  // Fallback for sign-in pages that cut the link to the opener.
  try { localStorage.setItem('atlas-oauth-result', JSON.stringify(result)); } catch { /* ignore */ }
  setTimeout(() => window.close(), 400);
})();
