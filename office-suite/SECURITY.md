# Security

How Isadek Office protects people from hostile documents, collaborators and networks, and how to
report a problem.

## Reporting

Please report vulnerabilities privately through GitHub's **Report a vulnerability** (Security
tab of this repository) rather than in a public issue. Include the version, the platform and a
document or steps that show the problem.

## The model

**Every document is untrusted** — files, pasted content, cloud files, templates from others, and
changes from other people in a live session.

- **Content.** HTML is sanitized on the way in (no scripts, frames, forms, event handlers,
  `javascript:` links or remote resources). Structured data (slides, sheets, notebooks, citations,
  ink) is rebuilt from known, typed fields. Every page has a Content Security Policy without
  inline scripts, so even a sanitizer gap can't run code.
- **Live editing.** Changes from other participants pass the same checks as a file before they
  are applied.
- **Active content waits for permission.** Python formulas, workbook scripts and web queries
  from a workbook someone else made — opened from a file or arriving in a live session — run
  only after the person allows them.
- **Python in cells** runs in a separate worker with no network, storage or page access (also
  removed from the prototypes Python can reach), and is stopped if it runs too long.
- **Spreadsheets are parsed** in a throwaway worker with a time limit, which contains the known
  advisories in the npm build of SheetJS `xlsx` 0.18.5 (prototype pollution, ReDoS). Upgrading
  to SheetJS 0.20.3 from `cdn.sheetjs.com` removes them.

## Desktop app (Electron)

- Windows are sandboxed, context-isolated and have no Node.js. They can't navigate, open other
  windows or attach webviews; links open in the browser (http, https and mailto only).
- Every message to the main process is checked to come from one of the app's own pages in a
  window's main frame.
- A page can only read or write files the person chose (dialogs, the system, the recent list).
- Saves are atomic: written to a temporary file, flushed and renamed, so a crash or full disk
  never leaves half a document.
- "Get data from the web" can't reach this computer, link-local addresses or cloud metadata
  services, also after redirects and DNS tricks, and is size- and time-limited.
- Permissions: clipboard, full screen and microphone; the camera only in Slides (Cameo).
- API keys and cloud tokens are encrypted with the operating system's keychain.
- Packaged builds set Electron fuses: no `ELECTRON_RUN_AS_NODE`, `NODE_OPTIONS` or `--inspect`;
  encrypted cookies; ASAR integrity checking; the app only loads from its ASAR.
- Updates come from GitHub Releases (or an https feed) and are verified by electron-updater.

## Servers

The web server and the live-editing server (`src/server/web-server.js`,
`src/main/collab-server.js`):

- Survive malformed requests and hostile messages: an error affects only that request or
  connection.
- Limit message size, rate per connection, presence size, operations per message, connections
  and idle sockets.
- The Assistant's API key stays on the server. Access codes are compared in constant time, and
  repeated wrong codes lock an address out.
- Send `X-Frame-Options`/`frame-ancestors`, `nosniff`, `Referrer-Policy`,
  `Cross-Origin-Resource-Policy` and `Permissions-Policy` with every response, and serve only the
  app's own files.

Forms (`src/main/forms-server.js`, on the web server and on the desktop app's "This computer"
server):

- A form's id (128 random bits) is its public link; its owner key (256 bits) is needed to read
  answers, change or delete the form, and only the key's SHA-256 is stored. The key is kept in the
  workbook, so anyone with the file can collect the answers.
- Every answer is checked against the form (types, options, lengths) before it is kept; answers
  per address per minute, forms per address per hour, and answers per form (count and bytes) are
  limited.
- Answers that look like formulas (`=`, `+`, `-`, `@`…) are written to cells as text.
- The desktop app's server serves the answer page and the forms API only (no other files); the
  answer page has a strict CSP (no inline script, connects only to its own server).
- A workbook from someone else asks before fetching answers from its form's server.

Run them behind HTTPS (a reverse proxy) when they are reachable from the internet. A live
session's link is its key: share it only with the people meant to join.

## Mobile

App data (sign-ins, notebooks, drafts) is left out of device backups. Mixed content is blocked.
Documents open from other apps through "Open with", not from web links.
