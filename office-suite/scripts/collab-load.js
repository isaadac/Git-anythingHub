'use strict';
// Load test for the live-editing server (src/main/collab-server.js).
//
//   node scripts/collab-load.js [--rooms 20] [--clients 10] [--rate 5] [--seconds 10] [--keys 200]
//                               [--size 200] [--url ws://host:port/]
//
// Starts a server (unless --url is given), opens rooms × clients WebSocket connections, and has
// every client change random keys of its room `rate` times a second for `seconds`. It measures
// how long each change takes to reach the other people in the room, and checks at the end that
// every copy of every room holds exactly the same document as the server, in the same order of
// changes (no gaps, nothing lost or duplicated). Exits 1 if anything is wrong.
const http = require('http');
const { createCollabServer } = require('../src/main/collab-server');

function options(argv = process.argv.slice(2)) {
  const o = { rooms: 20, clients: 10, rate: 5, seconds: 10, keys: 200, size: 200, url: null };
  for (let i = 0; i < argv.length; i += 2) {
    const k = argv[i].replace(/^--/, '');
    if (k in o) o[k] = k === 'url' ? argv[i + 1] : +argv[i + 1];
  }
  return o;
}

const sleep = ms => new Promise(r => setTimeout(r, ms));
const pct = (xs, p) => (xs.length ? xs[Math.min(xs.length - 1, Math.floor(xs.length * p))] : 0);

/** One person in a room: applies everything the server relays, in order, to its own copy. */
class Client {
  constructor(url, room, create, n, stats) {
    this.url = url; this.room = room; this.create = create; this.n = n; this.stats = stats;
    this.state = new Map(); this.seq = 0; this.sent = 0; this.errors = [];
  }
  open() {
    return new Promise((resolve, reject) => {
      const ws = this.ws = new WebSocket(this.url);
      ws.onerror = () => reject(new Error('connection failed'));
      ws.onopen = () => ws.send(JSON.stringify({ t: 'hello', room: this.room, create: this.create, app: 'write', name: 'Load ' + this.n, init: this.create ? { title: 'Load test' } : undefined }));
      ws.onmessage = e => {
        const m = JSON.parse(e.data);
        if (m.t === 'welcome') {
          this.id = m.id; this.seq = m.seq;
          for (const [k, v] of Object.entries(m.state)) this.state.set(k, v);
          resolve();
        } else if (m.t === 'op') {
          if (m.seq !== this.seq + 1) this.errors.push(`gap: expected ${this.seq + 1}, got ${m.seq}`);
          this.seq = m.seq;
          for (const o of m.ops) { if (o.v === null) this.state.delete(o.k); else this.state.set(o.k, o.v); }
          if (m.from !== this.id) for (const o of m.ops) if (o.v && o.v.at) this.stats.latency.push(Date.now() - o.v.at);
          this.stats.received++;
        } else if (m.t === 'error') { this.errors.push(m.error); reject(new Error(m.error)); }
      };
      ws.onclose = () => { this.closed = true; };
    });
  }
  change(keys, size) {
    const k = 'p' + Math.floor(Math.random() * keys);
    // Some changes delete a key, as removing a paragraph does.
    const v = Math.random() < 0.05 ? null : { at: Date.now(), by: this.n, text: 'x'.repeat(size) };
    this.ws.send(JSON.stringify({ t: 'op', n: ++this.sent, ops: [{ k, v }] }));
    this.stats.sent++;
  }
}

async function run(o = options()) {
  let server = null, url = o.url;
  if (!url) {
    server = createCollabServer({ server: http.createServer() });
    await new Promise(r => server.listen(0, '127.0.0.1', r));
    url = `ws://127.0.0.1:${server.address().port}/`;
  }
  const stats = { sent: 0, received: 0, latency: [] };
  const rooms = [];
  const t0 = Date.now();
  for (let r = 0; r < o.rooms; r++) {
    const id = 'load-' + r + '-' + Math.random().toString(36).slice(2, 12).padEnd(10, '0');
    const first = new Client(url, id, true, 0, stats);
    await first.open();
    const others = [];
    for (let c = 1; c < o.clients; c++) others.push(new Client(url, id, false, c, stats));
    await Promise.all(others.map(c => c.open()));
    rooms.push({ id, clients: [first, ...others] });
  }
  const connectMs = Date.now() - t0;
  const all = rooms.flatMap(r => r.clients);
  const mem0 = process.memoryUsage().rss;
  // Everyone types at once, each at its own steady pace.
  const end = Date.now() + o.seconds * 1000;
  const every = 1000 / o.rate;
  await Promise.all(all.map(async c => {
    await sleep(Math.random() * every);
    while (Date.now() < end) { c.change(o.keys, o.size); await sleep(every * (0.5 + Math.random())); }
  }));
  // Let the last changes arrive.
  const expected = () => rooms.every(r => { const s = server ? server.rooms.get(r.id).seq : null; return r.clients.every(c => s === null || c.seq === s); });
  for (let i = 0; i < 100 && !expected(); i++) await sleep(100);
  const secs = (Date.now() - end + o.seconds * 1000) / 1000;
  // Every copy must be the server's document.
  const problems = [];
  for (const r of rooms) {
    const ref = server ? server.rooms.get(r.id) : null;
    const want = ref ? JSON.stringify([...ref.state].map(([k, e]) => [k, e.v]).sort()) : JSON.stringify([...r.clients[0].state].sort());
    for (const c of r.clients) {
      if (c.errors.length) problems.push(`${r.id} client ${c.n}: ${c.errors.slice(0, 3).join('; ')}`);
      if (JSON.stringify([...c.state].sort()) !== want) problems.push(`${r.id} client ${c.n}: its document differs from the server's`);
      if (ref && c.seq !== ref.seq) problems.push(`${r.id} client ${c.n}: at change ${c.seq} of ${ref.seq}`);
    }
  }
  const lat = stats.latency.sort((a, b) => a - b);
  const result = {
    rooms: o.rooms, clients: all.length, connectMs,
    sent: stats.sent, delivered: stats.received, perSecond: Math.round(stats.received / secs),
    latencyMs: { p50: pct(lat, 0.5), p95: pct(lat, 0.95), p99: pct(lat, 0.99), max: lat[lat.length - 1] || 0 },
    processMemoryMB: Math.round((process.memoryUsage().rss - mem0) / 1048576),
    problems
  };
  for (const c of all) c.ws.close();
  if (server) await new Promise(r => { server.close(r); for (const room of server.rooms.values()) for (const c of room.clients) c.socket.destroy(); });
  return result;
}

module.exports = { run, options };

if (require.main === module) {
  run().then(r => {
    console.log(JSON.stringify(r, null, 2));
    process.exit(r.problems.length ? 1 : 0);
  }).catch(e => { console.error(e); process.exit(1); });
}
