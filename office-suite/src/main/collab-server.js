'use strict';
// Isadek Office live-editing relay. Each session (room) holds the document as a map of keys
// (paragraphs, cells, slides, shapes) to their latest value; clients send changed keys and the
// server puts every change in one order and relays it, so all copies converge (last write wins
// per key). It has no dependencies: a small WebSocket (RFC 6455) server on Node's http module.
//
// Run your own:  node src/main/collab-server.js [--port 8787] [--host 0.0.0.0]
// The desktop app can also host one on the local network (File → Live Session).
const http = require('http');
const crypto = require('crypto');

// backlog: bytes waiting to go to one connection; a person whose connection stalls is dropped
// beyond it (their app reconnects and catches up) instead of the server holding it all.
// rate: bytes one connection may send per minute (a burst up to a whole minute's worth);
// connections: open sockets in all; helloMs: how long a new socket may take to join a session.
const LIMITS = { message: 32e6, roomBytes: 200e6, rooms: 200, clientsPerRoom: 50, idleMs: 60 * 60 * 1000, keyLen: 300, backlog: 64e6,
  rate: 256e6, connections: 2000, helloMs: 15000, presence: 8000, opsPerMessage: 50000 };
const GUID = '258EAFA5-E914-47DA-95CA-C5AB0DC85B11';

// ---------- WebSocket framing ----------
function encodeFrame(text, opcode = 1) {
  const payload = Buffer.from(text);
  const n = payload.length;
  const head = n < 126 ? Buffer.from([0x80 | opcode, n]) : n < 65536 ? Buffer.from([0x80 | opcode, 126, n >> 8, n & 255]) : Buffer.alloc(10);
  if (n >= 65536) { head[0] = 0x80 | opcode; head[1] = 127; head.writeBigUInt64BE(BigInt(n), 2); }
  return Buffer.concat([head, payload]);
}

class Conn {
  constructor(socket, onMessage, onClose, backlog = LIMITS.backlog, rate = LIMITS.rate) {
    this.socket = socket;
    this.backlog = backlog;
    // Token bucket: what this connection may still send before it has to slow down.
    this.rate = rate;
    this.allowance = rate;
    this.lastTick = Date.now();
    this.buf = Buffer.alloc(0);
    this.parts = [];
    this.partBytes = 0;
    this.closed = false;
    this.onMessage = onMessage;
    this.onClose = onClose;
    socket.on('data', d => this.data(d));
    socket.on('close', () => this.end());
    socket.on('error', () => this.end());
    socket.setNoDelay(true);
  }
  data(d) {
    if (this.closed) return;
    const now = Date.now();
    this.allowance = Math.min(this.rate, this.allowance + (now - this.lastTick) * this.rate / 60000);
    this.lastTick = now;
    this.allowance -= d.length;
    // Far more than anyone typing or pasting sends: a runaway or hostile client.
    if (this.allowance < 0) return this.close(1008);
    this.buf = this.buf.length ? Buffer.concat([this.buf, d]) : d;
    for (;;) {
      const b = this.buf;
      if (b.length < 2) return;
      const fin = b[0] & 0x80, op = b[0] & 0x0f, masked = b[1] & 0x80;
      let len = b[1] & 0x7f, at = 2;
      if (len === 126) { if (b.length < 4) return; len = b.readUInt16BE(2); at = 4; }
      else if (len === 127) { if (b.length < 10) return; const big = b.readBigUInt64BE(2); if (big > BigInt(LIMITS.message)) return this.close(1009); len = Number(big); at = 10; }
      if (!masked) return this.close(1002); // clients must mask
      if (len > LIMITS.message || this.partBytes + len > LIMITS.message) return this.close(1009);
      if (b.length < at + 4 + len) return;
      const mask = b.subarray(at, at + 4);
      const payload = Buffer.from(b.subarray(at + 4, at + 4 + len));
      for (let i = 0; i < payload.length; i++) payload[i] ^= mask[i & 3];
      this.buf = b.subarray(at + 4 + len);
      if (op === 8) return this.close(1000);
      if (op === 9) { this.send(payload.toString(), 10); continue; }
      if (op === 10) continue;
      if (op === 1 || op === 2 || op === 0) {
        this.parts.push(payload);
        this.partBytes += payload.length;
        if (fin) {
          const msg = Buffer.concat(this.parts).toString('utf8');
          this.parts = []; this.partBytes = 0;
          let j;
          try { j = JSON.parse(msg); } catch { return this.close(1007); }
          // Whatever a message holds, it can't take the server down: the sender is let go.
          try { this.onMessage(this, j, msg.length); } catch { return this.close(1011); }
        }
        continue;
      }
      return this.close(1002);
    }
  }
  send(obj, opcode = 1) {
    if (this.closed) return;
    try { this.socket.write(encodeFrame(typeof obj === 'string' ? obj : JSON.stringify(obj), opcode)); } catch { this.end(); return; }
    if (this.socket.writableLength > this.backlog) this.end();
  }
  close(code = 1000) {
    if (this.closed) return;
    try { this.socket.write(Buffer.from([0x88, 2, code >> 8, code & 255])); } catch { /* gone */ }
    // Let the close frame (with its reason) go out first; then the socket goes regardless.
    this.socket.end();
    this.end(true);
  }
  end(gently) {
    if (this.closed) return;
    this.closed = true;
    clearInterval(this.ping);
    this.buf = Buffer.alloc(0);
    if (gently) setTimeout(() => this.socket.destroy(), 2000).unref(); else this.socket.destroy();
    this.onClose(this);
  }
}

// ---------- Sessions ----------
function createCollabServer(opts = {}) {
  const limits = { ...LIMITS, ...opts.limits };
  const rooms = new Map();
  // opts.server: share an existing HTTP server (the web server), else a server of its own.
  const server = opts.server || http.createServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Isadek Office live editing server\n');
  });
  const validRoom = r => typeof r === 'string' && /^[\w-]{16,64}$/.test(r);
  let open = 0; // sockets connected now
  const peerList = room => [...room.clients].map(c => ({ id: c.id, name: c.name, color: c.color, presence: c.presence || null }));
  const broadcast = (room, msg, except) => { const s = JSON.stringify(msg); for (const c of room.clients) if (c !== except) c.send(s); };

  function onMessage(c, m, size) {
    if (!m || typeof m !== 'object') return;
    if (m.t === 'hello') {
      if (c.room || !validRoom(m.room)) return c.send({ t: 'error', error: 'Invalid session' });
      let room = rooms.get(m.room);
      if (m.create) {
        if (room) return c.send({ t: 'error', error: 'That session already exists' });
        if (rooms.size >= limits.rooms) return c.send({ t: 'error', error: 'The server is full' });
        room = { id: m.room, app: String(m.app || '').slice(0, 20), state: new Map(), seq: 0, bytes: 0, clients: new Set(), idle: null };
        rooms.set(m.room, room);
        if (m.init && typeof m.init === 'object') {
          for (const [k, v] of Object.entries(m.init)) if (typeof k === 'string' && k.length <= limits.keyLen && v !== null) set(room, k, v);
        }
      } else if (!room) return c.send({ t: 'error', error: 'This live session has ended or the link is wrong' });
      if (room.app && m.app && room.app !== m.app) return c.send({ t: 'error', error: `This session is for Isadek ${room.app[0].toUpperCase() + room.app.slice(1)}` });
      if (room.clients.size >= limits.clientsPerRoom) return c.send({ t: 'error', error: 'This session is full' });
      clearTimeout(room.idle);
      c.room = room;
      c.id = crypto.randomBytes(6).toString('hex');
      c.name = String(m.name || 'Guest').slice(0, 60);
      c.color = /^#[0-9a-f]{6}$/i.test(m.color) ? m.color : '#888888';
      clearTimeout(c.helloTimer);
      room.clients.add(c);
      const state = {};
      for (const [k, e] of room.state) state[k] = e.v;
      c.send({ t: 'welcome', id: c.id, seq: room.seq, state, peers: peerList(room), app: room.app });
      broadcast(room, { t: 'peers', peers: peerList(room) }, c);
      return;
    }
    const room = c.room;
    if (!room) return;
    if (m.t === 'op' && Array.isArray(m.ops)) {
      if (m.ops.length > limits.opsPerMessage) return c.close(1009);
      const ops = m.ops.filter(o => o && typeof o === 'object' && typeof o.k === 'string' && o.k.length <= limits.keyLen);
      if (room.bytes + size > limits.roomBytes) return c.send({ t: 'error', error: 'The document is too large for this server' });
      for (const o of ops) set(room, o.k, o.v === undefined ? null : o.v);
      room.seq++;
      // seen: the last change the sender had when it made this one (so others can merge with it).
      broadcast(room, { t: 'op', from: c.id, n: Number.isInteger(m.n) ? m.n : undefined, seq: room.seq, seen: Number.isInteger(m.seen) ? m.seen : undefined, ops });
      return;
    }
    if (m.t === 'presence') {
      // Where someone is (a cell, a slide): small by nature, so anything big is dropped.
      const p = m.data && typeof m.data === 'object' && !Array.isArray(m.data) ? m.data : null;
      c.presence = p && JSON.stringify(p).length <= limits.presence ? p : null;
      broadcast(room, { t: 'presence', from: c.id, data: c.presence }, c);
    }
  }
  function set(room, k, v) {
    const old = room.state.get(k);
    if (old) room.bytes -= old.size;
    if (v === null) { room.state.delete(k); return; }
    const size = JSON.stringify(v).length + k.length;
    room.state.set(k, { v, size });
    room.bytes += size;
  }
  function onClose(c) {
    open--;
    clearTimeout(c.helloTimer);
    const room = c.room;
    if (!room) return;
    room.clients.delete(c);
    broadcast(room, { t: 'peers', peers: peerList(room) });
    // An empty session is kept for a while, so people can reconnect.
    if (!room.clients.size) { room.idle = setTimeout(() => { if (!room.clients.size) rooms.delete(room.id); }, limits.idleMs); room.idle.unref(); }
  }

  server.on('upgrade', (req, socket, head) => {
    socket.on('error', () => socket.destroy());
    const key = req.headers['sec-websocket-key'];
    if (!key || !/^[A-Za-z0-9+/=]{16,64}$/.test(key) || (req.headers.upgrade || '').toLowerCase() !== 'websocket') { socket.end('HTTP/1.1 400 Bad Request\r\nContent-Length: 0\r\nConnection: close\r\n\r\n'); return; }
    if (open >= limits.connections) { socket.end('HTTP/1.1 503 Service Unavailable\r\nContent-Length: 0\r\nConnection: close\r\n\r\n'); return; }
    open++;
    const accept = crypto.createHash('sha1').update(key + GUID).digest('base64');
    socket.write(`HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Accept: ${accept}\r\n\r\n`);
    const c = new Conn(socket, onMessage, onClose, limits.backlog, limits.rate);
    // A socket that never joins a session is let go.
    c.helloTimer = setTimeout(() => { if (!c.room) c.close(1008); }, limits.helloMs);
    if (head && head.length) c.data(head); // bytes that came with the handshake
    // Keep idle connections alive through proxies.
    c.ping = setInterval(() => c.send('', 9), 30000);
  });
  server.rooms = rooms;
  return server;
}

module.exports = { createCollabServer };

if (require.main === module) {
  const arg = (n, d) => { const i = process.argv.indexOf(n); return i > 0 ? process.argv[i + 1] : d; };
  const port = +arg('--port', process.env.PORT || 8787), host = arg('--host', '0.0.0.0');
  createCollabServer().listen(port, host, () => console.log(`Isadek Office live editing server on ws://${host}:${port}/`));
}
