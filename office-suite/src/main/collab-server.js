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

const LIMITS = { message: 32e6, roomBytes: 200e6, rooms: 200, clientsPerRoom: 50, idleMs: 60 * 60 * 1000, keyLen: 300 };
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
  constructor(socket, onMessage, onClose) {
    this.socket = socket;
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
          this.onMessage(this, j, msg.length);
        }
        continue;
      }
      return this.close(1002);
    }
  }
  send(obj, opcode = 1) {
    if (this.closed) return;
    try { this.socket.write(encodeFrame(typeof obj === 'string' ? obj : JSON.stringify(obj), opcode)); } catch { this.end(); }
  }
  close(code = 1000) {
    if (this.closed) return;
    try { this.socket.write(Buffer.from([0x88, 2, code >> 8, code & 255])); } catch { /* gone */ }
    this.socket.end();
    this.end();
  }
  end() {
    if (this.closed) return;
    this.closed = true;
    this.socket.destroy();
    this.onClose(this);
  }
}

// ---------- Sessions ----------
function createCollabServer(opts = {}) {
  const limits = { ...LIMITS, ...opts.limits };
  const rooms = new Map();
  const server = http.createServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Isadek Office live editing server\n');
  });
  const validRoom = r => typeof r === 'string' && /^[\w-]{16,64}$/.test(r);
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
      const ops = m.ops.filter(o => o && typeof o.k === 'string' && o.k.length <= limits.keyLen);
      if (room.bytes + size > limits.roomBytes) return c.send({ t: 'error', error: 'The document is too large for this server' });
      for (const o of ops) set(room, o.k, o.v === undefined ? null : o.v);
      room.seq++;
      broadcast(room, { t: 'op', from: c.id, n: m.n, seq: room.seq, ops });
      return;
    }
    if (m.t === 'presence') {
      c.presence = m.data && typeof m.data === 'object' ? m.data : null;
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
    const room = c.room;
    if (!room) return;
    room.clients.delete(c);
    broadcast(room, { t: 'peers', peers: peerList(room) });
    // An empty session is kept for a while, so people can reconnect.
    if (!room.clients.size) room.idle = setTimeout(() => { if (!room.clients.size) rooms.delete(room.id); }, limits.idleMs);
  }

  server.on('upgrade', (req, socket) => {
    const key = req.headers['sec-websocket-key'];
    if (!key || (req.headers.upgrade || '').toLowerCase() !== 'websocket') { socket.end('HTTP/1.1 400 Bad Request\r\n\r\n'); return; }
    const accept = crypto.createHash('sha1').update(key + GUID).digest('base64');
    socket.write(`HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Accept: ${accept}\r\n\r\n`);
    const c = new Conn(socket, onMessage, onClose);
    // Keep idle connections alive through proxies.
    const ping = setInterval(() => (c.closed ? clearInterval(ping) : c.send('', 9)), 30000);
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
