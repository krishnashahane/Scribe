// Yjs collaboration over WebSockets (y-websocket wire protocol), with
// server-side authorization: every connection is tied to a session and a
// document role; viewers and commenters receive updates but cannot send them.
import * as Y from 'yjs'
import * as syncProtocol from 'y-protocols/sync'
import * as awarenessProtocol from 'y-protocols/awareness'
import * as encoding from 'lib0/encoding'
import * as decoding from 'lib0/decoding'
import { WebSocketServer } from 'ws'
import { parseCookies, SESSION_COOKIE, userFromToken } from './auth.js'
import { canComment, canEdit, roleFor } from './db.js'

const MSG_SYNC = 0
const MSG_AWARENESS = 1

export function createCollab(db, { log = () => {} } = {}) {
  const rooms = new Map()

  function getRoom(docId) {
    let room = rooms.get(docId)
    if (room) return room
    const doc = new Y.Doc({ gc: true })
    const row = db.prepare('SELECT ystate FROM documents WHERE id = ?').get(docId)
    if (row?.ystate) Y.applyUpdate(doc, new Uint8Array(row.ystate))
    const awareness = new awarenessProtocol.Awareness(doc)
    awareness.setLocalState(null)
    room = { id: docId, doc, awareness, conns: new Map(), timer: null, dirty: false }
    const persist = () => {
      room.timer = null
      if (!room.dirty) return
      room.dirty = false
      db.prepare('UPDATE documents SET ystate = ?, updated = ? WHERE id = ?').run(Buffer.from(Y.encodeStateAsUpdate(doc)), Date.now(), docId)
    }
    room.persist = persist
    doc.on('update', (update, origin) => {
      room.dirty = true
      if (!room.timer) room.timer = setTimeout(persist, 800)
      const enc = encoding.createEncoder()
      encoding.writeVarUint(enc, MSG_SYNC)
      syncProtocol.writeUpdate(enc, update)
      const msg = encoding.toUint8Array(enc)
      for (const ws of room.conns.keys()) if (ws !== origin) send(ws, msg)
    })
    awareness.on('update', ({ added, updated, removed }, origin) => {
      const changed = added.concat(updated, removed)
      const meta = room.conns.get(origin)
      if (meta) {
        for (const id of added) meta.ids.add(id)
        for (const id of removed) meta.ids.delete(id)
      }
      const enc = encoding.createEncoder()
      encoding.writeVarUint(enc, MSG_AWARENESS)
      encoding.writeVarUint8Array(enc, awarenessProtocol.encodeAwarenessUpdate(awareness, changed))
      const msg = encoding.toUint8Array(enc)
      for (const ws of room.conns.keys()) send(ws, msg)
    })
    rooms.set(docId, room)
    return room
  }

  function send(ws, msg) {
    if (ws.readyState !== 1) return
    try {
      ws.send(msg)
    } catch {
      ws.close()
    }
  }

  function closeConn(room, ws) {
    const meta = room.conns.get(ws)
    if (!meta) return
    room.conns.delete(ws)
    awarenessProtocol.removeAwarenessStates(room.awareness, [...meta.ids], null)
    if (room.conns.size === 0) {
      if (room.timer) clearTimeout(room.timer)
      room.persist()
      room.doc.destroy()
      rooms.delete(room.id)
    }
  }

  const wss = new WebSocketServer({ noServer: true, maxPayload: 64 * 1024 * 1024 })

  function handleUpgrade(req, socket, head) {
    const url = new URL(req.url, 'http://localhost')
    const m = /^\/collab\/([A-Za-z0-9_-]{6,64})$/.exec(url.pathname)
    const reject = (code, text) => {
      socket.write(`HTTP/1.1 ${code} ${text}\r\nConnection: close\r\n\r\n`)
      socket.destroy()
    }
    if (!m) return reject(404, 'Not Found')
    // Cross-site WebSocket hijacking defence: the Origin must match the Host
    const origin = req.headers.origin
    if (origin && new URL(origin).host !== req.headers.host) return reject(403, 'Forbidden')
    const user = userFromToken(db, parseCookies(req.headers.cookie)[SESSION_COOKIE])
    if (!user) return reject(401, 'Unauthorized')
    const role = roleFor(db, m[1], user.id)
    if (!role) return reject(403, 'Forbidden')
    wss.handleUpgrade(req, socket, head, (ws) => onConnection(ws, m[1], user, role))
  }

  function onConnection(ws, docId, user, role) {
    const room = getRoom(docId)
    const readOnly = !canEdit(role)
    room.conns.set(ws, { ids: new Set(), user, role, readOnly })
    ws.binaryType = 'arraybuffer'
    ws.on('message', (data) => {
      try {
        const decoder = decoding.createDecoder(new Uint8Array(data))
        const type = decoding.readVarUint(decoder)
        if (type === MSG_SYNC) {
          const enc = encoding.createEncoder()
          encoding.writeVarUint(enc, MSG_SYNC)
          const sub = decoding.readVarUint(decoder)
          if (sub === syncProtocol.messageYjsSyncStep1) syncProtocol.readSyncStep1(decoder, enc, room.doc)
          else if (!readOnly && (sub === syncProtocol.messageYjsSyncStep2 || sub === syncProtocol.messageYjsUpdate)) {
            Y.applyUpdate(room.doc, decoding.readVarUint8Array(decoder), ws)
          }
          if (encoding.length(enc) > 1) send(ws, encoding.toUint8Array(enc))
        } else if (type === MSG_AWARENESS) {
          awarenessProtocol.applyAwarenessUpdate(room.awareness, decoding.readVarUint8Array(decoder), ws)
        }
      } catch (e) {
        log('collab message error', e.message)
        ws.close(1003, 'bad message')
      }
    })
    ws.on('close', () => closeConn(room, ws))
    ws.on('error', () => closeConn(room, ws))
    // heartbeat
    let alive = true
    ws.on('pong', () => (alive = true))
    const hb = setInterval(() => {
      if (!alive) {
        clearInterval(hb)
        return ws.terminate()
      }
      alive = false
      try {
        ws.ping()
      } catch {
        /* closed */
      }
    }, 30000)
    ws.on('close', () => clearInterval(hb))
    // initial sync
    const enc = encoding.createEncoder()
    encoding.writeVarUint(enc, MSG_SYNC)
    syncProtocol.writeSyncStep1(enc, room.doc)
    send(ws, encoding.toUint8Array(enc))
    const states = [...room.awareness.getStates().keys()]
    if (states.length) {
      const a = encoding.createEncoder()
      encoding.writeVarUint(a, MSG_AWARENESS)
      encoding.writeVarUint8Array(a, awarenessProtocol.encodeAwarenessUpdate(room.awareness, states))
      send(ws, encoding.toUint8Array(a))
    }
    void canComment
  }

  function flushAll() {
    for (const room of rooms.values()) {
      if (room.timer) clearTimeout(room.timer)
      room.persist()
    }
  }

  /** Revoke live access when a share is removed. */
  function kick(docId, userId) {
    const room = rooms.get(docId)
    if (!room) return
    for (const [ws, meta] of room.conns) if (meta.user.id === userId) ws.close(4403, 'access revoked')
  }

  return { handleUpgrade, flushAll, kick, rooms }
}
