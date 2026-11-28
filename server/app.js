// SCRIBE HTTP API. Runs as a Vercel Function (api/index.js) or under a plain Node server (server/index.js).
import express from 'express'
import path from 'node:path'
import fs from 'node:fs'
import { fileURLToPath } from 'node:url'
import * as Y from 'yjs'
import { canEdit, openDb, rateLimit, roleFor, toBuffer } from './db.js'
import {
  DUMMY_HASH, GUEST_DOMAIN, GUEST_TTL_MS, SESSION_COOKIE, clearCookie, createSession, destroySession, hashPassword, isGuestEmail,
  newId, parseCookies, sessionCookie, userFromToken, verifyPassword,
} from './auth.js'
import { AiError, aiConfig, runAi } from './ai.js'
import { CHUNK_BYTES, MAX_JSON_CHARS, MAX_MEDIA_BYTES, MAX_UPDATE_BYTES, check, jsonSize, params, sanitizeContent, sanitizeMeta, schemas } from './validate.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const PRESENCE_TTL_MS = 30_000
const SYNC_RESPONSE_BYTES = 3_000_000
const COMPACT_AFTER_ROWS = 400

export const SECURITY_HEADERS = {
  'Content-Security-Policy':
    "default-src 'self'; img-src 'self' data: blob: https:; style-src 'self' 'unsafe-inline'; font-src 'self' data:; script-src 'self'; connect-src 'self'; media-src 'self' data: blob:; worker-src 'self' blob:; manifest-src 'self'; frame-src 'none'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'",
  'Strict-Transport-Security': 'max-age=63072000; includeSubDomains; preload',
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'Referrer-Policy': 'no-referrer',
  'Cross-Origin-Opener-Policy': 'same-origin',
  'Cross-Origin-Resource-Policy': 'same-origin',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=(), payment=(), usb=(), serial=(), hid=(), bluetooth=(), midi=(), display-capture=(), browsing-topics=(), clipboard-read=(self), clipboard-write=(self)',
  'X-Permitted-Cross-Domain-Policies': 'none',
  'X-DNS-Prefetch-Control': 'off',
  'Origin-Agent-Cluster': '?1',
}

/**
 * @param {{ db?: any, dbUrl?: string, dbDir?: string, secureCookies?: boolean, serveStatic?: boolean, authRateLimit?: number, allowedOrigins?: string[] }} opts
 */
export async function createApp(opts = {}) {
  const db = opts.db ?? (await openDb({ url: opts.dbUrl ?? process.env.DATABASE_URL, dir: opts.dbDir }))
  const secureCookies = opts.secureCookies ?? (process.env.NODE_ENV === 'production' || !!process.env.VERCEL)
  const authRateLimit = opts.authRateLimit ?? 10
  const allowedOrigins = new Set(opts.allowedOrigins ?? (process.env.ALLOWED_ORIGINS || '').split(',').map((s) => s.trim()).filter(Boolean))
  const app = express()
  app.disable('x-powered-by')
  app.set('trust proxy', process.env.VERCEL ? true : 'loopback')

  const ah = (fn) => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next)

  app.use((req, res, next) => {
    for (const [k, v] of Object.entries(SECURITY_HEADERS)) res.setHeader(k, v)
    next()
  })

  // ---- strict CORS: same-origin only (plus an explicit allow-list); no wildcard
  app.use('/api', (req, res, next) => {
    const origin = req.get('origin')
    if (origin) {
      const sameHost = URL.canParse(origin) && new URL(origin).host === req.get('host')
      if (!sameHost && !allowedOrigins.has(origin)) return res.status(403).json({ error: 'Cross-origin request refused' })
      if (!sameHost) {
        res.setHeader('Access-Control-Allow-Origin', origin)
        res.setHeader('Access-Control-Allow-Credentials', 'true')
        res.setHeader('Vary', 'Origin')
      }
    }
    if (req.method === 'OPTIONS') {
      res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PATCH, PUT, DELETE')
      res.setHeader('Access-Control-Allow-Headers', 'content-type, x-scribe-request')
      return res.status(204).end()
    }
    next()
  })

  app.use('/api', express.json({ limit: '4mb' }))

  // ---- session + CSRF guard
  app.use(
    '/api',
    ah(async (req, res, next) => {
      req.user = await userFromToken(db, parseCookies(req.headers.cookie)[SESSION_COOKIE])
      // A custom header cannot be sent cross-site without a CORS preflight, which only allow-listed origins pass.
      if (!['GET', 'HEAD'].includes(req.method) && req.get('x-scribe-request') !== '1') return res.status(403).json({ error: 'Missing request header' })
      next()
    }),
  )
  const requireUser = (req, res, next) => (req.user ? next() : res.status(401).json({ error: 'Sign in required' }))
  const limited = (bucket, limit, windowMs, keyOf) =>
    ah(async (req, res, next) => {
      if (await rateLimit(db, `${bucket}:${keyOf(req)}`, limit, windowMs)) return next()
      res.setHeader('Retry-After', Math.ceil(windowMs / 1000))
      res.status(429).json({ error: 'Too many requests. Please wait a moment and try again.' })
    })
  const byIp = (req) => req.ip || 'unknown'
  const byUser = (req) => req.user?.id || req.ip || 'unknown'
  const authLimit = limited('auth', authRateLimit, 60_000, byIp)
  const guestLimit = limited('guest', Math.max(authRateLimit, 20), 60_000, byIp)
  const apiLimit = limited('api', 1200, 60_000, byUser)
  const blobLimit = limited('blob', 300, 60_000, byUser)
  const aiLimit = limited('ai', 30, 60_000, byUser)
  // sharing reveals whether an email has an account, so it gets its own tighter limit
  const shareLimit = limited('share', 30, 60_000, byUser)
  app.use('/api', (req, res, next) => (req.path.startsWith('/auth/') || req.path === '/health' ? next() : apiLimit(req, res, next)))

  /** Start a fresh session, ending the one this browser had before (no session fixation, no orphaned tokens). */
  const startSession = async (req, res, userId) => {
    await destroySession(db, parseCookies(req.headers.cookie)[SESSION_COOKIE])
    res.setHeader('Set-Cookie', sessionCookie(await createSession(db, userId), secureCookies))
  }
  const publicUser = (u) => (u ? { id: u.id, email: isGuestEmail(u.email) ? '' : u.email, name: u.name, guest: isGuestEmail(u.email) } : null)

  // ---------------------------------------------------------------- auth
  app.post(
    '/api/auth/register',
    authLimit,
    ah(async (req, res) => {
      const b = check(schemas.register, req.body, res)
      if (!b) return
      if (isGuestEmail(b.email)) return res.status(400).json({ error: 'Enter a valid email address.' })
      const id = newId()
      const name = b.name || b.email.split('@')[0]
      const { rows } = await db.query(
        'INSERT INTO users (id, email, name, pass_hash, created) VALUES ($1, $2, $3, $4, $5) ON CONFLICT (email) DO NOTHING RETURNING id',
        [id, b.email, name, await hashPassword(b.password), Date.now()],
      )
      if (!rows.length) return res.status(409).json({ error: 'An account with this email already exists.' })
      await startSession(req, res, id)
      res.status(201).json({ user: publicUser({ id, email: b.email, name }) })
    }),
  )

  app.post(
    '/api/auth/login',
    authLimit,
    ah(async (req, res) => {
      const b = check(schemas.login, req.body, res)
      if (!b) return
      const { rows } = await db.query('SELECT id, email, name, pass_hash FROM users WHERE email = $1', [b.email])
      const u = rows[0]
      const usable = u && !isGuestEmail(u.email)
      const ok = await verifyPassword(b.password, usable ? u.pass_hash : DUMMY_HASH)
      if (!usable || !ok) return res.status(401).json({ error: 'Email or password is incorrect.' })
      await startSession(req, res, u.id)
      res.json({ user: publicUser(u) })
    }),
  )

  app.post(
    '/api/auth/logout',
    ah(async (req, res) => {
      await destroySession(db, parseCookies(req.headers.cookie)[SESSION_COOKIE])
      res.setHeader('Set-Cookie', clearCookie(secureCookies))
      res.json({ ok: true })
    }),
  )

  app.get('/api/auth/me', (req, res) => res.json({ user: publicUser(req.user) }))

  // No-login start: a private local identity per browser. Guests cannot sign in with a password.
  app.post(
    '/api/auth/guest',
    guestLimit,
    ah(async (req, res) => {
      if (req.user) return res.json({ user: publicUser(req.user) })
      const id = newId()
      const email = `${id.toLowerCase()}${GUEST_DOMAIN}`
      await db.query('INSERT INTO users (id, email, name, pass_hash, created) VALUES ($1, $2, $3, $4, $5)', [id, email, 'You', '!guest', Date.now()])
      res.setHeader('Set-Cookie', sessionCookie(await createSession(db, id, GUEST_TTL_MS), secureCookies, GUEST_TTL_MS))
      res.status(201).json({ user: publicUser({ id, email, name: 'You' }) })
    }),
  )

  app.post(
    '/api/auth/upgrade',
    authLimit,
    requireUser,
    ah(async (req, res) => {
      if (!isGuestEmail(req.user.email)) return res.status(400).json({ error: 'This is already a full account.' })
      const b = check(schemas.upgrade, req.body, res)
      if (!b) return
      if (isGuestEmail(b.email)) return res.status(400).json({ error: 'Enter a valid email address.' })
      const exists = await db.query('SELECT 1 FROM users WHERE email = $1', [b.email])
      if (exists.rows.length) return res.status(409).json({ error: 'An account with this email already exists. Sign in instead.' })
      const name = b.name || b.email.split('@')[0]
      try {
        await db.query('UPDATE users SET email = $1, name = $2, pass_hash = $3 WHERE id = $4', [b.email, name, await hashPassword(b.password), req.user.id])
      } catch (e) {
        // another request took the email between the check and the update
        if (e?.code === '23505') return res.status(409).json({ error: 'An account with this email already exists. Sign in instead.' })
        throw e
      }
      // the year-long guest session becomes an ordinary account session
      await db.query('DELETE FROM sessions WHERE user_id = $1', [req.user.id])
      res.setHeader('Set-Cookie', sessionCookie(await createSession(db, req.user.id), secureCookies))
      res.json({ user: publicUser({ id: req.user.id, email: b.email, name }) })
    }),
  )

  // ---------------------------------------------------------------- documents
  const docAccess = (need) =>
    ah(async (req, res, next) => {
      const p = check(params.doc, { id: req.params.id }, res)
      if (!p) return
      const role = await roleFor(db, p.id, req.user.id)
      if (!role) return res.status(404).json({ error: 'Document not found' })
      if (need === 'edit' && !canEdit(role)) return res.status(403).json({ error: 'You have view-only access' })
      if (need === 'owner' && role !== 'owner') return res.status(403).json({ error: 'Only the owner can do this' })
      req.role = role
      next()
    })

  app.get(
    '/api/docs',
    requireUser,
    ah(async (req, res) => {
      const { rows } = await db.query(
        `SELECT d.id, d.title, d.updated, d.created, u.name AS owner_name,
                CASE WHEN d.owner_id = $1 THEN 'owner' ELSE s.role END AS role,
                (d.original_blob IS NOT NULL) AS has_original
         FROM documents d JOIN users u ON u.id = d.owner_id
         LEFT JOIN shares s ON s.doc_id = d.id AND s.user_id = $1
         WHERE d.deleted = FALSE AND (d.owner_id = $1 OR s.user_id IS NOT NULL)
         ORDER BY d.updated DESC LIMIT 500`,
        [req.user.id],
      )
      res.json({ docs: rows })
    }),
  )

  app.post(
    '/api/docs',
    requireUser,
    ah(async (req, res) => {
      const b = check(schemas.createDoc, req.body, res)
      if (!b) return
      const content = sanitizeContent(b.content)
      const meta = sanitizeMeta(b.meta)
      if (!content || jsonSize(content) + jsonSize(meta) > MAX_JSON_CHARS) return res.status(413).json({ error: 'Document is too large' })
      if (b.originalBlob) {
        const { rows } = await db.query("SELECT 1 FROM blobs WHERE id = $1 AND owner_id = $2 AND kind = 'original' AND complete = TRUE", [b.originalBlob, req.user.id])
        if (!rows.length) return res.status(400).json({ error: 'Original file upload not found' })
      }
      const update = b.yUpdate ? Buffer.from(b.yUpdate, 'base64') : null
      if (update && update.length > MAX_UPDATE_BYTES) return res.status(413).json({ error: 'Document is too large' })
      if (update) {
        try {
          Y.applyUpdate(new Y.Doc(), new Uint8Array(update))
        } catch {
          return res.status(400).json({ error: 'Invalid document state' })
        }
      }
      const id = newId()
      const now = Date.now()
      await db.tx(async (t) => {
        await t.query(
          'INSERT INTO documents (id, owner_id, title, created, updated, content_json, meta_json, original_blob, original_name) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)',
          [id, req.user.id, b.title || 'Untitled document', now, now, JSON.stringify(content), JSON.stringify(meta), b.originalBlob ?? null, b.originalName ?? null],
        )
        if (update) await t.query('INSERT INTO doc_updates (doc_id, data, size) VALUES ($1, $2, $3)', [id, update, update.length])
      })
      res.status(201).json({ id })
    }),
  )

  app.get(
    '/api/docs/:id',
    requireUser,
    docAccess('view'),
    ah(async (req, res) => {
      const { rows } = await db.query('SELECT id, title, created, updated, meta_json, content_json, original_blob, original_name FROM documents WHERE id = $1', [req.params.id])
      const d = rows[0]
      let originalParts = 0
      if (d.original_blob) originalParts = (await db.query('SELECT chunks FROM blobs WHERE id = $1', [d.original_blob])).rows[0]?.chunks ?? 0
      res.json({
        id: d.id, title: d.title, created: d.created, updated: d.updated, role: req.role,
        meta: JSON.parse(d.meta_json), content: JSON.parse(d.content_json), hasOriginal: !!d.original_blob, originalParts, originalName: d.original_name,
      })
    }),
  )

  app.patch(
    '/api/docs/:id',
    requireUser,
    docAccess('edit'),
    ah(async (req, res) => {
      const b = check(schemas.patchDoc, req.body, res)
      if (!b) return
      const sets = []
      const vals = []
      const add = (col, v) => {
        vals.push(v)
        sets.push(`${col} = $${vals.length}`)
      }
      if (b.title !== undefined) add('title', b.title || 'Untitled document')
      if (b.content !== undefined) {
        const c = sanitizeContent(b.content)
        if (!c || jsonSize(c) > MAX_JSON_CHARS) return res.status(413).json({ error: 'Document is too large' })
        add('content_json', JSON.stringify(c))
      }
      if (b.meta !== undefined) add('meta_json', JSON.stringify(sanitizeMeta(b.meta)))
      if (!sets.length) return res.json({ ok: true })
      add('updated', Date.now())
      vals.push(req.params.id)
      // column names above are fixed literals; only values are parameters
      await db.query(`UPDATE documents SET ${sets.join(', ')} WHERE id = $${vals.length}`, vals)
      res.json({ ok: true })
    }),
  )

  app.delete(
    '/api/docs/:id',
    requireUser,
    docAccess('owner'),
    ah(async (req, res) => {
      await db.query('UPDATE documents SET deleted = TRUE, updated = $1 WHERE id = $2', [Date.now(), req.params.id])
      res.json({ ok: true })
    }),
  )

  // original DOCX, downloaded in parts that each fit a serverless response
  app.get(
    '/api/docs/:id/original/:part',
    requireUser,
    docAccess('view'),
    ah(async (req, res) => {
      const p = check(params.part, req.params, res)
      if (!p) return
      const { rows } = await db.query('SELECT c.data FROM documents d JOIN blob_chunks c ON c.blob_id = d.original_blob WHERE d.id = $1 AND c.idx = $2', [p.id, p.part])
      if (!rows.length) return res.status(404).json({ error: 'Not found' })
      res.setHeader('Content-Type', 'application/octet-stream')
      res.setHeader('Cache-Control', 'private, no-store')
      res.send(toBuffer(rows[0].data))
    }),
  )

  // ---------------------------------------------------------------- real-time sync over HTTP
  // Clients push Yjs updates and pull everything after the last id they saw. Viewers can only pull.
  app.post(
    '/api/docs/:id/sync',
    requireUser,
    docAccess('view'),
    ah(async (req, res) => {
      const b = check(schemas.sync, req.body, res)
      if (!b) return
      const docId = req.params.id
      if (b.updates.length && !canEdit(req.role)) return res.status(403).json({ error: 'You have view-only access' })
      const decoded = b.updates.map((u) => Buffer.from(u, 'base64'))
      for (const u of decoded) {
        if (u.length > MAX_UPDATE_BYTES) return res.status(413).json({ error: 'Change is too large' })
        try {
          Y.decodeUpdate(new Uint8Array(u))
        } catch {
          return res.status(400).json({ error: 'Invalid update' })
        }
      }
      for (const u of decoded) await db.query('INSERT INTO doc_updates (doc_id, data, size) VALUES ($1, $2, $3)', [docId, u, u.length])
      if (decoded.length) await db.query('UPDATE documents SET updated = $1 WHERE id = $2', [Date.now(), docId])
      if (b.awareness) {
        await db.query(
          `INSERT INTO presence (doc_id, client_id, user_id, state, updated) VALUES ($1, $2, $3, $4, $5)
           ON CONFLICT (doc_id, client_id) DO UPDATE SET state = EXCLUDED.state, updated = EXCLUDED.updated
           WHERE presence.user_id = EXCLUDED.user_id`,
          [docId, b.awareness.clientId, req.user.id, Buffer.from(b.awareness.state, 'base64'), Date.now()],
        )
      }
      const { rows } = await db.query(
        `SELECT id, data FROM (
           SELECT id, data, size, SUM(size) OVER (ORDER BY id) AS run FROM doc_updates WHERE doc_id = $1 AND id > $2
         ) t WHERE run - size < $3 ORDER BY id LIMIT 500`,
        [docId, b.after, SYNC_RESPONSE_BYTES],
      )
      const last = rows.length ? Number(rows[rows.length - 1].id) : b.after
      const more = rows.length ? (await db.query('SELECT 1 FROM doc_updates WHERE doc_id = $1 AND id > $2 LIMIT 1', [docId, last])).rows.length > 0 : false
      const presence = await db.query('SELECT client_id, state FROM presence WHERE doc_id = $1 AND updated > $2 AND client_id <> $3', [
        docId, Date.now() - PRESENCE_TTL_MS, b.awareness?.clientId ?? -1,
      ])
      res.setHeader('Cache-Control', 'no-store')
      res.json({
        updates: rows.map((r) => ({ id: Number(r.id), data: toBuffer(r.data).toString('base64') })),
        last,
        more,
        presence: presence.rows.map((p) => ({ clientId: Number(p.client_id), state: toBuffer(p.state).toString('base64') })),
      })
      if (decoded.length) await compact(docId).catch(() => {})
    }),
  )

  /** Merge many small updates into one so new clients load quickly. */
  async function compact(docId) {
    const { rows } = await db.query('SELECT COUNT(*)::int AS n FROM doc_updates WHERE doc_id = $1', [docId])
    if (rows[0].n < COMPACT_AFTER_ROWS) return
    await db.tx(async (t) => {
      const all = await t.query('SELECT id, data FROM doc_updates WHERE doc_id = $1 ORDER BY id FOR UPDATE', [docId])
      if (!all.rows.length) return
      const merged = Buffer.from(Y.mergeUpdates(all.rows.map((r) => new Uint8Array(toBuffer(r.data)))))
      if (merged.length > MAX_UPDATE_BYTES) return
      const maxId = all.rows[all.rows.length - 1].id
      await t.query('DELETE FROM doc_updates WHERE doc_id = $1 AND id <= $2', [docId, maxId])
      await t.query('INSERT INTO doc_updates (doc_id, data, size) VALUES ($1, $2, $3)', [docId, merged, merged.length])
    })
  }

  // ---------------------------------------------------------------- binary uploads (chunked)
  const MAGIC = {
    'image/png': [0x89, 0x50, 0x4e, 0x47],
    'image/jpeg': [0xff, 0xd8, 0xff],
    'image/gif': [0x47, 0x49, 0x46],
    'image/webp': [0x52, 0x49, 0x46, 0x46],
    'image/bmp': [0x42, 0x4d],
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': [0x50, 0x4b, 0x03, 0x04],
  }

  app.post(
    '/api/blobs',
    requireUser,
    blobLimit,
    ah(async (req, res) => {
      const b = check(schemas.blob, req.body, res)
      if (!b) return
      const isMedia = b.kind === 'media'
      if (isMedia !== b.mime.startsWith('image/')) return res.status(400).json({ error: 'Invalid request' })
      if (isMedia && b.size > MAX_MEDIA_BYTES) return res.status(413).json({ error: 'Images must be smaller than 4 MB' })
      const id = newId(16)
      const chunks = Math.ceil(b.size / CHUNK_BYTES)
      await db.query('INSERT INTO blobs (id, owner_id, kind, mime, size, chunks, created) VALUES ($1, $2, $3, $4, $5, $6, $7)', [id, req.user.id, b.kind, b.mime, b.size, chunks, Date.now()])
      res.status(201).json({ id, chunks, chunkSize: CHUNK_BYTES })
    }),
  )

  app.put(
    '/api/blobs/:id/chunks/:idx',
    requireUser,
    blobLimit,
    express.raw({ type: 'application/octet-stream', limit: CHUNK_BYTES + 1024 }),
    ah(async (req, res) => {
      const p = check(params.blobChunk, req.params, res)
      if (!p) return
      const { rows } = await db.query('SELECT chunks, size, complete FROM blobs WHERE id = $1 AND owner_id = $2', [p.id, req.user.id])
      const blob = rows[0]
      if (!blob) return res.status(404).json({ error: 'Not found' })
      if (blob.complete || p.idx >= blob.chunks || !Buffer.isBuffer(req.body)) return res.status(400).json({ error: 'Invalid request' })
      const expected = p.idx === blob.chunks - 1 ? blob.size - CHUNK_BYTES * (blob.chunks - 1) : CHUNK_BYTES
      if (req.body.length !== expected) return res.status(400).json({ error: 'Invalid chunk size' })
      await db.query('INSERT INTO blob_chunks (blob_id, idx, data) VALUES ($1, $2, $3) ON CONFLICT (blob_id, idx) DO UPDATE SET data = EXCLUDED.data', [p.id, p.idx, req.body])
      res.json({ ok: true })
    }),
  )

  app.post(
    '/api/blobs/:id/complete',
    requireUser,
    ah(async (req, res) => {
      const p = check(params.blob, req.params, res)
      if (!p) return
      const { rows } = await db.query('SELECT mime, chunks FROM blobs WHERE id = $1 AND owner_id = $2', [p.id, req.user.id])
      const blob = rows[0]
      if (!blob) return res.status(404).json({ error: 'Not found' })
      const have = (await db.query('SELECT COUNT(*)::int AS n FROM blob_chunks WHERE blob_id = $1', [p.id])).rows[0].n
      if (have !== blob.chunks) return res.status(400).json({ error: 'Upload is incomplete' })
      const first = toBuffer((await db.query('SELECT data FROM blob_chunks WHERE blob_id = $1 AND idx = 0', [p.id])).rows[0].data)
      if (!MAGIC[blob.mime].every((byte, i) => first[i] === byte)) {
        await db.query('DELETE FROM blobs WHERE id = $1', [p.id])
        return res.status(400).json({ error: 'File content does not match its type' })
      }
      await db.query('UPDATE blobs SET complete = TRUE WHERE id = $1', [p.id])
      res.json({ id: p.id, url: blob.mime.startsWith('image/') ? `/api/media/${p.id}` : null })
    }),
  )

  // Images: addressed by an unguessable 128-bit id (capability URL) and only served to signed-in sessions.
  app.get(
    '/api/media/:id',
    requireUser,
    ah(async (req, res) => {
      const p = check(params.blob, req.params, res)
      if (!p) return
      const { rows } = await db.query(
        "SELECT b.mime, c.data FROM blobs b JOIN blob_chunks c ON c.blob_id = b.id WHERE b.id = $1 AND b.kind = 'media' AND b.complete = TRUE AND c.idx = 0",
        [p.id],
      )
      if (!rows.length) return res.status(404).json({ error: 'Not found' })
      res.setHeader('Content-Type', rows[0].mime)
      res.setHeader('Content-Disposition', 'inline')
      res.setHeader('Cache-Control', 'private, max-age=31536000, immutable')
      res.send(toBuffer(rows[0].data))
    }),
  )

  // ---------------------------------------------------------------- sharing
  app.get(
    '/api/docs/:id/shares',
    requireUser,
    docAccess('view'),
    ah(async (req, res) => {
      const { rows } = await db.query('SELECT u.id, u.email, u.name, s.role FROM shares s JOIN users u ON u.id = s.user_id WHERE s.doc_id = $1 ORDER BY u.email', [req.params.id])
      res.json({ shares: rows })
    }),
  )
  app.post(
    '/api/docs/:id/shares',
    requireUser,
    shareLimit,
    docAccess('owner'),
    ah(async (req, res) => {
      const b = check(schemas.share, req.body, res)
      if (!b) return
      const { rows } = await db.query('SELECT id FROM users WHERE email = $1', [b.email])
      const u = rows[0]
      if (!u || isGuestEmail(b.email)) return res.status(404).json({ error: 'No SCRIBE account uses that email.' })
      if (u.id === req.user.id) return res.status(400).json({ error: 'You already own this document.' })
      await db.query('INSERT INTO shares (doc_id, user_id, role) VALUES ($1, $2, $3) ON CONFLICT (doc_id, user_id) DO UPDATE SET role = EXCLUDED.role', [req.params.id, u.id, b.role])
      res.json({ ok: true })
    }),
  )
  app.delete(
    '/api/docs/:id/shares/:userId',
    requireUser,
    docAccess('owner'),
    ah(async (req, res) => {
      const p = check(params.docShare, req.params, res)
      if (!p) return
      await db.query('DELETE FROM shares WHERE doc_id = $1 AND user_id = $2', [p.id, p.userId])
      await db.query('DELETE FROM presence WHERE doc_id = $1 AND user_id = $2', [p.id, p.userId])
      res.json({ ok: true })
    }),
  )

  // ---------------------------------------------------------------- versions
  app.get(
    '/api/docs/:id/versions',
    requireUser,
    docAccess('view'),
    ah(async (req, res) => {
      const { rows } = await db.query(
        'SELECT v.id, v.created, v.label, v.auto, u.name AS author FROM versions v LEFT JOIN users u ON u.id = v.author_id WHERE v.doc_id = $1 ORDER BY v.created DESC LIMIT 200',
        [req.params.id],
      )
      res.json({ versions: rows })
    }),
  )
  app.post(
    '/api/docs/:id/versions',
    requireUser,
    docAccess('edit'),
    ah(async (req, res) => {
      const b = check(schemas.version, req.body, res)
      if (!b) return
      const content = sanitizeContent(b.content)
      if (!content) return res.status(400).json({ error: 'Invalid request' })
      const json = JSON.stringify(content)
      if (json.length > MAX_JSON_CHARS) return res.status(413).json({ error: 'Document is too large' })
      if (b.auto) {
        const last = await db.query('SELECT content_json FROM versions WHERE doc_id = $1 ORDER BY created DESC LIMIT 1', [req.params.id])
        if (last.rows[0]?.content_json === json) return res.json({ skipped: true })
      }
      const id = newId()
      await db.query('INSERT INTO versions (id, doc_id, created, author_id, label, auto, content_json, meta_json) VALUES ($1, $2, $3, $4, $5, $6, $7, $8)', [
        id, req.params.id, Date.now(), req.user.id, b.label || null, b.auto, json, JSON.stringify(sanitizeMeta(b.meta)),
      ])
      await db.query(
        'DELETE FROM versions WHERE doc_id = $1 AND auto = TRUE AND id NOT IN (SELECT id FROM versions WHERE doc_id = $1 AND auto = TRUE ORDER BY created DESC LIMIT 100)',
        [req.params.id],
      )
      res.status(201).json({ id })
    }),
  )
  app.get(
    '/api/docs/:id/versions/:vid',
    requireUser,
    docAccess('view'),
    ah(async (req, res) => {
      const p = check(params.docVersion, req.params, res)
      if (!p) return
      const { rows } = await db.query('SELECT id, created, label, content_json, meta_json FROM versions WHERE id = $1 AND doc_id = $2', [p.vid, p.id])
      if (!rows.length) return res.status(404).json({ error: 'Version not found' })
      const v = rows[0]
      res.json({ id: v.id, created: v.created, label: v.label, content: JSON.parse(v.content_json), meta: JSON.parse(v.meta_json) })
    }),
  )

  // ---------------------------------------------------------------- AI
  app.get('/api/ai/status', requireUser, (req, res) => {
    const c = aiConfig()
    res.json({ provider: c.provider, model: c.model ?? null, reason: c.reason ?? null })
  })
  app.post(
    '/api/ai',
    requireUser,
    aiLimit,
    ah(async (req, res) => {
      const b = check(schemas.ai, req.body, res)
      if (!b) return
      try {
        res.json(await runAi(b))
      } catch (e) {
        if (e instanceof AiError) return res.status(e.status).json({ error: e.message })
        throw e
      }
    }),
  )

  app.get('/api/health', (req, res) => res.json({ ok: true }))
  app.use('/api', (req, res) => res.status(404).json({ error: 'Not found' }))

  // ---- static client (local production server only; on Vercel the CDN serves it)
  const dist = path.join(here, '..', 'dist')
  if (opts.serveStatic && fs.existsSync(dist)) {
    app.use(express.static(dist, { index: false, maxAge: '1h' }))
    // a missing hashed asset is a real 404, not the app page (matches vercel.json)
    app.get(/^\/assets\//, (req, res) => res.status(404).type('text/plain').send('Not found'))
    app.get(/^(?!\/api).*/, (req, res) => res.sendFile(path.join(dist, 'index.html')))
  }

  // Generic errors only: no stack traces or internal messages reach the client.
  app.use((err, req, res, next) => {
    if (res.headersSent) return next(err)
    if (err?.type === 'entity.too.large') return res.status(413).json({ error: 'Request too large' })
    if (err?.type === 'entity.parse.failed') return res.status(400).json({ error: 'Invalid request' })
    if (process.env.NODE_ENV !== 'test') console.error('[scribe] unhandled error', err?.stack || err)
    res.status(500).json({ error: 'Something went wrong' })
  })

  return { app, db }
}
