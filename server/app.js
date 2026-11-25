import express from 'express'
import http from 'node:http'
import path from 'node:path'
import fs from 'node:fs'
import { fileURLToPath } from 'node:url'
import { openDb, roleFor, canEdit, canComment } from './db.js'
import {
  GUEST_DOMAIN, GUEST_TTL_MS, isGuestEmail, SESSION_COOKIE, clearCookie, createSession, destroySession, hashPassword, newId, parseCookies, rateLimiter, sessionCookie,
  userFromToken, verifyPassword,
} from './auth.js'
import { createCollab } from './collab.js'
import { AiError, aiConfig, runAi } from './ai.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const MAX_ORIGINAL_BYTES = 100 * 1024 * 1024

export function createApp({ dbFile = path.join(here, '..', 'data', 'scribe.db'), secureCookies = process.env.NODE_ENV === 'production', serveStatic = process.env.NODE_ENV === 'production', log = console.log, authRateLimit = 10 } = {}) {
  const db = openDb(dbFile)
  const collab = createCollab(db, { log })
  const app = express()
  app.disable('x-powered-by')
  app.set('trust proxy', 'loopback')

  // ---- security headers
  app.use((req, res, next) => {
    res.setHeader('X-Content-Type-Options', 'nosniff')
    res.setHeader('Referrer-Policy', 'same-origin')
    res.setHeader('X-Frame-Options', 'DENY')
    res.setHeader('Cross-Origin-Opener-Policy', 'same-origin')
    res.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=()')
    res.setHeader(
      'Content-Security-Policy',
      "default-src 'self'; img-src 'self' data: blob: https:; style-src 'self' 'unsafe-inline'; font-src 'self' data:; script-src 'self'; connect-src 'self' ws: wss:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'",
    )
    next()
  })

  app.use(express.json({ limit: '150mb' }))

  // ---- auth context + CSRF guard
  app.use('/api', (req, res, next) => {
    req.user = userFromToken(db, parseCookies(req.headers.cookie)[SESSION_COOKIE])
    if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method)) {
      // Custom header can't be sent cross-site without CORS preflight, which we never grant.
      if (req.get('x-scribe-request') !== '1') return res.status(403).json({ error: 'Missing request header' })
      const origin = req.get('origin')
      if (origin && new URL(origin).host !== req.get('host')) return res.status(403).json({ error: 'Cross-origin request refused' })
    }
    next()
  })
  const requireUser = (req, res, next) => (req.user ? next() : res.status(401).json({ error: 'Sign in required' }))
  const loginLimit = rateLimiter(authRateLimit, 60_000)
  const aiLimit = rateLimiter(30, 60_000)

  const str = (v, max) => (typeof v === 'string' ? v.trim().slice(0, max) : '')

  // ---- auth routes
  app.post('/api/auth/register', (req, res) => {
    if (!loginLimit('reg:' + req.ip)) return res.status(429).json({ error: 'Too many attempts, wait a minute.' })
    const email = str(req.body?.email, 200).toLowerCase()
    const name = str(req.body?.name, 100) || email.split('@')[0]
    const password = typeof req.body?.password === 'string' ? req.body.password : ''
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || isGuestEmail(email)) return res.status(400).json({ error: 'Enter a valid email address.' })
    if (password.length < 8 || password.length > 200) return res.status(400).json({ error: 'Password must be 8–200 characters.' })
    if (db.prepare('SELECT 1 FROM users WHERE email = ?').get(email)) return res.status(409).json({ error: 'An account with this email already exists.' })
    const id = newId()
    db.prepare('INSERT INTO users (id, email, name, pass_hash, created) VALUES (?, ?, ?, ?, ?)').run(id, email, name, hashPassword(password), Date.now())
    res.setHeader('Set-Cookie', sessionCookie(createSession(db, id), secureCookies))
    res.status(201).json({ user: publicUser({ id, email, name }) })
  })

  app.post('/api/auth/login', (req, res) => {
    if (!loginLimit('login:' + req.ip)) return res.status(429).json({ error: 'Too many attempts, wait a minute.' })
    const email = str(req.body?.email, 200).toLowerCase()
    const password = typeof req.body?.password === 'string' ? req.body.password : ''
    const u = db.prepare('SELECT id, email, name, pass_hash FROM users WHERE email = ?').get(email)
    if (!u || !verifyPassword(password, u.pass_hash)) return res.status(401).json({ error: 'Email or password is incorrect.' })
    res.setHeader('Set-Cookie', sessionCookie(createSession(db, u.id), secureCookies))
    res.json({ user: publicUser({ id: u.id, email: u.email, name: u.name }) })
  })

  app.post('/api/auth/logout', (req, res) => {
    destroySession(db, parseCookies(req.headers.cookie)[SESSION_COOKIE])
    res.setHeader('Set-Cookie', clearCookie(secureCookies))
    res.json({ ok: true })
  })

  const publicUser = (u) => (u ? { ...u, guest: isGuestEmail(u.email) } : null)
  app.get('/api/auth/me', (req, res) => res.json({ user: publicUser(req.user) }))

  // No-login start: a private local identity per browser, like a desktop word processor.
  // Guests cannot sign in with a password; they can upgrade to a full account later.
  const guestLimit = rateLimiter(Math.max(authRateLimit, 20), 60_000)
  app.post('/api/auth/guest', (req, res) => {
    if (req.user) return res.json({ user: publicUser(req.user) })
    if (!guestLimit('guest:' + req.ip)) return res.status(429).json({ error: 'Too many new sessions, wait a minute.' })
    const id = newId()
    db.prepare('INSERT INTO users (id, email, name, pass_hash, created) VALUES (?, ?, ?, ?, ?)').run(id, `${id.toLowerCase()}${GUEST_DOMAIN}`, 'You', '!guest', Date.now())
    res.setHeader('Set-Cookie', sessionCookie(createSession(db, id, GUEST_TTL_MS), secureCookies, GUEST_TTL_MS))
    res.status(201).json({ user: publicUser({ id, email: `${id.toLowerCase()}${GUEST_DOMAIN}`, name: 'You' }) })
  })

  // Turn the current guest into a full account, keeping all of its documents.
  app.post('/api/auth/upgrade', requireUser, (req, res) => {
    if (!isGuestEmail(req.user.email)) return res.status(400).json({ error: 'This is already a full account.' })
    const email = str(req.body?.email, 200).toLowerCase()
    const name = str(req.body?.name, 100) || email.split('@')[0]
    const password = typeof req.body?.password === 'string' ? req.body.password : ''
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || isGuestEmail(email)) return res.status(400).json({ error: 'Enter a valid email address.' })
    if (password.length < 8 || password.length > 200) return res.status(400).json({ error: 'Password must be 8–200 characters.' })
    if (db.prepare('SELECT 1 FROM users WHERE email = ?').get(email)) return res.status(409).json({ error: 'An account with this email already exists. Sign in instead.' })
    db.prepare('UPDATE users SET email = ?, name = ?, pass_hash = ? WHERE id = ?').run(email, name, hashPassword(password), req.user.id)
    res.json({ user: publicUser({ id: req.user.id, email, name }) })
  })

  // ---- documents
  const docAccess = (need) => (req, res, next) => {
    const role = roleFor(db, req.params.id, req.user.id)
    if (!role) return res.status(404).json({ error: 'Document not found' })
    if (need === 'edit' && !canEdit(role)) return res.status(403).json({ error: 'You have view-only access' })
    if (need === 'comment' && !canComment(role)) return res.status(403).json({ error: 'You cannot comment on this document' })
    if (need === 'owner' && role !== 'owner') return res.status(403).json({ error: 'Only the owner can do this' })
    req.role = role
    next()
  }

  app.get('/api/docs', requireUser, (req, res) => {
    const rows = db
      .prepare(
        `SELECT d.id, d.title, d.updated, d.created, u.name AS owner_name, CASE WHEN d.owner_id = ? THEN 'owner' ELSE s.role END AS role,
                d.original IS NOT NULL AS has_original
         FROM documents d JOIN users u ON u.id = d.owner_id LEFT JOIN shares s ON s.doc_id = d.id AND s.user_id = ?
         WHERE d.deleted = 0 AND (d.owner_id = ? OR s.user_id IS NOT NULL)
         ORDER BY d.updated DESC LIMIT 500`,
      )
      .all(req.user.id, req.user.id, req.user.id)
    res.json({ docs: rows.map((r) => ({ ...r, has_original: !!r.has_original })) })
  })

  app.post('/api/docs', requireUser, (req, res) => {
    const title = str(req.body?.title, 200) || 'Untitled document'
    const content = req.body?.content
    const meta = req.body?.meta
    if (!content || content.type !== 'doc') return res.status(400).json({ error: 'Invalid document content' })
    let ystate = null
    let original = null
    try {
      if (req.body?.yUpdate) ystate = Buffer.from(String(req.body.yUpdate), 'base64')
      if (req.body?.original) {
        original = Buffer.from(String(req.body.original), 'base64')
        if (original.length > MAX_ORIGINAL_BYTES) return res.status(413).json({ error: 'File too large' })
        if (original[0] !== 0x50 || original[1] !== 0x4b) return res.status(400).json({ error: 'Original file is not a DOCX package' })
      }
    } catch {
      return res.status(400).json({ error: 'Invalid encoded payload' })
    }
    const id = newId()
    const now = Date.now()
    db.prepare(
      'INSERT INTO documents (id, owner_id, title, created, updated, content_json, meta_json, ystate, original, original_name) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
    ).run(id, req.user.id, title, now, now, JSON.stringify(content), JSON.stringify(meta ?? {}), ystate, original, str(req.body?.originalName, 255) || null)
    res.status(201).json({ id })
  })

  app.get('/api/docs/:id', requireUser, docAccess('view'), (req, res) => {
    const d = db.prepare('SELECT id, title, created, updated, meta_json, content_json, original IS NOT NULL AS has_original, original_name FROM documents WHERE id = ?').get(req.params.id)
    res.json({
      id: d.id, title: d.title, created: d.created, updated: d.updated, role: req.role,
      meta: JSON.parse(d.meta_json), content: JSON.parse(d.content_json), hasOriginal: !!d.has_original, originalName: d.original_name,
    })
  })

  app.patch('/api/docs/:id', requireUser, docAccess('edit'), (req, res) => {
    const sets = []
    const vals = []
    if (req.body?.title !== undefined) {
      sets.push('title = ?')
      vals.push(str(req.body.title, 200) || 'Untitled document')
    }
    if (req.body?.content !== undefined) {
      if (req.body.content?.type !== 'doc') return res.status(400).json({ error: 'Invalid content' })
      sets.push('content_json = ?')
      vals.push(JSON.stringify(req.body.content))
    }
    if (req.body?.meta !== undefined) {
      sets.push('meta_json = ?')
      vals.push(JSON.stringify(req.body.meta))
    }
    if (!sets.length) return res.json({ ok: true })
    sets.push('updated = ?')
    vals.push(Date.now())
    db.prepare(`UPDATE documents SET ${sets.join(', ')} WHERE id = ?`).run(...vals, req.params.id)
    res.json({ ok: true })
  })

  app.delete('/api/docs/:id', requireUser, docAccess('owner'), (req, res) => {
    db.prepare('UPDATE documents SET deleted = 1, updated = ? WHERE id = ?').run(Date.now(), req.params.id)
    res.json({ ok: true })
  })

  app.get('/api/docs/:id/original', requireUser, docAccess('view'), (req, res) => {
    const d = db.prepare('SELECT original, original_name FROM documents WHERE id = ?').get(req.params.id)
    if (!d?.original) return res.status(404).json({ error: 'No original file' })
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    res.setHeader('Content-Disposition', `attachment; filename="original.docx"`)
    res.send(Buffer.from(d.original))
  })

  // ---- sharing (server-side RBAC)
  app.get('/api/docs/:id/shares', requireUser, docAccess('view'), (req, res) => {
    const rows = db.prepare('SELECT u.id, u.email, u.name, s.role FROM shares s JOIN users u ON u.id = s.user_id WHERE s.doc_id = ? ORDER BY u.email').all(req.params.id)
    res.json({ shares: rows })
  })
  app.post('/api/docs/:id/shares', requireUser, docAccess('owner'), (req, res) => {
    const email = str(req.body?.email, 200).toLowerCase()
    const role = req.body?.role
    if (!['viewer', 'commenter', 'editor'].includes(role)) return res.status(400).json({ error: 'Invalid role' })
    const u = db.prepare('SELECT id FROM users WHERE email = ?').get(email)
    if (!u) return res.status(404).json({ error: 'No SCRIBE account uses that email.' })
    if (u.id === req.user.id) return res.status(400).json({ error: 'You already own this document.' })
    db.prepare('INSERT INTO shares (doc_id, user_id, role) VALUES (?, ?, ?) ON CONFLICT(doc_id, user_id) DO UPDATE SET role = excluded.role').run(req.params.id, u.id, role)
    collab.kick(req.params.id, u.id) // reconnect with the new role
    res.json({ ok: true })
  })
  app.delete('/api/docs/:id/shares/:userId', requireUser, docAccess('owner'), (req, res) => {
    db.prepare('DELETE FROM shares WHERE doc_id = ? AND user_id = ?').run(req.params.id, req.params.userId)
    collab.kick(req.params.id, req.params.userId)
    res.json({ ok: true })
  })

  // ---- versions
  app.get('/api/docs/:id/versions', requireUser, docAccess('view'), (req, res) => {
    const rows = db
      .prepare('SELECT v.id, v.created, v.label, v.auto, u.name AS author FROM versions v LEFT JOIN users u ON u.id = v.author_id WHERE v.doc_id = ? ORDER BY v.created DESC LIMIT 200')
      .all(req.params.id)
    res.json({ versions: rows.map((r) => ({ ...r, auto: !!r.auto })) })
  })
  app.post('/api/docs/:id/versions', requireUser, docAccess('edit'), (req, res) => {
    const content = req.body?.content
    if (content?.type !== 'doc') return res.status(400).json({ error: 'Invalid content' })
    const auto = !!req.body?.auto
    const last = db.prepare('SELECT content_json FROM versions WHERE doc_id = ? ORDER BY created DESC LIMIT 1').get(req.params.id)
    const json = JSON.stringify(content)
    if (auto && last?.content_json === json) return res.json({ skipped: true })
    const id = newId()
    db.prepare('INSERT INTO versions (id, doc_id, created, author_id, label, auto, content_json, meta_json) VALUES (?, ?, ?, ?, ?, ?, ?, ?)').run(
      id, req.params.id, Date.now(), req.user.id, str(req.body?.label, 120) || null, auto ? 1 : 0, json, JSON.stringify(req.body?.meta ?? {}),
    )
    // keep at most 100 automatic versions per document
    db.prepare('DELETE FROM versions WHERE doc_id = ? AND auto = 1 AND id NOT IN (SELECT id FROM versions WHERE doc_id = ? AND auto = 1 ORDER BY created DESC LIMIT 100)').run(req.params.id, req.params.id)
    res.status(201).json({ id })
  })
  app.get('/api/docs/:id/versions/:vid', requireUser, docAccess('view'), (req, res) => {
    const v = db.prepare('SELECT id, created, label, content_json, meta_json FROM versions WHERE id = ? AND doc_id = ?').get(req.params.vid, req.params.id)
    if (!v) return res.status(404).json({ error: 'Version not found' })
    res.json({ id: v.id, created: v.created, label: v.label, content: JSON.parse(v.content_json), meta: JSON.parse(v.meta_json) })
  })

  // ---- AI
  app.get('/api/ai/status', requireUser, (req, res) => {
    const c = aiConfig()
    res.json({ provider: c.provider, model: c.model ?? null, reason: c.reason ?? null })
  })
  app.post('/api/ai', requireUser, async (req, res) => {
    if (!aiLimit('ai:' + req.user.id)) return res.status(429).json({ error: 'AI rate limit reached; wait a minute.' })
    try {
      const out = await runAi({ task: req.body?.task, text: req.body?.text, instruction: req.body?.instruction, context: req.body?.context })
      res.json(out)
    } catch (e) {
      res.status(e instanceof AiError ? e.status : 500).json({ error: e instanceof AiError ? e.message : 'AI request failed' })
    }
  })

  app.get('/api/health', (req, res) => res.json({ ok: true }))
  app.use('/api', (req, res) => res.status(404).json({ error: 'Not found' }))

  // ---- static client (production)
  const dist = path.join(here, '..', 'dist')
  if (serveStatic && fs.existsSync(dist)) {
    app.use(express.static(dist, { index: false, maxAge: '1h' }))
    app.get(/^(?!\/api|\/collab).*/, (req, res) => res.sendFile(path.join(dist, 'index.html')))
  }

  // eslint-disable-next-line no-unused-vars
  app.use((err, req, res, next) => {
    if (err?.type === 'entity.too.large') return res.status(413).json({ error: 'Request too large' })
    if (err?.type === 'entity.parse.failed') return res.status(400).json({ error: 'Malformed JSON' })
    log('server error', err)
    res.status(500).json({ error: 'Internal server error' })
  })

  const server = http.createServer(app)
  server.on('upgrade', (req, socket, head) => collab.handleUpgrade(req, socket, head))
  return { app, server, db, collab }
}
