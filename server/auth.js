import crypto from 'node:crypto'

const SCRYPT = { N: 16384, r: 8, p: 1, keylen: 64 }
export const SESSION_COOKIE = 'scribe_session'
const SESSION_TTL_MS = 1000 * 60 * 60 * 24 * 14
/** Guest (no-login) sessions last a year so a browser keeps its documents, like a desktop app. */
export const GUEST_TTL_MS = 1000 * 60 * 60 * 24 * 365
export const GUEST_DOMAIN = '@guest.scribe.local'
export const isGuestEmail = (email) => String(email).endsWith(GUEST_DOMAIN)

export function hashPassword(password) {
  const salt = crypto.randomBytes(16)
  const key = crypto.scryptSync(password, salt, SCRYPT.keylen, { N: SCRYPT.N, r: SCRYPT.r, p: SCRYPT.p })
  return `scrypt$${SCRYPT.N}$${SCRYPT.r}$${SCRYPT.p}$${salt.toString('base64')}$${key.toString('base64')}`
}

export function verifyPassword(password, stored) {
  const [alg, N, r, p, salt, key] = String(stored).split('$')
  if (alg !== 'scrypt') return false
  const expected = Buffer.from(key, 'base64')
  const actual = crypto.scryptSync(password, Buffer.from(salt, 'base64'), expected.length, { N: Number(N), r: Number(r), p: Number(p) })
  return expected.length === actual.length && crypto.timingSafeEqual(expected, actual)
}

const sha256 = (s) => crypto.createHash('sha256').update(s).digest('hex')
export const newId = (bytes = 12) => crypto.randomBytes(bytes).toString('base64url')

export function createSession(db, userId, ttl = SESSION_TTL_MS) {
  const token = crypto.randomBytes(32).toString('base64url')
  db.prepare('INSERT INTO sessions (token_hash, user_id, expires) VALUES (?, ?, ?)').run(sha256(token), userId, Date.now() + ttl)
  return token
}

export function destroySession(db, token) {
  if (token) db.prepare('DELETE FROM sessions WHERE token_hash = ?').run(sha256(token))
}

export function userFromToken(db, token) {
  if (!token) return null
  const row = db
    .prepare('SELECT u.id, u.email, u.name, s.expires FROM sessions s JOIN users u ON u.id = s.user_id WHERE s.token_hash = ?')
    .get(sha256(token))
  if (!row) return null
  if (row.expires < Date.now()) {
    destroySession(db, token)
    return null
  }
  return { id: row.id, email: row.email, name: row.name }
}

export function parseCookies(header) {
  const out = {}
  for (const part of String(header || '').split(';')) {
    const i = part.indexOf('=')
    if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim())
  }
  return out
}

export function sessionCookie(token, secure, ttl = SESSION_TTL_MS) {
  const attrs = [`${SESSION_COOKIE}=${token}`, 'Path=/', 'HttpOnly', 'SameSite=Strict', `Max-Age=${Math.floor(ttl / 1000)}`]
  if (secure) attrs.push('Secure')
  return attrs.join('; ')
}

export function clearCookie(secure) {
  return `${SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0${secure ? '; Secure' : ''}`
}

/** Fixed-window rate limiter keyed by IP + bucket (login, register, ai). */
export function rateLimiter(limit, windowMs) {
  const hits = new Map()
  return (key) => {
    const now = Date.now()
    const h = hits.get(key)
    if (!h || h.reset < now) {
      hits.set(key, { n: 1, reset: now + windowMs })
      return true
    }
    h.n++
    if (hits.size > 10000) for (const [k, v] of hits) if (v.reset < now) hits.delete(k)
    return h.n <= limit
  }
}
