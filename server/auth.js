import crypto from 'node:crypto'
import { promisify } from 'node:util'

const scrypt = promisify(crypto.scrypt)
const SCRYPT = { N: 16384, r: 8, p: 1, keylen: 64 }
export const SESSION_COOKIE = 'scribe_session'
export const SESSION_TTL_MS = 1000 * 60 * 60 * 24 * 14
/** Guest (no-login) sessions last a year so a browser keeps its documents, like a desktop app. */
export const GUEST_TTL_MS = 1000 * 60 * 60 * 24 * 365
export const GUEST_DOMAIN = '@guest.scribe.local'
export const isGuestEmail = (email) => String(email).endsWith(GUEST_DOMAIN)

export async function hashPassword(password) {
  const salt = crypto.randomBytes(16)
  const key = await scrypt(password, salt, SCRYPT.keylen, { N: SCRYPT.N, r: SCRYPT.r, p: SCRYPT.p })
  return `scrypt$${SCRYPT.N}$${SCRYPT.r}$${SCRYPT.p}$${salt.toString('base64')}$${key.toString('base64')}`
}

export async function verifyPassword(password, stored) {
  const [alg, N, r, p, salt, key] = String(stored).split('$')
  if (alg !== 'scrypt' || !salt || !key) return false
  const expected = Buffer.from(key, 'base64')
  const actual = await scrypt(password, Buffer.from(salt, 'base64'), expected.length, { N: Number(N), r: Number(r), p: Number(p) })
  return expected.length === actual.length && crypto.timingSafeEqual(expected, actual)
}

/** Dummy hash so unknown emails cost the same time as wrong passwords (no user enumeration by timing). */
export const DUMMY_HASH = 'scrypt$16384$8$1$AAAAAAAAAAAAAAAAAAAAAA==$' + Buffer.alloc(64).toString('base64')

const sha256 = (s) => crypto.createHash('sha256').update(s).digest('hex')
export const newId = (bytes = 12) => crypto.randomBytes(bytes).toString('base64url')

export async function createSession(db, userId, ttl = SESSION_TTL_MS) {
  const token = crypto.randomBytes(32).toString('base64url')
  await db.query('INSERT INTO sessions (token_hash, user_id, expires) VALUES ($1, $2, $3)', [sha256(token), userId, Date.now() + ttl])
  return token
}

export async function destroySession(db, token) {
  if (token) await db.query('DELETE FROM sessions WHERE token_hash = $1', [sha256(token)])
}

export async function userFromToken(db, token) {
  if (!token || typeof token !== 'string' || token.length > 100) return null
  const { rows } = await db.query(
    'SELECT u.id, u.email, u.name, s.expires FROM sessions s JOIN users u ON u.id = s.user_id WHERE s.token_hash = $1',
    [sha256(token)],
  )
  const row = rows[0]
  if (!row) return null
  if (row.expires < Date.now()) {
    await destroySession(db, token)
    return null
  }
  return { id: row.id, email: row.email, name: row.name }
}

export function parseCookies(header) {
  const out = {}
  for (const part of String(header || '').split(';')) {
    const i = part.indexOf('=')
    if (i <= 0) continue
    try {
      out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim())
    } catch {
      /* ignore malformed cookie values */
    }
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
