// @vitest-environment node
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import http from 'node:http'
import * as Y from 'yjs'
import { createApp, type Db } from '../../server/app.js'

let base = ''
let server: http.Server
let db: Db

beforeAll(async () => {
  const ctx = await createApp({ dbDir: ':memory:', authRateLimit: 1000, secureCookies: false })
  db = ctx.db
  server = http.createServer(ctx.app)
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', () => r()))
  base = `http://127.0.0.1:${(server.address() as any).port}`
}, 60_000)
afterAll(async () => {
  server.close()
  await db.close()
})

async function api(path: string, opts: { method?: string; body?: any; cookie?: string; csrf?: boolean; origin?: string } = {}) {
  const res = await fetch(base + path, {
    method: opts.method ?? 'GET',
    headers: {
      'content-type': 'application/json',
      ...(opts.cookie ? { cookie: opts.cookie } : {}),
      ...(opts.csrf === false ? {} : { 'x-scribe-request': '1' }),
      ...(opts.origin ? { origin: opts.origin } : {}),
    },
    body: opts.body ? JSON.stringify(opts.body) : undefined,
  })
  const cookie = res.headers.get('set-cookie')?.split(';')[0]
  return { status: res.status, json: await res.json().catch(() => null), cookie, headers: res.headers }
}

async function user(email: string) {
  const r = await api('/api/auth/register', { method: 'POST', body: { email, password: 'correct horse battery', name: email.split('@')[0] } })
  expect(r.status).toBe(201)
  return r.cookie!
}

const doc = { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Hello' }] }] }
const b64 = (u: Uint8Array) => Buffer.from(u).toString('base64')

describe('auth', () => {
  it('rate-limits repeated login attempts (database-backed, works across instances)', async () => {
    const limited = await createApp({ dbDir: ':memory:', authRateLimit: 3, secureCookies: false })
    const s = http.createServer(limited.app)
    await new Promise<void>((r) => s.listen(0, '127.0.0.1', () => r()))
    const url = `http://127.0.0.1:${(s.address() as any).port}/api/auth/login`
    const codes: number[] = []
    for (let i = 0; i < 5; i++) codes.push((await fetch(url, { method: 'POST', headers: { 'content-type': 'application/json', 'x-scribe-request': '1' }, body: JSON.stringify({ email: 'a@b.c', password: 'x' }) })).status)
    s.close()
    await limited.db.close()
    expect(codes).toEqual([401, 401, 401, 429, 429])
  }, 60_000)
  it('registers, logs in, rejects bad passwords, duplicates and invalid input', async () => {
    const c = await user('ada@example.com')
    expect((await api('/api/auth/me', { cookie: c })).json.user.email).toBe('ada@example.com')
    expect((await api('/api/auth/register', { method: 'POST', body: { email: 'ada@example.com', password: 'another password' } })).status).toBe(409)
    expect((await api('/api/auth/login', { method: 'POST', body: { email: 'ada@example.com', password: 'wrong password' } })).status).toBe(401)
    expect((await api('/api/auth/login', { method: 'POST', body: { email: 'ADA@example.com', password: 'correct horse battery' } })).status).toBe(200)
    const bad = await api('/api/auth/register', { method: 'POST', body: { email: 'shortpw@example.com', password: 'short' } })
    expect(bad.status).toBe(400)
    expect(bad.json.error).toMatch(/Password/)
    expect((await api('/api/auth/register', { method: 'POST', body: { email: 'not-an-email', password: 'longenough1' } })).status).toBe(400)
    expect((await api('/api/auth/register', { method: 'POST', body: { email: { $ne: 1 }, password: 'longenough1' } })).status).toBe(400)
  })
  it('sets an HttpOnly SameSite=Strict session cookie', async () => {
    const res = await fetch(base + '/api/auth/register', { method: 'POST', headers: { 'content-type': 'application/json', 'x-scribe-request': '1' }, body: JSON.stringify({ email: 'cookie@example.com', password: 'longenough1' }) })
    const sc = res.headers.get('set-cookie')!
    expect(sc).toMatch(/HttpOnly/)
    expect(sc).toMatch(/SameSite=Strict/)
  })
  it('refuses requests without the CSRF header or from other origins (strict CORS)', async () => {
    const c = await user('csrf@example.com')
    expect((await api('/api/docs', { method: 'POST', cookie: c, csrf: false, body: { title: 'x', content: doc } })).status).toBe(403)
    const cross = await api('/api/docs', { method: 'POST', cookie: c, origin: 'https://evil.example', body: { title: 'x', content: doc } })
    expect(cross.status).toBe(403)
    expect(cross.headers.get('access-control-allow-origin')).toBeNull()
    expect((await api('/api/docs', { cookie: c, origin: 'https://evil.example' })).status).toBe(403)
  })
  it('sends security headers and generic errors', async () => {
    const res = await api('/api/health')
    expect(res.headers.get('content-security-policy')).toContain("default-src 'self'")
    expect(res.headers.get('strict-transport-security')).toContain('max-age=')
    expect(res.headers.get('x-content-type-options')).toBe('nosniff')
    expect(res.headers.get('x-frame-options')).toBe('DENY')
    expect(res.headers.get('referrer-policy')).toBe('strict-origin-when-cross-origin')
    expect(res.headers.get('x-powered-by')).toBeNull()
    const malformed = await fetch(base + '/api/auth/login', { method: 'POST', headers: { 'content-type': 'application/json', 'x-scribe-request': '1' }, body: '{bad json' })
    const text = await malformed.text()
    expect(malformed.status).toBe(400)
    expect(text).not.toMatch(/at .*\.js|SyntaxError|stack/i)
  })
  it('requires a session on every document route', async () => {
    for (const [m, p] of [['GET', '/api/docs'], ['POST', '/api/docs'], ['GET', '/api/docs/abcdefgh'], ['POST', '/api/docs/abcdefgh/sync'], ['POST', '/api/blobs'], ['GET', '/api/media/abcdefghijkl'], ['POST', '/api/ai'], ['GET', '/api/ai/status']])
      expect((await api(p, { method: m, body: m === 'GET' ? undefined : {} })).status, `${m} ${p}`).toBe(401)
  })
})

describe('no-login guest sessions', () => {
  it('creates a private guest identity that cannot be logged into, and upgrades it keeping documents', async () => {
    const g = await api('/api/auth/guest', { method: 'POST' })
    expect(g.status).toBe(201)
    expect(g.json.user.guest).toBe(true)
    const cookie = g.cookie!
    expect((await api('/api/auth/guest', { method: 'POST', cookie })).json.user.id).toBe(g.json.user.id)
    const d = await api('/api/docs', { method: 'POST', cookie, body: { title: 'Guest doc', content: doc } })
    expect(d.status).toBe(201)
    const row = await db.query('SELECT email FROM users WHERE id = $1', [g.json.user.id])
    expect((await api('/api/auth/login', { method: 'POST', body: { email: row.rows[0].email, password: '!guest' } })).status).toBe(401)
    expect((await api('/api/auth/register', { method: 'POST', body: { email: 'x@guest.scribe.local', password: 'longenough1' } })).status).toBe(400)
    const other = (await api('/api/auth/guest', { method: 'POST' })).cookie!
    expect((await api(`/api/docs/${d.json.id}`, { cookie: other })).status).toBe(404)
    const up = await api('/api/auth/upgrade', { method: 'POST', cookie, body: { email: 'upgraded@example.com', password: 'correct horse battery', name: 'Up' } })
    expect(up.json.user).toMatchObject({ email: 'upgraded@example.com', guest: false })
    const login = await api('/api/auth/login', { method: 'POST', body: { email: 'upgraded@example.com', password: 'correct horse battery' } })
    expect((await api('/api/docs', { cookie: login.cookie })).json.docs.map((x: any) => x.title)).toContain('Guest doc')
  })
})

describe('documents and authorization', () => {
  it('enforces owner / editor / viewer / stranger permissions server-side', async () => {
    const owner = await user('owner@example.com')
    const editor = await user('editor@example.com')
    const viewer = await user('viewer@example.com')
    const stranger = await user('stranger@example.com')
    const id = (await api('/api/docs', { method: 'POST', cookie: owner, body: { title: 'Plan', content: doc, meta: {} } })).json.id
    expect((await api(`/api/docs/${id}`, { cookie: stranger })).status).toBe(404)
    expect((await api(`/api/docs/${id}`, { method: 'PATCH', cookie: stranger, body: { title: 'hijack' } })).status).toBe(404)
    await api(`/api/docs/${id}/shares`, { method: 'POST', cookie: owner, body: { email: 'editor@example.com', role: 'editor' } })
    await api(`/api/docs/${id}/shares`, { method: 'POST', cookie: owner, body: { email: 'viewer@example.com', role: 'viewer' } })
    expect((await api(`/api/docs/${id}/shares`, { method: 'POST', cookie: editor, body: { email: 'stranger@example.com', role: 'editor' } })).status).toBe(403)
    expect((await api(`/api/docs/${id}`, { cookie: viewer })).json.role).toBe('viewer')
    expect((await api(`/api/docs/${id}`, { method: 'PATCH', cookie: viewer, body: { title: 'nope' } })).status).toBe(403)
    expect((await api(`/api/docs/${id}`, { method: 'PATCH', cookie: editor, body: { title: 'Plan v2' } })).status).toBe(200)
    expect((await api(`/api/docs/${id}`, { method: 'PATCH', cookie: editor, body: { owner_id: 'x' } })).status).toBe(400)
    expect((await api(`/api/docs/${id}`, { method: 'DELETE', cookie: editor })).status).toBe(403)
    expect((await api('/api/docs', { cookie: viewer })).json.docs.map((d: any) => d.title)).toContain('Plan v2')
    expect((await api(`/api/docs/${id}`, { method: 'DELETE', cookie: owner })).status).toBe(200)
    expect((await api(`/api/docs/${id}`, { cookie: owner })).status).toBe(404)
  })

  it('rejects invalid ids and payloads; strips unsafe links and images server-side', async () => {
    const c = await user('payload@example.com')
    expect((await api("/api/docs/'%20OR%201=1--", { cookie: c })).status).toBe(400)
    expect((await api('/api/docs', { method: 'POST', cookie: c, body: { title: 'x', content: { type: 'html' } } })).status).toBe(400)
    const evil = {
      type: 'doc',
      content: [
        { type: 'paragraph', content: [{ type: 'text', text: 'x', marks: [{ type: 'link', attrs: { href: 'javascript:alert(1)' } }] }] },
        { type: 'paragraph', content: [{ type: 'image', attrs: { src: 'javascript:alert(1)' } }, { type: 'script', text: 'x' }] },
      ],
    }
    const id = (await api('/api/docs', { method: 'POST', cookie: c, body: { title: 'x', content: evil } })).json.id
    const got = JSON.stringify((await api(`/api/docs/${id}`, { cookie: c })).json.content)
    expect(got).not.toContain('javascript:')
    expect(got).not.toContain('script')
  })

  it('creates, lists and restores versions; dedupes identical auto versions', async () => {
    const c = await user('versions@example.com')
    const id = (await api('/api/docs', { method: 'POST', cookie: c, body: { title: 'v', content: doc } })).json.id
    const v1 = await api(`/api/docs/${id}/versions`, { method: 'POST', cookie: c, body: { content: doc, meta: {}, label: 'First' } })
    expect(v1.status).toBe(201)
    expect((await api(`/api/docs/${id}/versions`, { method: 'POST', cookie: c, body: { content: doc, meta: {}, auto: true } })).json.skipped).toBe(true)
    expect((await api(`/api/docs/${id}/versions`, { cookie: c })).json.versions).toHaveLength(1)
    expect((await api(`/api/docs/${id}/versions/${v1.json.id}`, { cookie: c })).json.content).toEqual(doc)
  })

  it('reports AI as unconfigured instead of faking output', async () => {
    const prev = process.env.ANTHROPIC_API_KEY
    delete process.env.ANTHROPIC_API_KEY
    const c = await user('ai@example.com')
    expect((await api('/api/ai/status', { cookie: c })).json.provider).toBe('none')
    expect((await api('/api/ai', { method: 'POST', cookie: c, body: { task: 'improve', text: 'hello' } })).status).toBe(503)
    expect((await api('/api/ai', { method: 'POST', cookie: c, body: { task: 'rm -rf', text: 'hello' } })).status).toBe(400)
    if (prev) process.env.ANTHROPIC_API_KEY = prev
  })
})

describe('chunked uploads', () => {
  const put = (id: string, idx: number, body: Uint8Array, cookie: string) =>
    fetch(`${base}/api/blobs/${id}/chunks/${idx}`, { method: 'PUT', headers: { 'content-type': 'application/octet-stream', 'x-scribe-request': '1', cookie }, body: body as BodyInit })

  it('stores images, serves them only to signed-in users, and rejects mismatched content', async () => {
    const c = await user('media@example.com')
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3])
    const b = await api('/api/blobs', { method: 'POST', cookie: c, body: { kind: 'media', mime: 'image/png', size: png.length } })
    expect((await put(b.json.id, 0, png, c)).status).toBe(200)
    const done = await api(`/api/blobs/${b.json.id}/complete`, { method: 'POST', cookie: c, body: {} })
    expect(done.json.url).toBe(`/api/media/${b.json.id}`)
    const img = await fetch(base + done.json.url, { headers: { cookie: c } })
    expect(img.headers.get('content-type')).toBe('image/png')
    expect(new Uint8Array(await img.arrayBuffer())).toEqual(png)
    expect((await fetch(base + done.json.url)).status).toBe(401)
    // HTML pretending to be a PNG is refused
    const fake = new TextEncoder().encode('<script>alert(1)</script>')
    const f = await api('/api/blobs', { method: 'POST', cookie: c, body: { kind: 'media', mime: 'image/png', size: fake.length } })
    await put(f.json.id, 0, fake, c)
    expect((await api(`/api/blobs/${f.json.id}/complete`, { method: 'POST', cookie: c, body: {} })).status).toBe(400)
    expect((await api('/api/blobs', { method: 'POST', cookie: c, body: { kind: 'media', mime: 'text/html', size: 10 } })).status).toBe(400)
    expect((await api('/api/blobs', { method: 'POST', cookie: c, body: { kind: 'media', mime: 'image/png', size: 5_000_000 } })).status).toBe(413)
  })

  it('stores a multi-chunk original DOCX and serves it back in parts to people with access', async () => {
    const c = await user('orig@example.com')
    const stranger = await user('orig-stranger@example.com')
    const bytes = new Uint8Array(3_000_000 + 1234)
    bytes.set([0x50, 0x4b, 0x03, 0x04])
    for (let i = 4; i < bytes.length; i += 4096) bytes[i] = i % 251
    const b = await api('/api/blobs', { method: 'POST', cookie: c, body: { kind: 'original', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', size: bytes.length } })
    expect(b.json.chunks).toBe(2)
    expect((await put(b.json.id, 0, bytes.subarray(0, 3_000_000), c)).status).toBe(200)
    expect((await put(b.json.id, 1, bytes.subarray(3_000_000), c)).status).toBe(200)
    expect((await api(`/api/blobs/${b.json.id}/complete`, { method: 'POST', cookie: c, body: {} })).status).toBe(200)
    // another user cannot attach someone else's upload
    expect((await api('/api/docs', { method: 'POST', cookie: stranger, body: { title: 'x', content: doc, originalBlob: b.json.id } })).status).toBe(400)
    const id = (await api('/api/docs', { method: 'POST', cookie: c, body: { title: 'x', content: doc, originalBlob: b.json.id, originalName: 'a.docx' } })).json.id
    const info = (await api(`/api/docs/${id}`, { cookie: c })).json
    expect(info.originalParts).toBe(2)
    const parts = await Promise.all([0, 1].map((i) => fetch(`${base}/api/docs/${id}/original/${i}`, { headers: { cookie: c } }).then((r) => r.arrayBuffer())))
    const joined = new Uint8Array(parts.reduce((n, p) => n + p.byteLength, 0))
    joined.set(new Uint8Array(parts[0]))
    joined.set(new Uint8Array(parts[1]), parts[0].byteLength)
    expect(joined).toEqual(bytes)
    expect((await fetch(`${base}/api/docs/${id}/original/0`, { headers: { cookie: stranger } })).status).toBe(404)
  }, 60_000)
})

describe('real-time collaboration over HTTP', () => {
  async function client(id: string, cookie: string) {
    const ydoc = new Y.Doc()
    let after = 0
    const sync = async (updates: Uint8Array[] = []) => {
      const r = await api(`/api/docs/${id}/sync`, { method: 'POST', cookie, body: { after, updates: updates.map(b64) } })
      if (r.status !== 200) return r.status
      for (const u of r.json.updates) Y.applyUpdate(ydoc, new Uint8Array(Buffer.from(u.data, 'base64')), 'remote')
      after = r.json.last
      return 200
    }
    await sync()
    return { ydoc, sync }
  }

  it('merges concurrent edits, blocks viewer writes, rejects strangers and invalid updates', async () => {
    const a = await user('collab-a@example.com')
    const bCookie = await user('collab-b@example.com')
    const v = await user('collab-v@example.com')
    const s = await user('collab-s@example.com')
    const id = (await api('/api/docs', { method: 'POST', cookie: a, body: { title: 'c', content: doc } })).json.id
    await api(`/api/docs/${id}/shares`, { method: 'POST', cookie: a, body: { email: 'collab-b@example.com', role: 'editor' } })
    await api(`/api/docs/${id}/shares`, { method: 'POST', cookie: a, body: { email: 'collab-v@example.com', role: 'viewer' } })
    const ca = await client(id, a)
    const cb = await client(id, bCookie)
    const cv = await client(id, v)
    const edit = (c: { ydoc: Y.Doc }, text: string) => {
      const sv = Y.encodeStateVector(c.ydoc)
      c.ydoc.getText('t').insert(0, text)
      return Y.encodeStateAsUpdate(c.ydoc, sv)
    }
    expect(await ca.sync([edit(ca, 'Hello ')])).toBe(200)
    expect(await cb.sync([edit(cb, 'World ')])).toBe(200)
    expect(await cv.sync([edit(cv, 'VIEWER ')])).toBe(403)
    await ca.sync()
    await cb.sync()
    const t = ca.ydoc.getText('t').toString()
    expect(t).toBe(cb.ydoc.getText('t').toString())
    expect(t).toContain('Hello')
    expect(t).toContain('World')
    expect(t).not.toContain('VIEWER')
    expect((await api(`/api/docs/${id}/sync`, { method: 'POST', cookie: s, body: { after: 0 } })).status).toBe(404)
    expect((await api(`/api/docs/${id}/sync`, { method: 'POST', cookie: a, body: { after: 0, updates: [b64(new Uint8Array([0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0x7f]))] } })).status).toBe(400)
    // a new client sees the full merged state
    const fresh = await client(id, a)
    expect(fresh.ydoc.getText('t').toString()).toBe(t)
  })

  it('compacts many small updates without losing content', async () => {
    const a = await user('compact@example.com')
    const id = (await api('/api/docs', { method: 'POST', cookie: a, body: { title: 'c', content: doc } })).json.id
    const c = await client(id, a)
    for (let i = 0; i < 420; i++) {
      const sv = Y.encodeStateVector(c.ydoc)
      c.ydoc.getText('t').insert(c.ydoc.getText('t').length, String(i % 10))
      await c.sync([Y.encodeStateAsUpdate(c.ydoc, sv)])
    }
    const rows = (await db.query('SELECT COUNT(*)::int AS n FROM doc_updates WHERE doc_id = $1', [id])).rows[0].n
    expect(rows).toBeLessThan(420)
    const fresh = await client(id, a)
    expect(fresh.ydoc.getText('t').toString()).toBe(c.ydoc.getText('t').toString())
  }, 120_000)
})
