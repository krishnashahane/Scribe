// @vitest-environment node
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import WebSocket from 'ws'
import * as Y from 'yjs'
import * as syncProtocol from 'y-protocols/sync'
import * as encoding from 'lib0/encoding'
import * as decoding from 'lib0/decoding'
import { createApp } from '../../server/app.js'

let base = ''
let ctx: ReturnType<typeof createApp>

beforeAll(async () => {
  ctx = createApp({ dbFile: ':memory:', log: () => {}, authRateLimit: 1000 })
  await new Promise<void>((r) => ctx.server.listen(0, '127.0.0.1', () => r()))
  const addr = ctx.server.address() as any
  base = `http://127.0.0.1:${addr.port}`
})
afterAll(() => ctx.server.close())

async function api(path: string, opts: { method?: string; body?: any; cookie?: string; csrf?: boolean } = {}) {
  const res = await fetch(base + path, {
    method: opts.method ?? 'GET',
    headers: {
      'content-type': 'application/json',
      ...(opts.cookie ? { cookie: opts.cookie } : {}),
      ...(opts.csrf === false ? {} : { 'x-scribe-request': '1' }),
    },
    body: opts.body ? JSON.stringify(opts.body) : undefined,
  })
  const cookie = res.headers.get('set-cookie')?.split(';')[0]
  return { status: res.status, json: await res.json().catch(() => null), cookie }
}

async function user(email: string) {
  const r = await api('/api/auth/register', { method: 'POST', body: { email, password: 'correct horse battery', name: email.split('@')[0] } })
  expect(r.status).toBe(201)
  return r.cookie!
}

const doc = { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Hello' }] }] }

describe('auth', () => {
  it('rate-limits repeated login attempts', async () => {
    const limited = createApp({ dbFile: ':memory:', log: () => {}, authRateLimit: 3 })
    await new Promise<void>((r) => limited.server.listen(0, '127.0.0.1', () => r()))
    const url = `http://127.0.0.1:${(limited.server.address() as any).port}/api/auth/login`
    const codes: number[] = []
    for (let i = 0; i < 5; i++) codes.push((await fetch(url, { method: 'POST', headers: { 'content-type': 'application/json', 'x-scribe-request': '1' }, body: JSON.stringify({ email: 'a@b.c', password: 'x' }) })).status)
    limited.server.close()
    expect(codes).toEqual([401, 401, 401, 429, 429])
  })
  it('registers, logs in, rejects bad passwords and duplicate emails', async () => {
    const c = await user('ada@example.com')
    expect((await api('/api/auth/me', { cookie: c })).json.user.email).toBe('ada@example.com')
    expect((await api('/api/auth/register', { method: 'POST', body: { email: 'ada@example.com', password: 'another password' } })).status).toBe(409)
    expect((await api('/api/auth/login', { method: 'POST', body: { email: 'ada@example.com', password: 'wrong password' } })).status).toBe(401)
    const ok = await api('/api/auth/login', { method: 'POST', body: { email: 'ADA@example.com', password: 'correct horse battery' } })
    expect(ok.status).toBe(200)
    expect((await api('/api/auth/register', { method: 'POST', body: { email: 'x@y.z', password: 'short' } })).status).toBe(400)
  })
  it('sets an HttpOnly SameSite=Strict session cookie', async () => {
    const res = await fetch(base + '/api/auth/register', { method: 'POST', headers: { 'content-type': 'application/json', 'x-scribe-request': '1' }, body: JSON.stringify({ email: 'cookie@example.com', password: 'longenough1' }) })
    const sc = res.headers.get('set-cookie')!
    expect(sc).toMatch(/HttpOnly/)
    expect(sc).toMatch(/SameSite=Strict/)
  })
  it('refuses state-changing requests without the CSRF header or from another origin', async () => {
    const c = await user('csrf@example.com')
    expect((await api('/api/docs', { method: 'POST', cookie: c, csrf: false, body: { title: 'x', content: doc } })).status).toBe(403)
    const res = await fetch(base + '/api/docs', { method: 'POST', headers: { 'content-type': 'application/json', 'x-scribe-request': '1', origin: 'https://evil.example', cookie: c }, body: JSON.stringify({ title: 'x', content: doc }) })
    expect(res.status).toBe(403)
  })
  it('sends security headers', async () => {
    const res = await fetch(base + '/api/health')
    expect(res.headers.get('content-security-policy')).toContain("default-src 'self'")
    expect(res.headers.get('x-content-type-options')).toBe('nosniff')
    expect(res.headers.get('x-frame-options')).toBe('DENY')
  })
})

describe('documents and authorization', () => {
  it('enforces owner / editor / viewer / stranger permissions server-side', async () => {
    const owner = await user('owner@example.com')
    const editor = await user('editor@example.com')
    const viewer = await user('viewer@example.com')
    const stranger = await user('stranger@example.com')
    const { json } = await api('/api/docs', { method: 'POST', cookie: owner, body: { title: 'Plan', content: doc, meta: {} } })
    const id = json.id
    expect((await api(`/api/docs/${id}`, { cookie: stranger })).status).toBe(404)
    expect((await api(`/api/docs/${id}`, { method: 'PATCH', cookie: stranger, body: { title: 'hijack' } })).status).toBe(404)
    await api(`/api/docs/${id}/shares`, { method: 'POST', cookie: owner, body: { email: 'editor@example.com', role: 'editor' } })
    await api(`/api/docs/${id}/shares`, { method: 'POST', cookie: owner, body: { email: 'viewer@example.com', role: 'viewer' } })
    expect((await api(`/api/docs/${id}/shares`, { method: 'POST', cookie: editor, body: { email: 'stranger@example.com', role: 'editor' } })).status).toBe(403)
    expect((await api(`/api/docs/${id}`, { cookie: viewer })).json.role).toBe('viewer')
    expect((await api(`/api/docs/${id}`, { method: 'PATCH', cookie: viewer, body: { title: 'nope' } })).status).toBe(403)
    expect((await api(`/api/docs/${id}`, { method: 'PATCH', cookie: editor, body: { title: 'Plan v2' } })).status).toBe(200)
    expect((await api(`/api/docs/${id}`, { method: 'DELETE', cookie: editor })).status).toBe(403)
    const list = await api('/api/docs', { cookie: viewer })
    expect(list.json.docs.map((d: any) => d.title)).toContain('Plan v2')
    expect((await api(`/api/docs/${id}`, { method: 'DELETE', cookie: owner })).status).toBe(200)
    expect((await api(`/api/docs/${id}`, { cookie: owner })).status).toBe(404)
  })

  it('rejects invalid payloads and non-DOCX originals', async () => {
    const c = await user('payload@example.com')
    expect((await api('/api/docs', { method: 'POST', cookie: c, body: { title: 'x', content: { type: 'html' } } })).status).toBe(400)
    expect((await api('/api/docs', { method: 'POST', cookie: c, body: { title: 'x', content: doc, original: Buffer.from('not a zip').toString('base64') } })).status).toBe(400)
  })

  it('stores the original DOCX untouched and serves it back', async () => {
    const c = await user('orig@example.com')
    const bytes = Buffer.from([0x50, 0x4b, 3, 4, 1, 2, 3, 4, 5])
    const { json } = await api('/api/docs', { method: 'POST', cookie: c, body: { title: 'x', content: doc, original: bytes.toString('base64'), originalName: 'a.docx' } })
    const res = await fetch(`${base}/api/docs/${json.id}/original`, { headers: { cookie: c } })
    expect(Buffer.from(await res.arrayBuffer()).equals(bytes)).toBe(true)
  })

  it('creates, lists and restores versions; dedupes identical auto versions', async () => {
    const c = await user('versions@example.com')
    const { json } = await api('/api/docs', { method: 'POST', cookie: c, body: { title: 'v', content: doc } })
    const v1 = await api(`/api/docs/${json.id}/versions`, { method: 'POST', cookie: c, body: { content: doc, meta: {}, label: 'First' } })
    expect(v1.status).toBe(201)
    const dup = await api(`/api/docs/${json.id}/versions`, { method: 'POST', cookie: c, body: { content: doc, meta: {}, auto: true } })
    expect(dup.json.skipped).toBe(true)
    const list = await api(`/api/docs/${json.id}/versions`, { cookie: c })
    expect(list.json.versions).toHaveLength(1)
    const got = await api(`/api/docs/${json.id}/versions/${v1.json.id}`, { cookie: c })
    expect(got.json.content).toEqual(doc)
  })

  it('reports AI as unconfigured instead of faking output', async () => {
    const prev = process.env.ANTHROPIC_API_KEY
    delete process.env.ANTHROPIC_API_KEY
    delete process.env.SCRIBE_AI_PROVIDER
    const c = await user('ai@example.com')
    expect((await api('/api/ai/status', { cookie: c })).json.provider).toBe('none')
    const r = await api('/api/ai', { method: 'POST', cookie: c, body: { task: 'improve', text: 'hello' } })
    expect(r.status).toBe(503)
    if (prev) process.env.ANTHROPIC_API_KEY = prev
  })
})

// ---------------------------------------------------------------- collaboration

function connect(id: string, cookie: string) {
  const ydoc = new Y.Doc()
  const ws = new WebSocket(`${base.replace('http', 'ws')}/collab/${id}`, { headers: { cookie } })
  ws.binaryType = 'arraybuffer'
  const synced = new Promise<void>((resolve, reject) => {
    ws.on('error', reject)
    ws.on('unexpected-response', (_r, res) => reject(new Error('HTTP ' + res.statusCode)))
    ws.on('message', (data: ArrayBuffer) => {
      const dec = decoding.createDecoder(new Uint8Array(data))
      if (decoding.readVarUint(dec) !== 0) return
      const enc = encoding.createEncoder()
      encoding.writeVarUint(enc, 0)
      const t = syncProtocol.readSyncMessage(dec, enc, ydoc, 'remote')
      if (encoding.length(enc) > 1) ws.send(encoding.toUint8Array(enc))
      if (t === syncProtocol.messageYjsSyncStep2) resolve()
    })
    ws.on('open', () => {
      const enc = encoding.createEncoder()
      encoding.writeVarUint(enc, 0)
      syncProtocol.writeSyncStep1(enc, ydoc)
      ws.send(encoding.toUint8Array(enc))
    })
  })
  ydoc.on('update', (u: Uint8Array, origin: any) => {
    if (origin === 'remote') return
    const enc = encoding.createEncoder()
    encoding.writeVarUint(enc, 0)
    syncProtocol.writeUpdate(enc, u)
    if (ws.readyState === 1) ws.send(encoding.toUint8Array(enc))
  })
  return { ydoc, ws, synced }
}
const wait = (ms: number) => new Promise((r) => setTimeout(r, ms))

describe('real-time collaboration', () => {
  it('syncs concurrent edits between editors, ignores viewer writes, rejects strangers, persists state', async () => {
    const a = await user('collab-a@example.com')
    const b = await user('collab-b@example.com')
    const v = await user('collab-v@example.com')
    const s = await user('collab-s@example.com')
    const { json } = await api('/api/docs', { method: 'POST', cookie: a, body: { title: 'c', content: doc } })
    await api(`/api/docs/${json.id}/shares`, { method: 'POST', cookie: a, body: { email: 'collab-b@example.com', role: 'editor' } })
    await api(`/api/docs/${json.id}/shares`, { method: 'POST', cookie: a, body: { email: 'collab-v@example.com', role: 'viewer' } })
    const ca = connect(json.id, a)
    const cb = connect(json.id, b)
    const cv = connect(json.id, v)
    await Promise.all([ca.synced, cb.synced, cv.synced])
    ca.ydoc.getText('t').insert(0, 'Hello ')
    cb.ydoc.getText('t').insert(0, 'World ')
    cv.ydoc.getText('t').insert(0, 'VIEWER ')
    await wait(400)
    const ta = ca.ydoc.getText('t').toString()
    expect(ta).toBe(cb.ydoc.getText('t').toString())
    expect(ta).toContain('Hello')
    expect(ta).toContain('World')
    expect(ta).not.toContain('VIEWER')
    await expect(connect(json.id, s).synced).rejects.toThrow(/403/)
    ca.ws.close()
    cb.ws.close()
    cv.ws.close()
    await wait(300)
    const row = ctx.db.prepare('SELECT ystate FROM documents WHERE id = ?').get(json.id) as any
    const restored = new Y.Doc()
    Y.applyUpdate(restored, new Uint8Array(row.ystate))
    expect(restored.getText('t').toString()).toBe(ta)
  })
})
