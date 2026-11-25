// PostgreSQL persistence.
// Production (Vercel): Neon/any Postgres via DATABASE_URL using `pg`.
// Local development and tests: embedded PGlite (real Postgres compiled to WASM), so one SQL dialect everywhere.
// Every statement is parameterised ($1, $2, …); no SQL is ever built from user input.

const MIGRATIONS = [
  `CREATE TABLE users (
     id TEXT PRIMARY KEY,
     email TEXT NOT NULL UNIQUE,
     name TEXT NOT NULL,
     pass_hash TEXT NOT NULL,
     created BIGINT NOT NULL
   );
   CREATE TABLE sessions (
     token_hash TEXT PRIMARY KEY,
     user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
     expires BIGINT NOT NULL
   );
   CREATE INDEX sessions_user ON sessions(user_id);
   CREATE TABLE documents (
     id TEXT PRIMARY KEY,
     owner_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
     title TEXT NOT NULL,
     created BIGINT NOT NULL,
     updated BIGINT NOT NULL,
     content_json TEXT NOT NULL,
     meta_json TEXT NOT NULL,
     original_blob TEXT,
     original_name TEXT,
     deleted BOOLEAN NOT NULL DEFAULT FALSE
   );
   CREATE INDEX documents_owner ON documents(owner_id, updated DESC);
   CREATE TABLE shares (
     doc_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
     user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
     role TEXT NOT NULL CHECK (role IN ('viewer','commenter','editor')),
     PRIMARY KEY (doc_id, user_id)
   );
   CREATE TABLE versions (
     id TEXT PRIMARY KEY,
     doc_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
     created BIGINT NOT NULL,
     author_id TEXT REFERENCES users(id) ON DELETE SET NULL,
     label TEXT,
     auto BOOLEAN NOT NULL DEFAULT FALSE,
     content_json TEXT NOT NULL,
     meta_json TEXT NOT NULL
   );
   CREATE INDEX versions_doc ON versions(doc_id, created DESC);
   -- Yjs CRDT updates, appended by editors and pulled by every client
   CREATE TABLE doc_updates (
     id BIGSERIAL PRIMARY KEY,
     doc_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
     data BYTEA NOT NULL,
     size INTEGER NOT NULL
   );
   CREATE INDEX doc_updates_doc ON doc_updates(doc_id, id);
   -- cursor/presence state per client, short-lived
   CREATE TABLE presence (
     doc_id TEXT NOT NULL REFERENCES documents(id) ON DELETE CASCADE,
     client_id BIGINT NOT NULL,
     user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
     state BYTEA NOT NULL,
     updated BIGINT NOT NULL,
     PRIMARY KEY (doc_id, client_id)
   );
   -- binary objects (images, original DOCX files), stored in chunks to fit serverless request limits
   CREATE TABLE blobs (
     id TEXT PRIMARY KEY,
     owner_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
     kind TEXT NOT NULL CHECK (kind IN ('media','original')),
     mime TEXT NOT NULL,
     size INTEGER NOT NULL,
     chunks INTEGER NOT NULL,
     complete BOOLEAN NOT NULL DEFAULT FALSE,
     created BIGINT NOT NULL
   );
   CREATE TABLE blob_chunks (
     blob_id TEXT NOT NULL REFERENCES blobs(id) ON DELETE CASCADE,
     idx INTEGER NOT NULL,
     data BYTEA NOT NULL,
     PRIMARY KEY (blob_id, idx)
   );
   CREATE TABLE rate_limits (
     key TEXT PRIMARY KEY,
     window_start BIGINT NOT NULL,
     count INTEGER NOT NULL
   );`,
]

/** Normalise bytea values (pg → Buffer, PGlite → Uint8Array). */
export const toBuffer = (v) => (v == null ? null : Buffer.isBuffer(v) ? v : Buffer.from(v))

async function migrate(db) {
  await db.query('CREATE TABLE IF NOT EXISTS schema_version (v INTEGER NOT NULL)')
  const { rows } = await db.query('SELECT v FROM schema_version')
  let v = rows[0]?.v ?? 0
  if (!rows.length) await db.query('INSERT INTO schema_version (v) VALUES (0)')
  while (v < MIGRATIONS.length) {
    await db.tx(async (t) => {
      // lock so concurrent cold starts don't run the same migration twice
      await t.query('LOCK TABLE schema_version IN EXCLUSIVE MODE')
      const cur = (await t.query('SELECT v FROM schema_version')).rows[0].v
      if (cur !== v) return
      await t.exec(MIGRATIONS[v])
      await t.query('UPDATE schema_version SET v = $1', [v + 1])
    })
    v = (await db.query('SELECT v FROM schema_version')).rows[0].v
  }
}

async function openPg(url) {
  const { default: pg } = await import('pg')
  // BIGINT columns hold millisecond timestamps (safe integers): return numbers, not strings
  pg.types.setTypeParser(20, (v) => Number(v))
  const pool = new pg.Pool({ connectionString: url, max: 5, idleTimeoutMillis: 10_000, connectionTimeoutMillis: 10_000 })
  const db = {
    query: (text, params = []) => pool.query(text, params),
    exec: (text) => pool.query(text),
    async tx(fn) {
      const client = await pool.connect()
      try {
        await client.query('BEGIN')
        const t = { query: (text, params = []) => client.query(text, params), exec: (text) => client.query(text) }
        const out = await fn(t)
        await client.query('COMMIT')
        return out
      } catch (e) {
        await client.query('ROLLBACK').catch(() => {})
        throw e
      } finally {
        client.release()
      }
    },
    close: () => pool.end(),
  }
  return db
}

async function openPglite(dir) {
  const { PGlite } = await import('@electric-sql/pglite')
  const pg = dir ? new PGlite(dir) : new PGlite()
  await pg.waitReady
  const wrap = (q) => ({
    query: async (text, params = []) => {
      const r = await q.query(text, params)
      return { rows: r.rows.map((row) => Object.fromEntries(Object.entries(row).map(([k, v]) => [k, typeof v === 'bigint' ? Number(v) : v]))) }
    },
    exec: (text) => q.exec(text),
  })
  const base = wrap(pg)
  return {
    ...base,
    tx: (fn) => pg.transaction((t) => fn(wrap(t))),
    close: () => pg.close(),
  }
}

/**
 * Open the database once per process (Vercel reuses warm instances).
 * DATABASE_URL → Postgres; otherwise PGlite (directory, or in-memory when dir is ':memory:').
 */
export async function openDb({ url = process.env.DATABASE_URL, dir } = {}) {
  const db = url ? await openPg(url) : await openPglite(dir === ':memory:' ? undefined : dir)
  await migrate(db)
  return db
}

export async function roleFor(db, docId, userId) {
  const { rows } = await db.query(
    `SELECT d.owner_id, s.role FROM documents d LEFT JOIN shares s ON s.doc_id = d.id AND s.user_id = $2
     WHERE d.id = $1 AND d.deleted = FALSE`,
    [docId, userId],
  )
  if (!rows.length) return null
  if (rows[0].owner_id === userId) return 'owner'
  return rows[0].role ?? null
}

export const canEdit = (role) => role === 'owner' || role === 'editor'

/**
 * Fixed-window rate limit stored in the database, so it holds across serverless instances.
 * Returns true when the request is allowed.
 */
export async function rateLimit(db, key, limit, windowMs) {
  const now = Date.now()
  const start = now - (now % windowMs)
  const { rows } = await db.query(
    `INSERT INTO rate_limits (key, window_start, count) VALUES ($1, $2, 1)
     ON CONFLICT (key) DO UPDATE SET
       count = CASE WHEN rate_limits.window_start = EXCLUDED.window_start THEN rate_limits.count + 1 ELSE 1 END,
       window_start = EXCLUDED.window_start
     RETURNING count`,
    [key, start],
  )
  if (Math.random() < 0.01) await db.query('DELETE FROM rate_limits WHERE window_start < $1', [now - 24 * 3600_000])
  return rows[0].count <= limit
}
