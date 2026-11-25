// SQLite persistence via Node's built-in node:sqlite (no native add-on to compile).
// All writes are parameterised; the schema is created and migrated on startup.
import { DatabaseSync } from 'node:sqlite'
import fs from 'node:fs'
import path from 'node:path'

const MIGRATIONS = [
  `CREATE TABLE users (
     id TEXT PRIMARY KEY,
     email TEXT NOT NULL UNIQUE COLLATE NOCASE,
     name TEXT NOT NULL,
     pass_hash TEXT NOT NULL,
     created INTEGER NOT NULL
   );
   CREATE TABLE sessions (
     token_hash TEXT PRIMARY KEY,
     user_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
     expires INTEGER NOT NULL
   );
   CREATE TABLE documents (
     id TEXT PRIMARY KEY,
     owner_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
     title TEXT NOT NULL,
     created INTEGER NOT NULL,
     updated INTEGER NOT NULL,
     content_json TEXT NOT NULL,
     meta_json TEXT NOT NULL,
     ystate BLOB,
     original BLOB,
     original_name TEXT,
     deleted INTEGER NOT NULL DEFAULT 0
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
     created INTEGER NOT NULL,
     author_id TEXT REFERENCES users(id) ON DELETE SET NULL,
     label TEXT,
     auto INTEGER NOT NULL DEFAULT 0,
     content_json TEXT NOT NULL,
     meta_json TEXT NOT NULL
   );
   CREATE INDEX versions_doc ON versions(doc_id, created DESC);`,
]

export function openDb(file) {
  if (file !== ':memory:') fs.mkdirSync(path.dirname(file), { recursive: true })
  const db = new DatabaseSync(file)
  db.exec('PRAGMA journal_mode = WAL; PRAGMA foreign_keys = ON; PRAGMA busy_timeout = 5000;')
  db.exec('CREATE TABLE IF NOT EXISTS schema_version (v INTEGER NOT NULL)')
  const row = db.prepare('SELECT v FROM schema_version').get()
  let v = row ? row.v : 0
  if (!row) db.prepare('INSERT INTO schema_version (v) VALUES (0)').run()
  while (v < MIGRATIONS.length) {
    db.exec('BEGIN')
    try {
      db.exec(MIGRATIONS[v])
      v++
      db.prepare('UPDATE schema_version SET v = ?').run(v)
      db.exec('COMMIT')
    } catch (e) {
      db.exec('ROLLBACK')
      throw e
    }
  }
  return db
}

/** Role of a user on a document: owner | editor | commenter | viewer | null */
export function roleFor(db, docId, userId) {
  const doc = db.prepare('SELECT owner_id FROM documents WHERE id = ? AND deleted = 0').get(docId)
  if (!doc) return null
  if (doc.owner_id === userId) return 'owner'
  const s = db.prepare('SELECT role FROM shares WHERE doc_id = ? AND user_id = ?').get(docId, userId)
  return s ? s.role : null
}

export const canEdit = (role) => role === 'owner' || role === 'editor'
export const canComment = (role) => canEdit(role) || role === 'commenter'
