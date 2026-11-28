import type { Db } from './app.js'

export declare function openDb(opts?: { url?: string; dir?: string }): Promise<Db>
export declare function roleFor(db: Db, docId: string, userId: string): Promise<'owner' | 'editor' | 'viewer' | 'commenter' | null>
export declare const canEdit: (role: string | null) => boolean
export declare function rateLimit(db: Db, key: string, limit: number, windowMs: number): Promise<boolean>
export declare function housekeep(db: Db, now?: number): Promise<void>
export declare const toBuffer: (v: unknown) => Buffer | null
