import type { Express } from 'express'

export interface Db {
  query(text: string, params?: unknown[]): Promise<{ rows: any[] }>
  exec(text: string): Promise<unknown>
  tx<T>(fn: (t: Pick<Db, 'query' | 'exec'>) => Promise<T>): Promise<T>
  close(): Promise<void>
}

export declare const SECURITY_HEADERS: Record<string, string>

export declare function createApp(opts?: {
  db?: Db
  dbUrl?: string
  dbDir?: string
  secureCookies?: boolean
  serveStatic?: boolean
  authRateLimit?: number
  allowedOrigins?: string[]
}): Promise<{ app: Express; db: Db }>
