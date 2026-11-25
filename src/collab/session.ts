import * as Y from 'yjs'
import { Awareness, applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness'
import { IndexeddbPersistence } from 'y-indexeddb'
import { getSchema } from '@tiptap/core'
import { prosemirrorJSONToYDoc } from '@tiptap/y-tiptap'
import { schemaExtensions } from '../model/extensions'
import type { DocMeta, PMNode } from '../model/types'
import { api, toBase64 } from '../api'

let schemaCache: ReturnType<typeof getSchema> | null = null
export const editorSchema = () => (schemaCache ??= getSchema(schemaExtensions({ collaborative: true })))

/** Build the initial Yjs state for a new document (body + meta), encoded for upload. */
export function initialYState(content: PMNode, meta: DocMeta): string {
  const ydoc = prosemirrorJSONToYDoc(editorSchema(), content, 'default')
  ydoc.getMap('scribe').set('meta', JSON.stringify(meta))
  const update = Y.encodeStateAsUpdate(ydoc)
  ydoc.destroy()
  return toBase64(update)
}

export type SyncStatus = 'connecting' | 'synced' | 'offline' | 'unsynced'

const fromBase64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0))
const MAX_UPDATE_BYTES = 1_900_000

/**
 * Real-time sync over plain HTTP (works on serverless hosts without WebSockets).
 * Local edits are batched and pushed; remote edits and cursors are pulled every ~1.2 s
 * (immediately after a push). Yjs merges updates, so ordering and duplicates are safe.
 */
export class HttpSyncProvider {
  awareness: Awareness
  synced = false
  connected = false
  private after = 0
  private pending: Uint8Array[] = []
  private timer: ReturnType<typeof setTimeout> | null = null
  private inflight = false
  private destroyed = false
  private awarenessDirty = true
  private listeners = new Map<string, Set<(...a: any[]) => void>>()
  private readOnly = false

  constructor(
    private docId: string,
    private ydoc: Y.Doc,
    readOnly = false,
  ) {
    this.readOnly = readOnly
    this.awareness = new Awareness(ydoc)
    ydoc.on('update', this.onUpdate)
    this.awareness.on('update', ({ added, updated, removed }: any, origin: unknown) => {
      if (origin !== this && added.concat(updated, removed).includes(ydoc.clientID)) {
        this.awarenessDirty = true
        this.schedule(150)
      }
    })
    window.addEventListener('online', this.kick)
    document.addEventListener('visibilitychange', this.kick)
    this.schedule(0)
  }

  on(event: 'status' | 'sync', fn: (...a: any[]) => void) {
    if (!this.listeners.has(event)) this.listeners.set(event, new Set())
    this.listeners.get(event)!.add(fn)
  }
  once(event: 'sync', fn: () => void) {
    const wrap = () => {
      this.listeners.get(event)?.delete(wrap)
      fn()
    }
    this.on(event, wrap)
  }
  private emit(event: string, ...args: any[]) {
    this.listeners.get(event)?.forEach((f) => f(...args))
  }
  setReadOnly(v: boolean) {
    this.readOnly = v
  }
  get hasPending() {
    return this.pending.length > 0 || this.inflight
  }

  private onUpdate = (update: Uint8Array, origin: unknown) => {
    if (origin === this || this.readOnly) return
    this.pending.push(update)
    this.schedule(250)
  }
  private kick = () => document.visibilityState !== 'hidden' && this.schedule(0)

  private schedule(ms: number) {
    if (this.destroyed) return
    if (this.timer) clearTimeout(this.timer)
    this.timer = setTimeout(() => this.tick(), ms)
  }

  private async tick() {
    this.timer = null
    if (this.inflight || this.destroyed) return this.schedule(300)
    this.inflight = true
    // merge pending local updates and split into request-sized pieces
    const batch = this.pending
    this.pending = []
    const merged = batch.length ? Y.mergeUpdates(batch) : null
    const updates = merged ? (merged.length > MAX_UPDATE_BYTES ? batch : [merged]) : []
    const awareness = this.awarenessDirty ? { clientId: this.ydoc.clientID, state: toBase64(encodeAwarenessUpdate(this.awareness, [this.ydoc.clientID])) } : undefined
    this.awarenessDirty = false
    let more = false
    try {
      const r = await api<{ updates: { id: number; data: string }[]; last: number; more: boolean; presence: { clientId: number; state: string }[] }>(`/api/docs/${this.docId}/sync`, {
        method: 'POST',
        body: { after: this.after, updates: updates.map(toBase64), awareness },
      })
      Y.transact(this.ydoc, () => {
        for (const u of r.updates) Y.applyUpdate(this.ydoc, fromBase64(u.data), this)
      }, this)
      for (const p of r.presence) applyAwarenessUpdate(this.awareness, fromBase64(p.state), this)
      this.after = r.last
      more = r.more
      if (!this.connected) {
        this.connected = true
        this.emit('status', { status: 'connected' })
      }
      if (!this.synced && !more) {
        this.synced = true
        this.emit('sync', true)
      }
      this.emit('status', { status: 'connected' })
    } catch (e: any) {
      if (e?.status === 403 && updates.length) {
        // the server refused our edits (view-only access): discard them and keep pulling
        this.readOnly = true
        this.pending = []
        this.inflight = false
        return this.schedule(0)
      }
      // keep unsent edits for the next attempt (they are also in IndexedDB)
      this.pending = [...batch, ...this.pending]
      this.awarenessDirty = this.awarenessDirty || !!awareness
      if (this.connected) {
        this.connected = false
        this.emit('status', { status: 'disconnected' })
      }
    } finally {
      this.inflight = false
    }
    const idle = document.visibilityState === 'hidden' ? 10_000 : 1200
    this.schedule(more || this.pending.length ? 0 : this.connected ? idle : 4000)
  }

  async flush() {
    if (this.timer) clearTimeout(this.timer)
    await this.tick()
  }

  destroy() {
    this.destroyed = true
    if (this.timer) clearTimeout(this.timer)
    this.ydoc.off('update', this.onUpdate)
    window.removeEventListener('online', this.kick)
    document.removeEventListener('visibilitychange', this.kick)
    this.awareness.destroy()
  }
}

export interface CollabSession {
  ydoc: Y.Doc
  provider: HttpSyncProvider
  local: IndexeddbPersistence
  destroy: () => void
  whenSynced: Promise<void>
}

/**
 * One collaborative session per open document:
 * - IndexedDB keeps a local copy of every change (crash recovery, offline edits)
 * - the HTTP provider syncs with the server, which stores the shared state in Postgres
 */
export function openSession(docId: string, onStatus: (s: SyncStatus) => void, readOnly = false): CollabSession {
  const ydoc = new Y.Doc()
  const local = new IndexeddbPersistence(`scribe-doc-${docId}`, ydoc)
  const provider = new HttpSyncProvider(docId, ydoc, readOnly)
  const emit = () => onStatus(!provider.connected ? 'offline' : !provider.synced ? 'connecting' : provider.hasPending ? 'unsynced' : 'synced')
  provider.on('status', emit)
  provider.on('sync', emit)
  ydoc.on('update', () => {
    emit()
    setTimeout(emit, 1500)
  })
  const whenSynced = new Promise<void>((resolve) => {
    if (provider.synced) resolve()
    else provider.once('sync', () => resolve())
  })
  return {
    ydoc,
    provider,
    local,
    whenSynced,
    destroy: () => {
      // push any last edits before tearing down (they are also safe in IndexedDB)
      provider.flush().finally(() => {
        provider.destroy()
        local.destroy()
        ydoc.destroy()
      })
    },
  }
}

const COLORS = ['#c2410c', '#0f766e', '#7c3aed', '#be123c', '#1d4ed8', '#15803d', '#a16207', '#0e7490']
export const userColor = (id: string) => COLORS[[...id].reduce((a, c) => a + c.charCodeAt(0), 0) % COLORS.length]
