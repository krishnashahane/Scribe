import * as Y from 'yjs'
import { WebsocketProvider } from 'y-websocket'
import { IndexeddbPersistence } from 'y-indexeddb'
import { getSchema } from '@tiptap/core'
import { prosemirrorJSONToYDoc } from '@tiptap/y-tiptap'
import { schemaExtensions } from '../model/extensions'
import type { DocMeta, PMNode } from '../model/types'
import { toBase64 } from '../api'

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

export interface CollabSession {
  ydoc: Y.Doc
  provider: WebsocketProvider
  local: IndexeddbPersistence
  destroy: () => void
  whenSynced: Promise<void>
}

/**
 * One collaborative session per open document:
 * - IndexedDB keeps a local copy of every change (crash recovery, offline edits)
 * - the WebSocket provider syncs with the server, which persists the shared state
 */
export function openSession(docId: string, onStatus: (s: SyncStatus) => void): CollabSession {
  const ydoc = new Y.Doc()
  const local = new IndexeddbPersistence(`scribe-doc-${docId}`, ydoc)
  const proto = location.protocol === 'https:' ? 'wss' : 'ws'
  const provider = new WebsocketProvider(`${proto}://${location.host}/collab`, docId, ydoc, { maxBackoffTime: 5000 })
  let pending = false
  const emit = () => onStatus(!provider.wsconnected ? 'offline' : !provider.synced ? 'connecting' : pending ? 'unsynced' : 'synced')
  provider.on('status', emit)
  provider.on('sync', () => {
    pending = false
    emit()
  })
  ydoc.on('update', (_u: Uint8Array, origin: unknown) => {
    if (origin !== provider) {
      pending = true
      emit()
      // y-websocket sends immediately; the server echo isn't sent back, so mark synced on the next tick when connected
      setTimeout(() => {
        if (provider.wsconnected && provider.ws?.bufferedAmount === 0) {
          pending = false
          emit()
        }
      }, 300)
    }
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
      provider.destroy()
      local.destroy()
      ydoc.destroy()
    },
  }
}

const COLORS = ['#c2410c', '#0f766e', '#7c3aed', '#be123c', '#1d4ed8', '#15803d', '#a16207', '#0e7490']
export const userColor = (id: string) => COLORS[[...id].reduce((a, c) => a + c.charCodeAt(0), 0) % COLORS.length]
