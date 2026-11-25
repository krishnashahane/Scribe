import { generateJSON } from '@tiptap/core'
import { api, toBase64 } from '../api'
import { TEMPLATES } from '../model/templates'
import { emptyMeta, type CompatibilityIssue, type DocMeta, type PMNode } from '../model/types'
import { importDocx } from '../docx/import'
import { sanitizeHtml, textToContent } from '../formats/formats'
import { schemaExtensions } from '../model/extensions'
import { initialYState } from '../collab/session'

export async function createDocument(title: string, content: PMNode, meta: DocMeta, original?: Uint8Array, originalName?: string) {
  const r = await api<{ id: string }>('/api/docs', {
    method: 'POST',
    body: { title, content, meta, yUpdate: initialYState(content, meta), original: original ? toBase64(original) : undefined, originalName },
  })
  return r.id
}

export async function createFromTemplate(templateId: string) {
  const t = TEMPLATES.find((x) => x.id === templateId) ?? TEMPLATES[0]
  const d = t.build()
  return createDocument(t.id.startsWith('blank') ? 'Document' : t.name, d.content, d.meta)
}

export const MAX_IMPORT = 100 * 1024 * 1024
export const OPEN_ACCEPT = '.docx,.docm,.dotx,.html,.htm,.txt,.md,.doc'

/** Import a file into a new SCRIBE document. Throws with a user-readable message. */
export async function importFile(file: File): Promise<{ id: string; report: CompatibilityIssue[] }> {
  if (file.size > MAX_IMPORT) throw new Error('That file is larger than 100 MB.')
  const name = file.name.replace(/\.(docx|docm|dotx|html?|txt|md)$/i, '')
  const ext = file.name.split('.').pop()?.toLowerCase()
  if (ext === 'docx' || ext === 'docm' || ext === 'dotx') {
    const bytes = new Uint8Array(await file.arrayBuffer())
    const r = await importDocx(bytes)
    const id = await createDocument(r.doc.meta.metadata.title || name, r.doc.content, r.doc.meta, bytes, file.name)
    return { id, report: r.report }
  }
  if (ext === 'html' || ext === 'htm') {
    const content = generateJSON(sanitizeHtml(await file.text()), schemaExtensions({ collaborative: true })) as PMNode
    return { id: await createDocument(name, content, emptyMeta()), report: [] }
  }
  if (ext === 'txt' || ext === 'md' || file.type.startsWith('text/')) {
    return { id: await createDocument(name, textToContent(await file.text()), emptyMeta()), report: [] }
  }
  if (ext === 'doc') throw new Error('Legacy .doc files are not supported. Save the file as .docx in Word or LibreOffice first.')
  throw new Error('Unsupported file type. SCRIBE opens .docx, .html and .txt files.')
}

export const openDoc = (id: string) => {
  localStorage.setItem('scribe-last-doc', id)
  location.hash = `#/doc/${id}`
}
