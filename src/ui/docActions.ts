import { generateJSON } from '@tiptap/core'
import { api } from '../api'
import { TEMPLATES } from '../model/templates'
import { emptyMeta, type CompatibilityIssue, type DocMeta, type PMNode } from '../model/types'
import { importDocx } from '../docx/import'
import { sanitizeHtml, textToContent } from '../formats/formats'
import { schemaExtensions } from '../model/extensions'
import { initialYState } from '../collab/session'
import { dataUrlToBytes, uploadImage, uploadOriginal } from '../blobs'

export async function createDocument(title: string, content: PMNode, meta: DocMeta, original?: Uint8Array, originalName?: string) {
  const originalBlob = original ? (await uploadOriginal(original)).id : undefined
  const r = await api<{ id: string }>('/api/docs', {
    method: 'POST',
    body: { title, content, meta, yUpdate: initialYState(content, meta), originalBlob, originalName },
  })
  return r.id
}

/**
 * Move embedded (data: URL) images to server storage so documents stay small.
 * Identical images are uploaded once. Images that cannot be uploaded stay embedded.
 */
export async function externalizeImages(root: PMNode, cache = new Map<string, string>()) {
  const visit = async (n: PMNode) => {
    if (n.type === 'image' && typeof n.attrs?.src === 'string' && n.attrs.src.startsWith('data:')) {
      const src = n.attrs.src
      let url = cache.get(src)
      if (!url) {
        const img = dataUrlToBytes(src)
        if (img) {
          try {
            url = await uploadImage(img.bytes, img.mime)
            cache.set(src, url)
          } catch {
            url = undefined
          }
        }
      }
      if (url) n.attrs = { ...n.attrs, src: url }
    }
    for (const c of n.content ?? []) await visit(c)
  }
  await visit(root)
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
    const cache = new Map<string, string>()
    await externalizeImages(r.doc.content, cache)
    for (const hf of Object.values(r.doc.meta.headerFooters)) await externalizeImages(hf, cache)
    const id = await createDocument(r.doc.meta.metadata.title || name, r.doc.content, r.doc.meta, bytes, file.name)
    return { id, report: r.report }
  }
  if (ext === 'html' || ext === 'htm') {
    const content = generateJSON(sanitizeHtml(await file.text()), schemaExtensions({ collaborative: true })) as PMNode
    await externalizeImages(content)
    return { id: await createDocument(name, content, emptyMeta()), report: [] }
  }
  if (ext === 'txt' || ext === 'md' || file.type.startsWith('text/')) {
    return { id: await createDocument(name, textToContent(await file.text()), emptyMeta()), report: [] }
  }
  if (ext === 'doc') throw new Error('Legacy .doc files are not supported. Save the file as .docx (for example with LibreOffice) first.')
  throw new Error('Unsupported file type. SCRIBE opens .docx, .html and .txt files.')
}

export const openDoc = (id: string) => {
  localStorage.setItem('scribe-last-doc', id)
  location.hash = `#/doc/${id}`
}
