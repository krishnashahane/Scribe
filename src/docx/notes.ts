// Footnotes and endnotes: write notes created or edited in SCRIBE into the notes parts.
// Notes that were not touched keep their original XML; notes are never removed from an
// original part (they may be referenced from preserved content SCRIBE does not model).

import type JSZip from 'jszip'
import type { Footnote, PMNode } from '../model/types'
import { NS, children, esc, parseXml, serializeXml, splitFragmentsNs } from './xml'

export type NoteKind = 'footnote' | 'endnote'

const SEPARATORS = (tag: string) =>
  `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}>` +
  `<w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`

/** Note XML; when the note came from the file, its first paragraph and run formatting are reused. */
export function noteXml(kind: NoteKind, note: Footnote): string {
  const tag = kind
  let pPr = '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
  let refRPr = '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
  let textRPr = '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
  if (note.rawXml) {
    try {
      const el = parseXml(note.rawXml, kind).documentElement
      const p = children(el, NS.w, 'p')[0]
      const pp = p && children(p, NS.w, 'pPr')[0]
      if (pp) pPr = serializeXml(pp)
      const runs = p ? children(p, NS.w, 'r') : []
      const refRun = runs.find((r) => children(r, NS.w, `${kind}Ref`).length)
      const rr = refRun && children(refRun, NS.w, 'rPr')[0]
      if (rr) refRPr = serializeXml(rr)
      const textRun = runs.find((r) => children(r, NS.w, 't').length)
      const tr = textRun && children(textRun, NS.w, 'rPr')[0]
      if (tr) textRPr = serializeXml(tr)
      else if (textRun) textRPr = ''
    } catch {
      /* fall back to default formatting */
    }
  }
  const lines = (note.text ?? '').split('\n')
  const paras = lines
    .map((line, i) => {
      const ref = i === 0 ? `<w:r>${refRPr}<w:${kind}Ref/></w:r>${line ? `<w:r>${textRPr}<w:t xml:space="preserve"> </w:t></w:r>` : ''}` : ''
      return `<w:p>${pPr}${ref}${line ? `<w:r>${textRPr}<w:t xml:space="preserve">${esc(line)}</w:t></w:r>` : ''}</w:p>`
    })
    .join('')
  return `<w:${tag} w:id="${esc(note.id)}">${paras}</w:${tag}>`
}

/** Note ids referenced in the body, headers/footers and preserved XML. */
export function referencedNotes(nodes: PMNode[], kind: NoteKind): Set<string> {
  const out = new Set<string>()
  const re = new RegExp(`${kind}Reference\\b[^>]*w:id="([^"]+)"`, 'g')
  const visit = (n: PMNode) => {
    if (n.type === 'noteRef' && (n.attrs?.kind === 'endnote' ? 'endnote' : 'footnote') === kind) out.add(String(n.attrs?.noteId))
    if (typeof n.attrs?.xml === 'string') for (const m of n.attrs.xml.matchAll(re)) out.add(m[1])
    n.content?.forEach(visit)
  }
  nodes.forEach(visit)
  return out
}

/**
 * Update or create the notes part. Returns the part path when one was written, or null when
 * nothing needed to change.
 */
export async function writeNotes(zip: JSZip, kind: NoteKind, notes: Footnote[], referenced: Set<string>, existingPath: string | null, dir: string): Promise<{ path: string; created: boolean } | null> {
  const tag = kind
  const rootTag = `${kind}s`
  const wanted = notes.filter((n) => referenced.has(n.id))
  const changed = wanted.filter((n) => n.edited || !n.rawXml)
  const original = existingPath ? await zip.file(existingPath)?.async('string') : null
  if (!changed.length) return null
  if (original) {
    const doc = parseXml(original, existingPath!)
    const root = doc.documentElement
    const byId = new Map(children(root, NS.w, tag).map((e) => [e.getAttributeNS(NS.w, 'id') ?? '', e]))
    for (const n of changed) {
      const el = doc.importNode(splitFragmentsNs(noteXml(kind, n))[0], true)
      const cur = byId.get(n.id)
      if (cur) root.replaceChild(el, cur)
      else root.appendChild(el)
    }
    const xml = serializeXml(doc)
    zip.file(existingPath!, xml.startsWith('<?xml') ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`)
    return { path: existingPath!, created: false }
  }
  const path = `${dir}${rootTag}.xml`
  const body = wanted.map((n) => (n.rawXml && !n.edited ? n.rawXml : noteXml(kind, n))).join('')
  zip.file(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${rootTag} xmlns:w="${NS.w}" xmlns:r="${NS.r}">${SEPARATORS(tag)}${body}</w:${rootTag}>`)
  return { path, created: true }
}

/** Next free numeric note id (ids -1 and 0 are the separators). */
export function nextNoteId(notes: Footnote[]): string {
  return String(Math.max(0, ...notes.map((n) => Number(n.id)).filter(Number.isFinite)) + 1)
}
