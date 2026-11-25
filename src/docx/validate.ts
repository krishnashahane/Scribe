import JSZip from 'jszip'
import { resolvePath } from './package'
import { NS, PPR_ORDER, RPR_ORDER, SECTPR_ORDER, TCPR_ORDER, children, parseXml } from './xml'

/**
 * Structural OOXML checks run on every export (and in tests):
 * well-formed XML, resolvable relationships, content types for every part,
 * property-child ordering and uniqueness, and required document structure.
 * Returns a list of problems; empty means the package passed.
 */
export async function validatePackage(data: Uint8Array): Promise<string[]> {
  const problems: string[] = []
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(data)
  } catch (e: any) {
    return [`not a zip: ${e.message}`]
  }
  const parts = Object.values(zip.files).filter((f) => !f.dir).map((f) => f.name)
  const ctXml = await zip.file('[Content_Types].xml')?.async('string')
  if (!ctXml) return ['missing [Content_Types].xml']
  const ct = parseXml(ctXml)
  const defaults = new Set(children(ct.documentElement, NS.ct, 'Default').map((d) => d.getAttribute('Extension')!.toLowerCase()))
  const overrides = new Set(children(ct.documentElement, NS.ct, 'Override').map((d) => d.getAttribute('PartName')!))
  for (const o of overrides) if (!parts.includes(o.slice(1))) problems.push(`content type override for missing part ${o}`)
  for (const p of parts) {
    if (p === '[Content_Types].xml') continue
    const ext = p.split('.').pop()!.toLowerCase()
    if (!overrides.has('/' + p) && !defaults.has(ext)) problems.push(`no content type for ${p}`)
  }
  const docs = new Map<string, Document>()
  for (const p of parts) {
    if (!/\.(xml|rels)$/i.test(p)) continue
    try {
      docs.set(p, parseXml(await zip.file(p)!.async('string'), p))
    } catch (e: any) {
      problems.push(`${p}: ${e.message}`)
    }
  }
  // relationships
  for (const [p, d] of docs) {
    if (!p.endsWith('.rels')) continue
    const base = p.replace(/_rels\/([^/]+)\.rels$/, '$1')
    const dir = base.includes('/') ? base.slice(0, base.lastIndexOf('/') + 1) : ''
    const ids = new Set<string>()
    for (const r of children(d.documentElement, NS.rel, 'Relationship')) {
      const id = r.getAttribute('Id')!
      if (ids.has(id)) problems.push(`${p}: duplicate relationship id ${id}`)
      ids.add(id)
      if (r.getAttribute('TargetMode') === 'External') continue
      const t = resolvePath(p === '_rels/.rels' ? '' : dir, r.getAttribute('Target')!)
      if (!parts.includes(t)) problems.push(`${p}: relationship ${id} targets missing part ${t}`)
    }
    // every r:id used in the source part must resolve
    const src = docs.get(base)
    if (src) {
      const xml = await zip.file(base)!.async('string')
      for (const m of xml.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)) if (!ids.has(m[1])) problems.push(`${base}: r:id ${m[1]} has no relationship`)
    }
  }
  const main = docs.get('word/document.xml')
  if (!main) problems.push('missing word/document.xml')
  else {
    const body = children(main.documentElement, NS.w, 'body')[0]
    if (!body) problems.push('document has no body')
    else {
      const last = children(body).at(-1)
      if (!last || last.localName !== 'sectPr') problems.push('body must end with sectPr')
    }
  }
  // ordering / duplicates of property containers
  const check = (d: Document, part: string, tag: string, order: readonly string[]) => {
    for (const e of Array.from(d.getElementsByTagNameNS(NS.w, tag))) {
      const names = children(e, NS.w).map((c) => c.localName)
      let prev = -1
      const seen = new Set<string>()
      for (const n of names) {
        const i = order.indexOf(n)
        if (i === -1) continue
        if (i < prev) {
          problems.push(`${part}: <w:${tag}> child <w:${n}> out of schema order (${names.join(',')})`)
          break
        }
        if (seen.has(n) && !['headerReference', 'footerReference'].includes(n)) {
          problems.push(`${part}: duplicate <w:${n}> in <w:${tag}>`)
          break
        }
        seen.add(n)
        prev = i
      }
    }
  }
  for (const [p, d] of docs) {
    if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/.test(p)) continue
    check(d, p, 'rPr', RPR_ORDER)
    check(d, p, 'pPr', PPR_ORDER)
    check(d, p, 'sectPr', SECTPR_ORDER)
    check(d, p, 'tcPr', TCPR_ORDER)
    for (const tc of Array.from(d.getElementsByTagNameNS(NS.w, 'tc'))) {
      const last = children(tc, NS.w).filter((c) => c.localName !== 'tcPr').at(-1)
      if (!last || last.localName !== 'p') problems.push(`${p}: table cell does not end with a paragraph`)
    }
    for (const tbl of Array.from(d.getElementsByTagNameNS(NS.w, 'tbl'))) {
      const gridCols = children(children(tbl, NS.w, 'tblGrid')[0], NS.w, 'gridCol').length
      if (!gridCols) problems.push(`${p}: table without tblGrid`)
    }
  }
  // comment anchors refer to existing comments
  const comments = docs.get('word/comments.xml')
  const commentIds = new Set(comments ? children(comments.documentElement, NS.w, 'comment').map((c) => c.getAttributeNS(NS.w, 'id')) : [])
  if (main) for (const r of Array.from(main.getElementsByTagNameNS(NS.w, 'commentReference'))) if (!commentIds.has(r.getAttributeNS(NS.w, 'id'))) problems.push(`commentReference ${r.getAttributeNS(NS.w, 'id')} has no comment`)
  // numbering references
  const numbering = docs.get('word/numbering.xml')
  const numIds = new Set(numbering ? children(numbering.documentElement, NS.w, 'num').map((n) => n.getAttributeNS(NS.w, 'numId')) : [])
  if (main) for (const n of Array.from(main.getElementsByTagNameNS(NS.w, 'numId'))) {
    const v = n.getAttributeNS(NS.w, 'val')
    if (v !== '0' && !numIds.has(v)) problems.push(`numId ${v} is not defined in numbering.xml`)
  }
  return problems
}
