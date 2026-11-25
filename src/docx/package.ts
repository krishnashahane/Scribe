import JSZip from 'jszip'
import { NS, attr, children, parseXml } from './xml'

// Limits that defend against zip bombs and absurd inputs.
export const LIMITS = {
  maxCompressedBytes: 200 * 1024 * 1024,
  maxUncompressedBytes: 1024 * 1024 * 1024,
  maxEntries: 10000,
  maxXmlPartBytes: 256 * 1024 * 1024,
  maxRatio: 200,
}

export interface Relationship {
  id: string
  type: string
  target: string
  external: boolean
}

export class DocxFormatError extends Error {}

export class OoxmlPackage {
  zip: JSZip
  contentTypes: Document | null = null
  private relCache = new Map<string, Relationship[]>()

  private constructor(zip: JSZip) {
    this.zip = zip
  }

  static async open(data: ArrayBuffer | Uint8Array): Promise<OoxmlPackage> {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data)
    if (bytes.byteLength > LIMITS.maxCompressedBytes) throw new DocxFormatError('File is too large to open safely.')
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
      if (bytes[0] === 0xd0 && bytes[1] === 0xcf) throw new DocxFormatError('This is a legacy .doc (binary Word 97-2003) file. Save it as .docx and try again.')
      throw new DocxFormatError('Not a DOCX file (missing ZIP signature).')
    }
    let zip: JSZip
    try {
      zip = await JSZip.loadAsync(bytes, { checkCRC32: true })
    } catch (e: any) {
      throw new DocxFormatError(`The file is damaged or not a valid DOCX archive: ${e?.message ?? e}`)
    }
    const entries = Object.values(zip.files)
    if (entries.length > LIMITS.maxEntries) throw new DocxFormatError('Archive contains too many entries.')
    let total = 0
    for (const f of entries) {
      const name = f.name
      if (name.includes('..') || name.startsWith('/') || /^[a-zA-Z]:/.test(name)) throw new DocxFormatError(`Unsafe path in archive: ${name}`)
      const size = (f as any)._data?.uncompressedSize ?? 0
      const csize = (f as any)._data?.compressedSize ?? 0
      total += size
      if (size > LIMITS.maxXmlPartBytes) throw new DocxFormatError(`Part ${name} is too large.`)
      if (csize > 1024 && size / csize > LIMITS.maxRatio) throw new DocxFormatError(`Part ${name} has a suspicious compression ratio (possible zip bomb).`)
    }
    if (total > LIMITS.maxUncompressedBytes) throw new DocxFormatError('Archive expands beyond the safe size limit.')
    const pkg = new OoxmlPackage(zip)
    const ct = zip.file('[Content_Types].xml')
    if (!ct) throw new DocxFormatError('Missing [Content_Types].xml — not an Open XML package.')
    pkg.contentTypes = parseXml(await ct.async('string'), '[Content_Types].xml')
    return pkg
  }

  has(path: string) {
    return !!this.zip.file(path)
  }

  async text(path: string): Promise<string | null> {
    const f = this.zip.file(path)
    return f ? f.async('string') : null
  }

  async xml(path: string): Promise<Document | null> {
    const t = await this.text(path)
    return t == null ? null : parseXml(t, path)
  }

  async bytes(path: string): Promise<Uint8Array | null> {
    const f = this.zip.file(path)
    return f ? f.async('uint8array') : null
  }

  /** Relationships of a part, e.g. rels('word/document.xml'). */
  async rels(partPath: string): Promise<Relationship[]> {
    if (this.relCache.has(partPath)) return this.relCache.get(partPath)!
    const dir = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/') + 1) : ''
    const base = partPath.slice(dir.length)
    const relPath = partPath === '' ? '_rels/.rels' : `${dir}_rels/${base}.rels`
    const doc = await this.xml(relPath)
    const out: Relationship[] = []
    if (doc) {
      for (const r of children(doc.documentElement, NS.rel, 'Relationship')) {
        const external = r.getAttribute('TargetMode') === 'External'
        const target = r.getAttribute('Target') || ''
        out.push({
          id: r.getAttribute('Id') || '',
          type: r.getAttribute('Type') || '',
          target: external ? target : resolvePath(dir, target),
          external,
        })
      }
    }
    this.relCache.set(partPath, out)
    return out
  }

  /** Locate the main document part through the package root relationships. */
  async mainDocumentPath(): Promise<string> {
    const rels = await this.rels('')
    const r = rels.find((x) => x.type.endsWith('/officeDocument'))
    const p = r?.target ?? 'word/document.xml'
    if (!this.has(p)) throw new DocxFormatError('The package has no main document part.')
    const ct = this.contentTypeOf(p)
    if (ct && /macroEnabled/i.test(ct)) {
      // .docm — macros are never executed; we preserve the vbaProject part untouched.
    }
    return p
  }

  contentTypeOf(path: string): string | null {
    if (!this.contentTypes) return null
    const root = this.contentTypes.documentElement
    for (const o of children(root, NS.ct, 'Override')) if (o.getAttribute('PartName') === '/' + path) return o.getAttribute('ContentType')
    const ext = path.split('.').pop()?.toLowerCase()
    for (const d of children(root, NS.ct, 'Default')) if (d.getAttribute('Extension')?.toLowerCase() === ext) return d.getAttribute('ContentType')
    return null
  }

  listParts(): string[] {
    return Object.values(this.zip.files)
      .filter((f) => !f.dir)
      .map((f) => f.name)
  }
}

export function resolvePath(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = (baseDir + target).split('/')
  const out: string[] = []
  for (const p of parts) {
    if (p === '..') out.pop()
    else if (p !== '.' && p !== '') out.push(p)
  }
  return out.join('/')
}

export function relativeTarget(fromDir: string, path: string): string {
  // fromDir like 'word/' ; path like 'word/media/image1.png' -> 'media/image1.png'
  return path.startsWith(fromDir) ? path.slice(fromDir.length) : '/' + path
}

export { attr }
