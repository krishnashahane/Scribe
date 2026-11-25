// Namespace-aware XML helpers for OOXML. Parsing refuses DOCTYPEs, which blocks
// entity-expansion ("billion laughs") and external-entity attacks.

export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  v: 'urn:schemas-microsoft-com:vml',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
} as const

export class XmlSecurityError extends Error {}

export function parseXml(text: string, partName = 'part'): Document {
  // Strip BOM
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1)
  if (/<!DOCTYPE/i.test(text.slice(0, 4096)) || /<!ENTITY/i.test(text)) {
    throw new XmlSecurityError(`${partName}: DOCTYPE/ENTITY declarations are not allowed in OOXML parts`)
  }
  const doc = new DOMParser().parseFromString(text, 'application/xml')
  const err = doc.getElementsByTagName('parsererror')[0]
  if (err) throw new Error(`${partName}: malformed XML (${(err.textContent || '').slice(0, 160)})`)
  return doc
}

export function serializeXml(node: Node): string {
  return new XMLSerializer().serializeToString(node)
}

export const isW = (el: Element | null | undefined, local: string) => !!el && el.namespaceURI === NS.w && el.localName === local

export function children(el: Element | null | undefined, ns?: string, local?: string): Element[] {
  if (!el) return []
  const out: Element[] = []
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
    if ((ns === undefined || c.namespaceURI === ns) && (local === undefined || c.localName === local)) out.push(c)
  }
  return out
}

export function child(el: Element | null | undefined, ns: string, local: string): Element | null {
  if (!el) return null
  for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.namespaceURI === ns && c.localName === local) return c
  return null
}

export const wChild = (el: Element | null | undefined, local: string) => child(el, NS.w, local)
export const wChildren = (el: Element | null | undefined, local?: string) => children(el, NS.w, local)

export function descendants(el: Element, ns: string, local: string): Element[] {
  return Array.from(el.getElementsByTagNameNS(ns, local))
}

/** w:val style attribute (also accepts un-namespaced attributes from sloppy producers). */
export function attr(el: Element | null | undefined, local: string, ns: string = NS.w): string | null {
  if (!el) return null
  return el.getAttributeNS(ns, local) ?? el.getAttribute(`w:${local}`) ?? el.getAttribute(local)
}

export const wVal = (el: Element | null | undefined) => attr(el, 'val')

/** OOXML on/off property: <w:b/>, <w:b w:val="true|1|on"/> are on; "false|0|off|none" are off. */
export function onOff(el: Element | null | undefined): boolean | null {
  if (!el) return null
  const v = wVal(el)
  if (v == null) return true
  return !['false', '0', 'off', 'none'].includes(v.toLowerCase())
}

export function intAttr(el: Element | null | undefined, local: string, ns: string = NS.w): number | null {
  const v = attr(el, local, ns)
  if (v == null || v === '') return null
  const n = parseInt(v, 10)
  return Number.isFinite(n) ? n : null
}

export function textContent(el: Element): string {
  let out = ''
  for (const t of descendants(el, NS.w, 't')) out += t.textContent ?? ''
  return out
}

// ---------- writing ----------

export function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // XML 1.0 forbids most control characters
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
}

export function el(name: string, attrs: Record<string, string | number | null | undefined> = {}, inner = ''): string {
  let s = `<${name}`
  for (const [k, v] of Object.entries(attrs)) if (v !== null && v !== undefined) s += ` ${k}="${esc(String(v))}"`
  return inner ? `${s}>${inner}</${name}>` : `${s}/>`
}

/**
 * Sort serialized child elements by the order the schema requires.
 * Word rejects files whose property children are out of sequence.
 */
export function orderChildren(xmlFragments: string[], order: readonly string[]): string {
  const rank = (frag: string) => {
    const m = /^<\s*(?:w:)?([A-Za-z0-9]+)/.exec(frag)
    const i = m ? order.indexOf(m[1]) : -1
    return i === -1 ? order.length : i
  }
  return xmlFragments
    .map((f, i) => ({ f, i, r: rank(f) }))
    .sort((a, b) => a.r - b.r || a.i - b.i)
    .map((x) => x.f)
    .join('')
}

export const RPR_ORDER = [
  'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss',
  'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs',
  'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
  'specVanish', 'oMath', 'rPrChange',
] as const

export const PPR_ORDER = [
  'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr',
  'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
  'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
  'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
  'sectPr', 'pPrChange',
] as const

export const SECTPR_ORDER = [
  'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders',
  'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter',
  'docGrid', 'printerSettings', 'sectPrChange',
] as const

export const TCPR_ORDER = [
  'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText',
  'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange',
] as const

export const TBLPR_ORDER = [
  'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
  'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription',
  'tblPrChange',
] as const

export const TRPR_ORDER = [
  'cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight', 'tblHeader',
  'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange',
] as const

/** Serialize the children of an element whose local names are NOT in `handled`. */
export function extraChildrenXml(parent: Element | null, handled: Set<string>): string[] {
  if (!parent) return []
  return children(parent)
    .filter((c) => !(c.namespaceURI === NS.w && handled.has(c.localName)))
    .map((c) => serializeXml(c))
}

/** Remove xmlns declarations that XMLSerializer adds to preserved fragments when the root already declares them. */
export function stripRedundantNs(frag: string, declared: Record<string, string>): string {
  return frag.replace(/\s+xmlns:([A-Za-z0-9]+)="([^"]*)"/g, (m, p, uri) => (declared[p] === uri ? '' : m))
}
