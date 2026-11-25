// Styles gallery: built-in styles for new documents, and writing edited styles into styles.xml.

import type { DocMeta, StyleDef } from '../model/types'
import { NS, children, esc, orderChildren, PPR_ORDER, RPR_ORDER, serializeXml, splitFragmentsNs } from '../docx/xml'

/** Styles every SCRIBE document can use; they match the style sheet written for new files. */
export const BUILTIN_STYLES: StyleDef[] = [
  { id: 'Normal', name: 'Normal', type: 'paragraph' },
  { id: 'Title', name: 'Title', type: 'paragraph', basedOn: 'Normal', fontSizePt: 28, spacingAfter: 120 },
  { id: 'Subtitle', name: 'Subtitle', type: 'paragraph', basedOn: 'Normal', fontSizePt: 14, color: '#5a6375', spacingAfter: 160 },
  { id: 'Heading1', name: 'heading 1', type: 'paragraph', basedOn: 'Normal', fontSizePt: 18, bold: true, color: '#1f2a44', spacingBefore: 360, spacingAfter: 80, outlineLevel: 0 },
  { id: 'Heading2', name: 'heading 2', type: 'paragraph', basedOn: 'Normal', fontSizePt: 15, bold: true, color: '#1f2a44', spacingBefore: 240, spacingAfter: 80, outlineLevel: 1 },
  { id: 'Heading3', name: 'heading 3', type: 'paragraph', basedOn: 'Normal', fontSizePt: 13, bold: true, color: '#1f2a44', spacingBefore: 240, spacingAfter: 80, outlineLevel: 2 },
  { id: 'Quote', name: 'Quote', type: 'paragraph', basedOn: 'Normal', italic: true, color: '#404040' },
  { id: 'IntenseQuote', name: 'Intense Quote', type: 'paragraph', basedOn: 'Normal', italic: true, color: '#2f5496', alignment: 'center' },
  { id: 'NoSpacing', name: 'No Spacing', type: 'paragraph', basedOn: 'Normal', spacingAfter: 0, lineSpacing: 1 },
  { id: 'ListParagraph', name: 'List Paragraph', type: 'paragraph', basedOn: 'Normal' },
]

export const headingLevelOfStyle = (s: Pick<StyleDef, 'name' | 'id' | 'outlineLevel'>): number | null => {
  const m = /^heading\s*([1-6])$/i.exec(s.name) ?? /^Heading([1-6])$/.exec(s.id)
  return m ? Number(m[1]) : null
}

export const displayName = (s: StyleDef) => {
  const lvl = headingLevelOfStyle(s)
  return lvl ? `Heading ${lvl}` : s.name
}

/** The document's paragraph styles plus any built-ins it does not define, in gallery order. */
export function galleryStyles(meta: DocMeta): StyleDef[] {
  const own = meta.styles.filter((s) => s.type === 'paragraph')
  const byName = new Map(own.map((s) => [s.name.toLowerCase(), s]))
  const out: StyleDef[] = []
  for (const b of BUILTIN_STYLES) {
    const mine = byName.get(b.name.toLowerCase()) ?? own.find((s) => s.id === b.id)
    out.push(mine ?? b)
  }
  for (const s of own) if (!out.includes(s)) out.push(s)
  return out
}

export function styleId(name: string, taken: Set<string>): string {
  const base = name.replace(/[^A-Za-z0-9]/g, '').slice(0, 30) || 'Style'
  let id = base
  for (let i = 2; taken.has(id); i++) id = base + i
  return id
}

const hex6 = (c?: string) => (c && /^#[0-9a-f]{6}$/i.test(c) ? c.slice(1).toUpperCase() : null)
const JC: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' }

/** pPr / rPr children SCRIBE manages for a style; everything else in the original style is kept. */
function managed(s: StyleDef) {
  const ppr: string[] = []
  const rpr: string[] = []
  const sp: Record<string, string | number | undefined> = {}
  if (s.spacingBefore != null) sp['w:before'] = Math.round(s.spacingBefore)
  if (s.spacingAfter != null) sp['w:after'] = Math.round(s.spacingAfter)
  if (s.lineSpacing != null) {
    sp['w:line'] = Math.round(s.lineSpacing * 240)
    sp['w:lineRule'] = 'auto'
  }
  if (Object.keys(sp).length) ppr.push(`<w:spacing${Object.entries(sp).map(([k, v]) => ` ${k}="${v}"`).join('')}/>`)
  if (s.alignment) ppr.push(`<w:jc w:val="${JC[s.alignment]}"/>`)
  const lvl = headingLevelOfStyle(s)
  if (s.outlineLevel != null || lvl) ppr.push(`<w:outlineLvl w:val="${s.outlineLevel ?? (lvl! - 1)}"/>`)
  if (s.fontFamily) rpr.push(`<w:rFonts w:ascii="${esc(s.fontFamily)}" w:hAnsi="${esc(s.fontFamily)}" w:eastAsia="${esc(s.fontFamily)}" w:cs="${esc(s.fontFamily)}"/>`)
  if (s.bold != null) rpr.push(s.bold ? '<w:b/><w:bCs/>' : '<w:b w:val="0"/><w:bCs w:val="0"/>')
  if (s.italic != null) rpr.push(s.italic ? '<w:i/><w:iCs/>' : '<w:i w:val="0"/><w:iCs w:val="0"/>')
  if (hex6(s.color)) rpr.push(`<w:color w:val="${hex6(s.color)}"/>`)
  if (s.fontSizePt) rpr.push(`<w:sz w:val="${Math.round(s.fontSizePt * 2)}"/><w:szCs w:val="${Math.round(s.fontSizePt * 2)}"/>`)
  if (s.underline != null) rpr.push(`<w:u w:val="${s.underline ? 'single' : 'none'}"/>`)
  return { ppr, rpr }
}

const MANAGED_PPR = new Set(['spacing', 'jc', 'outlineLvl'])
const MANAGED_RPR = new Set(['rFonts', 'b', 'bCs', 'i', 'iCs', 'color', 'sz', 'szCs', 'u'])

function expand(frags: string[]) {
  // "<w:b/><w:bCs/>" -> two fragments
  return frags.flatMap((f) => f.match(/<w:[A-Za-z]+\b[^>]*\/>/g) ?? [f])
}

/** New <w:style> element for a style created in SCRIBE. */
export function styleElementXml(s: StyleDef): string {
  const { ppr, rpr } = managed(s)
  const lvl = headingLevelOfStyle(s)
  return (
    `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${esc(s.id)}"><w:name w:val="${esc(s.name)}"/>` +
    (s.basedOn ? `<w:basedOn w:val="${esc(s.basedOn)}"/>` : '') +
    `<w:next w:val="${esc(s.next ?? (lvl ? 'Normal' : s.id))}"/><w:qFormat/>` +
    (ppr.length ? `<w:pPr>${orderChildren(expand(ppr), PPR_ORDER)}</w:pPr>` : '') +
    (rpr.length ? `<w:rPr>${orderChildren(expand(rpr), RPR_ORDER)}</w:rPr>` : '') +
    `</w:style>`
  )
}

/**
 * Apply edited and new styles to a parsed styles.xml document in place.
 * Edited styles keep every property SCRIBE does not manage (borders, tabs, numbering...).
 */
export function applyStyleEdits(styleDoc: Document, styles: StyleDef[], builtinXml: (id: string) => string | null): number {
  const root = styleDoc.documentElement
  const existing = new Map(children(root, NS.w, 'style').map((e) => [e.getAttributeNS(NS.w, 'styleId') ?? '', e]))
  let changed = 0
  for (const s of styles) {
    if (!s.modified && !s.custom) continue
    let el = existing.get(s.id)
    if (!el) {
      const xml = s.custom ? styleElementXml(s) : builtinXml(s.id) ?? styleElementXml({ ...s, custom: true })
      const frag = splitFragmentsNs(xml)[0]
      if (!frag) continue
      el = styleDoc.importNode(frag, true) as Element
      root.appendChild(el)
      existing.set(s.id, el)
      if (s.custom && !s.modified) {
        changed++
        continue
      }
    }
    // name for custom styles
    if (s.custom) {
      const name = children(el, NS.w, 'name')[0]
      if (name) name.setAttributeNS(NS.w, 'w:val', s.name)
    }
    const { ppr, rpr } = managed(s)
    rewrite(styleDoc, el, 'pPr', MANAGED_PPR, expand(ppr), PPR_ORDER, s)
    rewrite(styleDoc, el, 'rPr', MANAGED_RPR, expand(rpr), RPR_ORDER, s)
    changed++
  }
  return changed
}

const STYLE_ORDER = ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority', 'semiHidden', 'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose', 'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr']

/** Which managed element each style property controls; elements for undefined properties are kept as they are. */
function ownedBy(s: StyleDef, local: string): boolean {
  switch (local) {
    case 'spacing':
      return s.spacingBefore != null || s.spacingAfter != null || s.lineSpacing != null
    case 'jc':
      return !!s.alignment
    case 'outlineLvl':
      return s.outlineLevel != null || !!headingLevelOfStyle(s)
    case 'rFonts':
      return !!s.fontFamily
    case 'b':
    case 'bCs':
      return s.bold != null
    case 'i':
    case 'iCs':
      return s.italic != null
    case 'color':
      return !!hex6(s.color)
    case 'sz':
    case 'szCs':
      return !!s.fontSizePt
    case 'u':
      return s.underline != null
    default:
      return false
  }
}

function rewrite(doc: Document, style: Element, tag: 'pPr' | 'rPr', managedSet: Set<string>, frags: string[], order: readonly string[], def: StyleDef) {
  const cur = children(style, NS.w, tag)[0]
  const keep: string[] = []
  let spacingAttrs: Record<string, string> = {}
  for (const c of cur ? children(cur) : []) {
    const mine = c.namespaceURI === NS.w && managedSet.has(c.localName) && ownedBy(def, c.localName)
    if (!mine) keep.push(serializeXml(c))
    else if (c.localName === 'spacing') spacingAttrs = Object.fromEntries([...c.attributes].map((a) => [a.name, a.value]))
  }
  // spacing is shared by several properties: merge attributes SCRIBE did not set
  const merged = frags.map((f) => {
    if (!f.startsWith('<w:spacing')) return f
    const mineAttrs = Object.fromEntries([...f.matchAll(/(w:[A-Za-z]+)="([^"]*)"/g)].map((m) => [m[1], m[2]]))
    const all = { ...spacingAttrs, ...mineAttrs }
    if (!('w:line' in mineAttrs) && 'w:line' in spacingAttrs) all['w:lineRule'] = spacingAttrs['w:lineRule'] ?? 'auto'
    return `<w:spacing${Object.entries(all).filter(([k]) => !k.startsWith('xmlns')).map(([k, v]) => ` ${k}="${esc(v)}"`).join('')}/>`
  })
  const all = [...keep, ...merged]
  if (cur) style.removeChild(cur)
  if (!all.length) return
  const xml = `<w:${tag} xmlns:w="${NS.w}">${orderChildren(all, order)}</w:${tag}>`
  const el = doc.importNode(splitFragmentsNs(xml)[0], true) as Element
  const rank = STYLE_ORDER.indexOf(tag)
  const after = children(style).find((c) => STYLE_ORDER.indexOf(c.localName) > rank) ?? null
  style.insertBefore(el, after)
}
