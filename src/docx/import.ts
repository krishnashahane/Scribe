import type { CommentThread, CompatibilityIssue, DocMeta, Footnote, PMNode, ScribeDocument, SectionProps } from '../model/types'
import { SCHEMA_VERSION, defaultPage, emptyMeta, emuToPx } from '../model/types'
import { OoxmlPackage, type Relationship } from './package'
import { NUMFMT_CSS, type Numbering, type StyleSheet, fontFromRFonts, parseNumbering, parseStyles, parseTheme, resolvedStyles, type ThemeFonts } from './styles'
import { NS, attr, child, children, descendants, intAttr, onOff, parseXml, serializeXml, textContent, wChild, wChildren, wVal } from './xml'
import { C_NS, parseChart } from '../objects/chart'
import { DGM_NS, drawingRelId, parseSmartArt } from '../objects/smartart'
import { parseInlineShape } from '../objects/shapes'
import { parseSourcesXml } from '../features/citations'
import type { ObjectInfo } from '../model/extensions'

export interface ImportResult {
  doc: ScribeDocument
  report: CompatibilityIssue[]
  /** The untouched original package; the exporter rebuilds on top of it. */
  original: Uint8Array
}

type Mark = { type: string; attrs?: Record<string, any> }
type Inline = PMNode | { kind: 'pageBreak' }

const HIGHLIGHT: Record<string, string> = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
}
export const HIGHLIGHT_NAMES = HIGHLIGHT

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff',
}

const SIMPLE_FIELDS = /^\s*(PAGE|NUMPAGES|SECTIONPAGES|MERGEFIELD|CITATION)\b/i
const MAX_CUSTOM_XML = 256 * 1024

class Report {
  private map = new Map<string, CompatibilityIssue>()
  add(kind: CompatibilityIssue['kind'], feature: string, detail: string, severity: CompatibilityIssue['severity'] = kind === 'dropped' ? 'warning' : 'info') {
    const k = `${kind}|${feature}`
    const cur = this.map.get(k)
    if (cur) cur.count++
    else this.map.set(k, { kind, feature, detail, severity, count: 1 })
  }
  list() {
    return [...this.map.values()]
  }
}

interface FieldFrame {
  instr: string
  phase: 'instr' | 'result'
  /** raw XML of runs from begin through separate */
  rawBegin: string[]
  /** inlines collected during result phase while still a candidate for a field atom */
  result: Inline[]
  mode: 'candidate' | 'flushed'
  resultMarks: Mark[]
}

interface Ctx {
  pkg: OoxmlPackage
  part: string
  rels: Relationship[]
  sheet: StyleSheet
  theme: ThemeFonts
  numbering: Numbering
  report: Report
  media: Map<string, string>
  activeComments: Set<string>
  fields: FieldFrame[]
  noteLabels: Map<string, string>
  noteCounter: { footnote: number; endnote: number }
  sections: SectionProps[]
  headerFooters: Record<string, PMNode>
  /** while converting header/footer parts, section breaks aren't meaningful */
  inHeaderFooter: boolean
  /** charts and diagrams whose data parts are read after conversion */
  pending: { node: PMNode; rels: Relationship[] }[]
}

export async function importDocx(data: ArrayBuffer | Uint8Array): Promise<ImportResult> {
  const original = data instanceof Uint8Array ? data : new Uint8Array(data)
  const pkg = await OoxmlPackage.open(original)
  const main = await pkg.mainDocumentPath()
  const dir = main.slice(0, main.lastIndexOf('/') + 1)
  const rels = await pkg.rels(main)
  const relOf = (type: string) => rels.find((r) => r.type.endsWith('/' + type) && !r.external)?.target
  const report = new Report()

  const theme = parseTheme(await pkg.xml(relOf('theme') ?? dir + 'theme/theme1.xml'))
  const sheet = parseStyles(await pkg.xml(relOf('styles') ?? dir + 'styles.xml'), theme)
  const numbering = parseNumbering(await pkg.xml(relOf('numbering') ?? dir + 'numbering.xml'))

  const ctx: Ctx = {
    pkg, part: main, rels, sheet, theme, numbering, report,
    media: new Map(), activeComments: new Set(), fields: [], noteLabels: new Map(),
    noteCounter: { footnote: 0, endnote: 0 }, sections: [], headerFooters: {}, inHeaderFooter: false, pending: [],
  }

  const meta: DocMeta = emptyMeta()
  meta.defaults = { fontFamily: sheet.defaults.fontFamily, fontSizePt: sheet.defaults.fontSizePt, lineHeight: sheet.defaults.lineHeight }
  meta.styles = resolvedStyles(sheet)

  const docXml = await pkg.xml(main)
  if (!docXml) throw new Error('Main document part is empty')
  const body = wChild(docXml.documentElement, 'body')
  if (!body) throw new Error('Document has no <w:body>')

  // Notes first so references can be numbered
  meta.footnotes = await importNotes(ctx, relOf('footnotes'), 'footnote')
  meta.endnotes = await importNotes(ctx, relOf('endnotes'), 'endnote')
  meta.comments = await importComments(ctx, relOf('comments'))

  let blocks = await convertBlocks(body, ctx)
  // final body-level sectPr
  const finalSect = wChild(body, 'sectPr')
  ctx.sections.push(await parseSection(finalSect, ctx))
  closeOpenFields(ctx, blocks)
  if (!blocks.length) blocks = [{ type: 'paragraph' }]
  meta.sections = ctx.sections
  meta.headerFooters = ctx.headerFooters
  meta.metadata = await importCoreProps(pkg)
  await resolveObjects(ctx)

  // package-level features
  let customBytes = 0
  for (const p of pkg.listParts()) {
    if (/vbaProject\.bin$/i.test(p)) report.add('preserved', 'Macros (VBA)', 'Macro project kept byte-for-byte; SCRIBE never runs macros.', 'warning')
    if (/\/embeddings\//.test(p)) report.add('preserved', 'Embedded objects (OLE)', 'Embedded files are preserved unchanged.')
    if (/glossary\/document\.xml$/.test(p)) report.add('preserved', 'Building blocks / glossary', 'Kept unchanged.')
    if (/^customXml\/item\d+\.xml$/i.test(p)) {
      const doc = await pkg.xml(p).catch(() => null)
      if (!doc) continue
      const xml = serializeXml(doc.documentElement)
      const bib = parseSourcesXml(xml)
      if (bib) {
        meta.sources = bib.sources
        if (bib.style) meta.citationStyle = bib.style
        if (bib.sources.length) report.add('transformed', 'Bibliography sources', `${bib.sources.length} source(s) imported; edit them from References → Sources.`, 'info')
      } else if (xml.length <= MAX_CUSTOM_XML && customBytes + xml.length <= 4 * MAX_CUSTOM_XML) {
        customBytes += xml.length
        ;(meta.customXml ??= []).push({ path: p, xml })
        report.add('preserved', 'Custom XML parts', 'Kept unchanged; edit them from File → Custom XML data.')
      } else report.add('preserved', 'Custom XML parts', 'Large custom XML parts are kept unchanged but are too big to edit in SCRIBE.')
    }
  }
  if (meta.footnotes.length) report.add('transformed', 'Footnotes', `${meta.footnotes.length} footnote(s) imported. Untouched notes are kept exactly; a note you edit keeps its paragraph and character formatting but rich formatting inside it is simplified.`, 'info')
  if (meta.endnotes.length) report.add('transformed', 'Endnotes', `${meta.endnotes.length} endnote(s) imported. Untouched notes are kept exactly; a note you edit keeps its paragraph and character formatting but rich formatting inside it is simplified.`, 'info')
  if (meta.comments.length) report.add('transformed', 'Comments', 'Comments imported. Threaded replies and rich formatting inside comment text are flattened if you edit comments.')

  return {
    doc: { version: SCHEMA_VERSION, content: { type: 'doc', content: blocks }, meta },
    report: report.list(),
    original,
  }
}

// ---------------------------------------------------------------------------
// blocks

async function convertBlocks(container: Element, ctx: Ctx): Promise<PMNode[]> {
  const out: PMNode[] = []
  for (const c of children(container)) {
    if (c.namespaceURI === NS.w) {
      switch (c.localName) {
        case 'p':
          out.push(...(await convertParagraph(c, ctx)))
          continue
        case 'tbl':
          out.push(await convertTable(c, ctx))
          continue
        case 'sectPr':
        case 'tcPr':
        case 'trPr':
        case 'tblPr':
          continue // handled by caller
        case 'sdt': {
          const gallery = wVal(descendants(c, NS.w, 'docPartGallery')[0])
          if (gallery && /table of contents/i.test(gallery)) {
            ctx.report.add('preserved', 'Table of contents', 'The table of contents is kept exactly as it is until you choose References → Update table, which rebuilds it from the headings.')
            out.push(tocBlock(c))
          } else if (gallery && /bibliograph/i.test(gallery)) {
            const paras = descendants(c, NS.w, 'p').map((p) => textContent(p).trim()).filter(Boolean)
            const headed = descendants(c, NS.w, 'pStyle').some((st) => /heading/i.test(wVal(st) ?? ''))
            const title = headed && paras.length ? paras[0] : 'References'
            const entries = (headed ? paras.slice(1) : paras).map((t) => [{ text: t }])
            ctx.report.add('transformed', 'Bibliography', 'The bibliography is regenerated from the document’s sources when they change.')
            out.push({ type: 'bibliography', attrs: { title, entries } })
          } else {
            ctx.report.add('transformed', 'Content controls', 'Block content controls were unwrapped; their text is editable but control properties (locking, data binding) are not kept.')
            const content = wChild(c, 'sdtContent')
            if (content) out.push(...(await convertBlocks(content, ctx)))
          }
          continue
        }
        case 'customXml':
        case 'ins':
        case 'moveTo':
          out.push(...(await convertBlocks(c, ctx)))
          continue
        case 'bookmarkStart':
        case 'bookmarkEnd':
        case 'proofErr':
        case 'permStart':
        case 'permEnd':
          continue
        case 'commentRangeStart':
          ctx.activeComments.add(attr(c, 'id') || '')
          continue
        case 'commentRangeEnd':
          ctx.activeComments.delete(attr(c, 'id') || '')
          continue
        case 'altChunk':
          ctx.report.add('preserved', 'Embedded alternate content (altChunk)', 'Imported HTML/RTF chunk kept verbatim; it is not editable in SCRIBE.', 'warning')
          out.push(opaqueBlock(c, 'Embedded content', ctx))
          continue
      }
    }
    if (c.namespaceURI === NS.mc && c.localName === 'AlternateContent') {
      out.push(opaqueBlock(c, 'Alternate content', ctx))
      continue
    }
    ctx.report.add('preserved', `<${c.prefix ? c.prefix + ':' : ''}${c.localName}>`, 'Unknown block element preserved verbatim.')
    out.push(opaqueBlock(c, c.localName, ctx))
  }
  return buildLists(out, ctx)
}

/** Imported table of contents: entries are read for display; the original XML is kept for export. */
function tocBlock(sdt: Element): PMNode {
  const content = wChild(sdt, 'sdtContent')
  const entries: { level: number; text: string; page: string }[] = []
  let title = 'Contents'
  const paras = content ? descendants(content, NS.w, 'p') : []
  for (const p of paras) {
    const style = wVal(wChild(wChild(p, 'pPr'), 'pStyle')) ?? ''
    // visible text only (skip field instructions), split at the last tab into entry text and page number
    let text = ''
    for (const r of descendants(p, NS.w, 'r')) {
      for (const k of children(r)) {
        if (k.namespaceURI !== NS.w) continue
        if (k.localName === 't') text += k.textContent ?? ''
        else if (k.localName === 'tab') text += '\t'
      }
    }
    const level = /TOC\s*([1-9])$/i.exec(style)?.[1] ?? /^TOC([1-9])$/i.exec(style)?.[1]
    if (!level) {
      if (text.trim() && !entries.length && /heading/i.test(style)) title = text.trim()
      continue
    }
    const tab = text.lastIndexOf('\t')
    entries.push({ level: Number(level), text: (tab >= 0 ? text.slice(0, tab) : text).replace(/\t/g, ' ').trim(), page: tab >= 0 ? text.slice(tab + 1).trim() : '' })
  }
  return { type: 'toc', attrs: { title, levels: Math.max(3, ...entries.map((e) => e.level)), entries, rawXml: serializeXml(sdt), updated: false } }
}

function opaqueBlock(e: Element, label: string, _ctx: Ctx): PMNode {
  const preview = textContent(e).slice(0, 400)
  return { type: 'opaqueBlock', attrs: { xml: serializeXml(e), label, preview } }
}

const ALIGN: Record<string, string> = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', lowKashida: 'justify', mediumKashida: 'justify', highKashida: 'justify' }

const PPR_HANDLED = new Set(['pStyle', 'jc', 'ind', 'spacing', 'keepNext', 'pageBreakBefore', 'bidi', 'numPr', 'sectPr', 'rPr'])

function paragraphAttrs(pPr: Element | null, ctx: Ctx): { type: string; attrs: Record<string, any>; list?: { numId: string; ilvl: number } } {
  const attrs: Record<string, any> = {}
  let type = 'paragraph'
  const styleId = wVal(wChild(pPr, 'pStyle'))
  if (styleId) {
    attrs.styleId = styleId
    const lvl = ctx.sheet.headingLevels.get(styleId)
    if (lvl) {
      type = 'heading'
      attrs.level = lvl
    }
  }
  const jc = wVal(wChild(pPr, 'jc'))
  if (jc && ALIGN[jc]) attrs.textAlign = ALIGN[jc]
  const ind = wChild(pPr, 'ind')
  if (ind) {
    const left = intAttr(ind, 'left') ?? intAttr(ind, 'start')
    const right = intAttr(ind, 'right') ?? intAttr(ind, 'end')
    const first = intAttr(ind, 'firstLine')
    const hanging = intAttr(ind, 'hanging')
    if (left != null) attrs.indentLeft = left
    if (right != null) attrs.indentRight = right
    if (hanging != null) attrs.indentFirstLine = -hanging
    else if (first != null) attrs.indentFirstLine = first
    const extra = extraAttrs(ind, ['left', 'start', 'right', 'end', 'firstLine', 'hanging'])
    if (extra) attrs.indExtra = extra
  }
  const sp = wChild(pPr, 'spacing')
  if (sp) {
    const b = intAttr(sp, 'before')
    const a = intAttr(sp, 'after')
    const line = intAttr(sp, 'line')
    const rule = attr(sp, 'lineRule') || 'auto'
    if (b != null) attrs.spacingBefore = b
    if (a != null) attrs.spacingAfter = a
    if (line != null) {
      attrs.lineRule = rule
      attrs.lineSpacing = rule === 'auto' ? Math.round((line / 240) * 1000) / 1000 : line
    }
    const extra = extraAttrs(sp, ['before', 'after', 'line', 'lineRule'])
    if (extra) attrs.spacingExtra = extra
  }
  if (onOff(wChild(pPr, 'keepNext'))) attrs.keepNext = true
  if (onOff(wChild(pPr, 'pageBreakBefore'))) attrs.pageBreakBefore = true
  const fill = attr(wChild(pPr, 'shd'), 'fill')
  if (fill && fill !== 'auto' && /^[0-9a-fA-F]{6}$/.test(fill) && wChild(pPr, 'shd')!.attributes.length <= 3) attrs.shading = '#' + fill
  if (onOff(wChild(pPr, 'bidi'))) attrs.direction = 'rtl'
  const rPr = wChild(pPr, 'rPr')
  if (rPr) attrs.markRPr = serializeXml(rPr)
  const extra: string[] = []
  for (const c of children(pPr)) {
    if (c.namespaceURI === NS.w && PPR_HANDLED.has(c.localName)) continue
    if (c.namespaceURI === NS.w && c.localName === 'shd' && attrs.shading) continue
    extra.push(serializeXml(c))
  }
  if (extra.length) attrs.pPrExtra = extra.join('')
  let list: { numId: string; ilvl: number } | undefined
  const numPr = wChild(pPr, 'numPr')
  if (numPr) {
    const numId = wVal(wChild(numPr, 'numId'))
    const ilvl = Number(wVal(wChild(numPr, 'ilvl')) ?? 0)
    if (numId && numId !== '0') list = { numId, ilvl: Math.min(8, Math.max(0, ilvl || 0)) }
  }
  return { type, attrs, list }
}

function extraAttrs(e: Element, known: string[]): Record<string, string> | null {
  const out: Record<string, string> = {}
  for (const a of Array.from(e.attributes)) {
    if (a.namespaceURI === NS.w && known.includes(a.localName)) continue
    if (a.name.startsWith('xmlns')) continue
    out[a.name.includes(':') ? a.name : `w:${a.name}`] = a.value
  }
  return Object.keys(out).length ? out : null
}

async function convertParagraph(p: Element, ctx: Ctx): Promise<PMNode[]> {
  const pPr = wChild(p, 'pPr')
  const { type, attrs, list } = paragraphAttrs(pPr, ctx)
  const inlines = await convertInlines(p, ctx, [])
  // a simple field still open at the paragraph boundary cannot become an atom → keep it verbatim
  for (const f of ctx.fields) if (f.mode === 'candidate') {
    if (f.phase === 'result') flushField(f, inlines)
  }
  // split on page breaks
  const segments: Inline[][] = [[]]
  for (const i of inlines) {
    if ('kind' in i && i.kind === 'pageBreak') segments.push([])
    else segments[segments.length - 1].push(i)
  }
  const out: PMNode[] = []
  segments.forEach((seg, idx) => {
    if (idx > 0) out.push({ type: 'pageBreak' })
    if (idx === 0 && segments.length > 1 && seg.length === 0) return // break at start of paragraph
    const node: PMNode = { type, attrs: { ...attrs } }
    const content = mergeText(seg as PMNode[])
    if (content.length) node.content = content
    if (list) (node as any).__list = list
    out.push(node)
  })
  // section break carried by this paragraph
  const sectPr = wChild(pPr, 'sectPr')
  if (sectPr && !ctx.inHeaderFooter) {
    const sec = await parseSection(sectPr, ctx)
    ctx.sections.push(sec)
    out.push({ type: 'sectionBreak', attrs: { section: ctx.sections.length - 1, breakType: sec.breakType ?? 'nextPage' } })
  }
  return out
}

/** Merge adjacent text nodes with identical marks (keeps documents compact). */
function mergeText(nodes: PMNode[]): PMNode[] {
  const out: PMNode[] = []
  for (const n of nodes) {
    const prev = out[out.length - 1]
    if (n.type === 'text' && prev?.type === 'text' && JSON.stringify(prev.marks ?? []) === JSON.stringify(n.marks ?? [])) prev.text += n.text!
    else if (n.type !== 'text' || n.text) out.push(n)
  }
  return out
}

// ---------------------------------------------------------------------------
// inline content

const RPR_HANDLED = new Set(['b', 'i', 'u', 'strike', 'dstrike', 'color', 'sz', 'rFonts', 'highlight', 'shd', 'vertAlign', 'rStyle', 'caps', 'smallCaps', 'spacing'])

function marksFromRPr(rPr: Element | null, ctx: Ctx): Mark[] {
  if (!rPr) return []
  const marks: Mark[] = []
  const ts: Record<string, any> = {}
  const extra: string[] = []
  for (const c of children(rPr)) {
    if (c.namespaceURI !== NS.w || !RPR_HANDLED.has(c.localName)) {
      extra.push(serializeXml(c))
      continue
    }
    const v = wVal(c)
    switch (c.localName) {
      case 'b':
        if (onOff(c)) marks.push({ type: 'bold' })
        else extra.push(serializeXml(c))
        break
      case 'i':
        if (onOff(c)) marks.push({ type: 'italic' })
        else extra.push(serializeXml(c))
        break
      case 'u':
        if (v && v !== 'none') {
          marks.push({ type: 'underline' })
          if (v !== 'single' || c.attributes.length > 1) extra.push(serializeXml(c)) // keep style (double, wave...) and colour
        } else if (v === 'none') extra.push(serializeXml(c))
        break
      case 'strike':
      case 'dstrike':
        if (onOff(c)) marks.push({ type: 'strike' })
        if (c.localName === 'dstrike' || !onOff(c)) extra.push(serializeXml(c))
        break
      case 'color':
        if (v && v !== 'auto' && /^[0-9a-fA-F]{6}$/.test(v)) ts.color = '#' + v.toLowerCase()
        if (c.attributes.length > 1) extra.push(serializeXml(c)) // themeColor etc.
        break
      case 'sz': {
        const n = parseInt(v || '', 10)
        if (n > 0) ts.fontSize = `${n / 2}pt`
        break
      }
      case 'rFonts': {
        const f = fontFromRFonts(c, ctx.theme)
        if (f && (attr(c, 'ascii') || attr(c, 'hAnsi'))) ts.fontFamily = f
        // theme references and eastAsia/cs fonts are preserved raw
        if (attr(c, 'asciiTheme') || attr(c, 'hAnsiTheme') || attr(c, 'eastAsia') || attr(c, 'cs') || !(attr(c, 'ascii') || attr(c, 'hAnsi'))) {
          extra.push(serializeXml(c))
          if (!ts.fontFamily && f) ts.fontFamilyHint = f
        }
        break
      }
      case 'highlight':
        if (v && HIGHLIGHT[v]) marks.push({ type: 'highlight', attrs: { color: HIGHLIGHT[v] } })
        break
      case 'shd': {
        const fill = attr(c, 'fill')
        if (fill && fill !== 'auto' && /^[0-9a-fA-F]{6}$/.test(fill)) ts.backgroundColor = '#' + fill.toLowerCase()
        else extra.push(serializeXml(c))
        break
      }
      case 'vertAlign':
        if (v === 'superscript') marks.push({ type: 'superscript' })
        else if (v === 'subscript') marks.push({ type: 'subscript' })
        break
      case 'rStyle':
        if (v) ts.charStyle = v
        break
      case 'caps':
      case 'smallCaps':
        if (onOff(c)) ts.caps = c.localName
        else extra.push(serializeXml(c))
        break
      case 'spacing': {
        const n = intAttr(c, 'val')
        if (n) ts.letterSpacing = `${Math.round((n / 20) * 100) / 100}pt`
        break
      }
    }
  }
  if (ts.fontFamilyHint) {
    // render theme font in the editor without exporting it as a direct font
    delete ts.fontFamilyHint
  }
  if (extra.length) ts.rPrExtra = extra.join('')
  if (Object.keys(ts).length) marks.unshift({ type: 'textStyle', attrs: ts })
  return marks
}

function withComments(marks: Mark[], ctx: Ctx): Mark[] {
  if (!ctx.activeComments.size) return marks
  // ProseMirror allows one mark of a type per range unless excludes is empty; CommentMark has excludes ''
  return [...marks, ...[...ctx.activeComments].map((id) => ({ type: 'comment', attrs: { commentId: id } }))]
}

function text(t: string, marks: Mark[]): PMNode {
  const n: PMNode = { type: 'text', text: t }
  if (marks.length) n.marks = marks
  return n
}

async function convertInlines(container: Element, ctx: Ctx, inherited: Mark[]): Promise<Inline[]> {
  const out: Inline[] = []
  const push = (i: Inline | Inline[]) => {
    const list = Array.isArray(i) ? i : [i]
    const frame = ctx.fields[ctx.fields.length - 1]
    if (frame && frame.mode === 'candidate' && frame.phase === 'result') frame.result.push(...list)
    else if (frame && frame.phase === 'instr') {
      /* inline content inside the instruction part is part of rawBegin */
    } else out.push(...list)
  }
  for (const c of children(container)) {
    if (c.namespaceURI === NS.m) {
      ctx.report.add('preserved', 'Equations', 'Equations are kept exactly; double-click one to edit it.')
      push({ type: 'opaqueInline', attrs: { xml: serializeXml(c), label: 'Equation', preview: descendants(c, NS.m, 't').map((t) => t.textContent).join('') || '[equation]', object: { kind: 'equation' } } })
      continue
    }
    if (c.namespaceURI === NS.mc && c.localName === 'AlternateContent') {
      push(opaqueRunLevel(c, ctx))
      continue
    }
    if (c.namespaceURI !== NS.w) {
      ctx.report.add('preserved', `<${c.prefix}:${c.localName}>`, 'Unknown inline element preserved verbatim.')
      push({ type: 'opaqueInline', attrs: { xml: serializeXml(c), label: c.localName, preview: textContent(c) } })
      continue
    }
    switch (c.localName) {
      case 'pPr':
      case 'proofErr':
      case 'permStart':
      case 'permEnd':
      case 'bookmarkEnd':
        break
      case 'r':
        await convertRun(c, ctx, inherited, out, push)
        break
      case 'hyperlink': {
        const rid = attr(c, 'id', NS.r)
        const anchor = attr(c, 'anchor')
        let href: string | null = null
        if (rid) href = ctx.rels.find((r) => r.id === rid)?.target ?? null
        if (anchor) href = (href ?? '') + '#' + anchor
        const marks = href ? [...inherited, { type: 'link', attrs: { href, target: anchor && !rid ? null : '_blank' } }] : inherited
        if (href && !/^(https?:|mailto:|tel:|ftp:|#)/i.test(href)) {
          ctx.report.add('transformed', 'Unsafe or local hyperlinks', 'Hyperlinks with non-web targets (file:, javascript:) were converted to plain text for safety.', 'warning')
          push(await convertInlines(c, ctx, inherited))
        } else push(await convertInlines(c, ctx, marks))
        break
      }
      case 'ins':
      case 'moveTo':
      case 'del':
      case 'moveFrom': {
        const isIns = c.localName === 'ins' || c.localName === 'moveTo'
        if (c.localName.startsWith('move')) ctx.report.add('transformed', 'Tracked moves', 'Tracked moves are shown as an insertion plus a deletion.')
        ctx.report.add('preserved', 'Tracked changes', 'Insertions and deletions are shown and exported as revisions; accept or reject them from the Review tab.')
        const m = { type: isIns ? 'insertion' : 'deletion', attrs: { author: attr(c, 'author') || '', date: attr(c, 'date') || '', revId: attr(c, 'id') } }
        push(await convertInlines(c, ctx, [...inherited, m]))
        break
      }
      case 'smartTag':
      case 'customXml':
      case 'dir':
      case 'bdo':
        push(await convertInlines(c, ctx, inherited))
        break
      case 'sdt': {
        ctx.report.add('transformed', 'Content controls', 'Inline content controls were unwrapped; their text is editable but control properties are not kept.')
        const content = wChild(c, 'sdtContent')
        if (content) push(await convertInlines(content, ctx, inherited))
        break
      }
      case 'fldSimple': {
        const instr = attr(c, 'instr') || ''
        const result = textContent(c)
        if (SIMPLE_FIELDS.test(instr)) push({ type: 'field', attrs: { instr: instr.trim(), result } })
        else {
          ctx.report.add('preserved', 'Fields', 'Fields are preserved with their last calculated result.')
          push({ type: 'opaqueInline', attrs: { xml: serializeXml(c), label: 'Field', preview: result || `{${instr.trim()}}` } })
        }
        break
      }
      case 'bookmarkStart': {
        const name = attr(c, 'name') || ''
        if (name === '_GoBack') break
        ctx.report.add('transformed', 'Bookmarks', 'Bookmarks are kept as named positions; bookmarked text ranges collapse to their start point.')
        push({ type: 'bookmark', attrs: { name, bid: attr(c, 'id') } })
        break
      }
      case 'commentRangeStart':
        ctx.activeComments.add(attr(c, 'id') || '')
        break
      case 'commentRangeEnd':
        ctx.activeComments.delete(attr(c, 'id') || '')
        break
      default:
        ctx.report.add('preserved', `<w:${c.localName}>`, 'Unknown inline element preserved verbatim.')
        push({ type: 'opaqueInline', attrs: { xml: serializeXml(c), label: c.localName, preview: textContent(c) } })
    }
  }
  return out
}

function flushField(f: FieldFrame, out: Inline[]) {
  out.push({ type: 'opaqueInline', attrs: { xml: f.rawBegin.join(''), label: 'Field start', preview: '' } })
  out.push(...f.result)
  f.result = []
  f.mode = 'flushed'
}

function closeOpenFields(ctx: Ctx, _blocks: PMNode[]) {
  if (ctx.fields.length) {
    ctx.report.add('transformed', 'Unterminated fields', 'A field was not properly closed in the source file; its code was dropped and its result kept as text.', 'warning')
    ctx.fields = []
  }
}

/** inRunRPr: when the element came from inside a <w:r>, it is re-wrapped in a run so export stays schema-valid. */
function opaqueRunLevel(e: Element, ctx: Ctx, inRunRPr?: string): PMNode {
  const choice = child(e, NS.mc, 'Choice') ?? e
  const xml = inRunRPr === undefined ? serializeXml(e) : runXml(inRunRPr, serializeXml(e))
  let label = 'Object'
  let object: ObjectInfo | null = null
  const hasTextBox = descendants(e, NS.w, 'txbxContent').length > 0
  const chartRef = descendants(e, C_NS, 'chart')[0]
  const diagram = descendants(e, DGM_NS, 'relIds')[0]
  if (hasTextBox) {
    label = 'Text box'
    object = { kind: 'textbox' }
  } else if (chartRef) label = 'Chart'
  else if (diagram) label = 'SmartArt'
  else if (/wps:wsp|v:shape|v:rect/i.test(xml)) label = 'Shape'
  else if (/chart/i.test(xml)) label = 'Chart'
  const plural: Record<string, string> = { Object: 'Drawing objects', 'Text box': 'Text boxes', Chart: 'Charts', Shape: 'Shapes', SmartArt: 'SmartArt' }
  const editable: Record<string, string> = { 'Text box': 'double-click one to edit its text', Chart: 'double-click one to edit its data', SmartArt: 'double-click one to edit its text' }
  ctx.report.add('preserved', plural[label], `${plural[label]} keep their original position and appearance in the exported file${editable[label] ? `; ${editable[label]}` : '; SCRIBE shows them as placeholders'}.`)
  const node: PMNode = { type: 'opaqueInline', attrs: { xml, label, preview: textContent(choice).slice(0, 200), object } }
  if (chartRef || diagram) ctx.pending.push({ node, rels: ctx.rels })
  return node
}

async function convertRun(r: Element, ctx: Ctx, inherited: Mark[], _out: Inline[], push: (i: Inline | Inline[]) => void) {
  const rPr = wChild(r, 'rPr')
  const base = [...marksFromRPr(rPr, ctx), ...inherited]
  const marks = () => withComments(base, ctx)
  const rPrXml = rPr ? serializeXml(rPr) : ''
  for (const c of children(r)) {
    const frame = ctx.fields[ctx.fields.length - 1]
    // capture raw XML while inside a field instruction
    if (frame && frame.phase === 'instr' && !(c.namespaceURI === NS.w && c.localName === 'fldChar')) {
      frame.rawBegin.push(runXml(rPrXml, serializeXml(c)))
      if (c.namespaceURI === NS.w && c.localName === 'instrText') frame.instr += c.textContent ?? ''
      continue
    }
    if (c.namespaceURI === NS.mc) {
      push(opaqueRunLevel(c, ctx, rPrXml))
      continue
    }
    if (c.namespaceURI !== NS.w) {
      push({ type: 'opaqueInline', attrs: { xml: runXml(rPrXml, serializeXml(c)), label: c.localName, preview: '' } })
      continue
    }
    switch (c.localName) {
      case 'rPr':
      case 'lastRenderedPageBreak':
      case 'commentReference':
      case 'annotationRef':
        break
      case 't':
      case 'delText':
        if (c.textContent) push(text(c.textContent, marks()))
        break
      case 'tab':
        push(text('\t', marks()))
        break
      case 'noBreakHyphen':
        push(text('‑', marks()))
        break
      case 'softHyphen':
        push(text('­', marks()))
        break
      case 'br':
      case 'cr': {
        const t = attr(c, 'type')
        if (t === 'page' && !ctx.inHeaderFooter) push({ kind: 'pageBreak' })
        else if (t === 'column') push({ type: 'columnBreak' })
        else push({ type: 'hardBreak', marks: marks().length ? marks() : undefined } as PMNode)
        break
      }
      case 'drawing':
        push(await convertDrawing(c, ctx, rPrXml))
        break
      case 'pict':
      case 'object': {
        const img = descendants(c, NS.v, 'imagedata')[0]
        const rid = img && attr(img, 'id', NS.r)
        if (rid && c.localName === 'pict' && !/txbx|textbox/i.test(serializeXml(c))) {
          const src = await loadImage(rid, ctx)
          if (src) {
            ctx.report.add('transformed', 'Legacy VML images', 'Legacy VML images were converted to modern inline images.')
            push({ type: 'image', attrs: { src } })
            break
          }
        }
        push(opaqueRunLevel(c, ctx, rPrXml))
        break
      }
      case 'fldChar': {
        const t = attr(c, 'fldCharType')
        if (t === 'begin') {
          // nested field inside a candidate result → the outer one can no longer be a simple atom
          for (const f of ctx.fields) if (f.mode === 'candidate' && f.phase === 'result') {
            const sink: Inline[] = []
            flushField(f, sink)
            const parent = ctx.fields[ctx.fields.indexOf(f) - 1]
            if (parent && parent.mode === 'candidate' && parent.phase === 'result') parent.result.push(...sink)
            else push(sink)
          }
          ctx.fields.push({ instr: '', phase: 'instr', rawBegin: [runXml(rPrXml, serializeXml(c))], result: [], mode: 'candidate', resultMarks: base })
        } else if (t === 'separate' && frame) {
          frame.rawBegin.push(runXml(rPrXml, serializeXml(c)))
          frame.phase = 'result'
          if (!SIMPLE_FIELDS.test(frame.instr)) {
            ctx.report.add('preserved', 'Fields', 'Fields (TOC, REF, cross-references, dates…) are preserved with their last calculated result.')
            const sink: Inline[] = []
            flushField(frame, sink)
            pushOuter(ctx, frame, sink, push)
          }
        } else if (t === 'end' && frame) {
          ctx.fields.pop()
          const endXml = runXml(rPrXml, serializeXml(c))
          if (frame.phase === 'instr') {
            // field without a result (e.g. PAGE with no cached value)
            if (SIMPLE_FIELDS.test(frame.instr)) push({ type: 'field', attrs: { instr: frame.instr.trim(), result: '' } })
            else push({ type: 'opaqueInline', attrs: { xml: frame.rawBegin.join('') + endXml, label: 'Field', preview: `{${frame.instr.trim()}}` } })
          } else if (frame.mode === 'candidate') {
            const resultText = frame.result.map((n) => ('text' in n && (n as PMNode).text) || '').join('')
            push({ type: 'field', attrs: { instr: frame.instr.trim(), result: resultText } })
          } else {
            push({ type: 'opaqueInline', attrs: { xml: endXml, label: 'Field end', preview: '' } })
          }
        }
        break
      }
      case 'footnoteReference':
      case 'endnoteReference': {
        const kind = c.localName === 'footnoteReference' ? 'footnote' : 'endnote'
        const id = attr(c, 'id') || ''
        push({ type: 'noteRef', attrs: { noteId: id, kind, label: ctx.noteLabels.get(kind + id) ?? '' } })
        break
      }
      case 'footnoteRef':
      case 'endnoteRef':
        break
      case 'sym': {
        const ch = attr(c, 'char')
        const code = ch ? parseInt(ch, 16) : NaN
        // Symbol fonts map into the private-use area; keep the original element so the right glyph is rendered
        push({ type: 'opaqueInline', attrs: { xml: runXml(rPrXml, serializeXml(c)), label: 'Symbol', preview: Number.isFinite(code) && code < 0xf000 ? String.fromCodePoint(code) : '' } })
        break
      }
      default:
        ctx.report.add('preserved', `<w:${c.localName}> in run`, 'Unknown run content preserved verbatim.')
        push({ type: 'opaqueInline', attrs: { xml: runXml(rPrXml, serializeXml(c)), label: c.localName, preview: '' } })
    }
  }
}

function pushOuter(ctx: Ctx, _frame: FieldFrame, sink: Inline[], push: (i: Inline | Inline[]) => void) {
  const parent = ctx.fields[ctx.fields.length - 2]
  if (parent && parent.mode === 'candidate' && parent.phase === 'result') parent.result.push(...sink)
  else {
    // temporarily hide the frame so push() writes to the outer sink
    const top = ctx.fields.pop()!
    push(sink)
    ctx.fields.push(top)
  }
}

function runXml(rPrXml: string, inner: string) {
  return `<w:r xmlns:w="${NS.w}">${rPrXml}${inner}</w:r>`
}

async function loadImage(rid: string, ctx: Ctx): Promise<string | null> {
  const rel = ctx.rels.find((r) => r.id === rid)
  if (!rel || rel.external) return null
  if (ctx.media.has(rel.target)) return ctx.media.get(rel.target)!
  const ext = rel.target.split('.').pop()?.toLowerCase() ?? ''
  const mime = IMAGE_MIME[ext]
  if (!mime) return null
  const bytes = await ctx.pkg.bytes(rel.target)
  if (!bytes) return null
  const url = `data:${mime};base64,${toBase64(bytes)}`
  ctx.media.set(rel.target, url)
  return url
}

export function toBase64(bytes: Uint8Array): string {
  let s = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) s += String.fromCharCode(...bytes.subarray(i, i + chunk))
  return btoa(s)
}

async function convertDrawing(d: Element, ctx: Ctx, rPrXml: string): Promise<PMNode> {
  const holder = child(d, NS.wp, 'inline') ?? child(d, NS.wp, 'anchor')
  const blip = descendants(d, NS.a, 'blip')[0]
  const isPicture = !!descendants(d, NS.pic, 'pic')[0]
  const rid = blip && (attr(blip, 'embed', NS.r) || null)
  if (holder && isPicture && rid) {
    const src = await loadImage(rid, ctx)
    if (src) {
      const ext = child(holder, NS.wp, 'extent')
      const cx = Number(ext?.getAttribute('cx') || 0)
      const cy = Number(ext?.getAttribute('cy') || 0)
      const docPr = child(holder, NS.wp, 'docPr')
      const attrs: Record<string, any> = {
        src,
        alt: docPr?.getAttribute('descr') || null,
        title: docPr?.getAttribute('title') || null,
        docPrName: docPr?.getAttribute('name') || null,
        width: cx ? Math.round(emuToPx(cx)) : null,
        height: cy ? Math.round(emuToPx(cy)) : null,
      }
      if (holder.localName === 'anchor') {
        attrs.anchorXml = serializeXml(holder)
        const h = descendants(holder, NS.wp, 'positionH')[0]
        const align = h && child(h, NS.wp, 'align')?.textContent
        attrs.float = align === 'right' ? 'right' : align === 'center' ? 'center' : 'left'
        ctx.report.add('transformed', 'Floating images', 'Floating images keep their original anchor/wrap settings on export; SCRIBE previews them floated left or right.')
      }
      return { type: 'image', attrs }
    }
    ctx.report.add('preserved', 'Unsupported image formats', 'Images in formats browsers cannot show (EMF/WMF/TIFF) are preserved and shown as placeholders.', 'warning')
  }
  if (holder && !isPicture && holder.localName === 'inline') {
    const shape = parseInlineShape(holder)
    if (shape) return { type: 'shape', attrs: { ...shape } }
  }
  if (holder && !isPicture) return opaqueRunLevel(d, ctx, rPrXml)
  return { type: 'opaqueInline', attrs: { xml: runXml(rPrXml, serializeXml(d)), label: 'Image', preview: '[image]' } }
}

// ---------------------------------------------------------------------------
// lists

function listKind(numId: string, ilvl: number, ctx: Ctx): { type: 'bulletList' | 'orderedList'; css: string | null } {
  const lvl = ctx.numbering.levels.get(numId)?.get(ilvl)
  if (!lvl || lvl.numFmt === 'bullet') return { type: 'bulletList', css: null }
  return { type: 'orderedList', css: NUMFMT_CSS[lvl.numFmt] ?? 'decimal' }
}

function buildLists(blocks: PMNode[], ctx: Ctx): PMNode[] {
  const out: PMNode[] = []
  let stack: { level: number; list: PMNode; numId: string }[] = []
  for (const b of blocks) {
    const info = (b as any).__list as { numId: string; ilvl: number } | undefined
    if (!info) {
      stack = []
      out.push(b)
      continue
    }
    delete (b as any).__list
    const kind = listKind(info.numId, info.ilvl, ctx)
    while (stack.length && stack[stack.length - 1].level > info.ilvl) stack.pop()
    const top = stack[stack.length - 1]
    if (top && top.level === info.ilvl && (top.list.type !== kind.type || (info.ilvl === 0 && top.numId !== info.numId))) stack.pop()
    if (!stack.length || stack[stack.length - 1].level < info.ilvl) {
      const list: PMNode = { type: kind.type, attrs: { numId: info.numId, listStyle: kind.css }, content: [] }
      if (kind.type === 'orderedList') {
        const start = ctx.numbering.levels.get(info.numId)?.get(info.ilvl)?.start ?? 1
        list.attrs!.start = start
      }
      if (!stack.length) out.push(list)
      else {
        const parent = stack[stack.length - 1].list
        let item = parent.content![parent.content!.length - 1]
        if (!item) {
          item = { type: 'listItem', content: [{ type: 'paragraph' }] }
          parent.content!.push(item)
        }
        item.content!.push(list)
      }
      stack.push({ level: info.ilvl, list, numId: info.numId })
    }
    const cur = stack[stack.length - 1].list
    // list items hold paragraphs; headings in lists are kept as paragraphs with their style id
    const para: PMNode = b.type === 'heading' ? { ...b, type: 'paragraph', attrs: { ...b.attrs, level: undefined } } : b
    if (para.attrs) delete para.attrs.level
    cur.content!.push({ type: 'listItem', content: [para] })
  }
  return out
}

// ---------------------------------------------------------------------------
// tables

const TCPR_HANDLED = new Set(['gridSpan', 'vMerge', 'shd', 'vAlign', 'tcW'])
const VALIGN: Record<string, string> = { top: 'top', center: 'middle', bottom: 'bottom' }

async function convertTable(tbl: Element, ctx: Ctx): Promise<PMNode> {
  const tblPr = wChild(tbl, 'tblPr')
  const grid = wChildren(wChild(tbl, 'tblGrid'), 'gridCol').map((g) => intAttr(g, 'w') ?? 0)
  const tableAttrs: Record<string, any> = {}
  const gridEl = wChild(tbl, 'tblGrid')
  if (gridEl) tableAttrs.gridXml = serializeXml(gridEl)
  const style = wVal(wChild(tblPr, 'tblStyle'))
  if (style) tableAttrs.tblStyle = style
  const tblExtra = children(tblPr).filter((c) => !(c.namespaceURI === NS.w && c.localName === 'tblStyle')).map(serializeXml)
  if (tblExtra.length) tableAttrs.tblPrExtra = tblExtra.join('')
  const rows: PMNode[] = []
  // origin cells for vertical merges keyed by grid column
  const vOrigin = new Map<number, PMNode>()
  const rowEls: Element[] = []
  for (const c of children(tbl)) {
    if (isW(c, 'tr')) rowEls.push(c)
    else if (isW(c, 'sdt') || isW(c, 'customXml')) rowEls.push(...wChildren(wChild(c, 'sdtContent') ?? c, 'tr'))
  }
  for (const tr of rowEls) {
    const trPr = wChild(tr, 'trPr')
    const row: PMNode = { type: 'tableRow', attrs: {}, content: [] }
    if (trPr) row.attrs!.trPrExtra = serializeXml(trPr)
    let col = intAttr(wChild(trPr, 'gridBefore'), 'val') ?? 0
    if (col) ctx.report.add('transformed', 'Irregular table grids', 'Rows with skipped grid columns (gridBefore/gridAfter) may render misaligned in SCRIBE; widths are preserved.')
    const isHeader = onOff(wChild(trPr, 'tblHeader'))
    for (const tc of [...wChildren(tr, 'tc'), ...wChildren(wChild(tr, 'sdt'), 'sdtContent').flatMap((s) => wChildren(s, 'tc'))]) {
      const tcPr = wChild(tc, 'tcPr')
      const span = Math.max(1, intAttr(wChild(tcPr, 'gridSpan'), 'val') ?? 1)
      const vm = wChild(tcPr, 'vMerge')
      const vmVal = vm ? attr(vm, 'val') ?? 'continue' : null
      if (vmVal === 'continue') {
        const origin = vOrigin.get(col)
        if (origin) {
          origin.attrs!.rowspan = (origin.attrs!.rowspan ?? 1) + 1
          col += span
          continue
        }
      }
      const widths: number[] = []
      for (let i = 0; i < span; i++) widths.push(grid[col + i] ? Math.round((grid[col + i] / 1440) * 96) : 0)
      const attrs: Record<string, any> = { colspan: span, rowspan: 1, colwidth: widths.every((w) => w > 0) ? widths : null }
      const fill = attr(wChild(tcPr, 'shd'), 'fill')
      if (fill && fill !== 'auto' && /^[0-9a-fA-F]{6}$/.test(fill)) attrs.background = '#' + fill.toLowerCase()
      const va = wVal(wChild(tcPr, 'vAlign'))
      if (va && VALIGN[va]) attrs.verticalAlign = VALIGN[va]
      const extra = children(tcPr).filter((c) => !(c.namespaceURI === NS.w && TCPR_HANDLED.has(c.localName))).map(serializeXml)
      if (wChild(tcPr, 'shd') && !attrs.background) extra.push(serializeXml(wChild(tcPr, 'shd')!))
      if (extra.length) attrs.tcPrExtra = extra.join('')
      let content = await convertBlocks(tc, ctx)
      content = content.filter((b) => b.type !== 'sectionBreak')
      if (!content.length) content = [{ type: 'paragraph' }]
      const cell: PMNode = { type: isHeader ? 'tableHeader' : 'tableCell', attrs, content }
      row.content!.push(cell)
      for (let i = 0; i < span; i++) vOrigin.delete(col + i)
      if (vmVal === 'restart') vOrigin.set(col, cell)
      col += span
    }
    if (!row.content!.length) row.content!.push({ type: 'tableCell', attrs: { colspan: 1, rowspan: 1 }, content: [{ type: 'paragraph' }] })
    rows.push(row)
  }
  if (!rows.length) rows.push({ type: 'tableRow', content: [{ type: 'tableCell', content: [{ type: 'paragraph' }] }] })
  return { type: 'table', attrs: tableAttrs, content: rows }
}

function isW(e: Element, local: string) {
  return e.namespaceURI === NS.w && e.localName === local
}

// ---------------------------------------------------------------------------
// sections, headers, footers

async function parseSection(sectPr: Element | null, ctx: Ctx): Promise<SectionProps> {
  const page = defaultPage('Letter')
  const sec: SectionProps = { page }
  if (!sectPr) return sec
  sec.rawXml = serializeXml(sectPr)
  const pgSz = wChild(sectPr, 'pgSz')
  if (pgSz) {
    page.width = intAttr(pgSz, 'w') ?? page.width
    page.height = intAttr(pgSz, 'h') ?? page.height
    page.orientation = attr(pgSz, 'orient') === 'landscape' ? 'landscape' : page.width > page.height ? 'landscape' : 'portrait'
  }
  const m = wChild(sectPr, 'pgMar')
  if (m) {
    page.marginTop = Math.abs(intAttr(m, 'top') ?? page.marginTop)
    page.marginBottom = Math.abs(intAttr(m, 'bottom') ?? page.marginBottom)
    page.marginLeft = intAttr(m, 'left') ?? intAttr(m, 'start') ?? page.marginLeft
    page.marginRight = intAttr(m, 'right') ?? intAttr(m, 'end') ?? page.marginRight
    page.marginHeader = intAttr(m, 'header') ?? page.marginHeader
    page.marginFooter = intAttr(m, 'footer') ?? page.marginFooter
  }
  sec.breakType = wVal(wChild(sectPr, 'type')) || 'nextPage'
  sec.titlePage = !!onOff(wChild(sectPr, 'titlePg'))
  const cols = intAttr(wChild(sectPr, 'cols'), 'num')
  if (cols && cols > 1) {
    sec.columns = cols
    ctx.report.add('transformed', 'Multi-column sections', 'Columns are shown while editing and kept on export; within a column section, paragraphs, lists and tables move to the next column whole rather than splitting.')
  }
  const start = intAttr(wChild(sectPr, 'pgNumType'), 'start')
  if (start != null) sec.pageNumberStart = start
  for (const ref of [...wChildren(sectPr, 'headerReference'), ...wChildren(sectPr, 'footerReference')]) {
    const isHeader = ref.localName === 'headerReference'
    const type = attr(ref, 'type') || 'default'
    const rid = attr(ref, 'id', NS.r)
    const rel = ctx.rels.find((r) => r.id === rid)
    if (!rel) continue
    if (type === 'even') {
      ctx.report.add('preserved', 'Even-page headers/footers', 'Even-page headers and footers are preserved unchanged; SCRIBE shows the default header on all pages.')
      continue
    }
    if (!ctx.headerFooters[rel.target]) ctx.headerFooters[rel.target] = await importHeaderFooter(rel.target, ctx)
    const key = (isHeader ? 'header' : 'footer') + (type === 'first' ? 'First' : 'Default')
    ;(sec as any)[key] = rel.target
  }
  return sec
}

/** Read the data parts behind charts and diagrams so they can be previewed and edited. */
async function resolveObjects(ctx: Ctx) {
  const text = async (path: string) => {
    const d = await ctx.pkg.xml(path).catch(() => null)
    return d ? serializeXml(d.documentElement) : null
  }
  for (const { node, rels } of ctx.pending) {
    try {
      const el = parseXml(String(node.attrs!.xml), 'object').documentElement
      const ext = descendants(el, NS.wp, 'extent')[0]
      const width = Math.round(emuToPx(Number(ext?.getAttribute('cx') || 0))) || 480
      const height = Math.round(emuToPx(Number(ext?.getAttribute('cy') || 0))) || 280
      const chartRef = descendants(el, C_NS, 'chart')[0]
      if (chartRef) {
        const rel = rels.find((r) => r.id === chartRef.getAttributeNS(NS.r, 'id') && !r.external)
        const xml = rel && (await text(rel.target))
        if (!xml) continue
        const chart = parseChart(xml, rel.target)
        node.attrs!.object = { kind: 'chart', chart, width, height }
        node.attrs!.preview = chart.title || 'Chart'
        continue
      }
      const ids = descendants(el, DGM_NS, 'relIds')[0]
      if (ids) {
        const rel = rels.find((r) => r.id === ids.getAttributeNS(NS.r, 'dm') && !r.external)
        const xml = rel && (await text(rel.target))
        if (!xml) continue
        const drawId = drawingRelId(xml)
        const drawRel = drawId ? rels.find((r) => r.id === drawId && !r.external) : undefined
        const smartArt = parseSmartArt(xml, rel.target, drawRel?.target)
        node.attrs!.object = { kind: 'smartart', smartArt }
        node.attrs!.preview = smartArt.items.map((i) => i.text).filter(Boolean).join(' · ').slice(0, 200) || 'SmartArt'
      }
    } catch {
      ctx.report.add('preserved', 'Unreadable chart or diagram data', 'The object is kept exactly but its data could not be read for editing.', 'warning')
    }
  }
}

async function importHeaderFooter(path: string, ctx: Ctx): Promise<PMNode> {
  const xml = await ctx.pkg.xml(path)
  if (!xml) return { type: 'doc', content: [{ type: 'paragraph' }] }
  const sub: Ctx = { ...ctx, part: path, rels: await ctx.pkg.rels(path), inHeaderFooter: true, fields: [], activeComments: new Set() }
  const blocks = await convertBlocks(xml.documentElement, sub)
  return { type: 'doc', content: blocks.length ? blocks : [{ type: 'paragraph' }] }
}

async function importNotes(ctx: Ctx, path: string | undefined, kind: 'footnote' | 'endnote'): Promise<Footnote[]> {
  if (!path) return []
  const xml = await ctx.pkg.xml(path)
  if (!xml) return []
  const out: Footnote[] = []
  for (const n of wChildren(xml.documentElement, kind)) {
    const t = attr(n, 'type')
    if (t === 'separator' || t === 'continuationSeparator' || t === 'continuationNotice') continue
    const id = attr(n, 'id') || ''
    ctx.noteCounter[kind]++
    ctx.noteLabels.set(kind + id, String(ctx.noteCounter[kind]))
    out.push({ id, text: wChildren(n, 'p').map(textContent).join('\n').trim(), rawXml: serializeXml(n) })
  }
  return out
}

async function importComments(ctx: Ctx, path: string | undefined): Promise<CommentThread[]> {
  if (!path) return []
  const xml = await ctx.pkg.xml(path)
  if (!xml) return []
  const out: CommentThread[] = []
  for (const c of wChildren(xml.documentElement, 'comment')) {
    out.push({
      id: attr(c, 'id') || '',
      author: attr(c, 'author') || 'Unknown',
      initials: attr(c, 'initials') || undefined,
      date: attr(c, 'date') || '',
      text: wChildren(c, 'p').map(textContent).join('\n'),
      pristine: true,
    })
  }
  return out
}

async function importCoreProps(pkg: OoxmlPackage) {
  const rels = await pkg.rels('')
  const path = rels.find((r) => r.type.endsWith('/core-properties'))?.target ?? 'docProps/core.xml'
  const xml = await pkg.xml(path).catch(() => null)
  const meta: Record<string, string> = {}
  if (!xml) return meta
  const get = (ns: string, n: string) => xml.getElementsByTagNameNS(ns, n)[0]?.textContent ?? undefined
  const set = (k: string, v: string | undefined) => {
    if (v) meta[k] = v
  }
  set('title', get(NS.dc, 'title'))
  set('subject', get(NS.dc, 'subject'))
  set('creator', get(NS.dc, 'creator'))
  set('description', get(NS.dc, 'description'))
  set('keywords', get(NS.cp, 'keywords'))
  set('lastModifiedBy', get(NS.cp, 'lastModifiedBy'))
  set('created', get(NS.dcterms, 'created'))
  set('modified', get(NS.dcterms, 'modified'))
  return meta
}
