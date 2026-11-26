import JSZip from 'jszip'
import type { CompatibilityIssue, DocMeta, PMNode, ScribeDocument, SectionProps } from '../model/types'
import { defaultPage, pxToEmu, pxToTwip, twipToPx } from '../model/types'
import { parseNumbering } from './styles'
import { applyChart, chartRunXml, newChartXml, type ChartData } from '../objects/chart'
import { applySmartArtData, applySmartArtDrawing, type SmartArtData } from '../objects/smartart'
import { shapeRunXml, WPS_NS, type ShapeAttrs } from '../objects/shapes'
import { BIB_NS, sourcesXml } from '../features/citations'
import { applyStyleEdits } from '../features/styles'
import { HIGHLIGHT_NAMES } from './import'
import { resolvePath } from './package'
import {
  NS, PPR_ORDER, RPR_ORDER, SECTPR_ORDER, TBLPR_ORDER, TCPR_ORDER, attr, children, el, esc, orderChildren, parseXml,
  serializeXml, stripRedundantNs, wChildren,
} from './xml'

export interface ExportResult {
  data: Uint8Array
  report: CompatibilityIssue[]
}

const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  core: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  app: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
}

const CT = {
  document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  styles: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  settings: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  comments: 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
}

const IMG_EXT: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/tiff': 'tiff' }

const ROOT_NS: Record<string, string> = {
  w: NS.w, r: NS.r, wp: NS.wp, a: NS.a, pic: NS.pic, mc: NS.mc, m: NS.m, wps: WPS_NS,
}
const ROOT_NS_ATTRS = Object.entries(ROOT_NS).map(([p, u]) => `xmlns:${p}="${u}"`).join(' ')

// ---------------------------------------------------------------------------
// package helpers

class Rels {
  private items: { id: string; type: string; target: string; mode?: string }[] = []
  private n = 1
  constructor(public path: string, xml: string | null) {
    if (xml) {
      const d = parseXml(xml, path)
      for (const r of children(d.documentElement, NS.rel, 'Relationship')) {
        this.items.push({ id: r.getAttribute('Id')!, type: r.getAttribute('Type')!, target: r.getAttribute('Target')!, mode: r.getAttribute('TargetMode') || undefined })
      }
    }
  }
  add(type: string, target: string, external = false): string {
    const existing = this.items.find((r) => r.type === type && r.target === target && !!r.mode === external)
    if (existing) return existing.id
    let id: string
    do id = `rIdS${this.n++}`
    while (this.items.some((r) => r.id === id))
    this.items.push({ id, type, target, mode: external ? 'External' : undefined })
    return id
  }
  find(type: string) {
    return this.items.find((r) => r.type === type)
  }
  removeType(type: string) {
    const removed = this.items.filter((r) => r.type === type)
    this.items = this.items.filter((r) => r.type !== type)
    return removed
  }
  xml() {
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rel}">` +
      this.items.map((r) => el('Relationship', { Id: r.id, Type: r.type, Target: r.target, TargetMode: r.mode })).join('') +
      `</Relationships>`
    )
  }
}

class ContentTypes {
  defaults = new Map<string, string>()
  overrides = new Map<string, string>()
  constructor(xml: string | null) {
    if (xml) {
      const d = parseXml(xml, '[Content_Types].xml')
      for (const x of children(d.documentElement, NS.ct, 'Default')) this.defaults.set(x.getAttribute('Extension')!.toLowerCase(), x.getAttribute('ContentType')!)
      for (const x of children(d.documentElement, NS.ct, 'Override')) this.overrides.set(x.getAttribute('PartName')!, x.getAttribute('ContentType')!)
    } else {
      this.defaults.set('rels', CT.rels)
      this.defaults.set('xml', 'application/xml')
    }
  }
  ensureDefault(ext: string, type: string) {
    if (!this.defaults.has(ext)) this.defaults.set(ext, type)
  }
  override(part: string, type: string) {
    this.overrides.set('/' + part, type)
  }
  remove(part: string) {
    this.overrides.delete('/' + part)
  }
  xml() {
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${NS.ct}">` +
      [...this.defaults].map(([e, t]) => el('Default', { Extension: e, ContentType: t })).join('') +
      [...this.overrides].map(([p, t]) => el('Override', { PartName: p, ContentType: t })).join('') +
      `</Types>`
    )
  }
}

class Report {
  private map = new Map<string, CompatibilityIssue>()
  add(kind: CompatibilityIssue['kind'], feature: string, detail: string, severity: CompatibilityIssue['severity'] = 'warning') {
    const k = kind + feature
    const c = this.map.get(k)
    if (c) c.count++
    else this.map.set(k, { kind, feature, detail, severity, count: 1 })
  }
  list() {
    return [...this.map.values()]
  }
}

// ---------------------------------------------------------------------------
// numbering

interface ListAlloc {
  bulletNumId: string | null
  orderedAbstract: Map<string, string>
  bulletAbstract: string | null
  newAbstracts: string[]
  newNums: string[]
  nextNumId: number
  nextAbstractId: number
  existing: ReturnType<typeof parseNumbering>
}

const CSS_TO_FMT: Record<string, string[]> = {
  decimal: ['decimal', 'lowerLetter', 'lowerRoman'],
  'lower-alpha': ['lowerLetter', 'lowerRoman', 'decimal'],
  'upper-alpha': ['upperLetter', 'decimal', 'lowerLetter'],
  'lower-roman': ['lowerRoman', 'lowerLetter', 'decimal'],
  'upper-roman': ['upperRoman', 'upperLetter', 'decimal'],
}

function abstractXml(id: string, kind: 'bullet' | string): string {
  const lvls: string[] = []
  const bullets = ['\u2022', '\u25E6', '\u25AA']
  for (let l = 0; l < 9; l++) {
    const fmt = kind === 'bullet' ? 'bullet' : (CSS_TO_FMT[kind] ?? CSS_TO_FMT.decimal)[l % 3]
    const text = kind === 'bullet' ? bullets[l % 3] : `%${l + 1}.`
    const rPr = kind === 'bullet' ? el('w:rPr', {}, el('w:rFonts', { 'w:ascii': 'Arial', 'w:hAnsi': 'Arial', 'w:hint': 'default' })) : ''
    lvls.push(
      el('w:lvl', { 'w:ilvl': l }, el('w:start', { 'w:val': 1 }) + el('w:numFmt', { 'w:val': fmt }) + el('w:lvlText', { 'w:val': text }) + el('w:lvlJc', { 'w:val': 'left' }) +
        el('w:pPr', {}, el('w:ind', { 'w:left': 720 * (l + 1), 'w:hanging': 360 })) + rPr),
    )
  }
  return el('w:abstractNum', { 'w:abstractNumId': id }, el('w:multiLevelType', { 'w:val': 'hybridMultilevel' }) + lvls.join(''))
}

function allocList(alloc: ListAlloc, node: PMNode): string {
  const numId = node.attrs?.numId as string | null
  if (numId && alloc.existing.levels.has(numId)) {
    const fmt = alloc.existing.levels.get(numId)!.get(0)?.numFmt
    const isBullet = fmt === 'bullet'
    if (isBullet === (node.type === 'bulletList')) return numId
  } else if (numId && alloc.newNums.some((x) => x.includes(`w:numId="${numId}"`))) return numId
  if (node.type === 'bulletList') {
    if (!alloc.bulletAbstract) {
      alloc.bulletAbstract = String(alloc.nextAbstractId++)
      alloc.newAbstracts.push(abstractXml(alloc.bulletAbstract, 'bullet'))
    }
    if (!alloc.bulletNumId) {
      alloc.bulletNumId = String(alloc.nextNumId++)
      alloc.newNums.push(el('w:num', { 'w:numId': alloc.bulletNumId }, el('w:abstractNumId', { 'w:val': alloc.bulletAbstract })))
    }
    return alloc.bulletNumId
  }
  const css = (node.attrs?.listStyle as string) || 'decimal'
  let abs = alloc.orderedAbstract.get(css)
  if (!abs) {
    abs = String(alloc.nextAbstractId++)
    alloc.orderedAbstract.set(css, abs)
    alloc.newAbstracts.push(abstractXml(abs, css))
  }
  const id = String(alloc.nextNumId++)
  const start = Number(node.attrs?.start ?? 1) || 1
  alloc.newNums.push(
    el('w:num', { 'w:numId': id }, el('w:abstractNumId', { 'w:val': abs }) + el('w:lvlOverride', { 'w:ilvl': 0 }, el('w:startOverride', { 'w:val': start }))),
  )
  return id
}

// ---------------------------------------------------------------------------
// serialization context

interface Ctx {
  zip: JSZip
  ct: ContentTypes
  rels: Rels
  partDir: string
  report: Report
  meta: DocMeta
  lists: ListAlloc
  comments: Set<string>
  /** comment ids that actually got anchors */
  anchored: Set<string>
  ids: { bookmark: number; docPr: number }
  media: Map<string, string>
  inHeaderFooter: boolean
  headingStyle: (level: number) => string
  styleIsHeading: (id: string) => boolean
  usedStyles: Set<string>
  textWidthTwips: number
  /** chart / diagram parts already rewritten in this export */
  editedParts: Set<string>
}

const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'

/** Write a chart created in SCRIBE as a new chart part and return the run that places it. */
function newChartRun(obj: any, ctx: Ctx, rPr: string): string {
  const chart = obj.chart as ChartData
  let n = 1
  while (ctx.zip.file(`word/charts/chart${n}.xml`)) n++
  const part = `word/charts/chart${n}.xml`
  ctx.zip.file(part, newChartXml(chart))
  ctx.ct.override(part, CT_CHART)
  const target = ctx.partDir === 'word/' ? `charts/chart${n}.xml` : `${'../'.repeat(ctx.partDir.split('/').filter(Boolean).length)}word/charts/chart${n}.xml`
  const rid = ctx.rels.add(REL_CHART, target)
  return chartRunXml(rid, ctx.ids.docPr++, Number(obj.width) || 480, Number(obj.height) || 288, rPr)
}

/** Rewrite the data parts behind an edited chart or diagram. */
async function writeObjectEdits(obj: any, ctx: Ctx) {
  if (obj?.kind === 'chart' && obj.chart?.dirty) {
    const chart = obj.chart as ChartData
    if (ctx.editedParts.has(chart.part)) return
    const f = ctx.zip.file(chart.part)
    if (!f) return ctx.report.add('dropped', 'Chart edits', 'A chart’s data part was missing from the original file, so its edits could not be saved.')
    ctx.zip.file(chart.part, applyChart(await f.async('string'), chart))
    ctx.editedParts.add(chart.part)
    ctx.report.add('transformed', 'Edited charts', 'Edited chart data is stored inside the chart itself; the spreadsheet copy embedded in the original file is no longer linked.', 'info')
  }
  if (obj?.kind === 'smartart' && obj.smartArt?.dirty) {
    const sa = obj.smartArt as SmartArtData
    if (ctx.editedParts.has(sa.dataPart)) return
    const data = ctx.zip.file(sa.dataPart)
    if (!data) return ctx.report.add('dropped', 'SmartArt edits', 'A diagram’s data part was missing from the original file, so its edits could not be saved.')
    const dataXml = await data.async('string')
    const drawing = sa.drawingPart ? ctx.zip.file(sa.drawingPart) : null
    if (drawing) ctx.zip.file(sa.drawingPart!, applySmartArtDrawing(await drawing.async('string'), dataXml, sa.items))
    ctx.zip.file(sa.dataPart, applySmartArtData(dataXml, sa.items))
    ctx.editedParts.add(sa.dataPart)
  }
}

type Mark = { type: string; attrs?: Record<string, any> }

function markOf(marks: Mark[] | undefined, type: string) {
  return marks?.find((m) => m.type === type)
}

function splitFragments(xml: string | null | undefined): Element[] {
  if (!xml) return []
  try {
    const d = parseXml(`<x xmlns:w="${NS.w}" xmlns:r="${NS.r}">${xml}</x>`, 'fragment')
    return children(d.documentElement)
  } catch {
    return []
  }
}

function rPrXml(marks: Mark[] | undefined, ctx: Ctx, extraRun: string[] = []): string {
  const parts: string[] = [...extraRun]
  const ts = markOf(marks, 'textStyle')?.attrs ?? {}
  const raw = splitFragments(ts.rPrExtra)
  const rawBy = (n: string) => raw.find((e) => e.namespaceURI === NS.w && e.localName === n)
  const drop = new Set<Element>()
  const has = (t: string) => !!markOf(marks, t)
  const link = markOf(marks, 'link')

  if (ts.charStyle) parts.push(el('w:rStyle', { 'w:val': ts.charStyle }))
  else if (link) {
    parts.push(el('w:rStyle', { 'w:val': 'Hyperlink' }))
    ctx.usedStyles.add('Hyperlink')
  }

  const rawFonts = rawBy('rFonts')
  if (ts.fontFamily && rawFonts && attr(rawFonts, 'ascii') === ts.fontFamily) {
    /* unchanged font: the raw element (with eastAsia/cs/theme attributes) is re-emitted as-is */
  } else if (ts.fontFamily) {
    const keep: Record<string, string> = {}
    if (rawFonts) {
      for (const a of ['eastAsia', 'cs', 'hint']) {
        const v = attr(rawFonts, a)
        if (v) keep['w:' + a] = v
      }
      drop.add(rawFonts)
    }
    parts.push(el('w:rFonts', { 'w:ascii': ts.fontFamily, 'w:hAnsi': ts.fontFamily, ...keep }))
  }
  const onOffMark = (mark: string, tag: string) => {
    const r = rawBy(tag)
    if (has(mark)) {
      if (r) drop.add(r)
      parts.push(el('w:' + tag))
    }
  }
  onOffMark('bold', 'b')
  onOffMark('italic', 'i')
  if (ts.caps) {
    for (const t of ['caps', 'smallCaps']) if (rawBy(t)) drop.add(rawBy(t)!)
    parts.push(el('w:' + ts.caps))
  }
  // strike / dstrike
  const rawD = rawBy('dstrike')
  const rawS = rawBy('strike')
  if (has('strike')) {
    if (rawS) drop.add(rawS)
    if (!rawD) parts.push(el('w:strike'))
  } else if (rawD && rawD.getAttribute('w:val') !== 'false' && rawD.getAttribute('w:val') !== '0') drop.add(rawD)
  // colour
  const rawColor = rawBy('color')
  const color = ts.color
  if (color) {
    const hex = String(color).replace('#', '').toUpperCase()
    if (rawColor && (attr(rawColor, 'val') || '').toUpperCase() === hex) {
      /* unchanged: keep original element with theme attributes */
    } else {
      if (rawColor) drop.add(rawColor)
      if (/^[0-9A-F]{6}$/.test(hex)) parts.push(el('w:color', { 'w:val': hex }))
    }
  } else if (rawColor && /^[0-9a-fA-F]{6}$/.test(attr(rawColor, 'val') || '')) drop.add(rawColor) // user removed colour
  if (ts.letterSpacing) {
    const pt = parseFloat(ts.letterSpacing)
    if (Number.isFinite(pt)) parts.push(el('w:spacing', { 'w:val': Math.round(pt * 20) }))
  }
  if (ts.fontSize) {
    const pt = parseFloat(ts.fontSize)
    if (Number.isFinite(pt) && pt > 0) {
      const hp = Math.round(pt * 2)
      parts.push(el('w:sz', { 'w:val': hp }))
      if (!rawBy('szCs')) parts.push(el('w:szCs', { 'w:val': hp }))
    }
  }
  const hl = markOf(marks, 'highlight')
  if (hl) {
    const c = String(hl.attrs?.color || '#ffff00').toLowerCase()
    const name = Object.entries(HIGHLIGHT_NAMES).find(([, v]) => v === c)?.[0]
    if (name) parts.push(el('w:highlight', { 'w:val': name }))
    else parts.push(el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': c.replace('#', '').toUpperCase() }))
  }
  // underline
  const rawU = rawBy('u')
  if (has('underline')) {
    if (rawU && attr(rawU, 'val') !== 'none') {
      /* keep styled underline */
    } else {
      if (rawU) drop.add(rawU)
      parts.push(el('w:u', { 'w:val': 'single' }))
    }
  } else if (rawU && attr(rawU, 'val') !== 'none') drop.add(rawU)
  if (ts.backgroundColor && !hl) {
    const r = rawBy('shd')
    if (r) drop.add(r)
    parts.push(el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': String(ts.backgroundColor).replace('#', '').toUpperCase() }))
  }
  if (has('superscript')) parts.push(el('w:vertAlign', { 'w:val': 'superscript' }))
  else if (has('subscript')) parts.push(el('w:vertAlign', { 'w:val': 'subscript' }))
  const managed = new Set(parts.map((p) => /^<w:([A-Za-z]+)/.exec(p)?.[1]))
  for (const r of raw) {
    if (drop.has(r)) continue
    if (r.namespaceURI === NS.w && managed.has(r.localName)) continue
    parts.push(serializeXml(r))
  }
  return parts.length ? `<w:rPr>${orderChildren(parts, RPR_ORDER)}</w:rPr>` : ''
}

function textRuns(t: string, rPr: string, deleted: boolean): string {
  let out = ''
  let buf = ''
  const tTag = deleted ? 'w:delText' : 'w:t'
  const flush = () => {
    if (buf) out += `<w:r>${rPr}<${tTag} xml:space="preserve">${esc(buf)}</${tTag}></w:r>`
    buf = ''
  }
  for (const ch of t) {
    if (ch === '\t') {
      flush()
      out += `<w:r>${rPr}<w:tab/></w:r>`
    } else if (ch === '\n') {
      flush()
      out += `<w:r>${rPr}<w:br/></w:r>`
    } else if (ch === '\u2011') {
      flush()
      out += `<w:r>${rPr}<w:noBreakHyphen/></w:r>`
    } else if (ch === '\u00AD') {
      flush()
      out += `<w:r>${rPr}<w:softHyphen/></w:r>`
    } else buf += ch
  }
  flush()
  return out
}

// ---------------------------------------------------------------------------
// images

function dataUrlBytes(url: string): { bytes: Uint8Array; mime: string } | null {
  const m = /^data:([^;,]+);base64,(.*)$/s.exec(url)
  if (!m) return null
  const bin = atob(m[2])
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return { bytes, mime: m[1].toLowerCase() }
}

export function imageSize(b: Uint8Array): { w: number; h: number } | null {
  if (b[0] === 0x89 && b[1] === 0x50) return { w: (b[16] << 24) | (b[17] << 16) | (b[18] << 8) | b[19], h: (b[20] << 24) | (b[21] << 16) | (b[22] << 8) | b[23] }
  if (b[0] === 0x47 && b[1] === 0x49) return { w: b[6] | (b[7] << 8), h: b[8] | (b[9] << 8) }
  if (b[0] === 0xff && b[1] === 0xd8) {
    let i = 2
    while (i < b.length) {
      if (b[i] !== 0xff) return null
      const marker = b[i + 1]
      const len = (b[i + 2] << 8) | b[i + 3]
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) return { h: (b[i + 5] << 8) | b[i + 6], w: (b[i + 7] << 8) | b[i + 8] }
      i += 2 + len
    }
  }
  return null
}

function hashBytes(b: Uint8Array): string {
  // FNV-1a 32-bit, sufficient for deduplicating media inside one package
  let h = 0x811c9dc5
  for (let i = 0; i < b.length; i++) {
    h ^= b[i]
    h = Math.imul(h, 0x01000193)
  }
  return (h >>> 0).toString(16) + b.length.toString(16)
}

async function imageRun(node: PMNode, ctx: Ctx, rPr: string): Promise<string> {
  const src = String(node.attrs?.src || '')
  let data = dataUrlBytes(src)
  if (!data && (/^https?:/i.test(src) || src.startsWith('/api/media/'))) {
    try {
      const res = await fetch(src, { credentials: 'same-origin' })
      if (!res.ok) throw new Error(String(res.status))
      const buf = new Uint8Array(await res.arrayBuffer())
      data = { bytes: buf, mime: (res.headers.get('content-type') || '').split(';')[0] }
    } catch {
      data = null
    }
  }
  const ext = data && IMG_EXT[data.mime]
  if (!data || !ext) {
    ctx.report.add('dropped', 'Images that could not be embedded', 'Remote or unsupported images could not be embedded and were left out of the DOCX.')
    return ''
  }
  if (ext === 'webp') ctx.report.add('transformed', 'WebP images', 'WebP images are embedded as-is; older desktop editors may not display them.', 'info')
  const hash = hashBytes(data.bytes)
  let target = ctx.media.get(hash)
  if (!target) {
    target = `media/scribe_${hash}.${ext}`
    ctx.zip.file(resolvePath('word/', target), data.bytes)
    ctx.media.set(hash, target)
    ctx.ct.ensureDefault(ext, data.mime)
  }
  // media lives in word/media; all story parts (document, headers, footers) sit in word/
  const rid = ctx.rels.add(REL.image, target)
  let w = Number(node.attrs?.width) || 0
  let h = Number(node.attrs?.height) || 0
  if (!w || !h) {
    const nat = imageSize(data.bytes)
    if (nat) {
      if (!w && !h) {
        w = nat.w
        h = nat.h
      } else if (!h) h = Math.round((w * nat.h) / nat.w)
      else w = Math.round((h * nat.w) / nat.h)
    } else {
      w = w || 300
      h = h || 200
    }
  }
  const maxW = twipToPx(ctx.textWidthTwips)
  if (w > maxW && !node.attrs?.anchorXml) {
    h = Math.round((h * maxW) / w)
    w = Math.round(maxW)
  }
  const cx = pxToEmu(w)
  const cy = pxToEmu(h)
  const id = ctx.ids.docPr++
  if (node.attrs?.anchorXml) {
    let a = String(node.attrs.anchorXml)
    a = a.replace(/(r:embed=")[^"]*(")/, `$1${rid}$2`)
    a = a.replace(/(<wp:extent[^>]*\bcx=")\d+("[^>]*\bcy=")\d+/, `$1${cx}$2${cy}`)
    a = a.replace(/(<a:ext[^>]*\bcx=")\d+("[^>]*\bcy=")\d+/, `$1${cx}$2${cy}`)
    a = a.replace(/(<wp:docPr[^>]*\bid=")\d+/, `$1${id}`)
    return `<w:r>${rPr}<w:drawing>${a}</w:drawing></w:r>`
  }
  const name = esc(String(node.attrs?.docPrName || `Picture ${id}`))
  const descr = node.attrs?.alt ? ` descr="${esc(String(node.attrs.alt))}"` : ''
  const title = node.attrs?.title ? ` title="${esc(String(node.attrs.title))}"` : ''
  return (
    `<w:r>${rPr}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="${name}"${descr}${title}/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}.${ext}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  )
}

// ---------------------------------------------------------------------------
// inline content

async function inlinesXml(nodes: PMNode[] | undefined, ctx: Ctx, prefix = ''): Promise<string> {
  if (!nodes?.length) return prefix
  let out = prefix
  let openComments = new Set<string>()
  // group by (link, revision) wrappers
  let group: { key: string; link?: Mark; rev?: Mark; xml: string } | null = null
  const closeGroup = () => {
    if (!group) return
    let x = group.xml
    if (group.rev) {
      const tag = group.rev.type === 'insertion' ? 'w:ins' : 'w:del'
      x = el(tag, { 'w:id': group.rev.attrs?.revId ?? ctx.ids.bookmark++, 'w:author': group.rev.attrs?.author || 'SCRIBE', 'w:date': group.rev.attrs?.date || undefined }, x)
    }
    if (group.link) {
      const href = String(group.link.attrs?.href || '')
      if (href.startsWith('#')) x = el('w:hyperlink', { 'w:anchor': href.slice(1), 'w:history': 1 }, x)
      else if (/^(https?:|mailto:|tel:|ftp:)/i.test(href)) {
        const hashAt = href.indexOf('#')
        const base = hashAt > 0 && !/^mailto:/i.test(href) ? href.slice(0, hashAt) : href
        const anchor = hashAt > 0 && !/^mailto:/i.test(href) ? href.slice(hashAt + 1) : undefined
        x = el('w:hyperlink', { 'r:id': ctx.rels.add(REL.hyperlink, base, true), 'w:anchor': anchor, 'w:history': 1 }, x)
      }
    }
    out += x
    group = null
  }
  for (const n of nodes) {
    const marks = (n.marks ?? []) as Mark[]
    // comment ranges (not supported inside headers)
    if (!ctx.inHeaderFooter) {
      const ids = new Set(marks.filter((m) => m.type === 'comment' && ctx.comments.has(String(m.attrs?.commentId))).map((m) => String(m.attrs!.commentId)))
      for (const id of openComments) if (!ids.has(id)) {
        closeGroup()
        out += el('w:commentRangeEnd', { 'w:id': id }) + `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${esc(id)}"/></w:r>`
      }
      for (const id of ids) if (!openComments.has(id)) {
        closeGroup()
        out += el('w:commentRangeStart', { 'w:id': id })
        ctx.anchored.add(id)
      }
      openComments = ids
    }
    const link = markOf(marks, 'link')
    const rev = markOf(marks, 'insertion') ?? markOf(marks, 'deletion')
    const key = JSON.stringify([link?.attrs?.href ?? null, rev ? [rev.type, rev.attrs?.author, rev.attrs?.date, rev.attrs?.revId] : null])
    if (n.type === 'opaqueInline' && /^<(w:)?(hyperlink|fldSimple|ins|del|smartTag|sdt|bookmarkStart|bookmarkEnd)\b|^<m:/.test(String(n.attrs?.xml))) closeGroup()
    if (!group || group.key !== key) {
      closeGroup()
      group = { key, link, rev, xml: '' }
    }
    const rPr = rPrXml(marks, ctx)
    const deleted = rev?.type === 'deletion'
    switch (n.type) {
      case 'text':
        group.xml += textRuns(n.text || '', rPr, deleted)
        break
      case 'hardBreak':
        group.xml += `<w:r>${rPr}<w:br/></w:r>`
        break
      case 'columnBreak':
        group.xml += `<w:r>${rPr}<w:br w:type="column"/></w:r>`
        break
      case 'image':
        group.xml += await imageRun(n, ctx, rPr)
        break
      case 'bookmark': {
        const id = ctx.ids.bookmark++
        const name = String(n.attrs?.name || `_scribe${id}`).slice(0, 40)
        group.xml += el('w:bookmarkStart', { 'w:id': id, 'w:name': name }) + el('w:bookmarkEnd', { 'w:id': id })
        break
      }
      case 'field': {
        const instr = String(n.attrs?.instr || 'PAGE')
        group.xml +=
          `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> ${esc(instr.trim())} </w:instrText></w:r>` +
          `<w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r>` +
          (n.attrs?.result ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(String(n.attrs.result))}</w:t></w:r>` : '') +
          `<w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r>`
        break
      }
      case 'noteRef': {
        const tag = n.attrs?.kind === 'endnote' ? 'w:endnoteReference' : 'w:footnoteReference'
        const ref = `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`
        group.xml += `<w:r>${ref}<${tag} w:id="${esc(String(n.attrs?.noteId))}"/></w:r>`
        break
      }
      case 'opaqueInline':
        if (n.attrs?.object?.kind === 'chart' && n.attrs.object.chart?.isNew) {
          group.xml += newChartRun(n.attrs.object, ctx, rPr)
          break
        }
        await writeObjectEdits(n.attrs?.object, ctx)
        group.xml += String(n.attrs?.xml || '')
        break
      case 'shape':
        group.xml += shapeRunXml(n.attrs as ShapeAttrs, ctx.ids.docPr++, rPr)
        break
      default:
        ctx.report.add('dropped', `Inline ${n.type}`, 'This inline element has no DOCX equivalent and was skipped.')
    }
  }
  closeGroup()
  for (const id of openComments) out += el('w:commentRangeEnd', { 'w:id': id }) + `<w:r><w:commentReference w:id="${esc(id)}"/></w:r>`
  return out
}

// ---------------------------------------------------------------------------
// blocks

const JC: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' }

interface ParaOpts {
  list?: { numId: string; ilvl: number }
  sectPr?: string
  prefixRuns?: string
  forceStyle?: string
  extraPPr?: string[]
  runMarks?: Mark[]
}

async function paragraphXml(node: PMNode, ctx: Ctx, o: ParaOpts = {}): Promise<string> {
  const a = node.attrs ?? {}
  const parts: string[] = []
  let styleId: string | null = o.forceStyle ?? a.styleId ?? null
  if (node.type === 'heading') {
    const lvl = Number(a.level) || 1
    if (!styleId || !ctx.styleIsHeading(styleId) || headingLevelOf(ctx, styleId) !== lvl) styleId = ctx.headingStyle(lvl)
  } else if (styleId && ctx.styleIsHeading(styleId) && !o.forceStyle) styleId = null
  if (styleId) {
    parts.push(el('w:pStyle', { 'w:val': styleId }))
    ctx.usedStyles.add(styleId)
  }
  if (a.keepNext) parts.push(el('w:keepNext'))
  if (a.pageBreakBefore) parts.push(el('w:pageBreakBefore'))
  if (o.list) parts.push(el('w:numPr', {}, el('w:ilvl', { 'w:val': o.list.ilvl }) + el('w:numId', { 'w:val': o.list.numId })))
  if (a.shading) parts.push(el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': String(a.shading).replace('#', '').toUpperCase() }))
  if (a.direction === 'rtl') parts.push(el('w:bidi'))
  if (a.spacingBefore != null || a.spacingAfter != null || a.lineSpacing != null || a.spacingExtra) {
    const sp: Record<string, any> = { ...(a.spacingExtra ?? {}) }
    if (a.spacingBefore != null) sp['w:before'] = Math.round(a.spacingBefore)
    if (a.spacingAfter != null) sp['w:after'] = Math.round(a.spacingAfter)
    if (a.lineSpacing != null) {
      const rule = a.lineRule || 'auto'
      sp['w:line'] = rule === 'auto' ? Math.round(Number(a.lineSpacing) * 240) : Math.round(a.lineSpacing)
      sp['w:lineRule'] = rule
    }
    parts.push(el('w:spacing', sp))
  }
  if (a.indentLeft != null || a.indentRight != null || a.indentFirstLine != null || a.indExtra) {
    const ind: Record<string, any> = { ...(a.indExtra ?? {}) }
    if (a.indentLeft != null) ind['w:left'] = Math.round(a.indentLeft)
    if (a.indentRight != null) ind['w:right'] = Math.round(a.indentRight)
    if (a.indentFirstLine != null) {
      if (a.indentFirstLine < 0) ind['w:hanging'] = Math.round(-a.indentFirstLine)
      else ind['w:firstLine'] = Math.round(a.indentFirstLine)
    }
    parts.push(el('w:ind', ind))
  }
  if (a.textAlign && JC[a.textAlign]) parts.push(el('w:jc', { 'w:val': JC[a.textAlign] }))
  if (o.extraPPr) parts.push(...o.extraPPr)
  const managed = new Set(parts.map((p) => /^<w:([A-Za-z]+)/.exec(p)?.[1]))
  for (const f of splitFragments(a.pPrExtra)) {
    if (f.namespaceURI === NS.w && managed.has(f.localName)) continue
    if (f.namespaceURI === NS.w && f.localName === 'numPr') continue
    parts.push(serializeXml(f))
  }
  if (a.markRPr) parts.push(String(a.markRPr))
  if (o.sectPr) parts.push(o.sectPr)
  const pPr = parts.length ? `<w:pPr>${orderChildren(parts, PPR_ORDER)}</w:pPr>` : ''
  let content = node.content
  if (o.runMarks && content) content = content.map((c) => ({ ...c, marks: [...(c.marks ?? []), ...o.runMarks!] }))
  return `<w:p>${pPr}${await inlinesXml(content, ctx, o.prefixRuns ?? '')}</w:p>`
}

function headingLevelOf(ctx: Ctx, id: string): number | null {
  const s = ctx.meta.styles.find((x) => x.id === id)
  if (!s) {
    const m = /^Heading([1-6])$/.exec(id)
    return m ? Number(m[1]) : null
  }
  const m = /^heading\s*([1-6])$/i.exec(s.name)
  if (m) return Number(m[1])
  return s.outlineLevel != null && s.outlineLevel < 6 ? s.outlineLevel + 1 : null
}

const PAGE_BREAK_RUN = '<w:r><w:br w:type="page"/></w:r>'

async function blocksXml(blocks: PMNode[] | undefined, ctx: Ctx, opts: { topLevel?: boolean; sections?: SectionProps[] } = {}): Promise<string> {
  if (!blocks?.length) return ''
  let out = ''
  let pendingBreak = false
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i]
    const next = blocks[i + 1]
    const prefix = pendingBreak ? PAGE_BREAK_RUN : ''
    const isPara = b.type === 'paragraph' || b.type === 'heading'
    if (pendingBreak && !isPara) out += `<w:p>${PAGE_BREAK_RUN}</w:p>`
    pendingBreak = false
    const sectPr = opts.topLevel && next?.type === 'sectionBreak' && isPara ? sectPrXml(ctx, Number(next.attrs?.section ?? 0)) : undefined
    switch (b.type) {
      case 'paragraph':
      case 'heading':
        out += await paragraphXml(b, ctx, { prefixRuns: prefix, sectPr })
        break
      case 'pageBreak':
        if (ctx.inHeaderFooter) break
        pendingBreak = true
        break
      case 'sectionBreak': {
        const prev = blocks[i - 1]
        if (!(prev && (prev.type === 'paragraph' || prev.type === 'heading'))) out += `<w:p><w:pPr>${sectPrXml(ctx, Number(b.attrs?.section ?? 0))}</w:pPr></w:p>`
        break
      }
      case 'bulletList':
      case 'orderedList':
        out += await listXml(b, ctx, 0)
        break
      case 'table':
        out += await tableXml(b, ctx)
        break
      case 'blockquote':
        ctx.usedStyles.add('Quote')
        for (const c of b.content ?? []) {
          if (c.type === 'paragraph' || c.type === 'heading') out += await paragraphXml(c, ctx, { forceStyle: 'Quote' })
          else out += await blocksXml([c], ctx)
        }
        break
      case 'codeBlock': {
        const lines = (b.content ?? []).map((t) => t.text ?? '').join('').split('\n')
        for (const line of lines)
          out += await paragraphXml({ type: 'paragraph', attrs: { spacingAfter: 0, lineSpacing: 1, lineRule: 'auto' }, content: line ? [{ type: 'text', text: line }] : [] }, ctx, {
            runMarks: [{ type: 'textStyle', attrs: { fontFamily: 'Consolas' } }],
          })
        break
      }
      case 'horizontalRule':
        out += `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`
        break
      case 'opaqueBlock':
        out += String(b.attrs?.xml || '')
        break
      case 'bibliography': {
        const entries = (Array.isArray(b.attrs?.entries) ? b.attrs!.entries : []) as { text: string; italic?: boolean }[][]
        let inner = await paragraphXml({ type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: String(b.attrs?.title || 'References') }] }, ctx)
        for (const e of entries)
          inner += await paragraphXml({ type: 'paragraph', attrs: { indentLeft: 720, indentFirstLine: -720 }, content: e.filter((g) => g.text).map((g) => ({ type: 'text', text: String(g.text), ...(g.italic ? { marks: [{ type: 'italic' }] } : {}) })) }, ctx)
        if (!entries.length) inner += '<w:p/>'
        out += `<w:sdt><w:sdtPr><w:id w:val="${ctx.ids.bookmark++}"/><w:docPartObj><w:docPartGallery w:val="Bibliographies"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`
        break
      }
      default:
        ctx.report.add('dropped', `Block ${b.type}`, 'This block has no DOCX equivalent and was skipped.')
    }
  }
  if (pendingBreak) out += `<w:p>${PAGE_BREAK_RUN}</w:p>`
  return out
}

async function listXml(list: PMNode, ctx: Ctx, ilvl: number): Promise<string> {
  const numId = allocList(ctx.lists, list)
  let out = ''
  for (const item of list.content ?? []) {
    let first = true
    for (const c of item.content ?? []) {
      if (c.type === 'bulletList' || c.type === 'orderedList') out += await listXml(c, ctx, Math.min(8, ilvl + 1))
      else if ((c.type === 'paragraph' || c.type === 'heading') && first) {
        out += await paragraphXml(c, ctx, { list: { numId, ilvl } })
        first = false
      } else if (c.type === 'paragraph' || c.type === 'heading') {
        // continuation paragraph inside a list item: indented, unnumbered
        out += await paragraphXml({ ...c, attrs: { ...c.attrs, indentLeft: c.attrs?.indentLeft ?? 720 * (ilvl + 1) } }, ctx)
      } else out += await blocksXml([c], ctx)
    }
  }
  return out
}

async function tableXml(t: PMNode, ctx: Ctx): Promise<string> {
  const rows = t.content ?? []
  // occupancy grid for rowspans
  const occ: (null | { span: number; tcW: number })[][] = rows.map(() => [])
  let cols = 0
  const widths: number[] = []
  rows.forEach((row, r) => {
    let c = 0
    for (const cell of row.content ?? []) {
      while (occ[r][c]) c += occ[r][c]!.span
      const span = Number(cell.attrs?.colspan ?? 1) || 1
      const rs = Number(cell.attrs?.rowspan ?? 1) || 1
      const cw: number[] | null = cell.attrs?.colwidth ?? null
      for (let k = 0; k < span; k++) if (cw?.[k] && !widths[c + k]) widths[c + k] = cw[k]
      for (let rr = r + 1; rr < r + rs && rr < rows.length; rr++) occ[rr][c] = { span, tcW: 0 }
      c += span
    }
    while (occ[r][c]) c += occ[r][c]!.span
    cols = Math.max(cols, c)
  })
  const total = ctx.textWidthTwips
  const known = widths.filter(Boolean).map(pxToTwip)
  const knownSum = known.reduce((s, x) => s + x, 0)
  const missing = cols - known.length
  const fallback = missing > 0 ? Math.max(360, Math.round((total - knownSum) / missing) || Math.round(total / cols)) : 0
  const grid: number[] = []
  for (let i = 0; i < cols; i++) grid.push(widths[i] ? pxToTwip(widths[i]) : fallback)

  const a = t.attrs ?? {}
  const tblParts: string[] = []
  const style = a.tblStyle || (a.tblPrExtra ? null : 'TableGrid')
  if (style) {
    tblParts.push(el('w:tblStyle', { 'w:val': style }))
    ctx.usedStyles.add(style)
  }
  for (const f of splitFragments(a.tblPrExtra)) tblParts.push(serializeXml(f))
  const fresh = !a.tblPrExtra && !a.tblStyle
  if (fresh) tblParts.push(el('w:tblW', { 'w:w': 0, 'w:type': 'auto' }))
  if (fresh) tblParts.push(el('w:tblLook', { 'w:val': '04A0', 'w:firstRow': 1, 'w:lastRow': 0, 'w:firstColumn': 1, 'w:lastColumn': 0, 'w:noHBand': 0, 'w:noVBand': 1 }))
  // untouched imported tables without explicit widths keep their original grid
  const gridXml = !widths.some(Boolean) && a.gridXml && (a.gridXml.match(/<w:gridCol\b/g)?.length ?? 0) === cols
    ? String(a.gridXml)
    : `<w:tblGrid>${grid.map((g) => el('w:gridCol', { 'w:w': g })).join('')}</w:tblGrid>`
  let out = `<w:tbl><w:tblPr>${orderChildren(tblParts, TBLPR_ORDER)}</w:tblPr>${gridXml}`

  const occ2: (null | { span: number; w: number })[][] = rows.map(() => [])
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]
    let rowXml = ''
    const trPr = row.attrs?.trPrExtra ? String(row.attrs.trPrExtra) : ''
    let c = 0
    const emitContinuations = () => {
      while (occ2[r][c]) {
        const o = occ2[r][c]!
        rowXml += `<w:tc><w:tcPr>${el('w:tcW', { 'w:w': o.w, 'w:type': 'dxa' })}${o.span > 1 ? el('w:gridSpan', { 'w:val': o.span }) : ''}<w:vMerge/></w:tcPr><w:p/></w:tc>`
        c += o.span
      }
    }
    for (const cell of row.content ?? []) {
      emitContinuations()
      const span = Number(cell.attrs?.colspan ?? 1) || 1
      const rs = Number(cell.attrs?.rowspan ?? 1) || 1
      const w = grid.slice(c, c + span).reduce((s, x) => s + x, 0)
      const ca = cell.attrs ?? {}
      const tc: string[] = [el('w:tcW', { 'w:w': w, 'w:type': 'dxa' })]
      if (span > 1) tc.push(el('w:gridSpan', { 'w:val': span }))
      if (rs > 1) tc.push(el('w:vMerge', { 'w:val': 'restart' }))
      if (ca.background) tc.push(el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': String(ca.background).replace('#', '').toUpperCase() }))
      if (ca.verticalAlign) tc.push(el('w:vAlign', { 'w:val': ca.verticalAlign === 'middle' ? 'center' : ca.verticalAlign }))
      const managed = new Set(tc.map((p) => /^<w:([A-Za-z]+)/.exec(p)?.[1]))
      for (const f of splitFragments(ca.tcPrExtra)) if (!(f.namespaceURI === NS.w && managed.has(f.localName))) tc.push(serializeXml(f))
      let inner = await blocksXml(cell.content, ctx)
      if (!inner.endsWith('</w:p>') && !inner.endsWith('<w:p/>')) inner += '<w:p/>'
      rowXml += `<w:tc><w:tcPr>${orderChildren(tc, TCPR_ORDER)}</w:tcPr>${inner}</w:tc>`
      for (let rr = r + 1; rr < r + rs && rr < rows.length; rr++) occ2[rr][c] = { span, w }
      c += span
    }
    emitContinuations()
    out += `<w:tr>${trPr}${rowXml}</w:tr>`
  }
  return out + '</w:tbl>'
}

// ---------------------------------------------------------------------------
// sections & header/footer parts

function sectPrXml(ctx: Ctx, index: number): string {
  const sec = ctx.meta.sections[index] ?? { page: defaultPage() }
  const keep: string[] = []
  if (sec.rawXml) {
    try {
      const d = parseXml(sec.rawXml, 'sectPr')
      for (const c of children(d.documentElement)) {
        if (c.namespaceURI === NS.w) {
          if (['pgSz', 'pgMar', 'type', 'titlePg'].includes(c.localName)) continue
          if ((c.localName === 'headerReference' || c.localName === 'footerReference') && attr(c, 'type') !== 'even') continue
          if (c.localName === 'pgNumType' && sec.pageNumberStart != null) {
            c.setAttributeNS(NS.w, 'w:start', String(sec.pageNumberStart))
          }
        }
        keep.push(serializeXml(c))
      }
    } catch {
      ctx.report.add('transformed', 'Section properties', 'Unreadable section properties were regenerated.')
    }
  }
  const p = sec.page
  const refs: string[] = []
  const ref = (tag: 'headerReference' | 'footerReference', type: 'default' | 'first', key?: string) => {
    if (!key || !ctx.meta.headerFooters[key]) return
    const rid = ctx.rels.add(tag === 'headerReference' ? REL.header : REL.footer, headerTargets.get(key)!)
    refs.push(el('w:' + tag, { 'w:type': type, 'r:id': rid }))
  }
  ref('headerReference', 'default', sec.headerDefault)
  ref('headerReference', 'first', sec.headerFirst)
  ref('footerReference', 'default', sec.footerDefault)
  ref('footerReference', 'first', sec.footerFirst)
  const parts = [...refs, ...keep]
  if (sec.breakType && sec.breakType !== 'nextPage') parts.push(el('w:type', { 'w:val': sec.breakType }))
  parts.push(el('w:pgSz', { 'w:w': p.width, 'w:h': p.height, 'w:orient': p.orientation === 'landscape' ? 'landscape' : undefined }))
  parts.push(el('w:pgMar', { 'w:top': p.marginTop, 'w:right': p.marginRight, 'w:bottom': p.marginBottom, 'w:left': p.marginLeft, 'w:header': p.marginHeader, 'w:footer': p.marginFooter, 'w:gutter': 0 }))
  if (sec.pageNumberStart != null && !keep.some((k) => k.includes('pgNumType'))) parts.push(el('w:pgNumType', { 'w:start': sec.pageNumberStart }))
  if (sec.titlePage) parts.push(el('w:titlePg'))
  return `<w:sectPr>${orderChildren(parts, SECTPR_ORDER)}</w:sectPr>`
}

// header part key -> rel target (relative to word/)
let headerTargets = new Map<string, string>()

// ---------------------------------------------------------------------------
// fresh-package templates

function stylesXml(defaults: DocMeta['defaults']): string {
  const font = esc(defaults.fontFamily || 'Calibri')
  const sz = Math.round((defaults.fontSizePt || 11) * 2)
  const line = Math.round((defaults.lineHeight || 1.15) * 240)
  const heading = (n: number, size: number) =>
    `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${n === 1 ? 360 : 240}" w:after="80"/><w:outlineLvl w:val="${n - 1}"/></w:pPr>` +
    `<w:rPr><w:b/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/><w:color w:val="1F2A44"/></w:rPr></w:style>`
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${NS.w}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="${line}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>` +
    heading(1, 36) + heading(2, 30) + heading(3, 26) + heading(4, 24) + heading(5, 22) + heading(6, 22) +
    `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>` +
    QUOTE_STYLE + TABLE_GRID_STYLE + HYPERLINK_STYLE +
    `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
    `</w:styles>`
  )
}

const QUOTE_STYLE = `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>`
const TABLE_GRID_STYLE =
  `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>` +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
  `</w:tblBorders></w:tblPr></w:style>`
const HYPERLINK_STYLE = `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`

const SETTINGS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${NS.w}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>` +
  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`

function coreXml(meta: DocMeta['metadata']): string {
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z')
  const t = (tag: string, v?: string, extra = '') => (v ? `<${tag}${extra}>${esc(v)}</${tag}>` : '')
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    t('dc:title', meta.title) + t('dc:subject', meta.subject) + t('dc:creator', meta.creator || 'SCRIBE') + t('cp:keywords', meta.keywords) + t('dc:description', meta.description) +
    t('cp:lastModifiedBy', meta.lastModifiedBy || 'SCRIBE') +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${esc(meta.created || now)}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`
  )
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>SCRIBE</Application></Properties>`

// ---------------------------------------------------------------------------
// main entry

export async function exportDocx(doc: ScribeDocument, original?: Uint8Array | null): Promise<ExportResult> {
  const report = new Report()
  const zip = original ? await JSZip.loadAsync(original) : new JSZip()
  const ct = new ContentTypes(original ? await zip.file('[Content_Types].xml')!.async('string') : null)
  const rootRels = new Rels('_rels/.rels', original ? (await zip.file('_rels/.rels')?.async('string')) ?? null : null)
  let mainPath = 'word/document.xml'
  const mainRel = rootRels.find(REL.officeDocument)
  if (mainRel) mainPath = mainRel.target.replace(/^\//, '')
  else rootRels.add(REL.officeDocument, 'word/document.xml')
  const dir = mainPath.slice(0, mainPath.lastIndexOf('/') + 1)
  const relsPath = `${dir}_rels/${mainPath.slice(dir.length)}.rels`
  const rels = new Rels(relsPath, (await zip.file(relsPath)?.async('string')) ?? null)
  const meta = doc.meta

  // ---- styles
  let stylesTarget = rels.find(REL.styles)?.target
  let stylesXmlStr = stylesTarget ? (await zip.file(resolvePath(dir, stylesTarget))?.async('string')) ?? null : null
  if (!stylesXmlStr) {
    stylesXmlStr = stylesXml(meta.defaults)
    stylesTarget = 'styles.xml'
    rels.add(REL.styles, stylesTarget)
    ct.override(dir + stylesTarget, CT.styles)
  }
  const styleDoc = parseXml(stylesXmlStr, 'styles.xml')
  const styleIds = new Set(wChildren(styleDoc.documentElement, 'style').map((s) => attr(s, 'styleId')!))
  const headingMap = new Map<number, string>()
  for (const s of wChildren(styleDoc.documentElement, 'style')) {
    const name = attr(wChildren(s, 'name')[0], 'val') || ''
    const m = /^heading\s*([1-6])$/i.exec(name)
    if (m && !headingMap.has(Number(m[1]))) headingMap.set(Number(m[1]), attr(s, 'styleId')!)
  }

  // ---- numbering
  let numTarget = rels.find(REL.numbering)?.target
  const numXmlStr = numTarget ? (await zip.file(resolvePath(dir, numTarget))?.async('string')) ?? null : null
  const numDoc = numXmlStr ? parseXml(numXmlStr, 'numbering.xml') : null
  const existing = parseNumbering(numDoc)
  const maxId = (sel: string, a: string) => (numDoc ? Math.max(0, ...wChildren(numDoc.documentElement, sel).map((e) => Number(attr(e, a)) || 0)) : 0)
  const lists: ListAlloc = {
    bulletNumId: null, bulletAbstract: null, orderedAbstract: new Map(), newAbstracts: [], newNums: [],
    nextNumId: maxId('num', 'numId') + 1, nextAbstractId: maxId('abstractNum', 'abstractNumId') + 1, existing,
  }

  const firstPage = meta.sections[meta.sections.length - 1]?.page ?? defaultPage()
  const ctx: Ctx = {
    zip, ct, rels, partDir: dir, report, meta, lists,
    comments: new Set(meta.comments.map((c) => c.id)), anchored: new Set(),
    ids: { bookmark: 100000, docPr: 20000 }, media: new Map(), inHeaderFooter: false,
    headingStyle: (lvl) => {
      const id = headingMap.get(lvl) ?? `Heading${lvl}`
      ctx.usedStyles.add(id)
      return id
    },
    styleIsHeading: (id) => [...headingMap.values()].includes(id) || /^Heading[1-6]$/.test(id) || headingLevelOf(ctx, id) != null,
    usedStyles: new Set(),
    textWidthTwips: firstPage.width - firstPage.marginLeft - firstPage.marginRight,
    editedParts: new Set(),
  }

  // ---- header / footer parts
  headerTargets = new Map()
  const hfKinds = new Map<string, 'header' | 'footer'>()
  for (const s of meta.sections) {
    for (const k of ['headerDefault', 'headerFirst', 'footerDefault', 'footerFirst'] as const) {
      const key = s[k]
      if (key && meta.headerFooters[key]) hfKinds.set(key, k.startsWith('header') ? 'header' : 'footer')
    }
  }
  let hfCounter = 1
  for (const [key, kind] of hfKinds) {
    let part = key
    if (!/^word\/.+\.xml$/.test(key) || !zip.file(key)) {
      while (zip.file(`${dir}${kind}${hfCounter}.xml`)) hfCounter++
      part = `${dir}${kind}${hfCounter++}.xml`
    }
    const target = part.slice(dir.length)
    headerTargets.set(key, target)
    const partDir = part.slice(0, part.lastIndexOf('/') + 1)
    const pRelsPath = `${partDir}_rels/${part.slice(partDir.length)}.rels`
    const pRels = new Rels(pRelsPath, (await zip.file(pRelsPath)?.async('string')) ?? null)
    const sub: Ctx = { ...ctx, rels: pRels, partDir, inHeaderFooter: true }
    let inner = await blocksXml(meta.headerFooters[key].content, sub)
    if (!inner) inner = '<w:p/>'
    const origXml = await zip.file(part)?.async('string')
    const rootTag = kind === 'header' ? 'w:hdr' : 'w:ftr'
    const open = origXml ? openingTag(origXml, rootTag) : `<${rootTag} ${ROOT_NS_ATTRS}>`
    const decls = declaredNs(open)
    zip.file(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${ensureNs(open)}${stripRedundantNs(inner, { ...decls, ...ROOT_NS })}</${rootTag}>`)
    zip.file(pRelsPath, pRels.xml())
    ct.override(part, kind === 'header' ? CT.header : CT.footer)
  }

  // ---- body
  const lastIdx = Math.max(0, meta.sections.length - 1)
  let body = await blocksXml(doc.content.content, ctx, { topLevel: true, sections: meta.sections })
  body += sectPrXml(ctx, lastIdx)

  // ---- comments
  const commentsRel = rels.find(REL.comments)
  const origCommentsPath = commentsRel ? resolvePath(dir, commentsRel.target) : null
  const origCount = origCommentsPath ? ((await zip.file(origCommentsPath)?.async('string')) ?? '').match(/<w:comment\b/g)?.length ?? 0 : 0
  const allPristine = meta.comments.every((c) => c.pristine) && meta.comments.length === origCount
  if (!allPristine) {
    const used = meta.comments.filter((c) => ctx.anchored.has(c.id))
    const path = origCommentsPath ?? `${dir}comments.xml`
    if (used.length) {
      const xml =
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments ${ROOT_NS_ATTRS}>` +
        used
          .map((c) =>
            el('w:comment', { 'w:id': c.id, 'w:author': c.author || 'SCRIBE', 'w:date': c.date || undefined, 'w:initials': c.initials || initials(c.author) },
              (c.text || '').split('\n').map((line, i) => `<w:p>${i === 0 ? '<w:r><w:annotationRef/></w:r>' : ''}<w:r><w:t xml:space="preserve">${esc(line)}</w:t></w:r></w:p>`).join(''),
            ),
          )
          .join('') +
        `</w:comments>`
      zip.file(path, xml)
      if (!commentsRel) rels.add(REL.comments, path.slice(dir.length))
      ct.override(path, CT.comments)
    } else if (origCommentsPath) {
      zip.remove(origCommentsPath)
      ct.remove(origCommentsPath)
      rels.removeType(REL.comments)
    }
    // extended comment parts reference paragraph ids that no longer exist
    for (const t of ['commentsExtended', 'commentsIds', 'commentsExtensible']) {
      for (const r of rels.removeType(`http://schemas.microsoft.com/office/2011/relationships/${t}`).concat(rels.removeType(`http://schemas.microsoft.com/office/2016/09/relationships/${t}`), rels.removeType(`http://schemas.microsoft.com/office/2018/08/relationships/${t}`))) {
        const p = resolvePath(dir, r.target)
        zip.remove(p)
        ct.remove(p)
        report.add('transformed', 'Comment threads', 'Comments were edited, so reply and resolved-state metadata was rebuilt without threading.', 'info')
      }
    }
  }

  // ---- numbering write-back
  if (lists.newAbstracts.length || lists.newNums.length) {
    let numXml: string
    if (numDoc) {
      numXml = serializeXml(numDoc)
      const firstNum = numXml.search(/<w:num\b/)
      const abs = lists.newAbstracts.join('')
      numXml = firstNum >= 0 ? numXml.slice(0, firstNum) + abs + numXml.slice(firstNum) : numXml.replace(/<\/w:numbering>\s*$/, abs + '</w:numbering>')
      const cleanup = numXml.search(/<w:numIdMacAtCleanup\b/)
      const nums = lists.newNums.join('')
      numXml = cleanup >= 0 ? numXml.slice(0, cleanup) + nums + numXml.slice(cleanup) : numXml.replace(/<\/w:numbering>\s*$/, nums + '</w:numbering>')
    } else {
      numXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="${NS.w}">${lists.newAbstracts.join('')}${lists.newNums.join('')}</w:numbering>`
      numTarget = 'numbering.xml'
      rels.add(REL.numbering, numTarget)
      ct.override(dir + numTarget, CT.numbering)
    }
    zip.file(resolvePath(dir, numTarget!), numXml.startsWith('<?xml') ? numXml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${numXml}`)
  }

  // ---- edited / new styles
  const styleEdits = applyStyleEdits(styleDoc, meta.styles, (id) => stylesXml(meta.defaults).match(new RegExp(`<w:style w:type="paragraph"(?: w:default="1")? w:styleId="${id}">.*?</w:style>`))?.[0] ?? null)
  for (const s of meta.styles) if (s.modified || s.custom) styleIds.add(s.id)

  // ---- bibliography sources part
  if (meta.sources?.length) await writeSources(zip, ct, rels, dir, meta)

  // ---- edited custom XML parts
  for (const part of meta.customXml ?? []) {
    if (!part.edited || !/^customXml\/item\d+\.xml$/i.test(part.path) || !zip.file(part.path)) continue
    try {
      parseXml(part.xml, part.path)
    } catch {
      report.add('dropped', 'Custom XML edits', `${part.path} was not well-formed XML, so the original part was kept.`)
      continue
    }
    zip.file(part.path, part.xml.startsWith('<?xml') ? part.xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${part.xml}`)
  }

  // ---- styles referenced but missing
  const missing: string[] = []
  for (const id of ctx.usedStyles) {
    if (id.startsWith('__') || styleIds.has(id)) continue
    const m = /^Heading([1-6])$/.exec(id)
    if (m) missing.push(stylesXml(meta.defaults).match(new RegExp(`<w:style w:type="paragraph" w:styleId="${id}">.*?</w:style>`))![0])
    else if (id === 'Quote') missing.push(QUOTE_STYLE)
    else if (id === 'TableGrid') missing.push(TABLE_GRID_STYLE)
    else if (id === 'Hyperlink') missing.push(HYPERLINK_STYLE)
  }
  if (missing.length || !original || styleEdits) {
    const s = serializeXml(styleDoc).replace(/<\/w:styles>\s*$/, missing.join('') + '</w:styles>')
    zip.file(resolvePath(dir, stylesTarget!), s.startsWith('<?xml') ? s : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`)
  }

  // ---- document part
  const origDoc = original ? await zip.file(mainPath)?.async('string') : null
  let docOpen = `<w:document ${ROOT_NS_ATTRS}>`
  let background = ''
  if (origDoc) {
    docOpen = ensureNs(openingTag(origDoc, 'w:document'))
    const bg = /<w:background\b[\s\S]*?(?:\/>|<\/w:background>)/.exec(origDoc)
    if (bg) background = bg[0]
  }
  const decls = { ...declaredNs(docOpen), ...ROOT_NS }
  zip.file(mainPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${docOpen}${background}<w:body>${stripRedundantNs(body, decls)}</w:body></w:document>`)
  ct.override(mainPath, CT.document)

  // ---- settings / properties
  if (!rels.find(REL.settings)) {
    zip.file(dir + 'settings.xml', SETTINGS_XML)
    rels.add(REL.settings, 'settings.xml')
    ct.override(dir + 'settings.xml', CT.settings)
  }
  const coreRel = rootRels.find(REL.core)
  const corePath = coreRel ? coreRel.target.replace(/^\//, '') : 'docProps/core.xml'
  zip.file(corePath, coreXml(meta.metadata))
  if (!coreRel) rootRels.add(REL.core, corePath)
  ct.override(corePath, CT.core)
  if (!rootRels.find(REL.app)) {
    zip.file('docProps/app.xml', APP_XML)
    rootRels.add(REL.app, 'docProps/app.xml')
    ct.override('docProps/app.xml', CT.app)
  }
  for (const [e, t] of Object.entries({ png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif' })) if ([...ctx.media.values()].some((m) => m.endsWith('.' + e))) ct.ensureDefault(e, t)

  zip.file(relsPath, rels.xml())
  zip.file('_rels/.rels', rootRels.xml())
  zip.file('[Content_Types].xml', ct.xml())
  const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
  return { data, report: report.list() }
}

const REL_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
const REL_CUSTOM_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
const CT_CUSTOM_PROPS = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'

/** Write meta.sources into the package's bibliography part (creating it when needed). */
async function writeSources(zip: JSZip, ct: ContentTypes, rels: Rels, dir: string, meta: DocMeta) {
  const xml = sourcesXml(meta.sources!, meta.citationStyle ?? 'APA')
  for (const name of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/i.test(name)) continue
    const cur = await zip.file(name)!.async('string')
    if (cur.includes(BIB_NS) && /<(?:\w+:)?Sources\b/.test(cur)) {
      zip.file(name, xml)
      return
    }
  }
  let n = 1
  while (zip.file(`customXml/item${n}.xml`) || zip.file(`customXml/itemProps${n}.xml`)) n++
  zip.file(`customXml/item${n}.xml`, xml)
  const id = `{${crypto.randomUUID().toUpperCase()}}`
  zip.file(
    `customXml/itemProps${n}.xml`,
    `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="${id}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${BIB_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
  )
  const itemRels = new Rels(`customXml/_rels/item${n}.xml.rels`, null)
  itemRels.add(REL_CUSTOM_PROPS, `itemProps${n}.xml`)
  zip.file(`customXml/_rels/item${n}.xml.rels`, itemRels.xml())
  ct.override(`customXml/itemProps${n}.xml`, CT_CUSTOM_PROPS)
  ct.ensureDefault('xml', 'application/xml')
  rels.add(REL_CUSTOM_XML, `${'../'.repeat(dir.split('/').filter(Boolean).length)}customXml/item${n}.xml`)
}

function initials(name: string) {
  return name.split(/\s+/).map((p) => p[0] ?? '').join('').slice(0, 3).toUpperCase() || 'S'
}

function openingTag(xml: string, tag: string): string {
  const m = new RegExp(`<${tag}\\b[^>]*>`).exec(xml)
  return m ? m[0].replace(/\/>$/, '>') : `<${tag} ${ROOT_NS_ATTRS}>`
}

function declaredNs(open: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const m of open.matchAll(/xmlns:([A-Za-z0-9]+)="([^"]*)"/g)) out[m[1]] = m[2]
  return out
}

/** Make sure the prefixes SCRIBE emits are declared on the part root. */
function ensureNs(open: string): string {
  const have = declaredNs(open)
  const add = Object.entries(ROOT_NS)
    .filter(([p]) => !have[p])
    .map(([p, u]) => ` xmlns:${p}="${u}"`)
    .join('')
  return add ? open.replace(/>$/, add + '>') : open
}
