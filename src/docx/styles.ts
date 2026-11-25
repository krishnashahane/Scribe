import type { StyleDef } from '../model/types'
import { NS, attr, child, children, intAttr, onOff, wChild, wChildren, wVal } from './xml'

export interface ThemeFonts {
  major: string
  minor: string
}

export function parseTheme(doc: Document | null): ThemeFonts {
  const out = { major: 'Calibri Light', minor: 'Calibri' }
  if (!doc) return out
  const scheme = doc.getElementsByTagNameNS(NS.a, 'fontScheme')[0]
  const major = scheme && child(child(scheme, NS.a, 'majorFont'), NS.a, 'latin')
  const minor = scheme && child(child(scheme, NS.a, 'minorFont'), NS.a, 'latin')
  if (major?.getAttribute('typeface')) out.major = major.getAttribute('typeface')!
  if (minor?.getAttribute('typeface')) out.minor = minor.getAttribute('typeface')!
  return out
}

export function fontFromRFonts(rFonts: Element | null, theme: ThemeFonts): string | null {
  if (!rFonts) return null
  const direct = attr(rFonts, 'ascii') || attr(rFonts, 'hAnsi')
  if (direct) return direct
  const th = attr(rFonts, 'asciiTheme') || attr(rFonts, 'hAnsiTheme')
  if (th) return th.startsWith('major') ? theme.major : theme.minor
  return attr(rFonts, 'eastAsia') || attr(rFonts, 'cs')
}

export interface StyleSheet {
  styles: Map<string, StyleDef>
  defaults: { fontFamily: string; fontSizePt: number; lineHeight: number; spacingAfter: number | null }
  /** styleId -> heading level (1..6) */
  headingLevels: Map<string, number>
  defaultParagraphStyle: string | null
}

export function parseStyles(doc: Document | null, theme: ThemeFonts): StyleSheet {
  const sheet: StyleSheet = {
    styles: new Map(),
    defaults: { fontFamily: theme.minor, fontSizePt: 10, lineHeight: 1, spacingAfter: null },
    headingLevels: new Map(),
    defaultParagraphStyle: null,
  }
  if (!doc) return sheet
  const root = doc.documentElement
  const dd = wChild(root, 'docDefaults')
  const rPr = wChild(wChild(dd, 'rPrDefault'), 'rPr')
  const pPr = wChild(wChild(dd, 'pPrDefault'), 'pPr')
  const f = fontFromRFonts(wChild(rPr, 'rFonts'), theme)
  if (f) sheet.defaults.fontFamily = f
  const sz = intAttr(wChild(rPr, 'sz'), 'val')
  if (sz) sheet.defaults.fontSizePt = sz / 2
  const sp = wChild(pPr, 'spacing')
  const line = intAttr(sp, 'line')
  if (line && (attr(sp, 'lineRule') ?? 'auto') === 'auto') sheet.defaults.lineHeight = line / 240
  const after = intAttr(sp, 'after')
  if (after != null) sheet.defaults.spacingAfter = after

  const raw = new Map<string, Element>()
  for (const s of wChildren(root, 'style')) {
    const id = attr(s, 'styleId')
    if (!id) continue
    raw.set(id, s)
    const type = (attr(s, 'type') || 'paragraph') as StyleDef['type']
    const name = wVal(wChild(s, 'name')) || id
    const r = wChild(s, 'rPr')
    const p = wChild(s, 'pPr')
    const def: StyleDef = { id, name, type, basedOn: wVal(wChild(s, 'basedOn')) || undefined }
    const font = fontFromRFonts(wChild(r, 'rFonts'), theme)
    if (font) def.fontFamily = font
    const ssz = intAttr(wChild(r, 'sz'), 'val')
    if (ssz) def.fontSizePt = ssz / 2
    const b = onOff(wChild(r, 'b'))
    if (b != null) def.bold = b
    const i = onOff(wChild(r, 'i'))
    if (i != null) def.italic = i
    const c = wVal(wChild(r, 'color'))
    if (c && c !== 'auto' && /^[0-9a-fA-F]{6}$/.test(c)) def.color = '#' + c
    const ssp = wChild(p, 'spacing')
    const bf = intAttr(ssp, 'before')
    const af = intAttr(ssp, 'after')
    if (bf != null) def.spacingBefore = bf
    if (af != null) def.spacingAfter = af
    const ol = intAttr(wChild(p, 'outlineLvl'), 'val')
    if (ol != null) def.outlineLevel = ol
    sheet.styles.set(id, def)
    const isDefault = attr(s, 'default')
    if (type === 'paragraph' && isDefault && ['1', 'true', 'on'].includes(isDefault.toLowerCase())) sheet.defaultParagraphStyle = id
  }
  // heading detection (follows basedOn for outline level)
  for (const [id, def] of sheet.styles) {
    if (def.type !== 'paragraph') continue
    const m = /^heading\s*([1-6])$/i.exec(def.name)
    let level: number | null = m ? Number(m[1]) : null
    if (level == null) {
      const ol = resolve(sheet, id, 'outlineLevel')
      if (typeof ol === 'number' && ol >= 0 && ol < 6) level = ol + 1
    }
    if (level != null) sheet.headingLevels.set(id, level)
  }
  return sheet
}

export function resolve<K extends keyof StyleDef>(sheet: StyleSheet, id: string, key: K): StyleDef[K] | undefined {
  const seen = new Set<string>()
  let cur: StyleDef | undefined = sheet.styles.get(id)
  while (cur && !seen.has(cur.id)) {
    if (cur[key] !== undefined) return cur[key]
    seen.add(cur.id)
    cur = cur.basedOn ? sheet.styles.get(cur.basedOn) : undefined
  }
  return undefined
}

/** Flattened (inheritance-resolved) style list for rendering and the style picker. */
export function resolvedStyles(sheet: StyleSheet): StyleDef[] {
  const out: StyleDef[] = []
  for (const [id, def] of sheet.styles) {
    if (def.type !== 'paragraph' && def.type !== 'character') continue
    out.push({
      ...def,
      fontFamily: resolve(sheet, id, 'fontFamily'),
      fontSizePt: resolve(sheet, id, 'fontSizePt'),
      bold: resolve(sheet, id, 'bold'),
      italic: resolve(sheet, id, 'italic'),
      color: resolve(sheet, id, 'color'),
      spacingBefore: resolve(sheet, id, 'spacingBefore'),
      spacingAfter: resolve(sheet, id, 'spacingAfter'),
      outlineLevel: resolve(sheet, id, 'outlineLevel'),
    })
  }
  return out
}

// ---------- numbering ----------

export interface NumLevel {
  numFmt: string
  lvlText: string
  start: number
}

export interface Numbering {
  /** numId -> ilvl -> level */
  levels: Map<string, Map<number, NumLevel>>
}

export function parseNumbering(doc: Document | null): Numbering {
  const out: Numbering = { levels: new Map() }
  if (!doc) return out
  const root = doc.documentElement
  const abstract = new Map<string, Map<number, NumLevel>>()
  for (const a of wChildren(root, 'abstractNum')) {
    const id = attr(a, 'abstractNumId')!
    const lv = new Map<number, NumLevel>()
    for (const l of wChildren(a, 'lvl')) {
      lv.set(intAttr(l, 'ilvl') ?? 0, {
        numFmt: wVal(wChild(l, 'numFmt')) || 'decimal',
        lvlText: wVal(wChild(l, 'lvlText')) || '',
        start: intAttr(wChild(l, 'start'), 'val') ?? 1,
      })
    }
    abstract.set(id, lv)
  }
  for (const n of wChildren(root, 'num')) {
    const numId = attr(n, 'numId')!
    const absId = wVal(wChild(n, 'abstractNumId'))
    const base = new Map(abstract.get(absId || '') ?? [])
    for (const o of wChildren(n, 'lvlOverride')) {
      const il = intAttr(o, 'ilvl') ?? 0
      const l = wChild(o, 'lvl')
      const so = intAttr(wChild(o, 'startOverride'), 'val')
      if (so != null && !l) {
        const cur = base.get(il)
        base.set(il, { numFmt: cur?.numFmt ?? 'decimal', lvlText: cur?.lvlText ?? '', start: so })
      }
      if (l) base.set(il, { numFmt: wVal(wChild(l, 'numFmt')) || 'decimal', lvlText: wVal(wChild(l, 'lvlText')) || '', start: intAttr(wChild(l, 'start'), 'val') ?? 1 })
    }
    out.levels.set(numId, base)
  }
  return out
}

export const NUMFMT_CSS: Record<string, string> = {
  decimal: 'decimal',
  decimalZero: 'decimal-leading-zero',
  lowerLetter: 'lower-alpha',
  upperLetter: 'upper-alpha',
  lowerRoman: 'lower-roman',
  upperRoman: 'upper-roman',
  none: 'none',
}

export { children }
