// SCRIBE document model.
// The editable body is ProseMirror JSON (Tiptap schema defined in ./extensions.ts).
// Everything that is not body text lives in DocMeta so the editor, renderer and
// exporters all read one source of truth.

export type PMNode = {
  type: string
  attrs?: Record<string, any>
  content?: PMNode[]
  marks?: { type: string; attrs?: Record<string, any> }[]
  text?: string
}

/** Page geometry in twips (1/1440 inch), as in OOXML. */
export interface PageSetup {
  width: number
  height: number
  orientation: 'portrait' | 'landscape'
  marginTop: number
  marginRight: number
  marginBottom: number
  marginLeft: number
  marginHeader: number
  marginFooter: number
}

export interface SectionProps {
  page: PageSetup
  /** Header/footer content keys into DocMeta.headerFooters */
  headerDefault?: string
  footerDefault?: string
  headerFirst?: string
  footerFirst?: string
  titlePage?: boolean
  /** Section start type: nextPage | continuous | evenPage | oddPage | nextColumn */
  breakType?: string
  columns?: number
  pageNumberStart?: number
  /** Original <w:sectPr> XML; unknown children are re-emitted on export. */
  rawXml?: string
}

export interface CommentThread {
  id: string
  author: string
  initials?: string
  date: string
  text: string
  resolved?: boolean
  /** true when text/author unchanged since import (lets exporter keep original part) */
  pristine?: boolean
}

export interface Footnote {
  id: string
  /** Plain text preview for display. */
  text: string
  /** Original <w:footnote> XML (preserved verbatim if present). */
  rawXml?: string
}

export interface DocMetadata {
  title?: string
  subject?: string
  creator?: string
  keywords?: string
  description?: string
  lastModifiedBy?: string
  created?: string
  modified?: string
}

export interface CompatibilityIssue {
  severity: 'info' | 'warning' | 'error'
  /** preserved: kept verbatim; transformed: converted with possible fidelity change; dropped: lost */
  kind: 'preserved' | 'transformed' | 'dropped' | 'unsupported'
  feature: string
  detail: string
  count: number
}

export interface DocMeta {
  sections: SectionProps[]
  headerFooters: Record<string, PMNode>
  comments: CommentThread[]
  footnotes: Footnote[]
  endnotes: Footnote[]
  metadata: DocMetadata
  /** Default text style (from docDefaults) */
  defaults: { fontFamily: string; fontSizePt: number; lineHeight: number }
  /** Paragraph styles available in the style picker */
  styles: StyleDef[]
}

export interface StyleDef {
  id: string
  name: string
  type: 'paragraph' | 'character' | 'table' | 'numbering'
  basedOn?: string
  /** CSS-ish resolved props for rendering */
  fontFamily?: string
  fontSizePt?: number
  bold?: boolean
  italic?: boolean
  color?: string
  spacingBefore?: number
  spacingAfter?: number
  outlineLevel?: number
}

export interface ScribeDocument {
  version: number
  content: PMNode
  meta: DocMeta
}

export const SCHEMA_VERSION = 1

export const PAGE_SIZES: Record<string, { width: number; height: number; label: string }> = {
  A4: { width: 11906, height: 16838, label: 'A4 (210 × 297 mm)' },
  Letter: { width: 12240, height: 15840, label: 'Letter (8.5 × 11 in)' },
  Legal: { width: 12240, height: 20160, label: 'Legal (8.5 × 14 in)' },
  A5: { width: 8391, height: 11906, label: 'A5 (148 × 210 mm)' },
}

export function defaultPage(size: keyof typeof PAGE_SIZES = 'Letter'): PageSetup {
  const s = PAGE_SIZES[size]
  return {
    width: s.width,
    height: s.height,
    orientation: 'portrait',
    marginTop: 1440,
    marginRight: 1440,
    marginBottom: 1440,
    marginLeft: 1440,
    marginHeader: 720,
    marginFooter: 720,
  }
}

export function emptyMeta(): DocMeta {
  return {
    sections: [{ page: defaultPage() }],
    headerFooters: {},
    comments: [],
    footnotes: [],
    endnotes: [],
    metadata: {},
    defaults: { fontFamily: 'Calibri', fontSizePt: 11, lineHeight: 1.15 },
    styles: [],
  }
}

export function emptyDocument(): ScribeDocument {
  return {
    version: SCHEMA_VERSION,
    content: { type: 'doc', content: [{ type: 'paragraph' }] },
    meta: emptyMeta(),
  }
}

/** Upgrade older serialized documents to the current schema. */
export function migrateDocument(raw: any): ScribeDocument {
  if (!raw || typeof raw !== 'object') throw new Error('Invalid document payload')
  const doc: ScribeDocument = {
    version: SCHEMA_VERSION,
    content: raw.content ?? { type: 'doc', content: [{ type: 'paragraph' }] },
    meta: { ...emptyMeta(), ...(raw.meta ?? {}) },
  }
  if (!doc.meta.sections?.length) doc.meta.sections = [{ page: defaultPage() }]
  if (doc.content.type !== 'doc') throw new Error('Document root must be "doc"')
  return doc
}

export const twipToPx = (tw: number) => (tw / 1440) * 96
export const pxToTwip = (px: number) => Math.round((px / 96) * 1440)
export const emuToPx = (emu: number) => emu / 9525
export const pxToEmu = (px: number) => Math.round(px * 9525)
