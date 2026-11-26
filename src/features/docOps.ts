// Editor operations shared by the References, Mailings and styles UI.

import type { Editor } from '@tiptap/react'
import type { PageInfo } from '../layout/pagination'
import type { TocEntry } from '../model/extensions'
import type { DocMeta, Source, StyleDef } from '../model/types'
import { BIBLIOGRAPHY_TITLES, formatBibliography, formatCitation, parseCitationInstr } from './citations'
import { headingLevelOfStyle } from './styles'

/** Recompute every citation result and bibliography from the current sources and style. */
export function refreshCitations(editor: Editor, meta: Pick<DocMeta, 'sources' | 'citationStyle'>, opts: { retitle?: boolean } = {}) {
  const sources = meta.sources ?? []
  const style = meta.citationStyle ?? 'APA'
  const byTag = new Map(sources.map((s) => [s.tag, s]))
  const { state } = editor
  const tr = state.tr
  let changed = false
  state.doc.descendants((node, pos) => {
    if (node.type.name === 'field') {
      const c = parseCitationInstr(String(node.attrs.instr))
      if (!c) return
      const found = c.tags.map((t) => byTag.get(t)).filter((s): s is Source => !!s)
      const result = found.length ? formatCitation(found, style, c.page) : `(Missing source: ${c.tags.join(', ')})`
      if (result !== node.attrs.result) {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, result })
        changed = true
      }
    } else if (node.type.name === 'bibliography') {
      const entries = formatBibliography(sources, style)
      const title = opts.retitle ? BIBLIOGRAPHY_TITLES[style] : node.attrs.title
      if (JSON.stringify(entries) !== JSON.stringify(node.attrs.entries) || title !== node.attrs.title) {
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, entries, title })
        changed = true
      }
    }
  })
  if (changed) editor.view.dispatch(tr)
  return changed
}

export function citedTags(editor: Editor): Set<string> {
  const out = new Set<string>()
  editor.state.doc.descendants((n) => {
    if (n.type.name === 'field') parseCitationInstr(String(n.attrs.instr))?.tags.forEach((t) => out.add(t))
  })
  return out
}

/**
 * Apply a paragraph style to the selection. Built-in styles the document does not yet define are
 * added to its style list so they are written into the exported file.
 */
export function applyParagraphStyle(editor: Editor, style: StyleDef, meta: DocMeta, setMeta: (f: (m: DocMeta) => DocMeta) => void) {
  const lvl = headingLevelOfStyle(style)
  const known = meta.styles.some((s) => s.id === style.id)
  // headings, Normal and Quote are always available on export; other built-ins are written when first used
  if (!known && !lvl && style.id !== 'Normal' && style.id !== 'Quote')
    setMeta((m) => (m.styles.some((s) => s.id === style.id) ? m : { ...m, styles: [...m.styles, { ...style, custom: true }] }))
  const c = editor.chain().focus()
  if (style.id === 'Normal') return c.setParagraph().setParagraphFormat({ styleId: null }).run()
  if (lvl) return c.setHeading({ level: lvl as 1 | 2 | 3 | 4 | 5 | 6 }).setParagraphFormat({ styleId: style.id }).run()
  return c.setParagraph().setParagraphFormat({ styleId: style.id }).run()
}

/** Style properties taken from the selection ("Update style to match selection"). */
export function styleFromSelection(editor: Editor): Partial<StyleDef> {
  const ts = editor.getAttributes('textStyle')
  const para = editor.isActive('heading') ? editor.getAttributes('heading') : editor.getAttributes('paragraph')
  const out: Partial<StyleDef> = {}
  if (ts.fontFamily) out.fontFamily = String(ts.fontFamily).replace(/["']/g, '').split(',')[0].trim()
  if (ts.fontSize) out.fontSizePt = parseFloat(ts.fontSize)
  if (ts.color && /^#[0-9a-f]{6}$/i.test(ts.color)) out.color = ts.color
  out.bold = editor.isActive('bold')
  out.italic = editor.isActive('italic')
  out.underline = editor.isActive('underline')
  if (para.textAlign) out.alignment = para.textAlign
  if (para.spacingBefore != null) out.spacingBefore = para.spacingBefore
  if (para.spacingAfter != null) out.spacingAfter = para.spacingAfter
  if (para.lineSpacing != null && (!para.lineRule || para.lineRule === 'auto')) out.lineSpacing = para.lineSpacing
  return out
}

// ---------------------------------------------------------------------------
// Table of contents

/** Headings up to `levels`, with the number of the page each one is on (from the live layout). */
export function tocEntries(editor: Editor, pages: PageInfo[], zoom: number, levels: number): TocEntry[] {
  const out: TocEntry[] = []
  const root = editor.view.dom.getBoundingClientRect()
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'heading') return node.type.name !== 'paragraph'
    const level = Number(node.attrs.level) || 1
    if (level > levels || !node.textContent.trim()) return false
    let page = ''
    const dom = editor.view.nodeDOM(pos) as HTMLElement | null
    if (dom && pages.length) {
      const y = (dom.getBoundingClientRect().top - root.top) / (zoom || 1)
      const p = pages.find((pg) => y < pg.top + pg.height) ?? pages[pages.length - 1]
      page = String(p.number)
    }
    out.push({ level, text: node.textContent.trim(), page, target: node.attrs.nid ?? undefined })
    return false
  })
  return out
}

/** Rebuild every table of contents. `onlyOwn` skips imported tables that were never updated. */
export function updateTocs(editor: Editor, pages: PageInfo[], zoom: number, onlyOwn = false): number {
  const { state } = editor
  const tr = state.tr
  let n = 0
  state.doc.descendants((node, pos) => {
    if (node.type.name !== 'toc') return node.type.name !== 'paragraph'
    if (onlyOwn && node.attrs.rawXml && !node.attrs.updated) return false
    const entries = tocEntries(editor, pages, zoom, Number(node.attrs.levels) || 3)
    if (JSON.stringify(entries) !== JSON.stringify(node.attrs.entries) || !node.attrs.updated) {
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, entries, updated: true })
      n++
    }
    return false
  })
  if (n) editor.view.dispatch(tr.setMeta('addToHistory', !onlyOwn))
  return n
}

/** Scroll to the heading a table-of-contents entry points to. */
export function goToHeading(editor: Editor, nid: string): boolean {
  let at = -1
  editor.state.doc.descendants((node, pos) => {
    if (at !== -1) return false
    if (node.type.name === 'heading' && node.attrs.nid === nid) at = pos
    return true
  })
  if (at === -1) return false
  editor.chain().focus().setTextSelection(at + 1).run()
  ;(editor.view.nodeDOM(at) as HTMLElement | null)?.scrollIntoView({ block: 'start', behavior: 'smooth' })
  return true
}
