// Editor operations shared by the References, Mailings and styles UI.

import type { Editor } from '@tiptap/react'
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
