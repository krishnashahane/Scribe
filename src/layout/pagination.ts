import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import type { Node as PMNode } from '@tiptap/pm/model'
import type { SectionProps } from '../model/types'
import { twipToPx } from '../model/types'

/**
 * Pagination for a continuous ProseMirror document.
 *
 * The document is laid out once without page spacers, every leaf block is
 * measured, and spacer widget decorations are inserted where a page ends.
 * Paragraphs that cross a page boundary are split at the first line that does
 * not fit (inline block-level widget), so pages fill to the bottom like a
 * word processor. Explicit page breaks, "page break before", next-page section
 * breaks and per-section page sizes are honoured.
 *
 * Tables break between rows. A single row or atomic block taller than a page
 * overflows into the page gap.
 */

export interface PageInfo {
  index: number
  section: number
  /** page top relative to the editor content box (px) */
  top: number
  width: number
  height: number
  marginTop: number
  marginBottom: number
  marginLeft: number
  marginRight: number
  headerOffset: number
  footerOffset: number
  /** first page of its section (for "different first page") */
  firstOfSection: boolean
  /** displayed page number (honours section page-number restarts) */
  number: number
  /** footnote ids shown at the bottom of this page, in order */
  notes: string[]
  /** px reserved at the bottom of the text area for footnotes (0 when none) */
  noteSpace: number
  /** endnotes block position (layout px from page top) when this page holds them */
  endnotesTop?: number
}

/** px between the body text and the footnote separator */
export const NOTE_SEPARATOR = 14

export interface PaginationOptions {
  getSections: () => SectionProps[]
  /** gap between pages in px (0 while printing) */
  getGap: () => number
  /** CSS zoom applied to the page canvas; client rects are divided by it */
  getScale: () => number
  onLayout: (pages: PageInfo[]) => void
  /** rendered height in px of a footnote's text (0 if unknown) */
  getNoteHeight?: (id: string) => number
  /** rendered height in px of the endnotes block (0 when there are none) */
  getEndnotesHeight?: () => number
}

const key = new PluginKey<DecorationSet>('scribe-pagination')

interface Unit {
  pos: number
  node: PMNode
  /** inline position range for paragraph splitting */
  textblock: boolean
  section: number
  forceBreakBefore: boolean
  /** table row: number of grid columns (the page-break spacer is a table row) */
  rowCols?: number
  /** table row: position of its table */
  tablePos?: number
}

function pageGeom(sec: SectionProps | undefined) {
  const p = sec?.page
  return {
    width: twipToPx(p?.width ?? 12240),
    height: twipToPx(p?.height ?? 15840),
    marginTop: twipToPx(p?.marginTop ?? 1440),
    marginBottom: twipToPx(p?.marginBottom ?? 1440),
    marginLeft: twipToPx(p?.marginLeft ?? 1440),
    marginRight: twipToPx(p?.marginRight ?? 1440),
    headerOffset: twipToPx(p?.marginHeader ?? 720),
    footerOffset: twipToPx(p?.marginFooter ?? 720),
  }
}

function collectUnits(doc: PMNode): Unit[] {
  const units: Unit[] = []
  let section = 0
  let breakNext = false
  doc.forEach((node, offset) => {
    const visit = (n: PMNode, pos: number) => {
      if (n.type.name === 'pageBreak') {
        breakNext = true
        units.push({ pos, node: n, textblock: false, section, forceBreakBefore: false })
        return
      }
      if (n.type.name === 'sectionBreak') {
        units.push({ pos, node: n, textblock: false, section, forceBreakBefore: false })
        section = Number(n.attrs.section ?? section) + 1
        if (n.attrs.breakType !== 'continuous') breakNext = true
        return
      }
      if (n.type.name === 'table') {
        // tables break between rows
        const first = n.firstChild
        let cols = 0
        first?.forEach((cell) => (cols += Number(cell.attrs.colspan) || 1))
        n.forEach((row, o) => {
          units.push({ pos: pos + 1 + o, node: row, textblock: false, section, forceBreakBefore: o === 0 && breakNext, rowCols: Math.max(1, cols), tablePos: pos })
        })
        breakNext = false
        return
      }
      if (n.isTextblock || n.isAtom || n.type.name === 'horizontalRule') {
        units.push({ pos, node: n, textblock: n.isTextblock, section, forceBreakBefore: breakNext || !!n.attrs?.pageBreakBefore })
        breakNext = false
        return
      }
      // container (list, list item, blockquote): descend
      n.forEach((c, o) => visit(c, pos + 1 + o))
    }
    visit(node, offset)
  })
  return units
}

function spacer(height: number, kind: 'block' | 'inline' | 'row', page: number, cols = 1) {
  return () => {
    if (kind === 'row') {
      const tr = document.createElement('tr')
      tr.className = 'sc-page-spacer sc-page-spacer-row'
      tr.setAttribute('contenteditable', 'false')
      tr.setAttribute('aria-hidden', 'true')
      tr.dataset.page = String(page)
      const td = document.createElement('td')
      td.colSpan = cols
      td.style.height = `${Math.max(0, height)}px`
      tr.appendChild(td)
      return tr
    }
    const el = document.createElement(kind === 'block' ? 'div' : 'span')
    el.className = 'sc-page-spacer'
    el.setAttribute('contenteditable', 'false')
    el.setAttribute('aria-hidden', 'true')
    el.dataset.page = String(page)
    el.style.display = 'block'
    el.style.height = `${Math.max(0, height)}px`
    return el
  }
}

export function computeLayout(view: EditorView, opts: PaginationOptions): { decos: Decoration[]; pages: PageInfo[] } {
  const sections = opts.getSections()
  const gap = opts.getGap()
  const scale = opts.getScale() || 1
  const units = collectUnits(view.state.doc)
  const root = view.dom as HTMLElement
  const originClient = root.getBoundingClientRect().top + (parseFloat(getComputedStyle(root).paddingTop) || 0) * scale
  /** client px → unscaled layout px relative to the content box (without spacers) */
  const toLayout = (clientY: number) => (clientY - originClient) / scale
  const toClient = (layoutY: number) => originClient + layoutY * scale
  const pages: PageInfo[] = []
  const g0 = pageGeom(sections[0])
  const firstWidth = g0.width - g0.marginLeft - g0.marginRight
  const addPage = (section: number, firstOfSection: boolean) => {
    const g = pageGeom(sections[section] ?? sections[sections.length - 1])
    const prev = pages[pages.length - 1]
    const top = prev ? prev.top + prev.height + gap : -g.marginTop
    const restart = sections[section]?.pageNumberStart
    const number = firstOfSection && restart != null ? restart : (prev?.number ?? 0) + 1
    pages.push({ index: pages.length, section, top, ...g, firstOfSection, number, notes: [], noteSpace: 0 })
    return pages[pages.length - 1]
  }
  let page = addPage(0, true)
  const contentTop = (p: PageInfo) => p.top + p.marginTop
  const contentBottom = (p: PageInfo) => p.top + p.height - p.marginBottom - p.noteSpace
  // footnotes: each reference reserves the height of its note on the page where it lands
  const noteH = (id: string) => Math.max(0, opts.getNoteHeight?.(id) ?? 0)
  const reserve = (p: PageInfo, id: string) => {
    if (p.notes.includes(id)) return
    if (!p.notes.length) p.noteSpace += NOTE_SEPARATOR
    p.notes.push(id)
    p.noteSpace += noteH(id)
  }
  const release = (p: PageInfo, id: string) => {
    const i = p.notes.indexOf(id)
    if (i === -1) return
    p.notes.splice(i, 1)
    p.noteSpace -= noteH(id)
    if (!p.notes.length) p.noteSpace = 0
  }
  let lastBottom = 0
  const decos: Decoration[] = []
  let offset = 0

  for (const u of units) {
    if (u.node.type.name === 'pageBreak' || u.node.type.name === 'sectionBreak') continue
    const dom = view.nodeDOM(u.pos) as HTMLElement | null
    if (!dom || !(dom instanceof HTMLElement)) continue
    const rect = dom.getBoundingClientRect()
    if (!rect.height && !rect.width) continue
    let top = toLayout(rect.top) + offset
    let bottom = toLayout(rect.bottom) + offset
    // section width differs from the first section (e.g. landscape) → widen the block
    const sec = sections[u.section]
    if (sec && u.section !== 0) {
      const g = pageGeom(sec)
      const w = g.width - g.marginLeft - g.marginRight
      if (Math.abs(w - firstWidth) > 1 && u.pos >= 0 && u.rowCols === undefined) decos.push(Decoration.node(u.pos, u.pos + u.node.nodeSize, { style: `width:${w}px;max-width:none` }))
    }
    const newPageFor = (section: number) => {
      const first = section !== page.section
      page = addPage(section, first)
    }
    // footnote references inside this block: [document position, note id]
    const refs: [number, string][] = []
    if (opts.getNoteHeight)
      u.node.descendants((n, p) => {
        if (n.type.name === 'noteRef' && n.attrs.kind !== 'endnote') refs.push([u.pos + 1 + p, String(n.attrs.noteId)])
      })
    let refPage: PageInfo
    if (u.forceBreakBefore || u.section !== page.section) {
      if (top > contentTop(page) - 0.5 || u.section !== page.section) {
        newPageFor(u.section)
        const h = contentTop(page) - top
        decos.push(Decoration.widget(u.pos, u.rowCols ? spacer(h, 'row', page.index, u.rowCols) : spacer(h, 'block', page.index), { side: -1, key: `pb-${u.pos}-${Math.round(h)}`, ignoreSelection: true }))
        offset += h
        top += h
        bottom += h
      }
    }
    refPage = page
    for (const [, id] of refs) reserve(page, id)
    // overflow: move or split
    while (bottom > contentBottom(page) + 0.5) {
      if (top <= contentTop(page) + 0.5 && !u.textblock) {
        // atomic block taller than the page: let it overflow, continue on following pages
        while (bottom > contentBottom(page) + 0.5) newPageFor(u.section)
        break
      }
      let splitPos: number | null = null
      let lineTop = 0
      if (u.textblock && u.node.content.size > 0) {
        const r = splitTextblock(view, u, dom, toClient(contentBottom(page) - offset))
        if (r) {
          splitPos = r.pos
          lineTop = toLayout(r.lineTop) + offset
        }
      }
      const prevPage = page
      newPageFor(u.section)
      if (splitPos != null && lineTop > top + 0.5) {
        const h = contentTop(page) - lineTop
        for (const [p, id] of refs)
          if (p >= splitPos && refPage.notes.includes(id)) {
            release(refPage, id)
            reserve(page, id)
          }
        refPage = page
        decos.push(Decoration.widget(splitPos, spacer(h, 'inline', page.index), { side: -1, key: `ps-${splitPos}-${Math.round(h)}`, ignoreSelection: true, marks: [] }))
        offset += h
        bottom += h
        top = contentTop(page)
        // the remaining lines may again overflow: measure by line positions in the next loop iteration
        if (bottom > contentBottom(page) + 0.5) {
          // continue splitting relative to the new page
          continue
        }
      } else {
        if (top <= contentTop(prevPage) + 0.5) {
          // unsplittable and already at page top: overflow
          break
        }
        for (const [, id] of refs) {
          release(refPage, id)
          reserve(page, id)
        }
        refPage = page
        const h = contentTop(page) - top
        decos.push(Decoration.widget(u.pos, u.rowCols ? spacer(h, 'row', page.index, u.rowCols) : spacer(h, 'block', page.index), { side: -1, key: `pb-${u.pos}-${Math.round(h)}`, ignoreSelection: true }))
        offset += h
        top += h
        bottom += h
      }
    }
    lastBottom = bottom
  }
  // endnotes follow the last block, on a new page when they don't fit
  const endH = opts.getEndnotesHeight?.() ?? 0
  if (endH > 0) {
    let start = Math.max(lastBottom, contentTop(page)) + 24
    if (start + endH > contentBottom(page) && start > contentTop(page) + 24) {
      page = addPage(page.section, false)
      start = contentTop(page)
    }
    page.endnotesTop = start - page.top
  }
  return { decos, pages }
}

/** Find the document position of the first line of `u` that starts below `limitY` (client px). */
function splitTextblock(view: EditorView, u: Unit, dom: HTMLElement, limitClientY: number): { pos: number; lineTop: number } | null {
  const start = u.pos + 1
  const end = u.pos + u.node.nodeSize - 1
  // binary search positions by line bottom
  let lo = start
  let hi = end
  let best: number | null = null
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    let c
    try {
      c = view.coordsAtPos(mid, 1)
    } catch {
      break
    }
    if (c.bottom > limitClientY) {
      best = mid
      hi = mid - 1
    } else lo = mid + 1
  }
  if (best == null) return null
  // move to the start of that visual line
  const lineTopY = view.coordsAtPos(best, 1).top
  let pos = best
  while (pos > start) {
    const c = view.coordsAtPos(pos - 1, 1)
    if (c.top < lineTopY - 1) break
    pos--
  }
  if (pos <= start) return null // first line doesn't fit: move whole paragraph
  void dom
  return { pos, lineTop: lineTopY }
}

export const Pagination = Extension.create<PaginationOptions>({
  name: 'pagination',
  addOptions() {
    return { getSections: () => [], getGap: () => 24, getScale: () => 1, onLayout: () => {} }
  },
  addProseMirrorPlugins() {
    const opts = this.options
    let timer: ReturnType<typeof setTimeout> | null = null
    let lastSig = ''
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const meta = tr.getMeta(key)
            if (meta) return meta as DecorationSet
            return set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
        view(view) {
          const run = () => {
            timer = null
            if (!view.dom.isConnected) return
            const scroller = view.dom.closest('.sc-canvas') as HTMLElement | null
            const scrollTop = scroller?.scrollTop ?? 0
            // 1) remove spacers, 2) measure natural layout, 3) apply — synchronous, so nothing paints in between
            view.dispatch(view.state.tr.setMeta(key, DecorationSet.empty).setMeta('addToHistory', false))
            const { decos, pages } = computeLayout(view, opts)
            view.dispatch(view.state.tr.setMeta(key, DecorationSet.create(view.state.doc, decos)).setMeta('addToHistory', false))
            if (scroller) scroller.scrollTop = scrollTop
            const sig = JSON.stringify(pages.map((p) => [p.top, p.height, p.section, p.number, p.notes, p.noteSpace, p.endnotesTop]))
            if (sig !== lastSig) {
              lastSig = sig
              opts.onLayout(pages)
            }
          }
          const schedule = (delay: number) => {
            if (timer) clearTimeout(timer)
            timer = setTimeout(run, delay)
          }
          const onResize = () => schedule(50)
          window.addEventListener('resize', onResize)
          ;(view as any).__scribeRepaginate = () => schedule(0)
          ;(view as any).__scribeRepaginateNow = () => {
            if (timer) clearTimeout(timer)
            run()
          }
          document.fonts?.ready.then(() => schedule(0))
          schedule(0)
          return {
            update(v, prev) {
              if (v.state.doc !== prev.doc) schedule(120)
            },
            destroy() {
              if (timer) clearTimeout(timer)
              window.removeEventListener('resize', onResize)
            },
          }
        },
      }),
    ]
  },
})

export function repaginate(view: EditorView | undefined, now = false) {
  if (now) (view as any)?.__scribeRepaginateNow?.()
  else (view as any)?.__scribeRepaginate?.()
}
