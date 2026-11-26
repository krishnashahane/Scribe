import { Extension, Mark, Node, mergeAttributes } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import StarterKit from '@tiptap/starter-kit'
import { TextStyleKit } from '@tiptap/extension-text-style'
import Highlight from '@tiptap/extension-highlight'
import TextAlign from '@tiptap/extension-text-align'
import { Table, TableRow, TableCell, TableHeader } from '@tiptap/extension-table'
import Image from '@tiptap/extension-image'
import Subscript from '@tiptap/extension-subscript'
import Superscript from '@tiptap/extension-superscript'
import { twipToPx } from './types'
import { LIST_STYLES, VERTICAL_ALIGN, safeColor, safeFontFamily, safeKeyword, safeLength, safeNumber } from './safeCss'
import { chartSvg, type ChartData } from '../objects/chart'
import { ommlToMathml } from '../objects/math'
import { DEFAULT_SHAPE, SHAPE_KINDS, shapeSvg, type ShapeAttrs } from '../objects/shapes'

const BLOCK_TYPES = ['paragraph', 'heading']

/** Safe URL check for links and images (blocks javascript:, vbscript:, etc.). */
export function isSafeUrl(url: string, allowData = false): boolean {
  const u = url.trim().toLowerCase()
  if (u.startsWith('#') || u.startsWith('/')) return true
  if (allowData && /^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml);/.test(u)) return !u.startsWith('data:image/svg')
  return /^(https?:|mailto:|tel:|ftp:)/.test(u)
}

/** Paragraph-level formatting stored in OOXML units so round trips are lossless. */
export const ParagraphFormat = Extension.create({
  name: 'paragraphFormat',
  addGlobalAttributes() {
    // key: the node attribute; name: its data-* attribute in HTML
    const num = (key: string, name: string) => ({
      default: null,
      parseHTML: (el: HTMLElement) => {
        const v = el.getAttribute(`data-${name}`)
        return v == null || v === '' || !Number.isFinite(Number(v)) ? null : Number(v)
      },
      renderHTML: (attrs: Record<string, any>) => (safeNumber(attrs[key]) == null ? {} : { [`data-${name}`]: safeNumber(attrs[key]) }),
    })
    const str = (key: string, name: string) => ({
      default: null,
      parseHTML: (el: HTMLElement) => el.getAttribute(`data-${name}`),
      // only plain tokens are mirrored into data-* attributes (colours, style ids, keywords)
      renderHTML: (attrs: Record<string, any>) => (typeof attrs[key] === 'string' && /^[\w#.,() %-]{1,80}$/.test(attrs[key]) && (key !== 'shading' || safeColor(attrs[key])) ? { [`data-${name}`]: attrs[key] } : {}),
    })
    return [
      {
        types: BLOCK_TYPES,
        attributes: {
          styleId: str('styleId', 'style-id'),
          indentLeft: num('indentLeft', 'indent-left'),
          indentRight: num('indentRight', 'indent-right'),
          indentFirstLine: num('indentFirstLine', 'indent-first-line'),
          spacingBefore: num('spacingBefore', 'spacing-before'),
          spacingAfter: num('spacingAfter', 'spacing-after'),
          /** line spacing: for rule "auto" a multiple (1.15), otherwise twips */
          lineSpacing: num('lineSpacing', 'line-spacing'),
          lineRule: str('lineRule', 'line-rule'),
          pageBreakBefore: { default: null },
          keepNext: { default: null },
          shading: str('shading', 'shading'),
          direction: str('direction', 'dir'),
          // validated alignment (the text-align extension writes its value straight into a style attribute)
          textAlign: {
            default: null,
            parseHTML: (el: HTMLElement) => safeKeyword(el.style.textAlign, ['left', 'center', 'right', 'justify']),
            renderHTML: (a: Record<string, any>) => {
              const v = safeKeyword(a.textAlign, ['left', 'center', 'right', 'justify'])
              return v && v !== 'left' ? { style: `text-align:${v}` } : {}
            },
          },
          /** Raw unknown <w:pPr> children, re-emitted verbatim on DOCX export. */
          pPrExtra: { default: null, rendered: false },
          /** Paragraph-mark run properties (affects empty-paragraph height). */
          markRPr: { default: null, rendered: false },
          /** Attributes of <w:spacing>/<w:ind> that SCRIBE does not edit (e.g. beforeAutospacing). */
          spacingExtra: { default: null, rendered: false },
          indExtra: { default: null, rendered: false },
        },
      },
      {
        // style attribute assembled in one place so attributes don't clobber each other
        types: BLOCK_TYPES,
        attributes: {
          _style: {
            default: null,
            rendered: true,
            renderHTML: (a: Record<string, any>) => {
              // every value is validated: paragraph attributes come from files and collaborators
              const css: string[] = []
              const tw = (v: unknown) => safeNumber(v, -1e5, 1e5)
              const px = (v: number) => `${Math.round(twipToPx(v) * 100) / 100}px`
              if (tw(a.indentLeft) != null) css.push(`margin-left:${px(tw(a.indentLeft)!)}`)
              if (tw(a.indentRight) != null) css.push(`margin-right:${px(tw(a.indentRight)!)}`)
              if (tw(a.indentFirstLine) != null) css.push(`text-indent:${px(tw(a.indentFirstLine)!)}`)
              if (tw(a.spacingBefore) != null) css.push(`margin-top:${px(tw(a.spacingBefore)!)}`)
              if (tw(a.spacingAfter) != null) css.push(`margin-bottom:${px(tw(a.spacingAfter)!)}`)
              const ls = safeNumber(a.lineSpacing, 0, 1e5)
              if (ls != null) {
                if (!a.lineRule || a.lineRule === 'auto') css.push(`line-height:${Math.min(ls, 20)}`)
                else css.push(`line-height:${px(ls)}`)
              }
              if (safeColor(a.shading)) css.push(`background-color:${safeColor(a.shading)}`)
              if (a.direction === 'rtl') css.push('direction:rtl')
              return css.length ? { style: css.join(';') } : {}
            },
          },
        },
      },
    ]
  },
  addCommands() {
    return {
      setParagraphFormat:
        (attrs: Record<string, any>) =>
        ({ commands }: any) =>
          BLOCK_TYPES.map((t) => commands.updateAttributes(t, attrs)).some(Boolean),
      changeIndent:
        (deltaTwips: number) =>
        ({ state, tr, dispatch }: any) => {
          const { from, to } = state.selection
          let changed = false
          state.doc.nodesBetween(from, to, (node: any, pos: number) => {
            if (BLOCK_TYPES.includes(node.type.name)) {
              const next = Math.max(0, (node.attrs.indentLeft ?? 0) + deltaTwips)
              tr.setNodeMarkup(pos, undefined, { ...node.attrs, indentLeft: next || null })
              changed = true
            }
          })
          if (changed && dispatch) dispatch(tr)
          return changed
        },
    } as any
  },
})

/** Unknown run properties preserved on the textStyle mark. */
export const RunExtras = Extension.create({
  name: 'runExtras',
  addGlobalAttributes() {
    return [
      {
        types: ['textStyle'],
        attributes: {
          rPrExtra: { default: null, rendered: false },
          charStyle: { default: null, rendered: false },
          letterSpacing: {
            default: null,
            parseHTML: (el: HTMLElement) => safeLength(el.style.letterSpacing, ['pt', 'px', 'em']),
            renderHTML: (a: Record<string, any>) => (safeLength(a.letterSpacing, ['pt', 'px', 'em']) ? { style: `letter-spacing:${safeLength(a.letterSpacing, ['pt', 'px', 'em'])}` } : {}),
          },
          // the text-style kit writes these straight into style attributes: validate them (same names: these definitions take precedence)
          color: {
            default: null,
            parseHTML: (el: HTMLElement) => safeColor(el.style.color?.replace(/['"]+/g, '')),
            renderHTML: (a: Record<string, any>) => (safeColor(a.color) ? { style: `color:${safeColor(a.color)}` } : {}),
          },
          backgroundColor: {
            default: null,
            parseHTML: (el: HTMLElement) => safeColor(el.style.backgroundColor?.replace(/['"]+/g, '')),
            renderHTML: (a: Record<string, any>) => (safeColor(a.backgroundColor) ? { style: `background-color:${safeColor(a.backgroundColor)}` } : {}),
          },
          fontFamily: {
            default: null,
            parseHTML: (el: HTMLElement) => safeFontFamily(el.style.fontFamily),
            renderHTML: (a: Record<string, any>) => (safeFontFamily(a.fontFamily) ? { style: `font-family:${safeFontFamily(a.fontFamily)}` } : {}),
          },
          fontSize: {
            default: null,
            parseHTML: (el: HTMLElement) => safeLength(el.style.fontSize),
            renderHTML: (a: Record<string, any>) => (safeLength(a.fontSize) ? { style: `font-size:${safeLength(a.fontSize)}` } : {}),
          },
          caps: {
            default: null,
            parseHTML: (el: HTMLElement) => (el.style.textTransform === 'uppercase' ? 'caps' : el.style.fontVariant === 'small-caps' ? 'smallCaps' : null),
            renderHTML: (a: Record<string, any>) =>
              a.caps === 'caps' ? { style: 'text-transform:uppercase' } : a.caps === 'smallCaps' ? { style: 'font-variant:small-caps' } : {},
          },
        },
      },
    ]
  },
})

const atomInline = (name: string, attrs: Record<string, any>, render: (a: any) => [string, Record<string, any>, ...any[]]) =>
  Node.create({
    name,
    group: 'inline',
    inline: true,
    atom: true,
    selectable: true,
    addAttributes() {
      return attrs
    },
    parseHTML() {
      return [{ tag: `span[data-scribe="${name}"]` }]
    },
    renderHTML({ node }) {
      const [tag, a, ...rest] = render(node.attrs)
      return [tag, mergeAttributes({ 'data-scribe': name }, a), ...rest]
    },
  })

export const PageBreak = Node.create({
  name: 'pageBreak',
  group: 'block',
  atom: true,
  selectable: true,
  parseHTML() {
    return [{ tag: 'div[data-scribe="pageBreak"]' }, { tag: 'hr[data-page-break]' }]
  },
  renderHTML() {
    return ['div', { 'data-scribe': 'pageBreak', class: 'sc-page-break', contenteditable: 'false' }, 'Page break']
  },
  addCommands() {
    return {
      insertPageBreak:
        () =>
        ({ chain }: any) =>
          chain().insertContent([{ type: 'pageBreak' }, { type: 'paragraph' }]).run(),
    } as any
  },
  addKeyboardShortcuts() {
    return { 'Mod-Enter': () => (this.editor.commands as any).insertPageBreak() }
  },
})

export const SectionBreak = Node.create({
  name: 'sectionBreak',
  group: 'block',
  atom: true,
  selectable: true,
  addAttributes() {
    // index into DocMeta.sections of the section this break ENDS
    return { section: { default: 0 }, breakType: { default: 'nextPage' } }
  },
  parseHTML() {
    return [{ tag: 'div[data-scribe="sectionBreak"]' }]
  },
  renderHTML({ node }) {
    const label = node.attrs.breakType === 'continuous' ? 'Section break (continuous)' : 'Section break (next page)'
    return ['div', { 'data-scribe': 'sectionBreak', class: 'sc-section-break', contenteditable: 'false', 'data-section': node.attrs.section }, label]
  },
})

export const ColumnBreak = atomInline('columnBreak', {}, () => ['span', { class: 'sc-chip sc-column-break' }, '⏎ column'])

export const Bookmark = atomInline('bookmark', { name: { default: '' }, bid: { default: null } }, (a) => [
  'span',
  { class: 'sc-bookmark', id: `bm-${a.name}`, title: `Bookmark: ${a.name}` },
])

/** Document field (PAGE, NUMPAGES, REF, TOC entries...). The cached result is shown; instr is preserved. */
export const Field = atomInline('field', { instr: { default: '' }, result: { default: '' }, rawXml: { default: null, rendered: false } }, (a) => [
  'span',
  { class: 'sc-field', title: `Field: ${a.instr}`, 'data-instr': a.instr },
  a.result || `{${String(a.instr).trim().split(/\s+/)[0] || 'FIELD'}}`,
])

export const NoteRef = atomInline('noteRef', { noteId: { default: '' }, kind: { default: 'footnote' }, label: { default: '' } }, (a) => [
  'sup',
  { class: 'sc-note-ref', 'data-note': a.noteId, 'data-kind': a.kind === 'endnote' ? 'endnote' : 'footnote', title: a.kind === 'endnote' ? 'Endnote' : 'Footnote' },
  a.label || '*',
])

function roman(n: number): string {
  const map: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]
  let out = ''
  for (const [v, s] of map) while (n >= v) {
    out += s
    n -= v
  }
  return out
}

/** Footnote and endnote references are numbered in document order (1, 2, 3… per kind). */
export const NoteNumbering = Extension.create({
  name: 'noteNumbering',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('noteNumbering'),
        appendTransaction(trs, _old, state) {
          if (!trs.some((t) => t.docChanged)) return null
          const counters: Record<string, number> = { footnote: 0, endnote: 0 }
          const tr = state.tr
          let changed = false
          state.doc.descendants((node, pos) => {
            if (node.type.name !== 'noteRef') return
            const kind = node.attrs.kind === 'endnote' ? 'endnote' : 'footnote'
            const n = ++counters[kind]
            // endnotes use lowercase roman numerals, the default numbering for endnotes in DOCX files
            const label = kind === 'endnote' ? roman(n) : String(n)
            if (node.attrs.label !== label) {
              tr.setNodeMarkup(pos, undefined, { ...node.attrs, label })
              changed = true
            }
          })
          if (!changed) return null
          tr.setMeta('addToHistory', false)
          return tr
        },
      }),
    ]
  },
})

/** Preserved OOXML the editor does not understand (shapes, charts, equations, content controls...). */
export const OpaqueBlock = Node.create({
  name: 'opaqueBlock',
  group: 'block',
  atom: true,
  selectable: true,
  draggable: true,
  addAttributes() {
    return { xml: { default: '', rendered: false }, label: { default: 'Preserved content' }, preview: { default: '' } }
  },
  parseHTML() {
    return [{ tag: 'div[data-scribe="opaqueBlock"]' }]
  },
  renderHTML({ node }) {
    return [
      'div',
      { 'data-scribe': 'opaqueBlock', class: 'sc-opaque', contenteditable: 'false', title: 'Preserved from the original file. It will be written back unchanged.' },
      ['span', { class: 'sc-opaque-label' }, node.attrs.label],
      ['span', { class: 'sc-opaque-preview' }, node.attrs.preview || ''],
    ]
  },
})

const svgUrl = (svg: string) => `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`

/** Editable data extracted from a preserved object (chart data, diagram text, equation, text box). */
export type ObjectInfo =
  | { kind: 'chart'; chart: ChartData; width: number; height: number }
  | { kind: 'smartart'; smartArt: import('../objects/smartart').SmartArtData }
  | { kind: 'textbox' }
  | { kind: 'equation' }

function opaqueInlineSpec(a: Record<string, any>): [string, Record<string, any>, ...any[]] {
  const obj = a.object as ObjectInfo | null
  if (obj?.kind === 'chart' && obj.chart) {
    const w = Math.max(60, Math.round(obj.width || 480))
    const h = Math.max(40, Math.round(obj.height || 280))
    try {
      return ['img', { class: 'sc-chart', src: svgUrl(chartSvg(obj.chart, w, h)), width: w, height: h, alt: String(obj.chart.title || 'Chart'), title: 'Chart: double-click to edit its data' }]
    } catch {
      /* malformed chart data: fall back to the placeholder chip */
    }
  }
  const hint = obj?.kind === 'smartart' ? ' (double-click to edit its text)' : obj?.kind === 'textbox' ? ' (double-click to edit)' : obj?.kind === 'equation' ? ' (double-click to edit)' : ''
  return ['span', { class: `sc-chip sc-opaque-inline${obj ? ' is-editable' : ''}`, title: `${a.label}${hint || ' (kept from the original file)'}` }, a.preview || `[${a.label}]`]
}

export const OpaqueInline = Node.create({
  name: 'opaqueInline',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      xml: { default: '', rendered: false },
      label: { default: 'Object' },
      preview: { default: '' },
      object: { default: null, rendered: false },
    }
  },
  parseHTML() {
    return [{ tag: 'span[data-scribe="opaqueInline"]' }]
  },
  renderHTML({ node }) {
    const [tag, a, ...rest] = opaqueInlineSpec(node.attrs)
    return [tag, mergeAttributes({ 'data-scribe': 'opaqueInline' }, a), ...rest]
  },
  addNodeView() {
    return ({ node }) => {
      const obj = node.attrs.object as ObjectInfo | null
      if (obj?.kind === 'equation') {
        const mathml = ommlToMathml(String(node.attrs.xml))
        if (mathml) {
          const dom = document.createElement('span')
          dom.className = 'sc-equation'
          dom.setAttribute('data-scribe', 'opaqueInline')
          dom.title = 'Equation (double-click to edit)'
          // MathML is generated from escaped equation text only (see objects/math.ts)
          dom.innerHTML = mathml
          return { dom, update: () => false, ignoreMutation: () => true }
        }
      }
      const [tag, a, ...rest] = opaqueInlineSpec(node.attrs)
      const dom = document.createElement(tag)
      for (const [k, v] of Object.entries(a)) dom.setAttribute(k, String(v))
      dom.setAttribute('data-scribe', 'opaqueInline')
      if (typeof rest[0] === 'string') dom.textContent = rest[0]
      return { dom, update: () => false, ignoreMutation: () => true }
    }
  },
})

/** A drawing shape created in SCRIBE (exported as a native DrawingML shape). */
export const Shape = Node.create({
  name: 'shape',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: true,
  addAttributes() {
    return {
      kind: { default: 'rect' },
      width: { default: DEFAULT_SHAPE.width },
      height: { default: DEFAULT_SHAPE.height },
      fill: { default: DEFAULT_SHAPE.fill },
      stroke: { default: DEFAULT_SHAPE.stroke },
      strokeWidth: { default: DEFAULT_SHAPE.strokeWidth },
      text: { default: '' },
      arrow: { default: false },
    }
  },
  parseHTML() {
    return [{ tag: 'img[data-scribe="shape"]' }]
  },
  renderHTML({ node }) {
    const a = node.attrs as ShapeAttrs
    const kind = a.kind in SHAPE_KINDS ? a.kind : 'rect'
    const svg = shapeSvg({ ...a, kind })
    return ['img', { 'data-scribe': 'shape', class: 'sc-shape', src: svgUrl(svg), width: Math.round(a.width), height: Math.round(a.height), alt: a.text || SHAPE_KINDS[kind], title: `${SHAPE_KINDS[kind]} (double-click to edit)` }]
  },
})

export interface TocEntry {
  level: number
  text: string
  /** page number shown ('' until the table is updated) */
  page: string
  /** stable id (nid) of the heading this entry points to */
  target?: string
}

/**
 * Table of contents built from the document's headings. An imported table keeps its original
 * XML (rawXml) and is written back unchanged until it is updated in SCRIBE.
 */
export const TableOfContents = Node.create({
  name: 'toc',
  group: 'block',
  atom: true,
  selectable: true,
  addAttributes() {
    return {
      title: { default: 'Contents' },
      levels: { default: 3 },
      entries: { default: [], rendered: false },
      rawXml: { default: null, rendered: false },
      /** set when SCRIBE regenerated the entries: the exporter writes a new table */
      updated: { default: false, rendered: false },
    }
  },
  parseHTML() {
    return [{ tag: 'div[data-scribe="toc"]' }]
  },
  renderHTML({ node }) {
    const entries = (Array.isArray(node.attrs.entries) ? node.attrs.entries : []) as TocEntry[]
    return [
      'div',
      { 'data-scribe': 'toc', class: 'sc-toc', contenteditable: 'false', title: 'Table of contents (References → Update table). Click an entry to go to it.' },
      ['p', { class: 'sc-toc-title' }, String(node.attrs.title || 'Contents')],
      ...(entries.length
        ? entries.map((e) => [
            'p',
            { class: `sc-toc-entry is-level-${Math.min(9, Math.max(1, Number(e.level) || 1))}`, 'data-toc-target': e.target ?? '' },
            ['span', { class: 'sc-toc-text' }, String(e.text)],
            ['span', { class: 'sc-toc-leader' }],
            ['span', { class: 'sc-toc-page' }, String(e.page ?? '')],
          ])
        : [['p', { class: 'sc-toc-entry sc-muted' }, 'No headings yet. Use Heading styles, then choose Update table.']]),
    ]
  },
})

/** Bibliography generated from the document's sources. entries: [[{text, italic}]] */
export const Bibliography = Node.create({
  name: 'bibliography',
  group: 'block',
  atom: true,
  selectable: true,
  addAttributes() {
    return { title: { default: 'References' }, entries: { default: [], rendered: false } }
  },
  parseHTML() {
    return [{ tag: 'div[data-scribe="bibliography"]' }]
  },
  renderHTML({ node }) {
    const entries = (Array.isArray(node.attrs.entries) ? node.attrs.entries : []) as { text: string; italic?: boolean }[][]
    return [
      'div',
      { 'data-scribe': 'bibliography', class: 'sc-bibliography', contenteditable: 'false', title: 'Bibliography (updates from References → Sources)' },
      ['p', { class: 'sc-bib-title' }, String(node.attrs.title || 'References')],
      ...(entries.length
        ? entries.map((e) => ['p', { class: 'sc-bib-entry' }, ...e.map((seg) => (seg.italic ? ['em', {}, String(seg.text)] : ['span', {}, String(seg.text)]))])
        : [['p', { class: 'sc-bib-entry sc-muted' }, 'No sources yet. Add sources from the References tab.']]),
    ]
  },
})

export const CommentMark = Mark.create({
  name: 'comment',
  inclusive: false,
  excludes: '',
  addAttributes() {
    return { commentId: { default: null, parseHTML: (el) => el.getAttribute('data-comment') } }
  },
  parseHTML() {
    return [{ tag: 'span[data-comment]' }]
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes, { class: 'sc-comment', 'data-comment': HTMLAttributes.commentId })]
  },
})

const revisionAttrs = () => ({
  author: { default: '' },
  date: { default: '' },
  revId: { default: null },
})

export const InsertionMark = Mark.create({
  name: 'insertion',
  excludes: 'deletion',
  addAttributes: revisionAttrs,
  parseHTML() {
    return [{ tag: 'ins' }]
  },
  renderHTML({ HTMLAttributes }) {
    return ['ins', mergeAttributes(HTMLAttributes, { class: 'sc-ins', title: `Inserted by ${HTMLAttributes.author || 'unknown'}` })]
  },
})

export const DeletionMark = Mark.create({
  name: 'deletion',
  excludes: 'insertion',
  addAttributes: revisionAttrs,
  parseHTML() {
    return [{ tag: 'del' }]
  },
  renderHTML({ HTMLAttributes }) {
    return ['del', mergeAttributes(HTMLAttributes, { class: 'sc-del', title: `Deleted by ${HTMLAttributes.author || 'unknown'}` })]
  },
})

/** Accept / reject tracked changes. */
export const TrackChanges = Extension.create({
  name: 'trackChanges',
  addCommands() {
    const resolve =
      (accept: boolean, onlyInSelection: boolean) =>
      ({ state, tr, dispatch }: any) => {
        const ins = state.schema.marks.insertion
        const del = state.schema.marks.deletion
        const from = onlyInSelection ? state.selection.from : 0
        const to = onlyInSelection ? state.selection.to : state.doc.content.size
        const removals: [number, number][] = []
        state.doc.nodesBetween(from, to, (node: any, pos: number) => {
          if (!node.isText) return
          const s = Math.max(pos, from)
          const e = Math.min(pos + node.nodeSize, to)
          const isIns = node.marks.some((m: any) => m.type === ins)
          const isDel = node.marks.some((m: any) => m.type === del)
          if ((accept && isDel) || (!accept && isIns)) removals.push([s, e])
        })
        tr.removeMark(from, to, ins)
        tr.removeMark(from, to, del)
        for (const [s, e] of removals.reverse()) tr.delete(tr.mapping.map(s), tr.mapping.map(e))
        if (dispatch) dispatch(tr)
        return true
      }
    return {
      acceptAllChanges: () => resolve(true, false),
      rejectAllChanges: () => resolve(false, false),
      acceptChange: () => resolve(true, true),
      rejectChange: () => resolve(false, true),
    } as any
  },
})

export const ScribeImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      width: { default: null },
      height: { default: null },
      /** Original <wp:anchor> XML for floating images, re-emitted on export. */
      anchorXml: { default: null, rendered: false },
      float: {
        default: null,
        renderHTML: (a: any) => (a.float ? { 'data-float': a.float, class: `sc-img-float-${a.float}` } : {}),
        parseHTML: (el: HTMLElement) => el.getAttribute('data-float'),
      },
      docPrName: { default: null, rendered: false },
    }
  },
}).configure({ inline: true, allowBase64: true })

const cellAttrs = {
  background: {
    default: null,
    parseHTML: (el: HTMLElement) => safeColor(el.style.backgroundColor || el.getAttribute('data-background')),
    renderHTML: (a: any) => (safeColor(a.background) ? { style: `background-color:${safeColor(a.background)}` } : {}),
  },
  verticalAlign: {
    default: null,
    parseHTML: (el: HTMLElement) => safeKeyword(el.style.verticalAlign, VERTICAL_ALIGN),
    renderHTML: (a: any) => (safeKeyword(a.verticalAlign, VERTICAL_ALIGN) ? { style: `vertical-align:${safeKeyword(a.verticalAlign, VERTICAL_ALIGN)}` } : {}),
  },
  tcPrExtra: { default: null, rendered: false },
}

export const ScribeTable = Table.extend({
  addAttributes() {
    return { ...this.parent?.(), tblPrExtra: { default: null, rendered: false }, tblStyle: { default: null, rendered: false }, gridXml: { default: null, rendered: false } }
  },
}).configure({ resizable: true, allowTableNodeSelection: true })
export const ScribeTableRow = TableRow.extend({
  addAttributes() {
    return { ...this.parent?.(), trPrExtra: { default: null, rendered: false } }
  },
})
export const ScribeTableCell = TableCell.extend({
  addAttributes() {
    return { ...this.parent?.(), ...cellAttrs }
  },
})
export const ScribeTableHeader = TableHeader.extend({
  addAttributes() {
    return { ...this.parent?.(), ...cellAttrs }
  },
})

/** Lists remember their OOXML numbering definition so imported numbering round-trips. */
export const ListNumbering = Extension.create({
  name: 'listNumbering',
  addGlobalAttributes() {
    return [
      {
        types: ['bulletList', 'orderedList'],
        attributes: {
          numId: { default: null, rendered: false },
          listStyle: {
            default: null,
            parseHTML: (el: HTMLElement) => safeKeyword(el.style.listStyleType, LIST_STYLES),
            renderHTML: (a: any) => (safeKeyword(a.listStyle, LIST_STYLES) ? { style: `list-style-type:${safeKeyword(a.listStyle, LIST_STYLES)}` } : {}),
          },
        },
      },
    ]
  },
})

/** Highlight with a validated colour (the stock mark writes the value straight into a style attribute). */
export const SafeHighlight = Highlight.extend({
  addAttributes() {
    return {
      color: {
        default: null,
        parseHTML: (el: HTMLElement) => safeColor(el.getAttribute('data-color') || el.style.backgroundColor),
        renderHTML: (a: Record<string, any>) => (safeColor(a.color) ? { 'data-color': safeColor(a.color), style: `background-color:${safeColor(a.color)};color:inherit` } : {}),
      },
    }
  },
})

/** Stable block identifiers: every top-level and nested block gets a persistent id. */
export const StableIds = Extension.create({
  name: 'stableIds',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading', 'table', 'bulletList', 'orderedList', 'blockquote', 'codeBlock', 'opaqueBlock', 'pageBreak', 'sectionBreak', 'image', 'bibliography', 'toc'],
        attributes: { nid: { default: null, keepOnSplit: false, renderHTML: (a: any) => (a.nid ? { 'data-nid': a.nid } : {}), parseHTML: () => null } },
      },
    ]
  },
  addProseMirrorPlugins() {
    const key = new PluginKey('stableIds')
    const types = new Set(['paragraph', 'heading', 'table', 'bulletList', 'orderedList', 'blockquote', 'codeBlock', 'opaqueBlock', 'pageBreak', 'sectionBreak', 'bibliography', 'toc'])
    return [
      new Plugin({
        key,
        appendTransaction(transactions, _old, state) {
          if (!transactions.some((t) => t.docChanged)) return null
          const seen = new Set<string>()
          const tr = state.tr
          let changed = false
          state.doc.descendants((node, pos) => {
            if (!types.has(node.type.name)) return
            const id = node.attrs.nid
            if (!id || seen.has(id)) {
              tr.setNodeAttribute(pos, 'nid', newId())
              changed = true
            } else seen.add(id)
          })
          if (!changed) return null
          tr.setMeta('addToHistory', false)
          return tr
        },
      }),
    ]
  },
})

export function newId(): string {
  const a = new Uint8Array(6)
  crypto.getRandomValues(a)
  return Array.from(a, (b) => b.toString(36).padStart(2, '0')).join('').slice(0, 10)
}

/** All extensions that define the document schema (no UI-only plugins). */
export function schemaExtensions(opts: { collaborative?: boolean } = {}) {
  return [
    StarterKit.configure({
      undoRedo: opts.collaborative ? false : { depth: 500 },
      link: {
        openOnClick: false,
        autolink: true,
        protocols: ['mailto', 'tel'],
        isAllowedUri: (url: string) => isSafeUrl(url),
      },
      heading: { levels: [1, 2, 3, 4, 5, 6] },
    }),
    TextStyleKit.configure({ lineHeight: false }),
    SafeHighlight.configure({ multicolor: true }),
    TextAlign.configure({ types: ['heading', 'paragraph'], alignments: ['left', 'center', 'right', 'justify'] }),
    ScribeTable,
    ScribeTableRow,
    ScribeTableCell,
    ScribeTableHeader,
    ScribeImage,
    Subscript,
    Superscript,
    ParagraphFormat,
    RunExtras,
    ListNumbering,
    PageBreak,
    SectionBreak,
    ColumnBreak,
    Bookmark,
    Field,
    NoteRef,
    NoteNumbering,
    OpaqueBlock,
    OpaqueInline,
    Shape,
    Bibliography,
    TableOfContents,
    CommentMark,
    InsertionMark,
    DeletionMark,
    TrackChanges,
    StableIds,
  ]
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    paragraphFormat: {
      /** Set paragraph-level OOXML formatting attributes on the selected blocks */
      setParagraphFormat: (attrs: Record<string, any>) => ReturnType
      /** Shift the left indent of the selected blocks by a number of twips */
      changeIndent: (deltaTwips: number) => ReturnType
    }
    pageBreak: {
      insertPageBreak: () => ReturnType
    }
    trackChanges: {
      acceptAllChanges: () => ReturnType
      rejectAllChanges: () => ReturnType
      acceptChange: () => ReturnType
      rejectChange: () => ReturnType
    }
  }
}
