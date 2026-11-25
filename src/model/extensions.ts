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
    const num = (name: string) => ({
      default: null,
      parseHTML: (el: HTMLElement) => {
        const v = el.getAttribute(`data-${name}`)
        return v == null ? null : Number(v)
      },
      renderHTML: (attrs: Record<string, any>) => (attrs[name] == null ? {} : { [`data-${name}`]: attrs[name] }),
    })
    const str = (name: string) => ({
      default: null,
      parseHTML: (el: HTMLElement) => el.getAttribute(`data-${name}`),
      renderHTML: (attrs: Record<string, any>) => (attrs[name] == null ? {} : { [`data-${name}`]: attrs[name] }),
    })
    return [
      {
        types: BLOCK_TYPES,
        attributes: {
          styleId: str('style-id'),
          indentLeft: num('indent-left'),
          indentRight: num('indent-right'),
          indentFirstLine: num('indent-first-line'),
          spacingBefore: num('spacing-before'),
          spacingAfter: num('spacing-after'),
          /** line spacing: for rule "auto" a multiple (1.15), otherwise twips */
          lineSpacing: num('line-spacing'),
          lineRule: str('line-rule'),
          pageBreakBefore: { default: null },
          keepNext: { default: null },
          shading: str('shading'),
          direction: str('dir'),
          /** Raw unknown <w:pPr> children, re-emitted verbatim on DOCX export. */
          pPrExtra: { default: null, rendered: false },
          /** Paragraph-mark run properties (affects empty-paragraph height in Word). */
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
              const css: string[] = []
              if (a.indentLeft != null) css.push(`margin-left:${twipToPx(a.indentLeft)}px`)
              if (a.indentRight != null) css.push(`margin-right:${twipToPx(a.indentRight)}px`)
              if (a.indentFirstLine != null) css.push(`text-indent:${twipToPx(a.indentFirstLine)}px`)
              if (a.spacingBefore != null) css.push(`margin-top:${twipToPx(a.spacingBefore)}px`)
              if (a.spacingAfter != null) css.push(`margin-bottom:${twipToPx(a.spacingAfter)}px`)
              if (a.lineSpacing != null) {
                if (!a.lineRule || a.lineRule === 'auto') css.push(`line-height:${a.lineSpacing}`)
                else css.push(`line-height:${twipToPx(a.lineSpacing)}px`)
              }
              if (a.shading) css.push(`background-color:${a.shading}`)
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
            parseHTML: (el: HTMLElement) => el.style.letterSpacing || null,
            renderHTML: (a: Record<string, any>) => (a.letterSpacing ? { style: `letter-spacing:${a.letterSpacing}` } : {}),
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

/** Word field (PAGE, NUMPAGES, REF, TOC entries...). The cached result is shown; instr is preserved. */
export const Field = atomInline('field', { instr: { default: '' }, result: { default: '' }, rawXml: { default: null, rendered: false } }, (a) => [
  'span',
  { class: 'sc-field', title: `Field: ${a.instr}`, 'data-instr': a.instr },
  a.result || `{${String(a.instr).trim().split(/\s+/)[0] || 'FIELD'}}`,
])

export const NoteRef = atomInline('noteRef', { noteId: { default: '' }, kind: { default: 'footnote' }, label: { default: '' } }, (a) => [
  'sup',
  { class: 'sc-note-ref', 'data-note': a.noteId, title: a.kind === 'endnote' ? 'Endnote' : 'Footnote' },
  a.label || '*',
])

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

export const OpaqueInline = atomInline(
  'opaqueInline',
  { xml: { default: '', rendered: false }, label: { default: 'Object' }, preview: { default: '' } },
  (a) => ['span', { class: 'sc-chip sc-opaque-inline', title: `${a.label} (preserved from the original file)` }, a.preview || `[${a.label}]`],
)

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
    parseHTML: (el: HTMLElement) => el.style.backgroundColor || el.getAttribute('data-background'),
    renderHTML: (a: any) => (a.background ? { style: `background-color:${a.background}` } : {}),
  },
  verticalAlign: {
    default: null,
    parseHTML: (el: HTMLElement) => el.style.verticalAlign || null,
    renderHTML: (a: any) => (a.verticalAlign ? { style: `vertical-align:${a.verticalAlign}` } : {}),
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
            parseHTML: (el: HTMLElement) => el.style.listStyleType || null,
            renderHTML: (a: any) => (a.listStyle ? { style: `list-style-type:${a.listStyle}` } : {}),
          },
        },
      },
    ]
  },
})

/** Stable block identifiers: every top-level and nested block gets a persistent id. */
export const StableIds = Extension.create({
  name: 'stableIds',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading', 'table', 'bulletList', 'orderedList', 'blockquote', 'codeBlock', 'opaqueBlock', 'pageBreak', 'sectionBreak', 'image'],
        attributes: { nid: { default: null, keepOnSplit: false, renderHTML: (a: any) => (a.nid ? { 'data-nid': a.nid } : {}), parseHTML: () => null } },
      },
    ]
  },
  addProseMirrorPlugins() {
    const key = new PluginKey('stableIds')
    const types = new Set(['paragraph', 'heading', 'table', 'bulletList', 'orderedList', 'blockquote', 'codeBlock', 'opaqueBlock', 'pageBreak', 'sectionBreak'])
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
    Highlight.configure({ multicolor: true }),
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
    OpaqueBlock,
    OpaqueInline,
    CommentMark,
    InsertionMark,
    DeletionMark,
    TrackChanges,
    StableIds,
  ]
}
