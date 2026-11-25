import { useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useEditorState } from '@tiptap/react'
import {
  AlignCenter, AlignJustify, AlignLeft, AlignRight, Baseline, Bold, Bookmark, CheckCheck, Columns2, Highlighter, ImagePlus,
  IndentDecrease, IndentIncrease, Italic, Link, List, ListOrdered, MessageSquarePlus, Minus, PenLine, Printer, Redo2,
  RemoveFormatting, Ruler, Search, SeparatorHorizontal, Strikethrough, Subscript, Superscript, Table, Underline, Undo2, Ban,
  ZoomIn, ZoomOut, Moon, Sun, PanelLeft, Hash, Keyboard, FileWarning, Pilcrow, ClipboardPaste, Copy, Scissors, Paintbrush,
  AArrowUp, AArrowDown,
} from 'lucide-react'
import { sanitizeHtml } from '../formats/formats'
import { Menu } from './kit'
import type { StyleDef } from '../model/types'
import { IconButton, mod, useToast } from './kit'
import { isSuggesting } from '../editor/plugins'

const FONTS = ['Calibri', 'Arial', 'Cambria', 'Georgia', 'Garamond', 'Helvetica', 'Times New Roman', 'Verdana', 'Trebuchet MS', 'Segoe UI', 'Courier New', 'Consolas']
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]
const SPACING = [1, 1.15, 1.5, 2, 2.5, 3]

export type RibbonAction =
  | 'image' | 'link' | 'bookmark' | 'comment' | 'pageSetup' | 'header' | 'footer' | 'sectionBreak' | 'find' | 'wordCount'
  | 'print' | 'shortcuts' | 'report' | 'toggleSuggest' | 'acceptAll' | 'rejectAll' | 'zoomIn' | 'zoomOut' | 'zoomReset'
  | 'theme' | 'panel' | 'pageNumber' | 'orientation' | 'margins-normal' | 'margins-narrow' | 'margins-wide' | 'size-A4' | 'size-Letter' | 'marks'
  | 'formatPainter'

const TABS = ['Home', 'Insert', 'Layout', 'Review', 'View'] as const

export function Ribbon(props: {
  editor: Editor
  readOnly: boolean
  styles: StyleDef[]
  defaults: { fontFamily: string; fontSizePt: number }
  zoom: number
  theme: 'light' | 'dark'
  hasOriginal: boolean
  showMarks: boolean
  onAction: (a: RibbonAction) => void
  painterActive: boolean
}) {
  const toast = useToast()
  const { editor, readOnly, onAction } = props
  const [tab, setTab] = useState<(typeof TABS)[number]>('Home')
  const tabRefs = useRef<(HTMLButtonElement | null)[]>([])
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const ts = e.getAttributes('textStyle')
      const para = e.isActive('heading') ? e.getAttributes('heading') : e.getAttributes('paragraph')
      return {
        bold: e.isActive('bold'), italic: e.isActive('italic'), underline: e.isActive('underline'), strike: e.isActive('strike'),
        sub: e.isActive('subscript'), sup: e.isActive('superscript'), bullet: e.isActive('bulletList'), ordered: e.isActive('orderedList'),
        link: e.isActive('link'), table: e.isActive('table'),
        align: (para.textAlign as string) || 'left',
        heading: e.isActive('heading') ? (e.getAttributes('heading').level as number) : 0,
        styleId: (para.styleId as string) || '',
        font: (ts.fontFamily as string) || '',
        size: ts.fontSize ? parseFloat(ts.fontSize) : null,
        color: (ts.color as string) || '#000000',
        lineSpacing: para.lineRule === 'auto' || !para.lineRule ? (para.lineSpacing as number | null) : null,
        canUndo: e.can().undo?.() ?? false,
        canRedo: e.can().redo?.() ?? false,
        suggesting: isSuggesting(e.state),
      }
    },
  })
  const c = () => editor.chain().focus()
  const stepFont = (dir: 1 | -1) => {
    const cur = s.size ?? props.defaults.fontSizePt
    const next = dir > 0 ? SIZES.find((x) => x > cur) ?? cur + 12 : [...SIZES].reverse().find((x) => x < cur) ?? Math.max(1, cur - 1)
    c().setFontSize(`${next}pt`).run()
  }
  const changeCase = (mode: 'upper' | 'lower' | 'sentence' | 'title') => {
    const { from, to, empty } = editor.state.selection
    if (empty) return toast('Select the text to change.', 'info')
    const tr = editor.state.tr
    let atSentenceStart = true
    editor.state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.isTextblock) atSentenceStart = true
      if (!node.isText) return
      const s0 = Math.max(from, pos)
      const e0 = Math.min(to, pos + node.nodeSize)
      const text = node.text!.slice(s0 - pos, e0 - pos)
      let out = text
      if (mode === 'upper') out = text.toLocaleUpperCase()
      else if (mode === 'lower') out = text.toLocaleLowerCase()
      else if (mode === 'title') out = text.toLocaleLowerCase().replace(/(^|[\s\-–—"“(])(\p{L})/gu, (_m, a, b) => a + b.toLocaleUpperCase())
      else {
        out = ''
        for (const ch of text.toLocaleLowerCase()) {
          if (atSentenceStart && /\p{L}/u.test(ch)) {
            out += ch.toLocaleUpperCase()
            atSentenceStart = false
          } else {
            out += ch
            if (/[.!?]/.test(ch)) atSentenceStart = true
          }
        }
      }
      if (out !== text) tr.replaceWith(tr.mapping.map(s0), tr.mapping.map(e0), editor.schema.text(out, node.marks))
    })
    editor.view.dispatch(tr)
  }
  const clip = async (op: 'cut' | 'copy' | 'paste') => {
    editor.commands.focus()
    if (op !== 'paste') {
      if (!document.execCommand(op)) toast(`Use ${mod}${op === 'cut' ? 'X' : 'C'} to ${op}.`, 'info')
      return
    }
    try {
      if ('read' in navigator.clipboard) {
        const items = await navigator.clipboard.read()
        for (const item of items) {
          if (item.types.includes('text/html')) {
            const html = await (await item.getType('text/html')).text()
            c().insertContent(sanitizeHtml(html)).run()
            return
          }
        }
      }
      const text = await navigator.clipboard.readText()
      if (text) c().insertContent(text.split(/\r?\n/).map((line) => ({ type: 'paragraph', content: line ? [{ type: 'text', text: line }] : [] }))).run()
    } catch {
      toast(`The browser blocked clipboard access. Use ${mod}V to paste.`, 'info')
    }
  }
  const docFonts = [...new Set(props.styles.map((x) => x.fontFamily).filter(Boolean) as string[])]
  const fonts = [...new Set([props.defaults.fontFamily, ...docFonts, ...FONTS])]
  const paraStyles = props.styles.filter((x) => x.type === 'paragraph' && !/^heading\s*[1-6]$/i.test(x.name) && x.id !== 'Normal')

  const styleValue = s.heading ? `h${s.heading}` : s.styleId ? `s:${s.styleId}` : 'p'
  const applyStyle = (v: string) => {
    if (v === 'p') c().setParagraph().setParagraphFormat({ styleId: null }).run()
    else if (v.startsWith('h')) c().setHeading({ level: Number(v.slice(1)) as any }).run()
    else c().setParagraph().setParagraphFormat({ styleId: v.slice(2) }).run()
  }

  const onTabKey = (e: React.KeyboardEvent, i: number) => {
    const d = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0
    if (!d) return
    const n = (i + d + TABS.length) % TABS.length
    setTab(TABS[n])
    tabRefs.current[n]?.focus()
  }

  const dis = readOnly
  return (
    <div className="sc-ribbon">
      <div className="sc-tabs" role="tablist" aria-label="Ribbon">
        {TABS.map((t, i) => (
          <button key={t} ref={(el) => { tabRefs.current[i] = el }} role="tab" aria-selected={tab === t} tabIndex={tab === t ? 0 : -1} onKeyDown={(e) => onTabKey(e, i)} onClick={() => setTab(t)}>
            {t}
          </button>
        ))}
      </div>
      <div className="sc-toolbar" role="toolbar" aria-label={`${tab} tools`}>
        {tab === 'Home' && (
          <>
            <div className="sc-group" aria-label="Clipboard">
              <IconButton label="Paste" shortcut={`${mod}V`} icon={<ClipboardPaste size={17} />} disabled={dis} onClick={() => clip('paste')} />
              <IconButton label="Cut" shortcut={`${mod}X`} icon={<Scissors size={17} />} disabled={dis} onClick={() => clip('cut')} />
              <IconButton label="Copy" shortcut={`${mod}C`} icon={<Copy size={17} />} onClick={() => clip('copy')} />
              <IconButton label="Format painter" icon={<Paintbrush size={17} />} active={props.painterActive} disabled={dis} onClick={() => onAction('formatPainter')} />
            </div>
            <div className="sc-group">
              <IconButton label="Undo" shortcut={`${mod}Z`} icon={<Undo2 size={17} />} disabled={dis || !s.canUndo} onClick={() => c().undo().run()} />
              <IconButton label="Redo" shortcut={`${mod}Shift+Z`} icon={<Redo2 size={17} />} disabled={dis || !s.canRedo} onClick={() => c().redo().run()} />
            </div>
            <div className="sc-group">
              <select className="sc-select is-style" aria-label="Paragraph style" value={styleValue} disabled={dis} onChange={(e) => applyStyle(e.target.value)}>
                <option value="p">Normal text</option>
                {[1, 2, 3, 4, 5, 6].map((l) => (
                  <option key={l} value={`h${l}`}>
                    Heading {l}
                  </option>
                ))}
                {paraStyles.map((st) => (
                  <option key={st.id} value={`s:${st.id}`}>
                    {st.name}
                  </option>
                ))}
              </select>
              <select className="sc-select is-font" aria-label="Font" value={s.font || ''} disabled={dis} onChange={(e) => (e.target.value ? c().setFontFamily(e.target.value).run() : c().unsetFontFamily().run())}>
                <option value="">{props.defaults.fontFamily} (default)</option>
                {fonts.map((f) => (
                  <option key={f} value={f} style={{ fontFamily: f }}>
                    {f}
                  </option>
                ))}
              </select>
              <select className="sc-select is-size" aria-label="Font size" value={s.size ?? ''} disabled={dis} onChange={(e) => (e.target.value ? c().setFontSize(`${e.target.value}pt`).run() : c().unsetFontSize().run())}>
                <option value="">{props.defaults.fontSizePt}</option>
                {[...new Set([...SIZES, ...(s.size ? [s.size] : [])])].sort((a, b) => a - b).map((n) => (
                  <option key={n} value={n}>
                    {n}
                  </option>
                ))}
              </select>
            </div>
            <div className="sc-group">
              <IconButton label="Grow font" shortcut={`${mod}Shift+.`} icon={<AArrowUp size={17} />} disabled={dis} onClick={() => stepFont(1)} />
              <IconButton label="Shrink font" shortcut={`${mod}Shift+,`} icon={<AArrowDown size={17} />} disabled={dis} onClick={() => stepFont(-1)} />
              <Menu
                ariaLabel="Change case"
                className="sc-case-menu"
                label={<span title="Change case" className="sc-case-label">Aa</span>}
                items={[
                  { label: 'Sentence case.', onClick: () => changeCase('sentence'), disabled: dis },
                  { label: 'lowercase', onClick: () => changeCase('lower'), disabled: dis },
                  { label: 'UPPERCASE', onClick: () => changeCase('upper'), disabled: dis },
                  { label: 'Capitalize Each Word', onClick: () => changeCase('title'), disabled: dis },
                ]}
              />
            </div>
            <div className="sc-group">
              <IconButton label="Bold" shortcut={`${mod}B`} icon={<Bold size={17} />} active={s.bold} disabled={dis} onClick={() => c().toggleBold().run()} />
              <IconButton label="Italic" shortcut={`${mod}I`} icon={<Italic size={17} />} active={s.italic} disabled={dis} onClick={() => c().toggleItalic().run()} />
              <IconButton label="Underline" shortcut={`${mod}U`} icon={<Underline size={17} />} active={s.underline} disabled={dis} onClick={() => c().toggleUnderline().run()} />
              <IconButton label="Strikethrough" shortcut={`${mod}Shift+X`} icon={<Strikethrough size={17} />} active={s.strike} disabled={dis} onClick={() => c().toggleStrike().run()} />
              <IconButton label="Subscript" icon={<Subscript size={17} />} active={s.sub} disabled={dis} onClick={() => c().toggleSubscript().run()} />
              <IconButton label="Superscript" icon={<Superscript size={17} />} active={s.sup} disabled={dis} onClick={() => c().toggleSuperscript().run()} />
              <label className="sc-color" title="Text colour">
                <Baseline size={17} aria-hidden="true" />
                <input type="color" aria-label="Text colour" disabled={dis} value={/^#[0-9a-f]{6}$/i.test(s.color) ? s.color : '#000000'} onChange={(e) => c().setColor(e.target.value).run()} />
              </label>
              <label className="sc-color" title="Highlight colour">
                <Highlighter size={17} aria-hidden="true" />
                <input type="color" aria-label="Highlight colour" disabled={dis} defaultValue="#ffff00" onChange={(e) => c().setHighlight({ color: e.target.value }).run()} />
              </label>
              <IconButton label="Clear formatting" icon={<RemoveFormatting size={17} />} disabled={dis} onClick={() => c().unsetAllMarks().setParagraphFormat({ indentLeft: null, indentRight: null, indentFirstLine: null, spacingBefore: null, spacingAfter: null, lineSpacing: null, lineRule: null, shading: null }).run()} />
            </div>
            <div className="sc-group">
              <IconButton label="Align left" shortcut={`${mod}Shift+L`} icon={<AlignLeft size={17} />} active={s.align === 'left'} disabled={dis} onClick={() => c().setTextAlign('left').run()} />
              <IconButton label="Center" shortcut={`${mod}Shift+E`} icon={<AlignCenter size={17} />} active={s.align === 'center'} disabled={dis} onClick={() => c().setTextAlign('center').run()} />
              <IconButton label="Align right" shortcut={`${mod}Shift+R`} icon={<AlignRight size={17} />} active={s.align === 'right'} disabled={dis} onClick={() => c().setTextAlign('right').run()} />
              <IconButton label="Justify" shortcut={`${mod}Shift+J`} icon={<AlignJustify size={17} />} active={s.align === 'justify'} disabled={dis} onClick={() => c().setTextAlign('justify').run()} />
              <select className="sc-select is-spacing" aria-label="Line spacing" value={s.lineSpacing ?? ''} disabled={dis} onChange={(e) => c().setParagraphFormat(e.target.value ? { lineSpacing: Number(e.target.value), lineRule: 'auto' } : { lineSpacing: null, lineRule: null }).run()}>
                <option value="">Line spacing</option>
                {SPACING.map((v) => (
                  <option key={v} value={v}>
                    {v.toFixed(2).replace(/0$/, '')}
                  </option>
                ))}
              </select>
            </div>
            <div className="sc-group">
              <IconButton label="Bulleted list" shortcut={`${mod}Shift+8`} icon={<List size={17} />} active={s.bullet} disabled={dis} onClick={() => c().toggleBulletList().run()} />
              <IconButton label="Numbered list" shortcut={`${mod}Shift+7`} icon={<ListOrdered size={17} />} active={s.ordered} disabled={dis} onClick={() => c().toggleOrderedList().run()} />
              <IconButton
                label="Decrease indent"
                icon={<IndentDecrease size={17} />}
                disabled={dis}
                onClick={() => (s.bullet || s.ordered ? c().liftListItem('listItem').run() : c().changeIndent(-720).run())}
              />
              <IconButton
                label="Increase indent"
                icon={<IndentIncrease size={17} />}
                disabled={dis}
                onClick={() => (s.bullet || s.ordered ? c().sinkListItem('listItem').run() : c().changeIndent(720).run())}
              />
            </div>
            <div className="sc-group">
              <IconButton label="Find and replace" shortcut={`${mod}F`} icon={<Search size={17} />} onClick={() => onAction('find')} />
            </div>
          </>
        )}
        {tab === 'Insert' && (
          <>
            <div className="sc-group">
              <IconButton showLabel label="Table" icon={<Table size={17} />} disabled={dis} onClick={() => c().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()} />
              <IconButton showLabel label="Image" icon={<ImagePlus size={17} />} disabled={dis} onClick={() => onAction('image')} />
              <IconButton showLabel label="Link" shortcut={`${mod}K`} icon={<Link size={17} />} active={s.link} disabled={dis} onClick={() => onAction('link')} />
              <IconButton showLabel label="Bookmark" icon={<Bookmark size={17} />} disabled={dis} onClick={() => onAction('bookmark')} />
              <IconButton showLabel label="Comment" shortcut={`${mod}Alt+M`} icon={<MessageSquarePlus size={17} />} disabled={dis} onClick={() => onAction('comment')} />
            </div>
            <div className="sc-group">
              <IconButton showLabel label="Page break" shortcut={`${mod}Enter`} icon={<SeparatorHorizontal size={17} />} disabled={dis} onClick={() => c().insertPageBreak().run()} />
              <IconButton showLabel label="Section break" icon={<Columns2 size={17} />} disabled={dis} onClick={() => onAction('sectionBreak')} />
              <IconButton showLabel label="Horizontal line" icon={<Minus size={17} />} disabled={dis} onClick={() => c().setHorizontalRule().run()} />
              <IconButton showLabel label="Page numbers" icon={<Hash size={17} />} disabled={dis} onClick={() => onAction('pageNumber')} />
            </div>
            {s.table && <TableTools editor={editor} disabled={dis} />}
          </>
        )}
        {tab === 'Layout' && (
          <>
            <div className="sc-group">
              <IconButton showLabel label="Page setup…" icon={<Ruler size={17} />} disabled={dis} onClick={() => onAction('pageSetup')} />
              <IconButton showLabel label="Rotate page" icon={<Columns2 size={17} />} disabled={dis} onClick={() => onAction('orientation')} />
            </div>
            <div className="sc-group">
              <span className="sc-group-label">Size</span>
              <button className="sc-text-btn" disabled={dis} onClick={() => onAction('size-Letter')}>Letter</button>
              <button className="sc-text-btn" disabled={dis} onClick={() => onAction('size-A4')}>A4</button>
            </div>
            <div className="sc-group">
              <span className="sc-group-label">Margins</span>
              <button className="sc-text-btn" disabled={dis} onClick={() => onAction('margins-normal')}>Normal</button>
              <button className="sc-text-btn" disabled={dis} onClick={() => onAction('margins-narrow')}>Narrow</button>
              <button className="sc-text-btn" disabled={dis} onClick={() => onAction('margins-wide')}>Wide</button>
            </div>
            <div className="sc-group">
              <button className="sc-text-btn" disabled={dis} onClick={() => onAction('header')}>Edit header</button>
              <button className="sc-text-btn" disabled={dis} onClick={() => onAction('footer')}>Edit footer</button>
            </div>
          </>
        )}
        {tab === 'Review' && (
          <>
            <div className="sc-group">
              <IconButton showLabel label={s.suggesting ? 'Suggesting' : 'Track changes'} icon={<PenLine size={17} />} active={s.suggesting} disabled={dis} onClick={() => onAction('toggleSuggest')} />
              <IconButton showLabel label="Accept all" icon={<CheckCheck size={17} />} disabled={dis} onClick={() => onAction('acceptAll')} />
              <IconButton showLabel label="Reject all" icon={<Ban size={17} />} disabled={dis} onClick={() => onAction('rejectAll')} />
            </div>
            <div className="sc-group">
              <IconButton showLabel label="New comment" icon={<MessageSquarePlus size={17} />} disabled={dis} onClick={() => onAction('comment')} />
              <button className="sc-text-btn" onClick={() => onAction('wordCount')}>Word count</button>
              {props.hasOriginal && <IconButton showLabel label="Compatibility report" icon={<FileWarning size={17} />} onClick={() => onAction('report')} />}
            </div>
          </>
        )}
        {tab === 'View' && (
          <>
            <div className="sc-group">
              <IconButton label="Zoom out" icon={<ZoomOut size={17} />} onClick={() => onAction('zoomOut')} />
              <button className="sc-text-btn" onClick={() => onAction('zoomReset')} aria-label="Reset zoom">
                {Math.round(props.zoom * 100)}%
              </button>
              <IconButton label="Zoom in" icon={<ZoomIn size={17} />} onClick={() => onAction('zoomIn')} />
            </div>
            <div className="sc-group">
              <IconButton showLabel label="Navigation pane" icon={<PanelLeft size={17} />} onClick={() => onAction('panel')} />
              <IconButton showLabel label="Formatting marks" icon={<Pilcrow size={17} />} active={props.showMarks} onClick={() => onAction('marks')} />
              <IconButton showLabel label={props.theme === 'dark' ? 'Light theme' : 'Dark theme'} icon={props.theme === 'dark' ? <Sun size={17} /> : <Moon size={17} />} onClick={() => onAction('theme')} />
            </div>
            <div className="sc-group">
              <IconButton showLabel label="Print preview" shortcut={`${mod}P`} icon={<Printer size={17} />} onClick={() => onAction('print')} />
              <IconButton showLabel label="Shortcuts" icon={<Keyboard size={17} />} onClick={() => onAction('shortcuts')} />
            </div>
          </>
        )}
      </div>
    </div>
  )
}

function TableTools({ editor, disabled }: { editor: Editor; disabled: boolean }) {
  const c = () => editor.chain().focus()
  const b = (label: string, fn: () => boolean) => (
    <button className="sc-text-btn" disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={fn}>
      {label}
    </button>
  )
  return (
    <div className="sc-group" aria-label="Table tools">
      <span className="sc-group-label">Table</span>
      {b('Row above', () => c().addRowBefore().run())}
      {b('Row below', () => c().addRowAfter().run())}
      {b('Column left', () => c().addColumnBefore().run())}
      {b('Column right', () => c().addColumnAfter().run())}
      {b('Merge', () => c().mergeCells().run())}
      {b('Split', () => c().splitCell().run())}
      {b('Header row', () => c().toggleHeaderRow().run())}
      <label className="sc-color" title="Cell shading">
        <span className="sc-group-label">Shade</span>
        <input type="color" aria-label="Cell shading" disabled={disabled} defaultValue="#e8eef8" onChange={(e) => c().setCellAttribute('background', e.target.value).run()} />
      </label>
      {b('Delete row', () => c().deleteRow().run())}
      {b('Delete column', () => c().deleteColumn().run())}
      {b('Delete table', () => c().deleteTable().run())}
    </div>
  )
}
