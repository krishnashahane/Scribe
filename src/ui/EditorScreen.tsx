import { lazy, Suspense, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor, type Editor } from '@tiptap/react'
import { generateHTML } from '@tiptap/core'
import Collaboration from '@tiptap/extension-collaboration'
import CollaborationCaret from '@tiptap/extension-collaboration-caret'
import {
  ArrowLeft, Cloud, CloudOff, FileDown, History, Loader2, MessageSquare, PanelLeft, PenLine, Share2, Sparkles, SpellCheck2, Users, Footprints,
} from 'lucide-react'
import { nextNoteId } from '../docx/notes'

// The DOCX engine (and its zip library) loads on first use: opening, exporting or merging a file.
const docxEngine = () => Promise.all([import('../docx/export'), import('../docx/import'), import('../docx/validate')]).then(([e, i, v]) => ({ exportDocx: e.exportDocx, importDocx: i.importDocx, validatePackage: v.validatePackage }))
import { Proofing } from '../features/grammar'
import { mergeDocument } from '../features/merge'
import { BIBLIOGRAPHY_TITLES, citationInstr, formatBibliography, formatCitation } from '../features/citations'
import { applyParagraphStyle, citedTags, refreshCitations, styleFromSelection } from '../features/docOps'
import { galleryStyles, headingLevelOfStyle } from '../features/styles'
import { DEFAULT_SHAPE, SHAPE_KINDS, type ShapeAttrs, type ShapeKind } from '../objects/shapes'
import { ommlToTree, toLinear } from '../objects/math'
import { newChart } from '../objects/chart'
import type { ObjectInfo } from '../model/extensions'
import type { StyleDef } from '../model/types'
// object dialogs (equations, charts, shapes, drawing…) load the first time one is opened
const objectDialogs = () => import('./ObjectDialogs')
const ChartDialog = lazy(() => objectDialogs().then((m) => ({ default: m.ChartDialog })))
const CustomXmlDialog = lazy(() => objectDialogs().then((m) => ({ default: m.CustomXmlDialog })))
const DrawDialog = lazy(() => objectDialogs().then((m) => ({ default: m.DrawDialog })))
const EquationDialog = lazy(() => objectDialogs().then((m) => ({ default: m.EquationDialog })))
const ShapeDialog = lazy(() => objectDialogs().then((m) => ({ default: m.ShapeDialog })))
const SmartArtDialog = lazy(() => objectDialogs().then((m) => ({ default: m.SmartArtDialog })))
const TextBoxDialog = lazy(() => objectDialogs().then((m) => ({ default: m.TextBoxDialog })))
import { CitationDialog, MailMergePanel, ModifyStyleDialog, NotesPanel, ProofingPanel, SourcesDialog } from './FeatureUI'
import { createDocument } from './docActions'
import { api, type User } from '../api'
import { openSession, userColor, type CollabSession, type SyncStatus } from '../collab/session'
import { schemaExtensions, newId as newNodeId } from '../model/extensions'
import { emptyMeta, migrateDocument, PAGE_SIZES, twipToPx, type DocMeta, type PMNode, type ScribeDocument } from '../model/types'
import { Pagination, repaginate, type PageInfo } from '../layout/pagination'
import { FindReplace, SectionSync, Suggesting, isSuggesting, setSuggesting } from '../editor/plugins'
import { contentToText, documentToHtml, download, safeFilename, sanitizeHtml } from '../formats/formats'
import { Ribbon, type RibbonAction } from './Ribbon'
import { AiPanel, ChangesPanel, CommentsPanel, FindBar, PagesPanel, VersionsPanel } from './Panels'
import { createFromTemplate, importFile, OPEN_ACCEPT, openDoc } from './docActions'
import { downloadOriginal, MAX_IMAGE_BYTES, uploadImage } from '../blobs'
import {
  AccountDialog, CompatReport, HeaderFooterDialog, LinkDialog, PageSetupDialog, PropertiesDialog, ShareDialog, ShortcutsDialog, TextPrompt, WordCountDialog,
} from './Dialogs'
import { Menu, mod, useToast } from './kit'
import type { CompatibilityIssue } from '../model/types'

type Panel = 'pages' | 'comments' | 'changes' | 'versions' | 'ai' | 'proofing' | 'merge' | 'notes'
type DialogState =
  | { kind: 'pageSetup'; section: number }
  | { kind: 'hf'; which: 'header' | 'footer'; section: number; first: boolean }
  | { kind: 'link' }
  | { kind: 'bookmark' }
  | { kind: 'share' }
  | { kind: 'props' }
  | { kind: 'shortcuts' }
  | { kind: 'wordCount' }
  | { kind: 'report'; issues: CompatibilityIssue[]; title: string; openId?: string }
  | { kind: 'account' }
  | { kind: 'object'; pos: number }
  | { kind: 'equation' }
  | { kind: 'newChart' }
  | { kind: 'newShape'; shape: ShapeAttrs }
  | { kind: 'draw' }
  | { kind: 'customXml' }
  | { kind: 'sources'; thenCite?: boolean }
  | { kind: 'citation' }
  | { kind: 'style'; style: StyleDef; isNew: boolean }
  | null

interface DocInfo {
  title: string
  role: 'owner' | 'editor' | 'commenter' | 'viewer'
  hasOriginal: boolean
  originalParts: number
  originalName: string | null
  meta: DocMeta
}

export function EditorScreen({ docId, user, onUser, theme, onToggleTheme }: { docId: string; user: User; onUser: (u: User) => void; theme: 'light' | 'dark'; onToggleTheme: () => void }) {
  const [info, setInfo] = useState<DocInfo | null>(null)
  const [error, setError] = useState('')
  const [session, setSession] = useState<CollabSession | null>(null)
  const [status, setStatus] = useState<SyncStatus>('connecting')
  const [ready, setReady] = useState(false)

  useEffect(() => {
    let s: CollabSession | null = null
    let cancelled = false
    api<any>(`/api/docs/${docId}`)
      .then((d) => {
        if (cancelled) return
        localStorage.setItem('scribe-last-doc', docId)
        setInfo({ title: d.title, role: d.role, hasOriginal: d.hasOriginal, originalParts: d.originalParts ?? 0, originalName: d.originalName, meta: migrateDocument({ content: d.content, meta: d.meta }).meta })
        s = openSession(docId, setStatus, d.role === 'viewer' || d.role === 'commenter')
        setSession(s)
        // render once the local (IndexedDB) or remote state has arrived, so the editor never starts empty
        Promise.race([s.whenSynced, s.local.whenSynced.then(() => (s!.ydoc.getXmlFragment('default').length ? undefined : s!.whenSynced))]).then(() => !cancelled && setReady(true))
      })
      .catch((e) => {
        if (localStorage.getItem('scribe-last-doc') === docId) localStorage.removeItem('scribe-last-doc')
        setError(e.message)
      })
    return () => {
      cancelled = true
      s?.destroy()
    }
  }, [docId])

  if (error)
    return (
      <div className="sc-splash">
        <p className="sc-error">{error}</p>
        <a href="#/documents">Back to documents</a>
      </div>
    )
  if (!info || !session || !ready)
    return (
      <div className="sc-splash" aria-busy="true">
        <Loader2 className="sc-spin" />
        <span>Opening document…</span>
      </div>
    )
  return <DocEditor docId={docId} info={info} session={session} status={status} user={user} onUser={onUser} theme={theme} onToggleTheme={onToggleTheme} />
}

function DocEditor({ docId, info, session, status, user, onUser, theme, onToggleTheme }: { docId: string; info: DocInfo; session: CollabSession; status: SyncStatus; user: User; onUser: (u: User) => void; theme: 'light' | 'dark'; onToggleTheme: () => void }) {
  const toast = useToast()
  const readOnly = info.role === 'viewer' || info.role === 'commenter'
  const ymeta = useMemo(() => session.ydoc.getMap<string>('scribe'), [session])
  const readMeta = useCallback((): DocMeta => {
    const raw = ymeta.get('meta')
    if (!raw) return info.meta
    try {
      return migrateDocument({ content: { type: 'doc' }, meta: JSON.parse(raw) }).meta
    } catch {
      return info.meta
    }
  }, [ymeta, info.meta])
  const [meta, setMetaState] = useState<DocMeta>(readMeta)
  // Values read by editor plugins and timers outside React rendering (pagination, autosave, print).
  const [live] = useState(() => new LiveState(meta, Number(localStorage.getItem('scribe-zoom')) || 1))
  useLayoutEffect(() => live.setMeta(meta), [live, meta])
  const [title, setTitle] = useState(info.title)
  const [pages, setPages] = useState<PageInfo[]>([])
  // the navigation pane starts closed; remember the user's choice
  const [panel, setPanelState] = useState<Panel | null>(() => (localStorage.getItem('scribe-panel') as Panel | null) || null)
  const setPanel = (p: Panel | null | ((cur: Panel | null) => Panel | null)) =>
    setPanelState((cur) => {
      const next = typeof p === 'function' ? p(cur) : p
      if (next) localStorage.setItem('scribe-panel', next)
      else localStorage.removeItem('scribe-panel')
      return next
    })
  const [dialog, setDialog] = useState<DialogState>(null)
  const [find, setFindOpen] = useState<null | { replace: boolean }>(null)
  const [zoom, setZoom] = useState(() => live.zoom)
  useLayoutEffect(() => live.setZoom(zoom), [live, zoom])
  const [showMarks, setShowMarks] = useState(false)
  const [painter, setPainter] = useState<null | { marks: any[]; para: Record<string, any> }>(null)
  const [composingComment, setComposingComment] = useState(false)
  const [saveState, setSaveState] = useState<'saved' | 'saving' | 'error'>('saved')

  const imageInput = useRef<HTMLInputElement>(null)
  const openInput = useRef<HTMLInputElement>(null)
  const newDocument = async (template = 'blank') => {
    await saveSnapshot()
    createFromTemplate(template).then(openDoc).catch((e) => toast(e.message, 'error'))
  }
  const openFile = async (file: File) => {
    toast(`Opening ${file.name}…`)
    try {
      await saveSnapshot()
      const r = await importFile(file)
      if (r.report.length) setDialog({ kind: 'report', issues: r.report, title: 'Import report', openId: r.id })
      else openDoc(r.id)
    } catch (e: any) {
      toast(`Could not open ${file.name}: ${e.message}`, 'error')
    }
  }

  // ---- meta sync (Yjs map, last-writer-wins per edit)
  useEffect(() => {
    const on = () => setMetaState(readMeta())
    ymeta.observe(on)
    if (!ymeta.get('meta') && !readOnly) ymeta.set('meta', JSON.stringify(info.meta))
    return () => ymeta.unobserve(on)
  }, [ymeta, readMeta, readOnly, info.meta])
  const setMeta = useCallback(
    (f: (m: DocMeta) => DocMeta) => {
      if (readOnly) return
      const next = f(live.meta)
      live.setMeta(next)
      ymeta.set('meta', JSON.stringify(next))
      setMetaState(next)
      live.markDirty()
    },
    [ymeta, readOnly, live],
  )

  const editor = useEditor(
    {
      editable: !readOnly,
      autofocus: readOnly ? false : 'end',
      extensions: [
        ...schemaExtensions({ collaborative: true }),
        Collaboration.configure({ document: session.ydoc, field: 'default' }),
        CollaborationCaret.configure({ provider: session.provider, user: { name: user.name, color: userColor(user.id) } }),
        Pagination.configure({
          getSections: () => live.meta.sections,
          getGap: () => live.gap,
          getScale: () => live.zoom,
          onLayout: setPages,
          getNoteHeight: (id) => {
            const n = live.meta.footnotes.find((f) => f.id === id)
            return n ? measureNotes([{ label: '00', text: n.text }], contentWidth(live.meta)) : 0
          },
          getEndnotesHeight: () => (live.meta.endnotes.length ? measureNotes(live.meta.endnotes.map((n) => ({ label: '00', text: n.text })), contentWidth(live.meta), true) : 0),
        }),
        FindReplace,
        Proofing,
        Suggesting.configure({ author: () => user.name }),
        SectionSync.configure({
          onSections: (count, removed) => {
            const m = live.meta
            if (m.sections.length === count && !removed.length) return
            let sections = m.sections.filter((_, i) => !removed.includes(i))
            if (sections.length > count) sections = [...sections.slice(0, count - 1), sections[sections.length - 1]]
            while (sections.length < count) sections = [...sections.slice(0, -1), { ...sections[sections.length - 1] }, sections[sections.length - 1]]
            setMeta((mm) => ({ ...mm, sections }))
          },
        }),
      ],
      editorProps: {
        attributes: { class: 'sc-doc', spellcheck: 'true', 'aria-label': 'Document body', role: 'textbox', 'aria-multiline': 'true' },
        handlePaste: (view, event) => {
          // sanitize pasted HTML; files (images) are handled below
          const files = [...(event.clipboardData?.files ?? [])].filter((f) => f.type.startsWith('image/'))
          if (files.length) {
            files.forEach(insertImageFile)
            return true
          }
          void view
          return false
        },
        transformPastedHTML: (html) => sanitizeHtml(html),
        handleDoubleClickOn: (_view, _pos, node, nodePos) => {
          if (readOnly) return false
          if (node.type.name === 'shape' || (node.type.name === 'opaqueInline' && node.attrs.object)) {
            setDialog({ kind: 'object', pos: nodePos })
            return true
          }
          return false
        },
        handleDrop: (_view, event) => {
          const files = [...(event.dataTransfer?.files ?? [])].filter((f) => f.type.startsWith('image/'))
          if (!files.length) return false
          files.forEach(insertImageFile)
          return true
        },
      },
      onUpdate: () => {
        live.markDirty()
      },
    },
    [session],
  )

  const noteLabels = useNoteLabels(editor)
  const [focusNote, setFocusNote] = useState<string | null>(null)

  // ---- autosave snapshot (content lives in Yjs; this snapshot feeds listings, exports on other devices and versions)
  const snapshot = useCallback((): ScribeDocument => ({ version: 1, content: editor!.getJSON() as PMNode, meta: live.meta }), [editor, live])
  const saveSnapshot = useCallback(
    async (keepalive = false) => {
      if (readOnly || !editor || !live.takeDirty()) return
      setSaveState('saving')
      const d = snapshot()
      try {
        if (keepalive) {
          fetch(`/api/docs/${docId}`, { method: 'PATCH', keepalive: true, headers: { 'content-type': 'application/json', 'x-scribe-request': '1' }, body: JSON.stringify({ content: d.content, meta: d.meta }) })
          return
        }
        await api(`/api/docs/${docId}`, { method: 'PATCH', body: { content: d.content, meta: d.meta } })
        if (live.versionDue()) {
          await api(`/api/docs/${docId}/versions`, { method: 'POST', body: { content: d.content, meta: d.meta, auto: true } })
        }
        setSaveState('saved')
      } catch {
        live.markDirty()
        setSaveState('error')
      }
    },
    [editor, readOnly, docId, snapshot, live],
  )
  useEffect(() => {
    const t = setInterval(() => saveSnapshot(), 4000)
    const hide = () => document.visibilityState === 'hidden' && saveSnapshot(true)
    const unload = () => saveSnapshot(true)
    document.addEventListener('visibilitychange', hide)
    window.addEventListener('pagehide', unload)
    return () => {
      clearInterval(t)
      document.removeEventListener('visibilitychange', hide)
      window.removeEventListener('pagehide', unload)
      saveSnapshot(true)
    }
  }, [saveSnapshot])

  useEffect(() => {
    document.title = `${title} — SCRIBE`
  }, [title])
  useEffect(() => {
    repaginate(editor?.view)
  }, [meta.sections, meta.footnotes, meta.endnotes, editor])
  useEffect(() => {
    localStorage.setItem('scribe-zoom', String(zoom))
    repaginate(editor?.view)
  }, [zoom, editor])

  const renameTimer = useRef<ReturnType<typeof setTimeout>>(undefined)
  const rename = (t: string) => {
    setTitle(t)
    clearTimeout(renameTimer.current)
    renameTimer.current = setTimeout(() => api(`/api/docs/${docId}`, { method: 'PATCH', body: { title: t } }).catch((e) => toast(e.message, 'error')), 600)
  }

  // ---- images
  const insertImageFile = useCallback(
    (file: File) => {
      if (!/^image\/(png|jpe?g|gif|webp|bmp)$/.test(file.type)) return toast('Unsupported image type. Use PNG, JPEG, GIF, WebP or BMP.', 'error')
      if (file.size > 15 * 1024 * 1024) return toast('Images must be smaller than 15 MB.', 'error')
      if (file.size > MAX_IMAGE_BYTES) return toast('Images must be smaller than 4 MB.', 'error')
      const local = URL.createObjectURL(file)
      const img = new Image()
      img.onload = async () => {
        const p = live.meta.sections[0].page
        const maxW = twipToPx(p.width - p.marginLeft - p.marginRight)
        const w = Math.min(img.naturalWidth, maxW)
        const h = Math.round((img.naturalHeight * w) / img.naturalWidth)
        URL.revokeObjectURL(local)
        try {
          const src = await uploadImage(new Uint8Array(await file.arrayBuffer()), file.type)
          editor?.chain().focus().insertContent({ type: 'image', attrs: { src, width: Math.round(w), height: h, alt: file.name.replace(/\.[^.]+$/, '') } }).run()
        } catch (e: any) {
          toast(e.message, 'error')
        }
      }
      img.onerror = () => {
        URL.revokeObjectURL(local)
        toast('That image could not be read.', 'error')
      }
      img.src = local
    },
    [editor, toast, live],
  )

  // ---- exports
  const fetchOriginal = async () => (info.hasOriginal ? downloadOriginal(docId, info.originalParts) : null)
  const exportAs = async (fmt: 'docx' | 'html' | 'txt' | 'pdf') => {
    if (!editor) return
    try {
      if (fmt === 'pdf') return printDoc()
      const d = snapshot()
      d.meta = { ...d.meta, metadata: { ...d.meta.metadata, title, lastModifiedBy: user.name } }
      if (fmt === 'docx') {
        toast('Preparing DOCX…')
        const { exportDocx, validatePackage } = await docxEngine()
        const { data, report } = await exportDocx(d, await fetchOriginal())
        const problems = await validatePackage(data)
        if (problems.length) {
          toast(`Export stopped: the DOCX failed ${problems.length} structural check(s), so no file was downloaded. First problem: ${problems[0]}`, 'error')
          return
        }
        download(data as BlobPart, safeFilename(title, 'docx'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        if (report.length) setDialog({ kind: 'report', issues: report, title: 'Export report' })
        else toast('DOCX exported as a new file. The original upload is unchanged.', 'success')
      } else if (fmt === 'html') {
        download(documentToHtml(editor.getHTML(), d, title), safeFilename(title, 'html'), 'text/html')
      } else {
        download(contentToText(d.content), safeFilename(title, 'txt'), 'text/plain;charset=utf-8')
      }
    } catch (e: any) {
      toast(`Export failed: ${e.message}`, 'error')
    }
  }

  const printDoc = () => {
    const p = live.meta.sections[0].page
    const style = document.createElement('style')
    style.id = 'sc-print-page'
    style.textContent = `@page { size: ${(p.width / 1440).toFixed(3)}in ${(p.height / 1440).toFixed(3)}in; margin: 0 }`
    document.head.appendChild(style)
    document.documentElement.classList.add('is-printing')
    live.beginPrint()
    repaginate(editor?.view, true)
    const done = () => {
      document.documentElement.classList.remove('is-printing')
      style.remove()
      live.endPrint()
      repaginate(editor?.view, true)
      window.removeEventListener('afterprint', done)
    }
    window.addEventListener('afterprint', done)
    // let React render the gapless page frames before the print snapshot
    requestAnimationFrame(() => requestAnimationFrame(() => window.print()))
  }

  const saveVersion = useCallback(
    async (label?: string) => {
      if (!editor || readOnly) return
      const d = snapshot()
      await api(`/api/docs/${docId}/versions`, { method: 'POST', body: { content: d.content, meta: d.meta, label: label || `Saved by ${user.name}` } })
      toast('Version saved', 'success')
    },
    [editor, readOnly, snapshot, docId, user.name, toast],
  )

  const restore = async (content: PMNode, m: DocMeta) => {
    await saveVersion('Before restore').catch(() => {})
    editor?.commands.setContent(content as any)
    setMeta(() => migrateDocument({ content, meta: m }).meta)
    toast('Version restored. The previous state was saved as a version.', 'success')
  }

  const currentSection = () => {
    if (!editor) return 0
    let sec = 0
    const pos = editor.state.selection.from
    editor.state.doc.descendants((n, p) => {
      if (n.type.name === 'sectionBreak' && p < pos) sec = n.attrs.section + 1
      return n.isBlock && n.type.name !== 'table'
    })
    return Math.min(sec, live.meta.sections.length - 1)
  }

  const updateSections = (idx: number | 'all', f: (s: DocMeta['sections'][number]) => DocMeta['sections'][number]) =>
    setMeta((m) => ({ ...m, sections: m.sections.map((s, i) => (idx === 'all' || i === idx ? f(s) : s)) }))

  const bookmarks = () => {
    const out: string[] = []
    editor?.state.doc.descendants((n) => {
      if (n.type.name === 'bookmark' && n.attrs.name) out.push(n.attrs.name)
    })
    return out
  }

  const openHF = (which: 'header' | 'footer', section = currentSection(), page?: PageInfo) => {
    const sec = live.meta.sections[section]
    const first = !!(sec?.titlePage && page?.firstOfSection)
    setDialog({ kind: 'hf', which, section, first })
  }

  const onAction = (a: RibbonAction) => {
    if (!editor) return
    const sec = currentSection()
    switch (a) {
      case 'image':
        return imageInput.current?.click()
      case 'shape-rect':
      case 'shape-roundRect':
      case 'shape-ellipse':
      case 'shape-triangle':
      case 'shape-diamond':
      case 'shape-rightArrow':
      case 'shape-star5':
      case 'shape-line': {
        const kind = a.slice(6) as ShapeKind
        if (!(kind in SHAPE_KINDS)) return
        editor.chain().focus().insertContent({ type: 'shape', attrs: { ...DEFAULT_SHAPE, kind, ...(kind === 'line' ? { height: 4, stroke: '#1f2937', strokeWidth: 2 } : {}) } }).run()
        return toast('Shape inserted. Double-click it to change its colour, size or text.', 'info')
      }
      case 'link':
        return setDialog({ kind: 'link' })
      case 'bookmark':
        return setDialog({ kind: 'bookmark' })
      case 'comment':
        if (editor.state.selection.empty) return toast('Select the text you want to comment on.', 'info')
        setPanel('comments')
        return setComposingComment(true)
      case 'pageSetup':
        return setDialog({ kind: 'pageSetup', section: currentSection() })
      case 'header':
      case 'footer':
        return openHF(a, sec)
      case 'pageNumber': {
        const key = live.meta.sections[sec].footerDefault
        if (!key) {
          const k = `new:footer-${newNodeId()}`
          setMeta((m) => ({
            ...m,
            headerFooters: { ...m.headerFooters, [k]: { type: 'doc', content: [{ type: 'paragraph', attrs: { textAlign: 'center' }, content: [{ type: 'field', attrs: { instr: 'PAGE', result: '1' } }] }] } },
            sections: m.sections.map((s, i) => (i === sec ? { ...s, footerDefault: k } : s)),
          }))
          return toast('Page numbers added to the footer.', 'success')
        }
        return openHF('footer', sec)
      }
      case 'sectionBreak': {
        setMeta((m) => {
          const sections = [...m.sections]
          sections.splice(sec, 0, { ...sections[sec], rawXml: undefined })
          return { ...m, sections }
        })
        editor.chain().focus().insertContent([{ type: 'sectionBreak', attrs: { section: sec, breakType: 'nextPage' } }, { type: 'paragraph' }]).run()
        return
      }
      case 'orientation':
        return updateSections(sec, (s) => ({ ...s, page: { ...s.page, orientation: s.page.orientation === 'landscape' ? 'portrait' : 'landscape', width: s.page.height, height: s.page.width } }))
      case 'size-A4':
      case 'size-Letter': {
        const size = PAGE_SIZES[a.slice(5)]
        return updateSections(sec, (s) => ({ ...s, page: { ...s.page, ...(s.page.orientation === 'landscape' ? { width: size.height, height: size.width } : { width: size.width, height: size.height }) } }))
      }
      case 'margins-normal':
      case 'margins-narrow':
      case 'margins-wide': {
        const v = a === 'margins-normal' ? { t: 1440, s: 1440 } : a === 'margins-narrow' ? { t: 720, s: 720 } : { t: 1440, s: 2880 }
        return updateSections(sec, (s) => ({ ...s, page: { ...s.page, marginTop: v.t, marginBottom: v.t, marginLeft: v.s, marginRight: v.s } }))
      }
      case 'find':
        return setFindOpen({ replace: false })
      case 'wordCount':
        return setDialog({ kind: 'wordCount' })
      case 'print':
        return printDoc()
      case 'shortcuts':
        return setDialog({ kind: 'shortcuts' })
      case 'report':
        fetchOriginal()
          .then(async (b) => b && (await docxEngine()).importDocx(b))
          .then((r) => r && setDialog({ kind: 'report', issues: r.report, title: 'Compatibility report' }))
          .catch((e) => toast(e.message, 'error'))
        return
      case 'toggleSuggest': {
        const on = !isSuggesting(editor.state)
        setSuggesting(editor, on, user.name)
        if (on) setPanel('changes')
        return toast(on ? 'Track changes on — your edits are recorded as suggestions.' : 'Track changes off.', 'info')
      }
      case 'acceptAll':
        return (editor.commands as any).acceptAllChanges()
      case 'rejectAll':
        return (editor.commands as any).rejectAllChanges()
      case 'zoomIn':
        return setZoom((z) => Math.min(3, Math.round((z + 0.1) * 10) / 10))
      case 'zoomOut':
        return setZoom((z) => Math.max(0.3, Math.round((z - 0.1) * 10) / 10))
      case 'zoomReset':
        return setZoom(1)
      case 'theme':
        return onToggleTheme()
      case 'panel':
        return setPanel((p) => (p ? null : 'pages'))
      case 'marks':
        return setShowMarks((v) => !v)
      case 'equation':
        return setDialog({ kind: 'equation' })
      case 'columns-1':
      case 'columns-2':
      case 'columns-3': {
        const n = Number(a.slice(-1))
        updateSections(sec, (x) => ({ ...x, columns: n, columnsEdited: true }))
        return toast(n === 1 ? 'This section is set to one column.' : `This section is set to ${n} columns. Exported and printed files show the columns; SCRIBE shows the text in one column while you edit.`, 'info')
      }
      case 'chart':
        return setDialog({ kind: 'newChart' })
      case 'footnote':
      case 'endnote': {
        const kind = a
        const list = kind === 'footnote' ? live.meta.footnotes : live.meta.endnotes
        const id = nextNoteId(list)
        setMeta((m) => (kind === 'footnote' ? { ...m, footnotes: [...m.footnotes, { id, text: '', edited: true }] } : { ...m, endnotes: [...m.endnotes, { id, text: '', edited: true }] }))
        // a selected object (picture, chart, shape) is never replaced: the reference goes right after it.
        // No editor focus here: the note's text box in the panel takes the cursor.
        const sel = editor.state.selection
        const at = (sel as any).node ? sel.to : sel.from
        editor.commands.insertContentAt(sel.empty || (sel as any).node ? at : { from: sel.from, to: sel.to }, { type: 'noteRef', attrs: { noteId: id, kind, label: '' } })
        setPanel('notes')
        setFocusNote(`${kind}:${id}`)
        return
      }
      case 'textbox':
        return setDialog({ kind: 'newShape', shape: { ...DEFAULT_SHAPE, kind: 'rect', fill: '#ffffff', stroke: '#1f2937', strokeWidth: 1, width: 220, height: 90, text: 'Text box' } })
      case 'draw':
        return setDialog({ kind: 'draw' })
      case 'editObject': {
        const sel = editor.state.selection as any
        if (sel.node) setDialog({ kind: 'object', pos: sel.from })
        return
      }
      case 'citation':
        return setDialog({ kind: 'citation' })
      case 'sources':
        return setDialog({ kind: 'sources' })
      case 'bibliography': {
        const m = live.meta
        const style = m.citationStyle ?? 'APA'
        if (!(m.sources ?? []).length) toast('The bibliography lists your sources. Add them from References → Sources.', 'info')
        editor.chain().focus().insertContent([{ type: 'bibliography', attrs: { title: BIBLIOGRAPHY_TITLES[style], entries: formatBibliography(m.sources ?? [], style) } }, { type: 'paragraph' }]).run()
        return
      }
      case 'updateCitations':
        return toast(refreshCitations(editor, live.meta) ? 'Citations and bibliography updated.' : 'Citations and bibliography are up to date.', 'success')
      case 'citationStyle-APA':
      case 'citationStyle-MLA':
      case 'citationStyle-Chicago': {
        const style = a.slice('citationStyle-'.length) as 'APA' | 'MLA' | 'Chicago'
        setMeta((m) => ({ ...m, citationStyle: style }))
        refreshCitations(editor, { ...live.meta, citationStyle: style }, { retitle: true })
        return
      }
      case 'mailMerge':
        return setPanel('merge')
      case 'proofing':
        return setPanel('proofing')
      case 'customXml':
        return setDialog({ kind: 'customXml' })
      case 'newStyle':
        return setDialog({ kind: 'style', isNew: true, style: { id: '', name: 'New style', type: 'paragraph', basedOn: 'Normal', ...styleFromSelection(editor) } })
      case 'formatPainter': {
        if (painter) return setPainter(null)
        const { $from } = editor.state.selection
        const skip = ['link', 'comment', 'insertion', 'deletion']
        const marks = (editor.state.storedMarks ?? $from.marks()).filter((m) => !skip.includes(m.type.name)).map((m) => m.toJSON())
        const a = $from.parent.attrs
        const para = Object.fromEntries(['textAlign', 'indentLeft', 'indentRight', 'indentFirstLine', 'spacingBefore', 'spacingAfter', 'lineSpacing', 'lineRule', 'shading', 'styleId'].map((k) => [k, a[k] ?? null]))
        setPainter({ marks, para })
        return toast('Format painter: select the text to apply this formatting to (Esc to cancel).', 'info')
      }
    }
  }

  // ---- objects (charts, diagrams, equations, text boxes, shapes)
  /** Update the node a dialog was opened for; finds it again if collaborators moved it. */
  const updateNodeAt = (pos: number, original: any, attrs: Record<string, any>) => {
    if (!editor) return
    let at = pos
    const cur = editor.state.doc.nodeAt(pos)
    if (!cur || cur.type.name !== original.type.name || cur.attrs.xml !== original.attrs.xml) {
      at = -1
      editor.state.doc.descendants((n, p) => {
        if (at === -1 && n.type.name === original.type.name && n.attrs.xml === original.attrs.xml && JSON.stringify(n.attrs) === JSON.stringify(original.attrs)) at = p
        return at === -1
      })
      if (at === -1) return toast('That object was changed or removed by someone else.', 'error')
    }
    editor.view.dispatch(editor.state.tr.setNodeMarkup(at, undefined, { ...editor.state.doc.nodeAt(at)!.attrs, ...attrs }))
    setDialog(null)
    editor.commands.focus()
  }
  const insertEquation = (xml: string) => {
    let preview: string
    try {
      preview = toLinear(ommlToTree(xml))
    } catch {
      preview = 'Equation'
    }
    editor?.chain().focus().insertContent({ type: 'opaqueInline', attrs: { xml, label: 'Equation', preview, object: { kind: 'equation' } } }).run()
    setDialog(null)
  }
  const insertDrawing = async (png: Blob, w: number, h: number) => {
    try {
      if (png.size > MAX_IMAGE_BYTES) return toast('The drawing is too large. Draw a smaller area.', 'error')
      const p = live.meta.sections[0].page
      const maxW = twipToPx(p.width - p.marginLeft - p.marginRight)
      const scale = Math.min(1, maxW / w)
      const src = await uploadImage(new Uint8Array(await png.arrayBuffer()), 'image/png')
      editor?.chain().focus().insertContent({ type: 'image', attrs: { src, width: Math.round(w * scale), height: Math.round(h * scale), alt: 'Drawing' } }).run()
      setDialog(null)
    } catch (e: any) {
      toast(e.message, 'error')
    }
  }

  // ---- mail merge output
  const mergedDoc = () => {
    const data = live.meta.mailMerge
    if (!data || !editor) throw new Error('Load recipients first.')
    return mergeDocument(editor.getJSON() as PMNode, live.meta, data)
  }
  const mergeToNewDocument = async () => {
    try {
      toast('Merging…')
      const m = mergedDoc()
      await saveSnapshot()
      const original = await fetchOriginal()
      const id = await createDocument(`${title} (merged)`, m.content, m.meta, original ?? undefined, original ? `${title} (merged).docx` : undefined)
      toast(`Merged ${live.meta.mailMerge!.records.length.toLocaleString()} recipients into a new document.`, 'success')
      openDoc(id)
    } catch (e: any) {
      toast(`Merge failed: ${e.message}`, 'error')
    }
  }
  const mergeToDocx = async () => {
    try {
      toast('Preparing merged DOCX…')
      const m = mergedDoc()
      const { exportDocx, validatePackage } = await docxEngine()
      const { data } = await exportDocx({ version: 1, content: m.content, meta: { ...m.meta, metadata: { ...m.meta.metadata, title: `${title} (merged)`, lastModifiedBy: user.name } } }, await fetchOriginal())
      const problems = await validatePackage(data)
      if (problems.length) return toast(`Merge stopped: the DOCX failed a structural check (${problems[0]}).`, 'error')
      download(data as BlobPart, safeFilename(`${title} (merged)`, 'docx'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
      toast('Merged DOCX downloaded.', 'success')
    } catch (e: any) {
      toast(`Merge failed: ${e.message}`, 'error')
    }
  }

  // ---- styles
  const saveStyle = (st: StyleDef) => {
    setMeta((m) => {
      const exists = m.styles.some((x) => x.id === st.id)
      return { ...m, styles: exists ? m.styles.map((x) => (x.id === st.id ? { ...x, ...st } : x)) : [...m.styles, st] }
    })
    setDialog(null)
  }

  // ---- format painter: apply copied formatting to the next selection
  useEffect(() => {
    if (!editor || !painter) return
    const dom = editor.view.dom as HTMLElement
    dom.classList.add('is-painting')
    const apply = () =>
      setTimeout(() => {
        const { from, to, empty } = editor.state.selection
        if (empty) return
        const { schema } = editor.state
        const tr = editor.state.tr
        for (const name of ['bold', 'italic', 'underline', 'strike', 'textStyle', 'highlight', 'subscript', 'superscript']) if (schema.marks[name]) tr.removeMark(from, to, schema.marks[name])
        for (const m of painter.marks) tr.addMark(from, to, schema.markFromJSON(m))
        editor.state.doc.nodesBetween(from, to, (n, pos) => {
          if (n.type.name === 'paragraph' || n.type.name === 'heading') tr.setNodeMarkup(pos, undefined, { ...n.attrs, ...painter.para })
        })
        editor.view.dispatch(tr)
        setPainter(null)
      }, 0)
    const esc = (e: KeyboardEvent) => e.key === 'Escape' && setPainter(null)
    dom.addEventListener('mouseup', apply)
    dom.addEventListener('keyup', apply)
    window.addEventListener('keydown', esc)
    return () => {
      dom.classList.remove('is-painting')
      dom.removeEventListener('mouseup', apply)
      dom.removeEventListener('keyup', apply)
      window.removeEventListener('keydown', esc)
    }
  }, [editor, painter])

  // ---- global shortcuts
  useEffect(() => {
    const on = (e: KeyboardEvent) => {
      const m = e.metaKey || e.ctrlKey
      if (!m) return
      const k = e.key.toLowerCase()
      if (k === 'f' && !e.shiftKey) {
        e.preventDefault()
        setFindOpen({ replace: false })
      } else if (k === 'h' && !e.shiftKey && !e.altKey) {
        e.preventDefault()
        setFindOpen({ replace: true })
      } else if (k === 'k' && !readOnly) {
        e.preventDefault()
        setDialog({ kind: 'link' })
      } else if (k === 's') {
        e.preventDefault()
        saveVersion().catch((err) => toast(err.message, 'error'))
      } else if (e.shiftKey && (e.key === '>' || e.key === '<' || e.code === 'Period' || e.code === 'Comma') && editor && !readOnly) {
        e.preventDefault()
        const SZ = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]
        const cur = parseFloat(editor.getAttributes('textStyle').fontSize) || live.meta.defaults.fontSizePt
        const up = e.key === '>' || e.code === 'Period'
        const next = up ? SZ.find((x) => x > cur) ?? cur + 12 : [...SZ].reverse().find((x) => x < cur) ?? Math.max(1, cur - 1)
        editor.chain().focus().setFontSize(`${next}pt`).run()
      } else if (k === 'o' && !e.shiftKey) {
        e.preventDefault()
        openInput.current?.click()
      } else if (e.altKey && (k === 'n' || e.code === 'KeyN')) {
        e.preventDefault()
        newDocument()
      } else if (k === 'p') {
        e.preventDefault()
        printDoc()
      } else if (k === '/') {
        e.preventDefault()
        setDialog({ kind: 'shortcuts' })
      } else if (e.altKey && !readOnly && (e.code === 'KeyF' || e.code === 'KeyD')) {
        e.preventDefault()
        onAction(e.code === 'KeyF' ? 'footnote' : 'endnote')
      } else if (e.altKey && (k === 'm' || e.code === 'KeyM')) {
        e.preventDefault()
        onAction('comment')
      }
      // heading shortcuts (Mod-Alt-0…6) come from Tiptap's paragraph/heading extensions
    }
    window.addEventListener('keydown', on)
    return () => window.removeEventListener('keydown', on)
  })

  if (!editor) return null

  const styleCss = buildStyleCss(meta)
  const p0 = meta.sections[0].page
  const lastPage = pages[pages.length - 1]
  const totalHeight = lastPage ? lastPage.top + lastPage.height + twipToPx(p0.marginTop) : twipToPx(p0.height)
  const maxWidth = Math.max(twipToPx(p0.width), ...pages.map((p) => p.width))

  const onCanvasDoubleClick = (e: React.MouseEvent) => {
    if (readOnly) return
    const wrap = e.currentTarget.getBoundingClientRect()
    const y = (e.clientY - wrap.top) / zoom - twipToPx(p0.marginTop)
    const page = pages.find((p) => y >= p.top && y <= p.top + p.height)
    if (!page) return
    if (y < page.top + page.marginTop) openHF('header', page.section, page)
    else if (y > page.top + page.height - page.marginBottom) openHF('footer', page.section, page)
  }

  const statusLabel =
    status === 'offline' ? 'Offline — changes are kept on this device' : status === 'connecting' ? 'Connecting…' : saveState === 'error' ? 'Saved live; snapshot retrying' : saveState === 'saving' ? 'Saving…' : 'All changes saved'

  return (
    <div className={`sc-app ${showMarks ? 'show-marks' : ''}`}>
      <style>{styleCss}</style>
      <header className="sc-topbar">
        <a className="sc-btn" href="#/documents" aria-label="All documents" title="All documents">
          <ArrowLeft size={18} />
        </a>
        <span className="sc-logo is-small">SCRIBE</span>
        <input className="sc-title" value={title} onChange={(e) => rename(e.target.value)} readOnly={readOnly} aria-label="Document title" maxLength={200} />
        <span className={`sc-status is-${status}`} role="status" title={statusLabel}>
          {status === 'offline' ? <CloudOff size={15} /> : status === 'connecting' ? <Loader2 size={15} className="sc-spin" /> : <Cloud size={15} />}
          <span>{statusLabel}</span>
        </span>
        {readOnly && <span className="sc-chip-ui">View only</span>}
        <span className="sc-spacer" />
        <Menu
          ariaLabel="File menu"
          label={<span className="sc-menu-label">File</span>}
          items={[
            { label: 'New blank document', onClick: () => newDocument(), shortcut: `${mod}Alt+N` },
            { label: 'Open…', onClick: () => openInput.current?.click(), shortcut: `${mod}O` },
            { label: 'New from template…', onClick: () => (location.hash = '#/documents') },
            'sep',
            { label: 'Save version', onClick: () => saveVersion().catch((e) => toast(e.message, 'error')), shortcut: `${mod}S`, disabled: readOnly },
            { label: 'Version history', onClick: () => setPanel('versions') },
            'sep',
            { label: 'Download as DOCX (.docx)', onClick: () => exportAs('docx') },
            { label: 'Download as PDF (print)', onClick: () => exportAs('pdf'), shortcut: `${mod}P` },
            { label: 'Download as web page (.html)', onClick: () => exportAs('html') },
            { label: 'Download as plain text (.txt)', onClick: () => exportAs('txt') },
            'sep',
            { label: 'Page setup…', onClick: () => setDialog({ kind: 'pageSetup', section: currentSection() }), disabled: readOnly },
            { label: 'Document properties…', onClick: () => setDialog({ kind: 'props' }) },
            { label: 'Custom XML data…', onClick: () => setDialog({ kind: 'customXml' }) },
            { label: 'Share…', onClick: () => setDialog({ kind: 'share' }) },
            'sep',
            { label: user.guest ? 'Save to an account…' : 'Account', onClick: () => setDialog({ kind: 'account' }) },
            { label: 'Keyboard shortcuts', onClick: () => setDialog({ kind: 'shortcuts' }), shortcut: `${mod}/` },
            { label: 'All documents', onClick: () => (location.hash = '#/documents') },
          ]}
        />
        <button className="sc-btn with-label" onClick={() => exportAs('docx')} title="Download as DOCX" aria-label="Download as DOCX">
          <FileDown size={16} />
          <span>DOCX</span>
        </button>
        <button className="sc-primary with-icon" onClick={() => setDialog({ kind: 'share' })}>
          <Share2 size={16} /> Share
        </button>
      </header>
      <Ribbon
        editor={editor}
        readOnly={readOnly}
        styles={meta.styles}
        defaults={meta.defaults}
        zoom={zoom}
        theme={theme}
        hasOriginal={info.hasOriginal}
        showMarks={showMarks}
        painterActive={!!painter}
        onAction={onAction}
        meta={meta}
        onApplyStyle={(st) => applyParagraphStyle(editor, st, live.meta, setMeta)}
        onModifyStyle={(st) => setDialog({ kind: 'style', style: st, isNew: false })}
      />
      {find && <FindBar editor={editor} replace={find.replace} readOnly={readOnly} onClose={() => { setFindOpen(null); editor.commands.focus() }} />}
      <div className="sc-workspace">
        <nav className="sc-rail" aria-label="Side panels">
          {(
            [
              ['pages', <PanelLeft size={18} key="i" />, 'Pages & outline'],
              ['comments', <MessageSquare size={18} key="i" />, 'Comments'],
              ['changes', <PenLine size={18} key="i" />, 'Tracked changes'],
              ['versions', <History size={18} key="i" />, 'Version history'],
              ['notes', <Footprints size={18} key="i" />, 'Footnotes and endnotes'],
              ['proofing', <SpellCheck2 size={18} key="i" />, 'Proofing'],
              ['merge', <Users size={18} key="i" />, 'Mail merge'],
              ['ai', <Sparkles size={18} key="i" />, 'AI assistant'],
            ] as const
          ).map(([k, icon, label]) => (
            <button key={k} className={`sc-btn ${panel === k ? 'is-active' : ''}`} aria-label={label} title={label} aria-pressed={panel === k} onMouseDown={(e) => e.preventDefault()} onClick={() => setPanel(panel === k ? null : k)}>
              {icon}
            </button>
          ))}
        </nav>
        {panel && (
          <aside className="sc-panel" aria-label={panel}>
            {panel === 'pages' && (
              <PagesPanel
                editor={editor}
                pages={pages}
                onGoto={(p) => {
                  const canvas = document.querySelector('.sc-canvas') as HTMLElement
                  canvas?.scrollTo({ top: (p.top + twipToPx(p0.marginTop)) * zoom, behavior: 'smooth' })
                }}
              />
            )}
            {panel === 'comments' && <CommentsPanel editor={editor} meta={meta} setMeta={setMeta} user={user} readOnly={readOnly} composing={composingComment} onComposed={() => setComposingComment(false)} />}
            {panel === 'changes' && <ChangesPanel editor={editor} readOnly={readOnly} />}
            {panel === 'versions' && <VersionsPanel docId={docId} readOnly={readOnly} onSaveVersion={() => saveVersion()} onRestore={restore} />}
            {panel === 'ai' && <AiPanel editor={editor} readOnly={readOnly} />}
            {panel === 'notes' && <NotesPanel editor={editor} meta={meta} setMeta={setMeta} labels={noteLabels} readOnly={readOnly} focus={focusNote} onFocused={() => setFocusNote(null)} />}
            {panel === 'proofing' && <ProofingPanel editor={editor} readOnly={readOnly} />}
            {panel === 'merge' && <MailMergePanel editor={editor} meta={meta} setMeta={setMeta} readOnly={readOnly} onMergeToDoc={mergeToNewDocument} onMergeToDocx={mergeToDocx} />}
          </aside>
        )}
        <main className="sc-canvas" aria-label="Document">
          <div className="sc-zoom" style={{ zoom } as any}>
            <div className="sc-pages-wrap" style={{ width: maxWidth, minHeight: totalHeight }} onDoubleClick={onCanvasDoubleClick}>
              <PageFrames pages={pages} meta={meta} offsetTop={twipToPx(p0.marginTop)} labels={noteLabels} />
              <div
                className="sc-editor-host"
                style={{ width: twipToPx(p0.width), paddingTop: twipToPx(p0.marginTop), paddingLeft: twipToPx(p0.marginLeft), paddingRight: twipToPx(p0.marginRight), minHeight: totalHeight }}
                onMouseDown={(e) => {
                  if (e.target === e.currentTarget) {
                    e.preventDefault()
                    editor.commands.focus('end')
                  }
                }}
              >
                <EditorContent editor={editor} />
              </div>
            </div>
          </div>
        </main>
      </div>
      <footer className="sc-statusbar">
        <span>
          Page {pages.length ? Math.min(pages.length, currentPageIndex(editor, pages, zoom) + 1) : 1} of {Math.max(1, pages.length)}
        </span>
        <WordCountLive editor={editor} />
        <span className="sc-spacer" />
        <span>{Object.keys(meta.headerFooters).length ? 'Headers/footers on' : ''}</span>
        {(meta.sections[currentSection()]?.columns ?? 1) > 1 && (
          <span className="sc-chip-ui" title="Columns appear in exported and printed files; editing shows a single column">
            {meta.sections[currentSection()]!.columns} columns in the file
          </span>
        )}
        <label className="sc-zoom-slider">
          <span className="sc-visually-hidden">Zoom</span>
          <button className="sc-zoom-step" onClick={() => setZoom((z) => Math.max(0.3, Math.round((z - 0.1) * 10) / 10))} aria-label="Zoom out">−</button>
          <input type="range" min={30} max={300} step={10} value={Math.round(zoom * 100)} onChange={(e) => setZoom(Number(e.target.value) / 100)} aria-label="Zoom level" />
          <button className="sc-zoom-step" onClick={() => setZoom((z) => Math.min(3, Math.round((z + 0.1) * 10) / 10))} aria-label="Zoom in">+</button>
          <button className="sc-zoom-pct" onClick={() => setZoom(1)} title="Reset to 100%">{Math.round(zoom * 100)}%</button>
        </label>
      </footer>
      <input ref={openInput} type="file" accept={OPEN_ACCEPT} hidden onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ''; if (f) openFile(f) }} />
      <input ref={imageInput} type="file" accept="image/png,image/jpeg,image/gif,image/webp,image/bmp" hidden onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ''; if (f) insertImageFile(f) }} />

      {dialog?.kind === 'pageSetup' && (
        <PageSetupDialog
          sections={meta.sections}
          initialSection={dialog.section}
          onClose={() => setDialog(null)}
          onApply={(idx, page, o) => {
            updateSections(idx, (s) => ({ ...s, page, titlePage: o.titlePage }))
            setDialog(null)
          }}
        />
      )}
      {dialog?.kind === 'hf' && (() => {
        const sec = meta.sections[dialog.section]
        const prop = `${dialog.which}${dialog.first ? 'First' : 'Default'}` as 'headerDefault' | 'headerFirst' | 'footerDefault' | 'footerFirst'
        const key = sec[prop]
        return (
          <HeaderFooterDialog
            kind={dialog.which}
            first={dialog.first}
            content={key ? meta.headerFooters[key] ?? null : null}
            onClose={() => setDialog(null)}
            onSave={(content) => {
              setMeta((m) => {
                const hf = { ...m.headerFooters }
                const sections = m.sections.map((s) => ({ ...s }))
                if (!content) {
                  sections[dialog.section][prop] = undefined
                } else {
                  const k = key ?? `new:${dialog.which}-${newNodeId()}`
                  hf[k] = content
                  sections[dialog.section][prop] = k
                }
                return { ...m, headerFooters: hf, sections }
              })
              setDialog(null)
            }}
          />
        )
      })()}
      {dialog?.kind === 'link' && (
        <LinkDialog
          initial={editor.getAttributes('link').href ?? ''}
          bookmarks={bookmarks()}
          onClose={() => setDialog(null)}
          onRemove={() => {
            editor.chain().focus().extendMarkRange('link').unsetLink().run()
            setDialog(null)
          }}
          onApply={(href) => {
            const target = href.startsWith('#') ? null : '_blank'
            if (editor.state.selection.empty && !editor.isActive('link')) editor.chain().focus().insertContent({ type: 'text', text: href.replace(/^mailto:/, ''), marks: [{ type: 'link', attrs: { href, target } }] }).run()
            else editor.chain().focus().extendMarkRange('link').setLink({ href, target }).run()
            setDialog(null)
          }}
        />
      )}
      {dialog?.kind === 'bookmark' && (
        <TextPrompt
          title="Add bookmark"
          label="Bookmark name"
          onClose={() => setDialog(null)}
          validate={(v) => (!/^[A-Za-z_][A-Za-z0-9_]{0,39}$/.test(v) ? 'Use letters, digits and underscores; start with a letter.' : bookmarks().includes(v) ? 'A bookmark with this name exists.' : null)}
          onSubmit={(name) => {
            editor.chain().focus().insertContent({ type: 'bookmark', attrs: { name } }).run()
            setDialog(null)
            toast(`Bookmark “${name}” added. Link to it with Insert → Link.`, 'success')
          }}
        />
      )}
      {dialog?.kind === 'share' && <ShareDialog docId={docId} isOwner={info.role === 'owner'} onClose={() => setDialog(null)} />}
      {dialog?.kind === 'props' && (
        <PropertiesDialog
          title={title}
          meta={meta}
          readOnly={readOnly}
          onClose={() => setDialog(null)}
          onSave={(t, md) => {
            rename(t)
            setMeta((m) => ({ ...m, metadata: { ...m.metadata, ...md, title: t } }))
            setDialog(null)
          }}
        />
      )}
      {dialog?.kind === 'shortcuts' && <ShortcutsDialog onClose={() => setDialog(null)} />}
      {dialog?.kind === 'wordCount' && <WordCountDialog text={editor.state.selection.empty ? editor.getText() : editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, '\n')} onClose={() => setDialog(null)} />}
      {dialog?.kind === 'report' && (
        <CompatReport
          issues={dialog.issues}
          title={dialog.title}
          actionLabel={dialog.openId ? 'Open document' : 'Close'}
          onClose={() => {
            const id = dialog.openId
            setDialog(null)
            if (id) openDoc(id)
          }}
        />
      )}
      {dialog?.kind === 'account' && <AccountDialog user={user} onUser={onUser} onClose={() => setDialog(null)} />}
      <Suspense fallback={null}>
      {dialog?.kind === 'object' && (() => {
        const node = editor.state.doc.nodeAt(dialog.pos)
        const close = () => setDialog(null)
        if (!node) return null
        if (node.type.name === 'shape') return <ShapeDialog initial={node.attrs as ShapeAttrs} onClose={close} onSave={(a) => updateNodeAt(dialog.pos, node, a)} />
        const obj = node.attrs.object as ObjectInfo | null
        if (!obj) return null
        if (obj.kind === 'equation') return <EquationDialog xml={node.attrs.xml} onClose={close} onSave={(xml) => updateNodeAt(dialog.pos, node, { xml, preview: (() => { try { return toLinear(ommlToTree(xml)) } catch { return node.attrs.preview } })() })} />
        if (obj.kind === 'textbox') return <TextBoxDialog xml={node.attrs.xml} onClose={close} onSave={(xml, preview) => updateNodeAt(dialog.pos, node, { xml, preview })} />
        if (obj.kind === 'chart') return <ChartDialog chart={obj.chart} onClose={close} onSave={(chart) => updateNodeAt(dialog.pos, node, { object: { ...obj, chart }, preview: chart.title || 'Chart' })} />
        if (obj.kind === 'smartart') return <SmartArtDialog data={obj.smartArt} onClose={close} onSave={(smartArt) => updateNodeAt(dialog.pos, node, { object: { ...obj, smartArt }, preview: smartArt.items.map((i) => i.text).filter(Boolean).join(' · ').slice(0, 200) })} />
        return null
      })()}
      {dialog?.kind === 'newChart' && (
        <ChartDialog
          chart={newChart('column')}
          onClose={() => setDialog(null)}
          onSave={(chart) => {
            const p = live.meta.sections[0].page
            const width = Math.min(480, Math.round(twipToPx(p.width - p.marginLeft - p.marginRight)))
            editor.chain().focus().insertContent({ type: 'opaqueInline', attrs: { xml: '', label: 'Chart', preview: chart.title || 'Chart', object: { kind: 'chart', chart, width, height: Math.round(width * 0.6) } } }).run()
            setDialog(null)
          }}
        />
      )}
      {dialog?.kind === 'equation' && <EquationDialog xml={null} onClose={() => setDialog(null)} onSave={insertEquation} />}
      {dialog?.kind === 'newShape' && (
        <ShapeDialog
          initial={dialog.shape}
          onClose={() => setDialog(null)}
          onSave={(a) => {
            editor.chain().focus().insertContent({ type: 'shape', attrs: a }).run()
            setDialog(null)
          }}
        />
      )}
      {dialog?.kind === 'draw' && <DrawDialog onClose={() => setDialog(null)} onSave={insertDrawing} />}
      {dialog?.kind === 'customXml' && (
        <CustomXmlDialog
          parts={meta.customXml ?? []}
          readOnly={readOnly}
          onClose={() => setDialog(null)}
          onSave={(parts) => {
            setMeta((m) => ({ ...m, customXml: parts }))
            setDialog(null)
            toast('Custom XML saved. Edited parts are written when you export to DOCX.', 'success')
          }}
        />
      )}
      </Suspense>
      {dialog?.kind === 'sources' && (
        <SourcesDialog
          sources={meta.sources ?? []}
          style={meta.citationStyle ?? 'APA'}
          cited={citedTags(editor)}
          startNew={dialog.thenCite && !(meta.sources ?? []).length}
          onClose={() => setDialog(null)}
          onSave={(sources) => {
            setMeta((m) => ({ ...m, sources }))
            refreshCitations(editor, { ...live.meta, sources })
            setDialog(dialog.thenCite ? { kind: 'citation' } : null)
          }}
        />
      )}
      {dialog?.kind === 'citation' && (
        <CitationDialog
          sources={meta.sources ?? []}
          style={meta.citationStyle ?? 'APA'}
          onClose={() => setDialog(null)}
          onManage={() => setDialog({ kind: 'sources', thenCite: true })}
          onInsert={(tags, page) => {
            const found = tags.map((t) => (live.meta.sources ?? []).find((x) => x.tag === t)!).filter(Boolean)
            editor.chain().focus().insertContent({ type: 'field', attrs: { instr: citationInstr(tags, page), result: formatCitation(found, live.meta.citationStyle ?? 'APA', page) } }).run()
            setDialog(null)
          }}
        />
      )}
      {dialog?.kind === 'style' && (
        <ModifyStyleDialog
          style={dialog.style}
          styles={galleryStyles(meta)}
          isNew={dialog.isNew}
          onFromSelection={() => styleFromSelection(editor)}
          onClose={() => setDialog(null)}
          onSave={(st) => {
            saveStyle(st)
            if (dialog.isNew) applyParagraphStyle(editor, st, { ...live.meta, styles: [...live.meta.styles, st] }, setMeta)
          }}
          onDelete={
            dialog.style.custom && !dialog.isNew
              ? () => {
                  setMeta((m) => ({ ...m, styles: m.styles.filter((x) => x.id !== dialog.style.id) }))
                  setDialog(null)
                }
              : undefined
          }
        />
      )}
    </div>
  )
}

function currentPageIndex(editor: Editor, pages: PageInfo[], zoom: number) {
  try {
    const c = editor.view.coordsAtPos(editor.state.selection.from)
    const root = editor.view.dom.getBoundingClientRect()
    const y = (c.top - root.top) / zoom
    const i = pages.findIndex((p) => y < p.top + p.height)
    return i === -1 ? pages.length - 1 : i
  } catch {
    return 0
  }
}

function WordCountLive({ editor }: { editor: Editor }) {
  const [n, setN] = useState(0)
  useEffect(() => {
    let t: ReturnType<typeof setTimeout>
    const count = () => setN(editor.getText().match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu)?.length ?? 0)
    const on = () => {
      clearTimeout(t)
      t = setTimeout(count, 400)
    }
    count()
    editor.on('update', on)
    return () => {
      editor.off('update', on)
      clearTimeout(t)
    }
  }, [editor])
  return <span>{n.toLocaleString()} words</span>
}

/** Page backgrounds with headers and footers rendered per page (fields substituted). */
/** Footnote/endnote display order labels, from the references in the document. */
function useNoteLabels(editor: Editor | null) {
  const [labels, setLabels] = useState<Record<string, string>>({})
  useEffect(() => {
    if (!editor) return
    const read = () => {
      const out: Record<string, string> = {}
      editor.state.doc.descendants((n) => {
        if (n.type.name === 'noteRef') out[`${n.attrs.kind === 'endnote' ? 'endnote' : 'footnote'}:${n.attrs.noteId}`] = String(n.attrs.label || '')
      })
      setLabels((cur) => (JSON.stringify(cur) === JSON.stringify(out) ? cur : out))
    }
    read()
    editor.on('update', read)
    return () => {
      editor.off('update', read)
    }
  }, [editor])
  return labels
}

/** Position of a note in reading order (labels are assigned in document order). */
const noteOrder = (labels: Record<string, string>, kind: string, id: string) => {
  const keys = Object.keys(labels).filter((k) => k.startsWith(kind + ':'))
  const i = keys.indexOf(`${kind}:${id}`)
  return i === -1 ? 1e9 : i
}

const contentWidth = (m: DocMeta) => {
  const p = m.sections[0].page
  return twipToPx(p.width - p.marginLeft - p.marginRight)
}

let measurer: HTMLDivElement | null = null
/** Rendered height of note text in the page-notes style (same CSS as the page frames). */
function measureNotes(notes: { label: string; text: string }[], width: number, endnotes = false): number {
  if (typeof document === 'undefined') return 0
  if (!measurer) {
    measurer = document.createElement('div')
    measurer.setAttribute('aria-hidden', 'true')
    measurer.style.cssText = 'position:absolute;left:-10000px;top:0;visibility:hidden;pointer-events:none'
    document.body.appendChild(measurer)
  }
  measurer.className = `sc-page-notes sc-doc${endnotes ? ' is-endnotes' : ''}`
  measurer.style.width = `${width}px`
  measurer.replaceChildren()
  if (endnotes) {
    const h = document.createElement('div')
    h.className = 'sc-notes-title'
    h.textContent = 'Endnotes'
    measurer.appendChild(h)
  }
  for (const n of notes) {
    const p = document.createElement('p')
    const sup = document.createElement('sup')
    sup.textContent = n.label
    p.append(sup, ' ', n.text || ' ')
    measurer.appendChild(p)
  }
  return Math.ceil(measurer.getBoundingClientRect().height)
}

function NoteList({ notes, kind, labels }: { notes: { id: string; text: string }[]; kind: 'footnote' | 'endnote'; labels: Record<string, string> }) {
  return (
    <>
      {notes.map((n) => (
        <p key={n.id} data-note={n.id}>
          <sup>{labels[`${kind}:${n.id}`] || '*'}</sup> {n.text}
        </p>
      ))}
    </>
  )
}

function PageFrames({ pages, meta, offsetTop, labels }: { pages: PageInfo[]; meta: DocMeta; offsetTop: number; labels: Record<string, string> }) {
  const total = pages.length
  const extensions = useMemo(() => schemaExtensions(), [])
  const htmlCache = useMemo(() => {
    const c: Record<string, string> = {}
    for (const [k, v] of Object.entries(meta.headerFooters)) {
      try {
        c[k] = generateHTML(v as any, extensions)
      } catch {
        c[k] = ''
      }
    }
    return c
  }, [meta.headerFooters, extensions])
  const withFields = (html: string, page: PageInfo) =>
    html.replace(/(<span[^>]*data-instr="\s*(PAGE|NUMPAGES|SECTIONPAGES)[^"]*"[^>]*>)([^<]*)(<\/span>)/g, (_m, open, instr, _old, close) => `${open}${instr === 'PAGE' ? page.number : total}${close}`)
  return (
    <div className="sc-frames" aria-hidden="true">
      {pages.map((p) => {
        const sec = meta.sections[p.section] ?? meta.sections[0]
        const first = sec.titlePage && p.firstOfSection
        const hk = first ? sec.headerFirst : sec.headerDefault
        const fk = first ? sec.footerFirst : sec.footerDefault
        return (
          <div key={p.index} className="sc-page" style={{ top: p.top + offsetTop, width: p.width, height: p.height }}>
            {hk && htmlCache[hk] && (
              <div className="sc-page-header sc-doc" style={{ top: p.headerOffset, left: p.marginLeft, right: p.marginRight }} dangerouslySetInnerHTML={{ __html: withFields(htmlCache[hk], p) }} />
            )}
            {fk && htmlCache[fk] && (
              <div className="sc-page-footer sc-doc" style={{ bottom: p.footerOffset, left: p.marginLeft, right: p.marginRight }} dangerouslySetInnerHTML={{ __html: withFields(htmlCache[fk], p) }} />
            )}
            {p.notes.length > 0 && (
              <div className="sc-page-notes sc-doc" style={{ bottom: p.marginBottom, left: p.marginLeft, right: p.marginRight, height: p.noteSpace }}>
                <NoteList notes={p.notes.map((id) => meta.footnotes.find((f) => f.id === id)).filter((f): f is NonNullable<typeof f> => !!f)} kind="footnote" labels={labels} />
              </div>
            )}
            {p.endnotesTop !== undefined && (
              <div className="sc-page-notes sc-doc is-endnotes" style={{ top: p.endnotesTop, left: p.marginLeft, right: p.marginRight }}>
                <div className="sc-notes-title">Endnotes</div>
                <NoteList notes={meta.endnotes.filter((n) => labels[`endnote:${n.id}`]).sort((a, b) => noteOrder(labels, 'endnote', a.id) - noteOrder(labels, 'endnote', b.id))} kind="endnote" labels={labels} />
              </div>
            )}
            <span className="sc-page-number-hint">{p.index + 1}</span>
          </div>
        )
      })}
    </div>
  )
}

/** Style-sheet styles (with inheritance already resolved at import) → editor CSS. */
function buildStyleCss(meta: DocMeta) {
  const d = meta.defaults
  const esc = (s: string) => s.replace(/["\\]/g, '')
  const pg = meta.sections[0].page
  // a picture never grows past the text area of a page (the file keeps its real size)
  let css = `.sc-doc img{max-height:${Math.floor(twipToPx(pg.height - pg.marginTop - pg.marginBottom))}px;object-fit:contain}`
  css += `.sc-doc{font-family:"${esc(d.fontFamily)}",Calibri,Carlito,"Segoe UI",Arial,sans-serif;font-size:${d.fontSizePt}pt;line-height:${d.lineHeight}}`
  // the document's styles plus built-in gallery styles it doesn't define (Quote, Subtitle…)
  for (const s of galleryStyles(meta)) {
    if (s.type !== 'paragraph') continue
    const r: string[] = []
    if (s.fontFamily) r.push(`font-family:"${esc(s.fontFamily)}",sans-serif`)
    if (s.fontSizePt) r.push(`font-size:${s.fontSizePt}pt`)
    if (s.bold != null) r.push(`font-weight:${s.bold ? 700 : 400}`)
    if (s.italic != null) r.push(`font-style:${s.italic ? 'italic' : 'normal'}`)
    if (s.color) r.push(`color:${s.color}`)
    if (s.spacingBefore != null) r.push(`margin-top:${twipToPx(s.spacingBefore)}px`)
    if (s.spacingAfter != null) r.push(`margin-bottom:${twipToPx(s.spacingAfter)}px`)
    if (s.underline != null) r.push(`text-decoration:${s.underline ? 'underline' : 'none'}`)
    if (s.alignment) r.push(`text-align:${s.alignment}`)
    if (s.lineSpacing) r.push(`line-height:${s.lineSpacing}`)
    if (!r.length) continue
    css += `.sc-doc [data-style-id="${esc(s.id)}"]{${r.join(';')}}`
    // headings without an explicit style follow the document's (or edited) heading styles
    const lvl = headingLevelOfStyle(s)
    if (lvl && (s.modified || s.custom || meta.styles.includes(s))) css += `.sc-doc h${lvl}:not([data-style-id]){${r.join(';')}}`
  }
  return css
}

export { emptyMeta }

/** Mutable values read outside rendering by editor plugins, timers and the print flow. */
class LiveState {
  gap = 24
  /** unsaved changes since the last snapshot */
  private dirty = false
  markDirty() {
    this.dirty = true
  }
  /** returns whether there were unsaved changes, and clears the flag */
  takeDirty() {
    const d = this.dirty
    this.dirty = false
    return d
  }
  private printZoom = 1
  private lastVersionAt = Date.now()
  constructor(
    public meta: DocMeta,
    public zoom: number,
  ) {}
  setMeta(m: DocMeta) {
    this.meta = m
  }
  setZoom(z: number) {
    this.zoom = z
  }
  beginPrint() {
    this.printZoom = this.zoom
    this.zoom = 1
    this.gap = 0
  }
  endPrint() {
    this.zoom = this.printZoom
    this.gap = 24
  }
  /** true once every 10 minutes of editing: time for an automatic version */
  versionDue() {
    if (Date.now() - this.lastVersionAt < 10 * 60 * 1000) return false
    this.lastVersionAt = Date.now()
    return true
  }
}
