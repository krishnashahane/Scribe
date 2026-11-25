import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { EditorContent, useEditor, type Editor } from '@tiptap/react'
import { generateHTML } from '@tiptap/core'
import Collaboration from '@tiptap/extension-collaboration'
import CollaborationCaret from '@tiptap/extension-collaboration-caret'
import {
  ArrowLeft, Cloud, CloudOff, FileDown, History, Loader2, MessageSquare, PanelLeft, PenLine, Share2, Sparkles,
} from 'lucide-react'
import { api, type User } from '../api'
import { openSession, userColor, type CollabSession, type SyncStatus } from '../collab/session'
import { schemaExtensions, newId as newNodeId } from '../model/extensions'
import { emptyMeta, migrateDocument, PAGE_SIZES, twipToPx, type DocMeta, type PMNode, type ScribeDocument } from '../model/types'
import { Pagination, repaginate, type PageInfo } from '../layout/pagination'
import { FindReplace, SectionSync, Suggesting, isSuggesting, setSuggesting } from '../editor/plugins'
import { exportDocx } from '../docx/export'
import { importDocx } from '../docx/import'
import { validatePackage } from '../docx/validate'
import { contentToText, documentToHtml, download, safeFilename, sanitizeHtml } from '../formats/formats'
import { Ribbon, type RibbonAction } from './Ribbon'
import { AiPanel, ChangesPanel, CommentsPanel, FindBar, PagesPanel, VersionsPanel } from './Panels'
import {
  CompatReport, HeaderFooterDialog, LinkDialog, PageSetupDialog, PropertiesDialog, ShareDialog, ShortcutsDialog, TextPrompt, WordCountDialog,
} from './Dialogs'
import { Menu, mod, useToast } from './kit'
import type { CompatibilityIssue } from '../model/types'

type Panel = 'pages' | 'comments' | 'changes' | 'versions' | 'ai'
type DialogState =
  | { kind: 'pageSetup' }
  | { kind: 'hf'; which: 'header' | 'footer'; section: number; first: boolean }
  | { kind: 'link' }
  | { kind: 'bookmark' }
  | { kind: 'share' }
  | { kind: 'props' }
  | { kind: 'shortcuts' }
  | { kind: 'wordCount' }
  | { kind: 'report'; issues: CompatibilityIssue[]; title: string }
  | null

interface DocInfo {
  title: string
  role: 'owner' | 'editor' | 'commenter' | 'viewer'
  hasOriginal: boolean
  originalName: string | null
  meta: DocMeta
}

export function EditorScreen({ docId, user, theme, onToggleTheme }: { docId: string; user: User; theme: 'light' | 'dark'; onToggleTheme: () => void }) {
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
        setInfo({ title: d.title, role: d.role, hasOriginal: d.hasOriginal, originalName: d.originalName, meta: migrateDocument({ content: d.content, meta: d.meta }).meta })
        s = openSession(docId, setStatus)
        setSession(s)
        // render once the local (IndexedDB) or remote state has arrived, so the editor never starts empty
        Promise.race([s.whenSynced, s.local.whenSynced.then(() => (s!.ydoc.getXmlFragment('default').length ? undefined : s!.whenSynced))]).then(() => !cancelled && setReady(true))
      })
      .catch((e) => setError(e.message))
    return () => {
      cancelled = true
      s?.destroy()
    }
  }, [docId])

  if (error)
    return (
      <div className="sc-splash">
        <p className="sc-error">{error}</p>
        <a href="#/">Back to documents</a>
      </div>
    )
  if (!info || !session || !ready)
    return (
      <div className="sc-splash" aria-busy="true">
        <Loader2 className="sc-spin" />
        <span>Opening document…</span>
      </div>
    )
  return <DocEditor docId={docId} info={info} session={session} status={status} user={user} theme={theme} onToggleTheme={onToggleTheme} />
}

function DocEditor({ docId, info, session, status, user, theme, onToggleTheme }: { docId: string; info: DocInfo; session: CollabSession; status: SyncStatus; user: User; theme: 'light' | 'dark'; onToggleTheme: () => void }) {
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
  const metaRef = useRef(meta)
  metaRef.current = meta
  const [title, setTitle] = useState(info.title)
  const [pages, setPages] = useState<PageInfo[]>([])
  const [panel, setPanel] = useState<Panel | null>('pages')
  const [dialog, setDialog] = useState<DialogState>(null)
  const [find, setFindOpen] = useState<null | { replace: boolean }>(null)
  const [zoom, setZoom] = useState(() => Number(localStorage.getItem('scribe-zoom')) || 1)
  const zoomRef = useRef(zoom)
  zoomRef.current = zoom
  const gapRef = useRef(24)
  const [showMarks, setShowMarks] = useState(false)
  const [composingComment, setComposingComment] = useState(false)
  const [saveState, setSaveState] = useState<'saved' | 'saving' | 'error'>('saved')
  const dirty = useRef(false)
  const lastVersionAt = useRef(Date.now())
  const imageInput = useRef<HTMLInputElement>(null)

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
      const next = f(metaRef.current)
      metaRef.current = next
      ymeta.set('meta', JSON.stringify(next))
      setMetaState(next)
      dirty.current = true
    },
    [ymeta, readOnly],
  )

  const editor = useEditor(
    {
      editable: !readOnly,
      extensions: [
        ...schemaExtensions({ collaborative: true }),
        Collaboration.configure({ document: session.ydoc, field: 'default' }),
        CollaborationCaret.configure({ provider: session.provider, user: { name: user.name, color: userColor(user.id) } }),
        Pagination.configure({ getSections: () => metaRef.current.sections, getGap: () => gapRef.current, getScale: () => zoomRef.current, onLayout: setPages }),
        FindReplace,
        Suggesting.configure({ author: () => user.name }),
        SectionSync.configure({
          onSections: (count, removed) => {
            const m = metaRef.current
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
        handleDrop: (_view, event) => {
          const files = [...(event.dataTransfer?.files ?? [])].filter((f) => f.type.startsWith('image/'))
          if (!files.length) return false
          files.forEach(insertImageFile)
          return true
        },
      },
      onUpdate: () => {
        dirty.current = true
      },
    },
    [session],
  )

  // ---- autosave snapshot (content lives in Yjs; this snapshot feeds listings, exports on other devices and versions)
  const snapshot = useCallback((): ScribeDocument => ({ version: 1, content: editor!.getJSON() as PMNode, meta: metaRef.current }), [editor])
  const saveSnapshot = useCallback(
    async (keepalive = false) => {
      if (readOnly || !editor || !dirty.current) return
      dirty.current = false
      setSaveState('saving')
      const d = snapshot()
      try {
        if (keepalive) {
          fetch(`/api/docs/${docId}`, { method: 'PATCH', keepalive: true, headers: { 'content-type': 'application/json', 'x-scribe-request': '1' }, body: JSON.stringify({ content: d.content, meta: d.meta }) })
          return
        }
        await api(`/api/docs/${docId}`, { method: 'PATCH', body: { content: d.content, meta: d.meta } })
        if (Date.now() - lastVersionAt.current > 10 * 60 * 1000) {
          lastVersionAt.current = Date.now()
          await api(`/api/docs/${docId}/versions`, { method: 'POST', body: { content: d.content, meta: d.meta, auto: true } })
        }
        setSaveState('saved')
      } catch {
        dirty.current = true
        setSaveState('error')
      }
    },
    [editor, readOnly, docId, snapshot],
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
  }, [meta.sections, editor])
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
      const r = new FileReader()
      r.onload = () => {
        const src = String(r.result)
        const img = new Image()
        img.onload = () => {
          const p = metaRef.current.sections[0].page
          const maxW = twipToPx(p.width - p.marginLeft - p.marginRight)
          const w = Math.min(img.naturalWidth, maxW)
          const h = Math.round((img.naturalHeight * w) / img.naturalWidth)
          editor?.chain().focus().insertContent({ type: 'image', attrs: { src, width: Math.round(w), height: h, alt: file.name.replace(/\.[^.]+$/, '') } }).run()
        }
        img.onerror = () => toast('That image could not be read.', 'error')
        img.src = src
      }
      r.readAsDataURL(file)
    },
    [editor, toast],
  )

  // ---- exports
  const fetchOriginal = async () => {
    if (!info.hasOriginal) return null
    const res = await api<Response>(`/api/docs/${docId}/original`, { raw: true })
    return new Uint8Array(await res.arrayBuffer())
  }
  const exportAs = async (fmt: 'docx' | 'html' | 'txt' | 'pdf') => {
    if (!editor) return
    try {
      if (fmt === 'pdf') return printDoc()
      const d = snapshot()
      d.meta = { ...d.meta, metadata: { ...d.meta.metadata, title, lastModifiedBy: user.name } }
      if (fmt === 'docx') {
        toast('Preparing DOCX…')
        const { data, report } = await exportDocx(d, await fetchOriginal())
        const problems = await validatePackage(data)
        if (problems.length) {
          console.error('DOCX validation problems', problems)
          toast(`Export validation found ${problems.length} problem(s); the file was not downloaded. See the console for details.`, 'error')
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
    const p = metaRef.current.sections[0].page
    const style = document.createElement('style')
    style.id = 'sc-print-page'
    style.textContent = `@page { size: ${(p.width / 1440).toFixed(3)}in ${(p.height / 1440).toFixed(3)}in; margin: 0 }`
    document.head.appendChild(style)
    document.documentElement.classList.add('is-printing')
    const prevZoom = zoomRef.current
    zoomRef.current = 1
    gapRef.current = 0
    repaginate(editor?.view, true)
    const done = () => {
      document.documentElement.classList.remove('is-printing')
      style.remove()
      gapRef.current = 24
      zoomRef.current = prevZoom
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
    return Math.min(sec, metaRef.current.sections.length - 1)
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
    const sec = metaRef.current.sections[section]
    const first = !!(sec?.titlePage && page?.firstOfSection)
    setDialog({ kind: 'hf', which, section, first })
  }

  const onAction = (a: RibbonAction) => {
    if (!editor) return
    const sec = currentSection()
    switch (a) {
      case 'image':
        return imageInput.current?.click()
      case 'link':
        return setDialog({ kind: 'link' })
      case 'bookmark':
        return setDialog({ kind: 'bookmark' })
      case 'comment':
        if (editor.state.selection.empty) return toast('Select the text you want to comment on.', 'info')
        setPanel('comments')
        return setComposingComment(true)
      case 'pageSetup':
        return setDialog({ kind: 'pageSetup' })
      case 'header':
      case 'footer':
        return openHF(a, sec)
      case 'pageNumber': {
        const key = metaRef.current.sections[sec].footerDefault
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
          .then((b) => b && importDocx(b))
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
    }
  }

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
      } else if (k === 'p') {
        e.preventDefault()
        printDoc()
      } else if (k === '/') {
        e.preventDefault()
        setDialog({ kind: 'shortcuts' })
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
        <a className="sc-btn" href="#/" aria-label="Back to documents" title="Back to documents">
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
            { label: 'Save version', onClick: () => saveVersion().catch((e) => toast(e.message, 'error')), shortcut: `${mod}S`, disabled: readOnly },
            { label: 'Version history', onClick: () => setPanel('versions') },
            'sep',
            { label: 'Download as Word (.docx)', onClick: () => exportAs('docx') },
            { label: 'Download as PDF (print)', onClick: () => exportAs('pdf'), shortcut: `${mod}P` },
            { label: 'Download as web page (.html)', onClick: () => exportAs('html') },
            { label: 'Download as plain text (.txt)', onClick: () => exportAs('txt') },
            'sep',
            { label: 'Page setup…', onClick: () => setDialog({ kind: 'pageSetup' }), disabled: readOnly },
            { label: 'Document properties…', onClick: () => setDialog({ kind: 'props' }) },
            { label: 'Share…', onClick: () => setDialog({ kind: 'share' }) },
            'sep',
            { label: 'Keyboard shortcuts', onClick: () => setDialog({ kind: 'shortcuts' }), shortcut: `${mod}/` },
            { label: 'Back to documents', onClick: () => (location.hash = '#/') },
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
        onAction={onAction}
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
              ['ai', <Sparkles size={18} key="i" />, 'AI assistant'],
            ] as const
          ).map(([k, icon, label]) => (
            <button key={k} className={`sc-btn ${panel === k ? 'is-active' : ''}`} aria-label={label} title={label} aria-pressed={panel === k} onClick={() => setPanel(panel === k ? null : k)}>
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
          </aside>
        )}
        <main className="sc-canvas" aria-label="Document">
          <div className="sc-zoom" style={{ zoom } as any}>
            <div className="sc-pages-wrap" style={{ width: maxWidth, minHeight: totalHeight }} onDoubleClick={onCanvasDoubleClick}>
              <PageFrames pages={pages} meta={meta} offsetTop={twipToPx(p0.marginTop)} />
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
        <span>{Math.round(zoom * 100)}%</span>
      </footer>
      <input ref={imageInput} type="file" accept="image/png,image/jpeg,image/gif,image/webp,image/bmp" hidden onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ''; if (f) insertImageFile(f) }} />

      {dialog?.kind === 'pageSetup' && (
        <PageSetupDialog
          sections={meta.sections}
          initialSection={currentSection()}
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
      {dialog?.kind === 'report' && <CompatReport issues={dialog.issues} title={dialog.title} onClose={() => setDialog(null)} />}
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
function PageFrames({ pages, meta, offsetTop }: { pages: PageInfo[]; meta: DocMeta; offsetTop: number }) {
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
  let css = `.sc-doc{font-family:"${esc(d.fontFamily)}",Calibri,Carlito,"Segoe UI",Arial,sans-serif;font-size:${d.fontSizePt}pt;line-height:${d.lineHeight}}`
  for (const s of meta.styles) {
    if (s.type !== 'paragraph') continue
    const r: string[] = []
    if (s.fontFamily) r.push(`font-family:"${esc(s.fontFamily)}",sans-serif`)
    if (s.fontSizePt) r.push(`font-size:${s.fontSizePt}pt`)
    if (s.bold != null) r.push(`font-weight:${s.bold ? 700 : 400}`)
    if (s.italic != null) r.push(`font-style:${s.italic ? 'italic' : 'normal'}`)
    if (s.color) r.push(`color:${s.color}`)
    if (s.spacingBefore != null) r.push(`margin-top:${twipToPx(s.spacingBefore)}px`)
    if (s.spacingAfter != null) r.push(`margin-bottom:${twipToPx(s.spacingAfter)}px`)
    if (r.length) css += `.sc-doc [data-style-id="${esc(s.id)}"]{${r.join(';')}}`
  }
  return css
}

export { emptyMeta }
