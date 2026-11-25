import { useCallback, useEffect, useMemo, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useEditorState } from '@tiptap/react'
import { generateHTML } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { Check, Copy, Loader2, RotateCcw, Sparkles, Trash2, X } from 'lucide-react'
import { api } from '../api'
import type { CommentThread, DocMeta, PMNode } from '../model/types'
import type { PageInfo } from '../layout/pagination'
import { findKey, gotoMatch, listChanges, replaceMatch, resolveChange, setFind, type FindState } from '../editor/plugins'
import { schemaExtensions } from '../model/extensions'
import { Dialog, useToast } from './kit'

// ---------------------------------------------------------------------------
// Navigation: pages + outline

export function PagesPanel({ editor, pages, onGoto }: { editor: Editor; pages: PageInfo[]; onGoto: (page: PageInfo) => void }) {
  const outline = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const out: { level: number; text: string; pos: number }[] = []
      e.state.doc.descendants((n, pos) => {
        if (n.type.name === 'heading') out.push({ level: n.attrs.level, text: n.textContent || '(empty heading)', pos })
        return n.type.name !== 'heading' && n.type.name !== 'paragraph'
      })
      return out
    },
    equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b),
  })
  const goHeading = (pos: number) => {
    editor.chain().focus().setTextSelection(pos + 1).run()
    const dom = editor.view.nodeDOM(pos) as HTMLElement | null
    dom?.scrollIntoView({ block: 'start', behavior: 'smooth' })
  }
  return (
    <div className="sc-panel-body">
      <h3>Pages</h3>
      <ol className="sc-page-list">
        {pages.map((p) => (
          <li key={p.index}>
            <button onClick={() => onGoto(p)} aria-label={`Go to page ${p.index + 1}`}>
              <span className="sc-page-thumb" style={{ aspectRatio: `${p.width} / ${p.height}` }} aria-hidden="true">
                <i />
                <i />
                <i />
              </span>
              <span>Page {p.index + 1}</span>
              {p.section > 0 && <span className="sc-muted">§{p.section + 1}</span>}
            </button>
          </li>
        ))}
      </ol>
      <h3>Outline</h3>
      {outline.length === 0 ? (
        <p className="sc-muted">Headings appear here.</p>
      ) : (
        <ul className="sc-outline">
          {outline.map((h, i) => (
            <li key={i} style={{ paddingLeft: (h.level - 1) * 12 }}>
              <button onClick={() => goHeading(h.pos)}>{h.text}</button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

// ---------------------------------------------------------------------------
// Comments

function commentRanges(editor: Editor, id: string) {
  const ranges: { from: number; to: number }[] = []
  editor.state.doc.descendants((n, pos) => {
    if (n.isText && n.marks.some((m) => m.type.name === 'comment' && String(m.attrs.commentId) === id)) {
      const last = ranges[ranges.length - 1]
      if (last && last.to === pos) last.to = pos + n.nodeSize
      else ranges.push({ from: pos, to: pos + n.nodeSize })
    }
  })
  return ranges
}

export function CommentsPanel({ editor, meta, setMeta, user, readOnly, composing, onComposed }: { editor: Editor; meta: DocMeta; setMeta: (f: (m: DocMeta) => DocMeta) => void; user: { name: string }; readOnly: boolean; composing: boolean; onComposed: () => void }) {
  const [draft, setDraft] = useState('')
  const [editing, setEditing] = useState<string | null>(null)
  const [editText, setEditText] = useState('')
  const anchored = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const ids: string[] = []
      e.state.doc.descendants((n) => {
        if (n.isText) for (const m of n.marks) if (m.type.name === 'comment' && !ids.includes(String(m.attrs.commentId))) ids.push(String(m.attrs.commentId))
      })
      return ids
    },
    equalityFn: (a, b) => a.join() === b?.join(),
  })
  const add = () => {
    const { from, to, empty } = editor.state.selection
    if (empty || !draft.trim()) return
    const id = `c${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`
    editor.chain().focus().setTextSelection({ from, to }).setMark('comment', { commentId: id }).run()
    setMeta((m) => ({ ...m, comments: [...m.comments, { id, author: user.name, date: new Date().toISOString(), text: draft.trim() }] }))
    setDraft('')
    onComposed()
  }
  const select = (c: CommentThread) => {
    const r = commentRanges(editor, c.id)[0]
    if (!r) return
    editor.chain().focus().setTextSelection(r).run()
    ;(editor.view.domAtPos(r.from).node as HTMLElement)?.parentElement?.scrollIntoView({ block: 'center', behavior: 'smooth' })
  }
  const remove = (c: CommentThread) => {
    const tr = editor.state.tr
    for (const r of commentRanges(editor, c.id)) tr.removeMark(r.from, r.to, editor.schema.marks.comment.create({ commentId: c.id }))
    editor.view.dispatch(tr)
    setMeta((m) => ({ ...m, comments: m.comments.filter((x) => x.id !== c.id) }))
  }
  const update = (id: string, patch: Partial<CommentThread>) => setMeta((m) => ({ ...m, comments: m.comments.map((x) => (x.id === id ? { ...x, ...patch, pristine: false } : x)) }))
  const list = meta.comments.filter((c) => anchored.includes(c.id))
  const hasSelection = !editor.state.selection.empty
  return (
    <div className="sc-panel-body">
      <h3>Comments</h3>
      {!readOnly && (composing || hasSelection) && (
        <div className="sc-comment-new">
          <textarea autoFocus={composing} placeholder={hasSelection ? 'Comment on the selected text…' : 'Select text to comment on'} value={draft} onChange={(e) => setDraft(e.target.value)} disabled={!hasSelection} aria-label="New comment" />
          <button className="sc-primary" disabled={!hasSelection || !draft.trim()} onClick={add}>
            Comment
          </button>
        </div>
      )}
      {list.length === 0 && <p className="sc-muted">No comments. Select text and choose Insert → Comment.</p>}
      <ul className="sc-comments">
        {list.map((c) => (
          <li key={c.id} className={c.resolved ? 'is-resolved' : ''}>
            <button className="sc-comment-head" onClick={() => select(c)}>
              <strong>{c.author}</strong>
              <span className="sc-muted">{c.date ? new Date(c.date).toLocaleString() : ''}</span>
            </button>
            {editing === c.id ? (
              <div className="sc-comment-new">
                <textarea value={editText} onChange={(e) => setEditText(e.target.value)} aria-label="Edit comment" />
                <button className="sc-primary" onClick={() => { update(c.id, { text: editText }); setEditing(null) }}>Save</button>
                <button onClick={() => setEditing(null)}>Cancel</button>
              </div>
            ) : (
              <p>{c.text}</p>
            )}
            {!readOnly && editing !== c.id && (
              <div className="sc-comment-actions">
                <button className="sc-link" onClick={() => update(c.id, { resolved: !c.resolved })}>
                  {c.resolved ? 'Reopen' : 'Resolve'}
                </button>
                <button className="sc-link" onClick={() => { setEditing(c.id); setEditText(c.text) }}>Edit</button>
                <button className="sc-link is-danger" onClick={() => remove(c)}>
                  Delete
                </button>
              </div>
            )}
          </li>
        ))}
      </ul>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Tracked changes

export function ChangesPanel({ editor, readOnly }: { editor: Editor; readOnly: boolean }) {
  const changes = useEditorState({ editor, selector: ({ editor: e }) => listChanges(e.state.doc), equalityFn: (a, b) => JSON.stringify(a) === JSON.stringify(b) })
  return (
    <div className="sc-panel-body">
      <h3>Tracked changes</h3>
      {changes.length === 0 && <p className="sc-muted">No tracked changes. Turn on Review → Track changes to suggest edits.</p>}
      <ul className="sc-changes">
        {changes.map((ch, i) => (
          <li key={i} className={`is-${ch.type}`}>
            <button className="sc-comment-head" onClick={() => editor.chain().focus().setTextSelection({ from: ch.from, to: ch.to }).scrollIntoView().run()}>
              <strong>{ch.type === 'insertion' ? 'Inserted' : 'Deleted'}</strong>
              <span className="sc-muted">
                {ch.author || 'Unknown'} {ch.date ? '· ' + new Date(ch.date).toLocaleDateString() : ''}
              </span>
            </button>
            <p className="sc-change-text">{ch.text}</p>
            {!readOnly && (
              <div className="sc-comment-actions">
                <button className="sc-link" onClick={() => resolveChange(editor, ch, true)}>
                  <Check size={14} /> Accept
                </button>
                <button className="sc-link is-danger" onClick={() => resolveChange(editor, ch, false)}>
                  <X size={14} /> Reject
                </button>
              </div>
            )}
          </li>
        ))}
      </ul>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Version history

export function VersionsPanel({ docId, readOnly, onSaveVersion, onRestore }: { docId: string; readOnly: boolean; onSaveVersion: () => Promise<void>; onRestore: (content: PMNode, meta: DocMeta) => void }) {
  const [versions, setVersions] = useState<{ id: string; created: number; label: string | null; auto: boolean; author: string | null }[] | null>(null)
  const [preview, setPreview] = useState<{ id: string; created: number; html: string; content: PMNode; meta: DocMeta } | null>(null)
  const toast = useToast()
  const load = useCallback(
    () => api<{ versions: any[] }>(`/api/docs/${docId}/versions`).then((r) => setVersions(r.versions)).catch((e) => toast(e.message, 'error')),
    [docId, toast],
  )
  useEffect(() => {
    load()
  }, [load])
  const open = async (id: string) => {
    try {
      const v = await api<{ id: string; created: number; content: PMNode; meta: DocMeta }>(`/api/docs/${docId}/versions/${id}`)
      setPreview({ ...v, html: generateHTML(v.content as any, schemaExtensions()) })
    } catch (e: any) {
      toast(e.message, 'error')
    }
  }
  return (
    <div className="sc-panel-body">
      <h3>Version history</h3>
      {!readOnly && (
        <button className="sc-primary" onClick={() => onSaveVersion().then(load)}>
          Save current version
        </button>
      )}
      {versions === null ? (
        <p className="sc-muted">Loading…</p>
      ) : versions.length === 0 ? (
        <p className="sc-muted">No saved versions yet. SCRIBE also saves one automatically every 10 minutes while you edit.</p>
      ) : (
        <ul className="sc-versions">
          {versions.map((v) => (
            <li key={v.id}>
              <button onClick={() => open(v.id)}>
                <strong>{v.label || (v.auto ? 'Automatic save' : 'Saved version')}</strong>
                <span className="sc-muted">
                  {new Date(v.created).toLocaleString()} {v.author ? `· ${v.author}` : ''}
                </span>
              </button>
            </li>
          ))}
        </ul>
      )}
      {preview && (
        <Dialog
          title={`Version from ${new Date(preview.created).toLocaleString()}`}
          wide
          onClose={() => setPreview(null)}
          footer={
            <>
              <button onClick={() => setPreview(null)}>Close</button>
              {!readOnly && (
                <button className="sc-primary with-icon" onClick={() => { onRestore(preview.content, preview.meta); setPreview(null) }}>
                  <RotateCcw size={16} /> Restore this version
                </button>
              )}
            </>
          }
        >
          {/* HTML is produced by the editor schema serializer from stored JSON, so it contains no raw user HTML */}
          <div className="sc-version-preview ProseMirror" dangerouslySetInnerHTML={{ __html: preview.html }} />
        </Dialog>
      )}
    </div>
  )
}

// ---------------------------------------------------------------------------
// AI assistant

const AI_ACTIONS: { task: string; label: string }[] = [
  { task: 'improve', label: 'Improve writing' },
  { task: 'grammar', label: 'Fix spelling & grammar' },
  { task: 'summarize', label: 'Summarize' },
  { task: 'shorter', label: 'Make shorter' },
  { task: 'longer', label: 'Make longer' },
  { task: 'formal', label: 'Formal tone' },
  { task: 'friendly', label: 'Friendly tone' },
  { task: 'confident', label: 'Confident tone' },
  { task: 'continue', label: 'Continue writing' },
]

export function AiPanel({ editor, readOnly }: { editor: Editor; readOnly: boolean }) {
  const [status, setStatus] = useState<{ provider: string; model: string | null; reason: string | null } | null>(null)
  const [busy, setBusy] = useState(false)
  const [result, setResult] = useState<{ text: string; task: string; range: { from: number; to: number } | null; truncated?: boolean } | null>(null)
  const [error, setError] = useState('')
  const [prompt, setPrompt] = useState('')
  useEffect(() => {
    api('/api/ai/status').then(setStatus).catch((e) => setError(e.message))
  }, [])
  const selection = () => {
    const { from, to, empty } = editor.state.selection
    return empty ? null : { from, to, text: editor.state.doc.textBetween(from, to, '\n\n', ' ') }
  }
  const run = async (task: string, instruction?: string) => {
    setError('')
    const sel = selection()
    const docText = editor.state.doc.textBetween(0, editor.state.doc.content.size, '\n\n', ' ')
    const text = task === 'generate' ? '' : sel?.text ?? (task === 'continue' ? docText.slice(-6000) : docText)
    if (!text && task !== 'generate') return setError('Select some text, or write something first.')
    setBusy(true)
    try {
      const r = await api<{ text: string; truncated: boolean }>('/api/ai', { method: 'POST', body: { task, text, instruction, context: sel ? docText.slice(0, 8000) : undefined } })
      setResult({ text: r.text, task, range: sel ? { from: sel.from, to: sel.to } : null, truncated: r.truncated })
    } catch (e: any) {
      setError(e.message)
    } finally {
      setBusy(false)
    }
  }
  const toNodes = (text: string) => text.split(/\n{2,}/).map((para) => ({ type: 'paragraph', content: para.trim() ? para.split('\n').flatMap((line, i) => [...(i ? [{ type: 'hardBreak' }] : []), ...(line ? [{ type: 'text', text: line }] : [])]) : undefined }))
  const apply = (mode: 'replace' | 'insert') => {
    if (!result) return
    const nodes = toNodes(result.text)
    const chain = editor.chain().focus()
    if (mode === 'replace' && result.range) chain.insertContentAt(result.range, nodes.length === 1 ? nodes[0].content ?? [] : nodes).run()
    else {
      const at = result.range ? editor.state.doc.resolve(result.range.to).after(1) : editor.state.doc.content.size
      chain.insertContentAt(Math.min(at, editor.state.doc.content.size), nodes).run()
    }
    setResult(null)
  }
  if (status && status.provider === 'none')
    return (
      <div className="sc-panel-body">
        <h3>AI assistant</h3>
        <p className="sc-muted">AI is not configured on this server.</p>
        <p className="sc-muted">{status.reason}</p>
      </div>
    )
  return (
    <div className="sc-panel-body">
      <h3>
        <Sparkles size={16} aria-hidden="true" /> AI assistant
      </h3>
      {status?.model && <p className="sc-muted">Model: {status.model}. Works on the selection, or the whole document if nothing is selected.</p>}
      <div className="sc-ai-actions">
        {AI_ACTIONS.map((a) => (
          <button key={a.task} disabled={busy || readOnly} onClick={() => run(a.task)}>
            {a.label}
          </button>
        ))}
      </div>
      <form className="sc-ai-prompt" onSubmit={(e) => { e.preventDefault(); if (prompt.trim()) run(selection() ? 'custom' : 'generate', prompt.trim()) }}>
        <textarea placeholder={'Ask AI to write or change something…\ne.g. “Draft an intro about our Q3 results”'} value={prompt} onChange={(e) => setPrompt(e.target.value)} aria-label="AI instruction" disabled={readOnly} />
        <button className="sc-primary" disabled={busy || !prompt.trim() || readOnly}>
          {busy ? <Loader2 size={16} className="sc-spin" /> : 'Run'}
        </button>
      </form>
      {busy && <p className="sc-muted" role="status">Working…</p>}
      {error && <p className="sc-error" role="alert">{error}</p>}
      {result && (
        <div className="sc-ai-result">
          <div className="sc-ai-text">{result.text}</div>
          {result.truncated && <p className="sc-muted">The response hit its length limit and may be incomplete.</p>}
          <div className="sc-comment-actions">
            {result.range && <button className="sc-primary" onClick={() => apply('replace')}>Replace selection</button>}
            <button onClick={() => apply('insert')}>Insert below</button>
            <button onClick={() => navigator.clipboard.writeText(result.text)} aria-label="Copy result">
              <Copy size={14} />
            </button>
            <button onClick={() => setResult(null)} aria-label="Discard result">
              <Trash2 size={14} />
            </button>
          </div>
        </div>
      )}
    </div>
  )
}

// ---------------------------------------------------------------------------
// Find & replace bar

export function FindBar({ editor, replace: showReplaceInitially, readOnly, onClose }: { editor: Editor; replace: boolean; readOnly: boolean; onClose: () => void }) {
  const [query, setQuery] = useState('')
  const [rep, setRep] = useState('')
  const [showReplace, setShowReplace] = useState(showReplaceInitially)
  const [caseSensitive, setCase] = useState(false)
  const [wholeWord, setWhole] = useState(false)
  const toast = useToast()
  const st = useEditorState({ editor, selector: ({ editor: e }) => findKey.getState(e.state) as FindState })
  useEffect(() => {
    setFind(editor.view, { query, caseSensitive, wholeWord, current: 0 })
  }, [query, caseSensitive, wholeWord, editor])
  useEffect(() => () => setFind(editor.view, { query: '' }), [editor])
  const count = st?.matches.length ?? 0
  const label = useMemo(() => (query ? (count ? `${(st?.current ?? 0) + 1} of ${count}` : 'No results') : ''), [query, count, st])
  return (
    <div className="sc-findbar" role="search" onKeyDown={(e) => e.key === 'Escape' && onClose()}>
      <input
        autoFocus
        placeholder="Find in document"
        aria-label="Find"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault()
            gotoMatch(editor, e.shiftKey ? -1 : 1)
          }
        }}
      />
      <span className="sc-muted sc-find-count" aria-live="polite">
        {label}
      </span>
      <button onClick={() => gotoMatch(editor, -1)} disabled={!count} aria-label="Previous match">
        ↑
      </button>
      <button onClick={() => gotoMatch(editor, 1)} disabled={!count} aria-label="Next match">
        ↓
      </button>
      <label className="sc-check">
        <input type="checkbox" checked={caseSensitive} onChange={(e) => setCase(e.target.checked)} /> Match case
      </label>
      <label className="sc-check">
        <input type="checkbox" checked={wholeWord} onChange={(e) => setWhole(e.target.checked)} /> Whole words
      </label>
      {!readOnly && !showReplace && (
        <button className="sc-link" onClick={() => setShowReplace(true)}>
          Replace…
        </button>
      )}
      {!readOnly && showReplace && (
        <>
          <input placeholder="Replace with" aria-label="Replace with" value={rep} onChange={(e) => setRep(e.target.value)} />
          <button disabled={!count} onClick={() => replaceMatch(editor, rep, false)}>
            Replace
          </button>
          <button disabled={!count} onClick={() => toast(`Replaced ${replaceMatch(editor, rep, true)} occurrence(s)`, 'success')}>
            Replace all
          </button>
        </>
      )}
      <button onClick={onClose} aria-label="Close find">
        <X size={16} />
      </button>
    </div>
  )
}

export { TextSelection }
