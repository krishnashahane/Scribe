import { useEffect, useRef, useState } from 'react'
import { generateJSON } from '@tiptap/core'
import { FileText, LogOut, Moon, Sun, Trash2, Upload, Users } from 'lucide-react'
import { api, toBase64, type DocSummary, type User } from '../api'
import { TEMPLATES } from '../model/templates'
import { emptyMeta, type CompatibilityIssue, type DocMeta, type PMNode } from '../model/types'
import { importDocx } from '../docx/import'
import { sanitizeHtml, textToContent } from '../formats/formats'
import { schemaExtensions } from '../model/extensions'
import { initialYState } from '../collab/session'
import { Dialog, useToast } from './kit'
import { CompatReport } from './Dialogs'
import type { Theme } from './App'

export async function createDocument(title: string, content: PMNode, meta: DocMeta, original?: Uint8Array, originalName?: string) {
  const r = await api<{ id: string }>('/api/docs', {
    method: 'POST',
    body: { title, content, meta, yUpdate: initialYState(content, meta), original: original ? toBase64(original) : undefined, originalName },
  })
  return r.id
}

const MAX_IMPORT = 100 * 1024 * 1024

export function Home({ user, onSignOut, theme, onToggleTheme }: { user: User; onSignOut: () => void; theme: Theme; onToggleTheme: () => void }) {
  const [docs, setDocs] = useState<DocSummary[] | null>(null)
  const [busy, setBusy] = useState('')
  const [report, setReport] = useState<{ issues: CompatibilityIssue[]; id: string } | null>(null)
  const [confirmDelete, setConfirmDelete] = useState<DocSummary | null>(null)
  const fileRef = useRef<HTMLInputElement>(null)
  const toast = useToast()

  const load = () =>
    api<{ docs: DocSummary[] }>('/api/docs')
      .then((r) => setDocs(r.docs))
      .catch((e) => toast(e.message, 'error'))
  useEffect(() => {
    load()
    document.title = 'SCRIBE'
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const open = (id: string) => (location.hash = `#/doc/${id}`)

  const fromTemplate = async (tid: string) => {
    const t = TEMPLATES.find((x) => x.id === tid)!
    setBusy('Creating document…')
    try {
      const d = t.build()
      open(await createDocument(t.id === 'blank' || t.id === 'blank-a4' ? 'Untitled document' : t.name, d.content, d.meta))
    } catch (e: any) {
      toast(e.message, 'error')
    } finally {
      setBusy('')
    }
  }

  const importFile = async (file: File) => {
    if (file.size > MAX_IMPORT) return toast('That file is larger than 100 MB.', 'error')
    const name = file.name.replace(/\.(docx|html?|txt|md)$/i, '')
    setBusy(`Importing ${file.name}…`)
    try {
      const ext = file.name.split('.').pop()?.toLowerCase()
      if (ext === 'docx' || ext === 'docm' || ext === 'dotx') {
        const bytes = new Uint8Array(await file.arrayBuffer())
        const r = await importDocx(bytes)
        const title = r.doc.meta.metadata.title || name
        const id = await createDocument(title, r.doc.content, r.doc.meta, bytes, file.name)
        if (r.report.length) setReport({ issues: r.report, id })
        else open(id)
      } else if (ext === 'html' || ext === 'htm') {
        const html = sanitizeHtml(await file.text())
        const content = generateJSON(html, schemaExtensions({ collaborative: true })) as PMNode
        open(await createDocument(name, content, emptyMeta()))
      } else if (ext === 'txt' || ext === 'md' || file.type.startsWith('text/')) {
        open(await createDocument(name, textToContent(await file.text()), emptyMeta()))
      } else if (ext === 'doc') {
        toast('Legacy .doc files are not supported. Save the file as .docx in Word or LibreOffice first.', 'error')
      } else toast('Unsupported file type. SCRIBE opens .docx, .html and .txt files.', 'error')
    } catch (e: any) {
      toast(`Could not import ${file.name}: ${e.message}`, 'error')
    } finally {
      setBusy('')
    }
  }

  const remove = async (d: DocSummary) => {
    try {
      await api(`/api/docs/${d.id}`, { method: 'DELETE' })
      setDocs((x) => x?.filter((y) => y.id !== d.id) ?? null)
      toast('Document deleted', 'success')
    } catch (e: any) {
      toast(e.message, 'error')
    }
    setConfirmDelete(null)
  }

  return (
    <div className="sc-home">
      <header className="sc-home-bar">
        <span className="sc-logo">SCRIBE</span>
        <div className="sc-home-actions">
          <button className="sc-btn" onClick={onToggleTheme} aria-label={theme === 'dark' ? 'Use light theme' : 'Use dark theme'} title="Toggle theme">
            {theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
          </button>
          <span className="sc-user" title={user.email}>
            {user.name}
          </span>
          <button className="sc-btn with-label" onClick={onSignOut}>
            <LogOut size={16} />
            <span>Sign out</span>
          </button>
        </div>
      </header>
      <main
        className="sc-home-main"
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => {
          e.preventDefault()
          const f = e.dataTransfer.files[0]
          if (f) importFile(f)
        }}
      >
        <section aria-labelledby="new-h">
          <div className="sc-section-head">
            <h2 id="new-h">Start a document</h2>
            <button className="sc-primary with-icon" onClick={() => fileRef.current?.click()} disabled={!!busy}>
              <Upload size={16} /> Open file…
            </button>
            <input ref={fileRef} type="file" hidden accept=".docx,.docm,.dotx,.html,.htm,.txt,.md,.doc" onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ''; if (f) importFile(f) }} />
          </div>
          <div className="sc-templates">
            {TEMPLATES.map((t) => (
              <button key={t.id} className="sc-template" onClick={() => fromTemplate(t.id)} disabled={!!busy}>
                <div className={`sc-template-thumb t-${t.id}`} aria-hidden="true">
                  <i />
                  <i />
                  <i />
                  <i />
                </div>
                <strong>{t.name}</strong>
                <span>{t.description}</span>
              </button>
            ))}
          </div>
          {busy && (
            <p className="sc-busy" role="status">
              {busy}
            </p>
          )}
        </section>
        <section aria-labelledby="recent-h">
          <h2 id="recent-h">Recent documents</h2>
          {docs === null ? (
            <p className="sc-muted">Loading…</p>
          ) : docs.length === 0 ? (
            <p className="sc-muted">No documents yet. Start from a template above, or drop a .docx file here.</p>
          ) : (
            <ul className="sc-doclist">
              {docs.map((d) => (
                <li key={d.id}>
                  <button className="sc-doc-open" onClick={() => open(d.id)}>
                    <FileText size={18} aria-hidden="true" />
                    <span className="sc-doc-title">{d.title}</span>
                    <span className="sc-doc-meta">
                      {d.role !== 'owner' && (
                        <span className="sc-chip-ui">
                          <Users size={12} /> {d.owner_name} · {d.role}
                        </span>
                      )}
                      {d.has_original && <span className="sc-chip-ui">DOCX</span>}
                      {new Date(d.updated).toLocaleString()}
                    </span>
                  </button>
                  {d.role === 'owner' && (
                    <button className="sc-btn" aria-label={`Delete ${d.title}`} title="Delete" onClick={() => setConfirmDelete(d)}>
                      <Trash2 size={16} />
                    </button>
                  )}
                </li>
              ))}
            </ul>
          )}
        </section>
      </main>
      {report && <CompatReport issues={report.issues} title="Import report" onClose={() => open(report.id)} actionLabel="Open document" />}
      {confirmDelete && (
        <Dialog
          title="Delete document?"
          onClose={() => setConfirmDelete(null)}
          footer={
            <>
              <button onClick={() => setConfirmDelete(null)}>Cancel</button>
              <button className="sc-danger" onClick={() => remove(confirmDelete)}>
                Delete
              </button>
            </>
          }
        >
          <p>
            “{confirmDelete.title}” will be removed for you and everyone it is shared with.
          </p>
        </Dialog>
      )}
    </div>
  )
}
