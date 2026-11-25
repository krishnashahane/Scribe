import { useEffect, useRef, useState } from 'react'
import { FileText, LogIn, LogOut, Moon, Sun, Trash2, Upload, UserRound, Users } from 'lucide-react'
import { api, type DocSummary, type User } from '../api'
import { TEMPLATES } from '../model/templates'
import type { CompatibilityIssue } from '../model/types'
import { Dialog, useToast } from './kit'
import { AccountDialog, CompatReport } from './Dialogs'
import { createFromTemplate, importFile as importToDoc, OPEN_ACCEPT, openDoc } from './docActions'
import type { Theme } from './App'

export function Home({ user, onUser, onSignOut, theme, onToggleTheme }: { user: User; onUser: (u: User) => void; onSignOut: () => void; theme: Theme; onToggleTheme: () => void }) {
  const [account, setAccount] = useState(false)
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

  const open = openDoc

  const fromTemplate = async (tid: string) => {
    setBusy('Creating document…')
    try {
      open(await createFromTemplate(tid))
    } catch (e: any) {
      toast(e.message, 'error')
    } finally {
      setBusy('')
    }
  }

  const importFile = async (file: File) => {
    setBusy(`Opening ${file.name}…`)
    try {
      const r = await importToDoc(file)
      if (r.report.length) setReport({ issues: r.report, id: r.id })
      else open(r.id)
    } catch (e: any) {
      toast(`Could not open ${file.name}: ${e.message}`, 'error')
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
        <a className="sc-logo" href="#/" title="Back to your last document" style={{ textDecoration: 'none' }}>SCRIBE</a>
        <div className="sc-home-actions">
          <button className="sc-btn" onClick={onToggleTheme} aria-label={theme === 'dark' ? 'Use light theme' : 'Use dark theme'} title="Toggle theme">
            {theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
          </button>
          {user.guest ? (
            <>
              <button className="sc-btn with-label" onClick={() => setAccount(true)}>
                <UserRound size={16} />
                <span>Save to account</span>
              </button>
              <a className="sc-btn with-label" href="#/signin">
                <LogIn size={16} />
                <span>Sign in</span>
              </a>
            </>
          ) : (
            <>
              <span className="sc-user" title={user.email}>
                {user.name}
              </span>
              <button className="sc-btn with-label" onClick={onSignOut}>
                <LogOut size={16} />
                <span>Sign out</span>
              </button>
            </>
          )}
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
            <input ref={fileRef} type="file" hidden accept={OPEN_ACCEPT} onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ''; if (f) importFile(f) }} />
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
      {account && <AccountDialog user={user} onUser={onUser} onClose={() => setAccount(false)} />}
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
