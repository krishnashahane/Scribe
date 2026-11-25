import { useEffect, useMemo, useState } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import { api } from '../api'
import { isSafeUrl, schemaExtensions } from '../model/extensions'
import { PAGE_SIZES, type CompatibilityIssue, type DocMeta, type PageSetup, type PMNode, type SectionProps } from '../model/types'
import { Dialog, mod, useToast } from './kit'

export function CompatReport({ issues, title, onClose, actionLabel = 'Close' }: { issues: CompatibilityIssue[]; title: string; onClose: () => void; actionLabel?: string }) {
  const groups: Record<string, CompatibilityIssue[]> = { dropped: [], transformed: [], preserved: [], unsupported: [] }
  for (const i of issues) groups[i.kind].push(i)
  const label: Record<string, string> = {
    dropped: 'Lost — not kept in SCRIBE',
    transformed: 'Converted — kept with possible fidelity changes',
    preserved: 'Preserved — kept verbatim, limited editing',
    unsupported: 'Unsupported',
  }
  return (
    <Dialog title={title} onClose={onClose} wide footer={<button className="sc-primary" onClick={onClose}>{actionLabel}</button>}>
      {!issues.length ? (
        <p>Everything in this file is supported. Nothing was converted or lost.</p>
      ) : (
        <>
          <p className="sc-muted">SCRIBE never silently discards content. Elements it cannot edit are kept unchanged and written back when you export to DOCX. The original file is stored untouched.</p>
          {Object.entries(groups).map(([k, list]) =>
            list.length ? (
              <section key={k} className={`sc-report is-${k}`}>
                <h3>{label[k]}</h3>
                <ul>
                  {list.map((i, n) => (
                    <li key={n}>
                      <strong>{i.feature}</strong> {i.count > 1 && <span className="sc-chip-ui">×{i.count}</span>}
                      <div>{i.detail}</div>
                    </li>
                  ))}
                </ul>
              </section>
            ) : null,
          )}
        </>
      )}
    </Dialog>
  )
}

const INCH = 1440
const MM = 1440 / 25.4

export function PageSetupDialog({ sections, initialSection, onApply, onClose }: { sections: SectionProps[]; initialSection: number; onApply: (index: number | 'all', page: PageSetup, opts: { titlePage: boolean }) => void; onClose: () => void }) {
  const [idx, setIdx] = useState(initialSection)
  const [unit, setUnit] = useState<'in' | 'mm'>(() => (Math.abs(sections[initialSection].page.width - 11906) < 30 ? 'mm' : 'in'))
  const [page, setPage] = useState<PageSetup>({ ...sections[initialSection].page })
  const [titlePage, setTitlePage] = useState(!!sections[initialSection].titlePage)
  const [scope, setScope] = useState<'this' | 'all'>(sections.length > 1 ? 'this' : 'all')
  useEffect(() => {
    setPage({ ...sections[idx].page })
    setTitlePage(!!sections[idx].titlePage)
  }, [idx, sections])
  const f = unit === 'in' ? INCH : MM
  const fmt = (tw: number) => String(Math.round((tw / f) * 100) / 100)
  const set = (k: keyof PageSetup) => (e: React.ChangeEvent<HTMLInputElement>) => {
    const v = parseFloat(e.target.value)
    if (Number.isFinite(v) && v >= 0) setPage((p) => ({ ...p, [k]: Math.round(v * f) }))
  }
  const sizeKey = Object.entries(PAGE_SIZES).find(([, s]) => (Math.abs(s.width - page.width) < 30 && Math.abs(s.height - page.height) < 30) || (Math.abs(s.height - page.width) < 30 && Math.abs(s.width - page.height) < 30))?.[0] ?? 'custom'
  const setSize = (k: string) => {
    if (k === 'custom') return
    const s = PAGE_SIZES[k]
    setPage((p) => (p.orientation === 'landscape' ? { ...p, width: s.height, height: s.width } : { ...p, width: s.width, height: s.height }))
  }
  const setOrientation = (o: 'portrait' | 'landscape') =>
    setPage((p) => (p.orientation === o ? p : { ...p, orientation: o, width: p.height, height: p.width }))
  const valid = page.width > page.marginLeft + page.marginRight + 720 && page.height > page.marginTop + page.marginBottom + 720 && page.width >= 2880 && page.height >= 2880 && page.width <= 31680 && page.height <= 31680
  return (
    <Dialog
      title="Page setup"
      onClose={onClose}
      footer={
        <>
          <button onClick={onClose}>Cancel</button>
          <button className="sc-primary" disabled={!valid} onClick={() => onApply(scope === 'all' ? 'all' : idx, page, { titlePage })}>
            Apply
          </button>
        </>
      }
    >
      <div className="sc-form">
        {sections.length > 1 && (
          <label>
            Section
            <select value={idx} onChange={(e) => setIdx(Number(e.target.value))}>
              {sections.map((_, i) => (
                <option key={i} value={i}>
                  Section {i + 1}
                </option>
              ))}
            </select>
          </label>
        )}
        <label>
          Paper size
          <select value={sizeKey} onChange={(e) => setSize(e.target.value)}>
            {Object.entries(PAGE_SIZES).map(([k, s]) => (
              <option key={k} value={k}>
                {s.label}
              </option>
            ))}
            <option value="custom">Custom</option>
          </select>
        </label>
        <fieldset className="sc-radio-row">
          <legend>Orientation</legend>
          <label>
            <input type="radio" checked={page.orientation !== 'landscape'} onChange={() => setOrientation('portrait')} /> Portrait
          </label>
          <label>
            <input type="radio" checked={page.orientation === 'landscape'} onChange={() => setOrientation('landscape')} /> Landscape
          </label>
        </fieldset>
        <label>
          Units
          <select value={unit} onChange={(e) => setUnit(e.target.value as any)}>
            <option value="in">Inches</option>
            <option value="mm">Millimetres</option>
          </select>
        </label>
        <div className="sc-grid2">
          <label>
            Width
            <input type="number" step="0.01" value={fmt(page.width)} onChange={set('width')} />
          </label>
          <label>
            Height
            <input type="number" step="0.01" value={fmt(page.height)} onChange={set('height')} />
          </label>
          <label>
            Top margin
            <input type="number" step="0.01" value={fmt(page.marginTop)} onChange={set('marginTop')} />
          </label>
          <label>
            Bottom margin
            <input type="number" step="0.01" value={fmt(page.marginBottom)} onChange={set('marginBottom')} />
          </label>
          <label>
            Left margin
            <input type="number" step="0.01" value={fmt(page.marginLeft)} onChange={set('marginLeft')} />
          </label>
          <label>
            Right margin
            <input type="number" step="0.01" value={fmt(page.marginRight)} onChange={set('marginRight')} />
          </label>
          <label>
            Header distance
            <input type="number" step="0.01" value={fmt(page.marginHeader)} onChange={set('marginHeader')} />
          </label>
          <label>
            Footer distance
            <input type="number" step="0.01" value={fmt(page.marginFooter)} onChange={set('marginFooter')} />
          </label>
        </div>
        <label className="sc-check">
          <input type="checkbox" checked={titlePage} onChange={(e) => setTitlePage(e.target.checked)} /> Different first page header/footer
        </label>
        {sections.length > 1 && (
          <fieldset className="sc-radio-row">
            <legend>Apply to</legend>
            <label>
              <input type="radio" checked={scope === 'this'} onChange={() => setScope('this')} /> This section
            </label>
            <label>
              <input type="radio" checked={scope === 'all'} onChange={() => setScope('all')} /> Whole document
            </label>
          </fieldset>
        )}
        {!valid && <p className="sc-error">Margins leave no room for text, or the size is outside 2–22 inches.</p>}
      </div>
    </Dialog>
  )
}

/** Small rich-text editor for one header or footer. */
export function HeaderFooterDialog(props: {
  kind: 'header' | 'footer'
  first: boolean
  content: PMNode | null
  onSave: (content: PMNode | null) => void
  onClose: () => void
}) {
  const editor = useEditor({
    extensions: schemaExtensions(),
    content: props.content ?? { type: 'doc', content: [{ type: 'paragraph' }] },
    editorProps: { attributes: { class: 'sc-hf-editor', 'aria-label': `${props.kind} content` } },
  })
  const insertField = (instr: string) => editor?.chain().focus().insertContent({ type: 'field', attrs: { instr, result: '1' } }).run()
  const name = `${props.first ? 'First-page ' : ''}${props.kind}`
  return (
    <Dialog
      title={`Edit ${name}`}
      onClose={props.onClose}
      wide
      footer={
        <>
          <button className="sc-danger-quiet" onClick={() => props.onSave(null)}>
            Remove {props.kind}
          </button>
          <span className="sc-spacer" />
          <button onClick={props.onClose}>Cancel</button>
          <button className="sc-primary" onClick={() => editor && props.onSave(editor.getJSON() as PMNode)}>
            Save
          </button>
        </>
      }
    >
      <div className="sc-hf-tools" role="toolbar" aria-label={`${props.kind} tools`}>
        <button onClick={() => editor?.chain().focus().toggleBold().run()}>Bold</button>
        <button onClick={() => editor?.chain().focus().toggleItalic().run()}>Italic</button>
        <button onClick={() => editor?.chain().focus().setTextAlign('left').run()}>Left</button>
        <button onClick={() => editor?.chain().focus().setTextAlign('center').run()}>Center</button>
        <button onClick={() => editor?.chain().focus().setTextAlign('right').run()}>Right</button>
        <button onClick={() => insertField('PAGE')}>Insert page number</button>
        <button onClick={() => insertField('NUMPAGES')}>Insert page count</button>
      </div>
      <div className="sc-hf-box">
        <EditorContent editor={editor} />
      </div>
      <p className="sc-muted">Page number fields update on every page and are exported as real Word fields.</p>
    </Dialog>
  )
}

export function LinkDialog({ initial, bookmarks, onApply, onRemove, onClose }: { initial: string; bookmarks: string[]; onApply: (href: string) => void; onRemove: () => void; onClose: () => void }) {
  const [href, setHref] = useState(initial)
  const normalized = /^[a-z]+:|^#|^\//i.test(href) ? href : href.includes('@') && !href.includes('/') ? `mailto:${href}` : `https://${href}`
  const ok = href.trim() !== '' && isSafeUrl(normalized)
  return (
    <Dialog
      title="Link"
      onClose={onClose}
      footer={
        <>
          {initial && <button className="sc-danger-quiet" onClick={onRemove}>Remove link</button>}
          <span className="sc-spacer" />
          <button onClick={onClose}>Cancel</button>
          <button className="sc-primary" disabled={!ok} onClick={() => onApply(normalized)}>
            Apply
          </button>
        </>
      }
    >
      <form className="sc-form" onSubmit={(e) => { e.preventDefault(); if (ok) onApply(normalized) }}>
        <label>
          Address
          <input value={href} onChange={(e) => setHref(e.target.value)} placeholder="https://example.com or #bookmark" />
        </label>
        {bookmarks.length > 0 && (
          <label>
            Or link to a bookmark
            <select value={href.startsWith('#') ? href : ''} onChange={(e) => setHref(e.target.value)}>
              <option value="">—</option>
              {bookmarks.map((b) => (
                <option key={b} value={`#${b}`}>
                  {b}
                </option>
              ))}
            </select>
          </label>
        )}
        {href && !ok && <p className="sc-error">Only web, email, telephone and bookmark links are allowed.</p>}
      </form>
    </Dialog>
  )
}

export function TextPrompt({ title, label, initial = '', onSubmit, onClose, validate }: { title: string; label: string; initial?: string; onSubmit: (v: string) => void; onClose: () => void; validate?: (v: string) => string | null }) {
  const [v, setV] = useState(initial)
  const err = validate?.(v) ?? null
  return (
    <Dialog title={title} onClose={onClose} footer={<><button onClick={onClose}>Cancel</button><button className="sc-primary" disabled={!!err || !v.trim()} onClick={() => onSubmit(v.trim())}>OK</button></>}>
      <form className="sc-form" onSubmit={(e) => { e.preventDefault(); if (!err && v.trim()) onSubmit(v.trim()) }}>
        <label>
          {label}
          <input value={v} onChange={(e) => setV(e.target.value)} />
        </label>
        {err && v && <p className="sc-error">{err}</p>}
      </form>
    </Dialog>
  )
}

export function ShareDialog({ docId, isOwner, onClose }: { docId: string; isOwner: boolean; onClose: () => void }) {
  const [shares, setShares] = useState<{ id: string; email: string; name: string; role: string }[]>([])
  const [email, setEmail] = useState('')
  const [role, setRole] = useState('editor')
  const [error, setError] = useState('')
  const toast = useToast()
  const load = () => api<{ shares: any[] }>(`/api/docs/${docId}/shares`).then((r) => setShares(r.shares)).catch((e) => setError(e.message))
  useEffect(() => {
    load()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])
  const add = async (e: React.FormEvent) => {
    e.preventDefault()
    setError('')
    try {
      await api(`/api/docs/${docId}/shares`, { method: 'POST', body: { email, role } })
      setEmail('')
      toast('Access updated', 'success')
      load()
    } catch (err: any) {
      setError(err.message)
    }
  }
  const remove = async (id: string) => {
    await api(`/api/docs/${docId}/shares/${id}`, { method: 'DELETE' }).catch((e) => setError(e.message))
    load()
  }
  return (
    <Dialog title="Share" onClose={onClose} footer={<button className="sc-primary" onClick={onClose}>Done</button>}>
      {isOwner ? (
        <form className="sc-share-form" onSubmit={add}>
          <input type="email" required placeholder="Email of a SCRIBE user" value={email} onChange={(e) => setEmail(e.target.value)} aria-label="Email" />
          <select value={role} onChange={(e) => setRole(e.target.value)} aria-label="Access level">
            <option value="editor">Can edit</option>
            <option value="viewer">Can view</option>
          </select>
          <button className="sc-primary" type="submit">
            Share
          </button>
        </form>
      ) : (
        <p className="sc-muted">Only the owner can change who has access.</p>
      )}
      {error && <p className="sc-error" role="alert">{error}</p>}
      <ul className="sc-share-list">
        {shares.length === 0 && <li className="sc-muted">Not shared with anyone yet.</li>}
        {shares.map((s) => (
          <li key={s.id}>
            <span>
              <strong>{s.name}</strong> <span className="sc-muted">{s.email}</span>
            </span>
            <span className="sc-chip-ui">{s.role === 'editor' ? 'Can edit' : 'Can view'}</span>
            {isOwner && (
              <button className="sc-link" onClick={() => remove(s.id)}>
                Remove
              </button>
            )}
          </li>
        ))}
      </ul>
    </Dialog>
  )
}

export function PropertiesDialog({ title, meta, onSave, onClose, readOnly }: { title: string; meta: DocMeta; onSave: (title: string, m: DocMeta['metadata']) => void; onClose: () => void; readOnly: boolean }) {
  const [t, setT] = useState(title)
  const [m, setM] = useState({ ...meta.metadata })
  const field = (k: keyof DocMeta['metadata'], label: string) => (
    <label>
      {label}
      <input value={m[k] ?? ''} readOnly={readOnly} onChange={(e) => setM({ ...m, [k]: e.target.value })} />
    </label>
  )
  return (
    <Dialog title="Document properties" onClose={onClose} footer={<><button onClick={onClose}>Cancel</button>{!readOnly && <button className="sc-primary" onClick={() => onSave(t, m)}>Save</button>}</>}>
      <div className="sc-form">
        <label>
          Title
          <input value={t} readOnly={readOnly} onChange={(e) => setT(e.target.value)} />
        </label>
        {field('creator', 'Author')}
        {field('subject', 'Subject')}
        {field('keywords', 'Keywords')}
        {field('description', 'Description')}
        {meta.metadata.created && <p className="sc-muted">Created {new Date(meta.metadata.created).toLocaleString()}</p>}
      </div>
    </Dialog>
  )
}

export function WordCountDialog({ text, onClose }: { text: string; onClose: () => void }) {
  const stats = useMemo(() => {
    const words = text.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu)?.length ?? 0
    const cjk = text.match(/[぀-ヿ㐀-鿿가-힯]/g)?.length ?? 0
    return { words: words + cjk, chars: [...text.replace(/\n/g, '')].length, charsNoSpace: [...text.replace(/\s/g, '')].length, paragraphs: text.split(/\n+/).filter((l) => l.trim()).length }
  }, [text])
  return (
    <Dialog title="Word count" onClose={onClose} footer={<button className="sc-primary" onClick={onClose}>Close</button>}>
      <dl className="sc-stats">
        <dt>Words</dt>
        <dd>{stats.words.toLocaleString()}</dd>
        <dt>Characters</dt>
        <dd>{stats.chars.toLocaleString()}</dd>
        <dt>Characters (no spaces)</dt>
        <dd>{stats.charsNoSpace.toLocaleString()}</dd>
        <dt>Paragraphs</dt>
        <dd>{stats.paragraphs.toLocaleString()}</dd>
      </dl>
    </Dialog>
  )
}

export const SHORTCUTS: [string, string][] = [
  [`${mod}B / ${mod}I / ${mod}U`, 'Bold / italic / underline'],
  [`${mod}Shift+X`, 'Strikethrough'],
  [`${mod}Z / ${mod}Shift+Z`, 'Undo / redo'],
  [`${mod}Alt+1…6`, 'Heading 1–6'],
  [`${mod}Alt+0`, 'Normal text'],
  [`${mod}Shift+7 / 8`, 'Numbered / bulleted list'],
  ['Tab / Shift+Tab', 'Indent / outdent list item'],
  [`${mod}Shift+L / E / R / J`, 'Align left / center / right / justify'],
  [`${mod}K`, 'Insert link'],
  [`${mod}Enter`, 'Page break'],
  ['Shift+Enter', 'Line break'],
  [`${mod}F`, 'Find'],
  [`${mod}H`, 'Find and replace'],
  [`${mod}Alt+M`, 'Add comment'],
  [`${mod}S`, 'Save a named version'],
  [`${mod}P`, 'Print / export PDF'],
  [`${mod}/`, 'This list'],
]

export function ShortcutsDialog({ onClose }: { onClose: () => void }) {
  return (
    <Dialog title="Keyboard shortcuts" onClose={onClose} footer={<button className="sc-primary" onClick={onClose}>Close</button>}>
      <table className="sc-shortcuts">
        <tbody>
          {SHORTCUTS.map(([k, v]) => (
            <tr key={k}>
              <td>
                <kbd>{k}</kbd>
              </td>
              <td>{v}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </Dialog>
  )
}

/** Turn the local no-login identity into a full account (documents are kept), or sign in elsewhere. */
export function AccountDialog({ user, onUser, onClose }: { user: { name: string; email: string; guest?: boolean }; onUser: (u: any) => void; onClose: () => void }) {
  const [name, setName] = useState(user.guest ? '' : user.name)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const toast = useToast()
  const save = async (e: React.FormEvent) => {
    e.preventDefault()
    setError('')
    try {
      const r = await api<{ user: any }>('/api/auth/upgrade', { method: 'POST', body: { name, email, password } })
      onUser(r.user)
      toast('Account created. Your documents are saved to it.', 'success')
      onClose()
    } catch (err: any) {
      setError(err.message)
    }
  }
  if (!user.guest)
    return (
      <Dialog title="Account" onClose={onClose} footer={<button className="sc-primary" onClick={onClose}>Close</button>}>
        <p>
          Signed in as <strong>{user.name}</strong> ({user.email}).
        </p>
      </Dialog>
    )
  return (
    <Dialog title="Save your work to an account" onClose={onClose} footer={<><button onClick={() => (location.hash = '#/signin')}>Sign in instead</button><span className="sc-spacer" /><button onClick={onClose}>Not now</button><button className="sc-primary" type="submit" form="sc-upgrade">Create account</button></>}>
      <p className="sc-muted">You're using SCRIBE without an account. Your documents are stored on this computer's SCRIBE server and tied to this browser. Create an account to keep them if browser data is cleared, open them on other devices, and let people share documents with you.</p>
      <form id="sc-upgrade" className="sc-form" onSubmit={save}>
        <label>
          Name
          <input value={name} onChange={(e) => setName(e.target.value)} autoComplete="name" />
        </label>
        <label>
          Email
          <input type="email" required value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="email" />
        </label>
        <label>
          Password
          <input type="password" required minLength={8} value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="new-password" />
        </label>
        {error && <p className="sc-error" role="alert">{error}</p>}
      </form>
    </Dialog>
  )
}
