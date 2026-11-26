import { useEffect, useEffectEvent, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { useEditorState } from '@tiptap/react'
import { generateHTML } from '@tiptap/core'
import { ChevronLeft, ChevronRight, FileDown, FilePlus2, Footprints, Plus, SpellCheck2, Trash2, Upload, Users } from 'lucide-react'
import { TextSelection } from '@tiptap/pm/state'
import { Dialog, useToast } from './kit'
import type { CitationStyle, DocMeta, MailMergeData, PMNode, Source, SourcePerson, SourceType, StyleDef } from '../model/types'
import { SOURCE_TYPES, formatCitation, formatEntry, newTag } from '../features/citations'
import { fillRecord, mergeFieldInstr, recipientsFromCsv } from '../features/merge'
import { RULE_NAMES, proofState, setProofing, type DocIssue } from '../features/grammar'
import { displayName, headingLevelOfStyle, styleId } from '../features/styles'
import { schemaExtensions } from '../model/extensions'

// ---------------------------------------------------------------------------
// Sources

const blankSource = (): Source => ({ tag: '', type: 'Book', authors: [{ last: '', first: '' }], title: '', year: '' })

export function SourcesDialog({ sources, style, cited, startNew, onSave, onClose }: { sources: Source[]; style: CitationStyle; cited: Set<string>; startNew?: boolean; onSave: (s: Source[]) => void; onClose: () => void }) {
  const [list, setList] = useState<Source[]>(() => [...sources.map((s) => ({ ...s, authors: s.authors.map((a) => ({ ...a })) })), ...(startNew ? [blankSource()] : [])])
  const [sel, setSel] = useState<number>(startNew ? sources.length : sources.length ? 0 : -1)
  const cur = sel >= 0 ? list[sel] : null
  const update = (p: Partial<Source>) => setList(list.map((s, i) => (i === sel ? { ...s, ...p } : s)))
  const setAuthor = (k: number, p: Partial<SourcePerson>) => update({ authors: cur!.authors.map((a, i) => (i === k ? { ...a, ...p } : a)) })
  const add = () => {
    setList([...list, blankSource()])
    setSel(list.length)
  }
  const remove = () => {
    setList(list.filter((_, i) => i !== sel))
    setSel(Math.min(sel, list.length - 2))
  }
  const finalize = () => {
    const tags: string[] = []
    const out = list
      .filter((s) => s.title.trim() || s.authors.some((a) => a.last.trim()))
      .map((s) => {
        const authors = s.authors.filter((a) => a.last.trim() || a.first?.trim())
        const tag = s.tag && !tags.includes(s.tag) ? s.tag : newTag({ ...s, authors }, tags)
        tags.push(tag)
        return { ...s, tag, authors }
      })
    onSave(out)
  }
  const isContainer = cur?.type === 'JournalArticle' || cur?.type === 'InternetSite'
  return (
    <Dialog title="Sources" onClose={onClose} wide footer={<><button onClick={onClose}>Cancel</button><button className="sc-primary" onClick={finalize}>Save sources</button></>}>
      <div className="sc-sources">
        <div className="sc-source-list">
          <button className="sc-add-row" onClick={add}>
            <Plus size={14} /> New source
          </button>
          <ul role="listbox" aria-label="Sources">
            {list.map((s, i) => (
              <li key={i} role="option" aria-selected={i === sel}>
                <button className={i === sel ? 'is-active' : ''} onClick={() => setSel(i)}>
                  <strong>{s.authors[0]?.last || s.title || 'Untitled source'}</strong>
                  <span>
                    {s.year || 'n.d.'} · {SOURCE_TYPES[s.type]}
                    {s.tag && cited.has(s.tag) ? ' · cited' : ''}
                  </span>
                </button>
              </li>
            ))}
          </ul>
          {!list.length && <p className="sc-muted">No sources yet.</p>}
        </div>
        {cur ? (
          <div className="sc-form sc-source-form">
            <div className="sc-form-row">
              <label>
                Type
                <select value={cur.type} onChange={(e) => update({ type: e.target.value as SourceType })}>
                  {Object.entries(SOURCE_TYPES).map(([k, v]) => (
                    <option key={k} value={k}>
                      {v}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Year
                <input value={cur.year ?? ''} onChange={(e) => update({ year: e.target.value })} inputMode="numeric" />
              </label>
            </div>
            <fieldset className="sc-fieldset">
              <legend>Authors</legend>
              {cur.authors.map((a, k) => (
                <div className="sc-form-row" key={k}>
                  <input aria-label={`Author ${k + 1} last name`} placeholder="Last name or organisation" value={a.last} onChange={(e) => setAuthor(k, { last: e.target.value })} />
                  <input aria-label={`Author ${k + 1} first name`} placeholder="First name" value={a.first ?? ''} onChange={(e) => setAuthor(k, { first: e.target.value })} />
                  <button className="sc-icon-mini" aria-label={`Remove author ${k + 1}`} disabled={cur.authors.length === 1} onClick={() => update({ authors: cur.authors.filter((_, i) => i !== k) })}>
                    <Trash2 size={13} />
                  </button>
                </div>
              ))}
              <button className="sc-link-btn" onClick={() => update({ authors: [...cur.authors, { last: '', first: '' }] })}>
                + Add author
              </button>
            </fieldset>
            <label>
              Title
              <input value={cur.title} onChange={(e) => update({ title: e.target.value })} />
            </label>
            {isContainer && (
              <label>
                {cur.type === 'JournalArticle' ? 'Journal' : 'Website name'}
                <input value={cur.journal ?? ''} onChange={(e) => update({ journal: e.target.value })} />
              </label>
            )}
            {cur.type === 'JournalArticle' ? (
              <div className="sc-form-row">
                <label>
                  Volume
                  <input value={cur.volume ?? ''} onChange={(e) => update({ volume: e.target.value })} />
                </label>
                <label>
                  Issue
                  <input value={cur.issue ?? ''} onChange={(e) => update({ issue: e.target.value })} />
                </label>
                <label>
                  Pages
                  <input value={cur.pages ?? ''} onChange={(e) => update({ pages: e.target.value })} />
                </label>
              </div>
            ) : (
              <div className="sc-form-row">
                <label>
                  Publisher
                  <input value={cur.publisher ?? ''} onChange={(e) => update({ publisher: e.target.value })} />
                </label>
                <label>
                  City
                  <input value={cur.city ?? ''} onChange={(e) => update({ city: e.target.value })} />
                </label>
              </div>
            )}
            <div className="sc-form-row">
              <label>
                URL
                <input value={cur.url ?? ''} onChange={(e) => update({ url: e.target.value })} type="url" placeholder="https://" />
              </label>
              {cur.type === 'InternetSite' && (
                <label>
                  Accessed
                  <input type="date" value={cur.accessed ?? ''} onChange={(e) => update({ accessed: e.target.value })} />
                </label>
              )}
            </div>
            <div className="sc-source-preview">
              <span className="sc-muted">{style} preview</span>
              <p>
                {formatEntry(cur, style).map((s, i) => (s.italic ? <em key={i}>{s.text}</em> : <span key={i}>{s.text}</span>))}
              </p>
              <p className="sc-muted">In text: {formatCitation([cur], style)}</p>
            </div>
            <button className="sc-danger-quiet" onClick={remove} disabled={!!cur.tag && cited.has(cur.tag)} title={cur.tag && cited.has(cur.tag) ? 'This source is cited in the document' : undefined}>
              Delete source
            </button>
          </div>
        ) : (
          <div className="sc-empty-state">
            <p>Add the books, articles and web pages you cite. SCRIBE formats citations and the bibliography for you.</p>
            <button className="sc-primary" onClick={add}>
              Add a source
            </button>
          </div>
        )}
      </div>
    </Dialog>
  )
}

export function CitationDialog({ sources, style, onInsert, onManage, onClose }: { sources: Source[]; style: CitationStyle; onInsert: (tags: string[], page: string) => void; onManage: () => void; onClose: () => void }) {
  const [picked, setPicked] = useState<string[]>([])
  const [page, setPage] = useState('')
  const [q, setQ] = useState('')
  const shown = sources.filter((s) => `${s.title} ${s.authors.map((a) => `${a.first} ${a.last}`).join(' ')} ${s.year}`.toLowerCase().includes(q.toLowerCase()))
  const chosen = picked.map((t) => sources.find((s) => s.tag === t)!).filter(Boolean)
  return (
    <Dialog title="Insert citation" onClose={onClose} wide footer={<><button onClick={onManage}>Manage sources…</button><span className="sc-spacer" /><button onClick={onClose}>Cancel</button><button className="sc-primary" disabled={!picked.length} onClick={() => onInsert(picked, page)}>Insert</button></>}>
      {!sources.length ? (
        <div className="sc-empty-state">
          <p>There are no sources in this document yet.</p>
          <button className="sc-primary" onClick={onManage}>
            Add a source
          </button>
        </div>
      ) : (
        <div className="sc-form">
          <input placeholder="Search sources" value={q} onChange={(e) => setQ(e.target.value)} aria-label="Search sources" autoFocus />
          <ul className="sc-pick-list">
            {shown.map((s) => (
              <li key={s.tag}>
                <label className="sc-check">
                  <input type="checkbox" checked={picked.includes(s.tag)} onChange={(e) => setPicked(e.target.checked ? [...picked, s.tag] : picked.filter((t) => t !== s.tag))} />
                  <span>
                    {formatEntry(s, style).map((g, i) => (g.italic ? <em key={i}>{g.text}</em> : <span key={i}>{g.text}</span>))}
                  </span>
                </label>
              </li>
            ))}
          </ul>
          <label>
            Page or location (optional)
            <input value={page} onChange={(e) => setPage(e.target.value)} placeholder="e.g. 42 or 10-12" />
          </label>
          {chosen.length > 0 && <p className="sc-muted">Will insert: {formatCitation(chosen, style, page)}</p>}
        </div>
      )}
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// Mail merge

export function MailMergePanel({ editor, meta, setMeta, readOnly, onMergeToDoc, onMergeToDocx }: { editor: Editor; meta: DocMeta; setMeta: (f: (m: DocMeta) => DocMeta) => void; readOnly: boolean; onMergeToDoc: () => void; onMergeToDocx: () => void }) {
  const toast = useToast()
  const file = useRef<HTMLInputElement>(null)
  const data = meta.mailMerge
  const [rec, setRec] = useState(0)
  const [preview, setPreview] = useState(false)
  const fieldsInDoc = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      let n = 0
      e.state.doc.descendants((x) => {
        if (x.type.name === 'field' && /^\s*MERGEFIELD\b/i.test(String(x.attrs.instr))) n++
      })
      return n
    },
  })
  const load = async (f: File) => {
    if (f.size > 5 * 1024 * 1024) return toast('The CSV file is larger than 5 MB.', 'error')
    try {
      const d = recipientsFromCsv(await f.text(), f.name)
      setMeta((m) => ({ ...m, mailMerge: d }))
      setRec(0)
      toast(`${d.records.length.toLocaleString()} recipients loaded.`, 'success')
    } catch (e: any) {
      toast(e.message, 'error')
    }
  }
  const insertField = (name: string) => editor.chain().focus().insertContent({ type: 'field', attrs: { instr: mergeFieldInstr(name), result: `«${name}»` } }).run()
  const html = useMemo(() => {
    if (!preview || !data?.records.length) return ''
    const content = fillRecord(editor.getJSON() as PMNode, data.fields, data.records[Math.min(rec, data.records.length - 1)])
    try {
      return generateHTML(content as any, schemaExtensions())
    } catch {
      return ''
    }
  }, [preview, data, rec, editor])
  return (
    <div className="sc-panel-body">
      <h3>
        <Users size={16} aria-hidden="true" /> Mail merge
      </h3>
      <ol className="sc-steps">
        <li>
          <strong>Recipients</strong>
          {data ? (
            <p className="sc-muted">
              {data.name}: {data.records.length.toLocaleString()} recipients, {data.fields.length} fields.
            </p>
          ) : (
            <p className="sc-muted">Load a CSV file whose first row holds the field names (Name, Email, City…).</p>
          )}
          <div className="sc-inline-actions">
            <button disabled={readOnly} onClick={() => file.current?.click()}>
              <Upload size={14} /> {data ? 'Replace list' : 'Load CSV'}
            </button>
            {data && (
              <button disabled={readOnly} onClick={() => setMeta((m) => ({ ...m, mailMerge: undefined }))}>
                Remove
              </button>
            )}
          </div>
          <input ref={file} type="file" accept=".csv,.tsv,text/csv,text/plain" hidden onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ''; if (f) load(f) }} />
        </li>
        <li>
          <strong>Insert fields</strong>
          {data ? (
            <div className="sc-chips-row">
              {data.fields.map((f) => (
                <button key={f} className="sc-chip-btn" disabled={readOnly} onMouseDown={(e) => e.preventDefault() /* keep the cursor in the document */} onClick={() => insertField(f)}>
                  «{f}»
                </button>
              ))}
            </div>
          ) : (
            <p className="sc-muted">Fields appear after you load recipients.</p>
          )}
          <p className="sc-muted">{fieldsInDoc} merge field{fieldsInDoc === 1 ? '' : 's'} in the document.</p>
        </li>
        <li>
          <strong>Preview</strong>
          <div className="sc-inline-actions">
            <button disabled={!data?.records.length} aria-pressed={preview} className={preview ? 'is-active' : ''} onClick={() => setPreview(!preview)}>
              {preview ? 'Hide preview' : 'Preview results'}
            </button>
            {preview && data && (
              <span className="sc-pager">
                <button aria-label="Previous recipient" disabled={rec <= 0} onClick={() => setRec(rec - 1)}>
                  <ChevronLeft size={14} />
                </button>
                <span>
                  {rec + 1} / {data.records.length}
                </span>
                <button aria-label="Next recipient" disabled={rec >= data.records.length - 1} onClick={() => setRec(rec + 1)}>
                  <ChevronRight size={14} />
                </button>
              </span>
            )}
          </div>
          {preview && html && <div className="sc-merge-preview sc-doc" dangerouslySetInnerHTML={{ __html: html }} />}
        </li>
        <li>
          <strong>Finish</strong>
          <div className="sc-inline-actions">
            <button className="sc-primary" disabled={!data?.records.length || !fieldsInDoc} onClick={onMergeToDoc}>
              <FilePlus2 size={14} /> New document
            </button>
            <button disabled={!data?.records.length || !fieldsInDoc} onClick={onMergeToDocx}>
              <FileDown size={14} /> DOCX
            </button>
          </div>
          <p className="sc-muted">Each recipient starts on a new page.</p>
        </li>
      </ol>
    </div>
  )
}

// ---------------------------------------------------------------------------
// Proofing

export function ProofingPanel({ editor, readOnly }: { editor: Editor; readOnly: boolean }) {
  const s = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      const st = proofState(e.state)
      return { enabled: !!st?.enabled, issues: st?.issues ?? [] }
    },
    equalityFn: (a, b) => !!a && !!b && a.enabled === b.enabled && a.issues.length === b.issues.length && a.issues.every((x, i) => x.id === b.issues[i].id && x.from === b.issues[i].from),
  })
  const go = (i: DocIssue) => {
    editor.chain().focus().setTextSelection({ from: i.from, to: i.to }).run()
    ;(editor.view.domAtPos(i.from).node as HTMLElement)?.parentElement?.scrollIntoView({ block: 'center', behavior: 'smooth' })
  }
  const fix = (i: DocIssue) => {
    if (i.fix === undefined) return
    const cur = editor.state.doc.textBetween(i.from, i.to, '￼', '￼')
    if (cur.includes('￼')) return
    editor.chain().focus().insertContentAt({ from: i.from, to: i.to }, i.fix ? i.fix : '').run()
  }
  const counts = { grammar: 0, punctuation: 0, style: 0 }
  for (const i of s.issues) counts[i.kind]++
  return (
    <div className="sc-panel-body">
      <h3>
        <SpellCheck2 size={16} aria-hidden="true" /> Proofing
      </h3>
      <label className="sc-switch">
        <input type="checkbox" checked={s.enabled} onChange={(e) => setProofing(editor.view, { enable: e.target.checked })} />
        <span>Check grammar and style</span>
      </label>
      <p className="sc-muted">Runs on this device; nothing is sent anywhere. Spelling uses your browser’s spell checker.</p>
      {s.enabled && (
        <>
          <div className="sc-proof-summary">
            <span className="is-grammar">{counts.grammar} grammar</span>
            <span className="is-punctuation">{counts.punctuation} punctuation</span>
            <span className="is-style">{counts.style} style</span>
          </div>
          {!s.issues.length ? (
            <p className="sc-empty-note">No issues found.</p>
          ) : (
            <ul className="sc-proof-list">
              {s.issues.slice(0, 200).map((i) => (
                <li key={i.id + i.from} className={`is-${i.kind}`}>
                  <button className="sc-proof-item" onClick={() => go(i)}>
                    <strong>{RULE_NAMES[i.rule] ?? 'Suggestion'}</strong>
                    <span>{i.message}</span>
                    <q>{editor.state.doc.textBetween(i.from, i.to, ' ').slice(0, 80)}</q>
                  </button>
                  <div className="sc-comment-actions">
                    {i.fix !== undefined && (
                      <button className="sc-primary" disabled={readOnly} onClick={() => fix(i)}>
                        {i.fix ? `Change to “${i.fix.trim() || '␣'}”` : 'Remove'}
                      </button>
                    )}
                    <button onClick={() => setProofing(editor.view, { ignore: i.id })}>Ignore</button>
                    <button onClick={() => setProofing(editor.view, { disableRule: i.rule })}>Turn off rule</button>
                  </div>
                </li>
              ))}
            </ul>
          )}
          {s.issues.length > 200 && <p className="sc-muted">Showing the first 200 of {s.issues.length} issues.</p>}
        </>
      )}
    </div>
  )
}

// ---------------------------------------------------------------------------
// Styles gallery and Modify Style

export function StylesGallery({ styles, current, disabled, onApply, onModify, onNew }: { styles: StyleDef[]; current: string; disabled: boolean; onApply: (s: StyleDef) => void; onModify: (s: StyleDef) => void; onNew: () => void }) {
  const [open, setOpen] = useState(false)
  const shown = open ? styles : styles.slice(0, 6)
  return (
    <div className={`sc-gallery ${open ? 'is-open' : ''}`} role="listbox" aria-label="Styles">
      <div className="sc-gallery-items">
        {shown.map((s) => (
          <div key={s.id} className={`sc-gallery-item ${current === s.id ? 'is-active' : ''}`} role="option" aria-selected={current === s.id}>
            <button disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={() => onApply(s)} title={`Apply ${displayName(s)}`}>
              <span className="sc-gallery-sample" style={sampleStyle(s)}>
                AaBbCc
              </span>
              <span className="sc-gallery-name">{displayName(s)}</span>
            </button>
            <button className="sc-gallery-edit" disabled={disabled} aria-label={`Modify ${displayName(s)}`} title="Modify style" onMouseDown={(e) => e.preventDefault()} onClick={() => onModify(s)}>
              ⋯
            </button>
          </div>
        ))}
      </div>
      <div className="sc-gallery-side">
        <button onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)} aria-expanded={open} aria-label={open ? 'Show fewer styles' : 'Show all styles'} title={open ? 'Fewer styles' : 'All styles'}>
          {open ? '▴' : '▾'}
        </button>
        <button disabled={disabled} onMouseDown={(e) => e.preventDefault()} onClick={onNew} aria-label="New style" title="New style">
          +
        </button>
      </div>
    </div>
  )
}

function sampleStyle(s: StyleDef): React.CSSProperties {
  const lvl = headingLevelOfStyle(s)
  return {
    fontFamily: s.fontFamily ? `"${s.fontFamily}", sans-serif` : undefined,
    fontWeight: s.bold ? 700 : lvl ? 700 : 400,
    fontStyle: s.italic ? 'italic' : undefined,
    textDecoration: s.underline ? 'underline' : undefined,
    color: s.color,
    fontSize: s.fontSizePt ? `${Math.min(19, Math.max(11, s.fontSizePt * 0.85))}px` : '12px',
  }
}

const FONT_CHOICES = ['Calibri', 'Aptos', 'Arial', 'Cambria', 'Georgia', 'Garamond', 'Helvetica', 'Times New Roman', 'Verdana', 'Courier New']

export function ModifyStyleDialog({ style, styles, isNew, onFromSelection, onSave, onDelete, onClose }: { style: StyleDef; styles: StyleDef[]; isNew: boolean; onFromSelection: () => Partial<StyleDef>; onSave: (s: StyleDef) => void; onDelete?: () => void; onClose: () => void }) {
  const [s, setS] = useState<StyleDef>({ ...style })
  const set = (p: Partial<StyleDef>) => setS((x) => ({ ...x, ...p }))
  const taken = new Set(styles.map((x) => x.id))
  const nameTaken = isNew && styles.some((x) => displayName(x).toLowerCase() === s.name.trim().toLowerCase())
  const pt = (tw?: number) => (tw == null ? '' : String(Math.round((tw / 20) * 10) / 10))
  const save = () => {
    const out: StyleDef = { ...s, name: s.name.trim() }
    if (isNew) {
      out.id = styleId(out.name, taken)
      out.custom = true
      out.type = 'paragraph'
    } else out.modified = true
    onSave(out)
  }
  return (
    <Dialog
      title={isNew ? 'New style' : `Modify style: ${displayName(style)}`}
      onClose={onClose}
      wide
      footer={
        <>
          {onDelete && (
            <button className="sc-danger-quiet" onClick={onDelete}>
              Delete style
            </button>
          )}
          <button onClick={() => set(onFromSelection())}>Match selection</button>
          <span className="sc-spacer" />
          <button onClick={onClose}>Cancel</button>
          <button className="sc-primary" disabled={!s.name.trim() || nameTaken} onClick={save}>
            {isNew ? 'Create' : 'Save'}
          </button>
        </>
      }
    >
      <div className="sc-style-editor">
        <div className="sc-form">
          <div className="sc-form-row">
            <label>
              Name
              <input value={isNew ? s.name : displayName(s)} readOnly={!isNew && !s.custom} onChange={(e) => set({ name: e.target.value })} autoFocus={isNew} />
            </label>
            <label>
              Based on
              <select value={s.basedOn ?? ''} onChange={(e) => set({ basedOn: e.target.value || undefined })} disabled={!isNew && !s.custom}>
                <option value="">(none)</option>
                {styles.filter((x) => x.id !== s.id).map((x) => (
                  <option key={x.id} value={x.id}>
                    {displayName(x)}
                  </option>
                ))}
              </select>
            </label>
          </div>
          {nameTaken && <p className="sc-error">A style with this name already exists.</p>}
          <div className="sc-form-row">
            <label>
              Font
              <select value={s.fontFamily ?? ''} onChange={(e) => set({ fontFamily: e.target.value || undefined })}>
                <option value="">(inherit)</option>
                {[...new Set([...(s.fontFamily ? [s.fontFamily] : []), ...FONT_CHOICES])].map((f) => (
                  <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Size (pt)
              <input type="number" min={1} max={400} step={0.5} value={s.fontSizePt ?? ''} onChange={(e) => set({ fontSizePt: e.target.value ? Math.max(1, Math.min(400, Number(e.target.value))) : undefined })} />
            </label>
            <label>
              Colour
              <input type="color" value={s.color ?? '#000000'} onChange={(e) => set({ color: e.target.value })} />
            </label>
          </div>
          <div className="sc-toggle-row" role="group" aria-label="Font style">
            {(['bold', 'italic', 'underline'] as const).map((k) => (
              <button key={k} className={s[k] ? 'is-active' : ''} aria-pressed={!!s[k]} onClick={() => set({ [k]: !s[k] })} style={{ fontWeight: k === 'bold' ? 700 : undefined, fontStyle: k === 'italic' ? 'italic' : undefined, textDecoration: k === 'underline' ? 'underline' : undefined }}>
                {k[0].toUpperCase() + k.slice(1)}
              </button>
            ))}
            <span className="sc-spacer" />
            {(['left', 'center', 'right', 'justify'] as const).map((a) => (
              <button key={a} className={s.alignment === a ? 'is-active' : ''} aria-pressed={s.alignment === a} onClick={() => set({ alignment: s.alignment === a ? undefined : a })}>
                {a[0].toUpperCase() + a.slice(1)}
              </button>
            ))}
          </div>
          <div className="sc-form-row">
            <label>
              Space before (pt)
              <input type="number" min={0} max={200} value={pt(s.spacingBefore)} onChange={(e) => set({ spacingBefore: e.target.value === '' ? undefined : Math.round(Number(e.target.value) * 20) })} />
            </label>
            <label>
              Space after (pt)
              <input type="number" min={0} max={200} value={pt(s.spacingAfter)} onChange={(e) => set({ spacingAfter: e.target.value === '' ? undefined : Math.round(Number(e.target.value) * 20) })} />
            </label>
            <label>
              Line spacing
              <select value={s.lineSpacing ?? ''} onChange={(e) => set({ lineSpacing: e.target.value ? Number(e.target.value) : undefined })}>
                <option value="">(inherit)</option>
                {[1, 1.08, 1.15, 1.5, 2, 2.5, 3].map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
            </label>
          </div>
        </div>
        <div className="sc-style-preview" aria-label="Preview">
          <p className="sc-muted">Previous paragraph shows how text before this style looks.</p>
          <p
            style={{
              fontFamily: s.fontFamily ? `"${s.fontFamily}", sans-serif` : undefined,
              fontSize: s.fontSizePt ? `${s.fontSizePt}pt` : undefined,
              fontWeight: s.bold ? 700 : 400,
              fontStyle: s.italic ? 'italic' : 'normal',
              textDecoration: s.underline ? 'underline' : 'none',
              color: s.color,
              textAlign: s.alignment,
              lineHeight: s.lineSpacing,
              marginTop: s.spacingBefore != null ? s.spacingBefore / 15 : undefined,
              marginBottom: s.spacingAfter != null ? s.spacingAfter / 15 : undefined,
            }}
          >
            {isNew ? s.name || 'New style' : displayName(s)} — the quick brown fox jumps over the lazy dog.
          </p>
          <p className="sc-muted">Following paragraph text continues here.</p>
        </div>
      </div>
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// Footnotes and endnotes

export function NotesPanel({ editor, meta, setMeta, labels, readOnly, focus, onFocused }: { editor: Editor; meta: DocMeta; setMeta: (f: (m: DocMeta) => DocMeta) => void; labels: Record<string, string>; readOnly: boolean; focus: string | null; onFocused: () => void }) {
  const refs = useRef<Record<string, HTMLTextAreaElement | null>>({})
  const focused = useEffectEvent(() => onFocused())
  useEffect(() => {
    if (!focus) return
    const el = refs.current[focus]
    if (!el) return
    el.focus()
    el.scrollIntoView({ block: 'nearest' })
    focused()
    // labels: the note's text box appears once its reference has been numbered
  }, [focus, labels])
  const kinds = [
    ['footnote', 'Footnotes', meta.footnotes],
    ['endnote', 'Endnotes', meta.endnotes],
  ] as const
  // labels are assigned in document order, so the key order of `labels` is reading order
  const order = (kind: string, id: string) => {
    const i = Object.keys(labels).filter((k) => k.startsWith(kind + ':')).indexOf(`${kind}:${id}`)
    return i === -1 ? 1e9 : i
  }
  const setText = (kind: 'footnote' | 'endnote', id: string, text: string) =>
    setMeta((m) => {
      const upd = (l: typeof m.footnotes) => l.map((n) => (n.id === id ? { ...n, text, edited: true } : n))
      return kind === 'footnote' ? { ...m, footnotes: upd(m.footnotes) } : { ...m, endnotes: upd(m.endnotes) }
    })
  const refPositions = (kind: string, id: string) => {
    const out: number[] = []
    editor.state.doc.descendants((n, pos) => {
      if (n.type.name === 'noteRef' && String(n.attrs.noteId) === id && (n.attrs.kind === 'endnote' ? 'endnote' : 'footnote') === kind) out.push(pos)
    })
    return out
  }
  const goTo = (kind: string, id: string) => {
    const pos = refPositions(kind, id)[0]
    if (pos === undefined) return
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos + 1)).scrollIntoView())
    editor.commands.focus()
  }
  const remove = (kind: 'footnote' | 'endnote', id: string) => {
    const tr = editor.state.tr
    for (const pos of refPositions(kind, id).reverse()) tr.delete(pos, pos + 1)
    if (tr.docChanged) editor.view.dispatch(tr)
    setMeta((m) => (kind === 'footnote' ? { ...m, footnotes: m.footnotes.filter((n) => n.id !== id) } : { ...m, endnotes: m.endnotes.filter((n) => n.id !== id) }))
  }
  return (
    <div className="sc-panel-body">
      <h3>
        <Footprints size={16} aria-hidden="true" /> Footnotes and endnotes
      </h3>
      <p className="sc-muted">Insert them from References (footnote {'⌘/Ctrl+Alt+F'}, endnote {'⌘/Ctrl+Alt+D'}). Numbers follow the order of the references.</p>
      {kinds.map(([kind, title, list]) => {
        const shown = list.filter((n) => labels[`${kind}:${n.id}`]).sort((a, b) => order(kind, a.id) - order(kind, b.id))
        return (
          <section key={kind} className="sc-notes-section">
            <h4>
              {title} <span className="sc-chip-ui">{shown.length}</span>
            </h4>
            {!shown.length ? (
              <p className="sc-muted">None yet.</p>
            ) : (
              <ol className="sc-note-list">
                {shown.map((n) => (
                  <li key={n.id}>
                    <button className="sc-note-num" onClick={() => goTo(kind, n.id)} title="Go to the reference" aria-label={`Go to ${kind} ${labels[`${kind}:${n.id}`]}`}>
                      {labels[`${kind}:${n.id}`]}
                    </button>
                    <textarea
                      ref={(el) => {
                        refs.current[`${kind}:${n.id}`] = el
                      }}
                      aria-label={`${kind === 'footnote' ? 'Footnote' : 'Endnote'} ${labels[`${kind}:${n.id}`]} text`}
                      rows={2}
                      value={n.text}
                      readOnly={readOnly}
                      placeholder="Type the note…"
                      onChange={(e) => setText(kind, n.id, e.target.value)}
                    />
                    {!readOnly && (
                      <button className="sc-icon-mini" aria-label={`Delete ${kind} ${labels[`${kind}:${n.id}`]}`} title="Delete note and its reference" onClick={() => remove(kind, n.id)}>
                        <Trash2 size={13} />
                      </button>
                    )}
                  </li>
                ))}
              </ol>
            )}
          </section>
        )
      })}
    </div>
  )
}

export { type MailMergeData }
