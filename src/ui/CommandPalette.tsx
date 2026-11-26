import { useEffect, useMemo, useRef, useState } from 'react'
import { Search } from 'lucide-react'

export interface Command {
  id: string
  label: string
  group: string
  run: () => void
  shortcut?: string
  keywords?: string
  disabled?: boolean
}

/** Score how well `q` matches a command (lower is better); null = no match. */
function score(q: string, c: Command): number | null {
  const hay = `${c.label} ${c.group} ${c.keywords ?? ''}`.toLowerCase()
  const words = q.toLowerCase().split(/\s+/).filter(Boolean)
  if (!words.length) return 0
  let total = 0
  for (const w of words) {
    const i = hay.indexOf(w)
    if (i === -1) return null
    total += i === 0 ? 0 : /\s/.test(hay[i - 1]) ? 1 : 3
  }
  return total + (c.label.toLowerCase().startsWith(words[0]) ? -2 : 0)
}

/** Search every command by name and run it with Enter. */
export function CommandPalette({ getCommands, onClose }: { getCommands: () => Command[]; onClose: () => void }) {
  // the list is built once when the palette opens
  const [commands] = useState(getCommands)
  const [q, setQ] = useState('')
  const [sel, setSel] = useState(0)
  const input = useRef<HTMLInputElement>(null)
  const list = useRef<HTMLUListElement>(null)
  const results = useMemo(
    () =>
      commands
        .filter((c) => !c.disabled)
        .map((c) => ({ c, s: score(q, c) }))
        .filter((x): x is { c: Command; s: number } => x.s !== null)
        .sort((a, b) => a.s - b.s)
        .slice(0, 60)
        .map((x) => x.c),
    [q, commands],
  )
  useEffect(() => input.current?.focus(), [])
  useEffect(() => {
    list.current?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' })
  }, [sel])
  const run = (c: Command | undefined) => {
    if (!c) return
    onClose()
    // after the palette has closed and focus has returned to the document
    setTimeout(c.run, 0)
  }
  const onKey = (e: React.KeyboardEvent) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault()
      setSel((i) => Math.min(results.length - 1, i + 1))
    } else if (e.key === 'ArrowUp') {
      e.preventDefault()
      setSel((i) => Math.max(0, i - 1))
    } else if (e.key === 'Enter') {
      e.preventDefault()
      run(results[sel])
    } else if (e.key === 'Escape') {
      e.preventDefault()
      onClose()
    }
  }
  return (
    <div className="sc-overlay sc-palette-overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="sc-palette" role="dialog" aria-modal="true" aria-label="Search commands">
        <div className="sc-palette-input">
          <Search size={18} aria-hidden="true" />
          <input
            ref={input}
            value={q}
            onChange={(e) => {
              setQ(e.target.value)
              setSel(0)
            }}
            onKeyDown={onKey}
            placeholder="Search commands — e.g. footnote, two columns, export PDF"
            aria-label="Search commands"
            role="combobox"
            aria-expanded="true"
            aria-controls="sc-palette-list"
            aria-activedescendant={results[sel] ? `cmd-${results[sel].id}` : undefined}
          />
          <kbd>Esc</kbd>
        </div>
        <ul id="sc-palette-list" ref={list} role="listbox" className="sc-palette-list">
          {results.map((c, i) => (
            <li key={c.id} id={`cmd-${c.id}`} role="option" aria-selected={i === sel} className={i === sel ? 'is-active' : ''} onMouseEnter={() => setSel(i)} onMouseDown={(e) => e.preventDefault()} onClick={() => run(c)}>
              <span className="sc-palette-label">{c.label}</span>
              <span className="sc-palette-group">{c.group}</span>
              {c.shortcut && <kbd>{c.shortcut}</kbd>}
            </li>
          ))}
          {!results.length && <li className="sc-palette-empty">No matching commands.</li>}
        </ul>
      </div>
    </div>
  )
}
