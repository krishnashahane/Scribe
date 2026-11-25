import { createContext, useCallback, useContext, useEffect, useRef, useState, type ReactNode } from 'react'
import { X } from 'lucide-react'

export function IconButton(props: {
  label: string
  icon: ReactNode
  onClick?: () => void
  active?: boolean
  disabled?: boolean
  shortcut?: string
  showLabel?: boolean
  className?: string
}) {
  const title = props.shortcut ? `${props.label} (${props.shortcut})` : props.label
  return (
    <button
      type="button"
      className={`sc-btn ${props.active ? 'is-active' : ''} ${props.showLabel ? 'with-label' : ''} ${props.className ?? ''}`}
      aria-label={props.label}
      aria-pressed={props.active === undefined ? undefined : props.active}
      title={title}
      disabled={props.disabled}
      onMouseDown={(e) => e.preventDefault() /* keep editor selection */}
      onClick={props.onClick}
    >
      {props.icon}
      {props.showLabel && <span>{props.label}</span>}
    </button>
  )
}

export function Dialog(props: { title: string; onClose: () => void; children: ReactNode; footer?: ReactNode; wide?: boolean }) {
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const prev = document.activeElement as HTMLElement | null
    const el = ref.current
    const first = el?.querySelector<HTMLElement>('input, select, textarea, button:not(.sc-dialog-close)')
    first?.focus()
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        props.onClose()
      }
      if (e.key === 'Tab' && el) {
        const f = [...el.querySelectorAll<HTMLElement>('button, input, select, textarea, [tabindex]:not([tabindex="-1"])')].filter((x) => !x.hasAttribute('disabled'))
        if (!f.length) return
        if (e.shiftKey && document.activeElement === f[0]) {
          e.preventDefault()
          f[f.length - 1].focus()
        } else if (!e.shiftKey && document.activeElement === f[f.length - 1]) {
          e.preventDefault()
          f[0].focus()
        }
      }
    }
    document.addEventListener('keydown', onKey, true)
    return () => {
      document.removeEventListener('keydown', onKey, true)
      prev?.focus?.()
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])
  return (
    <div className="sc-overlay" onMouseDown={(e) => e.target === e.currentTarget && props.onClose()}>
      <div className={`sc-dialog ${props.wide ? 'is-wide' : ''}`} role="dialog" aria-modal="true" aria-label={props.title} ref={ref}>
        <header>
          <h2>{props.title}</h2>
          <button className="sc-dialog-close" aria-label="Close" onClick={props.onClose}>
            <X size={18} />
          </button>
        </header>
        <div className="sc-dialog-body">{props.children}</div>
        {props.footer && <footer>{props.footer}</footer>}
      </div>
    </div>
  )
}

type Toast = { id: number; text: string; kind: 'info' | 'error' | 'success' }
const ToastCtx = createContext<(text: string, kind?: Toast['kind']) => void>(() => {})

export function ToastProvider({ children }: { children: ReactNode }) {
  const [toasts, setToasts] = useState<Toast[]>([])
  const push = useCallback((text: string, kind: Toast['kind'] = 'info') => {
    const id = Date.now() + Math.random()
    setToasts((t) => [...t.slice(-3), { id, text, kind }])
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), kind === 'error' ? 8000 : 4000)
  }, [])
  return (
    <ToastCtx.Provider value={push}>
      {children}
      <div className="sc-toasts" role="status" aria-live="polite">
        {toasts.map((t) => (
          <div key={t.id} className={`sc-toast is-${t.kind}`}>
            {t.text}
          </div>
        ))}
      </div>
    </ToastCtx.Provider>
  )
}

export const useToast = () => useContext(ToastCtx)

export function Menu(props: { label: ReactNode; ariaLabel: string; items: ({ label: string; onClick: () => void; shortcut?: string; disabled?: boolean } | 'sep')[]; className?: string }) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    if (!open) return
    const close = (e: MouseEvent) => !ref.current?.contains(e.target as Node) && setOpen(false)
    const esc = (e: KeyboardEvent) => e.key === 'Escape' && setOpen(false)
    document.addEventListener('mousedown', close)
    document.addEventListener('keydown', esc)
    ref.current?.querySelector<HTMLElement>('[role=menuitem]')?.focus()
    return () => {
      document.removeEventListener('mousedown', close)
      document.removeEventListener('keydown', esc)
    }
  }, [open])
  const onKeyDown = (e: React.KeyboardEvent) => {
    const items = [...(ref.current?.querySelectorAll<HTMLElement>('[role=menuitem]:not([disabled])') ?? [])]
    const i = items.indexOf(document.activeElement as HTMLElement)
    if (e.key === 'ArrowDown') {
      e.preventDefault()
      items[(i + 1) % items.length]?.focus()
    } else if (e.key === 'ArrowUp') {
      e.preventDefault()
      items[(i - 1 + items.length) % items.length]?.focus()
    }
  }
  return (
    <div className={`sc-menu ${props.className ?? ''}`} ref={ref}>
      <button className="sc-menu-trigger" aria-haspopup="menu" aria-expanded={open} aria-label={props.ariaLabel} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)}>
        {props.label}
      </button>
      {open && (
        <div className="sc-menu-list" role="menu" onKeyDown={onKeyDown}>
          {props.items.map((it, i) =>
            it === 'sep' ? (
              <div key={i} className="sc-menu-sep" role="separator" />
            ) : (
              <button
                key={i}
                role="menuitem"
                disabled={it.disabled}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  setOpen(false)
                  it.onClick()
                }}
              >
                <span>{it.label}</span>
                {it.shortcut && <kbd>{it.shortcut}</kbd>}
              </button>
            ),
          )}
        </div>
      )}
    </div>
  )
}

export const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform)
export const mod = isMac ? '⌘' : 'Ctrl+'
