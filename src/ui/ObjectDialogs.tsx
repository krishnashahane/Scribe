import { useEffect, useMemo, useRef, useState } from 'react'
import { Plus, Trash2 } from 'lucide-react'
import { Dialog } from './kit'
import { MathSyntaxError, equationPieces, linearBlocker, ommlToTree, parseLinear, replaceEquation, setEquationPieces, toLinear, treeToMathml, treeToOmml } from '../objects/math'
import { setTextBoxParagraphs, textBoxIsSimple, textBoxParagraphs } from '../objects/textbox'
import { chartSvg, type ChartData, type ChartKind } from '../objects/chart'
import type { SmartArtData } from '../objects/smartart'
import { DEFAULT_SHAPE, SHAPE_KINDS, shapeSvg, type ShapeAttrs, type ShapeKind } from '../objects/shapes'
import type { CustomXmlPart } from '../model/types'
import { parseXml } from '../docx/xml'

const Footer = ({ onClose, onSave, disabled, label = 'Save' }: { onClose: () => void; onSave: () => void; disabled?: boolean; label?: string }) => (
  <>
    <button onClick={onClose}>Cancel</button>
    <button className="sc-primary" disabled={disabled} onClick={onSave}>
      {label}
    </button>
  </>
)

// ---------------------------------------------------------------------------
// Equations

const EXAMPLES: [string, string][] = [
  ['Fraction', 'a/b'],
  ['Power', 'x^2'],
  ['Subscript', 'x_i'],
  ['Root', '√(x+1)'],
  ['Sum', '\\sum_(i=1)^n i'],
  ['Integral', '\\int_0^1 f(x) dx'],
  ['Quadratic', 'x = (-b ± √(b^2-4ac))/(2a)'],
]

export function EquationDialog({ xml, onSave, onClose }: { xml: string | null; onSave: (xml: string) => void; onClose: () => void }) {
  const blocker = useMemo(() => (xml ? linearBlocker(xml) : null), [xml])
  const [text, setText] = useState(() => {
    if (!xml || blocker) return ''
    try {
      return toLinear(ommlToTree(xml))
    } catch {
      return ''
    }
  })
  const [pieces, setPieces] = useState<string[]>(() => (xml && blocker ? equationPieces(xml) : []))
  const input = useRef<HTMLTextAreaElement>(null)
  const parsed = useMemo(() => {
    if (blocker) return null
    if (!text.trim()) return { error: '' }
    try {
      return { tree: parseLinear(text.trim()) }
    } catch (e) {
      return { error: e instanceof MathSyntaxError ? e.message : 'This equation could not be read.' }
    }
  }, [text, blocker])
  const insert = (snippet: string) => {
    const el = input.current
    if (!el) return setText((t) => t + snippet)
    const s = el.selectionStart
    const e = el.selectionEnd
    setText(text.slice(0, s) + snippet + text.slice(e))
    requestAnimationFrame(() => {
      el.focus()
      el.setSelectionRange(s + snippet.length, s + snippet.length)
    })
  }
  const save = () => {
    if (blocker && xml) return onSave(setEquationPieces(xml, pieces))
    if (parsed && 'tree' in parsed && parsed.tree) onSave(xml ? replaceEquation(xml, parsed.tree) : treeToOmml(parsed.tree))
  }
  if (blocker)
    return (
      <Dialog title="Edit equation" onClose={onClose} wide footer={<Footer onClose={onClose} onSave={save} />}>
        <p className="sc-muted">{blocker} Change the text of each piece below; the structure is kept exactly.</p>
        <div className="sc-form sc-pieces">
          {pieces.map((p, i) => (
            <label key={i}>
              Piece {i + 1}
              <input value={p} onChange={(e) => setPieces(pieces.map((x, k) => (k === i ? e.target.value : x)))} />
            </label>
          ))}
        </div>
      </Dialog>
    )
  const ok = !!parsed && 'tree' in parsed && !!parsed.tree
  return (
    <Dialog title={xml ? 'Edit equation' : 'Insert equation'} onClose={onClose} wide footer={<Footer onClose={onClose} onSave={save} disabled={!ok} label={xml ? 'Save' : 'Insert'} />}>
      <div className="sc-form">
        <label>
          Equation
          <textarea ref={input} className="sc-mono" rows={3} value={text} onChange={(e) => setText(e.target.value)} placeholder="e.g. E = mc^2  or  (a+b)/2" spellCheck={false} autoFocus />
        </label>
        <div className="sc-chips-row" aria-label="Examples">
          {EXAMPLES.map(([label, s]) => (
            <button key={label} type="button" className="sc-chip-btn" onClick={() => insert(s)} title={s}>
              {label}
            </button>
          ))}
          {['α', 'β', 'π', 'θ', '∞', '≤', '≥', '≠', '±', '×', '→'].map((c) => (
            <button key={c} type="button" className="sc-chip-btn is-symbol" onClick={() => insert(c)}>
              {c}
            </button>
          ))}
        </div>
        <div className="sc-eq-preview" aria-live="polite">
          {/* MathML is generated from escaped text only */}
          {ok ? <span dangerouslySetInnerHTML={{ __html: treeToMathml(parsed!.tree!, true) }} /> : <span className={parsed?.error ? 'sc-error' : 'sc-muted'}>{parsed?.error || 'The preview appears here.'}</span>}
        </div>
        <details className="sc-help">
          <summary>Syntax</summary>
          <p>
            <code>a/b</code> fraction · <code>x^2</code> power · <code>x_i</code> subscript · <code>√x</code> or <code>\sqrt(x)</code> root · <code>\sum_(i=1)^n</code>, <code>\int_0^1</code> sums and integrals ·
            <code>( )</code> visible brackets, dropped around fraction parts and scripts · <code>{'{ }'}</code> invisible grouping · <code>\alpha</code>, <code>\pi</code>, <code>\le</code>… symbols · <code>\^</code> a literal character.
          </p>
        </details>
      </div>
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// Text boxes

export function TextBoxDialog({ xml, onSave, onClose }: { xml: string; onSave: (xml: string, preview: string) => void; onClose: () => void }) {
  const [text, setText] = useState(() => (textBoxParagraphs(xml) ?? []).join('\n'))
  const simple = useMemo(() => textBoxIsSimple(xml), [xml])
  return (
    <Dialog
      title="Edit text box"
      onClose={onClose}
      wide
      footer={<Footer onClose={onClose} onSave={() => onSave(setTextBoxParagraphs(xml, text.split('\n')), text.replace(/\n/g, ' ').slice(0, 200))} />}
    >
      <div className="sc-form">
        <label>
          Text (one paragraph per line)
          <textarea rows={8} value={text} onChange={(e) => setText(e.target.value)} autoFocus />
        </label>
        <p className="sc-muted">
          The text box keeps its size, position, border and fill. {simple ? 'Paragraph and character formatting is kept.' : 'It contains mixed formatting or fields; each edited paragraph takes the formatting of its first run.'}
        </p>
      </div>
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// Charts

const KIND_LABEL: Record<ChartKind, string> = { column: 'Column', bar: 'Bar', line: 'Line', area: 'Area', pie: 'Pie', doughnut: 'Doughnut' }

export function ChartDialog({ chart, onSave, onClose }: { chart: ChartData; onSave: (c: ChartData) => void; onClose: () => void }) {
  const [kind, setKind] = useState<ChartKind>(chart.kind)
  const [title, setTitle] = useState(chart.title)
  const [cats, setCats] = useState(chart.categories.length ? chart.categories : ['Category 1'])
  const [series, setSeries] = useState(chart.series.length ? chart.series.map((s) => ({ name: s.name, values: s.values.map((v) => (v == null ? '' : String(v))) })) : [{ name: 'Series 1', values: [''] }])
  const rows = cats.length
  const numeric = (v: string) => v.trim() === '' || Number.isFinite(Number(v.trim()))
  const invalid = series.some((s) => s.values.some((v) => !numeric(v)))
  const data: ChartData = {
    ...chart,
    kind,
    title,
    categories: cats,
    series: series.map((s) => ({ name: s.name, values: Array.from({ length: rows }, (_, i) => (s.values[i]?.trim() ? Number(s.values[i]) : null)) })),
    dirty: true,
  }
  if (!chart.editable)
    return (
      <Dialog title="Chart" onClose={onClose} footer={<button className="sc-primary" onClick={onClose}>Close</button>}>
        <p>{chart.reason}</p>
        <p className="sc-muted">The chart is kept exactly as it is in the original file.</p>
      </Dialog>
    )
  const setCell = (si: number, ri: number, v: string) => setSeries(series.map((s, k) => (k === si ? { ...s, values: Array.from({ length: rows }, (_, i) => (i === ri ? v : s.values[i] ?? '')) } : s)))
  return (
    <Dialog title={chart.isNew ? 'Chart' : `Edit ${KIND_LABEL[chart.kind].toLowerCase()} chart`} onClose={onClose} wide footer={<Footer onClose={onClose} onSave={() => onSave(data)} disabled={invalid} label="Save" />}>
      <div className="sc-chart-editor">
        <div className="sc-form">
          {chart.isNew && (
            <div className="sc-toggle-row" role="radiogroup" aria-label="Chart type">
              {(Object.keys(KIND_LABEL) as ChartKind[]).map((k) => (
                <button key={k} role="radio" aria-checked={kind === k} className={kind === k ? 'is-active' : ''} onClick={() => setKind(k)}>
                  {KIND_LABEL[k]}
                </button>
              ))}
            </div>
          )}
          <label>
            Chart title
            <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="No title" />
          </label>
          <div className="sc-grid-wrap">
            <table className="sc-data-grid">
              <thead>
                <tr>
                  <th scope="col">Category</th>
                  {series.map((s, si) => (
                    <th key={si} scope="col">
                      <input aria-label={`Series ${si + 1} name`} value={s.name} onChange={(e) => setSeries(series.map((x, k) => (k === si ? { ...x, name: e.target.value } : x)))} />
                      {series.length > 1 && (
                        <button className="sc-icon-mini" aria-label={`Remove series ${s.name || si + 1}`} onClick={() => setSeries(series.filter((_, k) => k !== si))}>
                          <Trash2 size={13} />
                        </button>
                      )}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {cats.map((c, ri) => (
                  <tr key={ri}>
                    <th scope="row">
                      <input aria-label={`Category ${ri + 1}`} value={c} onChange={(e) => setCats(cats.map((x, k) => (k === ri ? e.target.value : x)))} />
                    </th>
                    {series.map((s, si) => (
                      <td key={si}>
                        <input inputMode="decimal" aria-label={`${s.name || `Series ${si + 1}`}, ${c || `row ${ri + 1}`}`} className={numeric(s.values[ri] ?? '') ? '' : 'is-invalid'} value={s.values[ri] ?? ''} onChange={(e) => setCell(si, ri, e.target.value)} />
                      </td>
                    ))}
                    <td className="sc-row-tools">
                      {cats.length > 1 && (
                        <button className="sc-icon-mini" aria-label={`Remove ${c || `row ${ri + 1}`}`} onClick={() => {
                          setCats(cats.filter((_, k) => k !== ri))
                          setSeries(series.map((s) => ({ ...s, values: s.values.filter((_, k) => k !== ri) })))
                        }}>
                          <Trash2 size={13} />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="sc-inline-actions">
            <button onClick={() => setCats([...cats, `Category ${cats.length + 1}`])}>
              <Plus size={14} /> Row
            </button>
            {kind !== 'pie' && kind !== 'doughnut' && (
              <button onClick={() => setSeries([...series, { name: `Series ${series.length + 1}`, values: Array(rows).fill('') }])}>
                <Plus size={14} /> Series
              </button>
            )}
          </div>
          {invalid && <p className="sc-error">Values must be numbers (or empty).</p>}
        </div>
        <div className="sc-chart-preview" dangerouslySetInnerHTML={{ __html: chartSvg(data, 360, 240) }} />
      </div>
      <p className="sc-muted">{chart.isNew ? (kind === 'pie' || kind === 'doughnut' ? 'Pie and doughnut charts show the first series. ' : '') + 'The chart is saved as a real chart in DOCX files.' : 'The chart keeps its type, colours and formatting. Edited values are saved inside the chart.'}</p>
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// SmartArt

export function SmartArtDialog({ data, onSave, onClose }: { data: SmartArtData; onSave: (d: SmartArtData) => void; onClose: () => void }) {
  const [items, setItems] = useState(data.items)
  return (
    <Dialog title="Edit SmartArt text" onClose={onClose} wide footer={<Footer onClose={onClose} onSave={() => onSave({ ...data, items, dirty: true })} />}>
      {!items.length ? (
        <p className="sc-muted">This diagram has no text nodes.</p>
      ) : (
        <div className="sc-form">
          {items.map((it, i) => (
            <label key={it.id} style={{ paddingLeft: Math.min(4, it.depth) * 18 }}>
              {it.depth ? 'Sub-item' : 'Item'} {i + 1}
              <input value={it.text} onChange={(e) => setItems(items.map((x, k) => (k === i ? { ...x, text: e.target.value } : x)))} />
            </label>
          ))}
        </div>
      )}
      <p className="sc-muted">The layout, colours and shapes are kept. {data.drawingPart ? 'The diagram picture is updated with the new text.' : 'This file has no cached picture, so other editors redraw the diagram from the new text.'}</p>
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// Shapes

export function ShapeDialog({ initial, onSave, onClose }: { initial: ShapeAttrs; onSave: (a: ShapeAttrs) => void; onClose: () => void }) {
  const [a, setA] = useState<ShapeAttrs>(initial)
  const set = (p: Partial<ShapeAttrs>) => setA((x) => ({ ...x, ...p }))
  const isLine = a.kind === 'line'
  return (
    <Dialog title={`Format ${SHAPE_KINDS[a.kind].toLowerCase()}`} onClose={onClose} wide footer={<Footer onClose={onClose} onSave={() => onSave(a)} disabled={!(a.width > 0 && a.height > 0)} />}>
      <div className="sc-shape-editor">
        <div className="sc-form">
          <label>
            Shape
            <select value={a.kind} onChange={(e) => set({ kind: e.target.value as ShapeKind, ...(e.target.value === 'line' ? { height: 4 } : a.kind === 'line' ? { height: 96 } : {}) })}>
              {Object.entries(SHAPE_KINDS).map(([k, v]) => (
                <option key={k} value={k}>
                  {v}
                </option>
              ))}
            </select>
          </label>
          <div className="sc-form-row">
            <label>
              Width (px)
              <input type="number" min={4} max={1200} value={Math.round(a.width)} onChange={(e) => set({ width: Math.max(4, Math.min(1200, Number(e.target.value) || 0)) })} />
            </label>
            <label>
              Height (px)
              <input type="number" min={isLine ? 1 : 4} max={1600} value={Math.round(a.height)} onChange={(e) => set({ height: Math.max(isLine ? 1 : 4, Math.min(1600, Number(e.target.value) || 0)) })} />
            </label>
          </div>
          <div className="sc-form-row">
            {!isLine && (
              <label>
                Fill
                <span className="sc-color-field">
                  <input type="color" value={a.fill ?? '#ffffff'} disabled={!a.fill} onChange={(e) => set({ fill: e.target.value })} aria-label="Fill colour" />
                  <span className="sc-check">
                    <input type="checkbox" checked={!a.fill} onChange={(e) => set({ fill: e.target.checked ? null : DEFAULT_SHAPE.fill })} /> None
                  </span>
                </span>
              </label>
            )}
            <label>
              Outline
              <span className="sc-color-field">
                <input type="color" value={a.stroke ?? '#000000'} disabled={!a.stroke} onChange={(e) => set({ stroke: e.target.value })} aria-label="Outline colour" />
                <span className="sc-check">
                  <input type="checkbox" checked={!a.stroke} disabled={isLine} onChange={(e) => set({ stroke: e.target.checked ? null : DEFAULT_SHAPE.stroke })} /> None
                </span>
              </span>
            </label>
            <label>
              Weight (px)
              <input type="number" min={0} max={20} step={0.5} value={a.strokeWidth} onChange={(e) => set({ strokeWidth: Math.max(0, Math.min(20, Number(e.target.value) || 0)) })} />
            </label>
          </div>
          {isLine ? (
            <label className="sc-check">
              <input type="checkbox" checked={!!a.arrow} onChange={(e) => set({ arrow: e.target.checked })} /> Arrow head
            </label>
          ) : (
            <label>
              Text in shape
              <textarea rows={3} value={a.text} onChange={(e) => set({ text: e.target.value })} />
            </label>
          )}
        </div>
        <div className="sc-shape-preview" dangerouslySetInnerHTML={{ __html: shapeSvg({ ...a, width: Math.min(a.width, 280), height: Math.min(a.height, 220) }) }} />
      </div>
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// Freehand drawing -> PNG image

export function DrawDialog({ onSave, onClose }: { onSave: (png: Blob, w: number, h: number) => void; onClose: () => void }) {
  const canvas = useRef<HTMLCanvasElement>(null)
  const [color, setColor] = useState('#1f2937')
  const [size, setSize] = useState(3)
  const [eraser, setEraser] = useState(false)
  const [empty, setEmpty] = useState(true)
  const strokes = useRef<ImageData[]>([])
  const W = 640
  const H = 360
  useEffect(() => {
    const c = canvas.current!
    const ratio = window.devicePixelRatio || 1
    c.width = W * ratio
    c.height = H * ratio
    const ctx = c.getContext('2d')!
    ctx.scale(ratio, ratio)
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
  }, [])
  const pos = (e: React.PointerEvent) => {
    const r = canvas.current!.getBoundingClientRect()
    return { x: ((e.clientX - r.left) / r.width) * W, y: ((e.clientY - r.top) / r.height) * H }
  }
  const drawing = useRef(false)
  const down = (e: React.PointerEvent) => {
    const ctx = canvas.current!.getContext('2d')!
    strokes.current.push(ctx.getImageData(0, 0, canvas.current!.width, canvas.current!.height))
    if (strokes.current.length > 30) strokes.current.shift()
    drawing.current = true
    canvas.current!.setPointerCapture(e.pointerId)
    const p = pos(e)
    ctx.globalCompositeOperation = eraser ? 'destination-out' : 'source-over'
    ctx.strokeStyle = color
    ctx.lineWidth = eraser ? size * 4 : size
    ctx.beginPath()
    ctx.moveTo(p.x, p.y)
    ctx.lineTo(p.x + 0.01, p.y + 0.01)
    ctx.stroke()
    setEmpty(false)
  }
  const move = (e: React.PointerEvent) => {
    if (!drawing.current) return
    const ctx = canvas.current!.getContext('2d')!
    const p = pos(e)
    ctx.lineTo(p.x, p.y)
    ctx.stroke()
  }
  const up = () => {
    drawing.current = false
  }
  const undo = () => {
    const last = strokes.current.pop()
    if (last) canvas.current!.getContext('2d')!.putImageData(last, 0, 0)
  }
  const clear = () => {
    const c = canvas.current!
    const ctx = c.getContext('2d')!
    strokes.current.push(ctx.getImageData(0, 0, c.width, c.height))
    ctx.save()
    ctx.setTransform(1, 0, 0, 1, 0, 0)
    ctx.clearRect(0, 0, c.width, c.height)
    ctx.restore()
    setEmpty(true)
  }
  const save = () => {
    const c = canvas.current!
    // crop to the drawn area with a small margin
    const ctx = c.getContext('2d')!
    const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height)
    let x0 = width
    let y0 = height
    let x1 = -1
    let y1 = -1
    for (let y = 0; y < height; y++)
      for (let x = 0; x < width; x++)
        if (data[(y * width + x) * 4 + 3]) {
          if (x < x0) x0 = x
          if (x > x1) x1 = x
          if (y < y0) y0 = y
          if (y > y1) y1 = y
        }
    if (x1 < 0) return
    const pad = 8 * (window.devicePixelRatio || 1)
    x0 = Math.max(0, x0 - pad)
    y0 = Math.max(0, y0 - pad)
    x1 = Math.min(width - 1, x1 + pad)
    y1 = Math.min(height - 1, y1 + pad)
    const out = document.createElement('canvas')
    out.width = x1 - x0 + 1
    out.height = y1 - y0 + 1
    out.getContext('2d')!.drawImage(c, x0, y0, out.width, out.height, 0, 0, out.width, out.height)
    const ratio = window.devicePixelRatio || 1
    out.toBlob((b) => b && onSave(b, Math.round(out.width / ratio), Math.round(out.height / ratio)), 'image/png')
  }
  return (
    <Dialog title="Draw" onClose={onClose} wide footer={<Footer onClose={onClose} onSave={save} disabled={empty} label="Insert drawing" />}>
      <div className="sc-draw-tools" role="toolbar" aria-label="Pen">
        {['#1f2937', '#2563eb', '#dc2626', '#16a34a', '#d97706', '#7c3aed'].map((c) => (
          <button key={c} className={`sc-swatch ${color === c && !eraser ? 'is-active' : ''}`} style={{ background: c }} aria-label={`Pen colour ${c}`} onClick={() => { setColor(c); setEraser(false) }} />
        ))}
        <input type="color" value={color} onChange={(e) => { setColor(e.target.value); setEraser(false) }} aria-label="Custom pen colour" />
        <label className="sc-range">
          Size
          <input type="range" min={1} max={16} value={size} onChange={(e) => setSize(Number(e.target.value))} />
        </label>
        <button className={eraser ? 'is-active' : ''} aria-pressed={eraser} onClick={() => setEraser(!eraser)}>
          Eraser
        </button>
        <button onClick={undo}>Undo</button>
        <button onClick={clear}>Clear</button>
      </div>
      <canvas ref={canvas} className="sc-draw-canvas" style={{ aspectRatio: `${W} / ${H}` }} onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerCancel={up} aria-label="Drawing area" />
      <p className="sc-muted">Your drawing is inserted as a picture, cropped to what you drew.</p>
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// Custom XML data parts

export function CustomXmlDialog({ parts, readOnly, onSave, onClose }: { parts: CustomXmlPart[]; readOnly: boolean; onSave: (p: CustomXmlPart[]) => void; onClose: () => void }) {
  const [list, setList] = useState(parts)
  const [sel, setSel] = useState(0)
  const cur = list[sel]
  const error = useMemo(() => {
    if (!cur) return null
    try {
      parseXml(cur.xml, cur.path)
      return null
    } catch (e: any) {
      return String(e.message || 'Not well-formed XML')
    }
  }, [cur])
  const invalid = list.some((p) => {
    try {
      parseXml(p.xml, p.path)
      return false
    } catch {
      return true
    }
  })
  return (
    <Dialog title="Custom XML data" onClose={onClose} wide footer={readOnly ? <button className="sc-primary" onClick={onClose}>Close</button> : <Footer onClose={onClose} onSave={() => onSave(list)} disabled={invalid} />}>
      {!list.length ? (
        <p className="sc-muted">This document has no custom XML data parts. They come from files created by other applications, usually to store data bound to the document.</p>
      ) : (
        <div className="sc-form">
          <label>
            Part
            <select value={sel} onChange={(e) => setSel(Number(e.target.value))}>
              {list.map((p, i) => (
                <option key={p.path} value={i}>
                  {p.path}
                  {p.edited ? ' (edited)' : ''}
                </option>
              ))}
            </select>
          </label>
          <label>
            XML
            <textarea className="sc-mono" rows={14} spellCheck={false} readOnly={readOnly} value={cur.xml} onChange={(e) => setList(list.map((p, i) => (i === sel ? { ...p, xml: e.target.value, edited: true } : p)))} />
          </label>
          {error ? <p className="sc-error">{error}</p> : <p className="sc-muted">Well-formed. Only edited parts are rewritten when you export to DOCX.</p>}
        </div>
      )}
    </Dialog>
  )
}

// ---------------------------------------------------------------------------
// Symbols and special characters

const SYMBOL_SETS: [string, string][] = [
  ['Common', '©®™§¶†‡•…–—‘’“”«»‹›°±×÷≈≠≤≥∞√∑∏∫∂µ¼½¾¹²³'],
  ['Currency', '€£¥¢₹₩₽₺₪₫₴₦₱฿₿$¤'],
  ['Arrows', '←↑→↓↔↕⇐⇑⇒⇓⇔↩↪↺↻➔➜⟵⟶⟷'],
  ['Math', '∀∃∄∅∈∉∋⊂⊃⊆⊇∪∩∧∨¬⊕⊗⊥∥∠∴∵∝∼≅≡≪≫∇∆∑∏∫∮ℕℤℚℝℂ'],
  ['Greek', 'αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ'],
  ['Latin', 'ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÑÒÓÔÕÖØŒÙÚÛÜÝßàáâãäåæçèéêëìíîïñòóôõöøœùúûüýÿ'],
  ['Shapes', '■□▪▫▲△▶▷▼▽◀◁◆◇○●◎★☆♠♣♥♦✓✔✗✘☐☑☒'],
  ['Spaces & marks', '   ‑­‐‑‒¡¿‰′″‴'],
]
const SPACE_NAMES: Record<string, string> = { ' ': 'No-break space', ' ': 'Thin space', ' ': 'Hair space', '‑': 'No-break hyphen', '­': 'Soft hyphen' }

export function SymbolDialog({ onInsert, onClose }: { onInsert: (ch: string) => void; onClose: () => void }) {
  const [set, setSet] = useState(0)
  const [recent, setRecent] = useState<string[]>(() => {
    try {
      const r = JSON.parse(localStorage.getItem('scribe-recent-symbols') || '[]')
      return Array.isArray(r) ? r.filter((x) => typeof x === 'string' && [...x].length === 1).slice(0, 16) : []
    } catch {
      return []
    }
  })
  const pick = (ch: string) => {
    onInsert(ch)
    const next = [ch, ...recent.filter((x) => x !== ch)].slice(0, 16)
    setRecent(next)
    localStorage.setItem('scribe-recent-symbols', JSON.stringify(next))
  }
  const label = (ch: string) => SPACE_NAMES[ch] ?? `${ch} (U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')})`
  const grid = (chars: string[]) => (
    <div className="sc-symbol-grid" role="grid">
      {chars.map((ch) => (
        <button key={ch} className="sc-symbol" title={label(ch)} aria-label={label(ch)} onMouseDown={(e) => e.preventDefault()} onClick={() => pick(ch)}>
          {SPACE_NAMES[ch] ? <span className="sc-symbol-name">{SPACE_NAMES[ch].replace(/ (space|hyphen)$/, '')}</span> : ch}
        </button>
      ))}
    </div>
  )
  return (
    <Dialog title="Insert symbol" onClose={onClose} wide footer={<button className="sc-primary" onClick={onClose}>Done</button>}>
      <div className="sc-toggle-row" role="tablist" aria-label="Symbol sets">
        {SYMBOL_SETS.map(([name], i) => (
          <button key={name} role="tab" aria-selected={set === i} className={set === i ? 'is-active' : ''} onClick={() => setSet(i)}>
            {name}
          </button>
        ))}
      </div>
      {recent.length > 0 && (
        <>
          <p className="sc-muted sc-symbol-heading">Recently used</p>
          {grid(recent)}
        </>
      )}
      <p className="sc-muted sc-symbol-heading">{SYMBOL_SETS[set][0]}</p>
      {grid([...SYMBOL_SETS[set][1]])}
      <p className="sc-muted">Click a symbol to insert it at the cursor; the dialog stays open so you can add more.</p>
    </Dialog>
  )
}
