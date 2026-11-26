// Charts: read the cached data of a DrawingML chart part, write edited data back, and draw a preview.

import { children, parseXml, serializeXml } from '../docx/xml'

export const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

export type ChartKind = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut'

export interface ChartSeries {
  name: string
  values: (number | null)[]
}

export interface ChartData {
  /** package path of the chart part, e.g. word/charts/chart1.xml */
  part: string
  kind: ChartKind
  title: string
  categories: string[]
  series: ChartSeries[]
  /** false when the chart type or layout cannot be edited (scatter, combo charts...) */
  editable: boolean
  reason?: string
  /** set when the data was changed in SCRIBE; the exporter rewrites the part */
  dirty?: boolean
  /** created in SCRIBE: the exporter writes a new chart part (part is empty until then) */
  isNew?: boolean
}

export const NEW_CHART_SAMPLE: Pick<ChartData, 'categories' | 'series'> = {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
  ],
}

export function newChart(kind: ChartKind): ChartData {
  const pie = kind === 'pie' || kind === 'doughnut'
  return {
    part: '',
    kind,
    title: 'Chart title',
    categories: [...NEW_CHART_SAMPLE.categories],
    series: (pie ? NEW_CHART_SAMPLE.series.slice(0, 1) : NEW_CHART_SAMPLE.series).map((x) => ({ ...x, values: [...x.values] })),
    editable: true,
    isNew: true,
    dirty: true,
  }
}

/** Explicit slice colour: readers without chart style parts draw uncoloured slices otherwise. */
function dPtXml(k: number) {
  const col = CHART_COLORS[k % 8].slice(1).toUpperCase()
  return `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${col}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`
}

/** A complete chart part for a chart created in SCRIBE. */
export function newChartXml(d: ChartData): string {
  const esc = (v: string) => v.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  const lit = (tag: 'strLit' | 'numLit', vals: (string | number | null)[]) =>
    `<c:${tag}>${tag === 'numLit' ? '<c:formatCode>General</c:formatCode>' : ''}<c:ptCount val="${vals.length}"/>${vals
      .map((v, i) => (v === null || v === '' ? '' : `<c:pt idx="${i}"><c:v>${esc(String(v))}</c:v></c:pt>`))
      .join('')}</c:${tag}>`
  const pie = d.kind === 'pie' || d.kind === 'doughnut'
  const sers = d.series
    .map(
      (s, i) =>
        `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${esc(s.name)}</c:v></c:tx>` +
        (pie ? '' : `<c:spPr><a:solidFill><a:srgbClr val="${CHART_COLORS[i % 8].slice(1).toUpperCase()}"/></a:solidFill>${d.kind === 'line' ? '<a:ln w="28575"><a:solidFill><a:srgbClr val="' + CHART_COLORS[i % 8].slice(1).toUpperCase() + '"/></a:solidFill></a:ln>' : ''}</c:spPr>`) +
        (d.kind === 'column' || d.kind === 'bar' ? '<c:invertIfNegative val="0"/>' : '') +
        (pie ? d.categories.map((_, k) => dPtXml(k)).join('') : '') +
        (d.kind === 'line' ? `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="${CHART_COLORS[i % 8].slice(1).toUpperCase()}"/></a:solidFill></c:spPr></c:marker>` : '') +
        `<c:cat>${lit('strLit', d.categories)}</c:cat><c:val>${lit('numLit', s.values)}</c:val>` +
        (d.kind === 'line' ? '<c:smooth val="0"/>' : '') +
        `</c:ser>`,
    )
    .join('')
  const axes = '<c:axId val="50010001"/><c:axId val="50010002"/>'
  let group: string
  if (d.kind === 'column' || d.kind === 'bar')
    group = `<c:barChart><c:barDir val="${d.kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="150"/>${axes}</c:barChart>`
  else if (d.kind === 'line') group = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axes}</c:lineChart>`
  else if (d.kind === 'area') group = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${axes}</c:areaChart>`
  else if (d.kind === 'pie') group = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`
  else group = `<c:doughnutChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`
  const catPos = d.kind === 'bar' ? 'l' : 'b'
  const valPos = d.kind === 'bar' ? 'b' : 'l'
  const axXml = pie
    ? ''
    : `<c:catAx><c:axId val="50010001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
      `<c:valAx><c:axId val="50010002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
  const title = d.title.trim()
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>${d.title
        .split('\n')
        .map((l) => `<a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>${esc(l)}</a:t></a:r></a:p>`)
        .join('')}</c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>'
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${group}${axXml}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:chartSpace>`
  )
}

/** Inline drawing run that places a chart part (by relationship id) in the text. */
export function chartRunXml(rid: string, id: number, widthPx: number, heightPx: number, rPr = ''): string {
  const cx = Math.max(1, Math.round(widthPx * 9525))
  const cy = Math.max(1, Math.round(heightPx * 9525))
  return (
    `<w:r>${rPr}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="Chart ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" r:id="${rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  )
}

const GROUPS: Record<string, ChartKind> = {
  barChart: 'column', bar3DChart: 'column', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut',
}

const c = (el: Element | null | undefined, name: string) => (el ? children(el, C_NS, name)[0] ?? null : null)
const cs = (el: Element | null | undefined, name: string) => (el ? children(el, C_NS, name) : [])
const val = (el: Element | null) => el?.getAttribute('val') ?? null

/** The cached or literal point list of a tx / cat / val element. */
function cacheOf(container: Element | null): Element | null {
  if (!container) return null
  for (const name of ['strCache', 'numCache', 'strLit', 'numLit']) {
    const direct = c(container, name)
    if (direct) return direct
  }
  for (const ref of ['strRef', 'numRef']) {
    const r = c(container, ref)
    const cache = c(r, 'strCache') ?? c(r, 'numCache')
    if (cache) return cache
  }
  return null
}

function points(container: Element | null): Map<number, string> {
  const out = new Map<number, string>()
  const cache = cacheOf(container)
  if (cache) for (const pt of cs(cache, 'pt')) out.set(Number(pt.getAttribute('idx')), c(pt, 'v')?.textContent ?? '')
  return out
}

function count(container: Element | null): number {
  const n = Number(val(c(cacheOf(container), 'ptCount')))
  return Number.isFinite(n) ? n : 0
}

function richText(el: Element | null): string {
  if (!el) return ''
  return Array.from(el.getElementsByTagNameNS(A_NS, 'p'))
    .map((p) => Array.from(p.getElementsByTagNameNS(A_NS, 't')).map((t) => t.textContent ?? '').join(''))
    .join('\n')
}

export function parseChart(xml: string, part: string): ChartData {
  const doc = parseXml(xml, part)
  const chart = c(doc.documentElement, 'chart')
  const plot = c(chart, 'plotArea')
  const groups = plot ? children(plot).filter((g) => g.namespaceURI === C_NS && g.localName.endsWith('Chart')) : []
  const data: ChartData = { part, kind: 'column', title: richText(c(c(c(chart, 'title'), 'tx'), 'rich')), categories: [], series: [], editable: true }
  if (groups.length !== 1) {
    data.editable = false
    data.reason = groups.length ? 'Combination charts can only be kept, not edited.' : 'This chart has no plot area.'
  }
  const g = groups[0]
  if (!g) return data
  const kind = GROUPS[g.localName]
  if (!kind) {
    data.editable = false
    data.reason = 'Only column, bar, line, area, pie and doughnut charts can be edited.'
  } else data.kind = kind === 'column' && val(c(g, 'barDir')) === 'bar' ? 'bar' : kind
  const sers = cs(g, 'ser')
  let n = 0
  for (const s of sers) n = Math.max(n, count(c(s, 'cat')), count(c(s, 'val')), ...[...points(c(s, 'val')).keys()].map((k) => k + 1))
  const cat = sers.map((s) => c(s, 'cat')).find(Boolean) ?? null
  if (cat && (c(cat, 'multiLvlStrRef') || c(cat, 'multiLvlStrLit'))) {
    data.editable = false
    data.reason = 'Charts with multi-level categories can only be kept, not edited.'
  }
  const catPts = points(cat)
  n = Math.max(n, ...[...catPts.keys()].map((k) => k + 1), 0)
  data.categories = Array.from({ length: n }, (_, i) => catPts.get(i) ?? '')
  data.series = sers.map((s, i) => {
    const tx = c(s, 'tx')
    const name = tx ? (c(tx, 'v')?.textContent ?? points(tx).get(0) ?? '') : `Series ${i + 1}`
    const pts = points(c(s, 'val'))
    return { name, values: Array.from({ length: n }, (_, k) => (pts.has(k) && pts.get(k) !== '' && Number.isFinite(Number(pts.get(k))) ? Number(pts.get(k)) : null)) }
  })
  return data
}

const xmlEsc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

function el(doc: Document, name: string, attrs: Record<string, string> = {}): Element {
  const e = doc.createElementNS(C_NS, `c:${name}`)
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v)
  return e
}

function literal(doc: Document, kind: 'str' | 'num', values: (string | number | null)[]): Element {
  const lit = el(doc, kind === 'str' ? 'strLit' : 'numLit')
  if (kind === 'num') {
    const fc = el(doc, 'formatCode')
    fc.textContent = 'General'
    lit.appendChild(fc)
  }
  lit.appendChild(el(doc, 'ptCount', { val: String(values.length) }))
  values.forEach((v, i) => {
    if (v === null || v === '') return
    const pt = el(doc, 'pt', { idx: String(i) })
    const ve = el(doc, 'v')
    ve.textContent = String(v)
    pt.appendChild(ve)
    lit.appendChild(pt)
  })
  return lit
}

/** Write edited data into a chart part. The data becomes literal values inside the chart. */
export function applyChart(xml: string, data: ChartData): string {
  const doc = parseXml(xml, data.part)
  const root = doc.documentElement
  const chart = c(root, 'chart')
  const plot = c(chart, 'plotArea')
  const g = plot ? children(plot).find((x) => x.namespaceURI === C_NS && x.localName.endsWith('Chart')) : null
  if (!chart || !g) throw new Error('The chart part has no plot area.')
  // the embedded spreadsheet copy would overwrite edited values when the chart is next opened elsewhere
  for (const ext of cs(root, 'externalData')) root.removeChild(ext)
  // title
  const title = c(chart, 'title')
  if (data.title.trim()) {
    const t = title ?? chart.insertBefore(el(doc, 'title'), chart.firstElementChild)
    for (const k of children(t).filter((k) => k.namespaceURI === C_NS && k.localName === 'tx')) t.removeChild(k)
    const tx = el(doc, 'tx')
    const rich = el(doc, 'rich')
    rich.appendChild(doc.createElementNS(A_NS, 'a:bodyPr'))
    rich.appendChild(doc.createElementNS(A_NS, 'a:lstStyle'))
    for (const line of data.title.split('\n')) {
      const p = doc.createElementNS(A_NS, 'a:p')
      const r = doc.createElementNS(A_NS, 'a:r')
      const at = doc.createElementNS(A_NS, 'a:t')
      at.textContent = line
      r.appendChild(at)
      p.appendChild(r)
      rich.appendChild(p)
    }
    tx.appendChild(rich)
    t.insertBefore(tx, t.firstElementChild)
    if (!c(t, 'overlay')) {
      const layoutOrAfter = children(t).find((k) => k.localName !== 'tx' && k.localName !== 'layout') ?? null
      t.insertBefore(el(doc, 'overlay', { val: '0' }), layoutOrAfter)
    }
    const atd = c(chart, 'autoTitleDeleted')
    if (atd) atd.setAttribute('val', '0')
  } else if (title) {
    chart.removeChild(title)
    const atd = c(chart, 'autoTitleDeleted')
    if (atd) atd.setAttribute('val', '1')
    else chart.insertBefore(el(doc, 'autoTitleDeleted', { val: '1' }), chart.firstElementChild)
  }
  // series
  const existing = cs(g, 'ser')
  if (!existing.length) throw new Error('The chart has no data series to edit.')
  const template = existing[existing.length - 1]
  const usedIdx = Array.from(root.getElementsByTagNameNS(C_NS, 'ser')).flatMap((x) => [Number(val(c(x, 'idx'))) || 0, Number(val(c(x, 'order'))) || 0])
  let nextIdx = Math.max(0, ...usedIdx) + 1
  const sers: Element[] = []
  data.series.forEach((s, i) => {
    let ser = existing[i]
    if (!ser) {
      ser = template.cloneNode(true) as Element
      for (const k of children(ser).filter((k) => k.namespaceURI === C_NS && ['spPr', 'dPt', 'dLbls', 'trendline', 'errBars'].includes(k.localName))) ser.removeChild(k)
      c(ser, 'idx')?.setAttribute('val', String(nextIdx))
      c(ser, 'order')?.setAttribute('val', String(nextIdx))
      nextIdx++
      g.insertBefore(ser, sers[sers.length - 1].nextSibling)
    }
    sers.push(ser)
    // name
    const tx = el(doc, 'tx')
    const v = el(doc, 'v')
    v.textContent = s.name
    tx.appendChild(v)
    const oldTx = c(ser, 'tx')
    if (oldTx) ser.replaceChild(tx, oldTx)
    else ser.insertBefore(tx, c(ser, 'order')?.nextSibling ?? null)
    // categories
    const cat = el(doc, 'cat')
    cat.appendChild(literal(doc, 'str', data.categories))
    const oldCat = c(ser, 'cat')
    const valEl = c(ser, 'val')
    if (oldCat) ser.replaceChild(cat, oldCat)
    else ser.insertBefore(cat, valEl)
    // values
    const nv = el(doc, 'val')
    nv.appendChild(literal(doc, 'num', s.values))
    if (valEl) ser.replaceChild(nv, valEl)
    else ser.insertBefore(nv, children(ser).find((k) => ['smooth', 'shape', 'extLst'].includes(k.localName)) ?? null)
  })
  for (const extra of existing.slice(data.series.length)) g.removeChild(extra)
  // pie and doughnut slices added in SCRIBE get an explicit colour like the ones the file already had
  if (g.localName.includes('ie') || g.localName === 'doughnutChart') {
    for (const ser of cs(g, 'ser')) {
      const have = new Set(cs(ser, 'dPt').map((p) => Number(val(c(p, 'idx')))))
      if (!have.size && !c(ser, 'spPr')) continue
      for (let k = 0; k < data.categories.length; k++) {
        if (have.has(k)) continue
        const frag = parseXml(`<x xmlns:c="${C_NS}" xmlns:a="${A_NS}">${dPtXml(k)}</x>`, 'dPt').documentElement.firstElementChild!
        const anchor = c(ser, 'dLbls') ?? c(ser, 'cat') ?? c(ser, 'val')
        ser.insertBefore(doc.importNode(frag, true), anchor)
      }
    }
  }
  return serializeXml(doc.documentElement)
}

// ---------------------------------------------------------------------------
// preview

export const CHART_COLORS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e']

export function chartSvg(input: Pick<ChartData, 'kind' | 'title' | 'categories' | 'series'>, width = 480, height = 280): string {
  // normalise untrusted data before drawing
  const num = (v: unknown) => (v === null || v === undefined || v === '' || !Number.isFinite(Number(v)) ? null : Number(v))
  const d = {
    kind: input.kind,
    title: String(input.title ?? ''),
    categories: (Array.isArray(input.categories) ? input.categories : []).map((x) => String(x ?? '')),
    series: (Array.isArray(input.series) ? input.series : []).map((s) => ({ name: String(s?.name ?? ''), values: (Array.isArray(s?.values) ? s.values : []).map(num) })),
  }
  const W = width
  const H = height
  const top = d.title ? 34 : 14
  const parts: string[] = []
  parts.push(`<rect x="0" y="0" width="${W}" height="${H}" fill="#ffffff" stroke="#d0d5dd"/>`)
  if (d.title) parts.push(`<text x="${W / 2}" y="22" text-anchor="middle" font-size="14" font-weight="600" fill="#1f2937">${xmlEsc(d.title)}</text>`)
  const n = d.categories.length
  const legendH = 22
  if (d.kind === 'pie' || d.kind === 'doughnut') {
    const s = d.series[0]
    const vals = (s?.values ?? []).map((v) => Math.max(0, v ?? 0))
    const total = vals.reduce((a, b) => a + b, 0)
    const r = Math.max(10, Math.min(W / 2 - 20, (H - top - legendH - 10) / 2))
    const cx = W / 2
    const cy = top + r + 4
    let a0 = -Math.PI / 2
    vals.forEach((v, i) => {
      if (!total || !v) return
      const a1 = a0 + (v / total) * Math.PI * 2
      const large = a1 - a0 > Math.PI ? 1 : 0
      const p = (a: number) => `${(cx + r * Math.cos(a)).toFixed(1)} ${(cy + r * Math.sin(a)).toFixed(1)}`
      parts.push(v === total ? `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${CHART_COLORS[i % 8]}"/>` : `<path d="M${cx} ${cy} L${p(a0)} A${r} ${r} 0 ${large} 1 ${p(a1)} Z" fill="${CHART_COLORS[i % 8]}" stroke="#fff"/>`)
      a0 = a1
    })
    if (d.kind === 'doughnut') parts.push(`<circle cx="${cx}" cy="${cy}" r="${r * 0.5}" fill="#fff"/>`)
    parts.push(legend(d.categories, W, H))
  } else {
    const all = d.series.flatMap((s) => s.values.filter((v): v is number => v != null))
    const max = Math.max(0, ...all)
    const min = Math.min(0, ...all)
    const span = max - min || 1
    const left = 40
    const right = W - 12
    const bottom = H - legendH - 22
    const scale = (v: number) => bottom - ((v - min) / span) * (bottom - top)
    for (let k = 0; k <= 4; k++) {
      const v = min + (span * k) / 4
      const y = scale(v)
      parts.push(`<line x1="${left}" y1="${y.toFixed(1)}" x2="${right}" y2="${y.toFixed(1)}" stroke="#eceef2"/>`, `<text x="${left - 4}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="10" fill="#6b7280">${fmt(v)}</text>`)
    }
    if (d.kind === 'bar') {
      // horizontal bars
      const band = (bottom - top) / Math.max(1, n)
      const bh = (band * 0.7) / Math.max(1, d.series.length)
      const sx = (v: number) => left + ((v - min) / span) * (right - left)
      d.series.forEach((s, si) =>
        s.values.forEach((v, i) => {
          if (v == null) return
          const y = top + i * band + band * 0.15 + si * bh
          const x0 = sx(Math.max(min, 0))
          const x1 = sx(v)
          parts.push(`<rect x="${Math.min(x0, x1).toFixed(1)}" y="${y.toFixed(1)}" width="${Math.abs(x1 - x0).toFixed(1)}" height="${bh.toFixed(1)}" fill="${CHART_COLORS[si % 8]}"/>`)
        }),
      )
      d.categories.forEach((cat, i) => parts.push(`<text x="${left - 4}" y="${(top + i * band + band / 2 + 4).toFixed(1)}" text-anchor="end" font-size="10" fill="#374151">${xmlEsc(cat.slice(0, 10))}</text>`))
    } else {
      const band = (right - left) / Math.max(1, n)
      d.categories.forEach((cat, i) => parts.push(`<text x="${(left + i * band + band / 2).toFixed(1)}" y="${bottom + 14}" text-anchor="middle" font-size="10" fill="#374151">${xmlEsc(cat.slice(0, 14))}</text>`))
      if (d.kind === 'column') {
        const bw = (band * 0.7) / Math.max(1, d.series.length)
        d.series.forEach((s, si) =>
          s.values.forEach((v, i) => {
            if (v == null) return
            const x = left + i * band + band * 0.15 + si * bw
            const y0 = scale(0)
            const y1 = scale(v)
            parts.push(`<rect x="${x.toFixed(1)}" y="${Math.min(y0, y1).toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.abs(y1 - y0).toFixed(1)}" fill="${CHART_COLORS[si % 8]}"/>`)
          }),
        )
      } else {
        d.series.forEach((s, si) => {
          const pts = s.values.map((v, i) => (v == null ? null : `${(left + i * band + band / 2).toFixed(1)},${scale(v).toFixed(1)}`))
          const segs = pts.filter(Boolean) as string[]
          if (!segs.length) return
          if (d.kind === 'area') {
            const first = segs[0].split(',')[0]
            const last = segs[segs.length - 1].split(',')[0]
            parts.push(`<polygon points="${first},${scale(Math.max(min, 0)).toFixed(1)} ${segs.join(' ')} ${last},${scale(Math.max(min, 0)).toFixed(1)}" fill="${CHART_COLORS[si % 8]}" fill-opacity="0.55"/>`)
          }
          parts.push(`<polyline points="${segs.join(' ')}" fill="none" stroke="${CHART_COLORS[si % 8]}" stroke-width="2.5"/>`)
        })
      }
    }
    parts.push(`<line x1="${left}" y1="${bottom}" x2="${right}" y2="${bottom}" stroke="#9ca3af"/>`)
    parts.push(legend(d.series.map((s) => s.name), W, H))
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" font-family="Calibri, Carlito, Arial, sans-serif" role="img" aria-label="${xmlEsc(d.title || 'Chart')}">${parts.join('')}</svg>`
}

function legend(names: string[], W: number, H: number) {
  const items = names.slice(0, 8)
  const w = Math.min(120, (W - 20) / Math.max(1, items.length))
  const x0 = (W - w * items.length) / 2
  return items
    .map((nm, i) => `<rect x="${(x0 + i * w).toFixed(1)}" y="${H - 16}" width="10" height="10" fill="${CHART_COLORS[i % 8]}"/><text x="${(x0 + i * w + 14).toFixed(1)}" y="${H - 7}" font-size="10" fill="#374151">${xmlEsc(nm.slice(0, 16))}</text>`)
    .join('')
}

function fmt(v: number) {
  const a = Math.abs(v)
  if (a >= 1e6) return (v / 1e6).toFixed(1).replace(/\.0$/, '') + 'M'
  if (a >= 1e3) return (v / 1e3).toFixed(1).replace(/\.0$/, '') + 'k'
  return String(Math.round(v * 100) / 100)
}
