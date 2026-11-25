// SmartArt diagrams: edit the text of diagram nodes in the data part and keep the
// cached drawing (what most readers display) in sync.

import { children, parseXml, serializeXml } from '../docx/xml'

export const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
export const DSP_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

export interface SmartArtItem {
  id: string
  text: string
  depth: number
}

export interface SmartArtData {
  dataPart: string
  drawingPart?: string
  items: SmartArtItem[]
  dirty?: boolean
}

const kids = (el: Element | null | undefined, ns: string, name: string) => (el ? children(el, ns, name) : [])
const kid = (el: Element | null | undefined, ns: string, name: string) => kids(el, ns, name)[0] ?? null

function bodyText(body: Element | null): string {
  if (!body) return ''
  return kids(body, A_NS, 'p')
    .map((p) => Array.from(p.getElementsByTagNameNS(A_NS, 't')).map((t) => t.textContent ?? '').join(''))
    .join('\n')
}

const isDataPoint = (pt: Element) => {
  const t = pt.getAttribute('type') ?? 'node'
  return t === 'node' || t === 'asst'
}

/** The drawing part's relationship id, stored in the data part's extension list. */
export function drawingRelId(dataXml: string): string | null {
  const doc = parseXml(dataXml, 'diagram data')
  const ext = doc.getElementsByTagNameNS(DSP_NS, 'dataModelExt')[0]
  return ext?.getAttribute('relId') ?? null
}

export function parseSmartArt(dataXml: string, dataPart: string, drawingPart?: string): SmartArtData {
  const doc = parseXml(dataXml, dataPart)
  const root = doc.documentElement
  const pts = kids(kid(root, DGM_NS, 'ptLst'), DGM_NS, 'pt')
  const parent = new Map<string, string>()
  for (const cx of kids(kid(root, DGM_NS, 'cxnLst'), DGM_NS, 'cxn')) {
    const type = cx.getAttribute('type') ?? 'parOf'
    if (type === 'parOf') parent.set(cx.getAttribute('destId') ?? '', cx.getAttribute('srcId') ?? '')
  }
  const depthOf = (id: string) => {
    let d = 0
    let cur = parent.get(id)
    const seen = new Set<string>()
    while (cur && !seen.has(cur)) {
      seen.add(cur)
      const p = pts.find((x) => x.getAttribute('modelId') === cur)
      if (!p || p.getAttribute('type') === 'doc') break
      d++
      cur = parent.get(cur)
    }
    return d
  }
  const items = pts.filter(isDataPoint).map((pt) => {
    const id = pt.getAttribute('modelId') ?? ''
    return { id, text: bodyText(kid(pt, DGM_NS, 't')), depth: depthOf(id) }
  })
  return { dataPart, drawingPart, items }
}

function setBody(doc: Document, body: Element, text: string) {
  const old = kids(body, A_NS, 'p')
  const tpl = old[0] ?? null
  for (const p of old) body.removeChild(p)
  const lines = text.split('\n')
  for (const line of lines) {
    const p = doc.createElementNS(A_NS, 'a:p')
    const pPr = tpl && kid(tpl, A_NS, 'pPr')
    if (pPr) p.appendChild(pPr.cloneNode(true))
    const firstRun = tpl && kid(tpl, A_NS, 'r')
    const rPr = firstRun && kid(firstRun, A_NS, 'rPr')
    if (line) {
      const r = doc.createElementNS(A_NS, 'a:r')
      if (rPr) r.appendChild(rPr.cloneNode(true))
      const t = doc.createElementNS(A_NS, 'a:t')
      t.textContent = line
      r.appendChild(t)
      p.appendChild(r)
    }
    const end = tpl && kid(tpl, A_NS, 'endParaRPr')
    if (end) p.appendChild(end.cloneNode(true))
    body.appendChild(p)
  }
}

/** Write edited node text into the data part. */
export function applySmartArtData(dataXml: string, items: SmartArtItem[]): string {
  const doc = parseXml(dataXml, 'diagram data')
  const pts = kids(kid(doc.documentElement, DGM_NS, 'ptLst'), DGM_NS, 'pt')
  for (const it of items) {
    const pt = pts.find((p) => p.getAttribute('modelId') === it.id)
    if (!pt) continue
    let t = kid(pt, DGM_NS, 't')
    if (!t) {
      t = doc.createElementNS(DGM_NS, 'dgm:t')
      t.appendChild(doc.createElementNS(A_NS, 'a:bodyPr'))
      t.appendChild(doc.createElementNS(A_NS, 'a:lstStyle'))
      pt.insertBefore(t, kid(pt, DGM_NS, 'extLst'))
    }
    if (bodyText(t) !== it.text) setBody(doc, t, it.text)
  }
  return serializeXml(doc.documentElement)
}

/** Write edited node text into the cached drawing, using the original data part to map shapes to nodes. */
export function applySmartArtDrawing(drawingXml: string, originalDataXml: string, items: SmartArtItem[]): string {
  const data = parseXml(originalDataXml, 'diagram data')
  const presToNode = new Map<string, string>()
  for (const pt of kids(kid(data.documentElement, DGM_NS, 'ptLst'), DGM_NS, 'pt')) {
    if (pt.getAttribute('type') !== 'pres') continue
    const assoc = kid(pt, DGM_NS, 'prSet')?.getAttribute('presAssocID')
    if (assoc) presToNode.set(pt.getAttribute('modelId') ?? '', assoc)
  }
  const byId = new Map(items.map((i) => [i.id, i.text]))
  const doc = parseXml(drawingXml, 'diagram drawing')
  for (const sp of Array.from(doc.getElementsByTagNameNS(DSP_NS, 'sp'))) {
    const mid = sp.getAttribute('modelId') ?? ''
    const node = byId.has(mid) ? mid : presToNode.get(mid)
    if (!node || !byId.has(node)) continue
    const body = kid(sp, DSP_NS, 'txBody')
    if (!body) continue
    const text = byId.get(node)!
    if (bodyText(body) !== text) setBody(doc, body, text)
  }
  return serializeXml(doc.documentElement)
}
