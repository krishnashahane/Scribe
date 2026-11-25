// Drawing shapes: preset DrawingML geometries SCRIBE can draw, export and re-import.

import { NS, children, serializeXml } from '../docx/xml'

export const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
const A = NS.a

export const SHAPE_KINDS = {
  rect: 'Rectangle',
  roundRect: 'Rounded rectangle',
  ellipse: 'Oval',
  triangle: 'Triangle',
  diamond: 'Diamond',
  rightArrow: 'Arrow',
  star5: 'Star',
  line: 'Line',
} as const

export type ShapeKind = keyof typeof SHAPE_KINDS

export interface ShapeAttrs {
  kind: ShapeKind
  /** CSS pixels */
  width: number
  height: number
  /** #rrggbb or null for no fill */
  fill: string | null
  stroke: string | null
  /** CSS pixels */
  strokeWidth: number
  text: string
  /** for lines: draw an arrow head at the end */
  arrow?: boolean
}

export const DEFAULT_SHAPE: Omit<ShapeAttrs, 'kind'> = { width: 160, height: 96, fill: '#4472c4', stroke: '#2f528f', strokeWidth: 1.5, text: '', arrow: false }

const hex = (c: string | null) => (c && /^#[0-9a-f]{6}$/i.test(c) ? c : null)
const xmlEsc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/** SVG preview of a shape (all values are numbers or validated colours). */
export function shapeSvg(a: ShapeAttrs): string {
  const w = Math.max(4, Math.round(a.width))
  const h = Math.max(a.kind === 'line' ? 2 : 4, Math.round(a.height))
  const sw = Math.max(0, Math.min(20, Number(a.strokeWidth) || 0))
  const fill = a.kind === 'line' ? 'none' : hex(a.fill) ?? 'none'
  const stroke = hex(a.stroke) ?? 'none'
  const i = sw / 2
  let body = ''
  const pts = (p: [number, number][]) => p.map(([x, y]) => `${x.toFixed(1)},${y.toFixed(1)}`).join(' ')
  switch (a.kind) {
    case 'rect':
      body = `<rect x="${i}" y="${i}" width="${w - sw}" height="${h - sw}"/>`
      break
    case 'roundRect':
      body = `<rect x="${i}" y="${i}" width="${w - sw}" height="${h - sw}" rx="${Math.min(w, h) * 0.1667}"/>`
      break
    case 'ellipse':
      body = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2 - i}" ry="${h / 2 - i}"/>`
      break
    case 'triangle':
      body = `<polygon points="${pts([[w / 2, i], [w - i, h - i], [i, h - i]])}"/>`
      break
    case 'diamond':
      body = `<polygon points="${pts([[w / 2, i], [w - i, h / 2], [w / 2, h - i], [i, h / 2]])}"/>`
      break
    case 'rightArrow': {
      const head = Math.min(w * 0.5, h)
      body = `<polygon points="${pts([[i, h * 0.25], [w - head, h * 0.25], [w - head, i], [w - i, h / 2], [w - head, h - i], [w - head, h * 0.75], [i, h * 0.75]])}"/>`
      break
    }
    case 'star5': {
      const p: [number, number][] = []
      for (let k = 0; k < 10; k++) {
        const r = k % 2 ? 0.382 : 1
        const ang = -Math.PI / 2 + (k * Math.PI) / 5
        p.push([w / 2 + (w / 2 - i) * r * Math.cos(ang), h / 2 + (h / 2 - i) * r * Math.sin(ang) + h * 0.05])
      }
      body = `<polygon points="${pts(p)}"/>`
      break
    }
    case 'line':
      body = `<line x1="${i}" y1="${h / 2}" x2="${w - i - (a.arrow ? sw * 2 : 0)}" y2="${h / 2}"/>` + (a.arrow ? `<polygon fill="${stroke}" stroke="none" points="${pts([[w - sw * 4 - 2, h / 2 - sw * 2 - 2], [w, h / 2], [w - sw * 4 - 2, h / 2 + sw * 2 + 2]])}"/>` : '')
      break
  }
  // centred lines of text, coloured for contrast with the fill (the same rule the DOCX export uses)
  const lines = a.text && a.kind !== 'line' ? String(a.text).split('\n') : []
  const color = hex(a.fill) && luminance(hex(a.fill)!) < 0.5 ? '#ffffff' : '#1f2937'
  const lh = 18
  const text = lines.length
    ? `<text x="${w / 2}" text-anchor="middle" font-family="Calibri, Carlito, Arial, sans-serif" font-size="14" fill="${color}">${lines.map((l, k) => `<tspan x="${w / 2}" y="${(h / 2 - ((lines.length - 1) * lh) / 2 + k * lh + 5).toFixed(1)}">${xmlEsc(l)}</tspan>`).join('')}</text>`
    : ''
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" overflow="visible"><g fill="${fill}" stroke="${stroke}" stroke-width="${sw}">${body}</g>${text}</svg>`
}

const emu = (px: number) => Math.max(1, Math.round(px * 9525))
const srgb = (c: string) => `<a:srgbClr val="${c.slice(1).toUpperCase()}"/>`

/** Inline DrawingML shape run for DOCX export. */
export function shapeRunXml(a: ShapeAttrs, id: number, rPr = ''): string {
  const cx = emu(a.width)
  // a line's box is flat: a non-zero height would tilt it from corner to corner
  const cy = a.kind === 'line' ? 0 : emu(a.height)
  const fill = a.kind === 'line' ? '<a:noFill/>' : hex(a.fill) ? `<a:solidFill>${srgb(hex(a.fill)!)}</a:solidFill>` : '<a:noFill/>'
  const sw = Math.round(Math.max(0, a.strokeWidth) * 9525)
  const ln = hex(a.stroke) && sw ? `<a:ln w="${sw}"><a:solidFill>${srgb(hex(a.stroke)!)}</a:solidFill>${a.kind === 'line' && a.arrow ? '<a:tailEnd type="triangle"/>' : ''}</a:ln>` : '<a:ln><a:noFill/></a:ln>'
  const textColor = hex(a.fill) && luminance(hex(a.fill)!) < 0.5 ? 'FFFFFF' : '1F2937'
  const txbx =
    a.text && a.kind !== 'line'
      ? `<wps:txbx><w:txbxContent>${a.text
          .split('\n')
          .map((line) => `<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr>${line ? `<w:r><w:rPr><w:color w:val="${textColor}"/></w:rPr><w:t xml:space="preserve">${xmlEsc(line)}</w:t></w:r>` : ''}</w:p>`)
          .join('')}</w:txbxContent></wps:txbx>`
      : ''
  const name = `${SHAPE_KINDS[a.kind]} ${id}`
  return (
    `<w:r>${rPr}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="${xmlEsc(name)}"/><wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp><wps:cNvSpPr${a.text ? ' txBox="0"' : ''}/>` +
    `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${a.kind}"><a:avLst/></a:prstGeom>${fill}${ln}</wps:spPr>` +
    `${txbx}<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  )
}

function luminance(c: string) {
  const n = parseInt(c.slice(1), 16)
  return (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255
}

/**
 * Recognise an inline shape that uses only what SCRIBE can reproduce exactly (preset geometry,
 * solid fill, solid outline, plain centred text). Anything richer stays preserved verbatim.
 */
export function parseInlineShape(inline: Element): ShapeAttrs | null {
  const ext = children(inline, NS.wp, 'extent')[0]
  const gd = inline.getElementsByTagNameNS(A, 'graphicData')[0]
  if (!ext || !gd || gd.getAttribute('uri') !== WPS_NS) return null
  const wsp = children(gd, WPS_NS, 'wsp')[0]
  if (!wsp) return null
  // spacing, alt text and text-body settings must be the ones SCRIBE writes, or the shape stays preserved verbatim
  if (['distT', 'distB', 'distL', 'distR'].some((k) => (inline.getAttribute(k) ?? '0') !== '0')) return null
  const docPr = children(inline, NS.wp, 'docPr')[0]
  if (docPr && (docPr.getAttribute('descr') || docPr.getAttribute('title') || docPr.firstElementChild)) return null
  const bodyPr = children(wsp, WPS_NS, 'bodyPr')[0]
  const BODY: Record<string, string> = { rot: '0', vert: 'horz', wrap: 'square', lIns: '91440', tIns: '45720', rIns: '91440', bIns: '45720', anchor: 'ctr', anchorCtr: '0' }
  if (bodyPr) {
    if ([...bodyPr.attributes].some((x) => BODY[x.name] !== x.value)) return null
    if (children(bodyPr).some((k) => k.localName !== 'noAutofit')) return null
  }
  const allowed = new Set(['cNvSpPr', 'spPr', 'txbx', 'bodyPr'])
  if (children(wsp).some((k) => !(k.namespaceURI === WPS_NS && allowed.has(k.localName)))) return null
  const spPr = children(wsp, WPS_NS, 'spPr')[0]
  const geom = spPr && children(spPr, A, 'prstGeom')[0]
  const kind = geom?.getAttribute('prst') as ShapeKind | undefined
  if (!kind || !(kind in SHAPE_KINDS)) return null
  if (children(geom!, A, 'avLst')[0]?.firstElementChild) return null
  const spKids = children(spPr!).map((k) => k.localName)
  if (spKids.some((k) => !['xfrm', 'prstGeom', 'solidFill', 'noFill', 'ln'].includes(k))) return null
  const colorOf = (el: Element | undefined) => {
    if (!el) return undefined
    const kidsOf = children(el)
    if (kidsOf.length !== 1 || kidsOf[0].localName !== 'srgbClr' || kidsOf[0].firstElementChild) return undefined
    const v = kidsOf[0].getAttribute('val') ?? ''
    return /^[0-9A-Fa-f]{6}$/.test(v) ? '#' + v.toLowerCase() : undefined
  }
  const sf = children(spPr!, A, 'solidFill')[0]
  const fill = sf ? colorOf(sf) : null
  if (fill === undefined) return null
  const ln = children(spPr!, A, 'ln')[0]
  let stroke: string | null = null
  let strokeWidth = 0
  let arrow = false
  if (ln) {
    const lk = children(ln).map((k) => k.localName)
    if (lk.some((k) => !['solidFill', 'noFill', 'tailEnd'].includes(k))) return null
    if ([...ln.attributes].some((x) => x.name !== 'w')) return null
    const lsf = children(ln, A, 'solidFill')[0]
    if (lsf) {
      const col = colorOf(lsf)
      if (col === undefined) return null
      stroke = col
      strokeWidth = Math.round((Number(ln.getAttribute('w') || 9525) / 9525) * 100) / 100
    }
    const tail = children(ln, A, 'tailEnd')[0]
    if (tail) {
      if (tail.getAttribute('type') !== 'triangle' || kind !== 'line') return null
      arrow = true
    }
  }
  let text = ''
  const txbx = children(wsp, WPS_NS, 'txbx')[0]
  if (txbx) {
    const content = children(txbx, NS.w, 'txbxContent')[0]
    if (!content) return null
    const lines: string[] = []
    for (const p of children(content)) {
      if (!(p.namespaceURI === NS.w && p.localName === 'p')) return null
      let line = ''
      for (const k of children(p)) {
        if (k.localName === 'pPr') {
          if (children(k).some((x) => !['spacing', 'jc'].includes(x.localName))) return null
          if (children(k, NS.w, 'jc')[0] && children(k, NS.w, 'jc')[0].getAttributeNS(NS.w, 'val') !== 'center') return null
          continue
        }
        if (k.localName !== 'r') return null
        for (const x of children(k)) {
          if (x.localName === 'rPr') {
            if (children(x).some((y) => y.localName !== 'color')) return null
            continue
          }
          if (x.localName !== 't') return null
          line += x.textContent ?? ''
        }
      }
      lines.push(line)
    }
    text = lines.join('\n')
  }
  const px = (v: string | null) => Math.round((Number(v) || 0) / 9525)
  const height = kind === 'line' ? Math.max(4, px(ext.getAttribute('cy'))) : px(ext.getAttribute('cy'))
  return { kind, width: px(ext.getAttribute('cx')), height, fill: kind === 'line' ? null : fill, stroke, strokeWidth, text, arrow }
}

export { serializeXml }
