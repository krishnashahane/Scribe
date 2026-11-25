// Text boxes: read and rewrite the paragraphs inside <w:txbxContent>.
// A text box usually appears twice (modern shape + legacy VML fallback); both copies are kept in sync.

import { NS, children, parseXml, serializeXml } from '../docx/xml'

const W = NS.w

function contents(doc: Document): Element[] {
  return Array.from(doc.getElementsByTagNameNS(W, 'txbxContent'))
}

function paraText(p: Element): string {
  let s = ''
  const walk = (e: Element) => {
    for (const c of children(e)) {
      if (c.namespaceURI === W && (c.localName === 't')) s += c.textContent ?? ''
      else if (c.namespaceURI === W && c.localName === 'tab') s += '\t'
      else if (c.namespaceURI === W && (c.localName === 'br' || c.localName === 'cr')) s += '\n'
      else if (!(c.namespaceURI === W && (c.localName === 'pPr' || c.localName === 'rPr' || c.localName === 'instrText' || c.localName === 'delText'))) walk(c)
    }
  }
  walk(p)
  return s
}

/** One string per paragraph of the first text box copy. */
export function textBoxParagraphs(xml: string): string[] | null {
  const doc = parseXml(xml, 'text box')
  const first = contents(doc)[0]
  if (!first) return null
  return children(first, W, 'p').map(paraText)
}

/** True when the text box holds only simple paragraphs that a plain-text edit can rewrite without loss. */
export function textBoxIsSimple(xml: string): boolean {
  const doc = parseXml(xml, 'text box')
  for (const c of contents(doc)) {
    for (const k of children(c)) if (!(k.namespaceURI === W && k.localName === 'p')) return false
    for (const r of Array.from(c.getElementsByTagNameNS(W, 'p'))) {
      for (const k of children(r)) {
        if (k.namespaceURI !== W) return false
        if (k.localName === 'pPr' || k.localName === 'bookmarkStart' || k.localName === 'bookmarkEnd' || k.localName === 'proofErr') continue
        if (k.localName !== 'r') return false
        for (const x of children(k)) if (!(x.namespaceURI === W && ['rPr', 't', 'tab', 'br', 'lastRenderedPageBreak'].includes(x.localName))) return false
      }
    }
  }
  return true
}

/**
 * Replace the text of every text box copy. Each line becomes a paragraph; the paragraph and run
 * properties of the original first paragraph (or the matching paragraph) are kept.
 */
export function setTextBoxParagraphs(xml: string, lines: string[]): string {
  const doc = parseXml(xml, 'text box')
  const boxes = contents(doc)
  if (!boxes.length) throw new Error('This object has no text box content.')
  if (!lines.length) lines = [''] // a text box must hold at least one paragraph
  for (const box of boxes) {
    const old = children(box, W, 'p')
    for (const p of old) box.removeChild(p)
    const others = children(box)
    lines.forEach((line, i) => {
      const tpl = old[Math.min(i, old.length - 1)]
      const p = doc.createElementNS(W, 'w:p')
      const pPr = tpl && children(tpl, W, 'pPr')[0]
      if (pPr) p.appendChild(pPr.cloneNode(true))
      const firstRun = tpl && Array.from(tpl.getElementsByTagNameNS(W, 'r'))[0]
      const rPr = firstRun && children(firstRun, W, 'rPr')[0]
      const parts = line.split('\t')
      if (line) {
        const r = doc.createElementNS(W, 'w:r')
        if (rPr) r.appendChild(rPr.cloneNode(true))
        parts.forEach((txt, k) => {
          if (k > 0) r.appendChild(doc.createElementNS(W, 'w:tab'))
          if (txt) {
            const t = doc.createElementNS(W, 'w:t')
            t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
            t.textContent = txt
            r.appendChild(t)
          }
        })
        p.appendChild(r)
      }
      if (others.length) box.insertBefore(p, others[0])
      else box.appendChild(p)
    })
  }
  return serializeXml(doc.documentElement)
}
