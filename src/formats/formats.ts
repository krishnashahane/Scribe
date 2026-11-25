import DOMPurify from 'dompurify'
import type { PMNode, ScribeDocument } from '../model/types'
import { twipToPx } from '../model/types'

/** Plain text → document (one paragraph per line; blank lines kept). */
export function textToContent(text: string): PMNode {
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  return {
    type: 'doc',
    content: lines.map((l) => (l ? { type: 'paragraph', content: [{ type: 'text', text: l }] } : { type: 'paragraph' })),
  }
}

export function contentToText(n: PMNode): string {
  const out: string[] = []
  const inline = (x: PMNode): string => {
    if (x.type === 'text') return x.text ?? ''
    if (x.type === 'hardBreak') return '\n'
    if (x.type === 'field') return x.attrs?.result ?? ''
    if (x.type === 'opaqueInline') return x.attrs?.preview ?? ''
    if (x.type === 'noteRef') return `[${x.attrs?.label ?? '*'}]`
    return (x.content ?? []).map(inline).join('')
  }
  const block = (x: PMNode, depth = 0) => {
    switch (x.type) {
      case 'paragraph':
      case 'heading':
        out.push(inline(x))
        break
      case 'bulletList':
      case 'orderedList':
        ;(x.content ?? []).forEach((li, i) => {
          const bullet = x.type === 'bulletList' ? '•' : `${(x.attrs?.start ?? 1) + i}.`
          const [first, ...rest] = li.content ?? []
          out.push(`${'  '.repeat(depth)}${bullet} ${first ? inline(first) : ''}`)
          rest.forEach((r) => block(r, depth + 1))
        })
        break
      case 'table':
        for (const row of x.content ?? []) out.push((row.content ?? []).map((c) => (c.content ?? []).map(inline).join(' ')).join('\t'))
        break
      case 'pageBreak':
        out.push('\f')
        break
      case 'opaqueBlock':
        out.push(x.attrs?.preview ?? '')
        break
      case 'sectionBreak':
        break
      default:
        ;(x.content ?? []).forEach((c) => block(c, depth))
    }
  }
  ;(n.content ?? []).forEach((b) => block(b))
  return out.join('\n')
}

const PURIFY = {
  ALLOWED_TAGS: [
    'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'br', 'strong', 'b', 'em', 'i', 'u', 's', 'strike', 'del', 'ins', 'sub', 'sup', 'span', 'a',
    'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tr', 'td', 'th', 'colgroup', 'col', 'img', 'blockquote', 'pre', 'code', 'hr', 'div', 'mark', 'font',
  ],
  ALLOWED_ATTR: ['href', 'src', 'alt', 'title', 'style', 'colspan', 'rowspan', 'width', 'height', 'start', 'data-page-break', 'color', 'face', 'size', 'dir'],
  ALLOWED_URI_REGEXP: /^(?:(?:https?|mailto|tel):|#|data:image\/(?:png|jpe?g|gif|webp);base64,)/i,
}

/** Sanitise untrusted HTML (imports, paste) before it reaches the editor. */
export function sanitizeHtml(html: string): string {
  // explicit CSS page breaks in pasted or imported HTML become SCRIBE page breaks
  html = html.replace(/<br[^>]*style="[^"]*page-break-before:\s*always[^"]*"[^>]*>/gi, '<hr data-page-break>')
  return DOMPurify.sanitize(html, PURIFY) as unknown as string
}

/** Standalone HTML export with print-ready page CSS. */
export function documentToHtml(bodyHtml: string, doc: ScribeDocument, title: string): string {
  const p = doc.meta.sections[0].page
  const d = doc.meta.defaults
  const css = `
@page { size: ${(p.width / 1440).toFixed(2)}in ${(p.height / 1440).toFixed(2)}in; margin: ${(p.marginTop / 1440).toFixed(2)}in ${(p.marginRight / 1440).toFixed(2)}in ${(p.marginBottom / 1440).toFixed(2)}in ${(p.marginLeft / 1440).toFixed(2)}in; }
body { font-family: "${d.fontFamily}", Calibri, Carlito, sans-serif; font-size: ${d.fontSizePt}pt; line-height: ${d.lineHeight}; max-width: ${twipToPx(p.width - p.marginLeft - p.marginRight)}px; margin: 2rem auto; color: #1b1b1b; }
p { margin: 0 0 0.6em; } table { border-collapse: collapse; } td, th { border: 1px solid #999; padding: 4px 6px; vertical-align: top; }
img { max-width: 100%; } .sc-page-break, [data-scribe="pageBreak"] { break-after: page; height: 0; overflow: hidden; color: transparent; }
.sc-section-break { display: none; } .sc-opaque { border: 1px dashed #aaa; padding: 4px; color: #555; } .sc-ins { color: #1a7f37; } .sc-del { color: #b42318; }
`
  return `<!DOCTYPE html>\n<html lang="en"><head><meta charset="utf-8"><meta name="generator" content="SCRIBE"><title>${escapeHtml(title)}</title><style>${css}</style></head><body>\n${bodyHtml}\n</body></html>\n`
}

export function escapeHtml(s: string) {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!)
}

export function download(data: BlobPart, filename: string, mime: string) {
  const blob = new Blob([data], { type: mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 5000)
}

export function safeFilename(name: string, ext: string) {
  const cleaned = [...(name || 'Untitled')].map((ch) => (ch.charCodeAt(0) < 0x20 || '\\/:*?"<>|'.includes(ch) ? ' ' : ch)).join('')
  const base = cleaned.replace(/ {2,}/g, ' ').trim().slice(0, 120) || 'Untitled'
  return `${base}.${ext}`
}
