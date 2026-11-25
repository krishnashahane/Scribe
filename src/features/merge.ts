// Mail merge: recipients from CSV, MERGEFIELD fields in the document, and merged output.

import type { DocMeta, MailMergeData, PMNode } from '../model/types'

export const MAX_RECORDS = 2000
export const MAX_MERGE_CHARS = 600_000

/** RFC 4180 CSV (quoted fields, doubled quotes, CRLF/LF). Also accepts ; or tab separated files. */
export function parseCsv(text: string): string[][] {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1)
  const firstLine = text.slice(0, text.search(/\r?\n|$/))
  const counts = [',', ';', '\t'].map((d) => ({ d, n: splitCount(firstLine, d) }))
  const delim = counts.sort((a, b) => b.n - a.n)[0].n > 0 ? counts[0].d : ','
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  let i = 0
  while (i < text.length) {
    const ch = text[i]
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i += 2
          continue
        }
        quoted = false
        i++
        continue
      }
      field += ch
      i++
      continue
    }
    if (ch === '"' && field === '') {
      quoted = true
      i++
      continue
    }
    if (ch === delim) {
      row.push(field)
      field = ''
      i++
      continue
    }
    if (ch === '\r' || ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1
      continue
    }
    field += ch
    i++
  }
  if (quoted) throw new Error('The CSV file has an unclosed quotation mark.')
  if (field !== '' || row.length) {
    row.push(field)
    rows.push(row)
  }
  return rows.filter((r) => r.some((c) => c.trim() !== ''))
}

function splitCount(line: string, d: string) {
  let n = 0
  let q = false
  for (const ch of line) {
    if (ch === '"') q = !q
    else if (ch === d && !q) n++
  }
  return n
}

/** Field names usable in MERGEFIELD instructions (letters, digits, underscore). */
export function fieldName(header: string, used: Set<string>): string {
  let base = header.trim().replace(/\s+/g, '_').replace(/[^\p{L}\p{N}_]/gu, '').slice(0, 40) || 'Field'
  if (/^\p{N}/u.test(base)) base = 'F' + base
  let name = base
  for (let i = 2; used.has(name.toLowerCase()); i++) name = `${base}_${i}`
  used.add(name.toLowerCase())
  return name
}

export function recipientsFromCsv(text: string, name: string): MailMergeData {
  const rows = parseCsv(text)
  if (rows.length < 2) throw new Error('The CSV file needs a header row and at least one recipient.')
  const used = new Set<string>()
  const fields = rows[0].map((h) => fieldName(h, used))
  const records = rows.slice(1).map((r) => fields.map((_, i) => r[i] ?? ''))
  if (records.length > MAX_RECORDS) throw new Error(`Mail merge supports up to ${MAX_RECORDS.toLocaleString()} recipients per document.`)
  const data = { name, fields, records }
  if (JSON.stringify(data).length > MAX_MERGE_CHARS) throw new Error('The recipient list is too large (over 600 KB of text).')
  return data
}

export const mergeFieldInstr = (name: string) => `MERGEFIELD ${name}`

export function mergeFieldName(instr: string): string | null {
  const m = /^\s*MERGEFIELD\s+("?)([^"\\\s]+)\1/i.exec(instr)
  return m ? m[2] : null
}

/** Replace merge fields with a record's values. Unknown fields keep their «Name» placeholder. */
export function fillRecord(node: PMNode, fields: string[], record: string[]): PMNode {
  if (node.type === 'field') {
    const name = mergeFieldName(String(node.attrs?.instr ?? ''))
    if (name) {
      const i = fields.findIndex((f) => f.toLowerCase() === name.toLowerCase())
      const text = i >= 0 ? record[i] ?? '' : `«${name}»`
      return text ? { type: 'text', text, marks: node.marks } : { type: 'text', text: '​', marks: node.marks }
    }
  }
  if (!node.content) return node
  const content = node.content.map((c) => fillRecord(c, fields, record))
  // empty text nodes are invalid: drop zero-width placeholders created above
  return { ...node, content: content.filter((c) => !(c.type === 'text' && c.text === '​')) }
}

/** Mark-safe deep copy that renames bookmarks and drops comment marks so merged copies stay valid. */
function prepareCopy(node: PMNode, copy: number, keepNotes: boolean): PMNode | null {
  if (node.type === 'bookmark' && copy > 0) return { ...node, attrs: { ...node.attrs, name: `${String(node.attrs?.name ?? 'bm').slice(0, 34)}_${copy + 1}`, bid: null } }
  if (node.type === 'noteRef' && !keepNotes) return { type: 'text', text: `[${node.attrs?.label || '*'}]`, marks: node.marks }
  const marks = node.marks?.filter((m) => m.type !== 'comment')
  const out: PMNode = { ...node, ...(node.marks ? { marks } : {}) }
  if (node.content) out.content = node.content.map((c) => prepareCopy(c, copy, keepNotes)).filter((c): c is PMNode => !!c)
  return out
}

/**
 * One document containing every record, each starting on a new page. Sections are repeated per
 * record so page setup, headers and footers stay correct.
 */
export function mergeDocument(content: PMNode, meta: DocMeta, data: MailMergeData, records: number[] = data.records.map((_, i) => i)): { content: PMNode; meta: DocMeta } {
  const S = meta.sections.length
  const blocks: PMNode[] = []
  records.forEach((r, k) => {
    const filled = prepareCopy(fillRecord(content, data.fields, data.records[r]), k, k === 0)!
    for (const b of filled.content ?? []) {
      if (b.type === 'sectionBreak') blocks.push({ ...b, attrs: { ...b.attrs, section: Number(b.attrs?.section ?? 0) + k * S } })
      else blocks.push(b)
    }
    if (k < records.length - 1) blocks.push({ type: 'sectionBreak', attrs: { section: k * S + S - 1, breakType: 'nextPage' } })
  })
  // every record after the first starts on a new page
  const sections = records.flatMap((_, k) => meta.sections.map((s, i) => (k > 0 && i === 0 ? { ...s, breakType: 'nextPage' } : { ...s })))
  return {
    content: { type: 'doc', content: blocks.length ? blocks : [{ type: 'paragraph' }] },
    meta: { ...meta, sections, comments: [], mailMerge: undefined },
  }
}
