// Citations and bibliography: source records, APA / MLA / Chicago formatting, and the
// package bibliography part (b:Sources) used by the CITATION and bibliography fields.

import type { CitationStyle, Source, SourcePerson, SourceType } from '../model/types'
import { children, parseXml } from '../docx/xml'

export const BIB_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'

export const SOURCE_TYPES: Record<SourceType, string> = {
  Book: 'Book',
  JournalArticle: 'Journal article',
  InternetSite: 'Web page',
  Report: 'Report',
  Misc: 'Other',
}

/** A formatted run of bibliography text. */
export interface Seg {
  text: string
  italic?: boolean
}

const clean = (s: string | undefined) => (s ?? '').trim()
const endDot = (s: string) => (/[.!?]$/.test(s) ? s : s + '.')

function initials(first: string | undefined) {
  return clean(first)
    .split(/[\s-]+/)
    .filter(Boolean)
    .map((p) => p[0].toUpperCase() + '.')
    .join(' ')
}

const lastName = (p: SourcePerson) => clean(p.last) || clean(p.first)

// ---------------------------------------------------------------------------
// in-text citations

/** In-text citation for one or more sources, e.g. (Smith, 2020; Lee & Park, 2021). */
export function formatCitation(sources: Source[], style: CitationStyle, page?: string): string {
  if (!sources.length) return '(Missing source)'
  const parts = sources.map((s) => {
    const names = authorShort(s, style)
    const year = clean(s.year) || 'n.d.'
    if (style === 'MLA') return names || shortTitle(s)
    if (style === 'Chicago') return `${names || shortTitle(s)} ${year}`
    return `${names || shortTitle(s)}, ${year}`
  })
  const p = clean(page)
  const suffix = p ? (style === 'APA' ? `, p. ${p}` : style === 'MLA' ? ` ${p}` : `, ${p}`) : ''
  return `(${parts.join('; ')}${suffix})`
}

function shortTitle(s: Source) {
  const t = clean(s.title)
  return t.length > 30 ? `“${t.slice(0, 28).trim()}…”` : `“${t}”`
}

function authorShort(s: Source, style: CitationStyle) {
  const a = s.authors.filter((p) => lastName(p))
  if (!a.length) return ''
  if (a.length === 1) return lastName(a[0])
  if (a.length === 2) return style === 'APA' ? `${lastName(a[0])} & ${lastName(a[1])}` : `${lastName(a[0])} and ${lastName(a[1])}`
  if (style === 'Chicago' && a.length === 3) return `${lastName(a[0])}, ${lastName(a[1])}, and ${lastName(a[2])}`
  return `${lastName(a[0])} et al.`
}

// ---------------------------------------------------------------------------
// bibliography entries

function apaAuthors(a: SourcePerson[]) {
  const names = a.filter((p) => lastName(p)).map((p) => [clean(p.last), initials(p.first)].filter(Boolean).join(', '))
  if (!names.length) return ''
  if (names.length === 1) return names[0]
  if (names.length <= 20) return names.slice(0, -1).join(', ') + ', & ' + names[names.length - 1]
  return names.slice(0, 19).join(', ') + ', … ' + names[names.length - 1]
}

function mlaAuthors(a: SourcePerson[]) {
  const n = a.filter((p) => lastName(p))
  if (!n.length) return ''
  const inv = (p: SourcePerson) => [clean(p.last), clean(p.first)].filter(Boolean).join(', ')
  const fwd = (p: SourcePerson) => [clean(p.first), clean(p.last)].filter(Boolean).join(' ')
  if (n.length === 1) return inv(n[0])
  if (n.length === 2) return `${inv(n[0])}, and ${fwd(n[1])}`
  return `${inv(n[0])}, et al`
}

function chicagoAuthors(a: SourcePerson[]) {
  const n = a.filter((p) => lastName(p))
  if (!n.length) return ''
  const inv = (p: SourcePerson) => [clean(p.last), clean(p.first)].filter(Boolean).join(', ')
  const fwd = (p: SourcePerson) => [clean(p.first), clean(p.last)].filter(Boolean).join(' ')
  if (n.length === 1) return inv(n[0])
  if (n.length <= 10) return [inv(n[0]), ...n.slice(1, -1).map(fwd)].join(', ') + ', and ' + fwd(n[n.length - 1])
  return [inv(n[0]), ...n.slice(1, 7).map(fwd)].join(', ') + ', et al'
}

export function formatEntry(s: Source, style: CitationStyle): Seg[] {
  const out: Seg[] = []
  const t = (text: string, italic = false) => text && out.push({ text, italic })
  const title = clean(s.title)
  const year = clean(s.year)
  const pub = clean(s.publisher)
  const journal = clean(s.journal)
  const vol = clean(s.volume)
  const iss = clean(s.issue)
  const pages = clean(s.pages)
  const url = clean(s.url)
  const container = s.type === 'JournalArticle' || s.type === 'InternetSite'
  if (style === 'APA') {
    const au = apaAuthors(s.authors)
    if (au) t(endDot(au) + ' ')
    t(`(${year || 'n.d.'}). `)
    if (container) {
      t(endDot(title) + ' ')
      if (journal) {
        t(journal, true)
        if (vol) {
          t(', ')
          t(vol, true)
        }
        if (iss) t(`(${iss})`)
        if (pages) t(`, ${pages}`)
        t('. ')
      } else if (pub) t(endDot(pub) + ' ')
    } else {
      t(endDot(title), true)
      t(' ')
      if (pub) t(endDot(pub) + ' ')
    }
    if (url) t(url)
  } else if (style === 'MLA') {
    const au = mlaAuthors(s.authors)
    if (au) t(endDot(au) + ' ')
    if (container) {
      t(`“${endDot(title)}” `)
      if (journal) {
        t(journal, true)
        const bits = [vol && `vol. ${vol}`, iss && `no. ${iss}`, year, pages && `pp. ${pages}`].filter(Boolean)
        t(bits.length ? `, ${bits.join(', ')}. ` : '. ')
      } else {
        const bits = [pub, year].filter(Boolean)
        if (bits.length) t(endDot(bits.join(', ')) + ' ')
      }
    } else {
      t(endDot(title), true)
      t(' ')
      const bits = [pub, year].filter(Boolean)
      if (bits.length) t(endDot(bits.join(', ')) + ' ')
    }
    if (url) t(endDot(url.replace(/^https?:\/\//, '')))
  } else {
    const au = chicagoAuthors(s.authors)
    if (au) t(endDot(au) + ' ')
    t(`${year || 'n.d.'}. `)
    if (container) {
      t(`“${endDot(title)}” `)
      if (journal) {
        t(journal, true)
        t(`${vol ? ' ' + vol : ''}${iss ? ` (${iss})` : ''}${pages ? `: ${pages}` : ''}. `)
      } else if (pub) t(endDot(pub) + ' ')
    } else {
      t(endDot(title), true)
      t(' ')
      const place = [clean(s.city), pub].filter(Boolean).join(': ')
      if (place) t(endDot(place) + ' ')
    }
    if (url) t(endDot(url))
  }
  // trim trailing space
  const last = out[out.length - 1]
  if (last) last.text = last.text.replace(/\s+$/, '')
  return out.filter((x) => x.text)
}

const sortKey = (s: Source) => (lastName(s.authors[0] ?? { last: '' }) || clean(s.title)).toLocaleLowerCase() + ' ' + clean(s.year)

/** Bibliography entries in alphabetical order. */
export function formatBibliography(sources: Source[], style: CitationStyle): Seg[][] {
  return [...sources].sort((a, b) => sortKey(a).localeCompare(sortKey(b))).map((s) => formatEntry(s, style))
}

export const BIBLIOGRAPHY_TITLES: Record<CitationStyle, string> = { APA: 'References', MLA: 'Works Cited', Chicago: 'Bibliography' }

// ---------------------------------------------------------------------------
// CITATION field instructions:  CITATION Tag1 \l 1033 \m Tag2 \p 23

export function citationInstr(tags: string[], page?: string): string {
  const [first, ...rest] = tags
  return `CITATION ${first} \\l 1033${rest.map((t) => ` \\m ${t}`).join('')}${clean(page) ? ` \\p ${clean(page).replace(/\s+/g, '')}` : ''}`
}

export function parseCitationInstr(instr: string): { tags: string[]; page?: string } | null {
  const m = /^\s*CITATION\s+(\S+)(.*)$/i.exec(instr)
  if (!m) return null
  const tags = [m[1]]
  for (const x of m[2].matchAll(/\\m\s+(\S+)/g)) tags.push(x[1])
  const page = /\\p\s+(\S+)/.exec(m[2])?.[1]
  return { tags, page }
}

export function newTag(s: Pick<Source, 'authors' | 'year' | 'title'>, existing: string[]): string {
  const base = ((lastName(s.authors[0] ?? { last: '' }) || clean(s.title) || 'Src').replace(/[^\p{L}\p{N}]/gu, '').slice(0, 3) || 'Src') + (clean(s.year).slice(-2) || '')
  let tag = base
  for (let i = 1; existing.includes(tag); i++) tag = base + i
  return tag
}

// ---------------------------------------------------------------------------
// package bibliography part

const XML_ESC = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const STYLE_FILE: Record<CitationStyle, { name: string; version: string }> = {
  APA: { name: 'APA', version: '7' },
  MLA: { name: 'MLA', version: '9' },
  Chicago: { name: 'Chicago', version: '17' },
}
const TYPE_TO_B: Record<SourceType, string> = { Book: 'Book', JournalArticle: 'JournalArticle', InternetSite: 'InternetSite', Report: 'Report', Misc: 'Misc' }

export function sourcesXml(sources: Source[], style: CitationStyle): string {
  const st = STYLE_FILE[style]
  const tag = (name: string, v?: string) => (clean(v) ? `<b:${name}>${XML_ESC(clean(v))}</b:${name}>` : '')
  const body = sources
    .map((s, i) => {
      const people = s.authors.filter((p) => lastName(p))
      const authors = people.length
        ? `<b:Author><b:Author><b:NameList>${people.map((p) => `<b:Person>${tag('Last', p.last)}${tag('First', p.first)}</b:Person>`).join('')}</b:NameList></b:Author></b:Author>`
        : ''
      const journalTag = s.type === 'JournalArticle' ? tag('JournalName', s.journal) : s.type === 'InternetSite' ? tag('InternetSiteTitle', s.journal) : ''
      const accessed = clean(s.accessed) ? (() => {
        const d = new Date(clean(s.accessed))
        return Number.isNaN(d.getTime()) ? '' : `<b:YearAccessed>${d.getUTCFullYear()}</b:YearAccessed><b:MonthAccessed>${d.toLocaleString('en-US', { month: 'long', timeZone: 'UTC' })}</b:MonthAccessed><b:DayAccessed>${d.getUTCDate()}</b:DayAccessed>`
      })() : ''
      return `<b:Source>${tag('Tag', s.tag)}<b:SourceType>${TYPE_TO_B[s.type]}</b:SourceType><b:Guid>{${guid(s.tag + i)}}</b:Guid>${authors}${tag('Title', s.title)}${tag('Year', s.year)}${tag('City', s.city)}${tag('Publisher', s.publisher)}${journalTag}${tag('Volume', s.volume)}${tag('Issue', s.issue)}${tag('Pages', s.pages)}${tag('URL', s.url)}${accessed}<b:RefOrder>${i + 1}</b:RefOrder></b:Source>`
    })
    .join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<b:Sources xmlns:b="${BIB_NS}" xmlns="${BIB_NS}" StyleName="${st.name}" Version="${st.version}">${body}</b:Sources>`
}

/** Deterministic GUID-shaped id so repeated exports of the same sources are identical. */
function guid(seed: string) {
  let h1 = 0x811c9dc5
  let h2 = 0x01000193
  for (const ch of seed) {
    h1 = Math.imul(h1 ^ ch.codePointAt(0)!, 16777619) >>> 0
    h2 = Math.imul(h2 + ch.codePointAt(0)!, 2246822519) >>> 0
  }
  const hex = (h1.toString(16).padStart(8, '0') + h2.toString(16).padStart(8, '0') + (h1 ^ h2).toString(16).padStart(8, '0') + ((h1 + h2) >>> 0).toString(16).padStart(8, '0')).toUpperCase()
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-8${hex.slice(17, 20)}-${hex.slice(20, 32)}`
}

const B_TO_TYPE: Record<string, SourceType> = { Book: 'Book', BookSection: 'Book', JournalArticle: 'JournalArticle', ArticleInAPeriodical: 'JournalArticle', ConferenceProceedings: 'JournalArticle', InternetSite: 'InternetSite', DocumentFromInternetSite: 'InternetSite', Report: 'Report' }

export function parseSourcesXml(xml: string): { sources: Source[]; style?: CitationStyle } | null {
  const doc = parseXml(xml, 'bibliography sources')
  const root = doc.documentElement
  if (root.localName !== 'Sources' || root.namespaceURI !== BIB_NS) return null
  const txt = (el: Element, name: string) => children(el, BIB_NS, name)[0]?.textContent?.trim() || undefined
  const sources: Source[] = []
  for (const s of children(root, BIB_NS, 'Source')) {
    const tag = txt(s, 'Tag')
    if (!tag) continue
    const people: SourcePerson[] = []
    const author = children(children(s, BIB_NS, 'Author')[0] ?? s, BIB_NS, 'Author')[0]
    for (const nl of author ? children(author, BIB_NS, 'NameList') : []) for (const p of children(nl, BIB_NS, 'Person')) people.push({ last: txt(p, 'Last') ?? '', first: [txt(p, 'First'), txt(p, 'Middle')].filter(Boolean).join(' ') || undefined })
    const corp = author && children(author, BIB_NS, 'Corporate')[0]?.textContent?.trim()
    if (corp) people.push({ last: corp })
    const y = txt(s, 'YearAccessed')
    const mo = txt(s, 'MonthAccessed')
    const d = txt(s, 'DayAccessed')
    sources.push({
      tag,
      type: B_TO_TYPE[txt(s, 'SourceType') ?? ''] ?? 'Misc',
      authors: people,
      title: txt(s, 'Title') ?? '',
      year: txt(s, 'Year'),
      publisher: txt(s, 'Publisher'),
      city: txt(s, 'City'),
      journal: txt(s, 'JournalName') ?? txt(s, 'PeriodicalTitle') ?? txt(s, 'InternetSiteTitle') ?? txt(s, 'ConferenceName'),
      volume: txt(s, 'Volume'),
      issue: txt(s, 'Issue'),
      pages: txt(s, 'Pages'),
      url: txt(s, 'URL'),
      accessed: y ? [y, mo, d].filter(Boolean).join(' ') : undefined,
    })
  }
  const sn = (root.getAttribute('StyleName') ?? '').toUpperCase()
  const style: CitationStyle | undefined = sn.startsWith('APA') ? 'APA' : sn.startsWith('MLA') ? 'MLA' : sn.startsWith('CHICAGO') ? 'Chicago' : undefined
  return { sources, style }
}
