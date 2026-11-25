import { describe, expect, it } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import JSZip from 'jszip'
import { importDocx } from '../../src/docx/import'
import { exportDocx } from '../../src/docx/export'
import { validatePackage } from '../../src/docx/validate'
import { emptyDocument, type PMNode } from '../../src/model/types'

const fx = (n: string) => new Uint8Array(fs.readFileSync(path.resolve('tests/fixtures', n)))
const outDir = path.resolve('tests/output')
fs.mkdirSync(outDir, { recursive: true })

function walk(n: PMNode, f: (n: PMNode) => void) {
  f(n)
  n.content?.forEach((c) => walk(c, f))
}
function all(n: PMNode, type: string) {
  const out: PMNode[] = []
  walk(n, (x) => x.type === type && out.push(x))
  return out
}
function plain(n: PMNode): string {
  if (n.text) return n.text
  const inner = (n.content ?? []).map(plain).join('')
  return ['paragraph', 'heading'].includes(n.type) ? inner + '\n' : inner
}
/** Normalise for structural comparison: drop volatile ids and media rel names. */
function norm(n: PMNode): any {
  return JSON.parse(
    JSON.stringify(n, (k, v) => {
      if (k === 'nid' || k === 'docPrName') return undefined
      if (k === 'anchorXml' && typeof v === 'string') return v.replace(/r:embed="[^"]*"/, '').replace(/\bid="\d+"/g, '')
      if (k === 'bid' || k === 'revId') return undefined
      return v
    }),
  )
}

async function cycle(bytes: Uint8Array) {
  const a = await importDocx(bytes)
  const e = await exportDocx(a.doc, a.original)
  const b = await importDocx(e.data)
  return { a, e, b }
}

describe('DOCX import', () => {
  it('imports a simple document', async () => {
    const r = await importDocx(fx('simple.docx'))
    const h = all(r.doc.content, 'heading')
    expect(h[0].attrs?.level).toBe(1)
    expect(plain(r.doc.content)).toContain('Hello SCRIBE')
    expect(plain(r.doc.content)).toContain('Second paragraph.')
    expect(r.doc.meta.metadata.title).toBe('Simple document')
  })

  it('imports complex formatting, lists, tables, images, links, notes, comments and revisions', async () => {
    const r = await importDocx(fx('complex.docx'))
    const c = r.doc.content
    const marksOf = (t: string) => all(c, 'text').find((x) => x.text?.includes(t))?.marks?.map((m) => m.type) ?? []
    expect(marksOf('Bold')).toContain('bold')
    expect(marksOf('Italic')).toContain('italic')
    expect(marksOf('Strike')).toContain('strike')
    expect(marksOf('Highlighted')).toContain('highlight')
    const red = all(c, 'text').find((x) => x.text === 'Red 14pt Georgia')!
    const ts = red.marks!.find((m) => m.type === 'textStyle')!.attrs!
    expect(ts.color).toBe('#c00000')
    expect(ts.fontSize).toBe('14pt')
    expect(ts.fontFamily).toBe('Georgia')
    // paragraph formatting
    const just = all(c, 'paragraph').find((p) => plain(p).startsWith('Justified'))!
    expect(just.attrs).toMatchObject({ textAlign: 'justify', indentLeft: 720, indentRight: 360, indentFirstLine: 360, spacingBefore: 240, spacingAfter: 120, lineSpacing: 1.5 })
    const hanging = all(c, 'paragraph').find((p) => plain(p).startsWith('Hanging'))!
    expect(hanging.attrs?.indentFirstLine).toBe(-720)
    expect(plain(hanging)).toContain('\t')
    // lists
    const bl = all(c, 'bulletList')
    expect(bl.length).toBeGreaterThanOrEqual(2) // outer + nested
    const ol = all(c, 'orderedList')
    expect(ol.length).toBeGreaterThanOrEqual(2)
    expect(plain(ol[0])).toContain('Sub-step a')
    // tables: outer + nested, merges
    const tables = all(c, 'table')
    expect(tables.length).toBe(2)
    const cells = [...all(tables[0], 'tableCell'), ...all(tables[0], 'tableHeader')]
    expect(cells.some((x) => x.attrs?.colspan === 2)).toBe(true)
    expect(cells.some((x) => x.attrs?.rowspan === 2)).toBe(true)
    expect(all(tables[0], 'tableHeader')[0].attrs?.background).toBe('#d9e2f3')
    // images
    const img = all(c, 'image')[0]
    expect(img.attrs?.src).toMatch(/^data:image\/png;base64,/)
    expect(img.attrs?.width).toBe(120)
    expect(img.attrs?.alt).toBe('Gradient swatch')
    // links
    const links = all(c, 'text').flatMap((t) => t.marks ?? []).filter((m) => m.type === 'link')
    expect(links.some((l) => l.attrs?.href === 'https://example.com/docs?q=1')).toBe(true)
    expect(links.some((l) => l.attrs?.href === '#target')).toBe(true)
    // notes, comments, revisions, bookmarks
    expect(all(c, 'noteRef').length).toBe(2)
    expect(r.doc.meta.footnotes.map((f) => f.text)).toEqual(['This is footnote one.', 'Footnote two text.'])
    expect(r.doc.meta.comments[0]).toMatchObject({ author: 'Reviewer A', text: 'Please check this sentence.' })
    expect(marksOf('This sentence has a comment.')).toContain('comment')
    expect(marksOf('inserted words')).toContain('insertion')
    expect(marksOf('deleted words')).toContain('deletion')
    expect(all(c, 'bookmark').map((b) => b.attrs?.name)).toContain('target')
    // unicode + rtl
    expect(plain(c)).toContain('日本語テキスト')
    expect(plain(c)).toContain('🎉')
    expect(all(c, 'paragraph').find((p) => plain(p).includes('مرحبا'))!.attrs?.direction).toBe('rtl')
    // page & section breaks
    expect(all(c, 'pageBreak').length).toBe(1)
    expect(all(c, 'sectionBreak').length).toBe(1)
    expect(r.doc.meta.sections).toHaveLength(2)
    expect(r.doc.meta.sections[1].page.orientation).toBe('landscape')
    expect(r.doc.meta.sections[1].page.marginTop).toBe(1080)
    // header/footer with page fields
    const s0 = r.doc.meta.sections[0]
    expect(plain(r.doc.meta.headerFooters[s0.headerDefault!])).toContain('Complex Fixture — header')
    const fields = all(r.doc.meta.headerFooters[s0.footerDefault!], 'field')
    expect(fields.map((f) => f.attrs?.instr.split(' ')[0])).toEqual(['PAGE', 'NUMPAGES'])
  })

  it('reports and preserves unsupported OOXML instead of dropping it', async () => {
    const r = await importDocx(fx('unsupported.docx'))
    const op = all(r.doc.content, 'opaqueInline')
    expect(op.some((o) => o.attrs?.label === 'Text box' && o.attrs?.preview.includes('Text inside a text box'))).toBe(true)
    expect(op.some((o) => o.attrs?.label === 'Equation')).toBe(true)
    expect(op.some((o) => String(o.attrs?.xml).includes('w14:checkbox') || String(o.attrs?.xml).includes('checkbox'))).toBe(true)
    expect(plain(r.doc.content)).toContain('Inside custom XML')
    const features = r.report.map((x) => x.feature)
    expect(features).toContain('Text boxes')
    expect(features).toContain('Equations')
    expect(features).toContain('Custom XML parts')
  })

  it('preserves the table of contents field', async () => {
    const r = await importDocx(fx('toc.docx'))
    expect(r.report.some((x) => x.feature === 'Table of contents' || x.feature === 'Fields')).toBe(true)
  })
})

describe('DOCX security and malformed input', () => {
  it('rejects non-zip input with a clear message', async () => {
    await expect(importDocx(fx('not-a-zip.docx'))).rejects.toThrow(/Not a DOCX/)
  })
  it('rejects truncated archives', async () => {
    await expect(importDocx(fx('truncated.docx'))).rejects.toThrow()
  })
  it('rejects DOCTYPE / entity expansion (XXE, billion laughs)', async () => {
    await expect(importDocx(fx('xxe.docx'))).rejects.toThrow(/DOCTYPE/)
  })
  it('rejects malformed XML', async () => {
    await expect(importDocx(fx('malformed-xml.docx'))).rejects.toThrow(/malformed/i)
  })
  it('rejects zip bombs by compression ratio', async () => {
    await expect(importDocx(fx('zipbomb.docx'))).rejects.toThrow(/zip bomb|suspicious/i)
  })
  it('rejects legacy binary .doc', async () => {
    const doc = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0])
    await expect(importDocx(doc)).rejects.toThrow(/legacy binary \.doc/)
  })
})

describe('DOCX round trip', () => {
  for (const name of ['simple.docx', 'complex.docx', 'toc.docx', 'unsupported.docx', 'large.docx']) {
    it(`${name}: export is structurally valid and re-imports identically`, async () => {
      const { a, e, b } = await cycle(fx(name))
      fs.writeFileSync(path.join(outDir, 'roundtrip-' + name), e.data)
      const issues = await validatePackage(e.data)
      expect(issues).toEqual([])
      expect(plain(b.doc.content)).toBe(plain(a.doc.content))
      expect(norm(b.doc.content)).toEqual(norm(a.doc.content))
      expect(b.doc.meta.sections.map((s) => s.page)).toEqual(a.doc.meta.sections.map((s) => s.page))
      expect(Object.values(b.doc.meta.headerFooters).map(norm)).toEqual(Object.values(a.doc.meta.headerFooters).map(norm))
      expect(b.doc.meta.footnotes.map((f) => f.text)).toEqual(a.doc.meta.footnotes.map((f) => f.text))
      expect(b.doc.meta.comments.map((c) => [c.author, c.text])).toEqual(a.doc.meta.comments.map((c) => [c.author, c.text]))
    })
  }

  it('is stable across five repeated import/export cycles', async () => {
    let bytes: Uint8Array = fx('complex.docx')
    const first = await importDocx(bytes)
    for (let i = 0; i < 5; i++) {
      const r = await importDocx(bytes)
      bytes = (await exportDocx(r.doc, r.original)).data
    }
    const last = await importDocx(bytes)
    expect(norm(last.doc.content)).toEqual(norm(first.doc.content))
    expect(await validatePackage(bytes)).toEqual([])
  })

  it('keeps package parts it does not understand (custom XML, unknown elements) byte-for-byte', async () => {
    const { a, e } = await cycle(fx('unsupported.docx'))
    const zip = await JSZip.loadAsync(e.data)
    expect(await zip.file('customXml/item1.xml')!.async('string')).toBe('<root><value>kept</value></root>')
    const doc = await zip.file('word/document.xml')!.async('string')
    expect(doc).toContain('Text inside a text box')
    expect(doc).toContain('a²+b²=c²')
    expect(doc).toMatch(/checkbox/)
    expect(a.report.length).toBeGreaterThan(0)
  })

  it('exports edits: new formatting, list, table, image and section changes', async () => {
    const r = await importDocx(fx('simple.docx'))
    const content = r.doc.content.content!
    content.push(
      { type: 'paragraph', attrs: { textAlign: 'center' }, content: [{ type: 'text', text: 'Edited bold blue', marks: [{ type: 'bold' }, { type: 'textStyle', attrs: { color: '#1155cc', fontSize: '16pt', fontFamily: 'Arial' } }] }] },
      { type: 'orderedList', attrs: { start: 3 }, content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Third item' }] }] }] },
      { type: 'bulletList', content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'New bullet' }] }, { type: 'bulletList', content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Nested bullet' }] }] }] }] }] },
      {
        type: 'table',
        content: [
          { type: 'tableRow', content: [{ type: 'tableCell', attrs: { colspan: 1, rowspan: 2, colwidth: [200] }, content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Tall' }] }] }, { type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: [150], background: '#ffeeaa' }, content: [{ type: 'paragraph', content: [{ type: 'text', text: 'R1' }] }] }] },
          { type: 'tableRow', content: [{ type: 'tableCell', attrs: { colspan: 1, rowspan: 1, colwidth: [150] }, content: [{ type: 'paragraph', content: [{ type: 'text', text: 'R2' }] }] }] },
        ],
      },
      { type: 'paragraph', content: [{ type: 'image', attrs: { src: all((await importDocx(fx('complex.docx'))).doc.content, 'image')[0].attrs!.src, width: 60, height: 40 } }] },
    )
    r.doc.meta.sections[0].page.orientation = 'landscape'
    ;[r.doc.meta.sections[0].page.width, r.doc.meta.sections[0].page.height] = [r.doc.meta.sections[0].page.height, r.doc.meta.sections[0].page.width]
    const e = await exportDocx(r.doc, r.original)
    fs.writeFileSync(path.join(outDir, 'edited-simple.docx'), e.data)
    expect(await validatePackage(e.data)).toEqual([])
    const b = await importDocx(e.data)
    const t = all(b.doc.content, 'text').find((x) => x.text === 'Edited bold blue')!
    expect(t.marks!.map((m) => m.type)).toContain('bold')
    expect(t.marks!.find((m) => m.type === 'textStyle')!.attrs).toMatchObject({ color: '#1155cc', fontSize: '16pt', fontFamily: 'Arial' })
    expect(all(b.doc.content, 'orderedList')[0].attrs?.start).toBe(3)
    expect(all(b.doc.content, 'bulletList').length).toBe(2)
    const cells = all(b.doc.content, 'tableCell')
    expect(cells.find((c) => plain(c).includes('Tall'))!.attrs?.rowspan).toBe(2)
    expect(cells.find((c) => plain(c).includes('R1'))!.attrs?.background).toBe('#ffeeaa')
    expect(all(b.doc.content, 'image')[0].attrs?.width).toBe(60)
    expect(b.doc.meta.sections[0].page.orientation).toBe('landscape')
  })

  it('exports a brand-new document without an original package', async () => {
    const d = emptyDocument()
    d.content.content = [
      { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Fresh' }] },
      { type: 'paragraph', content: [{ type: 'text', text: 'Body', marks: [{ type: 'italic' }, { type: 'comment', attrs: { commentId: 'c1' } }] }] },
      { type: 'blockquote', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Quoted' }] }] },
      { type: 'pageBreak' },
      { type: 'paragraph', content: [{ type: 'text', text: 'Page two ' }, { type: 'text', text: 'link', marks: [{ type: 'link', attrs: { href: 'https://scribe.example/' } }] }] },
    ]
    d.meta.comments = [{ id: 'c1', author: 'Me', date: '2026-03-01T00:00:00Z', text: 'A new comment' }]
    d.meta.headerFooters['new:footer'] = { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Page ' }, { type: 'field', attrs: { instr: 'PAGE', result: '1' } }] }] }
    d.meta.sections[0].footerDefault = 'new:footer'
    d.meta.metadata.title = 'Fresh doc'
    const e = await exportDocx(d, null)
    fs.writeFileSync(path.join(outDir, 'fresh.docx'), e.data)
    expect(await validatePackage(e.data)).toEqual([])
    const b = await importDocx(e.data)
    expect(all(b.doc.content, 'heading')[0].attrs?.level).toBe(1)
    expect(all(b.doc.content, 'pageBreak')).toHaveLength(1)
    expect(b.doc.meta.comments[0].text).toBe('A new comment')
    expect(b.doc.meta.metadata.title).toBe('Fresh doc')
    const f = b.doc.meta.headerFooters[b.doc.meta.sections[0].footerDefault!]
    expect(all(f, 'field')[0].attrs?.instr).toBe('PAGE')
  })

  it('handles large documents within a reasonable time', async () => {
    const t0 = performance.now()
    const { b } = await cycle(fx('large.docx'))
    const ms = performance.now() - t0
    expect(all(b.doc.content, 'paragraph').length).toBeGreaterThanOrEqual(3000)
    expect(ms).toBeLessThan(20000)
  })
})
