import { describe, expect, it } from 'vitest'
import fs from 'node:fs'
import path from 'node:path'
import JSZip from 'jszip'
import { importDocx } from '../../src/docx/import'
import { exportDocx } from '../../src/docx/export'
import { validatePackage } from '../../src/docx/validate'
import { emptyDocument, type PMNode, type ScribeDocument } from '../../src/model/types'
import { parseLinear, toLinear, ommlToTree, treeToOmml } from '../../src/objects/math'
import { setTextBoxParagraphs } from '../../src/objects/textbox'
import { citationInstr, formatBibliography, formatCitation } from '../../src/features/citations'
import { mergeDocument, parseCsv, recipientsFromCsv, mergeFieldInstr } from '../../src/features/merge'

const fx = (n: string) => new Uint8Array(fs.readFileSync(path.resolve('tests/fixtures', n)))
const outDir = path.resolve('tests/output')
fs.mkdirSync(outDir, { recursive: true })
const save = (n: string, d: Uint8Array) => fs.writeFileSync(path.join(outDir, n), d)

function find(n: PMNode, pred: (x: PMNode) => boolean, out: PMNode[] = []): PMNode[] {
  if (pred(n)) out.push(n)
  for (const c of n.content ?? []) find(c, pred, out)
  return out
}
const para = (...content: PMNode[]): PMNode => ({ type: 'paragraph', content })
const doc = (blocks: PMNode[], meta: Partial<ScribeDocument['meta']> = {}): ScribeDocument => {
  const d = emptyDocument()
  return { ...d, content: { type: 'doc', content: blocks }, meta: { ...d.meta, ...meta } }
}
async function roundTrip(d: ScribeDocument, original?: Uint8Array, name?: string) {
  const { data, report } = await exportDocx(d, original)
  expect(await validatePackage(data)).toEqual([])
  if (name) save(name, data)
  return { back: await importDocx(data), data, report }
}

describe('charts, diagrams and custom XML', () => {
  it('imports chart data, diagram text, custom XML and sources', async () => {
    const r = await importDocx(fx('objects.docx'))
    const chart = find(r.doc.content, (n) => n.attrs?.object?.kind === 'chart')[0]
    expect(chart.attrs!.object.chart).toMatchObject({ part: 'word/charts/chart1.xml', kind: 'column', title: 'Quarterly sales', categories: ['Q1', 'Q2', 'Q3'], editable: true })
    expect(chart.attrs!.object.chart.series).toEqual([{ name: 'North', values: [10, 14, 9] }, { name: 'South', values: [7, 11, 15] }])
    const sa = find(r.doc.content, (n) => n.attrs?.object?.kind === 'smartart')[0]
    expect(sa.attrs!.object.smartArt.items.map((i: any) => i.text)).toEqual(['Plan', 'Build', 'Ship'])
    expect(sa.attrs!.object.smartArt.drawingPart).toBe('word/diagrams/drawing1.xml')
    expect(r.doc.meta.customXml?.map((p) => p.path)).toEqual(['customXml/item1.xml'])
    expect(r.doc.meta.sources?.[0]).toMatchObject({ tag: 'Knu97', title: 'The Art of Computer Programming', authors: [{ last: 'Knuth', first: 'Donald' }] })
  })

  it('writes edited chart data, diagram text and custom XML back into the package', async () => {
    const original = fx('objects.docx')
    const r = await importDocx(original)
    const chart = find(r.doc.content, (n) => n.attrs?.object?.kind === 'chart')[0]
    chart.attrs!.object.chart = { ...chart.attrs!.object.chart, title: 'Sales by quarter', categories: ['Q1', 'Q2', 'Q3', 'Q4'], series: [{ name: 'North', values: [11, 14, 9, 20] }, { name: 'South', values: [7, null, 15, 3] }, { name: 'East', values: [1, 2, 3, 4] }], dirty: true }
    const sa = find(r.doc.content, (n) => n.attrs?.object?.kind === 'smartart')[0]
    sa.attrs!.object.smartArt = { ...sa.attrs!.object.smartArt, items: sa.attrs!.object.smartArt.items.map((i: any) => (i.text === 'Build' ? { ...i, text: 'Build & test' } : i)), dirty: true }
    r.doc.meta.customXml = [{ path: 'customXml/item1.xml', xml: '<invoice xmlns="urn:example:invoice"><number>43</number></invoice>', edited: true }]
    const { back, data } = await roundTrip(r.doc, original, 'objects-edited.docx')
    const c2 = find(back.doc.content, (n) => n.attrs?.object?.kind === 'chart')[0].attrs!.object.chart
    expect(c2.title).toBe('Sales by quarter')
    expect(c2.categories).toEqual(['Q1', 'Q2', 'Q3', 'Q4'])
    expect(c2.series).toEqual([{ name: 'North', values: [11, 14, 9, 20] }, { name: 'South', values: [7, null, 15, 3] }, { name: 'East', values: [1, 2, 3, 4] }])
    const zip = await JSZip.loadAsync(data)
    const chartXml = await zip.file('word/charts/chart1.xml')!.async('string')
    expect(chartXml).not.toContain('externalData')
    expect(chartXml).toContain('<c:invertIfNegative val="0"/>') // untouched formatting stays
    expect(await zip.file('word/diagrams/drawing1.xml')!.async('string')).toContain('Build &amp; test')
    expect(await zip.file('word/diagrams/data1.xml')!.async('string')).toContain('Build &amp; test')
    expect(await zip.file('customXml/item1.xml')!.async('string')).toContain('<number>43</number>')
    const sa2 = find(back.doc.content, (n) => n.attrs?.object?.kind === 'smartart')[0].attrs!.object.smartArt
    expect(sa2.items.map((i: any) => i.text)).toEqual(['Plan', 'Build & test', 'Ship'])
  })

  it('leaves unedited objects byte-identical', async () => {
    const original = fx('objects.docx')
    const r = await importDocx(original)
    const { data } = await exportDocx(r.doc, original)
    const a = await JSZip.loadAsync(original)
    const b = await JSZip.loadAsync(data)
    for (const p of ['word/charts/chart1.xml', 'word/diagrams/data1.xml', 'word/diagrams/drawing1.xml', 'customXml/item1.xml'])
      expect(await b.file(p)!.async('string'), p).toBe(await a.file(p)!.async('string'))
  })
})

describe('text boxes and equations', () => {
  it('edits text box text in both copies and round-trips it', async () => {
    const original = fx('unsupported.docx')
    const r = await importDocx(original)
    const tb = find(r.doc.content, (n) => n.attrs?.object?.kind === 'textbox')[0]
    tb.attrs!.xml = setTextBoxParagraphs(tb.attrs!.xml, ['First line', 'Second line'])
    const { back, data } = await roundTrip(r.doc, original, 'textbox-edited.docx')
    const xml = await (await JSZip.loadAsync(data)).file('word/document.xml')!.async('string')
    expect(xml.match(/Second line/g)?.length).toBe(2) // modern shape + fallback
    expect(find(back.doc.content, (n) => n.attrs?.object?.kind === 'textbox')[0].attrs!.preview).toContain('First line')
  })

  it('inserts new equations that survive export', async () => {
    const xml = treeToOmml(parseLinear('x = (-b ± √(b^2-4ac))/(2a)'))
    const d = doc([para({ type: 'text', text: 'Root: ' }, { type: 'opaqueInline', attrs: { xml, label: 'Equation', preview: '', object: { kind: 'equation' } } })])
    const { back } = await roundTrip(d, undefined, 'equation.docx')
    const eq = find(back.doc.content, (n) => n.attrs?.object?.kind === 'equation')[0]
    expect(eq.attrs!.xml).toBe(xml)
    expect(toLinear(ommlToTree(eq.attrs!.xml))).toBe('x = (-b ± √(b^2-4ac))/2a')
  })
})

describe('shapes', () => {
  it('exports native shapes and re-imports them as editable shapes', async () => {
    const shapes = [
      { kind: 'rect', width: 160, height: 96, fill: '#4472c4', stroke: '#2f528f', strokeWidth: 1.5, text: 'Hello', arrow: false },
      { kind: 'ellipse', width: 80, height: 80, fill: null, stroke: '#dc2626', strokeWidth: 2, text: '', arrow: false },
      { kind: 'line', width: 200, height: 4, fill: null, stroke: '#1f2937', strokeWidth: 2, text: '', arrow: true },
      { kind: 'star5', width: 90, height: 90, fill: '#ffc000', stroke: null, strokeWidth: 0, text: '', arrow: false },
    ]
    const d = doc([para(...shapes.map((a) => ({ type: 'shape', attrs: a })))])
    const { back } = await roundTrip(d, undefined, 'shapes.docx')
    expect(find(back.doc.content, (n) => n.type === 'shape').map((n) => n.attrs)).toEqual(shapes)
  })
})

describe('citations and bibliography', () => {
  it('writes the sources part, citation fields and the bibliography, and reads them back', async () => {
    const sources = [
      { tag: 'Smi20', type: 'Book' as const, authors: [{ last: 'Smith', first: 'Jane' }], title: 'Writing Well', year: '2020', publisher: 'Acme Press' },
      { tag: 'Lee21', type: 'JournalArticle' as const, authors: [{ last: 'Lee', first: 'Ana' }, { last: 'Park', first: 'Min' }], title: 'On citations', year: '2021', journal: 'Journal of Notes', volume: '4', issue: '2', pages: '10-20' },
    ]
    expect(formatCitation([sources[0]], 'APA', '42')).toBe('(Smith, 2020, p. 42)')
    expect(formatCitation([sources[1]], 'APA')).toBe('(Lee & Park, 2021)')
    expect(formatCitation([sources[1]], 'MLA')).toBe('(Lee and Park)')
    const entries = formatBibliography(sources, 'APA')
    expect(entries.map((e) => e.map((s) => s.text).join(''))).toEqual(['Lee, A., & Park, M. (2021). On citations. Journal of Notes, 4(2), 10-20.', 'Smith, J. (2020). Writing Well. Acme Press.'])
    const d = doc(
      [
        para({ type: 'text', text: 'As shown ' }, { type: 'field', attrs: { instr: citationInstr(['Smi20'], '42'), result: '(Smith, 2020, p. 42)' } }),
        { type: 'bibliography', attrs: { title: 'References', entries } },
      ],
      { sources, citationStyle: 'APA' },
    )
    const { back, data } = await roundTrip(d, undefined, 'citations.docx')
    expect(back.doc.meta.sources?.map((s) => s.tag)).toEqual(['Smi20', 'Lee21'])
    expect(back.doc.meta.sources?.[1]).toMatchObject({ journal: 'Journal of Notes', volume: '4', issue: '2', pages: '10-20' })
    expect(back.doc.meta.citationStyle).toBe('APA')
    expect(find(back.doc.content, (n) => n.type === 'field')[0].attrs).toMatchObject({ instr: 'CITATION Smi20 \\l 1033 \\p 42', result: '(Smith, 2020, p. 42)' })
    const bib = find(back.doc.content, (n) => n.type === 'bibliography')[0]
    expect(bib.attrs!.title).toBe('References')
    expect(bib.attrs!.entries.map((e: any) => e.map((s: any) => s.text).join(''))).toEqual(entries.map((e) => e.map((s) => s.text).join('')))
    const zip = await JSZip.loadAsync(data)
    expect(Object.keys(zip.files).filter((f) => f.startsWith('customXml/') && !zip.files[f].dir).sort()).toEqual(['customXml/_rels/item1.xml.rels', 'customXml/item1.xml', 'customXml/itemProps1.xml'])
  })

  it('updates the existing sources part of an imported file instead of adding another', async () => {
    const original = fx('objects.docx')
    const r = await importDocx(original)
    r.doc.meta.sources = [...(r.doc.meta.sources ?? []), { tag: 'New1', type: 'Misc', authors: [], title: 'A new source' }]
    const { back, data } = await roundTrip(r.doc, original)
    expect(back.doc.meta.sources?.map((s) => s.tag)).toEqual(['Knu97', 'New1'])
    const zip = await JSZip.loadAsync(data)
    expect(Object.keys(zip.files).filter((f) => /^customXml\/item\d+\.xml$/.test(f)).length).toBe(2)
  })
})

describe('mail merge', () => {
  it('parses CSV with quotes, separators and line breaks', () => {
    expect(parseCsv('Name,City\r\n"Smith, Jane","New\nYork"\r\nLee,"Say ""hi"""\n')).toEqual([['Name', 'City'], ['Smith, Jane', 'New\nYork'], ['Lee', 'Say "hi"']])
    expect(parseCsv('a;b\n1;2')).toEqual([['a', 'b'], ['1', '2']])
    expect(recipientsFromCsv('First Name,E-mail,First Name\nA,b,c', 'x.csv').fields).toEqual(['First_Name', 'Email', 'First_Name_2'])
    expect(() => parseCsv('a,"b')).toThrow(/quotation/)
  })

  it('merges every record into its own section and exports a valid file', async () => {
    const data = recipientsFromCsv('Name,City\nAda,London\nGrace,New York\nLinus,Helsinki', 'r.csv')
    const content: PMNode = { type: 'doc', content: [para({ type: 'text', text: 'Dear ' }, { type: 'field', attrs: { instr: mergeFieldInstr('Name'), result: '«Name»' } }, { type: 'text', text: ' of ' }, { type: 'field', attrs: { instr: mergeFieldInstr('City'), result: '«City»' } }, { type: 'text', text: ',' }), para({ type: 'bookmark', attrs: { name: 'Sign' } }, { type: 'text', text: 'Regards' })] }
    const m = mergeDocument(content, emptyDocument().meta, data)
    expect(m.meta.sections.length).toBe(3)
    const { back } = await roundTrip({ version: 1, content: m.content, meta: m.meta }, undefined, 'merged.docx')
    const text = JSON.stringify(back.doc.content)
    for (const s of ['Dear Ada of London,', 'Dear Grace of New York,', 'Dear Linus of Helsinki,']) expect(text).toContain(s)
    expect(back.doc.meta.sections.length).toBe(3)
    expect(find(back.doc.content, (n) => n.type === 'bookmark').map((b) => b.attrs!.name)).toEqual(['Sign', 'Sign_2', 'Sign_3'])
    expect(find(back.doc.content, (n) => n.type === 'field').length).toBe(0)
  })

  it('keeps merge fields as fields through a round trip', async () => {
    const d = doc([para({ type: 'field', attrs: { instr: mergeFieldInstr('Name'), result: '«Name»' } })])
    const { back } = await roundTrip(d)
    expect(find(back.doc.content, (n) => n.type === 'field')[0].attrs).toMatchObject({ instr: 'MERGEFIELD Name', result: '«Name»' })
  })
})

describe('styles', () => {
  it('writes edited and new styles into styles.xml and reads them back', async () => {
    const d = doc([{ type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Title' }] }, { type: 'paragraph', attrs: { styleId: 'Callout' }, content: [{ type: 'text', text: 'Note' }] }], {
      styles: [
        { id: 'Heading1', name: 'heading 1', type: 'paragraph', fontSizePt: 22, color: '#c00000', bold: true, alignment: 'center', modified: true, outlineLevel: 0 },
        { id: 'Callout', name: 'Callout', type: 'paragraph', basedOn: 'Normal', italic: true, color: '#2f5496', spacingBefore: 120, spacingAfter: 240, custom: true },
      ],
    })
    const { back, data } = await roundTrip(d, undefined, 'styles.docx')
    const xml = await (await JSZip.loadAsync(data)).file('word/styles.xml')!.async('string')
    expect(xml).toMatch(/w:styleId="Callout"/)
    const h1 = back.doc.meta.styles.find((s) => s.id === 'Heading1')!
    expect(h1).toMatchObject({ fontSizePt: 22, color: '#c00000', bold: true, alignment: 'center' })
    expect(h1.spacingBefore).toBe(360) // unmanaged properties of the original definition are kept
    expect(back.doc.meta.styles.find((s) => s.id === 'Callout')).toMatchObject({ name: 'Callout', italic: true, color: '#2f5496', spacingBefore: 120, spacingAfter: 240 })
    expect(find(back.doc.content, (n) => n.attrs?.styleId === 'Callout').length).toBe(1)
  })
})
