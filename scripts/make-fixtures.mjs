// Generates the DOCX compatibility corpus in tests/fixtures using the independent
// `docx` library (not SCRIBE's own writer), so import tests exercise files SCRIBE did not produce.
import fs from 'node:fs'
import path from 'node:path'
import zlib from 'node:zlib'
import JSZip from 'jszip'
import {
  AlignmentType, BorderStyle, Bookmark, CommentRangeEnd, CommentRangeStart, CommentReference, DeletedTextRun, Document,
  ExternalHyperlink, Footer, FootnoteReferenceRun, Header, HeadingLevel, ImageRun, InsertedTextRun, InternalHyperlink,
  LevelFormat, Packer, PageBreak, PageNumber, PageOrientation, Paragraph, SectionType, ShadingType, Table, TableCell,
  TableOfContents, TableRow, TextRun, UnderlineType, VerticalMergeType, WidthType,
} from 'docx'

const out = path.resolve('tests/fixtures')
fs.mkdirSync(out, { recursive: true })

function png(w, h, rgb) {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    return c >>> 0
  })
  const crc = (buf) => {
    let c = 0xffffffff
    for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8)
    return (c ^ 0xffffffff) >>> 0
  }
  const chunk = (type, data) => {
    const len = Buffer.alloc(4)
    len.writeUInt32BE(data.length)
    const td = Buffer.concat([Buffer.from(type), data])
    const c = Buffer.alloc(4)
    c.writeUInt32BE(crc(td))
    return Buffer.concat([len, td, c])
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(w, 0)
  ihdr.writeUInt32BE(h, 4)
  ihdr[8] = 8
  ihdr[9] = 2
  const raw = Buffer.alloc((w * 3 + 1) * h)
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) raw.set([(rgb[0] + x) & 255, (rgb[1] + y) & 255, rgb[2]], y * (w * 3 + 1) + 1 + x * 3)
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))])
}

async function save(name, doc) {
  const buf = await Packer.toBuffer(doc)
  fs.writeFileSync(path.join(out, name), buf)
  process.stdout.write(`wrote ${name} ${buf.length} bytes\n`)
}

// 1. simple
await save(
  'simple.docx',
  new Document({
    creator: 'Fixture', title: 'Simple document',
    sections: [{ children: [new Paragraph({ text: 'Hello SCRIBE', heading: HeadingLevel.HEADING_1 }), new Paragraph('A plain paragraph of text.'), new Paragraph('Second paragraph.')] }],
  }),
)

// 2. complex formatting + lists + tables + images + links + headers/footers + notes + comments + revisions
const image = png(120, 80, [200, 40, 90])
await save(
  'complex.docx',
  new Document({
    creator: 'Fixture', title: 'Complex document', description: 'Round-trip corpus',
    numbering: {
      config: [
        {
          reference: 'steps',
          levels: [
            { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START },
            { level: 1, format: LevelFormat.LOWER_LETTER, text: '%2)', alignment: AlignmentType.START },
          ],
        },
      ],
    },
    footnotes: { 1: { children: [new Paragraph('This is footnote one.')] }, 2: { children: [new Paragraph('Footnote two text.')] } },
    comments: {
      children: [{ id: 0, author: 'Reviewer A', initials: 'RA', date: new Date('2026-01-02T10:00:00Z'), children: [new Paragraph('Please check this sentence.')] }],
    },
    sections: [
      {
        headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun('Complex Fixture — header')] })] }) },
        footers: {
          default: new Footer({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun('Page '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun(' of '), new TextRun({ children: [PageNumber.TOTAL_PAGES] })] })],
          }),
        },
        children: [
          new Paragraph({ text: 'Complex Fixture', heading: HeadingLevel.TITLE }),
          new Paragraph({ text: 'Formatting', heading: HeadingLevel.HEADING_1 }),
          new Paragraph({
            children: [
              new TextRun({ text: 'Bold', bold: true }), new TextRun(' '),
              new TextRun({ text: 'Italic', italics: true }), new TextRun(' '),
              new TextRun({ text: 'Underline', underline: { type: UnderlineType.SINGLE } }), new TextRun(' '),
              new TextRun({ text: 'Double underline', underline: { type: UnderlineType.DOUBLE } }), new TextRun(' '),
              new TextRun({ text: 'Strike', strike: true }), new TextRun(' '),
              new TextRun({ text: 'Red 14pt Georgia', color: 'C00000', size: 28, font: 'Georgia' }), new TextRun(' '),
              new TextRun({ text: 'Highlighted', highlight: 'yellow' }), new TextRun(' '),
              new TextRun({ text: 'Shaded', shading: { type: ShadingType.CLEAR, fill: 'DDEEFF', color: 'auto' } }), new TextRun(' H'),
              new TextRun({ text: '2', subScript: true }), new TextRun('O and E=mc'), new TextRun({ text: '2', superScript: true }),
              new TextRun({ text: ' SMALL CAPS', smallCaps: true }),
            ],
          }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun('Centered paragraph')] }),
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, indent: { left: 720, right: 360, firstLine: 360 }, spacing: { before: 240, after: 120, line: 360 }, children: [new TextRun('Justified, indented paragraph with 1.5 line spacing and custom space before and after. '.repeat(3))] }),
          new Paragraph({ indent: { left: 1440, hanging: 720 }, children: [new TextRun('Hanging indent paragraph.\tAfter a tab.')] }),
          new Paragraph({ text: 'Lists', heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ text: 'Bullet one', bullet: { level: 0 } }),
          new Paragraph({ text: 'Bullet nested', bullet: { level: 1 } }),
          new Paragraph({ text: 'Bullet two', bullet: { level: 0 } }),
          new Paragraph({ text: 'Step one', numbering: { reference: 'steps', level: 0 } }),
          new Paragraph({ text: 'Sub-step a', numbering: { reference: 'steps', level: 1 } }),
          new Paragraph({ text: 'Step two', numbering: { reference: 'steps', level: 0 } }),
          new Paragraph({ text: 'Tables', heading: HeadingLevel.HEADING_2 }),
          new Table({
            width: { size: 9000, type: WidthType.DXA },
            columnWidths: [3000, 3000, 3000],
            rows: [
              new TableRow({ tableHeader: true, children: ['Name', 'Role', 'Score'].map((t) => new TableCell({ shading: { fill: 'D9E2F3', type: ShadingType.CLEAR, color: 'auto' }, children: [new Paragraph({ children: [new TextRun({ text: t, bold: true })] })] })) }),
              new TableRow({ children: [new TableCell({ columnSpan: 2, children: [new Paragraph('Merged across two columns')] }), new TableCell({ rowSpan: 2, verticalMerge: VerticalMergeType.RESTART, children: [new Paragraph('Merged down')] })] }),
              new TableRow({ children: [new TableCell({ children: [new Paragraph('Ada')] }), new TableCell({ children: [new Paragraph('Engineer')] })] }),
              new TableRow({
                children: [
                  new TableCell({ children: [new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph('Nested A')] }), new TableCell({ children: [new Paragraph('Nested B')] })] })] }), new Paragraph('')] }),
                  new TableCell({ children: [new Paragraph('Grace')] }),
                  new TableCell({ children: [new Paragraph('99')] }),
                ],
              }),
            ],
          }),
          new Paragraph({ text: 'Media and links', heading: HeadingLevel.HEADING_2 }),
          new Paragraph({ children: [new TextRun('Inline image: '), new ImageRun({ type: 'png', data: image, transformation: { width: 120, height: 80 }, altText: { title: 'Swatch', description: 'Gradient swatch', name: 'swatch' } })] }),
          new Paragraph({ children: [new TextRun('Visit '), new ExternalHyperlink({ link: 'https://example.com/docs?q=1', children: [new TextRun({ text: 'example.com', style: 'Hyperlink' })] }), new TextRun(' or jump to '), new InternalHyperlink({ anchor: 'target', children: [new TextRun({ text: 'the bookmark', style: 'Hyperlink' })] }), new TextRun('.')] }),
          new Paragraph({ children: [new TextRun('Footnoted sentence'), new FootnoteReferenceRun(1), new TextRun(' and another'), new FootnoteReferenceRun(2), new TextRun('.')] }),
          new Paragraph({ children: [new CommentRangeStart(0), new TextRun('This sentence has a comment.'), new CommentRangeEnd(0), new TextRun({ children: [new CommentReference(0)] })] }),
          new Paragraph({ children: [new TextRun('Tracked: '), new InsertedTextRun({ text: 'inserted words', id: 1, author: 'Editor', date: '2026-02-01T09:00:00Z' }), new TextRun(' and '), new DeletedTextRun({ text: 'deleted words', id: 2, author: 'Editor', date: '2026-02-01T09:01:00Z' })] }),
          new Paragraph({ children: [new Bookmark({ id: 'target', children: [new TextRun('Bookmark target paragraph')] })] }),
          new Paragraph({ children: [new TextRun('Unicode: café — 日本語テキスト — Ελληνικά — emoji 🎉 — Ω≈ç√')] }),
          new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'مرحبا بالعالم — نص عربي', rightToLeft: true })] }),
          new Paragraph({ children: [new PageBreak()] }),
          new Paragraph({ text: 'After the page break', heading: HeadingLevel.HEADING_1 }),
          new Paragraph('Content on page two.'),
        ],
      },
      {
        properties: { type: SectionType.NEXT_PAGE, page: { size: { orientation: PageOrientation.LANDSCAPE, width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } } },
        children: [new Paragraph({ text: 'Landscape section', heading: HeadingLevel.HEADING_1 }), new Paragraph('This section is landscape with different margins.')],
      },
    ],
  }),
)

// 3. table of contents (field) + headings
await save(
  'toc.docx',
  new Document({
    features: { updateFields: true },
    sections: [{ children: [new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-3' }), new Paragraph({ text: 'Chapter 1', heading: HeadingLevel.HEADING_1 }), new Paragraph('Text'), new Paragraph({ text: 'Section 1.1', heading: HeadingLevel.HEADING_2 }), new Paragraph('More text')] }],
  }),
)

// 4. large document
const big = []
for (let i = 0; i < 3000; i++) {
  if (i % 100 === 0) big.push(new Paragraph({ text: `Chapter ${i / 100 + 1}`, heading: HeadingLevel.HEADING_1 }))
  big.push(new Paragraph({ children: [new TextRun(`Paragraph ${i}. `), new TextRun({ text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. ', bold: i % 3 === 0 }), new TextRun('Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.')] }))
}
await save('large.docx', new Document({ sections: [{ children: big }] }))

// 5. unsupported OOXML: inject a text box shape, an equation and an unknown extension element into a generated file
{
  const base = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('Before'), new Paragraph('Anchor paragraph'), new Paragraph('After')] }] }))
  const zip = await JSZip.loadAsync(base)
  let xml = await zip.file('word/document.xml').async('string')
  const textbox =
    '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="77" name="Text Box 77"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Text inside a text box</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape style="width:144pt;height:36pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Text inside a text box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>'
  const math = '<m:oMath><m:r><m:t>a²+b²=c²</m:t></m:r></m:oMath>'
  const unknown = '<w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checked w14:val="1"/></w14:checkbox>'
  xml = xml.replace(/(<w:p>(?:<w:pPr>.*?<\/w:pPr>)?<w:r><w:t[^>]*>Anchor paragraph<\/w:t><\/w:r>)/, `$1${textbox}${math}`)
  xml = xml.replace('<w:body>', `<w:body><w:customXml w:element="note"><w:p><w:r><w:t>Inside custom XML</w:t></w:r></w:p></w:customXml>`)
  xml = xml.replace(/(<w:t[^>]*>After<\/w:t><\/w:r>)/, `$1${unknown}`)
  if (!xml.includes('xmlns:wps=')) xml = xml.replace('<w:document ', '<w:document xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ')
  if (!xml.includes('xmlns:m=')) xml = xml.replace('<w:document ', '<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ')
  if (!xml.includes('xmlns:v=')) xml = xml.replace('<w:document ', '<w:document xmlns:v="urn:schemas-microsoft-com:vml" ')
  if (!xml.includes('xmlns:mc=')) xml = xml.replace('<w:document ', '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ')
  zip.file('word/document.xml', xml)
  zip.file('customXml/item1.xml', '<root><value>kept</value></root>')
  const buf = await zip.generateAsync({ type: 'nodebuffer' })
  fs.writeFileSync(path.join(out, 'unsupported.docx'), buf)
  process.stdout.write(`wrote unsupported.docx ${buf.length} bytes\n`)
}

// 6. malformed inputs
fs.writeFileSync(path.join(out, 'not-a-zip.docx'), 'This is plain text pretending to be a DOCX file.')
{
  const good = fs.readFileSync(path.join(out, 'simple.docx'))
  fs.writeFileSync(path.join(out, 'truncated.docx'), good.subarray(0, Math.floor(good.length / 2)))
  const zip = await JSZip.loadAsync(good)
  zip.file('word/document.xml', '<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;">]><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>&lol2;</w:t></w:r></w:p></w:body></w:document>')
  fs.writeFileSync(path.join(out, 'xxe.docx'), await zip.generateAsync({ type: 'nodebuffer' }))
  const z2 = await JSZip.loadAsync(good)
  z2.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>unclosed')
  fs.writeFileSync(path.join(out, 'malformed-xml.docx'), await z2.generateAsync({ type: 'nodebuffer' }))
  const z3 = await JSZip.loadAsync(good)
  z3.file('word/bomb.bin', Buffer.alloc(50 * 1024 * 1024, 0))
  fs.writeFileSync(path.join(out, 'zipbomb.docx'), await z3.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } }))
  process.stdout.write('wrote malformed fixtures\n')
}
void BorderStyle

// 7. editable objects: a column chart (with an embedded-workbook link), a SmartArt diagram with its
// cached drawing, a custom XML part and a bibliography sources part, assembled by hand
{
  const base = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('Chart below'), new Paragraph('CHART_HERE'), new Paragraph('DIAGRAM_HERE'), new Paragraph('End')] }] }))
  const zip = await JSZip.loadAsync(base)
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const ser = (i, name, vals) =>
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${'BC'[i]}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
    `<c:spPr><a:solidFill><a:srgbClr val="${['4472C4', 'ED7D31'][i]}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>` +
    `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>Sheet1!$${'BC'[i]}$2:$${'BC'[i]}$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>${vals.map((v, k) => `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val></c:ser>`
  const chartXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:roundedCorners val="0"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Quarterly sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` +
    `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser(0, 'North', [10, 14, 9])}${ser(1, 'South', [7, 11, 15])}<c:gapWidth val="150"/><c:axId val="111"/><c:axId val="222"/></c:barChart>` +
    `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx></c:plotArea>` +
    `<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`
  zip.file('word/charts/chart1.xml', chartXml)
  zip.file('word/charts/_rels/chart1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Workbook1.xlsx"/></Relationships>`)
  zip.file('word/embeddings/Workbook1.xlsx', Buffer.from('PK placeholder workbook bytes'))
  // SmartArt: three nodes under a doc root, presentation points mapping drawing shapes to nodes
  const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
  const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
  const node = (id, t) => `<dgm:pt modelId="${id}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p></dgm:t></dgm:pt>`
  const pres = (id, assoc) => `<dgm:pt modelId="${id}" type="pres"><dgm:prSet presAssocID="${assoc}" presName="node" presStyleLbl="node1" presStyleIdx="0" presStyleCnt="3"/><dgm:spPr/></dgm:pt>`
  const dataXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst><dgm:pt modelId="{D0}" type="doc"><dgm:prSet loTypeId="urn:diagrams/layout/default"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>` +
    node('{N1}', 'Plan') + node('{N2}', 'Build') + node('{N3}', 'Ship') + pres('{P1}', '{N1}') + pres('{P2}', '{N2}') + pres('{P3}', '{N3}') +
    `</dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="{C1}" srcId="{D0}" destId="{N1}" srcOrd="0" destOrd="0"/><dgm:cxn modelId="{C2}" srcId="{D0}" destId="{N2}" srcOrd="1" destOrd="0"/><dgm:cxn modelId="{C3}" srcId="{D0}" destId="{N3}" srcOrd="2" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/>` +
    `<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="${DSP}" relId="rIdDraw" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>`
  const shape = (id, x, t) => `<dsp:sp modelId="${id}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="1600000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2000"/><a:t>${t}</a:t></a:r></a:p></dsp:txBody></dsp:sp>`
  const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dsp:drawing xmlns:dsp="${DSP}" xmlns:dgm="${DGM}" xmlns:a="${A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${shape('{P1}', 0, 'Plan')}${shape('{P2}', 1800000, 'Build')}${shape('{P3}', 3600000, 'Ship')}</dsp:spTree></dsp:drawing>`
  zip.file('word/diagrams/data1.xml', dataXml)
  zip.file('word/diagrams/drawing1.xml', drawingXml)
  const minimal = (root, ns) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:${root} xmlns:dgm="${ns}" xmlns:a="${A}" uniqueId="urn:fixture"/>`
  zip.file('word/diagrams/layout1.xml', minimal('layoutDef', DGM))
  zip.file('word/diagrams/quickStyle1.xml', minimal('styleDef', DGM))
  zip.file('word/diagrams/colors1.xml', minimal('colorsDef', DGM))
  // custom XML and bibliography sources
  zip.file('customXml/item1.xml', '<?xml version="1.0" encoding="UTF-8"?>\n<invoice xmlns="urn:example:invoice"><number>42</number><customer>Acme</customer></invoice>')
  zip.file('customXml/item2.xml', `<?xml version="1.0" encoding="UTF-8"?>\n<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" StyleName="APA"><b:Source><b:Tag>Knu97</b:Tag><b:SourceType>Book</b:SourceType><b:Author><b:Author><b:NameList><b:Person><b:Last>Knuth</b:Last><b:First>Donald</b:First></b:Person></b:NameList></b:Author></b:Author><b:Title>The Art of Computer Programming</b:Title><b:Year>1997</b:Year><b:Publisher>Addison-Wesley</b:Publisher></b:Source></b:Sources>`)
  // relationships and content types
  let rels = await zip.file('word/_rels/document.xml.rels').async('string')
  rels = rels.replace('</Relationships>',
    `<Relationship Id="rIdChart" Type="${R}/chart" Target="charts/chart1.xml"/>` +
    `<Relationship Id="rIdDm" Type="${R}/diagramData" Target="diagrams/data1.xml"/><Relationship Id="rIdLo" Type="${R}/diagramLayout" Target="diagrams/layout1.xml"/>` +
    `<Relationship Id="rIdQs" Type="${R}/diagramQuickStyle" Target="diagrams/quickStyle1.xml"/><Relationship Id="rIdCs" Type="${R}/diagramColors" Target="diagrams/colors1.xml"/>` +
    `<Relationship Id="rIdDraw" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/>` +
    `<Relationship Id="rIdCx1" Type="${R}/customXml" Target="../customXml/item1.xml"/><Relationship Id="rIdCx2" Type="${R}/customXml" Target="../customXml/item2.xml"/></Relationships>`)
  zip.file('word/_rels/document.xml.rels', rels)
  let ct = await zip.file('[Content_Types].xml').async('string')
  ct = ct.replace('</Types>',
    '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>' +
    '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
    '<Override PartName="/word/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>' +
    '<Override PartName="/word/diagrams/layout1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>' +
    '<Override PartName="/word/diagrams/quickStyle1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/>' +
    '<Override PartName="/word/diagrams/colors1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>' +
    '<Override PartName="/word/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/></Types>')
  if (!/Extension="xml"/.test(ct)) ct = ct.replace('</Types>', '<Default Extension="xml" ContentType="application/xml"/></Types>')
  zip.file('[Content_Types].xml', ct)
  let xml = await zip.file('word/document.xml').async('string')
  const inline = (cx, cy, id, gd) => `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Object ${id}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="${A}">${gd}</a:graphic></wp:inline></w:drawing></w:r>`
  const chartRun = inline(4572000, 2743200, 501, `<a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rIdChart"/></a:graphicData>`)
  const dgmRun = inline(5486400, 1600200, 502, `<a:graphicData uri="${DGM}"><dgm:relIds xmlns:dgm="${DGM}" xmlns:r="${R}" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/></a:graphicData>`)
  xml = xml.replace(/<w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>CHART_HERE<\/w:t><\/w:r>/, chartRun).replace(/<w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>DIAGRAM_HERE<\/w:t><\/w:r>/, dgmRun)
  if (!xml.includes('xmlns:wp=')) xml = xml.replace('<w:document ', '<w:document xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ')
  zip.file('word/document.xml', xml)
  fs.writeFileSync(path.join(out, 'objects.docx'), await zip.generateAsync({ type: 'nodebuffer' }))
  process.stdout.write('wrote objects.docx\n')
}
