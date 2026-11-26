import { describe, expect, it } from 'vitest'
import { generateHTML } from '@tiptap/core'
import { schemaExtensions } from '../../src/model/extensions'
import { safeColor, safeFontFamily, safeLength } from '../../src/model/safeCss'

const render = (content: any[]) => generateHTML({ type: 'doc', content } as any, schemaExtensions())
const EVIL = 'red;background-image:url(https://evil.example/t)'

describe('style attribute injection', () => {
  it('validators accept real values and reject CSS injection', () => {
    for (const ok of ['#c00000', '#abc', 'red', 'rgb(10, 20, 30)', 'rgba(0,0,0,0.5)']) expect(safeColor(ok), ok).toBe(ok)
    for (const bad of [EVIL, 'url(x)', 'red;', 'expression(alert(1))', '#zzz', '']) expect(safeColor(bad), bad).toBeNull()
    expect(safeFontFamily('"Times New Roman", serif')).toBe('"Times New Roman", serif')
    expect(safeFontFamily('Arial;background:url(x)')).toBeNull()
    expect(safeLength('11pt')).toBe('11pt')
    expect(safeLength('11pt;color:red')).toBeNull()
  })

  it('never renders injected CSS from document attributes', () => {
    const html = render([
      {
        type: 'paragraph',
        attrs: { shading: EVIL, textAlign: `center;${EVIL}`, indentLeft: '1;x:url(y)', lineSpacing: 'calc(1)' },
        content: [
          { type: 'text', text: 'styled', marks: [{ type: 'textStyle', attrs: { color: EVIL, backgroundColor: EVIL, fontFamily: `Arial;${EVIL}`, fontSize: `12pt;${EVIL}`, letterSpacing: `1px;${EVIL}` } }, { type: 'highlight', attrs: { color: EVIL } }] },
        ],
      },
      { type: 'bulletList', attrs: { listStyle: `disc;${EVIL}` }, content: [{ type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'item' }] }] }] },
      { type: 'table', content: [{ type: 'tableRow', content: [{ type: 'tableCell', attrs: { background: EVIL, verticalAlign: `top;${EVIL}` }, content: [{ type: 'paragraph' }] }] }] },
      { type: 'paragraph', content: [{ type: 'shape', attrs: { kind: 'rect', width: 50, height: 50, fill: EVIL, stroke: '#000000', strokeWidth: 1, text: '<script>alert(1)</script>' } }] },
    ])
    expect(html).not.toContain('evil.example')
    expect(html).not.toContain('url(')
    expect(html).not.toContain('<script')
    expect(html).toContain('styled')
  })

  it('still renders legitimate formatting', () => {
    const html = render([
      { type: 'paragraph', attrs: { shading: '#ffff00', textAlign: 'center' }, content: [{ type: 'text', text: 'ok', marks: [{ type: 'textStyle', attrs: { color: '#c00000', fontFamily: 'Georgia', fontSize: '14pt' } }, { type: 'highlight', attrs: { color: '#00ff00' } }] }] },
    ])
    const d = new DOMParser().parseFromString(html, 'text/html')
    const p = d.querySelector('p')!
    const span = d.querySelector('span')!
    const mark = d.querySelector('mark')!
    expect(p.style.backgroundColor).toBe('rgb(255, 255, 0)')
    expect(p.style.textAlign).toBe('center')
    expect(span.style.color).toBe('rgb(192, 0, 0)')
    expect(span.style.fontFamily).toBe('Georgia')
    expect(span.style.fontSize).toBe('14pt')
    expect(mark.style.backgroundColor).toBe('rgb(0, 255, 0)')
  })
})

describe('server-side content sanitising', () => {
  it('strips injected style values from saved documents but keeps real formatting', async () => {
    // @ts-expect-error plain JS server module
    const { sanitizeContent } = await import('../../server/validate.js')
    const out = sanitizeContent({
      type: 'doc',
      content: [
        { type: 'paragraph', attrs: { shading: EVIL, textAlign: 'center' }, content: [{ type: 'text', text: 'x', marks: [{ type: 'textStyle', attrs: { color: EVIL, fontSize: '12pt', fontFamily: 'Georgia' } }, { type: 'link', attrs: { href: 'javascript:alert(1)' } }] }] },
      ],
    })
    const p = out.content[0]
    expect(p.attrs.shading).toBeNull()
    expect(p.attrs.textAlign).toBe('center')
    expect(p.content[0].marks).toEqual([{ type: 'textStyle', attrs: { color: null, fontSize: '12pt', fontFamily: 'Georgia' } }])
  })
})
