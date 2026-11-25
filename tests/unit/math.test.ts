import { describe, expect, it } from 'vitest'
import { MathSyntaxError, linearBlocker, ommlToMathml, ommlToTree, parseLinear, toLinear, treeToOmml } from '../../src/objects/math'

const rt = (s: string) => {
  const t = parseLinear(s)
  const omml = treeToOmml(t)
  const back = toLinear(ommlToTree(omml))
  expect(treeToOmml(parseLinear(back))).toBe(omml)
  return { t, omml, back }
}

describe('equations', () => {
  it('parses the linear syntax into structures', () => {
    expect(rt('a/b').omml).toContain('<m:f><m:num><m:r><m:t xml:space="preserve">a</m:t></m:r></m:num><m:den><m:r><m:t xml:space="preserve">b</m:t></m:r></m:den></m:f>')
    expect(rt('x^2+y_i').omml).toMatch(/<m:sSup><m:e>.*x.*<\/m:e><m:sup>.*2.*<\/m:sup><\/m:sSup>.*\+.*<m:sSub>/)
    expect(rt('(a+b)/(c+d)').omml).not.toContain('<m:d>')
    expect(rt('(a+b)^2').omml).toContain('<m:sSup><m:e><m:d>')
    expect(rt('√(b^2-4ac)').omml).toContain('<m:rad>')
    expect(rt('\\sum_(i=1)^n i^2').omml).toContain('<m:chr m:val="∑"/>')
    expect(rt('x = (-b ± √(b^2-4ac))/(2a)').omml).toContain('<m:f>')
    expect(rt('\\alpha+\\beta \\le \\pi').omml).toContain('α')
    expect(rt('e^(i\\pi)+1=0').omml).toContain('<m:sSup>')
    expect(rt('\\int_0^1 f(x) dx').omml).toContain('<m:limLoc m:val="subSup"/>')
    expect(rt('x_i^2').omml).toContain('<m:sSubSup>')
    expect(rt('2x^2+3x-1').omml).toContain('<m:sSup>')
    expect(rt('a\\^b \\/ c').omml).toContain('a^b / c')
    expect(rt('x^-1').omml).toContain('-1')
    expect(rt('\\sum_i x + y').omml).toContain('<m:nary>')
  })
  it('rejects malformed input with a readable error', () => {
    for (const bad of ['a/', '(a+b', 'a+b)', '^2', '\\nope', 'x^']) expect(() => parseLinear(bad), bad).toThrow(MathSyntaxError)
  })
  it('round-trips imported OMML and flags structures it cannot express', () => {
    const omml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>E=m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>'
    expect(linearBlocker(omml)).toBeNull()
    expect(toLinear(ommlToTree(omml))).toBe('E=mc^2')
    const matrix = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:m><m:mr><m:e><m:r><m:t>1</m:t></m:r></m:e></m:mr></m:m></m:oMath>'
    expect(linearBlocker(matrix)).toMatch(/matrices/)
    expect(ommlToMathml(matrix)).toContain('<mn>1</mn>')
    expect(ommlToMathml(omml)).toContain('<msup>')
  })
})
