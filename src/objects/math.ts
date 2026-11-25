// Equations: OMML (the DOCX math format) <-> a linear text syntax <-> MathML for display.
//
// Linear syntax
//   a/b            fraction (operands in parentheses lose them: (a+b)/(c+d))
//   x^2  x_i  x_i^2  superscript / subscript (script arguments in parentheses lose them: x^(n+1))
//   √x  √(a+b)  \sqrt(a+b)   square root
//   ∑_(i=1)^n x   ∫_0^1 f   ∏   n-ary operators with optional limits, followed by their operand
//   (a+b)          visible parentheses; [ ] and | | are plain text
//   {a+b}          invisible grouping, e.g. {a+b}^2
//   \alpha \pi \infty \le \ge \ne \pm \times \cdot \to \sum \int \prod   named symbols
//   \^ \_ \/ \( \) \{ \} \\ \√      literal characters

import { NS, children, parseXml, serializeXml } from '../docx/xml'

const M = NS.m

export type MathNode =
  | { t: 'row'; items: MathNode[] }
  | { t: 'text'; s: string }
  | { t: 'frac'; num: MathNode; den: MathNode }
  | { t: 'script'; base: MathNode; sub?: MathNode; sup?: MathNode }
  | { t: 'rad'; body: MathNode }
  | { t: 'nary'; op: string; sub?: MathNode; sup?: MathNode; body: MathNode }
  | { t: 'delim'; body: MathNode }

export class MathSyntaxError extends Error {}

const NARY = new Set(['∑', '∫', '∏', '∬', '∭', '∮', '⋃', '⋂'])
const SYMBOLS: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', iota: 'ι', kappa: 'κ',
  lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', chi: 'χ',
  psi: 'ψ', omega: 'ω', Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ',
  Psi: 'Ψ', Omega: 'Ω', infty: '∞', le: '≤', ge: '≥', ne: '≠', approx: '≈', pm: '±', mp: '∓', times: '×', div: '÷',
  cdot: '⋅', to: '→', rightarrow: '→', leftarrow: '←', in: '∈', notin: '∉', subset: '⊂', supset: '⊃', cup: '∪',
  cap: '∩', partial: '∂', nabla: '∇', forall: '∀', exists: '∃', degree: '°', sum: '∑', int: '∫', prod: '∏',
  iint: '∬', oint: '∮', sqrt: '√', equiv: '≡', propto: '∝', ldots: '…', cdots: '⋯',
}
const SPECIAL = new Set(['^', '_', '/', '(', ')', '{', '}', '\\', '√'])

const row = (items: MathNode[]): MathNode => (items.length === 1 ? items[0] : { t: 'row', items })

// ---------------------------------------------------------------------------
// linear -> tree

export function parseLinear(src: string): MathNode {
  // tokens: single characters, with escapes and \names resolved
  const toks: { ch: string; lit: boolean }[] = []
  for (let i = 0; i < src.length; ) {
    const cp = src.codePointAt(i)!
    const ch = String.fromCodePoint(cp)
    if (ch === '\\') {
      const name = /^[A-Za-z]+/.exec(src.slice(i + 1))?.[0]
      if (name) {
        const sym = SYMBOLS[name]
        if (!sym) throw new MathSyntaxError(`Unknown symbol \\${name}`)
        toks.push({ ch: sym, lit: false })
        i += 1 + name.length
        continue
      }
      const next = src.codePointAt(i + 1)
      if (next === undefined) throw new MathSyntaxError('A backslash must be followed by a symbol name or a character')
      toks.push({ ch: String.fromCodePoint(next), lit: true })
      i += 1 + String.fromCodePoint(next).length
      continue
    }
    toks.push({ ch, lit: false })
    i += ch.length
  }
  let pos = 0
  const peek = () => toks[pos]
  const isOp = (t: { ch: string; lit: boolean } | undefined, c: string) => !!t && !t.lit && t.ch === c
  const isWordChar = (t: { ch: string; lit: boolean } | undefined) => !!t && (t.lit || /[\p{L}\p{N}.′]/u.test(t.ch)) && !(!t.lit && (SPECIAL.has(t.ch) || NARY.has(t.ch)))

  const parseExpr = (closers: string[]): MathNode => {
    const items: MathNode[] = []
    while (pos < toks.length && !(closers.length && closers.some((c) => isOp(peek(), c)))) {
      items.push(...parseTerm())
      // fraction: binds the last operand with the next one
      while (isOp(peek(), '/')) {
        pos++
        const num = items.pop()
        if (!num) throw new MathSyntaxError('A fraction needs a numerator before “/”')
        const den = parseOperand(true)
        items.push({ t: 'frac', num: unwrapGroup(num), den })
      }
    }
    return row(mergeTexts(items))
  }
  /** One term: an atom plus any scripts. May return several text pieces before the scripted atom. */
  const parseTerm = (): MathNode[] => {
    const t = peek()
    if (!t) return []
    if (!t.lit && (t.ch === ')' || t.ch === '}')) throw new MathSyntaxError(`Unmatched “${t.ch}”`)
    if (!t.lit && (t.ch === '^' || t.ch === '_' || t.ch === '/')) throw new MathSyntaxError(`“${t.ch}” needs something before it`)
    if (!t.lit && NARY.has(t.ch)) {
      pos++
      const n: { t: 'nary'; op: string; sub?: MathNode; sup?: MathNode; body: MathNode } = { t: 'nary', op: t.ch, body: { t: 'row', items: [] } }
      for (let k = 0; k < 2; k++) {
        if (isOp(peek(), '_') && !n.sub) {
          pos++
          n.sub = parseOperand(true)
        } else if (isOp(peek(), '^') && !n.sup) {
          pos++
          n.sup = parseOperand(true)
        }
      }
      while (pos < toks.length && !peek().lit && /\s/.test(peek().ch)) pos++
      n.body = pos < toks.length && !isOp(peek(), ')') && !isOp(peek(), '}') ? row(parseTerm()) : { t: 'row', items: [] }
      return [n]
    }
    // gather plain text; the trailing word becomes the script base
    if (isWordChar(t) || (!SPECIAL.has(t.ch) && !t.lit)) {
      let s = ''
      while (pos < toks.length && (peek().lit || (!SPECIAL.has(peek().ch) && !NARY.has(peek().ch)))) s += toks[pos++].ch
      const hasScript = isOp(peek(), '^') || isOp(peek(), '_') || isOp(peek(), '/')
      if (!hasScript) return [{ t: 'text', s }]
      // fraction operands take the whole trailing word (ab/c); script bases take one letter or number (mc^2)
      const m = (isOp(peek(), '/') ? /[\p{L}\p{N}.′]+$/u : /(?:\p{N}+(?:\.\p{N}+)?|\p{L}′*)$/u).exec(s)
      const word = m ? m[0] : s.slice(-1)
      const before = s.slice(0, s.length - word.length)
      const base: MathNode = { t: 'text', s: word }
      return [...(before ? [{ t: 'text', s: before } as MathNode] : []), withScripts(base)]
    }
    return [withScripts(parseAtom())]
  }
  const withScripts = (base: MathNode): MathNode => {
    let sub: MathNode | undefined
    let sup: MathNode | undefined
    for (let k = 0; k < 2; k++) {
      if (isOp(peek(), '_') && !sub) {
        pos++
        sub = parseOperand(true)
      } else if (isOp(peek(), '^') && !sup) {
        pos++
        sup = parseOperand(true)
      }
    }
    if (isOp(peek(), '^') || isOp(peek(), '_')) throw new MathSyntaxError('Use braces for nested scripts, e.g. x^{a^b}')
    return sub || sup ? { t: 'script', base, sub, sup } : base
  }
  const parseAtom = (): MathNode => {
    const t = peek()
    if (!t) throw new MathSyntaxError('The equation ends too early')
    if (isOp(t, '(')) {
      pos++
      const body = parseExpr([')'])
      if (!isOp(peek(), ')')) throw new MathSyntaxError('Missing “)”')
      pos++
      return { t: 'delim', body }
    }
    if (isOp(t, '{')) {
      pos++
      const body = parseExpr(['}'])
      if (!isOp(peek(), '}')) throw new MathSyntaxError('Missing “}”')
      pos++
      return { t: 'row', items: body.t === 'row' ? body.items : [body] }
    }
    if (isOp(t, '√')) {
      pos++
      return { t: 'rad', body: parseOperand(true) }
    }
    if (isWordChar(t)) {
      let s = ''
      while (isWordChar(peek())) s += toks[pos++].ch
      return { t: 'text', s }
    }
    pos++
    return { t: 'text', s: t.ch }
  }
  /** Operand of / ^ _ √: one atom; parentheses around it are grouping and are dropped when strip is set. */
  const parseOperand = (strip: boolean): MathNode => {
    const t = peek()
    if (!t) throw new MathSyntaxError('Something is missing after an operator')
    if (isOp(t, ')') || isOp(t, '}')) throw new MathSyntaxError('Something is missing after an operator')
    if (!t.lit && NARY.has(t.ch)) return row(parseTerm())
    const a = isOp(t, '-') || isOp(t, '−') ? (pos++, { t: 'text', s: t.ch } as MathNode) : null
    const atom = parseAtom()
    const v = strip && atom.t === 'delim' ? atom.body : atom
    return a ? row(mergeTexts([a, ...(v.t === 'row' ? v.items : [v])])) : v
  }
  const result = parseExpr([])
  if (pos < toks.length) throw new MathSyntaxError(`Unexpected “${toks[pos].ch}”`)
  return result
}

function unwrapGroup(n: MathNode): MathNode {
  return n.t === 'delim' ? n.body : n
}

function mergeTexts(items: MathNode[]): MathNode[] {
  const out: MathNode[] = []
  for (const it of items) {
    const last = out[out.length - 1]
    if (it.t === 'text' && last?.t === 'text') out[out.length - 1] = { t: 'text', s: last.s + it.s }
    else out.push(it)
  }
  return out
}

// ---------------------------------------------------------------------------
// tree -> linear

const escText = (s: string) => [...s].map((c) => (SPECIAL.has(c) || NARY.has(c) ? '\\' + c : c)).join('')
const isWord = (s: string) => /^[\p{L}\p{N}.′]+$/u.test(s)
const isToken = (s: string) => /^(?:\p{N}+(?:\.\p{N}+)?|\p{L}′*)$/u.test(s)

export function toLinear(n: MathNode): string {
  switch (n.t) {
    case 'text':
      return escText(n.s)
    case 'row':
      return joinRow(n.items)
    case 'delim':
      return `(${toLinear(n.body)})`
    case 'frac':
      return `${operand(n.num)}/${operand(n.den)}`
    case 'rad':
      return `√${operand(n.body)}`
    case 'script': {
      const base = n.base.t === 'text' && isToken(n.base.s) ? escText(n.base.s) : n.base.t === 'delim' || n.base.t === 'rad' ? toLinear(n.base) : `{${toLinear(n.base)}}`
      return base + (n.sub ? `_${operand(n.sub)}` : '') + (n.sup ? `^${operand(n.sup)}` : '')
    }
    case 'nary': {
      const body = n.body.t === 'row' && n.body.items.length === 0 ? '' : n.body.t === 'row' ? `{${toLinear(n.body)}}` : bodyText(n.body)
      return escOp(n.op) + (n.sub ? `_${operand(n.sub)}` : '') + (n.sup ? `^${operand(n.sup)}` : '') + (body ? (/^[\\{(√∑∫∏]/.test(body) ? body : ' ' + body) : '')
    }
  }
}

const escOp = (op: string) => op

function bodyText(n: MathNode): string {
  const s = toLinear(n)
  // a plain-text operand followed by more text must be grouped so the parser stops at the right place
  return n.t === 'text' && !isWord(n.s) ? `{${s}}` : s
}

function joinRow(items: MathNode[]): string {
  let out = ''
  for (let i = 0; i < items.length; i++) {
    const it = items[i]
    let s = toLinear(it)
    const prev = items[i - 1]
    // a script/fraction that follows a word would grab the word's tail as its base: separate with braces
    if (prev?.t === 'text' && /[\p{L}\p{N}.′]$/u.test(prev.s)) {
      const digitBase = it.t === 'script' && it.base.t === 'text' && /^\p{N}/u.test(it.base.s) && /[\p{N}.]$/u.test(prev.s)
      const primeBase = it.t === 'script' && it.base.t === 'text' && /^′/u.test(it.base.s)
      if (it.t === 'frac' || digitBase || primeBase) s = `{}${s}`
    }
    // a nary operand would swallow following text: wrap the nary when more items follow
    if (it.t === 'nary' && i < items.length - 1) s = it.body.t === 'row' && it.body.items.length === 0 ? `${s}{}` : `{${s}}`
    out += s
  }
  return out
}

function operand(n: MathNode): string {
  if (n.t === 'text' && isWord(n.s)) return escText(n.s)
  if (n.t === 'delim') return `(${toLinear(n)})`
  if (n.t === 'rad') return toLinear(n)
  return `(${toLinear(n)})`
}

// ---------------------------------------------------------------------------
// OMML <-> tree

const SUPPORTED = new Set(['oMath', 'oMathPara', 'oMathParaPr', 'r', 't', 'rPr', 'f', 'fPr', 'num', 'den', 'sSup', 'sSupPr', 'sSub', 'sSubPr', 'sSubSup', 'sSubSupPr', 'sup', 'sub', 'e', 'rad', 'radPr', 'deg', 'degHide', 'nary', 'naryPr', 'chr', 'limLoc', 'subHide', 'supHide', 'grow', 'd', 'dPr', 'ctrlPr', 'jc', 'sty', 'scr', 'nor', 'lit', 'type', 'begChr', 'endChr'])

/** Why an OMML equation cannot be edited as linear text, or null when it can. */
export function linearBlocker(xml: string): string | null {
  let doc: Document
  try {
    doc = parseXml(xml, 'equation')
  } catch {
    return 'The equation XML could not be read.'
  }
  const bad = new Set<string>()
  for (const e of Array.from(doc.getElementsByTagName('*'))) {
    if (e.namespaceURI === M && !SUPPORTED.has(e.localName)) bad.add(e.localName)
    if (e.namespaceURI === M && e.localName === 'deg' && e.textContent) bad.add('nth roots')
    if (e.namespaceURI === M && e.localName === 'begChr' && e.getAttributeNS(M, 'val') !== '(') bad.add('brackets other than ( )')
    if (e.namespaceURI === M && e.localName === 'endChr' && e.getAttributeNS(M, 'val') !== ')') bad.add('brackets other than ( )')
    if (e.namespaceURI === M && e.localName === 'd' && children(e, M, 'e').length > 1) bad.add('separated delimiters')
    if (e.namespaceURI === M && e.localName === 'oMathPara' && children(e, M, 'oMath').length > 1) bad.add('multi-line equations')
    if (e.namespaceURI === M && e.localName === 'f' && children(children(e, M, 'fPr')[0], M, 'type').length) bad.add('stacked/linear fractions')
  }
  if (!bad.size) return null
  const names: Record<string, string> = { m: 'matrices', acc: 'accents', bar: 'bars', func: 'function names', limLow: 'limits', limUpp: 'limits', groupChr: 'group characters', eqArr: 'equation arrays', box: 'boxes', borderBox: 'boxes', sPre: 'pre-scripts', phant: 'phantoms' }
  return `It uses ${[...new Set([...bad].map((b) => names[b] ?? b))].join(', ')}, which can only be edited piece by piece.`
}

export function ommlToTree(xml: string): MathNode {
  const doc = parseXml(xml, 'equation')
  return seq(doc.documentElement)
}

function seq(el: Element | undefined): MathNode {
  if (!el) return { t: 'row', items: [] }
  const items: MathNode[] = []
  for (const c of children(el, M)) {
    const n = node(c)
    if (n) items.push(...(n.t === 'row' ? n.items : [n]))
  }
  return row(mergeTexts(items))
}

const kid = (el: Element, name: string) => children(el, M, name)[0]

function node(c: Element): MathNode | null {
  switch (c.localName) {
    case 'oMath':
    case 'e':
      return seq(c)
    case 'oMathPara':
      return seq(kid(c, 'oMath'))
    case 'r':
      return { t: 'text', s: children(c, M, 't').map((t) => t.textContent ?? '').join('') }
    case 'f':
      return { t: 'frac', num: seq(kid(c, 'num')), den: seq(kid(c, 'den')) }
    case 'sSup':
      return { t: 'script', base: seq(kid(c, 'e')), sup: seq(kid(c, 'sup')) }
    case 'sSub':
      return { t: 'script', base: seq(kid(c, 'e')), sub: seq(kid(c, 'sub')) }
    case 'sSubSup':
      return { t: 'script', base: seq(kid(c, 'e')), sub: seq(kid(c, 'sub')), sup: seq(kid(c, 'sup')) }
    case 'rad':
      return { t: 'rad', body: seq(kid(c, 'e')) }
    case 'nary': {
      const pr = kid(c, 'naryPr')
      const op = pr ? kid(pr, 'chr')?.getAttributeNS(M, 'val') || '∫' : '∫'
      const hide = (n: string) => !!pr && ['1', 'on', 'true'].includes(kid(pr, n)?.getAttributeNS(M, 'val') ?? (kid(pr, n) ? '1' : ''))
      const sub = seq(kid(c, 'sub'))
      const sup = seq(kid(c, 'sup'))
      const empty = (n: MathNode) => n.t === 'row' && !n.items.length
      return { t: 'nary', op, sub: hide('subHide') || empty(sub) ? undefined : sub, sup: hide('supHide') || empty(sup) ? undefined : sup, body: seq(kid(c, 'e')) }
    }
    case 'd':
      return { t: 'delim', body: seq(kid(c, 'e')) }
    default:
      return null
  }
}

const xmlEsc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export function treeToOmml(n: MathNode, display = false): string {
  const inner = omml(n)
  const math = `<m:oMath xmlns:m="${M}">${inner}</m:oMath>`
  return display ? `<m:oMathPara xmlns:m="${M}">${math.replace(` xmlns:m="${M}"`, '')}</m:oMathPara>` : math
}

function omml(n: MathNode): string {
  switch (n.t) {
    case 'row':
      return n.items.map(omml).join('')
    case 'text':
      return n.s ? `<m:r><m:t xml:space="preserve">${xmlEsc(n.s)}</m:t></m:r>` : ''
    case 'frac':
      return `<m:f><m:num>${omml(n.num)}</m:num><m:den>${omml(n.den)}</m:den></m:f>`
    case 'script': {
      const e = `<m:e>${omml(n.base)}</m:e>`
      if (n.sub && n.sup) return `<m:sSubSup>${e}<m:sub>${omml(n.sub)}</m:sub><m:sup>${omml(n.sup)}</m:sup></m:sSubSup>`
      if (n.sup) return `<m:sSup>${e}<m:sup>${omml(n.sup)}</m:sup></m:sSup>`
      return `<m:sSub>${e}<m:sub>${omml(n.sub!)}</m:sub></m:sSub>`
    }
    case 'rad':
      return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${omml(n.body)}</m:e></m:rad>`
    case 'nary': {
      const pr = [`<m:chr m:val="${xmlEsc(n.op)}"/>`, n.op === '∫' || n.op === '∬' || n.op === '∮' ? '<m:limLoc m:val="subSup"/>' : '<m:limLoc m:val="undOvr"/>', n.sub ? '' : '<m:subHide m:val="1"/>', n.sup ? '' : '<m:supHide m:val="1"/>'].join('')
      return `<m:nary><m:naryPr>${pr}</m:naryPr><m:sub>${n.sub ? omml(n.sub) : ''}</m:sub><m:sup>${n.sup ? omml(n.sup) : ''}</m:sup><m:e>${omml(n.body)}</m:e></m:nary>`
    }
    case 'delim':
      return `<m:d><m:e>${omml(n.body)}</m:e></m:d>`
  }
}

/** Replace the math content of an existing equation element, keeping its wrapper (inline vs display). */
export function replaceEquation(xml: string, tree: MathNode): string {
  const doc = parseXml(xml, 'equation')
  const root = doc.documentElement
  const display = root.localName === 'oMathPara'
  return treeToOmml(tree, display)
}

// ---------------------------------------------------------------------------
// piece-by-piece editing for equations the linear syntax cannot express

export function equationPieces(xml: string): string[] {
  const doc = parseXml(xml, 'equation')
  return Array.from(doc.getElementsByTagNameNS(M, 't')).map((t) => t.textContent ?? '')
}

export function setEquationPieces(xml: string, pieces: string[]): string {
  const doc = parseXml(xml, 'equation')
  const ts = Array.from(doc.getElementsByTagNameNS(M, 't'))
  if (ts.length !== pieces.length) throw new Error('The equation changed while it was being edited.')
  ts.forEach((t, i) => {
    t.textContent = pieces[i]
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
  })
  return serializeXml(doc.documentElement)
}

// ---------------------------------------------------------------------------
// display (MathML), built from escaped text only

const OPS = /[+\-−=<>≤≥≠≈±∓×÷⋅→←∈∉⊂⊃∪∩∀∃,;:!|[\]′'…⋯]/u

function tokensMathml(s: string): string {
  let out = ''
  for (const m of s.matchAll(/(\p{N}+(?:\.\p{N}+)?)|(\p{L})|(\s+)|(.)/gu)) {
    if (m[1]) out += `<mn>${xmlEsc(m[1])}</mn>`
    else if (m[2]) out += `<mi>${xmlEsc(m[2])}</mi>`
    else if (m[3]) out += `<mspace width="0.25em"/>`
    else out += OPS.test(m[4]) ? `<mo>${xmlEsc(m[4])}</mo>` : `<mi>${xmlEsc(m[4])}</mi>`
  }
  return out
}

function mml(n: MathNode): string {
  switch (n.t) {
    case 'row':
      return `<mrow>${n.items.map(mml).join('')}</mrow>`
    case 'text':
      return `<mrow>${tokensMathml(n.s)}</mrow>`
    case 'frac':
      return `<mfrac>${mml(n.num)}${mml(n.den)}</mfrac>`
    case 'script':
      if (n.sub && n.sup) return `<msubsup>${mml(n.base)}${mml(n.sub)}${mml(n.sup)}</msubsup>`
      if (n.sup) return `<msup>${mml(n.base)}${mml(n.sup)}</msup>`
      return `<msub>${mml(n.base)}${mml(n.sub!)}</msub>`
    case 'rad':
      return `<msqrt>${mml(n.body)}</msqrt>`
    case 'nary': {
      const op = `<mo largeop="true" movablelimits="false">${xmlEsc(n.op)}</mo>`
      const under = n.op === '∫' || n.op === '∬' || n.op === '∮'
      const lim = n.sub && n.sup ? `<${under ? 'msubsup' : 'munderover'}>${op}${mml(n.sub)}${mml(n.sup)}</${under ? 'msubsup' : 'munderover'}>` : n.sub ? `<${under ? 'msub' : 'munder'}>${op}${mml(n.sub)}</${under ? 'msub' : 'munder'}>` : n.sup ? `<${under ? 'msup' : 'mover'}>${op}${mml(n.sup)}</${under ? 'msup' : 'mover'}>` : op
      return `<mrow>${lim}${mml(n.body)}</mrow>`
    }
    case 'delim':
      return `<mrow><mo>(</mo>${mml(n.body)}<mo>)</mo></mrow>`
  }
}

export function treeToMathml(n: MathNode, display = false): string {
  return `<math xmlns="http://www.w3.org/1998/Math/MathML"${display ? ' display="block"' : ''}>${mml(n)}</math>`
}

/** MathML for an OMML equation; unsupported structures fall back to their text. */
export function ommlToMathml(xml: string): string | null {
  try {
    const display = /^<m:oMathPara\b/.test(xml.trim()) || /^<[^>]*oMathPara\b/.test(xml.trim())
    if (linearBlocker(xml)) {
      const text = equationPieces(xml).join('')
      return treeToMathml({ t: 'text', s: text }, display)
    }
    return treeToMathml(ommlToTree(xml), display)
  } catch {
    return null
  }
}
