// Grammar and style checker that runs entirely in the browser (no AI, no network).
// Rules are deliberately conservative: each one targets a pattern that is almost always a mistake.

import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { Node as PMNodeType } from '@tiptap/pm/model'

export type IssueKind = 'grammar' | 'punctuation' | 'style'

export interface TextIssue {
  rule: string
  kind: IssueKind
  /** offsets into the checked string */
  start: number
  end: number
  message: string
  /** replacement text; empty string deletes; undefined means no automatic fix */
  fix?: string
}

export const RULE_NAMES: Record<string, string> = {
  repeated: 'Repeated word',
  spaces: 'Extra spaces',
  spaceBefore: 'Space before punctuation',
  spaceAfter: 'Missing space after punctuation',
  capital: 'Sentence starts with a lowercase letter',
  pronounI: 'Lowercase “i”',
  article: '“a” or “an”',
  couldOf: '“could of”',
  confused: 'Commonly confused words',
  doublePunct: 'Doubled punctuation',
  longSentence: 'Very long sentence',
  wordy: 'Wordy phrase',
}

const ALLOWED_REPEATS = new Set(['had', 'that', 'is', 'do', 'bye', 'no', 'yeah', 'ha', 'very', 'so', 'knock', 'tut', 'well', 'blah', 'la', 'go', 'boo', 'chop', 'hush', 'ding', 'dong', 'murmur', 'pom', 'tsk'])
const ABBREVIATIONS = /\b(?:e\.g|i\.e|etc|vs|cf|approx|incl|Mr|Mrs|Ms|Dr|Prof|Sr|Jr|St|No|Fig|Vol|pp|p|ca|al|Inc|Ltd|Co|Corp|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|a\.m|p\.m)\.$/i

// a/an exceptions by sound, not spelling
const AN_BEFORE = /^(?:hour|honest|honou?r|heir|herb(?!al)|FBI|FAQ|HTML|HTTP|MBA|MP3|MRI|NGO|SQL|SMS|LED|X-?ray|RSVP|NBA|NFL|MSc|LLC|8|11|18|80)/i
const A_BEFORE = /^(?:uni(?!n)|use|usu|uti|ur(?:e|o)|eu|one(?![a-z])|once|ewe|u-|uk\b|us\b|ubiq|uran|uter|uvu|unanim|unicorn|unique|unit|univ|user|usual|utop)/i

const CONFUSED: [RegExp, string, string][] = [
  [/\b(could|should|would|must|might)\s+of\b/gi, '$1 have', 'Use “have” after modal verbs.'],
  [/\balot\b/gi, 'a lot', '“A lot” is two words.'],
  [/\birregardless\b/gi, 'regardless', 'Use “regardless”.'],
  [/\b(better|more|less|rather|other|greater|smaller|larger|faster|slower|higher|lower|older|younger)\s+then\b/gi, '$1 than', 'Use “than” for comparisons.'],
  [/\byour\s+welcome\b/gi, "you're welcome", 'Use “you’re” (you are).'],
  [/\bfor\s+all\s+intensive\s+purposes\b/gi, 'for all intents and purposes', 'The phrase is “for all intents and purposes”.'],
  [/\bcan\s+not\b/gi, 'cannot', '“Cannot” is usually one word.'],
  [/\bsuppose\s+to\b/gi, 'supposed to', 'Use “supposed to”.'],
  [/\buse\s+to\b(?=\s+(?:be|have|go|do|live|work|think|say|make|play|get|come))/gi, 'used to', 'Use “used to” for past habits.'],
]

const WORDY: [RegExp, string][] = [
  [/\bin order to\b/gi, 'to'],
  [/\bdue to the fact that\b/gi, 'because'],
  [/\bat this point in time\b/gi, 'now'],
  [/\bin the event that\b/gi, 'if'],
  [/\bfor the purpose of\b/gi, 'to'],
  [/\bin spite of the fact that\b/gi, 'although'],
  [/\bhas the ability to\b/gi, 'can'],
  [/\bit is important to note that\s*/gi, ''],
]

const keepCase = (orig: string, rep: string) => (orig[0] && orig[0] === orig[0].toUpperCase() && orig[0] !== orig[0].toLowerCase() ? rep[0].toUpperCase() + rep.slice(1) : rep)

/** Check one paragraph of text. Offsets refer to the given string. */
export function checkText(text: string, disabled: Set<string> = new Set()): TextIssue[] {
  const out: TextIssue[] = []
  const add = (i: TextIssue) => !disabled.has(i.rule) && out.push(i)
  let m: RegExpExecArray | null

  // repeated words: "the the"
  const rep = /\b([\p{L}']+)(\s+)\1\b/giu
  while ((m = rep.exec(text))) {
    const w = m[1].toLowerCase()
    if (ALLOWED_REPEATS.has(w) || /^\p{N}/u.test(w)) continue
    add({ rule: 'repeated', kind: 'grammar', start: m.index, end: m.index + m[0].length, message: `“${m[1]}” is repeated.`, fix: m[1] })
  }
  // multiple spaces between words (not leading indentation, not before tabs)
  const sp = /(?<=\S) {2,}(?=\S)/g
  while ((m = sp.exec(text))) add({ rule: 'spaces', kind: 'punctuation', start: m.index, end: m.index + m[0].length, message: 'There is more than one space here.', fix: ' ' })
  // space before punctuation
  const sb = /(?<=[\p{L}\p{N})”"’]) +([,;:!?](?![\p{L}\p{N}])|\.(?![.\p{N}]))/gu
  while ((m = sb.exec(text))) {
    if (m[1] === ':' || m[1] === ';' || m[1] === '!' || m[1] === '?') {
      // French typography uses a space before : ; ! ? — skip when the text looks French
      if (/\b(?:le|la|les|des|est|une|pour|avec)\b/i.test(text)) continue
    }
    add({ rule: 'spaceBefore', kind: 'punctuation', start: m.index, end: m.index + m[0].length, message: `Remove the space before “${m[1]}”.`, fix: m[1] })
  }
  // missing space after comma/semicolon before a letter (1,000 and URLs are fine)
  const sa = /([,;])(?=\p{L})/gu
  while ((m = sa.exec(text))) {
    const before = text.slice(Math.max(0, m.index - 40), m.index)
    if (/\S*(?:https?:|www\.|@)\S*$/.test(before)) continue
    add({ rule: 'spaceAfter', kind: 'punctuation', start: m.index, end: m.index + 1, message: `Add a space after “${m[1]}”.`, fix: m[1] + ' ' })
  }
  // doubled punctuation ",," "!!" is style; ".." (not "...") and ",," are errors
  const dp = /(?<![.])(\.\.)(?!\.)|,,|;;|::/g
  while ((m = dp.exec(text))) add({ rule: 'doublePunct', kind: 'punctuation', start: m.index, end: m.index + m[0].length, message: 'This punctuation mark is doubled.', fix: m[0][0] })
  // sentence capitalisation
  const sentence = /([.!?])(["”’)]*)(\s+)(\p{Ll})/gu
  while ((m = sentence.exec(text))) {
    const head = text.slice(0, m.index + 1)
    if (m[1] === '.' && ABBREVIATIONS.test(head.split(/\s/).pop() ?? '')) continue
    if (m[1] === '.' && /(?:^|\s)\p{L}\.$/u.test(head)) continue // initials: J. smith is rare, "U.S. lawmakers"
    if (/\.\.$/.test(head) || /…$/.test(head)) continue
    const at = m.index + m[1].length + m[2].length + m[3].length
    const word = /^\p{L}+/u.exec(text.slice(at))?.[0] ?? ''
    if (/^(?:e|i|iphone|ipad|ios|macos|npm|eBay)$/i.test(word) && word !== word.toLowerCase()) continue
    if (/^(?:iphone|ipad|ios|macos|npm|ebay|eg|ie)$/i.test(word)) continue
    add({ rule: 'capital', kind: 'grammar', start: at, end: at + 1, message: 'Start the sentence with a capital letter.', fix: m[4].toUpperCase() })
  }
  const first = /^(\s*["“(]?)(\p{Ll})(\p{L}*)/u.exec(text)
  if (first && text.length > 25 && /[.!?]["”’)]*\s*$/.test(text) && !/^(?:iphone|ipad|ios|macos|npm|ebay|e|i)$/i.test(first[2] + first[3])) {
    add({ rule: 'capital', kind: 'grammar', start: first[1].length, end: first[1].length + 1, message: 'Start the sentence with a capital letter.', fix: first[2].toUpperCase() })
  }
  // lowercase "i" as a pronoun
  const pi = /(?<![\p{L}\p{N}.'’-])i(?=(?:'m|’m|'ve|’ve|'ll|’ll|'d|’d)?(?![\p{L}\p{N}.'’-]))/gu
  while ((m = pi.exec(text))) {
    const prev = text.slice(Math.max(0, m.index - 2), m.index)
    if (/[([]$/.test(prev) || /\(\s?i\)/.test(text.slice(m.index - 1, m.index + 2))) continue // list markers (i)
    add({ rule: 'pronounI', kind: 'grammar', start: m.index, end: m.index + 1, message: 'The pronoun “I” is always capitalised.', fix: 'I' })
  }
  // a / an
  const art = /\b(a|an|A|An)\s+([\p{L}\p{N}][\p{L}\p{N}-]*)/gu
  while ((m = art.exec(text))) {
    const word = m[2]
    if (word.length === 1 && /[\p{L}]/u.test(word)) continue // "a b c", variables
    // acronyms may be spelled out (an FBI agent) or read as words (a NASA probe): only the listed ones are checked
    if (/^[A-Z0-9-]{2,}$/.test(word)) {
      if (AN_BEFORE.test(word) && m[1].toLowerCase() === 'a') add({ rule: 'article', kind: 'grammar', start: m.index, end: m.index + m[1].length, message: `Use “an” before “${word}”.`, fix: keepCase(m[1], 'an') })
      continue
    }
    const vowelSound = AN_BEFORE.test(word) || (/^[aeiou]/i.test(word) && !A_BEFORE.test(word))
    const isAn = m[1].toLowerCase() === 'an'
    if (/^[\p{N}]/u.test(word) && !AN_BEFORE.test(word)) {
      if (isAn) add({ rule: 'article', kind: 'grammar', start: m.index, end: m.index + m[1].length, message: `Use “a” before “${word}”.`, fix: keepCase(m[1], 'a') })
      continue
    }
    if (vowelSound && !isAn) add({ rule: 'article', kind: 'grammar', start: m.index, end: m.index + m[1].length, message: `Use “an” before “${word}”.`, fix: keepCase(m[1], 'an') })
    else if (!vowelSound && isAn && /^\p{L}/u.test(word) && !/^h/i.test(word)) add({ rule: 'article', kind: 'grammar', start: m.index, end: m.index + m[1].length, message: `Use “a” before “${word}”.`, fix: keepCase(m[1], 'a') })
  }
  // commonly confused words
  for (const [re, rep2, msg] of CONFUSED) {
    re.lastIndex = 0
    while ((m = re.exec(text))) {
      const fixText = m[0].replace(new RegExp(re.source, 'i'), rep2)
      add({ rule: re === CONFUSED[0][0] ? 'couldOf' : 'confused', kind: 'grammar', start: m.index, end: m.index + m[0].length, message: msg, fix: keepCase(m[0], fixText) })
    }
  }
  // wordy phrases (style)
  for (const [re, rep2] of WORDY) {
    re.lastIndex = 0
    while ((m = re.exec(text))) add({ rule: 'wordy', kind: 'style', start: m.index, end: m.index + m[0].length, message: rep2 ? `Consider “${rep2}”.` : 'This phrase can usually be removed.', fix: rep2 ? keepCase(m[0], rep2) : '' })
  }
  // long sentences (style, no automatic fix)
  const sent = /[^.!?]+[.!?]+/g
  while ((m = sent.exec(text))) {
    const words = m[0].trim().split(/\s+/).filter((w) => /[\p{L}\p{N}]/u.test(w)).length
    if (words > 40) {
      const lead = m[0].length - m[0].trimStart().length
      add({ rule: 'longSentence', kind: 'style', start: m.index + lead, end: m.index + m[0].length, message: `This sentence has ${words} words. Consider splitting it.` })
    }
  }
  // drop overlapping issues, keeping the first (more specific rules run first)
  out.sort((a, b) => a.start - b.start || b.end - a.end)
  const kept: TextIssue[] = []
  for (const i of out) {
    const clash = kept.find((k) => i.start < k.end && k.start < i.end && i.rule !== 'longSentence' && k.rule !== 'longSentence')
    if (!clash) kept.push(i)
  }
  return kept
}

// ---------------------------------------------------------------------------
// editor integration

export interface DocIssue extends TextIssue {
  from: number
  to: number
  id: string
}

interface ProofState {
  enabled: boolean
  issues: DocIssue[]
  ignored: Set<string>
  disabled: Set<string>
  decos: DecorationSet
}

export const proofKey = new PluginKey<ProofState>('proofing')

/**
 * Text of a textblock with a position map. Atoms (fields, images) become U+FFFC so they never join words.
 */
function blockText(node: PMNodeType, start: number): { text: string; pos: number[] } {
  let text = ''
  const pos: number[] = []
  node.forEach((child, offset) => {
    const at = start + 1 + offset
    if (child.isText) {
      const deleted = child.marks.some((m) => m.type.name === 'deletion')
      const t = child.text ?? ''
      for (let i = 0; i < t.length; i++) {
        if (deleted) continue
        text += t[i]
        pos.push(at + i)
      }
    } else {
      text += child.type.name === 'hardBreak' ? '\n' : '￼'
      pos.push(at)
    }
  })
  pos.push(start + node.nodeSize - 1)
  return { text, pos }
}

function issueId(i: TextIssue, text: string) {
  return `${i.rule}:${text.slice(Math.max(0, i.start - 12), i.end + 12)}`
}

function collect(doc: PMNodeType, ignored: Set<string>, disabled: Set<string>): DocIssue[] {
  const out: DocIssue[] = []
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    if (node.type.name === 'codeBlock') return false
    const { text, pos: map } = blockText(node, pos)
    if (!text.trim()) return false
    for (const i of checkText(text, disabled)) {
      if (text.slice(i.start, i.end).includes('￼')) continue
      const id = issueId(i, text)
      if (ignored.has(id)) continue
      out.push({ ...i, from: map[i.start], to: map[i.end - 1] + 1, id })
    }
    return false
  })
  return out
}

function decorations(doc: PMNodeType, issues: DocIssue[]) {
  return DecorationSet.create(
    doc,
    issues.map((i) => Decoration.inline(i.from, i.to, { class: `sc-proof is-${i.kind}`, 'data-proof': i.id, title: `${RULE_NAMES[i.rule] ?? 'Suggestion'}: ${i.message}` }, { id: i.id })),
  )
}

function recompute(state: EditorState, s: ProofState): ProofState {
  if (!s.enabled) return { ...s, issues: [], decos: DecorationSet.empty }
  const issues = collect(state.doc, s.ignored, s.disabled)
  return { ...s, issues, decos: decorations(state.doc, issues) }
}

type ProofMeta = { enable?: boolean; ignore?: string; disableRule?: string; refresh?: boolean }

export const Proofing = Extension.create({
  name: 'proofing',
  addProseMirrorPlugins() {
    let timer: ReturnType<typeof setTimeout> | undefined
    return [
      new Plugin<ProofState>({
        key: proofKey,
        state: {
          init: () => ({ enabled: false, issues: [], ignored: new Set(), disabled: new Set(), decos: DecorationSet.empty }),
          apply(tr: Transaction, s: ProofState, _old, state) {
            const meta = tr.getMeta(proofKey) as ProofMeta | undefined
            if (meta) {
              const next = { ...s }
              if (meta.enable !== undefined) next.enabled = meta.enable
              if (meta.ignore) next.ignored = new Set([...s.ignored, meta.ignore])
              if (meta.disableRule) next.disabled = new Set([...s.disabled, meta.disableRule])
              return recompute(state, next)
            }
            if (!s.enabled || !tr.docChanged) return s
            // keep underlines in place while typing; a debounced refresh recomputes
            const issues = s.issues
              .map((i) => ({ ...i, from: tr.mapping.map(i.from, 1), to: tr.mapping.map(i.to, -1) }))
              .filter((i) => i.to > i.from)
            return { ...s, issues, decos: s.decos.map(tr.mapping, tr.doc) }
          },
        },
        props: {
          decorations: (state) => proofKey.getState(state)?.decos,
        },
        view: () => ({
          update(view, prev) {
            const s = proofKey.getState(view.state)
            if (!s?.enabled || view.state.doc.eq(prev.doc)) return
            clearTimeout(timer)
            timer = setTimeout(() => {
              if (!view.isDestroyed) view.dispatch(view.state.tr.setMeta(proofKey, { refresh: true }).setMeta('addToHistory', false))
            }, 600)
          },
          destroy: () => clearTimeout(timer),
        }),
      }),
    ]
  },
})

export function setProofing(view: { state: EditorState; dispatch: (tr: Transaction) => void }, meta: ProofMeta) {
  view.dispatch(view.state.tr.setMeta(proofKey, meta).setMeta('addToHistory', false))
}

export function proofState(state: EditorState) {
  return proofKey.getState(state)
}
