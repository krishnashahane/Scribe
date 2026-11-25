import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { ReplaceStep } from '@tiptap/pm/transform'
import { Fragment, Slice, type Node as PMNode } from '@tiptap/pm/model'

// ---------------------------------------------------------------------------
// Find & replace

export interface FindState {
  query: string
  caseSensitive: boolean
  wholeWord: boolean
  matches: { from: number; to: number }[]
  current: number
}

export const findKey = new PluginKey<FindState>('scribe-find')

function escapeRe(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function findAll(doc: PMNode, query: string, caseSensitive: boolean, wholeWord: boolean) {
  const out: { from: number; to: number }[] = []
  if (!query) return out
  const re = new RegExp(wholeWord ? `\\b${escapeRe(query)}\\b` : escapeRe(query), caseSensitive ? 'gu' : 'giu')
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return
    // build the text of the block with a position map (atoms count as one char)
    let text = ''
    const map: number[] = []
    node.forEach((child, offset) => {
      if (child.isText) {
        for (let i = 0; i < child.text!.length; i++) {
          text += child.text![i]
          map.push(pos + 1 + offset + i)
        }
      } else {
        text += '￼'
        map.push(pos + 1 + offset)
      }
    })
    for (const m of text.matchAll(re)) {
      if (!m[0].length) continue
      out.push({ from: map[m.index!], to: map[m.index! + m[0].length - 1] + 1 })
    }
    return false
  })
  return out
}

export const FindReplace = Extension.create({
  name: 'findReplace',
  addProseMirrorPlugins() {
    return [
      new Plugin<FindState>({
        key: findKey,
        state: {
          init: () => ({ query: '', caseSensitive: false, wholeWord: false, matches: [], current: 0 }),
          apply(tr, prev) {
            const meta = tr.getMeta(findKey) as Partial<FindState> | undefined
            if (!meta && !tr.docChanged) return prev
            const next = { ...prev, ...(meta ?? {}) }
            next.matches = findAll(tr.doc, next.query, next.caseSensitive, next.wholeWord)
            if (next.current >= next.matches.length) next.current = 0
            return next
          },
        },
        props: {
          decorations(state) {
            const s = findKey.getState(state)!
            if (!s.matches.length) return null
            return DecorationSet.create(
              state.doc,
              s.matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === s.current ? 'sc-find sc-find-current' : 'sc-find' })),
            )
          },
        },
      }),
    ]
  },
})

export function setFind(view: { state: EditorState; dispatch: (tr: Transaction) => void }, patch: Partial<FindState>) {
  view.dispatch(view.state.tr.setMeta(findKey, patch).setMeta('addToHistory', false))
}

export function gotoMatch(editor: any, delta: number) {
  const s = findKey.getState(editor.state) as FindState
  if (!s.matches.length) return
  const current = (s.current + delta + s.matches.length) % s.matches.length
  const m = s.matches[current]
  editor.view.dispatch(editor.state.tr.setMeta(findKey, { current }).setSelection(TextSelection.create(editor.state.doc, m.from, m.to)).scrollIntoView())
}

export function replaceMatch(editor: any, replacement: string, all: boolean) {
  const s = findKey.getState(editor.state) as FindState
  if (!s.matches.length) return 0
  const targets = all ? [...s.matches].reverse() : [s.matches[s.current]]
  const tr = editor.state.tr
  for (const m of targets) {
    const marks = editor.state.doc.resolve(m.from + 1).marks()
    if (replacement) tr.replaceWith(m.from, m.to, editor.state.schema.text(replacement, marks))
    else tr.delete(m.from, m.to)
  }
  editor.view.dispatch(tr)
  return targets.length
}

// ---------------------------------------------------------------------------
// Suggesting mode (tracked changes)
//
// While enabled, typed text is marked as an insertion and deleted text is kept
// and marked as a deletion instead of being removed. Deleting your own pending
// insertion removes it for real. Supported for edits inside a single text block;
// structural edits across blocks (joining paragraphs, deleting tables) are
// applied directly and are not tracked.

export const suggestKey = new PluginKey<{ enabled: boolean; author: string }>('scribe-suggest')

export const Suggesting = Extension.create<{ author: () => string }>({
  name: 'suggesting',
  addOptions() {
    return { author: () => 'SCRIBE user' }
  },
  addProseMirrorPlugins() {
    const opts = this.options
    return [
      new Plugin({
        key: suggestKey,
        state: {
          init: () => ({ enabled: false, author: opts.author() }),
          apply(tr, prev) {
            const m = tr.getMeta(suggestKey)
            return m ? { ...prev, ...m } : prev
          },
        },
        appendTransaction(trs, oldState, newState) {
          const st = suggestKey.getState(newState)
          if (!st?.enabled) return null
          const schema = newState.schema
          const ins = schema.marks.insertion
          const del = schema.marks.deletion
          const userTrs = trs.filter((t) => t.docChanged && !t.getMeta('y-sync$') && !t.getMeta(suggestKey) && !t.getMeta('scribe-suggest-internal'))
          if (!userTrs.length) return null
          const attrs = { author: st.author || opts.author(), date: new Date().toISOString(), revId: null }
          const tr = newState.tr
          let changed = false
          let doc = oldState.doc
          for (const t of userTrs) {
            t.steps.forEach((step, i) => {
              const stepDoc = t.docs[i] ?? doc
              if (!(step instanceof ReplaceStep)) return
              const { from, to } = step as any
              const slice: Slice = (step as any).slice
              // map positions from this step into the final document
              const mapAfter = t.mapping.slice(i + 1)
              const finalMap = (p: number) => tr.mapping.map(mapAfter.map(p))
              // insertion: mark the inserted inline content
              if (slice.size > 0) {
                const insFrom = finalMap(from)
                const insTo = finalMap(from + slice.size)
                if (insTo > insFrom) {
                  tr.addMark(insFrom, insTo, ins.create(attrs))
                  tr.removeMark(insFrom, insTo, del)
                  changed = true
                }
              }
              // deletion: re-insert removed inline content marked as deleted
              if (to > from) {
                const $from = stepDoc.resolve(from)
                const $to = stepDoc.resolve(to)
                if ($from.sameParent($to) && $from.parent.isTextblock) {
                  const removed = stepDoc.slice(from, to).content
                  const kept: PMNode[] = []
                  removed.forEach((n) => {
                    if (n.isText && n.marks.some((m) => m.type === ins && m.attrs.author === attrs.author)) return // own suggestion: really delete
                    kept.push(n.mark(del.create(attrs).addToSet(n.marks.filter((m) => m.type !== ins))))
                  })
                  if (kept.length) {
                    const at = finalMap(from)
                    tr.insert(at, Fragment.from(kept))
                    changed = true
                    // keep the cursor before the re-inserted text for Backspace, after it for Delete
                    const sel = newState.selection
                    if (sel.empty) {
                      const size = kept.reduce((s, n) => s + n.nodeSize, 0)
                      const cursor = tr.mapping.map(sel.from, -1) === at + size && slice.size === 0 ? at : tr.mapping.map(sel.from)
                      tr.setSelection(TextSelection.create(tr.doc, Math.min(cursor, tr.doc.content.size)))
                    }
                  }
                }
              }
            })
            doc = t.doc
          }
          if (!changed) return null
          tr.setMeta('scribe-suggest-internal', true)
          return tr
        },
      }),
    ]
  },
})

export function setSuggesting(editor: any, enabled: boolean, author: string) {
  editor.view.dispatch(editor.state.tr.setMeta(suggestKey, { enabled, author }))
}

export function isSuggesting(state: EditorState) {
  return !!suggestKey.getState(state)?.enabled
}

export interface ChangeItem {
  type: 'insertion' | 'deletion'
  author: string
  date: string
  text: string
  from: number
  to: number
}

export function listChanges(doc: PMNode): ChangeItem[] {
  const out: ChangeItem[] = []
  doc.descendants((node, pos) => {
    if (!node.isText) return
    for (const m of node.marks) {
      if (m.type.name !== 'insertion' && m.type.name !== 'deletion') continue
      const last = out[out.length - 1]
      if (last && last.type === m.type.name && last.to === pos && last.author === m.attrs.author) {
        last.to = pos + node.nodeSize
        last.text += node.text
      } else out.push({ type: m.type.name as any, author: m.attrs.author, date: m.attrs.date, text: node.text!, from: pos, to: pos + node.nodeSize })
    }
  })
  return out
}

export function resolveChange(editor: any, c: ChangeItem, accept: boolean) {
  const { state } = editor
  const tr = state.tr.setMeta('scribe-suggest-internal', true)
  const remove = (accept && c.type === 'deletion') || (!accept && c.type === 'insertion')
  if (remove) tr.delete(c.from, c.to)
  else tr.removeMark(c.from, c.to, state.schema.marks[c.type])
  editor.view.dispatch(tr)
}

// ---------------------------------------------------------------------------
// Section bookkeeping: section breaks are numbered in document order and the
// caller is told which sections disappeared (break deleted) so it can merge them.

export const SectionSync = Extension.create<{ onSections: (count: number, removed: number[]) => void }>({
  name: 'sectionSync',
  addOptions() {
    return { onSections: () => {} }
  },
  addProseMirrorPlugins() {
    const opts = this.options
    return [
      new Plugin({
        appendTransaction(trs, _old, state) {
          if (!trs.some((t) => t.docChanged)) return null
          const breaks: { pos: number; section: number }[] = []
          state.doc.descendants((n, pos) => {
            if (n.type.name === 'sectionBreak') breaks.push({ pos, section: n.attrs.section })
            return n.isBlock && n.type.name !== 'table'
          })
          const tr = state.tr
          let changed = false
          const present = new Set(breaks.map((b) => b.section))
          const maxIdx = breaks.reduce((m, b) => Math.max(m, b.section), -1)
          const removed: number[] = []
          for (let i = 0; i <= maxIdx; i++) if (!present.has(i)) removed.push(i)
          breaks.forEach((b, i) => {
            if (b.section !== i) {
              tr.setNodeAttribute(b.pos, 'section', i)
              changed = true
            }
          })
          if (removed.length || changed) queueMicrotask(() => opts.onSections(breaks.length + 1, removed))
          if (!changed) return null
          tr.setMeta('addToHistory', false)
          return tr
        },
      }),
    ]
  },
})
