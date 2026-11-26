// Server-side input validation (zod) and sanitisation of stored document content.
import { z } from 'zod'

export const MAX_JSON_CHARS = 3_000_000 // fits Vercel's 4.5 MB request limit with room for base64 payloads
export const MAX_UPDATE_BYTES = 2_000_000
export const CHUNK_BYTES = 3_000_000
export const MAX_MEDIA_BYTES = 4_000_000 // a media response must fit in one serverless response
export const MAX_ORIGINAL_BYTES = 100 * 1024 * 1024

const id = z.string().regex(/^[A-Za-z0-9_-]{6,64}$/, 'Invalid id')
export const params = {
  doc: z.object({ id }),
  docVersion: z.object({ id, vid: id }),
  docShare: z.object({ id, userId: id }),
  blob: z.object({ id }),
  blobChunk: z.object({ id, idx: z.coerce.number().int().min(0).max(10_000) }),
  part: z.object({ id, part: z.coerce.number().int().min(0).max(10_000) }),
}

const email = z.string().trim().toLowerCase().max(200).pipe(z.email('Enter a valid email address.'))
const password = z.string().min(8, 'Password must be 8–200 characters.').max(200, 'Password must be 8–200 characters.')
const name = z.string().trim().max(100).optional().default('')
const title = z.string().trim().max(200)

export const schemas = {
  register: z.object({ email, password, name }),
  login: z.object({ email: z.string().trim().toLowerCase().max(200), password: z.string().max(200) }),
  upgrade: z.object({ email, password, name }),
  createDoc: z.object({
    title: title.optional().default(''),
    content: z.object({ type: z.literal('doc') }).loose(),
    meta: z.record(z.string(), z.unknown()).optional().default({}),
    yUpdate: z.string().max(Math.ceil((MAX_UPDATE_BYTES * 4) / 3) + 4).regex(/^[A-Za-z0-9+/=]*$/).optional(),
    originalBlob: id.optional(),
    originalName: z.string().max(255).optional(),
  }),
  patchDoc: z
    .object({
      title: title.optional(),
      content: z.object({ type: z.literal('doc') }).loose().optional(),
      meta: z.record(z.string(), z.unknown()).optional(),
    })
    .strict(),
  share: z.object({ email: z.string().trim().toLowerCase().max(200), role: z.enum(['viewer', 'editor']) }),
  version: z.object({
    content: z.object({ type: z.literal('doc') }).loose(),
    meta: z.record(z.string(), z.unknown()).optional().default({}),
    label: z.string().trim().max(120).optional(),
    auto: z.boolean().optional().default(false),
  }),
  sync: z.object({
    after: z.number().int().min(0),
    updates: z.array(z.string().max(Math.ceil((MAX_UPDATE_BYTES * 4) / 3) + 4).regex(/^[A-Za-z0-9+/=]*$/)).max(50).optional().default([]),
    awareness: z
      .object({ clientId: z.number().int().min(0).max(2 ** 53 - 1), state: z.string().max(20_000).regex(/^[A-Za-z0-9+/=]*$/) })
      .optional(),
  }),
  blob: z.object({
    kind: z.enum(['media', 'original']),
    mime: z.enum(['image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/bmp', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document']),
    size: z.number().int().min(1).max(MAX_ORIGINAL_BYTES),
  }),
  ai: z.object({
    task: z.enum(['improve', 'grammar', 'summarize', 'shorter', 'longer', 'formal', 'friendly', 'confident', 'continue', 'generate', 'custom']),
    text: z.string().max(60_000).optional().default(''),
    instruction: z.string().max(2000).optional(),
    context: z.string().max(20_000).optional(),
  }),
}

/** Parse or respond 400 with a safe, user-facing message. Returns null when invalid. */
export function check(schema, data, res) {
  const r = schema.safeParse(data ?? {})
  if (r.success) return r.data
  const msg = r.error.issues[0]?.message
  const friendly = msg && !/^(Invalid input|Expected|Too (big|small)|Required|Unrecognized)/i.test(msg) ? msg : 'Invalid request'
  res.status(400).json({ error: friendly })
  return null
}

// ---------------------------------------------------------------------------
// Document content sanitisation (defence in depth: the editor schema also validates)

const SAFE_LINK = /^(https?:|mailto:|tel:|ftp:|#)/i
const SAFE_IMG = /^(data:image\/(png|jpe?g|gif|webp|bmp);base64,[A-Za-z0-9+/=]+$|\/api\/media\/[A-Za-z0-9_-]{6,64}$|https:\/\/)/i
const NODE_TYPES = new Set([
  'doc', 'paragraph', 'heading', 'text', 'hardBreak', 'bulletList', 'orderedList', 'listItem', 'blockquote', 'codeBlock',
  'horizontalRule', 'table', 'tableRow', 'tableCell', 'tableHeader', 'image', 'pageBreak', 'sectionBreak', 'columnBreak',
  'bookmark', 'field', 'noteRef', 'opaqueBlock', 'opaqueInline', 'shape', 'bibliography', 'toc',
])

// values that are rendered into inline style attributes (mirrors src/model/safeCss.ts)
const COLOR = /^(?:#[0-9a-f]{3,8}|[a-z]{3,24}|(?:rgb|rgba|hsl|hsla)\(\s*[\d.%]+\s*(?:,\s*|\s+)[\d.%]+\s*(?:,\s*|\s+)[\d.%]+\s*(?:(?:,|\/)\s*[\d.%]+\s*)?\))$/i
const FONT = /^[\p{L}\p{N}\s,'"_.-]{1,200}$/u
const LENGTH = /^-?\d{1,4}(?:\.\d{1,3})?(?:pt|px|em|rem|%)$/i
const KEYWORDS = {
  textAlign: ['left', 'center', 'right', 'justify'],
  verticalAlign: ['top', 'middle', 'bottom', 'baseline'],
  listStyle: ['disc', 'circle', 'square', 'decimal', 'decimal-leading-zero', 'lower-alpha', 'upper-alpha', 'lower-latin', 'upper-latin', 'lower-roman', 'upper-roman', 'lower-greek', 'none'],
}
const STYLE_ATTRS = {
  color: (v) => COLOR.test(v),
  backgroundColor: (v) => COLOR.test(v),
  background: (v) => COLOR.test(v),
  shading: (v) => COLOR.test(v),
  fontFamily: (v) => FONT.test(v),
  fontSize: (v) => LENGTH.test(v),
  letterSpacing: (v) => LENGTH.test(v),
  textAlign: (v) => KEYWORDS.textAlign.includes(v),
  verticalAlign: (v) => KEYWORDS.verticalAlign.includes(v),
  listStyle: (v) => KEYWORDS.listStyle.includes(v),
}

/** Drop style-bearing attribute values that are not plain colours, fonts, lengths or keywords. */
function cleanStyleAttrs(attrs) {
  const out = { ...attrs }
  for (const [k, ok] of Object.entries(STYLE_ATTRS)) {
    if (out[k] == null) continue
    if (typeof out[k] !== 'string' || out[k].length > 200 || !ok(out[k].trim())) out[k] = null
  }
  return out
}

/** Remove unsafe URLs and unknown node types from ProseMirror JSON. Returns a new object. */
export function sanitizeContent(node, depth = 0) {
  if (!node || typeof node !== 'object' || depth > 200) return null
  if (typeof node.type !== 'string' || !NODE_TYPES.has(node.type)) return null
  const out = { type: node.type }
  if (node.attrs && typeof node.attrs === 'object') {
    out.attrs = cleanStyleAttrs(node.attrs)
    if (node.type === 'image' && !SAFE_IMG.test(String(out.attrs.src ?? ''))) return null
  }
  if (typeof node.text === 'string') out.text = node.text
  if (Array.isArray(node.marks)) {
    out.marks = node.marks
      .filter((m) => m && typeof m.type === 'string')
      .filter((m) => m.type !== 'link' || SAFE_LINK.test(String(m.attrs?.href ?? '')))
      .map((m) => (m.attrs && typeof m.attrs === 'object' ? { ...m, attrs: cleanStyleAttrs(m.attrs) } : m))
  }
  if (Array.isArray(node.content)) out.content = node.content.map((c) => sanitizeContent(c, depth + 1)).filter(Boolean)
  return out
}

export function sanitizeMeta(meta) {
  const m = { ...meta }
  if (m.headerFooters && typeof m.headerFooters === 'object') {
    const hf = {}
    for (const [k, v] of Object.entries(m.headerFooters)) {
      const clean = sanitizeContent(v)
      if (clean) hf[k.slice(0, 200)] = clean
    }
    m.headerFooters = hf
  }
  return m
}

export const jsonSize = (v) => JSON.stringify(v).length
