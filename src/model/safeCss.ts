// Validators for values that end up inside inline `style` attributes. Document content comes from
// imported files and from collaborators, so a value like `red;background:url(https://…)` must never
// reach the page as CSS (it could load remote resources or restyle the interface).

const NAMED_COLOR = /^[a-z]{3,24}$/i
const HEX = /^#[0-9a-f]{3,8}$/i
const FUNC = /^(?:rgb|rgba|hsl|hsla)\(\s*[\d.%]+\s*(?:,\s*|\s+)[\d.%]+\s*(?:,\s*|\s+)[\d.%]+\s*(?:(?:,|\/)\s*[\d.%]+\s*)?\)$/i

export function safeColor(v: unknown): string | null {
  if (typeof v !== 'string') return null
  const s = v.trim()
  if (s.length > 40) return null
  return HEX.test(s) || FUNC.test(s) || NAMED_COLOR.test(s) ? s : null
}

/** Font family lists: names, quotes, commas, spaces and hyphens only. */
export function safeFontFamily(v: unknown): string | null {
  if (typeof v !== 'string') return null
  const s = v.trim()
  return s.length > 0 && s.length <= 200 && /^[\p{L}\p{N}\s,'"_.-]+$/u.test(s) ? s : null
}

/** A CSS length such as 11pt, 14px, 1.2em. */
export function safeLength(v: unknown, units = ['pt', 'px', 'em', 'rem', '%']): string | null {
  if (typeof v !== 'string') return null
  const m = /^(-?\d{1,4}(?:\.\d{1,3})?)([a-z%]{1,3})$/i.exec(v.trim())
  return m && units.includes(m[2].toLowerCase()) ? `${m[1]}${m[2].toLowerCase()}` : null
}

export function safeKeyword(v: unknown, allowed: readonly string[]): string | null {
  return typeof v === 'string' && allowed.includes(v.trim().toLowerCase()) ? v.trim().toLowerCase() : null
}

export const LIST_STYLES = ['disc', 'circle', 'square', 'decimal', 'decimal-leading-zero', 'lower-alpha', 'upper-alpha', 'lower-latin', 'upper-latin', 'lower-roman', 'upper-roman', 'lower-greek', 'none'] as const
export const VERTICAL_ALIGN = ['top', 'middle', 'bottom', 'baseline'] as const

/** A finite number within bounds, else null. */
export function safeNumber(v: unknown, min = -1e6, max = 1e6): number | null {
  const n = typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' ? Number(v) : NaN
  return Number.isFinite(n) && n >= min && n <= max ? n : null
}
