export class ApiError extends Error {
  constructor(message: string, public status: number) {
    super(message)
  }
}

export async function api<T = any>(path: string, opts: { method?: string; body?: unknown; raw?: boolean } = {}): Promise<T> {
  let res: Response
  try {
    res = await fetch(path, {
      method: opts.method ?? 'GET',
      credentials: 'same-origin',
      headers: { 'content-type': 'application/json', 'x-scribe-request': '1' },
      body: opts.body === undefined ? undefined : JSON.stringify(opts.body),
    })
  } catch {
    throw new ApiError('Cannot reach the SCRIBE server. Check your connection.', 0)
  }
  if (opts.raw) {
    if (!res.ok) throw new ApiError(`Request failed (${res.status})`, res.status)
    return res as unknown as T
  }
  const data = await res.json().catch(() => ({}))
  if (!res.ok) throw new ApiError(data?.error || `Request failed (${res.status})`, res.status)
  return data as T
}

export interface User {
  id: string
  email: string
  name: string
  /** local no-login identity; can be upgraded to a full account */
  guest?: boolean
}

export interface DocSummary {
  id: string
  title: string
  updated: number
  created: number
  owner_name: string
  role: 'owner' | 'editor' | 'commenter' | 'viewer'
  has_original: boolean
}

export function toBase64(bytes: Uint8Array): string {
  let s = ''
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return btoa(s)
}
