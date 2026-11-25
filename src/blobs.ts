import { api, ApiError } from './api'

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

/** Upload bytes in chunks that fit serverless request limits. Returns the blob id and (for images) its URL. */
export async function uploadBlob(bytes: Uint8Array, kind: 'media' | 'original', mime: string): Promise<{ id: string; url: string | null }> {
  const meta = await api<{ id: string; chunks: number; chunkSize: number }>('/api/blobs', { method: 'POST', body: { kind, mime, size: bytes.length } })
  for (let i = 0; i < meta.chunks; i++) {
    const part = bytes.subarray(i * meta.chunkSize, Math.min(bytes.length, (i + 1) * meta.chunkSize))
    const res = await fetch(`/api/blobs/${meta.id}/chunks/${i}`, {
      method: 'PUT',
      credentials: 'same-origin',
      headers: { 'content-type': 'application/octet-stream', 'x-scribe-request': '1' },
      body: part as BodyInit,
    })
    if (!res.ok) throw new ApiError((await res.json().catch(() => ({})))?.error || 'Upload failed', res.status)
  }
  return api('/api/blobs/' + meta.id + '/complete', { method: 'POST', body: {} })
}

export const uploadOriginal = (bytes: Uint8Array) => uploadBlob(bytes, 'original', DOCX_MIME)

/** Download the stored original DOCX of a document (fetched in parts). */
export async function downloadOriginal(docId: string, parts: number): Promise<Uint8Array> {
  const chunks: Uint8Array[] = []
  for (let i = 0; i < parts; i++) {
    const res = await api<Response>(`/api/docs/${docId}/original/${i}`, { raw: true })
    chunks.push(new Uint8Array(await res.arrayBuffer()))
  }
  const out = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0))
  let o = 0
  for (const c of chunks) {
    out.set(c, o)
    o += c.length
  }
  return out
}

const IMAGE_MIMES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/bmp']
export const MAX_IMAGE_BYTES = 4_000_000

export function dataUrlToBytes(url: string): { bytes: Uint8Array; mime: string } | null {
  const m = /^data:([^;,]+);base64,(.*)$/s.exec(url)
  if (!m || !IMAGE_MIMES.includes(m[1].toLowerCase())) return null
  return { mime: m[1].toLowerCase(), bytes: Uint8Array.from(atob(m[2]), (c) => c.charCodeAt(0)) }
}

/** Store an image file on the server and return its URL. */
export async function uploadImage(bytes: Uint8Array, mime: string) {
  if (!IMAGE_MIMES.includes(mime)) throw new Error('Unsupported image type. Use PNG, JPEG, GIF, WebP or BMP.')
  if (bytes.length > MAX_IMAGE_BYTES) throw new Error('Images must be smaller than 4 MB.')
  return (await uploadBlob(bytes, 'media', mime)).url!
}
