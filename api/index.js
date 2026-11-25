// Vercel Function entry: every /api/* request is rewritten here (see vercel.json).
import { createApp } from '../server/app.js'

let ready
export default async function handler(req, res) {
  ready ??= createApp().catch((e) => {
    ready = undefined
    throw e
  })
  try {
    const { app } = await ready
    return app(req, res)
  } catch {
    res.statusCode = 503
    res.setHeader('content-type', 'application/json')
    res.end(JSON.stringify({ error: 'Service temporarily unavailable' }))
  }
}
