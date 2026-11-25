// Local server: the same API as the Vercel Function, plus the built client.
import http from 'node:http'
import { createApp } from './app.js'

const port = Number(process.env.PORT || 8787)
const host = process.env.HOST || '127.0.0.1'
const { app, db } = await createApp({ dbDir: process.env.SCRIBE_DB_DIR || 'data/pglite', serveStatic: process.env.NODE_ENV === 'production', secureCookies: false })
const server = http.createServer(app)
server.listen(port, host, () => process.stdout.write(`SCRIBE listening on http://${host}:${port}\n`))

const shutdown = () => {
  server.close(() => db.close().finally(() => process.exit(0)))
  setTimeout(() => process.exit(0), 3000).unref()
}
process.on('SIGINT', shutdown)
process.on('SIGTERM', shutdown)
