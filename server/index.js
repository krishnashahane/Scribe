import { createApp } from './app.js'

const port = Number(process.env.PORT || 8787)
const host = process.env.HOST || '127.0.0.1'
const { server, collab, db } = createApp({ dbFile: process.env.SCRIBE_DB || undefined })

server.listen(port, host, () => console.log(`SCRIBE server listening on http://${host}:${port}`))

const shutdown = () => {
  collab.flushAll()
  db.close()
  server.close(() => process.exit(0))
  setTimeout(() => process.exit(0), 2000).unref()
}
process.on('SIGINT', shutdown)
process.on('SIGTERM', shutdown)
