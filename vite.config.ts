import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: { '/api': 'http://localhost:8787' },
  },
  // no source maps in production builds (application code has no console calls: enforced by ESLint no-console)
  build: {
    sourcemap: false,
    chunkSizeWarningLimit: 700,
    rollupOptions: {
      output: {
        // libraries change rarely: separate files stay cached in the browser across SCRIBE updates
        manualChunks(id: string) {
          if (!id.includes('node_modules')) return undefined
          if (/node_modules\/(react|react-dom|scheduler)\//.test(id)) return 'vendor-react'
          if (/node_modules\/(yjs|lib0|y-protocols|y-indexeddb|@tiptap\/y-tiptap)\//.test(id)) return 'vendor-collab'
          if (/node_modules\/(@tiptap|prosemirror-|linkifyjs|orderedmap|rope-sequence|w3c-keyname|fast-equals)/.test(id)) return 'vendor-editor'
          if (/node_modules\/jszip\//.test(id)) return 'vendor-zip'
          return undefined
        },
      },
    },
  },
  test: {
    environment: 'jsdom',
    include: ['tests/unit/**/*.test.ts'],
    testTimeout: 30000,
  },
})
