import { defineConfig } from '@playwright/test'

export default defineConfig({
  testDir: 'tests/e2e',
  timeout: 90_000,
  workers: 1,
  reporter: 'list',
  use: { baseURL: 'http://127.0.0.1:8799', acceptDownloads: true, viewport: { width: 1400, height: 950 } },
  webServer: {
    command: 'rm -rf tests/output/e2e-db && NODE_ENV=production PORT=8799 SCRIBE_DB_DIR=tests/output/e2e-db node server/index.js',
    url: 'http://127.0.0.1:8799/api/health',
    reuseExistingServer: false,
    timeout: 30_000,
  },
})
