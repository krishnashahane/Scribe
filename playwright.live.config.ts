import { defineConfig } from '@playwright/test'

// Runs the end-to-end suite against a deployed URL: LIVE_URL=https://… npx playwright test -c playwright.live.config.ts
export default defineConfig({
  testDir: 'tests/e2e',
  timeout: 120_000,
  workers: 1,
  retries: 0,
  reporter: 'list',
  use: { baseURL: process.env.LIVE_URL, acceptDownloads: true, viewport: { width: 1400, height: 950 } },
})
