import { defineConfig } from '@playwright/test'

export default defineConfig({
  testDir: './tests',
  timeout: 120_000,
  retries: 1,
  workers: 1, // Sequential: PowerPoint Web supports one add-in WS connection at a time
  globalSetup: './global-setup.ts',
  use: {
    trace: 'retain-on-failure',
    video: 'retain-on-failure',
    ignoreHTTPSErrors: true, // mkcert self-signed certs
  },
  reporter: [['html', { open: 'never' }], ['list']],
})
