import { defineConfig } from '@playwright/test'

export default defineConfig({
  testDir: './tests',
  // Cold WAC can take 20-40s just to surface the sideload dialogs, and a
  // freshly-registered add-in needs a ribbon click + WS connect on top.
  timeout: 180_000,
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
