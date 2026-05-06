// @ts-check
import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './tests/e2e',
  outputDir: './artifacts/test-results',

  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: process.env.CI ? 1 : undefined,
  timeout: 1000 * 10000,
  reporter: [['html', { outputFolder: './artifacts/playwright-report', open: 'never' }]],

  use: {

    storageState: './auth.json',
    headless: false,
    ignoreHTTPSErrors: true,
    strict: false,
    testIdAttribute: 'data-testid',
    trace: 'on-first-retry',
    viewport: null
  },

  projects: [
    {
      name: 'chrome-con-perfil-real',
      use: {
        // Config manual (evita deviceScaleFactor del device cuando viewport es null)
        viewport: null,

        // Force real Chrome path
        executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',

        // Your actual profile
        userDataDir: 'C:\\Users\\radames\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 2',

        headless: false,
        args: [
          '--start-maximized',
          '--window-position=0,0',
          '--disable-restore-session-state', // clave!
          '--disable-session-crashed-bubble',
          '--no-first-run'
        ],
        ignoreDefaultArgs: ['--window-size'],

      },
    },
    {
      name: 'edge-con-perfil-real',
      use: {
        viewport: null,
        channel: 'msedge',
        browserName: 'chromium',
        headless: false,
        args: [
          '--start-maximized',
          '--window-position=0,0',
          '--disable-restore-session-state',
          '--disable-session-crashed-bubble',
          '--no-first-run'
        ],
        ignoreDefaultArgs: ['--window-size'],
      },
    },
  ],

});
