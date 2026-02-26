const { defineConfig } = require('@playwright/test');

module.exports = defineConfig({
  testDir: './tests',
  timeout: 15000,
  use: {
    headless: true,
  },
  projects: [
    {
      name: 'bridge-server',
      testMatch: '**/bridge-server.test.js',
      use: { browserName: 'chromium' },
    },
    {
      name: 'content-script',
      testMatch: '**/content-script.test.js',
      use: { browserName: 'chromium' },
    },
  ],
  reporter: 'list',
});
