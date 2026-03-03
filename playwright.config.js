const { defineConfig } = require('@playwright/test');

module.exports = defineConfig({
  testDir: './tests',
  timeout: 15000,
  // bridge-server と mcp-server が同じポート(3765)を使うため順次実行
  workers: 1,
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
    {
      name: 'content-script-dm',
      testMatch: '**/content-script-dm.test.js',
      use: { browserName: 'chromium' },
    },
    {
      name: 'content-script-attachments',
      testMatch: '**/content-script-attachments.test.js',
      use: { browserName: 'chromium' },
    },
    {
      name: 'content-script-inspect',
      testMatch: '**/content-script-inspect.test.js',
      use: { browserName: 'chromium' },
    },
    {
      name: 'popup',
      testMatch: '**/popup.test.js',
      use: { browserName: 'chromium' },
    },
    {
      name: 'content-script-token',
      testMatch: '**/content-script-token.test.js',
      use: { browserName: 'chromium' },
    },
    {
      name: 'mcp-server',
      testMatch: '**/mcp-server.test.js',
      use: { browserName: 'chromium' },
    },
    {
      name: 'mcp-server-download',
      testMatch: '**/mcp-server-download.test.js',
      use: { browserName: 'chromium' },
    },
  ],
  reporter: 'list',
});
