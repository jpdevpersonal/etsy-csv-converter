const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests/e2e",
  fullyParallel: false,
  reporter: "list",
  use: {
    baseURL: "http://127.0.0.1:4173",
    browserName: "chromium",
    headless: true,
  },
  webServer: {
    command: "node tests/helpers/static-server.js",
    url: "http://127.0.0.1:4173/",
    reuseExistingServer: true,
    timeout: 30000,
  },
});