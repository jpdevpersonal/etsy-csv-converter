const path = require("node:path");
const { test, expect } = require("@playwright/test");

const fixturePath = path.join(__dirname, "..", "fixtures", "etsy-good.csv");

test("served page includes the configured security headers", async ({ page }) => {
  const response = await page.goto("/");
  const headers = response.headers();

  expect(headers["referrer-policy"]).toBe("no-referrer");
  expect(headers["x-content-type-options"]).toBe("nosniff");
  expect(headers["x-frame-options"]).toBe("DENY");
  expect(headers["permissions-policy"]).toContain("camera=()");
  expect(headers["strict-transport-security"]).toContain("max-age=31536000");
  expect(headers["content-security-policy"]).toContain("default-src 'self'");
  expect(headers["content-security-policy"]).toContain("frame-ancestors 'none'");
  expect(headers["content-security-policy"]).toContain("object-src 'none'");
});