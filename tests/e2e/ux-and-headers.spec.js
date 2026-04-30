const path = require("node:path");
const { test, expect } = require("@playwright/test");

const fixturePath = path.join(__dirname, "..", "fixtures", "etsy-good.csv");

test("browser flow supports keyboard-first interaction and recovery", async ({ page }) => {
  await page.goto("/");

  const uploadShell = page.locator("#uploadShell");
  await uploadShell.focus();

  const fileChooserPromise = page.waitForEvent("filechooser");
  await page.keyboard.press("Enter");
  const fileChooser = await fileChooserPromise;
  await fileChooser.setFiles(fixturePath);

  await expect(page.locator("#fileCount")).toHaveText("1 of 15 CSVs");
  await page.getByRole("button", { name: /generate profit breakdown/i }).click();

  await expect(page.locator("#summarySection")).toBeVisible();
  await expect(page.locator("#summarySectionStatus")).toContainText(/profit breakdown is ready/i);

  const extraCostsToggle = page.locator("#extraCostsToggle");
  await extraCostsToggle.focus();
  await page.keyboard.press("Space");
  await expect(extraCostsToggle).toHaveAttribute("aria-expanded", "true");

  await page.getByRole("button", { name: /edit columns/i }).nth(0).click();
  await expect(page.locator("#mappingSection")).toBeVisible();
  await expect(page.locator("#summarySection")).toBeHidden();
});

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