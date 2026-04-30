const {
  createVirtualFile,
  loadApp,
  readFixture,
  uploadFixture,
} = require("../helpers/appHarness");

async function generateSummary(app, fixtureName = "etsy-good.csv") {
  await uploadFixture(app, fixtureName);
  app.document.getElementById("processButton").click();
}

function getVisibleEditColumnsButton(document) {
  return Array.from(document.querySelectorAll("button")).find((button) => {
    return /edit columns/i.test(button.textContent) && !button.classList.contains("hidden") && !button.closest(".hidden");
  });
}

describe("UX, privacy, and security hammer tests", () => {
  let app;

  afterEach(() => {
    jest.restoreAllMocks();
    if (app) app.cleanup();
    app = null;
  });

  test("drag-and-drop affordance visibly activates and clears", () => {
    app = loadApp();

    const uploadShell = app.document.getElementById("uploadShell");

    uploadShell.dispatchEvent(new app.window.Event("dragenter", { bubbles: true }));
    expect(uploadShell.classList.contains("dragover")).toBe(true);

    uploadShell.dispatchEvent(new app.window.Event("dragleave", { bubbles: true }));
    expect(uploadShell.classList.contains("dragover")).toBe(false);
  });

  test("upload zone is keyboard-activatable like the button it presents itself as", () => {
    app = loadApp();

    const uploadShell = app.document.getElementById("uploadShell");
    const csvInput = app.document.getElementById("csvInput");
    const clickSpy = jest.spyOn(csvInput, "click");

    uploadShell.dispatchEvent(new app.window.KeyboardEvent("keydown", { bubbles: true, key: "Enter" }));
    uploadShell.dispatchEvent(new app.window.KeyboardEvent("keydown", { bubbles: true, key: " " }));

    expect(clickSpy).toHaveBeenCalledTimes(2);
  });

  test("extra-costs disclosure behaves like an accessible button", async () => {
    app = loadApp();

    await generateSummary(app);

    const toggle = app.document.getElementById("extraCostsToggle");
    const panel = app.document.getElementById("extraCostsPanel");

    expect(toggle.getAttribute("aria-expanded")).toBe("false");
    expect(panel.classList.contains("is-open")).toBe(false);

    toggle.click();

    expect(panel.classList.contains("is-open")).toBe(true);
    expect(toggle.getAttribute("aria-expanded")).toBe("true");

    toggle.dispatchEvent(new app.window.KeyboardEvent("keydown", { bubbles: true, key: "Enter" }));

    expect(panel.classList.contains("is-open")).toBe(false);
    expect(toggle.getAttribute("aria-expanded")).toBe("false");

    toggle.dispatchEvent(new app.window.KeyboardEvent("keydown", { bubbles: true, key: " " }));

    expect(panel.classList.contains("is-open")).toBe(true);
    expect(toggle.getAttribute("aria-expanded")).toBe("true");
  });

  test("summary view keeps a visible path back to column editing", async () => {
    app = loadApp();

    await generateSummary(app);

    const editButton = getVisibleEditColumnsButton(app.document);

    expect(editButton).toBeDefined();

    editButton.click();

    expect(app.document.getElementById("summarySection").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("mappingSection").classList.contains("hidden")).toBe(false);
  });

  test("uploaded filenames are rendered as text, not executable markup", async () => {
    app = loadApp();

    const maliciousName = '<img data-evil="file-name" src="x" onerror="window.__fileXss = true">.csv';

    await app.hooks.handleSelectedFiles([createVirtualFile(maliciousName, readFixture("etsy-good.csv"))]);

    expect(app.document.querySelector('[data-evil="file-name"]')).toBeNull();
    expect(app.document.getElementById("uploadedFilesList").textContent).toContain(maliciousName);
    expect(app.window.__fileXss).toBeUndefined();
  });

  test("CSV header text is rendered safely inside mapping controls", async () => {
    app = loadApp();

    const maliciousHeader = '<img data-evil="header-name" src="x" onerror="window.__headerXss = true">';
    const csv = [
      `Date,Type,Amount,Fees & Taxes,Net,${maliciousHeader}`,
      "30-Apr-25,Sale,£1.00,--,£1.00,Safe value",
    ].join("\n");

    await app.hooks.handleSelectedFiles([createVirtualFile("malicious-headers.csv", csv)]);

    const descriptionSelect = app.document.getElementById("mapping-description");
    const optionTexts = Array.from(descriptionSelect.options).map((option) => option.textContent);

    expect(optionTexts.some((text) => text.includes("data-evil") && text.includes("header-name"))).toBe(true);
    expect(app.document.querySelector('[data-evil="header-name"]')).toBeNull();
    expect(app.window.__headerXss).toBeUndefined();
  });

  test("processing and download stay local and do not touch network or storage APIs", async () => {
    app = loadApp();

    const fetchSpy = jest.fn();
    const sendBeaconSpy = jest.fn();
    const xmlHttpRequestSpy = jest.fn();
    const webSocketSpy = jest.fn();
    const indexedDbOpenSpy = jest.fn();
    const setItemSpy = jest.spyOn(app.window.Storage.prototype, "setItem");
    const getItemSpy = jest.spyOn(app.window.Storage.prototype, "getItem");
    const removeItemSpy = jest.spyOn(app.window.Storage.prototype, "removeItem");

    Object.defineProperty(app.window, "fetch", {
      configurable: true,
      value: fetchSpy,
    });
    Object.defineProperty(app.window.navigator, "sendBeacon", {
      configurable: true,
      value: sendBeaconSpy,
    });
    Object.defineProperty(app.window, "XMLHttpRequest", {
      configurable: true,
      value: xmlHttpRequestSpy,
    });
    Object.defineProperty(app.window, "WebSocket", {
      configurable: true,
      value: webSocketSpy,
    });
    Object.defineProperty(app.window, "indexedDB", {
      configurable: true,
      value: { open: indexedDbOpenSpy },
    });

    await generateSummary(app);
    app.document.getElementById("downloadCsvButton").click();

    expect(fetchSpy).not.toHaveBeenCalled();
    expect(sendBeaconSpy).not.toHaveBeenCalled();
    expect(xmlHttpRequestSpy).not.toHaveBeenCalled();
    expect(webSocketSpy).not.toHaveBeenCalled();
    expect(indexedDbOpenSpy).not.toHaveBeenCalled();
    expect(setItemSpy).not.toHaveBeenCalled();
    expect(getItemSpy).not.toHaveBeenCalled();
    expect(removeItemSpy).not.toHaveBeenCalled();
    expect(app.window.document.cookie).toBe("");
  });

  test("pagehide clears in-memory CSV and summary data", async () => {
    app = loadApp();

    await generateSummary(app);

    expect(app.hooks.state.uploadedFiles).toHaveLength(1);
    expect(app.hooks.state.summaryRows).toHaveLength(1);

    app.window.dispatchEvent(new app.window.Event("pagehide"));

    expect(app.hooks.state.uploadedFiles).toHaveLength(0);
    expect(app.hooks.state.combinedHeaders).toHaveLength(0);
    expect(app.hooks.state.summaryRows).toHaveLength(0);
    expect(app.hooks.state.extraCosts).toEqual({});
  });

  test("visibilitychange clears in-memory CSV and summary data when the page is hidden", async () => {
    app = loadApp();

    await generateSummary(app);

    Object.defineProperty(app.document, "visibilityState", {
      configurable: true,
      value: "hidden",
    });

    app.document.dispatchEvent(new app.window.Event("visibilitychange"));

    expect(app.hooks.state.uploadedFiles).toHaveLength(0);
    expect(app.hooks.state.combinedHeaders).toHaveLength(0);
    expect(app.hooks.state.summaryRows).toHaveLength(0);
    expect(app.hooks.state.extraCosts).toEqual({});
  });

  test("submit another resets the interface to a clean first-run state", async () => {
    app = loadApp();

    app.document.getElementById("companyNameInput").value = "Northwind Studio";
    app.document.getElementById("companyNameInput").dispatchEvent(new app.window.Event("input", { bubbles: true }));

    await generateSummary(app);
    app.document.getElementById("downloadCsvButton").click();

    expect(app.document.getElementById("submitAnotherButton").classList.contains("hidden")).toBe(false);

    app.document.getElementById("submitAnotherButton").click();

    expect(app.document.getElementById("fileCount").textContent).toBe("0 of 15 CSVs");
    expect(app.document.getElementById("companyNameInput").value).toBe("");
    expect(app.document.getElementById("uploadedFilesWrap").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("summarySection").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("wizardStepUpload").classList.contains("is-active")).toBe(true);
    expect(app.document.getElementById("statusMessage").textContent).toBe("");
  });

  test("page-level privacy affordances are present and external links suppress referrers", () => {
    app = loadApp();

    const runtimeAssets = Array.from(
      app.document.querySelectorAll('script[src], link[rel="stylesheet"][href], img[src]'),
    )
      .map((element) => element.getAttribute("src") || element.getAttribute("href") || "")
      .filter((value) => /^https?:\/\//i.test(value));

    const externalLinks = Array.from(app.document.querySelectorAll('a[href^="http"]'));

    expect(app.document.querySelector('meta[name="referrer"]')?.getAttribute("content")).toBe("no-referrer");
    expect(runtimeAssets).toEqual([]);
    expect(externalLinks.length).toBeGreaterThan(0);
    externalLinks.forEach((link) => {
      expect(link.getAttribute("rel") || "").toContain("noreferrer");
      expect(link.getAttribute("referrerpolicy")).toBe("no-referrer");
    });

    const pageText = app.document.body.textContent;
    expect(pageText).toMatch(/processed locally in your browser/i);
    expect(pageText).toMatch(/not uploaded|nothing is uploaded or stored/i);
  });
});