const {
  createVirtualFile,
  loadApp,
  readFixture,
  uploadFixture,
} = require("../helpers/appHarness");

function getSummaryRowTexts(document) {
  return Array.from(document.querySelectorAll("#summaryTableBody tr td")).map((cell) => cell.textContent.trim());
}

async function generateSummary(app, fixtureName = "etsy-good.csv") {
  await uploadFixture(app, fixtureName);
  app.document.getElementById("processButton").click();
}

describe("UI workflow and user-error handling", () => {
  let app;

  afterEach(() => {
    if (app) app.cleanup();
    app = null;
  });

  test("starts with upload step active and no report generated", () => {
    app = loadApp();

    expect(app.document.getElementById("fileCount").textContent).toBe("0 of 15 CSVs");
    expect(app.document.getElementById("processButton").disabled).toBe(true);
    expect(app.document.getElementById("uploadedFilesWrap").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("mappingSection").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("summarySection").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("uploadSectionStatus").textContent).toBe("");
  });

  test("rejects non-CSV uploads before reading file contents", async () => {
    app = loadApp();

    await app.hooks.handleSelectedFiles([createVirtualFile("notes.txt", "not,a,csv")]);

    expect(app.document.getElementById("uploadSectionStatus").textContent).toContain("ending in .csv");
    expect(app.document.getElementById("fileCount").textContent).toBe("0 of 15 CSVs");
  });

  test("blocks CSV files with blank headers", async () => {
    app = loadApp();

    const blankHeaderCsv = [
      "Date,Type,,Info,Currency,Amount,Fees & Taxes,Net,Tax Details",
      "30-Apr-25,Sale,Payment for Order #3672216433,,GBP,£1.37,--,£1.37,--",
    ].join("\n");

    await app.hooks.handleSelectedFiles([createVirtualFile("blank-headers.csv", blankHeaderCsv)]);

    expect(app.document.getElementById("mappingSectionStatus").textContent).toContain(
      "check the file format",
    );
    expect(app.document.getElementById("processButton").disabled).toBe(true);
    expect(app.document.getElementById("uploadedFilesWrap").classList.contains("hidden")).toBe(true);
  });

  test("blocks the whole batch and shows the file name when one uploaded CSV has header-format errors", async () => {
    app = loadApp();

    const badHeaderCsv = [
      "Date,Type,,Info,Currency,Amount,Fees & Taxes,Net,Tax Details",
      "30-Apr-25,Sale,Payment for Order #3672216433,,GBP,£1.37,--,£1.37,--",
    ].join("\n");

    await app.hooks.handleSelectedFiles([
      createVirtualFile("good.csv", readFixture("etsy-good.csv")),
      createVirtualFile("bad-header.csv", badHeaderCsv),
    ]);

    expect(app.document.getElementById("mappingSectionStatus").textContent).toContain("bad-header.csv");
    expect(app.document.getElementById("mappingSectionStatus").textContent).toContain("check the file format");
    expect(app.document.getElementById("fileCount").textContent).toBe("0 of 15 CSVs");
    expect(app.document.getElementById("processButton").disabled).toBe(true);
    expect(app.document.getElementById("uploadedFilesWrap").classList.contains("hidden")).toBe(true);
  });

  test("auto-fills mappings and generates a summary from a standard Etsy export", async () => {
    app = loadApp();

    await generateSummary(app);

    expect(app.document.getElementById("mapping-date").value).toBe("Date");
    expect(app.document.getElementById("mapping-type").value).toBe("Type");
    expect(app.document.getElementById("mapping-amount").value).toBe("Amount");
    expect(app.document.getElementById("summarySection").classList.contains("hidden")).toBe(false);
    expect(app.document.getElementById("summarySectionStatus").textContent).toContain("profit breakdown is ready");
    expect(getSummaryRowTexts(app.document)).toEqual([
      "2025-04",
      "£1.37",
      "-£1.14",
      "£0.00",
      "£0.00",
      "£0.23",
      "£0.00",
      "£0.23",
    ]);
  });

  test("recalculates estimated profit when extra monthly costs are entered", async () => {
    app = loadApp();

    await generateSummary(app);

    app.document.getElementById("extraCostsToggle").click();

    const packagingInput = app.document.querySelector(
      '#costTableBody input[data-month="2025-04"][data-cost-key="packaging"]',
    );
    const adsInput = app.document.querySelector(
      '#costTableBody input[data-month="2025-04"][data-cost-key="ads"]',
    );

    packagingInput.value = "0.10";
    packagingInput.dispatchEvent(new app.window.Event("input", { bubbles: true }));
    adsInput.value = "0.05";
    adsInput.dispatchEvent(new app.window.Event("input", { bubbles: true }));

    expect(getSummaryRowTexts(app.document)).toEqual([
      "2025-04",
      "£1.37",
      "-£1.14",
      "£0.00",
      "£0.00",
      "£0.23",
      "£0.15",
      "£0.08",
    ]);
  });

  test("surfaces a validation error when a required mapping is cleared", async () => {
    app = loadApp();

    await uploadFixture(app, "etsy-good.csv");

    const amountSelect = app.document.getElementById("mapping-amount");
    amountSelect.value = "";
    amountSelect.dispatchEvent(new app.window.Event("change", { bubbles: true }));

    expect(app.document.getElementById("mappingSectionStatus").textContent).toContain(
      "Choose Date, Transaction type, and Amount",
    );
    expect(app.document.getElementById("processButton").disabled).toBe(true);
    expect(app.document.getElementById("mappingSection").classList.contains("hidden")).toBe(false);
  });

  test("shows a processing error when numeric cells are unreadable", async () => {
    app = loadApp();

    await uploadFixture(app, "etsy-invalid-amount.csv");
    app.document.getElementById("processButton").click();

    expect(app.document.getElementById("mappingSection").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("processSectionStatus").textContent).toContain(
      "money values are not in the expected format",
    );
    expect(app.document.getElementById("processSectionStatus").textContent).toContain(
      "Check the CSV format and try again, or delete the file from the upload list.",
    );
    expect(app.document.getElementById("processHelper").textContent).toBe("");
    expect(app.document.getElementById("processButton").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("cancelButton").classList.contains("hidden")).toBe(false);
    expect(app.document.getElementById("summarySection").classList.contains("hidden")).toBe(true);
  });

  test("shows the offending file name and blocks progress when one file in a multi-file batch has unreadable numeric values", async () => {
    app = loadApp();

    await app.hooks.handleSelectedFiles([
      createVirtualFile("good.csv", readFixture("etsy-good.csv")),
      createVirtualFile("broken-values.csv", readFixture("etsy-invalid-amount.csv")),
    ]);

    app.document.getElementById("processButton").click();

    expect(app.document.getElementById("processSectionStatus").textContent).toContain("broken-values.csv");
    expect(app.document.getElementById("processSectionStatus").textContent).toContain(
      "money values are not in the expected format",
    );
    expect(app.document.getElementById("processButton").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("cancelButton").classList.contains("hidden")).toBe(false);
    expect(app.document.getElementById("summarySection").classList.contains("hidden")).toBe(true);
  });

  test("removing a bad CSV after a processing error restores Step 2 and allows regeneration", async () => {
    app = loadApp();

    await app.hooks.handleSelectedFiles([
      createVirtualFile("good.csv", readFixture("etsy-good.csv")),
      createVirtualFile("broken-values.csv", readFixture("etsy-invalid-amount.csv")),
    ]);

    app.document.getElementById("processButton").click();
    app.document.querySelector('button[aria-label="Remove broken-values.csv"]').click();

    expect(app.document.getElementById("processSectionStatus").textContent).toBe("");
    expect(app.document.getElementById("processButton").classList.contains("hidden")).toBe(false);
    expect(app.document.getElementById("cancelButton").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("fileCount").textContent).toBe("1 of 15 CSVs");
    expect(app.document.getElementById("processCard").textContent).toContain("Step 2 of 3");

    app.document.getElementById("processButton").click();
    expect(app.document.getElementById("summarySection").classList.contains("hidden")).toBe(false);
  });

  test("uploading a new valid CSV after a processing error restarts Step 2", async () => {
    app = loadApp();

    await uploadFixture(app, "etsy-invalid-amount.csv");
    app.document.getElementById("processButton").click();

    await uploadFixture(app, "etsy-good.csv");

    expect(app.document.getElementById("processSectionStatus").textContent).toBe("");
    expect(app.document.getElementById("processButton").classList.contains("hidden")).toBe(false);
    expect(app.document.getElementById("cancelButton").classList.contains("hidden")).toBe(true);
    expect(app.document.getElementById("processCard").textContent).toContain("Step 2 of 3");
  });

  test("enforces the 15-file upload limit and warns about skipped files", async () => {
    app = loadApp();

    const csv = readFixture("etsy-good.csv");
    const files = Array.from({ length: 16 }, (_value, index) =>
      createVirtualFile(`batch-${index + 1}.csv`, csv),
    );

    await app.hooks.handleSelectedFiles(files);

    expect(app.document.getElementById("fileCount").textContent).toBe("15 of 15 CSVs");
    expect(app.document.getElementById("uploadSectionStatus").textContent).toContain("15-file limit");
    expect(app.document.querySelectorAll("#uploadedFilesList li")).toHaveLength(15);
  });

  test("shows an error when download is attempted before a report exists", () => {
    app = loadApp();

    app.document.getElementById("downloadCsvButton").click();

    expect(app.document.getElementById("summarySectionStatus").textContent).toContain(
      "Generate your report before downloading.",
    );
  });

  test("creates a downloadable CSV report once the summary exists", async () => {
    app = loadApp();

    await generateSummary(app);
    app.document.getElementById("downloadCsvButton").click();

    expect(app.window.URL.createObjectURL).toHaveBeenCalledTimes(1);
    expect(app.document.getElementById("downloadReadyMessage").textContent).toContain("downloaded");
    expect(app.document.getElementById("submitAnotherButton").classList.contains("hidden")).toBe(false);
  });
});