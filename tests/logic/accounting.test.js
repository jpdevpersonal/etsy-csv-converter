const { createVirtualFile, loadApp, readFixture } = require("../helpers/appHarness");

function expectMonthRow(row, expected) {
  expect(row.month).toBe(expected.month);
  expect(row.revenue).toBeCloseTo(expected.revenue, 2);
  expect(row.fees).toBeCloseTo(expected.fees, 2);
  expect(row.refunds).toBeCloseTo(expected.refunds, 2);
  expect(row.adjustments).toBeCloseTo(expected.adjustments, 2);
  expect(row.netReceived).toBeCloseTo(expected.netReceived, 2);
}

describe("app.js accounting logic", () => {
  let app;

  afterEach(() => {
    if (app) app.cleanup();
    app = null;
  });

  test("parseCsvText handles BOMs, semicolon delimiters, and quoted commas", () => {
    app = loadApp();

    const parsed = app.hooks.parseCsvText(
      "\uFEFFDate;Type;Title;Amount\n30-Apr-25;Fee;\"Transaction fee: Alpha, Beta\";--",
    );

    expect(parsed.headers).toEqual(["Date", "Type", "Title", "Amount"]);
    expect(parsed.rows).toHaveLength(1);
    expect(parsed.rows[0][2]).toBe("Transaction fee: Alpha, Beta");
  });

  test.each([
    ["dash placeholder", "--", 0],
    ["negative sterling", "-£0.05", -0.05],
    ["accounting parentheses", "(€12.50)", -12.5],
    ["decimal comma", "€1.234,56", 1234.56],
    ["comma thousands", "$1,234.56", 1234.56],
  ])("parseCurrencyValue handles %s", (_label, value, expected) => {
    app = loadApp();

    expect(app.hooks.parseCurrencyValue(value)).toEqual({ valid: true, value: expected });
  });

  test("detectMonthKey handles Etsy short dates, ISO dates, and day-first dates", () => {
    app = loadApp();

    expect(app.hooks.detectMonthKey("30-Apr-25")).toBe("2025-04");
    expect(app.hooks.detectMonthKey("2025-12-01")).toBe("2025-12");
    expect(app.hooks.detectMonthKey("29/02/2024")).toBe("2024-02");
  });

  test("guessMappings auto-detects the standard Etsy export columns", () => {
    app = loadApp();

    const parsed = app.hooks.parseCsvText(readFixture("etsy-good.csv"));
    const mappings = app.hooks.guessMappings(parsed.headers);

    expect(mappings).toMatchObject({
      date: "Date",
      type: "Type",
      amount: "Amount",
      fee: "Fees & Taxes",
      net: "Net",
    });
    expect(["Title", "Info"]).toContain(mappings.description);
  });

  test("buildMonthlySummary aggregates a standard Etsy export and skips blank-date rows", () => {
    app = loadApp();

    const parsed = app.hooks.parseCsvText(readFixture("etsy-good.csv"));
    const result = app.hooks.buildMonthlySummary(parsed.rows, parsed.headers, {
      date: "Date",
      type: "Type",
      amount: "Amount",
      fee: "Fees & Taxes",
      net: "Net",
      description: "Title",
    });

    expect(result.ok).toBe(true);
    expect(result.summaryRows).toHaveLength(1);
    expectMonthRow(result.summaryRows[0], {
      month: "2025-04",
      revenue: 1.37,
      fees: -1.14,
      refunds: 0,
      adjustments: 0,
      netReceived: 0.23,
    });
  });

  test("buildMonthlySummary handles refunds, adjustments, and deposit fallback extraction across months", () => {
    app = loadApp();

    const parsed = app.hooks.parseCsvText(readFixture("etsy-multi-month.csv"));
    const result = app.hooks.buildMonthlySummary(parsed.rows, parsed.headers, {
      date: "Date",
      type: "Type",
      amount: "Amount",
      fee: "Fees & Taxes",
      net: "Net",
      description: "Info",
    });

    expect(result.ok).toBe(true);
    expect(result.summaryRows).toHaveLength(2);
    expectMonthRow(result.summaryRows[0], {
      month: "2024-02",
      revenue: 25,
      fees: -1.9,
      refunds: 0,
      adjustments: 3,
      netReceived: 26.1,
    });
    expectMonthRow(result.summaryRows[1], {
      month: "2024-03",
      revenue: 12.5,
      fees: -1.75,
      refunds: -10,
      adjustments: 43.67,
      netReceived: 44.42,
    });
  });

  test("buildCombinedMonthlySummary merges multiple files into the same month buckets", () => {
    app = loadApp();

    const parsed = app.hooks.parseCsvText(readFixture("etsy-good.csv"));
    const guessedMappings = app.hooks.guessMappings(parsed.headers);
    const files = [
      { name: "april-a.csv", headers: parsed.headers, rows: parsed.rows, guessedMappings },
      { name: "april-b.csv", headers: parsed.headers, rows: parsed.rows, guessedMappings },
    ];

    const result = app.hooks.buildCombinedMonthlySummary(files, {
      date: "Date",
      type: "Type",
      amount: "Amount",
      fee: "Fees & Taxes",
      net: "Net",
      description: "Title",
    });

    expect(result.ok).toBe(true);
    expect(result.summaryRows).toHaveLength(1);
    expectMonthRow(result.summaryRows[0], {
      month: "2025-04",
      revenue: 2.74,
      fees: -2.28,
      refunds: 0,
      adjustments: 0,
      netReceived: 0.46,
    });
  });

  test("buildMonthlySummary rejects unreadable numeric values", () => {
    app = loadApp();

    const parsed = app.hooks.parseCsvText(readFixture("etsy-invalid-amount.csv"));
    const result = app.hooks.buildMonthlySummary(parsed.rows, parsed.headers, {
      date: "Date",
      type: "Type",
      amount: "Amount",
      fee: "Fees & Taxes",
      net: "Net",
      description: "Title",
    });

    expect(result.ok).toBe(false);
    expect(result.message).toContain("could not be read as numbers");
  });

  test("handleSelectedFiles accepts quoted Etsy exports with long-form month dates when the progress tracker is absent", async () => {
    app = loadApp();

    const sample = [
      'Date,Type,Title,Info,Currency,Amount,"Fees & Taxes",Net,"Tax Details"',
      '"30 April, 2025",VAT,"VAT: Processing Fee","Order #3672216433",GBP,--,-£0.05,-£0.05,--',
      '"30 April, 2025",Fee,"Processing fee","Order #3672216433",GBP,--,-£0.25,-£0.25,--',
      '"30 April, 2025",Sale,"Payment for Order #3672216433",,GBP,£1.37,--,£1.37,--',
    ].join("\n");

    await app.hooks.handleSelectedFiles([createVirtualFile("etsy-long-month.csv", sample)]);

    expect(app.hooks.state.uploadedFiles).toHaveLength(1);
    expect(app.hooks.state.uploadedFiles[0].headers).toEqual([
      "Date",
      "Type",
      "Title",
      "Info",
      "Currency",
      "Amount",
      "Fees & Taxes",
      "Net",
      "Tax Details",
    ]);
    expect(app.document.getElementById("uploadSectionStatus").textContent.trim()).toBe("");
  });
});