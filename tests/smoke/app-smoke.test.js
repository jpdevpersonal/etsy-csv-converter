const { loadApp, uploadFixture } = require("../helpers/appHarness");

describe("app harness smoke test", () => {
  let app;

  afterEach(() => {
    if (app) app.cleanup();
    app = null;
  });

  test("loads the browser script and blocks a CSV with no header row", async () => {
    app = loadApp();

    expect(app.hooks.parseCurrencyValue("-£0.05")).toEqual({ valid: true, value: -0.05 });

    await uploadFixture(app, "etsy-no-headers.csv");

    expect(app.document.getElementById("uploadSectionStatus").textContent).toContain(
      "check the file format",
    );
    expect(app.document.getElementById("processButton").disabled).toBe(true);
    expect(app.document.getElementById("fileCount").textContent).toBe("0 of 15 CSVs");
  });
});