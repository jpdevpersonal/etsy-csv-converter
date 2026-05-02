const fs = require("node:fs");
const path = require("node:path");
const { JSDOM } = require("jsdom");

const ROOT_DIR = path.resolve(__dirname, "../..");
const INDEX_PATH = path.join(ROOT_DIR, "index.html");
const APP_PATH = path.join(ROOT_DIR, "app.js");
const FIXTURES_DIR = path.join(ROOT_DIR, "tests", "fixtures");

const EXTERNAL_SCRIPT_PATTERN = /<script\s+src="app\.js\?[^\"]+"\s+defer><\/script>/i;

const TEST_HOOK_SOURCE = `
window.__appTestHooks__ = {
  MAX_FILE_COUNT,
  MAPPING_FIELDS,
  state,
  parseCsvText,
  parseCurrencyValue,
  detectDateValue,
  detectMonthKey,
  guessMappings,
  buildMonthlySummary,
  buildCombinedMonthlySummary,
  renderSummaryTable,
  renderCostTable,
  getSummaryWithCosts,
  getExportRows,
  handleSelectedFiles,
  resetToolState
};
`;

function createFakeFileReader() {
  return class FakeFileReader {
    readAsText(file) {
      Promise.resolve()
        .then(async () => {
          if (file && typeof file.__contents === "string") return file.__contents;
          if (file && typeof file.text === "function") return file.text();
          throw new Error("Unsupported test file object");
        })
        .then((contents) => {
          this.result = contents;
          if (typeof this.onload === "function") this.onload();
        })
        .catch((error) => {
          this.error = error;
          if (typeof this.onerror === "function") this.onerror(error);
        });
    }
  };
}

function loadApp() {
  const html = fs
    .readFileSync(INDEX_PATH, "utf8")
    .replace(EXTERNAL_SCRIPT_PATTERN, "");

  const dom = new JSDOM(html, {
    url: "http://localhost/",
    pretendToBeVisual: true,
    runScripts: "outside-only",
  });

  const { window } = dom;
  window.FileReader = createFakeFileReader();
  window.HTMLElement.prototype.scrollIntoView = jest.fn();
  window.URL.createObjectURL = jest.fn(() => "blob:test-download");
  window.URL.revokeObjectURL = jest.fn();
  window.HTMLAnchorElement.prototype.click = jest.fn();

  const source = fs.readFileSync(APP_PATH, "utf8");
  window.eval(`${source}\n${TEST_HOOK_SOURCE}`);

  return {
    window,
    document: window.document,
    hooks: window.__appTestHooks__,
    cleanup() {
      window.close();
    },
  };
}

function readFixture(name) {
  return fs.readFileSync(path.join(FIXTURES_DIR, name), "utf8");
}

function createVirtualFile(name, contents) {
  return { name, __contents: contents };
}

async function uploadFixture(app, name) {
  const contents = readFixture(name);
  await app.hooks.handleSelectedFiles([createVirtualFile(name, contents)]);
}

module.exports = {
  createVirtualFile,
  loadApp,
  readFixture,
  uploadFixture,
};