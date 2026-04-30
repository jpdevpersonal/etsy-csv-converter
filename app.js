/* ===================================================================
   Etsy CSV Monthly Accounts Cleaner — Application Logic
   SimpleBizToolkit.com

   Sections:
   1.  Configuration
   2.  State
   3.  DOM References
   4.  Event Listeners
   5.  UI Helpers
   6.  Upload Handling
   7.  CSV Parsing
   8.  Column Mapping
   9.  Monthly Summary Calculation
   10. Table Rendering
   11. Export — CSV, XLS, XLSX
   12. Download
   13. Utilities
   14. Initialise
   =================================================================== */

/* =================================================================
   1. Configuration
   ================================================================= */

const MAX_FILE_COUNT = 15;

const MAPPING_FIELDS = [
  {
    key: "date",
    label: "Date column",
    required: true,
    help: "When did the transaction happen? Usually labelled \"Date\" or \"Transaction date\".",
    guesses: ["date", "transaction date", "order date", "created date", "posted date"],
  },
  {
    key: "type",
    label: "Transaction type column",
    required: true,
    help: "What kind of transaction was it? Usually \"Type\" or \"Transaction type\" — this is how the tool tells sales from fees and refunds.",
    guesses: ["type", "transaction type", "activity type", "entry type", "kind"],
  },
  {
    key: "amount",
    label: "Amount column",
    required: true,
    help: "The monetary value of each row. Usually \"Amount\" or \"Gross\". Etsy placeholders like \"--\" are treated as 0.",
    guesses: ["amount", "gross", "total", "sale amount", "payment amount", "gross amount"],
  },
  {
    key: "fee",
    label: "Fee column",
    required: false,
    help: "Optional. A column that lists the Etsy fee charged per transaction. Etsy placeholders like \"--\" are treated as 0. Leave as \"Not used\" if your export doesn't have one.",
    guesses: ["fee", "fees", "fees & taxes", "fees and taxes", "etsy fee", "transaction fee", "fee amount"],
  },
  {
    key: "net",
    label: "Net column",
    required: false,
    help: "Optional. The amount received after fees. Etsy placeholders like \"--\" are treated as 0. Leave as \"Not used\" if your export doesn't include this.",
    guesses: ["net", "net amount", "amount net", "net total", "net received"],
  },
  {
    key: "description",
    label: "Description column",
    required: false,
    reviewIfMissing: true,
    help: "Optional. Any extra detail column such as notes or item title. Used to help classify ambiguous transactions.",
    guesses: ["description", "details", "memo", "notes", "title", "info"],
  },
];

const EXPORT_COLUMNS = [
  { key: "month",           label: "Month",            width: 14, type: "text"     },
  { key: "revenue",         label: "Revenue",           width: 16, type: "currency" },
  { key: "fees",            label: "Etsy Fees",         width: 16, type: "currency" },
  { key: "refunds",         label: "Refunds",           width: 16, type: "currency" },
  { key: "adjustments",     label: "Adjustments",       width: 16, type: "currency" },
  { key: "netReceived",     label: "Net Received",      width: 16, type: "currency" },
  { key: "extraCosts",      label: "Extra Costs",       width: 16, type: "currency" },
  { key: "estimatedProfit", label: "Estimated Profit",  width: 18, type: "currency" },
];

/* =================================================================
   2. State
   ================================================================= */

const state = {
  uploadedFiles: [],
  combinedHeaders: [],
  mappings: {
    date: "",
    type: "",
    amount: "",
    fee: "",
    net: "",
    description: "",
  },
  summaryRows: [],
  extraCosts: {},
  displaySymbol: "",
  companyName: "",
  mappingLocked: false,
  mappingVisible: false,
  mappingBlockMessage: "",
  hasDownloadedFile: false,
};

/* =================================================================
   3. DOM References
   ================================================================= */

const uploadCard        = document.getElementById("uploadCard");
const uploadButton      = document.getElementById("uploadButton");
const csvInput          = document.getElementById("csvInput");
const uploadShell       = document.getElementById("uploadShell");
const companyNameInput  = document.getElementById("companyNameInput");
const fileCount         = document.getElementById("fileCount");
const uploadedFilesWrap = document.getElementById("uploadedFilesWrap");
const uploadedFilesList = document.getElementById("uploadedFilesList");
const uploadSectionStatus = document.getElementById("uploadSectionStatus");

const mappingSection      = document.getElementById("mappingSection");
const mappingGrid         = document.getElementById("mappingGrid");
const mappingHelper       = document.getElementById("mappingHelper");
const mappingSectionStatus = document.getElementById("mappingSectionStatus");
const mappingLockBanner   = document.getElementById("mappingLockBanner");
const mappingResetButton  = document.getElementById("mappingResetButton");
const processCard         = document.getElementById("processCard");
const processButton       = document.getElementById("processButton");
const processHelper       = document.getElementById("processHelper");
const editMappingsButton  = document.getElementById("editMappingsButton");
const cancelButton        = document.getElementById("cancelButton");

const summarySection      = document.getElementById("summarySection");
const summarySectionStatus = document.getElementById("summarySectionStatus");
const extraCostsPanel     = document.getElementById("extraCostsPanel");
const extraCostsToggle    = document.getElementById("extraCostsToggle");
const costTableBody       = document.getElementById("costTableBody");
const summaryTableBody    = document.getElementById("summaryTableBody");
const summaryTableFoot    = document.getElementById("summaryTableFoot");

const downloadCsvButton  = document.getElementById("downloadCsvButton");
const downloadXlsButton  = document.getElementById("downloadXlsButton");
const downloadXlsxButton = document.getElementById("downloadXlsxButton");
const downloadReadyMessage = document.getElementById("downloadReadyMessage");
const submitAnotherButton = document.getElementById("submitAnotherButton");

const statusMessage     = document.getElementById("statusMessage");
const wizardStepUpload  = document.getElementById("wizardStepUpload");
const wizardStepMapping = document.getElementById("wizardStepMapping");
const wizardStepSummary = document.getElementById("wizardStepSummary");

const DEFAULT_MAPPING_HELPER_TEXT = "We've automatically matched the columns in your CSV to the fields below. The three starred fields (*) are required — if a dropdown shows Please choose a column, pick the closest match from your CSV before continuing.";
const BLOCKED_MAPPING_HELPER_TEXT = "We cannot continue because this file does not contain usable Etsy column headers. Export a fresh CSV from Etsy, then try again.";

/** Maps section keys to their inline status elements. */
const SECTION_STATUS_MAP = {
  upload:  uploadSectionStatus,
  mapping: mappingSectionStatus,
  summary: summarySectionStatus,
};

/* =================================================================
   4. Event Listeners
   ================================================================= */

// Upload button
uploadButton.addEventListener("click", () => csvInput.click());

// Company name
companyNameInput.addEventListener("input", () => {
  state.companyName = companyNameInput.value.trim();
});

// File input change
csvInput.addEventListener("change", async (event) => {
  await handleSelectedFiles(event.target.files || []);
  csvInput.value = "";
});

// Drag and drop on upload zone
["dragenter", "dragover"].forEach((eventName) => {
  uploadShell.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadShell.classList.add("dragover");
  });
});

["dragleave", "dragend", "drop"].forEach((eventName) => {
  uploadShell.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadShell.classList.remove("dragover");
  });
});

uploadShell.addEventListener("drop", async (event) => {
  await handleSelectedFiles(event.dataTransfer.files || []);
});

// Edit mappings button (unlock after summary is created)
editMappingsButton.addEventListener("click", () => {
  state.mappingLocked = false;
  state.mappingVisible = true;
  resetSummaryState();
  refreshMappingValidationState();
  focusSection(mappingSection);
});

// Extra costs collapsible toggle
if (extraCostsToggle) {
  extraCostsToggle.addEventListener("click", () => {
    extraCostsPanel.classList.toggle("is-open");
  });
}

// Cancel and Submit Another both perform a full reset
submitAnotherButton.addEventListener("click", resetToolState);
cancelButton.addEventListener("click", resetToolState);
if (mappingResetButton) mappingResetButton.addEventListener("click", resetToolState);

// Process / create summary
processButton.addEventListener("click", () => {
  if (!state.uploadedFiles.length) {
    announce("upload", "Choose at least one Etsy CSV before generating your profit breakdown.", "error");
    return;
  }

  updateMappingsFromInputs();

  const mappingValidation = refreshMappingValidationState({ scrollOnBlock: true });
  if (!mappingValidation.ok) return;

  const missingRequired = MAPPING_FIELDS
    .filter((field) => field.required && !state.mappings[field.key])
    .map((field) => field.label.replace(" column", ""));

  if (missingRequired.length) {
    announce("mapping", "Choose Date, Transaction type, and Amount before generating your report.", "error");
    return;
  }

  const result = buildCombinedMonthlySummary(state.uploadedFiles, state.mappings);
  if (!result.ok) {
    announce("mapping", result.message, "error");
    return;
  }

  state.summaryRows = result.summaryRows;
  ensureExtraCostBuckets();
  renderCostTable();
  renderSummaryTable();
  summarySection.classList.remove("hidden");
  state.mappingLocked = true;
  state.mappingVisible = false;
  refreshMappingValidationState();
  clearSectionStatus("upload");

  announce("summary", "Your profit breakdown is ready. Add extra costs if you want, or download your report now.", "success");
  clearDownloadReadyMessage();
  state.hasDownloadedFile = false;
  submitAnotherButton.classList.add("hidden");
  updateWizardProgress();
  focusSection(summarySection);
});

// Download buttons
downloadCsvButton.addEventListener("click", () => {
  if (!state.summaryRows.length) {
    announce("summary", "Generate your report before downloading.", "error");
    return;
  }
  downloadFile(convertSummaryToCsv(getExportRows()), buildDownloadName("csv"), "text/csv;charset=utf-8;");
  handleSuccessfulDownload();
});

downloadXlsButton.addEventListener("click", () => {
  if (!state.summaryRows.length) {
    announce("summary", "Generate your report before downloading.", "error");
    return;
  }
  downloadFile(createXlsWorkbookXml(getExportRows()), buildDownloadName("xls"), "application/vnd.ms-excel;charset=utf-8;");
  handleSuccessfulDownload();
});

downloadXlsxButton.addEventListener("click", () => {
  if (!state.summaryRows.length) {
    announce("summary", "Generate your report before downloading.", "error");
    return;
  }
  downloadFile(
    createXlsxWorkbook(getExportRows()),
    buildDownloadName("xlsx"),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  handleSuccessfulDownload();
});

/* =================================================================
   5. UI Helpers
   ================================================================= */

/** Write text to the global status bar at the top of the page. */
function showMessage(text, type) {
  statusMessage.textContent = text;
  statusMessage.dataset.type = type;
}

/** Set or clear a status element's text and data-type attribute. */
function setElementStatus(element, text, type) {
  if (!element) return;

  if (!text) {
    element.textContent = "";
    delete element.dataset.type;
    return;
  }

  element.textContent = text;
  element.dataset.type = type;
}

/** Clear a named section's inline status. */
function clearSectionStatus(sectionKey) {
  setElementStatus(SECTION_STATUS_MAP[sectionKey], "", "");
}

/** Write to both the global status bar and a section's inline status. */
function announce(sectionKey, text, type) {
  showMessage(text, type);
  setElementStatus(SECTION_STATUS_MAP[sectionKey], text, type);
}

/** Clear the global status bar. */
function clearMessage() {
  statusMessage.textContent = "";
  delete statusMessage.dataset.type;
}

/** Set the download-ready inline message. */
function setDownloadReadyMessage(text, type) {
  setElementStatus(downloadReadyMessage, text, type);
}

/** Clear the download-ready inline message. */
function clearDownloadReadyMessage() {
  setDownloadReadyMessage("", "");
}

/** Show the Submit Another button and set the download success message. */
function handleSuccessfulDownload() {
  state.hasDownloadedFile = true;
  submitAnotherButton.classList.remove("hidden");
  setDownloadReadyMessage("Your report has been downloaded. Check your Downloads folder.", "success");
}

/** Smooth-scroll a section element into view. */
function focusSection(element) {
  if (!element) return;
  element.scrollIntoView({ behavior: "smooth", block: "start" });
}

/** Update wizard progress step states based on current application state. */
function updateWizardProgress() {
  const hasFiles        = state.uploadedFiles.length > 0;
  const hasSummary      = state.summaryRows.length > 0;
  const hasMappingBlock = Boolean(state.mappingBlockMessage);

  wizardStepUpload.classList.toggle("is-active",   !hasFiles && !hasMappingBlock);
  wizardStepUpload.classList.toggle("is-complete",  hasFiles);

  wizardStepMapping.classList.toggle("is-active",  (hasFiles || hasMappingBlock) && !hasSummary);
  wizardStepMapping.classList.toggle("is-complete",  hasSummary);

  wizardStepSummary.classList.toggle("is-active",   hasSummary);
  wizardStepSummary.classList.toggle("is-complete",  false);
}

/** Reset only the summary section state (called when mappings change). */
function resetSummaryState() {
  state.summaryRows = [];
  state.extraCosts = {};
  state.hasDownloadedFile = false;
  summarySection.classList.add("hidden");
  clearSectionStatus("summary");
  clearDownloadReadyMessage();
  submitAnotherButton.classList.add("hidden");
}

/** Fully reset everything back to the initial upload state. */
function resetToolState() {
  state.uploadedFiles = [];
  state.combinedHeaders = [];
  state.mappings = { date: "", type: "", amount: "", fee: "", net: "", description: "" };
  state.summaryRows = [];
  state.extraCosts = {};
  state.displaySymbol = "";
  state.companyName = "";
  state.mappingLocked = false;
  state.mappingVisible = false;
  state.mappingBlockMessage = "";
  state.hasDownloadedFile = false;

  companyNameInput.value = "";
  csvInput.value = "";
  mappingGrid.innerHTML = "";
  uploadedFilesList.innerHTML = "";
  uploadedFilesWrap.classList.add("hidden");
  mappingSection.classList.add("hidden");
  processCard.classList.add("hidden");
  summarySection.classList.add("hidden");
  clearMessage();
  clearSectionStatus("upload");
  clearSectionStatus("mapping");
  clearSectionStatus("summary");
  clearDownloadReadyMessage();
  submitAnotherButton.classList.add("hidden");
  editMappingsButton.classList.add("hidden");
  if (mappingResetButton) mappingResetButton.classList.add("hidden");
  if (mappingHelper) mappingHelper.textContent = DEFAULT_MAPPING_HELPER_TEXT;

  updateUploadedFilesDisplay();
  updateProcessButtonState();
  updateWizardProgress();
  refreshMappingValidationState();
  focusSection(uploadCard);
}

/** Format a list of mismatch messages into a short readable string. */
function formatMismatchMessage(mismatches) {
  const preview   = mismatches.slice(0, 4).join(" ");
  const remaining = mismatches.length - 4;
  return remaining > 0
    ? `${preview} ${remaining} more mismatch${remaining === 1 ? "" : "es"} remaining.`
    : preview;
}

function showHeaderFormatBlock(message) {
  state.uploadedFiles = [];
  state.combinedHeaders = [];
  state.mappings = { date: "", type: "", amount: "", fee: "", net: "", description: "" };
  state.mappingLocked = false;
  state.mappingVisible = true;
  state.mappingBlockMessage = message;

  resetSummaryState();
  mappingGrid.innerHTML = "";
  uploadedFilesList.innerHTML = "";
  uploadedFilesWrap.classList.add("hidden");
  editMappingsButton.classList.add("hidden");

  updateUploadedFilesDisplay();
  updateProcessButtonState();
  updateWizardProgress();
  refreshMappingValidationState({ scrollOnBlock: true });
}

/* =================================================================
   6. Upload Handling
   ================================================================= */

/** Read, validate, and queue a single CSV file. */
async function handleFile(file) {
  if (state.uploadedFiles.length >= MAX_FILE_COUNT) {
    announce("upload", `You can add up to ${MAX_FILE_COUNT} CSV files. Remove some or generate a report first.`, "warning");
    return;
  }

  if (!file || !file.name.toLowerCase().endsWith(".csv")) {
    announce("upload", "Please choose an Etsy CSV file (ending in .csv).", "error");
    return;
  }

  clearMessage();
  clearSectionStatus("upload");
  state.mappingBlockMessage = "";

  try {
    const text   = await readCsvFile(file);
    const parsed = parseCsvText(text);

    if (!parsed.headers.length || !parsed.rows.length) {
      announce("upload", "We could not read any rows from that CSV. Please check the file and try again.", "error");
      return;
    }

    if (hasBlankHeaders(parsed.headers)) {
      showHeaderFormatBlock(
        "This CSV is missing one or more column headers. Please use the standard Etsy CSV export so the file includes proper header names.",
      );
      return;
    }

    if (isLikelyMissingHeaderRow(parsed.headers, parsed.rows)) {
      showHeaderFormatBlock(
        "This CSV does not appear to include a header row. Please use the standard Etsy CSV export so the file starts with Etsy column headers.",
      );
      return;
    }

    const fileEntry = {
      id: `${file.name}-${Date.now()}-${state.uploadedFiles.length}`,
      name: file.name,
      rows: parsed.rows,
      headers: parsed.headers,
      guessedMappings: guessMappings(parsed.headers),
      displaySymbol: detectDisplaySymbol(text),
    };

    state.uploadedFiles.push(fileEntry);
    state.combinedHeaders = combineHeaders(state.uploadedFiles);
    resetSummaryState();

    state.displaySymbol = state.displaySymbol || fileEntry.displaySymbol;
    state.mappings = buildCombinedMappings(state.mappings, state.combinedHeaders);

    updateUploadedFilesDisplay();
    updateProcessButtonState();
    updateWizardProgress();

    state.mappingLocked = false;
    const preCheck = getMappingValidationResult(state.mappings);
    state.mappingVisible = !preCheck.ok || requiresMappingReview(state.mappings);
    renderMappingFields(state.combinedHeaders);
    refreshMappingValidationState();

    focusSection(state.mappingVisible ? mappingSection : processCard);
  } catch {
    announce("upload", "We could not read that CSV. Please try again with a standard Etsy export.", "error");
  }
}

/** Handle multiple files from a drop or file-input selection. */
async function handleSelectedFiles(fileList) {
  const files = Array.from(fileList || []);
  if (!files.length) return;

  const availableSlots = Math.max(MAX_FILE_COUNT - state.uploadedFiles.length, 0);

  if (!availableSlots) {
    announce("upload", `You already have ${MAX_FILE_COUNT} CSVs selected. Remove some before adding more.`, "warning");
    return;
  }

  for (const file of files) {
    if (state.uploadedFiles.length >= MAX_FILE_COUNT) break;
    await handleFile(file);
    if (state.mappingBlockMessage) break;
  }

  if (files.length > availableSlots) {
    const skipped = files.length - availableSlots;
    announce(
      "upload",
      `Only ${availableSlots} file${availableSlots === 1 ? " was" : "s were"} added. The ${MAX_FILE_COUNT}-file limit was reached and ${skipped} file${skipped === 1 ? " was" : "s were"} skipped.`,
      "warning",
    );
  }
}

/** Update the file count pill and the uploaded files list. */
function updateUploadedFilesDisplay() {
  fileCount.textContent = `${state.uploadedFiles.length} of ${MAX_FILE_COUNT} CSVs`;

  uploadedFilesList.innerHTML = "";
  state.uploadedFiles.forEach((file, index) => {
    const item = document.createElement("li");
    item.textContent = `${index + 1}. ${file.name}`;
    uploadedFilesList.appendChild(item);
  });

  uploadedFilesWrap.classList.toggle("hidden", state.uploadedFiles.length === 0);
  uploadButton.disabled = state.uploadedFiles.length >= MAX_FILE_COUNT;
}

/** Update the process button enabled state and helper text. */
function updateProcessButtonState() {
  const hasFiles = state.uploadedFiles.length > 0;
  processButton.disabled = !hasFiles || state.mappingLocked || Boolean(state.mappingBlockMessage);
  processHelper.textContent = state.mappingBlockMessage
    ? "Upload a valid Etsy CSV with column headers to continue."
    : hasFiles
    ? "Required columns are Date, Transaction type, and Amount. Review the matches, then generate your report."
    : "Choose at least one CSV to continue.";
}

/* =================================================================
   7. CSV Parsing
   ================================================================= */

/** Read a File object as text using FileReader. */
function readCsvFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("File read failed"));
    reader.readAsText(file);
  });
}

/** Parse CSV text into { headers, rows }. Handles quoted fields and CRLF. */
function parseCsvText(text) {
  const delimiter = detectDelimiter(text);
  const rows = [];
  let currentCell = "";
  let currentRow  = [];
  let inQuotes    = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch   = text[i];
    const next = text[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        currentCell += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && ch === delimiter) {
      currentRow.push(currentCell);
      currentCell = "";
      continue;
    }

    if (!inQuotes && (ch === "\n" || ch === "\r")) {
      if (ch === "\r" && next === "\n") i += 1;
      currentRow.push(currentCell);
      currentCell = "";
      rows.push(currentRow);
      currentRow = [];
      continue;
    }

    currentCell += ch;
  }

  if (currentCell.length || currentRow.length) {
    currentRow.push(currentCell);
    rows.push(currentRow);
  }

  const cleanedRows = rows
    .map((row) => row.map((cell) => String(cell || "").replace(/^\uFEFF/, "").trim()))
    .filter((row) => row.some((cell) => cell !== ""));

  if (!cleanedRows.length) return { headers: [], rows: [] };

  const headers  = cleanedRows[0];
  const dataRows = cleanedRows.slice(1).map((row) => {
    if (row.length < headers.length) return [...row, ...new Array(headers.length - row.length).fill("")];
    return row.slice(0, headers.length);
  });

  return { headers, rows: dataRows };
}

/** Guess the delimiter by counting occurrences in the first line. */
function detectDelimiter(text) {
  const sample     = text.split(/\r?\n/).find((line) => line.trim().length) || "";
  const candidates = [",", ";", "\t", "|"];
  let best = ",";
  let bestScore = -1;

  candidates.forEach((candidate) => {
    const score = sample.split(candidate).length - 1;
    if (score > bestScore) {
      best = candidate;
      bestScore = score;
    }
  });

  return best;
}

/* =================================================================
   8. Column Mapping
   ================================================================= */

/** Guess column mappings for a set of headers using keyword scoring. */
function guessMappings(headers) {
  const result = {};
  MAPPING_FIELDS.forEach((field) => {
    result[field.key] = guessHeader(headers, field.guesses);
  });
  return result;
}

/** Merge headers from multiple files, deduplicating by normalised name. */
function combineHeaders(files) {
  const unique = [];
  const seen   = new Set();

  files.forEach((file) => {
    file.headers.forEach((header) => {
      const normalised = normaliseHeader(header);
      if (!seen.has(normalised)) {
        seen.add(normalised);
        unique.push(header);
      }
    });
  });

  return unique;
}

/** Re-guess mappings from a combined header list, keeping existing selections where valid. */
function buildCombinedMappings(currentMappings, headers) {
  const guessed = guessMappings(headers);
  const next    = {};

  MAPPING_FIELDS.forEach((field) => {
    const current = currentMappings[field.key];
    next[field.key] = findEquivalentHeader(headers, current) || guessed[field.key];
  });

  return next;
}

function requiresMappingReview(mappings) {
  return MAPPING_FIELDS.some((field) => field.reviewIfMissing && !mappings[field.key]);
}

/** Validate current mappings against uploaded files and update UI accordingly. */
function refreshMappingValidationState(options = {}) {
  const { scrollOnBlock = false } = options;
  const hasFiles = state.uploadedFiles.length > 0;
  const hasMappingBlock = Boolean(state.mappingBlockMessage);

  if (hasMappingBlock) {
    announce("mapping", state.mappingBlockMessage, "error");
    processButton.disabled = true;
    processHelper.textContent = "Upload a valid Etsy CSV with column headers to continue.";
    editMappingsButton.classList.add("hidden");
    if (mappingHelper) mappingHelper.textContent = BLOCKED_MAPPING_HELPER_TEXT;
    if (mappingLockBanner) mappingLockBanner.classList.add("hidden");
    if (mappingResetButton) mappingResetButton.classList.remove("hidden");
    mappingSection.classList.remove("hidden");
    mappingGrid.classList.add("hidden");
    processCard.classList.add("hidden");
    document.querySelectorAll("[data-mapping-key]").forEach((el) => { el.disabled = false; });
    if (scrollOnBlock) focusSection(mappingSection);
    return { ok: false, type: "blocked", message: state.mappingBlockMessage };
  }

  // No files — disable and clear
  if (!hasFiles) {
    clearSectionStatus("mapping");
    processButton.disabled = true;
    processHelper.textContent = "Choose at least one CSV to continue.";
    editMappingsButton.classList.add("hidden");
    if (mappingLockBanner) mappingLockBanner.classList.add("hidden");
    if (mappingResetButton) mappingResetButton.classList.add("hidden");
    mappingGrid.classList.add("hidden");
    processCard.classList.add("hidden");
    document.querySelectorAll("[data-mapping-key]").forEach((el) => { el.disabled = false; });
    return { ok: false, type: "empty" };
  }

  const validation    = getMappingValidationResult(state.mappings);
  const inputsLocked  = state.mappingLocked;

  document.querySelectorAll("[data-mapping-key]").forEach((el) => { el.disabled = inputsLocked; });
  const summaryActive = !summarySection.classList.contains("hidden");
  editMappingsButton.classList.toggle("hidden", !state.mappingLocked || summaryActive);
  if (mappingResetButton) mappingResetButton.classList.add("hidden");
  if (mappingHelper) mappingHelper.textContent = DEFAULT_MAPPING_HELPER_TEXT;
  // Lock banner only shown when there are column problems and locked
  if (mappingLockBanner) mappingLockBanner.classList.add("hidden");

  if (!validation.ok) {
    announce("mapping", validation.message, "error");
    processButton.disabled = true;
    processHelper.textContent = validation.type === "mismatch"
      ? "One or more columns don't match this CSV — use the dropdowns above to correct them."
      : "Select the Date, Transaction type, and Amount columns above to continue.";
    state.mappingVisible = true;
    mappingSection.classList.remove("hidden");
    mappingGrid.classList.remove("hidden");
    processCard.classList.remove("hidden");
    if (mappingLockBanner && state.mappingLocked) mappingLockBanner.classList.remove("hidden");
    if (scrollOnBlock) focusSection(mappingSection);
    return validation;
  }

  // Validation ok
  const summaryVisible = !summarySection.classList.contains("hidden");
  mappingSection.classList.toggle("hidden", !state.mappingVisible || summaryVisible);
  mappingGrid.classList.toggle("hidden", !state.mappingVisible || summaryVisible);
  processCard.classList.toggle("hidden", summaryVisible);
  setElementStatus(mappingSectionStatus, validation.message, "success");
  processButton.disabled = state.mappingLocked;
  processHelper.textContent = state.mappingLocked
    ? "Columns locked. Click Edit Columns to make changes."
    : state.mappingVisible || requiresMappingReview(state.mappings)
      ? "If your CSV uses an unusual column name, match it above or leave optional fields as Not used."
      : "";

  return validation;
}

/** Validate mappings and return { ok, type, message }. */
function getMappingValidationResult(mappings) {
  const missingRequired = MAPPING_FIELDS.filter((f) => f.required && !mappings[f.key]);

  if (missingRequired.length) {
    return {
      ok: false,
      type: "required",
      message: "Choose Date, Transaction type, and Amount before generating your report.",
    };
  }

  const mismatches = [];
  state.uploadedFiles.forEach((file) => {
    MAPPING_FIELDS.forEach((field) => {
      const selected = mappings[field.key];
      if (!selected) return;

      const fileHasHeader = Boolean(findEquivalentHeader(file.headers, selected));

      if (!fileHasHeader) {
        mismatches.push(`${file.name}: ${field.label} — no matching column found.`);
      }
    });
  });

  if (mismatches.length) {
    return { ok: false, type: "mismatch", message: formatMismatchMessage(mismatches) };
  }

  return {
    ok: true,
    type: "ok",
    message: state.mappingLocked
      ? "Required columns matched. This step is complete."
      : requiresMappingReview(mappings)
        ? "Required columns matched. One or more optional columns may need review above."
        : "Required columns matched — you can generate your report whenever you're ready.",
  };
}

/** Render the mapping select dropdowns into the mappingGrid element. */
function renderMappingFields(headers) {
  mappingGrid.innerHTML = "";

  MAPPING_FIELDS.forEach((field) => {
    const wrapper = document.createElement("div");
    wrapper.className = "field";

    const label = document.createElement("label");
    label.setAttribute("for", `mapping-${field.key}`);
    label.textContent = field.required ? `${field.label} *` : field.label;

    const note = document.createElement("small");
    note.textContent = field.help;

    const select = document.createElement("select");
    select.id = `mapping-${field.key}`;
    select.dataset.mappingKey = field.key;
    select.disabled = state.mappingLocked;

    const blankOption = document.createElement("option");
    blankOption.value = "";
    blankOption.textContent = field.required ? "Please choose a column" : "Not used";
    select.appendChild(blankOption);

    headers.forEach((header) => {
      const option = document.createElement("option");
      option.value = header;
      option.textContent = header;
      if (state.mappings[field.key] === header) option.selected = true;
      select.appendChild(option);
    });

    select.addEventListener("change", () => {
      state.mappingLocked = false;
      updateMappingsFromInputs();
      resetSummaryState();
      refreshMappingValidationState();
    });

    wrapper.appendChild(label);
    wrapper.appendChild(note);
    wrapper.appendChild(select);
    mappingGrid.appendChild(wrapper);
  });
}

/** Sync the state.mappings object from current select values in the DOM. */
function updateMappingsFromInputs() {
  document.querySelectorAll("[data-mapping-key]").forEach((el) => {
    state.mappings[el.dataset.mappingKey] = el.value;
  });
}

/** Resolve mappings for a single file, falling back to per-file guesses if needed. */
function resolveMappingsForFile(file, mappings) {
  const resolved = {};
  MAPPING_FIELDS.forEach((field) => {
    const selected = mappings[field.key];
    const matchedHeader = findEquivalentHeader(file.headers, selected);
    resolved[field.key] = matchedHeader
      ? matchedHeader
      : file.guessedMappings[field.key] || "";
  });
  return resolved;
}

/** Best-effort header match against a list of guess strings, scored by similarity. */
function guessHeader(headers, guesses) {
  let bestMatch = "";
  let bestScore = 0;

  headers.forEach((header) => {
    const normalised = normaliseHeader(header);
    guesses.forEach((guess) => {
      let score = 0;
      if (normalised === guess)                                           score = 100;
      else if (normalised.startsWith(guess) || normalised.endsWith(guess)) score = 80;
      else if (normalised.includes(guess))                                score = 70;

      if (score > bestScore) {
        bestScore = score;
        bestMatch = header;
      }
    });
  });

  return bestMatch;
}

/** Lowercase, collapse whitespace, and replace separators for consistent comparison. */
function normaliseHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function hasBlankHeaders(headers) {
  return headers.some((header) => !String(header || "").trim());
}

function hasRecognisableHeaderLabels(headers) {
  return MAPPING_FIELDS
    .filter((field) => field.required)
    .some((field) => Boolean(guessHeader(headers, field.guesses)));
}

function isLikelyMissingHeaderRow(headers, rows) {
  if (hasRecognisableHeaderLabels(headers)) return false;

  if (countDataLikeCells(headers) >= 2) return true;

  const sampleRows = rows.filter((row) => row.some((cell) => String(cell || "").trim())).slice(0, 3);
  return sampleRows.some((row) => countNonTextShapeMatches(headers, row) >= 2);
}

function countDataLikeCells(row) {
  return row.reduce((count, cell) => {
    const kind = inferCellKind(cell);
    return kind === "date" || kind === "number" || kind === "transaction"
      ? count + 1
      : count;
  }, 0);
}

function countNonTextShapeMatches(firstRow, secondRow) {
  const length = Math.min(firstRow.length, secondRow.length);
  let matches = 0;

  for (let index = 0; index < length; index += 1) {
    const firstKind = inferCellKind(firstRow[index]);
    const secondKind = inferCellKind(secondRow[index]);
    if (firstKind === secondKind && firstKind !== "text" && firstKind !== "empty") {
      matches += 1;
    }
  }

  return matches;
}

function inferCellKind(value) {
  const text = String(value || "").trim();
  if (!text) return "empty";
  if (detectDateValue(text)) return "date";

  const amount = parseCurrencyValue(text);
  if (amount.valid && amount.value !== null) return "number";

  if ([
    "sale", "payment", "order", "fee", "refund", "adjustment",
    "deposit", "payout", "transfer", "disbursement", "reserve", "chargeback",
  ].includes(normaliseHeader(text))) {
    return "transaction";
  }

  return "text";
}

function findEquivalentHeader(headers, selectedHeader) {
  if (!selectedHeader) return "";

  const target = normaliseHeader(selectedHeader);
  return headers.find((header) => normaliseHeader(header) === target) || "";
}

/* =================================================================
   9. Monthly Summary Calculation
   ================================================================= */

/** Build a combined monthly summary across all uploaded files. */
function buildCombinedMonthlySummary(files, mappings) {
  const months     = new Map();
  const skipped    = [];

  for (const file of files) {
    const resolved = resolveMappingsForFile(file, mappings);
    const missing  = MAPPING_FIELDS.filter((f) => f.required && !resolved[f.key]);

    if (missing.length) {
      return { ok: false, message: `Required columns are missing for "${file.name}". Please check the column matches.` };
    }

    const result = buildMonthlySummary(file.rows, file.headers, resolved);

    if (!result.ok) return { ok: false, message: `${file.name}: ${result.message}` };

    if (result.skippedBecauseIgnoredOnly) {
      skipped.push(file.name);
      continue;
    }

    result.summaryRows.forEach((row) => {
      const bucket = months.get(row.month) || createMonthBucket(row.month);
      months.set(row.month, mergeMonthlyBuckets(bucket, row));
    });
  }

  const summaryRows = Array.from(months.values()).sort((a, b) => a.month.localeCompare(b.month));

  if (!summaryRows.length) {
    return {
      ok: false,
      message: "These files only contained deposits, payouts, or other rows that are excluded from the monthly totals.",
    };
  }

  return { ok: true, summaryRows, skippedFiles: skipped };
}

/** Build a month-by-month summary for a single file. */
function buildMonthlySummary(rows, headers, mappings) {
  const idx = {
    date:        headers.indexOf(mappings.date),
    type:        headers.indexOf(mappings.type),
    amount:      headers.indexOf(mappings.amount),
    fee:         mappings.fee         ? headers.indexOf(mappings.fee)         : -1,
    net:         mappings.net         ? headers.indexOf(mappings.net)         : -1,
    description: mappings.description ? headers.indexOf(mappings.description) : -1,
  };

  if (idx.date < 0 || idx.type < 0 || idx.amount < 0) {
    return { ok: false, message: "Some required columns are missing. Please check date, transaction type, and amount." };
  }

  const months           = new Map();
  let   validRows        = 0;
  let   invalidDateCount = 0;
  let   invalidAmtCount  = 0;

  rows.forEach((row) => {
    if (!row.some((cell) => String(cell || "").trim())) return;

    const monthKey = detectMonthKey(row[idx.date]);
    if (!monthKey) { invalidDateCount += 1; return; }

    const amountVal = parseCurrencyValue(row[idx.amount]);
    const feeVal    = idx.fee >= 0 ? parseCurrencyValue(row[idx.fee]) : { valid: true, value: null };
    const netVal    = idx.net >= 0 ? parseCurrencyValue(row[idx.net]) : { valid: true, value: null };

    if (!amountVal.valid || !feeVal.valid || !netVal.valid) { invalidAmtCount += 1; return; }

    let primary = selectPrimaryValue(amountVal.value, feeVal.value, netVal.value);

    // For deposit/payout rows with a blank amount, try to extract from description text
    if (primary.value === null) {
      const descText = idx.description >= 0 ? row[idx.description] : "";
      const fallback = extractTransferAmount(row[idx.type], descText);
      if (fallback !== null) primary = { source: "derived", value: fallback };
    }

    if (primary.value === null) return;

    const typeText = String(row[idx.type] || "").trim().toLowerCase();
    const descText = idx.description >= 0 ? String(row[idx.description] || "").trim().toLowerCase() : "";
    const category = classifyTransaction(typeText, descText, primary.value);

    if (category === "ignore") return;

    const bucket = months.get(monthKey) || createMonthBucket(monthKey);
    const fee    = feeVal.value ?? 0;
    const net    = netVal.value;

    if      (category === "revenue")    bucket.revenue    += Math.max(primary.value, 0);
    else if (category === "refund")     bucket.refunds    += primary.value;
    else if (category === "fee")        bucket.fees       += primary.value;
    else                                bucket.adjustments += primary.value;

    if (primary.source === "amount" && fee) bucket.fees += fee;

    bucket.netReceived += calculateNetContribution(primary.source, primary.value, feeVal.value, net);
    validRows += 1;
    months.set(monthKey, bucket);
  });

  if (!validRows && invalidDateCount) {
    return { ok: false, message: "We could not read the date format. Please check that the correct date column is selected." };
  }
  if (invalidAmtCount > 0) {
    return { ok: false, message: "Some amount values could not be read as numbers. Please check the amount, fee, and net column matches." };
  }
  if (!validRows) {
    return { ok: true, summaryRows: [], skippedBecauseIgnoredOnly: true };
  }

  const summaryRows = Array.from(months.values()).sort((a, b) => a.month.localeCompare(b.month));
  return { ok: true, summaryRows, skippedBecauseIgnoredOnly: false };
}

function createMonthBucket(month) {
  return { month, revenue: 0, fees: 0, refunds: 0, adjustments: 0, netReceived: 0 };
}

function mergeMonthlyBuckets(existing, next) {
  existing.revenue     += next.revenue;
  existing.fees        += next.fees;
  existing.refunds     += next.refunds;
  existing.adjustments += next.adjustments;
  existing.netReceived += next.netReceived;
  return existing;
}

function selectPrimaryValue(amountValue, feeValue, netValue) {
  if (amountValue !== null) return { source: "amount", value: amountValue };
  if (feeValue    !== null) return { source: "fee",    value: feeValue    };
  if (netValue    !== null) return { source: "net",    value: netValue    };
  return { source: "none", value: null };
}

function calculateNetContribution(primarySource, primaryValue, feeValue, netValue) {
  if (netValue !== null) return netValue;
  if (primarySource === "amount") return primaryValue + (feeValue ?? 0);
  if (primaryValue  !== null)     return primaryValue;
  return 0;
}

/** Classify a transaction row into revenue, refund, fee, adjustment, or ignore. */
function classifyTransaction(type, description, amount) {
  const combined = `${type} ${description}`.trim();

  if (containsAny(type, ["deposit", "disbursement", "payout", "transfer"])) return "adjustment";
  if (containsAny(type, ["refund", "reversal", "chargeback"]))              return "refund";
  if (containsAny(type, ["sale", "payment", "order"]))                      return amount < 0 ? "refund" : "revenue";
  if (containsAny(type, ["fee", "vat", "tax", "marketing"]))                return "fee";
  if (containsAny(type, ["adjustment", "reserve"]))                         return "adjustment";

  if (containsAny(combined, ["refund", "reversal", "chargeback", "cancel", "cancelled", "returned"])) return "refund";
  if (containsAny(combined, ["fee", "listing", "regulatory", "processing", "offsite ads", "shipping label", "marketing"])) return "fee";
  if (containsAny(combined, ["adjustment", "reserve", "disbursement", "deposit", "transfer", "tax", "vat", "balance", "payout"])) return "adjustment";
  if (containsAny(combined, ["sale", "order", "payment", "credit", "shipping", "receipt"])) return amount < 0 ? "refund" : "revenue";

  return amount < 0 ? "adjustment" : "revenue";
}

function containsAny(text, words) {
  return words.some((word) => text.includes(word));
}

/** Try to extract a currency amount embedded in a deposit/payout description. */
function extractTransferAmount(typeText, descriptionText) {
  const type = String(typeText || "").trim().toLowerCase();
  if (!containsAny(type, ["deposit", "disbursement", "payout", "transfer"])) return null;

  const match = String(descriptionText || "").match(/[£$€]\s*\d[\d,.]*/);
  if (!match) return null;

  const parsed = parseCurrencyValue(match[0]);
  return parsed.valid ? parsed.value : null;
}

/* --- Date parsing ------------------------------------------------- */

function detectMonthKey(value) {
  const date = detectDateValue(value);
  if (!date) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function detectDateValue(value) {
  const text = String(value || "").trim();
  if (!text) return null;

  // "15 Jan, 2025" or "15 January 2025"
  let m = text.match(/^(\d{1,2})\s+([a-zA-Z]+),?\s+(\d{4})/);
  if (m) {
    const month = getMonthNumberFromName(m[2]);
    if (month) return safeDate(Number(m[3]), month, Number(m[1]));
  }

  // ISO: 2025-01-15
  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return safeDate(Number(m[1]), Number(m[2]), Number(m[3]));

  // 2025/01/15 or 2025.01.15
  m = text.match(/^(\d{4})[/.](\d{1,2})[/.](\d{1,2})/);
  if (m) return safeDate(Number(m[1]), Number(m[2]), Number(m[3]));

  // DD/MM/YYYY or MM/DD/YYYY (ambiguous — favour day-first when day > 12)
  m = text.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})/);
  if (m) {
    const first  = Number(m[1]);
    const second = Number(m[2]);
    const year   = normaliseYear(m[3]);
    return first > 12 && second <= 12 ? safeDate(year, second, first) : safeDate(year, first, second);
  }

  // Fall back to native Date parsing
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function normaliseYear(value) {
  const num = Number(value);
  return value.length === 2 ? (num >= 70 ? 1900 + num : 2000 + num) : num;
}

function getMonthNumberFromName(value) {
  const map = {
    january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
    july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
  };
  return map[String(value || "").trim().toLowerCase()] || 0;
}

function safeDate(year, month, day) {
  const d = new Date(year, month - 1, day);
  if (
    Number.isNaN(d.getTime()) ||
    d.getFullYear() !== year  ||
    d.getMonth()   !== month - 1 ||
    d.getDate()    !== day
  ) return null;
  return d;
}

/* --- Currency parsing -------------------------------------------- */

/** Parse a currency string into { valid, value }. Returns null for blank cells and 0 for Etsy dash placeholders. */
function parseCurrencyValue(value) {
  const text = String(value || "").trim();
  if (!text) return { valid: true, value: null };
  if (/^(?:[-–—]+|n\/a)$/i.test(text)) return { valid: true, value: 0 };

  let cleaned  = text.replace(/\u00A0/g, "").replace(/\s+/g, "");
  let negative = false;

  if (cleaned.startsWith("(") && cleaned.endsWith(")")) { negative = true; cleaned = cleaned.slice(1, -1); }
  if (cleaned.endsWith("-"))   { negative = true; cleaned = cleaned.slice(0, -1); }
  if (cleaned.startsWith("-")) { negative = true; cleaned = cleaned.slice(1); }

  cleaned = cleaned.replace(/[£$€]/g, "");

  const commaCount = (cleaned.match(/,/g) || []).length;
  const dotCount   = (cleaned.match(/\./g) || []).length;

  if (commaCount && dotCount) {
    cleaned = cleaned.lastIndexOf(",") > cleaned.lastIndexOf(".")
      ? cleaned.replace(/\./g, "").replace(/,/g, ".")
      : cleaned.replace(/,/g, "");
  } else if (commaCount && !dotCount) {
    const lastComma = cleaned.lastIndexOf(",");
    const decimals  = cleaned.length - lastComma - 1;
    cleaned = decimals <= 2 ? cleaned.replace(/,/g, ".") : cleaned.replace(/,/g, "");
  } else {
    cleaned = cleaned.replace(/,/g, "");
  }

  if (!/^\d*\.?\d+$/.test(cleaned)) return { valid: false, value: null };

  const parsed = Number(cleaned);
  return Number.isNaN(parsed) ? { valid: false, value: null } : { valid: true, value: negative ? -parsed : parsed };
}

/** Detect the most frequently occurring currency symbol in a CSV file. */
function detectDisplaySymbol(text) {
  const matches = text.match(/[£$€]/g) || [];
  if (!matches.length) return "";

  const counts = matches.reduce((acc, sym) => {
    acc[sym] = (acc[sym] || 0) + 1;
    return acc;
  }, {});

  return Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
}

/* =================================================================
   10. Table Rendering
   ================================================================= */

/** Ensure every summary month has an extraCosts bucket (and remove stale ones). */
function ensureExtraCostBuckets() {
  const valid = new Set(state.summaryRows.map((r) => r.month));
  Object.keys(state.extraCosts).forEach((m) => { if (!valid.has(m)) delete state.extraCosts[m]; });
  state.summaryRows.forEach((row) => {
    if (!state.extraCosts[row.month]) {
      state.extraCosts[row.month] = { packaging: 0, shippingSupplies: 0, ads: 0, subscriptions: 0, other: 0 };
    }
  });
}

/** Render the extra-costs input table. */
function renderCostTable() {
  costTableBody.innerHTML = "";

  const COST_FIELDS = [
    ["packaging",        "Packaging"],
    ["shippingSupplies", "Shipping Supplies"],
    ["ads",              "Ads"],
    ["subscriptions",    "Subscriptions"],
    ["other",            "Other"],
  ];

  state.summaryRows.forEach((row) => {
    const tr = document.createElement("tr");

    const monthTd = document.createElement("td");
    monthTd.className = "month-cell";
    monthTd.textContent = row.month;
    tr.appendChild(monthTd);

    COST_FIELDS.forEach(([key, label]) => {
      const td    = document.createElement("td");
      const input = document.createElement("input");
      input.type  = "number";
      input.min   = "0";
      input.step  = "0.01";
      input.value = state.extraCosts[row.month][key] || 0;
      input.dataset.month   = row.month;
      input.dataset.costKey = key;
      input.setAttribute("aria-label", `${row.month} ${label}`);
      input.addEventListener("input", handleCostInput);
      td.appendChild(input);
      tr.appendChild(td);
    });

    costTableBody.appendChild(tr);
  });
}

function handleCostInput(event) {
  const { month, costKey } = event.target.dataset;
  const value = Number(event.target.value || 0);
  state.extraCosts[month][costKey] = Number.isFinite(value) ? value : 0;
  renderSummaryTable();
}

function appendSummaryCell(rowElement, text, className = "") {
  const cell = document.createElement("td");
  if (className) cell.className = className;
  cell.textContent = text;
  rowElement.appendChild(cell);
}

function appendSummaryRow(container, row) {
  const tr = document.createElement("tr");
  appendSummaryCell(tr, row.month, "month-cell");
  appendSummaryCell(tr, formatCurrency(row.revenue));
  appendSummaryCell(tr, formatCurrency(row.fees));
  appendSummaryCell(tr, formatCurrency(row.refunds));
  appendSummaryCell(tr, formatCurrency(row.adjustments));
  appendSummaryCell(tr, formatCurrency(row.netReceived));
  appendSummaryCell(tr, formatCurrency(row.extraCosts));
  appendSummaryCell(tr, formatCurrency(row.estimatedProfit));
  container.appendChild(tr);
}

/** Render the monthly accounts summary table with current extra cost figures. */
function renderSummaryTable() {
  summaryTableBody.replaceChildren();
  summaryTableFoot.replaceChildren();

  const rows = getSummaryWithCosts();

  rows.forEach((row) => {
    appendSummaryRow(summaryTableBody, row);
  });

  if (!rows.length) return;

  const totals  = calculateSummaryTotals(rows);
  appendSummaryRow(summaryTableFoot, totals);
}

function getSummaryWithCosts() {
  return state.summaryRows.map((row) => {
    const extra = sumValues(state.extraCosts[row.month] || {});
    return { ...row, extraCosts: extra, estimatedProfit: row.netReceived - extra };
  });
}

function getExportRows() {
  const rows = getSummaryWithCosts();
  return rows.length ? [...rows, calculateSummaryTotals(rows)] : [];
}

function calculateSummaryTotals(rows) {
  return rows.reduce(
    (totals, row) => {
      totals.revenue          += Number(row.revenue          || 0);
      totals.fees             += Number(row.fees             || 0);
      totals.refunds          += Number(row.refunds          || 0);
      totals.adjustments      += Number(row.adjustments      || 0);
      totals.netReceived      += Number(row.netReceived      || 0);
      totals.extraCosts       += Number(row.extraCosts       || 0);
      totals.estimatedProfit  += Number(row.estimatedProfit  || 0);
      return totals;
    },
    { month: "Total", revenue: 0, fees: 0, refunds: 0, adjustments: 0, netReceived: 0, extraCosts: 0, estimatedProfit: 0 },
  );
}

function formatCurrency(value) {
  const num    = Number(value || 0);
  const fixed  = Math.abs(num).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return `${num < 0 ? "-" : ""}${state.displaySymbol}${fixed}`;
}

function sumValues(obj) {
  return Object.values(obj).reduce((total, v) => total + Number(v || 0), 0);
}

/* =================================================================
   11. Export — CSV, XLS, XLSX
   ================================================================= */

function convertSummaryToCsv(rows) {
  const headers = EXPORT_COLUMNS.map((col) => col.label);
  const lines   = [headers.join(",")];

  rows.forEach((row) => {
    lines.push(
      EXPORT_COLUMNS
        .map((col) => col.type === "currency" ? Number(row[col.key] || 0).toFixed(2) : row[col.key])
        .map(escapeCsvValue)
        .join(","),
    );
  });

  return lines.join("\r\n");
}

/** Build a structured list of display rows for XLS/XLSX output. */
function buildWorkbookDisplayRows(rows) {
  const bodyRows   = rows.filter((r) => r.month !== "Total");
  const totalRow   = rows.find((r) => r.month === "Total") || null;
  const title      = state.companyName
    ? `${state.companyName} • Etsy Monthly Accounts Summary`
    : "Etsy Monthly Accounts Summary";
  const metaLabel  = `Prepared ${new Date().toISOString().slice(0, 10)} • ${state.uploadedFiles.length || 1} file${state.uploadedFiles.length === 1 ? "" : "s"}`;

  const emptyCells = () => EXPORT_COLUMNS.map(() => ({ type: "text", value: "" }));
  const toCells    = (row) => EXPORT_COLUMNS.map((col) => ({
    type:  col.type,
    value: col.type === "currency" ? Number(row[col.key] || 0) : String(row[col.key] || ""),
  }));

  const displayRows = [];

  const titleCells = emptyCells();
  titleCells[0] = { type: "text", value: title };
  displayRows.push({ kind: "title",  height: 28, cells: titleCells });

  const metaCells = emptyCells();
  metaCells[0] = { type: "text", value: metaLabel };
  displayRows.push({ kind: "meta",   height: 20, cells: metaCells });

  displayRows.push({ kind: "spacer", height: 10, cells: emptyCells() });
  displayRows.push({
    kind: "header",
    height: 22,
    cells: EXPORT_COLUMNS.map((col) => ({ type: "text", value: col.label })),
  });

  bodyRows.forEach((row, i) => {
    displayRows.push({ kind: i % 2 === 0 ? "data" : "dataAlt", height: 20, cells: toCells(row) });
  });

  if (totalRow) {
    displayRows.push({ kind: "spacer", height: 10, cells: emptyCells() });
    displayRows.push({ kind: "total",  height: 22, cells: toCells(totalRow) });
  }

  return {
    displayRows,
    headerRowNumber:   4,
    firstDataRowNumber: 5,
    dataRowCount:      bodyRows.length,
    lastColumnName:    columnNumberToName(EXPORT_COLUMNS.length),
    usedRangeEndRow:   displayRows.length,
  };
}

/* --- XLS (SpreadsheetML) ----------------------------------------- */

function getXlsStyleId(rowKind, cellType) {
  if (rowKind === "title")   return "sTitle";
  if (rowKind === "meta")    return "sMeta";
  if (rowKind === "spacer")  return "sSpacer";
  if (rowKind === "header")  return "sHeader";
  if (rowKind === "total")   return cellType === "currency" ? "sTotalCurrency" : "sTotalLabel";
  if (rowKind === "dataAlt") return cellType === "currency" ? "sCurrencyAlt"   : "sTextAlt";
  return cellType === "currency" ? "sCurrency" : "sText";
}

function createXlsWorkbookXml(rows) {
  const symbol    = xmlEscape(state.displaySymbol || "");
  const dateStamp = xmlEscape(new Date().toISOString().slice(0, 10));
  const layout    = buildWorkbookDisplayRows(rows);
  const fmtStr    = xmlEscape(`${state.displaySymbol || ""}#,##0.00;[Red](${state.displaySymbol || ""}#,##0.00)`);

  const bodyRows = layout.displayRows.map((row) => {
    const cells = row.cells.map((cell) => {
      const styleId = getXlsStyleId(row.kind, cell.type);
      if (cell.type === "currency") {
        return `<Cell ss:StyleID="${styleId}"><Data ss:Type="Number">${Number(cell.value || 0).toFixed(2)}</Data></Cell>`;
      }
      return `<Cell ss:StyleID="${styleId}"><Data ss:Type="String">${xmlEscape(cell.value || "")}</Data></Cell>`;
    }).join("");
    return `<Row ss:Height="${row.height}">${cells}</Row>`;
  }).join("");

  const colDefs = EXPORT_COLUMNS.map((col) => `<Column ss:AutoFitWidth="0" ss:Width="${col.width * 7.4}"/>`).join("");

  return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>SimpleBizToolkit.com</Author>
  <Created>${dateStamp}T00:00:00Z</Created>
 </DocumentProperties>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center"/>
   <Font ss:FontName="Calibri" ss:Size="11" ss:Color="#2B2A28"/>
  </Style>
  <Style ss:ID="sTitle">
   <Font ss:FontName="Calibri" ss:Size="15" ss:Bold="1" ss:Color="#1A4B3E"/>
   <Interior ss:Color="#F4EFE7" ss:Pattern="Solid"/>
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
  </Style>
  <Style ss:ID="sMeta">
   <Font ss:FontName="Calibri" ss:Size="10" ss:Italic="1" ss:Color="#6D665E"/>
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
  </Style>
  <Style ss:ID="sSpacer">
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="sHeader">
   <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#1F6A52" ss:Pattern="Solid"/>
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D0C0"/>
    <Border ss:Position="Left"   ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D0C0"/>
    <Border ss:Position="Right"  ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D0C0"/>
    <Border ss:Position="Top"    ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9D0C0"/>
   </Borders>
  </Style>
  <Style ss:ID="sText">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Left"   ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Right"  ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Top"    ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
   </Borders>
  </Style>
  <Style ss:ID="sTextAlt">
   <Interior ss:Color="#FAF7F2" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Left"   ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Right"  ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Top"    ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
   </Borders>
  </Style>
  <Style ss:ID="sCurrency">
   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
   <NumberFormat ss:Format="${fmtStr}"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Left"   ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Right"  ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Top"    ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
   </Borders>
  </Style>
  <Style ss:ID="sCurrencyAlt">
   <Interior ss:Color="#FAF7F2" ss:Pattern="Solid"/>
   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
   <NumberFormat ss:Format="${fmtStr}"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Left"   ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Right"  ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
    <Border ss:Position="Top"    ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E7DED0"/>
   </Borders>
  </Style>
  <Style ss:ID="sTotalLabel">
   <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#1A4B3E"/>
   <Interior ss:Color="#EAF3EE" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D8D0"/>
    <Border ss:Position="Left"   ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D8D0"/>
    <Border ss:Position="Right"  ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D8D0"/>
    <Border ss:Position="Top"    ss:LineStyle="Continuous" ss:Weight="2" ss:Color="#1F6A52"/>
   </Borders>
  </Style>
  <Style ss:ID="sTotalCurrency">
   <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#1A4B3E"/>
   <Interior ss:Color="#EAF3EE" ss:Pattern="Solid"/>
   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>
   <NumberFormat ss:Format="${fmtStr}"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D8D0"/>
    <Border ss:Position="Left"   ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D8D0"/>
    <Border ss:Position="Right"  ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#C8D8D0"/>
    <Border ss:Position="Top"    ss:LineStyle="Continuous" ss:Weight="2" ss:Color="#1F6A52"/>
   </Borders>
  </Style>
 </Styles>
 <Worksheet ss:Name="Monthly Accounts">
  <Table>
   ${colDefs}
   ${bodyRows}
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Selected/>
   <FreezePanes/>
   <FrozenNoSplit/>
   <SplitHorizontal>${layout.headerRowNumber}</SplitHorizontal>
   <TopRowBottomPane>${layout.headerRowNumber}</TopRowBottomPane>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>`;
}

/* --- XLSX (OpenXML ZIP) ------------------------------------------ */

function getXlsxStyleId(rowKind, cellType) {
  if (rowKind === "title")   return 3;
  if (rowKind === "meta")    return 4;
  if (rowKind === "spacer")  return 9;
  if (rowKind === "header")  return 2;
  if (rowKind === "total")   return cellType === "currency" ? 8 : 7;
  if (rowKind === "dataAlt") return cellType === "currency" ? 6 : 5;
  return cellType === "currency" ? 1 : 0;
}

function createXlsxWorkbook(rows) {
  return createZipBlob([
    { name: "[Content_Types].xml",          content: createXlsxContentTypesXml() },
    { name: "_rels/.rels",                  content: createXlsxRootRelsXml()     },
    { name: "docProps/app.xml",             content: createXlsxAppXml()          },
    { name: "docProps/core.xml",            content: createXlsxCoreXml()         },
    { name: "xl/workbook.xml",              content: createXlsxWorkbookXml()     },
    { name: "xl/_rels/workbook.xml.rels",   content: createXlsxWorkbookRelsXml() },
    { name: "xl/styles.xml",               content: createXlsxStylesXml()       },
    { name: "xl/worksheets/sheet1.xml",    content: createXlsxSheetXml(rows)    },
  ]);
}

function createXlsxContentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml"          ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml"         ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/styles.xml"             ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/workbook.xml"           ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
}

function createXlsxRootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"           Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"        Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"      Target="docProps/app.xml"/>
</Relationships>`;
}

function createXlsxAppXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>SimpleBizToolkit.com</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr"><vt:lpstr>Monthly Accounts</vt:lpstr></vt:vector>
  </TitlesOfParts>
</Properties>`;
}

function createXlsxCoreXml() {
  const ts = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>SimpleBizToolkit.com</dc:creator>
  <cp:lastModifiedBy>SimpleBizToolkit.com</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${ts}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${ts}</dcterms:modified>
  <dc:title>Etsy CSV Monthly Accounts Cleaner</dc:title>
</cp:coreProperties>`;
}

function createXlsxWorkbookXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>
  <sheets><sheet name="Monthly Accounts" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
}

function createXlsxWorkbookRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"    Target="styles.xml"/>
</Relationships>`;
}

function createXlsxStylesXml() {
  const fmt = xmlEscape(`${state.displaySymbol || ""}#,##0.00;[Red](${state.displaySymbol || ""}#,##0.00)`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="${fmt}"/></numFmts>
  <fonts count="5">
    <font><sz val="11"/><color rgb="FF2B2A28"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="15"/><color rgb="FF1A4B3E"/><name val="Calibri"/><family val="2"/></font>
    <font><i/><sz val="10"/><color rgb="FF6D665E"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF1A4B3E"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F6A52"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFAF7F2"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEAF3EE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4EFE7"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFE7DED0"/></left><right style="thin"><color rgb="FFE7DED0"/></right>
      <top style="thin"><color rgb="FFE7DED0"/></top><bottom style="thin"><color rgb="FFE7DED0"/></bottom><diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFC8D8D0"/></left><right style="thin"><color rgb="FFC8D8D0"/></right>
      <top style="medium"><color rgb="FF1F6A52"/></top><bottom style="thin"><color rgb="FFC8D8D0"/></bottom><diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="2" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="4" fillId="4" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
}

function createXlsxSheetXml(rows) {
  const layout = buildWorkbookDisplayRows(rows);

  const sheetRows = layout.displayRows.map((row, rowIndex) => {
    const rowNum = rowIndex + 1;
    const cells  = row.cells.map((cell, colIndex) => {
      const ref     = `${columnNumberToName(colIndex + 1)}${rowNum}`;
      const styleId = getXlsxStyleId(row.kind, cell.type);
      if (cell.type === "currency") {
        return `<c r="${ref}" s="${styleId}"><v>${Number(cell.value || 0).toFixed(2)}</v></c>`;
      }
      return `<c r="${ref}" t="inlineStr" s="${styleId}"><is><t>${xmlEscape(cell.value || "")}</t></is></c>`;
    }).join("");
    return `<row r="${rowNum}" ht="${row.height}" customHeight="1">${cells}</row>`;
  });

  const colDefs = EXPORT_COLUMNS.map((col, i) =>
    `<col min="${i + 1}" max="${i + 1}" width="${col.width + 1.5}" customWidth="1"/>`,
  ).join("");

  const lastCell = `${layout.lastColumnName}${Math.max(layout.usedRangeEndRow, 1)}`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:${lastCell}"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="${layout.headerRowNumber}" topLeftCell="A${layout.firstDataRowNumber}" activePane="bottomLeft" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>${colDefs}</cols>
  <sheetData>${sheetRows.join("")}</sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;
}

/* --- ZIP builder -------------------------------------------------- */

function createZipBlob(files) {
  const encoder     = new TextEncoder();
  const localParts  = [];
  const centralParts = [];
  let   offset      = 0;

  files.forEach((file) => {
    const nameBytes = encoder.encode(file.name);
    const dataBytes = encoder.encode(file.content);
    const crc       = crc32(dataBytes);

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const lv          = new DataView(localHeader.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0,  true);
    lv.setUint16(8, 0,  true);
    lv.setUint16(10, 0, true);
    lv.setUint16(12, 0, true);
    lv.setUint32(14, crc,              true);
    lv.setUint32(18, dataBytes.length, true);
    lv.setUint32(22, dataBytes.length, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0,                true);
    localHeader.set(nameBytes, 30);

    localParts.push(localHeader, dataBytes);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const cv            = new DataView(centralHeader.buffer);
    cv.setUint32(0,  0x02014b50,    true);
    cv.setUint16(4,  20,            true);
    cv.setUint16(6,  20,            true);
    cv.setUint16(8,  0,             true);
    cv.setUint16(10, 0,             true);
    cv.setUint16(12, 0,             true);
    cv.setUint16(14, 0,             true);
    cv.setUint32(16, crc,           true);
    cv.setUint32(20, dataBytes.length, true);
    cv.setUint32(24, dataBytes.length, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0,             true);
    cv.setUint16(32, 0,             true);
    cv.setUint16(34, 0,             true);
    cv.setUint16(36, 0,             true);
    cv.setUint32(38, 0,             true);
    cv.setUint32(42, offset,        true);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    offset += localHeader.length + dataBytes.length;
  });

  const centralSize = centralParts.reduce((total, p) => total + p.length, 0);
  const endRecord   = new Uint8Array(22);
  const ev          = new DataView(endRecord.buffer);
  ev.setUint32(0,  0x06054b50,      true);
  ev.setUint16(4,  0,               true);
  ev.setUint16(6,  0,               true);
  ev.setUint16(8,  files.length,    true);
  ev.setUint16(10, files.length,    true);
  ev.setUint32(12, centralSize,     true);
  ev.setUint32(16, offset,          true);
  ev.setUint16(20, 0,               true);

  return new Blob([...localParts, ...centralParts, endRecord], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

/* =================================================================
   12. Download
   ================================================================= */

/** Trigger a file download in the browser. Accepts string or Blob content. */
function downloadFile(content, name, mimeType) {
  const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });
  const url  = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href  = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/** Build the download filename from the uploaded file name(s). */
function buildDownloadName(extension) {
  const base = state.uploadedFiles.length === 1
    ? state.uploadedFiles[0].name.replace(/\.csv$/i, "")
    : "etsy-multi-month-accounts";
  return `${base}-clean-monthly-accounts.${extension}`;
}

/* =================================================================
   13. Utilities
   ================================================================= */

/** Convert a 1-based column index to a spreadsheet column name (A, B, … Z, AA …). */
function columnNumberToName(value) {
  let result = "";
  let number = value;
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

/** Escape a string for safe use inside XML. */
function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** Escape a value for safe use as a CSV field. */
function escapeCsvValue(value) {
  const text = String(value ?? "");
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

/* --- CRC-32 (required for ZIP) ----------------------------------- */

const CRC32_TABLE = (() => {
  const table = [];
  for (let i = 0; i < 256; i += 1) {
    let crc = i;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }
    table.push(crc >>> 0);
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0 ^ -1;
  for (let i = 0; i < bytes.length; i += 1) {
    crc = (crc >>> 8) ^ CRC32_TABLE[(crc ^ bytes[i]) & 0xff];
  }
  return (crc ^ -1) >>> 0;
}

/* =================================================================
   14. Initialise
   ================================================================= */

// Wipe in-memory CSV data when the user navigates away or closes the tab.
// pagehide fires reliably on page unload; visibilitychange covers tab switches
// that precede a close, ensuring the GC can reclaim data promptly.
window.addEventListener("pagehide", () => {
  state.uploadedFiles = [];
  state.combinedHeaders = [];
  state.summaryRows = [];
  state.extraCosts = {};
});

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") {
    state.uploadedFiles = [];
    state.combinedHeaders = [];
    state.summaryRows = [];
    state.extraCosts = {};
  }
});

updateUploadedFilesDisplay();
updateProcessButtonState();
updateWizardProgress();
refreshMappingValidationState();
