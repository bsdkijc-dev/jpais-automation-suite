/**
 * Validates the selected Production Board row.
 */
function validateActiveProductionRow() {
  try {
    const context = getActiveProductionContext_();
    const validation = validateProductionContext_(context);

if (!validation.valid) {
  SpreadsheetApp.getUi().alert(
    "Validation Failed",
    validation.errors.join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return;
}
    setWorkflowStatus_(
  context.row,
  WORKFLOW_STATUS.VALIDATED
);
    logAudit_(
  "VALIDATE_ROW",
  context.documentId,
  "SUCCESS",
  "Active Production Board row validated."
);
refreshDashboardSafe_();
    SpreadsheetApp.getUi().alert(
      "Active Row Valid",
      [
        "Document ID: " + context.documentId,
        "Title: " + context.documentTitle,
        "Category: " + context.category,
        "Assigned Lead: " + context.assignedLead
      ].join("\n"),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    handleError_("Validation failed", error);
  }
}

function validateWorkbook_(spreadsheet) {
  const production = spreadsheet.getSheetByName(JPAIS.SHEETS.PRODUCTION);
  if (!production) {
    throw new Error('Required sheet "' + JPAIS.SHEETS.PRODUCTION + '" was not found.');
  }

  const headerMap = getHeaderMap_(production, 4);
  const missing = JPAIS.REQUIRED_PRODUCTION_HEADERS.filter(function(header) {
    return !headerMap[header];
  });

  if (missing.length) {
    throw new Error("Missing Production Board headers: " + missing.join(", "));
  }
}

function getActiveProductionContext_() {
  const spreadsheet = SpreadsheetApp.getActive();
  validateWorkbook_(spreadsheet);

  const sheet = spreadsheet.getActiveSheet();
  if (sheet.getName() !== JPAIS.SHEETS.PRODUCTION) {
    throw new Error('Open the "' + JPAIS.SHEETS.PRODUCTION + '" sheet first.');
  }

  const row = sheet.getActiveRange().getRow();
  if (row < 5) {
    throw new Error("Select a data row below the header.");
  }

  const headerMap = getHeaderMap_(sheet, 4);

  const context = {
    spreadsheet: spreadsheet,
    sheet: sheet,
    row: row,
    headerMap: headerMap,
    documentId: getByHeader_(sheet, row, headerMap, "Document ID"),
    documentTitle: getByHeader_(sheet, row, headerMap, "Document Title"),
    category: getByHeader_(sheet, row, headerMap, "Category"),
    assignedLead: getByHeader_(sheet, row, headerMap, "Assigned Lead"),
    status: getByHeader_(sheet, row, headerMap, "Status"),
    driveFolderUrl: getByHeader_(sheet, row, headerMap, "Drive Folder URL"),
    notebookLmUrl: getByHeader_(sheet, row, headerMap, "NotebookLM URL")
  };

  if (!context.documentId) throw new Error("Document ID is required.");
  if (!context.documentTitle) throw new Error("Document Title is required.");
  if (!context.category) throw new Error("Category is required.");

  return context;
}

function validateNoExistingPackageConflict_(context) {
  if (!context.driveFolderUrl) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Existing Package Detected",
    "This row already contains a Drive Folder URL. Continue and reuse the existing package?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    throw new Error("Operation cancelled by user.");
  }
}
