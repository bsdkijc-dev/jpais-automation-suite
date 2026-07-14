/**
 * Updates the active Production Board row after folder/package creation.
 */
function updateProductionRowAfterCreation_(context, packageInfo, complete, assets) {
  const headerMap = context.headerMap;
  const sheet = context.sheet;
  const row = context.row;

  setByHeader_(sheet, row, headerMap, "Drive Folder URL", packageInfo.documentFolder.getUrl());
  setByHeader_(sheet, row, headerMap, "Status", "In Progress");

  if (complete && assets) {
    setByHeader_(sheet, row, headerMap, "Executive Summary", "Completed");
    setByHeader_(sheet, row, headerMap, "Slides", "Completed");
    setByHeader_(sheet, row, headerMap, "FAQ", "Completed");
    setByHeader_(sheet, row, headerMap, "Quiz", "Completed");
    setByHeader_(sheet, row, headerMap, "Metadata", "Completed");
    setByHeader_(sheet, row, headerMap, "References", "Completed");

    const linksNote = [
      "Knowledge Package generated automatically.",
      "Executive Summary: " + assets.executiveSummary.url,
      "Slides: " + assets.presentation.url,
      "FAQ: " + assets.faq.url,
      "Quiz: " + assets.quiz.publishedUrl,
      "Metadata: " + assets.metadata.url,
      "References: " + assets.references.url,
      "Created: " + formatTimestamp_(new Date())
    ].join("\n");

    setByHeader_(sheet, row, headerMap, "Remarks", linksNote);
  } else {
    setByHeader_(
      sheet,
      row,
      headerMap,
      "Remarks",
      "Knowledge Package folder structure created on " + formatTimestamp_(new Date())
    );
  }
}

/**
 * Adds or updates a registry record.
 */
function upsertDocumentRegistry_(context, packageInfo, assets) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(JPAIS.SHEETS.DOCUMENT_REGISTRY);
  if (!sheet) return;

  const headerRow = 4;
  const headers = getHeaderMap_(sheet, headerRow);
  const row = findOrAppendRowByValue_(sheet, headers["Document ID"], context.documentId, 5);

  setByHeader_(sheet, row, headers, "Document ID", context.documentId);
  setByHeader_(sheet, row, headers, "Official Title", context.documentTitle);
  setByHeader_(sheet, row, headers, "Document Type", context.category);
  setByHeader_(sheet, row, headers, "Date Added", new Date());
  setByHeader_(sheet, row, headers, "Added By", Session.getActiveUser().getEmail() || context.assignedLead);
  setByHeader_(sheet, row, headers, "Version", "1.0");
  setByHeader_(sheet, row, headers, "Access Level", "Internal");
  setByHeader_(sheet, row, headers, "Notes", "Drive folder: " + packageInfo.documentFolder.getUrl());
}

/**
 * Adds one AI tracker row for every resource generated or planned.
 */
function upsertAIResourceTracker_(context, assets) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(JPAIS.SHEETS.AI_TRACKER);
  if (!sheet) return;

  const headers = getHeaderMap_(sheet, 4);

  const resources = [
    ["Executive Summary", "Google Docs", "Completed", assets.executiveSummary.url],
    ["Audio Overview", "NotebookLM / YouTube", "Not Started", ""],
    ["Presentation Slides", "Google Slides", "Completed", assets.presentation.url],
    ["Mind Map", "NotebookLM / Drive", "Not Started", ""],
    ["Infographic", "Canva / Drive", "Not Started", ""],
    ["FAQ", "Google Docs", "Completed", assets.faq.url],
    ["Quiz", "Google Forms", "Completed", assets.quiz.publishedUrl],
    ["Data Table", "Google Sheets", "Completed", assets.metadata.url],
    ["References", "Google Docs", "Completed", assets.references.url]
  ];

  resources.forEach(function(resource) {
    const existingRow = findAITrackerRow_(sheet, headers, context.documentId, resource[0]);
    const row = existingRow || Math.max(sheet.getLastRow() + 1, 5);

    setByHeader_(sheet, row, headers, "Document ID", context.documentId);
    setByHeader_(sheet, row, headers, "Document Title", context.documentTitle);
    setByHeader_(sheet, row, headers, "Resource Type", resource[0]);
    setByHeader_(sheet, row, headers, "Platform", resource[1]);
    setByHeader_(sheet, row, headers, "Assigned To", context.assignedLead);
    setByHeader_(sheet, row, headers, "Status", resource[2]);
    setByHeader_(sheet, row, headers, "Final URL", resource[3]);
    setByHeader_(sheet, row, headers, "Version", "1.0");
    setByHeader_(sheet, row, headers, "Created Date", new Date());
    setByHeader_(sheet, row, headers, "Approval Status", "Pending");
  });
}

function findAITrackerRow_(sheet, headers, documentId, resourceType) {
  if (!headers["Document ID"] || !headers["Resource Type"] || sheet.getLastRow() < 5) return null;

  const values = sheet.getRange(
    5,
    1,
    sheet.getLastRow() - 4,
    sheet.getLastColumn()
  ).getValues();

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][headers["Document ID"] - 1]).trim() === documentId &&
      String(values[i][headers["Resource Type"] - 1]).trim() === resourceType
    ) {
      return i + 5;
    }
  }
  return null;
}

/**
 * Adds or updates the Google Sites publication register.
 */
function upsertPublisherRegister_(context, packageInfo, assets) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(JPAIS.SHEETS.PUBLISHER);
  if (!sheet) return;

  const headers = getHeaderMap_(sheet, 4);
  const row = findOrAppendRowByValue_(sheet, headers["Document ID"], context.documentId, 5);

  setByHeader_(sheet, row, headers, "Page ID", "SITE-" + context.documentId);
  setByHeader_(sheet, row, headers, "Page Title", context.documentTitle);
  setByHeader_(sheet, row, headers, "Content Type", "Knowledge Package");
  setByHeader_(sheet, row, headers, "Document ID", context.documentId);
  setByHeader_(sheet, row, headers, "Page Owner", context.assignedLead || "Cecep Mustafa");
  setByHeader_(sheet, row, headers, "Publication Status", "Draft");
  setByHeader_(sheet, row, headers, "Summary URL", assets.executiveSummary.url);
  setByHeader_(sheet, row, headers, "Podcast URL", "");
  setByHeader_(sheet, row, headers, "Slides URL", assets.presentation.url);
  setByHeader_(sheet, row, headers, "Infographic URL", "");
  setByHeader_(sheet, row, headers, "Last Updated", new Date());
  setByHeader_(sheet, row, headers, "Notes", "Package folder: " + packageInfo.documentFolder.getUrl());
}

/**
 * Recalculates spreadsheet formulas.
 */
function refreshDashboard() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast("Dashboard refreshed.", "JPAIS", 4);
}
