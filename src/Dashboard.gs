/**
 * ==========================================================
 * Dashboard.gs
 * JPAIS Automation Suite v1.3.0
 * ==========================================================
 */

/**
 * Memperbarui ringkasan Dashboard dari Production Board
 * dan Audit Log.
 */
function refreshDashboard() {
  const spreadsheet = SpreadsheetApp.getActive();

  if (!spreadsheet) {
    throw new Error("Active spreadsheet tidak tersedia.");
  }

  const productionSheet = spreadsheet.getSheetByName(
    JPAIS.SHEETS.PRODUCTION
  );

  if (!productionSheet) {
    throw new Error(
      'Sheet "' + JPAIS.SHEETS.PRODUCTION + '" tidak ditemukan.'
    );
  }

  const dashboardSheet = getOrCreateDashboardSheet_(spreadsheet);
  const summary = buildDashboardSummary_(productionSheet);

  writeDashboardSummary_(dashboardSheet, summary);

  logAudit_(
    "REFRESH_DASHBOARD",
    "",
    "SUCCESS",
    "Dashboard refreshed successfully."
  );

  spreadsheet.toast(
    "Dashboard berhasil diperbarui.",
    "JPAIS",
    5
  );
}

/**
 * Mengambil atau membuat sheet Dashboard.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateDashboardSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName("Dashboard");

  if (!sheet) {
    sheet = spreadsheet.insertSheet("Dashboard");
  }

  return sheet;
}

/**
 * Menghitung ringkasan status Production Board.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Object}
 */
function buildDashboardSummary_(sheet) {
  const headerMap = getHeaderMap_(sheet, 4);
  const firstDataRow = 5;
  const lastRow = sheet.getLastRow();

  const summary = {
    totalDocuments: 0,
    draft: 0,
    validated: 0,
    generating: 0,
    packageReady: 0,
    qaReview: 0,
    approved: 0,
    published: 0
  };

  if (lastRow < firstDataRow) {
    return summary;
  }

  const documentIdColumn = headerMap["Document ID"];
  const workflowColumn = headerMap["Workflow Status"];

  if (!documentIdColumn) {
    throw new Error('Header "Document ID" tidak ditemukan.');
  }

  const documentIds = sheet
    .getRange(
      firstDataRow,
      documentIdColumn,
      lastRow - firstDataRow + 1,
      1
    )
    .getDisplayValues();

  const workflowValues = workflowColumn
    ? sheet
        .getRange(
          firstDataRow,
          workflowColumn,
          lastRow - firstDataRow + 1,
          1
        )
        .getDisplayValues()
    : [];

  documentIds.forEach(function(row, index) {
    const documentId = String(row[0] || "").trim();

    if (!documentId) return;

    summary.totalDocuments++;

    const status = workflowColumn
      ? String(workflowValues[index][0] || "").trim()
      : "";

    switch (status) {
      case WORKFLOW_STATUS.DRAFT:
        summary.draft++;
        break;
      case WORKFLOW_STATUS.VALIDATED:
        summary.validated++;
        break;
      case WORKFLOW_STATUS.GENERATING:
        summary.generating++;
        break;
      case WORKFLOW_STATUS.READY:
        summary.packageReady++;
        break;
      case WORKFLOW_STATUS.QA:
        summary.qaReview++;
        break;
      case WORKFLOW_STATUS.APPROVED:
        summary.approved++;
        break;
      case WORKFLOW_STATUS.PUBLISHED:
        summary.published++;
        break;
    }
  });

  return summary;
}

/**
 * Menulis data ringkasan ke Dashboard.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} summary
 */
function writeDashboardSummary_(sheet, summary) {
  sheet.clear();

  sheet.getRange("A1:B1")
    .merge()
    .setValue("JPAIS Dashboard")
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center");

  const rows = [
    ["Metric", "Value"],
    ["Total Documents", summary.totalDocuments],
    ["Draft", summary.draft],
    ["Validated", summary.validated],
    ["Generating Package", summary.generating],
    ["Package Ready", summary.packageReady],
    ["QA Review", summary.qaReview],
    ["Approved", summary.approved],
    ["Published", summary.published],
    ["Last Updated", new Date()]
  ];

  sheet.getRange(3, 1, rows.length, 2).setValues(rows);

  sheet.getRange(3, 1, 1, 2)
    .setBackground("#D9EAF7")
    .setFontWeight("bold");

  sheet.getRange(3, 1, rows.length, 2)
    .setBorder(true, true, true, true, true, true);

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 160);
  sheet.setFrozenRows(3);
}