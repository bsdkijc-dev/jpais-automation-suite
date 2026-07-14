/**
 * ==========================================================
 * Audit.gs
 * JPAIS Automation Suite v1.2.0
 * ==========================================================
 */

/**
 * Menambahkan satu catatan ke sheet Audit Log.
 *
 * @param {string} action Nama tindakan.
 * @param {string} documentId ID dokumen.
 * @param {string} result SUCCESS atau FAILED.
 * @param {string} message Keterangan tindakan.
 */
function logAudit_(action, documentId, result, message) {
  const spreadsheet = SpreadsheetApp.getActive();

  if (!spreadsheet) {
    console.warn("Audit skipped: active spreadsheet not available.");
    return;
  }

  const sheet = getOrCreateAuditLogSheet_(spreadsheet);
  const userEmail =
    Session.getActiveUser().getEmail() ||
    Session.getEffectiveUser().getEmail() ||
    "Unknown User";

  sheet.appendRow([
    new Date(),
    userEmail,
    action,
    documentId || "",
    result,
    message || ""
  ]);
}

/**
 * Mengambil atau membuat sheet Audit Log.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateAuditLogSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(JPAIS.SHEETS.AUDIT_LOG);

  if (sheet) {
    return sheet;
  }

  sheet = spreadsheet.insertSheet(JPAIS.SHEETS.AUDIT_LOG);

  const headers = [
    "Timestamp",
    "User",
    "Action",
    "Document ID",
    "Result",
    "Message"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  return sheet;
}