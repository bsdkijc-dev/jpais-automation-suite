/**
 * ==========================================================
 * Workflow.gs
 * JPAIS Automation Suite v1.2.0
 * ==========================================================
 */

/**
 * Mengubah status workflow pada Production Board.
 *
 * @param {number} rowIndex Nomor baris di spreadsheet.
 * @param {string} status Status workflow baru.
 */
function setWorkflowStatus_(rowIndex, status) {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(JPAIS.SHEETS.PRODUCTION);

  const headerMap = getHeaderMap_(sheet, 4);

  const column = headerMap["Workflow Status"];

  if (!column) {
    return;
  }

  sheet.getRange(rowIndex, column).setValue(status);

}
/**
 * Mengambil workflow status dari satu baris Production Board.
 *
 * @param {number} rowIndex Nomor baris spreadsheet.
 * @return {string} Workflow status.
 */
function getWorkflowStatus_(rowIndex) {
  const spreadsheet = SpreadsheetApp.getActive();

  if (!spreadsheet) {
    return "";
  }

  const sheet = spreadsheet.getSheetByName(
    JPAIS.SHEETS.PRODUCTION
  );

  if (!sheet) {
    return "";
  }

  const headerMap = getHeaderMap_(sheet, 4);
  const column = headerMap["Workflow Status"];

  if (!column) {
    return "";
  }

  return String(
    sheet.getRange(rowIndex, column).getDisplayValue()
  ).trim();
}