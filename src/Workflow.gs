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

  const headerMap = buildHeaderMap_(sheet);

  const column = headerMap["Workflow Status"];

  if (!column) {
    return;
  }

  sheet.getRange(rowIndex, column).setValue(status);

}