function getHeaderMap_(sheet, headerRow) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return {};

  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};

  headers.forEach(function(header, index) {
    const clean = String(header).trim();
    if (clean) map[clean] = index + 1;
  });

  return map;
}

function getByHeader_(sheet, row, headerMap, header) {
  if (!headerMap[header]) return "";
  return String(sheet.getRange(row, headerMap[header]).getDisplayValue()).trim();
}

function setByHeader_(sheet, row, headerMap, header, value) {
  if (!headerMap[header]) return;
  sheet.getRange(row, headerMap[header]).setValue(value);
}

function findOrAppendRowByValue_(sheet, column, value, firstDataRow) {
  if (!column) return Math.max(sheet.getLastRow() + 1, firstDataRow);

  const lastRow = sheet.getLastRow();
  if (lastRow >= firstDataRow) {
    const values = sheet.getRange(firstDataRow, column, lastRow - firstDataRow + 1, 1)
      .getDisplayValues();

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(value).trim()) {
        return firstDataRow + i;
      }
    }
  }

  return Math.max(lastRow + 1, firstDataRow);
}

function sanitizeName_(value) {
  return String(value)
    .replace(/[\\/:*?"<>|#%{}]/g, "-")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .substring(0, 120);
}

function formatTimestamp_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "dd MMM yyyy HH:mm"
  );
}

function handleError_(title, error) {
  console.error(error);
  SpreadsheetApp.getActive().toast(error.message || String(error), "JPAIS Error", 8);
  SpreadsheetApp.getUi().alert(
    title,
    error.message || String(error),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showPackageResultDialog_(context, packageInfo, assets) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:12px;line-height:1.5">' +
    '<h3 style="margin-top:0">Knowledge Package Created</h3>' +
    '<p><strong>' + escapeHtml_(context.documentTitle) + '</strong></p>' +
    '<ul>' +
    '<li><a target="_blank" href="' + packageInfo.documentFolder.getUrl() + '">Open package folder</a></li>' +
    '<li><a target="_blank" href="' + assets.executiveSummary.url + '">Open Executive Summary</a></li>' +
    '<li><a target="_blank" href="' + assets.presentation.url + '">Open Presentation</a></li>' +
    '<li><a target="_blank" href="' + assets.faq.url + '">Open FAQ</a></li>' +
    '<li><a target="_blank" href="' + assets.quiz.editUrl + '">Edit Quiz</a></li>' +
    '<li><a target="_blank" href="' + assets.metadata.url + '">Open Metadata Sheet</a></li>' +
    '<li><a target="_blank" href="' + assets.references.url + '">Open References</a></li>' +
    '<li><a target="_blank" href="' + assets.statusFile.url + '">Open STATUS.json</a></li>' +
    '</ul>' +
    '</div>'
  ).setWidth(480).setHeight(360);

  SpreadsheetApp.getUi().showModalDialog(html, "JPAIS Automation Suite");
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
