function getHeaderMap_(sheet, headerRow) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return {};

  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getDisplayValues()[0];
  return buildHeaderMapFromValues_(headers);
}

function buildHeaderMapFromValues_(headers) {
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
    const values = sheet
      .getRange(firstDataRow, column, lastRow - firstDataRow + 1, 1)
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
  const message = error && error.message ? error.message : String(error);
  SpreadsheetApp.getActive().toast(message, "JPAIS Error", 8);
  SpreadsheetApp.getUi().alert(
    title,
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function getFirstFileByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

function createTextFileIfMissing_(folder, fileName, content) {
  const existing = getFirstFileByName_(folder, fileName);
  if (existing) return existing;
  return folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
}

function createOrReplaceTextFile_(folder, fileName, content) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
  return folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
}

function moveFileToFolder_(fileId, destinationFolder) {
  DriveApp.getFileById(fileId).moveTo(destinationFolder);
}

function assetResult_(file, extra) {
  const result = {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
  if (extra) {
    Object.keys(extra).forEach(function(key) {
      result[key] = extra[key];
    });
  }
  return result;
}

function validateAssetSet_(assets) {
  const missing = JPAIS.REQUIRED_ASSETS.filter(function(key) {
    return !assets[key] || !assets[key].id || !assets[key].url;
  });

  if (missing.length) {
    throw new Error("Knowledge Package assets incomplete: " + missing.join(", "));
  }
}

function showPackageResultDialog_(context, packageInfo, assets) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:12px;line-height:1.5">' +
    '<h3 style="margin-top:0">Knowledge Package Ready</h3>' +
    '<p><strong>' + escapeHtml_(context.documentTitle) + '</strong></p>' +
    '<p>Generated assets were created or safely reused.</p>' +
    '<ul>' +
    '<li><a target="_blank" href="' + packageInfo.documentFolder.getUrl() + '">Open package folder</a></li>' +
    '<li><a target="_blank" href="' + assets.executiveSummary.url + '">Open Executive Summary</a></li>' +
    '<li><a target="_blank" href="' + assets.presentation.url + '">Open Presentation</a></li>' +
    '<li><a target="_blank" href="' + assets.faq.url + '">Open FAQ</a></li>' +
    '<li><a target="_blank" href="' + assets.quiz.editUrl + '">Edit Quiz</a></li>' +
    '<li><a target="_blank" href="' + assets.metadata.url + '">Open Metadata Sheet</a></li>' +
    '<li><a target="_blank" href="' + assets.references.url + '">Open References</a></li>' +
    '<li><a target="_blank" href="' + assets.readme.url + '">Open README</a></li>' +
    '<li><a target="_blank" href="' + assets.qaChecklist.url + '">Open QA Checklist</a></li>' +
    '<li><a target="_blank" href="' + assets.manifest.url + '">Open Manifest</a></li>' +
    '<li><a target="_blank" href="' + assets.statusFile.url + '">Open STATUS.json</a></li>' +
    '</ul>' +
    '</div>'
  ).setWidth(510).setHeight(500);

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
