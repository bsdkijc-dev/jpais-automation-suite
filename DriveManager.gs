/**
 * Ensures the standard root folder structure exists.
 */
function createStandardRootFolders_(rootFolder) {
  [
    "01_Regulations",
    "02_Research",
    "03_Policy_Briefs",
    "04_AI_Resources",
    "05_Repository",
    "06_Training",
    "07_Publication",
    "08_Archive"
  ].forEach(function(name) {
    getOrCreateFolder_(rootFolder, name);
  });
}

/**
 * Creates only the folder structure for the active Production Board row.
 */
function createFolderStructureOnly() {
  try {
    const context = getActiveProductionContext_();
    const packageInfo = ensureKnowledgePackageFolders_(context);
    updateProductionRowAfterCreation_(context, packageInfo, false);
    SpreadsheetApp.getUi().alert(
      "Folder Structure Created",
      packageInfo.documentFolder.getName() + "\n\n" + packageInfo.documentFolder.getUrl(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    handleError_("Folder creation failed", error);
  }
}

/**
 * Creates or reuses the correct main category folder and document package folders.
 */
function ensureKnowledgePackageFolders_(context) {
  const rootFolder = DriveApp.getFolderById(JPAIS.ROOT_FOLDER_ID);
  createStandardRootFolders_(rootFolder);

  const mainFolderName = JPAIS.MAIN_FOLDERS[context.category] || "05_Repository";
  const mainFolder = getOrCreateFolder_(rootFolder, mainFolderName);

  const documentFolderName =
    sanitizeName_(context.documentId) + "_" + sanitizeName_(context.documentTitle);
  const documentFolder = getOrCreateFolder_(mainFolder, documentFolderName);

  const subfolders = {};
  JPAIS.PACKAGE_FOLDERS.forEach(function(folderName) {
    subfolders[folderName] = getOrCreateFolder_(documentFolder, folderName);
  });

  return {
    rootFolder: rootFolder,
    mainFolder: mainFolder,
    documentFolder: documentFolder,
    subfolders: subfolders
  };
}

/**
 * Returns an existing direct child folder or creates it.
 */
function getOrCreateFolder_(parentFolder, folderName) {
  const iterator = parentFolder.getFoldersByName(folderName);
  return iterator.hasNext() ? iterator.next() : parentFolder.createFolder(folderName);
}

/**
 * Opens the root JPAIS Drive folder.
 */
function openJPAISDrive() {
  try {
    const folder = DriveApp.getFolderById(JPAIS.ROOT_FOLDER_ID);
    const html = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;padding:12px">' +
      '<p><strong>JPAIS Drive</strong></p>' +
      '<p><a href="' + folder.getUrl() + '" target="_blank">Open the JPAIS root folder</a></p>' +
      '</div>'
    ).setWidth(360).setHeight(140);

    SpreadsheetApp.getUi().showModalDialog(html, "JPAIS Drive");
  } catch (error) {
    handleError_("Unable to open JPAIS Drive", error);
  }
}
