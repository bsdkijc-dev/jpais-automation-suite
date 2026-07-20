/**
 * Adds the JPAIS menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("JPAIS")
    .addItem("Initialize Suite", "initializeJPAIS")
    .addSeparator()
    .addItem(
      "Create Complete Knowledge Package",
      "createCompleteKnowledgePackage"
    )
    .addItem(
      "Create Folder Structure Only",
      "createFolderStructureOnly"
    )
    .addSeparator()
    .addItem(
      "Update Resource Status",
      "updatePackageResourceStatus"
    )
    .addItem(
      "Rebuild STATUS.json from Board",
      "rebuildPackageStatusFromBoard"
    )
    .addSeparator()
    .addItem("Open JPAIS Drive", "openJPAISDrive")
    .addItem("Validate Active Row", "validateActiveProductionRow")
    .addItem("Refresh Dashboard", "refreshDashboard")
    .addToUi();
}

/**
 * One-time validation and initialization.
 */
function initializeJPAIS() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActive();
    validateWorkbook_(ss);
    const root = DriveApp.getFolderById(JPAIS.ROOT_FOLDER_ID);
    createStandardRootFolders_(root);
    ss.toast("JPAIS Automation Suite v" + JPAIS.VERSION + " initialized.", "JPAIS", 6);
    ui.alert(
      "Initialization Complete",
      "Workbook structure and JPAIS Drive root were validated successfully.",
      ui.ButtonSet.OK
    );
  } catch (error) {
    handleError_("Initialization failed", error);
  }
}
