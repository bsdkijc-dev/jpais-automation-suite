/**
 * Lightweight tests that can run inside Apps Script without external libraries.
 * Run runJPAISCoreTests() from the Apps Script editor or clasp run.
 */
function runJPAISCoreTests() {
  const tests = [
    testSanitizeName_,
    testBuildHeaderMap_,
    testPackageConfiguration_,
    testRequiredHeaders_,
    testStatusNormalization_
  ];

  const results = [];
  tests.forEach(function(testFn) {
    try {
      testFn();
      results.push({ name: testFn.name, status: "PASS" });
    } catch (error) {
      results.push({
        name: testFn.name,
        status: "FAIL",
        message: error.message || String(error)
      });
    }
  });

  console.log(JSON.stringify(results, null, 2));

  const failed = results.filter(function(result) {
    return result.status === "FAIL";
  });

  if (failed.length) {
    throw new Error(failed.length + " core test(s) failed.");
  }

  return results;
}

function testSanitizeName_() {
  assertEqual_(
    sanitizeName_("PERMA 1/2016: Mediation"),
    "PERMA-1-2016-Mediation",
    "sanitizeName_ should normalize spaces and prohibited characters"
  );
}

function testBuildHeaderMap_() {
  const map = buildHeaderMapFromValues_([
    "Document ID",
    "Document Title",
    "",
    "Status"
  ]);

  assertEqual_(map["Document ID"], 1, "Document ID column");
  assertEqual_(map["Document Title"], 2, "Document Title column");
  assertEqual_(map["Status"], 4, "Status column");
}

function testPackageConfiguration_() {
  assertEqual_(
    JPAIS.PACKAGE_FOLDERS.length,
    10,
    "JPAIS must define ten standard package folders"
  );

  assertTrue_(
    JPAIS.REQUIRED_ASSETS.indexOf("statusFile") !== -1,
    "statusFile must be a required asset"
  );
}

function testRequiredHeaders_() {
  ["Document ID", "Document Title", "Drive Folder URL"].forEach(function(header) {
    assertTrue_(
      JPAIS.REQUIRED_PRODUCTION_HEADERS.indexOf(header) !== -1,
      "Required header missing: " + header
    );
  });
}

function testStatusNormalization_() {
  assertEqual_(
    normalizeBoardStatus_("Not Started"),
    JPAIS_STATUS.PENDING,
    "Not Started should normalize to Pending"
  );
  assertEqual_(
    normalizeBoardStatus_("Completed"),
    "Completed",
    "Completed should remain Completed"
  );
}

/**
 * Integration test: validates the currently selected Production Board row
 * without creating files.
 */
function testActiveProductionRowIntegration() {
  const context = getActiveProductionContext_();
  assertTrue_(Boolean(context.documentId), "Document ID must exist");
  assertTrue_(Boolean(context.documentTitle), "Document Title must exist");
  assertTrue_(Boolean(context.category), "Category must exist");
  console.log(JSON.stringify(context, null, 2));
  return context;
}

/**
 * Integration test: verifies the standard package folder hierarchy.
 * This may create missing folders but does not generate content assets.
 */
function testPackageFolderStructureIntegration() {
  const context = getActiveProductionContext_();
  const packageInfo = ensureKnowledgePackageFolders_(context);

  JPAIS.PACKAGE_FOLDERS.forEach(function(folderName) {
    assertTrue_(
      Boolean(packageInfo.subfolders[folderName]),
      "Missing folder: " + folderName
    );
  });

  console.log(packageInfo.documentFolder.getUrl());
  return packageInfo.documentFolder.getUrl();
}

function assertEqual_(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(
      message + ". Expected: " + expected + "; Actual: " + actual
    );
  }
}

function assertTrue_(condition, message) {
  if (!condition) throw new Error(message);
}
