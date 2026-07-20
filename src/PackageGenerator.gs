/**
 * Creates or safely reuses a complete JPAIS Knowledge Package.
 *
 * The operation is idempotent:
 * - Existing package folders are reused.
 * - Existing Google Workspace assets are reused by exact file name.
 * - Team-authored content is never overwritten.
 * - STATUS.json and MANIFEST.csv are refreshed to reflect the current package.
 */
function createCompleteKnowledgePackage() {
  const lock = LockService.getDocumentLock();
  let lockAcquired = false;

  try {
    lock.waitLock(30000);
    lockAcquired = true;

    const context = getActiveProductionContext_();
    setWorkflowStatus_(
  context.row,
  WORKFLOW_STATUS.GENERATING
);
    validateNoExistingPackageConflict_(context);

    const packageInfo = ensureKnowledgePackageFolders_(context);
    const assets = createPackageAssets_(context, packageInfo);
    const quality = evaluatePackageQuality_(assets);

if (!quality.passed) {
  SpreadsheetApp.getUi().alert(
    "QA Checklist Failed",
    "Missing assets:\n\n• " + quality.missing.join("\n• "),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  logAudit_(
    "QA_CHECK",
    context.documentId,
    "FAILED",
    quality.missing.join(", ")
  );

  return;
}

logAudit_(
  "QA_CHECK",
  context.documentId,
  "SUCCESS",
  quality.passedCount + "/" + quality.total + " assets verified."
);
    validateAssetSet_(assets);

    updateProductionRowAfterCreation_(context, packageInfo, true, assets);
    upsertDocumentRegistry_(context, packageInfo, assets);
    upsertAIResourceTracker_(context, assets);
    upsertPublisherRegister_(context, packageInfo, assets);

    SpreadsheetApp.flush();
    setWorkflowStatus_(
  context.row,
  WORKFLOW_STATUS.READY
);
    logAudit_(
  "CREATE_PACKAGE",
  context.documentId,
  "SUCCESS",
  "Knowledge Package created or safely reused."
);
refreshDashboardSafe_();
    SpreadsheetApp.getActive().toast(
      "Knowledge Package created or safely reused.",
      "JPAIS",
      8
    );

    showPackageResultDialog_(context, packageInfo, assets);
  } catch (error) {
    handleError_("Knowledge Package creation failed", error);
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function createPackageAssets_(context, packageInfo) {
  const assets = {
    executiveSummary: getOrCreateExecutiveSummaryDoc_(
      context,
      packageInfo.subfolders["02_Executive_Summary"]
    ),
    presentation: getOrCreatePresentation_(
      context,
      packageInfo.subfolders["04_Presentation_Slides"]
    ),
    faq: getOrCreateFAQDoc_(
      context,
      packageInfo.subfolders["07_FAQ"]
    ),
    quiz: getOrCreateQuizForm_(
      context,
      packageInfo.subfolders["08_Quiz"]
    ),
    metadata: getOrCreateMetadataSheet_(
      context,
      packageInfo.subfolders["09_Data_Table"]
    ),
    references: getOrCreateReferencesDoc_(
      context,
      packageInfo.subfolders["10_References"]
    )
  };

  assets.readme = createPackageReadme_(context, packageInfo);
  assets.qaChecklist = createQAChecklist_(context, packageInfo);
  assets.manifest = createPackageManifest_(context, packageInfo, assets);
  assets.statusFile = createPackageStatusFile_(context, packageInfo, assets);

  return assets;
}

function getOrCreateExecutiveSummaryDoc_(context, destinationFolder) {
  const fileName = context.documentId + " - Executive Summary";
  const existing = getFirstFileByName_(destinationFolder, fileName);
  if (existing) return assetResult_(existing);

  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();

  body.appendParagraph("JPAIS EXECUTIVE SUMMARY")
    .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(context.documentTitle)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  appendMetadataTableToDoc_(body, context);

  [
    ["1. Executive Summary", "Provide a concise overview of the document and its institutional significance."],
    ["2. Background", "Describe the legal, policy, and institutional context."],
    ["3. Legal Issues", "Identify the principal legal or regulatory issues."],
    ["4. Legal Basis", "List the relevant constitutional, statutory, regulatory, and institutional authorities."],
    ["5. Analysis", "Present source-based legal and policy analysis."],
    ["6. Implementation Implications", "Explain institutional roles, resources, procedures, and timelines."],
    ["7. Risks", "Identify legal, operational, data, security, and implementation risks."],
    ["8. Recommendations", "Provide actionable, evidence-based recommendations."],
    ["9. Conclusion", "State the principal conclusion and next actions."]
  ].forEach(function(section) {
    body.appendParagraph(section[0])
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(section[1]);
  });

  appendQualityNotice_(body);
  doc.saveAndClose();
  moveFileToFolder_(doc.getId(), destinationFolder);

  return assetResult_(DriveApp.getFileById(doc.getId()));
}

function getOrCreatePresentation_(context, destinationFolder) {
  const fileName = context.documentId + " - Presentation";
  const existing = getFirstFileByName_(destinationFolder, fileName);
  if (existing) return assetResult_(existing);

  const presentation = SlidesApp.create(fileName);
  const slides = presentation.getSlides();
  const titleSlide = slides[0];
  const shapes = titleSlide.getShapes();

  if (shapes.length > 0) {
    shapes[0].getText().setText(context.documentTitle);
  }
  if (shapes.length > 1) {
    shapes[1].getText().setText(
      "JPAIS Knowledge Package\nDocument ID: " + context.documentId
    );
  }

  [
    ["Overview", "Purpose, scope, and institutional context."],
    ["Background", "Legal, policy, and operational background."],
    ["Legal Basis", "Relevant laws, regulations, policies, and authority."],
    ["Key Provisions or Findings", "Core provisions, issues, findings, or recommendations."],
    ["Implementation", "Institutional roles, procedures, resources, and timeline."],
    ["Risks and Mitigation", "Legal, operational, data, security, and implementation risks."],
    ["Recommendations", "Prioritized institutional actions."],
    ["Conclusion", "Principal conclusions and next actions."]
  ].forEach(function(section) {
    const slide = presentation.appendSlide(
      SlidesApp.PredefinedLayout.TITLE_AND_BODY
    );
    const placeholders = slide.getPlaceholders();
    if (placeholders.length > 0) {
      placeholders[0].asShape().getText().setText(section[0]);
    }
    if (placeholders.length > 1) {
      placeholders[1].asShape().getText().setText(section[1]);
    }
  });

  moveFileToFolder_(presentation.getId(), destinationFolder);
  return assetResult_(DriveApp.getFileById(presentation.getId()));
}

function getOrCreateMetadataSheet_(context, destinationFolder) {
  const fileName = context.documentId + " - Metadata";
  const existing = getFirstFileByName_(destinationFolder, fileName);
  if (existing) return assetResult_(existing);

  const spreadsheet = SpreadsheetApp.create(fileName);
  const sheet = spreadsheet.getSheets()[0];
  sheet.setName("Metadata");

  const rows = [
    ["Field", "Value"],
    ["Document ID", context.documentId],
    ["Official Title", context.documentTitle],
    ["Category", context.category],
    ["Assigned Lead", context.assignedLead || ""],
    ["Status", context.status || "In Progress"],
    ["Language", ""],
    ["Issuing Institution", ""],
    ["Document Number", ""],
    ["Year", ""],
    ["Subject Area", ""],
    ["Keywords", ""],
    ["Legal Basis", ""],
    ["Legal Status", ""],
    ["Source URL", ""],
    ["NotebookLM URL", context.notebookLmUrl || ""],
    ["Date Created", new Date()],
    ["Version", JPAIS.VERSION],
    ["Access Level", "Internal"],
    ["Review Status", "Pending"]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange("A1:B1")
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
  sheet.getRange("A1:B" + rows.length).setWrap(true);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 420);
  sheet.setFrozenRows(1);

  moveFileToFolder_(spreadsheet.getId(), destinationFolder);
  return assetResult_(DriveApp.getFileById(spreadsheet.getId()));
}

function getOrCreateFAQDoc_(context, destinationFolder) {
  const fileName = context.documentId + " - FAQ";
  const existing = getFirstFileByName_(destinationFolder, fileName);
  if (existing) return assetResult_(existing);

  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();

  body.appendParagraph("FREQUENTLY ASKED QUESTIONS")
    .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(context.documentTitle)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  for (let i = 1; i <= 10; i++) {
    body.appendParagraph(i + ". [Insert verified question]")
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(
      "[Insert concise, source-based answer. Include a citation or official source link.]"
    );
  }

  appendQualityNotice_(body);
  doc.saveAndClose();
  moveFileToFolder_(doc.getId(), destinationFolder);

  return assetResult_(DriveApp.getFileById(doc.getId()));
}

function getOrCreateReferencesDoc_(context, destinationFolder) {
  const fileName = context.documentId + " - References";
  const existing = getFirstFileByName_(destinationFolder, fileName);
  if (existing) return assetResult_(existing);

  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();

  body.appendParagraph("REFERENCES AND RELATED DOCUMENTS")
    .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(context.documentTitle)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  [
    "Official source document",
    "Constitution and legislation",
    "Related PERMA, SEMA, and SK KMA",
    "Related court decisions",
    "Judicial research and policy briefs",
    "Academic references",
    "International standards",
    "Comparative-law materials"
  ].forEach(function(item) {
    body.appendListItem(item + ": [Insert verified citation and URL]");
  });

  appendQualityNotice_(body);
  doc.saveAndClose();
  moveFileToFolder_(doc.getId(), destinationFolder);

  return assetResult_(DriveApp.getFileById(doc.getId()));
}

function getOrCreateQuizForm_(context, destinationFolder) {
  const fileName = context.documentId + " - Knowledge Quiz";
  const existing = getFirstFileByName_(destinationFolder, fileName);

  if (existing) {
    const existingForm = FormApp.openById(existing.getId());
    return assetResult_(existing, {
      editUrl: existingForm.getEditUrl(),
      publishedUrl: existingForm.getPublishedUrl()
    });
  }

  const form = FormApp.create(fileName);
  form.setDescription(
    "JPAIS knowledge assessment for " + context.documentTitle +
    ". Questions and answers must be verified before publication."
  );
  form.setIsQuiz(true);

  for (let i = 1; i <= 10; i++) {
    const item = form.addMultipleChoiceItem();
    item.setTitle(i + ". [Insert verified multiple-choice question]");
    item.setChoices([
      item.createChoice("Option A", true),
      item.createChoice("Option B", false),
      item.createChoice("Option C", false),
      item.createChoice("Option D", false)
    ]);
    item.setPoints(1);
    item.setRequired(true);
  }

  moveFileToFolder_(form.getId(), destinationFolder);
  return assetResult_(DriveApp.getFileById(form.getId()), {
    editUrl: form.getEditUrl(),
    publishedUrl: form.getPublishedUrl()
  });
}

function createPackageReadme_(context, packageInfo) {
  const content = [
    "# JPAIS Knowledge Package",
    "",
    "## Document Information",
    "",
    "- Document ID: " + context.documentId,
    "- Document Title: " + context.documentTitle,
    "- Category: " + context.category,
    "- Assigned Lead: " + (context.assignedLead || ""),
    "- Package Version: " + JPAIS.VERSION,
    "- Created: " + formatTimestamp_(new Date()),
    "",
    "## Purpose",
    "",
    "This folder is the controlled JPAIS workspace for the source document.",
    "",
    "## Working Rules",
    "",
    "- Use authoritative and verified sources.",
    "- Review all AI-generated content before publication.",
    "- Preserve consistent legal terminology.",
    "- Include citations and source links.",
    "- Do not delete approved resources.",
    "- Complete QA before publication.",
    "",
    "## Package Folder",
    "",
    packageInfo.documentFolder.getUrl()
  ].join("\n");

  const file = createTextFileIfMissing_(
    packageInfo.documentFolder,
    "README.md",
    content
  );
  return assetResult_(file);
}

function createQAChecklist_(context, packageInfo) {
  const content = [
    "# JPAIS Quality Assurance Checklist",
    "",
    "## Document",
    "- Document ID: " + context.documentId,
    "- Title: " + context.documentTitle,
    "- Assigned Lead: " + (context.assignedLead || ""),
    "",
    "## Source Verification",
    "- [ ] Official source uploaded",
    "- [ ] Title, number, year, and institution verified",
    "- [ ] Amendments and legal status checked",
    "",
    "## Legal Accuracy",
    "- [ ] Provisions represented accurately",
    "- [ ] Terminology is consistent",
    "- [ ] Quotations checked against source",
    "- [ ] Unsupported conclusions removed",
    "",
    "## AI Verification",
    "- [ ] Human review completed",
    "- [ ] Hallucinated facts or citations removed",
    "- [ ] Source content distinguished from analysis",
    "",
    "## Package Completeness",
    "- [ ] Original document",
    "- [ ] Executive summary",
    "- [ ] Audio overview",
    "- [ ] Presentation",
    "- [ ] Mind map",
    "- [ ] Infographic",
    "- [ ] FAQ",
    "- [ ] Quiz",
    "- [ ] Metadata",
    "- [ ] References",
    "",
    "## Approval",
    "- [ ] Self-review",
    "- [ ] Peer review",
    "- [ ] Project Lead review",
    "- [ ] Approved for publication"
  ].join("\n");

  const file = createTextFileIfMissing_(
    packageInfo.documentFolder,
    "QA_CHECKLIST.md",
    content
  );
  return assetResult_(file);
}

function createPackageManifest_(context, packageInfo, assets) {
  const rows = [
    ["Resource Type", "Resource Name", "Status", "URL", "Version"],
    ["Package Folder", packageInfo.documentFolder.getName(), "Available", packageInfo.documentFolder.getUrl(), JPAIS.VERSION],
    ["Original Document", "", "Pending", "", JPAIS.VERSION],
    ["Executive Summary", assets.executiveSummary.name, "Available", assets.executiveSummary.url, JPAIS.VERSION],
    ["Audio Overview", "", "Pending", "", JPAIS.VERSION],
    ["Presentation Slides", assets.presentation.name, "Available", assets.presentation.url, JPAIS.VERSION],
    ["Mind Map", "", "Pending", "", JPAIS.VERSION],
    ["Infographic", "", "Pending", "", JPAIS.VERSION],
    ["FAQ", assets.faq.name, "Available", assets.faq.url, JPAIS.VERSION],
    ["Quiz", assets.quiz.name, "Available", assets.quiz.publishedUrl, JPAIS.VERSION],
    ["Metadata", assets.metadata.name, "Available", assets.metadata.url, JPAIS.VERSION],
    ["References", assets.references.name, "Available", assets.references.url, JPAIS.VERSION]
  ];

  const csv = rows.map(function(row) {
    return row.map(csvEscape_).join(",");
  }).join("\n");

  const file = createOrReplaceTextFile_(
    packageInfo.documentFolder,
    "MANIFEST.csv",
    csv
  );
  return assetResult_(file);
}

function csvEscape_(value) {
  const text = value === null || value === undefined ? "" : String(value);
  return '"' + text.replace(/"/g, '""') + '"';
}

function appendMetadataTableToDoc_(body, context) {
  body.appendTable([
    ["Document ID", context.documentId],
    ["Category", context.category],
    ["Assigned Lead", context.assignedLead || ""],
    ["Version", JPAIS.VERSION],
    ["Status", "Draft"]
  ]);
}

function appendQualityNotice_(body) {
  body.appendParagraph("Quality Notice")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    "This is an AI-assisted working template. All legal statements, citations, and interpretations must be verified against authoritative sources before publication."
  ).setItalic(true);
}
