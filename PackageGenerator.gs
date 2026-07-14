/**
 * Creates the full JPAIS Knowledge Package for the active Production Board row.
 *
 * Generated assets:
 * - Folder hierarchy
 * - Executive Summary Google Doc
 * - Presentation Google Slides
 * - Metadata Google Sheet
 * - FAQ Google Doc
 * - References Google Doc
 * - Quiz Google Form
 *
 * Podcast, Mind Map, and Infographic remain NotebookLM/creative outputs;
 * their folders are created and tracker fields remain available for links.
 */
function createCompleteKnowledgePackage() {
  const lock = LockService.getDocumentLock();

  try {
    lock.waitLock(30000);

    const context = getActiveProductionContext_();
    validateNoExistingPackageConflict_(context);

    const packageInfo = ensureKnowledgePackageFolders_(context);
    const assets = createPackageAssets_(context, packageInfo);

    updateProductionRowAfterCreation_(context, packageInfo, true, assets);
    upsertDocumentRegistry_(context, packageInfo, assets);
    upsertAIResourceTracker_(context, assets);
    upsertPublisherRegister_(context, packageInfo, assets);

    SpreadsheetApp.flush();
    SpreadsheetApp.getActive().toast(
      "Complete Knowledge Package created successfully.",
      "JPAIS",
      8
    );

    showPackageResultDialog_(context, packageInfo, assets);
  } catch (error) {
    handleError_("Knowledge Package creation failed", error);
  } finally {
    lock.releaseLock();
  }
}

function createPackageAssets_(context, packageInfo) {
  const assets = {};

  assets.executiveSummary = createExecutiveSummaryDoc_(
    context,
    packageInfo.subfolders["02_Executive_Summary"]
  );

  assets.presentation = createPresentation_(
    context,
    packageInfo.subfolders["04_Presentation_Slides"]
  );

  assets.faq = createFAQDoc_(
    context,
    packageInfo.subfolders["07_FAQ"]
  );

  assets.quiz = createQuizForm_(
    context,
    packageInfo.subfolders["08_Quiz"]
  );

  assets.metadata = createMetadataSheet_(
    context,
    packageInfo.subfolders["09_Data_Table"]
  );

  assets.references = createReferencesDoc_(
    context,
    packageInfo.subfolders["10_References"]
  );

  assets.statusFile = createPackageStatusFile_(
    context,
    packageInfo,
    assets
  );

  return assets;
}

function createExecutiveSummaryDoc_(context, destinationFolder) {
  const doc = DocumentApp.create(context.documentId + " - Executive Summary");
  const body = doc.getBody();

  body.appendParagraph("JPAIS EXECUTIVE SUMMARY")
      .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(context.documentTitle)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  appendMetadataTableToDoc_(body, context);

  [
    ["1. Background", "Describe the institutional and legal background of the document."],
    ["2. Purpose", "State the document's principal objectives."],
    ["3. Legal Basis", "List the relevant constitutional, statutory, regulatory, and institutional authorities."],
    ["4. Key Provisions or Findings", "Summarize the most important provisions, findings, or recommendations."],
    ["5. Implementation Implications", "Explain implications for courts, judges, officials, researchers, and other stakeholders."],
    ["6. Risks and Considerations", "Identify legal, operational, institutional, budgetary, and implementation risks."],
    ["7. Conclusion", "Provide a concise evidence-based conclusion."]
  ].forEach(function(section) {
    body.appendParagraph(section[0]).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(section[1]);
  });

  body.appendParagraph("Quality Notice")
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    "This document is an AI-assisted working template. All legal statements, citations, and interpretations must be verified against official sources before publication."
  );

  doc.saveAndClose();
  moveFileToFolder_(doc.getId(), destinationFolder);

  return { id: doc.getId(), url: doc.getUrl(), name: doc.getName() };
}

function createPresentation_(context, destinationFolder) {
  const presentation = SlidesApp.create(context.documentId + " - Presentation");
  const slides = presentation.getSlides();

  const titleSlide = slides[0];
  titleSlide.getShapes()[0].getText().setText(context.documentTitle);
  if (titleSlide.getShapes().length > 1) {
    titleSlide.getShapes()[1].getText().setText(
      "JPAIS Knowledge Package\nDocument ID: " + context.documentId
    );
  }

  const sections = [
    ["Overview", "Purpose, scope, and institutional context."],
    ["Legal Basis", "Relevant laws, regulations, policies, and institutional authority."],
    ["Key Provisions or Findings", "Core provisions, issues, findings, or recommendations."],
    ["Implementation", "Institutional roles, procedures, resources, and timelines."],
    ["Risks and Mitigation", "Legal, operational, data, security, and implementation risks."],
    ["Conclusion", "Principal conclusions and next actions."]
  ];

  sections.forEach(function(section) {
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    const placeholders = slide.getPlaceholders();
    placeholders[0].asShape().getText().setText(section[0]);
    if (placeholders.length > 1) {
      placeholders[1].asShape().getText().setText(section[1]);
    }
  });

  moveFileToFolder_(presentation.getId(), destinationFolder);
  return {
    id: presentation.getId(),
    url: presentation.getUrl(),
    name: presentation.getName()
  };
}

function createMetadataSheet_(context, destinationFolder) {
  const spreadsheet = SpreadsheetApp.create(context.documentId + " - Metadata");
  const sheet = spreadsheet.getSheets()[0];
  sheet.setName("Metadata");

  const rows = [
    ["Field", "Value"],
    ["Document ID", context.documentId],
    ["Official Title", context.documentTitle],
    ["Category", context.category],
    ["Assigned Lead", context.assignedLead],
    ["Status", context.status || "In Progress"],
    ["Language", ""],
    ["Issuing Institution", ""],
    ["Document Number", ""],
    ["Year", ""],
    ["Subject Area", ""],
    ["Keywords", ""],
    ["Legal Status", ""],
    ["Source URL", ""],
    ["NotebookLM URL", context.notebookLmUrl || ""],
    ["Date Created", new Date()],
    ["Version", "1.0"],
    ["Access Level", "Internal"],
    ["Review Status", "Pending"]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange("A1:B1")
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
  sheet.getRange("A1:B" + rows.length).setWrap(true);
  sheet.autoResizeColumns(1, 2);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 420);
  sheet.setFrozenRows(1);

  moveFileToFolder_(spreadsheet.getId(), destinationFolder);
  return {
    id: spreadsheet.getId(),
    url: spreadsheet.getUrl(),
    name: spreadsheet.getName()
  };
}

function createFAQDoc_(context, destinationFolder) {
  const doc = DocumentApp.create(context.documentId + " - FAQ");
  const body = doc.getBody();
  body.appendParagraph("FREQUENTLY ASKED QUESTIONS")
      .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(context.documentTitle)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  for (let i = 1; i <= 10; i++) {
    body.appendParagraph(i + ". [Insert verified question]")
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("[Insert concise, source-based answer and citation.]");
  }

  body.appendParagraph(
    "All answers must be verified against official sources before publication."
  ).setItalic(true);

  doc.saveAndClose();
  moveFileToFolder_(doc.getId(), destinationFolder);
  return { id: doc.getId(), url: doc.getUrl(), name: doc.getName() };
}

function createReferencesDoc_(context, destinationFolder) {
  const doc = DocumentApp.create(context.documentId + " - References");
  const body = doc.getBody();
  body.appendParagraph("REFERENCES AND RELATED DOCUMENTS")
      .setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(context.documentTitle)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  [
    "Official source document",
    "Related laws and regulations",
    "Related PERMA, SEMA, and SK KMA",
    "Court decisions",
    "Judicial research and policy briefs",
    "Academic references",
    "International standards and comparative materials"
  ].forEach(function(item) {
    body.appendListItem(item + ": [Insert verified citation and URL]");
  });

  doc.saveAndClose();
  moveFileToFolder_(doc.getId(), destinationFolder);
  return { id: doc.getId(), url: doc.getUrl(), name: doc.getName() };
}

function createQuizForm_(context, destinationFolder) {
  const form = FormApp.create(context.documentId + " - Knowledge Quiz");
  form.setDescription(
    "JPAIS knowledge assessment for " + context.documentTitle +
    ". Questions and answers must be legally verified before publication."
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
  return {
    id: form.getId(),
    editUrl: form.getEditUrl(),
    publishedUrl: form.getPublishedUrl(),
    name: form.getTitle()
  };
}

function appendMetadataTableToDoc_(body, context) {
  const table = body.appendTable([
    ["Document ID", context.documentId],
    ["Category", context.category],
    ["Assigned Lead", context.assignedLead || ""],
    ["Version", "1.0"],
    ["Status", "Draft"]
  ]);
  table.getRow(0).getCell(0).setBackgroundColor("#D9EAF7");
}

function moveFileToFolder_(fileId, destinationFolder) {
  const file = DriveApp.getFileById(fileId);
  file.moveTo(destinationFolder);
}
