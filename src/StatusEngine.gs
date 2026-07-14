/**
 * JPAIS Automation Suite v1.0.1
 * Module 2.1 — Package Status Engine
 */
const JPAIS_STATUS = Object.freeze({
  PENDING: "Pending",
  IN_PROGRESS: "In Progress",
  COMPLETED: "Completed",
  REVIEWED: "Reviewed",
  APPROVED: "Approved",
  PUBLISHED: "Published",
  ARCHIVED: "Archived"
});

function createPackageStatusFile_(context, packageInfo, assets) {
  const status = buildInitialPackageStatus_(context, packageInfo, assets);
  const file = createOrReplaceTextFile_(packageInfo.documentFolder, "STATUS.json", JSON.stringify(status, null, 2));
  return {id:file.getId(), url:file.getUrl(), name:file.getName(), data:status};
}

function buildInitialPackageStatus_(context, packageInfo, assets) {
  const now = new Date();
  return {
    schema_version: "1.0",
    package: {
      document_id: context.documentId,
      title: context.documentTitle,
      category: context.category,
      version: "1.0",
      status: JPAIS_STATUS.IN_PROGRESS,
      owner: context.assignedLead || "",
      created_at: now.toISOString(),
      updated_at: now.toISOString(),
      package_folder_url: packageInfo.documentFolder.getUrl(),
      notebooklm_url: context.notebookLmUrl || ""
    },
    resources: {
      original_document: resourceStatus_("Original Document", JPAIS_STATUS.PENDING, ""),
      executive_summary: resourceStatus_("Executive Summary", assets.executiveSummary ? JPAIS_STATUS.COMPLETED : JPAIS_STATUS.PENDING, assets.executiveSummary ? assets.executiveSummary.url : ""),
      audio_overview: resourceStatus_("Audio Overview", JPAIS_STATUS.PENDING, ""),
      presentation_slides: resourceStatus_("Presentation Slides", assets.presentation ? JPAIS_STATUS.COMPLETED : JPAIS_STATUS.PENDING, assets.presentation ? assets.presentation.url : ""),
      mind_map: resourceStatus_("Mind Map", JPAIS_STATUS.PENDING, ""),
      infographic: resourceStatus_("Infographic", JPAIS_STATUS.PENDING, ""),
      faq: resourceStatus_("FAQ", assets.faq ? JPAIS_STATUS.COMPLETED : JPAIS_STATUS.PENDING, assets.faq ? assets.faq.url : ""),
      quiz: resourceStatus_("Quiz", assets.quiz ? JPAIS_STATUS.COMPLETED : JPAIS_STATUS.PENDING, assets.quiz ? assets.quiz.publishedUrl : ""),
      data_table: resourceStatus_("Data Table", assets.metadata ? JPAIS_STATUS.COMPLETED : JPAIS_STATUS.PENDING, assets.metadata ? assets.metadata.url : ""),
      references: resourceStatus_("References", assets.references ? JPAIS_STATUS.COMPLETED : JPAIS_STATUS.PENDING, assets.references ? assets.references.url : "")
    },
    workflow: {peer_review:JPAIS_STATUS.PENDING, lead_approval:JPAIS_STATUS.PENDING, publication:"Draft"},
    progress: {completed_resources:6, total_resources:10, completion_percentage:60},
    audit: {created_by:Session.getActiveUser().getEmail() || context.assignedLead || "", last_updated_by:Session.getActiveUser().getEmail() || context.assignedLead || "", last_action:"Package created"}
  };
}

function resourceStatus_(label, status, url) {
  return {label:label, status:status, url:url || "", owner:"", updated_at:new Date().toISOString()};
}

function updatePackageResourceStatus() {
  try {
    const context = getActiveProductionContext_();
    if (!context.driveFolderUrl) throw new Error("Drive Folder URL is empty. Create the Knowledge Package first.");
    const folder = DriveApp.getFolderById(extractDriveFolderId_(context.driveFolderUrl));
    const current = readPackageStatus_(folder);
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Update Resource Status", "Enter: resource_key, Status\nExample: audio_overview, Completed", ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;
    const parts = response.getResponseText().split(",");
    if (parts.length < 2) throw new Error("Use the format: resource_key, Status");
    const key = parts[0].trim();
    const value = parts.slice(1).join(",").trim();
    if (!current.resources[key]) throw new Error("Unknown resource key: " + key);
    if (Object.keys(JPAIS_STATUS).map(k=>JPAIS_STATUS[k]).indexOf(value) === -1) throw new Error("Invalid status: " + value);
    current.resources[key].status = value;
    current.resources[key].updated_at = new Date().toISOString();
    current.package.updated_at = new Date().toISOString();
    current.audit.last_updated_by = Session.getActiveUser().getEmail() || context.assignedLead || "";
    current.audit.last_action = "Updated " + key + " to " + value;
    recalculatePackageProgress_(current);
    writePackageStatus_(folder, current);
    syncStatusJsonToProductionBoard_(context, current);
    SpreadsheetApp.getActive().toast(key + " updated to " + value, "JPAIS Status Engine", 6);
  } catch (error) { handleError_("Status update failed", error); }
}

function rebuildPackageStatusFromBoard() {
  try {
    const context = getActiveProductionContext_();
    if (!context.driveFolderUrl) throw new Error("Drive Folder URL is empty.");
    const folder = DriveApp.getFolderById(extractDriveFolderId_(context.driveFolderUrl));
    const status = readPackageStatus_(folder);
    const mapping = {original_document:"Original PDF",executive_summary:"Executive Summary",audio_overview:"Podcast",presentation_slides:"Slides",mind_map:"Mind Map",infographic:"Infographic",faq:"FAQ",quiz:"Quiz",data_table:"Metadata",references:"References"};
    Object.keys(mapping).forEach(function(key){
      const value=getByHeader_(context.sheet,context.row,context.headerMap,mapping[key]);
      status.resources[key].status = value === "Not Started" ? JPAIS_STATUS.PENDING : (value || JPAIS_STATUS.PENDING);
      status.resources[key].updated_at = new Date().toISOString();
    });
    status.workflow.peer_review=getByHeader_(context.sheet,context.row,context.headerMap,"Peer Review")||JPAIS_STATUS.PENDING;
    status.workflow.lead_approval=getByHeader_(context.sheet,context.row,context.headerMap,"Lead Approval")||JPAIS_STATUS.PENDING;
    status.workflow.publication=getByHeader_(context.sheet,context.row,context.headerMap,"Publication Status")||"Draft";
    status.package.updated_at=new Date().toISOString();
    status.audit.last_action="Rebuilt from Production Board";
    recalculatePackageProgress_(status);
    writePackageStatus_(folder,status);
    SpreadsheetApp.getActive().toast("STATUS.json rebuilt from Production Board.","JPAIS Status Engine",6);
  } catch(error){handleError_("Status rebuild failed",error);}
}

function readPackageStatus_(folder){
  const files=folder.getFilesByName("STATUS.json");
  if(!files.hasNext()) throw new Error("STATUS.json was not found in the package folder.");
  return JSON.parse(files.next().getBlob().getDataAsString());
}
function writePackageStatus_(folder,obj){createOrReplaceTextFile_(folder,"STATUS.json",JSON.stringify(obj,null,2));}
function recalculatePackageProgress_(obj){
  const done=[JPAIS_STATUS.COMPLETED,JPAIS_STATUS.REVIEWED,JPAIS_STATUS.APPROVED,JPAIS_STATUS.PUBLISHED,JPAIS_STATUS.ARCHIVED];
  const keys=Object.keys(obj.resources); let c=0; keys.forEach(k=>{if(done.indexOf(obj.resources[k].status)!==-1)c++;});
  obj.progress.completed_resources=c; obj.progress.total_resources=keys.length; obj.progress.completion_percentage=keys.length?Math.round(c/keys.length*100):0;
}
function syncStatusJsonToProductionBoard_(context,obj){
  const map={original_document:"Original PDF",executive_summary:"Executive Summary",audio_overview:"Podcast",presentation_slides:"Slides",mind_map:"Mind Map",infographic:"Infographic",faq:"FAQ",quiz:"Quiz",data_table:"Metadata",references:"References"};
  Object.keys(map).forEach(k=>setByHeader_(context.sheet,context.row,context.headerMap,map[k],obj.resources[k].status));
  setByHeader_(context.sheet,context.row,context.headerMap,"Remarks","STATUS.json synchronized on "+formatTimestamp_(new Date()));
}
function extractDriveFolderId_(url){const m=String(url).match(/\/folders\/([a-zA-Z0-9_-]+)/); if(!m) throw new Error("Unable to extract folder ID from Drive Folder URL."); return m[1];}
function normalizeBoardStatus_(value) {
  const text = String(value || "").trim();

  if (!text || text === "Not Started") {
    return JPAIS_STATUS.PENDING;
  }

  return text;
}