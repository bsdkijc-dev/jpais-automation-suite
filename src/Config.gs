/**
 * JPAIS Automation Suite
 * Central configuration.
 */
const JPAIS = Object.freeze({
  VERSION: "1.1.0",
  ROOT_FOLDER_ID: "1ZS90tmExnkU2qlyEudY8ucpCnc853e8H",

  SHEETS: {
    PRODUCTION: "Production Board",
    DOCUMENT_REGISTRY: "Document Registry",
    ASSIGNMENTS: "Team Assignment",
    AI_TRACKER: "AI Resource Tracker",
    QA: "QA Review",
    PUBLISHER: "Google Sites Publisher"
  },

  MAIN_FOLDERS: {
    "Regulation": "01_Regulations",
    "Research": "02_Research",
    "Policy Brief": "03_Policy_Briefs",
    "Court Decision": "05_Repository",
    "Training": "06_Training",
    "International Cooperation": "07_Publication",
    "Technical Documentation": "05_Repository"
  },

  PACKAGE_FOLDERS: [
    "01_Original_Document",
    "02_Executive_Summary",
    "03_Audio_Overview",
    "04_Presentation_Slides",
    "05_Mind_Map",
    "06_Infographic",
    "07_FAQ",
    "08_Quiz",
    "09_Data_Table",
    "10_References"
  ],

  REQUIRED_ASSETS: [
    "executiveSummary",
    "presentation",
    "faq",
    "quiz",
    "metadata",
    "references",
    "readme",
    "qaChecklist",
    "manifest",
    "statusFile"
  ],

  REQUIRED_PRODUCTION_HEADERS: [
    "Document ID",
    "Document Title",
    "Category",
    "Assigned Lead",
    "Status",
    "Original PDF",
    "Executive Summary",
    "Podcast",
    "Slides",
    "Mind Map",
    "Infographic",
    "FAQ",
    "Quiz",
    "Metadata",
    "References",
    "Peer Review",
    "Lead Approval",
    "Publication Status",
    "Google Sites URL",
    "Drive Folder URL",
    "NotebookLM URL",
    "Remarks"
  ]
});
