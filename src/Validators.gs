/**
 * ==========================================================
 * Validators.gs
 * JPAIS Automation Suite v1.2.0
 * ==========================================================
 */

/**
 * Menjalankan seluruh validasi pada satu baris Production Board.
 *
 * @param {Object} context Context dari getActiveProductionContext_().
 * @return {Object} Hasil validasi.
 */
function validateProductionContext_(context) {
  const errors = [];
  const warnings = [];

  validateRequiredValue_(
    context.documentId,
    "Document ID",
    errors
  );

  validateRequiredValue_(
    context.documentTitle,
    "Document Title",
    errors
  );

  validateRequiredValue_(
    context.category,
    "Category",
    errors
  );

  validateRequiredValue_(
    context.assignedLead,
    "Assigned Lead",
    errors
  );

  validateDocumentIdFormat_(
    context.documentId,
    errors
  );

  validateDocumentIdUnique_(
    context,
    errors
  );

  validateCategory_(
    context.category,
    errors
  );

  validatePriority_(
    context.priority,
    warnings
  );

  validateWorkflowStatus_(
    context.workflowStatus,
    warnings
  );

  validateDriveFolderUrl_(
    context.driveFolderUrl,
    warnings
  );

  return {
    valid: errors.length === 0,
    errors: errors,
    warnings: warnings
  };
}

/**
 * Memeriksa nilai wajib.
 *
 * @param {*} value Nilai yang diperiksa.
 * @param {string} label Nama field.
 * @param {string[]} errors Daftar error.
 */
function validateRequiredValue_(value, label, errors) {
  if (!String(value || "").trim()) {
    errors.push(label + " wajib diisi.");
  }
}

/**
 * Memeriksa format Document ID.
 *
 * Contoh valid:
 * REG-TEST-002
 * PERMA-001-2016
 *
 * @param {string} documentId
 * @param {string[]} errors
 */
function validateDocumentIdFormat_(documentId, errors) {
  const value = String(documentId || "").trim();

  if (!value) return;

  const pattern = /^[A-Z0-9]+(?:-[A-Z0-9]+)*$/;

  if (!pattern.test(value)) {
    errors.push(
      "Document ID hanya boleh menggunakan huruf kapital, angka, dan tanda hubung."
    );
  }
}

/**
 * Memeriksa apakah Document ID unik pada Production Board.
 *
 * @param {Object} context
 * @param {string[]} errors
 */
function validateDocumentIdUnique_(context, errors) {
  if (!context.documentId) return;

  const sheet = context.sheet;
  const documentIdColumn = context.headerMap["Document ID"];

  if (!documentIdColumn) return;

  const firstDataRow = 5;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstDataRow) return;

  const values = sheet
    .getRange(
      firstDataRow,
      documentIdColumn,
      lastRow - firstDataRow + 1,
      1
    )
    .getDisplayValues();

  const duplicateRows = [];

  values.forEach(function(rowValue, index) {
    const sheetRow = firstDataRow + index;
    const value = String(rowValue[0] || "").trim();

    if (
      value === context.documentId &&
      sheetRow !== context.row
    ) {
      duplicateRows.push(sheetRow);
    }
  });

  if (duplicateRows.length) {
    errors.push(
      "Document ID sudah digunakan pada baris: " +
      duplicateRows.join(", ")
    );
  }
}

/**
 * Memeriksa kategori.
 *
 * @param {string} category
 * @param {string[]} errors
 */
function validateCategory_(category, errors) {
  const allowed = Object.keys(JPAIS.MAIN_FOLDERS);
  const value = String(category || "").trim();

  if (!value) return;

  if (allowed.indexOf(value) === -1) {
    errors.push(
      "Category tidak valid. Pilihan yang diizinkan: " +
      allowed.join(", ")
    );
  }
}

/**
 * Memeriksa priority.
 *
 * @param {string} priority
 * @param {string[]} warnings
 */
function validatePriority_(priority, warnings) {
  const allowed = [
    "Low",
    "Medium",
    "High",
    "Critical"
  ];

  const value = String(priority || "").trim();

  if (!value) {
    warnings.push("Priority belum diisi.");
    return;
  }

  if (allowed.indexOf(value) === -1) {
    warnings.push(
      "Priority tidak dikenali: " + value
    );
  }
}

/**
 * Memeriksa workflow status.
 *
 * @param {string} status
 * @param {string[]} warnings
 */
function validateWorkflowStatus_(status, warnings) {
  const value = String(status || "").trim();

  if (!value) {
    warnings.push("Workflow Status belum diisi.");
    return;
  }

  const allowed = Object.keys(WORKFLOW_STATUS).map(function(key) {
    return WORKFLOW_STATUS[key];
  });

  if (allowed.indexOf(value) === -1) {
    warnings.push(
      "Workflow Status tidak valid: " + value
    );
  }
}

/**
 * Memeriksa URL folder Google Drive.
 *
 * URL kosong diperbolehkan karena package mungkin belum dibuat.
 *
 * @param {string} url
 * @param {string[]} warnings
 */
function validateDriveFolderUrl_(url, warnings) {
  const value = String(url || "").trim();

  if (!value) return;

  const valid =
    /^https:\/\/drive\.google\.com\/drive\/folders\/[A-Za-z0-9_-]+/.test(value);

  if (!valid) {
    warnings.push(
      "Drive Folder URL bukan URL folder Google Drive yang valid."
    );
  }
}