/**
 * ==========================================================
 * Quality.gs
 * JPAIS Automation Suite v1.2.0
 * ==========================================================
 */

/**
 * Memeriksa kelengkapan aset Knowledge Package.
 *
 * @param {Object} assets Hasil dari createPackageAssets_().
 * @return {Object} Hasil pemeriksaan kualitas.
 */
function evaluatePackageQuality_(assets) {
  const checks = [
    checkQualityAsset_("Executive Summary", assets.executiveSummary),
    checkQualityAsset_("Presentation", assets.presentation),
    checkQualityAsset_("FAQ", assets.faq),
    checkQualityAsset_("Quiz", assets.quiz),
    checkQualityAsset_("Metadata", assets.metadata),
    checkQualityAsset_("References", assets.references),
    checkQualityAsset_("README", assets.readme),
    checkQualityAsset_("QA Checklist", assets.qaChecklist),
    checkQualityAsset_("Manifest", assets.manifest),
    checkQualityAsset_("STATUS.json", assets.statusFile)
  ];

  const failed = checks.filter(function(check) {
    return !check.passed;
  });

  return {
    passed: failed.length === 0,
    total: checks.length,
    passedCount: checks.length - failed.length,
    failedCount: failed.length,
    checks: checks,
    missing: failed.map(function(check) {
      return check.name;
    })
  };
}

/**
 * Memeriksa satu aset.
 *
 * @param {string} name Nama aset.
 * @param {Object} asset Objek hasil asset.
 * @return {Object}
 */
function checkQualityAsset_(name, asset) {
  const passed = Boolean(
    asset &&
    asset.id &&
    asset.url
  );

  return {
    name: name,
    passed: passed,
    fileId: passed ? asset.id : "",
    url: passed ? asset.url : ""
  };
}