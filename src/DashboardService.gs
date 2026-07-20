/**
 * DashboardService.gs
 * Helper untuk memperbarui Dashboard secara aman.
 */

function refreshDashboardSafe_() {
  try {
    refreshDashboard();
  } catch (error) {
    Logger.log("Dashboard refresh skipped: " + error);
  }
}