/**
 * Triggers.js — Auto-Refresh Dashboard
 * ═══════════════════════════════════════════════════════════════════════
 * Simple auto-refresh trigger for dashboard
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Auto refresh dashboard (called by time-based trigger)
 * This function recreates the dashboard to show latest data
 */
function autoRefreshDashboard() {
  try {
    Logger.log('🔄 Auto-refresh started...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);

    if (dashSheet) {
      ss.deleteSheet(dashSheet);
      SpreadsheetApp.flush();
    }

    createDashboard();
    Logger.log('✅ Dashboard refreshed successfully');

  } catch (error) {
    Logger.log('❌ Auto-refresh error: ' + error.message);
  }
}
