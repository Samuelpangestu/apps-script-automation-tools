/**
 * Menu.js — Menu Interface
 * ═══════════════════════════════════════════════════════════════════════
 * Simple menu for QA team management
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Create custom menu on open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('👥 QA Team Management')
    .addItem('🔨 Setup All Tabs', 'menuSetupAll')
    .addSeparator()
    .addItem('⚙️ Setup Config (Projects)', 'menuSetupConfig')
    .addItem('👤 Setup Team Members', 'menuSetupTeam')
    .addItem('📊 Create Dashboard', 'menuCreateDashboard')
    .addSeparator()
    .addItem('🔄 Refresh Dashboard Now', 'menuRefreshDashboard')
    .addItem('⏰ Configure Auto-Refresh', 'menuConfigureAutoRefresh')
    .addSeparator()
    .addItem('ℹ️ About', 'menuShowAbout')
    .addToUi();
}

/**
 * Setup all tabs at once
 */
function menuSetupAll() {
  try {
    // Create tabs sequentially with flush between each
    createConfigTab();
    Utilities.sleep(500); // Small delay to ensure completion

    createTeamMemberTab();
    Utilities.sleep(500);

    createDashboard();

    SpreadsheetApp.getActiveSpreadsheet().toast('All tabs created successfully!', 'Setup Complete', 3);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Setup failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error in setup: ' + error.message);
  }
}

/**
 * Setup Config tab only
 */
function menuSetupConfig() {
  try {
    createConfigTab();
    SpreadsheetApp.getActiveSpreadsheet().toast('Config - Projects tab created!', 'Success', 2);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Setup Team Members tab only
 */
function menuSetupTeam() {
  try {
    createTeamMemberTab();
    SpreadsheetApp.getActiveSpreadsheet().toast('Team Members tab created!', 'Success', 2);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Create dashboard
 */
function menuCreateDashboard() {
  try {
    createDashboard();
    SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard created!', 'Success', 2);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Refresh dashboard
 */
function menuRefreshDashboard() {
  try {
    createDashboard();
    SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard refreshed!', 'Success', 2);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to refresh: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show about dialog
 */
function menuShowAbout() {
  const ui = SpreadsheetApp.getUi();

  const message =
    '👥 QA TEAM MANAGEMENT SYSTEM\n\n' +
    '═══════════════════════════════\n\n' +
    'Version: 2.3.0\n' +
    'Author: QA Team\n\n' +
    '📋 KEY FEATURES\n\n' +
    '✅ Simple & clean interface\n' +
    '✅ Centralized project configuration\n' +
    '✅ Difficulty level definitions\n' +
    '✅ Copy-paste friendly team management\n' +
    '✅ Project distribution dashboard\n' +
    '✅ Scheduled auto-refresh dashboard\n' +
    '✅ Ready for KPI integration\n\n' +
    '═══════════════════════════════\n\n' +
    'TABS:\n\n' +
    '1. Config - Projects\n' +
    '   • Difficulty definitions\n' +
    '   • Project list (centralized)\n\n' +
    '2. Team Members\n' +
    '   • 6 simple columns\n' +
    '   • Copy-paste data directly\n\n' +
    '3. Dashboard\n' +
    '   • Team distribution per project\n' +
    '   • Auto-refresh available\n\n' +
    '═══════════════════════════════\n\n' +
    'WORKFLOW:\n\n' +
    '1. Menu → Setup All Tabs\n' +
    '2. Add/edit projects in Config\n' +
    '3. Add/paste team data\n' +
    '4. Assign projects (comma-separated)\n' +
    '5. View Dashboard\n\n' +
    'AUTO-REFRESH:\n\n' +
    'Menu → Configure Auto-Refresh\n' +
    'Choose: 1, 3, 6, 12, or 24 hours\n\n' +
    'DIFFICULTY LEVELS:\n\n' +
    '🟢 Easy: 1-2 QE\n' +
    '🟡 Medium: 1 PIC + 2-3 QE\n' +
    '🔴 Hard: 1 Lead + 1 PIC + 3-5 QE\n\n' +
    '═══════════════════════════════\n\n' +
    '© 2024 QA Department';

  ui.alert('About', message, ui.ButtonSet.OK);
}
