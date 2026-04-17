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
    .addItem('⚙️ Setup Config', 'menuSetupConfig')
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Delete all existing tabs first to avoid conflicts
    let configSheet = ss.getSheetByName(CONFIG_TAB_NAME);
    if (configSheet) ss.deleteSheet(configSheet);

    let teamSheet = ss.getSheetByName(TEAM_TAB_NAME);
    if (teamSheet) ss.deleteSheet(teamSheet);

    let dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    if (dashSheet) ss.deleteSheet(dashSheet);

    SpreadsheetApp.flush();
    Utilities.sleep(1000); // Wait for deletes to complete

    // Create tabs sequentially with flush between each
    createConfigTab();
    Utilities.sleep(500);

    createTeamMemberTab();
    Utilities.sleep(500);

    createDashboard();

    ss.toast('All tabs created successfully!', 'Setup Complete', 3);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Setup failed: ' + error.message + '\n\nCheck View > Logs for details', SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error in setup: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
  }
}

/**
 * Setup Config tab only
 */
function menuSetupConfig() {
  try {
    createConfigTab();
    SpreadsheetApp.getActiveSpreadsheet().toast('Config tab created!', 'Success', 2);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Setup Team Members tab only
 */
function menuSetupTeam() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // Delete existing if any
    let existingSheet = ss.getSheetByName(TEAM_TAB_NAME);
    if (existingSheet) ss.deleteSheet(existingSheet);
    SpreadsheetApp.flush();

    createTeamMemberTab();
    ss.toast('Team Members tab created!', 'Success', 2);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Create dashboard
 */
function menuCreateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // Delete existing if any
    let existingSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    if (existingSheet) ss.deleteSheet(existingSheet);
    SpreadsheetApp.flush();

    createDashboard();
    ss.toast('Dashboard created!', 'Success', 2);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Refresh dashboard
 */
function menuRefreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // Delete existing dashboard
    let existingSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    if (existingSheet) ss.deleteSheet(existingSheet);
    SpreadsheetApp.flush();

    createDashboard();
    ss.toast('Dashboard refreshed!', 'Success', 2);
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
    'Version: 3.0.0\n' +
    'Author: QA Team\n\n' +
    '📋 KEY FEATURES\n\n' +
    '✅ Simple & clean interface\n' +
    '✅ Centralized configuration (all in one place)\n' +
    '✅ Modul & Submodul structure\n' +
    '✅ Difficulty level definitions\n' +
    '✅ Copy-paste friendly team management\n' +
    '✅ Submodul distribution dashboard\n' +
    '✅ Scheduled auto-refresh dashboard\n' +
    '✅ Ready for KPI integration\n\n' +
    '═══════════════════════════════\n\n' +
    'TABS:\n\n' +
    '1. Config\n' +
    '   • Difficulty definitions\n' +
    '   • Modul list (INADigital, SIPGN)\n' +
    '   • Submodul list (centralized)\n\n' +
    '2. Team Members\n' +
    '   • 7 columns\n' +
    '   • Copy-paste data directly\n' +
    '   • Modul + Submodul assignment\n\n' +
    '3. Dashboard\n' +
    '   • Team distribution per submodul\n' +
    '   • Grouped by modul\n' +
    '   • Auto-refresh available\n\n' +
    '═══════════════════════════════\n\n' +
    'STRUCTURE:\n\n' +
    'Modul → Submodul → Team Members\n\n' +
    'Example:\n' +
    'INADigital → INAgov → Samuel, Irvan\n' +
    'SIPGN → Core System → Samuel\n\n' +
    'WORKFLOW:\n\n' +
    '1. Menu → Setup All Tabs\n' +
    '2. Edit modul/submodul in Config\n' +
    '3. Add/paste team data\n' +
    '4. Assign modul & submodul\n' +
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
