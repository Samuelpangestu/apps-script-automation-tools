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
    .addItem('🔄 Sync from Production & Refresh', 'menuSyncFromProduction')
    .addItem('👥 Generate Team Members from Config', 'menuGenerateTeamMembers')
    .addItem('🧪 Test Production Connection', 'menuTestSyncConnection')
    .addItem('⏰ Configure Auto-Sync', 'menuConfigureAutoRefresh')
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
    'Version: 4.0.0\n' +
    'Author: QA Team\n\n' +
    '📋 KEY FEATURES\n\n' +
    '✅ Simple & clean interface\n' +
    '✅ 3-level hierarchy (Project → Modul → Submodul)\n' +
    '✅ Numerical ratings (1-10) for Difficulty, Risk, Complexity\n' +
    '✅ Centralized configuration (all in one place)\n' +
    '✅ Copy-paste friendly team management\n' +
    '✅ Submodul distribution dashboard with color coding\n' +
    '✅ Production sync (pull data from production spreadsheet)\n' +
    '✅ Scheduled auto-sync & refresh dashboard\n' +
    '✅ Dashboard status monitoring in Config tab\n' +
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
    'LOCAL MODE (No Production):\n' +
    '1. Menu → Setup All Tabs\n' +
    '2. Edit project/modul/submodul in Config\n' +
    '3. Add/paste team data\n' +
    '4. View Dashboard\n\n' +
    'PRODUCTION SYNC MODE:\n' +
    '1. Menu → Setup All Tabs\n' +
    '2. Open Config tab → Cell B3\n' +
    '3. Paste Production Spreadsheet ID or URL\n' +
    '4. Menu → Sync from Production & Refresh\n' +
    '5. (Optional) Menu → Configure Auto-Sync\n\n' +
    'PRODUCTION SYNC:\n\n' +
    'Simple 3-step process:\n' +
    '1. Config tab → Paste ID in cell B3\n' +
    '2. Menu → Test Production Connection ✓\n' +
    '3. Menu → Sync from Production & Refresh\n\n' +
    'AUTO-SYNC:\n\n' +
    'Menu → Configure Auto-Sync\n' +
    'Choose: 1, 3, 6, 12, or 24 hours\n' +
    'System will auto-sync + refresh dashboard\n\n' +
    'DIFFICULTY LEVELS:\n\n' +
    '🟢 Easy: 1-2 QE\n' +
    '🟡 Medium: 1 PIC + 2-3 QE\n' +
    '🔴 Hard: 1 Lead + 1 PIC + 3-5 QE\n\n' +
    '═══════════════════════════════\n\n' +
    '© 2024 QA Department';

  ui.alert('About', message, ui.ButtonSet.OK);
}

/**
 * Sync from production now (simplified - reads ID from Config tab)
 */
function menuSyncFromProduction() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if Production ID is configured in Config tab
  const prodId = getProductionIdFromConfig();

  if (!prodId) {
    ui.alert(
      'Production Not Configured',
      'Please paste your Production Spreadsheet ID in the Config tab first.\n\n' +
      'Location: Config tab → Cell B3 (Production Spreadsheet ID)\n\n' +
      'You can paste either:\n' +
      '• Just the ID: 1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY\n' +
      '• Or full URL: https://docs.google.com/spreadsheets/d/ID/edit',
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const result = syncFromProduction();

    if (result.success) {
      ui.alert(
        'Sync Complete! ✅',
        result.message + '\n\n' +
        'Synced:\n' +
        '• Projects: ' + result.synced.projects + '\n' +
        '• Moduls: ' + result.synced.moduls + '\n' +
        '• Submoduls: ' + result.synced.submoduls + '\n\n' +
        'Dashboard will be refreshed automatically.',
        ui.ButtonSet.OK
      );

      // Refresh dashboard after sync
      let dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
      if (dashSheet) ss.deleteSheet(dashSheet);
      SpreadsheetApp.flush();

      createDashboard();
      ss.toast('Dashboard refreshed with synced data!', 'Success', 3);

    } else {
      ui.alert('Sync Failed', result.message + '\n\nPlease check:\n1. Production Spreadsheet ID is correct in Config tab (cell B3)\n2. You have access to the production spreadsheet', ui.ButtonSet.OK);
    }

  } catch (error) {
    ui.alert('Error', 'Sync failed: ' + error.message + '\n\nPlease verify the Production Spreadsheet ID in Config tab (cell B3).', ui.ButtonSet.OK);
  }
}

/**
 * Test sync connection (reads from Config tab)
 */
function menuTestSyncConnection() {
  const ui = SpreadsheetApp.getUi();

  // Check if Production ID is configured in Config tab
  const prodId = getProductionIdFromConfig();

  if (!prodId) {
    ui.alert(
      'Production Not Configured',
      'Please paste your Production Spreadsheet ID in the Config tab first.\n\n' +
      'Location: Config tab → Cell B3 (Production Spreadsheet ID)',
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const result = testSyncConnection();

    if (result.success) {
      ui.alert(
        'Connection Test: SUCCESS ✅',
        result.message + '\n\n' +
        'Production spreadsheet is accessible and has a Config tab.\n\n' +
        'You can safely sync data from this source.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Connection Test: FAILED ❌',
        result.message + '\n\n' +
        'Please check:\n' +
        '1. Production Spreadsheet ID in Config tab (cell B3)\n' +
        '2. You have access to the production spreadsheet',
        ui.ButtonSet.OK
      );
    }

  } catch (error) {
    ui.alert('Error', 'Test failed: ' + error.message + '\n\nPlease verify the Production Spreadsheet ID in Config tab.', ui.ButtonSet.OK);
  }
}

/**
 * Generate Team Members from Config
 */
function menuGenerateTeamMembers() {
  const ui = SpreadsheetApp.getUi();

  try {
    const result = generateTeamMembersFromConfig();

    if (result.success) {
      ui.alert(
        'Team Members Generated! ✅',
        result.message + '\n\n' +
        'Generated team members from PIC QA column in Config tab.\n\n' +
        'Notes:\n' +
        '• Multiple PICs per modul (e.g., "Adinda, Denta") are split automatically\n' +
        '• Default role is "Quality Engineer" - update manually if needed\n' +
        '• Email field is empty - please fill manually\n' +
        '• Projects, Moduls, and Submoduls are auto-aggregated per person',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Generation Failed', result.message, ui.ButtonSet.OK);
    }

  } catch (error) {
    ui.alert('Error', 'Failed to generate team members: ' + error.message, ui.ButtonSet.OK);
  }
}
