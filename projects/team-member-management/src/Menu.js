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
    .addItem('🌐 Configure Production Sync', 'menuConfigureProductionSync')
    .addItem('🔄 Sync from Production Now', 'menuSyncFromProduction')
    .addItem('🧪 Test Sync Connection', 'menuTestSyncConnection')
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
    '1. Menu → Setup All Tabs\n' +
    '2. (Optional) Configure Production Sync\n' +
    '3. Edit project/modul/submodul in Config\n' +
    '4. Add/paste team data\n' +
    '5. Assign projects/modul/submodul\n' +
    '6. View Dashboard\n\n' +
    'PRODUCTION SYNC:\n\n' +
    'Menu → Configure Production Sync\n' +
    'Enter production spreadsheet ID\n' +
    'Use "Sync from Production Now" to pull data\n\n' +
    'AUTO SYNC & REFRESH:\n\n' +
    'Menu → Configure Auto-Refresh\n' +
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
 * Configure production sync
 */
function menuConfigureProductionSync() {
  const ui = SpreadsheetApp.getUi();
  const currentSettings = getSyncSettings();

  let message = 'PRODUCTION SYNC CONFIGURATION\n\n';

  if (currentSettings.enabled) {
    message += 'Current Status: ENABLED ✅\n';
    message += 'Production Spreadsheet: ' + currentSettings.spreadsheetName + '\n';
    message += 'Spreadsheet ID: ' + currentSettings.spreadsheetId + '\n\n';
    message += 'Do you want to change or disable production sync?';
  } else {
    message += 'Current Status: DISABLED (Local mode) 📍\n\n';
    message += 'Do you want to enable production sync?';
  }

  const response = ui.alert(
    'Configure Production Sync',
    message,
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.YES) {
    // Configure new production spreadsheet
    const promptResponse = ui.prompt(
      'Production Spreadsheet ID',
      'Enter the ID of your production spreadsheet:\n\n' +
      '(Find it in the URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit)',
      ui.ButtonSet.OK_CANCEL
    );

    if (promptResponse.getSelectedButton() === ui.Button.OK) {
      const spreadsheetId = promptResponse.getResponseText().trim();

      if (!spreadsheetId) {
        ui.alert('Error', 'Spreadsheet ID cannot be empty', ui.ButtonSet.OK);
        return;
      }

      // Test connection
      try {
        const testSS = SpreadsheetApp.openById(spreadsheetId);
        const testConfig = testSS.getSheetByName(CONFIG_TAB_NAME);

        if (!testConfig) {
          ui.alert('Error', 'Config tab not found in the specified spreadsheet.\n\nMake sure the production spreadsheet has a "Config" tab.', ui.ButtonSet.OK);
          return;
        }

        const spreadsheetName = testSS.getName();

        // Save settings
        saveSyncSettings(spreadsheetId, spreadsheetName);

        ui.alert(
          'Success! ✅',
          'Production sync configured successfully.\n\n' +
          'Production: ' + spreadsheetName + '\n\n' +
          'You can now use "Sync from Production Now" to pull data.',
          ui.ButtonSet.OK
        );

      } catch (error) {
        ui.alert('Error', 'Failed to connect to production spreadsheet:\n\n' + error.message + '\n\nPlease check:\n1. Spreadsheet ID is correct\n2. You have access to the spreadsheet', ui.ButtonSet.OK);
      }
    }

  } else if (response === ui.Button.NO) {
    // Disable
    if (currentSettings.enabled) {
      disableSyncSettings();
      ui.alert('Success', 'Production sync has been disabled.\n\nDashboard will use local data only.', ui.ButtonSet.OK);
    } else {
      ui.alert('Info', 'Production sync is already disabled.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Sync from production now
 */
function menuSyncFromProduction() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

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
      ui.alert('Sync Failed', result.message, ui.ButtonSet.OK);
    }

  } catch (error) {
    ui.alert('Error', 'Sync failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Test sync connection
 */
function menuTestSyncConnection() {
  const ui = SpreadsheetApp.getUi();

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
        'Please check your production sync configuration.',
        ui.ButtonSet.OK
      );
    }

  } catch (error) {
    ui.alert('Error', 'Test failed: ' + error.message, ui.ButtonSet.OK);
  }
}
