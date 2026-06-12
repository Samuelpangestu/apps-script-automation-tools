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
    .addItem('🔨 Create All Tabs', 'menuSetupAll')
    .addSeparator()
    .addItem('⚡ Run All (Sync + Generate + Refresh)', 'menuRunAll')
    .addItem('🔄 Sync from QA Dashboard', 'menuSyncFromDashboard')
    .addItem('👥 Generate Team Members from PIC', 'menuGenerateTeamMembers')
    .addItem('📊 Refresh Dashboard', 'menuRefreshDashboard')
    .addSeparator()
    .addItem('⏰ Setup Auto-Sync & Refresh', 'menuSetupAutoSyncRefresh')
    .addSeparator()
    .addItem('📖 Parameter Guide', 'menuShowParameterGuide')
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
 * Run All: Sync + Generate + Refresh (End-to-End)
 */
function menuRunAll() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.alert(
    'Run All Tasks?',
    'This will execute all tasks in sequence:\n\n' +
    '1️⃣ Sync from QA Dashboard\n' +
    '2️⃣ Generate Team Members from PIC\n' +
    '3️⃣ Refresh Dashboard\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    ss.toast('Running all tasks...', 'In Progress', -1);

    // Step 1: Sync from QA Dashboard
    ss.toast('Step 1/3: Syncing from QA Dashboard...', 'In Progress', -1);
    const syncResult = syncFromDashboard();

    if (!syncResult.success) {
      ui.alert(
        'Sync Failed',
        'Step 1 failed: ' + syncResult.message + '\n\nProcess stopped.',
        ui.ButtonSet.OK
      );
      return;
    }

    ss.toast('Step 1/3: ✅ Synced ' + syncResult.synced.submoduls + ' modules', 'Success', 2);
    Utilities.sleep(1000);

    // Step 2: Generate Team Members
    ss.toast('Step 2/3: Generating Team Members from PIC...', 'In Progress', -1);
    const generateResult = generateTeamMembersFromConfig();

    if (!generateResult.success) {
      ui.alert(
        'Generate Failed',
        'Step 2 failed: ' + generateResult.message + '\n\nProcess stopped.',
        ui.ButtonSet.OK
      );
      return;
    }

    ss.toast('Step 2/3: ✅ Generated ' + generateResult.generated + ' team members', 'Success', 2);
    Utilities.sleep(1000);

    // Step 3: Refresh Dashboard
    ss.toast('Step 3/3: Refreshing Dashboard...', 'In Progress', -1);
    const dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    if (dashSheet) ss.deleteSheet(dashSheet);
    SpreadsheetApp.flush();

    createDashboard();
    ss.toast('Step 3/3: ✅ Dashboard refreshed', 'Success', 2);
    Utilities.sleep(1000);

    // Success summary
    ui.alert(
      'All Tasks Completed! ✅',
      '1️⃣ Synced: ' + syncResult.synced.submoduls + ' modules\n' +
      '2️⃣ Generated: ' + generateResult.generated + ' team members\n' +
      '3️⃣ Dashboard refreshed\n\n' +
      'All tasks completed successfully!',
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert('Error', 'Run All failed: ' + error.message, ui.ButtonSet.OK);
    Logger.log('Run All error: ' + error.stack);
  }
}

/**
 * Generate Team Members from Project Config (PIC QA)
 */
function menuGenerateTeamMembers() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.alert(
    'Generate Team Members from PIC?',
    'This will auto-generate Team Members based on PIC QA from Project Config.\n\n' +
    '✅ Manual data (NP, Email, etc.) will be preserved.\n' +
    '✅ Only auto-generated fields (Name, Projects, Modules) will be updated.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    ss.toast('Generating Team Members from PIC QA...', 'In Progress', -1);

    const result = generateTeamMembersFromConfig();

    if (result.success) {
      ss.toast('Generated ' + result.generated + ' team members from PIC QA!', 'Success', 5);

      ui.alert(
        'Team Members Generated! ✅',
        'Successfully generated ' + result.generated + ' team members.\n\n' +
        'Please fill in additional details:\n' +
        '• NP (Employee Number)\n' +
        '• Email\n' +
        '• Phone\n' +
        '• Join Date\n' +
        '• Title & Role',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Generation Failed', result.message, ui.ButtonSet.OK);
    }

  } catch (error) {
    ui.alert('Error', 'Failed to generate: ' + error.message, ui.ButtonSet.OK);
    Logger.log('Generate error: ' + error.stack);
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
 * Show parameter guide
 */
function menuShowParameterGuide() {
  const ui = SpreadsheetApp.getUi();

  const message =
    '📖 PARAMETER GUIDE FOR PROJECT CONFIGURATION\n\n' +
    '═══════════════════════════════════════════════════════\n\n' +
    'Gunakan skala 1-10 untuk setiap parameter:\n' +
    '• 1-3 = Low (rendah)\n' +
    '• 4-6 = Medium (sedang)\n' +
    '• 7-10 = High (tinggi)\n\n' +
    '═══════════════════════════════════════════════════════\n\n' +
    '🔷 DIFFICULTY (Tingkat Kesulitan)\n' +
    'Seberapa sulit modul ini untuk di-test?\n\n' +
    '1-3: Modul sederhana, UI straightforward\n' +
    '4-6: Modul dengan business logic moderate\n' +
    '7-10: Modul kompleks, banyak edge cases\n\n' +
    '═══════════════════════════════════════════════════════\n\n' +
    '🔷 RISK (Tingkat Risiko)\n' +
    'Seberapa besar dampak jika ada bug?\n\n' +
    '1-3: Low impact, non-critical features\n' +
    '4-6: Medium impact, affects user experience\n' +
    '7-10: High impact, critical business function\n\n' +
    '═══════════════════════════════════════════════════════\n\n' +
    '🔷 COMPLEXITY (Kompleksitas Teknis)\n' +
    'Seberapa kompleks arsitektur/integrasi modul?\n\n' +
    '1-3: Simple, standalone module\n' +
    '4-6: Multiple integrations, moderate logic\n' +
    '7-10: High complexity, many dependencies\n\n' +
    '═══════════════════════════════════════════════════════\n\n' +
    '🔷 AUTOMATION (Tingkat Automasi)\n' +
    'Seberapa banyak test yang perlu automation?\n\n' +
    '1-3: Manual testing cukup\n' +
    '4-6: Mix manual + automation\n' +
    '7-10: Heavy automation needed (API, E2E)\n\n' +
    '═══════════════════════════════════════════════════════\n\n' +
    '🔷 TEST COMPLEXITY (Kompleksitas Testing)\n' +
    'Seberapa kompleks test case yang diperlukan?\n\n' +
    '1-3: Basic functional testing\n' +
    '4-6: Moderate test scenarios\n' +
    '7-10: Complex scenarios, many test cases\n\n' +
    '═══════════════════════════════════════════════════════\n\n' +
    '📊 TOTAL SCORE & CATEGORY:\n\n' +
    'Total = Diff + Risk + Comp + Auto + Test (5-50)\n\n' +
    '🟢 Very Easy: 5-14\n' +
    '🔵 Easy: 15-24\n' +
    '🟡 Medium: 25-34\n' +
    '🟠 Hard: 35-44\n' +
    '🔴 Very Hard: 45-50\n\n' +
    '═══════════════════════════════════════════════════════\n\n' +
    'Contoh Pengisian:\n\n' +
    'SIPGN E2E DAPUR:\n' +
    '• Diff: 7 (modul cukup kompleks)\n' +
    '• Risk: 6 (medium impact)\n' +
    '• Comp: 7 (banyak integrasi)\n' +
    '• Auto: 8 (perlu E2E automation)\n' +
    '• Test: 7 (many test scenarios)\n' +
    'Total: 35 → 🟠 Hard\n\n' +
    'INADigital POS:\n' +
    '• Diff: 5 (moderate)\n' +
    '• Risk: 4 (medium)\n' +
    '• Comp: 5 (moderate)\n' +
    '• Auto: 3 (mostly manual)\n' +
    '• Test: 4 (basic scenarios)\n' +
    'Total: 21 → 🔵 Easy';

  ui.alert('Parameter Guide', message, ui.ButtonSet.OK);
}

/**
 * Show about dialog
 */
function menuShowAbout() {
  const ui = SpreadsheetApp.getUi();

  const message =
    '👥 QA TEAM MANAGEMENT SYSTEM\n\n' +
    '═══════════════════════════════\n\n' +
    'Version: 5.0.0 (Simplified)\n' +
    'Author: QA Team\n\n' +
    '📋 KEY FEATURES\n\n' +
    '✅ Simple & clean interface\n' +
    '✅ 3-level hierarchy (Project → Modul → Submodul)\n' +
    '✅ 5 parameters (Diff, Risk, Comp, Auto, Test) - Scale 1-10\n' +
    '✅ 5-category difficulty system (Very Easy to Very Hard)\n' +
    '✅ Parameter guide based on QA Mandays Calculator\n' +
    '✅ Centralized configuration (all in one place)\n' +
    '✅ Copy-paste friendly team management\n' +
    '✅ Auto-generate team members from Config (PIC QA)\n' +
    '✅ Submodul distribution dashboard with color coding\n' +
    '✅ Manual data input directly in Google Sheets\n\n' +
    '═══════════════════════════════\n\n' +
    'TABS:\n\n' +
    '1. Config\n' +
    '   • Parameter guide (Diff, Risk, Comp, Auto, Test)\n' +
    '   • Project/Modul/Submodul list (flat table)\n' +
    '   • PIC QA assignments\n' +
    '   • All parameters with 1-10 ratings\n\n' +
    '2. Team Members\n' +
    '   • 8 columns (No, Name, Role, Projects, Modul, Submodul, Email, Status)\n' +
    '   • Copy-paste data directly or auto-generate from Config\n' +
    '   • Multiple assignments per person supported\n\n' +
    '3. Dashboard\n' +
    '   • Team distribution per submodul\n' +
    '   • 5-parameter ratings with color coding\n' +
    '   • Grouped by project and modul\n' +
    '   • 5-category difficulty legend\n\n' +
    '═══════════════════════════════\n\n' +
    'STRUCTURE:\n\n' +
    'Project → Modul → Submodul → Team Members\n\n' +
    'Example:\n' +
    'INADigital → INAgov → INAgov → Irvan\n' +
    'SIPGN → 4 → 4.1 → Adinda, Denta\n\n' +
    'WORKFLOW:\n\n' +
    '1. Menu → Create All Tabs\n' +
    '2. Fill Config tab (see parameter guide)\n' +
    '3. Option A: Manual input team data in Team Members tab\n' +
    '   Option B: Menu → Generate Team Members from Config\n' +
    '4. Menu → Create Dashboard to view distribution\n\n' +
    'DIFFICULTY CATEGORIES:\n\n' +
    'Based on Total Score (5-50):\n' +
    '🟢 Very Easy: 5-14\n' +
    '🔵 Easy: 15-24\n' +
    '🟡 Medium: 25-34\n' +
    '🟠 Hard: 35-44\n' +
    '🔴 Very Hard: 45-50\n\n' +
    '═══════════════════════════════\n\n' +
    '© 2025 QA Department';

  ui.alert('About', message, ui.ButtonSet.OK);
}

/**
 * Sync from QA Dashboard Config Sheet (SSOT)
 */
function menuSyncFromDashboard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if Dashboard ID is configured
  const settings = getSyncSettings();

  if (!settings.enabled || !settings.spreadsheetId) {
    ui.alert(
      'QA Dashboard Not Configured',
      'Please paste your QA Dashboard Spreadsheet ID in the Project tab first.\n\n' +
      'Location: Project tab → Cell Q2 (QA Dashboard ID)\n\n' +
      'You can paste either:\n' +
      '• Just the ID: 1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY\n' +
      '• Or full URL: https://docs.google.com/spreadsheets/d/ID/edit',
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.alert(
    'Sync from QA Dashboard?',
    'This will sync Config data from:\n' +
    settings.spreadsheetName + '\n\n' +
    '⚠️ This will OVERWRITE local Project config data.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    ss.toast('Syncing from QA Dashboard...', 'In Progress', -1);

    const result = syncFromDashboard();

    if (result.success) {
      ss.toast(
        'Synced ' + result.synced.projects + ' projects, ' +
        result.synced.moduls + ' moduls, ' +
        result.synced.submoduls + ' submoduls from QA Dashboard!',
        'Success',
        5
      );

      // Ask if user wants to refresh dashboard
      const refreshResponse = ui.alert(
        'Sync Complete!',
        'Data synced successfully.\n\n' +
        'Refresh Dashboard now?',
        ui.ButtonSet.YES_NO
      );

      if (refreshResponse === ui.Button.YES) {
        menuRefreshDashboard();
      }

    } else {
      ui.alert('Sync Failed', result.message + '\n\nPlease check:\n1. QA Dashboard ID is correct in Project tab (cell Q2)\n2. You have access to the QA Dashboard spreadsheet', ui.ButtonSet.OK);
    }

  } catch (error) {
    ui.alert('Error', 'Sync failed: ' + error.message, ui.ButtonSet.OK);
    Logger.log('Sync error: ' + error.stack);
  }
}

/**
 * Test Dashboard Connection
 */
function menuTestDashboardConnection() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const settings = getSyncSettings();

  if (!settings.enabled || !settings.spreadsheetId) {
    ui.alert(
      'QA Dashboard Not Configured',
      'Please paste your QA Dashboard Spreadsheet ID in Project tab → Cell Q2\n\n' +
      'Example:\n' +
      '1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY',
      ui.ButtonSet.OK
    );
    return;
  }

  ss.toast('Testing connection to QA Dashboard...', 'Testing', -1);

  const result = testSyncConnection();

  if (result.success) {
    ui.alert(
      'Connection Successful! ✅',
      result.message + '\n\n' +
      'Active modules found: ' + (result.activeModules || 0) + '\n\n' +
      'You can now use "Sync from QA Dashboard" to pull data.',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'Connection Failed ❌',
      result.message + '\n\n' +
      'Please check:\n' +
      '1. Spreadsheet ID is correct (cell Q2)\n' +
      '2. You have access to the QA Dashboard\n' +
      '3. Dashboard has a "Config" tab',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Setup Auto-Sync & Refresh
 */
function menuSetupAutoSyncRefresh() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Auto-Sync & Refresh',
    'Configure automatic sync from QA Dashboard and refresh local dashboard.\n\n' +
    'Choose interval:\n' +
    '• Every 1 hour - Very frequent\n' +
    '• Every 3 hours - Moderate\n' +
    '• Every 6 hours - Less frequent\n' +
    '• Every 12 hours - Twice daily\n' +
    '• Every 24 hours - Once daily\n\n' +
    'Enable auto-sync & refresh?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    // Check if there's existing trigger to disable
    const triggers = ScriptApp.getProjectTriggers();
    let removed = false;
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'autoSyncAndRefresh') {
        ScriptApp.deleteTrigger(trigger);
        removed = true;
      }
    });

    if (removed) {
      ui.alert('Auto-Sync Disabled', 'Auto-sync & refresh has been turned off.', ui.ButtonSet.OK);
    }
    return;
  }

  const intervalResponse = ui.prompt(
    'Enter Interval',
    'Enter hours (1, 3, 6, 12, or 24):',
    ui.ButtonSet.OK_CANCEL
  );

  if (intervalResponse.getSelectedButton() !== ui.Button.OK) return;

  const hours = parseInt(intervalResponse.getResponseText().trim());

  if (![1, 3, 6, 12, 24].includes(hours)) {
    ui.alert('Invalid Input', 'Please enter: 1, 3, 6, 12, or 24', ui.ButtonSet.OK);
    return;
  }

  try {
    // Remove existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'autoSyncAndRefresh') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Create new trigger
    ScriptApp.newTrigger('autoSyncAndRefresh')
      .timeBased()
      .everyHours(hours)
      .create();

    ui.alert(
      'Auto-Sync Enabled! ✅',
      'Will auto-sync from QA Dashboard and refresh every ' + hours + ' hour(s).\n\n' +
      'Next sync in ' + hours + ' hour(s).',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Error', 'Failed to setup auto-sync: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Apply dropdown validation to Team Members
 */
function menuApplyDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Update helper columns in Project tab first
    updateProjectHelperColumns();

    // Apply formatting and validation to Team Members
    const result = applyTeamMemberFormatting();

    if (result.success) {
      ss.toast('Dropdown applied! Check columns K, L, M in Project tab for dropdown values.', 'Success', 5);
    } else {
      ss.toast(result.message, 'Warning', 3);
    }
  } catch (error) {
    ss.toast('Failed: ' + error.message, 'Error', 5);
    Logger.log('Error applying dropdown: ' + error.message);
  }
}

/**
 * Setup auto-refresh dashboard
 */
function menuSetupAutoRefresh() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Auto-Refresh Dashboard',
    'Configure automatic dashboard refresh.\n\n' +
    'Choose refresh interval:\n' +
    '• Every 1 hour - Very frequent\n' +
    '• Every 3 hours - Moderate\n' +
    '• Every 6 hours - Less frequent\n' +
    '• Every 12 hours - Twice daily\n' +
    '• Every 24 hours - Once daily\n\n' +
    'Do you want to enable auto-refresh?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    // Check if there's existing trigger to disable
    const triggers = ScriptApp.getProjectTriggers();
    let removed = false;
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'autoRefreshDashboard') {
        ScriptApp.deleteTrigger(trigger);
        removed = true;
      }
    });

    if (removed) {
      ui.alert('Auto-Refresh Disabled', 'Auto-refresh has been turned off.', ui.ButtonSet.OK);
    }
    return;
  }

  const intervalResponse = ui.prompt(
    'Enter Refresh Interval',
    'Enter hours (1, 3, 6, 12, or 24):',
    ui.ButtonSet.OK_CANCEL
  );

  if (intervalResponse.getSelectedButton() !== ui.Button.OK) return;

  const hours = parseInt(intervalResponse.getResponseText().trim());

  if (![1, 3, 6, 12, 24].includes(hours)) {
    ui.alert('Invalid Input', 'Please enter: 1, 3, 6, 12, or 24', ui.ButtonSet.OK);
    return;
  }

  try {
    // Remove existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'autoRefreshDashboard') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Create new trigger
    ScriptApp.newTrigger('autoRefreshDashboard')
      .timeBased()
      .everyHours(hours)
      .create();

    ui.alert(
      'Auto-Refresh Enabled! ✅',
      'Dashboard will auto-refresh every ' + hours + ' hour(s).\n\n' +
      'Next refresh in ' + hours + ' hour(s).',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Error', 'Failed to setup auto-refresh: ' + error.message, ui.ButtonSet.OK);
  }
}
