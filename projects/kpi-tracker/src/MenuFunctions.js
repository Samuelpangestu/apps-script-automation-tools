/**
 * MenuFunctions.js — Menu UI and Main Entry Point
 * ═══════════════════════════════════════════════════════════════════════
 * Main menu and initialization functions
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🎯 KPI Tracker')
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('🚀 Initial Setup (Create All Tabs)', 'initialSetup')
      .addSeparator()
      .addItem('📋 Create Config Tab', 'menuSetupConfigTab')
      .addItem('📊 Create KPI Definition Tab', 'menuSetupKPIDefinitionTab')
      .addItem('📈 Create Dashboard Tab', 'menuSetupDashboard'))
    .addSeparator()
    .addSubMenu(ui.createMenu('👥 Team Members')
      .addItem('🔄 Sync from QA Team Management', 'syncFromSSOT')
      .addSeparator()
      .addItem('ℹ️ Show Team Info', 'showSSOTInfo')
      .addItem('🔧 Test Connection', 'testSSOTConnection'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📈 KPI Tracking')
      .addItem('➕ Create Period Tracker', 'menuCreateKPITracker')
      .addSeparator()
      .addItem('🔄 Refresh Dashboard', 'menuRefreshDashboard'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📝 360 Review')
      .addItem('✨ Create Review Form', 'create360ReviewForm')
      .addItem('🔄 Update Reviewee List', 'update360RevieweeList')
      .addItem('ℹ️ Show Form Info', 'show360ReviewFormInfo'))
    .addSeparator()
    .addSubMenu(ui.createMenu('ℹ️ Help')
      .addItem('📖 User Guide', 'showUserGuide')
      .addItem('🔧 About', 'showAbout'))
    .addToUi();
}

/**
 * Initial setup - Run once when first using the tool
 * Creates all necessary tabs and structure
 */
function initialSetup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if tabs already exist
  const configExists = ss.getSheetByName(CONFIG_TAB_NAME);
  const kpiDefExists = ss.getSheetByName(KPI_DEF_TAB_NAME);
  const dashboardExists = ss.getSheetByName(DASHBOARD_TAB_NAME);

  let warningMsg = '';
  if (configExists || kpiDefExists || dashboardExists) {
    warningMsg = '⚠️ WARNING: Existing tabs will be DELETED and recreated.\n\n' +
      'Existing tabs found:\n' +
      (configExists ? '• Config\n' : '') +
      (kpiDefExists ? '• KPI Definition\n' : '') +
      (dashboardExists ? '• Dashboard\n' : '') +
      '\nAll data in these tabs will be LOST.\n\n';
  }

  const response = ui.alert(
    'KPI Tracker — Initial Setup',
    warningMsg +
    'This will create all necessary tabs for the KPI Tracker:\n' +
    '• Dashboard tab (overview & summary)\n' +
    '• Config tab (team members management)\n' +
    '• KPI Definition tab (maintainable KPI definitions)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    Logger.log('Starting initial setup...');

    SpreadsheetApp.getActiveSpreadsheet().toast('Creating tabs...', 'Setup', 3);

    // Create all tabs
    setupConfigTab();
    setupKPIDefinitionTab();
    createDashboard();

    Logger.log('✅ Initial setup complete');

    // Ask if user wants to sync from Team Management
    const syncResponse = ui.alert(
      'Setup Complete! ✅',
      'KPI Tracker has been set up successfully.\n\n' +
      'Do you want to sync team members from QA Team Management now?\n\n' +
      'Click YES to sync from QA Team Management\n' +
      'Click NO to use sample data (can sync later from menu)',
      ui.ButtonSet.YES_NO
    );

    if (syncResponse === ui.Button.YES) {
      syncFromSSOT();  // From SSOTSync.js
    } else {
      ui.alert(
        'Next Steps',
        '1. Sync team members: Menu → Team Members → Sync from QA Team Management\n' +
        '   OR update manually in Config tab\n' +
        '2. Review and adjust KPI definitions if needed\n' +
        '3. Create period tracker (Menu: KPI Tracking → Create Period Tracker)\n' +
        '4. Create 360 Review Form (Menu: 360 Review → Create Review Form)\n\n' +
        'Tip: Dashboard shows overview of all periods. Refresh anytime from menu.',
        ui.ButtonSet.OK
      );
    }

  } catch (e) {
    Logger.log('❌ Setup failed: ' + e.message);
    ui.alert('Setup Failed', 'Error: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Menu function to create Config tab
 */
function menuSetupConfigTab() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ss.getSheetByName(CONFIG_TAB_NAME);

  let msg = exists
    ? '⚠️ Config tab already exists and will be DELETED and recreated.\n\nAll data will be LOST.\n\nContinue?'
    : 'This will create the Config tab for team member management.\n\nContinue?';

  const response = ui.alert('Create Config Tab', msg, ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    setupConfigTab();
    ui.alert('Success! ✅', 'Config tab created successfully.', ui.ButtonSet.OK);
  } catch (e) {
    Logger.log('❌ Config tab creation failed: ' + e.message);
    ui.alert('Error', 'Failed to create Config tab:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Menu function to create KPI Definition tab
 */
function menuSetupKPIDefinitionTab() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ss.getSheetByName(KPI_DEF_TAB_NAME);

  let msg = exists
    ? '⚠️ KPI Definition tab already exists and will be DELETED and recreated.\n\nAll data will be LOST.\n\nContinue?'
    : 'This will create the KPI Definition tab with all pre-configured KPIs.\n\nContinue?';

  const response = ui.alert('Create KPI Definition Tab', msg, ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    setupKPIDefinitionTab();
    ui.alert('Success! ✅', 'KPI Definition tab created successfully.', ui.ButtonSet.OK);
  } catch (e) {
    Logger.log('❌ KPI Definition tab creation failed: ' + e.message);
    ui.alert('Error', 'Failed to create KPI Definition tab:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Menu function to create Dashboard tab
 */
function menuSetupDashboard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ss.getSheetByName(DASHBOARD_TAB_NAME);

  let msg = exists
    ? '⚠️ Dashboard tab already exists and will be DELETED and recreated.\n\nAll data will be LOST.\n\nContinue?'
    : 'This will create the Dashboard tab showing KPI overview.\n\nContinue?';

  const response = ui.alert('Create Dashboard Tab', msg, ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    createDashboard();
    ui.alert('Success! ✅', 'Dashboard tab created successfully.', ui.ButtonSet.OK);
  } catch (e) {
    Logger.log('❌ Dashboard creation failed: ' + e.message);
    ui.alert('Error', 'Failed to create Dashboard:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Show user guide
 */
function showUserGuide() {
  const ui = SpreadsheetApp.getUi();

  const guide =
    '📖 KPI TRACKER — USER GUIDE\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '🚀 GETTING STARTED:\n\n' +
    '1. Run "Initial Setup" from menu (once only)\n' +
    '2. Sync team members from SSOT (recommended)\n' +
    '   OR update manually in Config tab\n' +
    '3. Review KPI definitions (already pre-populated)\n' +
    '4. Create period tracker (Menu → KPI Tracking)\n' +
    '5. Create 360 Review Form\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '👥 TEAM MEMBERS (SSOT):\n\n' +
    '• SSOT = Single Source of Truth\n' +
    '• Central team member database\n' +
    '• Sync from SSOT: Menu → Team Members → Sync from SSOT\n' +
    '• Auto-maps roles and status\n' +
    '• Filters only QA/QE roles\n' +
    '• One-click synchronization\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '📊 DASHBOARD TAB:\n\n' +
    '• Overview of all tracking periods\n' +
    '• Quick stats: periods, members, KPIs\n' +
    '• Success rate per period\n' +
    '• Refresh anytime from menu\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '📋 CONFIG TAB:\n\n' +
    '• Team members data (synced from SSOT or manual)\n' +
    '• Can manually edit if needed\n' +
    '• Set status to "Aktif" or "Non-Aktif"\n' +
    '• Role dropdown auto-populated\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '📊 KPI DEFINITION TAB:\n\n' +
    '• All KPIs defined here per role\n' +
    '• MAINTAINABLE: Edit targets, formulas anytime\n' +
    '• Add new KPIs by adding rows\n' +
    '• Color-coded by role for easy reading\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '📈 KPI TRACKING:\n\n' +
    '• Create period tracker: Menu → KPI Tracking → Create Period Tracker\n' +
    '• Choose period type: Sprint / Monthly / Quarterly\n' +
    '• Fill "Actual" column with real values\n' +
    '• Achievement % and Status auto-calculate\n' +
    '• One sheet per period (e.g., "KPI - Sprint 24")\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '📝 360 REVIEW:\n\n' +
    '• Create form once: Menu → 360 Review → Create Form\n' +
    '• Form auto-linked to this spreadsheet\n' +
    '• Update reviewee list when team changes\n' +
    '• Responses automatically collected\n' +
    '• Weighted scoring: TL 35% + PM 25% + Peer 15% + Self 10% + Other 15%\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '🔄 UPDATING:\n\n' +
    '• Team members: Menu → Team Members → Sync from SSOT\n' +
    '  OR edit Config tab directly\n' +
    '• KPI definitions: Edit KPI Definition tab\n' +
    '• Period tracker: Fill Actual values\n' +
    '• Dashboard: Menu → KPI Tracking → Refresh Dashboard\n' +
    '• Reviewee list: Menu → 360 Review → Update Reviewee List\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '💡 TIPS:\n\n' +
    '• Keep Config tab updated for accurate tracking\n' +
    '• Review KPI definitions quarterly\n' +
    '• Create tracker per sprint/month/quarter\n' +
    '• Run 360 Review per semester/contract period\n' +
    '• Refresh Dashboard after updating trackers\n' +
    '• All changes are immediately reflected\n\n' +
    '═══════════════════════════════════════════';

  ui.alert(guide);
}

/**
 * Show about info
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();

  const about =
    '🎯 KPI TRACKER — QA DEPARTMENT PERURI\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'Version: 1.0.0\n' +
    'Created: ' + new Date().toISOString().split('T')[0] + '\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'FEATURES:\n\n' +
    '✅ Centralized team member management\n' +
    '✅ Maintainable KPI definitions per role\n' +
    '✅ Automated 360 Review form & scoring\n' +
    '✅ Flexible & scalable for team changes\n' +
    '✅ Pre-configured with QA Department KPIs\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'ROLES SUPPORTED:\n\n' +
    '• QA Team Lead (CoE)\n' +
    '• QA Lead (Project Dedicated)\n' +
    '• PIC Project (QE + Koordinator)\n' +
    '• Quality Engineer (QE)\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'KPI CATEGORIES:\n\n' +
    '• On-time Delivery Rate\n' +
    '• Defect Detection Efficiency\n' +
    '• Test Automation Coverage\n' +
    '• Defect Escape Rate\n' +
    '• 360 Review Score\n' +
    '• Test Pass Rate\n' +
    '• CoE Tool Adoption Rate\n' +
    '• Training Completion Rate\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'Built with ❤️ for QA Team PERURI';

  ui.alert(about);
}

/**
 * Test function to verify all dependencies
 */
function testSetup() {
  Logger.log('Testing KPI Tracker setup...');

  try {
    // Test Config functions
    Logger.log('Testing Config functions...');
    const members = getTeamMembers();
    Logger.log('✅ Found ' + members.length + ' team members');

    // Test KPI Definition functions
    Logger.log('Testing KPI Definition functions...');
    const roles = getAllRoles();
    Logger.log('✅ Found ' + roles.length + ' roles');

    roles.forEach(role => {
      const kpis = getKPIsForRole(role);
      Logger.log('  - ' + role + ': ' + kpis.length + ' KPIs');
    });

    // Test 360 Review functions
    Logger.log('Testing 360 Review functions...');
    const reviewees = getRevieweesFromConfig_();
    Logger.log('✅ Found ' + reviewees.length + ' reviewees');

    // Test KPI Tracker functions
    Logger.log('Testing KPI Tracker functions...');
    const trackers = getAllKPITrackers();
    Logger.log('✅ Found ' + trackers.length + ' tracker period(s)');

    // Test Dashboard functions
    Logger.log('Testing Dashboard functions...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(DASHBOARD_TAB_NAME);
    if (dashboard) {
      Logger.log('✅ Dashboard exists');
    } else {
      Logger.log('⚠️ Dashboard not found (run Initial Setup)');
    }

    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅ All tests passed!');
    Logger.log('═══════════════════════════════════════');

    SpreadsheetApp.getUi().alert(
      'Test Successful! ✅\n\n' +
      'Team members: ' + members.length + '\n' +
      'Roles: ' + roles.length + '\n' +
      'Reviewees: ' + reviewees.length + '\n' +
      'Tracker periods: ' + trackers.length + '\n' +
      'Dashboard: ' + (dashboard ? 'Yes' : 'No') + '\n\n' +
      'All functions working correctly.'
    );

  } catch (e) {
    Logger.log('❌ Test failed: ' + e.message);
    Logger.log(e.stack);
    SpreadsheetApp.getUi().alert('Test Failed: ' + e.message);
  }
}
