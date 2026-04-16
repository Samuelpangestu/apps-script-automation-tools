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
      .addItem('🚀 Initial Setup (Run Once)', 'initialSetup')
      .addSeparator()
      .addItem('📋 Setup Config Tab', 'setupConfigTab')
      .addItem('📊 Setup KPI Definition Tab', 'setupKPIDefinitionTab'))
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
  const response = ui.alert(
    'KPI Tracker — Initial Setup',
    'This will create all necessary tabs and structure for the KPI Tracker.\n\n' +
    'The following will be created:\n' +
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

    // Step 1: Setup Config tab
    ui.alert('Step 1/2: Creating Config tab...');
    setupConfigTab();

    // Step 2: Setup KPI Definition tab
    ui.alert('Step 2/2: Creating KPI Definition tab...');
    setupKPIDefinitionTab();

    Logger.log('✅ Initial setup complete');

    ui.alert(
      'Setup Complete! ✅',
      'KPI Tracker has been set up successfully.\n\n' +
      'Next steps:\n' +
      '1. Update team members in Config tab\n' +
      '2. Review and adjust KPI definitions if needed\n' +
      '3. Create 360 Review Form (Menu: 360 Review → Create Review Form)\n\n' +
      'Tip: All KPI definitions are maintainable in the KPI Definition tab.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log('❌ Setup failed: ' + e.message);
    ui.alert('Setup Failed', 'Error: ' + e.message, ui.ButtonSet.OK);
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
    '2. Update team members in Config tab\n' +
    '3. Review KPI definitions (already pre-populated)\n' +
    '4. Create 360 Review Form\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '📋 CONFIG TAB:\n\n' +
    '• Manage team members here\n' +
    '• Add/remove members as needed\n' +
    '• Set status to "Aktif" or "Non-Aktif"\n' +
    '• Role dropdown auto-populated\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '📊 KPI DEFINITION TAB:\n\n' +
    '• All KPIs defined here per role\n' +
    '• MAINTAINABLE: Edit targets, formulas anytime\n' +
    '• Add new KPIs by adding rows\n' +
    '• Color-coded by role for easy reading\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '📝 360 REVIEW:\n\n' +
    '• Create form once: Menu → 360 Review → Create Form\n' +
    '• Form auto-linked to this spreadsheet\n' +
    '• Update reviewee list when team changes\n' +
    '• Responses automatically collected\n' +
    '• Weighted scoring: TL 35% + PM 25% + Peer 15% + Self 10% + Other 15%\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '🔄 UPDATING:\n\n' +
    '• Team members: Edit Config tab directly\n' +
    '• KPI definitions: Edit KPI Definition tab\n' +
    '• Reviewee list: Menu → 360 Review → Update Reviewee List\n\n' +
    '═══════════════════════════════════════════\n\n' +
    '💡 TIPS:\n\n' +
    '• Keep Config tab updated for accurate tracking\n' +
    '• Review KPI definitions quarterly\n' +
    '• Run 360 Review per semester/contract period\n' +
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

    Logger.log('');
    Logger.log('═══════════════════════════════════════');
    Logger.log('✅ All tests passed!');
    Logger.log('═══════════════════════════════════════');

    SpreadsheetApp.getUi().alert(
      'Test Successful! ✅\n\n' +
      'Team members: ' + members.length + '\n' +
      'Roles: ' + roles.length + '\n' +
      'Reviewees: ' + reviewees.length + '\n\n' +
      'All functions working correctly.'
    );

  } catch (e) {
    Logger.log('❌ Test failed: ' + e.message);
    Logger.log(e.stack);
    SpreadsheetApp.getUi().alert('Test Failed: ' + e.message);
  }
}
