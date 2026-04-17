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
    .addItem('➕ Add Team Member', 'menuAddMember')
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
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup All Tabs',
    'This will create:\n\n' +
    '1. Config - Projects tab (project list + difficulty)\n' +
    '2. Team Members tab (team + project assignments)\n' +
    '3. Dashboard tab (project distribution)\n\n' +
    'Existing tabs will be replaced.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    createConfigTab();
    createTeamMemberTab();
    createDashboard();

    ui.alert(
      '✅ Setup Complete!',
      'All tabs have been created:\n\n' +
      '✓ Config - Projects\n' +
      '✓ Team Members\n' +
      '✓ Dashboard\n\n' +
      'You can now:\n' +
      '1. Add projects in Config tab\n' +
      '2. Add team members in Team Members tab\n' +
      '3. Assign projects (comma-separated)\n' +
      '4. View distribution in Dashboard',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Error', 'Setup failed: ' + error.message, ui.ButtonSet.OK);
    Logger.log('Error in setup: ' + error.message);
  }
}

/**
 * Setup Config tab only
 */
function menuSetupConfig() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Config Tab',
    'Create Config - Projects tab?\n\n' +
    'This will create a tab to manage:\n' +
    '• Project list\n' +
    '• Difficulty levels (Easy/Medium/Hard)\n' +
    '• Project descriptions\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    createConfigTab();
    ui.alert('Success', 'Config - Projects tab created!', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Setup Team Members tab only
 */
function menuSetupTeam() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Setup Team Members Tab',
    'Create Team Members tab?\n\n' +
    'This will create a tab to manage:\n' +
    '• Team member names\n' +
    '• Roles (QA Team Lead, QA Lead, PIC, QE)\n' +
    '• Project assignments (multiple)\n' +
    '• Email addresses\n' +
    '• Status\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    createTeamMemberTab();
    ui.alert('Success', 'Team Members tab created!', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Create dashboard
 */
function menuCreateDashboard() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Create Dashboard',
    'Create project distribution dashboard?\n\n' +
    'Shows:\n' +
    '• Team distribution by role\n' +
    '• Project assignments per role\n' +
    '• Difficulty levels\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    createDashboard();
  } catch (error) {
    ui.alert('Error', 'Failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Add team member
 */
function menuAddMember() {
  try {
    addTeamMember();
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Failed to add member: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Refresh dashboard
 */
function menuRefreshDashboard() {
  const ui = SpreadsheetApp.getUi();

  try {
    createDashboard();
    ui.alert('Success', 'Dashboard refreshed with latest data!', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed to refresh: ' + error.message, ui.ButtonSet.OK);
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
    'Version: 2.1.0\n' +
    'Author: QA Team\n\n' +
    '📋 KEY FEATURES\n\n' +
    '✅ Centralized project configuration\n' +
    '✅ Difficulty level definitions\n' +
    '✅ Simple team member management\n' +
    '✅ Auto-complete project names from Config\n' +
    '✅ Project distribution dashboard\n' +
    '✅ Scheduled auto-refresh dashboard\n' +
    '✅ Long-term maintainable\n' +
    '✅ Ready for KPI integration\n\n' +
    '═══════════════════════════════\n\n' +
    'TABS:\n\n' +
    '1. Config - Projects (CENTRALIZED)\n' +
    '   • Difficulty definitions table\n' +
    '   • Project list (single source)\n' +
    '   • Add/remove anytime\n\n' +
    '2. Team Members\n' +
    '   • 6 simple columns\n' +
    '   • Projects auto-complete from Config\n' +
    '   • Comma-separated for multiple\n\n' +
    '3. Dashboard\n' +
    '   • Team distribution per project\n' +
    '   • Manual or auto-refresh\n\n' +
    '═══════════════════════════════\n\n' +
    'QUICK START:\n\n' +
    '1. Menu → Setup All Tabs\n' +
    '2. Add projects in Config tab\n' +
    '3. Add team members\n' +
    '4. Type projects in Team Members\n' +
    '   (comma-separated, auto-complete)\n' +
    '5. Dashboard auto-updates\n\n' +
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
