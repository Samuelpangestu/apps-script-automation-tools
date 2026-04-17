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
    .addItem('🎨 Apply Team Formatting', 'menuApplyTeamFormatting')
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
 * Apply formatting to Team Members tab
 */
function menuApplyTeamFormatting() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Apply Team Formatting',
    'This will apply formatting and validation to Team Members tab:\n\n' +
    '✓ Auto-numbering\n' +
    '✓ Alternating row colors\n' +
    '✓ Data validation (dropdowns)\n' +
    '✓ Project helper text from Config\n' +
    '✓ Status color coding\n' +
    '✓ Cell borders\n\n' +
    'Use this after copy-pasting team member data.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const result = applyTeamMemberFormatting();

    if (result.success) {
      ui.alert(
        'Formatting Applied! ✅',
        'Team Members tab has been formatted.\n\n' +
        'Formatted: ' + result.formatted + ' team members\n\n' +
        'Applied:\n' +
        '• Auto-numbering\n' +
        '• Row styling\n' +
        '• Data validation\n' +
        '• Project helper text\n' +
        '• Status colors\n\n' +
        'You can now assign projects by typing in the Projects column.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('No Data', result.message, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Error', 'Failed to apply formatting: ' + error.message, ui.ButtonSet.OK);
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
    'Version: 2.2.0\n' +
    'Author: QA Team\n\n' +
    '📋 KEY FEATURES\n\n' +
    '✅ Centralized project configuration\n' +
    '✅ Difficulty level definitions\n' +
    '✅ Copy-paste friendly team management\n' +
    '✅ Auto-formatting after paste\n' +
    '✅ Project names from Config (single source)\n' +
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
    '   • Copy-paste data directly!\n' +
    '   • Apply formatting after paste\n\n' +
    '3. Dashboard\n' +
    '   • Team distribution per project\n' +
    '   • Manual or auto-refresh\n\n' +
    '═══════════════════════════════\n\n' +
    'WORKFLOW:\n\n' +
    '1. Menu → Setup All Tabs\n' +
    '2. Add projects in Config tab\n' +
    '3. Copy-paste team data to Team Members\n' +
    '4. Menu → Apply Team Formatting\n' +
    '5. Type projects (comma-separated)\n' +
    '6. Dashboard auto-updates\n\n' +
    'COPY-PASTE DATA:\n\n' +
    'Just paste your data in Team Members tab!\n' +
    'Then: Menu → Apply Team Formatting\n\n' +
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
