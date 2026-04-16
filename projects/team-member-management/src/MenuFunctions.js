/**
 * Team Member Management - Menu Functions
 *
 * This file contains all menu-related functions for the Team Member Management tool.
 * It provides a user-friendly interface to access all features through the Google Sheets menu.
 *
 * @author QA Team
 * @version 1.0.0
 */

/**
 * Creates custom menu when spreadsheet is opened
 * This function runs automatically when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('👥 Team Management')
    .addItem('🔨 Create Team Member Tab', 'menuCreateTeamMemberTab')
    .addSeparator()
    .addItem('➕ Add Team Member', 'menuAddTeamMember')
    .addItem('📊 View Summary', 'menuViewSummary')
    .addItem('📤 Export Data', 'menuExportData')
    .addSeparator()
    .addItem('ℹ️ About', 'menuShowAbout')
    .addItem('❓ Help', 'menuShowHelp')
    .addToUi();
}

/**
 * Menu handler: Create Team Member Tab
 * Creates a new Team Member tab from scratch with proper structure and formatting
 */
function menuCreateTeamMemberTab() {
  const ui = SpreadsheetApp.getUi();

  // Confirm action
  const response = ui.alert(
    'Create Team Member Tab',
    'This will create a new "Team Member" tab with professional structure and sample data.\n\n' +
    'If the tab already exists, it will be cleared and rebuilt.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    createTeamMemberTab();
    ui.alert(
      '✅ Success',
      'Team Member tab has been created successfully!\n\n' +
      'The tab includes:\n' +
      '• Professional header formatting\n' +
      '• Sample data for reference\n' +
      '• Data validation (dropdowns, checkboxes)\n' +
      '• Conditional formatting (color-coded status)\n' +
      '• Filter views\n\n' +
      'You can now start adding your team members!',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      '❌ Error',
      'Failed to create Team Member tab:\n\n' + error.message,
      ui.ButtonSet.OK
    );
    Logger.log('Error creating Team Member tab: ' + error.message);
  }
}

/**
 * Menu handler: Add Team Member
 * Adds a new row for a team member with default values
 */
function menuAddTeamMember() {
  const ui = SpreadsheetApp.getUi();

  try {
    const result = addTeamMember();

    if (result.success) {
      ui.alert(
        '✅ Success',
        'New team member row added successfully!\n\n' +
        'Row: ' + result.row + '\n\n' +
        'Default values:\n' +
        '• Join Date: Today\n' +
        '• Status: Onboard\n' +
        '• Role: Quality Engineer\n\n' +
        'Please fill in the member details.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '❌ Error',
        result.message,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    ui.alert(
      '❌ Error',
      'Failed to add team member:\n\n' + error.message,
      ui.ButtonSet.OK
    );
    Logger.log('Error adding team member: ' + error.message);
  }
}

/**
 * Menu handler: View Summary
 * Displays statistics about team members
 */
function menuViewSummary() {
  const ui = SpreadsheetApp.getUi();

  try {
    const summary = viewSummary();

    if (!summary.success) {
      ui.alert('❌ Error', summary.message, ui.ButtonSet.OK);
      return;
    }

    const data = summary.data;

    let message = '📊 TEAM MEMBER SUMMARY\n\n';
    message += '═══════════════════════════════\n\n';
    message += '👥 TOTAL MEMBERS\n';
    message += '   Total: ' + data.total + '\n';
    message += '   Active: ' + data.active + '\n';
    message += '   Inactive: ' + data.inactive + '\n\n';

    message += '📈 BY STATUS\n';
    for (const status in data.byStatus) {
      message += '   ' + status + ': ' + data.byStatus[status] + '\n';
    }
    message += '\n';

    message += '💼 BY ROLE\n';
    for (const role in data.byRole) {
      message += '   ' + role + ': ' + data.byRole[role] + '\n';
    }
    message += '\n';

    message += '🔧 BY LEVEL\n';
    for (const level in data.byLevel) {
      message += '   ' + level + ': ' + data.byLevel[level] + '\n';
    }
    message += '\n';

    message += '═══════════════════════════════\n';
    message += 'Generated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    ui.alert('Team Member Summary', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert(
      '❌ Error',
      'Failed to generate summary:\n\n' + error.message,
      ui.ButtonSet.OK
    );
    Logger.log('Error viewing summary: ' + error.message);
  }
}

/**
 * Menu handler: Export Data
 * Exports team member data to CSV format
 */
function menuExportData() {
  const ui = SpreadsheetApp.getUi();

  // Confirm action
  const response = ui.alert(
    'Export Team Member Data',
    'This will export all team member data to a CSV file.\n\n' +
    'The file will be created in the same folder as this spreadsheet.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const result = exportData();

    if (result.success) {
      ui.alert(
        '✅ Success',
        'Team member data has been exported successfully!\n\n' +
        'File: ' + result.fileName + '\n' +
        'Rows exported: ' + result.rowCount + '\n\n' +
        'The file has been created in the same folder as this spreadsheet.\n\n' +
        'File ID: ' + result.fileId,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '❌ Error',
        result.message,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    ui.alert(
      '❌ Error',
      'Failed to export data:\n\n' + error.message,
      ui.ButtonSet.OK
    );
    Logger.log('Error exporting data: ' + error.message);
  }
}

/**
 * Menu handler: Show About dialog
 * Displays information about the tool
 */
function menuShowAbout() {
  const ui = SpreadsheetApp.getUi();

  const message =
    '👥 TEAM MEMBER MANAGEMENT TOOL\n\n' +
    '═══════════════════════════════\n\n' +
    'Version: 1.0.0\n' +
    'Author: QA Team\n\n' +
    '📋 DESCRIPTION\n\n' +
    'A professional team member management system for tracking QA team information.\n\n' +
    'Features:\n' +
    '• Professional tab structure with color-coded formatting\n' +
    '• Comprehensive member information tracking\n' +
    '• Data validation with dropdowns and checkboxes\n' +
    '• Conditional formatting for status visualization\n' +
    '• Team statistics and summary reports\n' +
    '• CSV export functionality\n' +
    '• Portable design for multi-company use\n\n' +
    '═══════════════════════════════\n\n' +
    '© 2024 QA Department';

  ui.alert('About Team Member Management', message, ui.ButtonSet.OK);
}

/**
 * Menu handler: Show Help dialog
 * Displays comprehensive usage guide
 */
function menuShowHelp() {
  const ui = SpreadsheetApp.getUi();

  const message =
    '❓ TEAM MEMBER MANAGEMENT - HELP GUIDE\n\n' +
    '═══════════════════════════════\n\n' +
    '🔨 CREATE TEAM MEMBER TAB\n\n' +
    'Creates a new "Team Member" tab with:\n' +
    '• Professional header (blue background, white text)\n' +
    '• 15 columns for comprehensive tracking\n' +
    '• Sample data for reference\n' +
    '• Data validation (dropdowns, checkboxes)\n' +
    '• Conditional formatting (color-coded status)\n' +
    '• Filter views\n\n' +
    'If the tab exists, it will be cleared and rebuilt.\n\n' +
    '─────────────────────────────\n\n' +
    '➕ ADD TEAM MEMBER\n\n' +
    'Adds a new row for a team member with defaults:\n' +
    '• Join Date: Today\n' +
    '• Status: Onboard\n' +
    '• Role: Quality Engineer\n\n' +
    'Fill in the remaining details manually.\n\n' +
    '─────────────────────────────\n\n' +
    '📊 VIEW SUMMARY\n\n' +
    'Displays team statistics:\n' +
    '• Total members (active/inactive)\n' +
    '• Members by status\n' +
    '• Members by role\n' +
    '• Members by level\n\n' +
    '─────────────────────────────\n\n' +
    '📤 EXPORT DATA\n\n' +
    'Exports all team member data to CSV:\n' +
    '• Creates file in same folder as spreadsheet\n' +
    '• Includes all columns and data\n' +
    '• Can be opened in Excel or other tools\n\n' +
    '═══════════════════════════════\n\n' +
    '📋 COLUMN GUIDE\n\n' +
    '1. Name: Full name of team member\n' +
    '2. Join Date: Date joined the team\n' +
    '3. Role: QA role (dropdown)\n' +
    '4. Level: Seniority level (dropdown)\n' +
    '5. Project Assignment: Current projects\n' +
    '6. NIP: Employee ID number\n' +
    '7. Email: Work email address\n' +
    '8. Division/Unit: Department or division\n' +
    '9. Status: Employment status (dropdown)\n' +
    '10. Automation: Has automation skills (checkbox)\n' +
    '11. Slack Username: Slack handle\n' +
    '12. WhatsApp Number: Contact number\n' +
    '13. Role AMS: AMS-specific role\n' +
    '14. VPN ABC: Has VPN ABC access (checkbox)\n' +
    '15. VPN Huwawei: Has VPN Huwawei access (checkbox)\n\n' +
    '═══════════════════════════════\n\n' +
    '🎨 STATUS COLOR CODES\n\n' +
    '🟢 Green: Onboard (active)\n' +
    '🟡 Yellow: Tidak Ada Kabar (inactive)\n' +
    '🔴 Red: Resign/Contract End/Digispark\n\n' +
    '═══════════════════════════════\n\n' +
    'For more help, contact the QA Team.';

  ui.alert('Team Member Management - Help', message, ui.ButtonSet.OK);
}
