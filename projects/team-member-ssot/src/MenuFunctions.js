/**
 * MenuFunctions.js — Menu UI for Team Member SSOT
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('👥 Team Management')
    .addItem('🎨 Format Team Member Tab', 'formatTeamMemberTab')
    .addSeparator()
    .addItem('➕ Add New Team Member', 'addNewTeamMember')
    .addSeparator()
    .addItem('📊 View Summary', 'createSummary')
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

/**
 * Show about dialog
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();

  const about =
    '👥 TEAM MEMBER SSOT\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'Single Source of Truth for Team Management\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'FEATURES:\n\n' +
    '🎨 Professional Formatting\n' +
    '• Clean header styling\n' +
    '• Optimal column widths\n' +
    '• Alternating row colors\n' +
    '• Conditional formatting\n' +
    '• Data validation\n' +
    '• Freeze panes & filters\n\n' +
    '➕ Easy Team Management\n' +
    '• Add new members with one click\n' +
    '• Pre-formatted rows\n' +
    '• Auto-validation\n\n' +
    '📊 Summary Statistics\n' +
    '• Total active/inactive members\n' +
    '• Count by status\n' +
    '• Count by role\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'USAGE:\n\n' +
    '1. Format Tab: Menu → Format Team Member Tab\n' +
    '2. Add Member: Menu → Add New Team Member\n' +
    '3. View Stats: Menu → View Summary\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'STATUS COLOR CODING:\n\n' +
    '🟢 Green: Onboard (Active)\n' +
    '🟡 Yellow: Tidak Ada Kabar\n' +
    '🔴 Red: Resign, Contract End, Digispark\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'INTEGRATED WITH:\n\n' +
    '• KPI Tracker (auto-sync)\n' +
    '• QA Dashboard\n' +
    '• Other QA Tools\n\n' +
    '═══════════════════════════════════════════';

  ui.alert(about);
}

/**
 * Show help/usage guide
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();

  const help =
    '📖 TEAM MEMBER SSOT — HELP\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'FORMATTING THE TAB:\n\n' +
    '1. Menu → Team Management → Format Team Member Tab\n' +
    '2. Click YES to confirm\n' +
    '3. Wait for formatting to complete\n\n' +
    'This will apply:\n' +
    '• Professional header (blue background)\n' +
    '• Optimal column widths\n' +
    '• Alternating row colors (white/gray)\n' +
    '• Status color coding (green/yellow/red)\n' +
    '• Data validation dropdowns\n' +
    '• Freeze header row\n' +
    '• Enable filter\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'ADDING NEW TEAM MEMBER:\n\n' +
    '1. Menu → Team Management → Add New Team Member\n' +
    '2. Enter full name\n' +
    '3. Row auto-created with defaults:\n' +
    '   - Join Date: Today\n' +
    '   - Status: Onboard\n' +
    '   - Role: Quality Engineer\n' +
    '4. Fill in remaining details (Email, Project, etc.)\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'VIEWING SUMMARY:\n\n' +
    '1. Menu → Team Management → View Summary\n' +
    '2. See statistics:\n' +
    '   - Total members (active/inactive)\n' +
    '   - Count by status\n' +
    '   - Count by role\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'COLUMN DESCRIPTIONS:\n\n' +
    '• Name: Full name\n' +
    '• Join Date: Start date (YYYY-MM-DD)\n' +
    '• Title: Job title/position\n' +
    '• Lead/PIC: Project lead designation\n' +
    '• Project: Assigned projects (comma-separated)\n' +
    '• NP: Employee number\n' +
    '• Email: Primary work email\n' +
    '• Email 2: Secondary email\n' +
    '• Status: Employment status (dropdown)\n' +
    '• Automation: Automation access (checkbox)\n' +
    '• Github: GitHub username\n' +
    '• Phone: Contact number\n' +
    '• Role: Job role (dropdown)\n' +
    '• VPN ABC: VPN access (checkbox)\n' +
    '• VPN Huwawei: VPN access (checkbox)\n\n' +
    '═══════════════════════════════════════════\n\n' +
    'TIPS:\n\n' +
    '• Use dropdowns for Status and Role\n' +
    '• Use checkboxes for Automation and VPNs\n' +
    '• Keep section headers (MBG, Riset, etc.)\n' +
    '• Run formatting after major changes\n' +
    '• Summary shows live statistics\n\n' +
    '═══════════════════════════════════════════';

  ui.alert(help);
}
