/**
 * TeamMember.js — Simplified Team Member Management
 * ═══════════════════════════════════════════════════════════════════════
 * Manage QA team members with project assignments
 * ═══════════════════════════════════════════════════════════════════════
 */

const TEAM_TAB_NAME = 'Team Members';
const TEAM_HEADER_ROW = 1;
const TEAM_DATA_START_ROW = 2;

const TEAM_COLUMNS = {
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  NAME: { index: 2, letter: 'B', width: 200, header: 'Name' },
  ROLE: { index: 3, letter: 'C', width: 150, header: 'Role' },
  PROJECTS: { index: 4, letter: 'D', width: 350, header: 'Projects' },
  EMAIL: { index: 5, letter: 'E', width: 250, header: 'Email' },
  STATUS: { index: 6, letter: 'F', width: 100, header: 'Status' }
};

const TEAM_TOTAL_COLUMNS = 6;

/**
 * Create Team Members tab
 */
function createTeamMemberTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Check if Config tab exists
  if (!ss.getSheetByName(CONFIG_TAB_NAME)) {
    ui.alert('Config tab not found. Creating Config tab first...');
    createConfigTab();
  }

  // Delete existing if any
  let existingSheet = ss.getSheetByName(TEAM_TAB_NAME);
  if (existingSheet) {
    const response = ui.alert(
      'Tab already exists',
      'Team Members tab already exists. Delete and recreate?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
    ss.deleteSheet(existingSheet);
  }

  // Create new sheet
  const sheet = ss.insertSheet(TEAM_TAB_NAME, 1);

  // Setup structure
  sheet.setRowHeight(TEAM_HEADER_ROW, 40);
  sheet.setFrozenRows(TEAM_HEADER_ROW);

  // Create header
  const headers = [
    TEAM_COLUMNS.NO.header,
    TEAM_COLUMNS.NAME.header,
    TEAM_COLUMNS.ROLE.header,
    TEAM_COLUMNS.PROJECTS.header,
    TEAM_COLUMNS.EMAIL.header,
    TEAM_COLUMNS.STATUS.header
  ];

  sheet.getRange(TEAM_HEADER_ROW, 1, 1, TEAM_TOTAL_COLUMNS)
    .setValues([headers])
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Set column widths
  Object.values(TEAM_COLUMNS).forEach(col => {
    sheet.setColumnWidth(col.index, col.width);
  });

  // Add sample data
  const sampleData = [
    [1, 'Samuel Pangestu', 'QA Team Lead', 'SIPGN, INAgov, Emeterai', 'samuel.gonggom@inadigital.co.id', 'Active'],
    [2, 'Muhammad Lutfi', 'QA Team Lead', 'Shared Resource', 'muhamad.ramdani@inadigital.co.id', 'Active'],
    [3, 'Irvan Muhandis', 'PIC Project', 'INAgov, Wahana', 'irvan.muhandis@inadigital.co.id', 'Active'],
    [4, 'Muhammad Rizky', 'Quality Engineer', 'Wahana', 'muhammad.ferdiansyah@inadigital.co.id', 'Active']
  ];

  sheet.getRange(TEAM_DATA_START_ROW, 1, sampleData.length, TEAM_TOTAL_COLUMNS)
    .setValues(sampleData);

  // Apply formatting
  for (let i = 0; i < sampleData.length; i++) {
    const rowNum = TEAM_DATA_START_ROW + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, TEAM_TOTAL_COLUMNS)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Add data validation
  addTeamDataValidation(sheet);

  // Add conditional formatting for status
  addTeamConditionalFormatting(sheet);

  // Enable filter
  sheet.getRange(TEAM_HEADER_ROW, 1, 100, TEAM_TOTAL_COLUMNS).createFilter();

  Logger.log('✅ Team Members tab created');

  ui.alert(
    'Success! ✅',
    'Team Members tab has been created!\n\n' +
    'You can now:\n' +
    '• Add team members\n' +
    '• Assign projects (comma-separated)\n' +
    '• View dashboard for project distribution',
    ui.ButtonSet.OK
  );
}

/**
 * Add data validation
 */
function addTeamDataValidation(sheet) {
  // Role validation
  const roleRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.ROLE.index, 100);
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['QA Team Lead', 'QA Lead', 'PIC Project', 'Quality Engineer'], true)
    .build();
  roleRange.setDataValidation(roleRule);

  // Status validation
  const statusRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 100);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'On Leave'], true)
    .build();
  statusRange.setDataValidation(statusRule);

  // Project validation - get from Config (helper text)
  const projects = getActiveProjects();
  if (projects.length > 0) {
    const projectNames = projects.map(p => p.name);
    const helperText = 'Available projects: ' + projectNames.join(', ') + '\n\nSeparate multiple projects with comma (,)';

    const projectRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.PROJECTS.index, 100);
    const projectRule = SpreadsheetApp.newDataValidation()
      .setHelpText(helperText)
      .setAllowInvalid(true)
      .build();
    projectRange.setDataValidation(projectRule);
  }

  Logger.log('✅ Data validation added');
}

/**
 * Add conditional formatting
 */
function addTeamConditionalFormatting(sheet) {
  const activeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Active')
    .setBackground('#d4edda')
    .setRanges([sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 100, 1)])
    .build();

  const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Inactive')
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 100, 1)])
    .build();

  const leaveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('On Leave')
    .setBackground('#fff3cd')
    .setRanges([sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 100, 1)])
    .build();

  sheet.setConditionalFormatRules([activeRule, inactiveRule, leaveRule]);

  Logger.log('✅ Conditional formatting added');
}

/**
 * Add new team member
 */
function addTeamMember() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TEAM_TAB_NAME);

  if (!sheet) {
    ui.alert('Error', 'Team Members tab not found. Create it first.', ui.ButtonSet.OK);
    return;
  }

  // Get name
  const nameResponse = ui.prompt('Add Team Member', 'Enter name:', ui.ButtonSet.OK_CANCEL);
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;

  const name = nameResponse.getResponseText().trim();
  if (!name) {
    ui.alert('Error', 'Name cannot be empty', ui.ButtonSet.OK);
    return;
  }

  // Find next row
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;
  const no = lastRow - TEAM_HEADER_ROW + 1;

  // Add row
  const rowData = [no, name, 'Quality Engineer', '', '', 'Active'];
  sheet.getRange(newRow, 1, 1, TEAM_TOTAL_COLUMNS).setValues([rowData]);

  // Format row
  const bg = (newRow - TEAM_DATA_START_ROW) % 2 === 0 ? '#ffffff' : '#f8f9fa';
  sheet.getRange(newRow, 1, 1, TEAM_TOTAL_COLUMNS)
    .setBackground(bg)
    .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  ui.alert(
    'Success! ✅',
    'Team member added: ' + name + '\n\n' +
    'Next steps:\n' +
    '1. Fill in email and other details\n' +
    '2. Type project names in Projects column\n' +
    '   (comma-separated for multiple)\n' +
    '3. Available projects are shown in helper text\n\n' +
    'Projects auto-complete from Config tab.',
    ui.ButtonSet.OK
  );
}

/**
 * Get team members by role
 */
function getTeamMembersByRole(role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TEAM_TAB_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < TEAM_DATA_START_ROW) return [];

  const data = sheet.getRange(TEAM_DATA_START_ROW, 1, lastRow - TEAM_DATA_START_ROW + 1, TEAM_TOTAL_COLUMNS).getValues();

  const members = [];
  data.forEach(row => {
    const name = row[1] ? row[1].toString().trim() : '';
    const memberRole = row[2] ? row[2].toString().trim() : '';
    const projects = row[3] ? row[3].toString().trim() : '';
    const status = row[5] ? row[5].toString().trim() : '';

    if (name && status === 'Active' && memberRole === role) {
      members.push({
        name: name,
        role: memberRole,
        projects: projects.split(',').map(p => p.trim()).filter(p => p),
        email: row[4] ? row[4].toString().trim() : ''
      });
    }
  });

  return members;
}

/**
 * Get all active team members
 */
function getAllActiveTeamMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TEAM_TAB_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < TEAM_DATA_START_ROW) return [];

  const data = sheet.getRange(TEAM_DATA_START_ROW, 1, lastRow - TEAM_DATA_START_ROW + 1, TEAM_TOTAL_COLUMNS).getValues();

  const members = [];
  data.forEach(row => {
    const name = row[1] ? row[1].toString().trim() : '';
    const status = row[5] ? row[5].toString().trim() : '';

    if (name && status === 'Active') {
      members.push({
        name: name,
        role: row[2] ? row[2].toString().trim() : '',
        projects: row[3] ? row[3].toString().split(',').map(p => p.trim()).filter(p => p) : [],
        email: row[4] ? row[4].toString().trim() : ''
      });
    }
  });

  return members;
}
