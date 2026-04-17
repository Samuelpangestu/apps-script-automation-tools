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

  // Check if Config tab exists - create without alert to avoid conflicts
  if (!ss.getSheetByName(CONFIG_TAB_NAME)) {
    createConfigTab();
  }

  // Delete existing if any (when called from menuSetupAll)
  let existingSheet = ss.getSheetByName(TEAM_TAB_NAME);
  if (existingSheet) {
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
    // [No, Name, Project, Modul, Submodul, Status, QA Lead, NP, Email, Email2, HP, JoinDate, Title, Role, Lead, StatusHiring, Automation, Github, VPN_ABC, VPN_Huwawei]
    [1, 'Team Member 1', 'Project A', 'Project A', 'Project A', 'Active', 'QA Team Lead', 'EMP-001', 'member1@example.com', '', '+62 800-0000-0001', '2023-01-15', 'Senior Quality Engineer', 'Senior Quality Engineer', 'QA Team Lead', 'Onboard', '', 'member-1', TRUE, TRUE],
    [2, 'Team Member 2', 'Project B', 'Project B', 'Project B', 'Active', 'Team Member 1', 'EMP-002', 'member2@example.com', '', '+62 800-0000-0002', '2023-02-01', 'Quality Engineer', 'Quality Engineer', 'Team Member 1', 'Onboard', '', 'member-2', TRUE, TRUE],
    [3, 'Team Member 3', 'Project A, Project C', 'Project A, Project C', 'Project A, Project C', 'Active', 'Team Member 1', 'EMP-003', 'member3@example.com', '', '+62 800-0000-0003', '2023-03-10', 'Quality Engineer', 'Quality Engineer', 'Team Member 1', 'Onboard', '', 'member-3', FALSE, FALSE],
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
 * Apply formatting and validation to Team Members tab
 * Use this after copy-pasting data
 */
function applyTeamMemberFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TEAM_TAB_NAME);

  if (!sheet) {
    throw new Error('Team Members tab not found');
  }

  Logger.log('🎨 Applying formatting to Team Members...');

  const lastRow = sheet.getLastRow();
  if (lastRow < TEAM_DATA_START_ROW) {
    Logger.log('No data to format');
    return { success: false, message: 'No team member data found' };
  }

  const dataRowCount = lastRow - TEAM_DATA_START_ROW + 1;

  // Update numbering
  for (let i = 0; i < dataRowCount; i++) {
    const rowNum = TEAM_DATA_START_ROW + i;
    sheet.getRange(rowNum, TEAM_COLUMNS.NO.index).setValue(i + 1);
  }

  // Apply row formatting
  for (let i = 0; i < dataRowCount; i++) {
    const rowNum = TEAM_DATA_START_ROW + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, TEAM_TOTAL_COLUMNS)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }
  Logger.log('✅ Row formatting applied');

  // Reapply data validation
  addTeamDataValidation(sheet);

  // Reapply conditional formatting
  addTeamConditionalFormatting(sheet);

  Logger.log('✅ Formatting complete for ' + dataRowCount + ' rows');

  return {
    success: true,
    formatted: dataRowCount,
    message: 'Successfully formatted ' + dataRowCount + ' team members'
  };
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
