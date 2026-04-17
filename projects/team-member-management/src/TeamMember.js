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
  PROJECTS: { index: 4, letter: 'D', width: 200, header: 'Projects' },
  MODUL: { index: 5, letter: 'E', width: 200, header: 'Modul' },
  SUBMODUL: { index: 6, letter: 'F', width: 300, header: 'Submodul' },
  EMAIL: { index: 7, letter: 'G', width: 250, header: 'Email' },
  STATUS: { index: 8, letter: 'H', width: 100, header: 'Status' }
};

const TEAM_TOTAL_COLUMNS = 8;

/**
 * Create Team Members tab
 */
function createTeamMemberTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if Config tab exists
  if (!ss.getSheetByName(CONFIG_TAB_NAME)) {
    throw new Error('Config tab must be created first');
  }

  // Create new sheet (deletion handled by caller)
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
    TEAM_COLUMNS.MODUL.header,
    TEAM_COLUMNS.SUBMODUL.header,
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

  // Add sample data (Project → Modul → Submodul)
  const sampleData = [
    [1, 'Samuel Pangestu', 'QA Team Lead', 'Government Systems', 'SIPGN', 'Core System, Integration Module', 'samuel.gonggom@inadigital.co.id', 'Active'],
    [2, 'Muhammad Lutfi', 'QA Team Lead', 'Digital Peruri', 'INADigital', 'INAgov, Emeterai', 'muhamad.ramdani@inadigital.co.id', 'Active'],
    [3, 'Irvan Muhandis', 'PIC Project', 'Digital Peruri', 'INADigital', 'INAgov, Wahana', 'irvan.muhandis@inadigital.co.id', 'Active'],
    [4, 'Muhammad Rizky', 'Quality Engineer', 'Digital Peruri', 'INADigital', 'Wahana, COTS', 'muhammad.ferdiansyah@inadigital.co.id', 'Active']
  ];

  sheet.getRange(TEAM_DATA_START_ROW, 1, sampleData.length, TEAM_TOTAL_COLUMNS)
    .setValues(sampleData);

  // Apply basic formatting only
  for (let i = 0; i < sampleData.length; i++) {
    const rowNum = TEAM_DATA_START_ROW + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, TEAM_TOTAL_COLUMNS).setBackground(bg);
  }

  // Flush changes to ensure completion
  SpreadsheetApp.flush();

  Logger.log('✅ Team Members tab created');
}

/**
 * Add data validation
 */
function addTeamDataValidation(sheet) {
  // Role validation (apply to reasonable range)
  const roleRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.ROLE.index, 50);
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['QA Team Lead', 'QA Lead', 'PIC Project', 'Quality Engineer'], true)
    .build();
  roleRange.setDataValidation(roleRule);

  // Status validation
  const statusRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 50);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'On Leave'], true)
    .build();
  statusRange.setDataValidation(statusRule);

  // Project validation - get from Config (helper text)
  const projects = getActiveProjects();
  if (projects.length > 0) {
    const projectNames = projects.map(p => p.name);
    const helperText = 'Available projects: ' + projectNames.join(', ') + '\n\nSeparate multiple projects with comma (,)';

    const projectRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.PROJECTS.index, 50);
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
    .setRanges([sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 50, 1)])
    .build();

  const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Inactive')
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 50, 1)])
    .build();

  const leaveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('On Leave')
    .setBackground('#fff3cd')
    .setRanges([sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 50, 1)])
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
    const modul = row[4] ? row[4].toString().trim() : '';
    const submodul = row[5] ? row[5].toString().trim() : '';
    const status = row[7] ? row[7].toString().trim() : '';

    if (name && status === 'Active' && memberRole === role) {
      members.push({
        name: name,
        role: memberRole,
        projects: projects.split(',').map(p => p.trim()).filter(p => p),
        modul: modul.split(',').map(m => m.trim()).filter(m => m),
        submodul: submodul.split(',').map(s => s.trim()).filter(s => s),
        email: row[6] ? row[6].toString().trim() : ''
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
    const status = row[7] ? row[7].toString().trim() : '';

    if (name && status === 'Active') {
      members.push({
        name: name,
        role: row[2] ? row[2].toString().trim() : '',
        projects: row[3] ? row[3].toString().split(',').map(p => p.trim()).filter(p => p) : [],
        modul: row[4] ? row[4].toString().split(',').map(m => m.trim()).filter(m => m) : [],
        submodul: row[5] ? row[5].toString().split(',').map(s => s.trim()).filter(s => s) : [],
        email: row[6] ? row[6].toString().trim() : ''
      });
    }
  });

  return members;
}

/**
 * Generate Team Members from Config tab (auto-import from QA Portfolio Dashboard)
 * Handles multiple PICs per modul (e.g., "Adinda, Denta")
 */
function generateTeamMembersFromConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);
  const teamSheet = ss.getSheetByName(TEAM_TAB_NAME);

  if (!configSheet) {
    throw new Error('Config tab not found');
  }

  if (!teamSheet) {
    throw new Error('Team Members tab not found');
  }

  Logger.log('🔄 Generating Team Members from Config...');

  // Read Config data
  const lastRow = configSheet.getLastRow();
  if (lastRow < CONFIG_DATA_START_ROW) {
    Logger.log('⚠️ No data in Config tab');
    return { success: false, message: 'No data in Config tab' };
  }

  const configData = configSheet.getRange(CONFIG_DATA_START_ROW, 1, lastRow - CONFIG_DATA_START_ROW + 1, CONFIG_TOTAL_COLUMNS).getValues();

  // Parse and aggregate by PIC name
  const picMap = new Map(); // Map<picName, {projects: Set, moduls: Set, submoduls: Set}>

  configData.forEach(row => {
    const active = row[0];
    const projectName = row[1] ? row[1].toString().trim() : '';
    const modulName = row[2] ? row[2].toString().trim() : '';
    const submodulName = row[3] ? row[3].toString().trim() : '';
    const picQA = row[4] ? row[4].toString().trim() : '';
    const status = row[9] ? row[9].toString().trim() : '';

    if (active && projectName && modulName && submodulName && picQA && status === 'Active') {
      // Parse multiple PICs (comma-separated)
      const picNames = picQA.split(',').map(name => name.trim()).filter(name => name);

      picNames.forEach(picName => {
        if (!picMap.has(picName)) {
          picMap.set(picName, {
            projects: new Set(),
            moduls: new Set(),
            submoduls: new Set()
          });
        }

        const picData = picMap.get(picName);
        picData.projects.add(projectName);
        picData.moduls.add(modulName);
        picData.submoduls.add(submodulName);
      });
    }
  });

  if (picMap.size === 0) {
    Logger.log('⚠️ No PICs found in Config tab');
    return { success: false, message: 'No PICs found in Config' };
  }

  // Clear existing Team Members data (keep header)
  const teamLastRow = teamSheet.getLastRow();
  if (teamLastRow > TEAM_HEADER_ROW) {
    teamSheet.getRange(TEAM_DATA_START_ROW, 1, teamLastRow - TEAM_HEADER_ROW, TEAM_TOTAL_COLUMNS).clearContent();
  }

  // Generate Team Members data
  const teamData = [];
  let rowNum = 1;

  picMap.forEach((data, picName) => {
    const projects = Array.from(data.projects).join(', ');
    const moduls = Array.from(data.moduls).join(', ');
    const submoduls = Array.from(data.submoduls).join(', ');

    teamData.push([
      rowNum,
      picName,
      'Quality Engineer', // Default role - user can update manually
      projects,
      moduls,
      submoduls,
      '', // Email - to be filled manually
      'Active'
    ]);

    rowNum++;
  });

  // Write to Team Members tab
  if (teamData.length > 0) {
    teamSheet.getRange(TEAM_DATA_START_ROW, 1, teamData.length, TEAM_TOTAL_COLUMNS).setValues(teamData);

    // Apply formatting
    applyTeamMemberFormatting();

    Logger.log('✅ Generated ' + teamData.length + ' team members from Config');

    return {
      success: true,
      generated: teamData.length,
      message: 'Generated ' + teamData.length + ' team members from Config'
    };
  }

  return { success: false, message: 'No team members generated' };
}
