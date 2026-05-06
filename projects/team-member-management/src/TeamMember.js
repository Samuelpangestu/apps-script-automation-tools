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
  NP: { index: 2, letter: 'B', width: 120, header: 'NP' },
  NAME: { index: 3, letter: 'C', width: 180, header: 'Name' },
  EMAIL: { index: 4, letter: 'D', width: 220, header: 'Email' },
  EMAIL_2: { index: 5, letter: 'E', width: 220, header: 'Email 2' },
  HP: { index: 6, letter: 'F', width: 140, header: 'HP' },
  JOIN_DATE: { index: 7, letter: 'G', width: 110, header: 'Join Date' },
  TITLE: { index: 8, letter: 'H', width: 150, header: 'Title' },
  ROLE: { index: 9, letter: 'I', width: 150, header: 'Role' },
  LEAD_PIC: { index: 10, letter: 'J', width: 150, header: 'Lead/PIC' },
  PROJECTS: { index: 11, letter: 'K', width: 200, header: 'Project' },
  MODUL: { index: 12, letter: 'L', width: 180, header: 'Modul' },
  SUBMODUL: { index: 13, letter: 'M', width: 250, header: 'Submodul' },
  STATUS: { index: 14, letter: 'N', width: 120, header: 'Status' },
  STATUS_HIRING: { index: 15, letter: 'O', width: 130, header: 'Status Hiring' },
  AUTOMATION: { index: 16, letter: 'P', width: 140, header: 'Automation' },
  GITHUB: { index: 17, letter: 'Q', width: 180, header: 'Github Personal' },
  VPN_ABC: { index: 18, letter: 'R', width: 120, header: 'VPN ABC' },
  VPN_HUWAWEI: { index: 19, letter: 'S', width: 140, header: 'VPN Huwawei' }
};

const TEAM_TOTAL_COLUMNS = 19;

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
    TEAM_COLUMNS.NP.header,
    TEAM_COLUMNS.NAME.header,
    TEAM_COLUMNS.EMAIL.header,
    TEAM_COLUMNS.EMAIL_2.header,
    TEAM_COLUMNS.HP.header,
    TEAM_COLUMNS.JOIN_DATE.header,
    TEAM_COLUMNS.TITLE.header,
    TEAM_COLUMNS.ROLE.header,
    TEAM_COLUMNS.LEAD_PIC.header,
    TEAM_COLUMNS.PROJECTS.header,
    TEAM_COLUMNS.MODUL.header,
    TEAM_COLUMNS.SUBMODUL.header,
    TEAM_COLUMNS.STATUS.header,
    TEAM_COLUMNS.STATUS_HIRING.header,
    TEAM_COLUMNS.AUTOMATION.header,
    TEAM_COLUMNS.GITHUB.header,
    TEAM_COLUMNS.VPN_ABC.header,
    TEAM_COLUMNS.VPN_HUWAWEI.header
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

  // Add sample data with new structure
  const sampleData = [
    [1, 'NP001', 'Samuel Pangestu', 'samuel.gonggom@inadigital.co.id', '', '081234567890', '2023-01-15', 'Senior QA Engineer', 'QA Team Lead', 'Samuel', 'Government Systems', 'SIPGN', 'Core System, Integration Module', 'Active', 'Permanent', 'Selenium, Playwright', 'github.com/samuel', 'Active', 'Active'],
    [2, 'NP002', 'Muhammad Lutfi', 'muhamad.ramdani@inadigital.co.id', '', '081234567891', '2023-02-01', 'Senior QA Engineer', 'QA Team Lead', 'Lutfi', 'Digital Peruri', 'INADigital', 'INAgov, Emeterai', 'Active', 'Permanent', 'Cypress, JMeter', 'github.com/lutfi', 'Active', 'Active'],
    [3, 'NP003', 'Irvan Muhandis', 'irvan.muhandis@inadigital.co.id', '', '081234567892', '2023-03-10', 'QA Engineer', 'PIC Project', 'Irvan', 'Digital Peruri', 'INADigital', 'INAgov, Wahana', 'Active', 'Permanent', 'Postman, K6', 'github.com/irvan', 'Active', 'Active'],
    [4, 'NP004', 'Muhammad Rizky', 'muhammad.ferdiansyah@inadigital.co.id', '', '081234567893', '2023-04-20', 'QA Engineer', 'Quality Engineer', 'Rizky', 'Digital Peruri', 'INADigital', 'Wahana, COTS', 'Active', 'Contract', 'Manual Testing', 'github.com/rizky', 'Active', 'Pending']
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
 * Add data validation using helper columns from Project tab
 */
function addTeamDataValidation(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!configSheet) {
    Logger.log('⚠️ Project tab not found, skipping dropdown validation');
    return;
  }

  // Role validation (apply to reasonable range)
  const roleRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.ROLE.index, 100);
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['QA Team Lead', 'QA Lead', 'PIC Project', 'Quality Engineer', 'Senior Quality Engineer', 'Intern Quality Engineer', 'Security Engineer', 'UX Research'], true)
    .build();
  roleRange.setDataValidation(roleRule);

  // Status validation
  const statusRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.STATUS.index, 100);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'On Leave'], true)
    .build();
  statusRange.setDataValidation(statusRule);

  // Get helper column last rows
  const configLastRow = configSheet.getLastRow();

  // Project validation - from helper column K with multiple selections
  if (configLastRow >= 2) {
    const projectHelperRange = configSheet.getRange('K2:K' + configLastRow);
    const projectRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.PROJECTS.index, 100);

    // Get values from helper column
    const projectValues = projectHelperRange.getValues().filter(row => row[0]).map(row => row[0]);

    const projectRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(projectValues, true)
      .setAllowInvalid(false)
      .build();
    projectRange.setDataValidation(projectRule);
  }

  // Modul validation - from helper column L with multiple selections
  if (configLastRow >= 2) {
    const modulHelperRange = configSheet.getRange('L2:L' + configLastRow);
    const modulRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.MODUL.index, 100);

    // Get values from helper column
    const modulValues = modulHelperRange.getValues().filter(row => row[0]).map(row => row[0]);

    const modulRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(modulValues, true)
      .setAllowInvalid(false)
      .build();
    modulRange.setDataValidation(modulRule);
  }

  // Submodul validation - from helper column M with multiple selections
  if (configLastRow >= 2) {
    const submodulHelperRange = configSheet.getRange('M2:M' + configLastRow);
    const submodulRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.SUBMODUL.index, 100);

    // Get values from helper column
    const submodulValues = submodulHelperRange.getValues().filter(row => row[0]).map(row => row[0]);

    const submodulRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(submodulValues, true)
      .setAllowInvalid(false)
      .build();
    submodulRange.setDataValidation(submodulRule);
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
    const name = row[2] ? row[2].toString().trim() : ''; // Column C (Name)
    const memberRole = row[8] ? row[8].toString().trim() : ''; // Column I (Role)
    const projects = row[10] ? row[10].toString().trim() : ''; // Column K (Projects)
    const modul = row[11] ? row[11].toString().trim() : ''; // Column L (Modul)
    const submodul = row[12] ? row[12].toString().trim() : ''; // Column M (Submodul)
    const status = row[13] ? row[13].toString().trim() : ''; // Column N (Status)

    if (name && status === 'Active' && memberRole === role) {
      members.push({
        name: name,
        role: memberRole,
        projects: projects.split(',').map(p => p.trim()).filter(p => p),
        modul: modul.split(',').map(m => m.trim()).filter(m => m),
        submodul: submodul.split(',').map(s => s.trim()).filter(s => s),
        email: row[3] ? row[3].toString().trim() : '' // Column D (Email)
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
    const name = row[2] ? row[2].toString().trim() : ''; // Column C (Name)
    const status = row[13] ? row[13].toString().trim() : ''; // Column N (Status)

    if (name && status === 'Active') {
      members.push({
        name: name,
        role: row[8] ? row[8].toString().trim() : '', // Column I (Role)
        projects: row[10] ? row[10].toString().split(',').map(p => p.trim()).filter(p => p) : [], // Column K (Projects)
        modul: row[11] ? row[11].toString().split(',').map(m => m.trim()).filter(m => m) : [], // Column L (Modul)
        submodul: row[12] ? row[12].toString().split(',').map(s => s.trim()).filter(s => s) : [], // Column M (Submodul)
        email: row[3] ? row[3].toString().trim() : '' // Column D (Email)
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
    const status = row[11] ? row[11].toString().trim() : ''; // Column L (index 11)

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
      rowNum,           // No
      '',               // NP - to be filled manually
      picName,          // Name
      '',               // Email - to be filled manually
      '',               // Email 2
      '',               // HP
      '',               // Join Date
      '',               // Title
      'Quality Engineer', // Role - default
      '',               // Lead/PIC
      projects,         // Projects
      moduls,           // Modul
      submoduls,        // Submodul
      'Active',         // Status
      '',               // Status Hiring
      '',               // Automation
      '',               // Github Personal
      '',               // VPN ABC
      ''                // VPN Huwawei
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
