/**
 * TeamMember.js — Simplified Team Member Management
 * ═══════════════════════════════════════════════════════════════════════
 * Manage QA team members with project assignments
 * ═══════════════════════════════════════════════════════════════════════
 */

const TEAM_TAB_NAME = 'Team Members';
const TEAM_HEADER_ROW = 1;
const TEAM_DATA_START_ROW = 2;

// Auto-generated columns (left) + Manual fill columns (right)
const TEAM_COLUMNS = {
  // Auto-generated from Project Config (columns A-F)
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  NAME: { index: 2, letter: 'B', width: 180, header: 'Name' },
  PROJECTS: { index: 3, letter: 'C', width: 200, header: 'Project' },
  MODUL: { index: 4, letter: 'D', width: 180, header: 'Modul' },
  SUBMODUL: { index: 5, letter: 'E', width: 250, header: 'Submodul' },
  STATUS: { index: 6, letter: 'F', width: 120, header: 'Status' },

  // Manual fill columns (columns G-S)
  NP: { index: 7, letter: 'G', width: 120, header: 'NP' },
  EMAIL: { index: 8, letter: 'H', width: 220, header: 'Email' },
  EMAIL_2: { index: 9, letter: 'I', width: 220, header: 'Email 2' },
  HP: { index: 10, letter: 'J', width: 140, header: 'HP' },
  JOIN_DATE: { index: 11, letter: 'K', width: 110, header: 'Join Date' },
  TITLE: { index: 12, letter: 'L', width: 150, header: 'Title' },
  ROLE: { index: 13, letter: 'M', width: 150, header: 'Role' },
  LEAD_PIC: { index: 14, letter: 'N', width: 150, header: 'Lead/PIC' },
  STATUS_HIRING: { index: 15, letter: 'O', width: 130, header: 'Status Hiring' },
  AUTOMATION: { index: 16, letter: 'P', width: 140, header: 'Automation' },
  GITHUB: { index: 17, letter: 'Q', width: 180, header: 'Github Personal' },
  VPN_ABC: { index: 18, letter: 'R', width: 120, header: 'VPN ABC' },
  VPN_HUWAWEI: { index: 19, letter: 'S', width: 140, header: 'VPN Huwawei' }
};

const TEAM_TOTAL_COLUMNS = 19;
const TEAM_AUTO_GENERATED_COLS = 6; // Columns A-F are auto-generated

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

  // Create header (Auto-generated left, Manual fill right)
  const headers = [
    // Auto-generated columns
    TEAM_COLUMNS.NO.header,
    TEAM_COLUMNS.NAME.header,
    TEAM_COLUMNS.PROJECTS.header,
    TEAM_COLUMNS.MODUL.header,
    TEAM_COLUMNS.SUBMODUL.header,
    TEAM_COLUMNS.STATUS.header,
    // Manual fill columns
    TEAM_COLUMNS.NP.header,
    TEAM_COLUMNS.EMAIL.header,
    TEAM_COLUMNS.EMAIL_2.header,
    TEAM_COLUMNS.HP.header,
    TEAM_COLUMNS.JOIN_DATE.header,
    TEAM_COLUMNS.TITLE.header,
    TEAM_COLUMNS.ROLE.header,
    TEAM_COLUMNS.LEAD_PIC.header,
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

  // Add sample data with new structure (Auto-generated left, Manual fill right)
  const sampleData = [
    // [No, Name, Project, Modul, Submodul, Status, NP, Email, Email2, HP, JoinDate, Title, Role, Lead, StatusHiring, Automation, Github, VPN_ABC, VPN_Huwawei]
    [1, 'Samuel Pangestu', 'Government Systems', 'SIPGN', 'Core System, Integration Module', 'Active', 'NP001', 'samuel.gonggom@inadigital.co.id', '', '081234567890', '2023-01-15', 'Senior QA Engineer', 'QA Team Lead', 'Samuel', 'Permanent', 'Selenium, Playwright', 'github.com/samuel', 'Active', 'Active'],
    [2, 'Muhammad Lutfi', 'Digital Peruri', 'INADigital', 'INAgov, Emeterai', 'Active', 'NP002', 'muhamad.ramdani@inadigital.co.id', '', '081234567891', '2023-02-01', 'Senior QA Engineer', 'QA Team Lead', 'Lutfi', 'Permanent', 'Cypress, JMeter', 'github.com/lutfi', 'Active', 'Active'],
    [3, 'Irvan Muhandis', 'Digital Peruri', 'INADigital', 'INAgov, Wahana', 'Active', 'NP003', 'irvan.muhandis@inadigital.co.id', '', '081234567892', '2023-03-10', 'QA Engineer', 'PIC Project', 'Irvan', 'Permanent', 'Postman, K6', 'github.com/irvan', 'Active', 'Active'],
    [4, 'Muhammad Rizky', 'Digital Peruri', 'INADigital', 'Wahana, COTS', 'Active', 'NP004', 'muhammad.ferdiansyah@inadigital.co.id', '', '081234567893', '2023-04-20', 'QA Engineer', 'Quality Engineer', 'Rizky', 'Contract', 'Manual Testing', 'github.com/rizky', 'Active', 'Pending']
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
    const name = row[1] ? row[1].toString().trim() : ''; // Column B (Name)
    const projects = row[2] ? row[2].toString().trim() : ''; // Column C (Projects)
    const modul = row[3] ? row[3].toString().trim() : ''; // Column D (Modul)
    const submodul = row[4] ? row[4].toString().trim() : ''; // Column E (Submodul)
    const status = row[5] ? row[5].toString().trim() : ''; // Column F (Status)
    const email = row[7] ? row[7].toString().trim() : ''; // Column H (Email)
    const memberRole = row[12] ? row[12].toString().trim() : ''; // Column M (Role)

    if (name && status === 'Active' && memberRole === role) {
      members.push({
        name: name,
        role: memberRole,
        projects: projects.split(',').map(p => p.trim()).filter(p => p),
        modul: modul.split(',').map(m => m.trim()).filter(m => m),
        submodul: submodul.split(',').map(s => s.trim()).filter(s => s),
        email: email
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
    const name = row[1] ? row[1].toString().trim() : ''; // Column B (Name)
    const status = row[5] ? row[5].toString().trim() : ''; // Column F (Status)

    if (name && status === 'Active') {
      members.push({
        name: name,
        role: row[12] ? row[12].toString().trim() : '', // Column M (Role)
        projects: row[2] ? row[2].toString().split(',').map(p => p.trim()).filter(p => p) : [], // Column C (Projects)
        modul: row[3] ? row[3].toString().split(',').map(m => m.trim()).filter(m => m) : [], // Column D (Modul)
        submodul: row[4] ? row[4].toString().split(',').map(s => s.trim()).filter(s => s) : [], // Column E (Submodul)
        email: row[7] ? row[7].toString().trim() : '' // Column H (Email)
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
    const projectName = row[0] ? row[0].toString().trim() : ''; // Column A
    const modulName = row[1] ? row[1].toString().trim() : ''; // Column B
    const submodulName = row[2] ? row[2].toString().trim() : ''; // Column C
    const picQA = row[3] ? row[3].toString().trim() : ''; // Column D (PIC QA)
    const status = row[9] ? row[9].toString().trim() : ''; // Column J (Status)

    if (projectName && modulName && submodulName && picQA && status === 'Active') {
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

  // Read existing Team Members data to preserve manual fills (columns G-S)
  const teamLastRow = teamSheet.getLastRow();
  const existingData = teamLastRow >= TEAM_DATA_START_ROW ?
    teamSheet.getRange(TEAM_DATA_START_ROW, 1, teamLastRow - TEAM_DATA_START_ROW + 1, TEAM_TOTAL_COLUMNS).getValues() : [];

  // Build map of existing manual data by name
  const existingManualDataMap = new Map();
  existingData.forEach(row => {
    const name = row[1] ? row[1].toString().trim() : ''; // Column B (Name)
    if (name) {
      // Preserve manual fill columns (G-S: indices 6-18)
      existingManualDataMap.set(name, row.slice(6, 19)); // Columns G-S
    }
  });

  // Generate Team Members data (preserve manual data if exists)
  const teamData = [];
  let rowNum = 1;

  picMap.forEach((data, picName) => {
    const projects = Array.from(data.projects).join(', ');
    const moduls = Array.from(data.moduls).join(', ');
    const submoduls = Array.from(data.submoduls).join(', ');

    // Get existing manual data or use defaults
    const existingManual = existingManualDataMap.get(picName);

    // Build row: Auto-generated (A-F) + Manual (G-S)
    const row = [
      // Auto-generated columns (A-F)
      rowNum,           // A: No
      picName,          // B: Name
      projects,         // C: Projects
      moduls,           // D: Modul
      submoduls,        // E: Submodul
      'Active'          // F: Status
    ];

    // Manual fill columns (G-S) - preserve if exists, otherwise default
    if (existingManual) {
      row.push(...existingManual); // Preserve existing manual data
    } else {
      row.push(
        '',               // G: NP
        '',               // H: Email
        '',               // I: Email 2
        '',               // J: HP
        '',               // K: Join Date
        '',               // L: Title
        'Quality Engineer', // M: Role - default
        '',               // N: Lead/PIC
        '',               // O: Status Hiring
        '',               // P: Automation
        '',               // Q: Github Personal
        '',               // R: VPN ABC
        ''                // S: VPN Huwawei
      );
    }

    teamData.push(row);
    rowNum++;
  });

  // Clear existing data area first
  if (teamLastRow > TEAM_HEADER_ROW) {
    teamSheet.getRange(TEAM_DATA_START_ROW, 1, teamLastRow - TEAM_HEADER_ROW, TEAM_TOTAL_COLUMNS).clearContent();
  }

  // Write to Team Members tab
  if (teamData.length > 0) {
    teamSheet.getRange(TEAM_DATA_START_ROW, 1, teamData.length, TEAM_TOTAL_COLUMNS).setValues(teamData);

    // Apply formatting
    applyTeamMemberFormatting();

    Logger.log('✅ Generated ' + teamData.length + ' team members from Config (manual data preserved)');

    return {
      success: true,
      generated: teamData.length,
      message: 'Generated ' + teamData.length + ' team members from Config'
    };
  }

  return { success: false, message: 'No team members generated' };
}
