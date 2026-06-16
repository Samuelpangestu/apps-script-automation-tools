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
  // Auto-generated from Project Config (columns A-G)
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  NAME: { index: 2, letter: 'B', width: 180, header: 'Name' },
  PROJECTS: { index: 3, letter: 'C', width: 200, header: 'Project' },
  MODUL: { index: 4, letter: 'D', width: 180, header: 'Modul' },
  SUBMODUL: { index: 5, letter: 'E', width: 250, header: 'Submodul' },
  STATUS: { index: 6, letter: 'F', width: 120, header: 'Status' },
  QA_LEAD: { index: 7, letter: 'G', width: 180, header: 'Submodul Lead' },

  // Manual fill columns (columns H-T)
  NP: { index: 8, letter: 'H', width: 120, header: 'NP' },
  EMAIL: { index: 9, letter: 'I', width: 220, header: 'Email' },
  EMAIL_2: { index: 10, letter: 'J', width: 220, header: 'Email 2' },
  HP: { index: 11, letter: 'K', width: 140, header: 'HP' },
  JOIN_DATE: { index: 12, letter: 'L', width: 110, header: 'Join Date' },
  TITLE: { index: 13, letter: 'M', width: 150, header: 'Title' },
  ROLE: { index: 14, letter: 'N', width: 150, header: 'Role' },
  LEAD_PIC: { index: 15, letter: 'O', width: 150, header: 'Lead/PIC' },
  STATUS_HIRING: { index: 16, letter: 'P', width: 130, header: 'Status Hiring' },
  AUTOMATION: { index: 17, letter: 'Q', width: 140, header: 'Automation' },
  GITHUB: { index: 18, letter: 'R', width: 180, header: 'Github Personal' },
  VPN_ABC: { index: 19, letter: 'S', width: 120, header: 'VPN ABC' },
  VPN_HUWAWEI: { index: 20, letter: 'T', width: 140, header: 'VPN Huwawei' }
};

const TEAM_TOTAL_COLUMNS = 20;
const TEAM_AUTO_GENERATED_COLS = 7; // Columns A-G are auto-generated

// Column indices for array access (0-based)
const TEAM_COL = {
  NO: 0,             // Column A
  NAME: 1,           // Column B
  PROJECTS: 2,       // Column C
  MODUL: 3,          // Column D
  SUBMODUL: 4,       // Column E
  STATUS: 5,         // Column F
  QA_LEAD: 6,        // Column G (Submodul Lead)
  NP: 7,             // Column H
  EMAIL: 8,          // Column I
  EMAIL_2: 9,        // Column J
  HP: 10,            // Column K
  JOIN_DATE: 11,     // Column L
  TITLE: 12,         // Column M
  ROLE: 13,          // Column N
  LEAD_PIC: 14,      // Column O
  STATUS_HIRING: 15, // Column P
  AUTOMATION: 16,    // Column Q
  GITHUB: 17,        // Column R
  VPN_ABC: 18,       // Column S
  VPN_HUWAWEI: 19    // Column T
};

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
    TEAM_COLUMNS.QA_LEAD.header,
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
    // [No, Name, Project, Modul, Submodul, Status, QA Lead, NP, Email, Email2, HP, JoinDate, Title, Role, Lead, StatusHiring, Automation, Github, VPN_ABC, VPN_Huwawei]
    [1, 'Samuel Pangestu', 'Government Systems', 'SIPGN', 'Core System, Integration Module', 'Active', 'Samuel', 'NP001', 'samuel.gonggom@inadigital.co.id', '', '081234567890', '2023-01-15', 'Senior QA Engineer', 'QA Team Lead', 'Samuel', 'Permanent', 'Selenium, Playwright', 'github.com/samuel', 'Active', 'Active'],
    [2, 'Muhammad Lutfi', 'Digital Peruri', 'INADigital', 'INAgov, Emeterai', 'Active', 'Lutfi', 'NP002', 'muhamad.ramdani@inadigital.co.id', '', '081234567891', '2023-02-01', 'Senior QA Engineer', 'QA Team Lead', 'Lutfi', 'Permanent', 'Cypress, JMeter', 'github.com/lutfi', 'Active', 'Active'],
    [3, 'Irvan Muhandis', 'Digital Peruri', 'INADigital', 'INAgov, Wahana', 'Active', 'Lutfi', 'NP003', 'irvan.muhandis@inadigital.co.id', '', '081234567892', '2023-03-10', 'QA Engineer', 'PIC Project', 'Irvan', 'Permanent', 'Postman, K6', 'github.com/irvan', 'Active', 'Active'],
    [4, 'Muhammad Rizky', 'Digital Peruri', 'INADigital', 'Wahana, COTS', 'Active', 'Lutfi', 'NP004', 'muhammad.ferdiansyah@inadigital.co.id', '', '081234567893', '2023-04-20', 'QA Engineer', 'Quality Engineer', 'Rizky', 'Contract', 'Manual Testing', 'github.com/rizky', 'Active', 'Pending']
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
 * NOTE: Only applies validation to MANUAL columns (Role).
 * Auto-generated columns (Project, Modul, Submodul, Status) are skipped
 * because they can contain comma-separated values that don't match the helper lists.
 */
function addTeamDataValidation(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!configSheet) {
    Logger.log('⚠️ Project tab not found, skipping dropdown validation');
    return;
  }

  // IMPORTANT: Clear existing validation from auto-generated columns
  // These columns can contain comma-separated values and should not have validation
  Logger.log('🧹 Clearing validation from auto-generated columns (entire columns)...');
  sheet.getRange(1, TEAM_COLUMNS.PROJECTS.index, sheet.getMaxRows(), 1).clearDataValidations(); // Column C: Project (entire column)
  sheet.getRange(1, TEAM_COLUMNS.MODUL.index, sheet.getMaxRows(), 1).clearDataValidations();    // Column D: Modul (entire column)
  sheet.getRange(1, TEAM_COLUMNS.SUBMODUL.index, sheet.getMaxRows(), 1).clearDataValidations(); // Column E: Submodul (entire column)
  sheet.getRange(1, TEAM_COLUMNS.STATUS.index, sheet.getMaxRows(), 1).clearDataValidations();   // Column F: Status (entire column)
  sheet.getRange(1, TEAM_COLUMNS.QA_LEAD.index, sheet.getMaxRows(), 1).clearDataValidations();  // Column G: QA Lead (entire column)

  // Role validation (apply to reasonable range) - MANUAL COLUMN (now column N)
  const roleRange = sheet.getRange(TEAM_DATA_START_ROW, TEAM_COLUMNS.ROLE.index, 100);
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['QA Team Lead', 'QA Lead', 'PIC Project', 'Quality Engineer', 'Senior Quality Engineer', 'Intern Quality Engineer', 'Security Engineer', 'UX Research'], true)
    .build();
  roleRange.setDataValidation(roleRule);

  Logger.log('✅ Data validation added (manual columns only)');
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
    const name = row[TEAM_COL.NAME] ? row[TEAM_COL.NAME].toString().trim() : '';
    const projects = row[TEAM_COL.PROJECTS] ? row[TEAM_COL.PROJECTS].toString().trim() : '';
    const modul = row[TEAM_COL.MODUL] ? row[TEAM_COL.MODUL].toString().trim() : '';
    const submodul = row[TEAM_COL.SUBMODUL] ? row[TEAM_COL.SUBMODUL].toString().trim() : '';
    const status = row[TEAM_COL.STATUS] ? row[TEAM_COL.STATUS].toString().trim() : '';
    const email = row[TEAM_COL.EMAIL] ? row[TEAM_COL.EMAIL].toString().trim() : '';
    const memberRole = row[TEAM_COL.ROLE] ? row[TEAM_COL.ROLE].toString().trim() : '';

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
    const name = row[TEAM_COL.NAME] ? row[TEAM_COL.NAME].toString().trim() : '';
    const status = row[TEAM_COL.STATUS] ? row[TEAM_COL.STATUS].toString().trim() : '';

    if (name && status === 'Active') {
      members.push({
        name: name,
        role: row[TEAM_COL.ROLE] ? row[TEAM_COL.ROLE].toString().trim() : '',
        projects: row[TEAM_COL.PROJECTS] ? row[TEAM_COL.PROJECTS].toString().split(',').map(p => p.trim()).filter(p => p) : [],
        modul: row[TEAM_COL.MODUL] ? row[TEAM_COL.MODUL].toString().split(',').map(m => m.trim()).filter(m => m) : [],
        submodul: row[TEAM_COL.SUBMODUL] ? row[TEAM_COL.SUBMODUL].toString().split(',').map(s => s.trim()).filter(s => s) : [],
        qaLead: row[TEAM_COL.QA_LEAD] ? row[TEAM_COL.QA_LEAD].toString().split(',').map(q => q.trim()).filter(q => q) : [],
        email: row[TEAM_COL.EMAIL] ? row[TEAM_COL.EMAIL].toString().trim() : ''
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
  const picMap = new Map(); // Map<picName, {projects: Set, moduls: Set, submoduls: Set, qaLeads: Set}>

  configData.forEach(row => {
    const projectName = row[CONFIG_COL.PROJECT] ? row[CONFIG_COL.PROJECT].toString().trim() : '';
    const modulName = row[CONFIG_COL.MODUL] ? row[CONFIG_COL.MODUL].toString().trim() : '';
    const submodulName = row[CONFIG_COL.SUBMODUL] ? row[CONFIG_COL.SUBMODUL].toString().trim() : '';
    const picQA = row[CONFIG_COL.PIC_QA] ? row[CONFIG_COL.PIC_QA].toString().trim() : '';
    const status = row[CONFIG_COL.STATUS] ? row[CONFIG_COL.STATUS].toString().trim() : '';
    const qaLead = row[CONFIG_COL.QA_LEAD] ? row[CONFIG_COL.QA_LEAD].toString().trim() : '';

    if (projectName && modulName && submodulName && picQA && status === 'Active') {
      // Parse multiple PICs (comma-separated)
      const picNames = picQA.split(',').map(name => name.trim()).filter(name => name);

      picNames.forEach(picName => {
        if (!picMap.has(picName)) {
          picMap.set(picName, {
            projects: new Set(),
            moduls: new Set(),
            submoduls: new Set(),
            qaLeads: new Set()
          });
        }

        const picData = picMap.get(picName);
        picData.projects.add(projectName);
        picData.moduls.add(modulName);
        picData.submoduls.add(submodulName);
        if (qaLead) picData.qaLeads.add(qaLead); // Add QA Lead if exists
      });
    }
  });

  if (picMap.size === 0) {
    Logger.log('⚠️ No PICs found in Config tab');
    return { success: false, message: 'No PICs found in Config' };
  }

  // Read existing Team Members data to preserve manual fills (columns H-T)
  const teamLastRow = teamSheet.getLastRow();
  const existingData = teamLastRow >= TEAM_DATA_START_ROW ?
    teamSheet.getRange(TEAM_DATA_START_ROW, 1, teamLastRow - TEAM_DATA_START_ROW + 1, TEAM_TOTAL_COLUMNS).getValues() : [];

  // Build map of existing manual data by name
  const existingManualDataMap = new Map();
  existingData.forEach(row => {
    const name = row[TEAM_COL.NAME] ? row[TEAM_COL.NAME].toString().trim() : '';
    if (name) {
      // Preserve manual fill columns (H-T: indices 7-19)
      existingManualDataMap.set(name, row.slice(TEAM_COL.NP, TEAM_COL.VPN_HUWAWEI + 1)); // Columns H-T
    }
  });

  // Generate Team Members data (preserve manual data if exists)
  const teamData = [];
  let rowNum = 1;

  picMap.forEach((data, picName) => {
    const projects = Array.from(data.projects).join(', ');
    const moduls = Array.from(data.moduls).join(', ');
    const submoduls = Array.from(data.submoduls).join(', ');
    const qaLeads = Array.from(data.qaLeads).join(', '); // Aggregate QA Leads

    // Get existing manual data or use defaults
    const existingManual = existingManualDataMap.get(picName);

    // Build row: Auto-generated (A-G) + Manual (H-T)
    const row = [
      // Auto-generated columns (A-G)
      rowNum,           // A: No
      picName,          // B: Name
      projects,         // C: Projects
      moduls,           // D: Modul
      submoduls,        // E: Submodul
      'Active',         // F: Status
      qaLeads           // G: QA Lead
    ];

    // Manual fill columns (H-T) - preserve if exists, otherwise default
    if (existingManual) {
      row.push(...existingManual); // Preserve existing manual data
    } else {
      row.push(
        '',               // H: NP
        '',               // I: Email
        '',               // J: Email 2
        '',               // K: HP
        '',               // L: Join Date
        '',               // M: Title
        'Quality Engineer', // N: Role - default
        '',               // O: Lead/PIC
        '',               // P: Status Hiring
        '',               // Q: Automation
        '',               // R: Github Personal
        '',               // S: VPN ABC
        ''                // T: VPN Huwawei
      );
    }

    teamData.push(row);
    rowNum++;
  });

  // Clear existing data area first
  if (teamLastRow > TEAM_HEADER_ROW) {
    teamSheet.getRange(TEAM_DATA_START_ROW, 1, teamLastRow - TEAM_HEADER_ROW, TEAM_TOTAL_COLUMNS).clearContent();
  }

  // CRITICAL: Clear data validation from auto-generated columns BEFORE writing data
  // These columns can contain comma-separated values that will violate single-value validation rules
  Logger.log('🧹 Clearing validation from auto-generated columns (entire columns) before writing...');
  teamSheet.getRange(1, TEAM_COLUMNS.PROJECTS.index, teamSheet.getMaxRows(), 1).clearDataValidations(); // Column C: Project (entire column)
  teamSheet.getRange(1, TEAM_COLUMNS.MODUL.index, teamSheet.getMaxRows(), 1).clearDataValidations();    // Column D: Modul (entire column)
  teamSheet.getRange(1, TEAM_COLUMNS.SUBMODUL.index, teamSheet.getMaxRows(), 1).clearDataValidations(); // Column E: Submodul (entire column)
  teamSheet.getRange(1, TEAM_COLUMNS.STATUS.index, teamSheet.getMaxRows(), 1).clearDataValidations();   // Column F: Status (entire column)
  teamSheet.getRange(1, TEAM_COLUMNS.QA_LEAD.index, teamSheet.getMaxRows(), 1).clearDataValidations();  // Column G: QA Lead (entire column)
  SpreadsheetApp.flush(); // Ensure validation clearing is applied before writing data

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
