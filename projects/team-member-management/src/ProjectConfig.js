/**
 * ProjectConfig.js — Centralized Configuration Management
 * ═══════════════════════════════════════════════════════════════════════
 * Flat table format (Portfolio style) for team management
 * ═══════════════════════════════════════════════════════════════════════
 */

const CONFIG_TAB_NAME = 'Config';

// Flat table structure (Portfolio style)
const CONFIG_HEADER_ROW = 5;
const CONFIG_DATA_START_ROW = 6;
const CONFIG_COLUMNS = {
  ACTIVE: { index: 1, letter: 'A', width: 80, header: 'Active' },
  PROJECT: { index: 2, letter: 'B', width: 150, header: 'Project' },
  MODUL: { index: 3, letter: 'C', width: 150, header: 'Modul' },
  SUBMODUL: { index: 4, letter: 'D', width: 200, header: 'Submodul' },
  PIC_QA: { index: 5, letter: 'E', width: 150, header: 'PIC QA' },
  DIFFICULTY: { index: 6, letter: 'F', width: 80, header: 'Diff (1-10)' },
  RISK: { index: 7, letter: 'G', width: 80, header: 'Risk (1-10)' },
  COMPLEXITY: { index: 8, letter: 'H', width: 80, header: 'Comp (1-10)' },
  DESCRIPTION: { index: 9, letter: 'I', width: 250, header: 'Description' },
  STATUS: { index: 10, letter: 'J', width: 100, header: 'Status' }
};
const CONFIG_TOTAL_COLUMNS = 10;

/**
 * Create Config tab (flat table format - Portfolio style)
 */
function createConfigTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Delete existing if any
  let existingSheet = ss.getSheetByName(CONFIG_TAB_NAME);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

  // Create new sheet
  const sheet = ss.insertSheet(CONFIG_TAB_NAME, 0);

  let currentRow = 1;

  // Title
  sheet.getRange(currentRow, 1, 1, CONFIG_TOTAL_COLUMNS).merge()
    .setValue('⚙️ QA TEAM MANAGEMENT — Module Configuration')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Instruction row
  sheet.getRange(currentRow, 1, 1, CONFIG_TOTAL_COLUMNS).merge()
    .setValue('Spreadsheet ID ada di URL modul: docs.google.com/spreadsheets/d/[ID]/edit  |  Data akan otomatis diambil dari production saat sync')
    .setBackground('#fff3cd')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 30);
  currentRow++;

  // Production Spreadsheet ID field
  const syncSettings = getSyncSettings();
  sheet.getRange(currentRow, 1).setValue('Production Spreadsheet ID:').setFontWeight('bold').setBackground('#f8f9fa');
  sheet.getRange(currentRow, 2, 1, CONFIG_TOTAL_COLUMNS - 1).merge()
    .setValue(syncSettings.spreadsheetId || 'Paste production ID or URL here')
    .setBackground('#ffffff')
    .setFontStyle('italic')
    .setHorizontalAlignment('left');
  sheet.setRowHeight(currentRow, 25);
  currentRow += 2;

  // Main table header
  const headers = [
    CONFIG_COLUMNS.ACTIVE.header,
    CONFIG_COLUMNS.PROJECT.header,
    CONFIG_COLUMNS.MODUL.header,
    CONFIG_COLUMNS.SUBMODUL.header,
    CONFIG_COLUMNS.PIC_QA.header,
    CONFIG_COLUMNS.DIFFICULTY.header,
    CONFIG_COLUMNS.RISK.header,
    CONFIG_COLUMNS.COMPLEXITY.header,
    CONFIG_COLUMNS.DESCRIPTION.header,
    CONFIG_COLUMNS.STATUS.header
  ];

  sheet.getRange(CONFIG_HEADER_ROW, 1, 1, CONFIG_TOTAL_COLUMNS)
    .setValues([headers])
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(CONFIG_HEADER_ROW, 40);
  sheet.setFrozenRows(CONFIG_HEADER_ROW);

  // Sample data (flat table format with PIC QA)
  const sampleData = [
    [true, 'SIPGN', 'E2E DAPUR', 'E2E DAPUR', 'Dini', 7, 6, 7, 'E2E DAPUR module', 'Active'],
    [true, 'SIPGN', 'AtomAPI', 'AtomAPI PM', 'Arief', 6, 5, 6, 'AtomAPI Project Management', 'Active'],
    [true, 'SIPGN', '1', '1.1', 'Reizha', 8, 7, 8, 'SIPGN Module 1.1', 'Active'],
    [true, 'SIPGN', '1', '1.2', 'Reizha', 7, 6, 7, 'SIPGN Module 1.2', 'Active'],
    [true, 'SIPGN', '1', '1.3', 'Reizha', 6, 5, 6, 'SIPGN Module 1.3', 'Active'],
    [true, 'SIPGN', '2', '2.1', 'Farhan', 7, 6, 7, 'SIPGN Module 2.1', 'Active'],
    [true, 'SIPGN', '2', '2.2', 'Farhan', 6, 5, 6, 'SIPGN Module 2.2', 'Active'],
    [true, 'SIPGN', '4', '4.1', 'Adinda, Denta', 7, 7, 7, 'SIPGN Module 4.1', 'Active'],
    [true, 'SIPGN', '4', '4.2', 'Adinda, Denta', 6, 5, 6, 'SIPGN Module 4.2', 'Active'],
    [true, 'INADigital', 'INAgov', 'INAgov', 'Irvan', 8, 7, 8, 'Indonesia Government Portal', 'Active'],
    [true, 'INADigital', 'POS', 'POS', 'Fresma', 5, 4, 5, 'Point of Sale system', 'Active'],
    [true, 'INADigital', 'SCM', 'SCM', 'Fresma', 6, 5, 6, 'Supply Chain Management', 'Active'],
    [true, 'INADigital', 'PERURIID', 'PERURIID', 'Imam, Farhan', 7, 6, 7, 'Peruri ID system', 'Active'],
    [false, 'COTS', 'COTS', 'CODEBASE', 'Daffa, Zahwa', 3, 2, 3, 'COTS Codebase', 'Inactive'],
    [false, 'COTS', 'COTS', 'GEODIPA', 'Daffa, Zahwa', 3, 2, 3, 'COTS Geodipa', 'Inactive']
  ];

  sheet.getRange(CONFIG_DATA_START_ROW, 1, sampleData.length, CONFIG_TOTAL_COLUMNS)
    .setValues(sampleData);

  // Apply formatting
  for (let i = 0; i < sampleData.length; i++) {
    const rowNum = CONFIG_DATA_START_ROW + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, CONFIG_TOTAL_COLUMNS).setBackground(bg);
  }

  currentRow = CONFIG_DATA_START_ROW + sampleData.length;

  // Add 40 blank rows for future entries
  for (let i = 0; i < 40; i++) {
    const rowNum = currentRow + i;
    const bg = (sampleData.length + i) % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, CONFIG_TOTAL_COLUMNS).setBackground(bg);
  }

  // Set column widths
  Object.values(CONFIG_COLUMNS).forEach(col => {
    sheet.setColumnWidth(col.index, col.width);
  });

  // Data validation for Active column
  const activeRange = sheet.getRange(CONFIG_DATA_START_ROW, CONFIG_COLUMNS.ACTIVE.index, 100);
  const activeRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  activeRange.setDataValidation(activeRule);

  // Data validation for Status column
  const statusRange = sheet.getRange(CONFIG_DATA_START_ROW, CONFIG_COLUMNS.STATUS.index, 100);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'On Hold'], true)
    .build();
  statusRange.setDataValidation(statusRule);

  // Flush changes to ensure completion
  SpreadsheetApp.flush();

  Logger.log('✅ Config tab created (flat table format)');
}

/**
 * Get all active projects (unique list from config table)
 */
function getActiveProjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_DATA_START_ROW) return [];

  const data = sheet.getRange(CONFIG_DATA_START_ROW, 1, lastRow - CONFIG_DATA_START_ROW + 1, CONFIG_TOTAL_COLUMNS).getValues();

  const projectsMap = new Map();
  data.forEach(row => {
    const active = row[0];
    const projectName = row[1] ? row[1].toString().trim() : '';
    const status = row[9] ? row[9].toString().trim() : ''; // Column J (index 9)

    if (active && projectName && status === 'Active') {
      if (!projectsMap.has(projectName)) {
        projectsMap.set(projectName, {
          name: projectName,
          description: ''
        });
      }
    }
  });

  return Array.from(projectsMap.values());
}

/**
 * Get all active modul (unique list from config table)
 */
function getActiveModul() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_DATA_START_ROW) return [];

  const data = sheet.getRange(CONFIG_DATA_START_ROW, 1, lastRow - CONFIG_DATA_START_ROW + 1, CONFIG_TOTAL_COLUMNS).getValues();

  const modulsMap = new Map();
  data.forEach(row => {
    const active = row[0];
    const projectName = row[1] ? row[1].toString().trim() : '';
    const modulName = row[2] ? row[2].toString().trim() : '';
    const status = row[9] ? row[9].toString().trim() : ''; // Column J (index 9)

    if (active && projectName && modulName && status === 'Active') {
      const key = projectName + '|' + modulName;
      if (!modulsMap.has(key)) {
        modulsMap.set(key, {
          name: modulName,
          project: projectName,
          description: ''
        });
      }
    }
  });

  return Array.from(modulsMap.values());
}

/**
 * Get all active submodul (from config table)
 */
function getActiveSubmodul() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_DATA_START_ROW) return [];

  const data = sheet.getRange(CONFIG_DATA_START_ROW, 1, lastRow - CONFIG_DATA_START_ROW + 1, CONFIG_TOTAL_COLUMNS).getValues();

  const submoduls = [];
  data.forEach(row => {
    const active = row[0];
    const projectName = row[1] ? row[1].toString().trim() : '';
    const modulName = row[2] ? row[2].toString().trim() : '';
    const submodulName = row[3] ? row[3].toString().trim() : '';
    // row[4] is PIC QA - skip
    const difficulty = row[5] || 0;
    const risk = row[6] || 0;
    const complexity = row[7] || 0;
    const description = row[8] ? row[8].toString().trim() : '';
    const status = row[9] ? row[9].toString().trim() : ''; // Column J (index 9)

    if (active && projectName && modulName && submodulName && status === 'Active') {
      submoduls.push({
        name: submodulName,
        modul: modulName,
        project: projectName,
        difficulty: difficulty,
        risk: risk,
        complexity: complexity,
        description: description
      });
    }
  });

  return submoduls;
}

/**
 * Get submodul ratings
 */
function getSubmodulRatings(submodulName) {
  const submoduls = getActiveSubmodul();
  const found = submoduls.find(s => s.name === submodulName);
  return found ? {
    difficulty: found.difficulty,
    risk: found.risk,
    complexity: found.complexity
  } : { difficulty: 0, risk: 0, complexity: 0 };
}
