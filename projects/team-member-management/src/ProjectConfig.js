/**
 * ProjectConfig.js — Centralized Configuration Management
 * ═══════════════════════════════════════════════════════════════════════
 * Manage modul, submodul, and difficulty levels - all in one place
 * ═══════════════════════════════════════════════════════════════════════
 */

const CONFIG_TAB_NAME = 'Config';

// Project section
const PROJECT_HEADER_ROW = 9;
const PROJECT_DATA_START_ROW = 10;
const PROJECT_COLUMNS = {
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  PROJECT_NAME: { index: 2, letter: 'B', width: 200, header: 'Project' },
  DESCRIPTION: { index: 3, letter: 'C', width: 300, header: 'Description' },
  STATUS: { index: 4, letter: 'D', width: 100, header: 'Status' }
};
const PROJECT_TOTAL_COLUMNS = 4;

// Modul section
const MODUL_HEADER_ROW = 14; // Dynamic
const MODUL_DATA_START_ROW = 15;
const MODUL_COLUMNS = {
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  MODUL_NAME: { index: 2, letter: 'B', width: 200, header: 'Modul' },
  PROJECT: { index: 3, letter: 'C', width: 150, header: 'Project' },
  DESCRIPTION: { index: 4, letter: 'D', width: 300, header: 'Description' },
  STATUS: { index: 5, letter: 'E', width: 100, header: 'Status' }
};
const MODUL_TOTAL_COLUMNS = 5;

// Submodul section
const SUBMODUL_HEADER_ROW = 20; // Dynamic
const SUBMODUL_DATA_START_ROW = 21;
const SUBMODUL_COLUMNS = {
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  SUBMODUL_NAME: { index: 2, letter: 'B', width: 250, header: 'Submodul' },
  MODUL: { index: 3, letter: 'C', width: 150, header: 'Modul' },
  DIFFICULTY: { index: 4, letter: 'D', width: 100, header: 'Difficulty (1-10)' },
  RISK: { index: 5, letter: 'E', width: 100, header: 'Risk (1-10)' },
  COMPLEXITY: { index: 6, letter: 'F', width: 100, header: 'Complexity (1-10)' },
  DESCRIPTION: { index: 7, letter: 'G', width: 250, header: 'Description' },
  STATUS: { index: 8, letter: 'H', width: 100, header: 'Status' }
};
const SUBMODUL_TOTAL_COLUMNS = 8;

/**
 * Create Config tab (centralized configuration)
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
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('⚙️ CENTRALIZED CONFIGURATION')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 45);
  currentRow += 2;

  // Rating Scale Definitions
  sheet.getRange(currentRow, 1, 1, 8).merge()
    .setValue('📋 RATING SCALE DEFINITIONS (1-10)')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 30);
  currentRow++;

  const definitions = [
    ['Rating', 'Difficulty', 'Risk', 'Complexity', 'Team Size Recommendation'],
    [
      '1-3 (Low)',
      'Simple, well-documented',
      'Low impact if fails',
      'Few components, straightforward',
      '1-2 QE'
    ],
    [
      '4-6 (Medium)',
      'Moderate complexity',
      'Moderate business impact',
      'Multiple modules, some integration',
      '1 PIC + 2-3 QE'
    ],
    [
      '7-10 (High)',
      'Highly complex, critical',
      'High business/security impact',
      'Many integrations, critical systems',
      '1 Team Lead + 1 PIC + 3-5 QE'
    ]
  ];

  sheet.getRange(currentRow, 1, definitions.length, 5).setValues(definitions);

  // Format definition header
  sheet.getRange(currentRow, 1, 1, 5)
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  currentRow++;

  // Format definition rows
  for (let i = 0; i < definitions.length - 1; i++) {
    const rowNum = currentRow + i;
    let bg = '#ffffff';
    if (definitions[i + 1][0].includes('Low')) bg = '#d4edda';
    else if (definitions[i + 1][0].includes('Medium')) bg = '#fff3cd';
    else if (definitions[i + 1][0].includes('High')) bg = '#f8d7da';

    sheet.getRange(rowNum, 1, 1, 5)
      .setBackground(bg)
      .setWrap(true)
      .setVerticalAlignment('top');

    sheet.setRowHeight(rowNum, 60);
  }

  // Set column widths for definition section
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 200);

  currentRow += definitions.length;
  currentRow++; // Empty row

  // ═════════════════════════════════════════════════════════════
  // PROJECT LIST SECTION
  // ═════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 8).merge()
    .setValue('📁 PROJECT LIST')
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Project headers
  const projectHeaders = [
    PROJECT_COLUMNS.NO.header,
    PROJECT_COLUMNS.PROJECT_NAME.header,
    PROJECT_COLUMNS.DESCRIPTION.header,
    PROJECT_COLUMNS.STATUS.header
  ];

  sheet.getRange(currentRow, 1, 1, PROJECT_TOTAL_COLUMNS)
    .setValues([projectHeaders])
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;

  // Project data
  const projectData = [
    [1, 'Digital Peruri', 'Digital transformation initiatives', 'Active'],
    [2, 'Government Systems', 'Government integration projects', 'Active']
  ];

  const projectStartRow = currentRow;
  sheet.getRange(projectStartRow, 1, projectData.length, PROJECT_TOTAL_COLUMNS)
    .setValues(projectData);

  // Apply formatting
  for (let i = 0; i < projectData.length; i++) {
    const rowNum = projectStartRow + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, PROJECT_TOTAL_COLUMNS).setBackground(bg);
  }

  currentRow += projectData.length;

  // Add 5 blank rows for future project entries
  for (let i = 0; i < 5; i++) {
    const rowNum = currentRow + i;
    const bg = (projectData.length + i) % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, PROJECT_TOTAL_COLUMNS).setBackground(bg);
  }

  currentRow += 5;
  currentRow++; // Empty row

  // ═════════════════════════════════════════════════════════════
  // MODUL LIST SECTION
  // ═════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 8).merge()
    .setValue('📂 MODUL LIST')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Modul headers
  const modulHeaders = [
    MODUL_COLUMNS.NO.header,
    MODUL_COLUMNS.MODUL_NAME.header,
    MODUL_COLUMNS.PROJECT.header,
    MODUL_COLUMNS.DESCRIPTION.header,
    MODUL_COLUMNS.STATUS.header
  ];

  sheet.getRange(currentRow, 1, 1, MODUL_TOTAL_COLUMNS)
    .setValues([modulHeaders])
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;

  // Modul data
  const modulData = [
    [1, 'INADigital', 'Digital Peruri', 'Digital identity and services', 'Active'],
    [2, 'SIPGN', 'Government Systems', 'Government integration platform', 'Active']
  ];

  const modulStartRow = currentRow;
  sheet.getRange(modulStartRow, 1, modulData.length, MODUL_TOTAL_COLUMNS)
    .setValues(modulData);

  // Apply formatting
  for (let i = 0; i < modulData.length; i++) {
    const rowNum = modulStartRow + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, MODUL_TOTAL_COLUMNS).setBackground(bg);
  }

  currentRow += modulData.length;

  // Add 15 blank rows for future modul entries
  for (let i = 0; i < 15; i++) {
    const rowNum = currentRow + i;
    const bg = (modulData.length + i) % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, MODUL_TOTAL_COLUMNS).setBackground(bg);
  }

  currentRow += 15;
  currentRow++; // Empty row

  // ═════════════════════════════════════════════════════════════
  // SUBMODUL LIST SECTION
  // ═════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 8).merge()
    .setValue('📋 SUBMODUL LIST (with Ratings 1-10)')
    .setBackground('#fbbc04')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Submodul headers
  const submodulHeaders = [
    SUBMODUL_COLUMNS.NO.header,
    SUBMODUL_COLUMNS.SUBMODUL_NAME.header,
    SUBMODUL_COLUMNS.MODUL.header,
    SUBMODUL_COLUMNS.DIFFICULTY.header,
    SUBMODUL_COLUMNS.RISK.header,
    SUBMODUL_COLUMNS.COMPLEXITY.header,
    SUBMODUL_COLUMNS.DESCRIPTION.header,
    SUBMODUL_COLUMNS.STATUS.header
  ];

  sheet.getRange(currentRow, 1, 1, SUBMODUL_TOTAL_COLUMNS)
    .setValues([submodulHeaders])
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;

  // Sample submodul data (with ratings 1-10)
  const submodulData = [
    [1, 'INAgov', 'INADigital', 7, 6, 7, 'Indonesia Government Portal', 'Active'],
    [2, 'Emeterai', 'INADigital', 6, 5, 6, 'Electronic Stamp System', 'Active'],
    [3, 'Peruri ID', 'INADigital', 3, 2, 3, 'Digital Identity System', 'Active'],
    [4, 'Wahana', 'INADigital', 3, 2, 3, 'Wahana Project', 'Active'],
    [5, 'Digidoc 2.0', 'INADigital', 5, 4, 6, 'Digital Document Management', 'Active'],
    [6, 'Peruri Shield', 'INADigital', 8, 9, 8, 'Security Shield System', 'Active'],
    [7, 'Penjaminan Online', 'INADigital', 6, 5, 5, 'Online Guarantee System', 'Active'],
    [8, 'COTS', 'INADigital', 2, 1, 2, 'Commercial Off-The-Shelf', 'Active'],
    [9, 'Core System', 'SIPGN', 9, 8, 9, 'SIPGN Core System', 'Active'],
    [10, 'Integration Module', 'SIPGN', 8, 7, 8, 'SIPGN Integration', 'Active'],
    [11, 'Data Omnyx', 'SIPGN', 9, 8, 10, 'Omnyx Data Integration', 'Active']
  ];

  const submodulStartRow = currentRow;
  sheet.getRange(submodulStartRow, 1, submodulData.length, SUBMODUL_TOTAL_COLUMNS)
    .setValues(submodulData);

  // Apply formatting
  for (let i = 0; i < submodulData.length; i++) {
    const rowNum = submodulStartRow + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, SUBMODUL_TOTAL_COLUMNS).setBackground(bg);
  }

  currentRow += submodulData.length;

  // Add 25 blank rows for future submodul entries
  for (let i = 0; i < 25; i++) {
    const rowNum = currentRow + i;
    const bg = (submodulData.length + i) % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, SUBMODUL_TOTAL_COLUMNS).setBackground(bg);
  }

  // Set column widths
  sheet.setColumnWidth(1, SUBMODUL_COLUMNS.NO.width);
  sheet.setColumnWidth(2, SUBMODUL_COLUMNS.SUBMODUL_NAME.width);
  sheet.setColumnWidth(3, SUBMODUL_COLUMNS.MODUL.width);
  sheet.setColumnWidth(4, SUBMODUL_COLUMNS.DIFFICULTY.width);
  sheet.setColumnWidth(5, SUBMODUL_COLUMNS.RISK.width);
  sheet.setColumnWidth(6, SUBMODUL_COLUMNS.COMPLEXITY.width);
  sheet.setColumnWidth(7, SUBMODUL_COLUMNS.DESCRIPTION.width);
  sheet.setColumnWidth(8, SUBMODUL_COLUMNS.STATUS.width);

  // Flush changes to ensure completion
  SpreadsheetApp.flush();

  Logger.log('✅ Config tab created');
}

/**
 * Get all active projects
 */
function getActiveProjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < PROJECT_DATA_START_ROW) return [];

  const data = sheet.getRange(PROJECT_DATA_START_ROW, 1, 10, PROJECT_TOTAL_COLUMNS).getValues();

  const projects = [];
  data.forEach(row => {
    const projectName = row[1] ? row[1].toString().trim() : '';
    const status = row[3] ? row[3].toString().trim() : '';

    if (projectName && status === 'Active') {
      projects.push({
        name: projectName,
        description: row[2] ? row[2].toString().trim() : ''
      });
    }
  });

  return projects;
}

/**
 * Get all active modul
 */
function getActiveModul() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < MODUL_DATA_START_ROW) return [];

  const data = sheet.getRange(MODUL_DATA_START_ROW, 1, 10, MODUL_TOTAL_COLUMNS).getValues();

  const moduls = [];
  data.forEach(row => {
    const modulName = row[1] ? row[1].toString().trim() : '';
    const status = row[4] ? row[4].toString().trim() : '';

    if (modulName && status === 'Active') {
      moduls.push({
        name: modulName,
        project: row[2] ? row[2].toString().trim() : '',
        description: row[3] ? row[3].toString().trim() : ''
      });
    }
  });

  return moduls;
}

/**
 * Get all active submodul
 */
function getActiveSubmodul() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < SUBMODUL_DATA_START_ROW) return [];

  const data = sheet.getRange(SUBMODUL_DATA_START_ROW, 1, lastRow - SUBMODUL_DATA_START_ROW + 1, SUBMODUL_TOTAL_COLUMNS).getValues();

  const submoduls = [];
  data.forEach(row => {
    const submodulName = row[1] ? row[1].toString().trim() : '';
    const status = row[7] ? row[7].toString().trim() : '';

    if (submodulName && status === 'Active') {
      submoduls.push({
        name: submodulName,
        modul: row[2] ? row[2].toString().trim() : '',
        difficulty: row[3] || 0,
        risk: row[4] || 0,
        complexity: row[5] || 0,
        description: row[6] ? row[6].toString().trim() : ''
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
