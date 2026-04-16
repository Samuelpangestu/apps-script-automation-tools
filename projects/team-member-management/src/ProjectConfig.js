/**
 * ProjectConfig.js — Project Configuration Management
 * ═══════════════════════════════════════════════════════════════════════
 * Manage project list and difficulty levels
 * ═══════════════════════════════════════════════════════════════════════
 */

const CONFIG_TAB_NAME = 'Config - Projects';
const CONFIG_HEADER_ROW = 1;
const CONFIG_DATA_START_ROW = 2;

const CONFIG_COLUMNS = {
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  PROJECT_NAME: { index: 2, letter: 'B', width: 250, header: 'Project Name' },
  DIFFICULTY: { index: 3, letter: 'C', width: 120, header: 'Difficulty Level' },
  DESCRIPTION: { index: 4, letter: 'D', width: 300, header: 'Description' },
  STATUS: { index: 5, letter: 'E', width: 100, header: 'Status' }
};

const CONFIG_TOTAL_COLUMNS = 5;

/**
 * Create Config - Projects tab
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

  // Setup structure
  sheet.setRowHeight(CONFIG_HEADER_ROW, 40);
  sheet.setFrozenRows(CONFIG_HEADER_ROW);

  // Create header
  const headers = [
    CONFIG_COLUMNS.NO.header,
    CONFIG_COLUMNS.PROJECT_NAME.header,
    CONFIG_COLUMNS.DIFFICULTY.header,
    CONFIG_COLUMNS.DESCRIPTION.header,
    CONFIG_COLUMNS.STATUS.header
  ];

  sheet.getRange(CONFIG_HEADER_ROW, 1, 1, CONFIG_TOTAL_COLUMNS)
    .setValues([headers])
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Set column widths
  sheet.setColumnWidth(CONFIG_COLUMNS.NO.index, CONFIG_COLUMNS.NO.width);
  sheet.setColumnWidth(CONFIG_COLUMNS.PROJECT_NAME.index, CONFIG_COLUMNS.PROJECT_NAME.width);
  sheet.setColumnWidth(CONFIG_COLUMNS.DIFFICULTY.index, CONFIG_COLUMNS.DIFFICULTY.width);
  sheet.setColumnWidth(CONFIG_COLUMNS.DESCRIPTION.index, CONFIG_COLUMNS.DESCRIPTION.width);
  sheet.setColumnWidth(CONFIG_COLUMNS.STATUS.index, CONFIG_COLUMNS.STATUS.width);

  // Add sample projects
  const sampleProjects = [
    [1, 'SIPGN', 'Hard', 'Sistem Informasi Pemerintah', 'Active'],
    [2, 'INAgov', 'Medium', 'Indonesia Government Portal', 'Active'],
    [3, 'Emeterai', 'Medium', 'Electronic Stamp System', 'Active'],
    [4, 'Peruri ID', 'Easy', 'Digital Identity System', 'Active'],
    [5, 'Wahana', 'Easy', 'Wahana Project', 'Active'],
    [6, 'Digidoc 2.0', 'Medium', 'Digital Document Management', 'Active'],
    [7, 'Peruri Shield', 'Hard', 'Security Shield System', 'Active'],
    [8, 'Penjaminan Online', 'Medium', 'Online Guarantee System', 'Active'],
    [9, 'COTS', 'Easy', 'Commercial Off-The-Shelf', 'Active'],
    [10, 'Integrasi Data Omnyx', 'Hard', 'Omnyx Data Integration', 'Active']
  ];

  sheet.getRange(CONFIG_DATA_START_ROW, 1, sampleProjects.length, CONFIG_TOTAL_COLUMNS)
    .setValues(sampleProjects);

  // Apply formatting
  for (let i = 0; i < sampleProjects.length; i++) {
    const rowNum = CONFIG_DATA_START_ROW + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, CONFIG_TOTAL_COLUMNS)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Add data validation for Difficulty
  const difficultyRange = sheet.getRange(CONFIG_DATA_START_ROW, CONFIG_COLUMNS.DIFFICULTY.index, 100);
  const difficultyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Easy', 'Medium', 'Hard'], true)
    .build();
  difficultyRange.setDataValidation(difficultyRule);

  // Add data validation for Status
  const statusRange = sheet.getRange(CONFIG_DATA_START_ROW, CONFIG_COLUMNS.STATUS.index, 100);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Completed'], true)
    .build();
  statusRange.setDataValidation(statusRule);

  // Add conditional formatting for Difficulty
  const easyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Easy')
    .setBackground('#d4edda')
    .setRanges([sheet.getRange(CONFIG_DATA_START_ROW, CONFIG_COLUMNS.DIFFICULTY.index, 100, 1)])
    .build();

  const mediumRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Medium')
    .setBackground('#fff3cd')
    .setRanges([sheet.getRange(CONFIG_DATA_START_ROW, CONFIG_COLUMNS.DIFFICULTY.index, 100, 1)])
    .build();

  const hardRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Hard')
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange(CONFIG_DATA_START_ROW, CONFIG_COLUMNS.DIFFICULTY.index, 100, 1)])
    .build();

  sheet.setConditionalFormatRules([easyRule, mediumRule, hardRule]);

  Logger.log('✅ Config tab created');
}

/**
 * Get all active projects
 */
function getActiveProjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) {
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_DATA_START_ROW) {
    return [];
  }

  const data = sheet.getRange(CONFIG_DATA_START_ROW, 1, lastRow - CONFIG_DATA_START_ROW + 1, CONFIG_TOTAL_COLUMNS).getValues();

  const projects = [];
  data.forEach(row => {
    const projectName = row[1] ? row[1].toString().trim() : '';
    const status = row[4] ? row[4].toString().trim() : '';

    if (projectName && status === 'Active') {
      projects.push({
        name: projectName,
        difficulty: row[2] ? row[2].toString().trim() : '',
        description: row[3] ? row[3].toString().trim() : ''
      });
    }
  });

  return projects;
}

/**
 * Get project difficulty
 */
function getProjectDifficulty(projectName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) {
    return 'Unknown';
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_DATA_START_ROW) {
    return 'Unknown';
  }

  const data = sheet.getRange(CONFIG_DATA_START_ROW, 1, lastRow - CONFIG_DATA_START_ROW + 1, CONFIG_TOTAL_COLUMNS).getValues();

  for (let i = 0; i < data.length; i++) {
    const name = data[i][1] ? data[i][1].toString().trim() : '';
    if (name === projectName) {
      return data[i][2] ? data[i][2].toString().trim() : 'Unknown';
    }
  }

  return 'Unknown';
}
