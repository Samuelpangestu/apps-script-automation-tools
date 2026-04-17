/**
 * ProjectConfig.js — Project Configuration Management
 * ═══════════════════════════════════════════════════════════════════════
 * Manage project list and difficulty levels
 * ═══════════════════════════════════════════════════════════════════════
 */

const CONFIG_TAB_NAME = 'Config - Projects';
const CONFIG_HEADER_ROW = 9; // Updated due to definitions section
const CONFIG_DATA_START_ROW = 10;

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

  // Add title and difficulty definitions first
  let currentRow = 1;

  // Title
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue('⚙️ PROJECT CONFIGURATION')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 45);
  currentRow += 2;

  // Difficulty Definitions
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue('📋 DIFFICULTY LEVEL DEFINITIONS')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 30);
  currentRow++;

  const definitions = [
    ['Level', 'Description', 'Characteristics', 'Team Size', 'Example'],
    [
      '🟢 Easy',
      'Straightforward projects with minimal complexity',
      '• Well-documented requirements\n• Simple workflows\n• Minimal integrations\n• Low risk',
      '1-2 QE',
      'COTS, Wahana, Peruri ID'
    ],
    [
      '🟡 Medium',
      'Moderate complexity with some challenges',
      '• Multiple modules\n• Some integrations\n• Moderate risk\n• Regular updates needed',
      '1 PIC + 2-3 QE',
      'INAgov, Emeterai, Digidoc 2.0, Penjaminan Online'
    ],
    [
      '🔴 Hard',
      'Complex projects requiring significant expertise',
      '• High complexity\n• Multiple integrations\n• Critical systems\n• High risk\n• Requires senior oversight',
      '1 Team Lead + 1 PIC + 3-5 QE',
      'SIPGN, Peruri Shield, Integrasi Data Omnyx'
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
    if (definitions[i + 1][0].includes('Easy')) bg = '#d4edda';
    else if (definitions[i + 1][0].includes('Medium')) bg = '#fff3cd';
    else if (definitions[i + 1][0].includes('Hard')) bg = '#f8d7da';

    sheet.getRange(rowNum, 1, 1, 5)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      .setWrap(true)
      .setVerticalAlignment('top');

    sheet.setRowHeight(rowNum, 80);
  }

  // Set column widths for definition section
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 230);

  currentRow += definitions.length;
  currentRow++; // Empty row

  // Project List Section
  const projectListRow = currentRow;
  sheet.getRange(currentRow, 1, 1, 5).merge()
    .setValue('📁 PROJECT LIST')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Create header
  const headers = [
    CONFIG_COLUMNS.NO.header,
    CONFIG_COLUMNS.PROJECT_NAME.header,
    CONFIG_COLUMNS.DIFFICULTY.header,
    CONFIG_COLUMNS.DESCRIPTION.header,
    CONFIG_COLUMNS.STATUS.header
  ];

  sheet.getRange(currentRow, 1, 1, CONFIG_TOTAL_COLUMNS)
    .setValues([headers])
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  currentRow++;

  // Set column widths for project list
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

  const projectStartRow = currentRow;
  sheet.getRange(projectStartRow, 1, sampleProjects.length, CONFIG_TOTAL_COLUMNS)
    .setValues(sampleProjects);

  // Apply formatting
  for (let i = 0; i < sampleProjects.length; i++) {
    const rowNum = projectStartRow + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(rowNum, 1, 1, CONFIG_TOTAL_COLUMNS)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Add data validation for Difficulty (reasonable range)
  const difficultyRange = sheet.getRange(projectStartRow, CONFIG_COLUMNS.DIFFICULTY.index, 50);
  const difficultyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Easy', 'Medium', 'Hard'], true)
    .build();
  difficultyRange.setDataValidation(difficultyRule);

  // Add data validation for Status
  const statusRange = sheet.getRange(projectStartRow, CONFIG_COLUMNS.STATUS.index, 50);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Completed'], true)
    .build();
  statusRange.setDataValidation(statusRule);

  // Add conditional formatting for Difficulty
  const easyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Easy')
    .setBackground('#d4edda')
    .setRanges([sheet.getRange(projectStartRow, CONFIG_COLUMNS.DIFFICULTY.index, 50, 1)])
    .build();

  const mediumRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Medium')
    .setBackground('#fff3cd')
    .setRanges([sheet.getRange(projectStartRow, CONFIG_COLUMNS.DIFFICULTY.index, 50, 1)])
    .build();

  const hardRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Hard')
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange(projectStartRow, CONFIG_COLUMNS.DIFFICULTY.index, 50, 1)])
    .build();

  sheet.setConditionalFormatRules([easyRule, mediumRule, hardRule]);

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
