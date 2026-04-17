/**
 * ProjectConfig.js — Centralized Configuration Management
 * ═══════════════════════════════════════════════════════════════════════
 * Manage modul, submodul, and difficulty levels - all in one place
 * ═══════════════════════════════════════════════════════════════════════
 */

const CONFIG_TAB_NAME = 'Config';

// Modul section
const MODUL_HEADER_ROW = 9;
const MODUL_DATA_START_ROW = 10;
const MODUL_COLUMNS = {
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  MODUL_NAME: { index: 2, letter: 'B', width: 200, header: 'Modul' },
  STATUS: { index: 3, letter: 'C', width: 100, header: 'Status' }
};
const MODUL_TOTAL_COLUMNS = 3;

// Submodul section (starts after modul data + gap)
const SUBMODUL_HEADER_ROW = 14; // Will be calculated dynamically
const SUBMODUL_DATA_START_ROW = 15;
const SUBMODUL_COLUMNS = {
  NO: { index: 1, letter: 'A', width: 50, header: 'No' },
  SUBMODUL_NAME: { index: 2, letter: 'B', width: 250, header: 'Submodul' },
  MODUL: { index: 3, letter: 'C', width: 150, header: 'Modul' },
  DIFFICULTY: { index: 4, letter: 'D', width: 120, header: 'Difficulty Level' },
  DESCRIPTION: { index: 5, letter: 'E', width: 300, header: 'Description' },
  STATUS: { index: 6, letter: 'F', width: 100, header: 'Status' }
};
const SUBMODUL_TOTAL_COLUMNS = 6;

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

  // ═════════════════════════════════════════════════════════════
  // MODUL LIST SECTION
  // ═════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('📁 MODUL LIST')
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

  // Modul data (only 2 modul)
  const modulData = [
    [1, 'INADigital', 'Active'],
    [2, 'SIPGN', 'Active']
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
  currentRow++; // Empty row

  // ═════════════════════════════════════════════════════════════
  // SUBMODUL LIST SECTION
  // ═════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('📋 SUBMODUL LIST')
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

  // Sample submodul data
  const submodulData = [
    [1, 'INAgov', 'INADigital', 'Medium', 'Indonesia Government Portal', 'Active'],
    [2, 'Emeterai', 'INADigital', 'Medium', 'Electronic Stamp System', 'Active'],
    [3, 'Peruri ID', 'INADigital', 'Easy', 'Digital Identity System', 'Active'],
    [4, 'Wahana', 'INADigital', 'Easy', 'Wahana Project', 'Active'],
    [5, 'Digidoc 2.0', 'INADigital', 'Medium', 'Digital Document Management', 'Active'],
    [6, 'Peruri Shield', 'INADigital', 'Hard', 'Security Shield System', 'Active'],
    [7, 'Penjaminan Online', 'INADigital', 'Medium', 'Online Guarantee System', 'Active'],
    [8, 'COTS', 'INADigital', 'Easy', 'Commercial Off-The-Shelf', 'Active'],
    [9, 'Core System', 'SIPGN', 'Hard', 'SIPGN Core System', 'Active'],
    [10, 'Integration Module', 'SIPGN', 'Hard', 'SIPGN Integration', 'Active'],
    [11, 'Data Omnyx', 'SIPGN', 'Hard', 'Omnyx Data Integration', 'Active']
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

  // Set column widths
  sheet.setColumnWidth(1, SUBMODUL_COLUMNS.NO.width);
  sheet.setColumnWidth(2, SUBMODUL_COLUMNS.SUBMODUL_NAME.width);
  sheet.setColumnWidth(3, SUBMODUL_COLUMNS.MODUL.width);
  sheet.setColumnWidth(4, SUBMODUL_COLUMNS.DIFFICULTY.width);
  sheet.setColumnWidth(5, SUBMODUL_COLUMNS.DESCRIPTION.width);
  sheet.setColumnWidth(6, SUBMODUL_COLUMNS.STATUS.width);

  // Flush changes to ensure completion
  SpreadsheetApp.flush();

  Logger.log('✅ Config tab created');
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

  // Read modul data (rows 10-11 typically)
  const data = sheet.getRange(MODUL_DATA_START_ROW, 1, 5, MODUL_TOTAL_COLUMNS).getValues();

  const moduls = [];
  data.forEach(row => {
    const modulName = row[1] ? row[1].toString().trim() : '';
    const status = row[2] ? row[2].toString().trim() : '';

    if (modulName && status === 'Active') {
      moduls.push({
        name: modulName
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

  // Read submodul data (typically starting from row 15)
  const data = sheet.getRange(SUBMODUL_DATA_START_ROW, 1, lastRow - SUBMODUL_DATA_START_ROW + 1, SUBMODUL_TOTAL_COLUMNS).getValues();

  const submoduls = [];
  data.forEach(row => {
    const submodulName = row[1] ? row[1].toString().trim() : '';
    const modulName = row[2] ? row[2].toString().trim() : '';
    const status = row[5] ? row[5].toString().trim() : '';

    if (submodulName && status === 'Active') {
      submoduls.push({
        name: submodulName,
        modul: modulName,
        difficulty: row[3] ? row[3].toString().trim() : '',
        description: row[4] ? row[4].toString().trim() : ''
      });
    }
  });

  return submoduls;
}

/**
 * Get submodul difficulty
 */
function getSubmodulDifficulty(submodulName) {
  const submoduls = getActiveSubmodul();
  const found = submoduls.find(s => s.name === submodulName);
  return found ? found.difficulty : 'Unknown';
}
