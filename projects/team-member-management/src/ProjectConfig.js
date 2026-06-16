/**
 * ProjectConfig.js — Centralized Configuration Management
 * ═══════════════════════════════════════════════════════════════════════
 * Flat table format (Portfolio style) for team management
 * ═══════════════════════════════════════════════════════════════════════
 */

const CONFIG_TAB_NAME = 'Project';

// Flat table structure (Portfolio style)
const CONFIG_HEADER_ROW = 3;  // Row 3 (header)
const CONFIG_DATA_START_ROW = 4;  // Row 4 (data starts after header)
const CONFIG_COLUMNS = {
  PROJECT: { index: 1, letter: 'A', width: 200, header: 'Project' },
  MODUL: { index: 2, letter: 'B', width: 180, header: 'Modul' },
  SUBMODUL: { index: 3, letter: 'C', width: 220, header: 'Submodul' },
  PIC_QA: { index: 4, letter: 'D', width: 150, header: 'PIC QA' },
  DIFFICULTY: { index: 5, letter: 'E', width: 80, header: 'Diff (1-10)' },
  RISK: { index: 6, letter: 'F', width: 80, header: 'Risk (1-10)' },
  COMPLEXITY: { index: 7, letter: 'G', width: 80, header: 'Comp (1-10)' },
  AUTOMATION: { index: 8, letter: 'H', width: 80, header: 'Auto (1-10)' },
  TEST_COMPLEXITY: { index: 9, letter: 'I', width: 80, header: 'Test (1-10)' },
  STATUS: { index: 10, letter: 'J', width: 100, header: 'Status' },
  QA_LEAD: { index: 11, letter: 'K', width: 150, header: 'Submodul Lead' }
};
const CONFIG_TOTAL_COLUMNS = 11;

// Column indices for array access (0-based)
const CONFIG_COL = {
  PROJECT: 0,        // Column A
  MODUL: 1,          // Column B
  SUBMODUL: 2,       // Column C
  PIC_QA: 3,         // Column D
  DIFFICULTY: 4,     // Column E
  RISK: 5,           // Column F
  COMPLEXITY: 6,     // Column G
  AUTOMATION: 7,     // Column H
  TEST_COMPLEXITY: 8,// Column I
  STATUS: 9,         // Column J
  QA_LEAD: 10        // Column K (Submodul Lead)
};

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
    .setValue('📋 PROJECT CONFIGURATION — Define projects, moduls, submoduls with difficulty ratings (Status: Active/Inactive/On Hold)')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 40);
  currentRow++;

  // QA Dashboard Sync Config (Row 2, Column P)
  sheet.getRange(2, 16).setValue('QA Dashboard ID:')
    .setFontWeight('bold')
    .setFontSize(9)
    .setBackground('#f3f3f3');
  sheet.getRange(2, 17, 1, 3).merge()
    .setValue('1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY')
    .setBackground('#e8f5e9')
    .setFontWeight('bold')
    .setFontSize(9)
    .setNote('QA Dashboard Spreadsheet ID (Production)\n\nYou can also paste full URL:\nhttps://docs.google.com/spreadsheets/d/1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY/edit');
  sheet.setColumnWidth(16, 120); // Column P
  sheet.setColumnWidth(17, 300); // Column Q (merged with R, S)

  currentRow++;

  // Main table header
  const headers = [
    CONFIG_COLUMNS.PROJECT.header,
    CONFIG_COLUMNS.MODUL.header,
    CONFIG_COLUMNS.SUBMODUL.header,
    CONFIG_COLUMNS.PIC_QA.header,
    CONFIG_COLUMNS.DIFFICULTY.header,
    CONFIG_COLUMNS.RISK.header,
    CONFIG_COLUMNS.COMPLEXITY.header,
    CONFIG_COLUMNS.AUTOMATION.header,
    CONFIG_COLUMNS.TEST_COMPLEXITY.header,
    CONFIG_COLUMNS.STATUS.header,
    CONFIG_COLUMNS.QA_LEAD.header
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

  // Sample data (11 columns: Project, Modul, Submodul, PIC QA, 5 Parameters, Status, QA Lead)
  const sampleData = [
    ['SIPGN', 'E2E DAPUR', 'E2E DAPUR', 'Adinda', 7, 6, 7, 8, 7, 'Active', 'Dini'],
    ['SIPGN', 'AtomAPI', 'AtomAPI PM', 'Denta', 6, 5, 6, 5, 6, 'Active', 'Dini'],
    ['SIPGN', '1', '1.1', 'Lutfi', 8, 7, 8, 9, 8, 'Active', 'Samuel'],
    ['SIPGN', '1', '1.2', 'Samuel', 7, 6, 7, 7, 7, 'Active', 'Samuel'],
    ['SIPGN', '1', '1.3', 'Irvan', 6, 5, 6, 5, 6, 'Active', 'Samuel'],
    ['SIPGN', '2', '2.1', 'Rizky', 7, 6, 7, 6, 7, 'Active', 'Samuel'],
    ['SIPGN', '2', '2.2', 'Adinda', 6, 5, 6, 4, 5, 'Active', 'Samuel'],
    ['SIPGN', '4', '4.1', 'Lutfi', 7, 7, 7, 7, 7, 'Active', 'Lutfi'],
    ['SIPGN', '4', '4.2', 'Irvan', 6, 5, 6, 5, 6, 'Active', 'Lutfi'],
    ['INADigital', 'INAgov', 'INAgov', 'Irvan', 8, 7, 8, 9, 8, 'Active', 'Lutfi'],
    ['INADigital', 'POS', 'POS', 'Rizky', 5, 4, 5, 3, 4, 'Active', 'Lutfi'],
    ['INADigital', 'SCM', 'SCM', 'Samuel', 6, 5, 6, 5, 6, 'Active', 'Lutfi'],
    ['INADigital', 'PERURIID', 'PERURIID', 'Adinda', 7, 6, 7, 6, 7, 'Active', 'Lutfi'],
    ['COTS', 'COTS', 'CODEBASE', '', 3, 2, 3, 2, 2, 'Inactive', ''],
    ['COTS', 'COTS', 'GEODIPA', '', 3, 2, 3, 2, 2, 'Inactive', '']
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

  // Data validation for Status column
  const statusRange = sheet.getRange(CONFIG_DATA_START_ROW, CONFIG_COLUMNS.STATUS.index, 100);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'On Hold'], true)
    .build();
  statusRange.setDataValidation(statusRule);

  // Create helper columns for dropdowns (columns L, M, N, O) - moved from K-M to avoid conflict with QA Lead
  currentRow = 1;

  // Helper header
  sheet.getRange(currentRow, 12).setValue('Available Projects').setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange(currentRow, 13).setValue('Available Moduls').setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange(currentRow, 14).setValue('Available Submoduls').setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange(currentRow, 15).setValue('Available QA Leads').setFontWeight('bold').setBackground('#f3f3f3');

  // Get unique values from ALL rows (not just Active) to preserve historical data
  const uniqueProjects = [...new Set(sampleData.map(r => r[0]))].filter(Boolean);
  const uniqueModuls = [...new Set(sampleData.map(r => r[1]))].filter(Boolean);
  const uniqueSubmoduls = [...new Set(sampleData.map(r => r[2]))].filter(Boolean);
  const uniqueQALeads = [...new Set(sampleData.map(r => r[10]))].filter(Boolean);

  // Write helper lists
  currentRow = 2;
  uniqueProjects.forEach(project => {
    sheet.getRange(currentRow, 12).setValue(project);
    currentRow++;
  });

  currentRow = 2;
  uniqueModuls.forEach(modul => {
    sheet.getRange(currentRow, 13).setValue(modul);
    currentRow++;
  });

  currentRow = 2;
  uniqueSubmoduls.forEach(submodul => {
    sheet.getRange(currentRow, 14).setValue(submodul);
    currentRow++;
  });

  currentRow = 2;
  uniqueQALeads.forEach(qaLead => {
    sheet.getRange(currentRow, 15).setValue(qaLead);
    currentRow++;
  });

  // Set helper column widths
  sheet.setColumnWidth(12, 180); // Projects helper
  sheet.setColumnWidth(13, 180); // Moduls helper
  sheet.setColumnWidth(14, 220); // Submoduls helper
  sheet.setColumnWidth(15, 150); // QA Leads helper

  // Add header notes as guidance
  addProjectHeaderNotes(sheet);

  // Flush changes to ensure completion
  SpreadsheetApp.flush();

  Logger.log('✅ Config tab created (flat table format)');
}

/**
 * Add notes to Project tab headers for guidance
 */
function addProjectHeaderNotes(sheet) {
  // Difficulty note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.DIFFICULTY.index).setNote(
    '📊 DIFFICULTY (Tingkat Kesulitan)\n' +
    'Seberapa sulit modul ini untuk di-test?\n\n' +
    '1-3: Modul sederhana, UI straightforward\n' +
    '4-6: Modul dengan business logic moderate\n' +
    '7-10: Modul kompleks, banyak edge cases'
  );

  // Risk note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.RISK.index).setNote(
    '⚠️ RISK (Tingkat Risiko)\n' +
    'Seberapa besar dampak jika ada bug?\n\n' +
    '1-3: Low impact, non-critical features\n' +
    '4-6: Medium impact, affects user experience\n' +
    '7-10: High impact, critical business function'
  );

  // Complexity note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.COMPLEXITY.index).setNote(
    '🔧 COMPLEXITY (Kompleksitas Teknis)\n' +
    'Seberapa kompleks arsitektur/integrasi modul?\n\n' +
    '1-3: Simple, standalone module\n' +
    '4-6: Multiple integrations, moderate logic\n' +
    '7-10: High complexity, many dependencies'
  );

  // Automation note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.AUTOMATION.index).setNote(
    '🤖 AUTOMATION (Tingkat Automasi)\n' +
    'Seberapa banyak test yang perlu automation?\n\n' +
    '1-3: Manual testing cukup\n' +
    '4-6: Mix manual + automation\n' +
    '7-10: Heavy automation needed (API, E2E)'
  );

  // Test Complexity note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.TEST_COMPLEXITY.index).setNote(
    '🧪 TEST COMPLEXITY (Kompleksitas Testing)\n' +
    'Seberapa kompleks test case yang diperlukan?\n\n' +
    '1-3: Basic functional testing\n' +
    '4-6: Moderate test scenarios\n' +
    '7-10: Complex scenarios, many test cases'
  );

  // Status note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.STATUS.index).setNote(
    '📌 STATUS\n' +
    'Status project/modul/submodul:\n\n' +
    '• Active: Aktif, akan muncul di dropdown Team Members\n' +
    '• Inactive: Tidak aktif, tidak muncul di dropdown\n' +
    '• On Hold: Ditunda sementara'
  );

  // Project note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.PROJECT.index).setNote(
    '📁 PROJECT\n' +
    'Nama project utama.\n' +
    'Contoh: SIPGN, INADigital, COTS'
  );

  // Modul note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.MODUL.index).setNote(
    '📂 MODUL\n' +
    'Nama modul di dalam project.\n' +
    'Contoh: E2E DAPUR, INAgov, 1, 2, 4'
  );

  // Submodul note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.SUBMODUL.index).setNote(
    '📄 SUBMODUL\n' +
    'Nama submodul yang akan di-test.\n' +
    'Contoh: E2E DAPUR, INAgov, 1.1, 1.2, 2.1'
  );

  // PIC QA note
  sheet.getRange(CONFIG_HEADER_ROW, CONFIG_COLUMNS.PIC_QA.index).setNote(
    '👤 PIC QA\n' +
    'Person in charge untuk testing modul ini.\n\n' +
    'Bisa satu atau lebih nama (pisahkan dengan koma).\n' +
    'Contoh: Adinda, Denta, Irvan'
  );

  Logger.log('✅ Header notes added');
}

/**
 * Update helper columns in Project tab
 * Call this after adding/editing project data
 */
function updateProjectHelperColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!sheet) {
    Logger.log('⚠️ Project tab not found');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG_DATA_START_ROW) return;

  const data = sheet.getRange(CONFIG_DATA_START_ROW, 1, lastRow - CONFIG_DATA_START_ROW + 1, CONFIG_TOTAL_COLUMNS).getValues();

  // Get unique values from ALL rows (not just Active) to preserve historical data
  const uniqueProjects = [...new Set(data.filter(r => r[0]).map(r => r[0].toString().trim()))];
  const uniqueModuls = [...new Set(data.filter(r => r[1]).map(r => r[1].toString().trim()))];
  const uniqueSubmoduls = [...new Set(data.filter(r => r[2]).map(r => r[2].toString().trim()))];
  const uniqueQALeads = [...new Set(data.filter(r => r[10]).map(r => r[10].toString().trim()))];

  // Clear existing helper data (keep header) - moved to columns L-O (12-15) to avoid conflict with QA Lead (K/11)
  const helperLastRow = sheet.getLastRow();
  if (helperLastRow > 1) {
    sheet.getRange(2, 12, helperLastRow - 1, 4).clearContent();
  }

  // Write new helper lists
  let currentRow = 2;
  uniqueProjects.forEach(project => {
    sheet.getRange(currentRow, 12).setValue(project);
    currentRow++;
  });

  currentRow = 2;
  uniqueModuls.forEach(modul => {
    sheet.getRange(currentRow, 13).setValue(modul);
    currentRow++;
  });

  currentRow = 2;
  uniqueSubmoduls.forEach(submodul => {
    sheet.getRange(currentRow, 14).setValue(submodul);
    currentRow++;
  });

  currentRow = 2;
  uniqueQALeads.forEach(qaLead => {
    sheet.getRange(currentRow, 15).setValue(qaLead);
    currentRow++;
  });

  Logger.log('✅ Helper columns updated');
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
    const projectName = row[CONFIG_COL.PROJECT] ? row[CONFIG_COL.PROJECT].toString().trim() : '';
    const status = row[CONFIG_COL.STATUS] ? row[CONFIG_COL.STATUS].toString().trim() : '';

    if (projectName && status === 'Active') {
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
    const projectName = row[CONFIG_COL.PROJECT] ? row[CONFIG_COL.PROJECT].toString().trim() : '';
    const modulName = row[CONFIG_COL.MODUL] ? row[CONFIG_COL.MODUL].toString().trim() : '';
    const status = row[CONFIG_COL.STATUS] ? row[CONFIG_COL.STATUS].toString().trim() : '';

    if (projectName && modulName && status === 'Active') {
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
    const projectName = row[CONFIG_COL.PROJECT] ? row[CONFIG_COL.PROJECT].toString().trim() : '';
    const modulName = row[CONFIG_COL.MODUL] ? row[CONFIG_COL.MODUL].toString().trim() : '';
    const submodulName = row[CONFIG_COL.SUBMODUL] ? row[CONFIG_COL.SUBMODUL].toString().trim() : '';
    const picQA = row[CONFIG_COL.PIC_QA] ? row[CONFIG_COL.PIC_QA].toString().trim() : '';
    const difficulty = row[CONFIG_COL.DIFFICULTY] || 0;
    const risk = row[CONFIG_COL.RISK] || 0;
    const complexity = row[CONFIG_COL.COMPLEXITY] || 0;
    const automation = row[CONFIG_COL.AUTOMATION] || 0;
    const testComplexity = row[CONFIG_COL.TEST_COMPLEXITY] || 0;
    const status = row[CONFIG_COL.STATUS] ? row[CONFIG_COL.STATUS].toString().trim() : '';
    const qaLead = row[CONFIG_COL.QA_LEAD] ? row[CONFIG_COL.QA_LEAD].toString().trim() : '';

    if (projectName && modulName && submodulName && status === 'Active') {
      submoduls.push({
        name: submodulName,
        modul: modulName,
        project: projectName,
        picQA: picQA,
        qaLead: qaLead,
        difficulty: difficulty,
        risk: risk,
        complexity: complexity,
        automation: automation,
        testComplexity: testComplexity,
        description: ''
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
    complexity: found.complexity,
    automation: found.automation,
    testComplexity: found.testComplexity
  } : { difficulty: 0, risk: 0, complexity: 0, automation: 0, testComplexity: 0 };
}

/**
 * Calculate total score from 5 parameters
 * @param {number} difficulty - 1-10
 * @param {number} risk - 1-10
 * @param {number} complexity - 1-10
 * @param {number} automation - 1-10
 * @param {number} testComplexity - 1-10
 * @returns {number} Total score (5-50)
 */
function calculateTotalScore(difficulty, risk, complexity, automation, testComplexity) {
  return (difficulty || 0) + (risk || 0) + (complexity || 0) + (automation || 0) + (testComplexity || 0);
}

/**
 * Get difficulty category based on total score
 * @param {number} totalScore - Total score from 5 parameters (5-50)
 * @returns {string} Category: 'Very Easy', 'Easy', 'Medium', 'Hard', 'Very Hard'
 */
function getDifficultyCategory(totalScore) {
  if (totalScore <= 14) return 'Very Easy';
  if (totalScore <= 24) return 'Easy';
  if (totalScore <= 34) return 'Medium';
  if (totalScore <= 44) return 'Hard';
  return 'Very Hard';
}

/**
 * Get difficulty category color
 * @param {string} category - Category name
 * @returns {string} Color code
 */
function getCategoryColor(category) {
  switch (category) {
    case 'Very Easy': return '#d4edda'; // Light green
    case 'Easy': return '#cce5ff'; // Light blue
    case 'Medium': return '#fff3cd'; // Light yellow
    case 'Hard': return '#ffe0b2'; // Light orange
    case 'Very Hard': return '#f8d7da'; // Light red
    default: return '#ffffff';
  }
}

/**
 * Get difficulty category emoji
 * @param {string} category - Category name
 * @returns {string} Emoji icon
 */
function getCategoryEmoji(category) {
  switch (category) {
    case 'Very Easy': return '🟢';
    case 'Easy': return '🔵';
    case 'Medium': return '🟡';
    case 'Hard': return '🟠';
    case 'Very Hard': return '🔴';
    default: return '⚪';
  }
}
