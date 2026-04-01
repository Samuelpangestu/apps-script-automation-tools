/**
 * InitVAPT.js - Initialize VAPT tabs and configuration
 *
 * Usage:
 * 1. Open Apps Script Editor
 * 2. Run: initializeVAPTSetup()
 * 3. This will create:
 *    - VAPT tab (main dashboard)
 *    - VAPT History tab (for trends)
 *    - Config tab (if not exists) with VAPT spreadsheet ID
 */

/**
 * Initialize VAPT Setup - Creates all necessary tabs
 * Run this function from Apps Script editor
 */
function initializeVAPTSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('🔧 Starting VAPT Setup...');

  // 1. Create or verify Config tab
  let configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    Logger.log('Creating Config tab...');
    configSheet = ss.insertSheet('Config', 0);
    setupConfigTab_(configSheet);
  } else {
    Logger.log('✅ Config tab already exists');
  }

  // 2. Create VAPT tab
  let vaptSheet = ss.getSheetByName('VAPT');
  if (vaptSheet) {
    const response = SpreadsheetApp.getUi().alert(
      'VAPT Tab Exists',
      'VAPT tab already exists. Do you want to rebuild it?\n\n' +
      'WARNING: This will delete all existing data in VAPT tab.',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (response === SpreadsheetApp.getUi().Button.YES) {
      Logger.log('Deleting existing VAPT tab...');
      ss.deleteSheet(vaptSheet);
      Logger.log('Creating new VAPT tab...');
      buildVAPT(ss);
    } else {
      Logger.log('Skipping VAPT tab creation');
    }
  } else {
    Logger.log('Creating VAPT tab...');
    buildVAPT(ss);
  }

  // 3. Create VAPT History tab
  let historySheet = ss.getSheetByName('VAPT History');
  if (historySheet) {
    const response = SpreadsheetApp.getUi().alert(
      'VAPT History Exists',
      'VAPT History tab already exists. Do you want to rebuild it?\n\n' +
      'WARNING: This will delete all historical data.',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (response === SpreadsheetApp.getUi().Button.YES) {
      Logger.log('Deleting existing VAPT History tab...');
      ss.deleteSheet(historySheet);
      Logger.log('Creating new VAPT History tab...');
      buildVAPTHistory(ss);
    } else {
      Logger.log('Skipping VAPT History tab creation');
    }
  } else {
    Logger.log('Creating VAPT History tab...');
    buildVAPTHistory(ss);
  }

  // 4. Show completion message
  SpreadsheetApp.getUi().alert(
    '✅ VAPT Setup Complete!',
    'VAPT tabs have been created successfully.\n\n' +
    'Next Steps:\n' +
    '1. Go to Config tab\n' +
    '2. Paste your VAPT Spreadsheet ID in cell B4\n' +
    '3. Run refreshDashboard() to fetch VAPT data\n\n' +
    'Tabs Created:\n' +
    '• VAPT (main dashboard)\n' +
    '• VAPT History (trends)\n' +
    '• Config (credentials)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log('✅ VAPT Setup Complete!');
}

/**
 * Setup Config tab with VAPT configuration section
 */
function setupConfigTab_(ws) {
  ws.setTabColor('#607D8B');
  ws.clear();

  // Title
  ws.getRange(1, 1, 1, 4).merge()
    .setValue('QA DASHBOARD CONFIGURATION')
    .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(14).setHorizontalAlignment('center');
  ws.setRowHeight(1, 35);

  // Headers
  ws.getRange(2, 1).setValue('Setting').setFontWeight('bold').setBackground('#B0BEC5');
  ws.getRange(2, 2).setValue('Value').setFontWeight('bold').setBackground('#B0BEC5');
  ws.getRange(2, 3).setValue('Description').setFontWeight('bold').setBackground('#B0BEC5');
  ws.getRange(2, 4).setValue('Example').setFontWeight('bold').setBackground('#B0BEC5');

  // Column widths
  ws.setColumnWidth(1, 200);
  ws.setColumnWidth(2, 350);
  ws.setColumnWidth(3, 300);
  ws.setColumnWidth(4, 250);

  // VAPT Configuration Section
  ws.getRange(3, 1, 1, 4).merge()
    .setValue('🔒 VAPT CONFIGURATION')
    .setBackground('#FF6F00').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center');

  // VAPT Spreadsheet ID
  const vaptRow = 4;
  ws.getRange(vaptRow, 1).setValue('VAPT Spreadsheet ID')
    .setBackground('#FFF3E0').setFontWeight('bold');
  ws.getRange(vaptRow, 2).setValue('PASTE_VAPT_SPREADSHEET_ID_HERE')
    .setBackground('#FFFFFF').setFontStyle('italic').setFontColor('#999999');
  ws.getRange(vaptRow, 3).setValue('Spreadsheet ID yang berisi VAPT findings (Ad Hoc + Regular VAPT)')
    .setBackground('#FFF3E0').setWrap(true);
  ws.getRange(vaptRow, 4).setValue('17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk')
    .setBackground('#FFF3E0').setFontStyle('italic');

  // Instructions
  ws.getRange(6, 1, 1, 4).merge()
    .setValue('📋 INSTRUCTIONS')
    .setBackground('#1976D2').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center');

  ws.getRange(7, 1, 1, 4).merge()
    .setValue(
      '1. Paste your VAPT Spreadsheet ID in cell B4\n' +
      '2. Go to Extensions > Apps Script\n' +
      '3. Run: refreshDashboard()\n' +
      '4. VAPT data will be fetched and displayed in VAPT tab'
    )
    .setBackground('#E3F2FD').setWrap(true).setVerticalAlignment('top');
  ws.setRowHeight(7, 80);

  // How to get Spreadsheet ID
  ws.getRange(9, 1, 1, 4).merge()
    .setValue('❓ How to get Spreadsheet ID?')
    .setBackground('#4CAF50').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center');

  ws.getRange(10, 1, 1, 4).merge()
    .setValue(
      'From spreadsheet URL:\n' +
      'https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit\n\n' +
      'Example:\n' +
      'URL: https://docs.google.com/spreadsheets/d/17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk/edit\n' +
      'ID:  17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk'
    )
    .setBackground('#E8F5E9').setWrap(true).setVerticalAlignment('top')
    .setFontFamily('Courier New').setFontSize(9);
  ws.setRowHeight(10, 100);

  Logger.log('✅ Config tab created');
}

/**
 * Quick function to rebuild VAPT tab only
 */
function rebuildVAPTTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = SpreadsheetApp.getUi().alert(
    'Rebuild VAPT Tab',
    'This will delete and recreate the VAPT tab.\n\n' +
    'WARNING: All data in VAPT tab will be lost.\n\n' +
    'Continue?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response === SpreadsheetApp.getUi().Button.YES) {
    const vaptSheet = ss.getSheetByName('VAPT');
    if (vaptSheet) {
      ss.deleteSheet(vaptSheet);
    }
    buildVAPT(ss);

    SpreadsheetApp.getUi().alert(
      '✅ VAPT Tab Rebuilt',
      'VAPT tab has been recreated.\n\n' +
      'Run refreshDashboard() to fetch data.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    Logger.log('✅ VAPT tab rebuilt');
  } else {
    Logger.log('Rebuild cancelled');
  }
}

/**
 * Quick function to rebuild VAPT History tab only
 */
function rebuildVAPTHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = SpreadsheetApp.getUi().alert(
    'Rebuild VAPT History',
    'This will delete and recreate the VAPT History tab.\n\n' +
    'WARNING: All historical data will be lost.\n\n' +
    'Continue?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response === SpreadsheetApp.getUi().Button.YES) {
    const historySheet = ss.getSheetByName('VAPT History');
    if (historySheet) {
      ss.deleteSheet(historySheet);
    }
    buildVAPTHistory(ss);

    SpreadsheetApp.getUi().alert(
      '✅ VAPT History Rebuilt',
      'VAPT History tab has been recreated.\n\n' +
      'Run refreshDashboard() to start logging data.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    Logger.log('✅ VAPT History tab rebuilt');
  } else {
    Logger.log('Rebuild cancelled');
  }
}
