/**
 * TeamConfig.js — KPI Tracker Configuration Setup
 * ═══════════════════════════════════════════════════════════════════════
 * Setup centralized configuration for team members and KPI settings
 * ═══════════════════════════════════════════════════════════════════════
 */

// Constants
const CONFIG_TAB_NAME = 'Config';
const TEAM_START_ROW = 6;

/**
 * Create or update Config tab with team member structure
 */
function setupConfigTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let config = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!config) {
    config = ss.insertSheet(CONFIG_TAB_NAME, 0);
  }

  // Clear existing content
  config.clear();

  // Set column widths
  config.setColumnWidth(1, 50);   // A: No
  config.setColumnWidth(2, 200);  // B: Nama
  config.setColumnWidth(3, 220);  // C: Role
  config.setColumnWidth(4, 200);  // D: Email
  config.setColumnWidth(5, 100);  // E: Status
  config.setColumnWidth(6, 150);  // F: Start Date
  config.setColumnWidth(7, 150);  // G: End Date

  // Header section
  config.getRange('A1:G1').merge()
    .setValue('🎯 KPI TRACKER — QA DEPARTMENT PERURI')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  config.setRowHeight(1, 40);

  config.getRange('A2:G2').merge()
    .setValue('Configuration & Team Members Management')
    .setFontSize(10)
    .setFontStyle('italic')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');

  // Team Members Section
  config.getRange('A4:G4').merge()
    .setValue('📋 TEAM MEMBERS')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f1f3f4')
    .setHorizontalAlignment('left');

  // Table headers
  const headers = [
    ['No', 'Nama', 'Role', 'Email', 'Status', 'Start Date', 'End Date']
  ];

  config.getRange(5, 1, 1, 7).setValues(headers)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Add sample data
  const sampleData = [
    [1, 'John Doe', 'QA Team Lead (CoE)', 'john@example.com', 'Aktif', new Date(), ''],
    [2, 'Jane Smith', 'QA Lead (Project Dedicated)', 'jane@example.com', 'Aktif', new Date(), ''],
    [3, 'Bob Wilson', 'PIC Project (QE + Koordinator)', 'bob@example.com', 'Aktif', new Date(), ''],
    [4, 'Alice Brown', 'Quality Engineer (QE)', 'alice@example.com', 'Aktif', new Date(), '']
  ];

  config.getRange(TEAM_START_ROW, 1, sampleData.length, 7).setValues(sampleData);

  // Format data rows
  for (let i = 0; i < 50; i++) {
    const row = TEAM_START_ROW + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
    config.getRange(row, 1, 1, 7)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    // Number column
    config.getRange(row, 1).setHorizontalAlignment('center');

    // Status dropdown
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Aktif', 'Non-Aktif', 'Cuti'], true)
      .setAllowInvalid(false)
      .build();
    config.getRange(row, 5).setDataValidation(statusRule);

    // Role dropdown
    const roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'QA Team Lead (CoE)',
        'QA Lead (Project Dedicated)',
        'PIC Project (QE + Koordinator)',
        'Quality Engineer (QE)'
      ], true)
      .setAllowInvalid(false)
      .build();
    config.getRange(row, 3).setDataValidation(roleRule);
  }

  // Date format for Start Date and End Date columns
  config.getRange(TEAM_START_ROW, 6, 50, 2).setNumberFormat('yyyy-mm-dd');

  // 360 Review Form Settings Section
  const formRow = 60;
  config.getRange(formRow, 1, 1, 7).merge()
    .setValue('📝 360 REVIEW FORM SETTINGS')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f1f3f4')
    .setHorizontalAlignment('left');

  config.getRange(formRow + 1, 1).setValue('Form ID:');
  config.getRange(formRow + 1, 2, 1, 6).merge()
    .setValue('(otomatis terisi saat form dibuat)')
    .setFontStyle('italic')
    .setFontColor('#666666');

  config.getRange(formRow + 2, 1).setValue('Form URL:');
  config.getRange(formRow + 2, 2, 1, 6).merge();

  config.getRange(formRow + 3, 1).setValue('Form Edit URL:');
  config.getRange(formRow + 3, 2, 1, 6).merge();

  // Freeze header rows
  config.setFrozenRows(5);
  config.setFrozenColumns(1);

  Logger.log('✅ Config tab setup complete');
  return config;
}

/**
 * Get active team members from Config tab
 * @param {string} roleFilter - Optional role filter
 * @returns {Array} Array of team member objects
 */
function getTeamMembers(roleFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!config) {
    throw new Error('Config tab not found. Run setupConfigTab() first.');
  }

  const members = [];
  const lastRow = Math.min(config.getLastRow(), TEAM_START_ROW + 50);

  for (let row = TEAM_START_ROW; row <= lastRow; row++) {
    const name = config.getRange(row, 2).getValue().toString().trim();
    const role = config.getRange(row, 3).getValue().toString().trim();
    const email = config.getRange(row, 4).getValue().toString().trim();
    const status = config.getRange(row, 5).getValue().toString().trim();

    if (name && status === 'Aktif') {
      if (!roleFilter || role === roleFilter) {
        members.push({
          row: row,
          name: name,
          role: role,
          email: email,
          status: status
        });
      }
    }
  }

  return members;
}

/**
 * Get 360 Review Form settings from Config
 * @returns {Object} Form settings
 */
function getFormSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_TAB_NAME);

  return {
    formId: config.getRange(61, 2).getValue().toString().trim(),
    formUrl: config.getRange(62, 2).getValue().toString().trim(),
    formEditUrl: config.getRange(63, 2).getValue().toString().trim()
  };
}

/**
 * Save 360 Review Form settings to Config
 * @param {string} formId - Google Form ID
 * @param {string} formUrl - Public form URL
 * @param {string} formEditUrl - Edit form URL
 */
function saveFormSettings(formId, formUrl, formEditUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_TAB_NAME);

  config.getRange(61, 2).setValue(formId);
  config.getRange(62, 2).setValue(formUrl);
  config.getRange(63, 2).setValue(formEditUrl);

  Logger.log('✅ Form settings saved to Config');
}
