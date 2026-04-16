/**
 * TeamMemberFormat.js — Team Member Tab Formatting
 * ═══════════════════════════════════════════════════════════════════════
 * Professional formatting for Team Member SSOT tab
 * Makes it clean, manageable, and easy to read
 * ═══════════════════════════════════════════════════════════════════════
 */

const TAB_NAME = 'Team Member';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Column definitions with optimal widths
const COLUMNS = {
  NAME: { letter: 'A', width: 200, header: 'Name' },
  JOIN: { letter: 'B', width: 100, header: 'Join Date' },
  TITLE: { letter: 'C', width: 150, header: 'Title' },
  LEAD_PIC: { letter: 'D', width: 120, header: 'Lead/PIC' },
  PROJECT: { letter: 'E', width: 250, header: 'Project' },
  NP: { letter: 'F', width: 80, header: 'NP' },
  EMAIL: { letter: 'G', width: 250, header: 'Email' },
  EMAIL_2: { letter: 'H', width: 250, header: 'Email 2' },
  STATUS: { letter: 'I', width: 120, header: 'Status' },
  AUTOMATION: { letter: 'J', width: 100, header: 'Automation' },
  GITHUB: { letter: 'K', width: 180, header: 'Github' },
  HP: { letter: 'L', width: 130, header: 'Phone' },
  ROLE: { letter: 'M', width: 200, header: 'Role' },
  VPN_ABC: { letter: 'N', width: 90, header: 'VPN ABC' },
  VPN_HUWAWEI: { letter: 'O', width: 120, header: 'VPN Huwawei' }
};

/**
 * Main function: Format Team Member tab
 */
function formatTeamMemberTab() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Format Team Member Tab',
    'This will apply professional formatting to the Team Member tab.\n\n' +
    'Changes:\n' +
    '• Clean header styling\n' +
    '• Optimal column widths\n' +
    '• Data validation\n' +
    '• Conditional formatting\n' +
    '• Freeze panes\n' +
    '• Filters\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(TAB_NAME);

    if (!sheet) {
      throw new Error('Tab "' + TAB_NAME + '" not found');
    }

    ui.alert('Formatting in progress...');

    // Step 1: Set column widths
    setColumnWidths(sheet);

    // Step 2: Format header
    formatHeader(sheet);

    // Step 3: Format data rows
    formatDataRows(sheet);

    // Step 4: Add data validation
    addDataValidation(sheet);

    // Step 5: Add conditional formatting
    addConditionalFormatting(sheet);

    // Step 6: Freeze panes
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);

    // Step 7: Add filter
    const lastRow = sheet.getLastRow();
    const lastCol = 15; // Column O
    if (lastRow >= HEADER_ROW) {
      sheet.getRange(HEADER_ROW, 1, lastRow, lastCol).createFilter();
    }

    // Step 8: Clean up empty rows at the end
    cleanupEmptyRows(sheet);

    Logger.log('✅ Formatting complete');

    ui.alert(
      'Formatting Complete! ✅',
      'Team Member tab has been formatted successfully.\n\n' +
      'Changes applied:\n' +
      '✓ Professional header styling\n' +
      '✓ Optimal column widths\n' +
      '✓ Data validation for dropdowns\n' +
      '✓ Color-coded status (Green/Yellow/Red)\n' +
      '✓ Freeze header row\n' +
      '✓ Filter enabled\n\n' +
      'Tab is now clean and easy to read!',
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log('❌ Formatting failed: ' + e.message);
    ui.alert('Formatting Failed', 'Error: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Set optimal column widths
 */
function setColumnWidths(sheet) {
  Object.keys(COLUMNS).forEach(key => {
    const col = COLUMNS[key];
    const colIndex = col.letter.charCodeAt(0) - 64; // A=1, B=2, etc.
    sheet.setColumnWidth(colIndex, col.width);
  });

  Logger.log('✅ Column widths set');
}

/**
 * Format header row
 */
function formatHeader(sheet) {
  // Header range (row 1, all columns)
  const headerRange = sheet.getRange(HEADER_ROW, 1, 1, 15);

  // Set header values
  const headers = [
    [
      COLUMNS.NAME.header,
      COLUMNS.JOIN.header,
      COLUMNS.TITLE.header,
      COLUMNS.LEAD_PIC.header,
      COLUMNS.PROJECT.header,
      COLUMNS.NP.header,
      COLUMNS.EMAIL.header,
      COLUMNS.EMAIL_2.header,
      COLUMNS.STATUS.header,
      COLUMNS.AUTOMATION.header,
      COLUMNS.GITHUB.header,
      COLUMNS.HP.header,
      COLUMNS.ROLE.header,
      COLUMNS.VPN_ABC.header,
      COLUMNS.VPN_HUWAWEI.header
    ]
  ];

  headerRange.setValues(headers);

  // Style header
  headerRange
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Set header row height
  sheet.setRowHeight(HEADER_ROW, 40);

  Logger.log('✅ Header formatted');
}

/**
 * Format data rows with alternating colors
 */
function formatDataRows(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    return;
  }

  // Apply alternating row colors
  for (let row = DATA_START_ROW; row <= lastRow; row++) {
    const rowRange = sheet.getRange(row, 1, 1, 15);

    // Check if this is a section header (like "MBG", "Tidak Lolos", "Riset/Security")
    const nameCell = sheet.getRange(row, 1).getValue().toString().trim();

    if (nameCell === 'MBG' || nameCell.includes('Tidak Lolos') || nameCell.includes('Riset') || nameCell.includes('Security')) {
      // Section header - darker background
      rowRange
        .setBackground('#e0e0e0')
        .setFontWeight('bold')
        .setFontSize(12)
        .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    } else if (nameCell === '') {
      // Empty row - white background
      rowRange
        .setBackground('#ffffff')
        .setBorder(false, false, false, false, false, false);
    } else {
      // Regular data row - alternating colors
      const bg = (row - DATA_START_ROW) % 2 === 0 ? '#ffffff' : '#f8f9fa';
      rowRange
        .setBackground(bg)
        .setFontWeight('normal')
        .setFontSize(10)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  // Center align specific columns
  const centerCols = ['B', 'F', 'I', 'J', 'N', 'O']; // Join, NP, Status, Automation, VPNs
  centerCols.forEach(col => {
    const colIndex = col.charCodeAt(0) - 64;
    sheet.getRange(DATA_START_ROW, colIndex, lastRow - DATA_START_ROW + 1, 1)
      .setHorizontalAlignment('center');
  });

  // Format date column
  sheet.getRange(DATA_START_ROW, 2, lastRow - DATA_START_ROW + 1, 1)
    .setNumberFormat('yyyy-mm-dd');

  Logger.log('✅ Data rows formatted');
}

/**
 * Add data validation for dropdowns
 */
function addDataValidation(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    return;
  }

  // Status dropdown (Column I)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Onboard', 'Tidak Ada Kabar', 'Digispark', 'Resign', 'Contract End'], true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(DATA_START_ROW, 9, lastRow - DATA_START_ROW + 1, 1)
    .setDataValidation(statusRule);

  // Role dropdown (Column M)
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'QA Team Lead',
      'Senior Quality Engineer',
      'Quality Engineer',
      'PIC QE',
      'Lead Project',
      'Intern Quality Engineer',
      'Security Engineer',
      'UX Research'
    ], true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(DATA_START_ROW, 13, lastRow - DATA_START_ROW + 1, 1)
    .setDataValidation(roleRule);

  // Automation dropdown (Column J) - TRUE/FALSE
  const automationRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  sheet.getRange(DATA_START_ROW, 10, lastRow - DATA_START_ROW + 1, 1)
    .setDataValidation(automationRule);

  // VPN ABC dropdown (Column N) - TRUE/FALSE
  sheet.getRange(DATA_START_ROW, 14, lastRow - DATA_START_ROW + 1, 1)
    .setDataValidation(automationRule);

  // VPN Huwawei dropdown (Column O) - TRUE/FALSE
  sheet.getRange(DATA_START_ROW, 15, lastRow - DATA_START_ROW + 1, 1)
    .setDataValidation(automationRule);

  Logger.log('✅ Data validation added');
}

/**
 * Add conditional formatting for status
 */
function addConditionalFormatting(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    return;
  }

  // Clear existing conditional format rules
  sheet.clearConditionalFormatRules();

  const rules = [];

  // Status column (I) - Green for "Onboard"
  const onboardRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Onboard')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setRanges([sheet.getRange(DATA_START_ROW, 9, lastRow - DATA_START_ROW + 1, 1)])
    .build();
  rules.push(onboardRule);

  // Status column (I) - Yellow for "Tidak Ada Kabar"
  const noNewsRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Tidak Ada Kabar')
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([sheet.getRange(DATA_START_ROW, 9, lastRow - DATA_START_ROW + 1, 1)])
    .build();
  rules.push(noNewsRule);

  // Status column (I) - Red for "Resign", "Contract End", "Digispark"
  const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Resign')
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([sheet.getRange(DATA_START_ROW, 9, lastRow - DATA_START_ROW + 1, 1)])
    .build();
  rules.push(inactiveRule);

  const contractEndRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Contract End')
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([sheet.getRange(DATA_START_ROW, 9, lastRow - DATA_START_ROW + 1, 1)])
    .build();
  rules.push(contractEndRule);

  const digisparkRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Digispark')
    .setBackground('#f8d7da')
    .setFontColor('#721c24')
    .setRanges([sheet.getRange(DATA_START_ROW, 9, lastRow - DATA_START_ROW + 1, 1)])
    .build();
  rules.push(digisparkRule);

  // Apply all rules
  sheet.setConditionalFormatRules(rules);

  Logger.log('✅ Conditional formatting added');
}

/**
 * Clean up excessive empty rows at the end
 */
function cleanupEmptyRows(sheet) {
  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();

  // Keep some buffer rows (e.g., 20 empty rows)
  const bufferRows = 20;
  const rowsToDelete = maxRows - lastRow - bufferRows;

  if (rowsToDelete > 0) {
    sheet.deleteRows(lastRow + bufferRows + 1, rowsToDelete);
    Logger.log('✅ Deleted ' + rowsToDelete + ' empty rows');
  }
}

/**
 * Add new team member with pre-formatted row
 */
function addNewTeamMember() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAME);

  if (!sheet) {
    ui.alert('Error', 'Tab "' + TAB_NAME + '" not found', ui.ButtonSet.OK);
    return;
  }

  // Get name
  const nameResponse = ui.prompt(
    'Add New Team Member',
    'Enter full name:',
    ui.ButtonSet.OK_CANCEL
  );

  if (nameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const name = nameResponse.getResponseText().trim();
  if (!name) {
    ui.alert('Name cannot be empty.');
    return;
  }

  // Find next available row
  let insertRow = DATA_START_ROW;
  const lastRow = sheet.getLastRow();

  // Insert after last data row (before section headers)
  for (let row = DATA_START_ROW; row <= lastRow; row++) {
    const cellValue = sheet.getRange(row, 1).getValue().toString().trim();
    if (cellValue === 'MBG' || cellValue.includes('Tidak Lolos') || cellValue === '') {
      insertRow = row;
      break;
    }
  }

  // Insert new row
  sheet.insertRowBefore(insertRow);

  // Set default values
  const newRow = [
    name,                    // Name
    new Date(),             // Join Date
    '',                     // Title
    '',                     // Lead/PIC
    '',                     // Project
    '',                     // NP
    '',                     // Email
    '',                     // Email 2
    'Onboard',              // Status
    false,                  // Automation
    '',                     // Github
    '',                     // Phone
    'Quality Engineer',     // Role
    false,                  // VPN ABC
    false                   // VPN Huwawei
  ];

  sheet.getRange(insertRow, 1, 1, 15).setValues([newRow]);

  // Apply formatting
  const bg = (insertRow - DATA_START_ROW) % 2 === 0 ? '#ffffff' : '#f8f9fa';
  sheet.getRange(insertRow, 1, 1, 15)
    .setBackground(bg)
    .setFontWeight('normal')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  // Apply data validation
  addDataValidation(sheet);

  // Apply conditional formatting
  addConditionalFormatting(sheet);

  // Highlight new row temporarily
  sheet.setActiveRange(sheet.getRange(insertRow, 1, 1, 15));

  ui.alert(
    'Team Member Added! ✅',
    'Name: ' + name + '\n' +
    'Row: ' + insertRow + '\n\n' +
    'Please fill in the remaining details (Email, Project, etc.)',
    ui.ButtonSet.OK
  );

  Logger.log('✅ New team member added: ' + name + ' at row ' + insertRow);
}

/**
 * Create summary statistics
 */
function createSummary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAME);

  if (!sheet) {
    ui.alert('Error', 'Tab "' + TAB_NAME + '" not found', ui.ButtonSet.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 15).getValues();

  // Count statistics
  let totalActive = 0;
  let totalInactive = 0;
  const roleCount = {};
  const statusCount = {};

  data.forEach(row => {
    const name = row[0].toString().trim();
    const status = row[8].toString().trim();
    const role = row[12].toString().trim();

    // Skip section headers and empty rows
    if (!name || name === 'MBG' || name.includes('Tidak Lolos') || name.includes('Riset')) {
      return;
    }

    // Count status
    if (status === 'Onboard') {
      totalActive++;
    } else if (status) {
      totalInactive++;
    }

    statusCount[status] = (statusCount[status] || 0) + 1;

    // Count roles
    if (role) {
      roleCount[role] = (roleCount[role] || 0) + 1;
    }
  });

  // Build summary text
  let summary = '📊 TEAM MEMBER SUMMARY\n\n';
  summary += '═══════════════════════════════════\n\n';
  summary += 'TOTAL:\n';
  summary += '• Total Members: ' + (totalActive + totalInactive) + '\n';
  summary += '• Active (Onboard): ' + totalActive + '\n';
  summary += '• Inactive: ' + totalInactive + '\n\n';
  summary += '═══════════════════════════════════\n\n';
  summary += 'BY STATUS:\n';
  Object.keys(statusCount).sort().forEach(status => {
    if (status) {
      summary += '• ' + status + ': ' + statusCount[status] + '\n';
    }
  });
  summary += '\n═══════════════════════════════════\n\n';
  summary += 'BY ROLE:\n';
  Object.keys(roleCount).sort().forEach(role => {
    if (role) {
      summary += '• ' + role + ': ' + roleCount[role] + '\n';
    }
  });
  summary += '\n═══════════════════════════════════';

  ui.alert(summary);
  Logger.log('Summary created');
}
