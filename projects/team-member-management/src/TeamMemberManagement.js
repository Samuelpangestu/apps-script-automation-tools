/**
 * TeamMemberManagement.js — Team Member Management System
 * ═══════════════════════════════════════════════════════════════════════
 * Create and manage Team Member tab from scratch
 * Portable for use in any company/organization
 * ═══════════════════════════════════════════════════════════════════════
 */

const TAB_NAME = 'Team Member';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Column definitions with optimal widths
const COLUMNS = {
  NAME: { index: 1, letter: 'A', width: 200, header: 'Name' },
  JOIN: { index: 2, letter: 'B', width: 100, header: 'Join Date' },
  TITLE: { index: 3, letter: 'C', width: 150, header: 'Title' },
  LEAD_PIC: { index: 4, letter: 'D', width: 120, header: 'Lead/PIC' },
  PROJECT: { index: 5, letter: 'E', width: 250, header: 'Project' },
  NP: { index: 6, letter: 'F', width: 80, header: 'NP' },
  EMAIL: { index: 7, letter: 'G', width: 250, header: 'Email' },
  EMAIL_2: { index: 8, letter: 'H', width: 250, header: 'Email 2' },
  STATUS: { index: 9, letter: 'I', width: 120, header: 'Status' },
  AUTOMATION: { index: 10, letter: 'J', width: 100, header: 'Automation' },
  GITHUB: { index: 11, letter: 'K', width: 180, header: 'Github' },
  HP: { index: 12, letter: 'L', width: 130, header: 'Phone' },
  ROLE: { index: 13, letter: 'M', width: 200, header: 'Role' },
  VPN_ABC: { index: 14, letter: 'N', width: 90, header: 'VPN ABC' },
  VPN_HUWAWEI: { index: 15, letter: 'O', width: 120, header: 'VPN Huwawei' }
};

const TOTAL_COLUMNS = 15;

// Status options
const STATUS_OPTIONS = [
  'Onboard',
  'Tidak Ada Kabar',
  'Digispark',
  'Resign',
  'Contract End'
];

// Role options
const ROLE_OPTIONS = [
  'QA Team Lead',
  'Senior Quality Engineer',
  'Quality Engineer',
  'PIC QE',
  'Lead Project',
  'Intern Quality Engineer',
  'Security Engineer',
  'UX Research'
];

/**
 * Create Team Member tab from scratch
 */
function createTeamMemberTab() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if tab already exists
  let existingSheet = ss.getSheetByName(TAB_NAME);
  if (existingSheet) {
    const response = ui.alert(
      'Tab Already Exists',
      'A "' + TAB_NAME + '" tab already exists.\n\n' +
      'Do you want to REBUILD it?\n' +
      '(All existing data will be DELETED)\n\n' +
      'Click YES to rebuild\n' +
      'Click NO to cancel',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Delete existing sheet
    ss.deleteSheet(existingSheet);
    Logger.log('Existing sheet deleted');
  }

  ui.alert('Creating Team Member tab...');

  try {
    // Create new sheet
    const sheet = ss.insertSheet(TAB_NAME, 0);

    // Step 1: Set up structure
    setupStructure(sheet);

    // Step 2: Create header
    createHeader(sheet);

    // Step 3: Set column widths
    setColumnWidths(sheet);

    // Step 4: Add sample data (optional)
    addSampleData(sheet);

    // Step 5: Add data validation
    addDataValidation(sheet);

    // Step 6: Add conditional formatting
    addConditionalFormatting(sheet);

    // Step 7: Freeze and filter
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
    sheet.getRange(HEADER_ROW, 1, 100, TOTAL_COLUMNS).createFilter();

    // Step 8: Add instructions
    addInstructions(sheet);

    Logger.log('✅ Team Member tab created successfully');

    ui.alert(
      'Success! ✅',
      'Team Member tab has been created successfully!\n\n' +
      'Features:\n' +
      '✓ Professional header\n' +
      '✓ Data validation (dropdowns & checkboxes)\n' +
      '✓ Conditional formatting (color-coded status)\n' +
      '✓ Sample data included\n' +
      '✓ Ready to use!\n\n' +
      'You can now:\n' +
      '• Add team members (Menu → Add Team Member)\n' +
      '• Export data (Menu → Export Data)\n' +
      '• Import data (Menu → Import Data)',
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log('❌ Error creating tab: ' + e.message);
    ui.alert('Error', 'Failed to create tab:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Setup basic structure (rows and columns)
 */
function setupStructure(sheet) {
  // Set to 100 rows initially (will auto-expand as needed)
  const maxRows = 100;
  const currentRows = sheet.getMaxRows();

  if (currentRows < maxRows) {
    sheet.insertRowsAfter(currentRows, maxRows - currentRows);
  } else if (currentRows > maxRows) {
    sheet.deleteRows(maxRows + 1, currentRows - maxRows);
  }

  // Set to 15 columns (A-O)
  const maxCols = TOTAL_COLUMNS;
  const currentCols = sheet.getMaxColumns();

  if (currentCols < maxCols) {
    sheet.insertColumnsAfter(currentCols, maxCols - currentCols);
  } else if (currentCols > maxCols) {
    sheet.deleteColumns(maxCols + 1, currentCols - maxCols);
  }

  Logger.log('✅ Structure setup complete');
}

/**
 * Create and format header row
 */
function createHeader(sheet) {
  // Header values
  const headers = [[
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
  ]];

  const headerRange = sheet.getRange(HEADER_ROW, 1, 1, TOTAL_COLUMNS);
  headerRange.setValues(headers);

  // Format header
  headerRange
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.setRowHeight(HEADER_ROW, 40);

  Logger.log('✅ Header created');
}

/**
 * Set optimal column widths
 */
function setColumnWidths(sheet) {
  Object.keys(COLUMNS).forEach(key => {
    const col = COLUMNS[key];
    sheet.setColumnWidth(col.index, col.width);
  });

  Logger.log('✅ Column widths set');
}

/**
 * Add sample data for demonstration
 */
function addSampleData(sheet) {
  const sampleData = [
    ['John Doe', new Date('2024-01-15'), 'QA Team Lead', 'Lead', 'Project Alpha, Project Beta', 'N001', 'john.doe@company.com', '', 'Onboard', true, 'johndoe', '081234567890', 'QA Team Lead', true, true],
    ['Jane Smith', new Date('2024-02-01'), 'Senior QE', 'PIC', 'Project Gamma', 'N002', 'jane.smith@company.com', '', 'Onboard', true, 'janesmith', '081234567891', 'Senior Quality Engineer', true, false],
    ['Bob Wilson', new Date('2024-03-10'), 'Quality Engineer', '', 'Project Delta', 'N003', 'bob.wilson@company.com', '', 'Onboard', false, 'bobwilson', '081234567892', 'Quality Engineer', false, false],
    ['Alice Brown', new Date('2023-12-01'), 'QE Intern', '', 'Project Epsilon', 'I001', 'alice.brown@company.com', '', 'Onboard', false, '', '081234567893', 'Intern Quality Engineer', false, false]
  ];

  sheet.getRange(DATA_START_ROW, 1, sampleData.length, TOTAL_COLUMNS).setValues(sampleData);

  // Format data rows
  for (let i = 0; i < sampleData.length; i++) {
    const row = DATA_START_ROW + i;
    const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';

    sheet.getRange(row, 1, 1, TOTAL_COLUMNS)
      .setBackground(bg)
      .setFontWeight('normal')
      .setFontSize(10)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Center align specific columns
  const centerCols = [2, 6, 9, 10, 14, 15]; // Join, NP, Status, Automation, VPNs
  centerCols.forEach(colIndex => {
    sheet.getRange(DATA_START_ROW, colIndex, sampleData.length, 1)
      .setHorizontalAlignment('center');
  });

  // Format date column
  sheet.getRange(DATA_START_ROW, 2, sampleData.length, 1)
    .setNumberFormat('yyyy-mm-dd');

  Logger.log('✅ Sample data added');
}

/**
 * Add data validation for dropdowns and checkboxes
 */
function addDataValidation(sheet) {
  const maxRows = 100;

  // Status dropdown (Column I)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS_OPTIONS, true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(DATA_START_ROW, COLUMNS.STATUS.index, maxRows, 1)
    .setDataValidation(statusRule);

  // Role dropdown (Column M)
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ROLE_OPTIONS, true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(DATA_START_ROW, COLUMNS.ROLE.index, maxRows, 1)
    .setDataValidation(roleRule);

  // Checkboxes
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  sheet.getRange(DATA_START_ROW, COLUMNS.AUTOMATION.index, maxRows, 1)
    .setDataValidation(checkboxRule);

  sheet.getRange(DATA_START_ROW, COLUMNS.VPN_ABC.index, maxRows, 1)
    .setDataValidation(checkboxRule);

  sheet.getRange(DATA_START_ROW, COLUMNS.VPN_HUWAWEI.index, maxRows, 1)
    .setDataValidation(checkboxRule);

  Logger.log('✅ Data validation added');
}

/**
 * Add conditional formatting for status column
 */
function addConditionalFormatting(sheet) {
  const rules = [];
  const maxRows = 100;
  const statusRange = sheet.getRange(DATA_START_ROW, COLUMNS.STATUS.index, maxRows, 1);

  // Green: Onboard
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Onboard')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setRanges([statusRange])
    .build());

  // Yellow: Tidak Ada Kabar
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Tidak Ada Kabar')
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([statusRange])
    .build());

  // Red: Inactive statuses
  ['Resign', 'Contract End', 'Digispark'].forEach(status => {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground('#f8d7da')
      .setFontColor('#721c24')
      .setRanges([statusRange])
      .build());
  });

  sheet.setConditionalFormatRules(rules);

  Logger.log('✅ Conditional formatting added');
}

/**
 * Add instructions sheet
 */
function addInstructions(sheet) {
  const lastDataRow = sheet.getLastRow();
  const instructionsRow = lastDataRow + 3;

  sheet.getRange(instructionsRow, 1, 1, TOTAL_COLUMNS).merge()
    .setValue(
      '💡 INSTRUCTIONS: Use Menu → Add Team Member to add new members | ' +
      'Use dropdowns for Status and Role | ' +
      'Checkboxes for Automation and VPNs | ' +
      'Status colors: 🟢 Green (Active), 🟡 Yellow (Pending), 🔴 Red (Inactive)'
    )
    .setBackground('#fff9c4')
    .setFontStyle('italic')
    .setFontSize(9)
    .setWrap(true)
    .setVerticalAlignment('middle');

  sheet.setRowHeight(instructionsRow, 60);

  Logger.log('✅ Instructions added');
}

/**
 * Add new team member
 */
function addTeamMember() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAME);

  if (!sheet) {
    ui.alert('Error', 'Team Member tab not found.\n\nCreate it first: Menu → Create Team Member Tab', ui.ButtonSet.OK);
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

  // Find next row
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;

  // Default values
  const rowData = [
    name,                       // Name
    new Date(),                 // Join Date
    '',                         // Title
    '',                         // Lead/PIC
    '',                         // Project
    '',                         // NP
    '',                         // Email
    '',                         // Email 2
    'Onboard',                  // Status
    false,                      // Automation
    '',                         // Github
    '',                         // Phone
    'Quality Engineer',         // Role
    false,                      // VPN ABC
    false                       // VPN Huwawei
  ];

  sheet.getRange(newRow, 1, 1, TOTAL_COLUMNS).setValues([rowData]);

  // Format new row
  const bg = (newRow - DATA_START_ROW) % 2 === 0 ? '#ffffff' : '#f8f9fa';
  sheet.getRange(newRow, 1, 1, TOTAL_COLUMNS)
    .setBackground(bg)
    .setFontWeight('normal')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  // Center align specific columns
  [2, 6, 9, 10, 14, 15].forEach(colIndex => {
    sheet.getRange(newRow, colIndex).setHorizontalAlignment('center');
  });

  // Date format
  sheet.getRange(newRow, 2).setNumberFormat('yyyy-mm-dd');

  // Re-apply data validation
  addDataValidation(sheet);

  // Highlight new row
  sheet.setActiveRange(sheet.getRange(newRow, 1, 1, TOTAL_COLUMNS));

  ui.alert(
    'Success! ✅',
    'Team member added: ' + name + '\n' +
    'Row: ' + newRow + '\n\n' +
    'Please fill in the remaining details.',
    ui.ButtonSet.OK
  );

  Logger.log('✅ Added team member: ' + name + ' at row ' + newRow);
}

/**
 * Export team member data to CSV
 */
function exportData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAME);

  if (!sheet) {
    ui.alert('Error', 'Team Member tab not found.', ui.ButtonSet.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, TOTAL_COLUMNS).getValues();

  // Convert to CSV
  let csv = '';
  data.forEach(row => {
    csv += row.map(cell => {
      // Escape quotes and wrap in quotes if contains comma
      const str = cell.toString();
      if (str.includes(',') || str.includes('"')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',') + '\n';
  });

  // Show CSV in dialog (user can copy)
  const html = HtmlService.createHtmlOutput(
    '<h3>Team Member Data (CSV)</h3>' +
    '<p>Copy this data and save as .csv file:</p>' +
    '<textarea style="width:100%;height:400px;font-family:monospace;font-size:12px;">' +
    csv +
    '</textarea>'
  ).setWidth(600).setHeight(500);

  ui.showModalDialog(html, 'Export Team Member Data');

  Logger.log('✅ Data exported');
}

/**
 * View summary statistics
 */
function viewSummary() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAME);

  if (!sheet) {
    ui.alert('Error', 'Team Member tab not found.', ui.ButtonSet.OK);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    ui.alert('No data', 'No team members found.', ui.ButtonSet.OK);
    return;
  }

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, TOTAL_COLUMNS).getValues();

  let totalActive = 0;
  let totalInactive = 0;
  const statusCount = {};
  const roleCount = {};

  data.forEach(row => {
    const name = row[0].toString().trim();
    if (!name) return;

    const status = row[8].toString().trim();
    const role = row[12].toString().trim();

    if (status === 'Onboard') totalActive++;
    else if (status) totalInactive++;

    if (status) statusCount[status] = (statusCount[status] || 0) + 1;
    if (role) roleCount[role] = (roleCount[role] || 0) + 1;
  });

  let summary = '📊 TEAM MEMBER SUMMARY\n\n';
  summary += '═══════════════════════════════════\n\n';
  summary += 'TOTAL:\n';
  summary += '• Total Members: ' + (totalActive + totalInactive) + '\n';
  summary += '• Active (Onboard): ' + totalActive + '\n';
  summary += '• Inactive: ' + totalInactive + '\n\n';
  summary += '═══════════════════════════════════\n\n';
  summary += 'BY STATUS:\n';
  Object.keys(statusCount).sort().forEach(status => {
    summary += '• ' + status + ': ' + statusCount[status] + '\n';
  });
  summary += '\n═══════════════════════════════════\n\n';
  summary += 'BY ROLE:\n';
  Object.keys(roleCount).sort().forEach(role => {
    summary += '• ' + role + ': ' + roleCount[role] + '\n';
  });
  summary += '\n═══════════════════════════════════';

  ui.alert(summary);
}
