/**
 * StandupLog.js
 * Manage bi-daily standup log sheets
 * - Initialize standup sheet structure
 * - Auto-generate rows for standup sessions
 * - Protect existing data when generating new rows
 */

/**
 * Initialize standup sheet for a project
 * @param {string} projectName - 'Project A' or 'Project B'
 */
function initializeStandupSheet(projectName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = projectName + ' Standup';
  let standupSheet = ss.getSheetByName(sheetName);

  if (!standupSheet) {
    standupSheet = ss.insertSheet(sheetName);
    standupSheet.setTabColor(projectName === 'Project A' ? '#34A853' : '#FBBC04'); // Green for A, Yellow for B
  }

  // Clear existing structure (but keep data rows if any)
  const lastRow = standupSheet.getLastRow();
  if (lastRow <= 3) {
    standupSheet.clear();
  } else {
    // Only clear header rows
    standupSheet.getRange(1, 1, 3, 8).clear();
  }

  // ═══════════════════════════════════════════════════════════════
  // HEADER
  // ═══════════════════════════════════════════════════════════════

  standupSheet.getRange('A1:H1').merge();
  standupSheet.getRange('A1').setValue(`📋 ${projectName.toUpperCase()} - BI-DAILY STANDUP LOG`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(projectName === 'Project A' ? '#34A853' : '#FBBC04')
    .setFontColor('#FFFFFF');

  // Column headers
  const headers = [
    ['Date', 'Day', 'Person', 'Task Type', 'Description', 'Status', 'Blocker', 'Notes']
  ];

  standupSheet.getRange(3, 1, 1, 8).setValues(headers);
  standupSheet.getRange(3, 1, 1, 8)
    .setFontWeight('bold')
    .setBackground('#E8F0FE')
    .setFontColor('#1A73E8')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#4285F4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Set column widths
  standupSheet.setColumnWidth(1, 100);  // Date
  standupSheet.setColumnWidth(2, 90);   // Day
  standupSheet.setColumnWidth(3, 100);  // Person
  standupSheet.setColumnWidth(4, 110);  // Task Type
  standupSheet.setColumnWidth(5, 350);  // Description
  standupSheet.setColumnWidth(6, 100);  // Status
  standupSheet.setColumnWidth(7, 200);  // Blocker
  standupSheet.setColumnWidth(8, 200);  // Notes

  // Freeze header rows
  standupSheet.setFrozenRows(3);

  Logger.log(`✅ ${sheetName} initialized`);
}

/**
 * Generate standup rows for a specific date
 * @param {string} projectName - 'Project A' or 'Project B'
 * @param {Date} date - Date object for the standup session
 * @param {boolean} skipIfExists - If true, skip generation if rows already exist for this date
 * @returns {number} Number of rows generated
 */
function generateStandupRows(projectName, date, skipIfExists = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = projectName + ' Standup';
  const standupSheet = ss.getSheetByName(sheetName);

  if (!standupSheet) {
    throw new Error(`Sheet "${sheetName}" not found. Run initializeStandupSheet() first.`);
  }

  // Get project config
  const config = getProjectConfig(projectName);

  if (!config.teamMembers || config.teamMembers.length === 0) {
    throw new Error(`No team members configured for ${projectName}`);
  }

  // Format date
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const dayName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEEE');
  const dayShort = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEE');

  // Check if rows already exist for this date
  if (skipIfExists) {
    const lastRow = standupSheet.getLastRow();
    if (lastRow > 3) {
      const existingDates = standupSheet.getRange(4, 1, lastRow - 3, 1).getValues();
      const dateExists = existingDates.some(row => {
        const cellDate = row[0];
        if (cellDate instanceof Date) {
          const cellDateStr = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          return cellDateStr === dateStr;
        } else if (typeof cellDate === 'string') {
          return cellDate === dateStr;
        }
        return false;
      });

      if (dateExists) {
        Logger.log(`⚠️ Rows for ${dateStr} already exist in ${sheetName}. Skipping generation.`);
        return 0;
      }
    }
  }

  // Prepare rows data (3 rows per person: Done, In Progress, Blocker)
  const rowsData = [];

  config.teamMembers.forEach(person => {
    // Row 1: Done
    rowsData.push([
      dateStr,
      dayShort,
      person,
      '✅ Done',
      '',  // Description (empty for user to fill)
      'Completed',
      'None',
      ''   // Notes
    ]);

    // Row 2: In Progress
    rowsData.push([
      dateStr,
      dayShort,
      person,
      '📋 In Progress',
      '',  // Description (empty for user to fill)
      'Ongoing',
      'None',
      ''   // Notes
    ]);

    // Row 3: Blocker
    rowsData.push([
      dateStr,
      dayShort,
      person,
      '🚨 Blocker',
      '',  // Description (empty for user to fill)
      'Blocked',
      '',  // Blocker description (empty for user to fill)
      ''   // Notes
    ]);
  });

  // Insert rows at the end
  const lastRow = standupSheet.getLastRow();
  const startRow = lastRow + 1;

  standupSheet.getRange(startRow, 1, rowsData.length, 8).setValues(rowsData);

  // Apply formatting
  const dataRange = standupSheet.getRange(startRow, 1, rowsData.length, 8);

  // Borders
  dataRange.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Conditional formatting for Task Type
  for (let i = 0; i < rowsData.length; i++) {
    const currentRow = startRow + i;
    const taskType = rowsData[i][3];

    if (taskType === '✅ Done') {
      standupSheet.getRange(currentRow, 1, 1, 8).setBackground('#D9EAD3'); // Light green
    } else if (taskType === '📋 In Progress') {
      standupSheet.getRange(currentRow, 1, 1, 8).setBackground('#FFF2CC'); // Light yellow
    } else if (taskType === '🚨 Blocker') {
      standupSheet.getRange(currentRow, 1, 1, 8).setBackground('#F4CCCC'); // Light red
    }
  }

  // Center align Date, Day, Person, Task Type, Status columns
  standupSheet.getRange(startRow, 1, rowsData.length, 1).setHorizontalAlignment('center'); // Date
  standupSheet.getRange(startRow, 2, rowsData.length, 1).setHorizontalAlignment('center'); // Day
  standupSheet.getRange(startRow, 3, rowsData.length, 1).setHorizontalAlignment('center'); // Person
  standupSheet.getRange(startRow, 4, rowsData.length, 1).setHorizontalAlignment('center'); // Task Type
  standupSheet.getRange(startRow, 6, rowsData.length, 1).setHorizontalAlignment('center'); // Status

  // Wrap text for Description, Blocker, Notes
  standupSheet.getRange(startRow, 5, rowsData.length, 1).setWrap(true); // Description
  standupSheet.getRange(startRow, 7, rowsData.length, 1).setWrap(true); // Blocker
  standupSheet.getRange(startRow, 8, rowsData.length, 1).setWrap(true); // Notes

  Logger.log(`✅ Generated ${rowsData.length} rows for ${dateStr} (${dayName}) in ${sheetName}`);

  return rowsData.length;
}

/**
 * Get standup data for a specific date
 * @param {string} projectName - 'Project A' or 'Project B'
 * @param {Date} date - Date object
 * @returns {Array} Array of standup data objects
 */
function getStandupDataForDate(projectName, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = projectName + ' Standup';
  const standupSheet = ss.getSheetByName(sheetName);

  if (!standupSheet) {
    Logger.log(`⚠️ Sheet "${sheetName}" not found`);
    return [];
  }

  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const lastRow = standupSheet.getLastRow();
  if (lastRow <= 3) {
    Logger.log(`⚠️ No data in ${sheetName}`);
    return [];
  }

  // Read all data
  const allData = standupSheet.getRange(4, 1, lastRow - 3, 8).getValues();

  // Filter by date
  const standupData = [];

  allData.forEach(row => {
    const rowDate = row[0];
    let rowDateStr = '';

    if (rowDate instanceof Date) {
      rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof rowDate === 'string') {
      rowDateStr = rowDate;
    }

    if (rowDateStr === dateStr) {
      standupData.push({
        date: row[0],
        day: row[1],
        person: row[2],
        taskType: row[3],
        description: row[4],
        status: row[5],
        blocker: row[6],
        notes: row[7]
      });
    }
  });

  Logger.log(`Found ${standupData.length} rows for ${dateStr} in ${sheetName}`);

  return standupData;
}

/**
 * Get standup summary grouped by person
 * @param {Array} standupData - Array of standup data objects from getStandupDataForDate()
 * @returns {Array} Array of person summaries
 */
function getStandupSummaryByPerson(standupData) {
  const summaryMap = {};

  standupData.forEach(row => {
    const person = row.person;

    if (!summaryMap[person]) {
      summaryMap[person] = {
        person: person,
        done: [],
        inProgress: [],
        blockers: []
      };
    }

    const description = String(row.description || '').trim();

    if (row.taskType === '✅ Done' && description) {
      summaryMap[person].done.push(description);
    } else if (row.taskType === '📋 In Progress' && description) {
      summaryMap[person].inProgress.push(description);
    } else if (row.taskType === '🚨 Blocker' && description) {
      summaryMap[person].blockers.push(description);
    }
  });

  // Convert map to array
  const summaryArray = Object.values(summaryMap);

  return summaryArray;
}
