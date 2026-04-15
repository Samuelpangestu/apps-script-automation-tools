/**
 * StandupLog.js
 * Manage bi-daily standup log sheets
 * - Initialize standup sheet structure
 * - Auto-generate rows for standup sessions
 * - Protect existing data when generating new rows
 */

/**
 * Initialize standup sheet for a project
 * @param {string} projectName - 'SIPGN' or 'INADigital/Internal'
 */
function initializeStandupSheet(projectName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = projectName + ' Standup';
  let standupSheet = ss.getSheetByName(sheetName);

  if (!standupSheet) {
    standupSheet = ss.insertSheet(sheetName);
  }

  // Clear existing structure (but keep data rows if any)
  const lastRow = standupSheet.getLastRow();
  if (lastRow <= 4) {
    standupSheet.clear();
  } else {
    // Only clear header rows
    standupSheet.getRange(1, 1, 4, 6).clear();
  }

  // Get project config for team members
  let config;
  try {
    config = getProjectConfig(projectName);
    Logger.log(`Config for ${projectName}: ${JSON.stringify(config.teamMembers)}`);
  } catch (e) {
    Logger.log(`⚠️ Error getting config for ${projectName}: ${e.message}`);
    // Fallback to empty config
    config = { teamMembers: ['User 1', 'User 2', 'User 3'] };
  }

  // ═══════════════════════════════════════════════════════════════
  // HEADER
  // ═══════════════════════════════════════════════════════════════

  standupSheet.getRange('A1:G1').merge();
  standupSheet.getRange('A1').setValue(`📋 ${projectName.toUpperCase()} - BI-DAILY STANDUP LOG`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');

  // Reminder note
  standupSheet.getRange('A2:G2').merge();
  standupSheet.getRange('A2').setValue('📝 REMINDER: Fill Done Since Last, Plan Next 2 Days, Blockers (if any), and Notes. Update test execution progress and Jira tickets.')
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setBackground('#FFF3E0')
    .setFontColor('#E65100')
    .setWrap(true);

  // Column headers
  const headers = [
    ['Date', 'Day', 'Person', 'Done Since Last', 'Plan Next 2 Days', 'Blockers', 'Notes']
  ];

  standupSheet.getRange(4, 1, 1, 7).setValues(headers);
  standupSheet.getRange(4, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground('#E8F0FE')
    .setFontColor('#1A73E8')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#4285F4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Set column widths
  standupSheet.setColumnWidth(1, 100);  // Date
  standupSheet.setColumnWidth(2, 90);   // Day
  standupSheet.setColumnWidth(3, 120);  // Person
  standupSheet.setColumnWidth(4, 300);  // Done Since Last
  standupSheet.setColumnWidth(5, 300);  // Plan Next 2 Days
  standupSheet.setColumnWidth(6, 250);  // Blockers
  standupSheet.setColumnWidth(7, 200);  // Notes

  // Data validation for Day (dropdown)
  const dayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Senin', 'Rabu', 'Jumat'], true)
    .setAllowInvalid(false)
    .build();

  // Data validation for Person (dropdown from team members)
  const personRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(config.teamMembers, true)
    .setAllowInvalid(false)
    .build();

  // Apply validation to entire columns (starting from row 5)
  standupSheet.getRange('B5:B1000').setDataValidation(dayRule);
  standupSheet.getRange('C5:C1000').setDataValidation(personRule);

  // ═══════════════════════════════════════════════════════════════
  // CONDITIONAL FORMATTING
  // ═══════════════════════════════════════════════════════════════

  // Clear existing rules for this range
  const rules = standupSheet.getConditionalFormatRules();
  const filteredRules = rules.filter(rule => {
    const range = rule.getRanges()[0];
    return !(range.getA1Notation() === 'A5:A1000' ||
             range.getA1Notation() === 'B5:B1000' ||
             range.getA1Notation() === 'C5:C1000' ||
             range.getA1Notation() === 'F5:F1000');
  });

  // ─────────────────────────────────────────────────────────────
  // BLOCKERS COLUMN (F5:F1000) - Highlight if not empty
  // ─────────────────────────────────────────────────────────────
  const blockersRangeForFormatting = standupSheet.getRange('F5:F1000');

  // Blocker exists - Light Red
  const blockerRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('.')  // Any text with period (basic check for content)
    .setBackground('#F4CCCC')
    .setRanges([blockersRangeForFormatting])
    .build();

  // Alternative: Highlight any non-empty blocker
  const blockerNotEmptyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(LEN(F5)>0, F5<>"")')
    .setBackground('#F4CCCC')
    .setRanges([blockersRangeForFormatting])
    .build();

  // ─────────────────────────────────────────────────────────────
  // DAY COLUMN (B5:B1000) - Subtle colors per day
  // ─────────────────────────────────────────────────────────────
  const dayRangeForFormatting = standupSheet.getRange('B5:B1000');

  // Senin - Light Blue
  const seninRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Senin')
    .setBackground('#E8F4FD')
    .setRanges([dayRangeForFormatting])
    .build();

  // Rabu - Light Green
  const rabuRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Rabu')
    .setBackground('#E8F5E9')
    .setRanges([dayRangeForFormatting])
    .build();

  // Jumat - Light Purple
  const jumatRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Jumat')
    .setBackground('#F3E5F5')
    .setRanges([dayRangeForFormatting])
    .build();

  // ─────────────────────────────────────────────────────────────
  // PERSON COLUMN (C5:C1000) - Very subtle alternating per person
  // ─────────────────────────────────────────────────────────────
  const personRangeForFormatting = standupSheet.getRange('C5:C1000');

  // Create subtle alternating colors for each team member
  const personRules = [];
  const subtleColors = [
    '#F9F9FB',  // Very light gray-blue
    '#FBF9F9',  // Very light pink-gray
    '#F9FBF9',  // Very light green-gray
    '#FEFEF9',  // Very light yellow-gray
    '#F9F9FE',  // Very light purple-gray
    '#FEF9FB',  // Very light rose-gray
  ];

  config.teamMembers.forEach((member, index) => {
    const colorIndex = index % subtleColors.length;
    const personRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(member)
      .setBackground(subtleColors[colorIndex])
      .setRanges([personRangeForFormatting])
      .build();
    personRules.push(personRule);
  });

  // ─────────────────────────────────────────────────────────────
  // DATE COLUMN (A5:A1000) - Highlight weekends
  // ─────────────────────────────────────────────────────────────
  const dateRangeForFormatting = standupSheet.getRange('A5:A1000');
  const weekendRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR(WEEKDAY(A5)=1, WEEKDAY(A5)=7)')
    .setBackground('#FCE5CD')
    .setRanges([dateRangeForFormatting])
    .build();

  // Apply all conditional formatting rules
  filteredRules.push(blockerNotEmptyRule);  // Use the formula-based blocker rule
  filteredRules.push(seninRule);
  filteredRules.push(rabuRule);
  filteredRules.push(jumatRule);
  filteredRules.push(...personRules);
  filteredRules.push(weekendRule);

  standupSheet.setConditionalFormatRules(filteredRules);

  // Freeze header rows
  standupSheet.setFrozenRows(4);

  Logger.log(`✅ ${sheetName} initialized`);
}

/**
 * Generate standup rows for a specific date
 * @param {string} projectName - 'SIPGN' or 'INADigital/Internal'
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

  // Validate that date is a standup day (Monday/Wednesday/Friday)
  if (!['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
    throw new Error(`Cannot generate rows for ${dayName}. Standup only on Monday, Wednesday, Friday.`);
  }

  // Convert to Indonesian day name
  const dayNameIndo = {
    'Monday': 'Senin',
    'Wednesday': 'Rabu',
    'Friday': 'Jumat'
  }[dayName];

  // Check if rows already exist for this date
  if (skipIfExists) {
    const lastRow = standupSheet.getLastRow();
    if (lastRow > 4) {
      const existingDates = standupSheet.getRange(5, 1, lastRow - 4, 1).getValues();
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

  // Prepare rows data (3 rows per person)
  const rowsData = [];

  config.teamMembers.forEach(person => {
    // Generate 3 empty rows per person
    for (let i = 0; i < 3; i++) {
      rowsData.push([
        dateStr,        // Date
        dayNameIndo,    // Day (Senin/Rabu/Jumat)
        person,         // Person
        '',             // Done Since Last (empty for user to fill)
        '',             // Plan Next 2 Days (empty for user to fill)
        '',             // Blockers (empty for user to fill)
        ''              // Notes (empty for user to fill)
      ]);
    }
  });

  // Insert rows at the end
  const lastRow = standupSheet.getLastRow();
  const startRow = lastRow + 1;

  standupSheet.getRange(startRow, 1, rowsData.length, 7).setValues(rowsData);

  // Apply formatting
  const dataRange = standupSheet.getRange(startRow, 1, rowsData.length, 7);

  // Borders
  dataRange.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Center align Date, Day, Person columns
  standupSheet.getRange(startRow, 1, rowsData.length, 1).setHorizontalAlignment('center'); // Date
  standupSheet.getRange(startRow, 2, rowsData.length, 1).setHorizontalAlignment('center'); // Day
  standupSheet.getRange(startRow, 3, rowsData.length, 1).setHorizontalAlignment('center'); // Person

  // Wrap text for Done Since Last, Plan Next 2 Days, Blockers, and Notes
  standupSheet.getRange(startRow, 4, rowsData.length, 1).setWrap(true); // Done Since Last
  standupSheet.getRange(startRow, 5, rowsData.length, 1).setWrap(true); // Plan Next 2 Days
  standupSheet.getRange(startRow, 6, rowsData.length, 1).setWrap(true); // Blockers
  standupSheet.getRange(startRow, 7, rowsData.length, 1).setWrap(true); // Notes

  Logger.log(`✅ Generated ${rowsData.length} rows for ${dateStr} (${dayNameIndo}) in ${sheetName}`);

  return rowsData.length;
}

/**
 * Get standup data for a specific date
 * @param {string} projectName - 'SIPGN' or 'INADigital/Internal'
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
  if (lastRow <= 4) {
    Logger.log(`⚠️ No data in ${sheetName}`);
    return [];
  }

  // Read all data (7 columns now)
  const allData = standupSheet.getRange(5, 1, lastRow - 4, 7).getValues();

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
        doneSinceLast: row[3],
        planNext2Days: row[4],
        blockers: row[5],
        notes: row[6]
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
        plan: [],
        blockers: []
      };
    }

    const doneSinceLast = String(row.doneSinceLast || '').trim();
    const planNext2Days = String(row.planNext2Days || '').trim();
    const blockers = String(row.blockers || '').trim();

    // Collect done items
    if (doneSinceLast) {
      summaryMap[person].done.push(doneSinceLast);
    }

    // Collect plan items
    if (planNext2Days) {
      summaryMap[person].plan.push(planNext2Days);
    }

    // Collect blockers
    if (blockers) {
      summaryMap[person].blockers.push(blockers);
    }
  });

  // Convert map to array
  const summaryArray = Object.values(summaryMap);

  return summaryArray;
}
