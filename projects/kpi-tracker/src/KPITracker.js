/**
 * KPITracker.js — KPI Tracking per Period
 * ═══════════════════════════════════════════════════════════════════════
 * Create and manage KPI tracking sheets per period (Sprint/Month/Quarter)
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Create KPI Tracker sheet for a specific period
 * @param {string} period - Period name (e.g., "Sprint 24", "Q1 2026", "Jan 2026")
 * @param {string} periodType - Type: "sprint", "monthly", "quarterly"
 */
function createKPITrackerPeriod(period, periodType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'KPI - ' + period;

  // Check if sheet already exists
  let tracker = ss.getSheetByName(sheetName);
  if (tracker) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Sheet Already Exists',
      'Sheet "' + sheetName + '" sudah ada.\n\nOverwrite?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return;
    }
    ss.deleteSheet(tracker);
  }

  // Create new sheet
  tracker = ss.insertSheet(sheetName);

  // Set column widths
  tracker.setColumnWidth(1, 50);   // A: No
  tracker.setColumnWidth(2, 200);  // B: Name
  tracker.setColumnWidth(3, 220);  // C: Role
  tracker.setColumnWidth(4, 250);  // D: KPI Name
  tracker.setColumnWidth(5, 100);  // E: Target
  tracker.setColumnWidth(6, 100);  // F: Actual
  tracker.setColumnWidth(7, 100);  // G: Achievement %
  tracker.setColumnWidth(8, 80);   // H: Status
  tracker.setColumnWidth(9, 300);  // I: Notes

  // Header
  tracker.getRange('A1:I1').merge()
    .setValue('📊 KPI TRACKER — ' + period)
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  tracker.setRowHeight(1, 40);

  tracker.getRange('A2:I2').merge()
    .setValue('Period: ' + period + ' | Type: ' + periodType.toUpperCase())
    .setFontSize(10)
    .setFontStyle('italic')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');

  // Instructions
  tracker.getRange('A3:I3').merge()
    .setValue('💡 Fill "Actual" column with real values. Achievement % and Status will auto-calculate.')
    .setBackground('#fff9c4')
    .setFontStyle('italic')
    .setWrap(true);

  // Table headers
  const headers = [
    ['No', 'Name', 'Role', 'KPI Name', 'Target', 'Actual', 'Achievement %', 'Status', 'Notes']
  ];

  tracker.getRange(5, 1, 1, 9).setValues(headers)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Get active team members
  const members = getTeamMembers();  // From TeamConfig.js

  // Populate data rows
  let rowNum = 6;
  let counter = 1;

  members.forEach(member => {
    const kpis = getKPIsForRole(member.role);  // From KPIDefinition.js

    kpis.forEach(kpi => {
      // Skip KPIs that don't match the period type
      const frequency = kpi.frequency.toLowerCase();
      if (periodType === 'sprint' && !frequency.includes('sprint')) {
        return; // Skip non-sprint KPIs for sprint tracking
      }
      if (periodType === 'monthly' && !frequency.includes('bulan')) {
        return; // Skip non-monthly KPIs
      }
      if (periodType === 'quarterly' && !frequency.includes('kuartal')) {
        return; // Skip non-quarterly KPIs
      }

      const row = [
        counter,
        member.name,
        member.role,
        kpi.name,
        kpi.target,
        '', // Actual - to be filled by user
        '', // Achievement % - formula
        '', // Status - formula
        ''  // Notes
      ];

      tracker.getRange(rowNum, 1, 1, 9).setValues([row]);

      // Format
      const bg = counter % 2 === 0 ? '#ffffff' : '#f8f9fa';
      tracker.getRange(rowNum, 1, 1, 9)
        .setBackground(bg)
        .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
        .setVerticalAlignment('middle');

      // Center No column
      tracker.getRange(rowNum, 1).setHorizontalAlignment('center');

      // Bold target
      tracker.getRange(rowNum, 5).setFontWeight('bold');

      // Achievement % formula (column G)
      const actualCell = 'F' + rowNum;
      const targetCell = 'E' + rowNum;

      // Parse target to get numeric value and operator
      // Examples: "≥ 80%", "≤ 10%", "≥ 3.5/5"
      const achievementFormula =
        '=IF(ISBLANK(' + actualCell + '),"",IF(ISNUMBER(' + actualCell + '),ROUND((' + actualCell + '/' +
        'IF(REGEXMATCH(' + targetCell + ',"%"),VALUE(REGEXEXTRACT(' + targetCell + ',"\\d+\\.?\\d*")),' +
        'IF(REGEXMATCH(' + targetCell + ',"/"),VALUE(LEFT(REGEXEXTRACT(' + targetCell + ',"\\d+\\.?\\d*/"),FIND("/",REGEXEXTRACT(' + targetCell + ',"\\d+\\.?\\d*/"))-1)),' +
        'VALUE(REGEXEXTRACT(' + targetCell + ',"\\d+\\.?\\d*"))))) * 100,1) & "%","N/A"))';

      tracker.getRange(rowNum, 7).setFormula(achievementFormula);

      // Status formula (column H) - based on target operator
      const statusFormula =
        '=IF(ISBLANK(' + actualCell + '),"⚪ Pending",' +
        'IF(REGEXMATCH(' + targetCell + ',"≥|>="),' +
        '  IF(' + actualCell + '>=' + 'VALUE(REGEXEXTRACT(' + targetCell + ',"\\d+\\.?\\d*")),"✅ Met","❌ Not Met"),' +
        'IF(REGEXMATCH(' + targetCell + ',"≤|<="),' +
        '  IF(' + actualCell + '<=' + 'VALUE(REGEXEXTRACT(' + targetCell + ',"\\d+\\.?\\d*")),"✅ Met","❌ Not Met"),' +
        '"⚠️ Review")))';

      tracker.getRange(rowNum, 8).setFormula(statusFormula);

      // Number format for Actual column (column F)
      tracker.getRange(rowNum, 6).setNumberFormat('0.00');

      rowNum++;
      counter++;
    });
  });

  // Add summary section
  const summaryRow = rowNum + 2;
  tracker.getRange(summaryRow, 1, 1, 9).merge()
    .setValue('📈 SUMMARY')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#f1f3f4')
    .setHorizontalAlignment('left');

  // Summary stats
  const dataRange = '6:' + (rowNum - 1);
  tracker.getRange(summaryRow + 1, 1, 1, 2).merge().setValue('Total KPIs Tracked:');
  tracker.getRange(summaryRow + 1, 3).setFormula('=COUNTA(H' + dataRange + ')');

  tracker.getRange(summaryRow + 2, 1, 1, 2).merge().setValue('✅ Met Target:');
  tracker.getRange(summaryRow + 2, 3).setFormula('=COUNTIF(H' + dataRange + ',"✅ Met")');

  tracker.getRange(summaryRow + 3, 1, 1, 2).merge().setValue('❌ Not Met:');
  tracker.getRange(summaryRow + 3, 3).setFormula('=COUNTIF(H' + dataRange + ',"❌ Not Met")');

  tracker.getRange(summaryRow + 4, 1, 1, 2).merge().setValue('⚪ Pending:');
  tracker.getRange(summaryRow + 4, 3).setFormula('=COUNTIF(H' + dataRange + ',"⚪ Pending")');

  tracker.getRange(summaryRow + 5, 1, 1, 2).merge().setValue('Success Rate:');
  tracker.getRange(summaryRow + 5, 3).setFormula(
    '=IF(COUNTIF(H' + dataRange + ',"✅ Met")+COUNTIF(H' + dataRange + ',"❌ Not Met")=0,"-",' +
    'ROUND(COUNTIF(H' + dataRange + ',"✅ Met")/(COUNTIF(H' + dataRange + ',"✅ Met")+COUNTIF(H' + dataRange + ',"❌ Not Met"))*100,1)&"%")'
  );

  // Format summary
  tracker.getRange(summaryRow + 1, 1, 5, 2).setFontWeight('bold');
  tracker.getRange(summaryRow + 1, 3, 5, 1)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#e8f5e9');

  // Freeze header
  tracker.setFrozenRows(5);
  tracker.setFrozenColumns(1);

  Logger.log('✅ KPI Tracker created for period: ' + period);
  return tracker;
}

/**
 * Menu function to create KPI Tracker with user input
 */
function menuCreateKPITracker() {
  const ui = SpreadsheetApp.getUi();

  // Get period name
  const periodResponse = ui.prompt(
    'Create KPI Tracker',
    'Enter period name:\n(e.g., "Sprint 24", "Jan 2026", "Q1 2026")',
    ui.ButtonSet.OK_CANCEL
  );

  if (periodResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const period = periodResponse.getResponseText().trim();
  if (!period) {
    ui.alert('Period name cannot be empty.');
    return;
  }

  // Get period type
  const typeResponse = ui.alert(
    'Select Period Type',
    'Period: ' + period + '\n\nSelect tracking frequency:',
    ui.ButtonSet.YES_NO_CANCEL
  );

  let periodType;
  if (typeResponse === ui.Button.YES) {
    periodType = 'sprint';
  } else if (typeResponse === ui.Button.NO) {
    // Ask monthly or quarterly
    const monthlyOrQuarterly = ui.alert(
      'Monthly or Quarterly?',
      'Choose tracking type:',
      ui.ButtonSet.YES_NO
    );
    periodType = monthlyOrQuarterly === ui.Button.YES ? 'monthly' : 'quarterly';
  } else {
    return;
  }

  try {
    createKPITrackerPeriod(period, periodType);
    ui.alert(
      'Success! ✅',
      'KPI Tracker sheet created for period: ' + period + '\n' +
      'Type: ' + periodType.toUpperCase() + '\n\n' +
      'Fill the "Actual" column to track KPI achievement.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log('Error creating KPI Tracker: ' + e.message);
    ui.alert('Error', 'Failed to create KPI Tracker:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Get all KPI tracker sheets
 * @returns {Array} Array of tracker sheet objects
 */
function getAllKPITrackers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const trackers = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith('KPI - ')) {
      const period = name.substring(6); // Remove "KPI - " prefix
      trackers.push({
        sheet: sheet,
        name: name,
        period: period
      });
    }
  });

  return trackers;
}

/**
 * Calculate overall KPI achievement for a specific tracker sheet
 * @param {Sheet} sheet - Tracker sheet
 * @returns {Object} Achievement summary
 */
function calculateTrackerAchievement(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 6) {
    return null;
  }

  const statusRange = sheet.getRange(6, 8, lastRow - 5, 1).getValues();

  let total = 0;
  let met = 0;
  let notMet = 0;
  let pending = 0;

  statusRange.forEach(row => {
    const status = row[0].toString();
    if (status.includes('Met')) {
      total++;
      met++;
    } else if (status.includes('Not Met')) {
      total++;
      notMet++;
    } else if (status.includes('Pending')) {
      pending++;
    }
  });

  const successRate = total > 0 ? Math.round((met / total) * 100) : 0;

  return {
    total: total + pending,
    met: met,
    notMet: notMet,
    pending: pending,
    successRate: successRate
  };
}
