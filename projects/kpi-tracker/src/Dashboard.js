/**
 * Dashboard.js — KPI Dashboard & Overview
 * ═══════════════════════════════════════════════════════════════════════
 * Create summary dashboard showing KPI trends across all periods
 * ═══════════════════════════════════════════════════════════════════════
 */

const DASHBOARD_TAB_NAME = 'Dashboard';

/**
 * Create or update Dashboard tab
 */
function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashboard = ss.getSheetByName(DASHBOARD_TAB_NAME);

  if (!dashboard) {
    dashboard = ss.insertSheet(DASHBOARD_TAB_NAME, 0);
  }

  dashboard.clear();

  // Set column widths
  dashboard.setColumnWidth(1, 200);  // A: Period
  dashboard.setColumnWidth(2, 120);  // B: Total KPIs
  dashboard.setColumnWidth(3, 100);  // C: Met
  dashboard.setColumnWidth(4, 100);  // D: Not Met
  dashboard.setColumnWidth(5, 100);  // E: Pending
  dashboard.setColumnWidth(6, 120);  // F: Success Rate

  // Header
  dashboard.getRange('A1:F1').merge()
    .setValue('📊 KPI DASHBOARD — OVERVIEW')
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  dashboard.setRowHeight(1, 50);

  dashboard.getRange('A2:F2').merge()
    .setValue('Summary of all KPI tracking periods')
    .setFontSize(10)
    .setFontStyle('italic')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');

  // Quick Stats Section
  const statsRow = 4;
  dashboard.getRange(statsRow, 1, 1, 6).merge()
    .setValue('📈 QUICK STATS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f1f3f4')
    .setHorizontalAlignment('left');

  // Get all trackers
  const trackers = getAllKPITrackers();  // From KPITracker.js

  // Quick stats cards
  dashboard.getRange(statsRow + 1, 1).setValue('Total Periods Tracked:');
  dashboard.getRange(statsRow + 1, 2).setValue(trackers.length)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#e3f2fd');

  dashboard.getRange(statsRow + 1, 3).setValue('Active Team Members:');
  const activeMembers = getTeamMembers().length;  // From TeamConfig.js
  dashboard.getRange(statsRow + 1, 4).setValue(activeMembers)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#e8f5e9');

  dashboard.getRange(statsRow + 1, 5).setValue('Total KPIs Defined:');
  const allRoles = getAllRoles();  // From KPIDefinition.js
  let totalKPIs = 0;
  allRoles.forEach(role => {
    totalKPIs += getKPIsForRole(role).length;
  });
  dashboard.getRange(statsRow + 1, 6).setValue(totalKPIs)
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#fff3e0');

  // Periods Summary Section
  const tableRow = statsRow + 3;
  dashboard.getRange(tableRow, 1, 1, 6).merge()
    .setValue('📋 PERIODS SUMMARY')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f1f3f4')
    .setHorizontalAlignment('left');

  // Table headers
  const headers = [
    ['Period', 'Total KPIs', '✅ Met', '❌ Not Met', '⚪ Pending', 'Success Rate']
  ];

  dashboard.getRange(tableRow + 1, 1, 1, 6).setValues(headers)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Populate data
  let dataRow = tableRow + 2;

  if (trackers.length === 0) {
    dashboard.getRange(dataRow, 1, 1, 6).merge()
      .setValue('No KPI tracker periods found. Create one from Menu → KPI Tracker → Create Period Tracker')
      .setFontStyle('italic')
      .setHorizontalAlignment('center')
      .setBackground('#fff9c4');
  } else {
    trackers.forEach((tracker, idx) => {
      const achievement = calculateTrackerAchievement(tracker.sheet);  // From KPITracker.js

      if (achievement) {
        const row = [
          tracker.period,
          achievement.total,
          achievement.met,
          achievement.notMet,
          achievement.pending,
          achievement.successRate + '%'
        ];

        dashboard.getRange(dataRow, 1, 1, 6).setValues([row]);

        // Format
        const bg = idx % 2 === 0 ? '#ffffff' : '#f8f9fa';
        dashboard.getRange(dataRow, 1, 1, 6)
          .setBackground(bg)
          .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
          .setVerticalAlignment('middle');

        // Center numeric columns
        dashboard.getRange(dataRow, 2, 1, 5).setHorizontalAlignment('center');

        // Color code success rate
        const rate = achievement.successRate;
        let rateBg = '#c8e6c9'; // Green
        if (rate < 70) rateBg = '#ffcdd2'; // Red
        else if (rate < 85) rateBg = '#fff9c4'; // Yellow

        dashboard.getRange(dataRow, 6)
          .setBackground(rateBg)
          .setFontWeight('bold');

        dataRow++;
      }
    });

    // Overall summary
    const overallRow = dataRow + 1;
    dashboard.getRange(overallRow, 1).setValue('OVERALL:').setFontWeight('bold');

    dashboard.getRange(overallRow, 2).setFormula(
      '=SUM(B' + (tableRow + 2) + ':B' + (dataRow - 1) + ')'
    ).setFontWeight('bold').setHorizontalAlignment('center');

    dashboard.getRange(overallRow, 3).setFormula(
      '=SUM(C' + (tableRow + 2) + ':C' + (dataRow - 1) + ')'
    ).setFontWeight('bold').setHorizontalAlignment('center');

    dashboard.getRange(overallRow, 4).setFormula(
      '=SUM(D' + (tableRow + 2) + ':D' + (dataRow - 1) + ')'
    ).setFontWeight('bold').setHorizontalAlignment('center');

    dashboard.getRange(overallRow, 5).setFormula(
      '=SUM(E' + (tableRow + 2) + ':E' + (dataRow - 1) + ')'
    ).setFontWeight('bold').setHorizontalAlignment('center');

    dashboard.getRange(overallRow, 6).setFormula(
      '=IF(C' + overallRow + '+D' + overallRow + '=0,"-",ROUND(C' + overallRow + '/(C' + overallRow + '+D' + overallRow + ')*100,1)&"%")'
    ).setFontWeight('bold').setHorizontalAlignment('center');

    dashboard.getRange(overallRow, 1, 1, 6)
      .setBackground('#e8f5e9')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  // Team Performance Section
  const teamPerfRow = dataRow + 3;
  dashboard.getRange(teamPerfRow, 1, 1, 6).merge()
    .setValue('👥 TEAM MEMBERS')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#f1f3f4')
    .setHorizontalAlignment('left');

  // Team member list with roles
  const members = getTeamMembers();
  const teamHeaders = [['Name', 'Role', 'Status', 'Email', '', '']];
  dashboard.getRange(teamPerfRow + 1, 1, 1, 6).setValues(teamHeaders)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  let teamRow = teamPerfRow + 2;
  members.forEach((member, idx) => {
    const row = [member.name, member.role, member.status, member.email, '', ''];
    dashboard.getRange(teamRow, 1, 1, 6).setValues([row]);

    const bg = idx % 2 === 0 ? '#ffffff' : '#f8f9fa';
    dashboard.getRange(teamRow, 1, 1, 6)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    teamRow++;
  });

  // Tips section
  const tipsRow = teamRow + 2;
  dashboard.getRange(tipsRow, 1, 1, 6).merge()
    .setValue(
      '💡 TIPS:\n' +
      '• Create new period tracker: Menu → KPI Tracker → Create Period Tracker\n' +
      '• Update team members: Edit Config tab\n' +
      '• Refresh dashboard: Menu → Dashboard → Refresh Dashboard\n' +
      '• Colors: Green (≥85%), Yellow (70-84%), Red (<70%)'
    )
    .setBackground('#e3f2fd')
    .setFontStyle('italic')
    .setWrap(true)
    .setVerticalAlignment('top');
  dashboard.setRowHeight(tipsRow, 100);

  // Freeze header
  dashboard.setFrozenRows(3);

  Logger.log('✅ Dashboard created/updated');
  return dashboard;
}

/**
 * Menu function to refresh dashboard
 */
function menuRefreshDashboard() {
  try {
    createDashboard();
    SpreadsheetApp.getUi().alert(
      'Success! ✅',
      'Dashboard has been refreshed with latest data.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('Error refreshing dashboard: ' + e.message);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to refresh dashboard:\n' + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Get KPI performance by team member across all periods
 * @param {string} memberName - Team member name
 * @returns {Object} Performance summary
 */
function getMemberPerformance(memberName) {
  const trackers = getAllKPITrackers();
  let totalKPIs = 0;
  let metKPIs = 0;
  let notMetKPIs = 0;

  trackers.forEach(tracker => {
    const sheet = tracker.sheet;
    const lastRow = sheet.getLastRow();

    if (lastRow < 6) return;

    const data = sheet.getRange(6, 2, lastRow - 5, 7).getValues();

    data.forEach(row => {
      const name = row[0].toString().trim();
      const status = row[6].toString();

      if (name === memberName) {
        if (status.includes('Met')) {
          totalKPIs++;
          metKPIs++;
        } else if (status.includes('Not Met')) {
          totalKPIs++;
          notMetKPIs++;
        }
      }
    });
  });

  const successRate = totalKPIs > 0 ? Math.round((metKPIs / totalKPIs) * 100) : 0;

  return {
    name: memberName,
    total: totalKPIs,
    met: metKPIs,
    notMet: notMetKPIs,
    successRate: successRate
  };
}
