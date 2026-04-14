/**
 * VAPTTab.js - VAPT Findings Tracking Dashboard
 *
 * Track VAPT (Vulnerability Assessment & Penetration Testing) findings
 * Combines data from 2 sources: Ad Hoc VAPT and Regular VAPT
 *
 * Features:
 * - Summary metrics (total apps, total findings, by severity, by status)
 * - Combined table from Ad Hoc + Regular VAPT
 * - Findings by severity: Critical, High, Medium, Low, Info
 * - Findings by status: Ready to Retest, Open, Closed
 * - Production status tracking
 * - VAPT History for trendline analysis
 */

// ═══════════════════════════════════════════════════════════════════════
// BUILD VAPT TAB
// ═══════════════════════════════════════════════════════════════════════

function buildVAPT(ss) {
  const ws = ss.insertSheet('VAPT', 2);  // Insert after Bugs
  ws.setTabColor('#EF6C00');  // Orange color for security
  ws.clear();
  initVAPTHeaders_(ws);

  ws.getRange(10,1,1,8).merge()
      .setValue('▶ Run refreshDashboard() untuk mengisi data VAPT per-project')
      .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(9);
}

function initVAPTHeaders_(ws) {
  // Set column widths (8 columns total - added Project)
  ws.setColumnWidth(1, 120);  // Project
  ws.setColumnWidth(2, 200);  // Aplikasi
  ws.setColumnWidth(3, 90);   // Blocker
  ws.setColumnWidth(4, 70);   // Critical
  ws.setColumnWidth(5, 70);   // High
  ws.setColumnWidth(6, 70);   // Medium
  ws.setColumnWidth(7, 70);   // Low
  ws.setColumnWidth(8, 70);   // Info

  function h_(r,c,rCnt,cCnt,txt,bg,fg) {
    fg = fg || '#FFFFFF';
    const rng = cCnt>1 || rCnt>1 ? ws.getRange(r,c,rCnt,cCnt).merge() : ws.getRange(r,c);
    return rng.setValue(txt)
        .setBackground(bg).setFontColor(fg).setFontWeight('bold')
        .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#FFFFFF',SpreadsheetApp.BorderStyle.SOLID);
  }

  // Row 1 — last refresh
  ws.getRange(1,1,1,8).merge().setValue('Last refreshed: —')
      .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  // Row 2 — title with link
  ws.getRange(2,1,1,8).merge().setValue('🔒 PER-PROJECT VAPT BLOCKER TRACKING')
      .setBackground('#BF360C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(2,30);

  // Row 3 — subtitle + link to external
  ws.getRange(3,1,1,8).merge()
      .setValue('📋 Each project has its own VAPT spreadsheet')
      .setBackground('#FFF3E0').setFontColor('#E65100')
      .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(3,20);

  // Row 4-8 — SIMPLE SUMMARY
  ws.getRange(4,1,1,8).merge().setValue('🚨 SUMMARY (All Projects)')
      .setBackground('#FF6F00').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(4,22);

  const summaryLabels = [
    ['Total Blocker:', '—', 'Total Projects:', '—'],
    ['Apps with Blocker:', '—', 'Total Apps:', '—']
  ];

  summaryLabels.forEach((row, i) => {
    const r = 5 + i;
    ws.getRange(r,1).setValue(row[0]).setFontWeight('bold').setBackground('#FFEBEE').setHorizontalAlignment('right');
    ws.getRange(r,2).setValue(row[1]).setBackground('#FFFFFF').setFontSize(11).setFontWeight('bold');
    ws.getRange(r,4).setValue(row[2]).setFontWeight('bold').setBackground('#FFEBEE').setHorizontalAlignment('right');
    ws.getRange(r,5).setValue(row[3]).setBackground('#FFFFFF').setFontSize(10);
  });

  // Row 7 — blank
  ws.setRowHeight(7,8);

  // Row 8 — Target
  ws.getRange(8,1,1,8).merge().setValue('🎯 Target: 0 blocker di semua aplikasi!')
      .setBackground('#E8F5E9').setFontColor('#2E7D32').setFontWeight('bold')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(8,20);

  // Row 9 — blank
  ws.setRowHeight(9,8);
}

// ═══════════════════════════════════════════════════════════════════════
// WRITE VAPT DATA
// ═══════════════════════════════════════════════════════════════════════

function writeVAPT(ss, vaptData) {
  let ws = ss.getSheetByName('VAPT');
  if (!ws) { buildVAPT(ss); ws = ss.getSheetByName('VAPT'); }

  initVAPTHeaders_(ws);

  // Clear ALL data rows (from row 10 onwards)
  const lastRow = ws.getMaxRows();
  if (lastRow >= 10) ws.getRange(10, 1, lastRow - 9, 7).clearContent().clearFormat();

  // Update link to external VAPT spreadsheet
  const vaptSpreadsheetId = '17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk';
  const externalLink = 'https://docs.google.com/spreadsheets/d/' + vaptSpreadsheetId + '/edit';
  ws.getRange(3, 1).setFormula('=HYPERLINK("' + externalLink + '", "📋 View External VAPT Spreadsheet")');

  // Calculate blocker for each row
  const now = new Date();
  let data = vaptData.table.map(row => {
    // Calculate blocker: Medium + High + Critical yang Open
    const blocker = (row.open.medium || 0) + (row.open.high || 0) + (row.open.critical || 0);
    return {...row, blocker: blocker};
  });

  // Separate Ad Hoc and Regular, then sort each by blocker descending
  const adHocData = data.filter(row => row.type === 'Ad Hoc').sort((a, b) => b.blocker - a.blocker);
  const regularData = data.filter(row => row.type === 'Regular').sort((a, b) => b.blocker - a.blocker);

  // Update summary
  updateVAPTSummary_(ws, vaptData.summary, vaptData.adHocSummary, vaptData.regularSummary, data);

  let currentRow = 10;

  // ═══════════════════════════════════════════════════════════════════════
  // AD HOC VAPT SECTION
  // ═══════════════════════════════════════════════════════════════════════
  if (adHocData.length > 0) {
    // Section header
    ws.getRange(currentRow, 1, 1, 7).merge()
        .setValue('═══ AD HOC VAPT (Development/Improvement) ═══')
        .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.setRowHeight(currentRow, 24);
    currentRow++;

    // Table headers (7 columns - removed Status)
    const headers = ['Aplikasi', 'Blocker', 'Critical', 'High', 'Medium', 'Low', 'Info'];
    headers.forEach((lbl, i) => {
      ws.getRange(currentRow, i + 1)
          .setValue(lbl)
          .setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold')
          .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center')
          .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
    });
    ws.setRowHeight(currentRow, 22);
    currentRow++;

    // Data rows
    adHocData.forEach((row, i) => {
      writeVAPTRow_(ws, currentRow, row, i);
      currentRow++;
    });

    currentRow++; // Blank row after section
  }

  // ═══════════════════════════════════════════════════════════════════════
  // REGULAR VAPT SECTION
  // ═══════════════════════════════════════════════════════════════════════
  if (regularData.length > 0) {
    // Section header
    ws.getRange(currentRow, 1, 1, 7).merge()
        .setValue('═══ REGULAR VAPT (Quarterly/Routine) ═══')
        .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.setRowHeight(currentRow, 24);
    currentRow++;

    // Table headers (7 columns - removed Status)
    const headers = ['Aplikasi', 'Blocker', 'Critical', 'High', 'Medium', 'Low', 'Info'];
    headers.forEach((lbl, i) => {
      ws.getRange(currentRow, i + 1)
          .setValue(lbl)
          .setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold')
          .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center')
          .setBorder(true, true, true, true, false, false, '#FFFFFF', SpreadsheetApp.BorderStyle.SOLID);
    });
    ws.setRowHeight(currentRow, 22);
    currentRow++;

    // Data rows
    regularData.forEach((row, i) => {
      writeVAPTRow_(ws, currentRow, row, i);
      currentRow++;
    });
  }

  // Update last refresh timestamp
  ws.getRange(1, 1).setValue('Last refreshed: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss'));

  Logger.log('✅ VAPT tab updated: ' + data.length + ' applications (Ad Hoc: ' + adHocData.length + ', Regular: ' + regularData.length + ')');
}

/**
 * Write a single VAPT data row (7 columns - removed Status)
 */
function writeVAPTRow_(ws, row, data, index) {
  const bg = index % 2 === 0 ? '#FFF8E1' : '#FFFFFF';

  function cell(col, val) {
    return ws.getRange(row, col).setValue(val == null ? '' : val).setBackground(bg)
        .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Col 1: Aplikasi (left-aligned)
  ws.getRange(row, 1).setValue(data.aplikasi).setBackground(bg).setHorizontalAlignment('left')
      .setFontFamily('Arial').setFontSize(9)
      .setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);

  // Col 2: Blocker (colored: red if > 0, green if 0)
  const blockerCell = ws.getRange(row, 2);
  blockerCell.setValue(data.blocker).setBackground(bg).setFontFamily('Arial').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  if (data.blocker > 0) {
    blockerCell.setBackground('#FFCDD2').setFontColor('#C62828');  // Red
  } else {
    blockerCell.setBackground('#C8E6C9').setFontColor('#2E7D32');  // Green
  }

  // Col 3-7: OPEN findings (Critical, High, Medium, Low, Info)
  cell(3, data.open.critical);
  cell(4, data.open.high);
  cell(5, data.open.medium);
  cell(6, data.open.low);
  cell(7, data.open.info);

  ws.setRowHeight(row, 22);
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Update summary section with calculated metrics
 */
function updateVAPTSummary_(ws, summary, adHocSummary, regularSummary, data) {
  // Row 5: Total Blocker & Ad Hoc Blocker
  const totalBlockerCell = ws.getRange(5, 2);
  totalBlockerCell.setValue(summary.blocker || 0).setFontWeight('bold').setFontSize(12);
  if (summary.blocker > 0) {
    totalBlockerCell.setBackground('#FFCDD2').setFontColor('#C62828');  // Red
  } else {
    totalBlockerCell.setBackground('#C8E6C9').setFontColor('#2E7D32');  // Green
  }

  ws.getRange(5, 5).setValue(adHocSummary.blocker || 0).setFontWeight('bold');

  // Row 6: Apps with Blocker & Regular Blocker
  const appsWithBlocker = data ? data.filter(app => app.blocker > 0).length : 0;
  ws.getRange(6, 2).setValue(appsWithBlocker).setFontWeight('bold');
  ws.getRange(6, 5).setValue(regularSummary.blocker || 0).setFontWeight('bold');
}

// ═══════════════════════════════════════════════════════════════════════
// BUILD VAPT HISTORY TAB
// ═══════════════════════════════════════════════════════════════════════

function buildVAPTHistory(ss) {
  const ws = ss.insertSheet('VAPT History');
  ws.setTabColor('#BF360C');
  ws.clear();

  const hdrs = [
    'Timestamp','Type',
    'Total Findings','Critical','High','Medium','Low','Info',
    'Ready to Retest','Open','Closed',
    'Apps Total','Apps Done','Apps In Progress',
    'Prod Count'
  ];

  ws.getRange(1,1,1,hdrs.length).merge().setValue('VAPT HISTORY  —  Trend Data (auto-appended setiap refresh)')
      .setBackground('#BF360C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold')
      .setBackground('#EF6C00').setFontColor('#FFFFFF');
  ws.setFrozenRows(2);
  ws.setColumnWidth(1,130);
  ws.setColumnWidth(2,90);
  for(let c=3;c<=hdrs.length;c++)ws.setColumnWidth(c,80);

  // Add notes
  ws.getRange(2,2).setNote('Type\n\nAd Hoc, Regular, or Combined');
  ws.getRange(2,3).setNote('Total Findings\n\nSum of all findings (all severities, all statuses)');
  ws.getRange(2,9).setNote('Ready to Retest\n\nFindings waiting for retest after fix');
  ws.getRange(2,10).setNote('Open\n\nActive findings needing action');
  ws.getRange(2,11).setNote('Closed\n\nVerified and closed findings');
}

/**
 * Append daily snapshot to VAPT History
 */
function appendVAPTHistory(ss, vaptData) {
  let ws = ss.getSheetByName('VAPT History');
  if (!ws) {
    buildVAPTHistory(ss);
    ws = ss.getSheetByName('VAPT History');
  }

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

  // Append 3 rows: Ad Hoc, Regular, Combined
  const historyRows = [
    createHistoryRow_(ts, 'Ad Hoc', vaptData.adHocSummary),
    createHistoryRow_(ts, 'Regular', vaptData.regularSummary),
    createHistoryRow_(ts, 'Combined', vaptData.summary)
  ];

  // Get last row and append using setValues (safer than appendRow with merged cells)
  // IMPORTANT: Ensure we never write to rows 1-2 (headers with merged cells)
  const lastRow = ws.getLastRow();
  const startRow = Math.max(lastRow + 1, 3);  // Always start at row 3 minimum (after headers)
  const numCols = historyRows[0].length;

  // Unmerge target range first to avoid merge conflicts
  const targetRange = ws.getRange(startRow, 1, historyRows.length, numCols);
  targetRange.breakApart();  // Unmerge any merged cells in target area
  targetRange.setValues(historyRows);

  Logger.log('✅ VAPT History appended: ' + historyRows.length + ' rows at row ' + startRow);
}

/**
 * Create history row from summary data
 */
function createHistoryRow_(timestamp, type, summary) {
  // Calculate blocker (Critical + High + Medium OPEN findings only, not Low/Info)
  // This matches the blocker definition used in dashboard
  const blocker = summary.blocker || 0;

  return [
    timestamp,
    type,
    summary.totalFindings || 0,
    summary.bySeverity.critical || 0,
    summary.bySeverity.high || 0,
    summary.bySeverity.medium || 0,
    summary.bySeverity.low || 0,
    summary.bySeverity.info || 0,
    summary.byStatus.readyToRetest || 0,
    blocker,  // Use blocker instead of summary.byStatus.open (which includes Low/Info)
    summary.byStatus.closed || 0,
    summary.totalApps || 0,
    summary.byVaptStatus.done || 0,
    summary.byVaptStatus.inProgress || 0,
    summary.prodCount || 0
  ];
}
