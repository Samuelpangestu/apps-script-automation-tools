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

  ws.getRange(22,1,1,23).merge()
      .setValue('▶ Run refreshDashboard() untuk mengisi data VAPT')
      .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(21);
}

function initVAPTHeaders_(ws) {
  // Set column widths
  ws.setColumnWidth(1, 80);   // Type
  ws.setColumnWidth(2, 200);  // Aplikasi
  ws.setColumnWidth(3, 120);  // PIC VAPT
  ws.setColumnWidth(4, 100);  // VAPT Status
  ws.setColumnWidth(5, 90);   // Blocker (NEW)
  for (let c=6; c<=23; c++) ws.setColumnWidth(c, 65);

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
  ws.getRange(1,1,1,23).merge().setValue('Last refreshed: —')
      .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  // Row 2 — title
  h_(2,1,1,23,'🔒 VAPT BLOCKER TRACKING  |  Medium-Critical Open Findings','#BF360C','#FFFFFF');
  ws.getRange(2,1).setFontSize(13);
  ws.setRowHeight(2,30);

  // Row 3-20 — SUMMARY SECTION
  ws.getRange(3,1,1,23).merge().setValue('SUMMARY METRICS')
      .setBackground('#FF6F00').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(3,24);

  // Summary layout (will be populated by writeVAPT)
  const summaryLabels = [
    ['Total Applications:', '—'],
    ['', ''],
    ['🚨 VAPT BLOCKER:', '—'],
    ['   (Medium-Critical Open)', ''],
    ['', ''],
    ['BLOCKER BREAKDOWN:', ''],
    ['  Critical Open:', '—'],
    ['  High Open:', '—'],
    ['  Medium Open:', '—'],
    ['', ''],
    ['OTHER OPEN FINDINGS:', ''],
    ['  Low Open:', '—'],
    ['  Info Open:', '—'],
    ['', ''],
    ['CLOSED FINDINGS:', '—'],
    ['', ''],
    ['VAPT STATUS:', ''],
    ['  Done:', '—'],
    ['  In Progress:', '—']
  ];

  summaryLabels.forEach((row, i) => {
    const r = 4 + i;
    ws.getRange(r,1).setValue(row[0]).setFontWeight('bold').setBackground('#FFEBEE');
    ws.getRange(r,2).setValue(row[1]).setBackground('#FFFFFF');
  });

  // Row 21 — TABLE SECTION HEADER
  h_(21,1, 1,4, 'VAPT INFO',              '#263238');
  h_(21,5, 1,1, 'BLOCKER',                '#B71C1C');  // NEW: Red for blocker
  h_(21,6, 1,5, 'OPEN FINDINGS',          '#D32F2F');
  h_(21,11,1,5, 'CLOSED',                 '#388E3C');
  h_(21,16,1,5, 'READY TO RETEST',        '#FF6F00');
  h_(21,21,1,3, 'STATUS',                 '#37474F');
  ws.setRowHeight(21,22);

  // Row 22 — column headers
  const headers = [
    'Type','Aplikasi','PIC VAPT','Status',
    'Blocker',  // NEW: Blocker column (Med+High+Crit Open)
    'Crit','High','Med','Low','Info',  // Open
    'Crit','High','Med','Low','Info',  // Closed
    'Crit','High','Med','Low','Info',  // Ready to Retest
    'Prod','Report','Last Updated'
  ];
  headers.forEach((lbl,i) => h_(22,i+1,1,1,lbl,'#1565C0'));

  // Add notes
  ws.getRange(22,1).setNote('Type\n\nAd Hoc = Development/Improvement VAPT\nRegular = Quarterly/Routine VAPT');
  ws.getRange(22,4).setNote('VAPT Status\n\nDone, In Progress, Not Started, Todo');
  ws.getRange(22,5).setNote('VAPT BLOCKER\n\n🚨 Medium + High + Critical findings yang masih OPEN\n\nTarget: 0 blocker\n\nPriority tinggi untuk diperbaiki!');
  ws.getRange(22,6).setNote('Open Findings\n\nActive findings that need to be fixed\nBreakdown by severity');
  ws.getRange(22,11).setNote('Closed\n\nFindings that have been verified and closed');
  ws.getRange(22,16).setNote('Ready to Retest\n\nFindings waiting for retest after fix');
  ws.getRange(22,21).setNote('Production Status\n\nTRUE = In Production\nFALSE = Not yet in Production');

  ws.setRowHeight(22,26);
  ws.setFrozenRows(22);
}

// ═══════════════════════════════════════════════════════════════════════
// WRITE VAPT DATA
// ═══════════════════════════════════════════════════════════════════════

function writeVAPT(ss, vaptData) {
  let ws = ss.getSheetByName('VAPT');
  if (!ws) { buildVAPT(ss); ws = ss.getSheetByName('VAPT'); }

  initVAPTHeaders_(ws);

  // Clear ALL data rows
  const lastRow = ws.getMaxRows();
  if (lastRow>=23) ws.getRange(23,1,lastRow-22,23).clearContent().clearFormat();

  // Update summary section (rows 4-21)
  updateVAPTSummary_(ws, vaptData.summary);

  // Calculate blocker for each row and sort by blocker descending
  const now = new Date();
  let data = vaptData.table.map(row => {
    // Calculate blocker: Medium + High + Critical yang Open
    const blocker = (row.open.medium || 0) + (row.open.high || 0) + (row.open.critical || 0);
    return {...row, blocker: blocker};
  });

  // Sort by blocker descending (apps with most blockers first)
  data.sort((a, b) => b.blocker - a.blocker);

  data.forEach((row, i) => {
    const r = 23 + i;
    const bg = i%2===0 ? '#FFF8E1' : '#FFFFFF';

    function cell(col,val,fmt){
      const c=ws.getRange(r,col).setValue(val==null?'':val).setBackground(bg)
          .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
      if(fmt)c.setNumberFormat(fmt);
      return c;
    }

    // VAPT Info (col 1-4)
    cell(1, row.type);  // Ad Hoc or Regular
    ws.getRange(r,2).setValue(row.aplikasi).setBackground(bg).setHorizontalAlignment('left').setFontFamily('Arial').setFontSize(9);
    cell(3, row.picVapt);
    cell(4, row.status);

    // Blocker (col 5) - NEW: Medium + High + Critical Open
    const blockerCell = ws.getRange(r,5);
    blockerCell.setValue(row.blocker).setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold')
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    // Apply color: Red if > 0, Green if 0
    if (row.blocker > 0) {
      blockerCell.setBackground('#FFCDD2').setFontColor('#C62828');
    } else {
      blockerCell.setBackground('#C8E6C9').setFontColor('#2E7D32');
    }

    // Open (col 6-10)
    cell(6, row.open.critical);
    cell(7, row.open.high);
    cell(8, row.open.medium);
    cell(9, row.open.low);
    cell(10, row.open.info);

    // Closed (col 11-15)
    cell(11, row.closed.critical);
    cell(12, row.closed.high);
    cell(13, row.closed.medium);
    cell(14, row.closed.low);
    cell(15, row.closed.info);

    // Ready to Retest (col 16-20)
    cell(16, row.readyToRetest.critical);
    cell(17, row.readyToRetest.high);
    cell(18, row.readyToRetest.medium);
    cell(19, row.readyToRetest.low);
    cell(20, row.readyToRetest.info);

    // Status (col 21-23)
    cell(21, row.prod ? 'TRUE' : 'FALSE');
    ws.getRange(r,22).setValue(row.report || '').setBackground(bg).setHorizontalAlignment('left').setFontFamily('Arial').setFontSize(8);
    cell(23, Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'));

    ws.setRowHeight(r, 22);
  });

  // Conditional formatting for severity columns
  applyVAPTConditionalFormatting_(ws, 23, data.length);

  // Update last refresh timestamp
  ws.getRange(1,1).setValue('Last refreshed: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss'));

  Logger.log('✅ VAPT tab updated: ' + data.length + ' applications');
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Update summary section with calculated metrics
 */
function updateVAPTSummary_(ws, summary) {
  ws.getRange(4,2).setValue(summary.totalApps || 0);

  // VAPT BLOCKER (row 6) - HIGHLIGHT
  const blockerValue = summary.blocker || 0;
  const blockerCell = ws.getRange(6,2);
  blockerCell.setValue(blockerValue).setFontWeight('bold').setFontSize(14);
  if (blockerValue > 0) {
    blockerCell.setBackground('#FFCDD2').setFontColor('#C62828');  // Red alert
  } else {
    blockerCell.setBackground('#C8E6C9').setFontColor('#2E7D32');  // Green = good
  }

  // Blocker Breakdown
  ws.getRange(9,2).setValue(summary.blockerBreakdown.critical || 0)
      .setBackground('#FFCDD2').setFontWeight('bold');  // Critical Open
  ws.getRange(10,2).setValue(summary.blockerBreakdown.high || 0)
      .setBackground('#FFCDD2').setFontWeight('bold');  // High Open
  ws.getRange(11,2).setValue(summary.blockerBreakdown.medium || 0)
      .setBackground('#FFE0B2').setFontWeight('bold');  // Medium Open

  // Other Open Findings
  ws.getRange(14,2).setValue(summary.otherOpen.low || 0)
      .setBackground('#C8E6C9').setFontWeight('bold');  // Low Open
  ws.getRange(15,2).setValue(summary.otherOpen.info || 0)
      .setBackground('#C8E6C9').setFontWeight('bold');  // Info Open

  // Closed Findings
  ws.getRange(17,2).setValue(summary.totalClosed || 0);

  // VAPT Status
  ws.getRange(20,2).setValue(summary.byVaptStatus.done || 0);
  ws.getRange(21,2).setValue(summary.byVaptStatus.inProgress || 0);
}

/**
 * Apply conditional formatting for severity columns
 */
function applyVAPTConditionalFormatting_(ws, startRow, dataLength) {
  if (dataLength === 0) return;

  const endRow = startRow + dataLength - 1;

  // Critical columns: 6 (Open), 11 (Closed), 16 (Ready to Retest) - Red if > 0
  [6, 11, 16].forEach(col => {
    const range = ws.getRange(startRow, col, dataLength, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#FFCDD2')
        .setFontColor('#C62828')
        .setRanges([range])
        .build();
    ws.setConditionalFormatRules([...ws.getConditionalFormatRules(), rule]);
  });

  // High columns: 7 (Open), 12 (Closed), 17 (Ready to Retest) - Orange if > 0
  [7, 12, 17].forEach(col => {
    const range = ws.getRange(startRow, col, dataLength, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#FFE0B2')
        .setFontColor('#E65100')
        .setRanges([range])
        .build();
    ws.setConditionalFormatRules([...ws.getConditionalFormatRules(), rule]);
  });
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

  historyRows.forEach(row => ws.appendRow(row));

  Logger.log('✅ VAPT History appended: ' + historyRows.length + ' rows');
}

/**
 * Create history row from summary data
 */
function createHistoryRow_(timestamp, type, summary) {
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
    summary.byStatus.open || 0,
    summary.byStatus.closed || 0,
    summary.totalApps || 0,
    summary.byVaptStatus.done || 0,
    summary.byVaptStatus.inProgress || 0,
    summary.prodCount || 0
  ];
}
