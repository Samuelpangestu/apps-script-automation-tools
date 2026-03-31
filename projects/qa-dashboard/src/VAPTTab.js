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

  ws.getRange(22,1,1,22).merge()
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
  for (let c=5; c<=22; c++) ws.setColumnWidth(c, 65);

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
  ws.getRange(1,1,1,22).merge().setValue('Last refreshed: —')
      .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  // Row 2 — title
  h_(2,1,1,22,'🔒 VAPT FINDINGS  |  SECURITY VULNERABILITY TRACKING','#BF360C','#FFFFFF');
  ws.getRange(2,1).setFontSize(13);
  ws.setRowHeight(2,30);

  // Row 3-20 — SUMMARY SECTION
  ws.getRange(3,1,1,22).merge().setValue('SUMMARY METRICS')
      .setBackground('#FF6F00').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(3,24);

  // Summary layout (will be populated by writeVAPT)
  const summaryLabels = [
    ['Total Applications:', '—'],
    ['Total Findings:', '—'],
    ['', ''],
    ['BY SEVERITY:', ''],
    ['  Critical:', '—'],
    ['  High:', '—'],
    ['  Medium:', '—'],
    ['  Low:', '—'],
    ['  Info:', '—'],
    ['', ''],
    ['BY STATUS:', ''],
    ['  Ready to Retest:', '—'],
    ['  Open:', '—'],
    ['  Closed:', '—'],
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
  h_(21,5, 1,5, 'READY TO RETEST',        '#FF6F00');
  h_(21,10,1,5, 'OPEN',                   '#D32F2F');
  h_(21,15,1,5, 'CLOSED',                 '#388E3C');
  h_(21,20,1,3, 'STATUS',                 '#37474F');
  ws.setRowHeight(21,22);

  // Row 22 — column headers
  const headers = [
    'Type','Aplikasi','PIC VAPT','Status',
    'Crit','High','Med','Low','Info',
    'Crit','High','Med','Low','Info',
    'Crit','High','Med','Low','Info',
    'Prod','Report','Last Updated'
  ];
  headers.forEach((lbl,i) => h_(22,i+1,1,1,lbl,'#1565C0'));

  // Add notes
  ws.getRange(22,1).setNote('Type\n\nAd Hoc = Development/Improvement VAPT\nRegular = Quarterly/Routine VAPT');
  ws.getRange(22,4).setNote('VAPT Status\n\nDone, In Progress, Not Started, Todo');
  ws.getRange(22,5).setNote('Ready to Retest\n\nFindings waiting for retest after fix');
  ws.getRange(22,10).setNote('Open\n\nActive findings that need to be fixed');
  ws.getRange(22,15).setNote('Closed\n\nFindings that have been verified and closed');
  ws.getRange(22,20).setNote('Production Status\n\nTRUE = In Production\nFALSE = Not yet in Production');

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
  if (lastRow>=23) ws.getRange(23,1,lastRow-22,22).clearContent().clearFormat();

  // Update summary section (rows 4-21)
  updateVAPTSummary_(ws, vaptData.summary);

  // Write table data
  const now = new Date();
  const data = vaptData.table;

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

    // Ready to Retest (col 5-9)
    cell(5, row.readyToRetest.critical);
    cell(6, row.readyToRetest.high);
    cell(7, row.readyToRetest.medium);
    cell(8, row.readyToRetest.low);
    cell(9, row.readyToRetest.info);

    // Open (col 10-14)
    cell(10, row.open.critical);
    cell(11, row.open.high);
    cell(12, row.open.medium);
    cell(13, row.open.low);
    cell(14, row.open.info);

    // Closed (col 15-19)
    cell(15, row.closed.critical);
    cell(16, row.closed.high);
    cell(17, row.closed.medium);
    cell(18, row.closed.low);
    cell(19, row.closed.info);

    // Status (col 20-22)
    cell(20, row.prod ? 'TRUE' : 'FALSE');
    ws.getRange(r,21).setValue(row.report || '').setBackground(bg).setHorizontalAlignment('left').setFontFamily('Arial').setFontSize(8);
    cell(22, Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'));

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
  ws.getRange(5,2).setValue(summary.totalFindings || 0);

  // By Severity
  ws.getRange(7,2).setValue(summary.bySeverity.critical || 0);
  ws.getRange(8,2).setValue(summary.bySeverity.high || 0);
  ws.getRange(9,2).setValue(summary.bySeverity.medium || 0);
  ws.getRange(10,2).setValue(summary.bySeverity.low || 0);
  ws.getRange(11,2).setValue(summary.bySeverity.info || 0);

  // By Status
  ws.getRange(14,2).setValue(summary.byStatus.readyToRetest || 0);
  ws.getRange(15,2).setValue(summary.byStatus.open || 0);
  ws.getRange(16,2).setValue(summary.byStatus.closed || 0);

  // VAPT Status
  ws.getRange(19,2).setValue(summary.byVaptStatus.done || 0);
  ws.getRange(20,2).setValue(summary.byVaptStatus.inProgress || 0);

  // Apply color coding to summary values
  // Critical/High = Red, Medium = Orange, Low/Info = Green
  ws.getRange(7,2).setBackground('#FFCDD2').setFontWeight('bold');  // Critical
  ws.getRange(8,2).setBackground('#FFCDD2').setFontWeight('bold');  // High
  ws.getRange(9,2).setBackground('#FFE0B2').setFontWeight('bold');  // Medium
  ws.getRange(10,2).setBackground('#C8E6C9').setFontWeight('bold');  // Low
  ws.getRange(11,2).setBackground('#C8E6C9').setFontWeight('bold');  // Info
}

/**
 * Apply conditional formatting for severity columns
 */
function applyVAPTConditionalFormatting_(ws, startRow, dataLength) {
  if (dataLength === 0) return;

  const endRow = startRow + dataLength - 1;

  // Critical columns (5, 10, 15) - Red if > 0
  [5, 10, 15].forEach(col => {
    const range = ws.getRange(startRow, col, dataLength, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#FFCDD2')
        .setFontColor('#C62828')
        .setRanges([range])
        .build();
    ws.setConditionalFormatRules([...ws.getConditionalFormatRules(), rule]);
  });

  // High columns (6, 11, 16) - Orange if > 0
  [6, 11, 16].forEach(col => {
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
