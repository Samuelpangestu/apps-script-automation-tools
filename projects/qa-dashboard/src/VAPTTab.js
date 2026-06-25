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
// AGGREGATE VAPT DATA FROM MODULE DATA
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggregate VAPT data from all modules (from QATM Summary tabs)
 * Replaces external VAPT spreadsheet fetch
 * @param {Array} allData - Module data from pullModuleData_()
 * @returns {Object} {table: [], summary: {}}
 */
function aggregateVAPTFromModules_(allData) {
  const table = [];
  const summary = {
    blocker: 0,
    totalFindings: 0,
    bySeverity: { critical: 0, high: 0, medium: 0, low: 0, info: 0 },
    byStatus: { readyToRetest: 0, open: 0, closed: 0 },
    totalApps: 0,
    byVaptStatus: { done: 0, inProgress: 0 },
    prodCount: 0
  };

  // Process each module's VAPT data
  allData.forEach(mod => {
    if (!mod.vapt || mod.vapt.total === 0) {
      return; // Skip modules with no VAPT data
    }

    const v = mod.vapt;

    // Calculate blocker with multiple fallbacks:
    // 1. Use blockerCount if available and non-zero
    // 2. Sum blockerCritical + blockerHigh + blockerMedium if available
    // 3. Fallback: Calculate from open status (Critical + High + Medium from "Open" findings)
    let blocker = 0;
    if (v.blockerCount !== undefined && v.blockerCount > 0) {
      blocker = v.blockerCount;
    } else if ((v.blockerCritical || 0) + (v.blockerHigh || 0) + (v.blockerMedium || 0) > 0) {
      blocker = (v.blockerCritical || 0) + (v.blockerHigh || 0) + (v.blockerMedium || 0);
    } else {
      // Fallback: Assume all Critical/High/Medium are blockers if blocker fields not set
      // This handles old Summary tabs without VAPT Blocker Breakdown section
      blocker = (v.critical || 0) + (v.high || 0) + (v.medium || 0);
    }

    // Only add to table if there are findings
    if (v.total > 0) {
      table.push({
        project: mod.project || 'Unknown',
        aplikasi: mod.submodule || mod.module || mod.name,
        blocker: blocker,
        open: {
          critical: v.critical || 0,
          high: v.high || 0,
          medium: v.medium || 0,
          low: v.low || 0,
          info: v.informational || 0
        }
      });

      summary.totalApps++;
    }

    // Aggregate summary
    summary.blocker += blocker;
    summary.totalFindings += v.total || 0;
    summary.bySeverity.critical += v.critical || 0;
    summary.bySeverity.high += v.high || 0;
    summary.bySeverity.medium += v.medium || 0;
    summary.bySeverity.low += v.low || 0;
    summary.bySeverity.info += v.informational || 0;
    summary.byStatus.open += v.open || 0;
    summary.byStatus.closed += v.closed || 0;
    summary.byVaptStatus.done += v.done || 0;
    summary.byVaptStatus.inProgress += (v.todo || 0) + (v.onProgress || 0);
  });

  Logger.log('Aggregated VAPT from ' + allData.length + ' modules: ' +
             summary.totalApps + ' apps, ' + summary.totalFindings + ' findings, ' +
             summary.blocker + ' blockers');

  return { table, summary };
}

// ═══════════════════════════════════════════════════════════════════════
// BUILD VAPT TAB
// ═══════════════════════════════════════════════════════════════════════

function buildVAPT(ss) {
  // Delete existing VAPT tab if it exists
  const existingSheet = ss.getSheetByName('VAPT');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

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

  // Clear ALL data rows (from row 10 onwards) - 8 columns now
  const lastRow = ws.getMaxRows();
  if (lastRow >= 10) ws.getRange(10, 1, lastRow - 9, 8).clearContent().clearFormat();

  // Calculate blocker for each row
  const now = new Date();
  let data = vaptData.table.map(row => {
    // Calculate blocker: Critical + High + Medium yang Open
    const blocker = (row.open.critical || 0) + (row.open.high || 0) + (row.open.medium || 0);
    return {...row, blocker: blocker};
  });

  // Sort by Project, then by blocker descending
  data.sort((a, b) => {
    // First sort by project name
    if (a.project < b.project) return -1;
    if (a.project > b.project) return 1;
    // Then by blocker descending
    return b.blocker - a.blocker;
  });

  // Count unique projects
  const uniqueProjects = [...new Set(data.map(row => row.project))];
  const totalProjects = uniqueProjects.length;

  // Update summary
  updateVAPTSummary_(ws, vaptData.summary, data, totalProjects);

  let currentRow = 10;

  // ═══════════════════════════════════════════════════════════════════════
  // PER-PROJECT VAPT SECTIONS
  // ═══════════════════════════════════════════════════════════════════════

  // Group by project
  uniqueProjects.forEach(projectName => {
    const projectData = data.filter(row => row.project === projectName);

    // Section header for each project
    ws.getRange(currentRow, 1, 1, 8).merge()
        .setValue('═══ ' + projectName + ' ═══')
        .setBackground('#37474F').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.setRowHeight(currentRow, 24);
    currentRow++;

    // Table headers (8 columns - added Project)
    const headers = ['Project', 'Aplikasi', 'Blocker', 'Critical', 'High', 'Medium', 'Low', 'Info'];
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
    projectData.forEach((row, i) => {
      writeVAPTRow_(ws, currentRow, row, i);
      currentRow++;
    });

    currentRow++; // Blank row after section
  });

  // Update last refresh timestamp
  ws.getRange(1, 1).setValue('Last refreshed: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss'));

  Logger.log('✅ VAPT tab updated: ' + data.length + ' applications across ' + totalProjects + ' project(s)');
}

/**
 * Write a single VAPT data row (8 columns - added Project)
 */
function writeVAPTRow_(ws, row, data, index) {
  const bg = index % 2 === 0 ? '#FFF8E1' : '#FFFFFF';

  function cell(col, val) {
    return ws.getRange(row, col).setValue(val == null ? '' : val).setBackground(bg)
        .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Col 1: Project (left-aligned)
  ws.getRange(row, 1).setValue(data.project || '').setBackground(bg).setHorizontalAlignment('left')
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);

  // Col 2: Aplikasi (left-aligned)
  ws.getRange(row, 2).setValue(data.aplikasi).setBackground(bg).setHorizontalAlignment('left')
      .setFontFamily('Arial').setFontSize(9)
      .setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);

  // Col 3: Blocker (colored: red if > 0, green if 0)
  const blockerCell = ws.getRange(row, 3);
  blockerCell.setValue(data.blocker).setBackground(bg).setFontFamily('Arial').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold')
      .setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  if (data.blocker > 0) {
    blockerCell.setBackground('#FFCDD2').setFontColor('#C62828');  // Red
  } else {
    blockerCell.setBackground('#C8E6C9').setFontColor('#2E7D32');  // Green
  }

  // Col 4-8: OPEN findings (Critical, High, Medium, Low, Info)
  cell(4, data.open.critical);
  cell(5, data.open.high);
  cell(6, data.open.medium);
  cell(7, data.open.low);
  cell(8, data.open.info);

  ws.setRowHeight(row, 22);
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Update summary section with calculated metrics (per-project)
 */
function updateVAPTSummary_(ws, summary, data, totalProjects) {
  // Row 5: Total Blocker & Total Projects
  const totalBlockerCell = ws.getRange(5, 2);
  totalBlockerCell.setValue(summary.blocker || 0).setFontWeight('bold').setFontSize(12);
  if (summary.blocker > 0) {
    totalBlockerCell.setBackground('#FFCDD2').setFontColor('#C62828');  // Red
  } else {
    totalBlockerCell.setBackground('#C8E6C9').setFontColor('#2E7D32');  // Green
  }

  ws.getRange(5, 5).setValue(totalProjects || 0).setFontWeight('bold');

  // Row 6: Apps with Blocker & Total Apps
  const appsWithBlocker = data ? data.filter(app => app.blocker > 0).length : 0;
  const totalApps = data ? data.length : 0;
  ws.getRange(6, 2).setValue(appsWithBlocker).setFontWeight('bold');
  ws.getRange(6, 5).setValue(totalApps).setFontWeight('bold');
}

// ═══════════════════════════════════════════════════════════════════════
// BUILD VAPT HISTORY TAB
// ═══════════════════════════════════════════════════════════════════════

function buildVAPTHistory(ss) {
  // Delete existing VAPT History tab if it exists
  const existingSheet = ss.getSheetByName('VAPT History');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

  const ws = ss.insertSheet('VAPT History');
  ws.setTabColor('#BF360C');
  ws.clear();

  const hdrs = [
    'Timestamp','Project','Module','Submodule','Total',
    'Critical','High','Medium','Low','Informational',
    'Todo','On Progress','Done','Open','Closed',
    'Blocker Total','Blocker Critical','Blocker High','Blocker Medium'
  ];

  ws.getRange(1,1,1,hdrs.length).merge().setValue('VAPT HISTORY  —  Per-Submodule Trend Data from QATM Summary (auto-appended setiap refresh)')
      .setBackground('#BF360C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold')
      .setBackground('#EF6C00').setFontColor('#FFFFFF');
  ws.setFrozenRows(2);
  ws.setColumnWidth(1,140);  // Timestamp
  ws.setColumnWidth(2,120);  // Project
  ws.setColumnWidth(3,120);  // Module
  ws.setColumnWidth(4,140);  // Submodule
  for(let c=5;c<=hdrs.length;c++)ws.setColumnWidth(c,80);

  // Add notes
  ws.getRange(2,1).setNote('Timestamp\n\nAuto-appended saat refreshDashboard()');
  ws.getRange(2,2).setNote('Project\n\nProject name from Config col C');
  ws.getRange(2,3).setNote('Module\n\nModule name from Config col D');
  ws.getRange(2,4).setNote('Submodule\n\nSubmodule name from Config col E\nData source: QATM Summary sheet');
  ws.getRange(2,5).setNote('Total\n\nTotal VAPT findings from QATM Summary');
  ws.getRange(2,6).setNote('Critical\n\nCritical severity findings (Risk Level)');
  ws.getRange(2,7).setNote('High\n\nHigh severity findings (Risk Level)');
  ws.getRange(2,8).setNote('Medium\n\nMedium severity findings (Risk Level)');
  ws.getRange(2,9).setNote('Low\n\nLow severity findings (Risk Level)');
  ws.getRange(2,10).setNote('Informational\n\nInformational severity findings (Risk Level)');
  ws.getRange(2,11).setNote('Todo\n\nPending fixes (Status Fix)');
  ws.getRange(2,12).setNote('On Progress\n\nIn progress remediation (Status Fix)');
  ws.getRange(2,13).setNote('Done\n\nCompleted fixes (Status Fix)');
  ws.getRange(2,14).setNote('Open\n\nOpen findings (Re-VAPT status by Pentester)');
  ws.getRange(2,15).setNote('Closed\n\nClosed findings (Re-VAPT status by Pentester)');
}

/**
 * Append daily snapshot to VAPT History (per-submodule from QATM Summary)
 *
 * @param {Spreadsheet} ss - Dashboard spreadsheet
 * @param {Array} allData - Module data from pullModuleData_() containing vapt objects
 */
function appendVAPTHistory(ss, allData) {
  let ws = ss.getSheetByName('VAPT History');
  if (!ws) {
    buildVAPTHistory(ss);
    ws = ss.getSheetByName('VAPT History');
  }

  const now = new Date();
  const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Create one row per submodule with VAPT data
  const historyRows = [];

  if (!allData || allData.length === 0) {
    Logger.log('⚠️ No module data provided to appendVAPTHistory');
    return;
  }

  Logger.log('📊 Processing VAPT data for ' + allData.length + ' module(s)...');

  allData.forEach(mod => {
    // Skip modules without VAPT data or with total = 0
    if (!mod.vapt || mod.vapt.total === 0) {
      Logger.log('  ⏭️ Skipped ' + (mod.submodule || mod.name) + ': No VAPT data');
      return;
    }

    const row = [
      ts,                           // Timestamp
      mod.project || '',            // Project (Config col C)
      mod.module || '',             // Module (Config col D)
      mod.submodule || mod.name,    // Submodule (Config col E)
      mod.vapt.total || 0,          // Total findings
      mod.vapt.critical || 0,       // Critical
      mod.vapt.high || 0,           // High
      mod.vapt.medium || 0,         // Medium
      mod.vapt.low || 0,            // Low
      mod.vapt.informational || 0,  // Informational
      mod.vapt.todo || 0,           // Todo (Status Fix)
      mod.vapt.onProgress || 0,     // On Progress (Status Fix)
      mod.vapt.done || 0,           // Done (Status Fix)
      mod.vapt.open || 0,           // Open (Re-VAPT)
      mod.vapt.closed || 0,         // Closed (Re-VAPT)
      mod.vapt.blockerCount || 0,   // Blocker Total
      mod.vapt.blockerCritical || 0,// Blocker Critical
      mod.vapt.blockerHigh || 0,    // Blocker High
      mod.vapt.blockerMedium || 0   // Blocker Medium
    ];

    historyRows.push(row);
    Logger.log('  ✅ ' + (mod.submodule || mod.name) + ': Total=' + mod.vapt.total +
               ' (C=' + mod.vapt.critical + ' H=' + mod.vapt.high + ' M=' + mod.vapt.medium + ')');
  });

  if (historyRows.length === 0) {
    Logger.log('⚠️ No modules with VAPT data to append');
    return;
  }

  // SMART APPEND: Check if today's data already exists for each submodule
  const lastRow = ws.getLastRow();
  if (lastRow >= 3) {
    // Detect column count from header row
    const headerRow = ws.getRange(2, 1, 1, ws.getLastColumn()).getValues()[0];
    const numCols = headerRow.length; // Auto-detect: 15 old, 19 new
    const existingData = ws.getRange(3, 1, lastRow - 2, numCols).getValues();
    const todayRows = {};  // Map: "project_module_submodule" => row index

    // Find all rows from today
    existingData.forEach((row, i) => {
      const rowDate = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (rowDate === today) {
        const key = String(row[1] || '').trim() + '_' +
                   String(row[2] || '').trim() + '_' +
                   String(row[3] || '').trim();  // project_module_submodule
        todayRows[key] = i + 3;  // +3 because row 1-2 are headers
      }
    });

    // Update existing rows or collect new rows to append
    const newRows = [];
    historyRows.forEach(row => {
      const key = String(row[1] || '').trim() + '_' +
                 String(row[2] || '').trim() + '_' +
                 String(row[3] || '').trim();  // project_module_submodule

      if (todayRows[key]) {
        // Update existing row (overwrite with latest data)
        ws.getRange(todayRows[key], 1, 1, numCols).breakApart().setValues([row]);
        Logger.log('  🔄 Updated today\'s data for: ' + row[3]);
      } else {
        // Collect for batch append
        newRows.push(row);
      }
    });

    // Append only new rows
    if (newRows.length > 0) {
      const startRow = Math.max(lastRow + 1, 3);
      const targetRange = ws.getRange(startRow, 1, newRows.length, numCols);
      targetRange.breakApart();  // Unmerge any merged cells in target area
      targetRange.setValues(newRows);
      Logger.log('✅ VAPT History appended: ' + newRows.length + ' new rows at row ' + startRow);
    } else {
      Logger.log('✅ VAPT History updated: All submodules already had today\'s data');
    }
  } else {
    // No data yet, just append
    const numCols = 19; // Updated: 15 base + 4 blocker columns
    const startRow = 3;
    const targetRange = ws.getRange(startRow, 1, historyRows.length, numCols);
    targetRange.breakApart();
    targetRange.setValues(historyRows);
    Logger.log('✅ VAPT History appended: ' + historyRows.length + ' rows at row ' + startRow);
  }
}

/**
 * Create history row from per-project summary data
 */
function createHistoryRow_(timestamp, projectName, summary, allData) {
  // Calculate blocker (Critical + High + Medium OPEN findings)
  const blocker = summary.blocker || 0;

  // Count apps for this project
  const projectApps = allData ? allData.filter(app => app.project === projectName) : [];
  const totalApps = projectApps.length;
  const appsWithBlocker = projectApps.filter(app => app.blocker > 0).length;

  return [
    timestamp,
    projectName,
    blocker,
    summary.bySeverity.critical || 0,
    summary.bySeverity.high || 0,
    summary.bySeverity.medium || 0,
    summary.bySeverity.low || 0,
    summary.bySeverity.info || 0,
    totalApps,
    appsWithBlocker
  ];
}

/**
 * Cleanup VAPT History tab - keep only 90 days of data, one entry per day per project
 * Manual trigger from menu
 */
function cleanupVAPTHistoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('VAPT History');
  if(!ws){
    SpreadsheetApp.getUi().alert('VAPT History tab not found');
    return;
  }

  const lastRow = ws.getLastRow();
  if(lastRow<3){
    SpreadsheetApp.getUi().alert('No data to cleanup');
    return;
  }

  // Get all data (10 columns for VAPT History)
  const data = ws.getRange(3,1,lastRow-2,10).getValues();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 90);  // 90 days ago

  // Group by date+project, keep latest entry per day
  const dailyData = {};  // Map: "date|project" => row data

  data.forEach(row=>{
    const ts = new Date(row[0]);
    if(ts < cutoffDate) return;  // Skip old data

    const date = Utilities.formatDate(ts,Session.getScriptTimeZone(),'yyyy-MM-dd');
    const key = `${date}|${row[1]||''}`;  // date|project

    // Keep only latest entry for this day+project
    if(!dailyData[key] || new Date(dailyData[key][0]) < ts){
      dailyData[key] = row;
    }
  });

  // Convert back to array and sort by date (oldest first)
  const cleanedData = Object.values(dailyData).sort((a,b)=> new Date(a[0]) - new Date(b[0]));

  const rowsDeleted = data.length - cleanedData.length;

  if(rowsDeleted===0){
    SpreadsheetApp.getUi().alert('No duplicate or old data found. VAPT History is already clean!');
    return;
  }

  // Clear all data and rewrite cleaned data
  ws.getRange(3,1,lastRow-2,10).clearContent().breakApart();
  if(cleanedData.length>0){
    const targetRange = ws.getRange(3,1,cleanedData.length,10);
    targetRange.breakApart();  // Ensure no merged cells
    targetRange.setValues(cleanedData);
  }

  SpreadsheetApp.getUi().alert(
    `✅ VAPT History Cleanup Complete!\n\n` +
    `Rows before: ${data.length}\n` +
    `Rows after: ${cleanedData.length}\n` +
    `Deleted: ${rowsDeleted} rows\n\n` +
    `Retention: 90 days, 1 entry per day per project`
  );
  Logger.log(`✅ VAPT History cleanup: Deleted ${rowsDeleted} rows, kept ${cleanedData.length} rows`);
}
