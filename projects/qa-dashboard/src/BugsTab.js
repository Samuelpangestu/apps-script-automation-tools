/**
 * BugsTab.js - Bug Historical Tracking Dashboard
 *
 * Track bug metrics dengan historical comparison untuk monitor kecepatan perbaikan bugs
 *
 * Features:
 * - Total Bugs (all status kecuali Closed)
 * - Breakdown by Priority: Critical, High, Medium, Low, Lowest
 * - Total Bugs Blocker (Medium-Critical)
 * - Environment breakdown
 * - Delta tracking: lihat kenaikan/penurunan dari refresh sebelumnya
 * - Color coding: Hijau = turun (bagus), Merah = naik (butuh perhatian)
 */

// ═══════════════════════════════════════════════════════════════════════
// BUILD BUGS TAB
// ═══════════════════════════════════════════════════════════════════════

function buildBugs(ss) {
  const ws = ss.insertSheet('Bugs', 1);  // Insert after Overview (position 1)
  ws.setTabColor('#D32F2F');
  ws.clear();
  initBugsHeaders_(ws);

  ws.getRange(5,1,1,38).merge()
      .setValue('▶ Run refreshDashboard() untuk mengisi data')
      .setBackground('#FFEBEE').setFontColor('#C62828').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(4);
}

function ensureBugsColumns_(ws) {
  const requiredColumns = 38;
  const currentColumns = ws.getMaxColumns();
  if (currentColumns < requiredColumns) {
    ws.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
  }
}

function initBugsHeaders_(ws) {
  ensureBugsColumns_(ws);
  ws.setColumnWidths(1,38,90);
  ws.setColumnWidth(3,180);  // Submodul wider

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
  ws.getRange(1,1,1,38).merge().setValue('Last refreshed: —')
      .setBackground('#FFEBEE').setFontColor('#C62828').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  // Row 2 — title
  h_(2,1,1,38,'🐛 BUG TRACKING  |  HISTORICAL METRICS WITH DELTA','#B71C1C','#FFFFFF');
  ws.getRange(2,1).setFontSize(13);
  ws.setRowHeight(2,30);

  // Row 3 — group headers
  h_(3,1, 1,3, 'MODULE INFO',        '#263238');
  h_(3,4, 1,7, 'BUGS BY PRIORITY',   '#D32F2F');
  h_(3,11,1,2, 'DELTA (vs Previous)','#FF6F00');
  h_(3,13,1,3, 'ENVIRONMENT BUGS',   '#6A1B9A');
  h_(3,16,1,5, 'STATUS BREAKDOWN',   '#1976D2');
  h_(3,21,1,5, 'BLOCKER STATUS',     '#E91E63');
  h_(3,26,1,2, 'METADATA',           '#37474F');
  h_(3,28,1,11, 'PRODUCTION BREAKDOWN','#4A148C');
  ws.setRowHeight(3,22);

  // Row 4 — column headers
  const headers = [
    'Project','Modul','Submodul',
    'Total','Critical','High','Medium','Low','Lowest','Blocker',
    'Δ Total','Δ Blocker',
    'Dev','UAT','Prod',
    'Open','In Progress','Fixed','Reopen','Verified',
    'BLK Open','BLK InProg','BLK Fixed','BLK Reopen','BLK Verified',
    'Prev Total',
    'Last Updated','Prod Blocker','Prod Critical','Prod High','Prod Medium','Prod Low','Prod Lowest','Prod Open','Prod InProg','Prod Fixed','Prod Reopen','Prod Verified'
  ];
  headers.forEach((lbl,i) => h_(4,i+1,1,1,lbl,'#1565C0'));

  // Add notes
  ws.getRange(4,4).setNote('Total Bugs\n\nAll status kecuali Closed\nMakin kecil makin bagus');
  ws.getRange(4,5).setNote('Priority: Critical/Highest\n\nBugs paling kritis yang harus segera diperbaiki');
  ws.getRange(4,10).setNote('Total Bugs Blocker\n\nBugs dengan Priority Medium-Critical/Highest dengan status selain Closed dan Won\'t Fix\n\nIncludes: Open, In Progress, Reopen, Fixed, Ready to Test, dll.\n\nTarget: 0 blocker bugs');
  ws.getRange(4,11).setNote('Delta Total Bugs\n\n🟢 Negatif = Turun (Good!)\n🔴 Positif = Naik (Need Action)\n⚪ 0 = Stabil');
  ws.getRange(4,12).setNote('Delta Blocker Bugs\n\n🟢 Negatif = Turun (Good!)\n🔴 Positif = Naik (Critical!)\n⚪ 0 = Stabil');
  ws.getRange(4,13).setNote('Bugs di Development environment\n\nBugs found during development/testing');
  ws.getRange(4,14).setNote('Bugs di UAT environment\n\nBugs found during UAT testing');
  ws.getRange(4,15).setNote('Bugs di Production environment\n\n⚠️ CRITICAL: Bugs affecting live users\nTarget: 0 prod bugs');
  ws.getRange(4,16).setNote('Status: Open\n\nBugs yang baru dilaporkan dan belum ditangani');
  ws.getRange(4,17).setNote('Status: In Progress\n\nBugs yang sedang dikerjakan oleh developer');
  ws.getRange(4,18).setNote('Status: Fixed\n\nBugs yang sudah diperbaiki, menunggu verifikasi');
  ws.getRange(4,19).setNote('Status: Reopen\n\nBugs yang di-reopen setelah verifikasi gagal');
  ws.getRange(4,20).setNote('Status: Verified/Ready to Test\n\nBugs yang sudah diverifikasi fixed oleh QA');
  ws.getRange(4,21).setNote('Blocker Open\n\nBlocker bugs (Medium-Critical) dengan status Open');
  ws.getRange(4,22).setNote('Blocker In Progress\n\nBlocker bugs (Medium-Critical) yang sedang dikerjakan');
  ws.getRange(4,23).setNote('Blocker Fixed\n\nBlocker bugs (Medium-Critical) yang sudah diperbaiki');
  ws.getRange(4,24).setNote('Blocker Reopen\n\nBlocker bugs (Medium-Critical) yang direopen');
  ws.getRange(4,25).setNote('Blocker Verified\n\nBlocker bugs (Medium-Critical) yang sudah diverifikasi');
  ws.getRange(4,26).setNote('Previous Total Bugs\n\nTotal bugs dari refresh sebelumnya\nDigunakan untuk calculate delta');
  ws.getRange(4,27).setNote('Last Updated\n\nTimestamp saat data terakhir di-refresh');
  ws.getRange(4,28).setNote('Production Blocker\n\nBugs di Production dengan priority Medium-Critical/Highest dan status selain Closed/Won\'t Fix.');
  ws.getRange(4,29).setNote('Production Critical\n\nBugs di Production dengan priority Critical/Highest.');
  ws.getRange(4,30).setNote('Production High\n\nBugs di Production dengan priority High.');
  ws.getRange(4,31).setNote('Production Medium\n\nBugs di Production dengan priority Medium.');
  ws.getRange(4,32).setNote('Production Low\n\nBugs di Production dengan priority Low.');
  ws.getRange(4,33).setNote('Production Lowest\n\nBugs di Production dengan priority Lowest.');
  ws.getRange(4,34).setNote('Production Open\n\nBugs di Production dengan status Open.');
  ws.getRange(4,35).setNote('Production In Progress\n\nBugs di Production dengan status In Progress.');
  ws.getRange(4,36).setNote('Production Fixed\n\nBugs di Production dengan status Fixed.');
  ws.getRange(4,37).setNote('Production Reopen\n\nBugs di Production dengan status Reopen.');
  ws.getRange(4,38).setNote('Production Verified\n\nBugs di Production dengan status Verified/Ready to Test/Done VAPT.');

  ws.setRowHeight(4,26);
  ws.setFrozenRows(4);
}

// ═══════════════════════════════════════════════════════════════════════
// WRITE BUGS DATA WITH DELTA TRACKING
// ═══════════════════════════════════════════════════════════════════════

function writeBugs(ss, allData) {
  let ws = ss.getSheetByName('Bugs');
  if (!ws) { buildBugs(ss); ws = ss.getSheetByName('Bugs'); }

  const headerCheck = ws.getRange(4,1,1,38).getValues()[0];
  if (headerCheck[0] !== 'Project' ||
      headerCheck[2] !== 'Submodul' ||
      headerCheck[37] !== 'Prod Verified') {
    initBugsHeaders_(ws);
  }

  // Get previous data for delta calculation
  const previousData = getPreviousBugsData_(ws);

  // Clear only the previously used data area; static formatting stays intact.
  const lastRow = ws.getLastRow();
  if (lastRow>=5) ws.getRange(5,1,lastRow-4,38).breakApart().clearContent();

  const rules = [];
  const now = new Date();

  // ═══════════════════════════════════════════════════════════════
  // AGGREGATE BUGS BY SUBMODUL FROM ALL MODULES
  // ═══════════════════════════════════════════════════════════════
  const submodulBugs = aggregateBugsBySubmodul_(allData);
  const submodulList = Object.keys(submodulBugs).sort();

  Logger.log('Total unique submoduls: ' + submodulList.length);

  const dataRows = [];
  const backgrounds = [];
  const fontColors = [];
  const numberFormats = [];
  submodulList.forEach((submodulKey, i) => {
    const bg = i%2===0 ? '#FFF5F5' : '#FFFFFF';
    const submodulData = submodulBugs[submodulKey];
    const bs = submodulData.stats;
    const prev = previousData[submodulKey] || {total:0,blocker:0};
    const deltaTotal = (bs.total||0)-prev.total;
    const deltaBlocker = (bs.blocker||0)-prev.blocker;
    const submodulName = submodulData.submodul||'Unknown';
    const safeName = String(submodulName).replace(/"/g,'""');
    const submodulValue = submodulData.qatmUrl
      ? '=HYPERLINK("' + submodulData.qatmUrl + '","' + safeName + '")'
      : submodulName;

    dataRows.push([
      submodulData.project||'',submodulData.modul||'',submodulValue,
      bs.total||0,bs.critical||0,bs.high||0,bs.medium||0,bs.low||0,bs.lowest||0,bs.blocker||0,
      deltaTotal,deltaBlocker,
      bs.devBugs||0,bs.uatBugs||0,bs.prodBugs||0,
      bs.openBugs||0,bs.inProgressBugs||0,bs.fixedBugs||0,bs.reopenBugs||0,bs.verifiedBugs||0,
      bs.blockerOpenBugs||0,bs.blockerInProgressBugs||0,bs.blockerFixedBugs||0,
      bs.blockerReopenBugs||0,bs.blockerVerifiedBugs||0,
      prev.total||0,now,
      bs.prodBlockerBugs||0,bs.prodCriticalBugs||0,bs.prodHighBugs||0,bs.prodMediumBugs||0,
      bs.prodLowBugs||0,bs.prodLowestBugs||0,bs.prodOpenBugs||0,bs.prodInProgressBugs||0,
      bs.prodFixedBugs||0,bs.prodReopenBugs||0,bs.prodVerifiedBugs||0
    ]);

    const rowBackgrounds = new Array(38).fill(bg);
    rowBackgrounds[10] = deltaTotal<0?'#C8E6C9':deltaTotal>0?'#FFCDD2':'#F5F5F5';
    rowBackgrounds[11] = deltaBlocker<0?'#81C784':deltaBlocker>0?'#EF5350':'#F5F5F5';
    backgrounds.push(rowBackgrounds);

    const rowFontColors = new Array(38).fill('#000000');
    rowFontColors[2] = submodulData.qatmUrl ? '#1155CC' : '#000000';
    rowFontColors[10] = deltaTotal<0?'#1B5E20':deltaTotal>0?'#C62828':'#757575';
    rowFontColors[11] = deltaBlocker<0?'#1B5E20':deltaBlocker>0?'#B71C1C':'#757575';
    fontColors.push(rowFontColors);

    numberFormats.push(dataRows[i].map((value,index) => index===26?'yyyy-mm-dd hh:mm':index>=3?'0':'General'));
  });

  if (dataRows.length > 0) {
    const dataRange = ws.getRange(5,1,dataRows.length,38);
    dataRange.setValues(dataRows)
      .setBackgrounds(backgrounds)
      .setFontColors(fontColors)
      .setNumberFormats(numberFormats)
      .setFontFamily('Arial').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(5,1,dataRows.length,3).setHorizontalAlignment('left');
    ws.getRange(5,3,dataRows.length,1).setFontWeight('bold');
    ws.getRange(5,11,dataRows.length,2).setFontWeight('bold');

    rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setBackground('#FFF9C4').setFontColor('#F57F17')
        .setRanges([ws.getRange(5,4,dataRows.length,1)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true)
        .setRanges([ws.getRange(5,10,dataRows.length,1)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
        .setRanges([ws.getRange(5,5,dataRows.length,2)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setBackground('#EF5350').setFontColor('#FFFFFF').setBold(true)
        .setRanges([ws.getRange(5,15,dataRows.length,1)]).build()
    );
  }

  // TOTAL row
  if (submodulList.length > 0) {
    const tr = 5+submodulList.length;
    ws.getRange(tr,1,1,3).merge().setValue('TOTAL')
        .setBackground('#E8F5E9').setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');

    // Sum bugs across all submoduls
    const sumBugs = (key) => submodulList.reduce((a, submodulKey) => {
      const bs = submodulBugs[submodulKey].stats;
      return a + (bs[key]||0);
    }, 0);

    [[4,'total'],[5,'critical'],[6,'high'],[7,'medium'],[8,'low'],[9,'lowest'],[10,'blocker'],
     [13,'devBugs'],[14,'uatBugs'],[15,'prodBugs'],
     [16,'openBugs'],[17,'inProgressBugs'],[18,'fixedBugs'],[19,'reopenBugs'],[20,'verifiedBugs'],
     [21,'blockerOpenBugs'],[22,'blockerInProgressBugs'],[23,'blockerFixedBugs'],[24,'blockerReopenBugs'],[25,'blockerVerifiedBugs'],
     [28,'prodBlockerBugs'],[29,'prodCriticalBugs'],[30,'prodHighBugs'],[31,'prodMediumBugs'],[32,'prodLowBugs'],[33,'prodLowestBugs'],
     [34,'prodOpenBugs'],[35,'prodInProgressBugs'],[36,'prodFixedBugs'],[37,'prodReopenBugs'],[38,'prodVerifiedBugs']].forEach(([col,key]) => {
      ws.getRange(tr,col).setValue(sumBugs(key)).setNumberFormat('0')
          .setBackground('#E8F5E9').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    });

    // Total deltas
    const totalDeltaTotal = submodulList.reduce((a, submodulKey) => {
      const bs = submodulBugs[submodulKey].stats;
      const prev = previousData[submodulKey] || {total: 0, blocker: 0};
      return a + ((bs.total||0) - prev.total);
    }, 0);

    const totalDeltaBlocker = submodulList.reduce((a, submodulKey) => {
      const bs = submodulBugs[submodulKey].stats;
      const prev = previousData[submodulKey] || {total: 0, blocker: 0};
      return a + ((bs.blocker||0) - prev.blocker);
    }, 0);

    const totalDeltaTotalCell = ws.getRange(tr,11).setValue(totalDeltaTotal).setNumberFormat('0')
        .setBackground(totalDeltaTotal<0?'#C8E6C9':totalDeltaTotal>0?'#FFCDD2':'#F5F5F5')
        .setFontColor(totalDeltaTotal<0?'#1B5E20':totalDeltaTotal>0?'#C62828':'#757575')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');

    const totalDeltaBlockerCell = ws.getRange(tr,12).setValue(totalDeltaBlocker).setNumberFormat('0')
        .setBackground(totalDeltaBlocker<0?'#81C784':totalDeltaBlocker>0?'#EF5350':'#F5F5F5')
        .setFontColor(totalDeltaBlocker<0?'#1B5E20':totalDeltaBlocker>0?'#B71C1C':'#757575')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');

    ws.setRowHeight(tr,22);
  }

  ws.setConditionalFormatRules(rules);

  // Update last refresh timestamp
  ws.getRange(1,1).setValue('Last refreshed: ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));

  // Charts removed - will use Web App dashboard for visualization
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Aggregate bugs by Submodul from all modules' BugReport sheets
 * Returns: {submodulKey: {project, modul, submodul, stats: {...}}}
 */
function aggregateBugsBySubmodul_(allData) {
  const submodulBugs = {};

  allData.forEach(moduleData => {
    if (!moduleData.id) return;

    try {
      const rows = moduleData.bugRows || [];
      if (rows.length === 0 && !moduleData.bugReportGid) {
        Logger.log('No BugReport for module: ' + moduleData.name);
        return;
      }

      rows.forEach(row => {
        const status = String(row[3]).trim();  // col D = Status
        const priority = String(row[2]).trim();  // col C = Priority
        const submodul = String(row[5]).trim();  // col F = Submodul
        const env = String(row[8]).trim();  // col I = Environment

        if (!submodul || submodul === '') return;  // Skip if no submodul

        // Create submodulKey (use submodul as unique key)
        const submodulKey = `${moduleData.project}|${moduleData.module}|${submodul}`;

        // Initialize if first time seeing this submodul
        if (!submodulBugs[submodulKey]) {
          // Get BugReport GID
          const qatmUrl = 'https://docs.google.com/spreadsheets/d/' + moduleData.id +
            '/edit#gid=' + moduleData.bugReportGid;

          submodulBugs[submodulKey] = {
            project: moduleData.project || '',
            modul: moduleData.module || '',
            submodul: submodul,
            qatmUrl: qatmUrl,  // Store QATM URL with BugReport GID
            stats: {
              total: 0,
              critical: 0,
              high: 0,
              medium: 0,
              low: 0,
              lowest: 0,
              blocker: 0,
              devBugs: 0,
              uatBugs: 0,
              prodBugs: 0,
              openBugs: 0,
              inProgressBugs: 0,
              fixedBugs: 0,
              reopenBugs: 0,
              verifiedBugs: 0,
              blockerOpenBugs: 0,
              blockerInProgressBugs: 0,
              blockerFixedBugs: 0,
              blockerReopenBugs: 0,
              blockerVerifiedBugs: 0,
              prodBlockerBugs: 0,
              prodCriticalBugs: 0,
              prodHighBugs: 0,
              prodMediumBugs: 0,
              prodLowBugs: 0,
              prodLowestBugs: 0,
              prodOpenBugs: 0,
              prodInProgressBugs: 0,
              prodFixedBugs: 0,
              prodReopenBugs: 0,
              prodVerifiedBugs: 0
            }
          };
        }

        const stats = submodulBugs[submodulKey].stats;

        // Count only active bugs (not Closed)
        if (status !== 'Closed') {
          stats.total++;

          // By priority
          if (priority === 'Critical' || priority === 'Highest') stats.critical++;
          else if (priority === 'High') stats.high++;
          else if (priority === 'Medium') stats.medium++;
          else if (priority === 'Low') stats.low++;
          else if (priority === 'Lowest') stats.lowest++;

          // Blocker bugs (Medium-Critical, status all except Closed and Won't Fix)
          if (status !== 'Closed' && status !== "Won't Fix" &&
              ['Critical','Highest','High','Medium'].includes(priority)) {
            stats.blocker++;
          }

          // By environment
          const envLower = env.toLowerCase().trim();
          const isProduction = envLower === 'production' || envLower === 'prod';
          if (env === 'Development' || env === 'Dev') stats.devBugs++;
          else if (env === 'UAT') stats.uatBugs++;
          else if (isProduction) {
            stats.prodBugs++;
            if (priority === 'Critical' || priority === 'Highest') stats.prodCriticalBugs++;
            else if (priority === 'High') stats.prodHighBugs++;
            else if (priority === 'Medium') stats.prodMediumBugs++;
            else if (priority === 'Low') stats.prodLowBugs++;
            else if (priority === 'Lowest') stats.prodLowestBugs++;
            if (['Critical','Highest','High','Medium'].includes(priority) && status !== "Won't Fix") stats.prodBlockerBugs++;
          }

          // By status (case-insensitive, trim whitespace)
          const statusLower = status.toLowerCase().trim();
          const isBlockerPriority = ['Critical','Highest','High','Medium'].includes(priority);

          if (statusLower === 'open') {
            stats.openBugs++;
            if (isBlockerPriority) stats.blockerOpenBugs++;
            if (isProduction) stats.prodOpenBugs++;
          } else if (statusLower === 'in progress' || statusLower === 'in progress vapt') {
            stats.inProgressBugs++;
            if (isBlockerPriority) stats.blockerInProgressBugs++;
            if (isProduction) stats.prodInProgressBugs++;
          } else if (statusLower === 'fixed') {
            stats.fixedBugs++;
            if (isBlockerPriority) stats.blockerFixedBugs++;
            if (isProduction) stats.prodFixedBugs++;
          } else if (statusLower === 'reopen') {
            stats.reopenBugs++;
            if (isBlockerPriority) stats.blockerReopenBugs++;
            if (isProduction) stats.prodReopenBugs++;
          } else if (statusLower === 'verified' || statusLower === 'ready to test' || statusLower === 'done vapt') {
            stats.verifiedBugs++;
            if (isBlockerPriority) stats.blockerVerifiedBugs++;
            if (isProduction) stats.prodVerifiedBugs++;
          } else {
            // Log unmatched statuses for debugging
            if (!submodulBugs[submodulKey].unmatchedStatuses) {
              submodulBugs[submodulKey].unmatchedStatuses = {};
            }
            if (!submodulBugs[submodulKey].unmatchedStatuses[statusLower]) {
              submodulBugs[submodulKey].unmatchedStatuses[statusLower] = 0;
            }
            submodulBugs[submodulKey].unmatchedStatuses[statusLower]++;
          }
        }
      });

    } catch(e) {
      Logger.log('Error reading bugs from module ' + moduleData.name + ': ' + e.message);
    }
  });

  // Log unmatched statuses for debugging
  Object.keys(submodulBugs).forEach(key => {
    const unmatchedStatuses = submodulBugs[key].unmatchedStatuses;
    if (unmatchedStatuses && Object.keys(unmatchedStatuses).length > 0) {
      Logger.log('⚠️ Unmatched statuses in ' + key + ':');
      Object.keys(unmatchedStatuses).forEach(status => {
        Logger.log('  - "' + status + '": ' + unmatchedStatuses[status] + ' bugs');
      });
    }
  });

  return submodulBugs;
}

/**
 * Get previous bug data from Bugs tab for delta calculation
 * Returns: {submodulKey: {total: X, blocker: Y}}
 */
function getPreviousBugsData_(ws) {
  const previousData = {};

  try {
    const lastRow = ws.getLastRow();
    if (lastRow < 5) return previousData;  // No previous data

    // Read existing data (skip header rows)
    const data = ws.getRange(5, 1, lastRow-4, 27).getValues();

    data.forEach(row => {
      const project = String(row[0]).trim();
      const modul = String(row[1]).trim();
      const submodul = String(row[2]).trim();
      const total = row[3] || 0;  // col 4 = Total Bugs
      const blocker = row[9] || 0;  // col 10 = Blocker Bugs

      if (project && modul && submodul) {
        const submodulKey = `${project}|${modul}|${submodul}`;
        previousData[submodulKey] = {
          total: typeof total === 'number' ? total : 0,
          blocker: typeof blocker === 'number' ? blocker : 0
        };
      }
    });
  } catch(e) {
    Logger.log('Error reading previous bugs data: ' + e.message);
  }

  return previousData;
}

// ═══════════════════════════════════════════════════════════════════════
// CHARTS - BUG TREND VISUALIZATION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Build charts for bug trend visualization
 */
function buildBugsCharts_(ws, submodulList, submodulBugs, previousData) {
  if (!submodulList || submodulList.length === 0) return;

  // Remove existing charts
  ws.getCharts().forEach(c => ws.removeChart(c));

  const n = submodulList.length;
  const dRow = 5;
  const cRow = dRow + n + 4;  // Charts start 4 rows after data

  // Helper to safely add chart
  function tryChart_(fn) {
    try { fn(); } catch(e) { Logger.log('Chart error: ' + e.message); }
  }

  // ═══════════════════════════════════════════════════════════════
  // CHART 1: Total Bugs vs Blocker Bugs (Line Chart)
  // ═══════════════════════════════════════════════════════════════
  tryChart_(() => ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(ws.getRange(4, 3, n+1, 1))    // Submodul names (col C)
      .addRange(ws.getRange(4, 4, n+1, 1))    // Total Bugs (col D)
      .addRange(ws.getRange(4, 10, n+1, 1))   // Blocker Bugs (col J)
      .setPosition(cRow, 1, 0, 0)
      .setOption('title', '📊 Total Bugs vs Blocker Bugs Trend')
      .setOption('curveType', 'function')
      .setOption('lineWidth', 3)
      .setOption('pointSize', 5)
      .setOption('colors', ['#FF6B6B', '#C62828'])
      .setOption('legend', {position: 'top'})
      .setOption('hAxis', {title: 'Module', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Bug Count', minValue: 0})
      .setOption('width', 550)
      .setOption('height', 300)
      .build()));

  // ═══════════════════════════════════════════════════════════════
  // CHART 2: Delta Trend - Which Modules Improving/Worsening (Bar Chart)
  // ═══════════════════════════════════════════════════════════════
  tryChart_(() => ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(4, 3, n+1, 1))    // Submodul names (col C)
      .addRange(ws.getRange(4, 11, n+1, 1))   // Delta Total (col K)
      .addRange(ws.getRange(4, 12, n+1, 1))   // Delta Blocker (col L)
      .setPosition(cRow, 10, 0, 0)
      .setOption('title', '📈 Bug Delta: Progress Tracking (Negative = Good!)')
      .setOption('isStacked', false)
      .setOption('colors', ['#FF9800', '#F44336'])
      .setOption('legend', {position: 'top'})
      .setOption('hAxis', {title: 'Delta (Negative = Improvement)'})
      .setOption('width', 550)
      .setOption('height', 300)
      .build()));

  // ═══════════════════════════════════════════════════════════════
  // CHART 3: Priority Breakdown (Stacked Column Chart)
  // ═══════════════════════════════════════════════════════════════
  tryChart_(() => ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(ws.getRange(4, 3, n+1, 1))    // Submodul names (col C)
      .addRange(ws.getRange(4, 5, n+1, 1))    // Critical (col E)
      .addRange(ws.getRange(4, 6, n+1, 1))    // High (col F)
      .addRange(ws.getRange(4, 7, n+1, 1))    // Medium (col G)
      .addRange(ws.getRange(4, 8, n+1, 1))    // Low (col H)
      .addRange(ws.getRange(4, 9, n+1, 1))    // Lowest (col I)
      .setPosition(cRow + 22, 1, 0, 0)
      .setOption('title', '🎯 Bugs by Priority - Distribution per Module')
      .setOption('isStacked', true)
      .setOption('colors', ['#C62828', '#E53935', '#FB8C00', '#FFA726', '#FFB74D'])
      .setOption('legend', {position: 'top'})
      .setOption('hAxis', {title: 'Module', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Bug Count', minValue: 0})
      .setOption('width', 550)
      .setOption('height', 300)
      .build()));

  // ═══════════════════════════════════════════════════════════════
  // CHART 4: Environment Breakdown (Stacked Column Chart)
  // ═══════════════════════════════════════════════════════════════
  tryChart_(() => ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(ws.getRange(4, 3, n+1, 1))    // Submodul names (col C)
      .addRange(ws.getRange(4, 13, n+1, 1))   // Dev Bugs (col M)
      .addRange(ws.getRange(4, 14, n+1, 1))   // UAT Bugs (col N)
      .addRange(ws.getRange(4, 15, n+1, 1))   // Prod Bugs (col O)
      .setPosition(cRow + 22, 10, 0, 0)
      .setOption('title', '🌍 Bugs by Environment - Dev vs UAT vs PROD')
      .setOption('isStacked', true)
      .setOption('colors', ['#4CAF50', '#FF9800', '#F44336'])
      .setOption('legend', {position: 'top'})
      .setOption('hAxis', {title: 'Module', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Bug Count', minValue: 0})
      .setOption('width', 550)
      .setOption('height', 300)
      .build()));

  // ═══════════════════════════════════════════════════════════════
  // CHART 5: Current vs Previous Total Bugs (Combo Chart)
  // ═══════════════════════════════════════════════════════════════
  tryChart_(() => ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(ws.getRange(4, 3, n+1, 1))    // Submodul names (col C)
      .addRange(ws.getRange(4, 16, n+1, 1))   // Previous Total (col P)
      .addRange(ws.getRange(4, 4, n+1, 1))    // Current Total (col D)
      .setPosition(cRow + 44, 1, 0, 0)
      .setOption('title', '📉 Bug Trend: Previous vs Current Total')
      .setOption('seriesType', 'bars')
      .setOption('series', {
        0: {type: 'bars', color: '#BDBDBD', targetAxisIndex: 0},  // Previous (gray bars)
        1: {type: 'line', color: '#F44336', lineWidth: 3, pointSize: 6, targetAxisIndex: 0}  // Current (red line)
      })
      .setOption('legend', {position: 'top'})
      .setOption('hAxis', {title: 'Module', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Bug Count', minValue: 0})
      .setOption('width', 1150)
      .setOption('height', 350)
      .build()));

  Logger.log('✅ Bugs charts created successfully');
}

/**
 * Get detailed bug stats with environment breakdown
 * Reads from BugReport sheet directly for more detailed data
 */
function getBugDetailsForTab_(moduleId) {
  if (!moduleId) return null;

  try {
    const src = SpreadsheetApp.openById(moduleId);
    const bugSheet = src.getSheetByName('BugReport');

    if (!bugSheet) return null;

    const rows = bugSheet.getDataRange().getValues().slice(4).filter(r => r[0] && r[0] !== '');
    const cnt = (fn) => rows.filter(fn).length;

    // All bugs except Closed
    const activeBugs = rows.filter(r => r[3] !== 'Closed');

    return {
      total: activeBugs.length,
      critical: cnt(r => r[3] !== 'Closed' && (r[2] === 'Critical' || r[2] === 'Highest')),
      high: cnt(r => r[3] !== 'Closed' && r[2] === 'High'),
      medium: cnt(r => r[3] !== 'Closed' && r[2] === 'Medium'),
      low: cnt(r => r[3] !== 'Closed' && r[2] === 'Low'),
      lowest: cnt(r => r[3] !== 'Closed' && r[2] === 'Lowest'),
      blocker: cnt(r => r[3] !== 'Closed' && r[3] !== "Won't Fix" && ['Critical','Highest','High','Medium'].includes(r[2])),

      // Environment breakdown (col 9 = Environment, index 8)
      devBugs: cnt(r => r[3] !== 'Closed' && (r[8] === 'Development' || r[8] === 'Dev')),
      uatBugs: cnt(r => r[3] !== 'Closed' && r[8] === 'UAT'),
      prodBugs: cnt(r => r[3] !== 'Closed' && r[8] === 'Production'),
    };
  } catch(e) {
    Logger.log('Error getting bug details for module ' + moduleId + ': ' + e.message);
    return null;
  }
}
