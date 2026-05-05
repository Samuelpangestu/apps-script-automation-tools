/**
 * BroadcastVAPTTabs.js - Broadcast VAPT tabs to existing QATM spreadsheets
 *
 * NEW STRUCTURE - Adds 3 VAPT tabs + Summary section to ALL active QATM spreadsheets:
 * 1. VAPT - Helper (Dashboard/Tracking)
 * 2. VAPT - Detail Finding (Combined Regular + Ad Hoc)
 * 3. VAPT - Evidence (Combined Regular + Ad Hoc)
 * 4. VAPT Summary section in Summary tab (row 35+)
 *
 * Usage: Run broadcastVAPTTabsToAllQATMs() from QA Dashboard
 *
 * Features:
 * - Auto-reads QATM list from Dashboard Config tab
 * - Skips QATMs that already have new VAPT tabs
 * - Deletes old VAPT tabs if found (Detail Finding - VAPT, Evidence - VAPT)
 * - Adds VAPT metrics to Summary tab without impacting existing content
 * - Handles merged cells errors gracefully
 * - Provides detailed summary report
 */

/**
 * MAIN FUNCTION: Broadcast VAPT tabs to all active QATMs
 *
 * Automatically reads QATM list from Dashboard Config tab and adds 3 VAPT tabs to each.
 * QATMs that already have all 3 new tabs will be skipped.
 * Old VAPT tabs will be replaced with new structure.
 *
 * Run this from QA Dashboard: Extensions > Apps Script > broadcastVAPTTabsToAllQATMs
 */
function broadcastVAPTTabsToAllQATMs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Verify running from Dashboard
  const cfg = ss.getSheetByName('Config');
  if (!cfg) {
    ui.alert(
      '❌ Config Not Found',
      'This function must be run from QA Dashboard.\n\n' +
      'Config tab with QATM module list not found.',
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.alert(
    '🔒 Broadcast VAPT Tabs + Summary (NEW STRUCTURE)',
    'Add 3 NEW VAPT tabs + Summary section to ALL active QATMs:\n\n' +
    '📋 New Tabs:\n' +
    '• VAPT - Helper (Dashboard/Tracking)\n' +
    '• VAPT - Detail Finding (Simplified)\n' +
    '• VAPT - Evidence\n\n' +
    '📊 Summary Section (row 35+):\n' +
    '• Total findings\n' +
    '• By Risk Level, Status Fix, Status Re-VAPT\n\n' +
    '⚠️ OLD VAPT tabs will be REPLACED with new structure\n' +
    'QATMs with NEW tabs already will be skipped.\n' +
    'Time: ~1-2 minutes per QATM\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('❌ Broadcast cancelled.');
    return;
  }

  try {
    Logger.log('🔒 Starting VAPT Tabs Broadcast...');
    Logger.log('══════════════════════════════════════════');

    const cfgData = cfg.getDataRange().getValues();
    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;
    const errors = [];
    const skipped = [];

    // Process each QATM (starting from row 4, rows 1-3 are headers)
    for (let i = 3; i < cfgData.length; i++) {
      const active = cfgData[i][0] === true;  // Col A: Active
      const project = String(cfgData[i][2]).trim();  // Col C: Project
      const modul = String(cfgData[i][3]).trim();    // Col D: Modul
      const qatmId = String(cfgData[i][6]).trim();   // Col G: QATM Spreadsheet ID

      if (!active || !qatmId || qatmId.length < 10) {
        continue;  // Skip inactive or invalid entries
      }

      try {
        Logger.log('\n📂 Processing: ' + project + ' - ' + modul);
        Logger.log('   Spreadsheet ID: ' + qatmId);

        const qatmSs = SpreadsheetApp.openById(qatmId);
        Logger.log('   ✅ Opened: ' + qatmSs.getName());

        // Check if NEW VAPT tabs already exist
        const newTabs = [
          qatmSs.getSheetByName('VAPT - Helper'),
          qatmSs.getSheetByName('VAPT - Detail Finding'),
          qatmSs.getSheetByName('VAPT - Evidence')
        ];

        const newTabsExist = newTabs.filter(tab => tab !== null).length;

        if (newTabsExist === 3) {
          Logger.log('   ⏭️  All NEW VAPT tabs already exist - skipping');
          skipCount++;
          skipped.push(project + ' - ' + modul + ' (already has new VAPT tabs)');
          continue;
        }

        // Check and delete OLD VAPT tabs if they exist
        const oldTabs = [
          'Detail Finding - Regular VAPT',
          'Evidence - Regular VAPT',
          'Detail Finding - Ad Hoc VAPT',
          'Evidence - Ad Hoc VAPT',
          'Detail Finding - VAPT',
          'Evidence - VAPT'
        ];

        let deletedCount = 0;
        oldTabs.forEach(tabName => {
          const oldTab = qatmSs.getSheetByName(tabName);
          if (oldTab) {
            Logger.log('   🗑️  Deleting old tab: ' + tabName);
            qatmSs.deleteSheet(oldTab);
            deletedCount++;
          }
        });

        if (deletedCount > 0) {
          Logger.log('   ⚠️  Deleted ' + deletedCount + ' old VAPT tab(s) - creating new structure');
          SpreadsheetApp.flush();
          Utilities.sleep(500);
        }

        if (newTabsExist > 0 && newTabsExist < 3) {
          Logger.log('   ⚠️  Partial NEW VAPT tabs exist (' + newTabsExist + '/3) - will complete missing tabs');
        }

        // Create VAPT tabs
        Logger.log('   🔧 Creating VAPT tabs...');

        createVAPTTabsInQATM_(qatmSs);

        // Add VAPT summary section to Summary tab
        Logger.log('   📊 Adding VAPT summary section...');
        addVAPTSummarySection_(qatmSs, 35);

        successCount++;
        Logger.log('   ✅ SUCCESS: ' + project + ' - ' + modul);

        // Sleep to avoid rate limiting
        Utilities.sleep(1000);

      } catch (e) {
        Logger.log('   ❌ ERROR: ' + e.message);
        errorCount++;
        errors.push(project + ' - ' + modul + ' (' + e.message + ')');
      }
    }

    Logger.log('\n══════════════════════════════════════════');
    Logger.log('📊 BROADCAST SUMMARY');
    Logger.log('══════════════════════════════════════════');
    Logger.log('✅ Success: ' + successCount + ' QATM(s)');
    Logger.log('⏭️  Skipped: ' + skipCount + ' QATM(s)');
    Logger.log('❌ Errors: ' + errorCount + ' QATM(s)');

    // Show results
    let msg = '✅ Broadcast Complete!\n\n';
    msg += '📊 Summary:\n';
    msg += '• ✅ Added: ' + successCount + ' QATM(s)\n';
    msg += '• ⏭️  Skipped: ' + skipCount + ' QATM(s) (already exist)\n';
    msg += '• ❌ Errors: ' + errorCount + ' QATM(s)\n';

    if (errors.length > 0) {
      msg += '\n❌ Errors:\n';
      errors.slice(0, 5).forEach(err => msg += '• ' + err + '\n');
      if (errors.length > 5) {
        msg += '• ... +' + (errors.length - 5) + ' more (check log)\n';
      }
    }

    ui.alert('🔒 VAPT Broadcast', msg, ui.ButtonSet.OK);
    Logger.log('✅ Broadcast completed successfully');

  } catch (e) {
    Logger.log('❌ Broadcast error: ' + e.message);
    ui.alert(
      '❌ Error',
      'Broadcast failed:\n\n' + e.message + '\n\n' +
      'Check Execution log for details.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Create VAPT tabs in a single QATM spreadsheet
 * Internal helper function
 *
 * Creates 3 NEW VAPT tabs (only missing ones)
 */
function createVAPTTabsInQATM_(ss) {
  // Import functions from InitVAPTTabs.js
  // Note: In Apps Script, all .gs files are globally scoped,
  // so we can call functions from InitVAPTTabs.js directly

  try {
    // Check which NEW tabs are missing
    const newTabsConfig = [
      {name: 'VAPT - Helper', createFn: createVAPTHelper},
      {name: 'VAPT - Detail Finding', createFn: createDetailFindingVAPT},
      {name: 'VAPT - Evidence', createFn: createEvidenceVAPT}
    ];

    let createdCount = 0;

    newTabsConfig.forEach(tabConfig => {
      const existingTab = ss.getSheetByName(tabConfig.name);

      if (!existingTab) {
        Logger.log('     📝 Creating ' + tabConfig.name + '...');
        tabConfig.createFn(ss);
        SpreadsheetApp.flush();
        Utilities.sleep(500);
        createdCount++;
      } else {
        Logger.log('     ⏭️  ' + tabConfig.name + ' already exists - skipping');
      }
    });

    if (createdCount > 0) {
      Logger.log('     ✅ Created ' + createdCount + ' VAPT tab(s) successfully');
    } else {
      Logger.log('     ℹ️  No new tabs needed - all tabs already exist');
    }

  } catch (e) {
    throw new Error('Failed to create VAPT tabs: ' + e.message);
  }
}

/**
 * Add VAPT Summary Section to Summary tab
 * Called after VAPT tabs are created
 *
 * @param {Spreadsheet} ss - The QATM spreadsheet
 * @param {number} startRow - Row number to start VAPT section (default: 35)
 */
function addVAPTSummarySection_(ss, startRow) {
  const summarySheet = ss.getSheetByName('Summary');

  if (!summarySheet) {
    Logger.log('     ⚠️  Summary tab not found - skipping VAPT summary');
    return;
  }

  // Default start row if not specified
  startRow = startRow || 35;

  // Colors
  const headerBg = '#263238';  // Dark grey
  const vaptBg = '#FF6F00';     // Orange for VAPT
  const sectionBg = '#E3F2FD';  // Light blue for section headers
  const white = '#FFFFFF';

  try {
    Logger.log('     📊 Adding VAPT summary section at row ' + startRow);

    // ═══════════════════════════════════════════════════════════════
    // SECTION TITLE
    // ═══════════════════════════════════════════════════════════════

    const titleRange = summarySheet.getRange(startRow, 1, 1, 3);
    titleRange.merge()
      .setValue('🔒 VAPT FINDINGS SUMMARY')
      .setBackground(headerBg)
      .setFontColor(white)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    summarySheet.setRowHeight(startRow, 35);

    // ═══════════════════════════════════════════════════════════════
    // OVERVIEW - Total Findings
    // ═══════════════════════════════════════════════════════════════

    let currentRow = startRow + 2;

    // Overview header
    summarySheet.getRange(currentRow, 1, 1, 2).merge()
      .setValue('Overview')
      .setBackground(sectionBg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // Total findings
    summarySheet.getRange(currentRow, 1).setValue('Total VAPT Findings')
      .setFontWeight('bold');
    summarySheet.getRange(currentRow, 2)
      .setFormula('=COUNTA(\'VAPT - Detail Finding\'!A3:A1000)-COUNTBLANK(\'VAPT - Detail Finding\'!A3:A1000)')
      .setNumberFormat('0')
      .setHorizontalAlignment('center')
      .setBackground(vaptBg)
      .setFontColor(white)
      .setFontWeight('bold')
      .setFontSize(11);
    currentRow += 2;

    // ═══════════════════════════════════════════════════════════════
    // BY RISK LEVEL
    // ═══════════════════════════════════════════════════════════════

    summarySheet.getRange(currentRow, 1, 1, 2).merge()
      .setValue('By Risk Level (Adjusted Risk)')
      .setBackground(sectionBg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // Headers
    summarySheet.getRange(currentRow, 1).setValue('Risk Level').setFontWeight('bold');
    summarySheet.getRange(currentRow, 2).setValue('Count').setFontWeight('bold').setHorizontalAlignment('center');
    currentRow++;

    const riskLevels = [
      {level: 'Critical', color: '#FFEBEE'},
      {level: 'High', color: '#FFCDD2'},
      {level: 'Medium', color: '#FFF9C4'},
      {level: 'Low', color: '#FFF8E1'},
      {level: 'Informational', color: '#E3F2FD'}
    ];

    riskLevels.forEach(risk => {
      summarySheet.getRange(currentRow, 1).setValue(risk.level)
        .setBackground(risk.color)
        .setFontWeight('bold');

      // Count from VAPT - Detail Finding tab
      summarySheet.getRange(currentRow, 2)
        .setFormula('=COUNTIF(\'VAPT - Detail Finding\'!H:H,"' + risk.level + '")')
        .setNumberFormat('0')
        .setHorizontalAlignment('center');

      currentRow++;
    });

    currentRow++;

    // ═══════════════════════════════════════════════════════════════
    // BY STATUS FIX (DEV)
    // ═══════════════════════════════════════════════════════════════

    summarySheet.getRange(currentRow, 1, 1, 2).merge()
      .setValue('By Status Fix (Dev Team)')
      .setBackground(sectionBg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // Headers
    summarySheet.getRange(currentRow, 1).setValue('Status').setFontWeight('bold');
    summarySheet.getRange(currentRow, 2).setValue('Count').setFontWeight('bold').setHorizontalAlignment('center');
    currentRow++;

    const statusFix = ['Todo', 'On Progress Remediation', 'Ready to Retest', 'Done', 'Accepted', 'False Positive'];

    statusFix.forEach(status => {
      summarySheet.getRange(currentRow, 1).setValue(status);

      // Count from VAPT - Detail Finding tab
      summarySheet.getRange(currentRow, 2)
        .setFormula('=COUNTIF(\'VAPT - Detail Finding\'!E:E,"' + status + '")')
        .setNumberFormat('0')
        .setHorizontalAlignment('center');

      currentRow++;
    });

    currentRow++;

    // ═══════════════════════════════════════════════════════════════
    // BY STATUS RE-VAPT (PENTESTER)
    // ═══════════════════════════════════════════════════════════════

    summarySheet.getRange(currentRow, 1, 1, 2).merge()
      .setValue('By Status Re-VAPT (Pentester)')
      .setBackground(sectionBg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // Headers
    summarySheet.getRange(currentRow, 1).setValue('Status').setFontWeight('bold');
    summarySheet.getRange(currentRow, 2).setValue('Count').setFontWeight('bold').setHorizontalAlignment('center');
    currentRow++;

    // Open
    summarySheet.getRange(currentRow, 1).setValue('Open')
      .setBackground('#FFEBEE')
      .setFontWeight('bold');
    summarySheet.getRange(currentRow, 2)
      .setFormula('=COUNTIF(\'VAPT - Detail Finding\'!F:F,"Open")')
      .setNumberFormat('0')
      .setHorizontalAlignment('center');
    currentRow++;

    // Closed
    summarySheet.getRange(currentRow, 1).setValue('Closed')
      .setBackground('#E8F5E9')
      .setFontWeight('bold');
    summarySheet.getRange(currentRow, 2)
      .setFormula('=COUNTIF(\'VAPT - Detail Finding\'!F:F,"Closed")')
      .setNumberFormat('0')
      .setHorizontalAlignment('center');
    currentRow++;

    // ═══════════════════════════════════════════════════════════════
    // SET COLUMN WIDTHS
    // ═══════════════════════════════════════════════════════════════

    summarySheet.setColumnWidth(1, 200);  // Labels
    summarySheet.setColumnWidth(2, 80);   // Count

    // Add borders to the entire VAPT section
    const sectionRange = summarySheet.getRange(startRow, 1, currentRow - startRow, 2);
    sectionRange.setBorder(
      true, true, true, true, true, true,
      '#CFD8DC', SpreadsheetApp.BorderStyle.SOLID
    );

    Logger.log('     ✅ VAPT summary section added successfully');

  } catch (e) {
    Logger.log('     ❌ Error adding VAPT summary: ' + e.message);
    // Don't throw - summary is optional, tabs are more important
  }
}
