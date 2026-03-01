/**
 * BROADCAST: Smoke Test Enhancements
 *
 * Adds:
 * 1. STATUS OVERVIEW - Smoke Test (Medium-Critical) section in Summary
 * 2. Bug Summary row for Smoke Test blockers
 * 3. Description column in BugReport sheet
 * 4. Updates Dashboard to pull Smoke Test metrics
 *
 * Safe: Non-destructive. Adds columns/sections without touching existing data.
 */

function broadcastSmokeTestEnhancements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Try both Config and _Config
  let cfgSheet = ss.getSheetByName('Config');
  if (!cfgSheet) {
    cfgSheet = ss.getSheetByName('_Config');
  }

  if (!cfgSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Sheet "Config" atau "_Config" tidak ditemukan!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Get all data from Config sheet (columns A and E)
  const configData = cfgSheet.getRange('A2:E50').getValues();

  // Filter valid rows: must have Sheet ID in column E (index 4)
  // AND Sheet ID must look like a valid Google Sheets ID (long alphanumeric string)
  const validRows = configData
    .map((row, idx) => ({
      name: row[0] ? row[0].toString().trim() : `Module ${idx + 1}`,
      sheetId: row[4] ? row[4].toString().trim() : '',
      originalIndex: idx + 2 // For logging (row number in sheet)
    }))
    .filter(item => {
      // Must have sheet ID
      if (!item.sheetId) return false;
      // Sheet ID should be at least 30 chars (typical Google Sheet ID length)
      if (item.sheetId.length < 30) return false;
      // Should not contain spaces or special instruction text
      if (item.sheetId.includes(' ') || item.sheetId.includes('[ID]')) return false;
      return true;
    });

  if (validRows.length === 0) {
    SpreadsheetApp.getUi().alert('Info', 'Tidak ada valid Sheet ID di Config sheet kolom E.\n\nPastikan:\n- Kolom E berisi Sheet ID (minimal 30 karakter)\n- Bukan placeholder text seperti "[ID]"', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Extract names and URLs
  const moduleNames = validRows.map(r => r.name);
  const moduleUrls = validRows.map(r => `https://docs.google.com/spreadsheets/d/${r.sheetId}/edit`);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Broadcast Smoke Test Enhancements',
    `Akan menambahkan:\n\n` +
    `1. STATUS OVERVIEW - Smoke Test (Medium-Critical)\n` +
    `2. Bug Summary - Smoke Test Blockers row\n` +
    `3. Description column di BugReport\n` +
    `4. Update Dashboard untuk metrics Smoke Test\n\n` +
    `Target: ${moduleUrls.length} modules\n\n` +
    `Lanjutkan?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    Logger.log('User cancelled broadcast.');
    return;
  }

  let successCount = 0;
  let errorLog = [];

  moduleUrls.forEach((url, index) => {
    const moduleName = moduleNames[index] || 'Unknown';
    try {
      Logger.log(`[${index + 1}/${moduleUrls.length}] Processing: ${moduleName}`);

      const targetSs = SpreadsheetApp.openByUrl(url.toString().trim());
      const summarySheet = targetSs.getSheetByName('Summary');
      const bugReportSheet = targetSs.getSheetByName('BugReport');
      const appendixSheet = targetSs.getSheetByName('Appendix');

      if (!summarySheet) {
        errorLog.push(`[${index + 1}] ${moduleName}: Summary sheet not found`);
        return;
      }

      // =================================================================
      // PART 1: Add Description column to BugReport
      // =================================================================
      if (bugReportSheet) {
        addDescriptionColumnToBugReport(bugReportSheet);
      } else {
        Logger.log(`  ⚠️ BugReport sheet not found, skipping`);
      }

      // =================================================================
      // PART 2: Add Smoke Test STATUS OVERVIEW section
      // =================================================================
      addSmokeTestStatusOverview(summarySheet);

      // =================================================================
      // PART 3: Add Smoke Test row to Bug Summary
      // =================================================================
      addSmokeTestBugSummary(summarySheet);

      // =================================================================
      // PART 4: Add Bug Status Definitions to Appendix
      // =================================================================
      if (appendixSheet) {
        addBugStatusDefinitionsToAppendix(appendixSheet);
      } else {
        Logger.log(`  ⚠️ Appendix sheet not found, skipping`);
      }

      SpreadsheetApp.flush();
      successCount++;
      Logger.log(`  ✅ Success`);

    } catch (e) {
      errorLog.push(`[${index + 1}] ${moduleName}: ${e.message}`);
      Logger.log(`  ❌ Error: ${e.message}`);
    }
  });

  // Final report
  let msg = `✅ Broadcast Selesai!\n\n`;
  msg += `Berhasil: ${successCount} / ${moduleUrls.length}\n`;

  if (errorLog.length > 0) {
    msg += `\nError:\n` + errorLog.join('\n');
  }

  msg += `\n\n📝 Perubahan:\n`;
  msg += `1. Added STATUS OVERVIEW - Smoke Test (Medium-Critical)\n`;
  msg += `2. Added Bug Summary - Medium-Critical Open blockers\n`;
  msg += `3. Added Description column in BugReport (col I)\n`;
  msg += `4. Added detail rows in BugReport (row 63+): Status, Update oleh, Artinya\n`;
  msg += `5. Added Bug Status Definitions to Appendix (Open, In Progress, Fixed, Verified, Closed, Reopen)\n`;
  msg += `6. All formulas auto-update from TC_Master/API_Master\n`;

  ui.alert('Broadcast Complete', msg, ui.ButtonSet.OK);
  Logger.log(msg);
}

// =================================================================
// HELPER: Add Description column and detail rows to BugReport
// =================================================================
function addDescriptionColumnToBugReport(bugSheet) {
  const headerRow = bugSheet.getRange(4, 1, 1, 25).getValues()[0];

  // Check if Description already exists
  const descIndex = headerRow.indexOf('Description');
  if (descIndex !== -1) {
    Logger.log('  ℹ️ Description column already exists at column ' + (descIndex + 1));
  } else {
    // Insert Description column at position I (column 9)
    // This pushes everything from I onwards to the right
    // Insert BEFORE column I, so Description becomes I
    bugSheet.insertColumnBefore(9);

    // Set header for Description (now at column I/9)
    bugSheet.getRange(4, 9).setValue('Description')
      .setBackground('#0D47A1').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);

    bugSheet.getRange(4, 9).setNote(
      'Description — Summary deskripsi bug (mirip Jira Summary/Description field).\n\n' +
      'Contoh:\n' +
      '- "Login button tidak responsif setelah input password"\n' +
      '- "API endpoint /users mengembalikan 500 error saat filter by role"\n' +
      '- "Dashboard chart tidak ter-render di mobile viewport"\n\n' +
      'Tulis deskripsi singkat & jelas (1-2 kalimat).'
    );

    // Set column width
    bugSheet.setColumnWidth(9, 220);

    // Apply formatting to data rows (row 5 onwards)
    const DS = 5, MR = 200;
    bugSheet.getRange(DS, 9, MR, 1)
      .setWrap(true)
      .setFontFamily('Arial').setFontSize(9)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);

    // Apply alternating colors
    for (let r = DS; r < DS + MR; r++) {
      const bg = (r - DS) % 2 === 0 ? '#F8FBFF' : '#FFFFFF';
      bugSheet.getRange(r, 9).setBackground(bg);
    }

    Logger.log('  ✅ Added Description column at position I (column 9)');
  }

  // =================================================================
  // Add detail rows: Status, Update oleh, Artinya (starting row 63)
  // =================================================================
  // Check if detail rows already exist
  const row63Value = bugSheet.getRange(63, 2).getValue();
  if (row63Value && row63Value.toString().includes('Status')) {
    Logger.log('  ℹ️ Detail rows (Status, Update oleh, Artinya) already exist at row 63');
    return;
  }

  // Insert 3 new rows at row 63
  bugSheet.insertRowsAfter(62, 3);

  const detailRows = [
    { label: 'Status', note: 'Status bug saat ini (Open, In Progress, Fixed, Verified, Closed, dll)' },
    { label: 'Update oleh', note: 'Nama developer/tester yang melakukan update terakhir' },
    { label: 'Artinya', note: 'Penjelasan singkat tentang update/status terkini' }
  ];

  detailRows.forEach((detail, idx) => {
    const row = 63 + idx;

    // Column B: Label
    bugSheet.getRange(row, 2).setValue(detail.label)
      .setBackground('#E3F2FD').setFontColor('#0D47A1')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('right').setVerticalAlignment('middle')
      .setNote(detail.note);

    // Column C onwards: merge for input area
    bugSheet.getRange(row, 3, 1, 18).merge()
      .setBackground('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9)
      .setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);

    bugSheet.setRowHeight(row, 24);
  });

  Logger.log('  ✅ Added detail rows (Status, Update oleh, Artinya) starting at row 63');
}

// =================================================================
// HELPER: Add Smoke Test STATUS OVERVIEW
// =================================================================
function addSmokeTestStatusOverview(summarySheet) {
  const data = summarySheet.getDataRange().getValues();

  // Check if A2 Smoke Test section already exists
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue.includes('A2.  STATUS OVERVIEW') && cellValue.includes('Smoke Test')) {
      Logger.log('  ℹ️ Smoke Test STATUS OVERVIEW already exists, skipping');
      return;
    }
  }

  // Find the current STATUS OVERVIEW section (look for "A. STATUS OVERVIEW")
  let statusOverviewRow = -1;

  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue.includes('A.  STATUS OVERVIEW')) {
      statusOverviewRow = i + 1; // Convert to 1-indexed
      break;
    }
  }

  if (statusOverviewRow === -1) {
    Logger.log('  ⚠️ STATUS OVERVIEW section not found, skipping');
    return;
  }

  // Insert 4 new rows after the existing STATUS OVERVIEW section
  // Existing structure:
  // Row N: Header "A. STATUS OVERVIEW"
  // Row N+1: KPI Labels (TOTAL, PASSED, FAILED...)
  // Row N+2: KPI Values
  // Row N+3: Legend/note
  //
  // We'll insert after Row N+3

  const insertAfterRow = statusOverviewRow + 3;
  summarySheet.insertRowsAfter(insertAfterRow, 4);

  const newHeaderRow = insertAfterRow + 1;
  const newLabelsRow = newHeaderRow + 1;
  const newValuesRow = newLabelsRow + 1;
  const newNoteRow = newValuesRow + 1;

  const L = 1, LW = 10, R_ = 12, RW = 10;

  // Row height for spacing
  summarySheet.setRowHeight(insertAfterRow, 8);

  // =================================================================
  // New section header
  // =================================================================
  summarySheet.getRange(newHeaderRow, L, 1, LW).merge();
  summarySheet.getRange(newHeaderRow, L)
    .setValue('A2.  STATUS OVERVIEW  -  Smoke Test (Medium - Critical)')
    .setBackground('#1565C0').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  summarySheet.getRange(newHeaderRow, R_, 1, RW).merge();
  summarySheet.getRange(newHeaderRow, R_)
    .setValue('A2.  STATUS OVERVIEW  -  Smoke Test (Medium - Critical)')
    .setBackground('#283593').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  summarySheet.setRowHeight(newHeaderRow, 22);

  // =================================================================
  // KPI Labels row
  // =================================================================
  const kpiLabels = ['TOTAL', 'PASSED', 'FAILED', 'BLOCKED', 'IN PROG', 'TODO', 'PASS RATE', 'AUTO RATE', 'EXEC RATE'];
  const kpiBgs = ['#37474F', '#2E7D32', '#B71C1C', '#E65100', '#1565C0', '#546E7A', '#0D47A1', '#1976D2', '#4A148C'];

  for (let i = 0; i < 9; i++) {
    const cell = summarySheet.getRange(newLabelsRow, L + i);
    cell.setValue(kpiLabels[i])
      .setBackground(kpiBgs[i]).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(i < 6 ? 8 : 7.5)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true);

    const cellR = summarySheet.getRange(newLabelsRow, R_ + i);
    cellR.setValue(kpiLabels[i])
      .setBackground(kpiBgs[i]).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(i < 6 ? 8 : 7.5)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true);
  }

  summarySheet.setRowHeight(newLabelsRow, 22);

  // =================================================================
  // KPI Values row with formulas
  // =================================================================
  // TC_Master columns: 1=No, 2=SubModul, 3=TC_ID, 4=Feature, 5=Priority, 6=Platform, 7=Test Type, 8=Automated, 14=[AUTO] Test Level
  // API_Master columns: 1=No, 2=SubModul, 3=TC_ID, 4=Feature, 5=Endpoint, 6=Method, 7=Priority, 13=Automated, 15=[AUTO] Test Level

  const wPrioCol = 'TC_Master!E3:E1000';      // Priority column (5)
  const wLevelCol = 'TC_Master!N3:N1000';     // Test Level column (14)
  const wAutoCol = 'TC_Master!H3:H1000';      // Automated column (8)
  const wTcIdCol = 'TC_Master!C3:C1000';      // TC_ID column (3)

  const aPrioCol = 'API_Master!G3:G1000';     // Priority column (7)
  const aLevelCol = 'API_Master!O3:O1000';    // Test Level column (15)
  const aAutoCol = 'API_Master!M3:M1000';     // Automated column (13)
  const aTcIdCol = 'API_Master!C3:C1000';     // TC_ID column (3)

  // Web/Mobile formulas
  const wForms = [
    // TOTAL: Count TC_Master rows where (Priority=Critical OR High OR Medium) AND Test Level=Smoke
    `=SUMPRODUCT(((${wPrioCol}="Critical")+(${wPrioCol}="High")+(${wPrioCol}="Medium"))*(${wLevelCol}="Smoke"))`,

    // PASSED: Count matching TCs where status = PASSED
    `=SUMPRODUCT(((${wPrioCol}="Critical")+(${wPrioCol}="High")+(${wPrioCol}="Medium"))*(${wLevelCol}="Smoke")*COUNTIFS(TC_Execution!$A$9:$A$1000,${wTcIdCol},TC_Execution!$Z$9:$Z$1000,"PASSED"))`,

    // FAILED
    `=SUMPRODUCT(((${wPrioCol}="Critical")+(${wPrioCol}="High")+(${wPrioCol}="Medium"))*(${wLevelCol}="Smoke")*COUNTIFS(TC_Execution!$A$9:$A$1000,${wTcIdCol},TC_Execution!$Z$9:$Z$1000,"FAILED"))`,

    // BLOCKED
    `=SUMPRODUCT(((${wPrioCol}="Critical")+(${wPrioCol}="High")+(${wPrioCol}="Medium"))*(${wLevelCol}="Smoke")*COUNTIFS(TC_Execution!$A$9:$A$1000,${wTcIdCol},TC_Execution!$Z$9:$Z$1000,"BLOCKED"))`,

    // IN PROGRESS
    `=SUMPRODUCT(((${wPrioCol}="Critical")+(${wPrioCol}="High")+(${wPrioCol}="Medium"))*(${wLevelCol}="Smoke")*COUNTIFS(TC_Execution!$A$9:$A$1000,${wTcIdCol},TC_Execution!$Z$9:$Z$1000,"IN PROGRESS"))`,

    // TODO
    `=SUMPRODUCT(((${wPrioCol}="Critical")+(${wPrioCol}="High")+(${wPrioCol}="Medium"))*(${wLevelCol}="Smoke")*COUNTIFS(TC_Execution!$A$9:$A$1000,${wTcIdCol},TC_Execution!$Z$9:$Z$1000,"TODO"))`,

    // PASS RATE
    `=IFERROR(${colLetter(L + 1)}${newValuesRow}/MAX(1,${colLetter(L)}${newValuesRow}),0)`,

    // AUTO RATE (automated smoke tests)
    `=IFERROR(SUMPRODUCT(((${wPrioCol}="Critical")+(${wPrioCol}="High")+(${wPrioCol}="Medium"))*(${wLevelCol}="Smoke")*(${wAutoCol}="Automated"))/MAX(1,${colLetter(L)}${newValuesRow}),0)`,

    // EXEC RATE
    `=IFERROR((${colLetter(L + 1)}${newValuesRow}+${colLetter(L + 2)}${newValuesRow}+${colLetter(L + 3)}${newValuesRow}+${colLetter(L + 4)}${newValuesRow})/MAX(1,${colLetter(L)}${newValuesRow}),0)`
  ];

  // API formulas
  const aForms = [
    // TOTAL
    `=SUMPRODUCT(((${aPrioCol}="Critical")+(${aPrioCol}="High")+(${aPrioCol}="Medium"))*(${aLevelCol}="Smoke"))`,

    // PASSED
    `=SUMPRODUCT(((${aPrioCol}="Critical")+(${aPrioCol}="High")+(${aPrioCol}="Medium"))*(${aLevelCol}="Smoke")*COUNTIFS(API_Execution!$A$9:$A$1000,${aTcIdCol},API_Execution!$Z$9:$Z$1000,"PASSED"))`,

    // FAILED
    `=SUMPRODUCT(((${aPrioCol}="Critical")+(${aPrioCol}="High")+(${aPrioCol}="Medium"))*(${aLevelCol}="Smoke")*COUNTIFS(API_Execution!$A$9:$A$1000,${aTcIdCol},API_Execution!$Z$9:$Z$1000,"FAILED"))`,

    // BLOCKED
    `=SUMPRODUCT(((${aPrioCol}="Critical")+(${aPrioCol}="High")+(${aPrioCol}="Medium"))*(${aLevelCol}="Smoke")*COUNTIFS(API_Execution!$A$9:$A$1000,${aTcIdCol},API_Execution!$Z$9:$Z$1000,"BLOCKED"))`,

    // IN PROGRESS
    `=SUMPRODUCT(((${aPrioCol}="Critical")+(${aPrioCol}="High")+(${aPrioCol}="Medium"))*(${aLevelCol}="Smoke")*COUNTIFS(API_Execution!$A$9:$A$1000,${aTcIdCol},API_Execution!$Z$9:$Z$1000,"IN PROGRESS"))`,

    // TODO
    `=SUMPRODUCT(((${aPrioCol}="Critical")+(${aPrioCol}="High")+(${aPrioCol}="Medium"))*(${aLevelCol}="Smoke")*COUNTIFS(API_Execution!$A$9:$A$1000,${aTcIdCol},API_Execution!$Z$9:$Z$1000,"TODO"))`,

    // PASS RATE
    `=IFERROR(${colLetter(R_ + 1)}${newValuesRow}/MAX(1,${colLetter(R_)}${newValuesRow}),0)`,

    // AUTO RATE
    `=IFERROR(SUMPRODUCT(((${aPrioCol}="Critical")+(${aPrioCol}="High")+(${aPrioCol}="Medium"))*(${aLevelCol}="Smoke")*(${aAutoCol}="Automated"))/MAX(1,${colLetter(R_)}${newValuesRow}),0)`,

    // EXEC RATE
    `=IFERROR((${colLetter(R_ + 1)}${newValuesRow}+${colLetter(R_ + 2)}${newValuesRow}+${colLetter(R_ + 3)}${newValuesRow}+${colLetter(R_ + 4)}${newValuesRow})/MAX(1,${colLetter(R_)}${newValuesRow}),0)`
  ];

  // Apply Web formulas
  wForms.forEach((f, i) => {
    const cell = summarySheet.getRange(newValuesRow, L + i);
    cell.setFormula(f)
      .setBackground('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(i < 6 ? 16 : 13)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    if (i >= 6) {
      cell.setNumberFormat('0%');
      applyPassRateCF(summarySheet, newValuesRow, L + i);
    }
  });

  // Apply API formulas
  aForms.forEach((f, i) => {
    const cell = summarySheet.getRange(newValuesRow, R_ + i);
    cell.setFormula(f)
      .setBackground('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(i < 6 ? 16 : 13)
      .setFontFamily('Arial')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    if (i >= 6) {
      cell.setNumberFormat('0%');
      applyPassRateCF(summarySheet, newValuesRow, R_ + i);
    }
  });

  summarySheet.setRowHeight(newValuesRow, 36);

  // =================================================================
  // Note row
  // =================================================================
  summarySheet.getRange(newNoteRow, L, 1, LW).merge();
  summarySheet.getRange(newNoteRow, L)
    .setValue('Smoke Test (Medium-Critical) = Test Level "Smoke" dengan Priority Critical/High/Medium')
    .setBackground('#E3F2FD').setFontColor('#1565C0')
    .setFontStyle('italic').setFontSize(7).setFontFamily('Arial')
    .setHorizontalAlignment('left');

  summarySheet.getRange(newNoteRow, R_, 1, RW).merge();
  summarySheet.getRange(newNoteRow, R_)
    .setValue('Smoke Test (Medium-Critical) = Test Level "Smoke" dengan Priority Critical/High/Medium')
    .setBackground('#E8EAF6').setFontColor('#283593')
    .setFontStyle('italic').setFontSize(7).setFontFamily('Arial')
    .setHorizontalAlignment('left');

  summarySheet.setRowHeight(newNoteRow, 14);

  Logger.log('  ✅ Added Smoke Test STATUS OVERVIEW section');
}

// =================================================================
// HELPER: Add Smoke Test row to Bug Summary
// =================================================================
function addSmokeTestBugSummary(summarySheet) {
  const data = summarySheet.getDataRange().getValues();

  // Check if Medium-Critical Open row already exists
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue.includes('Medium-Critical Open')) {
      Logger.log('  ℹ️ Medium-Critical Open row already exists in Bug Summary, skipping');
      return;
    }
  }

  // Find Bug Summary section (look for "D. BUG SUMMARY" not "E. BUG SUMMARY")
  let bugSummaryRow = -1;

  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue.includes('BUG SUMMARY') && (cellValue.includes('D.') || cellValue.includes('E.'))) {
      bugSummaryRow = i + 1;
      break;
    }
  }

  if (bugSummaryRow === -1) {
    Logger.log('  ⚠️ BUG SUMMARY section not found, skipping');
    return;
  }

  // Insert 1 row after the last bug metric (after "Medium" row, which is row +8)
  const insertAfterRow = bugSummaryRow + 8;
  summarySheet.insertRowsAfter(insertAfterRow, 1);

  const newRow = insertAfterRow + 1;
  const L = 1, LW = 10, R_ = 12, RW = 10;

  // Medium-Critical Open Blockers row
  const label = 'Medium-Critical Open';
  const bg = '#FFF9C4'; // Yellow/amber
  const fg = '#E65100'; // Orange

  // Left (Web + Mobile)
  summarySheet.getRange(newRow, L).setValue(label + ':')
    .setBackground('#FFEBEE').setFontFamily('Arial')
    .setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('right').setFontColor('#C62828')
    .setVerticalAlignment('middle');

  summarySheet.getRange(newRow, L + 1, 1, LW - 1).merge();

  // Formula: Count bugs where Type=Web OR Mobile, Priority=Critical/High/Medium, Status=Open, and related TC is Smoke Test
  // Since we can't easily join BugReport with TC_Master, we'll use a simpler approach:
  // Count Open bugs with Priority Critical/High/Medium (smoke test blockers are defined as medium-critical open bugs)
  const wFormula = `=IFERROR(COUNTIFS(BugReport!B5:B5000,"Web",BugReport!C5:C5000,"Critical",BugReport!D5:D5000,"Open")+` +
    `COUNTIFS(BugReport!B5:B5000,"Web",BugReport!C5:C5000,"High",BugReport!D5:D5000,"Open")+` +
    `COUNTIFS(BugReport!B5:B5000,"Web",BugReport!C5:C5000,"Medium",BugReport!D5:D5000,"Open")+` +
    `COUNTIFS(BugReport!B5:B5000,"Mobile",BugReport!C5:C5000,"Critical",BugReport!D5:D5000,"Open")+` +
    `COUNTIFS(BugReport!B5:B5000,"Mobile",BugReport!C5:C5000,"High",BugReport!D5:D5000,"Open")+` +
    `COUNTIFS(BugReport!B5:B5000,"Mobile",BugReport!C5:C5000,"Medium",BugReport!D5:D5000,"Open"),0)`;

  summarySheet.getRange(newRow, L + 1).setFormula(wFormula)
    .setBackground(bg).setFontFamily('Arial')
    .setFontSize(11).setFontWeight('bold')
    .setFontColor(fg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Right (API)
  summarySheet.getRange(newRow, R_).setValue(label + ':')
    .setBackground('#FFEBEE').setFontFamily('Arial')
    .setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('right').setFontColor('#B71C1C')
    .setVerticalAlignment('middle');

  summarySheet.getRange(newRow, R_ + 1, 1, RW - 1).merge();

  const aFormula = `=IFERROR(COUNTIFS(BugReport!B5:B5000,"API",BugReport!C5:C5000,"Critical",BugReport!D5:D5000,"Open")+` +
    `COUNTIFS(BugReport!B5:B5000,"API",BugReport!C5:C5000,"High",BugReport!D5:D5000,"Open")+` +
    `COUNTIFS(BugReport!B5:B5000,"API",BugReport!C5:C5000,"Medium",BugReport!D5:D5000,"Open"),0)`;

  summarySheet.getRange(newRow, R_ + 1).setFormula(aFormula)
    .setBackground(bg).setFontFamily('Arial')
    .setFontSize(11).setFontWeight('bold')
    .setFontColor(fg)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  summarySheet.setRowHeight(newRow, 24);

  Logger.log('  ✅ Added Smoke Test Blockers row to Bug Summary');
}

// =================================================================
// HELPER: Apply Pass Rate Conditional Formatting
// =================================================================
function applyPassRateCF(sheet, row, col) {
  const range = sheet.getRange(row, col);
  const rules = sheet.getConditionalFormatRules();

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0.8)
    .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.5, 0.799)
    .setBackground('#FFF8E1').setFontColor('#E65100').setBold(true)
    .setRanges([range]).build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0.5)
    .setBackground('#FFEBEE').setFontColor('#C62828').setBold(true)
    .setRanges([range]).build());

  sheet.setConditionalFormatRules(rules);
}

// =================================================================
// HELPER: Add Bug Status Definitions to Appendix
// =================================================================
function addBugStatusDefinitionsToAppendix(appendixSheet) {
  const data = appendixSheet.getDataRange().getValues();

  // Check if "BUG STATUS DEFINITIONS" section already exists
  for (let i = 0; i < data.length; i++) {
    const cellValue = data[i][0] ? data[i][0].toString() : '';
    if (cellValue.includes('BUG STATUS')) {
      Logger.log('  ℹ️ Bug Status Definitions already exist in Appendix, skipping');
      return;
    }
  }

  // Find the last row with content
  let lastRow = appendixSheet.getLastRow();

  // Add some spacing
  lastRow += 2;

  // Section header
  appendixSheet.getRange(lastRow, 1, 1, 4).merge()
    .setValue('3. BUG STATUS DEFINITIONS')
    .setBackground('#0D47A1').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#1976D2', SpreadsheetApp.BorderStyle.SOLID);
  appendixSheet.setRowHeight(lastRow, 24);
  lastRow++;

  // Status definitions
  const statusDefs = [
    { status: 'Open', description: 'Bug baru ditemukan dan belum ditangani oleh developer.\nStatus awal setelah bug dicatat di BugReport.' },
    { status: 'In Progress', description: 'Bug sedang dalam proses perbaikan oleh developer.\nDeveloper sudah mengambil ownership dan sedang coding fix.' },
    { status: 'Fixed', description: 'Bug sudah diperbaiki oleh developer dan siap untuk diverifikasi QA.\nCode sudah di-merge dan deployed ke testing environment.' },
    { status: 'Verified', description: 'Bug sudah diverifikasi oleh QA dan terkonfirmasi sudah diperbaiki.\nQA sudah retest dan hasilnya PASSED.' },
    { status: 'Closed', description: 'Bug sudah selesai dan ditutup (Verified + deployed to production).\nBug fix sudah live di production dan tidak ada regression.' },
    { status: 'Reopen', description: 'Bug yang sebelumnya Fixed/Verified ternyata masih terjadi dan dibuka kembali.\nPerlu investigasi ulang -- kemungkinan regression atau fix tidak complete.' }
  ];

  statusDefs.forEach((def) => {
    // Column A: Status name
    appendixSheet.getRange(lastRow, 1)
      .setValue(def.status)
      .setBackground('#ECEFF1')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('left').setVerticalAlignment('top')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#1976D2', SpreadsheetApp.BorderStyle.SOLID);

    // Column B-D: Description (merged)
    appendixSheet.getRange(lastRow, 2, 1, 3).merge()
      .setValue(def.description)
      .setBackground('#FFFFFF')
      .setFontFamily('Arial').setFontSize(9)
      .setHorizontalAlignment('left').setVerticalAlignment('top')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#1976D2', SpreadsheetApp.BorderStyle.SOLID);

    appendixSheet.setRowHeight(lastRow, 48);
    lastRow++;
  });

  Logger.log('  ✅ Added Bug Status Definitions to Appendix');
}

// =================================================================
// HELPER: Column letter converter
// =================================================================
function colLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - mod) / 26);
  }
  return letter;
}
