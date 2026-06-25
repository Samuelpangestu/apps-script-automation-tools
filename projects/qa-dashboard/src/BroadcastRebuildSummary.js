/**
 * BroadcastRebuildSummary.js
 *
 * Rebuild Summary tabs in all QATM with latest template (includes VAPT Blocker Breakdown)
 * Preserves user input fields (Project, Module, etc.)
 */

function broadcastRebuildSummary() {
  const dashboardSs = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const modules = getModuleList_(dashboardSs);
  const targets = [];
  const seen = {};

  modules.forEach(mod => {
    if (!mod || !mod.id || seen[mod.id]) return;
    seen[mod.id] = true;
    targets.push(mod);
  });

  if (targets.length === 0) {
    ui.alert('No active QATM targets found in Config.');
    return;
  }

  const response = ui.alert(
    'Broadcast: Rebuild Summary Tabs',
    'Will REBUILD Summary tab in ' + targets.length + ' QATM sheets with latest template.\n\n' +
    '⚠️ WARNING:\n' +
    '• Summary tab will be DELETED and RECREATED\n' +
    '• Manual edits in Summary will be LOST\n' +
    '• Input fields (Project/Module) will be RESTORED\n' +
    '• Formulas rebuilt from template\n\n' +
    '✅ What will be added:\n' +
    '• F. VAPT FINDINGS SUMMARY section\n' +
    '• VAPT Blocker Breakdown (Critical/High/Medium)\n' +
    '• Updated formulas to reference "VAPT - Detail Finding"\n\n' +
    '✅ Safe:\n' +
    '• Other tabs (TC, BugReport, VAPT) unchanged\n' +
    '• User inputs preserved\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let created = 0;
  let skipped = 0;
  let failed = 0;
  const errors = [];

  targets.forEach(mod => {
    try {
      const qatmSs = SpreadsheetApp.openById(mod.id);
      const existingSummary = qatmSs.getSheetByName('Summary');

      if (!existingSummary) {
        skipped++;
        return;
      }

      // Save input values
      let savedInputs = {};
      try {
        const data = existingSummary.getDataRange().getValues();
        // Find input rows (usually B2:B10 in Summary)
        savedInputs = {
          project: getValue_(data, 'Project:', 1),
          module: getValue_(data, 'Modul:', 1),
          submodule: getValue_(data, 'Submodul:', 1),
          qaLead: getValue_(data, 'QA Lead:', 1),
          picQA: getValue_(data, 'PIC QA:', 1),
          environment: getValue_(data, 'Environment:', 1),
          issueTracker: getValue_(data, 'Issue Tracker', 1),
          testStatus: getValue_(data, 'Test Status:', 1),
          scopeNotes: getValue_(data, 'Scope / Notes:', 1),
        };
      } catch (e) {
        Logger.log('Could not save inputs from ' + mod.id + ': ' + e.message);
      }

      // Call rebuild function from QATM
      rebuildSummaryInQATM_(qatmSs, savedInputs);
      created++;

    } catch (error) {
      failed++;
      errors.push((mod.project || '') + ' / ' + (mod.module || '') + ': ' + error.message);
      Logger.log('Rebuild Summary broadcast failed [' + (mod.id || '-') + ']: ' + error.stack);
    }
  });

  ui.alert(
    'Summary Rebuild Broadcast Complete',
    'Rebuilt: ' + created + '\n' +
    'Skipped (no Summary tab): ' + skipped + '\n' +
    'Failed: ' + failed +
    (errors.length ? '\n\nErrors:\n' + errors.slice(0, 8).join('\n') : '') +
    '\n\n✅ Summary tabs now include VAPT Blocker Breakdown section.',
    ui.ButtonSet.OK
  );
}

function getValue_(data, label, colOffset) {
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      if (String(data[i][j]).includes(label)) {
        return data[i][j + colOffset] || '';
      }
    }
  }
  return '';
}

function rebuildSummaryInQATM_(ss, savedInputs) {
  // Delete existing Summary
  const existingSummary = ss.getSheetByName('Summary');
  if (existingSummary) {
    ss.deleteSheet(existingSummary);
    SpreadsheetApp.flush();
    Utilities.sleep(300);
  }

  // Recreate Summary using inline version of createSummary
  createSummaryFromTemplate_(ss);

  // Restore inputs
  const newSummary = ss.getSheetByName('Summary');
  if (newSummary && savedInputs && savedInputs.project) {
    try {
      newSummary.getRange('B2').setValue(savedInputs.project);
      newSummary.getRange('B3').setValue(savedInputs.module);
      newSummary.getRange('B4').setValue(savedInputs.submodule);
      newSummary.getRange('B5').setValue(savedInputs.qaLead);
      newSummary.getRange('B6').setValue(savedInputs.picQA);
      newSummary.getRange('B7').setValue(savedInputs.environment);
      newSummary.getRange('B8').setValue(savedInputs.issueTracker);
      newSummary.getRange('B9').setValue(savedInputs.testStatus);
      newSummary.getRange('B10').setValue(savedInputs.scopeNotes);
    } catch (e) {
      Logger.log('Could not restore inputs: ' + e.message);
    }
  }
}

function createSummaryFromTemplate_(ss) {
  // Inline creation of Summary tab with latest VAPT template
  // Based on createSummary from MasterQATCM.js

  const ws = ss.getSheetByName('Summary');
  if (ws) {
    ss.deleteSheet(ws);
    SpreadsheetApp.flush();
    Utilities.sleep(300);
  }

  const newSummary = ss.insertSheet('Summary', 0);

  // Set column widths
  newSummary.setColumnWidth(1, 220); // A - Labels
  newSummary.setColumnWidth(2, 180); // B - Values
  newSummary.setColumnWidth(3, 180); // C - Additional
  newSummary.setColumnWidth(4, 180); // D - Additional
  newSummary.setColumnWidth(5, 180); // E - Additional

  let R = 1;
  const L = 1; // Column A
  const LW = 5; // Label width in columns

  // Helper functions
  const m_ = (r, c, rows, cols) => {
    if (rows > 1 || cols > 1) {
      newSummary.getRange(r, c, rows, cols).merge();
    }
  };

  const bd = (range) => {
    return range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  };

  const h_ = (range, bg) => {
    return range.setBackground(bg || '#263238')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  };

  const lbl = (r, c, text, bg) => {
    bd(newSummary.getRange(r, c))
      .setValue(text)
      .setBackground(bg || '#F5F5F5')
      .setFontFamily('Arial')
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
  };

  // A. PROJECT INFORMATION
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, LW);
  h_(newSummary.getRange(R, L), '#263238').setValue('A.  PROJECT INFORMATION');
  newSummary.setRowHeight(R, 20); R++;

  const projectFields = [
    ['Project:', ''],
    ['Modul:', ''],
    ['Submodul:', ''],
    ['QA Lead:', ''],
    ['PIC QA:', ''],
    ['Environment:', ''],
    ['Issue Tracker', ''],
    ['Test Status:', ''],
    ['Scope / Notes:', ''],
  ];

  projectFields.forEach(([label, value]) => {
    lbl(R, L, label, '#E3F2FD');
    m_(R, L + 1, 1, LW - 1);
    bd(newSummary.getRange(R, L + 1))
      .setValue(value)
      .setBackground('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(9)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    newSummary.setRowHeight(R, 18); R++;
  });

  // B. TEST EXECUTION SUMMARY
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, LW);
  h_(newSummary.getRange(R, L), '#1B5E20').setValue('B.  TEST EXECUTION SUMMARY');
  newSummary.setRowHeight(R, 20); R++;

  const testFields = [
    ['Total Test Cases', '=IFERROR(COUNTA(FILTER(TC!A:A,TC!A:A<>\"\",TC!A:A<>\"TC ID\")),0)'],
    ['Executed (Passed + Failed)', '=IFERROR(COUNTIF(TC!I:I,\"Passed\")+COUNTIF(TC!I:I,\"Failed\"),0)'],
    ['Not Executed (N/A + Blocked)', '=IFERROR(COUNTIF(TC!I:I,\"N/A\")+COUNTIF(TC!I:I,\"Blocked\"),0)'],
    ['Pass Rate (%)', '=IFERROR(IF(COUNTIF(TC!I:I,\"Passed\")+COUNTIF(TC!I:I,\"Failed\")=0,0,COUNTIF(TC!I:I,\"Passed\")/(COUNTIF(TC!I:I,\"Passed\")+COUNTIF(TC!I:I,\"Failed\"))*100),0)'],
  ];

  testFields.forEach(([label, formula]) => {
    lbl(R, L, label, '#E8F5E9');
    m_(R, L + 1, 1, LW - 1);
    const cell = bd(newSummary.getRange(R, L + 1))
      .setFormula(formula)
      .setBackground('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    if (label.includes('%')) {
      cell.setNumberFormat('0.00"%"');
    } else {
      cell.setNumberFormat('0');
    }
    newSummary.setRowHeight(R, 20); R++;
  });

  // C. BUG REPORT SUMMARY
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, LW);
  h_(newSummary.getRange(R, L), '#B71C1C').setValue('C.  BUG REPORT SUMMARY');
  newSummary.setRowHeight(R, 20); R++;

  const bugFields = [
    ['Total Bugs Reported', '=IFERROR(COUNTA(FILTER(BugReport!A:A,BugReport!A:A<>\"\",BugReport!A:A<>\"Bug ID\")),0)'],
    ['Open Bugs', '=IFERROR(COUNTIF(BugReport!G:G,\"Open\")+COUNTIF(BugReport!G:G,\"Reopened\"),0)'],
    ['In Progress / Resolved', '=IFERROR(COUNTIF(BugReport!G:G,\"In Progress\")+COUNTIF(BugReport!G:G,\"Resolved\"),0)'],
    ['Closed', '=IFERROR(COUNTIF(BugReport!G:G,\"Closed\"),0)'],
    ['Bug Blocker Count', '=IFERROR(SUMPRODUCT((BugReport!G:G<>\"Closed\")*((BugReport!F:F=\"Blocker\")+(BugReport!F:F=\"Critical\"))),0)'],
  ];

  bugFields.forEach(([label, formula]) => {
    lbl(R, L, label, '#FFEBEE');
    m_(R, L + 1, 1, LW - 1);
    bd(newSummary.getRange(R, L + 1))
      .setFormula(formula)
      .setBackground('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(10)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0');
    newSummary.setRowHeight(R, 20); R++;
  });

  // D. OVERALL STATUS
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, LW);
  h_(newSummary.getRange(R, L), '#4A148C').setValue('D.  OVERALL STATUS');
  newSummary.setRowHeight(R, 20); R++;

  lbl(R, L, 'Test Closure Status', '#F3E5F5');
  m_(R, L + 1, 1, LW - 1);
  bd(newSummary.getRange(R, L + 1))
    .setValue('In Progress')
    .setBackground('#FFF9C4')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  newSummary.setRowHeight(R, 24); R++;

  // E. EXIT CRITERIA
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, LW);
  h_(newSummary.getRange(R, L), '#FF6F00').setValue('E.  EXIT CRITERIA');
  newSummary.setRowHeight(R, 20); R++;

  const exitFields = [
    ['All test cases executed', '☐ No  ☐ Yes'],
    ['Pass rate >= target', '☐ No  ☐ Yes'],
    ['No open blocker bugs', '☐ No  ☐ Yes'],
    ['No open VAPT blockers (Critical/High/Medium)', '☐ No  ☐ Yes'],
    ['All documentation complete', '☐ No  ☐ Yes'],
  ];

  exitFields.forEach(([label, value]) => {
    lbl(R, L, label, '#FFF3E0');
    m_(R, L + 1, 1, LW - 1);
    bd(newSummary.getRange(R, L + 1))
      .setValue(value)
      .setBackground('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(9)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    newSummary.setRowHeight(R, 18); R++;
  });

  // F. VAPT FINDINGS SUMMARY
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, LW);
  h_(newSummary.getRange(R, L), '#BF360C').setValue('F.  VAPT FINDINGS SUMMARY');
  newSummary.setRowHeight(R, 20); R++;

  // Total VAPT Findings
  lbl(R, L, 'Total VAPT Findings', '#FFEBEE');
  m_(R, L + 1, 1, LW - 1);
  bd(newSummary.getRange(R, L + 1))
    .setFormula('=IFERROR(COUNTA(FILTER(\'VAPT - Detail Finding\'!A:A,\'VAPT - Detail Finding\'!A:A<>\"\",\'VAPT - Detail Finding\'!A:A<>"Finding ID")),0)')
    .setBackground('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('0');
  newSummary.setRowHeight(R, 24); R++;

  // By Risk Level
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, 5);
  bd(newSummary.getRange(R, L))
    .setValue('By Risk Level (Adjusted Risk)')
    .setBackground('#FFCCBC')
    .setFontColor('#BF360C')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(9)
    .setHorizontalAlignment('left');
  newSummary.setRowHeight(R, 18); R++;

  // Risk level headers
  ['Risk Level', 'Count', 'Risk Level', 'Count'].forEach((h, i) => {
    const col = i < 2 ? L + i : L + i + 1;
    h_(newSummary.getRange(R, col), '#D84315').setValue(h).setFontSize(8);
  });
  newSummary.setRowHeight(R, 18); R++;

  // Risk level data
  const riskLevels = [
    ['Critical', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"Critical"),0)', 'Low', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"Low"),0)'],
    ['High', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"High"),0)', 'Informational', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"Informational"),0)'],
    ['Medium', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"Medium"),0)', '', ''],
  ];

  riskLevels.forEach((row, i) => {
    const rr = R + i;
    const bg = i % 2 === 0 ? '#FFF3E0' : '#FFFFFF';
    bd(newSummary.getRange(rr, L)).setValue(row[0]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
    bd(newSummary.getRange(rr, L + 1)).setFormula(row[1]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold').setNumberFormat('0');
    bd(newSummary.getRange(rr, L + 3)).setValue(row[2]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
    if (row[3]) {
      bd(newSummary.getRange(rr, L + 4)).setFormula(row[3]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold').setNumberFormat('0');
    }
    newSummary.setRowHeight(rr, 16);
  });
  R += 3;

  // By Status Fix
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, 5);
  bd(newSummary.getRange(R, L))
    .setValue('By Status Fix (Dev Team)')
    .setBackground('#E3F2FD')
    .setFontColor('#0D47A1')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(9)
    .setHorizontalAlignment('left');
  newSummary.setRowHeight(R, 18); R++;

  // Status Fix headers
  ['Status', 'Count', 'Status', 'Count'].forEach((h, i) => {
    const col = i < 2 ? L + i : L + i + 1;
    h_(newSummary.getRange(R, col), '#1565C0').setValue(h).setFontSize(8);
  });
  newSummary.setRowHeight(R, 18); R++;

  // Status Fix data
  const statusFix = [
    ['Todo', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"Todo"),0)', 'Done', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"Done"),0)'],
    ['On Progress Remediation', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"On Progress Remediation"),0)', 'Accepted', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"Accepted"),0)'],
    ['Ready to Retest', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"Ready to Retest"),0)', 'False Positive', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"False Positive"),0)'],
  ];

  statusFix.forEach((row, i) => {
    const rr = R + i;
    const bg = i % 2 === 0 ? '#F8F9FA' : '#FFFFFF';
    bd(newSummary.getRange(rr, L)).setValue(row[0]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
    bd(newSummary.getRange(rr, L + 1)).setFormula(row[1]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold').setNumberFormat('0');
    bd(newSummary.getRange(rr, L + 3)).setValue(row[2]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
    bd(newSummary.getRange(rr, L + 4)).setFormula(row[3]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold').setNumberFormat('0');
    newSummary.setRowHeight(rr, 16);
  });
  R += 3;

  // By Status Re-VAPT
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, 5);
  bd(newSummary.getRange(R, L))
    .setValue('By Status Re-VAPT (Pentester)')
    .setBackground('#F3E5F5')
    .setFontColor('#6A1B9A')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(9)
    .setHorizontalAlignment('left');
  newSummary.setRowHeight(R, 18); R++;

  const reVaptRow = R;
  bd(newSummary.getRange(reVaptRow, L)).setValue('Open').setBackground('#F8F9FA').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
  bd(newSummary.getRange(reVaptRow, L + 1)).setFormula('=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!I:I,"Open"),0)').setBackground('#F8F9FA').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold').setNumberFormat('0');
  bd(newSummary.getRange(reVaptRow, L + 3)).setValue('Closed').setBackground('#F8F9FA').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
  bd(newSummary.getRange(reVaptRow, L + 4)).setFormula('=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!I:I,"Closed"),0)').setBackground('#F8F9FA').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold').setNumberFormat('0');
  newSummary.setRowHeight(reVaptRow, 16); R++;

  // VAPT Blockers Breakdown
  newSummary.setRowHeight(R, 6); R++;
  m_(R, L, 1, 5);
  bd(newSummary.getRange(R, L))
    .setValue('VAPT Blockers (Status: Todo / On Progress / Ready to Retest / Accepted)')
    .setBackground('#FFCDD2')
    .setFontColor('#C62828')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(9)
    .setHorizontalAlignment('left');
  newSummary.setRowHeight(R, 18); R++;

  lbl(R, L, 'VAPT Blocker Count', '#FFEBEE');
  m_(R, L + 1, 1, LW - 1);
  bd(newSummary.getRange(R, L + 1))
    .setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*((\'VAPT - Detail Finding\'!E:E="Critical")+(\'VAPT - Detail Finding\'!E:E="High")+(\'VAPT - Detail Finding\'!E:E="Medium"))),0)')
    .setBackground('#FFCDD2')
    .setFontFamily('Arial')
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#B71C1C')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('0');
  newSummary.setRowHeight(R, 28); R++;

  lbl(R, L, 'VAPT Blocker Critical', '#FFF3E0');
  m_(R, L + 1, 1, LW - 1);
  bd(newSummary.getRange(R, L + 1))
    .setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*(\'VAPT - Detail Finding\'!E:E="Critical")),0)')
    .setBackground('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('0');
  newSummary.setRowHeight(R, 22); R++;

  lbl(R, L, 'VAPT Blocker High', '#FFF3E0');
  m_(R, L + 1, 1, LW - 1);
  bd(newSummary.getRange(R, L + 1))
    .setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*(\'VAPT - Detail Finding\'!E:E="High")),0)')
    .setBackground('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('0');
  newSummary.setRowHeight(R, 22); R++;

  lbl(R, L, 'VAPT Blocker Medium', '#FFF3E0');
  m_(R, L + 1, 1, LW - 1);
  bd(newSummary.getRange(R, L + 1))
    .setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*(\'VAPT - Detail Finding\'!E:E="Medium")),0)')
    .setBackground('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('0');
  newSummary.setRowHeight(R, 22); R++;

  lbl(R, L, 'VAPT Non Blocker Count', '#E8F5E9');
  m_(R, L + 1, 1, LW - 1);
  bd(newSummary.getRange(R, L + 1))
    .setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*((\'VAPT - Detail Finding\'!E:E="Low")+(\'VAPT - Detail Finding\'!E:E="Informational"))),0)')
    .setBackground('#FFFFFF')
    .setFontFamily('Arial')
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('0');
  newSummary.setRowHeight(R, 22); R++;

  // Note
  m_(R, L, 1, LW);
  newSummary.getRange(R, L)
    .setValue('Target: VAPT Blocker Count = 0 sebelum closure sign-off. Blocker = severity Critical/High/Medium dengan status selain Done/False Positive.')
    .setBackground('#FFF8E1')
    .setFontColor('#E65100')
    .setFontStyle('italic')
    .setFontSize(7)
    .setFontFamily('Arial')
    .setHorizontalAlignment('left');
  newSummary.setRowHeight(R, 14); R++;

  SpreadsheetApp.flush();
}
