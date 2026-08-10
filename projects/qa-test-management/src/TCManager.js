/**
 * QA Test Management - TC Manager v6
 *
 * Sync TC_Master dan TC_Execution untuk insert/delete operations
 * Prevents screenshot/result misalignment
 *
 * Features:
 * - Insert TC (single & bulk) - leaves rows empty for ARRAYFORMULA
 * - Delete TC (single & multiple rows)
 * - Mark as Deprecated (single & multiple rows)
 * - Session reminder to use TC Manager
 */

// ═══════════════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🔧 TC Manager')
    .addItem('📋 Insert TC Here', 'insertTCAtPosition')
    .addItem('📋 Bulk Insert', 'bulkInsertTC')
    .addSeparator()
    .addItem('🗑️ Delete TC', 'deleteTC')
    .addItem('⚠️ Mark as Deprecated', 'markAsDeprecated')
    .addSeparator()
    .addItem('ℹ️ Help', 'showHelp')
    .addToUi();

  // Show reminder once per session
  showSessionReminder();
}

// ═══════════════════════════════════════════════════════════════════════
// SESSION REMINDER (Show once per 6 hours)
// ═══════════════════════════════════════════════════════════════════════

function showSessionReminder() {
  const cache = CacheService.getScriptCache();
  const shown = cache.get('tc_manager_reminder_shown');

  if (!shown) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '⚠️ Important Warning',
      '🚫 DO NOT insert/delete rows manually!\n\n' +
      'Always use 🔧 TC Manager menu:\n' +
      '• Insert TC Here\n' +
      '• Bulk Insert\n' +
      '• Delete TC\n\n' +
      'Manual insert/delete breaks TC alignment and test results.',
      ui.ButtonSet.OK
    );
    cache.put('tc_manager_reminder_shown', 'true', 21600); // 6 hours
  }
}

function getActiveSheetPair(ss) {
  const activeSheet = ss.getActiveSheet();
  const activeName = activeSheet && activeSheet.getName();
  const pairs = {
    TC_Master: { executionName: 'TC_Execution', scenarioColumn: 11 },
    API_Master: { executionName: 'API_Execution', scenarioColumn: 13 },
  };
  const pair = pairs[activeName];

  if (!pair) return null;

  const execution = ss.getSheetByName(pair.executionName);
  if (!activeSheet || !execution) return null;

  return {
    master: activeSheet,
    execution: execution,
    masterName: activeName,
    executionName: pair.executionName,
    scenarioColumn: pair.scenarioColumn,
  };
}

// ═══════════════════════════════════════════════════════════════════════
// INSERT TC AT POSITION (SINGLE)
// ═══════════════════════════════════════════════════════════════════════

function insertTCAtPosition() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPair = getActiveSheetPair(ss);

  if (!sheetPair) {
    ui.alert('❌ Error', 'Select TC_Master or API_Master first.', ui.ButtonSet.OK);
    return;
  }

  const tcMaster = sheetPair.master;
  const tcExecution = sheetPair.execution;
  const activeRow = tcMaster.getActiveRange().getRow();

  if (activeRow < 3) {
    ui.alert('⚠️ Invalid', 'Please select row 3 or below in ' + sheetPair.masterName, ui.ButtonSet.OK);
    return;
  }

  const tcIdAtPosition = tcMaster.getRange(activeRow, 3).getValue();
  const execStartRow = 9;
  const execRow = execStartRow + (activeRow - 3);

  const response = ui.alert(
    '📋 Insert TC',
    '➕ INSERT 1 ROW BEFORE:\n\n' +
    '📍 ' + sheetPair.masterName + ':\n' +
    '   Row: ' + activeRow + '\n' +
    '   TC_ID: ' + tcIdAtPosition + '\n\n' +
    '📍 ' + sheetPair.executionName + ':\n' +
    '   Row: ' + execRow + '\n\n' +
    '⚠️ Rows below will shift down in BOTH sheets!\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Insert rows (leave empty - ARRAYFORMULA will auto-populate)
  tcMaster.insertRowBefore(activeRow);
  tcExecution.insertRowBefore(execRow);

  ui.alert(
    '✅ Inserted!',
    '📍 ' + sheetPair.masterName + ' row: ' + activeRow + '\n' +
    '📍 ' + sheetPair.executionName + ' row: ' + execRow + '\n\n' +
    'Fill in TC_ID and details now.',
    ui.ButtonSet.OK
  );

  tcMaster.setActiveRange(tcMaster.getRange(activeRow, 3)); // Focus on TC_ID column
}

// ═══════════════════════════════════════════════════════════════════════
// BULK INSERT (MULTIPLE ROWS)
// ═══════════════════════════════════════════════════════════════════════

function bulkInsertTC() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPair = getActiveSheetPair(ss);

  if (!sheetPair) {
    ui.alert('❌ Error', 'Select TC_Master or API_Master first.', ui.ButtonSet.OK);
    return;
  }

  const tcMaster = sheetPair.master;
  const tcExecution = sheetPair.execution;
  const activeRow = tcMaster.getActiveRange().getRow();

  if (activeRow < 3) {
    ui.alert('⚠️ Invalid', 'Please select row 3 or below in ' + sheetPair.masterName, ui.ButtonSet.OK);
    return;
  }

  // Ask how many rows to insert
  const numRowsResponse = ui.prompt(
    '📋 Bulk Insert',
    'How many rows to insert?\n\n' +
    'Will insert BEFORE row ' + activeRow,
    ui.ButtonSet.OK_CANCEL
  );

  if (numRowsResponse.getSelectedButton() !== ui.Button.OK) return;

  const numRows = parseInt(numRowsResponse.getResponseText());

  if (isNaN(numRows) || numRows < 1 || numRows > 100) {
    ui.alert('❌ Invalid', 'Please enter number between 1-100', ui.ButtonSet.OK);
    return;
  }

  const tcIdAtPosition = tcMaster.getRange(activeRow, 3).getValue();
  const execStartRow = 9;
  const execRow = execStartRow + (activeRow - 3);

  const response = ui.alert(
    '📋 Bulk Insert Confirmation',
    '➕ INSERT ' + numRows + ' ROWS BEFORE:\n\n' +
    '📍 ' + sheetPair.masterName + ':\n' +
    '   Starting row: ' + activeRow + '\n' +
    '   Current TC_ID: ' + tcIdAtPosition + '\n' +
    '   New rows: ' + activeRow + ' to ' + (activeRow + numRows - 1) + '\n\n' +
    '📍 ' + sheetPair.executionName + ':\n' +
    '   Starting row: ' + execRow + '\n' +
    '   New rows: ' + execRow + ' to ' + (execRow + numRows - 1) + '\n\n' +
    '⚠️ Rows below will shift down!\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Insert rows one by one (leave empty - ARRAYFORMULA will auto-populate)
  for (let i = 0; i < numRows; i++) {
    tcMaster.insertRowBefore(activeRow);
    tcExecution.insertRowBefore(execRow);
  }

  ui.alert(
    '✅ Bulk Insert Complete!',
    'Inserted ' + numRows + ' rows:\n\n' +
    '📍 ' + sheetPair.masterName + ': rows ' + activeRow + '-' + (activeRow + numRows - 1) + '\n' +
    '📍 ' + sheetPair.executionName + ': rows ' + execRow + '-' + (execRow + numRows - 1) + '\n\n' +
    'Fill in TC details now.',
    ui.ButtonSet.OK
  );

  tcMaster.setActiveRange(tcMaster.getRange(activeRow, 3)); // Focus on first new row
}

// ═══════════════════════════════════════════════════════════════════════
// DELETE TC
// ═══════════════════════════════════════════════════════════════════════

function deleteTC() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPair = getActiveSheetPair(ss);

  if (!sheetPair) {
    ui.alert('❌ Error', 'Select TC_Master or API_Master first.', ui.ButtonSet.OK);
    return;
  }

  const tcMaster = sheetPair.master;
  const tcExecution = sheetPair.execution;
  const activeRange = tcMaster.getActiveRange();
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();

  if (startRow < 3) {
    ui.alert('⚠️ Invalid', 'Please select row 3 or below', ui.ButtonSet.OK);
    return;
  }

  const execStartRow = 9;
  const execRow = execStartRow + (startRow - 3);

  // Collect TC_IDs for confirmation
  const tcIds = [];
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const tcId = tcMaster.getRange(row, 3).getValue();
    if (tcId) tcIds.push(tcId);
  }

  if (tcIds.length === 0) {
    ui.alert('⚠️ No TCs', 'No TC_IDs found in selected rows', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '⚠️ DELETE TC',
    '🗑️ DELETE:\n\n' +
    '📍 ' + sheetPair.masterName + ' rows: ' + startRow + ' to ' + (startRow + numRows - 1) + '\n' +
    '📍 ' + sheetPair.executionName + ' rows: ' + execRow + ' to ' + (execRow + numRows - 1) + '\n' +
    '📋 Total TCs: ' + tcIds.length + '\n' +
    '🔖 TC_IDs: ' + tcIds.join(', ') + '\n\n' +
    '⚠️ WARNING:\n' +
    '- Test results will be LOST!\n' +
    '- Rows below will shift up!\n\n' +
    '💡 TIP: Use "Mark as Deprecated" instead?\n\n' +
    'Continue DELETE?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Delete rows from bottom to top (to avoid row index shifting issues)
  for (let i = numRows - 1; i >= 0; i--) {
    const rowToDelete = startRow + i;
    const execRowToDelete = execRow + i;
    tcMaster.deleteRow(rowToDelete);
    tcExecution.deleteRow(execRowToDelete);
  }

  ui.alert(
    '✅ Deleted',
    'Deleted ' + tcIds.length + ' TC(s)\n\n' +
    'TC_IDs: ' + tcIds.join(', ') + '\n\n' +
    'Deleted from both ' + sheetPair.masterName + ' and ' + sheetPair.executionName + '.',
    ui.ButtonSet.OK
  );
}

// ═══════════════════════════════════════════════════════════════════════
// MARK AS DEPRECATED
// ═══════════════════════════════════════════════════════════════════════

function markAsDeprecated() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPair = getActiveSheetPair(ss);

  if (!sheetPair) {
    ui.alert('❌ Error', 'Select TC_Master or API_Master first.', ui.ButtonSet.OK);
    return;
  }

  const tcMaster = sheetPair.master;
  const activeRange = tcMaster.getActiveRange();
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();

  if (startRow < 3) {
    ui.alert('⚠️ Invalid', 'Select row 3 or below', ui.ButtonSet.OK);
    return;
  }

  // Collect TC_IDs for confirmation
  const tcIds = [];
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const tcId = tcMaster.getRange(row, 3).getValue();
    if (tcId) tcIds.push(tcId);
  }

  if (tcIds.length === 0) {
    ui.alert('⚠️ No TCs', 'No TC_IDs found in selected rows', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '⚠️ Mark as Deprecated',
    '📍 Selected rows: ' + startRow + ' to ' + (startRow + numRows - 1) + '\n' +
    '📋 Total TCs: ' + tcIds.length + '\n' +
    '🔖 TC_IDs: ' + tcIds.join(', ') + '\n\n' +
    'This will:\n' +
    '✓ Add [DEPRECATED] prefix to scenario\n' +
    '✓ Gray out the rows\n' +
    '✓ Keep test results intact\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  let markedCount = 0;

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const scenarioCell = tcMaster.getRange(row, sheetPair.scenarioColumn);
    const currentScenario = scenarioCell.getValue();

    if (currentScenario && !currentScenario.toString().includes('[DEPRECATED]')) {
      scenarioCell.setValue('[DEPRECATED] ' + currentScenario);
      tcMaster.getRange(row, 1, 1, tcMaster.getLastColumn())
        .setBackground('#F5F5F5')
        .setFontColor('#999999');
      markedCount++;
    }
  }

  ui.alert(
    '✅ Marked',
    'Deprecated ' + markedCount + ' TC(s)\n\n' +
    'Rows: ' + startRow + '-' + (startRow + numRows - 1),
    ui.ButtonSet.OK
  );
}

// ═══════════════════════════════════════════════════════════════════════
// HELP
// ═══════════════════════════════════════════════════════════════════════

function showHelp() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    'ℹ️ TC Manager Help',
    '📋 INSERT TC HERE\n' +
    '• Insert 1 row at cursor position\n' +
    '• Syncs TC_Master + TC_Execution\n' +
    '• Shows exact row numbers\n\n' +
    '📋 BULK INSERT\n' +
    '• Insert multiple rows at once\n' +
    '• Enter number of rows (1-100)\n' +
    '• Faster for multiple TCs\n\n' +
    '🗑️ DELETE TC\n' +
    '• Deletes from both sheets\n' +
    '• Supports single & multiple rows\n' +
    '• Test results will be lost\n' +
    '• Use with caution!\n\n' +
    '⚠️ MARK AS DEPRECATED\n' +
    '• Safer than delete\n' +
    '• Supports single & multiple rows\n' +
    '• Adds [DEPRECATED] prefix\n' +
    '• Keeps test results intact\n\n' +
    '💡 WHY?\n' +
    'Manual insert/delete breaks alignment.\n' +
    'TC Manager keeps screenshots aligned with correct TC_ID.',
    ui.ButtonSet.OK
  );
}
