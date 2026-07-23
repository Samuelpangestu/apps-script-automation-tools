/**
 * QA Test Management - TC Manager (Clean Version)
 *
 * Solves alignment issue antara TC_Master dan TC_Execution
 * User pakai menu ini untuk add/insert/delete TC agar sync tetap terjaga
 */

// ═══════════════════════════════════════════════════════════════════════
// MENU - CLEAN VERSION (Only TC Manager)
// ═══════════════════════════════════════════════════════════════════════

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // TC Manager Menu
  ui.createMenu('🔧 TC Manager')
    .addItem('➕ Add New TC (at bottom)', 'addNewTC')
    .addItem('📋 Insert TC Here (at cursor)', 'insertTCAtPosition')
    .addSeparator()
    .addItem('🗑️ Delete TC (current row)', 'deleteTC')
    .addItem('⚠️ Mark as Deprecated', 'markAsDeprecated')
    .addSeparator()
    .addSubMenu(ui.createMenu('ℹ️ Help')
      .addItem('📖 How to Use', 'showHelp')
      .addItem('⚡ Why Use This?', 'showWhy'))
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════
// ADD NEW TC (SAFEST - APPEND AT BOTTOM)
// ═══════════════════════════════════════════════════════════════════════

function addNewTC() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tcMaster = ss.getSheetByName('TC_Master');
  const tcExecution = ss.getSheetByName('TC_Execution');

  if (!tcMaster || !tcExecution) {
    ui.alert('❌ Error', 'TC_Master or TC_Execution sheet not found!', ui.ButtonSet.OK);
    return;
  }

  // Find last row with data in TC_Master (column C = TC_ID)
  const lastRow = tcMaster.getRange('C:C').getValues().filter(String).length + 2;
  const newRow = lastRow + 1;

  // Calculate corresponding row in TC_Execution
  const execStartRow = 9;
  const execNewRow = execStartRow + (newRow - 3);

  // Insert row in TC_Master
  tcMaster.insertRowAfter(lastRow);

  // Insert row in TC_Execution
  tcExecution.insertRowAfter(execNewRow - 1);

  // Copy formulas from previous row (columns B-G)
  if (execNewRow > execStartRow) {
    const sourceRange = tcExecution.getRange(execNewRow - 1, 2, 1, 6);
    const targetRange = tcExecution.getRange(execNewRow, 2, 1, 6);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  }

  ui.alert(
    '✅ New TC Added!',
    'TC_Master row: ' + newRow + '\n' +
    'TC_Execution row: ' + execNewRow + '\n\n' +
    'Next: Fill in TC_ID and details',
    ui.ButtonSet.OK
  );

  // Move cursor to TC_ID column
  tcMaster.setActiveRange(tcMaster.getRange(newRow, 3));
}

// ═══════════════════════════════════════════════════════════════════════
// INSERT TC AT POSITION
// ═══════════════════════════════════════════════════════════════════════

function insertTCAtPosition() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tcMaster = ss.getSheetByName('TC_Master');
  const tcExecution = ss.getSheetByName('TC_Execution');

  if (!tcMaster || !tcExecution) {
    ui.alert('❌ Error', 'Sheets not found!', ui.ButtonSet.OK);
    return;
  }

  const activeRow = tcMaster.getActiveRange().getRow();

  if (activeRow < 3) {
    ui.alert('⚠️ Invalid', 'Select row 3 or below', ui.ButtonSet.OK);
    return;
  }

  const tcIdAtPosition = tcMaster.getRange(activeRow, 3).getValue();

  const response = ui.alert(
    '📋 Insert TC',
    'Insert BEFORE row ' + activeRow + '\nTC_ID: ' + tcIdAtPosition + '\n\n' +
    '⚠️ Rows below will shift down!',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const execStartRow = 9;
  const execRow = execStartRow + (activeRow - 3);

  // Insert rows
  tcMaster.insertRowBefore(activeRow);
  tcExecution.insertRowBefore(execRow);

  // Copy formulas
  if (execRow > execStartRow) {
    const sourceRange = tcExecution.getRange(execRow - 1, 2, 1, 6);
    const targetRange = tcExecution.getRange(execRow, 2, 1, 6);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  }

  ui.alert('✅ Inserted!', 'Row: ' + activeRow, ui.ButtonSet.OK);
  tcMaster.setActiveRange(tcMaster.getRange(activeRow, 3));
}

// ═══════════════════════════════════════════════════════════════════════
// DELETE TC
// ═══════════════════════════════════════════════════════════════════════

function deleteTC() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tcMaster = ss.getSheetByName('TC_Master');
  const tcExecution = ss.getSheetByName('TC_Execution');

  if (!tcMaster || !tcExecution) {
    ui.alert('❌ Error', 'Sheets not found!', ui.ButtonSet.OK);
    return;
  }

  const activeRow = tcMaster.getActiveRange().getRow();

  if (activeRow < 3) {
    ui.alert('⚠️ Invalid', 'Select row 3 or below', ui.ButtonSet.OK);
    return;
  }

  const tcId = tcMaster.getRange(activeRow, 3).getValue();
  const scenario = tcMaster.getRange(activeRow, 11).getValue();

  const response = ui.alert(
    '⚠️ DELETE TC',
    'Delete:\n' +
    'TC_ID: ' + tcId + '\n' +
    'Scenario: ' + scenario + '\n\n' +
    '⚠️ WARNING:\n' +
    '- Deletes from both TC_Master & TC_Execution\n' +
    '- Test results will be LOST!\n\n' +
    '💡 Use "Mark as Deprecated" instead?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const execStartRow = 9;
  const execRow = execStartRow + (activeRow - 3);

  // Delete rows
  tcMaster.deleteRow(activeRow);
  tcExecution.deleteRow(execRow);

  ui.alert('✅ Deleted', 'TC_ID: ' + tcId, ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════
// MARK AS DEPRECATED (SAFER ALTERNATIVE)
// ═══════════════════════════════════════════════════════════════════════

function markAsDeprecated() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tcMaster = ss.getSheetByName('TC_Master');

  if (!tcMaster) {
    ui.alert('❌ Error', 'TC_Master not found!', ui.ButtonSet.OK);
    return;
  }

  const activeRow = tcMaster.getActiveRange().getRow();

  if (activeRow < 3) {
    ui.alert('⚠️ Invalid', 'Select row 3 or below', ui.ButtonSet.OK);
    return;
  }

  const tcId = tcMaster.getRange(activeRow, 3).getValue();
  const scenarioCell = tcMaster.getRange(activeRow, 11);
  const currentScenario = scenarioCell.getValue();

  const response = ui.alert(
    'Mark as Deprecated',
    'TC_ID: ' + tcId + '\n\n' +
    'Adds [DEPRECATED] prefix to scenario',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Add prefix if not already there
  if (!currentScenario.toString().includes('[DEPRECATED]')) {
    scenarioCell.setValue('[DEPRECATED] ' + currentScenario);

    // Gray out row
    tcMaster.getRange(activeRow, 1, 1, tcMaster.getLastColumn())
      .setBackground('#F5F5F5')
      .setFontColor('#999999');
  }

  ui.alert('✅ Marked', 'TC_ID: ' + tcId + ' deprecated', ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════
// HELP FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

function showHelp() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    '📖 How to Use TC Manager',
    '1️⃣ Add New TC (at bottom)\n' +
    '   - Safest option\n' +
    '   - No impact on existing results\n\n' +
    '2️⃣ Insert TC Here\n' +
    '   - Insert at cursor position\n' +
    '   - Rows below shift down\n\n' +
    '3️⃣ Delete TC\n' +
    '   - Removes TC permanently\n' +
    '   - Use with caution!\n\n' +
    '4️⃣ Mark as Deprecated\n' +
    '   - Safer than delete\n' +
    '   - Keeps TC but marks obsolete\n\n' +
    '💡 Best Practice:\n' +
    '- Use "Add New TC" for new cases\n' +
    '- Use "Mark as Deprecated" vs delete',
    ui.ButtonSet.OK
  );
}

function showWhy() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    '⚡ Why TC Manager?',
    '❌ PROBLEM:\n' +
    'Manual insert/delete breaks alignment:\n' +
    '- TC_ID shifts position\n' +
    '- Screenshots map to wrong TC\n' +
    '- Results misaligned\n\n' +
    '✅ SOLUTION:\n' +
    'TC Manager syncs both sheets:\n' +
    '- Insert in TC_Master + TC_Execution\n' +
    '- Delete in both sheets\n' +
    '- Perfect alignment maintained\n\n' +
    '📊 Example:\n' +
    'TC-001 screenshot at row 10\n' +
    'Manual insert → TC-001 shifts to row 11\n' +
    'Screenshot still at row 10 = WRONG!\n\n' +
    'With TC Manager:\n' +
    'Both sheets sync → Screenshot stays correct ✓',
    ui.ButtonSet.OK
  );
}
