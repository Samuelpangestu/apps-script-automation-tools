/**
 * DebugNotes.js - Debug why notes not applying
 */

/**
 * Force update ONE specific note for testing
 */
function forceUpdateOneNote() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Overview');

  if (!ws) {
    SpreadsheetApp.getUi().alert('❌ Overview tab not found!');
    return;
  }

  try {
    // Force update note at specific cell
    // Assuming row 4 is header row, let's update a few cells

    // Clear ALL notes in header row first
    Logger.log('Clearing all notes in row 4...');
    ws.getRange(4, 1, 1, 30).clearNote();

    // Add test note to column 1 (should be "Project" or similar)
    Logger.log('Adding test note to D4...');
    const cell = ws.getRange(4, 4); // Row 4, Col 4 (D4)
    const currentValue = cell.getValue();

    const testNote = '🔥 UPDATED NOTE - COMPREHENSIVE VERSION\n\n' +
      'Column: ' + currentValue + '\n' +
      'Time: ' + new Date().toLocaleString() + '\n\n' +
      'If you see this, notes UPDATE is working!\n\n' +
      'This is the NEW comprehensive version with:\n' +
      '✅ Detailed explanations\n' +
      '✅ Formulas\n' +
      '✅ Thresholds\n' +
      '✅ Action items';

    cell.setNote(testNote);

    Logger.log('Test note added successfully');

    SpreadsheetApp.getUi().alert(
      '✅ Test Complete',
      'Test note added to Overview row 4, column 4 (' + currentValue + ')\n\n' +
      'Hover mouse di cell tersebut untuk lihat note.\n\n' +
      'Jika note muncul dengan timestamp baru, berarti update working!',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * Show current Dashboard structure for debugging
 */
function showDashboardStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Overview');

  if (!ws) {
    SpreadsheetApp.getUi().alert('❌ Overview not found');
    return;
  }

  // Get headers from row 4
  const headers = ws.getRange(4, 1, 1, 20).getValues()[0];

  let msg = 'OVERVIEW STRUCTURE (Row 4 headers):\n\n';
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]) {
      msg += 'Col ' + (i + 1) + ' (' + String.fromCharCode(65 + i) + '): ' + headers[i] + '\n';
    }
  }

  // Check which cells have notes
  msg += '\n\nCELLS WITH EXISTING NOTES:\n';
  for (let i = 1; i <= 20; i++) {
    const note = ws.getRange(4, i).getNote();
    if (note) {
      const col = String.fromCharCode(64 + i);
      const preview = note.substring(0, 50) + (note.length > 50 ? '...' : '');
      msg += col + '4: ' + preview + '\n';
    }
  }

  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Dashboard Structure', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Clear ALL notes from Overview tab
 */
function clearAllOverviewNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Overview');

  if (!ws) {
    SpreadsheetApp.getUi().alert('❌ Overview not found');
    return;
  }

  const response = SpreadsheetApp.getUi().alert(
    '⚠️ Clear All Notes?',
    'Akan menghapus SEMUA notes di Overview tab.\n\n' +
    'Lanjutkan?',
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response !== SpreadsheetApp.getUi().Button.YES) {
    return;
  }

  try {
    // Clear all notes in the sheet
    ws.getDataRange().clearNote();

    SpreadsheetApp.getUi().alert(
      '✅ Cleared',
      'Semua notes di Overview telah dihapus.\n\n' +
      'Sekarang run applyComprehensiveNotes() untuk add notes baru.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    Logger.log('All Overview notes cleared');

  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}
