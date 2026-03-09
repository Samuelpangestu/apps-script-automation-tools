/**
 * fixConfig.js - Fix redundant columns in Config
 *
 * Jalankan fixConfigStructure() untuk cleanup kolom yang redundan
 */

function fixConfigStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');

  if (!cfg) {
    SpreadsheetApp.getUi().alert('Config tab tidak ditemukan!');
    return;
  }

  // Get current headers (row 3)
  const headers = cfg.getRange(3, 1, 1, cfg.getLastColumn()).getValues()[0];

  Logger.log('Current headers: ' + headers.join(' | '));

  // Expected structure
  const expected = [
    'Active', 'Jira Sync', 'Project', 'Modul', 'Submodul', 'PIC QA',
    'Spreadsheet ID', 'Link', 'Jira Instance', 'Jira Project Key'
  ];

  // Check if structure is already correct
  const current = headers.slice(0, 10).map(h => String(h).trim());
  const isCorrect = JSON.stringify(current) === JSON.stringify(expected);

  if (isCorrect) {
    SpreadsheetApp.getUi().alert('✅ Config structure sudah benar!\n\nTidak perlu cleanup.');
    Logger.log('✅ Config structure is correct');
    return;
  }

  Logger.log('❌ Config structure needs fixing');
  Logger.log('Current (first 13): ' + current.concat(headers.slice(10, 13)).join(' | '));
  Logger.log('Expected: ' + expected.join(' | '));

  // Show current structure
  const msg =
    '⚠️ Struktur Config perlu diperbaiki!\n\n' +
    'Current headers (A-M):\n' +
    headers.slice(0, 13).map((h, i) => String.fromCharCode(65+i) + ': ' + h).join('\n') +
    '\n\nExpected headers (A-J):\n' +
    expected.map((h, i) => String.fromCharCode(65+i) + ': ' + h).join('\n') +
    '\n\n────────────────────────────\n' +
    'Fix yang akan dilakukan:\n' +
    '1. Hapus kolom H (Jira Sync duplicate)\n' +
    '2. Hapus kolom L-M (Jira Instance & Project duplicate)\n' +
    '3. Pindahkan kolom I, J, K ke posisi yang benar\n\n' +
    'Lanjutkan?';

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Fix Config Structure', msg, ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    Logger.log('User cancelled fix');
    return;
  }

  // Delete redundant columns (work backwards to maintain indices)
  // Delete col M (Jira Project duplicate)
  if (headers.length >= 13 && String(headers[12]).includes('Jira Project')) {
    cfg.deleteColumn(13);
    Logger.log('Deleted col M (Jira Project duplicate)');
  }

  // Delete col L (Jira Instance duplicate)
  if (headers.length >= 12 && String(headers[11]).includes('Jira Instance')) {
    cfg.deleteColumn(12);
    Logger.log('Deleted col L (Jira Instance duplicate)');
  }

  // Delete col H (Jira Sync duplicate)
  const colH = String(cfg.getRange(3, 8).getValue()).trim();
  if (colH.includes('Jira Sync')) {
    cfg.deleteColumn(8);
    Logger.log('Deleted col H (Jira Sync duplicate)');
  }

  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  // Verify new structure
  const newHeaders = cfg.getRange(3, 1, 1, 10).getValues()[0].map(h => String(h).trim());

  Logger.log('New headers: ' + newHeaders.join(' | '));

  // Check if matches expected
  const fixed = JSON.stringify(newHeaders) === JSON.stringify(expected);

  if (fixed) {
    ui.alert(
      '✅ Config structure berhasil diperbaiki!\n\n' +
      'New structure (A-J):\n' +
      newHeaders.map((h, i) => String.fromCharCode(65+i) + ': ' + h).join('\n') +
      '\n\nSekarang bisa jalankan:\n' +
      '• jiraSetup() untuk add Google Chat section\n' +
      '• setupDailyBlockerNotification() untuk setup trigger'
    );
    Logger.log('✅ Config structure fixed successfully');
  } else {
    ui.alert(
      '⚠️ Manual adjustment diperlukan\n\n' +
      'Current structure:\n' +
      newHeaders.map((h, i) => String.fromCharCode(65+i) + ': ' + h).join('\n') +
      '\n\nExpected:\n' +
      expected.map((h, i) => String.fromCharCode(65+i) + ': ' + h).join('\n') +
      '\n\nSilakan adjust manual atau hubungi support.'
    );
    Logger.log('⚠️ Manual adjustment needed');
  }
}
