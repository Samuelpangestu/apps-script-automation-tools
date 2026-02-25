/**
 * AppendixRevert.js
 * ─────────────────────────────────────────────────────────────────────────
 * REVERT SCRIPT - Remove section "0. HIERARKI QA" dari semua modul
 *
 * Use this untuk mengembalikan Appendix ke kondisi semula
 * (sebelum ada section "0. HIERARKI QA")
 *
 * FUNCTION:
 *   removeAppendixSection() -- Hapus section "0. HIERARKI QA" dari semua modul
 * ─────────────────────────────────────────────────────────────────────────
 */

// ── Helper: Safe UI alert ─────────────────────────────────────────────────
function safeAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log('='.repeat(60));
    Logger.log('INFO: ' + message);
    Logger.log('='.repeat(60));
  }
}

// ── Auto-detect Config column layout ─────────────────────────────────────
function getModulesFromConfig_(cfg) {
  var data = cfg.getDataRange().getValues();
  if (data.length < 4) return [];

  var headerRow = null;
  var headerIdx = -1;
  for (var h = 0; h <= 3; h++) {
    var row = data[h].map(function(c) { return String(c).trim().toUpperCase(); });
    if (row.indexOf('SPREADSHEET ID') !== -1 || row.indexOf('SPREADSHEET_ID') !== -1) {
      headerRow = row;
      headerIdx = h;
      break;
    }
  }

  var COL_ACTIVE = 0;
  var COL_ID     = -1;
  var COL_NAME   = -1;

  if (headerRow) {
    headerRow.forEach(function(h, i) {
      if (h === 'SPREADSHEET ID' || h === 'SPREADSHEET_ID') COL_ID = i;
      if (h === 'SUBMODULE' || h === 'SUBMODUL' || h === 'MODUL NAME' || h === 'MODULE NAME') COL_NAME = i;
      if (h === 'PROJECT' && COL_NAME === -1) COL_NAME = i;
    });
  }

  var SHEETS_ID_RE = /^[A-Za-z0-9_\-]{20,}$/;
  if (COL_ID === -1) {
    for (var col = 0; col < (data[0] || []).length; col++) {
      var matches = 0;
      for (var r = headerIdx + 1; r < Math.min(data.length, headerIdx + 10); r++) {
        if (SHEETS_ID_RE.test(String(data[r][col]).trim())) matches++;
      }
      if (matches >= 3) { COL_ID = col; break; }
    }
  }

  if (COL_ID === -1) {
    Logger.log('ERROR: Cannot detect Spreadsheet ID column in Config');
    return [];
  }

  var modules = [];
  for (var i = headerIdx + 1; i < data.length; i++) {
    var active = String(data[i][COL_ACTIVE]).trim().toUpperCase();
    if (active !== 'Y' && active !== 'YES') continue;

    var id   = String(data[i][COL_ID] || '').trim();
    var name = COL_NAME !== -1 ? String(data[i][COL_NAME] || '').trim() : 'Module_' + (i - headerIdx);
    if (id && SHEETS_ID_RE.test(id)) {
      modules.push({ name: name, id: id });
    }
  }

  Logger.log('Config: found ' + modules.length + ' active modules (ID col=' + (COL_ID + 1) + ')');
  return modules;
}

// ── REVERT FUNCTION: Remove Appendix Section ──────────────────────────────
/**
 * Remove section "0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE"
 * dari semua modul aktif.
 *
 * WARNING: Ini akan menghapus section tersebut PERMANENTLY!
 * Gunakan ini hanya jika ingin revert ke kondisi Appendix semula.
 */
function removeAppendixSection() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');

  if (!cfg) {
    safeAlert_('Config tab tidak ditemukan.\nJalankan dari QA Dashboard.');
    return;
  }

  // Confirmation prompt
  var ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log('WARNING: Running without UI confirmation (triggered from script/trigger)');
  }

  if (ui) {
    var response = ui.alert(
      'REVERT APPENDIX SECTION',
      'Ini akan MENGHAPUS section "0. HIERARKI QA" dari semua modul aktif.\n\n' +
      'Appendix akan kembali ke kondisi SEMULA (sebelum ada section ini).\n\n' +
      'Lanjutkan?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      Logger.log('CANCELLED: User cancelled revert operation');
      return;
    }
  }

  var modules = getModulesFromConfig_(cfg);
  if (modules.length === 0) {
    safeAlert_('Tidak ada modul aktif ditemukan di Config.');
    return;
  }

  var SECTION_TITLE = '0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE';
  var done = 0, skipped = 0, failed = 0;

  modules.forEach(function(mod) {
    var name    = mod.name;
    var cleanId = mod.id.replace(/[^A-Za-z0-9_\-]/g, '');

    if (cleanId.length < 20) {
      Logger.log('SKIP ' + name + ': ID tidak valid');
      skipped++;
      return;
    }

    try {
      var modSS = SpreadsheetApp.openById(cleanId);
      var ws    = modSS.getSheetByName('Appendix');

      if (!ws) {
        Logger.log('SKIP ' + name + ': tab Appendix tidak ditemukan');
        skipped++;
        return;
      }

      // Find section
      var lastRow = ws.getLastRow();
      var data    = lastRow > 0 ? ws.getRange(1, 1, lastRow, 1).getValues() : [];
      var sectionRow = -1;

      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === SECTION_TITLE) {
          sectionRow = i + 1; // 1-indexed
          break;
        }
      }

      if (sectionRow === -1) {
        Logger.log('SKIP ' + name + ': section tidak ditemukan (already removed)');
        skipped++;
        return;
      }

      // Delete section (header + content rows + spacer)
      // Delete up to 10 rows to accommodate any format
      var deleteCount = Math.min(10, lastRow - sectionRow + 1);
      ws.deleteRows(sectionRow, deleteCount);
      SpreadsheetApp.flush();

      Logger.log('REMOVED ' + name + ': deleted ' + deleteCount + ' rows starting at row ' + sectionRow);
      done++;

    } catch (e) {
      Logger.log('ERROR ' + name + ': ' + e.message);
      failed++;
    }
  });

  safeAlert_(
    'REVERT COMPLETE\n\n' +
    'Section "0. HIERARKI QA" dihapus dari:\n\n' +
    'Berhasil : ' + done    + ' modul\n' +
    'Dilewati : ' + skipped + ' modul\n' +
    'Gagal    : ' + failed  + ' modul\n\n' +
    'Appendix sekarang kembali ke kondisi SEMULA.\n' +
    'Section "0. HIERARKI QA" sudah tidak ada.\n\n' +
    'Lihat Apps Script Logs untuk detail.'
  );
}
