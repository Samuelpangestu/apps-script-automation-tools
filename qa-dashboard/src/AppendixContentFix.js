/**
 * AppendixContentFix.js
 * ─────────────────────────────────────────────────────────────────────────
 * Fix khusus untuk update content Appendix section "0. HIERARKI QA"
 * dengan TC_ID reference yang lengkap dan jelas.
 *
 * FUNCTION:
 *   fixAppendixContent() -- Update content di section "0. HIERARKI QA" semua modul
 * ─────────────────────────────────────────────────────────────────────────
 */

// ── Helper: Safe UI alert (fallback to Logger if no UI context) ───────────
function safeAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log('='.repeat(60));
    Logger.log('INFO: ' + message);
    Logger.log('='.repeat(60));
  }
}

// ── Content untuk Appendix ────────────────────────────────────────────────
var SECTION_TITLE = '0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE';

var CONTENT_ROWS = [
  {
    label: 'Definisi',
    content:
      'Project   = Inisiatif / client / program kerja. Contoh: SIPGN, INAGOV\n' +
      'Module    = Pengelompokan domain fungsional dalam project.\n' +
      '            Kosongkan ("-") jika project flat (tidak punya layer domain).\n' +
      'SubModule = Unit terkecil yang berdiri sendiri -- 1 aplikasi atau 1 domain.\n' +
      '            ANCHOR utama untuk TC_ID, Coverage, dan Dashboard.\n' +
      'Feature   = Fitur besar dalam SubModule. Dibedakan di kolom Feature, bukan TC_ID.',
    height: 90
  },
  {
    label: 'Pola A -- Project Berlayer\n(SIPGN)',
    content:
      'Project  : SIPGN\n' +
      '  Module 1  : Manajemen Gizi\n' +
      '    SubModule 1.1 : Aplikasi Nutritionist\n' +
      '      Feature: Meal Plan, Menu Management\n' +
      '    SubModule 1.2 : Aplikasi Courier\n' +
      '      Feature: Pick Up, Delivery, Return\n' +
      '    SubModule 1.3 : Aplikasi Beneficiary\n' +
      '  Module 2  : Manajemen Distribusi\n' +
      '    SubModule 2.1 : ...',
    height: 115
  },
  {
    label: 'Pola B -- Project Flat\n(INAGOV)',
    content:
      'Project   : INAGOV\n' +
      '  Module  : - (kosong)\n' +
      '    SubModule : Talenta\n' +
      '      Feature: Rekrutmen, Penggajian\n' +
      '    SubModule : e-Office\n' +
      '    SubModule : SIMPEG\n' +
      '\n' +
      'Pada pola flat, SubModule setara dengan Module di pola berlayer.\n' +
      'Kolom Module di Summary dan Config dikosongkan.',
    height: 115
  },
  {
    label: 'TC_ID -- Format:\n[SubModule].[3-digit]',
    content:
      'Opsi 1 -- Numerik (project berlayer seperti SIPGN)\n' +
      '  1.1.001  = SubModule 1.1, TC ke-1\n' +
      '  1.2.015  = SubModule 1.2, TC ke-15\n' +
      '\n' +
      'Opsi 2 -- Inisial (jika SubModule punya nama, bukan nomor)\n' +
      '  Gunakan 2-3 huruf kapital dari nama SubModule:\n' +
      '  PO.001   = Portal, TC ke-1\n' +
      '  SA.001   = Layanan SmartASN, TC ke-1\n' +
      '  BO.001   = BackOffice, TC ke-1\n' +
      '\n' +
      'Aturan inisial:\n' +
      '  - 2-3 huruf kapital, UNIK per project\n' +
      '  - Daftarkan inisial di kolom SubModul Summary\n' +
      '  - Konsisten di TC_Master, API_Master, dan Execution\n' +
      '\n' +
      'Pilih SATU format dan gunakan KONSISTEN dalam satu project.\n' +
      'Harus UNIK. Jangan ubah TC_ID jika sudah ada hasil di Execution.\n' +
      'API prefix wajib: API.1.1.001 / API.PO.001 / API.SA.001',
    height: 230
  }
];

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

  Logger.log('Config: found ID col=' + (COL_ID + 1) + (COL_NAME !== -1 ? ', Name col=' + (COL_NAME + 1) : ''));

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

// ── Main Function: Fix Appendix Content ───────────────────────────────────
/**
 * Update content di section "0. HIERARKI QA" untuk semua modul aktif.
 * Function ini hanya update CONTENT cells, tidak mengubah struktur atau formatting.
 */
function fixAppendixContent() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');

  if (!cfg) {
    safeAlert_('Config tab tidak ditemukan.\nJalankan dari QA Dashboard.');
    return;
  }

  var modules = getModulesFromConfig_(cfg);
  if (modules.length === 0) {
    safeAlert_('Tidak ada modul aktif ditemukan di Config.');
    return;
  }

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

      // Find section header
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
        Logger.log('SKIP ' + name + ': section "0. HIERARKI QA" tidak ditemukan');
        skipped++;
        return;
      }

      Logger.log('UPDATING ' + name + ': found section at row ' + sectionRow);

      // Update each content row
      var currentRow = sectionRow + 1; // Start after header
      var updated = 0;

      CONTENT_ROWS.forEach(function(rowData, idx) {
        // Update label (column A)
        ws.getRange(currentRow, 1).setValue(rowData.label);

        // Merge B:D first (important!)
        ws.getRange(currentRow, 2, 1, 3).merge();

        // Update content (column B, now merged B:D)
        ws.getRange(currentRow, 2).setValue(rowData.content);

        // Set row height
        ws.setRowHeight(currentRow, rowData.height);

        updated++;
        currentRow++;
      });

      SpreadsheetApp.flush();
      Logger.log('OK: ' + name + ' -- updated ' + updated + ' content rows');
      done++;

    } catch (e) {
      Logger.log('ERROR ' + name + ': ' + e.message);
      failed++;
    }
  });

  safeAlert_(
    'fixAppendixContent selesai\n\n' +
    'Berhasil : ' + done    + ' modul\n' +
    'Dilewati : ' + skipped + ' modul\n' +
    'Gagal    : ' + failed  + ' modul\n\n' +
    'Content yang diupdate:\n' +
    '1. Definisi (90px)\n' +
    '2. Pola A -- Project Berlayer (115px)\n' +
    '3. Pola B -- Project Flat (115px)\n' +
    '4. TC_ID Format Reference (230px)\n\n' +
    'Lihat Apps Script Logs untuk detail.'
  );
}
