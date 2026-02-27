/**
 * AppendixFullSetup.js
 * ─────────────────────────────────────────────────────────────────────────
 * Setup FULL Appendix content yang lengkap untuk semua modul.
 * Include semua section: Hierarki QA, Status, Priority, RBAC, dll.
 *
 * FUNCTION:
 *   setupFullAppendix() -- Setup complete Appendix dengan semua sections
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

// ── PERURI Branding ───────────────────────────────────────────────────────
var PERURI = {
  primary:   '#0D47A1',
  mid:       '#1565C0',
  accent:    '#1976D2',
  light:     '#E3F2FD',
  text:      '#FFFFFF',
  copyright: '(c) QA INA Digital  |  Template ini merupakan properti QA Team INA Digital  |  Dilarang digunakan/disebarluaskan tanpa izin  |  departemen.qa@inadigital.co.id'
};

// ── Appendix Content Sections ─────────────────────────────────────────────
var APPENDIX_SECTIONS = [
  {
    title: 'APPENDIX  .  Definisi, Konvensi & Panduan',
    isSectionHeader: true,
    height: 28
  },
  {
    title: '0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE',
    rows: [
      {
        label: 'Definisi',
        content: 'Project   = Inisiatif / client / program kerja. Contoh: SIPGN, INAGOV\n' +
                 'Module    = Pengelompokan domain fungsional dalam project.\n' +
                 '            Kosongkan ("-") jika project flat (tidak punya layer domain).\n' +
                 'SubModule = Unit terkecil yang berdiri sendiri -- 1 aplikasi atau 1 domain.\n' +
                 '            ANCHOR utama untuk TC_ID, Coverage, dan Dashboard.\n' +
                 'Feature   = Fitur besar dalam SubModule. Dibedakan di kolom Feature, bukan TC_ID.',
        height: 90
      },
      {
        label: 'Pola A -- Project Berlayer\n(SIPGN)',
        content: 'Project  : SIPGN\n' +
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
        content: 'Project   : INAGOV\n' +
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
        content: 'Opsi 1 -- Numerik (project berlayer seperti SIPGN)\n' +
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
    ]
  },
  {
    title: '1. STRUKTUR TAB',
    rows: [
      {
        label: 'PerfTest',
        content: 'Rekam hasil performance test. Isi threshold di baris 11 sesuai SLA.\n' +
                 'Kolom STATUS otomatis PASS/FAIL per skenario. OVERALL RESULT terhubung ke tab Summary.\n' +
                 'Metrik yang dicek: RPS, Error Rate, P90, P95, P99, VU, CPU, Memory.',
        height: 60
      },
      {
        label: 'Appendix',
        content: 'Dokumen ini.',
        height: 30
      }
    ]
  },
  {
    title: '2. STATUS EKSEKUSI',
    rows: [
      {
        label: 'PASSED',
        content: 'Skenario berhasil -- hasil sesuai Expected Result.',
        height: 30
      },
      {
        label: 'IN PROGRESS',
        content: 'Ada skenario yang sudah PASSED di run sebelumnya namun masih ada TODO -- eksekusi belum selesai.',
        height: 40
      },
      {
        label: 'FAILED',
        content: 'Skenario gagal -- hasil tidak sesuai Expected Result. Wajib buat bug report.',
        height: 40
      },
      {
        label: 'BLOCKED',
        content: 'Tidak bisa dieksekusi -- environment down, dependensi belum siap, atau data belum ada.',
        height: 40
      },
      {
        label: 'TODO',
        content: 'Belum dieksekusi pada run ini.',
        height: 30
      }
    ]
  },
  {
    title: '3. TEST LEVEL  .  OTOMATIS DARI PRIORITY',
    rows: [
      {
        label: 'Smoke',
        content: 'Priority Critical / High / Medium → Smoke.\n' +
                 'Test subset cepat untuk memvalidasi fungsi utama sebelum release atau setelah deployment.',
        height: 50
      },
      {
        label: 'Regression',
        content: 'Priority Low / Lowest → Regression.\n' +
                 'Full test cycle dijalankan sebelum release ke Production.',
        height: 50
      }
    ]
  },
  {
    title: '4. PRIORITY',
    rows: [
      {
        label: 'Critical  [BLOCKER]',
        content: 'Fungsi utama tidak bisa digunakan. Release DITAHAN jika FAIL/BLOCKED.',
        height: 40
      },
      {
        label: 'High      [BLOCKER]',
        content: 'Fungsi penting terganggu. Jika FAIL: perlu approval PM untuk release.',
        height: 40
      },
      {
        label: 'Medium    [POTENTIAL BLOCKER]',
        content: 'Fungsi terganggu sebagian. Jika FAIL sebelum UAT: flagged ke tech lead.',
        height: 40
      },
      {
        label: 'Low',
        content: 'Masalah minor: UI, typo, UX kurang optimal.',
        height: 30
      },
      {
        label: 'Lowest',
        content: 'Nice to have. Fix opsional.',
        height: 30
      }
    ]
  },
  {
    title: '5. ROLE (RBAC) -- KONTROL AKSES',
    rows: [
      {
        label: 'Tujuan',
        content: 'Kolom Role di TC_Master & API_Master bukan role developer, melainkan PERAN PENGGUNA (RBAC) yang menjalankan skenario tersebut.\n\n' +
                 'Digunakan untuk memverifikasi bahwa access control berjalan benar.',
        height: 60
      },
      {
        label: 'Contoh Role',
        content: 'Admin: bisa create/edit/delete semua data\n' +
                 'User/Operator: bisa create & edit data sendiri\n' +
                 'Viewer: hanya bisa read/view\n' +
                 'Guest: akses terbatas (belum login)\n' +
                 'Supervisor: approval flow\n' +
                 'Super Admin: full access termasuk system settings',
        height: 90
      },
      {
        label: 'Skenario RBAC',
        content: 'Untuk setiap endpoint/fitur sensitif, buat TC untuk:\n' +
                 '1. Role yang BERHAK -- harus dapat akses (expected: 200/201)\n' +
                 '2. Role yang TIDAK berhak -- harus ditolak (expected: 403 Forbidden)\n' +
                 '3. Tanpa token/login -- harus ditolak (expected: 401 Unauthorized)',
        height: 75
      }
    ]
  },
  {
    title: '6. AUTOMATION STATUS',
    rows: [
      {
        label: 'Automated',
        content: 'Script sudah dibuat dan dapat dijalankan via CI/CD atau manual run.',
        height: 40
      },
      {
        label: 'Manual',
        content: 'Tidak akan diautomasi -- memerlukan penilaian manusia (exploratory, visual, UX).',
        height: 40
      },
      {
        label: 'To Do',
        content: 'Direncanakan untuk diautomasi, belum dikerjakan.',
        height: 30
      },
      {
        label: 'Cannot be Automated',
        content: 'Secara teknis tidak bisa diautomasi (scan QR fisik, biometrik, hardware-dependent).',
        height: 40
      }
    ]
  },
  {
    title: '9. PERFORMANCE TEST -- METRIK',
    rows: [
      {
        label: 'RPS (req/s)',
        content: 'Requests Per Second -- jumlah request per detik yang berhasil diproses.\n' +
                 'Threshold: MINIMUM. Jika actual < threshold → FAIL.',
        height: 50
      },
      {
        label: 'Error Rate (%)',
        content: 'Persentase request yang mengembalikan error (status ≥ 400 atau timeout).\n' +
                 'Threshold: MAKSIMUM. Jika actual > threshold → FAIL.',
        height: 50
      },
      {
        label: 'P90 / P95 / P99',
        content: 'Percentile response time -- 90% / 95% / 99% dari semua request selesai dalam waktu ini.\n' +
                 'Threshold: MAKSIMUM (ms). Semakin kecil semakin baik.',
        height: 50
      },
      {
        label: 'VU (Virtual User)',
        content: 'Jumlah concurrent user yang disimulasikan dalam satu skenario test.\n' +
                 'Diisi sebagai config, bukan sebagai threshold PASS/FAIL.',
        height: 50
      },
      {
        label: 'CPU Usage (%)',
        content: 'Persentase penggunaan CPU server selama test berlangsung.\n' +
                 'Threshold: MAKSIMUM. Jika actual > threshold → FAIL.',
        height: 50
      },
      {
        label: 'Memory Usage (%)',
        content: 'Persentase penggunaan memory server selama test berlangsung.\n' +
                 'Threshold: MAKSIMUM. Jika actual > threshold → FAIL.',
        height: 50
      }
    ]
  },
  {
    title: '10. HTTP METHOD',
    rows: [
      {
        label: 'GET',
        content: 'Mengambil data. Tidak mengubah state.',
        height: 30
      },
      {
        label: 'POST',
        content: 'Membuat resource baru.',
        height: 30
      },
      {
        label: 'PUT',
        content: 'Update resource secara penuh (replace).',
        height: 30
      },
      {
        label: 'PATCH',
        content: 'Update resource sebagian (partial).',
        height: 30
      },
      {
        label: 'DELETE',
        content: 'Menghapus resource.',
        height: 30
      }
    ]
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

// ── Main Function: Setup Full Appendix ────────────────────────────────────
/**
 * Setup complete Appendix content untuk semua modul aktif.
 * Menghapus content lama dan rebuild dari scratch dengan format proper.
 */
function setupFullAppendix() {
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

      Logger.log('SETUP ' + name + ': building full Appendix...');

      // Clear existing content (keep row 1 as header row)
      var lastRow = ws.getLastRow();
      if (lastRow > 1) {
        ws.deleteRows(2, lastRow - 1);
        SpreadsheetApp.flush();
      }

      var currentRow = 1;

      // Main title row
      ws.insertRowsBefore(currentRow + 1, 1);
      currentRow++;
      ws.getRange(currentRow, 1, 1, 4).merge();
      ws.getRange(currentRow, 1)
        .setValue('APPENDIX  .  Definisi, Konvensi & Panduan')
        .setBackground(PERURI.primary).setFontColor(PERURI.text)
        .setFontWeight('bold').setFontSize(11).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      ws.setRowHeight(currentRow, 32);
      currentRow++;

      // Process each section
      APPENDIX_SECTIONS.forEach(function(section) {
        if (section.isSectionHeader) return; // Skip, already added

        // Section title
        ws.insertRowsBefore(currentRow + 1, 1);
        currentRow++;
        ws.getRange(currentRow, 1, 1, 4).merge();
        ws.getRange(currentRow, 1)
          .setValue(section.title)
          .setBackground(PERURI.mid).setFontColor(PERURI.text)
          .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
          .setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);
        ws.setRowHeight(currentRow, 24);
        currentRow++;

        // Section rows
        section.rows.forEach(function(rowData) {
          ws.insertRowsBefore(currentRow + 1, 1);
          currentRow++;

          // Label (column A)
          ws.getRange(currentRow, 1)
            .setValue(rowData.label)
            .setBackground(PERURI.light).setFontColor(PERURI.primary)
            .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
            .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true)
            .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);

          // Content (columns B:D merged)
          ws.getRange(currentRow, 2, 1, 3).merge();
          ws.getRange(currentRow, 2)
            .setValue(rowData.content)
            .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(9)
            .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true)
            .setBorder(true, true, true, true, false, false, '#BBDEFB', SpreadsheetApp.BorderStyle.SOLID);

          ws.setRowHeight(currentRow, rowData.height);
        });

        // Spacer after each section
        ws.insertRowsBefore(currentRow + 1, 1);
        currentRow++;
        ws.setRowHeight(currentRow, 6);
      });

      // Copyright footer
      ws.insertRowsBefore(currentRow + 1, 2);
      currentRow++;
      currentRow++;
      ws.getRange(currentRow, 1, 1, 4).merge();
      ws.getRange(currentRow, 1)
        .setValue(PERURI.copyright)
        .setBackground('#F5F5F5').setFontColor('#757575')
        .setFontSize(8).setFontFamily('Arial').setFontStyle('italic')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true, true, true, true, false, false, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
      ws.setRowHeight(currentRow, 40);

      // Set column widths
      ws.setColumnWidth(1, 180); // Label column
      ws.setColumnWidth(2, 500); // Content column (merged B:D)
      ws.setColumnWidth(3, 100);
      ws.setColumnWidth(4, 100);

      SpreadsheetApp.flush();
      Logger.log('OK: ' + name + ' -- full Appendix setup complete');
      done++;

    } catch (e) {
      Logger.log('ERROR ' + name + ': ' + e.message);
      failed++;
    }
  });

  safeAlert_(
    'setupFullAppendix selesai\n\n' +
    'Berhasil : ' + done    + ' modul\n' +
    'Dilewati : ' + skipped + ' modul\n' +
    'Gagal    : ' + failed  + ' modul\n\n' +
    'Appendix sekarang memiliki:\n' +
    '- 10 sections lengkap\n' +
    '- Proper formatting & colors\n' +
    '- Copyright footer\n' +
    '- Merged cells B:D untuk content\n\n' +
    'Lihat Apps Script Logs untuk detail.'
  );
}
