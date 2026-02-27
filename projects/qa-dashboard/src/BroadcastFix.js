/**
 * broadcast_appendix_fix.js
 * ─────────────────────────────────────────────────────────────────────────
 * Paste ke Apps Script editor QA Dashboard, lalu run broadcastFixAppendix().
 * ─────────────────────────────────────────────────────────────────────────
 */

var SECTION_TITLE = '0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE';

var HIER_ROWS = [
  [
    'Definisi',
    'Project   = Inisiatif / client / program kerja. Contoh: SIPGN, INAGOV\n' +
    'Module    = Pengelompokan domain fungsional dalam project.\n' +
    '            Kosongkan ("-") jika project flat (tidak punya layer domain).\n' +
    'SubModule = Unit terkecil yang berdiri sendiri -- 1 aplikasi atau 1 domain.\n' +
    '            ANCHOR utama untuk TC_ID, Coverage, dan Dashboard.\n' +
    'Feature   = Fitur besar dalam SubModule. Dibedakan di kolom Feature, bukan TC_ID.'
  ],
  [
    'Pola A -- Project Berlayer (SIPGN)',
    'Project  : SIPGN\n' +
    '  Module 1  : Manajemen Gizi\n' +
    '    SubModule 1.1 : Aplikasi Nutritionist\n' +
    '      Feature: Meal Plan, Menu Management\n' +
    '    SubModule 1.2 : Aplikasi Courier\n' +
    '      Feature: Pick Up, Delivery, Return\n' +
    '    SubModule 1.3 : Aplikasi Beneficiary\n' +
    '  Module 2  : Manajemen Distribusi\n' +
    '    SubModule 2.1 : ...'
  ],
  [
    'Pola B -- Project Flat (INAGOV)',
    'Project   : INAGOV\n' +
    '  Module  : - (kosong)\n' +
    '    SubModule : Portal           -> inisial: PO\n' +
    '    SubModule : Layanan SmartASN -> inisial: SA\n' +
    '    SubModule : BackOffice       -> inisial: BO\n' +
    '\n' +
    'Pada pola flat, SubModule setara dengan Module di pola berlayer.\n' +
    'Kolom Module di Summary dan Config dikosongkan.'
  ],
  [
    'Format TC_ID',
    'Format dasar : [SubModule].[3-digit]\n' +
    '\n' +
    'SubModule bisa berupa:\n' +
    '  Numerik : 1.1.001  1.2.015  2.1.001     (pola berlayer)\n' +
    '  Nama    : Talenta.001  eOffice.001       (pola flat, nama pendek)\n' +
    '  Inisial : PO.001  SA.001  BO.001         (pola flat, nama panjang)\n' +
    '\n' +
    'Contoh inisial INAGOV:\n' +
    '  Portal           -> PO   TC_ID: PO.001  PO.002\n' +
    '  Layanan SmartASN -> SA   TC_ID: SA.001  SA.002\n' +
    '  BackOffice       -> BO   TC_ID: BO.001  BO.002\n' +
    '\n' +
    'API prefix wajib: API.PO.001 / API.SA.001 / API.1.1.001\n' +
    '\n' +
    'Aturan:\n' +
    '  - Pilih SATU format dan gunakan konsisten per project\n' +
    '  - Harus UNIK -- jangan pernah reuse TC_ID yang sudah ada\n' +
    '  - Jangan ubah TC_ID jika sudah ada hasil di Execution'
  ],
];

var TC_MASTER_DESC =
  'Master list test case Web / Mobile.\n' +
  'Kolom [INPUT]: SubModul, TC_ID, Feature, Priority, Platform, Test Type, Automation, Version, Role (RBAC), Scenario, Steps, Expected Result.\n' +
  'Kolom [AUTO]: Test Level (kolom N) -- jangan diedit.\n' +
  'Format TC_ID: [SubModule].[3-digit]  contoh: 1.1.001  PO.001  Talenta.001\n' +
  'Gunakan inisial jika nama SubModule terlalu panjang (contoh: PO, SA, BO).\n' +
  'Role = peran RBAC yang menjalankan skenario, contoh: Admin, User, Viewer.';

var API_MASTER_DESC =
  'Master list test case API. Method (E) dan Endpoint URL (F) terpisah.\n' +
  'Kolom [INPUT]: SubModul, TC_ID, Feature, Method, Endpoint, Priority, Auth, Test Type, Automation, Version, Role (RBAC), Scenario.\n' +
  'Kolom [AUTO]: Test Level (kolom N) -- jangan diedit.\n' +
  'Format TC_ID: API.[SubModule].[3-digit]  contoh: API.1.1.001  API.PO.001  API.SA.001\n' +
  'Prefix API wajib. SubModule harus identik dengan TC_Master.\n' +
  'Role = peran RBAC yang diuji aksesnya, contoh: Admin, Super Admin, User, Viewer.';

// ── Auto-detect Spreadsheet ID column from Config header row ─────────────
function getIdColIndex_(headerRow) {
  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i]).trim().toUpperCase();
    if (h === 'SPREADSHEET ID' || h === 'SPREADSHEET_ID' || h === 'ID') {
      return i;
    }
  }
  // Fallback: try col E (index 4) then col F (index 5)
  return 4;
}

// ── Helpers ───────────────────────────────────────────────────────────────
function sectionHeader_(ws, row, title) {
  ws.getRange(row, 1, 1, 4).merge();
  ws.getRange(row, 1)
    .setValue(title)
    .setBackground('#1565C0').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(row, 24);
}

function contentRow_(ws, row, label, desc) {
  ws.getRange(row, 1)
    .setValue(label)
    .setBackground('#E3F2FD').setFontColor('#0D47A1')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true)
    .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
  ws.getRange(row, 2, 1, 3).merge();
  ws.getRange(row, 2)
    .setValue(desc)
    .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(9)
    .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true)
    .setBorder(true,true,true,true,false,false,'#BBDEFB',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(row, 80);
}

function safeAlert_(msg) {
  // getUi() fails when called from certain triggers — use Logger as fallback
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {
    Logger.log('RESULT:\n' + msg);
    Browser.msgBox(msg);
  }
}

// ── Main ──────────────────────────────────────────────────────────────────
function broadcastFixAppendix() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');
  if (!cfg) {
    Logger.log('ERROR: Config tab tidak ditemukan.');
    return;
  }

  var allData   = cfg.getDataRange().getValues();
  var headerRow = allData[2] || []; // row 3 = headers (index 2)
  var idColIdx  = getIdColIndex_(headerRow);

  Logger.log('Config: Spreadsheet ID detected at col index ' + idColIdx +
             ' (col ' + String.fromCharCode(65 + idColIdx) + ')');

  var done = 0, skipped = 0, failed = 0;
  var log  = [];

  for (var i = 3; i < allData.length; i++) {
    var row    = allData[i];
    var active = String(row[0]).trim().toUpperCase();
    var id     = String(row[idColIdx]).trim();

    // Also try adjacent col if id looks wrong
    if (!id || id === '' || id.length < 10) {
      if (idColIdx + 1 < row.length) id = String(row[idColIdx + 1]).trim();
    }

    var name = String(row[3]).trim() || String(row[1]).trim() || ('row ' + (i+1));

    if (active !== 'Y') { skipped++; continue; }
    if (!id || id === 'PASTE_SPREADSHEET_ID_HERE' || id.length < 10) {
      Logger.log('SKIP ' + name + ': Spreadsheet ID tidak valid ("' + id + '")');
      skipped++;
      continue;
    }

    try {
      var modSS = SpreadsheetApp.openById(id);
      var ws    = modSS.getSheetByName('Appendix');

      if (!ws) {
        Logger.log('SKIP ' + name + ': tab Appendix tidak ditemukan');
        skipped++;
        continue;
      }

      // Idempotent check
      var lastRow = ws.getLastRow();
      var colAVals = lastRow > 0
        ? ws.getRange(1, 1, lastRow, 1).getValues().flat().map(String)
        : [];

      if (colAVals.some(function(v) { return v.trim() === SECTION_TITLE; })) {
        Logger.log('SKIP ' + name + ': section sudah ada');
        skipped++;
        continue;
      }

      // Insert rows at position 2
      var insertCount = 1 + HIER_ROWS.length + 1;
      ws.insertRowsBefore(2, insertCount);

      var r = 2;
      sectionHeader_(ws, r, SECTION_TITLE);
      r++;

      HIER_ROWS.forEach(function(rd) {
        contentRow_(ws, r, rd[0], rd[1]);
        r++;
      });
      ws.setRowHeight(r, 8); // spacer
      r++;

      // Update TC_Master and API_Master description rows in Appendix
      var newLast = ws.getLastRow();
      var colA2   = ws.getRange(1, 1, newLast, 1).getValues().flat();
      colA2.forEach(function(cellVal, idx) {
        var v = String(cellVal).trim();
        if (v === 'TC_Master')  ws.getRange(idx+1, 2).setValue(TC_MASTER_DESC);
        if (v === 'API_Master') ws.getRange(idx+1, 2).setValue(API_MASTER_DESC);
      });

      SpreadsheetApp.flush();
      Logger.log('OK: ' + name);
      log.push('OK  ' + name);
      done++;

    } catch(e) {
      Logger.log('ERROR ' + name + ': ' + e.message);
      log.push('ERR ' + name + ': ' + e.message);
      failed++;
    }
  }

  var summary =
    'broadcastFixAppendix selesai\n\n' +
    'Berhasil : ' + done    + ' modul\n' +
    'Dilewati : ' + skipped + ' modul\n' +
    'Gagal    : ' + failed  + ' modul\n\n' +
    'Lihat Apps Script Logs untuk detail.';

  Logger.log(summary);
  safeAlert_(summary);
}