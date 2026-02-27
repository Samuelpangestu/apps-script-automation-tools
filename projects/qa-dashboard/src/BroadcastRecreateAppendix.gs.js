/**
 * broadcast_recreate_appendix.js
 * ─────────────────────────────────────────────────────────────────────────
 * Paste ke Apps Script editor QA Dashboard.
 * Run broadcastRecreateAppendix() untuk delete + recreate tab Appendix
 * di semua modul aktif dengan konten lengkap & terkini.
 * ─────────────────────────────────────────────────────────────────────────
 */

// ── Auto-detect Spreadsheet ID column ────────────────────────────────────
function getIdColIdx_(headerRow) {
  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i]).trim().toUpperCase();
    if (h === 'SPREADSHEET ID' || h === 'SPREADSHEET_ID' || h === 'ID') return i;
  }
  return 4; // fallback col E
}

function safeAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch(e) { try { Browser.msgBox(msg); } catch(e2) { Logger.log(msg); } }
}

// ── Style helpers ─────────────────────────────────────────────────────────
function apxHdr_(ws, r, title, bg) {
  ws.getRange(r,1,1,4).merge();
  ws.getRange(r,1).setValue(title)
    .setBackground(bg||'#37474F').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(r,24);
}

function apxRow_(ws, r, label, desc, labelBg, rowH) {
  ws.getRange(r,1).setValue(label)
    .setBackground(labelBg||'#ECEFF1').setFontColor('#000000')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true)
    .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  ws.getRange(r,2,1,3).merge();
  ws.getRange(r,2).setValue(desc)
    .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(9)
    .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true)
    .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(r, rowH||52);
}

function apxStatusRow_(ws, r, label, desc, bg, rowH) {
  ws.getRange(r,1).setValue(label)
    .setBackground(bg||'#F5F5F5').setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  ws.getRange(r,2,1,3).merge();
  ws.getRange(r,2).setValue(desc)
    .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(9)
    .setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(r, rowH||34);
}

// ── Build full Appendix content ───────────────────────────────────────────
function buildAppendix_(ss) {
  // Delete existing Appendix
  var existing = ss.getSheetByName('Appendix');
  if (existing) ss.deleteSheet(existing);

  // Re-create
  var ws = ss.insertSheet('Appendix');
  ws.setTabColor('#1A237E');
  ss.moveActiveSheet(ss.getSheets().length); // move to end

  // Title row
  ws.getRange(1,1,1,4).merge();
  ws.getRange(1,1).setValue('APPENDIX  .  Definisi, Konvensi & Panduan')
    .setBackground('#0D47A1').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  ws.setRowHeight(1,30);

  var r = 2;

  // ── SECTION 0: HIERARKI QA ──────────────────────────────────────────
  apxHdr_(ws, r, '0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE', '#0D47A1'); r++;

  apxRow_(ws, r, 'Definisi',
    'Project   = Inisiatif / client / program kerja. Contoh: SIPGN, INAGOV\n' +
    'Module    = Pengelompokan domain fungsional dalam project.\n' +
    '            Kosongkan ("-") jika project flat (tidak punya layer domain).\n' +
    'SubModule = Unit terkecil yang berdiri sendiri -- 1 aplikasi atau 1 domain.\n' +
    '            ANCHOR utama untuk TC_ID, Coverage, dan Dashboard.\n' +
    'Feature   = Fitur besar dalam SubModule. Dibedakan di kolom Feature, bukan TC_ID.',
    '#E3F2FD', 72); r++;

  apxRow_(ws, r, 'Pola A -- Project Berlayer (SIPGN)',
    'Project  : SIPGN\n' +
    '  Module 1  : Manajemen Gizi\n' +
    '    SubModule 1.1 : Aplikasi Nutritionist   Feature: Meal Plan, Menu Management\n' +
    '    SubModule 1.2 : Aplikasi Courier        Feature: Pick Up, Delivery, Return\n' +
    '    SubModule 1.3 : Aplikasi Beneficiary\n' +
    '  Module 2  : Manajemen Distribusi\n' +
    '    SubModule 2.1 : ...',
    '#E3F2FD', 80); r++;

  apxRow_(ws, r, 'Pola B -- Project Flat (INAGOV)',
    'Project   : INAGOV\n' +
    '  Module  : - (kosong)\n' +
    '    SubModule : Portal            inisial: PO\n' +
    '    SubModule : Layanan SmartASN  inisial: SA\n' +
    '    SubModule : BackOffice        inisial: BO\n' +
    '\n' +
    'Pada pola flat, SubModule setara dengan Module di pola berlayer.\n' +
    'Kolom Module di Summary dan Config dikosongkan.',
    '#E3F2FD', 80); r++;

  apxRow_(ws, r, 'Format TC_ID',
    'Format dasar : [SubModule].[3-digit]\n' +
    '\n' +
    'SubModule bisa berupa:\n' +
    '  Numerik : 1.1.001  1.2.015  2.1.001       (pola berlayer)\n' +
    '  Nama    : Talenta.001  eOffice.001         (pola flat, nama pendek)\n' +
    '  Inisial : PO.001  SA.001  BO.001           (pola flat, nama panjang)\n' +
    '\n' +
    'Contoh inisial INAGOV:\n' +
    '  Portal           -> PO    TC_ID: PO.001  PO.002\n' +
    '  Layanan SmartASN -> SA    TC_ID: SA.001  SA.002\n' +
    '  BackOffice       -> BO    TC_ID: BO.001  BO.002\n' +
    '\n' +
    'API prefix wajib: API.PO.001 / API.SA.001 / API.1.1.001\n' +
    '\n' +
    'Aturan:\n' +
    '  - Pilih SATU format dan gunakan konsisten per project\n' +
    '  - Harus UNIK -- jangan pernah reuse TC_ID yang sudah ada\n' +
    '  - Jangan ubah TC_ID jika sudah ada hasil di Execution\n' +
    '  - Urutan: Positive (001), Negative (002), Edge Case (003)',
    '#E3F2FD', 120); r++;
  r++; // spacer

  // ── SECTION 1: STRUKTUR TAB ─────────────────────────────────────────
  apxHdr_(ws, r, '1. STRUKTUR TAB', '#0D47A1'); r++;

  apxRow_(ws, r, 'TC_Master',
    'Master list test case Web / Mobile.\n' +
    'Kolom [INPUT]: SubModul, TC_ID, Feature, Priority, Platform, Test Type, Automation, Version, Role (RBAC), Scenario, Steps, Expected Result.\n' +
    'Kolom [AUTO]: Test Level (kolom N) -- jangan diedit.\n' +
    'Format TC_ID: [SubModule].[3-digit]  contoh: 1.1.001  PO.001  Talenta.001\n' +
    'Gunakan inisial jika nama SubModule terlalu panjang (contoh: PO, SA, BO).\n' +
    'Role = peran RBAC yang menjalankan skenario, contoh: Admin, User, Viewer.',
    null, 72); r++;

  apxRow_(ws, r, 'TC_Execution',
    'Kolom identitas sync otomatis dari TC_Master. Isi kolom staging dengan PASSED / FAILED / BLOCKED / TODO.\n' +
    'Tambah kolom ke kanan untuk setiap run. LATEST STATUS di kolom Z otomatis.\n' +
    'Kolom Screenshot auto-link ke tanggal run.\n' +
    'IN PROGRESS otomatis jika ada PASSED dan TODO di skenario yang sama.',
    null, 60); r++;

  apxRow_(ws, r, 'API_Master',
    'Master list test case API. Method (E) dan Endpoint URL (F) terpisah.\n' +
    'Kolom [INPUT]: SubModul, TC_ID, Feature, Method, Endpoint, Priority, Auth, Test Type, Automation, Version, Role (RBAC), Scenario.\n' +
    'Kolom [AUTO]: Test Level (kolom N) -- jangan diedit.\n' +
    'Format TC_ID: API.[SubModule].[3-digit]  contoh: API.1.1.001  API.PO.001  API.SA.001\n' +
    'Prefix API wajib. SubModule harus identik dengan TC_Master.\n' +
    'Role = peran RBAC yang diuji aksesnya, contoh: Admin, Super Admin, User, Viewer.',
    null, 72); r++;

  apxRow_(ws, r, 'API_Execution',
    'Sama seperti TC_Execution namun untuk API. Sync otomatis dari API_Master.\n' +
    'Kolom Screenshot auto-link ke tanggal run.',
    null, 44); r++;

  apxRow_(ws, r, 'Summary',
    'Isi bagian Test Plan sebelum memulai eksekusi (Project, Modul, PIC, Jira, status).\n' +
    'Ringkasan otomatis: coverage SubModul & Feature, run history dengan IN PROGRESS.\n' +
    'Perf Test Status otomatis dari tab PerfTest.',
    null, 52); r++;

  apxRow_(ws, r, 'PerfTest',
    'Rekam hasil performance test. Isi threshold di baris 11 sesuai SLA.\n' +
    'Kolom STATUS otomatis PASS/FAIL per skenario. OVERALL RESULT terhubung ke tab Summary.\n' +
    'Metrik yang dicek: RPS, Error Rate, P90, P95, P99, VU, CPU, Memory.',
    null, 52); r++;

  apxRow_(ws, r, 'BugReport',
    'Rekam semua bug yang ditemukan selama eksekusi. Kolom Type, Priority, Status, Linked TC_ID.\n' +
    'Summary bug count (Critical, High, Medium) otomatis tampil di Summary dan Dashboard.',
    null, 44); r++;

  apxRow_(ws, r, 'Appendix', 'Dokumen ini -- panduan konvensi, format, dan aturan pengisian.', null, 34); r++;
  r++; // spacer

  // ── SECTION 2: STATUS EKSEKUSI ──────────────────────────────────────
  apxHdr_(ws, r, '2. STATUS EKSEKUSI', '#0D47A1'); r++;
  [
    ['PASSED',      'Skenario berhasil -- hasil sesuai Expected Result.', '#C8E6C9'],
    ['IN PROGRESS', 'Ada skenario PASSED di run sebelumnya namun masih ada TODO -- eksekusi belum selesai.', '#E3F2FD'],
    ['FAILED',      'Skenario gagal -- hasil tidak sesuai Expected Result. Wajib buat bug report.', '#FFCDD2'],
    ['BLOCKED',     'Tidak bisa dieksekusi -- environment down, dependensi belum siap, atau data belum ada.', '#FFE0B2'],
    ['TODO',        'Belum dieksekusi pada run ini.', '#F5F5F5'],
  ].forEach(function(row) {
    apxStatusRow_(ws, r, row[0], row[1], row[2]); r++;
  });
  r++;

  // ── SECTION 3: SCENARIO NAMING STANDARD ─────────────────────────────
  apxHdr_(ws, r, '3. SCENARIO NAMING STANDARD & GHERKIN FORMAT', '#0D47A1'); r++;

  apxRow_(ws, r, 'Formula',
    'Happy Path : [Role] + Successfully + [Verb] + [Object] + (from/to [Location])\n' +
    'Negative   : [Role] + Failed to + [Verb] + [Object] + with [Condition]\n' +
    '\n' +
    'RULES\n' +
    '  - Role   -> Title Case  (Nutritionist, Courier, Beneficiary)\n' +
    '  - Verb   -> Active verb (Create, Pick Up, Confirm, Return)\n' +
    '  - Object -> Title Case  (Menu Plan, Meal Box, Food)\n' +
    '  - Location -> Optional, gunakan "from" atau "to"\n' +
    '\n' +
    'DO NOT USE\n' +
    '  X  success, succeed     (use: Successfully / Failed to)\n' +
    '  X  do, perform, process (sebelum verb utama)',
    null, 100); r++;

  apxRow_(ws, r, 'Gherkin Format',
    'Kolom Scenario -- tulis Given dan When saja:\n' +
    '  Given : Pre-kondisi / state awal sebelum aksi dimulai\n' +
    '          Contoh: Given user sudah login sebagai Courier\n' +
    '  When  : Aksi yang dilakukan oleh aktor\n' +
    '          Contoh: When user klik tombol Pick Up\n' +
    '  And   : Aksi tambahan (opsional)\n' +
    '\n' +
    'Kolom Expected Result -- tulis Then:\n' +
    '  Then  : Hasil / perubahan state setelah aksi selesai\n' +
    '          Contoh: Then Meal Box berhasil di-pickup dan status berubah ke On Delivery\n' +
    '\n' +
    'JANGAN tulis Then di kolom Scenario -- Then hanya di Expected Result.',
    null, 100); r++;

  apxRow_(ws, r, 'Contoh Lengkap',
    'Scenario  : Courier Successfully Picks Up Food from SPPG\n' +
    'Steps     : Given Courier sudah login dan ada pickup request\n' +
    '            When Courier membuka detail order dan klik "Pick Up"\n' +
    '            And Courier konfirmasi jumlah Meal Box\n' +
    'Expected  : Then status order berubah ke "On Delivery"\n' +
    '            Then notifikasi dikirim ke Beneficiary\n' +
    '\n' +
    'Scenario  : Nutritionist Failed to Create Menu with Incomplete Data\n' +
    'Steps     : Given Nutritionist sudah login\n' +
    '            When Nutritionist membuka form Create Menu\n' +
    '            And Nutritionist submit form tanpa mengisi field wajib\n' +
    'Expected  : Then form tidak tersimpan\n' +
    '            Then muncul pesan error pada field yang kosong',
    '#F1F8E9', 110); r++;
  r++;

  // ── SECTION 4: TEST LEVEL ───────────────────────────────────────────
  apxHdr_(ws, r, '4. TEST LEVEL  .  OTOMATIS DARI PRIORITY', '#546E7A'); r++;
  apxRow_(ws, r, 'Smoke',
    'Priority Critical / High / Medium -> Smoke.\n' +
    'Test subset cepat untuk memvalidasi fungsi utama sebelum release atau setelah deployment.',
    '#FFF8F0', 44); r++;
  apxRow_(ws, r, 'Regression',
    'Priority Low / Lowest -> Regression.\n' +
    'Full test cycle dijalankan sebelum release ke Production.',
    '#F1F8E9', 44); r++;
  r++;

  // ── SECTION 5: PRIORITY ─────────────────────────────────────────────
  apxHdr_(ws, r, '5. PRIORITY', '#0D47A1'); r++;
  [
    ['Critical  [BLOCKER]',          'Fungsi utama tidak bisa digunakan. Release DITAHAN jika FAIL/BLOCKED.', '#FFEBEE'],
    ['High      [BLOCKER]',          'Fungsi penting terganggu. Jika FAIL: perlu approval PM untuk release.', '#FFF3E0'],
    ['Medium    [POTENTIAL BLOCKER]','Fungsi terganggu sebagian. Jika FAIL sebelum UAT: flagged ke tech lead.', '#FFF8E1'],
    ['Low',                          'Masalah minor: UI, typo, UX kurang optimal.', '#F1F8E9'],
    ['Lowest',                       'Nice to have. Fix opsional.', '#ECEFF1'],
  ].forEach(function(row) {
    apxStatusRow_(ws, r, row[0], row[1], row[2]); r++;
  });
  r++;

  // ── SECTION 6: ROLE (RBAC) ──────────────────────────────────────────
  apxHdr_(ws, r, '6. ROLE (RBAC) -- KONTROL AKSES', '#0D47A1'); r++;
  apxRow_(ws, r, 'Tujuan',
    'Kolom Role di TC_Master & API_Master bukan role developer, melainkan PERAN PENGGUNA (RBAC)\n' +
    'yang menjalankan skenario tersebut.\n' +
    'Digunakan untuk memverifikasi bahwa access control berjalan benar.',
    null, 52); r++;
  apxRow_(ws, r, 'Contoh Role',
    'Admin       : bisa create/edit/delete semua data\n' +
    'User/Operator: bisa create & edit data sendiri\n' +
    'Viewer      : hanya bisa read/view\n' +
    'Guest       : akses terbatas (belum login)\n' +
    'Supervisor  : approval flow\n' +
    'Super Admin : full access termasuk system settings',
    null, 72); r++;
  apxRow_(ws, r, 'Skenario RBAC',
    'Untuk setiap endpoint/fitur sensitif, buat TC untuk:\n' +
    '1. Role yang BERHAK     -- harus dapat akses (expected: 200/201)\n' +
    '2. Role yang TIDAK berhak -- harus ditolak  (expected: 403 Forbidden)\n' +
    '3. Tanpa token/login    -- harus ditolak    (expected: 401 Unauthorized)',
    null, 60); r++;
  r++;

  // ── SECTION 7: AUTOMATION STATUS ───────────────────────────────────
  apxHdr_(ws, r, '7. AUTOMATION STATUS', '#0D47A1'); r++;
  [
    ['Automated',           'Script sudah dibuat dan dapat dijalankan via CI/CD atau manual run.'],
    ['Manual',              'Tidak akan diautomasi -- memerlukan penilaian manusia (exploratory, visual, UX).'],
    ['To Do',               'Direncanakan untuk diautomasi, belum dikerjakan.'],
    ['Cannot be Automated', 'Secara teknis tidak bisa diautomasi (scan QR fisik, biometrik, hardware-dependent).'],
  ].forEach(function(row) {
    apxRow_(ws, r, row[0], row[1], null, 44); r++;
  });
  r++;

  // ── SECTION 8: BUG REPORT ───────────────────────────────────────────
  apxHdr_(ws, r, '8. BUG REPORT -- PANDUAN PENGISIAN', '#B71C1C'); r++;
  apxRow_(ws, r, 'Kapan buat bug?',
    'Wajib buat bug report untuk setiap TC dengan status FAILED.\n' +
    'Opsional untuk BLOCKED jika penyebabnya adalah bug (bukan environment).',
    null, 44); r++;
  apxRow_(ws, r, 'Bug Type',
    'Functional   : Fungsi tidak berjalan sesuai spesifikasi\n' +
    'UI/UX        : Tampilan rusak, layout salah, teks terpotong\n' +
    'Performance  : Response time melebihi threshold\n' +
    'Security     : Akses tidak terotorisasi, data bocor\n' +
    'Data         : Data salah, tidak konsisten, atau hilang\n' +
    'Integration  : Gagal di titik integrasi antar service/API',
    null, 72); r++;
  apxRow_(ws, r, 'Bug Priority',
    'Critical : Blocker -- sistem crash, data corruption, tidak bisa login\n' +
    'High     : Fungsi utama tidak bisa digunakan\n' +
    'Medium   : Fungsi terganggu sebagian, ada workaround\n' +
    'Low      : UI issue, typo, minor UX\n' +
    'Lowest   : Nice to fix, tidak mengganggu operasional',
    null, 60); r++;
  r++;

  // ── SECTION 9: PERFORMANCE TEST ─────────────────────────────────────
  apxHdr_(ws, r, '9. PERFORMANCE TEST -- METRIK', '#4A148C'); r++;
  [
    ['RPS (req/s)',     'Requests Per Second -- jumlah request per detik yang berhasil diproses.\nThreshold: MINIMUM. Jika actual < threshold -> FAIL.'],
    ['Error Rate (%)',  'Persentase request yang mengembalikan error (status >= 400 atau timeout).\nThreshold: MAKSIMUM. Jika actual > threshold -> FAIL.'],
    ['P90 / P95 / P99','Percentile response time -- 90%/95%/99% dari semua request selesai dalam waktu ini.\nThreshold: MAKSIMUM (ms). Semakin kecil semakin baik.'],
    ['VU (Virtual User)','Jumlah concurrent user yang disimulasikan dalam satu skenario test.\nDiisi sebagai config, bukan sebagai threshold PASS/FAIL.'],
    ['CPU Usage (%)',   'Persentase penggunaan CPU server selama test berlangsung.\nThreshold: MAKSIMUM. Jika actual > threshold -> FAIL.'],
    ['Memory Usage (%)','Persentase penggunaan memory server selama test berlangsung.\nThreshold: MAKSIMUM. Jika actual > threshold -> FAIL.'],
  ].forEach(function(row) {
    apxRow_(ws, r, row[0], row[1], null, 52); r++;
  });
  r++;

  // ── SECTION 10: HTTP METHOD ──────────────────────────────────────────
  apxHdr_(ws, r, '10. HTTP METHOD', '#0D47A1'); r++;
  [
    ['GET',    'Mengambil data. Tidak mengubah state.',              '#E8F0FE'],
    ['POST',   'Membuat resource baru.',                             '#E8F5E9'],
    ['PUT',    'Update resource secara penuh (replace).',            '#FFF8E1'],
    ['PATCH',  'Update resource sebagian (partial).',                '#F3E5F5'],
    ['DELETE', 'Menghapus resource.',                                '#FCE4EC'],
  ].forEach(function(row) {
    apxStatusRow_(ws, r, row[0], row[1], row[2], 28); r++;
  });

  // Column widths
  [100, 140, 80, 200].forEach(function(w, i) { ws.setColumnWidth(i+1, w); });

  return r;
}

// ── Main broadcast ────────────────────────────────────────────────────────
function broadcastRecreateAppendix() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');
  if (!cfg) {
    Logger.log('ERROR: Config tab tidak ditemukan.');
    safeAlert_('Config tab tidak ditemukan.\nPastikan dijalankan dari QA Dashboard.');
    return;
  }

  var allData   = cfg.getDataRange().getValues();
  var headerRow = allData[2] || [];
  var idColIdx  = getIdColIdx_(headerRow);
  Logger.log('Config: ID col detected at index ' + idColIdx +
             ' (col ' + String.fromCharCode(65 + idColIdx) + ')');

  var done = 0, skipped = 0, failed = 0;

  for (var i = 3; i < allData.length; i++) {
    var row    = allData[i];
    var active = String(row[0]).trim().toUpperCase();
    var id     = String(row[idColIdx]).trim();

    // Try adjacent col if id looks wrong
    if ((!id || id.length < 10) && idColIdx + 1 < row.length) {
      id = String(row[idColIdx + 1]).trim();
    }

    var name = String(row[3]).trim() || String(row[1]).trim() || ('row ' + (i+1));

    if (active !== 'Y') { skipped++; continue; }
    if (!id || id === 'PASTE_SPREADSHEET_ID_HERE' || id.length < 10) {
      Logger.log('SKIP ' + name + ': ID tidak valid ("' + id + '")');
      skipped++;
      continue;
    }

    try {
      var modSS = SpreadsheetApp.openById(id);
      buildAppendix_(modSS);
      SpreadsheetApp.flush();
      Logger.log('OK: ' + name);
      done++;
    } catch(e) {
      Logger.log('ERROR ' + name + ': ' + e.message);
      failed++;
    }
  }

  var summary =
    'broadcastRecreateAppendix selesai\n\n' +
    'Berhasil : ' + done    + ' modul\n' +
    'Dilewati : ' + skipped + ' modul (inactive / ID kosong)\n' +
    'Gagal    : ' + failed  + ' modul\n\n' +
    'Lihat Apps Script Logs untuk detail.';

  Logger.log(summary);
  safeAlert_(summary);
}