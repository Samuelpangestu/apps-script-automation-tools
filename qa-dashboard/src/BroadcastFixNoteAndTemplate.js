/**
 * broadcast_all_fixes.js
 * ─────────────────────────────────────────────────────────────────────────
 * Paste ke Apps Script editor QA PORTFOLIO DASHBOARD (bukan modul).
 * Satu file untuk semua perubahan terbaru.
 *
 * FUNGSI UTAMA (jalankan satu-satu):
 *   1. broadcastAllFixes()         → jalankan semua fix sekaligus (recommended)
 *   ── atau satu per satu: ──
 *   2. fixDashboardLeadColumn()    → tambah kolom QA Lead di Config + Overview Dashboard
 *   3. broadcastApiExpectedResult()→ tambah kolom Expected Result di API_Master semua modul
 *   4. broadcastRecreateAppendix() → recreate Appendix semua modul (+ section Gherkin Examples)
 *
 * PERUBAHAN:
 *   Dashboard Config  : tambah kolom "QA Lead" (F) antara PIC/Team dan Spreadsheet ID
 *   Dashboard Overview: tambah kolom "Lead" di group MODULE INFO
 *   API_Master        : tambah kolom "Expected Result" (N) antara Scenario dan Examples
 *   Appendix          : section baru 3b — Gherkin Examples / Scenario Outline
 *
 * SAFE:
 *   - Setiap fix cek apakah sudah dilakukan → skip jika sudah
 *   - Tidak menghapus data yang ada
 * ─────────────────────────────────────────────────────────────────────────
 */

// ══════════════════════════════════════════════════════════════════════════
// MASTER RUNNER — jalankan semua fix
// ══════════════════════════════════════════════════════════════════════════
function broadcastAllFixes() {
  var results = [];
  Logger.log('broadcastAllFixes mulai...');

  // Step 1: Dashboard Lead Column
  try {
    var dashResult = fixDashboardLeadColumn_();
    results.push('✅  Dashboard Lead  : ' + dashResult);
  } catch(e) {
    results.push('❌  Dashboard Lead  : ' + e.message);
    Logger.log('Dashboard Lead ERROR: ' + e.message);
  }

  // Step 2: API Expected Result (broadcast to all modules)
  try {
    var apiResult = broadcastApiExpectedResult_(false); // false = no alert
    results.push('✅  API Expected    : ' + apiResult);
  } catch(e) {
    results.push('❌  API Expected    : ' + e.message);
    Logger.log('API Expected ERROR: ' + e.message);
  }

  // Step 3: Appendix recreation (broadcast to all modules)
  try {
    var apxResult = broadcastRecreateAppendix_(false); // false = no alert
    results.push('✅  Appendix        : ' + apxResult);
  } catch(e) {
    results.push('❌  Appendix        : ' + e.message);
    Logger.log('Appendix ERROR: ' + e.message);
  }

  safeAlert_('\n📋  Broadcast All Fixes — Selesai\n\n' + results.join('\n'));
}


// ══════════════════════════════════════════════════════════════════════════
// FIX 1: DASHBOARD — TAMBAH KOLOM QA LEAD
// ══════════════════════════════════════════════════════════════════════════
function fixDashboardLeadColumn() {
  var result = fixDashboardLeadColumn_();
  safeAlert_('✅  ' + result);
}

function fixDashboardLeadColumn_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  // ── Fix Config tab ─────────────────────────────────────────────────────
  var cfg = ss.getSheetByName('Config');
  if (!cfg) throw new Error('Tab Config tidak ditemukan. Pastikan dijalankan dari QA Dashboard.');

  var cfgHeaders = cfg.getRange(3, 1, 1, cfg.getLastColumn()).getValues()[0];
  var hasLead    = cfgHeaders.some(function(h){ return String(h).trim() === 'QA Lead'; });

  if (!hasLead) {
    // Insert col after col 5 (PIC/Team → col 5, Lead → col 6, Spreadsheet ID → col 7)
    cfg.insertColumnAfter(5);

    // Style header
    cfg.getRange(3, 6).setValue('QA Lead')
      .setBackground('#0D47A1').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    cfg.getRange(3, 6).setNote(
      'QA Lead\nNama QA Lead / person in charge untuk modul ini.\n' +
      'Diisi otomatis dari Summary tab (B4) saat refresh.\nBisa diisi manual sebagai default.'
    );
    cfg.setColumnWidth(6, 130);

    // Re-merge title & instruction rows to cover new column
    try {
      var lastCfgCol = cfg.getLastColumn();
      cfg.getRange(1, 1, 1, lastCfgCol).merge();
      cfg.getRange(2, 1, 1, lastCfgCol).merge();
    } catch(e) { Logger.log('Config merge warning: ' + e.message); }

    log.push('Config: kolom QA Lead ditambahkan (col F)');
  } else {
    log.push('Config: QA Lead sudah ada — skip');
  }

  // ── Fix Overview tab ────────────────────────────────────────────────────
  var ov = ss.getSheetByName('Overview');
  if (ov) {
    var ovHeaders = ov.getRange(4, 1, 1, ov.getLastColumn()).getValues()[0];
    var hasLeadOv = false;
    var leadOvIdx = -1;
    ovHeaders.forEach(function(h, i) {
      if (String(h).trim() === 'Lead') { hasLeadOv = true; leadOvIdx = i + 1; }
    });

    if (!hasLeadOv) {
      // Insert col 5 in Overview (after PIC/Team at col 4)
      ov.insertColumnAfter(4);

      // Style header at row 4 col 5
      ov.getRange(4, 5).setValue('Lead')
        .setBackground('#0D47A1').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');

      // Fix group header — MODULE INFO span: was 4, now 5
      try {
        ov.getRange(3, 1, 1, 4).breakApart();
        ov.getRange(3, 1, 1, 5).merge().setValue('MODULE INFO')
          .setBackground('#263238').setFontColor('#FFFFFF')
          .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
      } catch(e) { Logger.log('Overview group merge warning: ' + e.message); }

      // Fix title row merge
      try {
        var lastOvCol = ov.getLastColumn();
        ov.getRange(2, 1, 1, lastOvCol - 1).breakApart();
        ov.getRange(2, 1, 1, lastOvCol).merge();
        ov.getRange(1, 1, 1, lastOvCol - 1).breakApart();
        ov.getRange(1, 1, 1, lastOvCol).merge();
      } catch(e) { Logger.log('Overview title merge warning: ' + e.message); }

      ov.setColumnWidth(5, 110);
      log.push('Overview: kolom Lead ditambahkan (col E)');
    } else {
      log.push('Overview: Lead sudah ada di col ' + leadOvIdx + ' — skip');
    }
  } else {
    log.push('Overview: tab tidak ditemukan — skip');
  }

  var msg = log.join('\n');
  Logger.log('fixDashboardLeadColumn: ' + msg);
  return msg;
}


// ══════════════════════════════════════════════════════════════════════════
// FIX 2: API_MASTER — TAMBAH KOLOM EXPECTED RESULT
// ══════════════════════════════════════════════════════════════════════════
function broadcastApiExpectedResult() {
  var result = broadcastApiExpectedResult_(true);
  // alert is shown inside if showAlert=true
}

function broadcastApiExpectedResult_(showAlert) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');
  if (!cfg) throw new Error('Config tab tidak ditemukan');

  var allData  = cfg.getDataRange().getValues();
  var hdrRow   = allData[2] || [];
  var idColIdx = detectIdCol_(hdrRow);

  var ok = 0, skip = 0, err = 0, errList = [];

  for (var i = 3; i < allData.length; i++) {
    var row    = allData[i];
    if (String(row[0]).trim().toUpperCase() !== 'Y') continue;
    var id = String(row[idColIdx]).trim();
    if (!id || id.length < 10 || id === 'PASTE_SPREADSHEET_ID_HERE') continue;

    try {
      var res = fixApiExpectedResult_(id);
      if (res === 'skipped') skip++;
      else ok++;
    } catch(e) {
      err++;
      errList.push('  • ' + id.substring(0, 25) + ': ' + e.message);
      Logger.log('API Expected ERR [' + id.substring(0,20) + ']: ' + e.message);
    }
  }

  var summary = 'API Expected Result\nBerhasil: ' + ok + '  |  Skip: ' + skip + '  |  Gagal: ' + err;
  if (errList.length) summary += '\n' + errList.join('\n');
  Logger.log(summary);
  if (showAlert) safeAlert_('✅  ' + summary);
  return summary;
}

function fixApiExpectedResult_(spreadsheetId) {
  var src = SpreadsheetApp.openById(spreadsheetId);
  var ws  = src.getSheetByName('API_Master');
  if (!ws) throw new Error('Tab API_Master tidak ditemukan');

  var hdr      = ws.getRange(2, 1, 1, ws.getLastColumn()).getValues()[0];
  var col13    = String(hdr[12] || '').trim(); // Scenario
  var col14    = String(hdr[13] || '').trim(); // should become Expected Result

  if (col14.toLowerCase().indexOf('expected') > -1) return 'skipped';
  if (col13.toLowerCase().indexOf('scenario') === -1)
    throw new Error('Kolom 13 bukan Scenario ("' + col13 + '") — periksa manual');

  // Insert col 14 (after Scenario)
  ws.insertColumnAfter(13);

  ws.getRange(2, 14)
    .setValue('Expected Result')
    .setBackground('#283593').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  ws.getRange(2, 14).setNote(
    '[INPUT WAJIB] Expected Result dalam format Gherkin:\n' +
    '  Then : Hasil / response setelah request dikirim\n\n' +
    'Sertakan:\n' +
    '  - HTTP status code\n' +
    '  - Struktur response body (key penting)\n\n' +
    'Contoh Positive:\n' +
    '  Then 201 Created\n' +
    '  body: { id, status: "active", created_at }\n\n' +
    'Contoh Negative:\n' +
    '  Then 400 Bad Request\n' +
    '  body: { error: "field X required" }\n\n' +
    'Contoh Auth:\n' +
    '  Then 401 Unauthorized\n' +
    '  body: { message: "Token invalid or expired" }'
  );
  ws.setColumnWidth(14, 220);

  var lastRow = ws.getLastRow();
  if (lastRow >= 3) {
    ws.getRange(3, 14, lastRow - 2, 1)
      .setFontFamily('Arial').setFontSize(9)
      .setVerticalAlignment('middle').setWrap(true);
  }

  // Update title merge to cover new col count
  try {
    var newCols = ws.getLastColumn();
    ws.getRange(1, 1, 1, newCols - 1).breakApart();
    ws.getRange(1, 1, 1, newCols).merge();
  } catch(e) { Logger.log('API_Master title merge warning: ' + e.message); }

  // Fix API_Execution TestLevel sync: was col O (15), now col P (16)
  var execWs = src.getSheetByName('API_Execution');
  if (execWs) {
    var DS = 9, MR = 1000;
    var syncCell = execWs.getRange(DS, 6);
    var formula  = syncCell.getFormula();
    if (formula.indexOf('API_Master!O') > -1) {
      syncCell.setFormula(
        formula.replace(/API_Master!O(\d+):O(\d+)/g, 'API_Master!P$1:P$2')
               .replace(/API_Master!O(\d+)/g, 'API_Master!P$1')
      );
    }
  }

  return 'ok';
}


// ══════════════════════════════════════════════════════════════════════════
// FIX 3: APPENDIX — RECREATE DENGAN GHERKIN EXAMPLES
// ══════════════════════════════════════════════════════════════════════════
function broadcastRecreateAppendix() {
  var result = broadcastRecreateAppendix_(true);
}

function broadcastRecreateAppendix_(showAlert) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');
  if (!cfg) throw new Error('Config tab tidak ditemukan');

  var allData  = cfg.getDataRange().getValues();
  var hdrRow   = allData[2] || [];
  var idColIdx = detectIdCol_(hdrRow);

  var ok = 0, skip = 0, err = 0, errList = [];

  for (var i = 3; i < allData.length; i++) {
    var row    = allData[i];
    var active = String(row[0]).trim().toUpperCase();
    var id     = String(row[idColIdx]).trim();

    if (!id || id.length < 10) id = String(row[idColIdx + 1] || '').trim();
    if (active !== 'Y' || !id || id === 'PASTE_SPREADSHEET_ID_HERE' || id.length < 10) {
      skip++; continue;
    }

    try {
      var modSS = SpreadsheetApp.openById(id);
      buildAppendix_(modSS);
      ok++;
      Logger.log('Appendix OK: ' + id.substring(0, 25));
    } catch(e) {
      err++;
      errList.push('  • ' + id.substring(0, 25) + ': ' + e.message);
      Logger.log('Appendix ERR: ' + id.substring(0, 25) + ' — ' + e.message);
    }
  }

  var summary = 'Appendix recreated\nBerhasil: ' + ok + '  |  Skip: ' + skip + '  |  Gagal: ' + err;
  if (errList.length) summary += '\n' + errList.join('\n');
  Logger.log(summary);
  if (showAlert) safeAlert_('✅  ' + summary);
  return summary;
}


// ══════════════════════════════════════════════════════════════════════════
// BUILD APPENDIX — full content (dipanggil oleh broadcastRecreateAppendix_)
// ══════════════════════════════════════════════════════════════════════════
function buildAppendix_(ss) {
  var existing = ss.getSheetByName('Appendix');
  if (existing) ss.deleteSheet(existing);
  var ws = ss.insertSheet('Appendix');
  ws.setTabColor('#37474F');

  var r = 1;
  ws.setColumnWidth(1, 8);
  ws.setColumnWidth(2, 180);
  ws.setColumnWidth(3, 430);
  ws.setColumnWidth(4, 90);

  // Title
  ws.setRowHeight(r, 36);
  ws.getRange(r, 1, 1, 4).merge()
    .setValue('  📋  QA Test Management  —  Panduan Penggunaan  (Template v38)')
    .setBackground('#0D47A1').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(13).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;

  ws.setRowHeight(r, 16);
  ws.getRange(r, 1, 1, 4).merge()
    .setValue('  Dokumen ini berisi panduan pengisian semua tab. Update otomatis dari Dashboard broadcast.')
    .setBackground('#E3F2FD').setFontColor('#1565C0')
    .setFontSize(9).setFontFamily('Arial').setFontStyle('italic')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  r++;
  r++;

  // ── SECTION 0: HIERARKI ────────────────────────────────────────────────
  apxHdr_(ws, r, '0. HIERARKI QA  —  PROJECT / MODULE / SUBMODULE', '#0D47A1'); r++;
  apxRow_(ws, r, 'Struktur 4 level',
    'Project  →  Module  →  SubModule  →  Feature\n\n' +
    'PROJECT    = Inisiatif / client / program kerja  (SIPGN, INAGOV, COTS)\n' +
    'MODULE     = Domain besar dalam project  (1 - Manajemen Gizi, 2 - Distribusi)\n' +
    '             Kosongkan jika project flat  (INAGOV, COTS)\n' +
    'SUBMODULE  = Unit terkecil — 1 spreadsheet QA = 1 SubModule\n' +
    '             Berlayer  : 1.1 Aplikasi Nutritionist, 1.2 Aplikasi Courier\n' +
    '             Flat      : Talenta, e-Office, SIMPEG\n' +
    'FEATURE    = Fitur / halaman spesifik  (Login Page, Dashboard, Profile)',
    null, 85); r++;
  apxRow_(ws, r, 'Aturan SubModule',
    '• Gunakan kode yang IDENTIK di TC_Master, API_Master, dan Summary\n' +
    '• Perbedaan spasi / kapitalisasi = dihitung sebagai modul berbeda di Dashboard\n' +
    '• Berlayer: gunakan format n.m (1.1, 1.2, 2.1)\n' +
    '• Flat: gunakan nama langsung tanpa angka (Talenta bukan 1-Talenta)',
    '#E3F2FD', 70); r++;
  r++;

  // ── SECTION 1: STRUKTUR TAB ────────────────────────────────────────────
  apxHdr_(ws, r, '1. STRUKTUR TAB', '#0D47A1'); r++;
  var tabs = [
    ['Summary',       'KPI dan ringkasan status testing. Auto-hitung dari TC_Master + Execution.\nAda di SETIAP spreadsheet modul. Sumber data untuk Dashboard.'],
    ['TC_Master',     'Input test case Web / Mobile. Tidak ada kolom eksekusi di sini.\nKolom: No | SubModul | TC_ID | Feature | Priority | Platform | Test Type |\n       Automated | Version | Role (RBAC) | Scenario | Steps/Gherkin |\n       Expected Result | Examples | [AUTO] Test Level'],
    ['TC_Execution',  'Status eksekusi Web / Mobile. Kolom A-G auto-sync dari TC_Master.\nKolom H dst = status per run. Tambah kolom baru untuk setiap run baru.'],
    ['API_Master',    'Input test case API. Tidak ada kolom eksekusi.\nKolom: No | SubModul | TC_ID | Feature | Method | Endpoint URL | Priority |\n       Auth | Test Type | Automated | Version | Role (RBAC) | Scenario |\n       Expected Result | Examples | [AUTO] Test Level'],
    ['API_Execution', 'Status eksekusi API. Kolom A-F auto-sync dari API_Master.'],
    ['PerfTest',      'Hasil performance test. Input manual: tool, duration, VU, response time, status.'],
    ['BugReport',     'Daftar bug yang ditemukan. Status: Open / In Progress / Fixed / Verified.'],
    ['Appendix',      'Dokumen ini. Panduan pengisian semua tab. Di-update via Dashboard broadcast.'],
  ];
  tabs.forEach(function(t) { apxRow_(ws, r, t[0], t[1], null, 58); r++; });
  r++;

  // ── SECTION 2: STATUS EKSEKUSI ─────────────────────────────────────────
  apxHdr_(ws, r, '2. STATUS EKSEKUSI', '#0D47A1'); r++;
  var statuses = [
    ['PASS',    '#C8E6C9', '#1B5E20', 'Test case lulus — semua kondisi di Expected Result terpenuhi'],
    ['FAIL',    '#FFCDD2', '#C62828', 'Test case gagal — ada kondisi Expected Result yang tidak terpenuhi'],
    ['BLOCKED', '#FFE0B2', '#E65100', 'Tidak bisa dieksekusi karena dependency belum siap'],
    ['IN PROG', '#E3F2FD', '#1565C0', 'Sedang dalam proses eksekusi'],
    ['TODO',    '#F5F5F5', '#757575', 'Belum dieksekusi'],
    ['N/A',     '#ECEFF1', '#546E7A', 'Tidak applicable untuk kondisi atau environment saat ini'],
  ];
  statuses.forEach(function(s) { apxStatusRow_(ws, r, s[0], s[3], s[1], 28); r++; });
  r++;

  // ── SECTION 3: SCENARIO NAMING & GHERKIN ──────────────────────────────
  apxHdr_(ws, r, '3. SCENARIO NAMING STANDARD  &  GHERKIN FORMAT', '#0D47A1'); r++;
  apxRow_(ws, r, 'Formula Penamaan',
    'Happy Path  :  [Role] + Successfully + [Verb] + [Object] + (from/to [Location])\n' +
    'Negative    :  [Role] + Failed to + [Verb] + [Object] + with [Condition]\n\n' +
    'Contoh OK :\n' +
    '  ✓  Nutritionist Successfully Creates Meal Plan\n' +
    '  ✓  Courier Successfully Picks Up Food from SPPG\n' +
    '  ✓  Admin Failed to Delete User with Invalid ID -- 404\n\n' +
    'Jangan gunakan:\n' +
    '  ✗  success / succeed   →  pakai Successfully\n' +
    '  ✗  do, perform, process  →  langsung verb utama',
    null, 90); r++;
  apxRow_(ws, r, 'Gherkin — Steps (Given + When)',
    'Tulis di kolom "Steps / Gherkin":\n\n' +
    '  Given  : Pre-kondisi / state awal\n' +
    '           Contoh: Given user sudah login sebagai Nutritionist\n' +
    '  When   : Aksi yang dilakukan aktor\n' +
    '           Contoh: When user klik tombol Pick Up dan konfirmasi\n' +
    '  And    : Aksi tambahan jika diperlukan\n\n' +
    'JANGAN tulis Then di kolom Steps — Then ada di Expected Result.',
    null, 80); r++;
  apxRow_(ws, r, 'Gherkin — Expected Result (Then)',
    'Tulis di kolom "Expected Result":\n\n' +
    '  Then : Hasil / perubahan state setelah aksi selesai\n\n' +
    'Tips:\n' +
    '  • Spesifik — sebutkan elemen UI, pesan, atau status yang muncul\n' +
    '  • Contoh Web: Then halaman dashboard tampil, nama user muncul di header\n' +
    '  • Contoh API: Then 201 Created, body: { id, status: "active", created_at }\n' +
    '  • Bukan hanya "berhasil" atau "success"',
    null, 75); r++;
  r++;

  // ── SECTION 3b: GHERKIN EXAMPLES ──────────────────────────────────────
  apxHdr_(ws, r, '3b. GHERKIN EXAMPLES  —  SCENARIO OUTLINE', '#004D40'); r++;
  apxRow_(ws, r, 'Kapan pakai Examples?',
    'Gunakan Scenario Outline + Examples ketika skenario yang sama perlu dijalankan\n' +
    'dengan data berbeda — menghindari duplikasi TC.\n\n' +
    'Cocok untuk:\n' +
    '  • Form validation dengan variasi input\n' +
    '  • Kalkulasi dengan variasi angka / data\n' +
    '  • Multi-role access check (siapa boleh, siapa ditolak)',
    null, 65); r++;
  apxRow_(ws, r, 'Format Dasar',
    'Scenario Outline: [Role] Successfully/Failed to [Verb] [Object]\n' +
    '  Given  <pre_kondisi>\n' +
    '  When   <aksi> dengan "<input>"\n' +
    '  Then   <expected>\n\n' +
    'Examples:\n' +
    '  | input        | expected         |\n' +
    '  | data valid   | 201 Created      |\n' +
    '  | data kosong  | 400 Bad Request  |\n' +
    '  | duplikat     | 409 Conflict     |',
    null, 95); r++;
  apxRow_(ws, r, 'Contoh: Validasi Form Login',
    'Scenario Outline: User Attempts to Log In with <email> and <password>\n' +
    '  Given  user berada di halaman Login\n' +
    '  When   user mengisi email "<email>" dan password "<password>" lalu klik Login\n' +
    '  Then   <expected_result>\n\n' +
    'Examples:\n' +
    '  | email           | password   | expected_result                          |\n' +
    '  | user@sipgn.id   | Pass@123   | redirect ke Dashboard                    |\n' +
    '  | user@sipgn.id   | wrongpass  | muncul pesan "Password salah"            |\n' +
    '  | invalid-email   | Pass@123   | muncul pesan "Format email tidak valid"  |\n' +
    '  | (kosong)        | (kosong)   | tombol Login disabled                    |',
    '#E0F2F1', 105); r++;
  apxRow_(ws, r, 'Contoh: API Multi-role Access',
    'Scenario Outline: <role> Tries to Access Meal Plan Data -- <status>\n' +
    '  Given  user memiliki token Bearer sebagai <role>\n' +
    '  When   GET /api/v1/meal-plan\n' +
    '  Then   response status adalah <status>\n\n' +
    'Examples:\n' +
    '  | role          | status |\n' +
    '  | Nutritionist  | 200    |\n' +
    '  | Admin         | 200    |\n' +
    '  | Courier       | 403    |\n' +
    '  | (no token)    | 401    |',
    '#E0F2F1', 90); r++;
  apxRow_(ws, r, 'Aturan Penggunaan',
    '1. Nama kolom tabel pakai huruf kecil + underscore:  <input_value>  bukan  <Input Value>\n' +
    '2. Setiap baris di tabel = 1 test case yang dieksekusi tersendiri\n' +
    '3. Maksimal 5-6 baris per tabel — jika lebih, split jadi beberapa Scenario Outline\n' +
    '4. TC_ID: tambah suffix huruf per baris\n' +
    '   Contoh: PO.001 (outline)  →  PO.001a, PO.001b, PO.001c per baris\n' +
    '5. Kolom Expected Result di TC_Master: tulis pola umum\n' +
    '   Detail per baris cukup ada di tabel Examples\n' +
    '6. Tulis tabel Examples di kolom "Examples" yang tersedia di TC_Master / API_Master',
    null, 90); r++;
  r++;

  // ── SECTION 4: TEST LEVEL ─────────────────────────────────────────────
  apxHdr_(ws, r, '4. TEST LEVEL  —  OTOMATIS DARI PRIORITY', '#546E7A'); r++;
  apxRow_(ws, r, 'Smoke Test',
    'TC dengan Priority: Critical / High / Medium\n' +
    'Wajib dieksekusi di setiap run. Gagal = blocker release.',
    '#FFF8F0', 38); r++;
  apxRow_(ws, r, 'Regression Test',
    'TC dengan Priority: Low / Lowest\n' +
    'Dieksekusi di full regression cycle. Gagal = catatan, bukan blocker.',
    '#F1F8E9', 38); r++;
  r++;

  // ── SECTION 5: PRIORITY ────────────────────────────────────────────────
  apxHdr_(ws, r, '5. PRIORITY', '#0D47A1'); r++;
  var priorities = [
    ['Critical', '#FFCDD2', '#C62828', 'Blocker utama. WAJIB PASS sebelum release. FAIL/BLOCKED = release DITAHAN.'],
    ['High',     '#FFE0B2', '#D84315', 'Harus PASS di sprint yang sama. FAIL = perlu approval PM untuk release.'],
    ['Medium',   '#FFF9C4', '#E65100', 'Potential blocker. Fix sebelum UAT. FAIL = flagged ke tech lead.'],
    ['Low',      '#F1F8E9', '#2E7D32', 'Non-blocker. Fix di sprint berikutnya.'],
    ['Lowest',   '#ECEFF1', '#546E7A', 'Nice to have. Opsional.'],
  ];
  priorities.forEach(function(p) { apxStatusRow_(ws, r, p[0], p[3], p[1], 28); r++; });
  r++;

  // ── SECTION 6: ROLE RBAC ──────────────────────────────────────────────
  apxHdr_(ws, r, '6. ROLE (RBAC)  —  KONTROL AKSES', '#0D47A1'); r++;
  apxRow_(ws, r, 'Isi kolom Role dengan',
    'Peran pengguna yang menjalankan skenario — bukan role developer atau tester.\n\n' +
    'Contoh: Admin, Super Admin, User, Viewer, Operator, Supervisor, Guest\n\n' +
    'Gunakan untuk:\n' +
    '  • Memastikan coverage test per role\n' +
    '  • Verifikasi RBAC berjalan benar\n' +
    '  • Pastikan ada TC 401 (tanpa auth) dan 403 (role salah) untuk endpoint sensitif',
    null, 65); r++;
  apxRow_(ws, r, 'Aturan konsistensi',
    '• Nama role harus konsisten di seluruh TC — "Super Admin" ≠ "Superadmin"\n' +
    '• Gunakan Title Case\n' +
    '• Definisikan daftar role di awal project dan patuhi sampai akhir',
    '#E3F2FD', 50); r++;
  r++;

  // ── SECTION 7: AUTOMATION STATUS ──────────────────────────────────────
  apxHdr_(ws, r, '7. AUTOMATION STATUS', '#0D47A1'); r++;
  var autoStatuses = [
    ['Automated',            '#C8E6C9', '#1B5E20', 'Script automation sudah ada dan berjalan'],
    ['To Do',                '#FFF9C4', '#F57F17', 'Direncanakan untuk diautomasi, belum dikerjakan'],
    ['Manual',               '#E3F2FD', '#1565C0', 'Diputuskan tetap manual — tidak akan diautomasi'],
    ['Cannot be Automated',  '#F5F5F5', '#9E9E9E', 'Secara teknis tidak memungkinkan untuk diautomasi'],
  ];
  autoStatuses.forEach(function(a) { apxStatusRow_(ws, r, a[0], a[3], a[1], 28); r++; });
  r++;

  // ── SECTION 8: BUG REPORT ─────────────────────────────────────────────
  apxHdr_(ws, r, '8. BUG REPORT  —  PANDUAN PENGISIAN', '#B71C1C'); r++;
  apxRow_(ws, r, 'Kapan isi BugReport?',
    'Isi BugReport setiap kali menemukan defect saat eksekusi TC.\n' +
    'Bisa dari TC_Execution maupun dari exploratory testing.',
    null, 38); r++;
  apxRow_(ws, r, 'Severity',
    'Blocker  → Aplikasi tidak bisa digunakan sama sekali\n' +
    'Critical → Fitur utama rusak, tidak ada workaround\n' +
    'High     → Fitur penting terganggu, ada workaround terbatas\n' +
    'Medium   → Fitur sekunder terganggu\n' +
    'Low      → Tampilan / UX issue, tidak mempengaruhi fungsi',
    null, 65); r++;
  r++;

  // ── SECTION 9: PERFORMANCE TEST ───────────────────────────────────────
  apxHdr_(ws, r, '9. PERFORMANCE TEST  —  METRIK', '#4A148C'); r++;
  apxRow_(ws, r, 'Kolom PerfTest',
    'Tool         : k6, JMeter, Locust, Artillery\n' +
    'Duration     : Durasi test dalam menit\n' +
    'VU           : Virtual Users (jumlah user simultan)\n' +
    'Avg Response : Rata-rata response time dalam ms\n' +
    'P95          : 95th percentile response time\n' +
    'Error Rate   : Persentase request yang gagal\n' +
    'Result       : PASS / FAIL berdasarkan threshold yang disepakati',
    null, 75); r++;
  r++;

  // ── SECTION 10: HTTP METHOD ───────────────────────────────────────────
  apxHdr_(ws, r, '10. HTTP METHOD', '#0D47A1'); r++;
  var methods = [
    ['GET',    '#E8F0FE', '#1A237E', 'Ambil / baca data. Tidak mengubah state.'],
    ['POST',   '#E8F5E9', '#1B5E20', 'Buat resource baru. Response: 201 Created.'],
    ['PUT',    '#FFF8E1', '#E65100', 'Ganti seluruh resource. Idempotent.'],
    ['PATCH',  '#F3E5F5', '#4A148C', 'Update sebagian field resource. Idempotent.'],
    ['DELETE', '#FCE4EC', '#880E4F', 'Hapus resource. Response: 200 OK atau 204 No Content.'],
  ];
  methods.forEach(function(m) { apxStatusRow_(ws, r, m[0], m[3], m[1], 28); r++; });
  r++;

  // ── Footer ─────────────────────────────────────────────────────────────
  ws.setRowHeight(r, 20);
  ws.getRange(r, 1, 1, 4).merge()
    .setValue('  QA Team  ·  Template v38  ·  2026  ·  Auto-updated via Dashboard Broadcast')
    .setBackground('#0D47A1').setFontColor('#BBDEFB')
    .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}


// ══════════════════════════════════════════════════════════════════════════
// APPENDIX STYLE HELPERS
// ══════════════════════════════════════════════════════════════════════════
function apxHdr_(ws, r, title, bg) {
  ws.setRowHeight(r, 22);
  ws.getRange(r, 1, 1, 4).merge()
    .setValue('  ' + title)
    .setBackground(bg || '#37474F').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
}

function apxRow_(ws, r, label, desc, labelBg, rowH) {
  ws.setRowHeight(r, rowH || 48);
  ws.getRange(r, 1).setValue('')
    .setBackground(labelBg || '#FAFAFA');
  ws.getRange(r, 2).setValue(label)
    .setBackground(labelBg || '#FAFAFA')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
  ws.getRange(r, 3).setValue(desc)
    .setBackground('#FFFFFF')
    .setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
  ws.getRange(r, 4).setValue('');
}

function apxStatusRow_(ws, r, label, desc, bg, rowH) {
  ws.setRowHeight(r, rowH || 28);
  ws.getRange(r, 1).setValue('').setBackground(bg || '#EEEEEE');
  ws.getRange(r, 2).setValue(label)
    .setBackground(bg || '#EEEEEE')
    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange(r, 3).setValue(desc)
    .setBackground('#FFFFFF')
    .setFontSize(9).setFontFamily('Arial')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
  ws.getRange(r, 4).setValue('');
}


// ══════════════════════════════════════════════════════════════════════════
// SHARED HELPERS
// ══════════════════════════════════════════════════════════════════════════
function detectIdCol_(headerRow) {
  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i]).trim().toUpperCase();
    if (h === 'SPREADSHEET ID' || h === 'SPREADSHEET_ID') return i;
  }
  return 6; // fallback col G (after Active, Project, Module, SubModule, PIC, Lead)
}

function safeAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch(e) { try { Browser.msgBox(msg); } catch(e2) { Logger.log(msg); } }
}