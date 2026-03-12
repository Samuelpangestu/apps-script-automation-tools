/**
 * addDashboardNotes.js - Tambahkan notes/descriptions ke Dashboard headers
 * Run: addNotesToDashboard()
 *
 * Objective: Menambahkan tooltip notes di Dashboard untuk membantu stakeholder
 * memahami KPI dan metrics yang ditampilkan
 *
 * Tabs covered:
 * - Overview: Portfolio overview dengan semua modules
 * - Smoke: Smoke test dedicated view
 * - Blockers: Bug blocker tracking
 * - Coverage: Test coverage analysis
 * - Config: Module configuration
 */

/**
 * Main function: Tambahkan notes ke SEMUA dashboard tabs
 */
function addNotesToDashboard() {
  Logger.log('🔖 Adding notes to ALL Dashboard tabs...');

  addNotesToOverview();
  addNotesToSmoke();
  addNotesToBlockers();
  addNotesToCoverage();
  addNotesToConfig();

  SpreadsheetApp.getUi().alert(
    '✅ Dashboard Notes Added',
    'Notes telah ditambahkan ke SEMUA tabs:\n\n' +
    '✅ Overview\n' +
    '✅ Smoke\n' +
    '✅ Blockers\n' +
    '✅ Coverage\n' +
    '✅ Config\n\n' +
    'Hover mouse di header untuk melihat description lengkap.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log('✅ All dashboard notes added successfully');
}

/**
 * Add notes to Overview tab
 */
function addNotesToOverview() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Overview');

  if (!ws) {
    Logger.log('⚠️ Overview tab not found, skipping notes');
    return;
  }

  Logger.log('📝 Adding notes to Overview tab...');

  // Find header row (typically row 3, after title and column group headers)
  const headerRow = 3;

  // Overview section headers
  const overviewHeaders = [
    ['#', 'Nomor urut modul dalam dashboard\nUrutan sesuai Config sheet'],
    ['Project', 'Nama project yang sedang ditest\nContoh: SINET Mobile, Customer Portal Web, Payment Gateway\nSumber: Summary tab > Project field dari masing-masing QATM'],
    ['Modul', 'Kode/nama modul utama dalam project\nContoh: 1 (Login & Auth), 2 (Transaction), 3.1 (Reporting)\nDigunakan untuk grouping dan filtering test cases\nSumber: Config sheet kolom Modul'],
    ['Submodul', 'Sub-modul detail dalam modul utama\nBisa lebih dari 1, dipisahkan koma\nContoh: 1.1 (Login Page), 1.2 (Register), 1.3 (Forgot Password)\nMemberikan breakdown lebih detail untuk tracking\nSumber: Summary tab > Submodul field'],
    ['PIC QA', 'Person In Charge QA\nNama QA engineer yang bertanggung jawab melakukan testing modul ini\nBertanggung jawab: eksekusi test, report bug, update status\nSumber: Summary tab > PIC QA field'],
    ['QA Lead', 'QA Team Lead\nNama QA Lead yang supervise dan review testing modul ini\nBertanggung jawab: approve test plan, review hasil, quality assurance\nSumber: Summary tab > QA Lead field (jika ada)'],
    ['Last Updated', 'Timestamp terakhir kali data modul ini di-refresh dari source QATM sheet\nFormat: YYYY-MM-DD HH:MM:SS\nAuto-update saat refresh dashboard\nJika timestamp lama (>24 jam), kemungkinan ada issue sync']
  ];

  // TEST STATUS section
  const testStatusHeaders = [
    ['Status', 'Status keseluruhan testing untuk modul ini\n\nPilihan status:\n• Not Started = belum mulai testing\n• In Progress = testing sedang berjalan\n• Completed = testing selesai untuk sprint/cycle ini\n• On Hold = testing ditunda sementara\n\nSumber: Summary tab > Test Status field'],
    ['Total TC', 'Total Test Cases = jumlah total test case untuk modul ini\nGabungan dari:\n• Web/Mobile test cases (TC_Master)\n• API test cases (API_Master)\n\nAngka ini adalah baseline untuk menghitung Pass Rate dan Exec Rate\nSumber: Summary tab > TOTAL counter (Web/Mobile + API)'],
    ['Passed', 'Jumlah test case dengan hasil PASSED\n= Test case yang berhasil/sukses dijalankan\n= Sistem berfungsi sesuai expected result\n\nGabungan Web/Mobile PASSED + API PASSED\nSemakin tinggi angka ini semakin baik\nSumber: Summary tab > PASSED counter'],
    ['Failed', 'Jumlah test case dengan hasil FAILED\n= Test case yang gagal\n= Ditemukan bug atau sistem tidak sesuai expected\n\nPerlu action: investigate dan fix bug yang ditemukan\nFailed tinggi = quality issue, perlu perhatian tim development\nSumber: Summary tab > FAILED counter'],
    ['Pass Rate', 'Pass Rate = Persentase keberhasilan testing\nFormula: (PASSED / TOTAL) × 100%\n\n📊 THRESHOLD:\n🟢 Hijau ≥80% = BAIK, quality bagus, ready untuk release\n🟡 Kuning 50-79% = PERLU PERHATIAN, masih ada banyak issue\n🔴 Merah <50% = CRITICAL, quality buruk, NOT READY, perlu rework major\n\nMetric utama untuk menilai quality dan readiness produk\nTarget minimum untuk release: ≥80%'],
    ['Exec Rate', 'Execution Rate = Persentase progress eksekusi testing\nFormula: ((PASSED + FAILED + BLOCKED + IN PROGRESS) / TOTAL) × 100%\n= % test case yang sudah dapat hasil (bukan TODO)\n\n📊 THRESHOLD:\n🟢 Hijau ≥70% = progress baik\n🟡 Kuning 40-69% = progress medium\n🔴 Merah <40% = progress lambat, perlu percepatan\n\nBeda dengan Pass Rate:\n• Pass Rate = quality metric (berapa % sukses)\n• Exec Rate = progress metric (berapa % sudah dikerjakan)\n\nExec Rate tinggi tapi Pass Rate rendah = testing sudah banyak, tapi quality jelek']
  ];

  // BUGS section
  const bugsHeaders = [
    ['Total', 'Total Bugs = jumlah total bugs yang dilaporkan untuk modul ini\nDihitung dari BugReport sheet\nMeliputi semua status: Open, In Progress, Fixed, Verified, Reopen\n(tidak termasuk Closed)\n\nSemakin rendah semakin baik\nTotal bugs tinggi = indikator quality issue atau complexity tinggi'],
    ['Open', 'Bugs dengan status OPEN\n= Bug baru yang dilaporkan QA\n= Belum dikerjakan developer\n= Belum di-assign atau belum started\n\nPerlu action: Developer perlu ambil dan mulai fix\nOpen lama (>3 hari) tanpa action = bottleneck\nSumber: BugReport sheet, Status = "Open"'],
    ['In Prog', 'Bugs dengan status IN PROGRESS\n= Bug sedang dikerjakan developer\n= Developer sudah assigned dan actively fixing\n\nStatus positif = bug sedang ditangani\nIn Progress lama (>5 hari) = kemungkinan stuck, perlu follow up\nSumber: BugReport sheet, Status = "In Progress"'],
    ['Fixed', 'Bugs dengan status FIXED\n= Bug sudah diperbaiki oleh developer\n= Menunggu verifikasi dari QA\n= Belum confirmed fix berhasil\n\nPerlu action: QA perlu test ulang dan verifikasi fix\nFixed menumpuk = QA bottleneck, perlu verifikasi lebih cepat\nSumber: BugReport sheet, Status = "Fixed"'],
    ['Verified', 'Bugs dengan status VERIFIED\n= Bug sudah diverifikasi QA, fix berhasil\n= Bug confirmed solved\n= Menunggu deployment atau closure\n\nStatus positif = bug resolution sukses\nVerified tinggi = tim working well, bug cycle cepat\nSumber: BugReport sheet, Status = "Verified"'],
    ['Blocker', 'OPEN BLOCKER = Critical/High/Medium Priority Bugs yang masih terbuka\n\n📊 FORMULA:\nBugs dengan kondisi:\n• Priority: Critical, High, atau Medium\n• Status: Open, In Progress, Reopen, Fixed, atau Verified\n• Environment: SEMUA (Dev, Staging, Production, UAT, dll)\n\n⚠️ CRITICAL METRIC:\nBlocker > 0 = Ada bugs kritis yang MENGHALANGI testing!\n\n🚨 IMPACT:\n• Testing tidak bisa lanjut\n• Features tidak bisa di-test\n• Schedule testing terhambat\n• Risk delivery delay\n\n✅ ACTION REQUIRED:\nBlocker bugs harus jadi TOP PRIORITY\nDaily standup: report blocker dan track resolution\nEscalate ke management jika stuck >2 hari\n\n💡 OVERLAP dengan PROD:\nBlocker BISA include bugs dari Production environment\nContoh: Bug Critical di Production → masuk Blocker DAN PROD\n• Blocker = focus on Priority (Critical/High/Medium)\n• PROD = focus on Environment (Production only)\n\n🎯 Target: Blocker = 0'],
    ['Reopen', 'Bugs dengan status REOPEN\n= Bug yang sudah di-fix tapi muncul lagi\n= Fix sebelumnya tidak berhasil atau incomplete\n= Regression - bug kembali muncul\n\n⚠️ QUALITY INDICATOR:\nReopen tinggi = indikator:\n• Fix quality buruk\n• Developer tidak test dengan baik\n• Root cause tidak teridentifikasi\n• Regression testing tidak adequate\n\n📋 ACTION:\n• Perlu root cause analysis lebih dalam\n• Review testing approach developer\n• Strengthen regression test suite\n• Code review lebih ketat\n\nTarget: Reopen minimal, idealnya 0\nSumber: BugReport sheet, Status = "Reopen"'],
    ['PROD', 'PRODUCTION BUGS = Bugs di Production Environment yang masih aktif\n\n📊 FORMULA:\nBugs dengan kondisi:\n• Environment = "Production" (exact match)\n• Status: Open, In Progress, Reopen, Fixed, atau Verified\n  (belum Closed atau Won\'t Fix)\n• Priority: SEMUA (Critical, High, Medium, Low, Lowest)\n\n🚨🚨 HIGHEST PRIORITY - URGENT! 🚨🚨\n\nIMPACT:\n• Real users di production terimpact\n• Business operation terganggu\n• Revenue/reputation risk\n• SLA breach risk\n• Customer complaints\n\nSEVERITY:\n• PROD bugs > 0 = CRITICAL situation\n• PROD bugs Critical priority = EMERGENCY\n• Membutuhkan immediate action dan war room jika perlu\n\n✅ PROTOCOL:\n1. Immediate notification ke all stakeholders\n2. Assign senior developer immediately\n3. Daily updates mandatory\n4. Hotfix deployment prioritized\n5. Post-mortem setelah resolved\n\n📧 NOTIFICATION:\nPROD bugs otomatis muncul di:\n• Email notification (red alert box)\n• Google Chat notification (red banner)\n• Dashboard (red highlight)\n\n💡 OVERLAP dengan BLOCKER:\nPROD bugs dengan Priority Critical/High/Medium juga masuk Blocker\nContoh: Bug Medium di Production → masuk PROD DAN Blocker\n• PROD = focus on Environment (Production only)\n• Blocker = focus on Priority (Critical/High/Medium, all env)\n\n🎯 TARGET: PROD bugs = 0\nJika PROD bugs > 0 → all hands on deck!']
  ];

  // Set notes based on column positions
  // These positions are based on typical Dashboard layout
  // Adjust if your layout is different

  let col = 1;

  // Overview section
  overviewHeaders.forEach(([label, note]) => {
    setNote(ws, headerRow, col, note);
    col++;
  });

  // Test Status section (typically starts after col 7)
  testStatusHeaders.forEach(([label, note]) => {
    setNote(ws, headerRow, col, note);
    col++;
  });

  // Bugs section (continues after Test Status)
  bugsHeaders.forEach(([label, note]) => {
    setNote(ws, headerRow, col, note);
    col++;
  });

  // Add notes to section headers (row 2)
  setNote(ws, 2, 1, 'Informasi dasar modul: Project, Modul, Submodul, PIC, dan Last Update');
  setNote(ws, 2, 8, 'Metrics testing: Status, jumlah test case, Pass Rate, Execution Rate\n\nThresholds:\n• Pass Rate: Hijau ≥80%, Kuning 50-79%, Merah <50%\n• Exec Rate: Hijau ≥70%, Kuning 40-69%, Merah <40%');
  setNote(ws, 2, 14, 'Bug metrics: Total bugs, status breakdown, dan critical metrics\n\n⚠️ Perhatian khusus:\n• Blocker = bugs priority tinggi yang masih open\n• PROD = bugs di production environment (urgent!)');

  // Add note to Dashboard title
  setNote(ws, 1, 1, 'QA Portfolio Dashboard\n\nDashboard agregasi real-time dari semua modul testing aktif\nData di-refresh otomatis setiap 10 menit atau manual via menu Dashboard > Refresh All Data\n\nColor coding:\n• Hijau = metrics baik/target tercapat\n• Kuning = perlu perhatian\n• Merah = critical, butuh action segera');

  Logger.log('✅ Overview notes added');
}

// ═══════════════════════════════════════════════════════════════════════
// SMOKE TAB NOTES
// ═══════════════════════════════════════════════════════════════════════

/**
 * Add notes to Smoke tab
 */
function addNotesToSmoke() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Smoke');

  if (!ws) {
    Logger.log('⚠️ Smoke tab not found, skipping notes');
    return;
  }

  Logger.log('📝 Adding notes to Smoke tab...');

  // Smoke tab title
  setNote(ws, 1, 1, '🔥 SMOKE TEST DASHBOARD\n\nDedicated view untuk Smoke Test results\n\nSmoke Test = Subset critical test cases (Priority: Critical, High, Medium)\nTujuan: Quick validation bahwa core functionality berfungsi\nIdeal untuk: Pre-release validation, daily builds, CI/CD gates\n\nTarget: Smoke Pass Rate ≥80%');

  // Section headers (row 2)
  setNote(ws, 2, 1, 'Module Information: Project, Modul, PIC QA');
  setNote(ws, 2, 4, 'WEB/MOBILE SMOKE TEST\n\nTest cases Web/Mobile dengan Priority: Critical, High, Medium\nCore functionality yang harus selalu berfungsi');
  setNote(ws, 2, 7, 'API SMOKE TEST\n\nTest cases API dengan Priority: Critical, High, Medium\nCore API endpoints yang critical untuk system');

  // Column headers (row 3)
  const smokeHeaders = [
    [1, 'Project', 'Nama project yang sedang ditest'],
    [2, 'Modul', 'Kode/nama modul dalam project'],
    [3, 'PIC QA', 'QA engineer yang handle testing module ini'],
    // Web Smoke
    [4, 'Total', 'Total smoke test cases Web/Mobile\n= TC dengan Priority: Critical, High, Medium\n\nSmoke suite adalah subset dari total TC'],
    [5, 'Pass%', 'Web Smoke Pass Rate\n= (PASSED / Total Smoke) × 100%\n\n🟢 ≥80% = PASS, core functionality OK\n🟡 50-79% = WARNING, ada issue di core features\n🔴 <50% = FAIL, core functionality broken\n\nTarget untuk release: ≥80%'],
    [6, 'Exec%', 'Web Smoke Execution Rate\n= % smoke test yang sudah dijalankan\n\n🟢 ≥70% = Progress baik\n🟡 40-69% = Progress sedang\n🔴 <40% = Progress lambat'],
    // API Smoke
    [7, 'Total', 'Total smoke test cases API\n= API TC dengan Priority: Critical, High, Medium\n\nCore API endpoints yang critical'],
    [8, 'Pass%', 'API Smoke Pass Rate\n= (PASSED / Total Smoke) × 100%\n\n🟢 ≥80% = PASS, core APIs OK\n🟡 50-79% = WARNING, ada issue di core APIs\n🔴 <50% = FAIL, core APIs broken\n\nTarget untuk release: ≥80%'],
    [9, 'Exec%', 'API Smoke Execution Rate\n= % smoke API test yang sudah dijalankan\n\n🟢 ≥70% = Progress baik\n🟡 40-69% = Progress sedang\n🔴 <40% = Progress lambat']
  ];

  smokeHeaders.forEach(([col, label, note]) => {
    setNote(ws, 3, col, note);
  });

  Logger.log('✅ Smoke notes added');
}

// ═══════════════════════════════════════════════════════════════════════
// BLOCKERS TAB NOTES
// ═══════════════════════════════════════════════════════════════════════

/**
 * Add notes to Blockers tab
 */
function addNotesToBlockers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Blockers');

  if (!ws) {
    Logger.log('⚠️ Blockers tab not found, skipping notes');
    return;
  }

  Logger.log('📝 Adding notes to Blockers tab...');

  // Blockers tab title
  setNote(ws, 1, 1, '🚨 BLOCKER & PROD BUGS DASHBOARD\n\nCritical bugs yang menghalangi testing dan bugs di Production\n\nBLOCKER = Priority tinggi (Critical/High/Medium) yang masih open\nPROD BUGS = Bugs di Production environment (URGENT!)\n\nDashboard ini untuk:\n• Daily standup discussion\n• Prioritization bugs yang perlu immediate action\n• Track resolution progress\n• Escalation ke management\n\n⚠️ Target: Blocker = 0, PROD = 0');

  // Section headers
  setNote(ws, 2, 1, 'Module Information');
  setNote(ws, 2, 4, 'OPEN BLOCKERS\n\nBugs dengan Priority Critical/High/Medium yang masih open\nStatus: Open, In Progress, Reopen, Fixed, Verified\nEnvironment: SEMUA\n\n⚠️ Menghalangi testing, perlu immediate action');
  setNote(ws, 2, 8, 'PRODUCTION BUGS\n\nBugs di Production environment (any priority)\nStatus: Belum Closed\n\n🚨🚨 HIGHEST PRIORITY - Real users terimpact!');

  // Column headers
  const blockerHeaders = [
    [1, 'Project', 'Nama project'],
    [2, 'Modul', 'Kode modul'],
    [3, 'PIC QA', 'QA engineer responsible'],
    [4, 'Total Blocker', 'Total bugs dengan Priority Critical/High/Medium yang masih open\n\nStatus yang dihitung:\n• Open\n• In Progress\n• Reopen\n• Fixed (belum verified)\n• Verified (belum closed)\n\nEnvironment: SEMUA (Dev, Staging, Prod, UAT)\n\n🎯 Target: 0 blockers'],
    [5, 'Critical', 'Bugs dengan Priority = Critical\n\n🔴 CRITICAL:\n• System down\n• Data loss\n• Security breach\n• Major functionality broken\n\nAction: Hotfix <24 hours'],
    [6, 'High', 'Bugs dengan Priority = High\n\n🟠 HIGH:\n• Major feature broken\n• High business impact\n• Significant user impact\n\nAction: Fix <48 hours'],
    [7, 'Medium', 'Bugs dengan Priority = Medium\n\n🟡 MEDIUM:\n• Minor feature issue\n• Workaround available\n• Moderate impact\n\nAction: Fix in current sprint'],
    [8, 'Total PROD', 'Total bugs di Production environment\n\nEnvironment = "Production"\nStatus: Belum Closed\nPriority: SEMUA\n\n🚨 URGENT! Real users terimpact\n\n🎯 Target: 0 PROD bugs'],
    [9, 'PROD Critical', 'Production bugs dengan Priority Critical\n\n🚨🚨 EMERGENCY!\n• War room required\n• All hands on deck\n• Immediate hotfix\n• Hourly updates mandatory'],
    [10, 'PROD High', 'Production bugs dengan Priority High\n\n🚨 VERY URGENT\n• Senior dev assigned\n• Daily updates required\n• Hotfix prioritized'],
    [11, 'PROD Medium', 'Production bugs dengan Priority Medium\n\nURGENT\n• Track daily\n• Include in next release']
  ];

  blockerHeaders.forEach(([col, label, note]) => {
    setNote(ws, 3, col, note);
  });

  Logger.log('✅ Blockers notes added');
}

// ═══════════════════════════════════════════════════════════════════════
// COVERAGE TAB NOTES
// ═══════════════════════════════════════════════════════════════════════

/**
 * Add notes to Coverage tab
 */
function addNotesToCoverage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Coverage');

  if (!ws) {
    Logger.log('⚠️ Coverage tab not found, skipping notes');
    return;
  }

  Logger.log('📝 Adding notes to Coverage tab...');

  // Coverage tab title
  setNote(ws, 1, 1, '📊 TEST COVERAGE DASHBOARD\n\nAnalysis coverage testing per module\n\nMetrics:\n• Pass Rate: Quality metric (% test yang pass)\n• Execution Rate: Progress metric (% test yang sudah dijalankan)\n• TC Breakdown: Status distribution\n\nGunakan untuk:\n• Identify modules dengan coverage rendah\n• Track testing progress\n• Quality assessment per module\n• Resource allocation\n\n🎯 Target: Pass Rate ≥80%, Exec Rate ≥70%');

  // Section headers
  setNote(ws, 2, 1, 'Module Information');
  setNote(ws, 2, 4, 'WEB/MOBILE TEST COVERAGE\n\nCoverage untuk Web dan Mobile testing\n\nMetrics untuk assess quality dan progress');
  setNote(ws, 2, 8, 'API TEST COVERAGE\n\nCoverage untuk API testing\n\nMetrics untuk assess API quality dan progress');

  // Column headers
  const coverageHeaders = [
    [1, 'Project', 'Nama project'],
    [2, 'Modul', 'Kode modul'],
    [3, 'PIC QA', 'QA engineer responsible'],
    // Web Coverage
    [4, 'Total TC', 'Total test cases Web/Mobile\n\nSemua TC di TC_Master sheet'],
    [5, 'Passed', 'Test cases yang PASSED\n\n= Berfungsi sesuai expected\n= Quality indicator'],
    [6, 'Failed', 'Test cases yang FAILED\n\n= Ditemukan bugs\n= Perlu fixing'],
    [7, 'Blocked', 'Test cases yang BLOCKED\n\n= Tidak bisa ditest karena blocker bugs\n= Dependency issue'],
    [8, 'Pass%', 'Web Pass Rate\n= (Passed / Total) × 100%\n\n🟢 ≥80% = Good quality\n🟡 50-79% = Need attention\n🔴 <50% = Poor quality\n\nQuality metric utama'],
    [9, 'Exec%', 'Web Execution Rate\n= ((Passed + Failed + Blocked) / Total) × 100%\n\n🟢 ≥70% = Good progress\n🟡 40-69% = Medium progress\n🔴 <40% = Slow progress\n\nProgress metric'],
    // API Coverage
    [10, 'Total TC', 'Total test cases API\n\nSemua API TC di API_Master sheet'],
    [11, 'Passed', 'API tests yang PASSED\n\n= API berfungsi sesuai spec\n= Quality indicator'],
    [12, 'Failed', 'API tests yang FAILED\n\n= API bugs ditemukan\n= Perlu fixing'],
    [13, 'Blocked', 'API tests yang BLOCKED\n\n= Tidak bisa ditest\n= Dependency issue'],
    [14, 'Pass%', 'API Pass Rate\n= (Passed / Total) × 100%\n\n🟢 ≥80% = Good quality\n🟡 50-79% = Need attention\n🔴 <50% = Poor quality\n\nQuality metric utama'],
    [15, 'Exec%', 'API Execution Rate\n= ((Passed + Failed + Blocked) / Total) × 100%\n\n🟢 ≥70% = Good progress\n🟡 40-69% = Medium progress\n🔴 <40% = Slow progress\n\nProgress metric']
  ];

  coverageHeaders.forEach(([col, label, note]) => {
    setNote(ws, 3, col, note);
  });

  Logger.log('✅ Coverage notes added');
}

// ═══════════════════════════════════════════════════════════════════════
// CONFIG TAB NOTES
// ═══════════════════════════════════════════════════════════════════════

/**
 * Add notes to Config tab
 */
function addNotesToConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Config');

  if (!ws) {
    Logger.log('⚠️ Config tab not found, skipping notes');
    return;
  }

  Logger.log('📝 Adding notes to Config tab...');

  // Config tab title
  setNote(ws, 1, 1, '⚙️ MODULE CONFIGURATION\n\nKonfigurasi untuk semua modules dalam dashboard\n\nYang perlu diisi:\n• Spreadsheet ID dari QATM sheet masing-masing module\n• Jira project key (jika sync Jira enabled)\n• Notification settings (webhook URL, email, enable flags)\n• Auto refresh settings\n\nData module (Project, Modul, PIC QA) akan otomatis di-pull dari QATM Summary saat refresh');

  // Section header notes already exist in the column headers
  // Just add comprehensive notes to title and important columns

  setNote(ws, 3, 1, 'Active\n\nTRUE = Module aktif, data akan di-pull saat refresh\nFALSE = Module di-skip, tidak muncul di dashboard\n\nGunakan FALSE untuk:\n• Module yang sudah selesai/archived\n• Module yang temporary inactive\n• Testing purposes');

  setNote(ws, 3, 2, 'Jira Sync\n\nTRUE = Sync bugs dari Jira ke QATM BugReport\nFALSE = Skip Jira sync untuk module ini\n\nRequires:\n• Jira Instance diisi (digitalperuri/bgn-peruri)\n• Jira Project diisi (project key)\n• Credentials diisi di Credentials tab');

  setNote(ws, 3, 7, 'Spreadsheet ID\n\nID dari QATM Google Sheet module ini\n\nCara ambil ID:\n1. Buka QATM sheet di browser\n2. Copy dari URL:\n   https://docs.google.com/spreadsheets/d/[ID_INI]/edit\n3. Paste ID ke kolom ini\n\n⚠️ WAJIB diisi agar data bisa di-pull');

  setNote(ws, 3, 12, 'Google Chat Webhook URL\n\nWebhook URL untuk kirim notification ke Google Chat Space\n\nCara buat webhook:\n1. Buka Google Chat Space\n2. Space Settings > Apps & integrations\n3. Webhooks > Add webhook\n4. Copy URL yang generated\n5. Paste ke kolom ini\n\nFormat: https://chat.googleapis.com/v1/spaces/.../messages?key=...\n\n💡 PER-MODULE: Setiap module bisa punya webhook berbeda!');

  setNote(ws, 3, 13, 'Notif Time (Hour)\n\nJam berapa notification dikirim (0-23)\n\nContoh:\n• 7 = 07:00 pagi\n• 15 = 15:00 (3 sore)\n• 18 = 18:00 (6 sore)\n\nNotifikasi akan dikirim daily pada jam ini jika Enable Notifikasi = TRUE');

  setNote(ws, 3, 14, 'Enable Notifikasi\n\nTRUE = Kirim Google Chat notification untuk module ini\nFALSE = Skip notification untuk module ini\n\n💡 PER-MODULE CONFIG:\n• Module dengan TRUE akan kirim notif terpisah\n• Module dengan FALSE tidak akan kirim notif\n• Support multiple webhooks (setiap module bisa beda webhook)\n\nNotif akan dikirim jika ada Blocker atau PROD bugs');

  setNote(ws, 3, 15, 'Email Recipients\n\nEmail address untuk menerima notification\n\nFormat:\n• Single email: user@company.com\n• Multiple emails: user1@company.com,user2@company.com\n  (pisahkan dengan koma, tanpa spasi)\n\n💡 PER-MODULE: Setiap module bisa punya recipients berbeda!');

  setNote(ws, 3, 16, 'Enable Email\n\nTRUE = Kirim Email notification untuk module ini\nFALSE = Skip email untuk module ini\n\n💡 PER-MODULE CONFIG:\n• Module dengan TRUE akan kirim email terpisah\n• Module dengan FALSE tidak akan kirim email\n• Support different recipients per module\n\nEmail akan dikirim jika ada Blocker atau PROD bugs');

  setNote(ws, 3, 17, 'Refresh Interval (Minutes)\n\nInterval auto-refresh dalam menit\n\nContoh:\n• 10 = Refresh setiap 10 menit\n• 30 = Refresh setiap 30 menit\n• 60 = Refresh setiap 1 jam\n\nAuto refresh akan pull data terbaru dari QATM sheets\n\n⚠️ Jangan terlalu sering (<5 menit) untuk avoid quota issues');

  setNote(ws, 3, 18, 'Enable Auto Refresh\n\nTRUE = Enable auto-refresh untuk module ini\nFALSE = Disable auto-refresh, manual refresh only\n\nAuto refresh akan pull data dari QATM sesuai interval\n\n💡 Disable jika:\n• Module jarang update\n• Want to save quota\n• Testing purposes');

  Logger.log('✅ Config notes added');
}

/**
 * Set note on a cell, skip if cell doesn't exist
 * UPDATED: Clear existing note first to ensure overwrite
 */
function setNote(sheet, row, col, noteText) {
  try {
    const cell = sheet.getRange(row, col);
    if (cell) {
      // Clear existing note first
      cell.clearNote();
      // Add new note
      cell.setNote(noteText);
      Logger.log('  ✅ Note updated: ' + sheet.getName() + '!' + cell.getA1Notation());
    }
  } catch (e) {
    Logger.log('  ⚠️ Could not set note at ' + sheet.getName() + ' row ' + row + ', col ' + col + ': ' + e.message);
  }
}

/**
 * Test function - Add note to specific cell for debugging
 */
function testAddSingleNote() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Overview');

  if (!ws) {
    SpreadsheetApp.getUi().alert('Overview not found');
    return;
  }

  // Test add note to cell A1
  ws.getRange('D4').setNote('TEST NOTE - Pass Rate explanation\n\nIf you can see this, notes are working!');

  SpreadsheetApp.getUi().alert(
    'Test Complete',
    'Note added to Overview!D4\n\nHover mouse di cell D4 (PIC QA header) untuk lihat note.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log('Test note added to Overview!D4');
}
