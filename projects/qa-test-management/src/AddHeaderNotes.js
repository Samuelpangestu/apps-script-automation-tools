/**
 * AddHeaderNotes.js - Tambahkan notes/descriptions ke semua headers
 * Run: addNotesToAllHeaders()
 *
 * Objective: Menambahkan tooltip notes di setiap header kolom untuk membantu
 * reviewer, user, dan stakeholder memahami maksud kolom dan nilainya
 */

/**
 * Main function: Tambahkan notes ke semua tab
 */
function addNotesToAllHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('🔖 Adding notes to all headers...');

  addNotesToSummary(ss);
  addNotesToTCMaster(ss);
  addNotesToTCExecution(ss);
  addNotesToAPIMaster(ss);
  addNotesToAPIExecution(ss);
  addNotesToBugReport(ss);
  addNotesToPerfTest(ss);

  SpreadsheetApp.getUi().alert(
    '✅ Header Notes Added',
    'Notes telah ditambahkan ke semua headers di semua tab.\n\n' +
    'Hover mouse di header untuk melihat description.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log('✅ All header notes added successfully');
}

/**
 * Add notes to Summary tab
 */
function addNotesToSummary(ss) {
  const ws = ss.getSheetByName('Summary');
  if (!ws) return;

  Logger.log('  📄 Adding notes to Summary tab...');

  // Test Description headers - already have some notes, ensure consistency
  // Row 2: Project, Modul, Submodul, etc.
  setNote(ws, 2, 1, 'Nama project yang sedang ditest\nContoh: SINET, Customer Portal, dll');
  setNote(ws, 3, 1, 'Modul utama yang ditest\nContoh: 1, 2, 3.1, Order Management, dll');
  setNote(ws, 4, 1, 'Sub-modul detail (bisa lebih dari 1)\nPisahkan dengan koma jika lebih dari 1 submodul\nContoh: 1.1,1.2,1.3');
  setNote(ws, 5, 1, 'QA Team Lead yang bertanggung jawab atas testing');
  setNote(ws, 6, 1, 'Person In Charge - QA yang melakukan testing');
  setNote(ws, 7, 1, 'Environment testing\nPilihan: Dev, Staging/UAT, Production');
  setNote(ws, 8, 1, 'URL untuk tracking issues/bugs\nContoh: Jira, GitHub Issues, dll');
  setNote(ws, 9, 1, 'Status keseluruhan testing\nPilihan: Not Started, In Progress, Completed, On Hold');
  setNote(ws, 10, 1, 'Catatan tambahan tentang scope testing atau hal penting lainnya');

  // API side (right column)
  setNote(ws, 2, 12, 'Base URL untuk API testing\nContoh: https://api.example.com');
  setNote(ws, 3, 12, 'Versi API yang ditest\nContoh: v1, v2, 2.0, dll');
  setNote(ws, 4, 12, 'Metode autentikasi yang digunakan\nContoh: Bearer Token, OAuth2, API Key, dll');
  setNote(ws, 5, 12, 'Environment API testing\nPilihan: Dev, Staging/UAT, Production');
  setNote(ws, 6, 12, 'Link ke Postman Collection atau API documentation');
  setNote(ws, 7, 12, 'Status keseluruhan API testing\nPilihan: Not Started, In Progress, Completed, On Hold');
  setNote(ws, 8, 12, 'Hasil Performance Test (auto-calculated dari PerfTest tab)\nHijau = PASS, Merah = FAIL');
  setNote(ws, 9, 12, 'Catatan tambahan untuk API testing');

  // Status Overview KPI headers (row 13)
  setNote(ws, 13, 1, 'Total test cases di TC_Master');
  setNote(ws, 13, 2, 'Jumlah test case dengan status PASSED');
  setNote(ws, 13, 3, 'Jumlah test case dengan status FAILED');
  setNote(ws, 13, 4, 'Jumlah test case dengan status BLOCKED (tidak bisa dijalankan)');
  setNote(ws, 13, 5, 'Jumlah test case dengan status IN PROGRESS (sedang dijalankan)');
  setNote(ws, 13, 6, 'Jumlah test case dengan status TODO (belum dijalankan)');
  setNote(ws, 13, 7, 'Pass Rate = PASSED / TOTAL\nPersentase test case yang berhasil\nHijau ≥80%, Kuning 50-79%, Merah <50%');
  setNote(ws, 13, 8, 'Auto Rate = Automated / TOTAL\nPersentase test case yang sudah otomatis');
  setNote(ws, 13, 9, 'Execution Rate = (PASSED + FAILED + BLOCKED + IN PROGRESS) / TOTAL\nPersentase test case yang sudah dapat hasil (bukan TODO)');

  // API side KPIs
  setNote(ws, 13, 12, 'Total API test cases di API_Master');
  setNote(ws, 13, 13, 'Jumlah API test dengan status PASSED');
  setNote(ws, 13, 14, 'Jumlah API test dengan status FAILED');
  setNote(ws, 13, 15, 'Jumlah API test dengan status BLOCKED');
  setNote(ws, 13, 16, 'Jumlah API test dengan status IN PROGRESS');
  setNote(ws, 13, 17, 'Jumlah API test dengan status TODO');
  setNote(ws, 13, 18, 'Pass Rate untuk API = PASSED / TOTAL\nHijau ≥80%, Kuning 50-79%, Merah <50%');
  setNote(ws, 13, 19, 'Auto Rate untuk API = Automated / TOTAL');
  setNote(ws, 13, 20, 'Execution Rate untuk API\n= (PASSED + FAILED + BLOCKED + IN PROGRESS) / TOTAL');

  // Bug Summary headers - will be dynamically positioned, typically around row 20+
  // Find BUG SUMMARY section
  const bugRow = findRowWithText(ws, 'BUG SUMMARY');
  if (bugRow > 0) {
    setNote(ws, bugRow + 1, 1, 'Total bugs yang dilaporkan di BugReport tab');
    setNote(ws, bugRow + 1, 2, 'Bugs dengan status Open (baru dilaporkan)');
    setNote(ws, bugRow + 1, 3, 'Bugs dengan status In Progress (sedang dikerjakan)');
    setNote(ws, bugRow + 1, 4, 'Bugs dengan status Reopen (muncul lagi setelah fix)');
    setNote(ws, bugRow + 1, 5, 'Bugs dengan status Fixed (sudah diperbaiki)');
    setNote(ws, bugRow + 1, 6, 'Bugs dengan status Verified (sudah diverifikasi fix-nya)');
    setNote(ws, bugRow + 1, 7, 'Bugs dengan status Closed (sudah selesai)');

    // Open Blocker
    const blockerRow = findRowWithText(ws, 'OPEN BLOCKER');
    if (blockerRow > 0) {
      setNote(ws, blockerRow, 1, 'Jumlah bugs yang masih Open/In Progress/Reopen dengan priority Critical/High/Medium\nBugs ini harus segera ditangani karena menghalangi testing');
    }

    // PROD BUGS
    const prodRow = findRowWithText(ws, 'PROD BUGS');
    if (prodRow > 0) {
      setNote(ws, prodRow, 1, 'Jumlah bugs di Production environment yang masih Open/In Progress/Reopen/Fixed/Verified\n⚠️ CRITICAL: Bugs di production perlu prioritas tinggi');
    }
  }

  Logger.log('    ✓ Summary notes added');
}

/**
 * Add notes to TC_Master tab
 */
function addNotesToTCMaster(ss) {
  const ws = ss.getSheetByName('TC_Master');
  if (!ws) return;

  Logger.log('  📄 Adding notes to TC_Master tab...');

  // Header row typically at row 2
  const headers = [
    ['Test Case ID', 'Unique identifier untuk test case\nContoh: TC001, TC002, dll'],
    ['Modul', 'Modul yang ditest\nHarus sesuai dengan Summary > Modul'],
    ['Test Scenario', 'Deskripsi scenario yang akan ditest\nJelaskan apa yang akan diuji'],
    ['Test Steps', 'Langkah-langkah detail untuk menjalankan test\nBisa multi-line, pisahkan dengan enter'],
    ['Priority', 'Prioritas test case\nCritical = paling penting, High, Medium, Low'],
    ['Precondition', 'Kondisi awal yang harus dipenuhi sebelum test\nContoh: User sudah login, Data sudah ada, dll'],
    ['Expected Result', 'Hasil yang diharapkan dari test\nApa yang harus terjadi jika sistem bekerja dengan benar'],
    ['Test Level', 'Jenis testing\nManual = dijalankan manual oleh QA\nAutomated = dijalankan otomatis dengan script'],
    ['Test Type', 'Kategori test\nFunctional, Integration, Regression, Smoke, dll'],
    ['Environment', 'Environment tempat test dijalankan\nDev, Staging/UAT, Production'],
    ['Feature', 'Nama feature atau fitur yang ditest'],
    ['Notes', 'Catatan tambahan jika ada\nContoh: Known issue, dependency, dll']
  ];

  // Find header row
  const headerRow = 2; // Typically row 2
  headers.forEach((h, idx) => {
    setNote(ws, headerRow, idx + 1, h[1]);
  });

  Logger.log('    ✓ TC_Master notes added');
}

/**
 * Add notes to TC_Execution tab
 */
function addNotesToTCExecution(ss) {
  const ws = ss.getSheetByName('TC_Execution');
  if (!ws) return;

  Logger.log('  📄 Adding notes to TC_Execution tab...');

  // Execution headers (kolom Sprint 1, Sprint 2, dst)
  // Row 8 typically has headers
  setNote(ws, 8, 1, 'Test Case ID dari TC_Master\nAuto-populated, jangan diubah');
  setNote(ws, 8, 2, 'Modul dari TC_Master\nAuto-populated');
  setNote(ws, 8, 3, 'Test Scenario dari TC_Master\nAuto-populated');
  setNote(ws, 8, 4, 'Priority dari TC_Master\nAuto-populated');

  // Sprint columns (dynamic, typically start from col 5+)
  // Add note to row 7 (sprint period info row)
  for (let col = 5; col <= 30; col++) {
    const val = ws.getRange(8, col).getValue();
    if (val && String(val).includes('Sprint')) {
      setNote(ws, 8, col, 'Status hasil test untuk sprint ini\nPilihan:\nPASSED = test berhasil\nFAILED = test gagal\nBLOCKED = test tidak bisa dijalankan\nIN PROGRESS = sedang dijalankan\nTODO = belum dijalankan');
    }
  }

  Logger.log('    ✓ TC_Execution notes added');
}

/**
 * Add notes to API_Master tab
 */
function addNotesToAPIMaster(ss) {
  const ws = ss.getSheetByName('API_Master');
  if (!ws) return;

  Logger.log('  📄 Adding notes to API_Master tab...');

  const headers = [
    ['API Test ID', 'Unique identifier untuk API test\nContoh: API001, API002, dll'],
    ['Modul', 'Modul API yang ditest'],
    ['Endpoint', 'API endpoint yang ditest\nContoh: /api/v1/users, /api/orders, dll'],
    ['Method', 'HTTP method\nGET, POST, PUT, PATCH, DELETE, dll'],
    ['Test Scenario', 'Scenario yang ditest\nContoh: Get user by ID, Create new order, dll'],
    ['Request Body / Params', 'Body atau parameter yang dikirim ke API\nFormat JSON untuk body, key=value untuk params'],
    ['Priority', 'Prioritas API test\nCritical, High, Medium, Low'],
    ['Expected Status Code', 'HTTP status code yang diharapkan\nContoh: 200, 201, 400, 404, 500, dll'],
    ['Expected Response', 'Response yang diharapkan dari API\nBisa berupa JSON structure atau deskripsi'],
    ['Test Level', 'Manual atau Automated'],
    ['Feature', 'Nama feature yang ditest'],
    ['Notes', 'Catatan tambahan']
  ];

  const headerRow = 2;
  headers.forEach((h, idx) => {
    setNote(ws, headerRow, idx + 1, h[1]);
  });

  Logger.log('    ✓ API_Master notes added');
}

/**
 * Add notes to API_Execution tab
 */
function addNotesToAPIExecution(ss) {
  const ws = ss.getSheetByName('API_Execution');
  if (!ws) return;

  Logger.log('  📄 Adding notes to API_Execution tab...');

  setNote(ws, 8, 1, 'API Test ID dari API_Master\nAuto-populated');
  setNote(ws, 8, 2, 'Modul dari API_Master\nAuto-populated');
  setNote(ws, 8, 3, 'Endpoint dari API_Master\nAuto-populated');
  setNote(ws, 8, 4, 'Method dari API_Master\nAuto-populated');
  setNote(ws, 8, 5, 'Test Scenario dari API_Master\nAuto-populated');
  setNote(ws, 8, 6, 'Priority dari API_Master\nAuto-populated');

  // Sprint columns
  for (let col = 7; col <= 30; col++) {
    const val = ws.getRange(8, col).getValue();
    if (val && String(val).includes('Sprint')) {
      setNote(ws, 8, col, 'Status hasil API test untuk sprint ini\nPASSED / FAILED / BLOCKED / IN PROGRESS / TODO');
    }
  }

  Logger.log('    ✓ API_Execution notes added');
}

/**
 * Add notes to BugReport tab
 */
function addNotesToBugReport(ss) {
  const ws = ss.getSheetByName('BugReport');
  if (!ws) return;

  Logger.log('  📄 Adding notes to BugReport tab...');

  const headers = [
    ['Bug ID', 'Unique identifier untuk bug\nBisa dari Jira (contoh: SQA-2026) atau manual (contoh: BUG001)'],
    ['Type', 'Tipe issue: Bug atau Improvement'],
    ['Priority', 'Prioritas bug: Critical, High, Medium, Low\nCritical = sistem down/tidak bisa digunakan\nHigh = fitur utama tidak berfungsi\nMedium = fitur minor tidak berfungsi\nLow = cosmetic/UI issue'],
    ['Status', 'Status bug saat ini:\nOpen = baru dilaporkan\nIn Progress = sedang dikerjakan developer\nReopen = muncul lagi setelah fix\nFixed = sudah diperbaiki developer\nVerified = sudah diverifikasi QA\nClosed = selesai'],
    ['Feature', 'Nama feature/fitur tempat bug ditemukan'],
    ['Submodul', 'Sub-modul detail tempat bug terjadi'],
    ['Title / Summary', 'Judul singkat bug\nContoh: Login button tidak responsive, Error 500 saat submit form'],
    ['Description', 'Deskripsi detail bug\nJelaskan apa yang terjadi'],
    ['Environment', 'Environment tempat bug ditemukan\nDev, Staging/UAT, Production\n⚠️ Production bugs prioritas tinggi!'],
    ['Steps to Reproduce', 'Langkah-langkah untuk reproduce bug\n1. Buka halaman X\n2. Klik button Y\n3. Bug muncul'],
    ['Expected Result', 'Hasil yang seharusnya terjadi (tanpa bug)'],
    ['Actual Result', 'Hasil aktual yang terjadi (dengan bug)'],
    ['Related Test Case', 'Test Case ID yang menemukan bug ini\nContoh: TC001, API005, dll'],
    ['Reported By', 'Nama QA yang melaporkan bug'],
    ['Assigned To', 'Nama developer yang ditugaskan fix bug ini'],
    ['Date Found', 'Tanggal bug ditemukan'],
    ['Date Fixed', 'Tanggal bug selesai diperbaiki'],
    ['Sprint / Release', 'Sprint atau release target untuk fix bug'],
    ['Link / URL', 'Link ke Jira ticket atau URL terkait bug'],
    ['Jira Key', 'Jira ticket key (untuk sync otomatis)\nContoh: SQA-2026, PROJ-123'],
    ['Last Synced', 'Timestamp terakhir sync dari Jira\nAuto-updated oleh sistem'],
    ['Screenshot', 'Link ke screenshot atau attachment']
  ];

  const headerRow = 4; // BugReport headers typically at row 4
  headers.forEach((h, idx) => {
    setNote(ws, headerRow, idx + 1, h[1]);
  });

  Logger.log('    ✓ BugReport notes added');
}

/**
 * Add notes to PerfTest tab
 */
function addNotesToPerfTest(ss) {
  const ws = ss.getSheetByName('PerfTest');
  if (!ws) return;

  Logger.log('  📄 Adding notes to PerfTest tab...');

  // PerfTest headers typically at row 15
  setNote(ws, 15, 1, 'Nomor urut test');
  setNote(ws, 15, 2, 'Nama scenario performance test\nContoh: Login 100 users concurrent, Load 1000 products');
  setNote(ws, 15, 3, 'Endpoint atau page yang ditest');
  setNote(ws, 15, 4, 'HTTP Method (GET/POST/etc) atau action type');
  setNote(ws, 15, 5, 'Jumlah concurrent users/threads\nContoh: 10, 50, 100, 500');
  setNote(ws, 15, 6, 'Durasi test\nContoh: 5 minutes, 30 minutes, 1 hour');
  setNote(ws, 15, 7, 'Target response time (ms)\nContoh: <500ms, <1000ms, <2000ms');
  setNote(ws, 15, 8, 'Actual average response time (ms)\nHasil test yang didapat');
  setNote(ws, 15, 9, 'Target throughput (req/sec)\nBerapa request per detik yang harus ditangani');
  setNote(ws, 15, 10, 'Actual throughput (req/sec)\nHasil actual dari test');
  setNote(ws, 15, 11, 'Error rate (%)\n% request yang gagal/error\nTarget biasanya <1%');
  setNote(ws, 15, 12, 'Hasil test: PASS atau FAIL\nPASS jika semua metrics memenuhi target\nFAIL jika ada metrics yang tidak memenuhi');
  setNote(ws, 15, 13, 'Catatan tambahan\nBottleneck, issue, atau observasi penting');

  Logger.log('    ✓ PerfTest notes added');
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Set note on a cell, skip if cell doesn't exist
 */
function setNote(sheet, row, col, noteText) {
  try {
    const cell = sheet.getRange(row, col);
    if (cell) {
      cell.setNote(noteText);
    }
  } catch (e) {
    // Silently skip if cell doesn't exist
    Logger.log('    ⚠️ Could not set note at row ' + row + ', col ' + col + ': ' + e.message);
  }
}

/**
 * Find row containing specific text in column 1
 */
function findRowWithText(sheet, searchText) {
  try {
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).indexOf(searchText) >= 0) {
        return i + 1; // Return 1-indexed row
      }
    }
  } catch (e) {
    Logger.log('    ⚠️ Error finding text "' + searchText + '": ' + e.message);
  }
  return -1;
}
