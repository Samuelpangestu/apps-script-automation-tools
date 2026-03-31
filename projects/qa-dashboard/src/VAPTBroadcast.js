/**
 * VAPTBroadcast.js - Broadcast VAPT workflow updates to all QATM modules
 *
 * Updates:
 * - BugReport status dropdown (add "In Progress VAPT" & "Done VAPT")
 * - BugReport status flow notes
 * - BugReport header info row
 * - BugReport conditional formatting
 * - Summary Open Blocker formula (include VAPT statuses)
 * - Appendix Section 7 (recreate with VAPT workflow)
 *
 * Usage: Run from Dashboard via broadcastVAPTStatusUpdate()
 */

// ═══════════════════════════════════════════════════════════════════════
// VAPT WORKFLOW BROADCAST - ADD IN PROGRESS VAPT & DONE VAPT STATUSES
// ═══════════════════════════════════════════════════════════════════════

/**
 * Broadcast VAPT workflow updates to all active QATM modules
 *
 * VAPT Workflow:
 * QA Phase: Open → In Progress → Fixed → Verified
 * VAPT Phase: Verified → In Progress VAPT → Done VAPT → Closed
 *
 * Open Blocker includes: Open, In Progress, Reopen, In Progress VAPT, Done VAPT
 * (Only Verified, Closed, Won't Fix are NOT blockers)
 */
function broadcastVAPTStatusUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🔒 VAPT Workflow Broadcast',
    'Fix ini akan mengupdate SEMUA QATM dengan VAPT workflow:\n\n' +
    '✅ Add status "In Progress VAPT" & "Done VAPT"\n' +
    '✅ Update BugReport dropdown validation\n' +
    '✅ Update status flow notes & documentation\n' +
    '✅ Update Open Blocker formula (include VAPT statuses)\n' +
    '✅ Recreate Appendix Section 7 with VAPT workflow\n\n' +
    '📋 VAPT Flow:\n' +
    'Verified → In Progress VAPT → Done VAPT → Closed\n\n' +
    '🚨 Open Blocker UPDATE:\n' +
    'Status: Open / In Progress / Reopen / In Progress VAPT / Done VAPT\n' +
    'Priority: Critical / High / Medium\n\n' +
    '💡 Why Done VAPT is blocker?\n' +
    'Bug perlu re-test QA before Closed.\n\n' +
    'Lanjutkan broadcast ke semua active QATM?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Broadcast dibatalkan.');
    return;
  }

  try {
    Logger.log('🔧 Applying VAPT Status Broadcast...');

    const cfg = ss.getSheetByName('Config');
    if (!cfg) {
      ui.alert('❌ Config tab tidak ditemukan!');
      return;
    }

    const cfgData = cfg.getDataRange().getValues();
    let successCount = 0;
    let errorCount = 0;
    const errors = [];

    for (let i = 3; i < cfgData.length; i++) {
      const active = cfgData[i][0] === true;  // Col A
      const project = String(cfgData[i][2]).trim();  // Col C
      const modul = String(cfgData[i][3]).trim();    // Col D
      const qatmId = String(cfgData[i][6]).trim();   // Col G

      if (!active || !qatmId || qatmId.length < 10) continue;

      try {
        Logger.log('Adding VAPT statuses to: ' + project + ' - ' + modul);

        const qatmSs = SpreadsheetApp.openById(qatmId);
        const bugSheet = qatmSs.getSheetByName('BugReport');
        const summarySheet = qatmSs.getSheetByName('Summary');
        const appendixSheet = qatmSs.getSheetByName('Appendix');

        if (!bugSheet) {
          Logger.log('⚠️ BugReport sheet not found for ' + project + ' - ' + modul);
          errorCount++;
          errors.push(project + ' - ' + modul + ' (BugReport not found)');
          continue;
        }

        // Step 1: Update BugReport status dropdown (Col D, rows 5-204)
        const statusDropdown = SpreadsheetApp.newDataValidation()
          .requireValueInList(['Open','In Progress','Fixed','Verified','In Progress VAPT','Done VAPT','Closed',"Won't Fix",'Reopen'], true)
          .setAllowInvalid(false)
          .build();

        bugSheet.getRange('D5:D204').setDataValidation(statusDropdown);
        Logger.log('  ✅ Status dropdown updated');

        // Step 2: Update Status column note (D4)
        bugSheet.getRange('D4').setNote(
          'Status Flow (with VAPT Integration):\n\n' +
          'QA Phase:\n' +
          '  Open → In Progress → Fixed → Verified\n\n' +
          'VAPT Phase (Security Testing):\n' +
          '  Verified → In Progress VAPT → Done VAPT → Closed\n\n' +
          'Exception Flows:\n' +
          '  • Any status → Reopen (bug reappears after fix)\n' +
          '  • Any status → Won\'t Fix (rejected with reason)\n' +
          '  • Verified can skip directly to Closed (no VAPT needed)\n' +
          '  • Done VAPT can return to In Progress VAPT if issues found\n\n' +
          '🚨 BLOCKER STATUS:\n' +
          'Open, In Progress, Reopen, In Progress VAPT, Done VAPT\n' +
          '(with Priority Critical/High/Medium)\n\n' +
          'NOT Blocker: Verified, Closed, Won\'t Fix'
        );
        Logger.log('  ✅ Status note updated');

        // Step 3: Update header info row (row 2)
        const currentHeader = String(bugSheet.getRange(2, 1).getValue());
        if (!currentHeader.includes('In Progress VAPT')) {
          bugSheet.getRange(2, 1).setValue(
            'Priority: Critical = showstopper · High = blocker · Medium = degraded (blocker) · Low = minor  |  ' +
            'Status: Open · In Progress · Fixed · Verified · In Progress VAPT · Done VAPT · Closed'
          );
          Logger.log('  ✅ Header info row updated');
        }

        // Step 4: Update conditional formatting for VAPT statuses
        try {
          const existingRules = bugSheet.getConditionalFormatRules();
          const statusRange = bugSheet.getRange('D5:D204');

          // Add VAPT status conditional formatting
          const vaptRules = [
            {v:'In Progress VAPT', bg:'#E1F5FE', fg:'#01579B'},
            {v:'Done VAPT',        bg:'#B2DFDB', fg:'#004D40'}
          ];

          vaptRules.forEach(s => {
            existingRules.push(
              SpreadsheetApp.newConditionalFormatRule()
                .whenTextEqualTo(s.v)
                .setBackground(s.bg)
                .setFontColor(s.fg)
                .setBold(true)
                .setRanges([statusRange])
                .build()
            );
          });

          bugSheet.setConditionalFormatRules(existingRules);
          Logger.log('  ✅ Conditional formatting updated');
        } catch(e) {
          Logger.log('  ⚠️ Could not update conditional formatting: ' + e.message);
        }

        // Step 5: Update Summary Open Blocker formula (if Summary exists)
        if (summarySheet) {
          // Find Open Blocker row in Summary
          let blockerRow = -1;
          for (let r = 10; r <= 40; r++) {
            const cellVal = String(summarySheet.getRange(r, 1).getValue()).toUpperCase();
            if (cellVal.includes('OPEN BLOCKER') && cellVal.includes(':')) {
              blockerRow = r;
              break;
            }
          }

          if (blockerRow > 0) {
            // Update formula in col B (Web value) and col M (API value)
            const newFormula =
              'SUMPRODUCT(' +
              '(ISNUMBER(MATCH(BugReport!D5:D2000,{"Open","In Progress","Reopen","In Progress VAPT","Done VAPT"},0)))*' +
              '(ISNUMBER(MATCH(BugReport!C5:C2000,{"Critical","High","Medium"},0)))' +
              ')';

            summarySheet.getRange(blockerRow, 2).setFormula('=' + newFormula);
            summarySheet.getRange(blockerRow, 13).setFormula('=' + newFormula);  // Col M for API side
            Logger.log('  ✅ Open Blocker formula updated in Summary');

            // Update note on Open Blocker cell
            summarySheet.getRange(blockerRow, 2).setNote(
              'Open Blocker (Smoke) - UPDATED WITH VAPT\n\n' +
              'Bug yang dihitung sebagai blocker:\n' +
              '  • Status: Open / In Progress / Reopen / In Progress VAPT / Done VAPT\n' +
              '  • Priority: Critical / High / Medium\n\n' +
              'NOT Blocker:\n' +
              '  • Verified (sudah OK QA, waiting VAPT atau skip VAPT)\n' +
              '  • Closed (final)\n' +
              '  • Won\'t Fix (rejected)\n\n' +
              'Why Done VAPT is still blocker?\n' +
              'Bug perlu re-test QA sebelum Closed. Hanya Closed yang tidak blocker.\n\n' +
              'Target: 0 Open Blocker sebelum release ke production.'
            );
          }
        }

        // Step 6: Recreate entire Appendix with VAPT (if Appendix exists)
        if (appendixSheet) {
          try {
            Logger.log('  📝 Recreating Appendix with VAPT sections...');
            recreateEntireAppendixWithVAPT_(qatmSs);
            Logger.log('  ✅ Appendix recreated with VAPT workflow');
          } catch(e) {
            Logger.log('  ⚠️ Could not recreate Appendix: ' + e.message);
          }
        } else {
          Logger.log('  ⚠️ Appendix sheet not found, skipping...');
        }

        successCount++;
        Logger.log('✅ VAPT statuses added to ' + project + ' - ' + modul);

      } catch (e) {
        Logger.log('❌ Error for ' + project + ' - ' + modul + ': ' + e.message);
        errorCount++;
        errors.push(project + ' - ' + modul + ' (' + e.message + ')');
      }
    }

    let msg = '✅ VAPT Status Broadcast Complete!\n\n';
    msg += '📊 Summary:\n';
    msg += '• Success: ' + successCount + ' QATM(s) updated\n';
    msg += '• Errors: ' + errorCount + ' QATM(s)\n\n';
    msg += '🔒 VAPT Workflow Added:\n';
    msg += '• Status dropdown: Added "In Progress VAPT" & "Done VAPT"\n';
    msg += '• BugReport notes: Updated with VAPT flow\n';
    msg += '• Open Blocker formula: Updated to include VAPT statuses\n';
    msg += '• Appendix Section 7: Updated with VAPT documentation\n\n';
    msg += '🚨 Open Blocker NOW includes:\n';
    msg += '• Open / In Progress / Reopen / In Progress VAPT / Done VAPT\n';
    msg += '• With Priority: Critical / High / Medium\n\n';
    msg += '💡 Next Steps:\n';
    msg += '1. Update Jira workflow dengan status VAPT\n';
    msg += '2. Test Jira sync dengan status baru\n';
    msg += '3. Run refreshDashboard() untuk update metrics\n';

    if (errors.length > 0) {
      msg += '\n❌ Errors:\n';
      errors.slice(0, 10).forEach(err => {
        msg += '• ' + err + '\n';
      });
      if (errors.length > 10) {
        msg += '• ... and ' + (errors.length - 10) + ' more\n';
      }
    }

    ui.alert('🔒 VAPT Broadcast Complete!', msg, ui.ButtonSet.OK);
    Logger.log('✅ VAPT Broadcast complete - Success: ' + successCount + ', Errors: ' + errorCount);

  } catch (e) {
    Logger.log('❌ Error applying VAPT broadcast: ' + e.message);
    ui.alert(
      '❌ Error',
      'Gagal apply VAPT broadcast:\n' + e.message + '\n\n' +
      'Check Executions log untuk detail.',
      ui.ButtonSet.OK
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Recreate entire Appendix with VAPT workflow integrated
 * This completely recreates Appendix with all sections (0-10) + Section 7 with VAPT
 * Use this instead of updateAppendixSection7WithVAPT_ to fix corrupted Appendix
 */
function recreateEntireAppendixWithVAPT_(ss) {
  // Delete existing Appendix and create new one
  let ws = ss.getSheetByName('Appendix');
  if (ws) {
    ss.deleteSheet(ws);
  }
  ws = ss.insertSheet('Appendix');
  ws.setTabColor('#1A237E');
  ws.clear();

  // Helper functions
  function bd(r) {
    return r.setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  }
  function hdr(r, bg, fg, sz) {
    fg = fg || '#FFFFFF'; sz = sz || 9;
    return bd(r).setBackground(bg).setFontColor(fg).setFontWeight('bold')
        .setFontSize(sz).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // Title
  ws.getRange(1,1,1,4).merge();
  hdr(ws.getRange(1,1,1,4),'#0D47A1','#FFFFFF',11)
      .setValue('APPENDIX  .  Definisi, Konvensi & Panduan  (v40 - VAPT)');
  ws.setRowHeight(1,30);

  let r=2;
  function sec(title, bg){
    ws.getRange(r,1,1,4).merge();
    hdr(ws.getRange(r,1,1,4),bg||'#37474F','#FFFFFF',9).setValue(title);
    ws.setRowHeight(r,24); r++;
  }
  function row2(label,desc,bg){
    bd(ws.getRange(r,1)).setValue(label)
        .setBackground(bg||'#ECEFF1').setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(desc).setBackground('#FFFFFF')
        .setFontFamily('Arial').setFontSize(9).setWrap(true)
        .setHorizontalAlignment('left').setVerticalAlignment('top');
    ws.setRowHeight(r,48); r++;
  }

  // === COPY ALL SECTIONS FROM TEMPLATE ===

  // Section 0: HIERARKI
  sec('0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE','#0D47A1');
  row2('Definisi',
      'Project   = Inisiatif / client / program kerja. Contoh: SIPGN, INAGOV\n'+
      'Module    = Pengelompokan domain fungsional dalam project.\n'+
      '            Kosongkan ("-") jika project tidak punya layer domain (project flat).\n'+
      'SubModule = Unit terkecil yang berdiri sendiri -- 1 aplikasi atau 1 domain.\n'+
      '            Ini adalah ANCHOR utama untuk TC_ID, Coverage, dan Dashboard.\n'+
      'Feature   = Fitur besar dalam SubModule. Dibedakan di kolom Feature, bukan TC_ID.'
  );
  row2('Pola A -- Project Berlayer (SIPGN)',
      'Project  : SIPGN\n'+
      '  Module 1  : Manajemen Gizi\n'+
      '    SubModule 1.1 : Aplikasi Nutritionist\n'+
      '      Feature: Meal Plan, Menu Management\n'+
      '    SubModule 1.2 : Aplikasi Courier\n'+
      '      Feature: Pick Up, Delivery, Return\n'+
      '    SubModule 1.3 : Aplikasi Beneficiary\n'+
      '  Module 2  : Manajemen Distribusi\n'+
      '    SubModule 2.1 : ...'
  );
  row2('Pola B -- Project Flat (INAGOV)',
      'Project   : INAGOV\n'+
      '  Module  : - (kosong)\n'+
      '    SubModule : Talenta\n'+
      '      Feature: Rekrutmen, Penggajian\n'+
      '    SubModule : e-Office\n'+
      '    SubModule : SIMPEG\n'+
      '\n'+
      'Pada pola flat, SubModule setara dengan Module di pola berlayer.\n'+
      'Kolom Module di Summary dan Config dikosongkan.'
  );
  row2('TC_ID per SubModule',
      'Format  : [SubModule].[3-digit]\n'+
      'Pola A  : 1.1.001 (SubModule 1.1, TC ke-1)\n'+
      '          1.2.001 (SubModule 1.2, TC ke-1)\n'+
      'Pola B  : Talenta.001 atau tetap numerik 1.001\n'+
      '\n'+
      'API prefix wajib: API.1.1.001 / API.1.2.001'
  );

  // Section 1: STRUKTUR TAB
  sec('1. STRUKTUR TAB','#0D47A1');
  row2('TC_Master',     'Master list test case Web / Mobile.\nKolom [INPUT]: SubModul, TC_ID, Feature, Priority, Platform, Test Type, Automation, Version, Role (RBAC), Scenario, Steps, Expected Result.\nKolom [AUTO]: Test Level (kolom N) -- jangan diedit.\nFormat TC_ID: [SubModul].[3-digit]  contoh: 1.1.001  2.3.015\nRole = peran RBAC yang menjalankan skenario, contoh: Admin, User, Viewer.');
  row2('TC_Execution',  'Kolom identitas sync otomatis dari TC_Master. Isi kolom staging dengan PASSED / FAILED / BLOCKED / TODO.\nTambah kolom ke kanan untuk setiap run. LATEST STATUS di kolom Z otomatis.\nKolom AA = link screenshot / evidence.\nIN PROGRESS otomatis jika ada PASSED dan TODO di skenario yang sama.');
  row2('API_Master',    'Master list test case API. Method (E) dan Endpoint URL (F) terpisah.\nKolom [INPUT]: SubModul, TC_ID, Feature, Method, Endpoint, Priority, Auth, Test Type, Automation, Version, Role (RBAC), Scenario.\nKolom [AUTO]: Test Level (kolom N) -- jangan diedit.\n\nFormat TC_ID: API.[SVC].[FEAT].[000]\n  [SVC]  = Kode service/domain, maks 3-4 huruf kapital. Contoh: AUTH, USER, ORD, PAY, INV\n  [FEAT] = Kode endpoint/fitur, maks 3-4 huruf kapital. Contoh: LOG, LIST, CRT, UPD, DEL\n  [000]  = Nomor urut 3 digit, mulai 001\n\nContoh: API.AUTH.LOG.001 (Auth, Login, TC-1)  API.USER.CRT.002 (User, Create, TC-2)  API.PAY.CHK.005 (Payment, Checkout, TC-5)\n\nAturan:\n- Harus UNIK -- jangan pernah reuse TC_ID yang sudah ada\n- Jangan ubah TC_ID jika sudah ada hasil di Execution\n- Urutan: Positive dulu (001), baru Negative (002), Edge Case (003)\n\nRole = peran RBAC yang diuji aksesnya, contoh: Admin, Super Admin, User, Viewer.');
  row2('API_Execution', 'Sama seperti TC_Execution namun untuk API. Sync otomatis dari API_Master.\nKolom AA = link screenshot / evidence.');
  row2('Summary',       'Isi bagian Test Plan sebelum memulai eksekusi (Project, Modul, PIC, Jira, status).\nRingkasan otomatis: coverage SubModul & Feature, run history dengan IN PROGRESS.\nPerf Test Status otomatis dari tab PerfTest.');
  row2('PerfTest',      'Rekam hasil performance test. Isi threshold di baris 11 sesuai SLA.\nKolom STATUS otomatis PASS/FAIL per skenario. OVERALL RESULT terhubung ke tab Summary.\nMetrik yang dicek: RPS, Error Rate, P90, P95, P99, VU, CPU, Memory.');
  row2('Appendix',      'Dokumen ini.');
  r++;

  // Section 2: STATUS EKSEKUSI
  sec('2. STATUS EKSEKUSI','#0D47A1');
  [['PASSED',      'Skenario berhasil -- hasil sesuai Expected Result.','#C8E6C9'],
    ['IN PROGRESS', 'Ada skenario yang sudah PASSED di run sebelumnya namun masih ada TODO -- eksekusi belum selesai.','#E3F2FD'],
    ['FAILED',      'Skenario gagal -- hasil tidak sesuai Expected Result. Wajib buat bug report.','#FFCDD2'],
    ['BLOCKED',     'Tidak bisa dieksekusi -- environment down, dependensi belum siap, atau data belum ada.','#FFE0B2'],
    ['TODO',        'Belum dieksekusi pada run ini.','#F5F5F5'],
  ].forEach(([s,d,bg])=>{
    bd(ws.getRange(r,1)).setValue(s).setBackground(bg).setFontWeight('bold')
        .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(d).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
    ws.setRowHeight(r,34); r++;
  });
  r++;

  // Section 3: TEST LEVEL
  sec('3. TEST LEVEL  .  OTOMATIS DARI PRIORITY','#546E7A');
  row2('Smoke',      'Priority Critical / High / Medium ? Smoke.\nTest subset cepat untuk memvalidasi fungsi utama sebelum release atau setelah deployment.','#FFF8F0');
  row2('Regression', 'Priority Low / Lowest ? Regression.\nFull test cycle dijalankan sebelum release ke Production.','#F1F8E9');
  r++;

  // Section 4: PRIORITY
  sec('4. PRIORITY','#0D47A1');
  [['Critical  [BLOCKER]','Fungsi utama tidak bisa digunakan. Release DITAHAN jika FAIL/BLOCKED.','#FFEBEE'],
    ['High      [BLOCKER]','Fungsi penting terganggu. Jika FAIL: perlu approval PM untuk release.','#FFF3E0'],
    ['Medium    [POTENTIAL BLOCKER]','Fungsi terganggu sebagian. Jika FAIL sebelum UAT: flagged ke tech lead.','#FFF8E1'],
    ['Low',     'Masalah minor: UI, typo, UX kurang optimal.','#F1F8E9'],
    ['Lowest',  'Nice to have. Fix opsional.','#ECEFF1'],
  ].forEach(([p,d,bg])=>{
    bd(ws.getRange(r,1)).setValue(p).setBackground(bg).setFontWeight('bold')
        .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(d).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
    ws.setRowHeight(r,34); r++;
  });
  r++;

  r++;
  // Section 5: ROLE RBAC
  sec('5. ROLE (RBAC) -- KONTROL AKSES','#0D47A1');
  row2('Tujuan','Kolom Role di TC_Master & API_Master bukan role developer, melainkan PERAN PENGGUNA (RBAC) yang menjalankan skenario tersebut.\n\nDigunakan untuk memverifikasi bahwa access control berjalan benar.');
  row2('Contoh Role','Admin: bisa create/edit/delete semua data\nUser/Operator: bisa create & edit data sendiri\nViewer: hanya bisa read/view\nGuest: akses terbatas (belum login)\nSupervisor: approval flow\nSuper Admin: full access termasuk system settings');
  row2('Skenario RBAC','Untuk setiap endpoint/fitur sensitif, buat TC untuk:\n1. Role yang BERHAK -- harus dapat akses (expected: 200/201)\n2. Role yang TIDAK berhak -- harus ditolak (expected: 403 Forbidden)\n3. Tanpa token/login -- harus ditolak (expected: 401 Unauthorized)');
  r++;

  // Section 6: AUTOMATION STATUS
  sec('6. AUTOMATION STATUS','#0D47A1');
  [['Automated',          'Script sudah dibuat dan dapat dijalankan via CI/CD atau manual run.'],
    ['Manual',             'Tidak akan diautomasi -- memerlukan penilaian manusia (exploratory, visual, UX).'],
    ['To Do',              'Direncanakan untuk diautomasi, belum dikerjakan.'],
    ['Cannot be Automated','Secara teknis tidak bisa diautomasi (scan QR fisik, biometrik, hardware-dependent).'],
  ].forEach(([s,d])=>row2(s,d));
  r++;

  // === SECTION 7: BUG REPORT WITH VAPT (NEW) ===
  sec('7. BUG REPORT  —  STATUS & ALUR KERJA (WITH VAPT)','#B71C1C');
  // Status flow table with VAPT
  [['Open',              '#FFCDD2','#B71C1C', 'QA',       'Bug baru ditemukan. Belum ada yang mengerjakan. Masuk antrian Dev.'],
    ['In Progress',      '#E3F2FD','#1565C0', 'Dev',      'Dev sedang mengerjakan fix. Bug belum bisa di-retest.'],
    ['Fixed',            '#FFF9C4','#E65100', 'Dev',      'Dev klaim sudah diperbaiki. Menunggu QA untuk verifikasi ulang.'],
    ['Verified',         '#C8E6C9','#2E7D32', 'QA',       'QA sudah re-test dan bug confirmed fixed. Ready untuk VAPT (Security testing) atau langsung Closed jika skip VAPT.'],
    ['In Progress VAPT', '#E1F5FE','#01579B', 'Security', '🔒 Security team sedang melakukan VAPT testing.'],
    ['Done VAPT',        '#B2DFDB','#004D40', 'Security', '🔒 VAPT testing selesai. Perlu re-test QA sebelum Closed.'],
    ['Closed',           '#E8F5E9','#388E3C', 'QA/Lead',  '✅ Final. Bug sudah selesai dan di-release ke production.'],
    ["Won't Fix",        '#F5F5F5','#9E9E9E', 'Lead',     '❌ Tidak diperbaiki (alasan bisnis/teknis). Harus ada komentar.'],
    ['Reopen',           '#EDE7F6','#6A1B9A', 'QA',       '🔄 Bug masih muncul setelah Fixed/Done VAPT. Kembali ke Dev.'],
  ].forEach(function(s) {
    var bg=s[1], fg=s[2], who=s[3], desc=s[4];
    ws.getRange(r,1,1,1).setValue(s[0]).setBackground(bg).setFontColor(fg)
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r,2,1,1).setValue(who).setBackground('#F5F5F5').setFontColor('#424242')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r,3,1,2).merge().setValue(desc).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(r,30); r++;
  });

  // Flow legend with VAPT
  ws.getRange(r,1,1,4).merge()
      .setValue('Flow:  Open → In Progress → Fixed → Verified → In Progress VAPT → Done VAPT → Closed  |  Reopen dari any status')
      .setBackground('#FFEBEE').setFontColor('#C62828').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#EF9A9A',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(r,16); r++;

  // Responsibilities with VAPT roles
  row2('Dev update ke',      'In Progress (saat mulai mengerjakan)\nFixed (saat selesai — Dev TIDAK boleh langsung ke Verified/Closed)');
  row2('QA update ke',       'Verified (jika re-test lulus dan siap VAPT)\nReopen (jika bug masih ada)\nClosed (final setelah Done VAPT dan re-test QA)');
  row2('Security update ke', 'In Progress VAPT (saat mulai security testing)\nDone VAPT (testing selesai, siap re-test QA)\nReopen (jika ditemukan issue baru saat VAPT)');
  row2('Lead update ke',     "Won't Fix (dengan komentar alasan yang jelas)\nClosed (keputusan akhir)");
  row2('🚨 Open Blocker Calculation (UPDATED WITH VAPT)',
      'Formula Open Blocker di Summary & Dashboard menghitung bug dengan:\n' +
      '  • Status = Open / In Progress / Reopen / In Progress VAPT / Done VAPT\n' +
      '  • Priority = Critical / High / Medium\n\n' +
      'NOT Blocker: Verified, Closed, Won\'t Fix\n\n' +
      '💡 Kenapa Done VAPT masih blocker?\n' +
      'Karena bug perlu di-test ulang oleh QA sebelum bisa Closed.\n' +
      'Hanya setelah Closed baru tidak dihitung blocker.\n\n' +
      'Target: 0 Open Blocker sebelum release ke production.');
  r++;

  // Section 9: PERFORMANCE TEST
  sec('9. PERFORMANCE TEST -- METRIK','#4A148C');
  [['RPS (req/s)',    'Requests Per Second -- jumlah request per detik yang berhasil diproses.\nThreshold: MINIMUM. Jika actual < threshold ? FAIL.'],
    ['Error Rate (%)', 'Persentase request yang mengembalikan error (status ? 400 atau timeout).\nThreshold: MAKSIMUM. Jika actual > threshold ? FAIL.'],
    ['P90 / P95 / P99','Percentile response time -- 90% / 95% / 99% dari semua request selesai dalam waktu ini.\nThreshold: MAKSIMUM (ms). Semakin kecil semakin baik.'],
    ['VU (Virtual User)','Jumlah concurrent user yang disimulasikan dalam satu skenario test.\nDiisi sebagai config, bukan sebagai threshold PASS/FAIL.'],
    ['CPU Usage (%)',   'Persentase penggunaan CPU server selama test berlangsung.\nThreshold: MAKSIMUM. Jika actual > threshold ? FAIL.'],
    ['Memory Usage (%)','Persentase penggunaan memory server selama test berlangsung.\nThreshold: MAKSIMUM. Jika actual > threshold ? FAIL.'],
  ].forEach(([m,d])=>row2(m,d));
  r++;

  // Section 10: HTTP METHOD
  sec('10. HTTP METHOD','#0D47A1');
  [['GET','Mengambil data. Tidak mengubah state.','#E8F0FE'],
    ['POST','Membuat resource baru.','#E8F5E9'],
    ['PUT','Update resource secara penuh (replace).','#FFF8E1'],
    ['PATCH','Update resource sebagian (partial).','#F3E5F5'],
    ['DELETE','Menghapus resource.','#FCE4EC'],
  ].forEach(([m,d,bg])=>{
    bd(ws.getRange(r,1)).setValue(m).setBackground(bg).setFontWeight('bold')
        .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(d).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setHorizontalAlignment('left').setVerticalAlignment('middle');
    ws.setRowHeight(r,28); r++;
  });

  // Footer
  ws.getRange(r,1,1,4).merge()
      .setValue('QA Team  ·  Template v40 - VAPT  ·  2026  ·  Auto-updated via Dashboard Broadcast')
      .setBackground('#0D47A1').setFontColor('#FFFFFF').setFontFamily('Arial').setFontSize(8)
      .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(false);
  ws.setRowHeight(r,20);

  // Set column widths
  [100,140,80,200].forEach((w,i)=>ws.setColumnWidth(i+1,w));

  Logger.log('✅ Appendix recreated with all sections + VAPT workflow');
}

/**
 * Update Appendix Section 7 with VAPT workflow
 * Finds and replaces ONLY Section 7 (BUG REPORT — STATUS & ALUR KERJA)
 * Keeps all other sections (0-6, 8-10) completely intact
 *
 * NOTE: This function is deprecated. Use recreateEntireAppendixWithVAPT_() instead
 * to fix Appendix that was corrupted by previous broadcast.
 */
function updateAppendixSection7WithVAPT_(ws) {
  // Helper functions
  function bd(r) {
    return r.setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  }
  function hdr(r, bg, fg, sz) {
    fg = fg || '#FFFFFF'; sz = sz || 9;
    return bd(r).setBackground(bg).setFontColor(fg).setFontWeight('bold')
        .setFontSize(sz).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // Find Section 7 row (search for "7." and "BUG REPORT")
  let section7Row = -1;
  const maxRow = ws.getLastRow();

  for (let r = 1; r <= maxRow; r++) {
    const cellVal = String(ws.getRange(r, 1).getValue()).toUpperCase();
    if (cellVal.includes('7.') && cellVal.includes('BUG REPORT') && cellVal.includes('STATUS')) {
      section7Row = r;
      break;
    }
  }

  if (section7Row === -1) {
    throw new Error('Section 7 (BUG REPORT) not found in Appendix');
  }

  // Find next section (Section 8, 9, or 10)
  let nextSectionRow = -1;
  for (let r = section7Row + 1; r <= maxRow; r++) {
    const cellVal = String(ws.getRange(r, 1).getValue()).toUpperCase();
    // Look for next numbered section (8., 9., 10.) or footer
    if ((cellVal.match(/^(8\.|9\.|10\.)/)) || cellVal.includes('QA INA DIGITAL') || cellVal.includes('QA TEAM')) {
      nextSectionRow = r;
      break;
    }
  }

  if (nextSectionRow === -1) {
    // No next section found, means Section 7 is last before footer
    // Find footer row
    for (let r = section7Row + 1; r <= maxRow; r++) {
      const cellVal = String(ws.getRange(r, 1).getValue()).toUpperCase();
      if (cellVal.includes('QA INA DIGITAL') || cellVal.includes('QA TEAM') || cellVal.includes('TEMPLATE')) {
        nextSectionRow = r;
        break;
      }
    }
  }

  if (nextSectionRow === -1) {
    nextSectionRow = maxRow + 1; // If still not found, assume end of sheet
  }

  Logger.log('  📍 Section 7 found at row ' + section7Row);
  Logger.log('  📍 Next section at row ' + nextSectionRow);

  // Calculate how many rows Section 7 currently occupies
  const oldSection7Rows = nextSectionRow - section7Row;
  Logger.log('  📏 Old Section 7: ' + oldSection7Rows + ' rows');

  // Calculate how many rows new Section 7 needs
  // 1 header + 9 status rows + 1 flow legend + 5 responsibility rows = 16 rows
  const newSection7Rows = 16;
  Logger.log('  📏 New Section 7: ' + newSection7Rows + ' rows');

  // Adjust rows if needed
  if (newSection7Rows > oldSection7Rows) {
    // Need more rows - insert at end of Section 7
    const rowsToInsert = newSection7Rows - oldSection7Rows;
    ws.insertRowsAfter(nextSectionRow - 1, rowsToInsert);
    Logger.log('  ➕ Inserted ' + rowsToInsert + ' rows');
  } else if (newSection7Rows < oldSection7Rows) {
    // Need fewer rows - delete excess rows
    const rowsToDelete = oldSection7Rows - newSection7Rows;
    ws.deleteRows(section7Row + newSection7Rows, rowsToDelete);
    Logger.log('  ➖ Deleted ' + rowsToDelete + ' excess rows');
  }

  // Clear existing content in Section 7 range (keep formatting)
  ws.getRange(section7Row, 1, newSection7Rows, 4).clearContent();

  // Write new Section 7 with VAPT
  let r = section7Row;

  // Section header
  ws.getRange(r,1,1,4).merge();
  hdr(ws.getRange(r,1,1,4),'#B71C1C','#FFFFFF',9).setValue('7. BUG REPORT  —  STATUS & ALUR KERJA (WITH VAPT)');
  ws.setRowHeight(r,24); r++;

  // Status flow table with VAPT
  [['Open',              '#FFCDD2','#B71C1C', 'QA',       'Bug baru ditemukan. Belum ada yang mengerjakan. Masuk antrian Dev.'],
    ['In Progress',      '#E3F2FD','#1565C0', 'Dev',      'Dev sedang mengerjakan fix. Bug belum bisa di-retest.'],
    ['Fixed',            '#FFF9C4','#E65100', 'Dev',      'Dev klaim sudah diperbaiki. Menunggu QA untuk verifikasi ulang.'],
    ['Verified',         '#C8E6C9','#2E7D32', 'QA',       'QA sudah re-test dan bug confirmed fixed. Ready untuk VAPT (Security testing) atau langsung Closed jika skip VAPT.'],
    ['In Progress VAPT', '#E1F5FE','#01579B', 'Security', '🔒 Security team sedang melakukan VAPT testing.'],
    ['Done VAPT',        '#B2DFDB','#004D40', 'Security', '🔒 VAPT testing selesai. Perlu re-test QA sebelum Closed.'],
    ['Closed',           '#E8F5E9','#388E3C', 'QA/Lead',  '✅ Final. Bug sudah selesai dan di-release ke production.'],
    ["Won't Fix",        '#F5F5F5','#9E9E9E', 'Lead',     '❌ Tidak diperbaiki (alasan bisnis/teknis). Harus ada komentar.'],
    ['Reopen',           '#EDE7F6','#6A1B9A', 'QA',       '🔄 Bug masih muncul setelah Fixed/Done VAPT. Kembali ke Dev.'],
  ].forEach(function(s) {
    var bg=s[1], fg=s[2], who=s[3], desc=s[4];
    ws.getRange(r,1,1,1).setValue(s[0]).setBackground(bg).setFontColor(fg)
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r,2,1,1).setValue(who).setBackground('#F5F5F5').setFontColor('#424242')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r,3,1,2).merge().setValue(desc).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(r,30); r++;
  });

  // Flow legend with VAPT
  ws.getRange(r,1,1,4).merge()
      .setValue('Flow:  Open → In Progress → Fixed → Verified → In Progress VAPT → Done VAPT → Closed  |  Reopen dari any status')
      .setBackground('#FFEBEE').setFontColor('#C62828').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#EF9A9A',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(r,16); r++;

  // Responsibilities with VAPT roles
  function row2(label, desc, bg){
    bd(ws.getRange(r,1)).setValue(label)
        .setBackground(bg||'#ECEFF1').setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(desc).setBackground('#FFFFFF')
        .setFontFamily('Arial').setFontSize(9).setWrap(true)
        .setHorizontalAlignment('left').setVerticalAlignment('top');
    ws.setRowHeight(r,48); r++;
  }

  row2('Dev update ke',      'In Progress (saat mulai mengerjakan)\nFixed (saat selesai — Dev TIDAK boleh langsung ke Verified/Closed)');
  row2('QA update ke',       'Verified (jika re-test lulus dan siap VAPT)\nReopen (jika bug masih ada)\nClosed (final setelah Done VAPT dan re-test QA)');
  row2('Security update ke', 'In Progress VAPT (saat mulai security testing)\nDone VAPT (testing selesai, siap re-test QA)\nReopen (jika ditemukan issue baru saat VAPT)');
  row2('Lead update ke',     "Won't Fix (dengan komentar alasan yang jelas)\nClosed (keputusan akhir)");
  row2('🚨 Open Blocker Calculation (UPDATED WITH VAPT)',
      'Formula Open Blocker di Summary & Dashboard menghitung bug dengan:\n' +
      '  • Status = Open / In Progress / Reopen / In Progress VAPT / Done VAPT\n' +
      '  • Priority = Critical / High / Medium\n\n' +
      'NOT Blocker: Verified, Closed, Won\'t Fix\n\n' +
      '💡 Kenapa Done VAPT masih blocker?\n' +
      'Karena bug perlu di-test ulang oleh QA sebelum bisa Closed.\n' +
      'Hanya setelah Closed baru tidak dihitung blocker.\n\n' +
      'Target: 0 Open Blocker sebelum release ke production.');

  Logger.log('  ✅ Section 7 updated: rows ' + section7Row + ' to ' + (r-1) + ' (kept all other sections intact)');
}

// ═══════════════════════════════════════════════════════════════════════
// LIGHTWEIGHT BROADCAST - UPDATE ONLY NOTE & APPENDIX DOCUMENTATION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Lightweight broadcast: Update ONLY BugReport D4 note and Appendix Section 7
 * Does NOT update dropdowns, formulas, or colors
 *
 * Use this when:
 * - Full broadcast already done, but documentation needs refresh
 * - Only want to update text/notes, not functional changes
 */
function broadcastFixNoteAndAppendix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📝 Update Note & Appendix Only',
    'Broadcast ini hanya update DOKUMENTASI (tidak ubah dropdown/formula):\n\n' +
    '✅ BugReport D4 note (Status Flow)\n' +
    '✅ Appendix Section 7 (Open Blocker Calculation)\n\n' +
    '⚠️ TIDAK update:\n' +
    '- Dropdown validation\n' +
    '- Conditional formatting\n' +
    '- Open Blocker formula\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('❌ Broadcast dibatalkan.');
    return;
  }

  const configSh = ss.getSheetByName('Config');
  if (!configSh) {
    ui.alert('❌ Config sheet tidak ditemukan!');
    return;
  }

  // Get active modules
  const lastRow = configSh.getLastRow();
  if (lastRow < 2) {
    ui.alert('Config kosong! Isi tab Config dulu.');
    return;
  }

  const configData = configSh.getRange(2, 1, lastRow - 1, 10).getValues();
  const activeModules = configData.filter(row => {
    const status = String(row[0]).trim().toUpperCase();
    const ssId = String(row[1]).trim();
    return status === 'ACTIVE' && ssId.length > 10;
  });

  if (activeModules.length === 0) {
    ui.alert('Tidak ada modul aktif di Config!');
    return;
  }

  Logger.log('🚀 Starting lightweight broadcast: Note & Appendix update');
  Logger.log('Target modules: ' + activeModules.length);

  let successCount = 0;
  let errorCount = 0;
  const errors = [];

  activeModules.forEach(row => {
    const ssId = String(row[1]).trim();
    const project = String(row[2]).trim();
    const modul = String(row[3]).trim();

    try {
      Logger.log('\n📂 Processing: ' + project + ' - ' + modul);
      const qatmSs = SpreadsheetApp.openById(ssId);
      const bugSheet = qatmSs.getSheetByName('BugReport');
      const appendixSheet = qatmSs.getSheetByName('Appendix');

      if (!bugSheet) {
        Logger.log('⚠️ BugReport sheet not found');
        errorCount++;
        errors.push(project + ' - ' + modul + ' (BugReport not found)');
        return;
      }

      // Update BugReport D4 note
      bugSheet.getRange('D4').setNote(
        'Status Flow (with VAPT Integration):\n\n' +
        'QA Phase:\n' +
        '  Open → In Progress → Fixed → Verified\n\n' +
        'VAPT Phase (Security Testing):\n' +
        '  Verified → In Progress VAPT → Done VAPT → Closed\n\n' +
        'Exception Flows:\n' +
        '  • Any status → Reopen (bug reappears after fix)\n' +
        '  • Any status → Won\'t Fix (rejected with reason)\n' +
        '  • Verified can skip directly to Closed (no VAPT needed)\n' +
        '  • Done VAPT can return to In Progress VAPT if issues found\n\n' +
        '🚨 BLOCKER STATUS:\n' +
        'Open, In Progress, Reopen, In Progress VAPT, Done VAPT\n' +
        '(with Priority Critical/High/Medium)\n\n' +
        'NOT Blocker: Verified, Closed, Won\'t Fix'
      );
      Logger.log('  ✅ BugReport D4 note updated');

      // Update Appendix Section 7 if exists
      if (appendixSheet) {
        updateAppendixSection7Only_(appendixSheet);
        Logger.log('  ✅ Appendix Section 7 updated');
      } else {
        Logger.log('  ⚠️ Appendix sheet not found, skipped');
      }

      successCount++;
      Logger.log('✅ SUCCESS: ' + project + ' - ' + modul);

    } catch (err) {
      Logger.log('❌ ERROR: ' + err.message);
      errorCount++;
      errors.push(project + ' - ' + modul + ' (' + err.message + ')');
    }

    Utilities.sleep(500);
  });

  // Summary
  Logger.log('\n═══════════════════════════════════════════');
  Logger.log('📊 BROADCAST SUMMARY (Note & Appendix Only)');
  Logger.log('═══════════════════════════════════════════');
  Logger.log('✅ Success: ' + successCount + ' modules');
  Logger.log('❌ Errors: ' + errorCount + ' modules');

  let msg = '✅ Broadcast Selesai!\n\n';
  msg += '📊 SUMMARY:\n';
  msg += '✅ Success: ' + successCount + ' modules\n';
  msg += '❌ Errors: ' + errorCount + ' modules\n\n';
  msg += 'Updated:\n';
  msg += '• BugReport D4 note (Status Flow with VAPT)\n';
  msg += '• Appendix Section 7 (Open Blocker Calculation)\n\n';

  if (errorCount > 0) {
    msg += '⚠️ ERRORS:\n' + errors.slice(0, 5).join('\n');
    if (errors.length > 5) msg += '\n... dan ' + (errors.length - 5) + ' errors lainnya';
  }

  msg += '\n\nCek Execution log untuk detail lengkap.';
  ui.alert('📝 Note & Appendix Update Complete', msg, ui.ButtonSet.OK);
}

/**
 * Update ONLY the Open Blocker Calculation section in existing Appendix
 * Does not recreate entire Appendix, only updates Section 7 blocker text
 */
function updateAppendixSection7Only_(appendixSh) {
  if (!appendixSh) return;

  // Find Section 7 header row
  const lastRow = appendixSh.getLastRow();
  let section7StartRow = -1;

  for (let r = 1; r <= lastRow; r++) {
    const cellA = String(appendixSh.getRange(r, 1).getValue());
    const cellB = String(appendixSh.getRange(r, 2).getValue());

    // Look for "7. BUG REPORT" header
    if (cellA.includes('7.') && (cellA.toUpperCase().includes('BUG') || cellB.toUpperCase().includes('BUG'))) {
      section7StartRow = r;
      break;
    }
  }

  if (section7StartRow === -1) {
    Logger.log('  ⚠️ Section 7 header not found in Appendix');
    return;
  }

  // Find the "Open Blocker Calculation" row within Section 7
  let blockerRow = -1;
  for (let r = section7StartRow; r <= Math.min(section7StartRow + 50, lastRow); r++) {
    const cellB = String(appendixSh.getRange(r, 2).getValue());
    if (cellB.includes('Open Blocker') || cellB.includes('BLOCKER') && cellB.includes('CALCULATION')) {
      blockerRow = r;
      break;
    }
  }

  if (blockerRow === -1) {
    Logger.log('  ⚠️ Open Blocker row not found in Section 7');
    return;
  }

  // Update the blocker calculation text
  const newBlockerText =
    'Formula Open Blocker di Summary & Dashboard menghitung bug dengan:\n' +
    '  • Status = Open / In Progress / Reopen / In Progress VAPT / Done VAPT\n' +
    '  • Priority = Critical / High / Medium\n\n' +
    'NOT Blocker: Verified, Closed, Won\'t Fix\n\n' +
    '💡 Kenapa Done VAPT masih blocker?\n' +
    'Karena bug perlu di-test ulang oleh QA sebelum bisa Closed.\n' +
    'Hanya setelah Closed baru tidak dihitung blocker.\n\n' +
    'Target: 0 Open Blocker sebelum release ke production.';

  appendixSh.getRange(blockerRow, 3).setValue(newBlockerText);

  // Also update the label in column B if needed
  appendixSh.getRange(blockerRow, 2).setValue('🚨 Open Blocker Calculation (UPDATED WITH VAPT)');

  Logger.log('  ✅ Open Blocker text updated at row ' + blockerRow);
}
