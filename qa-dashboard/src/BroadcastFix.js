/**
 * broadcast_notes_fix.js
 * ─────────────────────────────────────────────────────────────────────────
 * Paste ke Apps Script editor QA Dashboard, lalu run broadcastFixNotes().
 * Update notes pada kolom SubModul, TC_ID, Scenario, Steps, Expected Result
 * di TC_Master dan API_Master semua modul aktif -- tanpa menyentuh data lain.
 *
 * FUNCTIONS:
 *   broadcastFixAll()       -- jalankan semua fix sekaligus (RECOMMENDED)
 *   broadcastFixNotes()     -- broadcast notes ke semua modul aktif di Config
 *   broadcastFixAppendix()  -- tambah section Hierarki QA ke Appendix
 *   fixNotesSingleSheet()   -- update sheet yang sedang aktif saja
 * ─────────────────────────────────────────────────────────────────────────
 */

// ── Helper: Safe UI alert (fallback to Logger if no UI context) ───────────
function safeAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    // When called from trigger or Apps Script Editor, getUi() fails
    // Fallback to Logger so script can still complete
    Logger.log('='.repeat(60));
    Logger.log('INFO: ' + message);
    Logger.log('='.repeat(60));
  }
}

// ── Note content ──────────────────────────────────────────────────────────

var SUBMODUL_NOTE_TC =
    'SubModule -- Level ke-3 dalam hierarki QA:\n' +
    '\n' +
    'Hierarki: Project > Module > SubModule > Feature\n' +
    '\n' +
    'SubModule = unit terkecil yang berdiri sendiri,\n' +
    'biasanya 1 aplikasi atau 1 domain dalam project.\n' +
    '\n' +
    'Format berlayer (contoh: SIPGN)\n' +
    '  1.1 = Module 1, SubModule ke-1 (Aplikasi Nutritionist)\n' +
    '  1.2 = Module 1, SubModule ke-2 (Aplikasi Courier)\n' +
    '  2.1 = Module 2, SubModule ke-1\n' +
    '\n' +
    'Format flat -- Module dikosongkan (contoh: INAGOV)\n' +
    '  Nama atau INISIAL jika nama terlalu panjang:\n' +
    '  1. Portal           -> PO\n' +
    '  2. Layanan SmartASN -> SA\n' +
    '  3. BackOffice       -> BO\n' +
    '\n' +
    'Aturan inisial:\n' +
    '  - 2-4 huruf kapital, unik per project\n' +
    '  - Konsisten di TC_Master, API_Master, dan Execution\n' +
    '  - Dokumentasikan di Appendix\n' +
    '\n' +
    'Gunakan nilai KONSISTEN di TC_Master dan API_Master\n' +
    'agar coverage Dashboard akurat.';

var SUBMODUL_NOTE_API =
    'SubModule -- sama dengan SubModule di TC_Master.\n' +
    '\n' +
    'Hierarki: Project > Module > SubModule > Feature\n' +
    '\n' +
    'Gunakan kode yang IDENTIK dengan TC_Master\n' +
    'agar coverage tergabung dengan benar di Summary.\n' +
    '\n' +
    'Contoh berlayer       : 1.1 / 1.2 / 2.1\n' +
    'Contoh flat (nama)    : Talenta / e-Office\n' +
    'Contoh flat (inisial) : PO / SA / BO\n' +
    '  Portal           -> PO\n' +
    '  Layanan SmartASN -> SA\n' +
    '  BackOffice       -> BO';

var TC_ID_NOTE =
    'TC_ID -- Format: [SubModule].[3-digit]\n' +
    '\n' +
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
    'API prefix wajib: API.1.1.001 / API.PO.001 / API.SA.001';

var TC_ID_NOTE_API =
    'TC_ID API -- Format: API.[SubModule].[3-digit]\n' +
    '\n' +
    'Opsi 1 -- Numerik:\n' +
    '  API.1.1.001  API.1.2.015\n' +
    '\n' +
    'Opsi 2 -- Inisial SubModule:\n' +
    '  API.PO.001   = Portal, TC ke-1\n' +
    '  API.SA.001   = Layanan SmartASN, TC ke-1\n' +
    '  API.BO.001   = BackOffice, TC ke-1\n' +
    '\n' +
    'Aturan inisial:\n' +
    '  - 2-3 huruf kapital, UNIK per project\n' +
    '  - Harus IDENTIK dengan inisial yang dipakai di TC_Master\n' +
    '  - Konsisten di TC_Master, API_Master, dan Execution\n' +
    '\n' +
    'Prefix API wajib untuk semua test case API.\n' +
    'Harus UNIK. Jangan ubah TC_ID jika sudah ada hasil di Execution.';

var SCENARIO_NOTE =
    'SCENARIO NAMING STANDARD\n' +
    '\n' +
    'FORMULA\n' +
    '  Happy Path : [Role] + Successfully + [Verb] + [Object] + (from/to [Location])\n' +
    '  Negative   : [Role] + Failed to + [Verb] + [Object] + with [Condition]\n' +
    '\n' +
    'RULES\n' +
    '  - Role   -> Title Case  (Nutritionist, Courier, Beneficiary)\n' +
    '  - Verb   -> Active verb (Create, Pick Up, Confirm, Return)\n' +
    '  - Object -> Title Case  (Menu Plan, Meal Box, Food)\n' +
    '  - Location -> Optional, gunakan "from" atau "to"\n' +
    '\n' +
    'DO NOT USE\n' +
    '  X  success, succeed     (use: Successfully / Failed to)\n' +
    '  X  do, perform, process (sebelum verb utama)\n' +
    '\n' +
    'GHERKIN\n' +
    '  Given : Pre-kondisi / state awal sebelum aksi dimulai\n' +
    '         Contoh: Given user sudah login sebagai Nutritionist\n' +
    '  When  : Aksi yang dilakukan oleh aktor\n' +
    '         Contoh: When user mengisi form dan klik Submit\n' +
    '\n' +
    'JANGAN tulis Then di sini -- Then ada di kolom Expected Result.\n' +
    '\n' +
    'EXAMPLE\n' +
    '  OK  Nutritionist Successfully Creates Meal Plan\n' +
    '  OK  Courier Successfully Picks Up Food from SPPG\n' +
    '  OK  Nutritionist Failed to Create Menu with Incomplete Data';

var SCENARIO_NOTE_API =
    'SCENARIO NAMING STANDARD\n' +
    '\n' +
    'FORMULA\n' +
    '  Happy Path : [Role] + Successfully + [Verb] + [Object]\n' +
    '  Negative   : [Role] + Failed to + [Verb] + [Object] + with [Condition]\n' +
    '\n' +
    'DO NOT USE\n' +
    '  X  success, succeed     (use: Successfully / Failed to)\n' +
    '  X  do, perform, process (sebelum verb utama)\n' +
    '\n' +
    'GHERKIN\n' +
    '  Given : Pre-kondisi / token / state awal\n' +
    '         Contoh: Given user memiliki token Bearer valid\n' +
    '  When  : Request yang dikirim\n' +
    '         Contoh: When POST /api/v1/login dengan payload valid\n' +
    '\n' +
    'JANGAN tulis Then di sini.\n' +
    'Sertakan expected HTTP status di akhir skenario.\n' +
    '\n' +
    'EXAMPLE\n' +
    '  OK  Nutritionist Successfully Creates Meal Plan -- 201\n' +
    '  OK  User Failed to Login with Wrong Password -- 401';

var STEPS_NOTE =
    '[INPUT WAJIB] Steps dalam format Gherkin:\n' +
    '  Given : Pre-kondisi sebelum aksi\n' +
    '         Contoh: Given user sudah login sebagai Courier\n' +
    '  When  : Aksi yang dilakukan\n' +
    '         Contoh: When user klik tombol Pick Up\n' +
    '  And   : Aksi tambahan jika diperlukan\n' +
    '         Contoh: And user konfirmasi dialog\n' +
    '\n' +
    'JANGAN tulis Then di sini -- Then ada di kolom Expected Result.';

var EXPECTED_NOTE =
    '[INPUT WAJIB] Isi dengan format Then Gherkin:\n' +
    '  Then : Hasil / perubahan state setelah aksi selesai\n' +
    '\n' +
    'Tips:\n' +
    '  - Spesifik: sebutkan elemen UI, pesan, atau status yang muncul\n' +
    '  - Contoh: Then halaman dashboard tampil, nama user muncul di header\n' +
    '  - Contoh: Then muncul toast "Berhasil disimpan" dan data terupdate';

// ── Column header → note mapping ──────────────────────────────────────────
var TC_MASTER_NOTES = {
    'SUBMODUL':        SUBMODUL_NOTE_TC,
    'SUBMODULE':       SUBMODUL_NOTE_TC,
    'TC_ID':           TC_ID_NOTE,
    'SCENARIO':        SCENARIO_NOTE,
    'STEPS / GHERKIN': STEPS_NOTE,
    'STEPS':           STEPS_NOTE,
    'EXPECTED RESULT': EXPECTED_NOTE,
    'EXPECTED':        EXPECTED_NOTE,
};

var API_MASTER_NOTES = {
    'SUBMODUL':        SUBMODUL_NOTE_API,
    'SUBMODULE':       SUBMODUL_NOTE_API,
    'TC_ID':           TC_ID_NOTE_API,
    'SCENARIO':        SCENARIO_NOTE_API,
};

// ── Appendix section content ──────────────────────────────────────────────
var APPENDIX_SECTION_TITLE = '0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE';

var APPENDIX_ROWS = [
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
        '    SubModule : Talenta\n' +
        '      Feature: Rekrutmen, Penggajian\n' +
        '    SubModule : e-Office\n' +
        '    SubModule : SIMPEG\n' +
        '\n' +
        'Pada pola flat, SubModule setara dengan Module di pola berlayer.\n' +
        'Kolom Module di Summary dan Config dikosongkan.'
    ],
    [
        'TC_ID per SubModule',
        'Format  : [SubModule].[3-digit]\n' +
        'Pola A  : 1.1.001  (SubModule 1.1, TC ke-1)\n' +
        '          1.2.001  (SubModule 1.2, TC ke-1)\n' +
        'Pola B  : Talenta.001  atau tetap numerik 1.001\n' +
        '\n' +
        'Boleh pakai Inisial jika nama SubModule terlalu panjang:\n' +
        '  Portal           -> PO  -> PO.001\n' +
        '  Layanan SmartASN -> SA  -> SA.001\n' +
        '  BackOffice       -> BO  -> BO.001\n' +
        '\n' +
        'Aturan inisial:\n' +
        '  - 2-3 huruf kapital, unik per project\n' +
        '  - Konsisten di TC_Master, API_Master, dan Execution\n' +
        '  - Dokumentasikan mapping di Appendix (baris ini)\n' +
        '\n' +
        'API prefix wajib: API.1.1.001 / API.PO.001 / API.SA.015'
    ],
];

// ── Auto-detect Config column layout ─────────────────────────────────────
/**
 * Reads Config and auto-detects which column contains Spreadsheet IDs.
 * Returns array of { name, id, project, module, submodule, team }.
 */
function getModulesFromConfig_(cfg) {
    var data = cfg.getDataRange().getValues();
    if (data.length < 4) return [];

    // Find header row (scan rows 0–3)
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

    // Heuristic: scan data rows for column whose cells look like Sheets IDs
    var SHEETS_ID_RE = /^[A-Za-z0-9_\-]{20,}$/;
    if (COL_ID === -1) {
        var dataStart = headerIdx >= 0 ? headerIdx + 1 : 3;
        var colScores = [];
        for (var ci = 0; ci < (data[dataStart] || []).length; ci++) colScores[ci] = 0;
        for (var ri = dataStart; ri < Math.min(dataStart + 10, data.length); ri++) {
            data[ri].forEach(function(cell, ci) {
                var v = String(cell).trim();
                if (v.length > 20 && SHEETS_ID_RE.test(v)) colScores[ci]++;
            });
        }
        var maxScore = 0;
        colScores.forEach(function(s, ci) { if (s > maxScore) { maxScore = s; COL_ID = ci; } });
        Logger.log('Auto-detected Spreadsheet ID column: ' + (COL_ID + 1) + ' (score: ' + maxScore + ')');
    }

    if (COL_ID === -1) {
        Logger.log('ERROR: Cannot detect Spreadsheet ID column in Config');
        return [];
    }

    if (COL_NAME === -1) COL_NAME = 1;

    var modules = [];
    var dataStart2 = headerIdx >= 0 ? headerIdx + 1 : 3;
    for (var r = dataStart2; r < data.length; r++) {
        var row    = data[r];
        var active = String(row[COL_ACTIVE] || '').trim().toUpperCase();
        var id     = String(row[COL_ID]     || '').trim();

        if (active !== 'Y' || !id || id === 'PASTE_SPREADSHEET_ID_HERE' || id.length < 20) continue;

        var project   = '';
        var module_   = '';
        var submodule = '';
        var team      = '';
        if (headerRow) {
            headerRow.forEach(function(h, i) {
                var v = String(row[i] || '').trim();
                if (h === 'PROJECT')                                             project   = v;
                if (h === 'MODULE')                                              module_   = v;
                if (h === 'SUBMODULE' || h === 'SUBMODUL')                      submodule = v;
                if (h === 'PIC / TEAM / SQUAD' || h === 'PIC' || h === 'TEAM') team      = v;
                if ((h === 'MODUL NAME' || h === 'MODUL') && !submodule)        submodule = v;
            });
        }
        var name = submodule || project || String(row[COL_NAME] || '').trim() || id;
        modules.push({ name: name, id: id, project: project, module: module_, submodule: submodule, team: team });
    }

    Logger.log('Config: found ' + modules.length + ' active modules (ID col=' + (COL_ID + 1) + ')');
    return modules;
}

// ── Apply notes by matching header text ──────────────────────────────────
function applyNotes_(ws, noteMap) {
    var lastCol = ws.getLastColumn();
    if (lastCol < 1) return 0;
    var headers = ws.getRange(2, 1, 1, lastCol).getValues()[0];
    var count = 0;
    headers.forEach(function(h, i) {
        var key = String(h).trim().toUpperCase();
        if (noteMap[key]) {
            ws.getRange(2, i + 1).setNote(noteMap[key]);
            Logger.log('  [' + ws.getName() + '] col ' + (i + 1) + ' (' + h + ') -> note set');
            count++;
        }
    });
    return count;
}

// ── broadcastFixNotes ─────────────────────────────────────────────────────
function broadcastFixNotes() {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var cfg = ss.getSheetByName('Config');
    if (!cfg) {
        safeAlert_('Config tab tidak ditemukan.\nPastikan script dijalankan dari QA Dashboard.');
        return;
    }

    var modules = getModulesFromConfig_(cfg);
    if (modules.length === 0) {
        safeAlert_('Tidak ada modul aktif ditemukan di Config.\nCek kolom Active (Y/N) dan Spreadsheet ID.');
        return;
    }

    var done = 0, skipped = 0, failed = 0;

    modules.forEach(function(mod) {
        var cleanId = mod.id.replace(/[^A-Za-z0-9_\-]/g, '');
        if (cleanId.length < 20) {
            Logger.log('SKIP ' + mod.name + ': ID tidak valid [' + mod.id + ']');
            skipped++;
            return;
        }
        Logger.log('Trying: ' + mod.name + ' | ID=' + cleanId);

        try {
            var modSS     = SpreadsheetApp.openById(cleanId);
            var tcMaster  = modSS.getSheetByName('TC_Master');
            var apiMaster = modSS.getSheetByName('API_Master');

            if (!tcMaster && !apiMaster) {
                Logger.log('SKIP ' + mod.name + ': TC_Master dan API_Master tidak ditemukan');
                skipped++;
                return;
            }

            var total = 0;
            if (tcMaster)  total += applyNotes_(tcMaster,  TC_MASTER_NOTES);
            if (apiMaster) total += applyNotes_(apiMaster, API_MASTER_NOTES);

            SpreadsheetApp.flush();
            Logger.log('OK: ' + mod.name + ' (' + total + ' notes updated)');
            done++;
        } catch (e) {
            Logger.log('ERROR ' + mod.name + ' [ID=' + cleanId + ']: ' + e.message);
            failed++;
        }
    });

    safeAlert_(
        'broadcastFixNotes selesai\n\n' +
        'Berhasil : ' + done    + ' modul\n' +
        'Dilewati : ' + skipped + ' modul\n' +
        'Gagal    : ' + failed  + ' modul\n\n' +
        'Lihat Apps Script Logs untuk detail.'
    );
}

// ── fixNotesSingleSheet ───────────────────────────────────────────────────
function fixNotesSingleSheet() {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var tcMaster  = ss.getSheetByName('TC_Master');
    var apiMaster = ss.getSheetByName('API_Master');

    if (!tcMaster && !apiMaster) {
        safeAlert_('TC_Master dan API_Master tidak ditemukan di spreadsheet ini.');
        return;
    }

    var total = 0;
    if (tcMaster)  total += applyNotes_(tcMaster,  TC_MASTER_NOTES);
    if (apiMaster) total += applyNotes_(apiMaster, API_MASTER_NOTES);

    safeAlert_(
        'Done -- ' + total + ' notes diupdate di:\n' +
        (tcMaster  ? '  - TC_Master\n'  : '') +
        (apiMaster ? '  - API_Master\n' : '') +
        '\nSheet: ' + ss.getName()
    );
}

// ── broadcastFixAppendix ──────────────────────────────────────────────────
/**
 * Menambahkan section "0. HIERARKI QA" ke tab Appendix semua modul aktif.
 * Idempotent -- jika section sudah ada, modul dilewati.
 */
function broadcastFixAppendix() {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var cfg = ss.getSheetByName('Config');
    if (!cfg) {
        safeAlert_('Config tab tidak ditemukan.\nJalankan dari QA Dashboard.');
        return;
    }

    var modules = getModulesFromConfig_(cfg);
    if (modules.length === 0) {
        safeAlert_('Tidak ada modul aktif ditemukan di Config.\nCek kolom Active (Y/N) dan Spreadsheet ID.');
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

            // Idempotent: cek apakah section sudah ada
            var lastRow  = ws.getLastRow();
            var existing = lastRow > 0
                ? ws.getRange(1, 1, lastRow, 1).getValues().map(function(r) { return String(r[0]); })
                : [];
            if (existing.some(function(v) { return v.trim() === APPENDIX_SECTION_TITLE; })) {
                Logger.log('SKIP ' + name + ': section sudah ada');
                skipped++;
                return;
            }

            // Insert section at row 2 (preserve row 1 header)
            var insertCount = 1 + APPENDIX_ROWS.length + 1; // header + rows + spacer
            ws.insertRowsBefore(2, insertCount);

            var r = 2;

            // Section header
            ws.getRange(r, 1, 1, 4).merge();
            ws.getRange(r, 1)
                .setValue(APPENDIX_SECTION_TITLE)
                .setBackground('#1565C0').setFontColor('#FFFFFF')
                .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
                .setHorizontalAlignment('left').setVerticalAlignment('middle')
                .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);
            ws.setRowHeight(r, 24);
            r++;

            // Content rows
            APPENDIX_ROWS.forEach(function(rowData) {
                ws.getRange(r, 1)
                    .setValue(rowData[0])
                    .setBackground('#E3F2FD').setFontColor('#0D47A1')
                    .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
                    .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true)
                    .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);
                ws.getRange(r, 2, 1, 3).merge();
                ws.getRange(r, 2)
                    .setValue(rowData[1])
                    .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(9)
                    .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true)
                    .setBorder(true, true, true, true, false, false, '#BBDEFB', SpreadsheetApp.BorderStyle.SOLID);
                ws.setRowHeight(r, 70);
                r++;
            });

            // Spacer row
            ws.setRowHeight(r, 8);

            SpreadsheetApp.flush();
            Logger.log('OK: ' + name + ' -- section ditambahkan');
            done++;

        } catch (e) {
            Logger.log('ERROR ' + name + ': ' + e.message);
            failed++;
        }
    });

    safeAlert_(
        'broadcastFixAppendix selesai\n\n' +
        'Berhasil : ' + done    + ' modul\n' +
        'Dilewati : ' + skipped + ' modul (sudah ada / no Appendix)\n' +
        'Gagal    : ' + failed  + ' modul\n\n' +
        'Lihat Apps Script Logs untuk detail.'
    );
}

// ── broadcastFixAll ───────────────────────────────────────────────────────
/**
 * Shortcut: jalankan broadcastFixNotes() + broadcastFixAppendix() sekaligus.
 */
function broadcastFixAll() {
    broadcastFixNotes();
    broadcastFixAppendix();
    safeAlert_('broadcastFixAll selesai.\nCek Logs untuk detail per modul.');
}
