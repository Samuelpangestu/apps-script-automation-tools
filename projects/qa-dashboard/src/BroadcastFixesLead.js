/**
 * broadcast_fix_lead_and_examples.js
 * ─────────────────────────────────────────────────────────────────────────
 * Paste ke Apps Script editor QA PORTFOLIO DASHBOARD.
 * Dua fix terpisah — bisa dijalankan satu-satu atau sekaligus.
 *
 * FUNGSI:
 *   broadcastBothFixes()          → jalankan Fix 1 + Fix 2 sekaligus
 *   ── atau satu per satu: ──
 *   fixDashboardLeadManualOnly()  → Fix 1: QA Team Lead = manual input saja
 *   broadcastRemoveExamples()     → Fix 2: hapus kolom Examples dari TC_Master
 *                                          dan API_Master semua modul aktif
 *
 * PERUBAHAN:
 *   Fix 1 — Dashboard Config:
 *     • Kolom QA Team Lead tidak lagi auto-sync dari Summary B4
 *     • Manual input saja di Config
 *     • Note kolom diupdate
 *
 *   Fix 2 — Semua modul (TC_Master & API_Master):
 *     • Kolom Examples dihapus dari TC_Master (was col N)
 *     • Kolom Examples dihapus dari API_Master (was col O)
 *     • Contoh Scenario Outline dipindah ke note kolom Scenario
 *     • TC_Execution & API_Execution sync TestLevel formula dikoreksi
 *
 * SAFE:
 *     • Cek header sebelum delete — skip jika kolom tidak ada
 *     • Tidak menghapus data kolom lain
 * ─────────────────────────────────────────────────────────────────────────
 */


// ══════════════════════════════════════════════════════════════════════════
// MASTER RUNNER
// ══════════════════════════════════════════════════════════════════════════
function broadcastBothFixes() {
  var results = [];
  Logger.log('broadcastBothFixes mulai...');

  // Fix 1: Dashboard Lead
  try {
    var r1 = fixDashboardLeadManualOnly_();
    results.push('✅  Dashboard Lead  : ' + r1);
  } catch(e) {
    results.push('❌  Dashboard Lead  : ' + e.message);
    Logger.log('Fix 1 ERROR: ' + e.message);
  }

  // Fix 2: Remove Examples column from all modules
  try {
    var r2 = broadcastRemoveExamples_(false);
    results.push('✅  Remove Examples : ' + r2);
  } catch(e) {
    results.push('❌  Remove Examples : ' + e.message);
    Logger.log('Fix 2 ERROR: ' + e.message);
  }

  var msg = 'broadcastBothFixes selesai\n\n' + results.join('\n');
  Logger.log(msg);
  safeAlertFix_(msg);
}


// ══════════════════════════════════════════════════════════════════════════
// FIX 1 — DASHBOARD: QA TEAM LEAD = MANUAL INPUT ONLY
// ══════════════════════════════════════════════════════════════════════════
function fixDashboardLeadManualOnly() {
  var result = fixDashboardLeadManualOnly_();
  safeAlertFix_('✅  ' + result);
}

function fixDashboardLeadManualOnly_() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');
  if (!cfg) throw new Error('Config tab tidak ditemukan.');

  var log = [];

  // Find QA Team Lead column in Config header row (row 3)
  var lastCol  = cfg.getLastColumn();
  var hdrRow   = cfg.getRange(3, 1, 1, lastCol).getValues()[0];
  var leadCol  = -1;
  hdrRow.forEach(function(h, i) {
    if (String(h).trim() === 'QA Team Lead') leadCol = i + 1;
  });

  if (leadCol === -1) {
    log.push('Config: kolom QA Team Lead tidak ditemukan — skip');
  } else {
    // Update header note — remove "auto-sync" mention
    cfg.getRange(3, leadCol).setNote(
      'QA Team Lead\n' +
      'Nama QA Team Lead untuk modul ini.\n' +
      'Isi manual sesuai PIC QA — tidak auto-sync dari modul.'
    );
    log.push('Config: note QA Team Lead diupdate (manual input only)');
  }

  var result = log.join('\n');
  Logger.log('fixDashboardLeadManualOnly: ' + result);
  return result;
}


// ══════════════════════════════════════════════════════════════════════════
// FIX 2 — BROADCAST: HAPUS KOLOM EXAMPLES DARI SEMUA MODUL
// ══════════════════════════════════════════════════════════════════════════
function broadcastRemoveExamples() {
  broadcastRemoveExamples_(true);
}

function broadcastRemoveExamples_(showAlert) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');
  if (!cfg) throw new Error('Config tab tidak ditemukan');

  var allData  = cfg.getDataRange().getValues();
  var hdrRow   = allData[2] || [];
  var idColIdx = detectIdColFix_(hdrRow);

  var ok = 0, skip = 0, err = 0, errList = [];

  for (var i = 3; i < allData.length; i++) {
    var row = allData[i];
    if (String(row[0]).trim().toUpperCase() !== 'Y') continue;
    var id = String(row[idColIdx]).trim();
    if (!id || id.length < 10 || id === 'PASTE_SPREADSHEET_ID_HERE') continue;

    try {
      var res = removeExamplesFromSheet_(id);
      if (res === 'skipped') {
        skip++;
        Logger.log('SKIP [' + id.substring(0,20) + ']: Examples column not found');
      } else {
        ok++;
        Logger.log('OK   [' + id.substring(0,20) + ']: ' + res);
      }
    } catch(e) {
      err++;
      errList.push('  • ' + id.substring(0,25) + ': ' + e.message);
      Logger.log('ERR  [' + id.substring(0,20) + ']: ' + e.message);
    }
  }

  var summary =
    'Remove Examples\n' +
    'Berhasil : ' + ok   + ' modul\n' +
    'Skip     : ' + skip + ' modul (kolom tidak ada)\n' +
    'Gagal    : ' + err  + ' modul';
  if (errList.length) summary += '\n' + errList.join('\n');
  Logger.log(summary);
  if (showAlert) safeAlertFix_('✅  ' + summary);
  return summary;
}


// ── Core fix for one spreadsheet ──────────────────────────────────────────
function removeExamplesFromSheet_(spreadsheetId) {
  var src = SpreadsheetApp.openById(spreadsheetId);
  var log = [];

  var tcDone  = removeTcMasterExamples_(src);
  var apiDone = removeApiMasterExamples_(src);

  if (tcDone === 'skipped' && apiDone === 'skipped') return 'skipped';
  log.push('TC:' + tcDone + ' API:' + apiDone);
  return log.join(' ');
}


// ── TC_Master: remove Examples column ────────────────────────────────────
function removeTcMasterExamples_(src) {
  var ws = src.getSheetByName('TC_Master');
  if (!ws) return 'skipped (no TC_Master)';

  var lastCol = ws.getLastColumn();
  var hdr     = ws.getRange(2, 1, 1, lastCol).getValues()[0];

  // Find Examples col — should be between Expected Result and [AUTO] Test Level
  var examplesCol = -1;
  hdr.forEach(function(h, i) {
    if (String(h).trim() === 'Examples') examplesCol = i + 1;
  });

  if (examplesCol === -1) {
    // Also check if Scenario note needs updating even without Examples col
    updateTcScenarioNote_(ws);
    return 'skipped';
  }

  // Verify column identity before deleting
  var colName = String(hdr[examplesCol - 1]).trim();
  if (colName !== 'Examples') {
    throw new Error('TC_Master col ' + examplesCol + ' is "' + colName + '" not "Examples" — aborted');
  }

  // Delete Examples column
  ws.deleteColumn(examplesCol);

  // Update Scenario column note (col 11) with Outline guidance
  updateTcScenarioNote_(ws);

  // Fix title row merge after column deletion
  fixTitleMerge_(ws);

  // Fix TC_Execution TestLevel sync formula
  // After deletion: TestLevel shifts from col O (15) to col N (14)
  var execWs = src.getSheetByName('TC_Execution');
  if (execWs) fixTcExecutionSync_(execWs);

  return 'TC Examples removed (was col ' + examplesCol + ')';
}


// ── API_Master: remove Examples column ───────────────────────────────────
function removeApiMasterExamples_(src) {
  var ws = src.getSheetByName('API_Master');
  if (!ws) return 'skipped (no API_Master)';

  var lastCol = ws.getLastColumn();
  var hdr     = ws.getRange(2, 1, 1, lastCol).getValues()[0];

  var examplesCol = -1;
  hdr.forEach(function(h, i) {
    if (String(h).trim() === 'Examples') examplesCol = i + 1;
  });

  if (examplesCol === -1) {
    updateApiScenarioNote_(ws);
    return 'skipped';
  }

  var colName = String(hdr[examplesCol - 1]).trim();
  if (colName !== 'Examples') {
    throw new Error('API_Master col ' + examplesCol + ' is "' + colName + '" not "Examples" — aborted');
  }

  // Delete Examples column
  ws.deleteColumn(examplesCol);

  // Update Scenario col note (col 13)
  updateApiScenarioNote_(ws);

  // Fix title row merge
  fixTitleMerge_(ws);

  // Fix API_Execution TestLevel sync formula
  // After deletion: TestLevel shifts from col P (16) to col O (15)
  var execWs = src.getSheetByName('API_Execution');
  if (execWs) fixApiExecutionSync_(execWs);

  return 'API Examples removed (was col ' + examplesCol + ')';
}


// ── Update TC_Master Scenario note (col 11) ───────────────────────────────
function updateTcScenarioNote_(ws) {
  ws.getRange(2, 11).setNote(
    'SCENARIO NAMING STANDARD\n\n' +
    'Happy Path : [Role] Successfully [Verb] [Object]\n' +
    'Negative   : [Role] Failed to [Verb] [Object] with [Condition]\n\n' +
    'Rules:\n' +
    '  • Role, Object → Title Case  (Nutritionist, Meal Plan)\n' +
    '  • Verb → active  (Create, Pick Up, Confirm, Submit)\n' +
    '  • Do not use: success/succeed, do, perform, process\n\n' +
    'Standard examples:\n' +
    '  Nutritionist Successfully Creates Meal Plan\n' +
    '  Admin Failed to Delete User with Invalid ID\n\n' +
    '───────────────────────────────────────\n' +
    'SCENARIO OUTLINE (write here in this column)\n' +
    'Use when same steps + same outcome type, different data.\n' +
    'Rule: all Examples = Positive only OR Negative only.\n\n' +
    'Negative example:\n' +
    'User Failed to Log In with <invalid_credential>\n' +
    'Examples:\n' +
    '- Invalid password\n' +
    '- Empty password\n' +
    '- Expired session\n\n' +
    'Positive example:\n' +
    'Admin Successfully Creates User with Role <role>\n' +
    'Examples:\n' +
    '- Viewer\n' +
    '- Operator\n' +
    '- Supervisor'
  );
}


// ── Update API_Master Scenario note (col 13) ─────────────────────────────
function updateApiScenarioNote_(ws) {
  ws.getRange(2, 13).setNote(
    'SCENARIO NAMING STANDARD\n\n' +
    'Happy Path : [Role] Successfully [Verb] [Object] -- [status]\n' +
    'Negative   : [Role] Failed to [Verb] [Object] with [Condition] -- [status]\n\n' +
    'Standard examples:\n' +
    '  User Successfully Creates Meal Plan -- 201\n' +
    '  User Failed to Authenticate with Invalid Token -- 401\n\n' +
    '───────────────────────────────────────\n' +
    'SCENARIO OUTLINE (write here in this column)\n' +
    'Use when same steps + same outcome type, different data.\n' +
    'Rule: all Examples = Positive only OR Negative only.\n\n' +
    'Negative example (all 401):\n' +
    'User Failed to Authenticate with <auth_condition> -- 401\n' +
    'Examples:\n' +
    '- Invalid token\n' +
    '- Empty token\n' +
    '- Expired token\n\n' +
    'Positive example (all 201):\n' +
    'Admin Successfully Creates User with Role <role> -- 201\n' +
    'Examples:\n' +
    '- Viewer\n' +
    '- Operator\n' +
    '- Supervisor\n\n' +
    'Do NOT mix 401 and 403 in one Outline.'
  );
}


// ── Fix title row merge after column deletion ─────────────────────────────
function fixTitleMerge_(ws) {
  try {
    var newCols = ws.getLastColumn();
    var mrs = ws.getRange(1, 1, 1, newCols + 1).getMergedRanges();
    mrs.forEach(function(mr) { if (mr.getRow() === 1) mr.breakApart(); });
    ws.getRange(1, 1, 1, newCols).merge();
  } catch(e) {
    Logger.log('fixTitleMerge (non-fatal): ' + e.message);
  }
}


// ── Fix TC_Execution TestLevel sync: col O (15) → col N (14) ─────────────
function fixTcExecutionSync_(execWs) {
  var DS = 9, MR = 1000;
  var syncCell = execWs.getRange(DS, 7); // col G = TestLevel sync
  var formula  = syncCell.getFormula();
  // Was syncing from TC_Master!O (col 15 = TestLevel after Examples existed)
  // After Examples removed: TestLevel is at TC_Master!N (col 14)
  if (formula.indexOf('TC_Master!O') > -1) {
    syncCell.setFormula(
      formula.replace(/TC_Master!O(\d+):O(\d+)/g, 'TC_Master!N$1:N$2')
             .replace(/TC_Master!O(\d+)/g, 'TC_Master!N$1')
    );
    Logger.log('TC_Execution TestLevel sync: O → N');
  } else {
    Logger.log('TC_Execution sync already OK: ' + formula.substring(0, 50));
  }
}


// ── Fix API_Execution TestLevel sync: col P (16) → col O (15) ────────────
function fixApiExecutionSync_(execWs) {
  var DS = 9, MR = 1000;
  var syncCell = execWs.getRange(DS, 6); // col F = TestLevel sync
  var formula  = syncCell.getFormula();
  // Was syncing from API_Master!P (col 16 = TestLevel after Examples existed)
  // After Examples removed: TestLevel is at API_Master!O (col 15)
  if (formula.indexOf('API_Master!P') > -1) {
    syncCell.setFormula(
      formula.replace(/API_Master!P(\d+):P(\d+)/g, 'API_Master!O$1:O$2')
             .replace(/API_Master!P(\d+)/g, 'API_Master!O$1')
    );
    Logger.log('API_Execution TestLevel sync: P → O');
  } else {
    Logger.log('API_Execution sync already OK: ' + formula.substring(0, 50));
  }
}


// ══════════════════════════════════════════════════════════════════════════
// SINGLE SHEET TEST (untuk debugging satu modul)
// ══════════════════════════════════════════════════════════════════════════
function fixSingleSheet() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Fix Single Sheet',
    'Masukkan Spreadsheet ID modul:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var id = resp.getResponseText().trim();
  if (!id) { ui.alert('ID tidak boleh kosong.'); return; }

  try {
    var result = removeExamplesFromSheet_(id);
    ui.alert(result === 'skipped'
      ? '⏭  Skip — kolom Examples tidak ditemukan di TC_Master maupun API_Master.'
      : '✅  Berhasil: ' + result);
  } catch(e) {
    ui.alert('❌  Error: ' + e.message);
  }
}


// ══════════════════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════════════════
function detectIdColFix_(headerRow) {
  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i]).trim().toUpperCase();
    if (h === 'SPREADSHEET ID' || h === 'SPREADSHEET_ID') return i;
  }
  return 6; // fallback col G (after Active, Project, Module, SubModule, PIC, QA Team Lead)
}

function safeAlertFix_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch(e) { try { Browser.msgBox(msg); } catch(e2) { Logger.log(msg); } }
}