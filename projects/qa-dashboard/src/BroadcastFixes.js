/**
 * BroadcastFixes.js - Consolidated broadcast fixes for QA Dashboard
 *
 * Semua broadcast fixes dalam 1 file untuk kemudahan maintenance.
 * Run fixes sesuai kebutuhan dari menu atau manual.
 */

// ═══════════════════════════════════════════════════════════════════════
// EXTERNAL QA REPORT TAB
// ═══════════════════════════════════════════════════════════════════════

function broadcastExternalTestReportTab() {
  const dashboardSs = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const modules = getModuleList_(dashboardSs);
  const targets = [];
  const seen = {};

  modules.forEach(mod => {
    if (!mod || !mod.id || seen[mod.id]) return;
    seen[mod.id] = true;
    targets.push(mod);
  });

  if (targets.length === 0) {
    ui.alert('No active QATM targets found in Config.');
    return;
  }

  const response = ui.alert(
    'Broadcast External Test Report Tab',
    'Akan membuat tab "External Test Report" di ' + targets.length + ' QATM aktif.\n\n' +
    'Aman untuk existing data:\n' +
    '- Jika tab sudah ada, akan di-skip.\n' +
    '- Tidak rebuild QATM.\n' +
    '- Tidak mengubah TC/API/Summary/BugReport.\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let created = 0;
  let skipped = 0;
  let failed = 0;
  const errors = [];

  targets.forEach(mod => {
    try {
      const qatmSs = SpreadsheetApp.openById(mod.id);
      if (qatmSs.getSheetByName('External Test Report')) {
        skipped++;
        return;
      }
      createExternalTestReportTab_(qatmSs);
      created++;
    } catch (error) {
      failed++;
      errors.push((mod.project || '') + ' / ' + (mod.module || '') + ' / ' + (mod.submodule || mod.name || '') + ': ' + error.message);
      Logger.log('External Test Report broadcast failed [' + (mod.id || '-') + ']: ' + error.stack);
    }
  });

  ui.alert(
    'External Test Report Broadcast Complete',
    'Created: ' + created + '\n' +
    'Skipped existing: ' + skipped + '\n' +
    'Failed: ' + failed +
    (errors.length ? '\n\nErrors:\n' + errors.slice(0, 8).join('\n') : ''),
    ui.ButtonSet.OK
  );
}

function createExternalTestReportTab_(ss) {
  const ws = ss.insertSheet('External Test Report');
  ws.clear();
  ws.setTabColor('#455A64');

  [190,360,160,160,240,180].forEach((width, index) => ws.setColumnWidth(index + 1, width));
  ws.getRange(1,1,1,6).merge()
      .setValue('EXTERNAL TEST REPORT')
      .setBackground('#263238').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,30);

  ws.getRange(2,1,1,6).merge()
      .setValue('Manual evidence untuk scope yang dites oleh external team. Dashboard/PDF membaca field ini jika External QA aktif.')
      .setBackground('#ECEFF1').setFontColor('#455A64').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(2,18);

  const statusList = ['Not Started','In Review','Approved','Rejected','Not Applicable'];
  const overallList = ['Not Started','In Review','Ready for Closure','Approved','Rejected','Not Applicable'];
  const fields = [
    ['External Team / Vendor:', '', null],
    ['Status Review:', 'Not Started', statusList],
    ['Functional Evidence URL:', '', null],
    ['Functional Review Status:', 'Not Started', statusList],
    ['Performance Evidence URL:', '', null],
    ['Performance Review Status:', 'Not Started', statusList],
    ['VAPT Evidence URL:', '', null],
    ['VAPT Review Status:', 'Not Started', statusList],
    ['Overall Status:', 'Not Started', overallList],
    ['Reviewer:', '', null],
    ['Review Date:', '', null],
    ['Notes:', '', null],
  ];

  fields.forEach(([label, value, list], index) => {
    const row = 4 + index;
    ws.getRange(row,1)
      .setValue(label)
      .setBackground('#CFD8DC').setFontColor('#263238').setFontWeight('bold')
      .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('right')
      .setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(row,2,1,5).merge();
    ws.getRange(row,2)
      .setValue(value)
      .setBackground('#FFFFFF').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
      .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    if (list) {
      ws.getRange(row,2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build());
      ws.getRange(row,2,1,5).setBorder(true,true,true,true,false,false,'#1976D2',SpreadsheetApp.BorderStyle.SOLID);
    }
    ws.setRowHeight(row, row === 15 ? 58 : 24);
  });

  const rules = [];
  [
    ['Approved','#C8E6C9','#1B5E20'],
    ['Ready for Closure','#C8E6C9','#1B5E20'],
    ['In Review','#E3F2FD','#1565C0'],
    ['Rejected','#FFCDD2','#B71C1C'],
    ['Not Started','#F5F5F5','#616161'],
    ['Not Applicable','#ECEFF1','#455A64'],
  ].forEach(([value, bg, fg]) => {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(value).setBackground(bg).setFontColor(fg).setBold(true)
        .setRanges([ws.getRange('B4:B15')]).build());
  });
  ws.setConditionalFormatRules(rules);
  ws.setFrozenRows(2);
}

// ═══════════════════════════════════════════════════════════════════════
// V1: ACTIVE FILTER FIX & DASHBOARD IMPROVEMENTS
// ═══════════════════════════════════════════════════════════════════════

/**
 * V1 Broadcast Fix - Apply all improvements
 * - Fix Active filter bug (getMaxRows vs getLastRow)
 * - Add PROD BUGS column to Overview
 * - Add comprehensive notes to headers
 * - Fix Smoke Test percentages
 * - Update formulas
 */
function applyV1BroadcastFix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📢 V1: Dashboard Improvements',
    'Fix ini akan apply:\n\n' +
    '✅ Fix Active=FALSE module filtering\n' +
    '✅ Add PROD BUGS column\n' +
    '✅ Add comprehensive header notes\n' +
    '✅ Fix Smoke Test percentage display\n' +
    '✅ Update Dashboard formulas\n\n' +
    'Dashboard akan di-rebuild. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Fix dibatalkan.');
    return;
  }

  try {
    Logger.log('🔧 Applying V1 Broadcast Fix...');

    // Rebuild Overview with improvements
    buildOverview(ss);

    // Rebuild other tabs
    buildSmoke(ss);
    buildBlockers(ss);
    buildCoverage(ss);

    // Notes already added by init*Headers_() functions
    // addNotesToDashboard(); // REMOVED - function deleted during cleanup

    // Refresh data dengan fixed logic
    refreshDashboard();

    ui.alert(
      '✅ V1 Fix Applied!',
      'Dashboard berhasil di-update dengan improvements:\n\n' +
      '✅ Active filter fixed\n' +
      '✅ PROD BUGS column added\n' +
      '✅ Header notes added\n' +
      '✅ Smoke percentages fixed\n\n' +
      '💡 Refresh data sudah dijalankan.',
      ui.ButtonSet.OK
    );

    Logger.log('✅ V1 Broadcast Fix applied successfully');

  } catch (e) {
    Logger.log('❌ Error applying V1 fix: ' + e.message);
    ui.alert(
      '❌ Error',
      'Gagal apply fix:\n' + e.message + '\n\n' +
      'Check Executions log untuk detail.',
      ui.ButtonSet.OK
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════
// V2: PER-MODULE NOTIFICATION SUPPORT
// ═══════════════════════════════════════════════════════════════════════

/**
 * V2 Broadcast Fix - Per-Module Notification Support
 * - Support multiple webhook URLs (per module)
 * - Enable Notifikasi per module (kolom N)
 * - Enable Email per module (kolom P)
 * - Module dengan FALSE tidak akan kirim notif
 */
function applyV2PerModuleNotifFix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📢 V2: Per-Module Notification Support',
    'Fix ini akan mengupdate notification system untuk support:\n\n' +
    '✅ Multiple webhook URLs (per module)\n' +
    '✅ Enable Notifikasi per module (kolom N)\n' +
    '✅ Enable Email per module (kolom P)\n' +
    '✅ Module dengan FALSE tidak akan kirim notif\n\n' +
    '⚠️ PENTING:\n' +
    '• Webhook dan enable flags dibaca dari Config SETIAP ROW\n' +
    '• Pastikan kolom L-P sudah terisi di Config untuk setiap module\n\n' +
    'Lanjutkan apply fix?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Fix dibatalkan.');
    return;
  }

  try {
    Logger.log('🔧 Applying V2 Per-Module Notification fix...');

    // Verify Config structure
    const cfg = ss.getSheetByName('Config');
    if (!cfg) {
      throw new Error('Config sheet not found');
    }

    // Check if columns L-P exist
    const headers = cfg.getRange(3, 1, 1, 20).getValues()[0];
    let missingCols = [];
    if (!String(headers[11]).includes('Webhook')) missingCols.push('L (Webhook)');
    if (!String(headers[13]).includes('Notif')) missingCols.push('N (Enable Notif)');
    if (!String(headers[15]).includes('Email')) missingCols.push('P (Enable Email)');

    if (missingCols.length > 0) {
      ui.alert(
        '⚠️ Config Structure Issue',
        'Kolom notification tidak ditemukan di Config:\n' +
        missingCols.join(', ') + '\n\n' +
        'Pastikan Config sudah di-rebuild dengan createDashboard() terbaru.',
        ui.ButtonSet.OK
      );
      return;
    }

    // Count modules with notifications enabled
    const cfgData = cfg.getDataRange().getValues();
    let chatEnabledCount = 0;
    let emailEnabledCount = 0;
    let uniqueWebhooks = new Set();

    for (let i = 3; i < cfgData.length; i++) {
      const project = String(cfgData[i][2]).trim();
      const modul = String(cfgData[i][3]).trim();
      if (!project || !modul) continue;

      const webhook = String(cfgData[i][11]).trim();
      const chatEnabled = cfgData[i][13] === true;
      const emailEnabled = cfgData[i][15] === true;

      if (chatEnabled) {
        chatEnabledCount++;
        if (webhook && webhook.includes('chat.googleapis.com')) {
          uniqueWebhooks.add(webhook);
        }
      }
      if (emailEnabled) emailEnabledCount++;
    }

    ui.alert(
      '✅ V2 Per-Module Notification Applied!',
      'Notification system sekarang support per-module config.\n\n' +
      '📊 Summary:\n' +
      '• Modules dengan Chat enabled: ' + chatEnabledCount + '\n' +
      '• Modules dengan Email enabled: ' + emailEnabledCount + '\n' +
      '• Unique webhook URLs: ' + uniqueWebhooks.size + '\n\n' +
      '🔍 How it works:\n' +
      '• Setiap module akan kirim notifikasi terpisah\n' +
      '• Menggunakan webhook dari kolom L di row module tersebut\n' +
      '• Hanya kirim jika Enable Notifikasi/Email = TRUE\n\n' +
      '💡 Test notification: Menu > Notifications > Test Notification Now',
      ui.ButtonSet.OK
    );

    Logger.log('✅ V2 Per-Module Notification fix applied successfully');
    Logger.log('   Chat enabled modules: ' + chatEnabledCount);
    Logger.log('   Email enabled modules: ' + emailEnabledCount);
    Logger.log('   Unique webhooks: ' + uniqueWebhooks.size);

  } catch (e) {
    Logger.log('❌ Error applying V2 fix: ' + e.message);
    ui.alert(
      '❌ Error',
      'Gagal apply fix:\n' + e.message + '\n\n' +
      'Check Executions log untuk detail.',
      ui.ButtonSet.OK
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════
// CONFIG FIXES
// ═══════════════════════════════════════════════════════════════════════

/**
 * Fix Config - Add missing columns or repair structure
 */
function fixConfigStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🔧 Fix Config Structure',
    'Akan memperbaiki struktur Config tab:\n\n' +
    '• Add missing notification columns (L-P)\n' +
    '• Add missing Jira columns\n' +
    '• Fix data validation\n' +
    '• Fix conditional formatting\n\n' +
    'Config akan di-rebuild. Data existing akan preserved.\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Dibatalkan.');
    return;
  }

  try {
    const cfg = ss.getSheetByName('Config');
    if (!cfg) {
      buildConfig(ss);
      ui.alert('✅ Config created!', 'Config tab berhasil dibuat.', ui.ButtonSet.OK);
      return;
    }

    // Backup data
    const data = cfg.getDataRange().getValues();
    const moduleData = [];

    for (let i = 3; i < data.length; i++) {
      if (data[i][2] && data[i][3]) { // Has Project and Modul
        moduleData.push(data[i]);
      }
    }

    // Rebuild Config
    ss.deleteSheet(cfg);
    buildConfig(ss);

    // Restore data
    const newCfg = ss.getSheetByName('Config');
    if (moduleData.length > 0) {
      const numCols = Math.min(moduleData[0].length, 20); // Max 20 cols
      moduleData.forEach((row, i) => {
        newCfg.getRange(4 + i, 1, 1, numCols).setValues([row.slice(0, numCols)]);
      });
    }

    ui.alert(
      '✅ Config Fixed!',
      'Config structure berhasil diperbaiki.\n\n' +
      'Data modules: ' + moduleData.length + ' rows preserved.\n\n' +
      'Silakan cek Config tab dan lengkapi data yang missing.',
      ui.ButtonSet.OK
    );

    Logger.log('✅ Config structure fixed - ' + moduleData.length + ' modules preserved');

  } catch (e) {
    Logger.log('❌ Error fixing Config: ' + e.message);
    ui.alert('❌ Error', 'Gagal fix Config:\n' + e.message, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Apply ALL fixes at once
 */
function applyAllFixes() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🚀 Apply All Fixes',
    'Akan apply semua fixes:\n\n' +
    '1. V1: Dashboard Improvements\n' +
    '2. V2: Per-Module Notifications\n' +
    '3. Config Structure Fix\n\n' +
    'Proses ini akan:\n' +
    '• Rebuild semua tabs\n' +
    '• Update notification system\n' +
    '• Fix Config structure\n' +
    '• Refresh all data\n\n' +
    '⚠️ Estimasi waktu: 2-5 menit\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Dibatalkan.');
    return;
  }

  try {
    Logger.log('🚀 Starting Apply All Fixes...');

    // Step 1: Fix Config
    Logger.log('Step 1/3: Fixing Config structure...');
    fixConfigStructure();

    // Step 2: V1 Improvements
    Logger.log('Step 2/3: Applying V1 improvements...');
    applyV1BroadcastFix();

    // Step 3: V2 Notifications
    Logger.log('Step 3/3: Applying V2 notification support...');
    applyV2PerModuleNotifFix();

    ui.alert(
      '✅ All Fixes Applied!',
      'Semua fixes berhasil dijalankan:\n\n' +
      '✅ Config structure fixed\n' +
      '✅ V1 Dashboard improvements applied\n' +
      '✅ V2 Per-module notifications enabled\n\n' +
      '💡 Dashboard siap digunakan!\n' +
      'Test notification: Menu > Notifications > Test',
      ui.ButtonSet.OK
    );

    Logger.log('✅ All fixes completed successfully');

  } catch (e) {
    Logger.log('❌ Error during Apply All: ' + e.message);
    ui.alert(
      '❌ Error',
      'Error saat menjalankan fixes:\n' + e.message + '\n\n' +
      'Beberapa fixes mungkin sudah applied.\n' +
      'Check Executions log untuk detail.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Apply comprehensive notes to all dashboard tabs
 */
function applyComprehensiveNotes() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📝 Add Comprehensive Notes',
    'Akan menambahkan comprehensive notes ke SEMUA tabs:\n\n' +
    '✅ Overview - Module info & metrics\n' +
    '✅ Smoke - Smoke test explanations\n' +
    '✅ Blockers - Blocker & PROD bugs guide\n' +
    '✅ Coverage - Coverage metrics guide\n' +
    '✅ Config - Configuration help\n\n' +
    '💡 Notes akan muncul sebagai tooltips saat hover di headers.\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Dibatalkan.');
    return;
  }

  try {
    Logger.log('📝 Notes already added by init*Headers_() functions during tab creation');

    // addNotesToDashboard(); // REMOVED - function deleted during cleanup

    Logger.log('✅ Notes are already present in all tabs');

  } catch (e) {
    Logger.log('❌ Error in notes section: ' + e.message);
    ui.alert(
      '❌ Error',
      'Error in notes section:\n' + e.message + '\n\n' +
      'Check Executions log untuk detail.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Show available fixes
 */
function showAvailableFixes() {
  const ui = SpreadsheetApp.getUi();

  ui.alert(
    '📋 Available Broadcast Fixes',
    'Broadcast fixes yang tersedia:\n\n' +
    '1️⃣ applyV1BroadcastFix()\n' +
    '   - Fix Active filter bug\n' +
    '   - Add PROD BUGS column\n' +
    '   - Add header notes\n' +
    '   - Fix Smoke percentages\n\n' +
    '2️⃣ applyV2PerModuleNotifFix()\n' +
    '   - Per-module webhook URLs\n' +
    '   - Per-module enable flags\n' +
    '   - Support multiple webhooks\n\n' +
    '3️⃣ fixConfigStructure()\n' +
    '   - Fix Config tab structure\n' +
    '   - Add missing columns\n\n' +
    '📝 applyComprehensiveNotes()\n' +
    '   - Add comprehensive tooltips ke ALL tabs\n' +
    '   - Overview, Smoke, Blockers, Coverage, Config\n\n' +
    '🚀 applyAllFixes()\n' +
    '   - Apply semua fixes sekaligus\n\n' +
    '4️⃣ broadcastManualSyncButton()\n' +
    '   - Add "🔄 Sync Jira" button ke tiap QATM\n' +
    '   - Manual sync per-QATM\n\n' +
    'Run dari Extensions > Apps Script atau Script Editor.',
    ui.ButtonSet.OK
  );
}

// ═══════════════════════════════════════════════════════════════════════
// V4: MANUAL JIRA SYNC - REMOTE SYNC FROM DASHBOARD
// ═══════════════════════════════════════════════════════════════════════

/**
 * Sync Single QATM from Dashboard
 * Call this function to manually sync a specific QATM
 */
function syncSingleQATMFromDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Get list of active modules
  const cfg = ss.getSheetByName('Config');
  if (!cfg) {
    ui.alert('❌ Config tab tidak ditemukan!');
    return;
  }

  const cfgData = cfg.getDataRange().getValues();
  const modules = [];

  for (let i = 3; i < cfgData.length; i++) {
    const active = cfgData[i][0] === true;
    const jiraSync = cfgData[i][1] === true;
    const project = String(cfgData[i][2]).trim();
    const modul = String(cfgData[i][3]).trim();
    const qatmId = String(cfgData[i][6]).trim();

    if (active && jiraSync && qatmId && qatmId.length > 10) {
      modules.push({
        label: project + ' - ' + modul,
        qatmId: qatmId,
        index: i
      });
    }
  }

  if (modules.length === 0) {
    ui.alert('❌ Tidak ada modul dengan Jira Sync aktif.');
    return;
  }

  // Show selection dialog
  let msg = 'Pilih QATM untuk di-sync:\n\n';
  modules.forEach((m, idx) => {
    msg += (idx + 1) + '. ' + m.label + '\n';
  });
  msg += '\nMasukkan nomor (1-' + modules.length + '):';

  const response = ui.prompt('Sync Single QATM', msg, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const selection = parseInt(response.getResponseText());
  if (isNaN(selection) || selection < 1 || selection > modules.length) {
    ui.alert('❌ Nomor tidak valid!');
    return;
  }

  const selectedModule = modules[selection - 1];
  const rowIndex = selectedModule.index;

  // Get credentials
  const dashboardSs = SpreadsheetApp.getActiveSpreadsheet();
  const credSheet = dashboardSs.getSheetByName('Credentials');
  const jiraInstance = String(cfgData[rowIndex][8]).trim().toLowerCase();
  const jiraProjectKey = String(cfgData[rowIndex][9]).trim().toUpperCase();
  const moduleName = String(cfgData[rowIndex][3]).trim();

  let jiraEmail = '';
  let jiraApiToken = '';

  if (credSheet) {
    const credData = credSheet.getDataRange().getValues();
    for (let j = 3; j < credData.length; j++) {
      const inst = String(credData[j][0]).trim().toLowerCase();
      if (inst === jiraInstance) {
        jiraEmail = String(credData[j][1]).trim();
        jiraApiToken = String(credData[j][2]).trim();
        break;
      }
    }
  }

  if (!jiraEmail || !jiraApiToken) {
    ui.alert('❌ Credentials tidak ditemukan untuk instance: ' + jiraInstance);
    return;
  }

  ui.alert(
    'Sync QATM: ' + selectedModule.label,
    'Memulai sync...\n\nIni akan memakan waktu beberapa menit.',
    ui.ButtonSet.OK
  );

  // Perform sync
  try {
    const result = syncQATMRemote_(selectedModule.qatmId, jiraInstance, jiraProjectKey, jiraEmail, jiraApiToken, moduleName);
    ui.alert('✅ Sync Selesai!', result, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Error', 'Sync gagal:\n\n' + e.message, ui.ButtonSet.OK);
    Logger.log('❌ Sync error: ' + e.message);
  }
}

/**
 * Sync All Active QATMs from Dashboard
 */
function syncAllQATMsFromDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Sync All QATMs',
    'Ini akan sync SEMUA QATM yang aktif dengan Jira Sync enabled.\n\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const cfg = ss.getSheetByName('Config');
  if (!cfg) {
    ui.alert('❌ Config tab tidak ditemukan!');
    return;
  }

  const cfgData = cfg.getDataRange().getValues();
  const dashboardSs = SpreadsheetApp.getActiveSpreadsheet();
  const credSheet = dashboardSs.getSheetByName('Credentials');

  let successCount = 0;
  let errorCount = 0;
  const errors = [];

  for (let i = 3; i < cfgData.length; i++) {
    const active = cfgData[i][0] === true;
    const jiraSync = cfgData[i][1] === true;
    const project = String(cfgData[i][2]).trim();
    const modul = String(cfgData[i][3]).trim();
    const qatmId = String(cfgData[i][6]).trim();
    const jiraInstance = String(cfgData[i][8]).trim().toLowerCase();
    const jiraProjectKey = String(cfgData[i][9]).trim().toUpperCase();

    if (!active || !jiraSync || !qatmId || qatmId.length < 10) continue;

    let jiraEmail = '';
    let jiraApiToken = '';

    if (credSheet) {
      const credData = credSheet.getDataRange().getValues();
      for (let j = 3; j < credData.length; j++) {
        const inst = String(credData[j][0]).trim().toLowerCase();
        if (inst === jiraInstance) {
          jiraEmail = String(credData[j][1]).trim();
          jiraApiToken = String(credData[j][2]).trim();
          break;
        }
      }
    }

    if (!jiraEmail || !jiraApiToken) {
      errorCount++;
      errors.push(project + ' - ' + modul + ' (missing credentials)');
      continue;
    }

    try {
      Logger.log('Syncing: ' + project + ' - ' + modul);
      syncQATMRemote_(qatmId, jiraInstance, jiraProjectKey, jiraEmail, jiraApiToken, modul);
      successCount++;
      Logger.log('✅ Synced: ' + project + ' - ' + modul);
    } catch (e) {
      errorCount++;
      errors.push(project + ' - ' + modul + ' (' + e.message + ')');
      Logger.log('❌ Error: ' + project + ' - ' + modul + ': ' + e.message);
    }
  }

  let msg = '✅ Sync All QATMs Selesai!\n\n';
  msg += '📊 Summary:\n';
  msg += '• Success: ' + successCount + ' QATM(s)\n';
  msg += '• Errors: ' + errorCount + ' QATM(s)\n';

  if (errors.length > 0) {
    msg += '\n❌ Errors:\n';
    errors.forEach(function(err) {
      msg += '• ' + err + '\n';
    });
  }

  ui.alert('Sync Complete', msg, ui.ButtonSet.OK);
}

/**
 * Remote sync function - syncs QATM from Dashboard
 */
function syncQATMRemote_(qatmId, jiraInstance, jiraProjectKey, jiraEmail, jiraApiToken, moduleName) {
  const qatmSs = SpreadsheetApp.openById(qatmId);
  const bugSheet = qatmSs.getSheetByName('BugReport');

  if (!bugSheet) {
    throw new Error('BugReport sheet not found');
  }

  Logger.log('Fetching bugs from Jira...');
  const issues = fetchJiraIssues_(jiraInstance, jiraProjectKey, moduleName, jiraEmail, jiraApiToken);

  if (!issues) {
    throw new Error('Failed to fetch issues from Jira');
  }

  Logger.log('Fetched ' + issues.length + ' issue(s)');

  const bugIndex = buildBugIndex_(bugSheet);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const instUrl = getJiraInstanceUrl_(jiraInstance);
  let inserted = 0;
  let updated = 0;

  issues.forEach(function(issue) {
    const statusName = issue.fields.status && issue.fields.status.name;
    const isClosedStatus = statusName === 'Closed' || statusName === "Won't Fix";

    if (isClosedStatus) return;

    if (bugIndex[issue.key] !== undefined) {
      updateBugRow_(bugSheet, bugIndex[issue.key], issue, now, instUrl, jiraInstance);
      updated++;
    } else {
      insertBugRow_(bugSheet, issue, now, instUrl, jiraInstance, moduleName);
      inserted++;
    }
  });

  const deleted = cleanupClosedBugs_(bugSheet, issues);

  let result = '📊 Hasil Sync:\n\n';
  result += '• Inserted: ' + inserted + ' bug baru\n';
  result += '• Updated: ' + updated + ' bug\n';
  if (deleted > 0) {
    result += '• Deleted: ' + deleted + ' (Closed/Won\'t Fix)\n';
  }
  result += '\n✅ Total: ' + (inserted + updated) + ' active bugs';

  return result;
}

/**
 * Broadcast: Add Config Tab to all QATM spreadsheets
 *
 * Creates Config tab with Jira configuration
 * Users can then sync from Dashboard using syncSingleQATMFromDashboard()
 */
function broadcastManualSyncButton() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🔄 Add Manual Sync Button',
    'Fix ini akan:\n\n' +
    '✅ Add button "🔄 Sync Jira" di Config tab (R4)\n' +
    '✅ Button untuk manual sync Jira per-QATM\n' +
    '✅ Auto sync tetap berjalan normal\n\n' +
    'Button akan di-broadcast ke semua active QATM.\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Broadcast dibatalkan.');
    return;
  }

  try {
    const cfg = ss.getSheetByName('Config');
    if (!cfg) {
      ui.alert('❌ Config tab tidak ditemukan!');
      return;
    }

    // Get active modules
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
        Logger.log('Adding sync button to: ' + project + ' - ' + modul);

        const qatmSs = SpreadsheetApp.openById(qatmId);

        // Get Jira config for this module from Dashboard Config
        const jiraSync = cfgData[i][1] === true;  // Col B
        const jiraInstance = String(cfgData[i][8]).trim().toLowerCase();  // Col I
        const jiraProjectKey = String(cfgData[i][9]).trim().toUpperCase();  // Col J

        if (!jiraSync || !jiraInstance || !jiraProjectKey) {
          Logger.log('⚠️ Jira Sync not configured for ' + project + ' - ' + modul);
          errorCount++;
          errors.push(project + ' - ' + modul + ' (Jira Sync not configured)');
          continue;
        }

        // Get credentials from Dashboard
        const dashboardSs = SpreadsheetApp.getActiveSpreadsheet();
        const credSheet = dashboardSs.getSheetByName('Credentials');
        let jiraEmail = '';
        let jiraApiToken = '';

        if (credSheet) {
          const credData = credSheet.getDataRange().getValues();
          for (let j = 3; j < credData.length; j++) {
            const inst = String(credData[j][0]).trim().toLowerCase();
            if (inst === jiraInstance) {
              jiraEmail = String(credData[j][1]).trim();
              jiraApiToken = String(credData[j][2]).trim();
              break;
            }
          }
        }

        if (!jiraEmail || !jiraApiToken) {
          Logger.log('⚠️ Jira credentials not found for instance: ' + jiraInstance);
          errorCount++;
          errors.push(project + ' - ' + modul + ' (missing credentials for ' + jiraInstance + ')');
          continue;
        }

        addSyncButtonToQATM_(qatmSs, project, modul, jiraInstance, jiraProjectKey, jiraEmail, jiraApiToken);
        successCount++;
        Logger.log('✅ Config added to ' + project + ' - ' + modul);

      } catch (e) {
        Logger.log('❌ Error for ' + project + ' - ' + modul + ': ' + e.message);
        errorCount++;
        errors.push(project + ' - ' + modul + ' (' + e.message + ')');
      }
    }

    let msg = '✅ Broadcast Manual Sync Button Complete!\n\n';
    msg += '📊 Summary:\n';
    msg += '• Success: ' + successCount + ' QATM(s)\n';
    msg += '• Errors: ' + errorCount + ' QATM(s)\n';

    if (errors.length > 0) {
      msg += '\n❌ Errors:\n';
      errors.forEach(err => {
        msg += '• ' + err + '\n';
      });
    }

    ui.alert('Broadcast Complete', msg, ui.ButtonSet.OK);
    Logger.log('✅ Broadcast complete - Success: ' + successCount + ', Errors: ' + errorCount);

  } catch (e) {
    ui.alert('❌ Error', 'Error during broadcast: ' + e.message, ui.ButtonSet.OK);
    Logger.log('❌ Broadcast error: ' + e.message);
  }
}

/**
 * Add Sync Button and Config to a single QATM
 * Creates Config tab if it doesn't exist
 * Populates Jira configuration
 * Adds visual sync button indicator
 */
function addSyncButtonToQATM_(qatmSs, project, modul, jiraInstance, jiraProjectKey, jiraEmail, jiraApiToken) {
  // Get or create Config sheet
  let qatmConfig = qatmSs.getSheetByName('Config');

  if (!qatmConfig) {
    Logger.log('Creating Config tab...');
    qatmConfig = qatmSs.insertSheet('Config');
  }

  // Setup Config tab structure
  // Row 1: Title
  qatmConfig.getRange('A1').setValue('JIRA SYNC CONFIGURATION')
           .setFontSize(14)
           .setFontWeight('bold')
           .setBackground('#0D47A1')
           .setFontColor('#FFFFFF');
  qatmConfig.getRange('A1:F1').merge();

  // Row 2: Info
  qatmConfig.getRange('A2').setValue('⚠️ This configuration is auto-generated from Dashboard. DO NOT edit manually.')
           .setFontSize(9)
           .setFontStyle('italic')
           .setBackground('#FFF3E0')
           .setFontColor('#E65100');
  qatmConfig.getRange('A2:F2').merge();

  // Row 3: Labels
  const labels = [
    ['Label', 'Jira Instance', 'Project Key', 'Email', 'API Token', 'Module'],
  ];
  qatmConfig.getRange('A3:F3').setValues(labels)
           .setFontWeight('bold')
           .setBackground('#E3F2FD')
           .setFontColor('#0D47A1')
           .setHorizontalAlignment('center');

  // Row 4: Values
  const values = [
    ['Config', jiraInstance, jiraProjectKey, jiraEmail, jiraApiToken, modul],
  ];
  qatmConfig.getRange('A4:F4').setValues(values);

  // Format value cells
  qatmConfig.getRange('B4:F4')
           .setBackground('#FFFFFF')
           .setBorder(true, true, true, true, false, false, '#1976D2', SpreadsheetApp.BorderStyle.SOLID);

  // Set column widths
  qatmConfig.setColumnWidth(1, 80);   // A
  qatmConfig.setColumnWidth(2, 120);  // B - Instance
  qatmConfig.setColumnWidth(3, 100);  // C - Project Key
  qatmConfig.setColumnWidth(4, 200);  // D - Email
  qatmConfig.setColumnWidth(5, 300);  // E - API Token
  qatmConfig.setColumnWidth(6, 150);  // F - Module

  // Hide API Token column for security (user can unhide if needed)
  qatmConfig.hideColumns(5);

  // Add instructions label
  qatmConfig.getRange('H3').setValue('Manual Sync:')
                           .setFontWeight('bold')
                           .setFontSize(9)
                           .setHorizontalAlignment('center')
                           .setBackground('#E3F2FD')
                           .setFontColor('#0D47A1');

  qatmConfig.setColumnWidth(8, 150);  // H

  // Create visual button indicator in H4
  const buttonCell = qatmConfig.getRange('H4');
  buttonCell.setValue('Use Menu Above ↑\n🔄 Jira Sync');
  buttonCell.setFontWeight('bold')
           .setFontSize(9)
           .setWrap(true)
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle')
           .setBackground('#4285F4')
           .setFontColor('#FFFFFF')
           .setBorder(true, true, true, true, false, false, '#1967D2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  qatmConfig.setRowHeight(4, 40);

  // Add usage instructions in row 6
  const instructions = 'CARA MANUAL SYNC JIRA:\n\n' +
    '1️⃣ Klik menu di atas: 🔄 Jira Sync > Sync Now\n' +
    '2️⃣ Tunggu beberapa menit (tergantung jumlah bug)\n' +
    '3️⃣ Selesai! Cek tab BugReport untuk hasil sync\n\n' +
    'YANG DILAKUKAN SAAT SYNC:\n' +
    '✅ Fetch bugs dari Jira (project ' + jiraProjectKey + ', module ' + modul + ')\n' +
    '✅ Insert bug baru ke BugReport\n' +
    '✅ Update bug yang sudah ada\n' +
    '✅ Hapus bug Closed/Won\'t Fix\n\n' +
    'AUTO SYNC dari Dashboard tetap berjalan seperti biasa.\n' +
    'Manual sync ini untuk sync on-demand jika diperlukan.';

  qatmConfig.getRange('A6').setValue(instructions)
           .setWrap(true)
           .setFontSize(9)
           .setBackground('#F5F5F5')
           .setVerticalAlignment('top');
  qatmConfig.getRange('A6:H10').merge();
  qatmConfig.setRowHeight(6, 150);

  // Add project info
  qatmConfig.getRange('A12').setValue('Project: ' + project + ' / Module: ' + modul)
           .setFontWeight('bold')
           .setFontSize(10);

  Logger.log('✅ Config tab created/updated with Jira configuration');
}

/**
 * Manual Sync Function - Call this from QATM directly
 * This function should be added to each QATM's Apps Script
 *
 * Usage: Create custom menu in QATM with this function
 *
 * Expected Config tab structure:
 * Row 3: Labels
 * Row 4: Values
 * B4 = Jira Instance (digitalperuri / bgn-peruri)
 * C4 = Jira Project Key (SQA / BGN / etc)
 * D4 = Jira Email
 * E4 = Jira API Token
 * F4 = Module Name/Number
 */
function manualSyncJiraFromQATM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // Get QATM Config
    const config = ss.getSheetByName('Config');
    if (!config) {
      ui.alert('❌ Config tab not found!\n\nPlease create Config tab first via broadcast from Dashboard.');
      return;
    }

    // Read Jira config from QATM Config tab
    // Row 3 = Labels, Row 4 = Values
    const jiraInstance = String(config.getRange('B4').getValue()).trim().toLowerCase();
    const jiraProjectKey = String(config.getRange('C4').getValue()).trim().toUpperCase();
    const jiraEmail = String(config.getRange('D4').getValue()).trim();
    const jiraApiToken = String(config.getRange('E4').getValue()).trim();
    const moduleName = String(config.getRange('F4').getValue()).trim();

    // Validate config
    if (!jiraInstance || !jiraProjectKey || !jiraEmail || !jiraApiToken || !moduleName) {
      let msg = '❌ Jira configuration incomplete!\n\nPlease fill in Config tab (row 4):\n\n';
      if (!jiraInstance) msg += '• B4: Jira Instance (digitalperuri / bgn-peruri)\n';
      if (!jiraProjectKey) msg += '• C4: Jira Project Key (SQA / BGN / etc)\n';
      if (!jiraEmail) msg += '• D4: Jira Email (your.email@company.com)\n';
      if (!jiraApiToken) msg += '• E4: Jira API Token (ATATT3xFf...)\n';
      if (!moduleName) msg += '• F4: Module Name/Number (1 / Portal+SSO / etc)\n';

      ui.alert(msg);
      return;
    }

    ui.alert(
      '🔄 Manual Jira Sync',
      'Starting Jira sync for this QATM...\n\n' +
      'Instance: ' + jiraInstance + '\n' +
      'Project: ' + jiraProjectKey + '\n' +
      'Module: ' + moduleName + '\n\n' +
      'This may take a few minutes.',
      ui.ButtonSet.OK
    );

    // Call the actual sync function
    const result = syncJiraForCurrentQATM_(jiraInstance, jiraProjectKey, jiraApiToken, jiraEmail, moduleName);

    ui.alert(
      '✅ Sync Complete!',
      result,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ Error', 'Sync failed:\n\n' + e.message + '\n\nCheck Execution log for details.', ui.ButtonSet.OK);
    Logger.log('❌ Manual sync error: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
  }
}

/**
 * Helper: Sync Jira for current QATM
 * Reuses logic from JiraSync.js but for single QATM
 * This is the reference implementation - actual code should be deployed to QATM template
 */
function syncJiraForCurrentQATM_(jiraInstance, jiraProjectKey, jiraApiToken, email, moduleName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bugSheet = ss.getSheetByName('BugReport');

  if (!bugSheet) {
    Logger.log('❌ BugReport sheet not found');
    return '❌ BugReport sheet not found in this spreadsheet.';
  }

  Logger.log('══════════════════════════════════════════');
  Logger.log('🔄 MANUAL JIRA SYNC');
  Logger.log('══════════════════════════════════════════');
  Logger.log('Instance: ' + jiraInstance);
  Logger.log('Project:  ' + jiraProjectKey);
  Logger.log('Modul:    ' + (moduleName || '(not specified)'));
  Logger.log('');

  try {
    // Fetch issues from Jira
    const issues = fetchJiraIssues_(jiraInstance, jiraProjectKey, moduleName, email, jiraApiToken);

    if (!issues) {
      Logger.log('❌ Failed to fetch issues from Jira');
      return '❌ Failed to fetch issues from Jira.\n\nCheck:\n• Jira Instance URL\n• Project Key\n• API Token\n• Module name';
    }

    Logger.log('✅ Fetched ' + issues.length + ' issue(s) from Jira');

    // Build index of existing bugs in sheet
    const bugIndex = buildBugIndex_(bugSheet);
    Logger.log('📊 Found ' + Object.keys(bugIndex).length + ' existing bug(s) in sheet');

    // Sync bugs
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const instUrl = getJiraInstanceUrl_(jiraInstance);
    let inserted = 0;
    let updated = 0;

    const jiraKeys = new Set();

    issues.forEach(issue => {
      const statusName = issue.fields.status && issue.fields.status.name;
      const isClosedStatus = statusName === 'Closed' || statusName === "Won't Fix";

      jiraKeys.add(issue.key);

      // Skip closed bugs - they'll be cleaned up below
      if (isClosedStatus) {
        return;
      }

      // Update or insert active bugs
      if (bugIndex[issue.key] !== undefined) {
        updateBugRow_(bugSheet, bugIndex[issue.key], issue, now, instUrl, jiraInstance);
        updated++;
      } else {
        insertBugRow_(bugSheet, issue, now, instUrl, jiraInstance, moduleName);
        inserted++;
      }
    });

    // Cleanup closed bugs
    const deleted = cleanupClosedBugs_(bugSheet, issues);

    Logger.log('✅ Sync complete:');
    Logger.log('  • Inserted: ' + inserted);
    Logger.log('  • Updated: ' + updated);
    Logger.log('  • Deleted (Closed/Won\'t Fix): ' + deleted);
    Logger.log('══════════════════════════════════════════');

    let result = '📊 Sync Results:\n\n';
    result += '• Inserted: ' + inserted + ' new bug(s)\n';
    result += '• Updated: ' + updated + ' bug(s)\n';
    if (deleted > 0) {
      result += '• Deleted: ' + deleted + ' (Closed/Won\'t Fix)\n';
    }
    result += '\n✅ Total: ' + (inserted + updated) + ' active bug(s)';

    return result;

  } catch (e) {
    Logger.log('❌ Error during sync: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    return '❌ Error during sync:\n\n' + e.message;
  }
}

/**
 * Fetch issues from Jira using REST API
 */
function fetchJiraIssues_(instance, projectKey, moduleName, email, apiToken) {
  const baseUrl = getJiraInstanceUrl_(instance);
  const modulField = getJiraModulField_(instance);

  // Build JQL query
  const isNumeric = /^\d+$/.test(moduleName);
  const modulValue = isNumeric ? moduleName : '"' + moduleName + '"';

  const jql = 'project = "' + projectKey + '" AND issuetype = Bug AND "' + modulField + '" = ' + modulValue + ' ORDER BY priority ASC, updated DESC';

  Logger.log('JQL: ' + jql);
  Logger.log('');

  const auth = Utilities.base64Encode(email + ':' + apiToken);
  const headers = {
    'Authorization': 'Basic ' + auth,
    'Content-Type': 'application/json'
  };

  const fields = 'summary,description,priority,status,assignee,reporter,resolutiondate,key,created,updated,labels,components,environment';
  const customFields = getJiraCustomFields_(instance);
  const allFields = fields + customFields;

  const allIssues = [];
  let nextPageToken = null;

  do {
    let url = baseUrl + '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) +
      '&fields=' + encodeURIComponent(allFields) + '&maxResults=100';

    if (nextPageToken) {
      url += '&nextPageToken=' + encodeURIComponent(nextPageToken);
    }

    let response;
    try {
      response = UrlFetchApp.fetch(url, {headers: headers, muteHttpExceptions: true});
    } catch (e) {
      Logger.log('❌ Fetch error: ' + e.message);
      return null;
    }

    if (response.getResponseCode() !== 200) {
      Logger.log('❌ Jira API error ' + response.getResponseCode() + ': ' + response.getContentText().substring(0, 200));
      return null;
    }

    let data;
    try {
      data = JSON.parse(response.getContentText());
    } catch (e) {
      Logger.log('❌ JSON parse error: ' + e.message);
      return null;
    }

    (data.issues || []).forEach(issue => allIssues.push(issue));

    nextPageToken = data.nextPageToken || null;

    if (nextPageToken) {
      Utilities.sleep(300);  // Rate limiting
    }

  } while (nextPageToken);

  return allIssues;
}

/**
 * Build index of existing bugs by Jira key
 */
function buildBugIndex_(bugSheet) {
  const data = bugSheet.getDataRange().getValues();
  const index = {};

  // Start from row 6 (assuming header rows 1-5)
  for (let i = 5; i < data.length; i++) {
    const jiraKey = String(data[i][19]).trim();  // Column T (index 19) = Jira Key
    if (jiraKey && jiraKey !== '') {
      index[jiraKey] = i + 1;  // Store 1-based row number
    }
  }

  return index;
}

/**
 * Update existing bug row
 */
function updateBugRow_(bugSheet, rowNum, issue, timestamp, instUrl, instance) {
  const row = [];

  // Column A: Bug ID (keep existing)
  // Column B: Type
  row[1] = 'Functional';  // Default type

  // Column C: Priority
  row[2] = normalizePriority_(issue.fields.priority && issue.fields.priority.name);

  // Column D: Status
  row[3] = normalizeStatus_(issue.fields.status && issue.fields.status.name);

  // Column E: Feature
  const fieldMap = getJiraFieldMap_(instance);
  row[4] = getCustomFieldValue_(issue, fieldMap.feature);

  // Column F: Submodul
  row[5] = getCustomFieldValue_(issue, fieldMap.submodul);

  // Column G: Title
  row[6] = issue.fields.summary || '';

  // Column H: Description
  row[7] = convertADF_(issue.fields.description);

  // Column I: Environment
  row[8] = getCustomFieldValue_(issue, fieldMap.environment) || issue.fields.environment || '';

  // Column J: Steps
  row[9] = getCustomFieldValue_(issue, fieldMap.steps);

  // Column K: Expected
  row[10] = getCustomFieldValue_(issue, fieldMap.expected);

  // Column L: Actual
  row[11] = getCustomFieldValue_(issue, fieldMap.actual);

  // Column M: Test Case (keep existing)
  // Column N: Reported By
  row[13] = issue.fields.reporter && issue.fields.reporter.displayName || '';

  // Column O: Assigned To
  row[14] = issue.fields.assignee && issue.fields.assignee.displayName || '';

  // Column P: Date Found
  row[15] = issue.fields.created ? issue.fields.created.substring(0, 10) : '';

  // Column Q: Date Fixed (keep existing)
  // Column R: Sprint (keep existing)

  // Column S: Link
  row[18] = instUrl + '/browse/' + issue.key;

  // Column T: Jira Key
  row[19] = issue.key;

  // Column U: Last Synced
  row[20] = timestamp;

  // Update row (starting from column B)
  for (let col = 1; col < row.length; col++) {
    if (row[col] !== undefined) {
      bugSheet.getRange(rowNum, col + 1).setValue(row[col]);
    }
  }
}

/**
 * Insert new bug row
 */
function insertBugRow_(bugSheet, issue, timestamp, instUrl, instance, moduleName) {
  const lastRow = bugSheet.getLastRow();
  const newRow = lastRow + 1;

  const row = [];

  // Column A: Bug ID (auto-generated or use Jira key)
  row[0] = 'BUG-' + String(newRow - 5).padStart(4, '0');

  // Column B: Type
  row[1] = 'Functional';

  // Column C: Priority
  row[2] = normalizePriority_(issue.fields.priority && issue.fields.priority.name);

  // Column D: Status
  row[3] = normalizeStatus_(issue.fields.status && issue.fields.status.name);

  // Column E: Feature
  const fieldMap = getJiraFieldMap_(instance);
  row[4] = getCustomFieldValue_(issue, fieldMap.feature);

  // Column F: Submodul
  row[5] = getCustomFieldValue_(issue, fieldMap.submodul) || moduleName || '';

  // Column G: Title
  row[6] = issue.fields.summary || '';

  // Column H: Description
  row[7] = convertADF_(issue.fields.description);

  // Column I: Environment
  row[8] = getCustomFieldValue_(issue, fieldMap.environment) || issue.fields.environment || '';

  // Column J: Steps
  row[9] = getCustomFieldValue_(issue, fieldMap.steps);

  // Column K: Expected
  row[10] = getCustomFieldValue_(issue, fieldMap.expected);

  // Column L: Actual
  row[11] = getCustomFieldValue_(issue, fieldMap.actual);

  // Column M: Test Case
  row[12] = '';

  // Column N: Reported By
  row[13] = issue.fields.reporter && issue.fields.reporter.displayName || '';

  // Column O: Assigned To
  row[14] = issue.fields.assignee && issue.fields.assignee.displayName || '';

  // Column P: Date Found
  row[15] = issue.fields.created ? issue.fields.created.substring(0, 10) : '';

  // Column Q: Date Fixed
  row[16] = '';

  // Column R: Sprint
  row[17] = '';

  // Column S: Link
  row[18] = instUrl + '/browse/' + issue.key;

  // Column T: Jira Key
  row[19] = issue.key;

  // Column U: Last Synced
  row[20] = timestamp;

  // Column V: Screenshot
  row[21] = '';

  // Insert row
  bugSheet.getRange(newRow, 1, 1, row.length).setValues([row]);
}

/**
 * Cleanup bugs that are Closed/Won't Fix in Jira
 */
function cleanupClosedBugs_(bugSheet, jiraIssues) {
  const closedKeys = new Set();

  jiraIssues.forEach(issue => {
    const statusName = issue.fields.status && issue.fields.status.name;
    if (statusName === 'Closed' || statusName === "Won't Fix") {
      closedKeys.add(issue.key);
    }
  });

  if (closedKeys.size === 0) {
    return 0;
  }

  const data = bugSheet.getDataRange().getValues();
  const rowsToDelete = [];

  // Find rows with closed bugs (from bottom to top for safe deletion)
  for (let i = data.length - 1; i >= 5; i--) {
    const jiraKey = String(data[i][19]).trim();
    if (closedKeys.has(jiraKey)) {
      rowsToDelete.push(i + 1);  // Store 1-based row number
    }
  }

  // Delete rows
  rowsToDelete.forEach(rowNum => {
    bugSheet.deleteRow(rowNum);
  });

  return rowsToDelete.length;
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

function getJiraInstanceUrl_(instance) {
  const instances = {
    'digitalperuri': 'https://digitalperuri.atlassian.net',
    'bgn-peruri': 'https://bgn-peruri.atlassian.net'
  };
  return instances[instance] || instances['digitalperuri'];
}

function getJiraModulField_(instance) {
  const fields = {
    'digitalperuri': 'cf[10097]',
    'bgn-peruri': 'cf[10289]'
  };
  return fields[instance] || 'cf[10097]';
}

function getJiraCustomFields_(instance) {
  const fields = {
    'digitalperuri': ',customfield_11090,customfield_10095,customfield_10560,customfield_10561,customfield_10562,customfield_11354',
    'bgn-peruri': ',customfield_10298,customfield_10291,customfield_10292,customfield_10293,customfield_10294,customfield_10300'
  };
  return fields[instance] || fields['digitalperuri'];
}

function getJiraFieldMap_(instance) {
  const maps = {
    'digitalperuri': {
      feature: 'customfield_11090',
      environment: 'customfield_10095',
      steps: 'customfield_10560',
      expected: 'customfield_10561',
      actual: 'customfield_10562',
      submodul: 'customfield_11354'
    },
    'bgn-peruri': {
      feature: 'customfield_10298',
      environment: 'customfield_10291',
      steps: 'customfield_10292',
      expected: 'customfield_10293',
      actual: 'customfield_10294',
      submodul: 'customfield_10300'
    }
  };
  return maps[instance] || maps['digitalperuri'];
}

function getCustomFieldValue_(issue, fieldId) {
  if (!fieldId || !issue.fields) return '';
  const value = issue.fields[fieldId];
  if (!value) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'object' && value.value) return value.value;
  if (typeof value === 'object' && value.content) return convertADF_(value);
  return '';
}

function normalizePriority_(priority) {
  if (!priority) return '';
  const map = {
    'Highest': 'Critical',
    'Critical': 'Critical',
    'High': 'High',
    'Medium': 'Medium',
    'Low': 'Low',
    'Lowest': 'Low',
    'Minor': 'Low',
    'Trivial': 'Low'
  };
  return map[priority] || priority;
}

function normalizeStatus_(status) {
  if (!status) return '';
  const sl = status.toLowerCase();
  if (['open', 'to do', 'backlog', 'new'].includes(sl)) return 'Open';
  if (['in progress', 'in review', 'review', 'testing'].includes(sl)) return 'In Progress';
  if (['fixed', 'ready for qa', 'ready for review'].includes(sl)) return 'Fixed';
  if (['verified', 'qa verified'].includes(sl)) return 'Verified';
  if (['closed', 'done'].includes(sl)) return 'Closed';
  if (["won't fix", 'wontfix', 'not a bug', 'invalid'].includes(sl)) return "Won't Fix";
  if (['reopened', 'reopen'].includes(sl)) return 'Reopen';
  return 'Open';
}

function convertADF_(adf) {
  if (!adf) return '';
  if (typeof adf === 'string') return adf;

  function extract(node) {
    if (!node) return '';
    if (node.type === 'text') return node.text || '';
    if (node.type === 'hardBreak') return '\n';
    if (node.type === 'paragraph') return (node.content || []).map(extract).join('') + '\n';
    if (node.type === 'bulletList' || node.type === 'orderedList') {
      return (node.content || []).map((item, i) =>
        (node.type === 'orderedList' ? (i + 1) + '. ' : ' • ') +
        (item.content || []).map(extract).join('')
      ).join('\n') + '\n';
    }
    if (node.type === 'mediaSingle' || node.type === 'mediaInline' || node.type === 'media') {
      return '[Image]';
    }
    if (node.content) return node.content.map(extract).join('');
    return '';
  }

  try {
    return extract(adf).trim();
  } catch (e) {
    return '';
  }
}

// ═══════════════════════════════════════════════════════════════════════
// TC_MASTER NOTES FIX
// ═══════════════════════════════════════════════════════════════════════

/**
 * Broadcast Fix: Apply TC_Master notes fix to all active QATMs
 * Fixes column notes that are shifted/misaligned in TC_Master tab
 */
function broadcastFixTCNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Fix TC_Master Column Notes',
    'Fix ini akan memperbaiki TC_Master column notes yang bergeser/tertukar.\n\n' +
    'Proses:\n' +
    '- Read semua QATM yang aktif dari Config\n' +
    '- Fix notes di header TC_Master (row 2)\n' +
    '- Notes dipindahkan ke kolom yang benar\n\n' +
    'Estimasi waktu: 30-60 detik per QATM.\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const cfg = ss.getSheetByName('Config');
    if (!cfg) {
      ui.alert('Config tab tidak ditemukan!');
      return;
    }

    const cfgData = cfg.getDataRange().getValues();
    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;
    const errors = [];

    Logger.log('Starting Broadcast Fix TC_Master Notes...');
    Logger.log('══════════════════════════════════════════');

    for (let i = 3; i < cfgData.length; i++) {
      const active = cfgData[i][0] === true;
      const project = String(cfgData[i][2]).trim();
      const modul = String(cfgData[i][3]).trim();
      const qatmId = String(cfgData[i][6]).trim();

      if (!active || !qatmId || qatmId.length < 10) continue;

      try {
        Logger.log('Processing ' + (i - 2) + '/' + (cfgData.length - 3) + ': ' + project + ' - ' + modul);

        const qatmSs = SpreadsheetApp.openById(qatmId);
        const ws = qatmSs.getSheetByName('TC_Master');

        if (!ws) {
          Logger.log('  TC_Master tab not found - skipping');
          skipCount++;
          continue;
        }

        applyTCNotesFix_(ws);
        successCount++;
        Logger.log('  Fixed: ' + qatmSs.getName());

      } catch (e) {
        Logger.log('  Error: ' + e.message);
        errorCount++;
        errors.push(project + ' - ' + modul + ' (' + e.message + ')');
      }
    }

    Logger.log('══════════════════════════════════════════');
    Logger.log('BROADCAST FIX COMPLETE');
    Logger.log('Success: ' + successCount + ', Skipped: ' + skipCount + ', Errors: ' + errorCount);

    let msg = 'Broadcast Fix TC_Master Notes Complete!\n\n';
    msg += 'Summary:\n';
    msg += '- Success: ' + successCount + ' QATM(s)\n';
    msg += '- Skipped: ' + skipCount + ' QATM(s)\n';
    msg += '- Errors: ' + errorCount + ' QATM(s)\n';

    if (errors.length > 0) {
      msg += '\nErrors:\n';
      errors.forEach(err => msg += '- ' + err + '\n');
    }

    ui.alert('Broadcast Complete', msg, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Error', 'Error during broadcast: ' + e.message, ui.ButtonSet.OK);
    Logger.log('Broadcast error: ' + e.message);
  }
}

/**
 * Helper: Apply TC notes fix to a worksheet
 * Fixes the column notes that are shifted/misaligned
 * Notes are applied to ROW 2 (header row for data columns)
 */
function applyTCNotesFix_(ws) {
  const correctNotes = [
    'Row number. Auto-filled when data is entered.',
    'SubModule — 3rd level in QA hierarchy:\n  Project > Module > SubModule > Feature\n\nSubModule = smallest standalone unit (one app or one domain).\n\nLayered project (e.g. SIPGN):\n  1.1 = Module 1, SubModule 1 (Nutritionist App)\n  1.2 = Module 1, SubModule 2 (Courier App)\n\nFlat project (e.g. INAGOV):\n  Direct name: Talenta, e-Office, SIMPEG\n\nMust be IDENTICAL in TC_Master and API_Master\nfor Dashboard coverage to merge correctly.',
    'TEST CASE ID\nUnique identifier untuk test case\n\nFORMAT RECOMMENDED:\nTC001, TC002, TC003...\natau TC-LOGIN-001, TC-DASHBOARD-001\n\nPENTING:\n• Harus unique dalam 1 modul\n• Jangan diubah setelah dibuat (untuk traceability)\n• Gunakan format yang konsisten',
    'Feature — Specific feature or page name.\nExamples: Login Page, Checkout Flow, User Management\nUsed for grouping Coverage per Feature in Summary.',
    'Priority & Impact:\n\nCRITICAL  → Must PASS before release. FAIL = release BLOCKED.\nHIGH      → Must PASS in same sprint. FAIL = needs PM approval.\nMEDIUM    → Potential blocker. Fix before UAT.\nLOW       → Non-blocker. Fix in next sprint.\nLOWEST    → Nice to have. Optional.\n\nAuto Test Level:\nCritical / High / Medium → Smoke Test\nLow / Lowest             → Regression Test',
    'Platform: Web / Mobile / Web & Mobile',
    'Test Type:\n  Positive   = happy path (valid data, expected flow)\n  Negative   = error case (invalid input, rejected action)\n  Edge Case  = boundary condition',
    'Automation Status:\n  Automated           = script exists and runs\n  To Do               = planned, not yet done\n  Manual              = decided to stay manual\n  Cannot be Automated = technically not possible',
    'Version: App version when TC was created. e.g. v1.0, v2.3',
    'Role (RBAC) — The user role executing this scenario.\n\nExamples: Admin, Super Admin, User, Viewer, Operator, Supervisor, Guest\n\nUsed to:\n  • Verify test coverage per role\n  • Confirm access control (RBAC) is correct\n  • Ensure 403 Forbidden for unauthorized roles',
    'SCENARIO NAMING STANDARD\n\nHappy Path : [Role] Successfully [Verb] [Object]\nNegative   : [Role] Failed to [Verb] [Object] with [Condition]\n\nRules:\n  • Role, Object → Title Case  (Nutritionist, Meal Plan)\n  • Verb → active  (Create, Pick Up, Confirm, Submit)\n  • Do not use: success/succeed, do, perform, process\n\nStandard examples:\n  Nutritionist Successfully Creates Meal Plan\n  Admin Failed to Delete User with Invalid ID\n\n───────────────────────────────────────\nSCENARIO OUTLINE (write here in this column)\nUse when same steps + same outcome type, different data.\nRule: all Examples = Positive only OR Negative only.\n\nNegative example:\nUser Failed to Log In with <invalid_credential>\nExamples:\n- Invalid password\n- Empty password\n- Expired session\n\nPositive example:\nAdmin Successfully Creates User with Role <role>\nExamples:\n- Viewer\n- Operator\n- Supervisor',
    'Steps / Gherkin Syntax\n\nUse BDD format:\nGiven [initial context]\nWhen [action taken]\nThen [expected outcome]\n\nExample:\nGiven user is on login page\nWhen user enters valid credentials\nAnd clicks login button\nThen user is redirected to dashboard',
    'Expected Result — What should happen when test passes.\n\nBe specific and measurable:\n  Good: "User sees success message: Account created"\n  Good: "Page loads within 3 seconds"\n  Bad: "It works" (too vague)',
    '[AUTO] Test Level\nAuto-calculated based on Priority:\n\nPriority = Critical/High/Medium → Smoke Test\nPriority = Low/Lowest → Regression Test\n\nDO NOT edit manually - this is auto-filled by formula.'
  ];

  ws.getRange(2, 1, 1, correctNotes.length).clearNote();

  for (let col = 1; col <= correctNotes.length; col++) {
    ws.getRange(2, col).setNote(correctNotes[col - 1]);
  }

  Logger.log('  All column notes updated (14 columns) at row 2');
}

// ═══════════════════════════════════════════════════════════════════════
// VAPT TABS BROADCAST & INDIVIDUAL TAB CREATION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Broadcast: Add VAPT tabs + Summary section to all active QATMs
 *
 * Adds 2 VAPT tabs + Summary metrics to ALL QATMs
 * Run from QA Dashboard: Menu > Broadcast Fixes > Add VAPT Tabs to All QATMs
 */
function broadcastVAPTTabsToAllQATMs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Verify running from Dashboard
  const cfg = ss.getSheetByName('Config');
  if (!cfg) {
    ui.alert(
      '❌ Config Not Found',
      'This function must be run from QA Dashboard.\n\n' +
      'Config tab with QATM module list not found.',
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.alert(
    '🔒 Broadcast VAPT Tabs + Summary',
    'Add 2 VAPT tabs + Summary section to ALL active QATMs:\n\n' +
    '📋 Tabs:\n' +
    '• Detail Finding - VAPT\n' +
    '• Evidence - VAPT\n\n' +
    '📊 Summary Section (row 35+):\n' +
    '• Total findings\n' +
    '• By Risk Level, Status Fix, Status Re-VAPT\n\n' +
    'QATMs with existing tabs will be skipped.\n' +
    'Time: ~1-2 minutes per QATM\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('❌ Broadcast cancelled.');
    return;
  }

  try {
    Logger.log('🔒 Starting VAPT Tabs Broadcast...');
    Logger.log('══════════════════════════════════════════');

    const cfgData = cfg.getDataRange().getValues();
    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;
    const errors = [];
    const skipped = [];

    // Process each QATM (starting from row 4, rows 1-3 are headers)
    for (let i = 3; i < cfgData.length; i++) {
      const active = cfgData[i][0] === true;  // Col A: Active
      const project = String(cfgData[i][2]).trim();  // Col C: Project
      const modul = String(cfgData[i][3]).trim();    // Col D: Modul
      const qatmId = String(cfgData[i][6]).trim();   // Col G: QATM Spreadsheet ID

      if (!active || !qatmId || qatmId.length < 10) {
        continue;  // Skip inactive or invalid entries
      }

      try {
        Logger.log('\n📂 Processing: ' + project + ' - ' + modul);
        Logger.log('   Spreadsheet ID: ' + qatmId);

        const qatmSs = SpreadsheetApp.openById(qatmId);
        Logger.log('   ✅ Opened: ' + qatmSs.getName());

        // Check if VAPT tabs already exist
        const existingTabs = [
          qatmSs.getSheetByName('Detail Finding - VAPT'),
          qatmSs.getSheetByName('Evidence - VAPT')
        ];

        const tabsExist = existingTabs.filter(tab => tab !== null).length;

        if (tabsExist === 2) {
          Logger.log('   ⏭️  All VAPT tabs already exist - skipping');
          skipCount++;
          skipped.push(project + ' - ' + modul + ' (already has VAPT tabs)');
          continue;
        }

        if (tabsExist > 0 && tabsExist < 2) {
          Logger.log('   ⚠️  Partial VAPT tabs exist (' + tabsExist + '/2) - recreating all');
        }

        // Create VAPT tabs via BroadcastVAPTTabs.js helper
        Logger.log('   🔧 Creating VAPT tabs...');
        createVAPTTabsInQATM_(qatmSs);

        // Add VAPT summary section via BroadcastVAPTTabs.js helper
        Logger.log('   📊 Adding VAPT summary section...');
        addVAPTSummarySection_(qatmSs, 35);

        successCount++;
        Logger.log('   ✅ SUCCESS: ' + project + ' - ' + modul);

        // Sleep to avoid rate limiting
        Utilities.sleep(1000);

      } catch (e) {
        Logger.log('   ❌ ERROR: ' + e.message);
        errorCount++;
        errors.push(project + ' - ' + modul + ' (' + e.message + ')');
      }
    }

    Logger.log('\n══════════════════════════════════════════');
    Logger.log('📊 BROADCAST SUMMARY');
    Logger.log('══════════════════════════════════════════');
    Logger.log('✅ Success: ' + successCount + ' QATM(s)');
    Logger.log('⏭️  Skipped: ' + skipCount + ' QATM(s)');
    Logger.log('❌ Errors: ' + errorCount + ' QATM(s)');

    // Show results
    let msg = '✅ Broadcast Complete!\n\n';
    msg += '📊 Summary:\n';
    msg += '• ✅ Added: ' + successCount + ' QATM(s)\n';
    msg += '• ⏭️  Skipped: ' + skipCount + ' QATM(s) (already exist)\n';
    msg += '• ❌ Errors: ' + errorCount + ' QATM(s)\n';

    if (errors.length > 0) {
      msg += '\n❌ Errors:\n';
      errors.slice(0, 5).forEach(err => msg += '• ' + err + '\n');
      if (errors.length > 5) {
        msg += '• ... +' + (errors.length - 5) + ' more (check log)\n';
      }
    }

    ui.alert('🔒 VAPT Broadcast', msg, ui.ButtonSet.OK);
    Logger.log('✅ Broadcast completed successfully');

  } catch (e) {
    Logger.log('❌ Broadcast error: ' + e.message);
    ui.alert(
      '❌ Error',
      'Broadcast failed:\n\n' + e.message + '\n\n' +
      'Check Execution log for details.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Create single VAPT tab: Detail Finding - VAPT
 * Useful for adding just this tab to existing QATM or new project
 */
function createSingleTabDetailFindingVAPT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📝 Create Detail Finding - VAPT',
    'Create "Detail Finding - VAPT" tab in this spreadsheet?\n\n' +
    '⚠️ If tab already exists, it will be recreated (existing data will be lost).',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // Delete existing tab if present
    const existing = ss.getSheetByName('Detail Finding - VAPT');
    if (existing) {
      ss.deleteSheet(existing);
      SpreadsheetApp.flush();
    }

    createDetailFindingVAPT(ss);
    SpreadsheetApp.flush();

    ui.alert(
      '✅ Tab Created!',
      'Detail Finding - VAPT tab has been created successfully.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ Error', 'Failed to create tab:\n\n' + e.message, ui.ButtonSet.OK);
    Logger.log('❌ Error creating Detail Finding - VAPT: ' + e.message);
  }
}

/**
 * Create single VAPT tab: Evidence - VAPT
 */
function createSingleTabEvidenceVAPT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📝 Create Evidence - VAPT',
    'Create "Evidence - VAPT" tab in this spreadsheet?\n\n' +
    '⚠️ If tab already exists, it will be recreated (existing data will be lost).',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    const existing = ss.getSheetByName('Evidence - VAPT');
    if (existing) {
      ss.deleteSheet(existing);
      SpreadsheetApp.flush();
    }

    createEvidenceVAPT(ss);
    SpreadsheetApp.flush();

    ui.alert(
      '✅ Tab Created!',
      'Evidence - VAPT tab has been created successfully.',
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ Error', 'Failed to create tab:\n\n' + e.message, ui.ButtonSet.OK);
    Logger.log('❌ Error creating Evidence - VAPT: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// V3: NEW VAPT STRUCTURE (3 TABS)
// ═══════════════════════════════════════════════════════════════════════

/**
 * V3 Broadcast - NEW VAPT Structure (3 tabs)
 * - VAPT - Helper (Dashboard/Tracking)
 * - VAPT - Detail Finding (32 columns)
 * - VAPT - Evidence (26 columns)
 * - Delete old VAPT tabs
 * - Delete Config & Sheet 2
 * - Reorder tabs: Summary, VAPT Helper, Bug Report, ...
 * - Update VAPT Summary section
 */
function broadcastV3NewVAPTStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const cfg = ss.getSheetByName('Config');
  if (!cfg) {
    ui.alert('❌ Error', 'Config tab not found. Run from QA Dashboard.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '🔒 V3: NEW VAPT Structure',
    'Broadcast 3 NEW VAPT tabs to active QATMs:\n\n' +
    '📋 New Tabs:\n' +
    '• VAPT - Helper (Dashboard/Tracking)\n' +
    '• VAPT - Detail Finding (32 columns)\n' +
    '• VAPT - Evidence (26 columns)\n\n' +
    '⏭️  Will Skip:\n' +
    '• QATMs that already have all 3 V3 tabs\n\n' +
    '🗑️ Will Delete (if not skipped):\n' +
    '• ALL old VAPT tabs\n' +
    '• Config tab\n' +
    '• Sheet 2\n\n' +
    '🔄 Tab Order: Summary, VAPT Helper, Bug Report, ...\n\n' +
    '⚠️ VAPT data will be lost (if not skipped)!\n' +
    'Time: ~2 min/QATM\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Broadcast cancelled.');
    return;
  }

  try {
    const cfgData = cfg.getDataRange().getValues();
    let successCount = 0;
    let skipCount = 0;
    let errorCount = 0;
    const errors = [];
    const skipped = [];
    const generated = [];

    for (let i = 3; i < cfgData.length; i++) {
      const active = cfgData[i][0] === true;
      const project = String(cfgData[i][2]).trim();
      const modul = String(cfgData[i][3]).trim();
      const qatmId = String(cfgData[i][6]).trim();

      if (!active || !qatmId || qatmId.length < 10) continue;

      try {
        const qatmSs = SpreadsheetApp.openById(qatmId);

        // Check if NEW VAPT tabs already exist
        const existingHelper = qatmSs.getSheetByName('VAPT - Helper');
        const existingDetail = qatmSs.getSheetByName('VAPT - Detail Finding');
        const existingEvidence = qatmSs.getSheetByName('VAPT - Evidence');

        if (existingHelper && existingDetail && existingEvidence) {
          // All 3 V3 tabs already exist - skip
          skipCount++;
          skipped.push(project + ' - ' + modul);
          continue;
        }

        // Delete ALL old VAPT tabs
        const oldVAPTTabs = [
          'Detail Finding - Regular VAPT', 'Evidence - Regular VAPT',
          'Detail Finding - Ad Hoc VAPT', 'Evidence - Ad Hoc VAPT',
          'Detail Finding - VAPT', 'Evidence - VAPT',
          'VAPT - Helper', 'VAPT - Detail Finding', 'VAPT - Evidence'
        ];
        oldVAPTTabs.forEach(name => {
          const tab = qatmSs.getSheetByName(name);
          if (tab) qatmSs.deleteSheet(tab);
        });

        // Delete Config & Sheet 2
        ['Config', 'Sheet 2'].forEach(name => {
          const tab = qatmSs.getSheetByName(name);
          if (tab) qatmSs.deleteSheet(tab);
        });

        SpreadsheetApp.flush();
        Utilities.sleep(100);

        // Create NEW VAPT tabs
        createVAPTHelper(qatmSs);
        SpreadsheetApp.flush();
        Utilities.sleep(100);

        createDetailFindingVAPT(qatmSs);
        SpreadsheetApp.flush();
        Utilities.sleep(100);

        createEvidenceVAPT(qatmSs);
        SpreadsheetApp.flush();

        // Reorder tabs: Summary, VAPT Helper, Bug Report, ...
        const summary = qatmSs.getSheetByName('Summary');
        const vaptHelper = qatmSs.getSheetByName('VAPT - Helper');
        const bugReport = qatmSs.getSheetByName('Bug Report');

        if (summary) qatmSs.setActiveSheet(summary);
        if (vaptHelper) qatmSs.moveActiveSheet(2);
        if (bugReport) {
          qatmSs.setActiveSheet(bugReport);
          qatmSs.moveActiveSheet(3);
        }

        SpreadsheetApp.flush();

        // Update VAPT Summary section
        addVAPTSummarySection_(qatmSs, 35);

        successCount++;
        generated.push(project + ' - ' + modul);
        Utilities.sleep(500);

      } catch (e) {
        errorCount++;
        errors.push(project + ' - ' + modul + ': ' + e.message);
      }
    }

    let msg = '✅ V3 Broadcast Complete!\n\n';
    msg += '📊 Summary:\n';
    msg += '• ✅ Generated: ' + successCount + ' QATM(s)\n';
    msg += '• ⏭️  Skipped: ' + skipCount + ' QATM(s)\n';
    msg += '• ❌ Errors: ' + errorCount + ' QATM(s)\n';

    if (generated.length > 0) {
      msg += '\n✅ Generated:\n';
      generated.slice(0, 5).forEach(name => msg += '• ' + name + '\n');
      if (generated.length > 5) msg += '• ... +' + (generated.length - 5) + ' more\n';
    }

    if (skipped.length > 0) {
      msg += '\n⏭️  Skipped (already have V3 tabs):\n';
      skipped.slice(0, 5).forEach(name => msg += '• ' + name + '\n');
      if (skipped.length > 5) msg += '• ... +' + (skipped.length - 5) + ' more\n';
    }

    if (errors.length > 0) {
      msg += '\n❌ Errors:\n';
      errors.slice(0, 3).forEach(err => msg += '• ' + err + '\n');
      if (errors.length > 3) msg += '• ... +' + (errors.length - 3) + ' more\n';
    }

    ui.alert('🔒 V3 Broadcast', msg, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('❌ Error', 'Broadcast failed:\n\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Add VAPT Summary Section to Summary tab
 */
function addVAPTSummarySection_(ss, startRow) {
  const summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) return;

  startRow = startRow || 35;

  const headerBg = '#263238';
  const vaptBg = '#FF6F00';
  const sectionBg = '#E3F2FD';
  const white = '#FFFFFF';

  try {
    const titleRange = summarySheet.getRange(startRow, 1, 1, 3);
    titleRange.merge()
      .setValue('🔒 VAPT FINDINGS SUMMARY')
      .setBackground(headerBg)
      .setFontColor(white)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    summarySheet.setRowHeight(startRow, 35);

    let currentRow = startRow + 2;

    summarySheet.getRange(currentRow, 1, 1, 2).merge()
      .setValue('Overview')
      .setBackground(sectionBg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    summarySheet.getRange(currentRow, 1).setValue('Total VAPT Findings').setFontWeight('bold');
    summarySheet.getRange(currentRow, 2)
      .setFormula('=COUNTA(\'VAPT - Detail Finding\'!A3:A1000)-COUNTBLANK(\'VAPT - Detail Finding\'!A3:A1000)')
      .setNumberFormat('0')
      .setHorizontalAlignment('center')
      .setBackground(vaptBg)
      .setFontColor(white)
      .setFontWeight('bold')
      .setFontSize(11);
    currentRow += 2;

    summarySheet.getRange(currentRow, 1, 1, 2).merge()
      .setValue('By Risk Level (Adjusted Risk)')
      .setBackground(sectionBg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    summarySheet.getRange(currentRow, 1).setValue('Risk Level').setFontWeight('bold');
    summarySheet.getRange(currentRow, 2).setValue('Count').setFontWeight('bold').setHorizontalAlignment('center');
    currentRow++;

    const riskLevels = [
      {level: 'Critical', color: '#FFEBEE'},
      {level: 'High', color: '#FFCDD2'},
      {level: 'Medium', color: '#FFF9C4'},
      {level: 'Low', color: '#FFF8E1'},
      {level: 'Informational', color: '#E3F2FD'}
    ];

    riskLevels.forEach(risk => {
      summarySheet.getRange(currentRow, 1).setValue(risk.level)
        .setBackground(risk.color)
        .setFontWeight('bold');
      summarySheet.getRange(currentRow, 2)
        .setFormula('=COUNTIF(\'VAPT - Detail Finding\'!H:H,"' + risk.level + '")')
        .setNumberFormat('0')
        .setHorizontalAlignment('center');
      currentRow++;
    });

    currentRow++;

    summarySheet.getRange(currentRow, 1, 1, 2).merge()
      .setValue('By Status Fix (Dev Team)')
      .setBackground(sectionBg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    summarySheet.getRange(currentRow, 1).setValue('Status').setFontWeight('bold');
    summarySheet.getRange(currentRow, 2).setValue('Count').setFontWeight('bold').setHorizontalAlignment('center');
    currentRow++;

    const statusFix = ['Todo', 'On Progress Remediation', 'Ready to Retest', 'Done', 'Accepted', 'False Positive'];
    statusFix.forEach(status => {
      summarySheet.getRange(currentRow, 1).setValue(status);
      summarySheet.getRange(currentRow, 2)
        .setFormula('=COUNTIF(\'VAPT - Detail Finding\'!E:E,"' + status + '")')
        .setNumberFormat('0')
        .setHorizontalAlignment('center');
      currentRow++;
    });

    currentRow++;

    summarySheet.getRange(currentRow, 1, 1, 2).merge()
      .setValue('By Status Re-VAPT (Pentester)')
      .setBackground(sectionBg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    summarySheet.getRange(currentRow, 1).setValue('Status').setFontWeight('bold');
    summarySheet.getRange(currentRow, 2).setValue('Count').setFontWeight('bold').setHorizontalAlignment('center');
    currentRow++;

    summarySheet.getRange(currentRow, 1).setValue('Open')
      .setBackground('#FFEBEE')
      .setFontWeight('bold');
    summarySheet.getRange(currentRow, 2)
      .setFormula('=COUNTIF(\'VAPT - Detail Finding\'!F:F,"Open")')
      .setNumberFormat('0')
      .setHorizontalAlignment('center');
    currentRow++;

    summarySheet.getRange(currentRow, 1).setValue('Closed')
      .setBackground('#E8F5E9')
      .setFontWeight('bold');
    summarySheet.getRange(currentRow, 2)
      .setFormula('=COUNTIF(\'VAPT - Detail Finding\'!F:F,"Closed")')
      .setNumberFormat('0')
      .setHorizontalAlignment('center');
    currentRow++;

    summarySheet.setColumnWidth(1, 200);
    summarySheet.setColumnWidth(2, 80);

    const sectionRange = summarySheet.getRange(startRow, 1, currentRow - startRow, 2);
    sectionRange.setBorder(
      true, true, true, true, true, true,
      '#CFD8DC', SpreadsheetApp.BorderStyle.SOLID
    );

  } catch (e) {
    // Silent skip
  }
}

// ═══════════════════════════════════════════════════════════════════════
// BROADCAST FIX: ADD VAPT BLOCKER BREAKDOWN TO SUMMARY
// ═══════════════════════════════════════════════════════════════════════

function broadcastAddVAPTBlockerBreakdown() {
  const dashboardSs = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const modules = getModuleList_(dashboardSs);
  const targets = [];
  const seen = {};

  modules.forEach(mod => {
    if (!mod || !mod.id || seen[mod.id]) return;
    seen[mod.id] = true;
    targets.push(mod);
  });

  if (targets.length === 0) {
    ui.alert('No active QATM targets found in Config.');
    return;
  }

  const response = ui.alert(
    'Broadcast: Add VAPT Blocker Breakdown',
    'Will ADD VAPT Blocker Breakdown section to Summary tab in ' + targets.length + ' QATM sheets.\\n\\n' +
    '✅ What will be added:\\n' +
    '• VAPT Blocker Count (total blockers)\\n' +
    '• VAPT Blocker Critical\\n' +
    '• VAPT Blocker High\\n' +
    '• VAPT Blocker Medium\\n' +
    '• VAPT Non Blocker Count\\n\\n' +
    'Formula: Blocker = Critical/High/Medium with status NOT Done/False Positive\\n\\n' +
    '⚠️ WARNING:\\n' +
    '• Section will be appended at row 35+\\n' +
    '• Check for conflicts with existing data\\n\\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let added = 0;
  let skipped = 0;
  let failed = 0;
  const errors = [];

  targets.forEach(mod => {
    try {
      const qatmSs = SpreadsheetApp.openById(mod.id);
      const summarySheet = qatmSs.getSheetByName('Summary');

      if (!summarySheet) {
        skipped++;
        return;
      }

      // Check if VAPT tab exists
      const vaptSheet = qatmSs.getSheetByName('VAPT - Detail Finding') || qatmSs.getSheetByName('Detail Finding - VAPT');
      if (!vaptSheet) {
        skipped++;
        return;
      }

      // Add blocker breakdown section
      addVAPTBlockerBreakdownToSummary_(qatmSs);
      added++;

    } catch (error) {
      failed++;
      errors.push((mod.project || '') + ' / ' + (mod.module || '') + ': ' + error.message);
      Logger.log('Add VAPT Blocker Breakdown failed [' + (mod.id || '-') + ']: ' + error.stack);
    }
  });

  ui.alert(
    'VAPT Blocker Breakdown Broadcast Complete',
    'Added: ' + added + '\\n' +
    'Skipped (no Summary/VAPT tab): ' + skipped + '\\n' +
    'Failed: ' + failed +
    (errors.length ? '\\n\\nErrors:\\n' + errors.slice(0, 8).join('\\n') : '') +
    '\\n\\n✅ Summary tabs now include VAPT Blocker Breakdown section.',
    ui.ButtonSet.OK
  );
}

function addVAPTBlockerBreakdownToSummary_(ss) {
  const summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) return;

  // Determine VAPT tab name
  let vaptTabName = 'VAPT - Detail Finding';
  if (!ss.getSheetByName(vaptTabName)) {
    vaptTabName = 'Detail Finding - VAPT';
    if (!ss.getSheetByName(vaptTabName)) return;
  }

  // Find insertion point after VAPT FINDINGS SUMMARY section
  let startRow = 99; // Default fallback
  const data = summarySheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const rowText = data[i].join(' ').toUpperCase();
    if (rowText.includes('BY STATUS RE-VAPT') || rowText.includes('STATUS RE-VAPT')) {
      startRow = i + 5; // Insert 5 rows after "By Status Re-VAPT" section
      break;
    }
  }
  // Safety check: if startRow too low, use minimum 99
  if (startRow < 99) startRow = 99;

  const sectionBg = '#FFCDD2';
  const labelBg = '#FFEBEE';
  const labelBg2 = '#FFF3E0';
  const labelBg3 = '#E8F5E9';
  const valueBg = '#FFFFFF';
  const blockerHighlight = '#FFCDD2';
  const textDark = '#B71C1C';
  const noteBg = '#FFF8E1';
  const noteText = '#E65100';

  try {
    let R = startRow;

    // Section header
    summarySheet.getRange(R, 1, 1, 5).merge()
      .setValue('VAPT Blockers (Status: Todo / On Progress / Ready to Retest / Accepted)')
      .setBackground(sectionBg)
      .setFontColor('#C62828')
      .setFontWeight('bold')
      .setFontFamily('Arial')
      .setFontSize(9)
      .setHorizontalAlignment('left')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    summarySheet.setRowHeight(R, 18);
    R++;

    // VAPT Blocker Count (total)
    summarySheet.getRange(R, 1)
      .setValue('VAPT Blocker Count')
      .setBackground(labelBg)
      .setFontFamily('Arial')
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    summarySheet.getRange(R, 2, 1, 4).merge()
      .setFormula('=IFERROR(SUMPRODUCT((\'' + vaptTabName + '\'!E:E<>"Done")*(\'' + vaptTabName + '\'!E:E<>"False Positive")*((\'' + vaptTabName + '\'!H:H="Critical")+(\'' + vaptTabName + '\'!H:H="High")+(\'' + vaptTabName + '\'!H:H="Medium"))),0)')
      .setBackground(blockerHighlight)
      .setFontFamily('Arial')
      .setFontSize(14)
      .setFontWeight('bold')
      .setFontColor(textDark)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    summarySheet.setRowHeight(R, 28);
    R++;

    // VAPT Blocker Critical
    summarySheet.getRange(R, 1)
      .setValue('VAPT Blocker Critical')
      .setBackground(labelBg2)
      .setFontFamily('Arial')
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    summarySheet.getRange(R, 2, 1, 4).merge()
      .setFormula('=IFERROR(SUMPRODUCT((\'' + vaptTabName + '\'!E:E<>"Done")*(\'' + vaptTabName + '\'!E:E<>"False Positive")*(\'' + vaptTabName + '\'!H:H="Critical")),0)')
      .setBackground(valueBg)
      .setFontFamily('Arial')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    summarySheet.setRowHeight(R, 22);
    R++;

    // VAPT Blocker High
    summarySheet.getRange(R, 1)
      .setValue('VAPT Blocker High')
      .setBackground(labelBg2)
      .setFontFamily('Arial')
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    summarySheet.getRange(R, 2, 1, 4).merge()
      .setFormula('=IFERROR(SUMPRODUCT((\'' + vaptTabName + '\'!E:E<>"Done")*(\'' + vaptTabName + '\'!E:E<>"False Positive")*(\'' + vaptTabName + '\'!H:H="High")),0)')
      .setBackground(valueBg)
      .setFontFamily('Arial')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    summarySheet.setRowHeight(R, 22);
    R++;

    // VAPT Blocker Medium
    summarySheet.getRange(R, 1)
      .setValue('VAPT Blocker Medium')
      .setBackground(labelBg2)
      .setFontFamily('Arial')
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    summarySheet.getRange(R, 2, 1, 4).merge()
      .setFormula('=IFERROR(SUMPRODUCT((\'' + vaptTabName + '\'!E:E<>"Done")*(\'' + vaptTabName + '\'!E:E<>"False Positive")*(\'' + vaptTabName + '\'!H:H="Medium")),0)')
      .setBackground(valueBg)
      .setFontFamily('Arial')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    summarySheet.setRowHeight(R, 22);
    R++;

    // VAPT Non Blocker Count
    summarySheet.getRange(R, 1)
      .setValue('VAPT Non Blocker Count')
      .setBackground(labelBg3)
      .setFontFamily('Arial')
      .setFontSize(9)
      .setFontWeight('bold')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    summarySheet.getRange(R, 2, 1, 4).merge()
      .setFormula('=IFERROR(SUMPRODUCT((\'' + vaptTabName + '\'!E:E<>"Done")*(\'' + vaptTabName + '\'!E:E<>"False Positive")*((\'' + vaptTabName + '\'!H:H="Low")+(\'' + vaptTabName + '\'!H:H="Informational"))),0)')
      .setBackground(valueBg)
      .setFontFamily('Arial')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setNumberFormat('0')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    summarySheet.setRowHeight(R, 22);
    R++;

    // Note
    summarySheet.getRange(R, 1, 1, 5).merge()
      .setValue('Target: VAPT Blocker Count = 0 sebelum closure sign-off. Blocker = severity Critical/High/Medium dengan status selain Done/False Positive.')
      .setBackground(noteBg)
      .setFontColor(noteText)
      .setFontStyle('italic')
      .setFontSize(7)
      .setFontFamily('Arial')
      .setHorizontalAlignment('left')
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    summarySheet.setRowHeight(R, 14);

    SpreadsheetApp.flush();

  } catch (e) {
    throw new Error('Failed to add VAPT Blocker Breakdown: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// MENU FUNCTIONS - CREATE NEW VAPT TABS
// ═══════════════════════════════════════════════════════════════════════

function menuCreateVAPTHelper() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📝 Create VAPT - Helper',
    'Create "VAPT - Helper" tab?\n\n' +
    '⚠️ If exists, will be recreated (data lost).',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const existing = ss.getSheetByName('VAPT - Helper');
    if (existing) {
      ss.deleteSheet(existing);
      SpreadsheetApp.flush();
    }

    createVAPTHelper(ss);
    SpreadsheetApp.flush();

    ui.alert('✅ Tab Created!', 'VAPT - Helper created successfully.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Error', 'Failed:\n\n' + e.message, ui.ButtonSet.OK);
  }
}

function menuCreateVAPTDetailFinding() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📝 Create VAPT - Detail Finding',
    'Create "VAPT - Detail Finding" tab?\n\n' +
    '⚠️ If exists, will be recreated (data lost).',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const existing = ss.getSheetByName('VAPT - Detail Finding');
    if (existing) {
      ss.deleteSheet(existing);
      SpreadsheetApp.flush();
    }

    createDetailFindingVAPT(ss);
    SpreadsheetApp.flush();

    ui.alert('✅ Tab Created!', 'VAPT - Detail Finding created successfully.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Error', 'Failed:\n\n' + e.message, ui.ButtonSet.OK);
  }
}

function menuCreateVAPTEvidence() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '📝 Create VAPT - Evidence',
    'Create "VAPT - Evidence" tab?\n\n' +
    '⚠️ If exists, will be recreated (data lost).',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const existing = ss.getSheetByName('VAPT - Evidence');
    if (existing) {
      ss.deleteSheet(existing);
      SpreadsheetApp.flush();
    }

    createEvidenceVAPT(ss);
    SpreadsheetApp.flush();

    ui.alert('✅ Tab Created!', 'VAPT - Evidence created successfully.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Error', 'Failed:\n\n' + e.message, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// VAPT BLOCKER BREAKDOWN - ADD TO SUMMARY TAB
// ═══════════════════════════════════════════════════════════════════════

/**
 * Broadcast: Add VAPT Blocker Breakdown to Summary tabs
 *
 * Adds new section to Summary tab with:
 * - VAPT Blocker Count (correct formula: status != Done/False Positive, severity Critical/High/Medium)
 * - VAPT Blocker Critical/High/Medium breakdown
 * - VAPT Non Blocker Count (Low/Info not Done/False Positive)
 *
 * Safe: Only adds section if not already present (checks for "VAPT Blocker Count" label)
 */
function broadcastVAPTBlockerBreakdown() {
  const dashboardSs = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const modules = getModuleList_(dashboardSs);
  const targets = [];
  const seen = {};

  modules.forEach(mod => {
    if (!mod || !mod.id || seen[mod.id]) return;
    seen[mod.id] = true;
    targets.push(mod);
  });

  if (targets.length === 0) {
    ui.alert('No active QATM targets found in Config.');
    return;
  }

  const response = ui.alert(
    'Broadcast VAPT Blocker Breakdown to Summary',
    'Will add VAPT blocker breakdown section to ' + targets.length + ' QATM Summary tabs.\n\n' +
    'Safe for existing data:\n' +
    '- Only adds if section not present\n' +
    '- Checks for existing "VAPT Blocker Count" label\n' +
    '- Requires "Detail Finding - VAPT" tab\n\n' +
    'New fields:\n' +
    '✓ VAPT Blocker Count (status != Done/False Positive)\n' +
    '✓ VAPT Blocker Critical/High/Medium\n' +
    '✓ VAPT Non Blocker Count\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let created = 0;
  let skipped = 0;
  let failed = 0;
  const errors = [];

  targets.forEach(mod => {
    try {
      const qatmSs = SpreadsheetApp.openById(mod.id);
      const summarySheet = qatmSs.getSheetByName('Summary');

      if (!summarySheet) {
        skipped++;
        return;
      }

      // Check if already exists
      const dataRange = summarySheet.getDataRange();
      const values = dataRange.getValues();
      let hasBlockerSection = false;

      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          if (String(values[i][j]).includes('VAPT Blocker Count')) {
            hasBlockerSection = true;
            break;
          }
        }
        if (hasBlockerSection) break;
      }

      if (hasBlockerSection) {
        skipped++;
        return;
      }

      // Check if Detail Finding - VAPT exists
      if (!qatmSs.getSheetByName('Detail Finding - VAPT')) {
        errors.push((mod.project || '') + ' / ' + (mod.module || '') + ': Missing Detail Finding - VAPT tab');
        failed++;
        return;
      }

      addVAPTBlockerSection_(summarySheet);
      created++;
    } catch (error) {
      failed++;
      errors.push((mod.project || '') + ' / ' + (mod.module || '') + ': ' + error.message);
      Logger.log('VAPT Blocker broadcast failed [' + (mod.id || '-') + ']: ' + error.stack);
    }
  });

  ui.alert(
    'VAPT Blocker Breakdown Broadcast Complete',
    'Created: ' + created + '\n' +
    'Skipped (already exists or no Summary): ' + skipped + '\n' +
    'Failed: ' + failed +
    (errors.length ? '\n\nErrors:\n' + errors.slice(0, 8).join('\n') : ''),
    ui.ButtonSet.OK
  );
}

function addVAPTBlockerSection_(ws) {
  // Find last row with content
  let lastRow = ws.getMaxRows();
  const values = ws.getRange(1, 1, lastRow, 10).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const rowHasData = values[i].some(cell => cell !== '' && cell != null);
    if (rowHasData) {
      lastRow = i + 1;
      break;
    }
  }

  // Remove footer if exists
  if (lastRow > 0) {
    const footerCheck = String(values[lastRow - 1][0]).toLowerCase();
    if (footerCheck.includes('peruri') || footerCheck.includes('qa team')) {
      ws.deleteRow(lastRow);
      lastRow--;
    }
  }

  let R = lastRow + 2; // Start after last content with gap
  const L = 1; // Left column
  const LW = 10; // Left width

  function bd(range) {
    return range.setBorder(true, true, true, true, false, false, '#1976D2', SpreadsheetApp.BorderStyle.SOLID);
  }

  function h_(range, bg) {
    return bd(range).setBackground(bg || '#BF360C').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  function lbl(row, col, text, bg) {
    bd(ws.getRange(row, col)).setValue(text)
        .setBackground(bg || '#E3F2FD').setFontColor('#0D47A1').setFontWeight('bold')
        .setFontFamily('Arial').setFontSize(9)
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
    ws.setRowHeight(row, 24);
  }

  function m_(row, col, nr, nc) {
    return ws.getRange(row, col, nr, nc).merge();
  }

  // Section header
  ws.setRowHeight(R, 6); R++;
  m_(R, L, 1, LW);
  h_(ws.getRange(R, L), '#BF360C').setValue('F.  VAPT BLOCKER BREAKDOWN');
  ws.setRowHeight(R, 20); R++;

  // VAPT Blocker Count
  lbl(R, L, 'VAPT Blocker Count', '#FFEBEE');
  m_(R, L + 1, 1, LW - 1);
  bd(ws.getRange(R, L + 1)).setFormula('=IFERROR(SUMPRODUCT((\'Detail Finding - VAPT\'!H:H<>"Done")*(\'Detail Finding - VAPT\'!H:H<>"False Positive")*((\'Detail Finding - VAPT\'!E:E="Critical")+(\'Detail Finding - VAPT\'!E:E="High")+(\'Detail Finding - VAPT\'!E:E="Medium"))),0)')
      .setBackground('#FFCDD2').setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
      .setFontColor('#B71C1C').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R, 28); R++;

  // VAPT Blocker Critical
  lbl(R, L, 'VAPT Blocker Critical', '#FFF3E0');
  m_(R, L + 1, 1, LW - 1);
  bd(ws.getRange(R, L + 1)).setFormula('=IFERROR(SUMPRODUCT((\'Detail Finding - VAPT\'!H:H<>"Done")*(\'Detail Finding - VAPT\'!H:H<>"False Positive")*(\'Detail Finding - VAPT\'!E:E="Critical")),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R, 22); R++;

  // VAPT Blocker High
  lbl(R, L, 'VAPT Blocker High', '#FFF3E0');
  m_(R, L + 1, 1, LW - 1);
  bd(ws.getRange(R, L + 1)).setFormula('=IFERROR(SUMPRODUCT((\'Detail Finding - VAPT\'!H:H<>"Done")*(\'Detail Finding - VAPT\'!H:H<>"False Positive")*(\'Detail Finding - VAPT\'!E:E="High")),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R, 22); R++;

  // VAPT Blocker Medium
  lbl(R, L, 'VAPT Blocker Medium', '#FFF3E0');
  m_(R, L + 1, 1, LW - 1);
  bd(ws.getRange(R, L + 1)).setFormula('=IFERROR(SUMPRODUCT((\'Detail Finding - VAPT\'!H:H<>"Done")*(\'Detail Finding - VAPT\'!H:H<>"False Positive")*(\'Detail Finding - VAPT\'!E:E="Medium")),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R, 22); R++;

  // VAPT Non Blocker Count
  lbl(R, L, 'VAPT Non Blocker Count', '#E8F5E9');
  m_(R, L + 1, 1, LW - 1);
  bd(ws.getRange(R, L + 1)).setFormula('=IFERROR(SUMPRODUCT((\'Detail Finding - VAPT\'!H:H<>"Done")*(\'Detail Finding - VAPT\'!H:H<>"False Positive")*((\'Detail Finding - VAPT\'!E:E="Low")+(\'Detail Finding - VAPT\'!E:E="Informational"))),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R, 22); R++;

  // Note
  m_(R, L, 1, LW);
  ws.getRange(R, L).setValue('Target: VAPT Blocker Count = 0 before closure sign-off. Blocker = Critical/High/Medium with status not Done/False Positive.')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic').setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(R, 14); R++;

  // Re-add footer
  R++;
  const PERURI_COPYRIGHT = '(c) QA INA Digital  |  Template ini merupakan properti QA Team INA Digital  |  Dilarang digunakan/disebarluaskan tanpa izin  |  departemen.qa@inadigital.co.id';
  ws.getRange(R, 1, 1, 21).merge();
  ws.getRange(R, 1)
      .setValue(PERURI_COPYRIGHT)
      .setBackground('#0D47A1')
      .setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(8)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(false);
  ws.setRowHeight(R, 20);
  ws.getRange(R, 1, 1, 21)
      .setBorder(true, false, false, false, false, false, '#1976D2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
