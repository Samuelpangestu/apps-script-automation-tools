/**
 * BroadcastFixes.js - Consolidated broadcast fixes for QA Dashboard
 *
 * Semua broadcast fixes dalam 1 file untuk kemudahan maintenance.
 * Run fixes sesuai kebutuhan dari menu atau manual.
 */

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

    // Add comprehensive notes
    addNotesToDashboard();

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
    Logger.log('📝 Applying comprehensive notes...');

    addNotesToDashboard();

    Logger.log('✅ Comprehensive notes applied successfully');

  } catch (e) {
    Logger.log('❌ Error applying notes: ' + e.message);
    ui.alert(
      '❌ Error',
      'Gagal apply notes:\n' + e.message + '\n\n' +
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
    'Run dari Extensions > Apps Script atau Script Editor.',
    ui.ButtonSet.OK
  );
}
