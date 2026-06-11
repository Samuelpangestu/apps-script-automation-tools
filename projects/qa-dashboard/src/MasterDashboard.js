/**
 * QA_Portfolio_Dashboard.js  —  v40
 * ═══════════════════════════════════════════════════════════════════════
 * Paste SELURUH FILE ini ke Apps Script QA PORTFOLIO DASHBOARD.
 *
 * SETUP PERTAMA KALI:
 *   createDashboard()    → buat semua tab (Config, Overview, Smoke, ...)
 *
 * REFRESH DATA:
 *   refreshDashboard()   → pull data dari semua modul aktif di Config
 *
 * AUTO TRIGGER:
 *   setupTrigger()       → auto refresh setiap 1 jam
 *
 * TABS YANG DIBUAT:
 *   Config   — daftar modul (Active, Project, Module, SubModule, PIC, QA Lead, SpreadsheetID)
 *   Overview — KPI portfolio: WEB | SMOKE WEB | API | SMOKE API | BUGS (25 col)
 *   Smoke    — dedicated Smoke Test view + 5 charts per modul
 *   Failure Scenario — TC Priority Critical/High/Medium yang FAILED/BLOCKED
 *   Coverage — coverage per SubModul
 *   History  — trend data tiap refresh (termasuk Smoke trend)
 *   _Raw     — cache internal (jangan edit manual)
 *
 * LAYOUT OVERVIEW (25 kolom):
 *   Col  1-4 : MODULE INFO  (SubModule, Project, Module, PIC/Team)
 *   Col  5-9 : WEB/MOBILE  (Total, Pass, Fail, Block, Pass%)
 *   Col 10-12: SMOKE WEB   (Total, Pass%, Exec%)
 *   Col 13-17: API          (Total, Pass, Fail, Block, Pass%)
 *   Col 18-20: SMOKE API    (Total, Pass%, Exec%)
 *   Col 21   : PERF
 *   Col 22-24: BUGS         (Total, Blocker, Critical)
 *   Col 25   : NOTES
 *
 * QA TEAM LEAD:
 *   - Dibaca dari Config col F (manual input)
 *   - ATAU auto-filled dari Summary baris B4 saat refresh
 *   - Tampil di tab Smoke col "QA Lead"
 * ═══════════════════════════════════════════════════════════════════════
 */


// ═══════════════════════════════════════════════════════════════════════
// CUSTOM MENU - Auto-loads when spreadsheet opens
// ═══════════════════════════════════════════════════════════════════════

/**
 * Force reload menu - Run this manually if menu doesn't update after deployment
 *
 * HOW TO USE:
 * 1. Open Apps Script Editor
 * 2. Select function: forceReloadMenu
 * 3. Click Run (▶️)
 */
function forceReloadMenu() {
  onOpen();
  SpreadsheetApp.getUi().alert(
    '✅ Menu Reloaded!',
    'Menu sudah di-reload dengan versi terbaru.\n\n' +
    'Kalau menu masih belum update, coba:\n' +
    '1. Close semua tab Dashboard spreadsheet\n' +
    '2. Clear browser cache (Ctrl+Shift+Del)\n' +
    '3. Buka Dashboard lagi dari Google Drive',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 QA Dashboard')
    .addSubMenu(ui.createMenu('📊 Dashboard')
      .addItem('Create Dashboard (First Time)', 'createDashboard')
      .addSeparator()
      .addItem('▶️ Manual Sync + Refresh', 'manualSyncAndRefresh')
      .addItem('⚙️ Setup External QA Config', 'menuSetupExternalQAConfig')
      .addItem('🚀 Setup Auto-Refresh Trigger', 'setupAutoRefreshTrigger')
      .addSeparator()
      .addItem('⚙️ Refresh Bug Only (No Jira Sync)', 'refreshBugOnly')
      .addItem('🔒 Refresh VAPT Only', 'refreshVAPTOnly'))
    .addSubMenu(ui.createMenu('🔔 Notifications')
      .addItem('⚙️ Setup Blocker Notification (Daily)', 'setupDailyBlockerNotification')
      .addItem('⚙️ Setup Test Execution Notification (Daily)', 'setupTestExecutionNotification')
      .addSeparator()
      .addItem('✅ Test Blocker Notification Now', 'sendBlockerNotification')
      .addItem('✅ Test Execution Notification Now', 'sendTestExecutionNotification')
      .addSeparator()
      .addItem('❌ Remove Blocker Notification Trigger', 'removeDailyBlockerNotification')
      .addItem('❌ Remove Test Execution Notification Trigger', 'removeTestExecutionNotification')
      .addSeparator()
      .addItem('📱 WhatsApp: Get Groups', 'menuTestGetGroups')
      .addItem('📱 WhatsApp: Send Test', 'menuTestSendToGroup'))
    .addSubMenu(ui.createMenu('🔄 Jira Sync')
      .addItem('Sync All Modules from Jira', 'syncAllJira')
      .addItem('Show Jira JQL for Module', 'showJiraJQL'))
    .addSubMenu(ui.createMenu('🔧 Rebuild Individual Tabs')
      .addItem('Config', 'rebuildConfig')
      .addItem('Credentials', 'rebuildCredentials')
      .addItem('Overview', 'rebuildOverview')
      .addItem('Bugs', 'rebuildBugs')
      .addItem('VAPT', 'rebuildVAPT')
      .addItem('VAPT History', 'rebuildVAPTHistory')
      .addItem('Smoke', 'rebuildSmoke')
      .addItem('Failure Scenario', 'rebuildFailureScenario')
      .addItem('Coverage', 'rebuildCoverage')
      .addItem('History', 'rebuildHistory')
      .addItem('Automation Runs', 'rebuildAutomationRuns')
      .addItem('_Raw', 'rebuildRaw')
      .addSeparator()
      .addItem('🔒 VAPT - Helper', 'menuCreateVAPTHelper')
      .addItem('🔒 VAPT - Detail Finding', 'menuCreateVAPTDetailFinding')
      .addItem('🔒 VAPT - Evidence', 'menuCreateVAPTEvidence'))
    .addSubMenu(ui.createMenu('🔧 Broadcast Fixes')
      .addItem('Fix BUG BLOCKER (Rename + Formula)', 'broadcastBugBlockerFix')
      .addSeparator()
      .addItem('External QA: Create Report Tab', 'broadcastExternalTestReportTab')
      .addSeparator()
      .addItem('V3: NEW VAPT Structure (3 Tabs)', 'broadcastV3NewVAPTStructure'))
    .addSubMenu(ui.createMenu('🧹 Data Cleanup')
      .addItem('Cleanup History Data (90 days)', 'cleanupHistoryData')
      .addItem('Cleanup VAPT History Data (90 days)', 'cleanupVAPTHistoryData'))
    .addSubMenu(ui.createMenu('⚙️ Settings')
      .addItem('Set Web App Dashboard URL', 'menuSetWebAppUrl')
      .addItem('Set Automation Ingest Token', 'menuSetAutomationIngestToken'))
    .addToUi();
}

/**
 * Menu function to open Web App Dashboard in new tab
 */
function menuOpenWebAppDashboard() {
  const webAppUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';
  const ui = SpreadsheetApp.getUi();

  if (!webAppUrl) {
    const response = ui.alert(
      '📊 Web App Dashboard URL Not Set',
      'Web App Dashboard URL has not been configured yet.\n\n' +
      'Would you like to set it now?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      menuSetWebAppUrl();
    }
    return;
  }

  // Show URL in alert box for user to copy and open in browser
  ui.alert(
    '📊 Interactive Dashboard',
    'Copy URL berikut dan buka di browser:\n\n' + webAppUrl + '\n\n' +
    '💡 Tip: Klik link di tab Overview (row 1) untuk langsung membuka dashboard.',
    ui.ButtonSet.OK
  );
}

/**
 * Menu function to set Web App Dashboard URL
 * URL will be displayed as a hyperlink in Overview tab row 1
 */
function menuSetWebAppUrl() {
  const ui = SpreadsheetApp.getUi();

  // Get current URL
  const currentUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';

  const promptText = currentUrl
    ? 'Current URL:\n' + currentUrl + '\n\nEnter new Web App Dashboard URL (or leave blank to clear):'
    : 'Enter Web App Dashboard URL:';

  const response = ui.prompt(
    '📊 Set Web App Dashboard URL',
    promptText,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const newUrl = response.getResponseText().trim();

    if (newUrl) {
      // Validate URL format
      if (!newUrl.startsWith('http://') && !newUrl.startsWith('https://')) {
        ui.alert('❌ Invalid URL', 'URL must start with http:// or https://', ui.ButtonSet.OK);
        return;
      }

      PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', newUrl);
      ui.alert(
        '✅ URL Set Successfully!',
        'Web App Dashboard URL has been set.\n\n' +
        'Refresh data to update the link in Overview tab.',
        ui.ButtonSet.OK
      );

      // Refresh Overview tab header to show new link
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const ws = ss.getSheetByName('Overview');
      if (ws) {
        initOverviewHeaders_(ws);
      }
    } else {
      // Clear URL
      PropertiesService.getScriptProperties().deleteProperty('WEB_APP_URL');
      ui.alert('✅ URL Cleared', 'Web App Dashboard URL has been removed.', ui.ButtonSet.OK);

      // Refresh Overview tab header
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const ws = ss.getSheetByName('Overview');
      if (ws) {
        initOverviewHeaders_(ws);
      }
    }
  }
}

function menuSetAutomationIngestToken() {
  const ui = SpreadsheetApp.getUi();
  const currentToken = PropertiesService.getScriptProperties().getProperty('AUTOMATION_INGEST_TOKEN') || '';
  const response = ui.prompt(
    'Set Automation Ingest Token',
    (currentToken ? 'Current token is already configured.\n\n' : '') +
      'Enter the shared token Jenkins will send as DASHBOARD_INGEST_TOKEN.\nLeave blank to clear it.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const token = response.getResponseText().trim();
  if (token) {
    PropertiesService.getScriptProperties().setProperty('AUTOMATION_INGEST_TOKEN', token);
    ui.alert('Automation ingestion token saved.');
  } else {
    PropertiesService.getScriptProperties().deleteProperty('AUTOMATION_INGEST_TOKEN');
    ui.alert('Automation ingestion token cleared.');
  }
}

function menuSetupExternalQAConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName('Config');
    if (!ws) throw new Error('Config tab not found. Please run Create Dashboard first.');
    setupExternalQAConfigSection_(ws);
    safeAlert_('External QA Config ready at Config!AL:AN.\nExisting values were preserved.');
  } catch (error) {
    safeAlert_('Setup External QA Config failed:\n' + error.message);
    Logger.log('Setup External QA Config error: ' + error.stack);
  }
}

/**
 * Quick Start Guide - REMOVED
 * Reason: showModalDialog() requires special permissions that cause authorization issues
 * Alternative: Users can refer to README.md or GitHub documentation
 */

/**
 * Setup Notifications - Guided wizard for Google Chat & Email notifications
 */
function setupNotifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');

  if (!cfg) {
    SpreadsheetApp.getUi().alert('❌ Config tab not found!\n\nPlease run "Create Dashboard" first.');
    return;
  }

  const ui = SpreadsheetApp.getUi();

  const msg =
    '⚠️ DEPRECATED MENU\n\n' +
    'Menu ini sudah digabung dengan Setup Notifications yang baru.\n\n' +
    'Silakan gunakan menu:\n' +
    '🎯 QA Dashboard > 🔔 Notifications > ⚙️ Setup Notifications\n\n' +
    'Menu baru sudah include:\n' +
    '• Setup instruksi lengkap\n' +
    '• Config Google Chat, Email, WhatsApp\n' +
    '• Create trigger otomatis\n\n' +
    'Buka Config tab untuk manual config?';

  const response = ui.alert('Setup Notifications', msg, ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    ss.setActiveSheet(cfg);
    ss.setActiveRange(cfg.getRange('L4'));
  }
}

// ═══════════════════════════════════════════════════════════════════════
// SETUP & REFRESH
// ═══════════════════════════════════════════════════════════════════════

/**
 * Create Dashboard from scratch
 *
 * INCLUDES ALL LATEST FIXES:
 * ✅ "Failure Scenario" tab (not "Blockers") - fix 12-maret
 * ✅ Comprehensive notes on all headers - fix 12-maret
 * ✅ No notes on data rows - fix 12-maret
 * ✅ Submodul sync from Jira - fix 12-maret-2 (via JiraSync.js)
 *
 * For EXISTING dashboards created with old code, run:
 * - fixDashboard12Maret() - untuk rename & notes fixes
 * - fixSubmodulJiraSync() - untuk Submodul field sync dari Jira
 */
function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Cleanup old tabs (including old naming: 'Blockers', 'Scenario Failure')
  ['Overview','Bugs','VAPT','VAPT History','Smoke','Failure Scenario','Scenario Failure','Blockers','Coverage','History','Automation Runs','_Raw','Config','Credentials'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  // Build all tabs with latest structure (no flush - avoid timeout)
  buildConfig(ss);
  buildCredentials(ss);
  buildOverview(ss);
  buildBugs(ss);  // NEW: Bugs tab with historical tracking
  buildVAPT(ss);  // NEW: VAPT findings tracking
  buildVAPTHistory(ss);  // NEW: VAPT history for trendline
  buildSmoke(ss);
  buildFailureScenario(ss);  // Uses "Failure Scenario" naming
  buildCoverage(ss);
  buildHistory(ss);
  buildAutomationRuns(ss);
  buildRaw(ss);

  // Notes are added by init*Headers_() functions in each build*() function
  // addNotesToDashboard(); // REMOVED - function deleted during cleanup

  ss.setActiveSheet(ss.getSheetByName('Config'));
  safeAlert_(
    'Dashboard berhasil dibuat!\n\n' +
    'Langkah selanjutnya:\n' +
    '1. Isi tab Config dengan Spreadsheet ID modul\n' +
    '2. Isi tab Credentials dengan Jira credentials\n' +
    '3. Data Modul/Submodul/QA Lead/PIC QA akan otomatis dari QATM Summary\n' +
    '4. Jalankan refreshDashboard()\n\n' +
    '=== WEB APP SETUP (opsional) ===\n' +
    '5. Deploy as Web App (lihat DEPLOYMENT_NOTES.md)\n' +
    '6. Jalankan setupWebAppUrl() dan paste deployment URL\n' +
    '7. Web App URL akan muncul di WhatsApp notifications\n\n' +
    'PENTING: Deploy Web App dari Testing Script, BUKAN Production!\n' +
    '(Production Script error saat deployment)'
  );
}

/**
 * STEP-BY-STEP CREATE DASHBOARD (untuk debug timeout issues)
 *
 * Jika createDashboard() timeout, jalankan step by step:
 * 1. step1_deleteOldSheets()
 * 2. step2_createConfigAndCredentials()
 * 3. step3_createDataTabs()
 * 4. step4_addNotes()
 */

// STEP 1: Delete all old sheets

function refreshDashboard() {
  const startTime = new Date();
  Logger.log('refreshDashboard START: ' + startTime);

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const modules = getModuleList_(ss);
  if (modules.length === 0) {
    safeAlert_('Belum ada modul aktif di Config.\nIsi tab Config dulu lalu refresh.');
    return;
  }

  // ═══ STEP 1: Pull Module Data ═══
  let t1 = new Date();
  const allData = [];
  modules.forEach(mod => {
    Logger.log('Pulling: ' + mod.name + ' [' + mod.id + ']');
    try {
      allData.push(pullModuleData_(mod));
      Logger.log('OK: ' + mod.name);
    } catch(e) {
      Logger.log('ERROR ' + mod.name + ': ' + e.message);
      allData.push(emptyModuleData_(mod, 'ERROR: ' + e.message));
    }
    Utilities.sleep(150);
  });
  Logger.log('⏱️  Pull module data: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  // ═══ STEP 2: Write Tabs ═══
  t1 = new Date();
  writeOverview(ss, allData);
  Logger.log('⏱️  writeOverview: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  writeBugs(ss, allData);
  Logger.log('⏱️  writeBugs: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  refreshVAPTData();
  Logger.log('⏱️  refreshVAPTData: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  writeSmoke(ss, allData);
  Logger.log('⏱️  writeSmoke: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  writeFailureScenario(ss, allData);
  Logger.log('⏱️  writeFailureScenario: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  writeCoverage(ss, allData);
  Logger.log('⏱️  writeCoverage: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  appendHistory(ss, allData);
  Logger.log('⏱️  appendHistory: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  updateRaw(ss, allData);
  Logger.log('⏱️  updateRaw: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  updateConfig(ss, allData);
  Logger.log('⏱️  updateConfig: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  // ═══ STEP 3: Update Timestamps ═══
  const ts = 'Last refreshed: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss');
  ['Overview','Bugs','VAPT','Smoke'].forEach(name => {
    const sh = ss.getSheetByName(name);
    // Overview has web app link in row 1, "Last refreshed" is in row 2
    if (sh) sh.getRange(name === 'Overview' ? 2 : 1, 1).setValue(ts);
  });

  const totalTime = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('refreshDashboard DONE');
  Logger.log('⏱️  TOTAL TIME: ' + totalTime + 's');
  safeAlert_('Refresh selesai! ' + allData.length + ' modul di-update.\n\nTotal time: ' + totalTime + 's');
}

/**
 * Refresh Bug Only (No VAPT refresh, No Jira Sync)
 * Updates all tabs except VAPT, same as refreshDashboard but skips VAPT data
 *
 * What gets updated:
 * - Overview, Bugs, Smoke, Failure Scenario, Coverage tabs
 * - History (append new snapshot)
 * - Config (Project/Module/Submodule/PIC QA from QATM Summary)
 * - _Raw cache
 *
 * What gets skipped:
 * - VAPT data refresh
 * - Jira sync
 */
function refreshBugOnly() {
  const startTime = new Date();
  Logger.log('refreshBugOnly START: ' + startTime);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modules = getModuleList_(ss);
  if (modules.length === 0) {
    safeAlert_('Belum ada modul aktif di Config.\nIsi tab Config dulu lalu refresh.');
    return;
  }

  // ═══ STEP 1: Pull Module Data ═══
  let t1 = new Date();
  const allData = [];
  modules.forEach(mod => {
    Logger.log('Pulling: ' + mod.name + ' [' + mod.id + ']');
    try {
      allData.push(pullModuleData_(mod));
      Logger.log('OK: ' + mod.name);
    } catch(e) {
      Logger.log('ERROR ' + mod.name + ': ' + e.message);
      allData.push(emptyModuleData_(mod, 'ERROR: ' + e.message));
    }
    Utilities.sleep(150);
  });
  Logger.log('⏱️  Pull module data: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  // ═══ STEP 2: Write Tabs (Same as refreshDashboard, except VAPT) ═══
  t1 = new Date();
  writeOverview(ss, allData);
  Logger.log('⏱️  writeOverview: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  writeBugs(ss, allData);
  Logger.log('⏱️  writeBugs: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  // Skip refreshVAPTData() - this is Bug Only refresh

  t1 = new Date();
  writeSmoke(ss, allData);
  Logger.log('⏱️  writeSmoke: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  writeFailureScenario(ss, allData);
  Logger.log('⏱️  writeFailureScenario: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  writeCoverage(ss, allData);
  Logger.log('⏱️  writeCoverage: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  appendHistory(ss, allData);
  Logger.log('⏱️  appendHistory: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  updateRaw(ss, allData);
  Logger.log('⏱️  updateRaw: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  t1 = new Date();
  updateConfig(ss, allData);
  Logger.log('⏱️  updateConfig: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  // ═══ STEP 3: Update Timestamp ═══
  const ts = 'Last refreshed: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss');
  ['Overview', 'Bugs', 'Smoke'].forEach(name => {
    const sh = ss.getSheetByName(name);
    // Overview has web app link in row 1, "Last refreshed" is in row 2
    if (sh) sh.getRange(name === 'Overview' ? 2 : 1, 1).setValue(ts);
  });

  const totalTime = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('refreshBugOnly DONE');
  Logger.log('⏱️  TOTAL TIME: ' + totalTime + 's');
  safeAlert_('Refresh Bug selesai! ' + allData.length + ' modul di-update.\n\nTotal time: ' + totalTime + 's');
}

/**
 * Refresh VAPT Only (No bug data refresh)
 * Use this when you only need to update VAPT findings
 */
function refreshVAPTOnly() {
  const startTime = new Date();
  Logger.log('refreshVAPTOnly START: ' + startTime);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ═══ Refresh VAPT Data ═══
  let t1 = new Date();
  refreshVAPTData();
  Logger.log('⏱️  refreshVAPTData: ' + ((new Date() - t1) / 1000).toFixed(1) + 's');

  // ═══ Update Timestamp ═══
  const ts = 'Last refreshed: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss');
  const vaptSheet = ss.getSheetByName('VAPT');
  if (vaptSheet) vaptSheet.getRange(1,1).setValue(ts);

  const totalTime = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('refreshVAPTOnly DONE');
  Logger.log('⏱️  TOTAL TIME: ' + totalTime + 's');
  safeAlert_('Refresh VAPT selesai!\n\nTotal time: ' + totalTime + 's');
}

/**
 * Sequential execution: Jira Sync → Dashboard Refresh
 * This ensures dashboard always has fresh data from Jira
 */
function refreshDashboardWithJiraSync() {
  Logger.log('═══════════════════════════════════════════════════');
  Logger.log('START: Sequential Jira Sync + Dashboard Refresh');
  Logger.log('═══════════════════════════════════════════════════');

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // ── Step 1: Jira Sync ────────────────────────────────────────────────
    Logger.log('');
    Logger.log('📊 STEP 1/2: Syncing Jira for modules with Jira Sync = TRUE...');
    Logger.log('─────────────────────────────────────────────────────');

    // Check if _runSync_ exists (from JiraSync.js)
    if (typeof _runSync_ !== 'function') {
      Logger.log('⚠️  Jira sync skipped - _runSync_ function not available');
      Logger.log('   (This is normal if Jira integration is not set up)');
    } else {
      // _runSync_ will only sync modules with Active = TRUE and Jira Sync = TRUE
      const syncResults = _runSync_(ss, false); // Sync Title, Desc, Priority, Assignee, Submodul
      Logger.log('✅ Jira sync completed:');
      syncResults.forEach(r => Logger.log('   ' + r));
    }

    // ── Step 2: Dashboard Refresh ────────────────────────────────────────
    Logger.log('');
    Logger.log('🔄 STEP 2/2: Refreshing Dashboard (all active modules)...');
    Logger.log('─────────────────────────────────────────────────────');

    refreshDashboard();

    Logger.log('');
    Logger.log('═══════════════════════════════════════════════════');
    Logger.log('✅ DONE: Sequential Jira Sync + Dashboard Refresh');
    Logger.log('═══════════════════════════════════════════════════');

  } catch(e) {
    Logger.log('❌ ERROR in sequential sync: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    safeAlert_('❌ Error during sync + refresh:\n\n' + e.message + '\n\nCheck Executions log for details.');
  }
}

function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');

  if (!cfg) {
    safeAlert_('❌ Config tab not found!\n\nPlease run "Create Dashboard" first.');
    return;
  }

  // Read refresh config from Q4-R4
  const refreshInterval = parseInt(cfg.getRange(4, 17).getValue()) || 10; // Q4 - default 10 minutes
  const enabledVal = cfg.getRange(4, 18).getValue(); // R4
  const enabled = (typeof enabledVal === 'boolean') ? enabledVal : false;

  // Delete existing triggers (both old and new function names)
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'refreshDashboard' || fn === 'refreshDashboardWithJiraSync') {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (!enabled) {
    safeAlert_('ℹ️  Auto-refresh trigger dihapus.\n\nSet "Enable Auto Refresh" = TRUE di Config (kolom R4) untuk aktifkan kembali.');
    Logger.log('Auto-refresh trigger removed (disabled in config)');
    return;
  }

  // Validate interval (1-60 minutes)
  if (refreshInterval < 1 || refreshInterval > 60) {
    safeAlert_('❌ Invalid interval!\n\nRefresh interval harus antara 1-60 menit.\nCurrent value: ' + refreshInterval);
    return;
  }

  // Create new trigger - SEQUENTIAL: Jira Sync → Dashboard Refresh
  ScriptApp.newTrigger('refreshDashboardWithJiraSync')
    .timeBased()
    .everyMinutes(refreshInterval)
    .create();

  Logger.log('✅ Trigger created for sequential auto-refresh every ' + refreshInterval + ' minutes');
  safeAlert_(
    '✅ Auto-Refresh Trigger Aktif!\n\n' +
    '⏱️  Interval: Setiap ' + refreshInterval + ' menit\n\n' +
    '🔄 Execution Flow (Sequential):\n' +
    '  1️⃣ Jira Sync → Sync data dari Jira ke QATM modules\n' +
    '  2️⃣ Dashboard Refresh → Pull data terbaru ke Dashboard\n\n' +
    '📝 Config: Q4 (interval) | R4 (enable/disable)\n\n' +
    '💡 Tip: Dashboard akan selalu punya data terbaru dari Jira!'
  );
}

// ═══════════════════════════════════════════════════════════════════════════
// BUTTON FUNCTIONS - One-click operations
// ═══════════════════════════════════════════════════════════════════════════

/**
 * 🔘 BUTTON 1: One-click setup auto-refresh trigger
 * Reads config from Q4-R4, creates time-based trigger
 */
function setupAutoRefreshTrigger() {
  Logger.log('Button clicked: Setup Auto-Refresh Trigger');
  setupTrigger();
}

/**
 * 🔘 BUTTON 2: One-click manual sync + refresh
 * Executes sequential: Jira Sync → Dashboard Refresh
 */
function manualSyncAndRefresh() {
  Logger.log('Button clicked: Manual Sync + Refresh');

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🔄 Manual Sync + Refresh',
    'Ini akan menjalankan:\n\n' +
    '1️⃣ Jira Sync → Sync semua QATM modules dari Jira\n' +
    '2️⃣ Dashboard Refresh → Update semua dashboard tabs\n\n' +
    'Proses bisa memakan waktu beberapa menit.\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    Logger.log('Manual sync cancelled by user');
    return;
  }

  // Execute sequential sync + refresh
  refreshDashboardWithJiraSync();
}

function safeAlert_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

/**
 * Helper function to rebuild a tab (delete old + create new)
 */
function rebuildTab_(ss, tabName, buildFunction, successMessage) {
  // Delete old tab if exists
  const oldSheet = ss.getSheetByName(tabName);
  if (oldSheet) {
    ss.deleteSheet(oldSheet);
  }

  // Create new tab
  buildFunction(ss);

  // Show success message
  safeAlert_(successMessage || tabName + ' tab rebuilt!');
}

// Wrapper functions for manual execution from menu
function rebuildConfig() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'Config', buildConfig, 'Config tab rebuilt!');
}
function rebuildCredentials() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'Credentials', buildCredentials, 'Credentials tab rebuilt!');
}
function rebuildOverview() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'Overview', buildOverview, 'Overview tab rebuilt!');
}
function rebuildBugs() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'Bugs', buildBugs, 'Bugs tab rebuilt!');
}
function rebuildVAPT() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'VAPT', buildVAPT, 'VAPT tab rebuilt!\n\nStructure: Per-Project (8 columns)');
}
function rebuildVAPTHistory() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'VAPT History', buildVAPTHistory, 'VAPT History tab rebuilt!\n\nStructure: Per-Project tracking');
}
function rebuildSmoke() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'Smoke', buildSmoke, 'Smoke tab rebuilt!');
}
function rebuildFailureScenario() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'Failure Scenario', buildFailureScenario, 'Failure Scenario tab rebuilt!');
}
function rebuildCoverage() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'Coverage', buildCoverage, 'Coverage tab rebuilt!');
}
function rebuildHistory() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), 'History', buildHistory, 'History tab rebuilt!');
}
function rebuildAutomationRuns() {
  rebuildTab_(getActiveOrDashboardSpreadsheet_(), 'Automation Runs', buildAutomationRuns, 'Automation Runs tab rebuilt!');
}
function rebuildRaw() {
  rebuildTab_(SpreadsheetApp.getActiveSpreadsheet(), '_Raw', buildRaw, '_Raw tab rebuilt!');
}

function getActiveOrDashboardSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  return active || getDashboardSpreadsheet_();
}

/**
 * Update Config structure: add Project column (col C) and shift existing data
 * SAFE: Does not delete existing data, only restructures columns
 */

function getModuleList_(ss) {
  const cfg = ss.getSheetByName('Config');
  if (!cfg) return [];
  const data = cfg.getDataRange().getValues();
  const headerRow = data[2] || [];
  const headerIndex = {};
  headerRow.forEach((header, idx) => {
    const key = normalizeAutomationValue_(header);
    if (key) headerIndex[key] = idx;
  });
  const readOptional = (row, names) => {
    for (let i = 0; i < names.length; i++) {
      const idx = headerIndex[normalizeAutomationValue_(names[i])];
      if (idx !== undefined) {
        const value = String(row[idx] || '').trim();
        if (value) return value;
      }
    }
    return '';
  };
  const readBoolOptional = (row, names) => {
    for (let i = 0; i < names.length; i++) {
      const idx = headerIndex[normalizeAutomationValue_(names[i])];
      if (idx !== undefined) {
        const value = row[idx];
        if (value === true || value === false) return value;
        const text = String(value || '').trim().toUpperCase();
        if (text === 'TRUE' || text === 'YES' || text === 'YA') return true;
        if (text === 'FALSE' || text === 'NO' || text === 'TIDAK') return false;
      }
    }
    return false;
  };
  const modules = [];
  for (let i = 3; i < data.length; i++) {
    const active    = data[i][0] === true;  // col A = Active (TRUE/FALSE)
    const jiraSync  = data[i][1] === true;  // col B = Jira Sync (TRUE/FALSE)
    const project   = String(data[i][2]).trim();  // col C = Project
    const module    = String(data[i][3]).trim();  // col D = Modul
    const submodule = String(data[i][4]).trim();  // col E = Submodul
    const team      = String(data[i][5]).trim();  // col F = PIC QA
    const id        = String(data[i][6]).trim();  // col G = Spreadsheet ID
    const jiraInst  = String(data[i][8]).trim();  // col I = Jira Instance
    const jiraProj  = String(data[i][9]).trim();  // col J = Jira Project

    if (!active || !id || id.length < 10 || id === 'PASTE_SPREADSHEET_ID_HERE') continue;

    modules.push({
      name: submodule||module||'Unknown',
      id,
      project,
      module,
      submodule,
      team,
      lead: '',  // QA Lead dihapus dari Config (hanya ada PIC QA)
      active: true,
      jiraSync,
      jiraInst,
      jiraProj,
      externalQA: {
        isExternal: readBoolOptional(data[i], ['External QA', 'External Team']),
        retestByPeruri: readBoolOptional(data[i], ['Retest QA Peruri', 'Need Retest QA Peruri', 'Retest by QA Peruri']),
        notes: readOptional(data[i], ['External QA Notes', 'External Notes'])
      },
      automationContracts: {
        all: readOptional(data[i], ['Automation Contract', 'Automation Key', 'Automation Alias', 'Jenkins Job', 'Jenkins Job Pattern']),
        web: readOptional(data[i], ['Web Automation Contract', 'Web Automation Key', 'Web Jenkins Job', 'Web Jenkins Job Pattern']),
        api: readOptional(data[i], ['API Automation Contract', 'API Automation Key', 'API Jenkins Job', 'API Jenkins Job Pattern'])
      }
    });
  }
  return modules;
}


// ═══════════════════════════════════════════════════════════════════════
// DATA PULL — per module
// ═══════════════════════════════════════════════════════════════════════

function pullModuleData_(mod) {
  const src  = SpreadsheetApp.openById(mod.id);
  const tcm  = src.getSheetByName('TC_Master');
  const tce  = src.getSheetByName('TC_Execution');
  const apim = src.getSheetByName('API_Master');
  const apie = src.getSheetByName('API_Execution');
  const perf = src.getSheetByName('PerfTest');
  const summ = src.getSheetByName('Summary');
  const externalReportSheet = src.getSheetByName('External Test Report');
  const bugr = src.getSheetByName('BugReport');

  const SUMM_KPI_ROW = 13;  // Summary row 13: main Web+API KPI values

  let projectName = '', moduleName = '', submoduleName = '', picQA = mod.team||'', qaLead = mod.lead||'';
  let wTotal=0,wPassed=0,wFailed=0,wBlocked=0,wInProg=0,wTodo=0,wPassRate=0,wAutoRate=0,wExecRate=0;
  let aTotal=0,aPassed=0,aFailed=0,aBlocked=0,aInProg=0,aTodo=0,aPassRate=0,aAutoRate=0,aExecRate=0;
  let wSmokeTotal=0,wSmokePassed=0,wSmokeFailed=0,wSmokeBlocked=0,wSmokeInProg=0,wSmokeTodo=0;
  let wSmokePassRate=0,wSmokeAutoRate=0,wSmokeExecRate=0;
  let aSmokeTotal=0,aSmokePassed=0,aSmokePassRate=0,aSmokeAutoRate=0,aSmokeExecRate=0;
  let perfResult = '--';
  let scopeNotes = '';
  let externalTestReport = emptyExternalTestReport_();

  try {
    if (summ) {
      const proj = summ.getRange(2,2).getValue();  // B2 = Project (baru)
      const modul = summ.getRange(3,2).getValue();  // B3 = Modul (baru)
      const submod = summ.getRange(4,2).getValue();  // B4 = Submodul (baru)
      const ql  = summ.getRange(5,2).getValue();  // B5 = QA Lead (baru, sebelumnya B4)
      const pic = summ.getRange(6,2).getValue();  // B6 = PIC QA (baru, sebelumnya B5)

      if (proj && String(proj).trim())     projectName = String(proj).trim();
      if (modul && String(modul).trim())   moduleName = String(modul).trim();
      if (submod && String(submod).trim()) submoduleName = String(submod).trim();
      if (ql  && String(ql).trim())        qaLead        = String(ql).trim(); // Summary overrides Config
      if (pic && String(pic).trim())       picQA         = String(pic).trim();
      scopeNotes = readSummaryFieldByLabel_(summ, 'Scope / Notes:');

      const perfVal = summ.getRange(8,13).getValue(); // M8 = Perf result
      if (perfVal && String(perfVal).trim()) perfResult = String(perfVal).trim();

      // Web KPI — row 13, cols A-I (1-9)
      const wKpi = summ.getRange(SUMM_KPI_ROW,1,1,9).getValues()[0];
      wTotal=+wKpi[0]||0; wPassed=+wKpi[1]||0; wFailed=+wKpi[2]||0;
      wBlocked=+wKpi[3]||0; wInProg=+wKpi[4]||0; wTodo=+wKpi[5]||0;
      wPassRate=+wKpi[6]||0; wAutoRate=+wKpi[7]||0; wExecRate=+wKpi[8]||0;

      // API KPI — row 13, cols L-T (12-20)
      const aKpi = summ.getRange(SUMM_KPI_ROW,12,1,9).getValues()[0];
      aTotal=+aKpi[0]||0; aPassed=+aKpi[1]||0; aFailed=+aKpi[2]||0;
      aBlocked=+aKpi[3]||0; aInProg=+aKpi[4]||0; aTodo=+aKpi[5]||0;
      aPassRate=+aKpi[6]||0; aAutoRate=+aKpi[7]||0; aExecRate=+aKpi[8]||0;

      Logger.log(mod.name + ' | wTotal=' + wTotal + ' wPass=' + pct_(wPassRate) + ' | aTotal=' + aTotal + ' aPass=' + pct_(aPassRate));

      // Smoke KPI — cari row secara DINAMIS (v2 — paling robust)
      //
      // STRATEGI 2 LANGKAH:
      //   1. textFinder('SMOKE TEST') → dapat posisi smoke header secara pasti
      //      (tidak bisa nyangkut di STATUS OVERVIEW — section itu tidak punya teks ini)
      //   2. Dari header, scan ke bawah (+1 s/d +5): cari row pertama
      //      di mana col A bernilai NUMBER (bukan text label "TOTAL")
      //      → itu adalah baris values yang sesungguhnya
      //
      // Ini menggantikan pendekatan lama (scan label row dari row 14) yang
      // bisa salah tangkap STATUS OVERVIEW label row jika sheet punya extra rows.
      //
      // Layout Summary smoke section (bisa bervariasi per versi broadcast):
      //   Row +0 : header  "A1. SMOKE TEST — Web / Mobile ..."  ← textFinder
      //   Row +1 : labels  TOTAL PASSED ... PASS RATE EXEC RATE  (text, di-skip)
      //   Row +2 : values  5  2  2  0  0  1  40%  80%  80%       ← target
      //  -- atau versi lama tanpa label row:
      //   Row +0 : header
      //   Row +1 : values  ← langsung ketemu di scan pertama
      //
      // Web (L=col 1): A=Total  G=Pass%  I=Exec%
      // API (R=col12): L=Total  R=Pass%  T=Exec%
      try {
        // Cari smokeValRow secara dinamis:
        //
        // KENAPA TIDAK "cari numeric di col A":
        //   Broadcast Fix D merges seluruh row (col 1-10) dengan formula COUNTIFS
        //   → col A di baris Open Blocker = NUMBER (bug count), bukan smoke total.
        //   Scan numeric akan nyangkut di situ.
        //
        // STRATEGI BENAR: cari LABEL ROW (col A = "TOTAL", col G = "PASS RATE"),
        //   lalu smokeValRow = labelRow + 1.
        //
        // Layout Summary setelah broadcast Fix D:
        //   Row 15: SMOKE TEST header
        //   Row 16: "Open Blocker (Smoke) ↓"   → col A = text, skip
        //   Row 17: =COUNTIFS(...)              → col A = NUMBER (bug count), skip ← jebakan
        //   Row 18: TOTAL PASSED... PASS RATE   → col A = "TOTAL" ← cari ini
        //   Row 19: 52 22 0 0 14 0 40% 80% 80% → smokeValRow = 18+1 = 19 ✓
        //
        // Layout fresh sheet (tanpa broadcast Fix D):
        //   Row 15: SMOKE TEST header
        //   Row 16: TOTAL PASSED... PASS RATE   → col A = "TOTAL" ← cari ini
        //   Row 17: 1 ...                       → smokeValRow = 16+1 = 17 ✓
        //
        // Scan dimulai dari row 14 (setelah Status Overview labels di row 12).

        let smokeLabelRow = -1;
        const SCAN_START = 14, SCAN_LEN = 22;  // scan rows 14–35
        const scanGrid = summ.getRange(SCAN_START, 1, SCAN_LEN, 9).getValues();
        for (let i = 0; i < scanGrid.length; i++) {
          const c1 = String(scanGrid[i][0]).trim().toUpperCase();  // col A
          const c7 = String(scanGrid[i][6]).trim().toUpperCase();  // col G
          if (c1 === 'TOTAL' && c7 === 'PASS RATE') {
            smokeLabelRow = SCAN_START + i;
            break;
          }
        }
        if (smokeLabelRow === -1) throw new Error('Smoke label row not found (TOTAL+PASS RATE) in rows 14-35');
        const smokeValRow = smokeLabelRow + 1;
        Logger.log(mod.name + ' | smokeLabelRow=' + smokeLabelRow + ' smokeValRow=' + smokeValRow);

        // Web / Mobile smoke (col 1–9)
        const wSm = summ.getRange(smokeValRow, 1, 1, 9).getValues()[0];
        wSmokeTotal    = +wSm[0]||0;   // col A = TOTAL
        wSmokePassed   = +wSm[1]||0;   // col B = PASSED
        wSmokeFailed   = +wSm[2]||0;   // col C = FAILED
        wSmokeBlocked  = +wSm[3]||0;   // col D = BLOCKED
        wSmokeInProg   = +wSm[4]||0;   // col E = IN PROG
        wSmokeTodo     = +wSm[5]||0;   // col F = TODO
        wSmokePassRate = +wSm[6]||0;   // col G = PASS RATE
        wSmokeAutoRate = +wSm[7]||0;   // col H = AUTO RATE
        wSmokeExecRate = +wSm[8]||0;   // col I = EXEC RATE

        // API smoke (col 12–20, same layout offset R_=12)
        const aSm = summ.getRange(smokeValRow, 12, 1, 9).getValues()[0];
        aSmokeTotal    = +aSm[0]||0;   // col L = TOTAL
        aSmokePassed   = +aSm[1]||0;   // col M = PASSED
        aSmokePassRate = +aSm[6]||0;   // col R = PASS RATE
        aSmokeAutoRate = +aSm[7]||0;   // col S = AUTO RATE
        aSmokeExecRate = +aSm[8]||0;   // col T = EXEC RATE

        Logger.log(mod.name
            + ' | smokeValRow=' + smokeValRow
            + ' | wSmokeTotal=' + wSmokeTotal
            + ' wSmokePass='  + pct_(wSmokePassRate)
            + ' wSmokeExec='  + pct_(wSmokeExecRate)
            + ' | aTotal='    + aSmokeTotal
            + ' aPass='       + pct_(aSmokePassRate));
      } catch(se) {
        Logger.log('Smoke KPI skip [' + mod.name + ']: ' + se.message);
      }

    } else {
      // Fallback: hitung dari raw sheets
      const wS = getSheetStats_(tcm,tce,'TC'), aS = getSheetStats_(apim,apie,'API');
      ({total:wTotal,passed:wPassed,failed:wFailed,blocked:wBlocked,inprog:wInProg,todo:wTodo,passRate:wPassRate,autoRate:wAutoRate,execRate:wExecRate} = wS);
      ({total:aTotal,passed:aPassed,failed:aFailed,blocked:aBlocked,inprog:aInProg,todo:aTodo,passRate:aPassRate,autoRate:aAutoRate,execRate:aExecRate} = aS);
      perfResult = getPerfResult_(perf);
    }
  } catch(e) {
    Logger.log('pullModuleData_ error [' + mod.name + ']: ' + e.message);
    try {
      const wS = getSheetStats_(tcm,tce,'TC'), aS = getSheetStats_(apim,apie,'API');
      wTotal=wS.total; wPassed=wS.passed; wFailed=wS.failed; wBlocked=wS.blocked; wPassRate=wS.passRate;
      aTotal=aS.total; aPassed=aS.passed; aFailed=aS.failed; aBlocked=aS.blocked; aPassRate=aS.passRate;
    } catch(e2) {}
  }
  externalTestReport = readExternalTestReport_(externalReportSheet);

  return {
    name:mod.name, team:picQA, lead:qaLead, id:mod.id,
    project:projectName, module:moduleName, submodule:submoduleName,
    automationContracts: mod.automationContracts || {},
    externalQA: mod.externalQA || {isExternal:false,retestByPeruri:false,notes:''},
    scopeNotes,
    externalTestReport,
    refreshed:new Date(),
    wTotal,wPassed,wFailed,wBlocked,wInProg,wTodo,wPassRate,wAutoRate,wExecRate,
    aTotal,aPassed,aFailed,aBlocked,aInProg,aTodo,aPassRate,aAutoRate,aExecRate,
    wSmokeTotal,wSmokePassed,wSmokeFailed,wSmokeBlocked,wSmokeInProg,wSmokeTodo,
    wSmokePassRate,wSmokeAutoRate,wSmokeExecRate,
    aSmokeTotal,aSmokePassed,aSmokePassRate,aSmokeAutoRate,aSmokeExecRate,
    perfResult,
    blockers: getBlockers_(tcm,tce,apim,apie,mod.name),
    coverage: getCoverage_(tcm,tce,apim,apie),
    bugStats: getBugStats_(bugr, summ),
    error: '',
  };
}

function emptyModuleData_(mod, errorMsg) {
  return {
    name:mod.name,team:mod.team||'',lead:mod.lead||'',id:mod.id,
    sprint:'',project:mod.project||'',module:mod.module||'',submodule:mod.submodule||'',
    automationContracts: mod.automationContracts || {},
    externalQA: mod.externalQA || {isExternal:false,retestByPeruri:false,notes:''},
    scopeNotes:'',
    externalTestReport: emptyExternalTestReport_(),
    refreshed:new Date(),error:errorMsg,
    wTotal:0,wPassed:0,wFailed:0,wBlocked:0,wInProg:0,wTodo:0,wPassRate:0,wAutoRate:0,wExecRate:0,
    aTotal:0,aPassed:0,aFailed:0,aBlocked:0,aInProg:0,aTodo:0,aPassRate:0,aAutoRate:0,aExecRate:0,
    wSmokeTotal:0,wSmokePassed:0,wSmokeFailed:0,wSmokeBlocked:0,wSmokeInProg:0,wSmokeTodo:0,
    wSmokePassRate:0,wSmokeAutoRate:0,wSmokeExecRate:0,
    aSmokeTotal:0,aSmokePassed:0,aSmokePassRate:0,aSmokeAutoRate:0,aSmokeExecRate:0,
    perfResult:'--',blockers:[],coverage:[],
    bugStats:{total:0,open:0,inprog:0,fixed:0,verified:0,critical:0,high:0,medium:0,low:0,blocker:0},
  };
}

function readSummaryFieldByLabel_(summarySheet, label) {
  if (!summarySheet) return '';
  try {
    const cell = summarySheet.createTextFinder(label).matchEntireCell(true).findNext();
    if (!cell) return '';
    const value = summarySheet.getRange(cell.getRow(), cell.getColumn() + 1).getDisplayValue();
    return String(value || '').trim();
  } catch(e) {
    return '';
  }
}

function emptyExternalTestReport_() {
  return {
    externalTeam: '',
    statusReview: '',
    functionalEvidenceUrl: '',
    functionalReviewStatus: '',
    performanceEvidenceUrl: '',
    performanceReviewStatus: '',
    vaptEvidenceUrl: '',
    vaptReviewStatus: '',
    overallStatus: '',
    reviewer: '',
    reviewDate: '',
    notes: ''
  };
}

function readExternalTestReport_(sheet) {
  const report = emptyExternalTestReport_();
  if (!sheet) return report;
  const mapping = {
    'External Team / Vendor:': 'externalTeam',
    'Status Review:': 'statusReview',
    'Functional Evidence URL:': 'functionalEvidenceUrl',
    'Functional Review Status:': 'functionalReviewStatus',
    'Performance Evidence URL:': 'performanceEvidenceUrl',
    'Performance Review Status:': 'performanceReviewStatus',
    'VAPT Evidence URL:': 'vaptEvidenceUrl',
    'VAPT Review Status:': 'vaptReviewStatus',
    'Overall Status:': 'overallStatus',
    'Reviewer:': 'reviewer',
    'Review Date:': 'reviewDate',
    'Notes:': 'notes'
  };
  Object.keys(mapping).forEach(label => {
    try {
      const cell = sheet.createTextFinder(label).matchEntireCell(true).findNext();
      if (!cell) return;
      report[mapping[label]] = String(sheet.getRange(cell.getRow(), cell.getColumn() + 1).getDisplayValue() || '').trim();
    } catch(e) {}
  });
  return report;
}


// ═══════════════════════════════════════════════════════════════════════
// STAT HELPERS
// ═══════════════════════════════════════════════════════════════════════

function getSheetStats_(masterSheet, execSheet, type) {
  const empty = {total:0,passed:0,failed:0,blocked:0,inprog:0,todo:0,passRate:0,autoRate:0,execRate:0};
  if (!masterSheet || !execSheet) return empty;
  try {
    const mData = masterSheet.getDataRange().getValues().slice(2);
    const eData = execSheet.getDataRange().getValues().slice(8);
    const statusMap = {};
    eData.forEach(r => { if(r[0]&&r[25]) statusMap[r[0]] = String(r[25]).trim(); });
    const autoCol = type==='TC' ? 7 : 9;
    let total=0,passed=0,failed=0,blocked=0,inprog=0,todo=0,auto=0;
    mData.forEach(r => {
      if (!r[2]) return; total++;
      if (r[autoCol]==='Automated') auto++;
      const st = statusMap[r[2]]||'TODO';
      if (st==='PASSED') passed++;
      else if (st==='FAILED') failed++;
      else if (st==='BLOCKED') blocked++;
      else if (st==='IN PROGRESS') inprog++;
      else todo++;
    });
    return {total,passed,failed,blocked,inprog,todo,
      passRate:total?passed/total:0, autoRate:total?auto/total:0,
      execRate:total?(passed+failed+blocked+inprog)/total:0};
  } catch(e) { return empty; }
}

function getBlockers_(tcm, tce, apim, apie, moduleName) {
  const bl = [];
  try {
    if (tcm&&tce) {
      const sm={}; tce.getDataRange().getValues().slice(8).forEach(r=>{if(r[0]&&r[25])sm[r[0]]=r[25];});
      tcm.getDataRange().getValues().slice(2).forEach(r=>{
        if(!r[2]||(r[4]!=='Critical'&&r[4]!=='High'))return;
        const st=sm[r[2]]||'TODO';
        if(st==='FAILED'||st==='BLOCKED') bl.push({module:moduleName,type:'Web',tcId:r[2],prio:r[4],feature:r[3],scenario:String(r[10]).substring(0,80),status:st});
      });
    }
  } catch(e) {}
  try {
    if (apim&&apie) {
      const sm={}; apie.getDataRange().getValues().slice(8).forEach(r=>{if(r[0]&&r[25])sm[r[0]]=r[25];});
      apim.getDataRange().getValues().slice(2).forEach(r=>{
        if(!r[2]||(r[6]!=='Critical'&&r[6]!=='High'))return;
        const st=sm[r[2]]||'TODO';
        if(st==='FAILED'||st==='BLOCKED') bl.push({module:moduleName,type:'API',tcId:r[2],prio:r[6],feature:r[3],scenario:String(r[12]).substring(0,80),status:st});
      });
    }
  } catch(e) {}
  return bl;
}

function getCoverage_(tcm, tce, apim, apie) {
  const cov=[];
  try {
    if (tcm) {
      const sm={}; (tce?tce.getDataRange().getValues().slice(8):[]).forEach(r=>{if(r[0]&&r[25])sm[r[0]]=r[25];});
      const map={};
      tcm.getDataRange().getValues().slice(2).forEach(r=>{
        if(!r[2]||!r[1])return;
        if(!map[r[1]])map[r[1]]={sub:r[1],total:0,passed:0,failed:0,auto:0,type:'Web'};
        map[r[1]].total++;
        if(r[7]==='Automated')map[r[1]].auto++;
        const st=sm[r[2]]||'';
        if(st==='PASSED')map[r[1]].passed++;
        if(st==='FAILED')map[r[1]].failed++;
      });
      Object.values(map).forEach(v=>cov.push(v));
    }
  } catch(e) {}
  try {
    if (apim) {
      const sm={}; (apie?apie.getDataRange().getValues().slice(8):[]).forEach(r=>{if(r[0]&&r[25])sm[r[0]]=r[25];});
      const map={};
      apim.getDataRange().getValues().slice(2).forEach(r=>{
        if(!r[2]||!r[1])return;
        if(!map[r[1]])map[r[1]]={sub:r[1],total:0,passed:0,failed:0,auto:0,type:'API'};
        map[r[1]].total++;
        if(r[9]==='Automated')map[r[1]].auto++;
        const st=sm[r[2]]||'';
        if(st==='PASSED')map[r[1]].passed++;
        if(st==='FAILED')map[r[1]].failed++;
      });
      Object.values(map).forEach(v=>cov.push(v));
    }
  } catch(e) {}
  return cov;
}

function getPerfResult_(perfSheet) {
  if (!perfSheet) return '--';
  try {
    const data = perfSheet.getDataRange().getValues().slice(15,45);
    if (!data.some(r=>r[4]&&r[4]!=='')) return '--';
    return data.some(r=>r[11]==='FAIL') ? 'FAIL' : 'PASS';
  } catch(e) { return '--'; }
}

function getBugStats_(bugSheet, summarySheet) {
  const empty={
    total:0,critical:0,high:0,medium:0,low:0,lowest:0,blocker:0,
    devBugs:0,uatBugs:0,prodBugs:0,
    open:0,inprog:0,fixed:0,reopen:0,verified:0,
    blockerOpenBugs:0,blockerInProgressBugs:0,blockerFixedBugs:0,blockerReopenBugs:0,blockerVerifiedBugs:0,
    prodBlockerBugs:0,prodCriticalBugs:0,prodHighBugs:0,prodMediumBugs:0,prodLowBugs:0,prodLowestBugs:0,
    prodOpenBugs:0,prodInProgressBugs:0,prodFixedBugs:0,prodReopenBugs:0,prodVerifiedBugs:0
  };
  if (!bugSheet) return empty;
  try {
    const rows=bugSheet.getDataRange().getValues().slice(4).filter(r=>r[0]&&r[0]!=='');
    const cnt=(fn)=>rows.filter(fn).length;
    const statusOf = r => String(r[3] || '').toLowerCase().trim();
    const priorityOf = r => String(r[2] || '').trim();
    const envOf = r => String(r[8] || '').toLowerCase().trim();
    const isActive = r => statusOf(r) !== 'closed';
    const isWontFix = r => statusOf(r) === "won't fix";
    const isBlockerPriority = r => ['Critical','Highest','High','Medium'].includes(priorityOf(r));
    const isProduction = r => {
      const env = envOf(r);
      return isActive(r) && (env === 'production' || env === 'prod');
    };
    const hasStatus = (r, statuses) => statuses.includes(statusOf(r));

    // Get blocker count from Summary sheet (more accurate - uses formula)
    let blockerCount = cnt(r=>isActive(r) && !isWontFix(r) && isBlockerPriority(r));
    let prodBugsCount = cnt(isProduction);

    // Try to read from Summary sheet if available
    if (summarySheet) {
      try {
        // Read BUG BLOCKER
        const blockerCell = summarySheet.createTextFinder('Open Blocker:').matchEntireCell(false).findNext();
        if (blockerCell) {
          const blockerRow = blockerCell.getRow();
          const blockerValue = summarySheet.getRange(blockerRow, 2).getValue();  // Same row, col B
          if (typeof blockerValue === 'number' && blockerValue >= 0) {
            blockerCount = blockerValue;
          }
        }

        // Read PROD BUGS
        const prodCell = summarySheet.createTextFinder('PROD BUGS:').matchEntireCell(false).findNext();
        if (prodCell) {
          const prodRow = prodCell.getRow();
          const prodValue = summarySheet.getRange(prodRow, 2).getValue();  // Same row, col B
          if (typeof prodValue === 'number' && prodValue >= 0) {
            prodBugsCount = prodValue;
          }
        }
      } catch(e) {
        Logger.log('Failed to read bug metrics from Summary: ' + e.message);
      }
    }

    return {
      total:   cnt(isActive),
      critical:cnt(r=>isActive(r) && (priorityOf(r)==='Critical' || priorityOf(r)==='Highest')),
      high:    cnt(r=>isActive(r) && priorityOf(r)==='High'),
      medium:  cnt(r=>isActive(r) && priorityOf(r)==='Medium'),
      low:     cnt(r=>isActive(r) && priorityOf(r)==='Low'),
      lowest:  cnt(r=>isActive(r) && priorityOf(r)==='Lowest'),
      blocker: blockerCount,
      devBugs: cnt(r=>isActive(r) && (envOf(r)==='development' || envOf(r)==='dev')),
      uatBugs: cnt(r=>isActive(r) && envOf(r)==='uat'),
      prodBugs: prodBugsCount,
      open:    cnt(r=>hasStatus(r, ['open'])),
      inprog:  cnt(r=>hasStatus(r, ['in progress','in progress vapt'])),
      fixed:   cnt(r=>hasStatus(r, ['fixed'])),
      reopen:  cnt(r=>hasStatus(r, ['reopen'])),
      verified:cnt(r=>hasStatus(r, ['verified','ready to test','done vapt'])),
      blockerOpenBugs:       cnt(r=>isBlockerPriority(r) && hasStatus(r, ['open'])),
      blockerInProgressBugs: cnt(r=>isBlockerPriority(r) && hasStatus(r, ['in progress','in progress vapt'])),
      blockerFixedBugs:      cnt(r=>isBlockerPriority(r) && hasStatus(r, ['fixed'])),
      blockerReopenBugs:     cnt(r=>isBlockerPriority(r) && hasStatus(r, ['reopen'])),
      blockerVerifiedBugs:   cnt(r=>isBlockerPriority(r) && hasStatus(r, ['verified','ready to test','done vapt'])),
      prodBlockerBugs:       cnt(r=>isProduction(r) && !isWontFix(r) && isBlockerPriority(r)),
      prodCriticalBugs:      cnt(r=>isProduction(r) && (priorityOf(r)==='Critical' || priorityOf(r)==='Highest')),
      prodHighBugs:          cnt(r=>isProduction(r) && priorityOf(r)==='High'),
      prodMediumBugs:        cnt(r=>isProduction(r) && priorityOf(r)==='Medium'),
      prodLowBugs:           cnt(r=>isProduction(r) && priorityOf(r)==='Low'),
      prodLowestBugs:        cnt(r=>isProduction(r) && priorityOf(r)==='Lowest'),
      prodOpenBugs:          cnt(r=>isProduction(r) && hasStatus(r, ['open'])),
      prodInProgressBugs:    cnt(r=>isProduction(r) && hasStatus(r, ['in progress','in progress vapt'])),
      prodFixedBugs:         cnt(r=>isProduction(r) && hasStatus(r, ['fixed'])),
      prodReopenBugs:        cnt(r=>isProduction(r) && hasStatus(r, ['reopen'])),
      prodVerifiedBugs:      cnt(r=>isProduction(r) && hasStatus(r, ['verified','ready to test','done vapt'])),
    };
  } catch(e) { return empty; }
}

function pct_(v) { return Math.round((v||0)*100)+'%'; }


// ═══════════════════════════════════════════════════════════════════════
// CONFIG TAB
// ═══════════════════════════════════════════════════════════════════════

function buildConfig(ss) {
  const ws = ss.insertSheet('Config');
  ws.setTabColor('#37474F');
  ws.clear();
  function hdr(r,c,txt,w,note){
    const cell = ws.getRange(r,c);
    cell.setValue(txt).setBackground('#0D47A1').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setColumnWidth(c,w);
    if(note) cell.setNote(note);
  }
  ws.getRange(1,1,1,10).merge().setValue('QA PORTFOLIO DASHBOARD  —  Module Config')
      .setBackground('#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,32);
  ws.getRange(2,1,1,10).merge()
      .setValue('Spreadsheet ID ada di URL modul: docs.google.com/spreadsheets/d/[ID]/edit  |  Data akan otomatis diambil dari QATM Summary saat refresh')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic').setFontSize(8).setFontFamily('Arial');
  ws.setRowHeight(2,16);

  hdr(3,1,'Active',55,'TRUE = aktif di-pull saat refresh\nFALSE = skip');
  hdr(3,2,'Jira Sync',65,'TRUE = sync bugs dari Jira\nFALSE = skip Jira sync');
  hdr(3,3,'Project',120,'Auto dari Summary B2 (QATM)');
  hdr(3,4,'Modul',120,'Auto dari Summary B3 (QATM)');
  hdr(3,5,'Submodul',140,'Auto dari Summary B4 (QATM)');
  hdr(3,6,'PIC QA',110,'Auto dari Summary B6 (QATM)');
  hdr(3,7,'Spreadsheet ID',320,'URL: https://docs.google.com/spreadsheets/d/[ID]/edit');
  hdr(3,8,'Link',55);
  hdr(3,9,'Jira Instance',110,'digitalperuri / bgn-peruri');
  hdr(3,10,'Jira Project',90,'Project key (contoh: TEST)');
  ws.setRowHeight(3,22);

  // Sample data dengan TRUE/FALSE
  [[true,false,'SIMPER','Mobile App','Login, Dashboard','Budi','1evhTCv0gyfsTxkh5SusXvK_GD68HRJkH9QFZB3-jDmg','','digitalperuri','TEST'],
    [true,true,'E-Meterai','Backend API','User Management','Dedi','PASTE_SPREADSHEET_ID_HERE','','bgn-peruri','SQA'],
    [false,false,'Portal','Web Portal','Admin Panel','Fandi','PASTE_SPREADSHEET_ID_HERE','','-','-'],
  ].forEach((row,i)=>{
    ws.getRange(4+i,1,1,row.length).setValues([row]);
    ws.setRowHeight(4+i,22);

    // Conditional formatting untuk Active (col A)
    const activeCell = ws.getRange(4+i,1);
    activeCell.setHorizontalAlignment('center').setFontWeight('bold');

    // Conditional formatting untuk Jira Sync (col B)
    const jiraCell = ws.getRange(4+i,2);
    jiraCell.setHorizontalAlignment('center').setFontWeight('bold');
  });

  // Apply conditional formatting rules untuk Active column
  const activeRange = ws.getRange('A4:A1000');
  const activeTrueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A4=TRUE')
    .setBackground('#C8E6C9')
    .setFontColor('#2E7D32')
    .setRanges([activeRange])
    .build();
  const activeFalseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A4=FALSE')
    .setBackground('#FFCDD2')
    .setFontColor('#C62828')
    .setRanges([activeRange])
    .build();

  // Apply conditional formatting rules untuk Jira Sync column
  const jiraRange = ws.getRange('B4:B1000');
  const jiraTrueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B4=TRUE')
    .setBackground('#BBDEFB')
    .setFontColor('#1565C0')
    .setRanges([jiraRange])
    .build();
  const jiraFalseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B4=FALSE')
    .setBackground('#F5F5F5')
    .setFontColor('#757575')
    .setRanges([jiraRange])
    .build();

  ws.setConditionalFormatRules([activeTrueRule, activeFalseRule, jiraTrueRule, jiraFalseRule]);

  // Data validation untuk Active dan Jira Sync
  const dvBool = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ws.getRange('A4:A1000').setDataValidation(dvBool);
  ws.getRange('B4:B1000').setDataValidation(dvBool);

  // Data validation untuk Jira Instance
  const dvInstance = SpreadsheetApp.newDataValidation()
    .requireValueInList(['digitalperuri', 'bgn-peruri', '-'], true).build();
  ws.getRange('I4:I1000').setDataValidation(dvInstance);

  // ─────────────────────────────────────────────────────────────────────
  // GOOGLE CHAT NOTIFICATION SECTION (Kolom L-N, di samping kanan)
  // ─────────────────────────────────────────────────────────────────────

  const chatCol = 12; // Start at column L (12)

  // Section header (row 1, merged L1:N1)
  ws.getRange(1, chatCol, 1, 3).merge()
    .setValue('GOOGLE CHAT NOTIFICATION')
    .setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');

  // Info row (row 2, merged L2:N2)
  ws.getRange(2, chatCol, 1, 3).merge()
    .setValue('💬  Notifikasi blocker ke Google Chat per-PROJECT. Cukup isi config di 1 module per project (sisanya diabaikan).')
    .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  // Column headers (row 3)
  const chatHeaders = [
    ['Google Chat Webhook URL', 300, 'Buat webhook di Google Chat Space:\nSpace Settings > Apps & integrations > Webhooks\n\nFormat: https://chat.googleapis.com/v1/spaces/.../messages?key=...'],
    ['Notification Schedule', 120, 'FORMAT SCHEDULE:\n• Single: 7 → 1x/hari jam 7:00\n• Multiple: 7,12,18 → 3x/hari jam 7:00, 12:00, 18:00\n• Interval: 4h → Setiap 4 jam (support: 1h, 2h, 4h, 6h, 8h, 12h)\n\nContoh:\n  9,17 → Pagi & sore\n  0,6,12,18 → 4x sehari\n  6h → Setiap 6 jam'],
    ['Enable Notifikasi', 100, 'TRUE = aktif notifikasi harian\nFALSE = nonaktifkan']
  ];

  chatHeaders.forEach(([h, w, note], i) => {
    const col = chatCol + i;
    const headerCell = ws.getRange(3, col);

    headerCell
      .setValue(h)
      .setBackground('#1976D2').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);

    ws.setColumnWidth(col, w);
    if (note) headerCell.setNote(note);
  });

  // Data row with defaults (row 4)
  ws.getRange(4, chatCol, 1, 3)
    .setValues([['https://chat.googleapis.com/v1/spaces/...', '7,12,18', false]])
    .setBackground('#E3F2FD')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);

  ws.getRange(4, chatCol).setFontFamily('Courier New').setFontSize(8);
  ws.getRange(4, chatCol + 1).setHorizontalAlignment('center');
  ws.getRange(4, chatCol + 2).setHorizontalAlignment('center').setFontWeight('bold');

  // Data validation for Enable (checkbox) - column N (chatCol + 2)
  const dvChatEnable = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ws.getRange(4, chatCol + 2).setDataValidation(dvChatEnable);

  // No data validation for Schedule (column M) - accepts text format like "7,12,18" or "4h"

  // Conditional formatting for Enable checkbox (N4)
  const chatEnableRange = ws.getRange(4, chatCol + 2, 1, 1);
  const chatEnableTrueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$N4=TRUE')
    .setBackground('#C8E6C9')
    .setFontColor('#2E7D32')
    .setRanges([chatEnableRange])
    .build();
  const chatEnableFalseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$N4=FALSE')
    .setBackground('#F5F5F5')
    .setFontColor('#757575')
    .setRanges([chatEnableRange])
    .build();

  // Merge with existing rules
  const existingRules = ws.getConditionalFormatRules();
  ws.setConditionalFormatRules([...existingRules, chatEnableTrueRule, chatEnableFalseRule]);

  // ─────────────────────────────────────────────────────────────────────
  // EMAIL NOTIFICATION SECTION (Kolom O-P, di samping kanan Google Chat)
  // ─────────────────────────────────────────────────────────────────────

  const emailCol = 15; // Start at column O (15)

  // Section header (row 1, merged O1:P1)
  ws.getRange(1, emailCol, 1, 2).merge()
    .setValue('EMAIL NOTIFICATION')
    .setBackground('#43A047').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');

  // Info row (row 2, merged O2:P2)
  ws.getRange(2, emailCol, 1, 2).merge()
    .setValue('📧  Email notifikasi blocker per-PROJECT. Cukup isi config di 1 module per project (multiple recipients pisah koma)')
    .setBackground('#E8F5E9').setFontColor('#2E7D32').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  // Column headers (row 3)
  const emailHeaders = [
    ['Email Recipients (comma separated)', 260, 'Masukkan email addresses dipisah koma\nContoh: dev@company.com, qa@company.com, manager@company.com'],
    ['Enable Email', 85, 'TRUE = aktif email harian\nFALSE = nonaktifkan']
  ];

  emailHeaders.forEach(([h, w, note], i) => {
    const col = emailCol + i;
    const headerCell = ws.getRange(3, col);

    headerCell
      .setValue(h)
      .setBackground('#66BB6A').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#A5D6A7', SpreadsheetApp.BorderStyle.SOLID);

    ws.setColumnWidth(col, w);
    if (note) headerCell.setNote(note);
  });

  // Data row with defaults (row 4)
  ws.getRange(4, emailCol, 1, 2)
    .setValues([['user@example.com, team@example.com', false]])
    .setBackground('#E8F5E9')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#A5D6A7', SpreadsheetApp.BorderStyle.SOLID);

  ws.getRange(4, emailCol).setFontFamily('Courier New').setFontSize(8);
  ws.getRange(4, emailCol + 1).setHorizontalAlignment('center').setFontWeight('bold');

  // Data validation for Enable Email (checkbox) - column P (emailCol + 1)
  const dvEmailEnable = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ws.getRange(4, emailCol + 1).setDataValidation(dvEmailEnable);

  // Conditional formatting for Enable Email checkbox (P4)
  const emailEnableRange = ws.getRange(4, emailCol + 1, 1, 1);
  const emailEnableTrueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$P4=TRUE')
    .setBackground('#C8E6C9')
    .setFontColor('#2E7D32')
    .setRanges([emailEnableRange])
    .build();
  const emailEnableFalseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$P4=FALSE')
    .setBackground('#F5F5F5')
    .setFontColor('#757575')
    .setRanges([emailEnableRange])
    .build();

  // Merge all rules
  const allRules = ws.getConditionalFormatRules();
  ws.setConditionalFormatRules([...allRules, emailEnableTrueRule, emailEnableFalseRule]);

  // ─────────────────────────────────────────────────────────────────────
  // AUTO REFRESH SECTION (Kolom Q-R, di samping kanan Email)
  // ─────────────────────────────────────────────────────────────────────

  const refreshCol = 17; // Start at column Q (17)

  // Section header (row 1, merged Q1:R1)
  ws.getRange(1, refreshCol, 1, 2).merge()
    .setValue('AUTO REFRESH')
    .setBackground('#FF6F00').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');

  // Info row (row 2, merged Q2:R2)
  ws.getRange(2, refreshCol, 1, 2).merge()
    .setValue('🔄  Auto-refresh dashboard & Jira sync setiap X menit')
    .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  // Column headers (row 3)
  const refreshHeaders = [
    ['Refresh Interval (Minutes)', 150, 'Berapa menit sekali auto-refresh?\nContoh: 10 = refresh setiap 10 menit\nMin: 1, Max: 60'],
    ['Enable Auto Refresh', 120, 'TRUE = aktif auto-refresh\nFALSE = nonaktifkan (manual only)']
  ];

  refreshHeaders.forEach(([h, w, note], i) => {
    const col = refreshCol + i;
    const headerCell = ws.getRange(3, col);

    headerCell
      .setValue(h)
      .setBackground('#FF8F00').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#FFCC80', SpreadsheetApp.BorderStyle.SOLID);

    ws.setColumnWidth(col, w);
    if (note) headerCell.setNote(note);
  });

  // Data row with defaults (row 4)
  ws.getRange(4, refreshCol, 1, 2)
    .setValues([[10, false]])
    .setBackground('#FFF3E0')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#FFCC80', SpreadsheetApp.BorderStyle.SOLID);

  ws.getRange(4, refreshCol).setHorizontalAlignment('center').setFontWeight('bold');
  ws.getRange(4, refreshCol + 1).setHorizontalAlignment('center').setFontWeight('bold');

  // Data validation for Interval (1-60 minutes)
  const dvInterval = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 60)
    .build();
  ws.getRange(4, refreshCol).setDataValidation(dvInterval);

  // Data validation for Enable Auto Refresh (checkbox) - column R (refreshCol + 1)
  const dvRefreshEnable = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ws.getRange(4, refreshCol + 1).setDataValidation(dvRefreshEnable);

  // Conditional formatting for Enable Auto Refresh checkbox (R4)
  const refreshEnableRange = ws.getRange(4, refreshCol + 1, 1, 1);
  const refreshEnableTrueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$R4=TRUE')
    .setBackground('#C8E6C9')
    .setFontColor('#2E7D32')
    .setRanges([refreshEnableRange])
    .build();
  const refreshEnableFalseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$R4=FALSE')
    .setBackground('#F5F5F5')
    .setFontColor('#757575')
    .setRanges([refreshEnableRange])
    .build();

  // Merge all rules including refresh
  const finalRules = ws.getConditionalFormatRules();
  ws.setConditionalFormatRules([...finalRules, refreshEnableTrueRule, refreshEnableFalseRule]);

  // ─────────────────────────────────────────────────────────────────────
  // WHATSAPP NOTIFICATION SECTION (Kolom S-U, per-module config)
  // ─────────────────────────────────────────────────────────────────────

  const waCol = 19; // Start at column S (19)

  // Section header (row 1, merged S1:U1)
  ws.getRange(1, waCol, 1, 3).merge()
    .setValue('WHATSAPP NOTIFICATION')
    .setBackground('#25D366').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');

  // Info row (row 2, merged S2:U2)
  ws.getRange(2, waCol, 1, 3).merge()
    .setValue('📱  Notifikasi blocker ke WhatsApp Group per-PROJECT via Fonnte API. Cukup isi config di 1 module per project (sisanya diabaikan). Token shared row 4.')
    .setBackground('#E8F5E9').setFontColor('#1B5E20').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  // Column headers (row 3)
  const waHeaders = [
    ['WhatsApp Group ID', 200, 'Group ID WhatsApp (per-PROJECT)\nFormat: 120363xxxxxxxxx@g.us\n\nCara dapat:\n1. Menu → Notifications → WhatsApp: Get Groups\n2. Atau dari invite link grup\n\n💡 TIP: Cukup isi di 1 module per project (biasanya module pertama).\nSemua module dalam 1 project akan kirim ke grup yang sama.\n\nRow 4 = Global fallback (opsional)'],
    ['Fonnte Token', 160, 'Token API dari Fonnte Dashboard (SHARED - row 4 only)\nLogin: https://fonnte.com/dashboard\nMenu: Device → Token\n\nToken di row 4 dipakai untuk SEMUA project\nRow 5+ kosongkan (tidak perlu diisi)'],
    ['Enable WA', 85, 'TRUE = aktif notifikasi WhatsApp untuk project ini\nFALSE = nonaktifkan\n\n💡 TIP: Cukup centang di 1 module per project']
  ];

  waHeaders.forEach(([h, w, note], i) => {
    const col = waCol + i;
    const headerCell = ws.getRange(3, col);

    headerCell
      .setValue(h)
      .setBackground('#43A047').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#81C784', SpreadsheetApp.BorderStyle.SOLID);

    ws.setColumnWidth(col, w);
    if (note) headerCell.setNote(note);
  });

  // Data row with defaults (row 4)
  ws.getRange(4, waCol, 1, 3)
    .setValues([['120363289471046194@g.us', 'TDdpPB6Gbyn7KYbqLgN8', false]])
    .setBackground('#E8F5E9')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#81C784', SpreadsheetApp.BorderStyle.SOLID);

  ws.getRange(4, waCol).setFontFamily('Courier New').setFontSize(8);
  ws.getRange(4, waCol + 1).setFontFamily('Courier New').setFontSize(8);
  ws.getRange(4, waCol + 2).setHorizontalAlignment('center').setFontWeight('bold');

  // Data validation for Enable WA (checkbox) - column U (waCol + 2)
  const dvWAEnable = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ws.getRange(4, waCol + 2).setDataValidation(dvWAEnable);

  // Conditional formatting for Enable WA checkbox (U4)
  const waEnableRange = ws.getRange(4, waCol + 2, 1, 1);
  const waEnableTrueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$U4=TRUE')
    .setBackground('#C8E6C9')
    .setFontColor('#2E7D32')
    .setRanges([waEnableRange])
    .build();
  const waEnableFalseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$U4=FALSE')
    .setBackground('#F5F5F5')
    .setFontColor('#757575')
    .setRanges([waEnableRange])
    .build();

  // Merge all rules including WhatsApp
  const allConfigRules = ws.getConditionalFormatRules();
  ws.setConditionalFormatRules([...allConfigRules, waEnableTrueRule, waEnableFalseRule]);

  // ─────────────────────────────────────────────────────────────────────
  // VAPT SPREADSHEET SECTION (Kolom V-W, untuk VAPT data source)
  // ─────────────────────────────────────────────────────────────────────

  const vaptCol = 22; // Start at column V (22)

  // Section header (row 1, merged V1:W1)
  ws.getRange(1, vaptCol, 1, 2).merge()
    .setValue('VAPT DATA SOURCE')
    .setBackground('#EF6C00').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');

  // Info row (row 2, merged V2:W2)
  ws.getRange(2, vaptCol, 1, 2).merge()
    .setValue('🔒  VAPT Report per-PROJECT. Cukup isi Spreadsheet ID di 1 module per project. Jika kosong = no VAPT.')
    .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  // Column headers (row 3)
  const vaptHeaders = [
    ['VAPT Spreadsheet ID', 320, 'Spreadsheet ID VAPT report untuk PROJECT ini\nContains: Ad Hoc VAPT + Regular VAPT tabs\n\n💡 TIP: Cukup isi di 1 module per project.\nSetiap project bisa punya VAPT spreadsheet sendiri.\n\nFormat: 17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk\nJika kosong = project tidak punya VAPT report'],
    ['Enable VAPT', 85, 'TRUE = include VAPT data in notifications\nFALSE = skip VAPT\n\n💡 TIP: Cukup centang di 1 module per project']
  ];

  vaptHeaders.forEach(([h, w, note], i) => {
    const col = vaptCol + i;
    const headerCell = ws.getRange(3, col);

    headerCell
      .setValue(h)
      .setBackground('#FF6F00').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#FFB74D', SpreadsheetApp.BorderStyle.SOLID);

    ws.setColumnWidth(col, w);
    if (note) headerCell.setNote(note);
  });

  // Data row with default (row 4)
  ws.getRange(4, vaptCol, 1, 2)
    .setValues([['17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk', true]])
    .setBackground('#FFF3E0')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#FFB74D', SpreadsheetApp.BorderStyle.SOLID);

  ws.getRange(4, vaptCol).setFontFamily('Courier New').setFontSize(8);
  ws.getRange(4, vaptCol + 1).setHorizontalAlignment('center').setFontWeight('bold');

  // Data validation for Enable VAPT (checkbox) - column W (vaptCol + 1)
  const dvVAPTEnable = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ws.getRange(4, vaptCol + 1).setDataValidation(dvVAPTEnable);

  // Conditional formatting for Enable VAPT checkbox (W4)
  const vaptEnableRange = ws.getRange(4, vaptCol + 1, 1, 1);
  const vaptEnableTrueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$W4=TRUE')
    .setBackground('#C8E6C9')
    .setFontColor('#2E7D32')
    .setRanges([vaptEnableRange])
    .build();
  const vaptEnableFalseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$W4=FALSE')
    .setBackground('#F5F5F5')
    .setFontColor('#757575')
    .setRanges([vaptEnableRange])
    .build();

  // Merge all rules including VAPT
  const finalConfigRules = ws.getConditionalFormatRules();
  ws.setConditionalFormatRules([...finalConfigRules, vaptEnableTrueRule, vaptEnableFalseRule]);

  // ─────────────────────────────────────────────────────────────────────
  // AUTOMATION CONTRACT SECTION (Kolom X-Z)
  // ─────────────────────────────────────────────────────────────────────

  const autoContractCol = 24; // Start at column X (24)

  ws.getRange(1, autoContractCol, 1, 3).merge()
    .setValue('AUTOMATION CONTRACT')
    .setBackground('#00695C').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');

  ws.getRange(2, autoContractCol, 1, 3).merge()
    .setValue('Mapping Jenkins automation result ke Project/Modul/Submodul dashboard. Isi jika nama job/tag berbeda dari nama di dashboard.')
    .setBackground('#E0F2F1').setFontColor('#00695C').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  const contractHeaders = [
    ['Automation Contract', 180, 'Alias umum untuk Web/API jika sama.\nContoh: qa-web-4-menuplanner-regression-prod atau menuplanner'],
    ['Web Automation Contract', 180, 'Alias khusus Web automation.\nDipakai untuk match payload channel=web dari Jenkins.'],
    ['API Automation Contract', 180, 'Alias khusus API automation.\nDipakai untuk match payload channel=api dari Jenkins.']
  ];

  contractHeaders.forEach(([h, w, note], i) => {
    const col = autoContractCol + i;
    const headerCell = ws.getRange(3, col);
    headerCell
      .setValue(h)
      .setBackground('#00897B').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#80CBC4', SpreadsheetApp.BorderStyle.SOLID);
    ws.setColumnWidth(col, w);
    if (note) headerCell.setNote(note);
  });

  // ─────────────────────────────────────────────────────────────────────
  // EXTERNAL QA SECTION (Kolom AL-AN)
  // ─────────────────────────────────────────────────────────────────────

  const externalQACol = 38; // Start at column AL (after Y:AK automation contract block)
  ensureSheetColumns_(ws, externalQACol + 2);

  ws.getRange(1, externalQACol, 1, 3).merge()
    .setValue('EXTERNAL QA')
    .setBackground('#455A64').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');

  ws.getRange(2, externalQACol, 1, 3).merge()
    .setValue('Flag untuk scope yang dites external team dan apakah perlu retest QA Peruri. Detail evidence ada di QATM tab External Test Report.')
    .setBackground('#ECEFF1').setFontColor('#455A64').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  const externalQAHeaders = [
    ['External QA', 90, 'TRUE = testing utama dilakukan external team.\nFALSE = testing internal QA Peruri seperti biasa.'],
    ['Retest QA Peruri', 120, 'TRUE = tetap perlu retest internal QA Peruri.\nFALSE = tidak perlu retest internal; hasil test QATM boleh kosong jika External QA TRUE.'],
    ['External QA Notes', 220, 'Catatan tambahan di level dashboard config. Scope / Notes utama tetap di QATM Summary.']
  ];

  externalQAHeaders.forEach(([h, w, note], i) => {
    const col = externalQACol + i;
    const headerCell = ws.getRange(3, col);
    headerCell
      .setValue(h)
      .setBackground('#607D8B').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#B0BEC5', SpreadsheetApp.BorderStyle.SOLID);
    ws.setColumnWidth(col, w);
    if (note) headerCell.setNote(note);
  });

  ws.getRange(4, externalQACol, 3, 3)
    .setValues([[false, true, ''], [false, true, ''], [false, true, '']])
    .setBackground('#ECEFF1')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#B0BEC5', SpreadsheetApp.BorderStyle.SOLID);
  ws.getRange(4, externalQACol, 997, 2).setDataValidation(dvBool);
  ws.getRange(4, externalQACol, 997, 2).setHorizontalAlignment('center').setFontWeight('bold');
}


// ═══════════════════════════════════════════════════════════════════════
// CREDENTIALS TAB — untuk Jira credentials
// ═══════════════════════════════════════════════════════════════════════

function buildCredentials(ss) {
  const ws = ss.insertSheet('Credentials');
  ws.setTabColor('#4A148C');

  // SIMPLIFIED VERSION - Minimal formatting to avoid timeout
  // Title
  ws.getRange(1,1).setValue('JIRA CREDENTIALS - Isi email & token per instance Jira')
    .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold');

  // Warning
  ws.getRange(2,1).setValue('⚠️ PENTING: Jangan share spreadsheet ke publik jika token diisi di sini')
    .setFontColor('#6A1B9A');

  // Headers - Simple, no borders, no notes
  ws.getRange(3,1,1,4).setValues([['Instance','Email Atlassian','API Token','Notif Email']])
    .setBackground('#6A1B9A').setFontColor('#FFFFFF').setFontWeight('bold');

  // Column widths
  ws.setColumnWidth(1, 140);
  ws.setColumnWidth(2, 220);
  ws.setColumnWidth(3, 400);
  ws.setColumnWidth(4, 280);

  // Sample data - Batch operation, no individual formatting
  ws.getRange(4,1,2,4).setValues([
    ['digitalperuri','email@company.com','ATATT3xFf...(paste token disini)','qa@company.com'],
    ['bgn-peruri',   'email@company.com','ATATT3xFf...(paste token disini)','']
  ]);

  // NO data validation, NO protection - to avoid timeout
  // Users can add manually if needed
}


// ═══════════════════════════════════════════════════════════════════════
// OVERVIEW TAB — build + write + charts
// ═══════════════════════════════════════════════════════════════════════

function buildOverview(ss) {
  const ws = ss.insertSheet('Overview');
  ws.setTabColor('#0D47A1');
  ws.clear();
  initOverviewHeaders_(ws);
  ws.getRange(5,1,1,26).merge()
      .setValue('▶ Run refreshDashboard() untuk mengisi data')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(4);
}

/**
 * Get Web App Dashboard URL from script properties
 * URL should be set via: PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', 'https://...')
 */
const DEFAULT_DASHBOARD_WEB_URL = 'https://qa-platform.inadigital.co.id/';

function getWebAppUrl_() {
  try {
    const configuredUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';
    if (!configuredUrl || configuredUrl.indexOf('script.google.com') !== -1) {
      return DEFAULT_DASHBOARD_WEB_URL;
    }
    return configuredUrl;
  } catch (e) {
    Logger.log('Error getting web app URL: ' + e.toString());
    return DEFAULT_DASHBOARD_WEB_URL;
  }
}

function initOverviewHeaders_(ws) {
  const lastCol = Math.max(ws.getLastColumn()||1, 34);
  try { ws.getRange(1,1,5,lastCol).breakApart(); } catch(e) {}
  ws.getRange(1,1,5,lastCol).clearContent().clearFormat();

  function h_(r,c,nr,nc,txt,bg,fg,sz){
    const rng=(nr>1||nc>1)?ws.getRange(r,c,nr,nc).merge():ws.getRange(r,c);
    rng.setValue(txt||'').setBackground(bg||'#0D47A1').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  }

  // Col widths — legacy 26 cols + external QA metadata.
  [80,80,90,80, 48,52,56,52, 48,52,48,48,60, 56,60,52, 48,52,48,48,60, 56,60,52, 60, 140,
    70,80,150,180,120,120,120,160]
      .forEach((w,i)=>ws.setColumnWidth(i+1,w));

  // Row 1 — Web App Dashboard Link
  const webAppUrl = getWebAppUrl_();
  if (webAppUrl) {
    ws.getRange(1,1,1,34).merge()
        .setFormula('=HYPERLINK("' + webAppUrl + '","📊 Open Interactive Dashboard (Charts & Trends)")')
        .setBackground('#4CAF50').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  } else {
    ws.getRange(1,1,1,34).merge()
        .setValue('📊 Interactive Dashboard Available (Deploy web app to get link)')
        .setBackground('#FF9800').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  }
  ws.setRowHeight(1,28);

  // Row 2 — last refresh
  ws.getRange(2,1,1,34).merge().setValue('Last refreshed: —')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(2,16);

  // Row 3 — title
  h_(3,1,1,34,'QA DASHBOARD  |  PORTFOLIO OVERVIEW','#0D47A1','#FFFFFF',13);
  ws.setRowHeight(3,30);

  // Row 4 — group headers
  // NEW: Bugs moved after MODULE INFO, added Prod Bugs
  h_(4,1, 1,4, 'MODULE INFO',    '#263238');
  h_(4,5, 1,4, 'BUGS',           '#B71C1C');
  h_(4,9, 1,5, 'WEB / MOBILE',   '#1565C0');
  h_(4,14,1,3, '🔥 SMOKE WEB',   '#BF360C');
  h_(4,17,1,5, 'API',             '#283593');
  h_(4,22,1,3, '🔥 SMOKE API',   '#4A148C');
  h_(4,25,1,1, 'PERF',            '#004D40');
  h_(4,26,1,1, 'NOTES',           '#37474F');
  h_(4,27,1,8, 'EXTERNAL QA',      '#455A64');
  ws.setRowHeight(4,22);

  // Row 5 — column headers
  // NEW: Added Prod Bugs column
  ['Project','Modul','Submodul','PIC QA',
    'Bugs','Blocker','Critical','Prod',
    'Total','Pass','Fail','Block','Pass%',
    'Total','Pass%','Exec%',
    'Total','Pass','Fail','Block','Pass%',
    'Total','Pass%','Exec%',
    'Perf','Notes',
    'External','Retest','Scope Notes','External Notes','Func Review','Perf Review','VAPT Review','Overall'
  ].forEach((lbl,i)=>h_(5,i+1,1,1,lbl,'#1565C0'));
  ws.getRange(5,5).setNote('Total bugs (all status kecuali Closed)');
  ws.getRange(5,6).setNote('Bug Blocker (Priority Critical/High/Medium, Status Open/In Progress/Reopen)');
  ws.getRange(5,7).setNote('Bugs dengan Priority Critical');
  ws.getRange(5,8).setNote('Bugs di Production environment (belum Closed)');
  ws.getRange(5,14).setNote('Smoke Web: TC Priority Critical+High+Medium');
  ws.getRange(5,15).setNote('Smoke Web Pass Rate (target ≥80%)');
  ws.getRange(5,16).setNote('Smoke Web Exec Rate (% TC sudah ada hasil)');
  ws.getRange(5,22).setNote('Smoke API: TC Priority Critical+High+Medium');
  ws.getRange(5,23).setNote('Smoke API Pass Rate (target ≥80%)');
  ws.getRange(5,24).setNote('Smoke API Exec Rate');
  ws.getRange(5,27).setNote('TRUE jika testing dilakukan external team.');
  ws.getRange(5,28).setNote('TRUE jika tetap perlu retest oleh QA Peruri.');
  ws.getRange(5,29).setNote('Manual Scope / Notes dari QATM Summary.');
  ws.getRange(5,30).setNote('Notes dari Config External QA Notes dan tab External Test Report.');
  ws.setRowHeight(5,26);
  ws.setFrozenRows(5);
}

function writeOverview(ss, allData) {
  let ws = ss.getSheetByName('Overview');
  if (!ws) { buildOverview(ss); ws = ss.getSheetByName('Overview'); }

  initOverviewHeaders_(ws);  // safe rebuild — breakApart dulu

  // Clear ALL data rows (not just lastRow which may have old inactive module data)
  const lastRow = ws.getMaxRows();
  if (lastRow>=6) ws.getRange(6,1,lastRow-5,34).clearContent().clearFormat();

  const rules = [];

  allData.forEach((d,i)=>{
    const r  = 6+i;
    const bg = i%2===0 ? '#F9FAFB' : '#FFFFFF';
    const bs = d.bugStats||{};
    const hasSmoke = d.wSmokeTotal>0||d.aSmokeTotal>0;

    function cell(col,val,fmt){
      const c=ws.getRange(r,col).setValue(val==null?'':val).setBackground(bg)
          .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
      if(fmt)c.setNumberFormat(fmt);
      return c;
    }

    // NEW LAYOUT: Project, Modul, Submodul, PIC QA | BUGS (4) | WEB (5) | SMOKE WEB (3) | API (5) | SMOKE API (3) | PERF | NOTES
    cell(1,d.project||d.sprint||'');
    cell(2,d.module||'');

    // Submodule with hyperlink to QATM BugReport
    const submodulCell = ws.getRange(r,3);
    if (d.id) {
      try {
        const qatmSs = SpreadsheetApp.openById(d.id);
        const bugSheet = qatmSs.getSheetByName('BugReport');
        if (bugSheet) {
          const bugReportGid = bugSheet.getSheetId();
          const qatmUrl = 'https://docs.google.com/spreadsheets/d/' + d.id + '/edit#gid=' + bugReportGid;
          submodulCell.setFormula('=HYPERLINK("' + qatmUrl + '","' + (d.submodule||d.name) + '")');
          submodulCell.setFontColor('#1155CC');  // Blue link color
        } else {
          submodulCell.setValue(d.submodule||d.name);
        }
      } catch (e) {
        submodulCell.setValue(d.submodule||d.name);
      }
    } else {
      submodulCell.setValue(d.submodule||d.name);
    }
    submodulCell.setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);

    cell(4,d.team||'');

    // Bugs (col 5-8) - added prodBugs, all should be count format
    cell(5,bs.total||0,'0'); cell(6,bs.blocker||0,'0'); cell(7,bs.critical||0,'0'); cell(8,bs.prodBugs||0,'0');

    // Web (col 9-13)
    cell(9,d.wTotal,'0'); cell(10,d.wPassed,'0'); cell(11,d.wFailed,'0'); cell(12,d.wBlocked,'0'); // Fixed: all should be count format
    cell(13,d.error?'ERR':d.wPassRate,'0%');

    // Smoke Web (col 14-16)
    cell(14,hasSmoke?d.wSmokeTotal:'--','0'); // Fixed: should be count, not percentage
    cell(15,hasSmoke?d.wSmokePassRate:'--',hasSmoke?'0%':null);
    cell(16,hasSmoke?d.wSmokeExecRate:'--',hasSmoke?'0%':null);

    // API (col 17-21)
    cell(17,d.aTotal,'0'); cell(18,d.aPassed,'0'); cell(19,d.aFailed,'0'); cell(20,d.aBlocked,'0'); // Fixed: all should be count format
    cell(21,d.error?'ERR':d.aPassRate,'0%');

    // Smoke API (col 22-24)
    cell(22,hasSmoke?d.aSmokeTotal:'--','0'); // Fixed: should be count, not percentage
    cell(23,hasSmoke?d.aSmokePassRate:'--',hasSmoke?'0%':null);
    cell(24,hasSmoke?d.aSmokeExecRate:'--',hasSmoke?'0%':null);

    // Perf (col 25)
    cell(25,d.perfResult);

    // Notes (col 26)
    ws.getRange(r,26).setValue(d.error||'').setBackground(bg).setFontFamily('Arial').setFontSize(8)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    const externalQA = d.externalQA || {};
    const externalReport = d.externalTestReport || {};
    cell(27, externalQA.isExternal === true);
    cell(28, externalQA.retestByPeruri === true);
    ws.getRange(r,29).setValue(d.scopeNotes || '').setBackground(bg).setFontFamily('Arial').setFontSize(8)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r,30).setValue([externalQA.notes, externalReport.notes].filter(Boolean).join('\n')).setBackground(bg).setFontFamily('Arial').setFontSize(8)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    cell(31, externalReport.functionalReviewStatus || '');
    cell(32, externalReport.performanceReviewStatus || '');
    cell(33, externalReport.vaptReviewStatus || '');
    cell(34, externalReport.overallStatus || '');
    ws.setRowHeight(r,22);

    // RAG Pass%
    [13,21].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.8,0.5)));
    // RAG Smoke Pass%
    [15,23].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.8,0.5)));
    // RAG Smoke Exec%
    [16,24].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.7,0.4)));
    // Failed > 0
    [11,19].forEach(col=>rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFCDD2').setFontColor('#C62828').setBold(true)
        .setRanges([ws.getRange(r,col)]).build()));
    // Blocked > 0
    [12,20].forEach(col=>rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
        .setRanges([ws.getRange(r,col)]).build()));
    // Blocker/Critical/ProdBugs > 0
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
        .setRanges([ws.getRange(r,6),ws.getRange(r,7),ws.getRange(r,8)]).build());
    // Perf
    [['PASS','#C8E6C9','#1B5E20'],['FAIL','#FFCDD2','#C62828'],['--','#F5F5F5','#9E9E9E']]
        .forEach(([v,bg2,fg])=>rules.push(SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(v).setBackground(bg2).setFontColor(fg).setBold(true)
            .setRanges([ws.getRange(r,25)]).build()));
  });

  // TOTAL row - updated for new layout with prodBugs
  if (allData.length > 0) {
    const tr = 6+allData.length;  // Fixed: was 5+length (conflict with last data row)
    ws.getRange(tr,1,1,4).merge().setValue('TOTAL / AVERAGE')
        .setBackground('#E3F2FD').setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    // Bugs totals (col 5-8) - added prodBugs, add number format
    [[5,'total'],[6,'blocker'],[7,'critical'],[8,'prodBugs']].forEach(([col,key])=>
        ws.getRange(tr,col).setValue(allData.reduce((a,d)=>a+((d.bugStats||{})[key]||0),0)).setNumberFormat('0')
            .setBackground('#DDEEFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center'));
    // Web totals (col 9-12) - Fixed: add number format
    [[9,'wTotal'],[10,'wPassed'],[11,'wFailed'],[12,'wBlocked']].forEach(([col,key])=>{
      ws.getRange(tr,col).setValue(allData.reduce((a,d)=>a+(d[key]||0),0)).setNumberFormat('0')
          .setBackground('#DDEEFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    });
    // Smoke Web Total (col 14) - Fixed: should be count sum, not average
    ws.getRange(tr,14).setValue(allData.reduce((a,d)=>a+(d.wSmokeTotal||0),0)).setNumberFormat('0')
        .setBackground('#FFF3E0').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    // API totals (col 17-20) - Fixed: add number format
    [[17,'aTotal'],[18,'aPassed'],[19,'aFailed'],[20,'aBlocked']].forEach(([col,key])=>{
      ws.getRange(tr,col).setValue(allData.reduce((a,d)=>a+(d[key]||0),0)).setNumberFormat('0')
          .setBackground('#DDEEFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    });
    // Smoke API Total (col 22) - Fixed: should be count sum, not average
    ws.getRange(tr,22).setValue(allData.reduce((a,d)=>a+(d.aSmokeTotal||0),0)).setNumberFormat('0')
        .setBackground('#FFF3E0').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    // Averages for Pass%, Smoke Pass%, Smoke Exec%
    const avg=(key)=>allData.reduce((a,d)=>a+(d[key]||0),0)/allData.length;
    [[13,'wPassRate'],[21,'aPassRate'],[15,'wSmokePassRate'],[16,'wSmokeExecRate'],[23,'aSmokePassRate'],[24,'aSmokeExecRate']].forEach(([col,key])=>
        ws.getRange(tr,col).setValue(avg(key)).setNumberFormat('0%')
            .setBackground(col>=14&&col<=16||col>=22&&col<=24?'#FFF3E0':'#DDEEFF')
            .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center'));
    ws.setRowHeight(tr,22);
  }

  ws.setConditionalFormatRules(rules);
  // Charts removed - will use Web App dashboard for visualization
  // buildOverviewCharts_(ws, allData);
}

function buildSmoke(ss) {
  const ws = ss.insertSheet('Smoke');
  ws.setTabColor('#BF360C');
  ws.clear();
  initSmokeHeaders_(ws);
  ws.getRange(5,1,1,13).merge()
      .setValue('▶ Run refreshDashboard() untuk mengisi data')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(4);
}

function initSmokeHeaders_(ws) {
  function h_(r,c,nr,nc,txt,bg,fg,sz){
    const rng=(nr>1||nc>1)?ws.getRange(r,c,nr,nc).merge():ws.getRange(r,c);
    rng.setValue(txt||'').setBackground(bg||'#BF360C').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#FFCCBC',SpreadsheetApp.BorderStyle.SOLID);
  }
  const lastCol=Math.max(ws.getLastColumn()||1,13);
  try{ws.getRange(1,1,4,lastCol).breakApart();}catch(e){}
  ws.getRange(1,1,4,lastCol).clearContent().clearFormat();

  [90,85,100,90, 62,70,60, 62,70,60, 62,62, 175]
      .forEach((w,i)=>ws.setColumnWidth(i+1,w));

  ws.getRange(1,1,1,13).merge().setValue('Last refreshed: —')
      .setBackground('#FBE9E7').setFontColor('#BF360C').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  h_(2,1,1,13,'🔥  SMOKE TEST DASHBOARD  —  Priority: Critical + High + Medium','#BF360C','#FFFFFF',13);
  ws.setRowHeight(2,30);

  h_(3,1, 1,4,'MODULE INFO',           '#263238');
  h_(3,5, 1,3,'SMOKE WEB / MOBILE',    '#BF360C');
  h_(3,8, 1,3,'SMOKE API',             '#4A148C');
  h_(3,11,1,2,'OPEN BLOCKER (BUG)',    '#B71C1C');
  h_(3,13,1,1,'STATUS',                '#37474F');
  ws.setRowHeight(3,20);

  ['Project','Modul','Submodul','PIC QA',
    'Total','Pass%','Exec%',
    'Total','Pass%','Exec%',
    'Web Bug','API Bug','Smoke Status'
  ].forEach((lbl,i)=>h_(4,i+1,1,1,lbl,'#1565C0'));
  ws.getRange(4,4).setNote('PIC QA\nAuto-filled dari Summary B6');
  ws.getRange(4,5).setNote('Total TC Smoke Web (Critical+High+Medium)');
  ws.getRange(4,6).setNote('Pass Rate Smoke Web (target ≥80%)');
  ws.getRange(4,7).setNote('Exec Rate Smoke Web (% TC sudah ada hasil)');
  ws.getRange(4,8).setNote('Total TC Smoke API (Critical+High+Medium)');
  ws.getRange(4,9).setNote('Pass Rate Smoke API (target ≥80%)');
  ws.getRange(4,11).setNote('Bug Open/InProg/Reopen priority Med-Critical (Web type)');
  ws.getRange(4,12).setNote('Bug Open/InProg/Reopen priority Critical');
  ws.setRowHeight(4,26);
  ws.setFrozenRows(4);
}

function writeSmoke(ss, allData) {
  let ws = ss.getSheetByName('Smoke');
  if (!ws) { buildSmoke(ss); ws = ss.getSheetByName('Smoke'); }

  initSmokeHeaders_(ws);

  // Clear ALL data rows to remove inactive modules
  const lastRow=ws.getMaxRows();
  if (lastRow>=5) ws.getRange(5,1,lastRow-4,13).clearContent().clearFormat();

  const rules=[];

  allData.forEach((d,i)=>{
    const r=5+i, bg=i%2===0?'#FFF8F6':'#FFFFFF';
    const bs=d.bugStats||{};
    const hasW=d.wSmokeTotal>0, hasA=d.aSmokeTotal>0;

    function cell(col,val,fmt){
      const c=ws.getRange(r,col).setValue(val==null?'':val).setBackground(bg)
          .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setBorder(true,true,true,true,false,false,'#FFCCBC',SpreadsheetApp.BorderStyle.SOLID);
      if(fmt)c.setNumberFormat(fmt);
      return c;
    }

    cell(1,d.project||d.sprint||'');
    cell(2,d.module||'');
    ws.getRange(r,3).setValue(d.submodule||d.name).setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#FFCCBC',SpreadsheetApp.BorderStyle.SOLID);
    cell(4,d.team||'');        // PIC QA

    cell(5, hasW?d.wSmokeTotal:'--');
    cell(6, hasW?d.wSmokePassRate:'--', hasW?'0%':null);
    cell(7, hasW?d.wSmokeExecRate:'--', hasW?'0%':null);

    cell(8, hasA?d.aSmokeTotal:'--');
    cell(9, hasA?d.aSmokePassRate:'--', hasA?'0%':null);
    cell(10,hasA?d.aSmokeExecRate:'--', hasA?'0%':null);

    cell(11,bs.blocker||0);
    cell(12,bs.critical||0);

    // Smoke Status badge
    const smokePct = hasW ? d.wSmokePassRate : (hasA ? d.aSmokePassRate : null);
    const status = d.error           ? '❌ Error'
        : smokePct===null    ? '⬜ Belum ada'
            : smokePct>=0.8      ? '✅ Ready'
                : smokePct>=0.5      ? '⚠️ Perlu perhatian'
                    : '🚨 Blocker';
    const stBg   = d.error?'#FFCDD2':smokePct===null?'#F5F5F5':smokePct>=0.8?'#C8E6C9':smokePct>=0.5?'#FFF9C4':'#FFCDD2';
    const stFg   = d.error?'#C62828':smokePct===null?'#757575':smokePct>=0.8?'#1B5E20':smokePct>=0.5?'#E65100':'#B71C1C';
    ws.getRange(r,13).setValue(status).setBackground(stBg).setFontColor(stFg)
        .setFontFamily('Arial').setFontSize(8).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#FFCCBC',SpreadsheetApp.BorderStyle.SOLID);

    ws.setRowHeight(r,22);

    [6,9].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.8,0.5)));
    [7,10].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.7,0.4)));
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
        .setRanges([ws.getRange(r,11),ws.getRange(r,12)]).build());
  });

  // Summary row
  if (allData.length>0) {
    const tr=5+allData.length;
    ws.getRange(tr,1,1,4).merge().setValue('RATA-RATA / TOTAL')
        .setBackground('#FCE4EC').setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    const avg=(key)=>allData.reduce((a,d)=>a+(d[key]||0),0)/allData.length;
    const sum=(key)=>allData.reduce((a,d)=>a+(d[key]||0),0);
    [[5,sum('wSmokeTotal'),null],[6,avg('wSmokePassRate'),'0%'],
      [8,sum('aSmokeTotal'),null],[9,avg('aSmokePassRate'),'0%'],
      [11,sum(d=>(d.bugStats||{}).blocker||0),null],[12,sum(d=>(d.bugStats||{}).critical||0),null]
    ].forEach(([col,val,fmt])=>{
      const c=ws.getRange(tr,col).setValue(typeof val==='function'?allData.reduce((a,d)=>a+val(d),0):val)
          .setBackground('#FCE4EC').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
      if(fmt)c.setNumberFormat(fmt);
    });
    ws.setRowHeight(tr,22);
  }

  ws.setConditionalFormatRules(rules);
  // Charts removed - will use Web App dashboard for visualization
  // buildSmokeCharts_(ws, allData);
}

function buildFailureScenario(ss) {
  const ws=ss.insertSheet('Failure Scenario'); ws.setTabColor('#B71C1C'); ws.clear();
  function h_(c,txt){ws.getRange(2,c).setValue(txt).setBackground('#B71C1C').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#E57373',SpreadsheetApp.BorderStyle.SOLID);}
  [120,80,85,75,75,100,250,85].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  ws.getRange(1,1,1,8).merge().setValue('FAILURE SCENARIO  —  Priority: Critical / High / Medium  |  Status: FAILED / BLOCKED')
      .setBackground('#B71C1C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(12).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,28);
  ['Modul','Type','TC_ID','Priority','Status','Feature','Scenario','Refresh'].forEach((t,i)=>h_(i+1,t));
  ws.setRowHeight(2,20);
  ws.getRange(3,1,1,8).merge().setValue('▶ Run refreshDashboard()')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(2);
}

function writeFailureScenario(ss, allData) {
  const ws=ss.getSheetByName('Failure Scenario'); if(!ws)return;
  // Clear ALL data rows to remove inactive modules
  const lastRow=ws.getMaxRows();
  if(lastRow>=3)ws.getRange(3,1,lastRow-2,8).clearContent().clearFormat();
  const all=[];
  allData.forEach(d=>d.blockers.forEach(b=>all.push({...b,refreshed:d.refreshed})));
  if(all.length===0){
    ws.getRange(3,1,1,8).merge().setValue('✅ Tidak ada failure scenario! Semua Priority Critical/High/Medium TC passed.')
        .setBackground('#C8E6C9').setFontColor('#1B5E20').setFontWeight('bold')
        .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.setRowHeight(3,28);
  } else {
    all.sort((a,b)=>{if(a.status!==b.status)return a.status==='FAILED'?-1:1;if(a.prio!==b.prio)return a.prio==='Critical'?-1:1;return a.module.localeCompare(b.module);});
    const rules=[];
    all.forEach((b,i)=>{
      const r=3+i,bg=i%2===0?'#FFF8F8':'#FFFFFF';
      [b.module,b.type,b.tcId,b.prio,b.status,b.feature,b.scenario,
        Utilities.formatDate(b.refreshed,Session.getScriptTimeZone(),'dd/MM HH:mm')
      ].forEach((v,ci)=>ws.getRange(r,ci+1).setValue(v).setBackground(bg).setFontFamily('Arial').setFontSize(9)
          .setHorizontalAlignment(ci>1&&ci<6?'center':'left').setVerticalAlignment('middle').setWrap(ci===6)
          .setBorder(true,true,true,true,false,false,'#E57373',SpreadsheetApp.BorderStyle.SOLID));
      ws.setRowHeight(r,20);
      const sR=ws.getRange(r,5),pR=ws.getRange(r,4);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FAILED').setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true).setRanges([sR]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('BLOCKED').setBackground('#FFE0B2').setFontColor('#E65100').setBold(true).setRanges([sR]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Critical').setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true).setRanges([pR]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('High').setBackground('#FFE0B2').setFontColor('#E65100').setBold(true).setRanges([pR]).build());
    });
    ws.setConditionalFormatRules(rules);
  }
  ws.getRange(1,1,1,8).clear();
  ws.getRange(1,1,1,8).merge().setValue(
      'BLOCKER ALERT  |  Total:'+all.length+'  FAILED:'+all.filter(b=>b.status==='FAILED').length+
      '  BLOCKED:'+all.filter(b=>b.status==='BLOCKED').length+
      '  Critical:'+all.filter(b=>b.prio==='Critical').length+
      '  High:'+all.filter(b=>b.prio==='High').length
  ).setBackground('#B71C1C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
}


// ═══════════════════════════════════════════════════════════════════════
// COVERAGE TAB
// ═══════════════════════════════════════════════════════════════════════

function buildCoverage(ss) {
  const ws=ss.insertSheet('Coverage'); ws.setTabColor('#1B5E20'); ws.clear();
  [90,85,100,60,60,60,60,65].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  ws.getRange(1,1,1,8).merge().setValue('COVERAGE PER SUBMODUL  —  All Modules')
      .setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(12).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,28);
  ['Project','Modul','Submodul','Type','Total','Passed','Failed','Auto%'].forEach((h,i)=>
      ws.getRange(2,i+1).setValue(h).setBackground('#2E7D32').setFontColor('#FFFFFF')
          .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center')
          .setBorder(true,true,true,true,false,false,'#81C784',SpreadsheetApp.BorderStyle.SOLID));
  ws.setRowHeight(2,20);
  ws.getRange(3,1,1,8).merge().setValue('▶ Run refreshDashboard()')
      .setBackground('#F1F8E9').setFontColor('#33691E').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(2);
}

function writeCoverage(ss, allData) {
  const ws=ss.getSheetByName('Coverage'); if(!ws)return;
  // Clear ALL data rows to remove inactive modules
  const lastRow=ws.getMaxRows();
  if(lastRow>=3)ws.getRange(3,1,lastRow-2,8).clearContent().clearFormat();
  let r=3; const rules=[];
  allData.forEach(d=>{
    if(!d.coverage||d.coverage.length===0)return;
    d.coverage.forEach((cov,i)=>{
      const bg=i%2===0?'#F1F8E9':'#FFFFFF';
      const autoRate=cov.total>0?cov.auto/cov.total:0;
      [d.project||'',d.module||'',d.submodule||d.name,cov.type,cov.total,cov.passed,cov.failed,autoRate].forEach((v,ci)=>{
        const c=ws.getRange(r,ci+1).setValue(v).setBackground(bg).setFontFamily('Arial').setFontSize(9)
            .setHorizontalAlignment(ci<3?'left':'center').setVerticalAlignment('middle')
            .setBorder(true,true,true,true,false,false,'#81C784',SpreadsheetApp.BorderStyle.SOLID);
        if(ci===7)c.setNumberFormat('0%');
      });
      rules.push(...ragRules_(ws.getRange(r,8),0.8,0.5));
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
          .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true).setRanges([ws.getRange(r,7)]).build());
      ws.setRowHeight(r,20); r++;
    });
  });
  ws.setConditionalFormatRules(rules);
}


// ═══════════════════════════════════════════════════════════════════════
// HISTORY TAB — trend data + chart
// ═══════════════════════════════════════════════════════════════════════

function buildHistory(ss) {
  const ws=ss.insertSheet('History'); ws.setTabColor('#4A148C'); ws.clear();
  const hdrs=getHistoryHeaders_();
  ensureSheetColumns_(ws, hdrs.length);
  ws.getRange(1,1,2,hdrs.length).clearContent().breakApart();
  ws.getRange(1,1,1,hdrs.length).merge().setValue('HISTORY  —  Trend Data (auto-appended setiap refresh)')
      .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold')
      .setBackground('#6A1B9A').setFontColor('#FFFFFF');
  ws.setFrozenRows(2);
  ws.setColumnWidth(1,130);
  [2,3,4,5].forEach(c=>ws.setColumnWidth(c,90));
  for(let c=6;c<=hdrs.length;c++)ws.setColumnWidth(c,72);
}

function buildAutomationRuns(ss) {
  const existing = ss.getSheetByName('Automation Runs');
  const ws = existing || ss.insertSheet('Automation Runs');
  ws.setTabColor('#00695C');
  ws.clear();
  const hdrs = getAutomationRunHeaders_();
  ensureSheetColumns_(ws, hdrs.length);
  ws.getRange(1,1,1,hdrs.length).merge()
      .setValue('AUTOMATION RUNS  —  Raw Jenkins automation execution results')
      .setBackground('#00695C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold')
      .setBackground('#00897B').setFontColor('#FFFFFF');
  ws.setFrozenRows(2);
  ws.setColumnWidth(1,130);
  [2,3,4,5,6,7,8,9,10,11,12,13].forEach(c=>ws.setColumnWidth(c,120));
  [14,15].forEach(c=>ws.setColumnWidth(c,260));
  for(let c=16;c<=hdrs.length;c++)ws.setColumnWidth(c,90);
}

function getAutomationRunHeaders_() {
  return ['Timestamp','Project','Modul','Submodul','Channel','Suite','Environment',
    'Contract Key','Tag','Job Name','Build Number','Build URL','Report URL',
    'Status','Total','Passed','Failed','Skipped','Broken','Flaky','Pass Rate','Source','Raw Payload'];
}

function ensureAutomationRunsSheet_(ss) {
  let ws = ss.getSheetByName('Automation Runs');
  if (!ws) {
    buildAutomationRuns(ss);
    ws = ss.getSheetByName('Automation Runs');
  }
  const hdrs = getAutomationRunHeaders_();
  ensureSheetColumns_(ws, hdrs.length);
  const current = ws.getRange(2,1,1,hdrs.length).getValues()[0];
  if (current.join('|') !== hdrs.join('|')) {
    ws.getRange(1,1,2,hdrs.length).breakApart().clearContent();
    ws.getRange(1,1,1,hdrs.length).merge()
        .setValue('AUTOMATION RUNS  —  Raw Jenkins automation execution results')
        .setBackground('#00695C').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold')
        .setBackground('#00897B').setFontColor('#FFFFFF');
    ws.setFrozenRows(2);
  }
  return ws;
}

function ensureSheetColumns_(ws, requiredColumns) {
  const currentColumns = ws.getMaxColumns();
  if (currentColumns < requiredColumns) {
    ws.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
  }
}

function setupExternalQAConfigSection_(ws) {
  const externalQACol = 38; // AL
  ensureSheetColumns_(ws, externalQACol + 2);

  try { ws.getRange(1, externalQACol, 2, 3).breakApart(); } catch(e) {}
  ws.getRange(1, externalQACol, 1, 3).merge()
    .setValue('EXTERNAL QA')
    .setBackground('#455A64').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.getRange(2, externalQACol, 1, 3).merge()
    .setValue('Flag untuk scope yang dites external team dan apakah perlu retest QA Peruri. Detail evidence ada di QATM tab External Test Report.')
    .setBackground('#ECEFF1').setFontColor('#455A64').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  const headers = [
    ['External QA', 90, 'TRUE = testing utama dilakukan external team.\nFALSE = testing internal QA Peruri seperti biasa.'],
    ['Retest QA Peruri', 120, 'TRUE = tetap perlu retest internal QA Peruri.\nFALSE = tidak perlu retest internal; hasil test QATM boleh kosong jika External QA TRUE.'],
    ['External QA Notes', 220, 'Catatan tambahan di level dashboard config. Scope / Notes utama tetap di QATM Summary.']
  ];
  headers.forEach(([h, w, note], i) => {
    const col = externalQACol + i;
    const headerCell = ws.getRange(3, col);
    headerCell
      .setValue(h)
      .setBackground('#607D8B').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#B0BEC5', SpreadsheetApp.BorderStyle.SOLID);
    ws.setColumnWidth(col, w);
    if (note) headerCell.setNote(note);
  });

  const dvBool = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ws.getRange(4, externalQACol, 997, 2).setDataValidation(dvBool);
  ws.getRange(4, externalQACol, 997, 2).setHorizontalAlignment('center').setFontWeight('bold');
}

function getHistoryHeaders_() {
  return ['Timestamp','Project','Modul','Submodul','PIC QA',
    'webTestCases','webPassed','webFailed','webBlocked','webPassRate','webExecRate','webAutomationRate',
    'apiTestCases','apiPassed','apiFailed','apiBlocked','apiPassRate','apiExecRate','apiAutomationRate',
    'smokeWebTestCases','smokeWebPassed','smokeWebFailed','smokeWebBlocked','smokeWebPassRate','smokeWebExecRate',
    'smokeApiTestCases','smokeApiPassed','smokeApiPassRate','smokeApiExecRate',
    'Perf',
    'totalBugs','critical','high','medium','low','lowest','blocker',
    'dev','uat','prod',
    'open','inProgress','fixed','reopen','verified',
    'blkOpen','blkInProgress','blkFixed','blkReopen','blkVerified',
    'prodBlocker','prodCritical','prodHigh','prodMedium','prodLow','prodLowest',
    'prodOpen','prodInProgress','prodFixed','prodReopen','prodVerified',
    'healthScore',
    'webInProgress','webTodo','webExecuted','webNotRun',
    'apiInProgress','apiTodo','apiExecuted','apiNotRun',
    'totalExecuted','totalNotRun','totalTodo','totalExecutable','totalExecutionRate',
    'smokeWebInProgress','smokeWebTodo','smokeApiInProgress','smokeApiTodo',
    'webAutomationPassed','webAutomationFailed','webAutomationPassRate','webAutomationStatus',
    'apiAutomationPassed','apiAutomationFailed','apiAutomationPassRate','apiAutomationStatus',
    'webDevPassed','webDevFailed','webDevPassRate','webDevStatus',
    'apiDevPassed','apiDevFailed','apiDevPassRate','apiDevStatus',
    'webStgPassed','webStgFailed','webStgPassRate','webStgStatus',
    'apiStgPassed','apiStgFailed','apiStgPassRate','apiStgStatus',
    'webProdPassed','webProdFailed','webProdPassRate','webProdStatus',
    'apiProdPassed','apiProdFailed','apiProdPassRate','apiProdStatus',
    'externalQA','retestQAPeruri','scopeNotes','externalQANotes',
    'externalTeam','externalStatusReview','functionalEvidenceUrl','functionalReviewStatus',
    'performanceEvidenceUrl','performanceReviewStatus','vaptEvidenceUrl','vaptReviewStatus',
    'externalOverallStatus','externalReviewer','externalReviewDate','externalReportNotes'];
}

function ensureHistoryHeaders_(ws) {
  const hdrs = getHistoryHeaders_();
  ensureSheetColumns_(ws, hdrs.length);
  const maxCols = Math.max(ws.getLastColumn(), hdrs.length);

  ws.getRange(1,1,2,maxCols).breakApart().clearContent();
  ws.getRange(1,1,1,hdrs.length).merge().setValue('HISTORY  —  Trend Data (auto-appended setiap refresh)')
      .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold')
      .setBackground('#6A1B9A').setFontColor('#FFFFFF');
  ws.setFrozenRows(2);
  ws.setColumnWidth(1,130);
  [2,3,4,5].forEach(c=>ws.setColumnWidth(c,90));
  for(let c=6;c<=hdrs.length;c++)ws.setColumnWidth(c,72);
}

function appendHistory(ss, allData) {
  const ws=ss.getSheetByName('History'); if(!ws)return;
  ensureHistoryHeaders_(ws);
  const hdrs = getHistoryHeaders_();
  const HISTORY_COLS = hdrs.length;
  const now = new Date();
  const ts=Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');
  const today=Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const automationRunsByKey = getLatestAutomationRunsByDashboardKey_(ss);

  // Build all rows at once (batch operation)
  const rows = allData.map(d => {
    const bs=d.bugStats||{};
    const webDevRun = getAutomationRunForModule_(automationRunsByKey, d, 'web', 'dev');
    const apiDevRun = getAutomationRunForModule_(automationRunsByKey, d, 'api', 'dev');
    const webStgRun = getAutomationRunForModule_(automationRunsByKey, d, 'web', 'stg');
    const apiStgRun = getAutomationRunForModule_(automationRunsByKey, d, 'api', 'stg');
    const webProdRun = getAutomationRunForModule_(automationRunsByKey, d, 'web', 'prod');
    const apiProdRun = getAutomationRunForModule_(automationRunsByKey, d, 'api', 'prod');
    const webRun = firstAutomationRun_(webProdRun, webStgRun, webDevRun);
    const apiRun = firstAutomationRun_(apiProdRun, apiStgRun, apiDevRun);
    const prodBlocker = bs.prodBlockerBugs || 0;
    const prodCritical = bs.prodCriticalBugs || 0;
    const prodHigh = bs.prodHighBugs || 0;
    const prodMedium = bs.prodMediumBugs || 0;
    const avgQualityRate = ((d.aPassRate || 0) + (d.wPassRate || 0) + (d.aAutoRate || 0) + (d.wAutoRate || 0) + (d.wExecRate || 0) + (d.aExecRate || 0)) / 6;
    const healthScore = Math.max(0, Math.round((avgQualityRate * 100) - (prodCritical * 8) - (prodHigh * 5) - (prodMedium * 2)));
    return [ts,d.project||'',d.module||'',d.submodule||d.name,d.team||'',
      d.wTotal||0,d.wPassed||0,d.wFailed||0,d.wBlocked||0,d.wPassRate||0,d.wExecRate||0,d.wAutoRate||0,
      d.aTotal||0,d.aPassed||0,d.aFailed||0,d.aBlocked||0,d.aPassRate||0,d.aExecRate||0,d.aAutoRate||0,
      d.wSmokeTotal||0,d.wSmokePassed||0,d.wSmokeFailed||0,d.wSmokeBlocked||0,d.wSmokePassRate||0,d.wSmokeExecRate||0,
      d.aSmokeTotal||0,d.aSmokePassed||0,d.aSmokePassRate||0,d.aSmokeExecRate||0,
      d.perfResult,
      bs.total||0,bs.critical||0,bs.high||0,bs.medium||0,bs.low||0,bs.lowest||0,bs.blocker||0,
      bs.devBugs||0,bs.uatBugs||0,bs.prodBugs||0,
      bs.open||0,bs.inprog||0,bs.fixed||0,bs.reopen||0,bs.verified||0,
      bs.blockerOpenBugs||0,bs.blockerInProgressBugs||0,bs.blockerFixedBugs||0,bs.blockerReopenBugs||0,bs.blockerVerifiedBugs||0,
      prodBlocker,prodCritical,prodHigh,prodMedium,bs.prodLowBugs||0,bs.prodLowestBugs||0,
      bs.prodOpenBugs||0,bs.prodInProgressBugs||0,bs.prodFixedBugs||0,bs.prodReopenBugs||0,bs.prodVerifiedBugs||0,
      healthScore,
      d.wInProg||0,d.wTodo||0,(d.wPassed||0)+(d.wFailed||0)+(d.wBlocked||0),Math.max(0,(d.wTotal||0)-((d.wPassed||0)+(d.wFailed||0)+(d.wBlocked||0))),
      d.aInProg||0,d.aTodo||0,(d.aPassed||0)+(d.aFailed||0)+(d.aBlocked||0),Math.max(0,(d.aTotal||0)-((d.aPassed||0)+(d.aFailed||0)+(d.aBlocked||0))),
      ((d.wPassed||0)+(d.wFailed||0)+(d.wBlocked||0)+(d.aPassed||0)+(d.aFailed||0)+(d.aBlocked||0)),
      Math.max(0,((d.wTotal||0)+(d.aTotal||0))-((d.wPassed||0)+(d.wFailed||0)+(d.wBlocked||0)+(d.aPassed||0)+(d.aFailed||0)+(d.aBlocked||0))),
      (d.wTodo||0)+(d.aTodo||0),(d.wTotal||0)+(d.aTotal||0),((d.wTotal||0)+(d.aTotal||0))>0?(((d.wPassed||0)+(d.wFailed||0)+(d.wBlocked||0)+(d.aPassed||0)+(d.aFailed||0)+(d.aBlocked||0))/((d.wTotal||0)+(d.aTotal||0))):0,
      d.wSmokeInProg||0,d.wSmokeTodo||0,d.aSmokeInProg||0,d.aSmokeTodo||0,
      webRun ? webRun.passed : '',webRun ? webRun.failed : '',webRun ? webRun.passRate : '',webRun ? webRun.status : 'Coming Soon',
      apiRun ? apiRun.passed : '',apiRun ? apiRun.failed : '',apiRun ? apiRun.passRate : '',apiRun ? apiRun.status : 'Coming Soon',
      webDevRun ? webDevRun.passed : '',webDevRun ? webDevRun.failed : '',webDevRun ? webDevRun.passRate : '',webDevRun ? webDevRun.status : 'Coming Soon',
      apiDevRun ? apiDevRun.passed : '',apiDevRun ? apiDevRun.failed : '',apiDevRun ? apiDevRun.passRate : '',apiDevRun ? apiDevRun.status : 'Coming Soon',
      webStgRun ? webStgRun.passed : '',webStgRun ? webStgRun.failed : '',webStgRun ? webStgRun.passRate : '',webStgRun ? webStgRun.status : 'Coming Soon',
      apiStgRun ? apiStgRun.passed : '',apiStgRun ? apiStgRun.failed : '',apiStgRun ? apiStgRun.passRate : '',apiStgRun ? apiStgRun.status : 'Coming Soon',
      webProdRun ? webProdRun.passed : '',webProdRun ? webProdRun.failed : '',webProdRun ? webProdRun.passRate : '',webProdRun ? webProdRun.status : 'Coming Soon',
      apiProdRun ? apiProdRun.passed : '',apiProdRun ? apiProdRun.failed : '',apiProdRun ? apiProdRun.passRate : '',apiProdRun ? apiProdRun.status : 'Coming Soon',
      (d.externalQA||{}).isExternal === true,(d.externalQA||{}).retestByPeruri === true,d.scopeNotes || '',(d.externalQA||{}).notes || '',
      (d.externalTestReport||{}).externalTeam || '',(d.externalTestReport||{}).statusReview || '',
      (d.externalTestReport||{}).functionalEvidenceUrl || '',(d.externalTestReport||{}).functionalReviewStatus || '',
      (d.externalTestReport||{}).performanceEvidenceUrl || '',(d.externalTestReport||{}).performanceReviewStatus || '',
      (d.externalTestReport||{}).vaptEvidenceUrl || '',(d.externalTestReport||{}).vaptReviewStatus || '',
      (d.externalTestReport||{}).overallStatus || '',(d.externalTestReport||{}).reviewer || '',
      (d.externalTestReport||{}).reviewDate || '',(d.externalTestReport||{}).notes || ''];
  });

  // SMART APPEND: Check if today's data already exists
  const lastRow = ws.getLastRow();
  if(lastRow>=3){
    const existingData = ws.getRange(3,1,lastRow-2,HISTORY_COLS).getValues();
    const todayRows = {};  // Map: "project-module-submodule" => row index

    // Find all rows from today
    existingData.forEach((row,i)=>{
      const rowDate = Utilities.formatDate(new Date(row[0]),Session.getScriptTimeZone(),'yyyy-MM-dd');
      if(rowDate===today){
        const key=`${row[1]||''}-${row[2]||''}-${row[3]||''}`;  // project-module-submodule
        todayRows[key]=i+3;  // +3 because row 1-2 are headers
      }
    });

    // Update existing rows or collect new rows to append
    const newRows=[];
    rows.forEach(row=>{
      const key=`${row[1]||''}-${row[2]||''}-${row[3]||''}`;
      if(todayRows[key]){
        // FULL UPDATE - Replace entire row with latest data
        // This ensures ALL columns (test execution, bugs, automation) are updated
        ws.getRange(todayRows[key],1,1,HISTORY_COLS).setValues([row]);
      }else{
        // Collect for batch append
        newRows.push(row);
      }
    });

    // Append only new rows
    if(newRows.length>0){
      const startRow=lastRow+1;
      ws.getRange(startRow,1,newRows.length,HISTORY_COLS).setValues(newRows);
    }
  }else{
    // No data yet, just append
    ws.getRange(3,1,rows.length,HISTORY_COLS).setValues(rows);
  }

  // Apply number formatting to percentage columns
  const newLastRow = ws.getLastRow();
  if(newLastRow>=3){
    for(const col of getHistoryPercentageColumns_(hdrs))ws.getRange(3,col,newLastRow-2,1).setNumberFormat('0%');
  }

  // Charts removed - will use Web App dashboard for visualization
}

/**
 * DEPRECATED: No longer used - keeping for reference
 * Previously used to merge only automation columns, but now we do full row update
 */
function mergeHistoryAutomationColumns_(existingRow, latestRow, hdrs) {
  const merged = existingRow.slice();
  merged[0] = latestRow[0];
  [
    'webAutomationPassed',
    'webAutomationFailed',
    'webAutomationPassRate',
    'webAutomationStatus',
    'apiAutomationPassed',
    'apiAutomationFailed',
    'apiAutomationPassRate',
    'apiAutomationStatus',
    'webDevPassed',
    'webDevFailed',
    'webDevPassRate',
    'webDevStatus',
    'apiDevPassed',
    'apiDevFailed',
    'apiDevPassRate',
    'apiDevStatus',
    'webStgPassed',
    'webStgFailed',
    'webStgPassRate',
    'webStgStatus',
    'apiStgPassed',
    'apiStgFailed',
    'apiStgPassRate',
    'apiStgStatus',
    'webProdPassed',
    'webProdFailed',
    'webProdPassRate',
    'webProdStatus',
    'apiProdPassed',
    'apiProdFailed',
    'apiProdPassRate',
    'apiProdStatus',
    'externalQA',
    'retestQAPeruri',
    'scopeNotes',
    'externalQANotes',
    'externalTeam',
    'externalStatusReview',
    'functionalEvidenceUrl',
    'functionalReviewStatus',
    'performanceEvidenceUrl',
    'performanceReviewStatus',
    'vaptEvidenceUrl',
    'vaptReviewStatus',
    'externalOverallStatus',
    'externalReviewer',
    'externalReviewDate',
    'externalReportNotes'
  ].forEach(header => {
    const index = hdrs.indexOf(header);
    if (index >= 0) merged[index] = latestRow[index];
  });
  return merged;
}

function getHistoryPercentageColumns_(hdrs) {
  return hdrs
    .map((header, index) => /(?:Rate|PassRate)$/.test(header) ? index + 1 : null)
    .filter(Boolean);
}

function firstAutomationRun_() {
  for (let i = 0; i < arguments.length; i++) {
    if (arguments[i]) return arguments[i];
  }
  return null;
}

function getLatestAutomationRunsByDashboardKey_(ss) {
  const ws = ss.getSheetByName('Automation Runs');
  const byKey = {};
  if (!ws || ws.getLastRow() < 3) return byKey;

  const values = ws.getRange(3,1,ws.getLastRow()-2,Math.min(ws.getLastColumn(), getAutomationRunHeaders_().length)).getValues();
  values.forEach(row => {
    const timestamp = parseAutomationTimestamp_(row[0]);
    if (!timestamp) return;

    const channel = normalizeAutomationValue_(row[4]);
    if (channel !== 'web' && channel !== 'api') return;

    const run = {
      timestamp,
      project: String(row[1] || ''),
      module: String(row[2] || ''),
      submodule: String(row[3] || ''),
      channel,
      environment: normalizeAutomationEnvironment_(row[6]),
      contractKey: String(row[7] || ''),
      tag: String(row[8] || ''),
      jobName: String(row[9] || ''),
      status: String(row[13] || ''),
      total: Number(row[14]) || 0,
      passed: Number(row[15]) || 0,
      failed: Number(row[16]) || 0,
      skipped: Number(row[17]) || 0,
      broken: Number(row[18]) || 0,
      flaky: Number(row[19]) || 0,
      passRate: parseRate_(row[20])
    };

    const aliases = getAutomationRunAliases_(run);
    aliases.forEach(alias => {
      const key = run.environment + '|' + channel + '|' + alias;
      if (!byKey[key] || byKey[key].timestamp < timestamp) byKey[key] = run;
    });
  });

  return byKey;
}

function getAutomationRunForModule_(runsByKey, moduleData, channel, environment) {
  const aliases = getDashboardAutomationAliases_(moduleData, channel);
  const normalizedEnvironment = normalizeAutomationEnvironment_(environment);
  for (let i = 0; i < aliases.length; i++) {
    const run = runsByKey[normalizedEnvironment + '|' + channel + '|' + aliases[i]];
    if (run) return run;
  }
  return null;
}

function getDashboardAutomationAliases_(moduleData, channel) {
  const aliases = [];
  const add = value => {
    const normalized = normalizeAutomationValue_(value);
    if (normalized && aliases.indexOf(normalized) === -1) aliases.push(normalized);
  };

  add([moduleData.project, moduleData.module, moduleData.submodule || moduleData.name].join('|'));
  add([moduleData.project, moduleData.module].join('|'));
  add([moduleData.module, moduleData.submodule || moduleData.name].join('|'));
  add(moduleData.submodule || moduleData.name);
  add(moduleData.module);

  const contracts = moduleData.automationContracts || {};
  add(contracts[channel]);
  add(contracts.all);

  return aliases;
}

function getAutomationRunAliases_(run) {
  const aliases = [];
  const add = value => {
    const normalized = normalizeAutomationValue_(value);
    if (normalized && aliases.indexOf(normalized) === -1) aliases.push(normalized);
  };

  add([run.project, run.module, run.submodule].join('|'));
  add([run.project, run.module].join('|'));
  add([run.module, run.submodule].join('|'));
  add(run.submodule);
  add(run.module);
  add(run.contractKey);
  add(run.tag);
  add(run.jobName);

  return aliases;
}

function normalizeAutomationValue_(value) {
  return String(value || '')
    .trim()
    .replace(/^@+/, '')
    .replace(/([a-z0-9])([A-Z])/g, '$1-$2')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function normalizeAutomationEnvironment_(value) {
  const normalized = normalizeAutomationValue_(value);
  if (normalized === 'staging' || normalized === 'stage') return 'stg';
  if (normalized === 'production') return 'prod';
  if (normalized === 'development') return 'dev';
  return normalized;
}

function parseAutomationTimestamp_(value) {
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function parseRate_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  if (typeof value === 'number') return value > 1 ? value / 100 : value;
  const text = String(value).trim();
  if (text.endsWith('%')) {
    const parsed = parseFloat(text.slice(0, -1));
    return isNaN(parsed) ? 0 : parsed / 100;
  }
  const parsed = parseFloat(text);
  return isNaN(parsed) ? 0 : (parsed > 1 ? parsed / 100 : parsed);
}

/**
 * Cleanup History tab - keep only 90 days of data, one entry per day per module
 * Manual trigger from menu
 */
function cleanupHistoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('History');
  if(!ws){
    SpreadsheetApp.getUi().alert('History tab not found');
    return;
  }

  const lastRow = ws.getLastRow();
  if(lastRow<3){
    SpreadsheetApp.getUi().alert('No data to cleanup');
    return;
  }

  // Get all data
  const historyCols = ws.getLastColumn();
  const data = ws.getRange(3,1,lastRow-2,historyCols).getValues();
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 90);  // 90 days ago

  // Group by date+module, keep latest entry per day
  const dailyData = {};  // Map: "date|project|module|submodule" => row data

  data.forEach(row=>{
    const ts = new Date(row[0]);
    if(ts < cutoffDate) return;  // Skip old data

    const date = Utilities.formatDate(ts,Session.getScriptTimeZone(),'yyyy-MM-dd');
    const key = `${date}|${row[1]||''}|${row[2]||''}|${row[3]||''}`;

    // Keep only latest entry for this day+module
    if(!dailyData[key] || new Date(dailyData[key][0]) < ts){
      dailyData[key] = row;
    }
  });

  // Convert back to array and sort by date (oldest first)
  const cleanedData = Object.values(dailyData).sort((a,b)=> new Date(a[0]) - new Date(b[0]));

  const rowsDeleted = data.length - cleanedData.length;

  if(rowsDeleted===0){
    SpreadsheetApp.getUi().alert('No duplicate or old data found. History is already clean!');
    return;
  }

  // Clear all data and rewrite cleaned data
  ws.getRange(3,1,lastRow-2,historyCols).clearContent();
  if(cleanedData.length>0){
    ws.getRange(3,1,cleanedData.length,historyCols).setValues(cleanedData);
  }

  // Reapply number formatting
  const newLastRow = ws.getLastRow();
  if(newLastRow>=3){
    for(const col of [10,11,12,17,18,19,24,25,28,29])ws.getRange(3,col,newLastRow-2,1).setNumberFormat('0%');
  }

  SpreadsheetApp.getUi().alert(
    `✅ History Cleanup Complete!\n\n` +
    `Rows before: ${data.length}\n` +
    `Rows after: ${cleanedData.length}\n` +
    `Deleted: ${rowsDeleted} rows\n\n` +
    `Retention: 90 days, 1 entry per day per module`
  );
  Logger.log(`✅ History cleanup: Deleted ${rowsDeleted} rows, kept ${cleanedData.length} rows`);
}


// ═══════════════════════════════════════════════════════════════════════
// RAW TAB
// ═══════════════════════════════════════════════════════════════════════

function buildRaw(ss) {
  const ws=ss.insertSheet('_Raw'); ws.setTabColor('#546E7A'); ws.clear();
  ws.getRange(1,1).setValue('Internal cache — jangan edit manual.')
      .setBackground('#546E7A').setFontColor('#FFFFFF').setFontSize(9).setFontFamily('Arial');
  ws.setRowHeight(1,16);
}

function updateRaw(ss, allData) {
  const ws=ss.getSheetByName('_Raw'); if(!ws)return;
  ws.clearContents();
  ws.getRange(1,1).setValue('Refreshed: '+new Date());
  const hdrs=['Modul','PIC','QA Lead','Project','wTotal','wPass','wFail','wBlock','wPass%',
    'aTotal','aPass','aFail','aBlock','aPass%',
    'wSmokeTotal','wSmokePass%','wSmokeExec%',
    'aSmokeTotal','aSmokePass%','aSmokeExec%',
    'Perf','Bugs','Blocker','Error'];
  ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold').setBackground('#607D8B').setFontColor('#FFFFFF');
  allData.forEach((d,i)=>{
    const bs=d.bugStats||{};
    ws.getRange(3+i,1,1,hdrs.length).setValues([[
      d.name,d.team,d.lead,d.sprint,
      d.wTotal,d.wPassed,d.wFailed,d.wBlocked,d.wPassRate,
      d.aTotal,d.aPassed,d.aFailed,d.aBlocked,d.aPassRate,
      d.wSmokeTotal,d.wSmokePassRate,d.wSmokeExecRate,
      d.aSmokeTotal,d.aSmokePassRate,d.aSmokeExecRate,
      d.perfResult,bs.total||0,bs.blocker||0,d.error
    ]]);
  });
}


// ═══════════════════════════════════════════════════════════════════════
// UPDATE CONFIG — write back PIC + QA Lead dari Summary
// ═══════════════════════════════════════════════════════════════════════

function updateConfig(ss, allData) {
  const cfg=ss.getSheetByName('Config'); if(!cfg)return;
  const cfgData=cfg.getDataRange().getValues();
  allData.forEach(d=>{
    for(let i=3;i<cfgData.length;i++){
      if(String(cfgData[i][6]).trim()===d.id){  // col G = Spreadsheet ID
        // Update data dari QATM Summary - ALWAYS replace, even if empty
        // This ensures Config tab always reflects current QATM state
        cfg.getRange(i+1,3).setValue(d.project || '');    // col C = Project
        cfg.getRange(i+1,4).setValue(d.module || '');     // col D = Modul
        cfg.getRange(i+1,5).setValue(d.submodule || '');  // col E = Submodul
        cfg.getRange(i+1,6).setValue(d.team || '');       // col F = PIC QA
        break;
      }
    }
  });
}


// ═══════════════════════════════════════════════════════════════════════
// SHARED HELPERS
// ═══════════════════════════════════════════════════════════════════════

function ragRules_(rng, greenMin, yellowMin) {
  return [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(greenMin)
        .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(yellowMin, greenMin-0.001)
        .setBackground('#FFF9C4').setFontColor('#E65100').setBold(true).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(yellowMin)
        .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true).setRanges([rng]).build(),
  ];
}
