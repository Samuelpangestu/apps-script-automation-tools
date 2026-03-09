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
 *   Blockers — TC Critical/High yang FAILED/BLOCKED
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

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 QA Dashboard')
    .addSubMenu(ui.createMenu('📊 Dashboard')
      .addItem('Create Dashboard (First Time)', 'createDashboard')
      .addItem('Refresh Data from Modules', 'refreshDashboard')
      .addSeparator()
      .addItem('Setup Auto-Refresh Trigger (10 min)', 'setupTrigger'))
    .addSubMenu(ui.createMenu('🔔 Notifications')
      .addItem('Setup Notifications (Chat & Email)', 'setupNotifications')
      .addItem('Test Notification Now', 'sendBlockerNotification')
      .addSeparator()
      .addItem('Setup Daily Notification Trigger', 'setupDailyBlockerNotification')
      .addItem('Remove Notification Trigger', 'removeDailyBlockerNotification'))
    .addSubMenu(ui.createMenu('🔄 Jira Sync')
      .addItem('Sync All Modules from Jira', 'syncAllJira')
      .addItem('Show Jira JQL for Module', 'showJiraJQL'))
    .addSubMenu(ui.createMenu('🔧 Broadcast Fixes')
      .addItem('Fix BUG BLOCKER (Rename + Formula)', 'broadcastBugBlockerFix'))
    .addSeparator()
    .addItem('📚 Quick Start Guide', 'showQuickStartGuide')
    .addToUi();
}

/**
 * Quick Start Guide - Shows step-by-step setup instructions
 */
function showQuickStartGuide() {
  const html = `
    <div style="font-family: Arial; padding: 20px; line-height: 1.6;">
      <h2 style="color: #1976D2;">🚀 QA Dashboard - Quick Start Guide</h2>

      <h3 style="color: #0D47A1;">📋 Step 1: Create Dashboard</h3>
      <p>Menu: <b>🎯 QA Dashboard > 📊 Dashboard > Create Dashboard (First Time)</b></p>
      <p>This creates all tabs: Config, Overview, Smoke, Blockers, Coverage, History</p>

      <h3 style="color: #0D47A1;">📝 Step 2: Add Modules to Config</h3>
      <p>Go to <b>Config</b> tab and fill in:</p>
      <ul>
        <li><b>Column A (Active):</b> Check to enable module</li>
        <li><b>Column B (Jira Sync):</b> Check to enable Jira sync</li>
        <li><b>Column G (Spreadsheet ID):</b> Paste module spreadsheet ID</li>
        <li><b>Column I-J (Jira Instance & Project):</b> For Jira sync</li>
      </ul>

      <h3 style="color: #0D47A1;">🔄 Step 3: Refresh Data</h3>
      <p>Menu: <b>🎯 QA Dashboard > 📊 Dashboard > Refresh Data from Modules</b></p>
      <p>This pulls data from all active modules</p>

      <h3 style="color: #0D47A1;">🔔 Step 4: Setup Notifications (Optional)</h3>
      <p>Menu: <b>🎯 QA Dashboard > 🔔 Notifications > Setup Notifications</b></p>
      <p>Configure Google Chat webhook and/or email recipients for daily blocker alerts</p>

      <h3 style="color: #0D47A1;">⏰ Step 5: Setup Auto-Refresh (Optional)</h3>
      <p>Go to <b>Config</b> tab, scroll right to columns Q-R:</p>
      <ul>
        <li><b>Column Q (Refresh Interval):</b> Set interval in minutes (1-60)</li>
        <li><b>Column R (Enable Auto Refresh):</b> Check to enable</li>
      </ul>
      <p>Then run: <b>Menu > 🎯 QA Dashboard > 📊 Dashboard > Setup Auto-Refresh Trigger</b></p>
      <p>Dashboard will auto-refresh (including Jira sync) every X minutes</p>

      <hr style="margin: 20px 0;">
      <p style="color: #666; font-size: 12px;">
        <b>Need help?</b> Check the documentation or contact your QA team lead.
      </p>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'QA Dashboard - Quick Start Guide');
}

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
    '🔔 SETUP NOTIFICATIONS & AUTO-REFRESH\n\n' +
    'Di Config tab (scroll right), ada 3 section:\n\n' +
    '📱 GOOGLE CHAT (kolom L-N)\n' +
    '1. Buat webhook di Google Chat Space:\n' +
    '   Space Settings > Apps & integrations > Webhooks\n' +
    '2. Isi webhook URL di kolom L4\n' +
    '3. Set waktu notifikasi (jam 0-23) di M4\n' +
    '4. Check N4 untuk aktifkan\n\n' +
    '📧 EMAIL (kolom O-P)\n' +
    '1. Isi email recipients (pisah koma) di O4\n' +
    '   Contoh: dev@company.com, qa@company.com\n' +
    '2. Check P4 untuk aktifkan\n\n' +
    '🔄 AUTO REFRESH (kolom Q-R)\n' +
    '1. Set interval (1-60 menit) di Q4\n' +
    '2. Check R4 untuk aktifkan\n' +
    '3. Run: Setup Auto-Refresh Trigger dari menu\n\n' +
    '────────────────────────────────\n' +
    'Test notification:\n' +
    '🎯 QA Dashboard > 🔔 Notifications > Test Notification Now\n\n' +
    'Setup daily blocker notification:\n' +
    '🎯 QA Dashboard > 🔔 Notifications > Setup Daily Notification Trigger\n\n' +
    'Open Config tab now?';

  const response = ui.alert('Setup Notifications', msg, ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    ss.setActiveSheet(cfg);
    ss.setActiveRange(cfg.getRange('L4'));
  }
}

// ═══════════════════════════════════════════════════════════════════════
// SETUP & REFRESH
// ═══════════════════════════════════════════════════════════════════════

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Overview','Smoke','Blockers','Coverage','History','_Raw','Config','Credentials'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });
  buildConfig(ss);
  buildCredentials(ss);
  buildOverview(ss);
  buildSmoke(ss);
  buildBlockers(ss);
  buildCoverage(ss);
  buildHistory(ss);
  buildRaw(ss);
  ss.setActiveSheet(ss.getSheetByName('Config'));
  safeAlert_('Dashboard berhasil dibuat!\n\nLangkah selanjutnya:\n1. Isi tab Config dengan Spreadsheet ID modul\n2. Isi tab Credentials dengan Jira credentials\n3. Data Modul/Submodul/QA Lead/PIC QA akan otomatis dari QATM Summary\n4. Jalankan refreshDashboard()');
}

function refreshDashboard() {
  Logger.log('refreshDashboard START: ' + new Date());
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const modules = getModuleList_(ss);
  if (modules.length === 0) {
    safeAlert_('Belum ada modul aktif di Config.\nIsi tab Config dulu lalu refresh.');
    return;
  }

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

  writeOverview(ss, allData);
  writeSmoke(ss, allData);
  writeBlockers(ss, allData);
  writeCoverage(ss, allData);
  appendHistory(ss, allData);
  updateRaw(ss, allData);
  updateConfig(ss, allData);  // write back PIC + QA Lead from Summary

  const ts = 'Last refreshed: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss');
  ['Overview','Smoke'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.getRange(1,1).setValue(ts);
  });

  Logger.log('refreshDashboard DONE');
  safeAlert_('Refresh selesai! ' + allData.length + ' modul di-update.\n\nQA Lead otomatis diisi dari Summary B4 (jika tersedia).');
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

  // Delete existing trigger
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshDashboard') ScriptApp.deleteTrigger(t);
  });

  if (!enabled) {
    safeAlert_('ℹ️  Trigger auto-refresh dihapus.\n\nSet "Enable Auto Refresh" = TRUE di Config (kolom R4) untuk aktifkan kembali.');
    Logger.log('Auto-refresh trigger removed (disabled in config)');
    return;
  }

  // Validate interval (1-60 minutes)
  if (refreshInterval < 1 || refreshInterval > 60) {
    safeAlert_('❌ Invalid interval!\n\nRefresh interval harus antara 1-60 menit.\nCurrent value: ' + refreshInterval);
    return;
  }

  // Create new trigger
  ScriptApp.newTrigger('refreshDashboard').timeBased().everyMinutes(refreshInterval).create();

  Logger.log('✅ Trigger created for auto-refresh every ' + refreshInterval + ' minutes');
  safeAlert_('✅ Trigger set! Auto-refresh setiap ' + refreshInterval + ' menit.\n\n' +
             'Dashboard akan otomatis refresh (termasuk Jira sync untuk modul yang aktif).\n\n' +
             'Untuk ubah interval atau disable, edit Config kolom Q4-R4.');
}

function safeAlert_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// Wrapper functions for manual execution from menu
function rebuildConfig() { buildConfig(SpreadsheetApp.getActiveSpreadsheet()); safeAlert_('Config tab rebuilt!'); }
function rebuildCredentials() { buildCredentials(SpreadsheetApp.getActiveSpreadsheet()); safeAlert_('Credentials tab rebuilt!'); }
function rebuildOverview() { buildOverview(SpreadsheetApp.getActiveSpreadsheet()); safeAlert_('Overview tab rebuilt!'); }
function rebuildSmoke() { buildSmoke(SpreadsheetApp.getActiveSpreadsheet()); safeAlert_('Smoke tab rebuilt!'); }
function rebuildBlockers() { buildBlockers(SpreadsheetApp.getActiveSpreadsheet()); safeAlert_('Blockers tab rebuilt!'); }
function rebuildCoverage() { buildCoverage(SpreadsheetApp.getActiveSpreadsheet()); safeAlert_('Coverage tab rebuilt!'); }
function rebuildHistory() { buildHistory(SpreadsheetApp.getActiveSpreadsheet()); safeAlert_('History tab rebuilt!'); }
function rebuildRaw() { buildRaw(SpreadsheetApp.getActiveSpreadsheet()); safeAlert_('_Raw tab rebuilt!'); }

/**
 * Update Config structure: add Project column (col C) and shift existing data
 * SAFE: Does not delete existing data, only restructures columns
 */
function updateConfigStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');
  if (!cfg) {
    safeAlert_('Config tab not found! Run createDashboard() first.');
    return;
  }

  // Check if already updated (col C header = "Project")
  const c3 = String(cfg.getRange(3, 3).getValue()).trim();
  if (c3 === 'Project') {
    safeAlert_('Config already updated! Column C is already "Project".');
    return;
  }

  Logger.log('Updating Config structure: adding Project column at C');

  // Get all existing data from row 4 onwards (skip headers)
  const lastRow = Math.max(cfg.getLastRow(), 4);
  const existingData = [];

  if (lastRow >= 4) {
    // Read existing data: cols C-F (old: Modul, Submodul, QA Lead, PIC QA)
    for (let i = 4; i <= lastRow; i++) {
      existingData.push({
        row: i,
        modul: cfg.getRange(i, 3).getValue(),      // old C = Modul
        submodul: cfg.getRange(i, 4).getValue(),   // old D = Submodul
        qaLead: cfg.getRange(i, 5).getValue(),     // old E = QA Lead
        picQA: cfg.getRange(i, 6).getValue()       // old F = PIC QA
      });
    }
  }

  // Update headers
  cfg.getRange(3, 3).setValue('Project').setNote('Auto dari Summary B2 (QATM)');
  cfg.getRange(3, 4).setValue('Modul').setNote('Auto dari Summary B3 (QATM)');
  cfg.getRange(3, 5).setValue('Submodul').setNote('Auto dari Summary B4 (QATM)');
  cfg.getRange(3, 6).setValue('PIC QA').setNote('Auto dari Summary B6 (QATM)');
  cfg.setColumnWidth(3, 120); // Project (compact)
  cfg.setColumnWidth(4, 120); // Modul (compact)
  cfg.setColumnWidth(5, 140); // Submodul (compact)
  cfg.setColumnWidth(6, 110); // PIC QA (compact)

  // Shift existing data: insert empty Project column, move data right
  existingData.forEach(row => {
    cfg.getRange(row.row, 3).setValue('');           // C = Project (empty for now)
    cfg.getRange(row.row, 4).setValue(row.modul);    // D = Modul (from old C)
    cfg.getRange(row.row, 5).setValue(row.submodul); // E = Submodul (from old D)
    cfg.getRange(row.row, 6).setValue(row.picQA);    // F = PIC QA (from old F, skip QA Lead)
  });

  Logger.log('Config structure updated: ' + existingData.length + ' rows migrated');
  safeAlert_('✅ Config updated!\n\nProject column added at C.\nData shifted: Modul→D, Submodul→E, PIC QA→F\nQA Lead removed (akan otomatis dari Summary saat refresh)\n\n' + existingData.length + ' rows migrated.');
}

/**
 * Update Overview headers and layout (safe, does not delete data rows)
 * Rebuilds headers with new layout: Bugs moved after MODULE INFO
 */
function updateOverviewHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Overview');
  if (!ws) {
    safeAlert_('Overview tab not found! Run createDashboard() first.');
    return;
  }

  // Check if already updated (check if Bugs is at col 5)
  const c5 = String(ws.getRange(4, 5).getValue()).trim();
  if (c5 === 'Bugs') {
    safeAlert_('Overview layout already updated!');
    return;
  }

  Logger.log('Updating Overview layout: moving Bugs columns and adjusting widths');

  // Call initOverviewHeaders_ to rebuild headers completely
  initOverviewHeaders_(ws);

  Logger.log('Overview layout updated');
  safeAlert_('✅ Overview layout updated!\n\nNew layout:\n- Bugs moved after PIC QA (col 5-7)\n- Column widths adjusted\n- Headers: Project, Modul, Submodul, PIC QA, Bugs...\n\nJalankan refreshDashboard() untuk update data.');
}

/**
 * 🚀 ALL-IN-ONE MIGRATION SCRIPT for QA Dashboard
 *
 * Jalankan fungsi ini SEKALI untuk migrate Dashboard production ke struktur baru.
 * Script ini AMAN dan IDEMPOTENT (bisa dijalankan berkali-kali tanpa error).
 *
 * Yang dilakukan:
 * 1. Update Config: tambah kolom Project, shift data existing
 * 2. Update Overview: update headers ke Project, Modul, Submodul, PIC QA
 * 3. Refresh data dari semua QATM modules
 *
 * PREREQUISITE:
 * - Pastikan sudah jalankan broadcast fixes di semua QATM modules dulu
 * - QATM Summary sudah struktur baru (Project, Modul, Submodul, QA Lead, PIC QA)
 */
function migrateDashboardToNewStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];

  Logger.log('═══════════════════════════════════════════════════════════');
  Logger.log('🚀 DASHBOARD MIGRATION START: ' + new Date());
  Logger.log('═══════════════════════════════════════════════════════════');

  // STEP 1: Update Config structure
  Logger.log('\n📋 STEP 1: Updating Config structure...');
  try {
    const cfg = ss.getSheetByName('Config');
    if (!cfg) {
      results.push('❌ Config tab not found!');
      Logger.log('❌ Config tab not found!');
    } else {
      const c3 = String(cfg.getRange(3, 3).getValue()).trim();
      if (c3 === 'Project') {
        results.push('✅ Config already updated (skipped)');
        Logger.log('✅ Config already updated (skipped)');
      } else {
        updateConfigStructure();
        results.push('✅ Config structure updated');
        Logger.log('✅ Config structure updated');
      }
    }
  } catch(e) {
    results.push('❌ Config update failed: ' + e.message);
    Logger.log('❌ Config update failed: ' + e.message);
  }

  // STEP 2: Update Overview headers
  Logger.log('\n📊 STEP 2: Updating Overview headers...');
  try {
    const ws = ss.getSheetByName('Overview');
    if (!ws) {
      results.push('❌ Overview tab not found!');
      Logger.log('❌ Overview tab not found!');
    } else {
      const a4 = String(ws.getRange(4, 1).getValue()).trim();
      if (a4 === 'Project') {
        results.push('✅ Overview headers already updated (skipped)');
        Logger.log('✅ Overview headers already updated (skipped)');
      } else {
        updateOverviewHeaders();
        results.push('✅ Overview headers updated');
        Logger.log('✅ Overview headers updated');
      }
    }
  } catch(e) {
    results.push('❌ Overview headers update failed: ' + e.message);
    Logger.log('❌ Overview headers update failed: ' + e.message);
  }

  // STEP 3: Refresh dashboard data
  Logger.log('\n🔄 STEP 3: Refreshing dashboard data from QATM modules...');
  try {
    const modules = getModuleList_(ss);
    if (modules.length === 0) {
      results.push('⚠️ No active modules found in Config');
      Logger.log('⚠️ No active modules found in Config');
    } else {
      Logger.log('Found ' + modules.length + ' active modules');
      refreshDashboard();
      results.push('✅ Dashboard refreshed (' + modules.length + ' modules)');
      Logger.log('✅ Dashboard refreshed (' + modules.length + ' modules)');
    }
  } catch(e) {
    results.push('❌ Dashboard refresh failed: ' + e.message);
    Logger.log('❌ Dashboard refresh failed: ' + e.message);
  }

  Logger.log('\n═══════════════════════════════════════════════════════════');
  Logger.log('🎉 DASHBOARD MIGRATION COMPLETE: ' + new Date());
  Logger.log('═══════════════════════════════════════════════════════════');

  const summary = '🚀 DASHBOARD MIGRATION COMPLETE\n\n' + results.join('\n') +
                  '\n\n📝 Next Steps:\n1. Verify Config tab (Project column should be at C)\n2. Verify Overview tab (headers: Project, Modul, Submodul, PIC QA)\n3. Check data in Overview matches QATM Summary';

  Logger.log('\n' + summary);
  safeAlert_(summary);
}


// ═══════════════════════════════════════════════════════════════════════
// MODULE LIST
// ═══════════════════════════════════════════════════════════════════════

function getModuleList_(ss) {
  const cfg = ss.getSheetByName('Config');
  if (!cfg) return [];
  const data = cfg.getDataRange().getValues();
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
      jiraProj
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
  const bugr = src.getSheetByName('BugReport');

  const SUMM_KPI_ROW = 13;  // Summary row 13: main Web+API KPI values

  let projectName = '', moduleName = '', submoduleName = '', picQA = mod.team||'', qaLead = mod.lead||'';
  let wTotal=0,wPassed=0,wFailed=0,wBlocked=0,wInProg=0,wTodo=0,wPassRate=0,wAutoRate=0,wExecRate=0;
  let aTotal=0,aPassed=0,aFailed=0,aBlocked=0,aInProg=0,aTodo=0,aPassRate=0,aAutoRate=0,aExecRate=0;
  let wSmokeTotal=0,wSmokePassed=0,wSmokeFailed=0,wSmokeBlocked=0,wSmokeInProg=0,wSmokeTodo=0;
  let wSmokePassRate=0,wSmokeAutoRate=0,wSmokeExecRate=0;
  let aSmokeTotal=0,aSmokePassed=0,aSmokePassRate=0,aSmokeAutoRate=0,aSmokeExecRate=0;
  let perfResult = '--';

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

  return {
    name:mod.name, team:picQA, lead:qaLead, id:mod.id,
    project:projectName, module:moduleName, submodule:submoduleName,
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
  const empty={total:0,open:0,inprog:0,fixed:0,verified:0,critical:0,high:0,medium:0,low:0,blocker:0};
  if (!bugSheet) return empty;
  try {
    const rows=bugSheet.getDataRange().getValues().slice(4).filter(r=>r[0]&&r[0]!=='');
    const cnt=(fn)=>rows.filter(fn).length;

    // Get blocker count from Summary sheet cell A17 (BUG BLOCKER row)
    let blockerCount = cnt(r=>['Open','In Progress','Reopen','Fixed','Verified'].includes(r[3])&&['Critical','High','Medium'].includes(r[2]));

    // Try to read from Summary sheet A17 if available (more accurate - uses formula)
    if (summarySheet) {
      try {
        const blockerCell = summarySheet.createTextFinder('BUG BLOCKER').matchEntireCell(false).findNext();
        if (blockerCell) {
          const blockerRow = blockerCell.getRow();
          const blockerValue = summarySheet.getRange(blockerRow + 1, 1).getValue();  // Next row, col A
          if (typeof blockerValue === 'number' && blockerValue >= 0) {
            blockerCount = blockerValue;
          }
        }
      } catch(e) {
        Logger.log('Failed to read BUG BLOCKER from Summary: ' + e.message);
      }
    }

    return {
      total:rows.length,
      open:    cnt(r=>r[3]==='Open'),
      inprog:  cnt(r=>r[3]==='In Progress'),
      fixed:   cnt(r=>r[3]==='Fixed'),
      verified:cnt(r=>r[3]==='Verified'),
      critical:cnt(r=>r[2]==='Critical'),
      high:    cnt(r=>r[2]==='High'),
      medium:  cnt(r=>r[2]==='Medium'),
      low:     cnt(r=>r[2]==='Low'),
      blocker: blockerCount,
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
    ws.getRange(r,c).setValue(txt).setBackground('#0D47A1').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setColumnWidth(c,w);
    if(note) ws.getRange(r,c).setNote(note);
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
    .setValue('💬  Notifikasi blocker otomatis dikirim ke Google Chat Space setiap hari.')
    .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  // Column headers (row 3)
  const chatHeaders = [
    ['Google Chat Webhook URL', 300, 'Buat webhook di Google Chat Space:\nSpace Settings > Apps & integrations > Webhooks\n\nFormat: https://chat.googleapis.com/v1/spaces/.../messages?key=...'],
    ['Notif Time (Hour)', 85, 'Jam berapa notifikasi dikirim (0-23)\nContoh: 15 = jam 3 sore'],
    ['Enable Notifikasi', 100, 'TRUE = aktif notifikasi harian\nFALSE = nonaktifkan']
  ];

  chatHeaders.forEach(([h, w, note], i) => {
    const col = chatCol + i;
    ws.getRange(3, col)
      .setValue(h)
      .setBackground('#1976D2').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);
    ws.setColumnWidth(col, w);
    if (note) ws.getRange(3, col).setNote(note);
  });

  // Data row with defaults (row 4)
  ws.getRange(4, chatCol, 1, 3)
    .setValues([['https://chat.googleapis.com/v1/spaces/...', 15, false]])
    .setBackground('#E3F2FD')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, '#90CAF9', SpreadsheetApp.BorderStyle.SOLID);

  ws.getRange(4, chatCol).setFontFamily('Courier New').setFontSize(8);
  ws.getRange(4, chatCol + 1).setHorizontalAlignment('center');
  ws.getRange(4, chatCol + 2).setHorizontalAlignment('center').setFontWeight('bold');

  // Data validation for Enable (checkbox) - column N (chatCol + 2)
  const dvChatEnable = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ws.getRange(4, chatCol + 2).setDataValidation(dvChatEnable);

  // Data validation for Hour (0-23) - column M (chatCol + 1)
  const dvHour = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 23).build();
  ws.getRange(4, chatCol + 1).setDataValidation(dvHour);

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
    .setValue('📧  Email notifikasi blocker harian (multiple recipients, pisah dengan koma)')
    .setBackground('#E8F5E9').setFontColor('#2E7D32').setFontStyle('italic')
    .setFontSize(8).setHorizontalAlignment('center');

  // Column headers (row 3)
  const emailHeaders = [
    ['Email Recipients (comma separated)', 260, 'Masukkan email addresses dipisah koma\nContoh: dev@company.com, qa@company.com, manager@company.com'],
    ['Enable Email', 85, 'TRUE = aktif email harian\nFALSE = nonaktifkan']
  ];

  emailHeaders.forEach(([h, w, note], i) => {
    const col = emailCol + i;
    ws.getRange(3, col)
      .setValue(h)
      .setBackground('#66BB6A').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#A5D6A7', SpreadsheetApp.BorderStyle.SOLID);
    ws.setColumnWidth(col, w);
    if (note) ws.getRange(3, col).setNote(note);
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
    ws.getRange(3, col)
      .setValue(h)
      .setBackground('#FF8F00').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true)
      .setBorder(true, true, true, true, false, false, '#FFCC80', SpreadsheetApp.BorderStyle.SOLID);
    ws.setColumnWidth(col, w);
    if (note) ws.getRange(3, col).setNote(note);
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
}


// ═══════════════════════════════════════════════════════════════════════
// CREDENTIALS TAB — untuk Jira credentials
// ═══════════════════════════════════════════════════════════════════════

function buildCredentials(ss) {
  const ws = ss.insertSheet('Credentials');
  ws.setTabColor('#4A148C');
  ws.clear();

  // Title
  ws.getRange(1,1,1,4).merge().setValue('JIRA CREDENTIALS  —  Isi email & token per instance Jira')
    .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(12).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,30);

  // Warning
  ws.getRange(2,1,1,4).merge()
    .setValue('⚠️  PENTING: Pastikan hanya owner yang bisa lihat tab ini. Jangan share spreadsheet ke publik jika token diisi di sini. Buat token di: https://id.atlassian.com/manage-profile/security/api-tokens')
    .setBackground('#F3E5F5').setFontColor('#6A1B9A').setFontStyle('italic').setFontSize(8).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  ws.setRowHeight(2,40);

  // Headers
  function hdr(c,txt,bg,w,note){
    ws.getRange(3,c).setValue(txt)
      .setBackground(bg).setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#CE93D8',SpreadsheetApp.BorderStyle.SOLID);
    ws.setColumnWidth(c,w);
    if(note) ws.getRange(3,c).setNote(note);
  }

  hdr(1,'Instance','#6A1B9A',140,'Nilai: digitalperuri atau bgn-peruri');
  hdr(2,'Email Atlassian','#7B1FA2',220,'Email login ke id.atlassian.com');
  hdr(3,'API Token','#7B1FA2',400,'Buat token di:\nhttps://id.atlassian.com/manage-profile/security/api-tokens');
  hdr(4,'Notif Email','#1565C0',280,'Penerima notif harian jam 07.00.\nBisa multiple, pisah koma.');
  ws.setRowHeight(3,24);

  // Sample data
  [['digitalperuri','email@company.com','ATATT3xFf...(paste token disini)','qa@company.com'],
   ['bgn-peruri',   'email@company.com','ATATT3xFf...(paste token disini)',''],
  ].forEach(([inst,em,tok,notif], i) => {
    const r = 4+i;
    ws.getRange(r,1,1,4).setValues([[inst,em,tok,notif]])
      .setBackground(i%2===0 ? '#F3E5F5' : '#FFFFFF')
      .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#CE93D8',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r,3).setFontFamily('Courier New').setFontSize(8);
    ws.setRowHeight(r,24);
  });

  // Data validation untuk Instance
  const dvInstance = SpreadsheetApp.newDataValidation()
    .requireValueInList(['digitalperuri','bgn-peruri'], true).build();
  ws.getRange('A4:A100').setDataValidation(dvInstance);

  // Hide tab from normal users
  ws.protect().setDescription('Credentials - Only owners should edit').setWarningOnly(true);
}


// ═══════════════════════════════════════════════════════════════════════
// OVERVIEW TAB — build + write + charts
// ═══════════════════════════════════════════════════════════════════════

function buildOverview(ss) {
  const ws = ss.insertSheet('Overview');
  ws.setTabColor('#0D47A1');
  ws.clear();
  initOverviewHeaders_(ws);
  ws.getRange(5,1,1,25).merge()
      .setValue('▶ Run refreshDashboard() untuk mengisi data')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(4);
}

function initOverviewHeaders_(ws) {
  const lastCol = Math.max(ws.getLastColumn()||1, 25);
  try { ws.getRange(1,1,4,lastCol).breakApart(); } catch(e) {}
  ws.getRange(1,1,4,lastCol).clearContent().clearFormat();

  function h_(r,c,nr,nc,txt,bg,fg,sz){
    const rng=(nr>1||nc>1)?ws.getRange(r,c,nr,nc).merge():ws.getRange(r,c);
    rng.setValue(txt||'').setBackground(bg||'#0D47A1').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  }

  // Col widths — 25 cols (COMPACT VERSION)
  // NEW LAYOUT: Project, Modul, Submodul, PIC QA | BUGS (3) | WEB (5) | SMOKE WEB (3) | API (5) | SMOKE API (3) | PERF | NOTES
  [80,80,90,80, 48,52,56, 48,52,48,48,60, 56,60,52, 48,52,48,48,60, 56,60,52, 60, 140]
      .forEach((w,i)=>ws.setColumnWidth(i+1,w));

  // Row 1 — last refresh
  ws.getRange(1,1,1,25).merge().setValue('Last refreshed: —')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  // Row 2 — title
  h_(2,1,1,25,'QA DASHBOARD  |  PORTFOLIO OVERVIEW','#0D47A1','#FFFFFF',13);
  ws.setRowHeight(2,30);

  // Row 3 — group headers
  // NEW: Bugs moved after MODULE INFO
  h_(3,1, 1,4, 'MODULE INFO',    '#263238');
  h_(3,5, 1,3, 'BUGS',           '#B71C1C');
  h_(3,8, 1,5, 'WEB / MOBILE',   '#1565C0');
  h_(3,13,1,3, '🔥 SMOKE WEB',   '#BF360C');
  h_(3,16,1,5, 'API',             '#283593');
  h_(3,21,1,3, '🔥 SMOKE API',   '#4A148C');
  h_(3,24,1,1, 'PERF',            '#004D40');
  h_(3,25,1,1, 'NOTES',           '#37474F');
  ws.setRowHeight(3,22);

  // Row 4 — column headers
  // NEW: Bugs moved after PIC QA
  ['Project','Modul','Submodul','PIC QA',
    'Bugs','Blocker','Critical',
    'Total','Pass','Fail','Block','Pass%',
    'Total','Pass%','Exec%',
    'Total','Pass','Fail','Block','Pass%',
    'Total','Pass%','Exec%',
    'Perf','Notes'
  ].forEach((lbl,i)=>h_(4,i+1,1,1,lbl,'#1565C0'));
  ws.getRange(4,13).setNote('Smoke Web: TC Priority Critical+High+Medium');
  ws.getRange(4,14).setNote('Smoke Web Pass Rate (target ≥80%)');
  ws.getRange(4,15).setNote('Smoke Web Exec Rate (% TC sudah ada hasil)');
  ws.getRange(4,21).setNote('Smoke API: TC Priority Critical+High+Medium');
  ws.getRange(4,22).setNote('Smoke API Pass Rate (target ≥80%)');
  ws.getRange(4,23).setNote('Smoke API Exec Rate');
  ws.setRowHeight(4,26);
  ws.setFrozenRows(4);
}

function writeOverview(ss, allData) {
  let ws = ss.getSheetByName('Overview');
  if (!ws) { buildOverview(ss); ws = ss.getSheetByName('Overview'); }

  initOverviewHeaders_(ws);  // safe rebuild — breakApart dulu

  const lastRow = Math.max(ws.getLastRow(),5);
  if (lastRow>=5) ws.getRange(5,1,lastRow-4,25).clearContent().clearFormat();

  const rules = [];

  allData.forEach((d,i)=>{
    const r  = 5+i;
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

    // NEW LAYOUT: Project, Modul, Submodul, PIC QA | BUGS (3) | WEB (5) | SMOKE WEB (3) | API (5) | SMOKE API (3) | PERF | NOTES
    cell(1,d.project||d.sprint||'');
    cell(2,d.module||'');
    ws.getRange(r,3).setValue(d.submodule||d.name).setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    cell(4,d.team||'');

    // Bugs (moved here - col 5-7)
    cell(5,bs.total||0); cell(6,bs.blocker||0); cell(7,bs.critical||0);

    // Web (col 8-12)
    cell(8,d.wTotal); cell(9,d.wPassed); cell(10,d.wFailed); cell(11,d.wBlocked);
    cell(12,d.error?'ERR':d.wPassRate,'0%');

    // Smoke Web (col 13-15)
    cell(13,hasSmoke?d.wSmokeTotal:'--');
    cell(14,hasSmoke?d.wSmokePassRate:'--',hasSmoke?'0%':null);
    cell(15,hasSmoke?d.wSmokeExecRate:'--',hasSmoke?'0%':null);

    // API (col 16-20)
    cell(16,d.aTotal); cell(17,d.aPassed); cell(18,d.aFailed); cell(19,d.aBlocked);
    cell(20,d.error?'ERR':d.aPassRate,'0%');

    // Smoke API (col 21-23)
    cell(21,hasSmoke?d.aSmokeTotal:'--');
    cell(22,hasSmoke?d.aSmokePassRate:'--',hasSmoke?'0%':null);
    cell(23,hasSmoke?d.aSmokeExecRate:'--',hasSmoke?'0%':null);

    // Perf (col 24)
    cell(24,d.perfResult);

    // Notes (col 25)
    ws.getRange(r,25).setValue(d.error||'').setBackground(bg).setFontFamily('Arial').setFontSize(8)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(r,22);

    // RAG Pass%
    [12,20].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.8,0.5)));
    // RAG Smoke Pass%
    [14,22].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.8,0.5)));
    // RAG Smoke Exec%
    [15,23].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.7,0.4)));
    // Failed > 0
    [10,18].forEach(col=>rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFCDD2').setFontColor('#C62828').setBold(true)
        .setRanges([ws.getRange(r,col)]).build()));
    // Blocked > 0
    [11,19].forEach(col=>rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
        .setRanges([ws.getRange(r,col)]).build()));
    // Blocker/Critical > 0
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
        .setRanges([ws.getRange(r,6),ws.getRange(r,7)]).build());
    // Perf
    [['PASS','#C8E6C9','#1B5E20'],['FAIL','#FFCDD2','#C62828'],['--','#F5F5F5','#9E9E9E']]
        .forEach(([v,bg2,fg])=>rules.push(SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(v).setBackground(bg2).setFontColor(fg).setBold(true)
            .setRanges([ws.getRange(r,24)]).build()));
  });

  // TOTAL row - updated for new layout
  if (allData.length > 0) {
    const tr = 5+allData.length;
    ws.getRange(tr,1,1,4).merge().setValue('TOTAL / AVERAGE')
        .setBackground('#E3F2FD').setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    // Bugs totals (col 5-7)
    [[5,'total'],[6,'blocker'],[7,'critical']].forEach(([col,key])=>
        ws.getRange(tr,col).setValue(allData.reduce((a,d)=>a+((d.bugStats||{})[key]||0),0))
            .setBackground('#DDEEFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center'));
    // Web totals (col 8-11)
    [[8,'wTotal'],[9,'wPassed'],[10,'wFailed'],[11,'wBlocked']].forEach(([col,key])=>{
      ws.getRange(tr,col).setValue(allData.reduce((a,d)=>a+(d[key]||0),0))
          .setBackground('#DDEEFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    });
    // API totals (col 16-19)
    [[16,'aTotal'],[17,'aPassed'],[18,'aFailed'],[19,'aBlocked']].forEach(([col,key])=>{
      ws.getRange(tr,col).setValue(allData.reduce((a,d)=>a+(d[key]||0),0))
          .setBackground('#DDEEFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    });
    // Averages for Pass%, Smoke Pass%
    const avg=(key)=>allData.reduce((a,d)=>a+(d[key]||0),0)/allData.length;
    [[12,'wPassRate'],[20,'aPassRate'],[14,'wSmokePassRate'],[22,'aSmokePassRate']].forEach(([col,key])=>
        ws.getRange(tr,col).setValue(avg(key)).setNumberFormat('0%')
            .setBackground(col>=13&&col<=15||col>=21&&col<=23?'#FFF3E0':'#DDEEFF')
            .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center'));
    ws.setRowHeight(tr,22);
  }

  ws.setConditionalFormatRules(rules);
  buildOverviewCharts_(ws, allData);
}

function buildOverviewCharts_(ws, allData) {
  if (!allData||allData.length===0) return;
  ws.getCharts().forEach(c=>ws.removeChart(c));
  const n=allData.length, dRow=5, cRow=dRow+n+4;

  // Charts now use col 2 (Modul) as labels instead of col 1 (Project)
  tryChart_(()=>ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(4,2,n+1,1))    // Modul
      .addRange(ws.getRange(4,12,n+1,1))   // Web Pass%
      .addRange(ws.getRange(4,20,n+1,1))   // API Pass%
      .setPosition(cRow,1,0,0)
      .setOption('title','Pass Rate — Web vs API (per Modul)')
      .setOption('hAxis',{title:'Pass Rate',format:'#%',minValue:0,maxValue:1})
      .setOption('colors',['#1565C0','#283593'])
      .setOption('legend',{position:'top'})
      .setOption('width',460).setOption('height',270).build()));

  tryChart_(()=>ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(4,2,n+1,1))    // Modul
      .addRange(ws.getRange(4,14,n+1,1))   // Smoke Web Pass%
      .addRange(ws.getRange(4,22,n+1,1))   // Smoke API Pass%
      .setPosition(cRow,9,0,0)
      .setOption('title','🔥 Smoke Pass Rate — Web vs API (per Modul)')
      .setOption('hAxis',{title:'Pass Rate',format:'#%',minValue:0,maxValue:1})
      .setOption('colors',['#BF360C','#4A148C'])
      .setOption('legend',{position:'top'})
      .setOption('width',460).setOption('height',270).build()));

  tryChart_(()=>ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(ws.getRange(4,2,n+1,1))    // Modul
      .addRange(ws.getRange(4,9,n+1,1))    // Web Passed
      .addRange(ws.getRange(4,10,n+1,1))   // Web Failed
      .addRange(ws.getRange(4,11,n+1,1))   // Web Blocked
      .setPosition(cRow+20,1,0,0)
      .setOption('title','Web/Mobile TC Status (per Modul)')
      .setOption('isStacked',true)
      .setOption('colors',['#4CAF50','#F44336','#FF9800'])
      .setOption('legend',{position:'top'})
      .setOption('width',460).setOption('height',250).build()));

  if (allData.some(d=>(d.bugStats||{}).blocker>0))
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(ws.getRange(4,2,n+1,1))    // Modul
        .addRange(ws.getRange(4,6,n+1,1))    // Blocker
        .addRange(ws.getRange(4,7,n+1,1))    // Critical
        .setPosition(cRow+20,9,0,0)
        .setOption('title','🚨 Open Blocker & Critical Bugs (per Modul)')
        .setOption('colors',['#FF9800','#F44336'])
        .setOption('legend',{position:'top'})
        .setOption('width',460).setOption('height',250).build()));
}


// ═══════════════════════════════════════════════════════════════════════
// SMOKE TAB — dedicated smoke view + 5 charts
// ═══════════════════════════════════════════════════════════════════════

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

  const lastRow=Math.max(ws.getLastRow(),5);
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
  buildSmokeCharts_(ws, allData);
}

function buildSmokeCharts_(ws, allData) {
  if (!allData||allData.length===0) return;
  ws.getCharts().forEach(c=>ws.removeChart(c));
  const n=allData.length, dRow=5, cRow=dRow+n+4, tmp=15;

  // Chart 1 — Smoke Pass Rate per Modul (Web + API) - use col 2
  tryChart_(()=>ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(dRow,2,n,1))    // Modul
      .addRange(ws.getRange(dRow,6,n,1))    // Smoke Web Pass%
      .addRange(ws.getRange(dRow,9,n,1))    // Smoke API Pass%
      .setPosition(cRow,1,0,0)
      .setOption('title','🔥 Smoke Pass Rate per Modul (Web & API)')
      .setOption('hAxis',{title:'Pass Rate',format:'#%',minValue:0,maxValue:1})
      .setOption('series',{0:{color:'#BF360C',labelInLegend:'Web/Mobile'},1:{color:'#4A148C',labelInLegend:'API'}})
      .setOption('legend',{position:'top'})
      .setOption('chartArea',{left:150,top:40,right:20,bottom:30})
      .setOption('width',510).setOption('height',Math.max(230,n*32+90)).build()));

  // Chart 2 — Smoke Total TC distribution (donut) - use module
  const d2=[['Modul','Smoke Total']];
  allData.forEach(d=>{if(d.wSmokeTotal>0)d2.push([d.module||d.name,d.wSmokeTotal]);});
  if (d2.length>1) {
    d2.forEach((row,ri)=>ws.getRange(cRow+ri,tmp,1,2).setValues([row]));
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(ws.getRange(cRow,tmp,d2.length,2))
        .setPosition(cRow,9,0,0)
        .setOption('title','Smoke TC Distribution (Web)')
        .setOption('pieHole',0.45).setOption('pieSliceText','percentage')
        .setOption('legend',{position:'right'})
        .setOption('width',340).setOption('height',220).build()));
  }

  // Chart 3 — Smoke Status Breakdown (stacked column) - use module
  const d3=[['Modul','Passed','Failed','Blocked','In Prog','Todo']];
  allData.forEach(d=>{if(d.wSmokeTotal>0)d3.push([d.module||d.name,d.wSmokePassed||0,d.wSmokeFailed||0,d.wSmokeBlocked||0,d.wSmokeInProg||0,d.wSmokeTodo||0]);});
  if (d3.length>1) {
    const sRow=cRow+Math.max(allData.length,3)+3;
    d3.forEach((row,ri)=>ws.getRange(sRow+ri,tmp,1,6).setValues([row]));
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(ws.getRange(sRow,tmp,d3.length,6))
        .setPosition(cRow+18,1,0,0)
        .setOption('title','Smoke TC Status Breakdown (Web/Mobile)')
        .setOption('isStacked',true)
        .setOption('series',{0:{color:'#4CAF50'},1:{color:'#F44336'},2:{color:'#FF9800'},3:{color:'#2196F3'},4:{color:'#9E9E9E'}})
        .setOption('legend',{position:'top'})
        .setOption('width',510).setOption('height',260).build()));
  }

  // Chart 4 — Open Blocker Bugs - use col 2
  if (allData.some(d=>(d.bugStats||{}).blocker>0))
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(ws.getRange(dRow,2,n,1))    // Modul
        .addRange(ws.getRange(dRow,11,n,1))   // Web blocker
        .setPosition(cRow+18,9,0,0)
        .setOption('title','🚨 Open Blocker Bugs per Modul')
        .setOption('hAxis',{title:'Jumlah Bug',minValue:0})
        .setOption('colors',['#B71C1C'])
        .setOption('legend',{position:'top'})
        .setOption('width',340).setOption('height',260).build()));

  // Chart 5 — Exec Rate vs target 100% - use module
  const d5=[['Modul','Exec%','Target']];
  allData.forEach(d=>{if(d.wSmokeTotal>0)d5.push([d.module||d.name,d.wSmokeExecRate,1]);});
  if (d5.length>1) {
    const eRow=cRow+Math.max(allData.length,3)*2+6;
    d5.forEach((row,ri)=>ws.getRange(eRow+ri,tmp,1,3).setValues([row]));
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(ws.getRange(eRow,tmp,d5.length,3))
        .setPosition(cRow+36,1,0,0)
        .setOption('title','Smoke Exec Rate vs Target 100%')
        .setOption('vAxis',{format:'#%',minValue:0,maxValue:1,title:'Exec Rate'})
        .setOption('series',{0:{color:'#0288D1',labelInLegend:'Exec Rate'},1:{color:'#E53935',type:'line',labelInLegend:'Target 100%',lineWidth:2}})
        .setOption('legend',{position:'top'})
        .setOption('width',510).setOption('height',240).build()));
  }
}


// ═══════════════════════════════════════════════════════════════════════
// BLOCKERS TAB
// ═══════════════════════════════════════════════════════════════════════

function buildBlockers(ss) {
  const ws=ss.insertSheet('Blockers'); ws.setTabColor('#B71C1C'); ws.clear();
  function h_(c,txt){ws.getRange(2,c).setValue(txt).setBackground('#B71C1C').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#E57373',SpreadsheetApp.BorderStyle.SOLID);}
  [120,80,85,75,75,100,250,85].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  ws.getRange(1,1,1,8).merge().setValue('BLOCKER ALERT  —  Critical & High  |  Status: FAILED / BLOCKED')
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

function writeBlockers(ss, allData) {
  const ws=ss.getSheetByName('Blockers'); if(!ws)return;
  const lastRow=Math.max(ws.getLastRow(),3);
  if(lastRow>=3)ws.getRange(3,1,lastRow-2,8).clearContent().clearFormat();
  const all=[];
  allData.forEach(d=>d.blockers.forEach(b=>all.push({...b,refreshed:d.refreshed})));
  if(all.length===0){
    ws.getRange(3,1,1,8).merge().setValue('✅ Tidak ada blocker! Semua Critical & High TC passed.')
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
  const lastRow=Math.max(ws.getLastRow(),3);
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
  const hdrs=['Timestamp','Project','Modul','Submodul','PIC QA',
    'wPass%','wExec%','aPass%','aExec%',
    'wSmokePass%','wSmokeExec%','aSmokePass%','aSmokeExec%',
    'Perf','Bugs','Open','Blocker','Critical'];
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
  const ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');
  allData.forEach(d=>{
    const bs=d.bugStats||{};
    ws.appendRow([ts,d.project||'',d.module||'',d.submodule||d.name,d.team||'',
      d.wPassRate,d.wExecRate,d.aPassRate,d.aExecRate,
      d.wSmokePassRate,d.wSmokeExecRate,d.aSmokePassRate,d.aSmokeExecRate,
      d.perfResult,bs.total||0,bs.open||0,bs.blocker||0,bs.critical||0]);
  });
  const lastRow=ws.getLastRow();
  if(lastRow>=3){
    for(const col of [7,8,9,10,11,12,13,14])ws.getRange(3,col,lastRow-2,1).setNumberFormat('0%');
    tryChart_(()=>{
      ws.getCharts().forEach(c=>ws.removeChart(c));
      ws.insertChart(ws.newChart()
          .setChartType(Charts.ChartType.LINE)
          .addRange(ws.getRange(2,1,lastRow-1,1))
          .addRange(ws.getRange(2,7,lastRow-1,1))    // wPass%
          .addRange(ws.getRange(2,9,lastRow-1,1))    // aPass%
          .addRange(ws.getRange(2,11,lastRow-1,1))   // wSmokePass%
          .addRange(ws.getRange(2,13,lastRow-1,1))   // aSmokePass%
          .setPosition(3,21,0,0)
          .setOption('title','Pass Rate Trend Over Time')
          .setOption('curveType','function')
          .setOption('series',{
            0:{color:'#1565C0',labelInLegend:'Web Pass%'},
            1:{color:'#283593',labelInLegend:'API Pass%'},
            2:{color:'#BF360C',labelInLegend:'Smoke Web%',lineDashStyle:[6,3]},
            3:{color:'#4A148C',labelInLegend:'Smoke API%',lineDashStyle:[6,3]}
          })
          .setOption('vAxis',{title:'Pass Rate',format:'#%',minValue:0,maxValue:1})
          .setOption('hAxis',{title:'Refresh Time'})
          .setOption('legend',{position:'top'})
          .setOption('width',720).setOption('height',380).build());
    });
  }
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
        // Update data dari QATM Summary
        if(d.project)   cfg.getRange(i+1,3).setValue(d.project);    // col C = Project
        if(d.module)    cfg.getRange(i+1,4).setValue(d.module);     // col D = Modul
        if(d.submodule) cfg.getRange(i+1,5).setValue(d.submodule);  // col E = Submodul
        if(d.team)      cfg.getRange(i+1,6).setValue(d.team);       // col F = PIC QA
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

function tryChart_(fn) {
  try { fn(); } catch(e) { Logger.log('Chart skip: ' + e.message); }
}

function getActiveModules_() {
  return getModuleList_(SpreadsheetApp.getActiveSpreadsheet());
}
