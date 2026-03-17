/**
 * notifications.js
 * Send blocker & PROD BUGS notifications via Google Chat & Email
 */

/**
 * Send blocker notification to Google Chat and Email
 * Run manually for testing or via daily trigger
 *
 * NEW v2: Aggregates modules by webhook - 1 message per webhook with multiple modules
 */
function sendBlockerNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');

  if (!cfg) {
    Logger.log('Config tab not found');
    return;
  }

  // Get blocker data from Overview tab (includes per-module config)
  const overview = ss.getSheetByName('Overview');
  if (!overview) {
    Logger.log('Overview tab not found');
    return;
  }

  const blockerData = getBlockerData_(overview, cfg);

  if (blockerData.totalBlockers === 0 && blockerData.totalProdBugs === 0) {
    Logger.log('No blockers or PROD bugs found - notification skipped');
    SpreadsheetApp.getUi().alert(
      '✅ All Clear!',
      'Tidak ada Open Blocker atau PROD BUGS.\n\n' +
      'Notification tidak dikirim (tidak ada yang perlu di-alert).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // ── DEBUG: Log blocker data ──────────────────────────────────────────────
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('DEBUG: Blocker Data Summary');
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('Total modules with bugs: ' + blockerData.modules.length);
  blockerData.modules.forEach((mod, i) => {
    Logger.log('Module ' + (i+1) + ':');
    Logger.log('  - Project: ' + mod.project);
    Logger.log('  - Module: ' + mod.module);
    Logger.log('  - Blockers: ' + mod.blocker);
    Logger.log('  - PROD Bugs: ' + mod.prodBugs);
    Logger.log('  - Chat Enabled: ' + mod.chatEnabled);
    Logger.log('  - Chat Webhook: ' + (mod.chatWebhook ? mod.chatWebhook.substring(0, 50) + '...' : 'NONE'));
    Logger.log('  - Email Enabled: ' + mod.emailEnabled);
  });

  // ── GET GLOBAL WHATSAPP CONFIG (before grouping) ────────────────────────
  // Get GLOBAL WhatsApp Config from Config row 4, col S-U (19-21)
  let whatsappGroupId = null;
  let fontteToken = null;
  let whatsappEnabled = false;

  if (cfg) {
    whatsappGroupId = String(cfg.getRange(4, 19).getValue()).trim(); // S4 = WhatsApp Group ID
    fontteToken = String(cfg.getRange(4, 20).getValue()).trim();      // T4 = Fonnte Token
    whatsappEnabled = cfg.getRange(4, 21).getValue() === true;        // U4 = Enable WA (checkbox)
  }

  // ── GROUP BY WEBHOOK ────────────────────────────────────────────────────
  // Aggregate modules with same webhook → send 1 message per webhook
  // ────────────────────────────────────────────────────────────────────────

  const webhookGroups = {};   // { webhookUrl: [module1, module2, ...] }
  const emailGroups = {};     // { emailRecipients: [module1, module2, ...] }

  blockerData.modules.forEach(module => {
    // Group by Google Chat webhook
    if (module.chatEnabled && module.chatWebhook) {
      if (!webhookGroups[module.chatWebhook]) {
        webhookGroups[module.chatWebhook] = [];
      }
      webhookGroups[module.chatWebhook].push(module);
    }

    // Group by Email recipients
    if (module.emailEnabled && module.emailRecipients) {
      if (!emailGroups[module.emailRecipients]) {
        emailGroups[module.emailRecipients] = [];
      }
      emailGroups[module.emailRecipients].push(module);
    }
  });

  // ── DEBUG: Log webhook grouping ──────────────────────────────────────────
  Logger.log('');
  Logger.log('Webhook Groups: ' + Object.keys(webhookGroups).length);
  Object.keys(webhookGroups).forEach((webhook, i) => {
    Logger.log('Webhook ' + (i+1) + ': ' + webhook.substring(0, 50) + '...');
    Logger.log('  - Modules: ' + webhookGroups[webhook].length);
    webhookGroups[webhook].forEach(m => {
      Logger.log('    • ' + m.project + ' - ' + m.module);
    });
  });
  Logger.log('WhatsApp Config (Global):');
  Logger.log('  - Group ID: ' + (whatsappGroupId || 'NOT SET'));
  Logger.log('  - Token: ' + (fontteToken ? fontteToken.substring(0, 10) + '...' : 'NOT SET'));
  Logger.log('  - Enabled: ' + whatsappEnabled);
  Logger.log('═══════════════════════════════════════════════');

  // ── SEND AGGREGATED NOTIFICATIONS ───────────────────────────────────────
  let chatSentCount = 0;
  let emailSentCount = 0;
  let whatsappSentCount = 0;
  let skippedCount = blockerData.modules.filter(m => !m.chatEnabled && !m.emailEnabled).length;

  // Send 1 Google Chat message per webhook (aggregated modules)
  Object.keys(webhookGroups).forEach(webhookUrl => {
    const modules = webhookGroups[webhookUrl];
    Logger.log('Sending Google Chat to webhook: ' + webhookUrl + ' (' + modules.length + ' modules)');

    const aggregatedData = {
      modules: modules,
      totalBlockers: modules.reduce((sum, m) => sum + m.blocker, 0),
      totalProdBugs: modules.reduce((sum, m) => sum + m.prodBugs, 0),
      timestamp: blockerData.timestamp
    };

    if (sendGoogleChatNotification_(webhookUrl, aggregatedData)) {
      chatSentCount++;
    }
  });

  // Send 1 Email per recipient group (aggregated modules)
  Object.keys(emailGroups).forEach(recipients => {
    const modules = emailGroups[recipients];
    Logger.log('Sending Email to: ' + recipients + ' (' + modules.length + ' modules)');

    const aggregatedData = {
      modules: modules,
      totalBlockers: modules.reduce((sum, m) => sum + m.blocker, 0),
      totalProdBugs: modules.reduce((sum, m) => sum + m.prodBugs, 0),
      timestamp: blockerData.timestamp
    };

    if (sendEmailNotification_(recipients, aggregatedData)) {
      emailSentCount++;
    }
  });

  // Send 1 WhatsApp message (GLOBAL config - all modules to 1 group)
  if (whatsappEnabled && whatsappGroupId && fontteToken && whatsappGroupId.includes('@g.us')) {
    Logger.log('Sending WhatsApp to global group: ' + whatsappGroupId + ' (' + blockerData.modules.length + ' modules)');

    if (sendWhatsAppNotification_(whatsappGroupId, blockerData, fontteToken)) {
      whatsappSentCount = 1;
    }
  } else if (whatsappEnabled) {
    Logger.log('⚠️ WhatsApp enabled but config incomplete - Group ID: ' + (whatsappGroupId || 'MISSING') + ', Token: ' + (fontteToken ? 'SET' : 'MISSING'));
  }

  // Show result
  let msg = '📤 Notifications Sent!\n\n';
  if (chatSentCount > 0) msg += '✅ Google Chat: ' + chatSentCount + ' message(s) sent\n';
  if (emailSentCount > 0) msg += '✅ Email: ' + emailSentCount + ' message(s) sent\n';
  if (whatsappSentCount > 0) msg += '✅ WhatsApp: ' + whatsappSentCount + ' message(s) sent\n';
  if (skippedCount > 0) msg += 'ℹ️ Skipped: ' + skippedCount + ' module(s) (notifications disabled)\n';
  msg += '\n📊 Summary:\n';
  msg += '• Total Open Blockers: ' + blockerData.totalBlockers + '\n';
  msg += '• Total PROD BUGS: ' + blockerData.totalProdBugs + '\n';
  msg += '• Modules with issues: ' + blockerData.modules.length + '\n';
  msg += '• Unique webhooks: ' + Object.keys(webhookGroups).length + '\n';
  msg += '• Unique email groups: ' + Object.keys(emailGroups).length + '\n';
  if (whatsappSentCount > 0) msg += '• WhatsApp: Sent to global group\n';
  msg += '• Modules skipped: ' + skippedCount;

  SpreadsheetApp.getUi().alert('Notification Sent', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log('✅ Notifications sent successfully - Chat: ' + chatSentCount + ', Email: ' + emailSentCount + ', WhatsApp: ' + whatsappSentCount + ', Skipped: ' + skippedCount);
}

/**
 * Setup daily blocker notification trigger
 *
 * SUPPORTS FLEXIBLE SCHEDULING:
 * - Single: "7" or 7 → 1x per day at 7:00
 * - Multiple: "7,12,18" → 3x per day at 7:00, 12:00, 18:00
 * - Interval: "4h" → Every 4 hours (supports: 1h, 2h, 4h, 6h, 8h, 12h)
 */
function setupDailyBlockerNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');
  const ui = SpreadsheetApp.getUi();

  if (!cfg) {
    ui.alert('❌ Config tab not found!');
    return;
  }

  // ═══════════════════════════════════════════════════════════════════════
  // SHOW SETUP INSTRUCTIONS FIRST
  // ═══════════════════════════════════════════════════════════════════════
  const instructionsMsg =
    '🔔 SETUP NOTIFICATIONS\n\n' +
    'Buka Config tab dan isi konfigurasi berikut:\n\n' +
    '📱 GOOGLE CHAT (kolom L-N) - Per Module\n' +
    '  L = Webhook URL (dari Google Chat Space)\n' +
    '  M = Schedule (9,14,18 atau 4h)\n' +
    '  N = ☑ Enable Chat\n\n' +
    '📧 EMAIL (kolom O-P) - Per Module\n' +
    '  O = Email recipients (comma separated)\n' +
    '  P = ☑ Enable Email\n\n' +
    '📲 WHATSAPP (kolom S-U row 4) - GLOBAL\n' +
    '  S = Group ID (120363xxx@g.us)\n' +
    '  T = Fonnte Token\n' +
    '  U = ☑ Enable WhatsApp\n\n' +
    '⏰ SCHEDULE FORMAT:\n' +
    '  • Single: 9\n' +
    '  • Multiple: 9,14,18\n' +
    '  • Interval: 4h (1h/2h/4h/6h/8h/12h)\n\n' +
    '────────────────────────────────\n' +
    'Sudah setup config?\n' +
    'YES = Lanjut create trigger\n' +
    'NO = Buka Config tab dulu';

  const response = ui.alert('Setup Notifications', instructionsMsg, ui.ButtonSet.YES_NO);

  if (response === ui.Button.NO) {
    // Open Config tab for user to setup
    ss.setActiveSheet(cfg);
    ss.setActiveRange(cfg.getRange('L4'));
    ui.alert(
      'ℹ️ Config Tab Opened',
      'Silakan isi konfigurasi notification di Config tab.\n\n' +
      'Setelah selesai, jalankan menu ini lagi untuk create trigger.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Check if ANY module has notifications enabled (Chat/Email per-module, WhatsApp global)
  const cfgData = cfg.getDataRange().getValues();
  let hasAnyEnabled = false;
  let scheduleStr = '7'; // Default schedule

  // Check per-module Chat/Email notifications
  for (let i = 3; i < cfgData.length; i++) {
    const chatEnabled = cfgData[i][13] === true; // Col N
    const emailEnabled = cfgData[i][15] === true; // Col P
    const schedule = String(cfgData[i][12]).trim(); // Col M - now text format

    if (chatEnabled || emailEnabled) {
      hasAnyEnabled = true;
      scheduleStr = schedule || '7'; // Use schedule from first enabled module
      break;
    }
  }

  // Check GLOBAL WhatsApp notification (row 4, col U)
  const whatsappEnabled = cfgData[3] && cfgData[3][20] === true; // U4 (index 20)

  if (whatsappEnabled) {
    hasAnyEnabled = true;
  }

  if (!hasAnyEnabled) {
    const openConfig = ui.alert(
      '⚠️ No Notifications Enabled',
      'Tidak ada notification yang aktif.\n\n' +
      'Aktifkan minimal 1 channel di Config tab:\n' +
      '• Kolom N = ☑ Enable Google Chat (per-module)\n' +
      '• Kolom P = ☑ Enable Email (per-module)\n' +
      '• Kolom U row 4 = ☑ Enable WhatsApp (global)\n\n' +
      'Buka Config tab sekarang?',
      ui.ButtonSet.YES_NO
    );

    if (openConfig === ui.Button.YES) {
      ss.setActiveSheet(cfg);
      ss.setActiveRange(cfg.getRange('N4'));
    }
    return;
  }

  // ═══════════════════════════════════════════════════════════════════════
  // PARSE SCHEDULE STRING
  // ═══════════════════════════════════════════════════════════════════════

  const parsedSchedule = parseSchedule_(scheduleStr);

  if (!parsedSchedule.success) {
    SpreadsheetApp.getUi().alert(
      '❌ Invalid Schedule Format',
      'Schedule format tidak valid: "' + scheduleStr + '"\n\n' +
      '✅ FORMAT VALID:\n' +
      '• Single: 7 atau "7"\n' +
      '• Multiple: 7,12,18\n' +
      '• Interval: 4h (support: 1h, 2h, 4h, 6h, 8h, 12h)\n\n' +
      'Error: ' + parsedSchedule.error,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // ═══════════════════════════════════════════════════════════════════════
  // DELETE ALL EXISTING NOTIFICATION TRIGGERS
  // ═══════════════════════════════════════════════════════════════════════

  let removedCount = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendBlockerNotification') {
      ScriptApp.deleteTrigger(t);
      removedCount++;
    }
  });

  Logger.log('Removed ' + removedCount + ' existing notification trigger(s)');

  // ═══════════════════════════════════════════════════════════════════════
  // CREATE NEW TRIGGERS BASED ON SCHEDULE TYPE
  // ═══════════════════════════════════════════════════════════════════════

  let createdTriggers = 0;
  let scheduleDescription = '';

  if (parsedSchedule.type === 'interval') {
    // Interval-based trigger (e.g., every 4 hours)
    ScriptApp.newTrigger('sendBlockerNotification')
      .timeBased()
      .everyHours(parsedSchedule.intervalHours)
      .create();

    createdTriggers = 1;
    scheduleDescription = 'Setiap ' + parsedSchedule.intervalHours + ' jam';

    Logger.log('✅ Created interval trigger: every ' + parsedSchedule.intervalHours + ' hours');

  } else if (parsedSchedule.type === 'specific') {
    // Specific hour(s) trigger (e.g., 7, 12, 18)
    parsedSchedule.hours.forEach(hour => {
      ScriptApp.newTrigger('sendBlockerNotification')
        .timeBased()
        .atHour(hour)
        .everyDays(1)
        .create();

      createdTriggers++;
      Logger.log('✅ Created daily trigger at hour: ' + hour);
    });

    if (parsedSchedule.hours.length === 1) {
      scheduleDescription = 'Setiap hari jam ' + parsedSchedule.hours[0] + ':00';
    } else {
      scheduleDescription = parsedSchedule.hours.length + 'x per hari: jam ' +
                           parsedSchedule.hours.map(h => h + ':00').join(', ');
    }
  }

  // ═══════════════════════════════════════════════════════════════════════
  // SUCCESS CONFIRMATION
  // ═══════════════════════════════════════════════════════════════════════

  ui.alert(
    '✅ Notifications Setup Complete!',
    '📅 SCHEDULE: ' + scheduleDescription + '\n' +
    '🔧 Triggers created: ' + createdTriggers + '\n\n' +
    '📤 Active Channels:\n' +
    '• Google Chat (per-module)\n' +
    '• Email (per-module)\n' +
    '• WhatsApp (global)\n\n' +
    '💡 Test now: Menu > Notifications > Test Notification Now\n' +
    '🔍 View triggers: Apps Script Editor > Triggers tab\n\n' +
    'Notifikasi akan otomatis terkirim sesuai schedule yang sudah di-set.',
    ui.ButtonSet.OK
  );

  Logger.log('✅ Successfully created ' + createdTriggers + ' notification trigger(s): ' + scheduleDescription);
}

/**
 * Parse schedule string into structured format
 *
 * Supports:
 * - Single: "7" or 7 → {type: 'specific', hours: [7]}
 * - Multiple: "7,12,18" → {type: 'specific', hours: [7, 12, 18]}
 * - Interval: "4h" → {type: 'interval', intervalHours: 4}
 *
 * @param {string|number} scheduleStr - Schedule string from Config
 * @return {Object} - {success: bool, type: 'specific'|'interval', hours: [], intervalHours: num, error: str}
 */
function parseSchedule_(scheduleStr) {
  const str = String(scheduleStr).trim();

  // ── INTERVAL FORMAT: "4h", "6h", etc. ──────────────────────────────────
  if (str.match(/^\d+h$/i)) {
    const hours = parseInt(str.replace(/h$/i, ''));

    // Apps Script only supports: 1, 2, 4, 6, 8, 12 hours
    const validIntervals = [1, 2, 4, 6, 8, 12];

    if (!validIntervals.includes(hours)) {
      return {
        success: false,
        error: 'Interval ' + hours + 'h tidak didukung. Valid: 1h, 2h, 4h, 6h, 8h, 12h'
      };
    }

    return {
      success: true,
      type: 'interval',
      intervalHours: hours
    };
  }

  // ── SPECIFIC HOUR(S) FORMAT: "7", "7,12,18" ────────────────────────────

  // Split by comma and parse each hour
  const parts = str.split(',').map(p => p.trim()).filter(p => p.length > 0);

  if (parts.length === 0) {
    return {
      success: false,
      error: 'Schedule kosong'
    };
  }

  const hours = [];

  for (let i = 0; i < parts.length; i++) {
    const hourStr = parts[i];

    // Check if it's a valid number
    if (!hourStr.match(/^\d+$/)) {
      return {
        success: false,
        error: 'Invalid hour: "' + hourStr + '". Harus angka 0-23.'
      };
    }

    const hour = parseInt(hourStr);

    // Validate hour range (0-23)
    if (hour < 0 || hour > 23) {
      return {
        success: false,
        error: 'Hour ' + hour + ' diluar range. Valid: 0-23.'
      };
    }

    // Check duplicate
    if (hours.includes(hour)) {
      return {
        success: false,
        error: 'Duplicate hour: ' + hour
      };
    }

    hours.push(hour);
  }

  // Sort hours ascending
  hours.sort((a, b) => a - b);

  return {
    success: true,
    type: 'specific',
    hours: hours
  };
}

/**
 * Remove daily blocker notification trigger
 */
function removeDailyBlockerNotification() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendBlockerNotification') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  if (removed > 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Trigger Removed',
      'Daily blocker notification trigger telah dihapus.\n\n' +
      'Triggers removed: ' + removed,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log('Removed ' + removed + ' blocker notification trigger(s)');
  } else {
    SpreadsheetApp.getUi().alert(
      'ℹ️ No Trigger Found',
      'Tidak ada daily blocker notification trigger yang aktif.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Read notification config from Config tab
 * Google Chat: columns L-N (12-14), row 4
 * Email: columns O-P (15-16), row 4
 */
function readNotificationConfig_(cfg) {
  const config = {
    chatEnabled: false,
    chatWebhook: null,
    emailEnabled: false,
    emailRecipients: null,
    notifHour: 7
  };

  try {
    // Google Chat config: row 4, columns L-N (12-14)
    // L4 = Webhook URL, M4 = Hour, N4 = Enable (checkbox boolean)
    const chatWebhook = String(cfg.getRange(4, 12).getValue()).trim(); // L4
    const notifHour = parseInt(cfg.getRange(4, 13).getValue()) || 7;   // M4
    const chatEnabled = cfg.getRange(4, 14).getValue();                 // N4 (checkbox boolean)

    if (chatWebhook && chatWebhook.includes('chat.googleapis.com')) {
      config.chatWebhook = chatWebhook;
      config.notifHour = notifHour;
      config.chatEnabled = (chatEnabled === true); // Checkbox boolean
    }

    // Email config: row 4, columns O-P (15-16)
    // O4 = Email recipients, P4 = Enable (checkbox boolean)
    const emailRecipients = String(cfg.getRange(4, 15).getValue()).trim(); // O4
    const emailEnabled = cfg.getRange(4, 16).getValue();                    // P4 (checkbox boolean)

    if (emailRecipients && emailRecipients.includes('@')) {
      config.emailRecipients = emailRecipients;
      config.emailEnabled = (emailEnabled === true); // Checkbox boolean
    }

  } catch (e) {
    Logger.log('Error reading notification config: ' + e.message);
  }

  Logger.log('Notification config read: chatEnabled=' + config.chatEnabled + ', emailEnabled=' + config.emailEnabled);
  return config;
}

/**
 * Get blocker data from Bugs tab (or Overview as fallback) and Config for QATM URLs + per-module notification config
 *
 * NEW: Reads from Bugs tab for accurate Modul/Submodul names (e.g., "1 - Portal + SSO" instead of "SIPGN - 1.1,1.2,1.3")
 * NEW: Includes per-module webhook URLs and enable flags from Config
 */
function getBlockerData_(overview, cfg) {
  // Read from Bugs tab instead of Overview for accurate module names
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bugsTab = ss.getSheetByName('Bugs');

  let data;
  if (!bugsTab) {
    Logger.log('⚠️⚠️⚠️ Bugs tab not found - using Overview as fallback ⚠️⚠️⚠️');
    data = overview.getDataRange().getValues();
  } else {
    Logger.log('✅✅✅ Reading from Bugs tab for accurate module names ✅✅✅');
    data = bugsTab.getDataRange().getValues();
    Logger.log('Bugs tab rows: ' + data.length);
  }

  // Build module map from Config for QATM URLs + notification config
  const moduleMap = {};
  if (cfg) {
    const cfgData = cfg.getDataRange().getValues();
    for (let i = 3; i < cfgData.length; i++) {
      const active = cfgData[i][0] === true;           // Col A = Active checkbox
      const project = String(cfgData[i][2]).trim();    // Col C
      const modul = String(cfgData[i][3]).trim();      // Col D
      const qatmId = String(cfgData[i][6]).trim();     // Col G
      const jiraInstance = String(cfgData[i][8]).trim(); // Col I (8) = Jira Instance

      // Per-module notification config
      // Col L (11) = Google Chat Webhook URL
      // Col N (13) = Enable Notifikasi (checkbox)
      // Col O (14) = Email Recipients
      // Col P (15) = Enable Email (checkbox)
      // WhatsApp is GLOBAL config (row 4, col S-U), not per-module
      const chatWebhook = String(cfgData[i][11]).trim();
      const chatEnabled = cfgData[i][13] === true;
      const emailRecipients = String(cfgData[i][14]).trim();
      const emailEnabled = cfgData[i][15] === true;

      // Only add active modules with valid QATM ID
      if (active && qatmId && qatmId.length > 10) {
        const key = project + '|' + modul;
        moduleMap[key] = {
          qatmUrl: 'https://docs.google.com/spreadsheets/d/' + qatmId + '/edit',
          bugReportGid: '2', // BugReport is typically GID 2
          jiraInstance: jiraInstance || null,
          chatWebhook: chatWebhook && chatWebhook.includes('chat.googleapis.com') ? chatWebhook : null,
          chatEnabled: chatEnabled,
          emailRecipients: emailRecipients && emailRecipients.includes('@') ? emailRecipients : null,
          emailEnabled: emailEnabled
        };
        Logger.log('Config loaded: ' + key + ' (chatEnabled=' + chatEnabled + ', emailEnabled=' + emailEnabled + ')');
      }
    }
  }

  Logger.log('Total configs in moduleMap: ' + Object.keys(moduleMap).length);

  const modules = [];
  let totalBlockers = 0;
  let totalProdBugs = 0;

  // Determine if reading from Bugs tab or Overview
  const isFromBugs = bugsTab ? true : false;

  // Start from row 5 for both Bugs tab and Overview (index 4)
  // Bugs tab structure: Row 1-4 are headers, data starts at row 5
  const startRow = 4;

  for (let i = startRow; i < data.length; i++) {
    const row = data[i];

    // Stop at TOTAL row or empty rows
    if (String(row[0] || '').includes('TOTAL') || String(row[0] || '').includes('AVERAGE')) break;

    let project, modul, moduleName, blocker, prodBugs, critical, high, medium;

    if (isFromBugs) {
      // Reading from Bugs tab structure
      // Col B (index 1) = Modul number, Col C (index 2) = Submodul name
      if (!row[1] && !row[2]) continue; // Empty row

      modul = String(row[1]).trim();          // Col B = Modul (e.g., "1", "3", "4")
      moduleName = String(row[2]).trim();     // Col C = Submodul (e.g., "Portal + SSO")

      // Skip rows with NO submodule name
      if (!moduleName || moduleName === '') {
        Logger.log('⚠️ Skipping row with empty moduleName');
        continue;
      }

      project = modul || moduleName;          // Use modul number as project, or moduleName if modul is empty
      critical = parseInt(row[4]) || 0;       // Col E = Critical
      high = parseInt(row[5]) || 0;           // Col F = High
      medium = parseInt(row[6]) || 0;         // Col G = Medium
      blocker = parseInt(row[9]) || 0;        // Col J = Blocker (was row[7] which is Low column)
      prodBugs = parseInt(row[11]) || 0;      // Col L = PROD BUGS
    } else {
      // Reading from Overview tab structure (fallback)
      if (!row[1] && !row[2]) continue; // Empty row

      project = row[0] || '';
      modul = row[1] || '';                   // Col B = Modul
      moduleName = row[2] || modul || 'Unknown'; // Col C = Submodule
      critical = parseInt(row[3]) || 0;       // Col D = Critical (estimate from Overview)
      high = parseInt(row[4]) || 0;           // Col E = High (estimate from Overview)
      medium = 0;                             // Not available in Overview
      blocker = parseInt(row[5]) || 0;        // Col F = Blocker
      prodBugs = parseInt(row[7]) || 0;       // Col H = PROD BUGS
    }

    if (blocker > 0 || prodBugs > 0) {
      // Find matching config by modul number OR submodule name (for rows with empty modul)
      let moduleInfo = null;
      let configProjectName = '';
      let matchedConfigKey = '';

      if (modul && modul !== '') {
        // Try matching by modul number first
        Object.keys(moduleMap).forEach(key => {
          const parts = key.split('|');
          const configProject = parts[0]; // Get project from "PROJECT|MODUL"
          const configModul = parts[1];    // Get modul from "PROJECT|MODUL"
          if (configModul === modul) {
            moduleInfo = moduleMap[key];
            configProjectName = configProject;
            matchedConfigKey = key;
          }
        });
      }

      // If no match and modul is empty, try finding by QATM that contains this submodule
      // (This is a fallback - ideally Dashboard should have modul filled)
      if (!moduleInfo) {
        Logger.log('⚠️ No direct config match for modul="' + modul + '", moduleName="' + moduleName + '" - will try to find QATM by reading all configs');
        // For now, leave moduleInfo empty - user should fill modul column in Dashboard
        moduleInfo = {};
      }

      Logger.log('Found module with bugs: configKey=' + matchedConfigKey + ', projectName=' + configProjectName + ', moduleName=' + moduleName + ', blocker=' + blocker + ', prodBugs=' + prodBugs + ', configFound=' + (moduleInfo.qatmUrl ? 'YES' : 'NO'));

      modules.push({
        project: configProjectName || project,  // Use project name from Config (e.g., "SIPGN")
        module: modul,                           // Module number (e.g., "4")
        submodule: moduleName,                   // Submodule name (e.g., "AI Surveillance")
        blocker: blocker,
        prodBugs: prodBugs,
        critical: critical,                      // Critical bugs count
        high: high,                              // High bugs count
        medium: medium,                          // Medium bugs count
        qatmUrl: moduleInfo.qatmUrl || null,
        bugReportGid: moduleInfo.bugReportGid || '2',
        jiraInstance: moduleInfo.jiraInstance || null,
        // Per-module notification config
        chatWebhook: moduleInfo.chatWebhook || null,
        chatEnabled: moduleInfo.chatEnabled || false,
        emailRecipients: moduleInfo.emailRecipients || null,
        emailEnabled: moduleInfo.emailEnabled || false
        // WhatsApp is GLOBAL config (read from Config row 4, col S-U)
      });
      totalBlockers += blocker;
      totalProdBugs += prodBugs;
    }
  }

  return {
    modules: modules,
    totalBlockers: totalBlockers,
    totalProdBugs: totalProdBugs,
    timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };
}

/**
 * Get detailed bug data from QATM BugReport sheets
 * Returns bugs grouped by module → environment → submodule
 *
 * NEW: Deduplicate QATM URLs to avoid reading same bugs multiple times
 */
function getBugDetailsFromQATM_(modules) {
  const bugDetails = [];

  // Deduplicate by QATM URL - only read each QATM once
  const processedQATMs = {};

  modules.forEach(moduleInfo => {
    if (!moduleInfo.qatmUrl) {
      Logger.log('Skipping ' + moduleInfo.project + '-' + moduleInfo.module + ' - no QATM URL');
      return;
    }

    // Check if this QATM was already processed
    if (processedQATMs[moduleInfo.qatmUrl]) {
      Logger.log('⚠️ Skipping duplicate QATM: ' + moduleInfo.project + '-' + moduleInfo.module + ' (already read from another module)');
      return;
    }

    try {
      // Extract spreadsheet ID from URL
      const match = moduleInfo.qatmUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) {
        Logger.log('Invalid QATM URL for ' + moduleInfo.project + '-' + moduleInfo.module);
        return;
      }

      const qatmId = match[1];

      // Mark this QATM as processed
      processedQATMs[moduleInfo.qatmUrl] = true;
      Logger.log('✅ Reading QATM: ' + moduleInfo.project + '-' + moduleInfo.module);

      const qatmSs = SpreadsheetApp.openById(qatmId);
      const bugSheet = qatmSs.getSheetByName('BugReport');

      if (!bugSheet) {
        Logger.log('BugReport sheet not found in ' + moduleInfo.project + '-' + moduleInfo.module);
        return;
      }

      // Get actual BugReport GID
      const bugReportGid = bugSheet.getSheetId();

      // Read bug data (skip header rows 1-4)
      const data = bugSheet.getDataRange().getValues();
      const bugs = [];

      for (let i = 4; i < data.length; i++) {
        const row = data[i];

        // Col A = Jira Key, B = Title, C = Priority, D = Status, E = Assignee
        // Col F = Submodul, G = Category, H = Severity, I = Environment
        const jiraKey = String(row[0]).trim();
        const title = String(row[1]).trim();
        const priority = String(row[2]).trim();
        const status = String(row[3]).trim();
        const submodul = String(row[5]).trim();
        const environment = String(row[8]).trim();

        // Skip empty rows or closed bugs
        if (!jiraKey || !title) continue;
        if (status === 'Closed' || status === "Won't Fix") continue;

        // Only include blocker bugs (Medium-Critical)
        if (!['Critical', 'Highest', 'High', 'Medium'].includes(priority)) continue;

        bugs.push({
          jiraKey: jiraKey,
          title: title,
          priority: priority,
          status: status,
          submodul: submodul || 'Uncategorized',
          environment: environment || 'Unknown',
          jiraUrl: getJiraUrl_(jiraKey, moduleInfo.jiraInstance)
        });
      }

      if (bugs.length > 0) {
        bugDetails.push({
          project: moduleInfo.project,
          module: moduleInfo.module,
          qatmUrl: moduleInfo.qatmUrl,
          bugReportGid: bugReportGid,  // Use actual GID from sheet
          bugs: bugs
        });
      }

    } catch (e) {
      Logger.log('Error reading bugs from ' + moduleInfo.project + '-' + moduleInfo.module + ': ' + e.message);
    }
  });

  return bugDetails;
}

/**
 * Get Jira URL from Jira key (assumes instance from Config)
 */
function getJiraUrl_(jiraKey, jiraInstance) {
  // Try to extract Jira instance from key prefix
  if (jiraKey.startsWith('BGN-')) {
    return 'https://bgn-peruri.atlassian.net/browse/' + jiraKey;
  } else if (jiraKey.startsWith('DPE-') || jiraKey.startsWith('DP-')) {
    return 'https://digitalperuri.atlassian.net/browse/' + jiraKey;
  }

  // Fallback - use jiraInstance if provided
  if (jiraInstance) {
    return 'https://' + jiraInstance + '.atlassian.net/browse/' + jiraKey;
  }

  // Default fallback
  return 'https://atlassian.net/browse/' + jiraKey;
}

/**
 * Send Google Chat notification - Plain text format (copy-paste friendly)
 * NEW v4: Single message, all modules, WhatsApp-friendly
 */
function sendGoogleChatNotification_(webhookUrl, blockerData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardUrl = ss.getUrl();
    const dashboardBugsUrl = dashboardUrl + '#gid=' + getDashboardBugsGid_(ss);
    const dashboardOverviewUrl = dashboardUrl + '#gid=' + getDashboardOverviewGid_(ss);

    // Get detailed bug data from QATM BugReport sheets
    const bugDetails = getBugDetailsFromQATM_(blockerData.modules);

    // Group bugs by PROJECT + MODULE (not by submodule)
    // This ensures all submodules under same module are grouped together
    // Example: All "SIPGN - 4" submodules (AI Surveillance, Menu Planner, PoP) in one group
    const moduleGroups = {};

    bugDetails.forEach(moduleData => {
      // Use project+module as moduleKey (e.g., "SIPGN|4")
      const moduleKey = moduleData.project + '|' + moduleData.module;

      if (!moduleGroups[moduleKey]) {
        moduleGroups[moduleKey] = {
          project: moduleData.project,  // Project name (e.g., "SIPGN")
          module: moduleData.module,     // Module number (e.g., "4")
          qatmUrl: moduleData.qatmUrl,
          bugReportGid: moduleData.bugReportGid,
          prodSubmodules: {},
          devSubmodules: {}
        };
      }

      // Group bugs by submodule and environment
      moduleData.bugs.forEach(bug => {
        const isProd = (bug.environment === 'Production' || bug.environment === 'Prod');
        const submodules = isProd ? moduleGroups[moduleKey].prodSubmodules : moduleGroups[moduleKey].devSubmodules;

        if (!submodules[bug.submodul]) {
          submodules[bug.submodul] = [];
        }
        submodules[bug.submodul].push(bug);
      });
    });

    // Build plain text message
    let message = '';

    // ══════════════════════════════════════════════════════════════════════
    // HEADER - Daily Bug Report Summary (WhatsApp-ready)
    // ══════════════════════════════════════════════════════════════════════
    message += '📊 *DAILY BUG REPORT*\n';
    message += '📅 ' + blockerData.timestamp + '\n';
    message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

    // Calculate totals from blockerData (same data source as Email/WhatsApp for consistency)
    const totalCritical = blockerData.modules.reduce((sum, m) => sum + (m.critical || 0), 0);
    const totalHigh = blockerData.modules.reduce((sum, m) => sum + (m.high || 0), 0);
    const totalMedium = blockerData.modules.reduce((sum, m) => sum + (m.medium || 0), 0);

    // Header summary
    message += '📈 *SUMMARY*\n';
    message += '🐛 Total Bugs: ' + blockerData.totalBlockers + '  |  🚨 PROD: ' + blockerData.totalProdBugs + '\n';
    message += 'Critical🔴 ' + totalCritical + '  High🟠 ' + totalHigh + '  Medium🟡 ' + totalMedium + '\n\n';

    // Per-submodule breakdown (using blockerData for consistency)
    message += '📋 *BY SUBMODULE:*\n';
    blockerData.modules.forEach(module => {
      const submoduleName = module.submodule || 'Unknown';
      const total = module.blocker || 0;
      const critical = module.critical || 0;
      const high = module.high || 0;
      const medium = module.medium || 0;
      const prod = module.prodBugs || 0;

      if (total > 0 || prod > 0) {
        message += '*' + submoduleName + ':* ' + total + ' bugs';

        if (prod > 0) {
          message += ' (🚨 ' + prod + ' PROD)';
        }

        message += '\n   Critical🔴 ' + critical + '  High🟠 ' + high + '  Medium🟡 ' + medium + '\n';
      }
    });

    message += '\n🔗 <' + dashboardBugsUrl + '|📊 View Dashboard>\n';
    message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

    // ══════════════════════════════════════════════════════════════════════
    // PRODUCTION BUGS SECTION (ALL MODULES) - PRIORITAS PERTAMA
    // ══════════════════════════════════════════════════════════════════════
    let hasProdBugs = false;
    Object.keys(moduleGroups).forEach(key => {
      if (Object.keys(moduleGroups[key].prodSubmodules).length > 0) hasProdBugs = true;
    });

    if (hasProdBugs) {
      message += '🚨🚨🚨 *PRODUCTION BUGS* 🚨🚨🚨\n';
      message += '<users/all>\n';
      message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

      Object.keys(moduleGroups).forEach(moduleKey => {
        const moduleData = moduleGroups[moduleKey];
        const prodSubmodules = moduleData.prodSubmodules;

        if (Object.keys(prodSubmodules).length === 0) return;

        message += '*' + moduleData.project + ' - ' + moduleData.module + '*\n';

        Object.keys(prodSubmodules).forEach(submodulName => {
          const bugs = prodSubmodules[submodulName];

          // Group by severity
          const criticalBugs = bugs.filter(b => b.priority === 'Critical' || b.priority === 'Highest');
          const highBugs = bugs.filter(b => b.priority === 'High');
          const mediumBugs = bugs.filter(b => b.priority === 'Medium');

          message += '📌 ' + submodulName + ': ' + bugs.length + ' bug' + (bugs.length > 1 ? 's' : '') + ' (Critical🔴 ' + criticalBugs.length + '  High🟠 ' + highBugs.length + '  Medium🟡 ' + mediumBugs.length + ')\n';
        });

        message += '<' + moduleData.qatmUrl + '#gid=' + moduleData.bugReportGid + '|📋 View in QATM>\n\n';
      });
    }

    // ══════════════════════════════════════════════════════════════════════
    // PER MODULE - ALL BUGS (grouped by submodule with environment label)
    // ══════════════════════════════════════════════════════════════════════
    Object.keys(moduleGroups).forEach(moduleKey => {
      const moduleData = moduleGroups[moduleKey];
      const prodSubmodules = moduleData.prodSubmodules;
      const devSubmodules = moduleData.devSubmodules;

      if (Object.keys(prodSubmodules).length === 0 && Object.keys(devSubmodules).length === 0) return;

      message += '═══════════════════════════════════\n';
      message += '*' + moduleData.project + ' - ' + moduleData.module + '*\n';
      message += '═══════════════════════════════════\n\n';

      // Combine all submodules with environment label
      const allSubmodules = [];

      // Add PROD submodules
      Object.keys(prodSubmodules).forEach(submodulName => {
        allSubmodules.push({
          name: submodulName,
          env: 'Production',
          envIcon: '🚨',
          bugs: prodSubmodules[submodulName]
        });
      });

      // Add DEV submodules
      Object.keys(devSubmodules).forEach(submodulName => {
        allSubmodules.push({
          name: submodulName,
          env: 'Development',
          envIcon: '⚠️',
          bugs: devSubmodules[submodulName]
        });
      });

      // Display all submodules
      allSubmodules.forEach(submodul => {
        const bugs = submodul.bugs;

        // Group by severity
        const criticalBugs = bugs.filter(b => b.priority === 'Critical' || b.priority === 'Highest');
        const highBugs = bugs.filter(b => b.priority === 'High');
        const mediumBugs = bugs.filter(b => b.priority === 'Medium');

        message += submodul.envIcon + ' *' + submodul.name + '* (' + submodul.env + ')\n';
        message += '   ' + bugs.length + ' bug' + (bugs.length > 1 ? 's' : '') + ' • Critical🔴 ' + criticalBugs.length + '  High🟠 ' + highBugs.length + '  Medium🟡 ' + mediumBugs.length + '\n\n';

        // Critical bugs first
        if (criticalBugs.length > 0) {
          message += '   *Critical 🔴:*\n';
          criticalBugs.forEach((bug, i) => {
            message += '     ' + (i + 1) + '. <' + bug.jiraUrl + '|' + bug.jiraKey + '> - ' + bug.title + '\n';
          });
          message += '\n';
        }

        // High bugs
        if (highBugs.length > 0) {
          message += '   *High 🟠:*\n';
          highBugs.forEach((bug, i) => {
            message += '     ' + (i + 1) + '. <' + bug.jiraUrl + '|' + bug.jiraKey + '> - ' + bug.title + '\n';
          });
          message += '\n';
        }

        // Medium bugs
        if (mediumBugs.length > 0) {
          message += '   *Medium 🟡:*\n';
          mediumBugs.forEach((bug, i) => {
            message += '     ' + (i + 1) + '. <' + bug.jiraUrl + '|' + bug.jiraKey + '> - ' + bug.title + '\n';
          });
          message += '\n';
        }
      });

      message += '<' + moduleData.qatmUrl + '#gid=' + moduleData.bugReportGid + '|📋 View in QATM>\n\n';
    });

    // ══════════════════════════════════════════════════════════════════════
    // FOOTER: Dashboard Links (shortened)
    // ══════════════════════════════════════════════════════════════════════
    message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
    message += '🔗 <' + dashboardOverviewUrl + '|📊 Dashboard Overview>  •  <' + dashboardBugsUrl + '|🐛 Bugs Tab>\n';
    message += '_Automated Daily Report_';

    const payload = {
      text: message
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      Logger.log('✅ Google Chat notification sent successfully');
      return true;
    } else {
      Logger.log('❌ Google Chat notification failed: ' + responseCode + ' - ' + response.getContentText());
      return false;
    }
  } catch (e) {
    Logger.log('❌ Error sending Google Chat notification: ' + e.message);
    return false;
  }
}

/**
 * Send Email notification - PER MODULE format (concise)
 */
function sendEmailNotification_(recipients, blockerData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardUrl = ss.getUrl();

    // Build subject based on aggregated data
    let subject = '🚨 QA Blocker Alert: ' + blockerData.totalBlockers + ' blockers';
    if (blockerData.totalProdBugs > 0) {
      subject = '🚨🚨 URGENT PROD BUGS: ' + blockerData.totalProdBugs + ' bugs affecting production';
    }

    let body = '<html><body style="font-family: Arial, sans-serif; max-width: 700px;">';

    // Header
    if (blockerData.totalProdBugs > 0) {
      body += '<div style="background: #D32F2F; color: white; padding: 20px; text-align: center;">';
      body += '<h2 style="margin: 0;">🚨🚨 PRODUCTION BUGS ALERT 🚨🚨</h2>';
      body += '</div>';
    } else {
      body += '<div style="background: #FF9800; color: white; padding: 15px; text-align: center;">';
      body += '<h2 style="margin: 0;">⚠️ QA BLOCKER ALERT</h2>';
      body += '</div>';
    }

    // Calculate total severity counts
    const totalCritical = blockerData.modules.reduce((sum, m) => sum + (m.critical || 0), 0);
    const totalHigh = blockerData.modules.reduce((sum, m) => sum + (m.high || 0), 0);
    const totalMedium = blockerData.modules.reduce((sum, m) => sum + (m.medium || 0), 0);

    // Summary section
    body += '<div style="background: ' + (blockerData.totalProdBugs > 0 ? '#FFCDD2' : '#FFF3E0') + '; border-left: 4px solid ' + (blockerData.totalProdBugs > 0 ? '#D32F2F' : '#FF9800') + '; padding: 20px; margin: 0;">';
    body += '<h3 style="margin: 0 0 10px 0; color: ' + (blockerData.totalProdBugs > 0 ? '#D32F2F' : '#E65100') + ';">📊 Summary</h3>';
    body += '<table style="width: 100%; font-size: 14px;">';
    body += '<tr><td><strong>Total Blockers:</strong></td><td><strong style="font-size: 18px; color: #E65100;">' + blockerData.totalBlockers + '</strong></td></tr>';

    // Severity breakdown
    body += '<tr><td><strong>Severity:</strong></td><td>';
    if (totalCritical > 0) {
      body += '<span style="background: #D32F2F; color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold; margin-right: 4px;">🔴 ' + totalCritical + '</span> ';
    }
    if (totalHigh > 0) {
      body += '<span style="background: #FF9800; color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold; margin-right: 4px;">🟠 ' + totalHigh + '</span> ';
    }
    if (totalMedium > 0) {
      body += '<span style="background: #FFC107; color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold;">🟡 ' + totalMedium + '</span>';
    }
    body += '</td></tr>';

    if (blockerData.totalProdBugs > 0) {
      body += '<tr><td><strong style="color: #D32F2F;">PROD Bugs:</strong></td><td><span style="background: #D32F2F; color: white; padding: 4px 12px; border-radius: 3px; font-weight: bold; font-size: 18px;">' + blockerData.totalProdBugs + '</span></td></tr>';
    }
    body += '<tr><td><strong>Modules Affected:</strong></td><td>' + blockerData.modules.length + '</td></tr>';
    body += '<tr><td><strong>Timestamp:</strong></td><td>' + blockerData.timestamp + '</td></tr>';
    body += '</table>';
    body += '</div>';

    // Alert message
    if (blockerData.totalProdBugs > 0) {
      body += '<div style="background: #FFF9C4; border: 2px solid #F57C00; padding: 15px; margin: 20px 0; border-radius: 4px;">';
      body += '<h4 style="margin-top: 0; color: #E65100;">🚨 EMERGENCY - IMMEDIATE ACTION REQUIRED!</h4>';
      body += '<p style="color: #D32F2F; font-weight: bold; margin: 10px 0;">Production bugs affect LIVE USERS</p>';
      body += '<h4 style="color: #E65100;">⚡ PROTOCOL:</h4>';
      body += '<ul style="margin: 8px 0;">';
      body += '<li>Drop other work - TOP PRIORITY</li>';
      body += '<li>Assign senior developer immediately</li>';
      body += '<li>Hourly status updates required</li>';
      body += '<li>Hotfix deployment prioritized</li>';
      body += '<li><strong>Target: Resolve within 4-8 hours</strong></li>';
      body += '</ul>';
      body += '</div>';
    } else {
      body += '<div style="background: #E3F2FD; border-left: 3px solid #1976D2; padding: 15px; margin: 20px 0;">';
      body += '<h4 style="margin-top: 0; color: #1565C0;">📋 ACTION REQUIRED</h4>';
      body += '<p style="color: #E65100; font-weight: bold;">Critical/High/Medium bugs blocking test execution</p>';
      body += '<ul style="margin: 8px 0;">';
      body += '<li>Impact: Testing cannot proceed</li>';
      body += '<li>Daily standup: Report status</li>';
      body += '<li>Escalate if stuck >2 days</li>';
      body += '<li><strong>Target: Resolve within 48 hours</strong></li>';
      body += '</ul>';
      body += '</div>';
    }

    // Modules breakdown
    body += '<div style="margin: 20px 0;">';
    body += '<h3 style="color: #424242; border-bottom: 2px solid #1976D2; padding-bottom: 8px;">📦 Modules Breakdown</h3>';

    blockerData.modules.forEach((module, index) => {
      const bgColor = index % 2 === 0 ? '#F5F5F5' : '#FFFFFF';

      // Get actual BugReport GID from QATM
      let actualBugReportGid = module.bugReportGid || '2';
      if (module.qatmUrl) {
        try {
          const match = module.qatmUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
          if (match) {
            const qatmId = match[1];
            const qatmSs = SpreadsheetApp.openById(qatmId);
            const bugSheet = qatmSs.getSheetByName('BugReport');
            if (bugSheet) {
              actualBugReportGid = bugSheet.getSheetId();
            }
          }
        } catch (e) {
          Logger.log('Error getting BugReport GID for ' + module.project + '-' + module.module + ': ' + e.message);
        }
      }

      body += '<div style="background: ' + bgColor + '; border: 1px solid #E0E0E0; padding: 15px; margin: 10px 0; border-radius: 4px;">';
      body += '<h4 style="margin: 0 0 10px 0; color: #1976D2;">' + module.project + ' - Module ' + module.module + '</h4>';
      body += '<table style="width: 100%; font-size: 13px;">';
      body += '<tr><td style="width: 40%;"><strong>Submodule:</strong></td><td>' + module.submodule + '</td></tr>';
      body += '<tr><td><strong>Blockers:</strong></td><td><span style="background: #FF9800; color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold;">' + module.blocker + '</span></td></tr>';

      // Severity breakdown
      if ((module.critical || 0) > 0 || (module.high || 0) > 0 || (module.medium || 0) > 0) {
        body += '<tr><td><strong>Severity:</strong></td><td>';
        if ((module.critical || 0) > 0) {
          body += '<span style="background: #D32F2F; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin-right: 4px;">🔴 ' + module.critical + '</span> ';
        }
        if ((module.high || 0) > 0) {
          body += '<span style="background: #FF9800; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin-right: 4px;">🟠 ' + module.high + '</span> ';
        }
        if ((module.medium || 0) > 0) {
          body += '<span style="background: #FFC107; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px;">🟡 ' + module.medium + '</span>';
        }
        body += '</td></tr>';
      }

      if (module.prodBugs > 0) {
        body += '<tr><td><strong>PROD Bugs:</strong></td><td><span style="background: #D32F2F; color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold;">' + module.prodBugs + '</span></td></tr>';
      }
      if (module.qatmUrl) {
        body += '<tr><td colspan="2"><a href="' + module.qatmUrl + '#gid=' + actualBugReportGid + '" style="color: #1976D2; font-weight: bold;">📋 View Bug Report</a></td></tr>';
      }
      body += '</table>';
      body += '</div>';
    });
    body += '</div>';

    // Priority levels explanation
    body += '<div style="background: #F5F5F5; border: 1px solid #E0E0E0; padding: 15px; margin: 20px 0;">';
    body += '<h3 style="color: #424242; margin-top: 0;">📊 PRIORITY LEVELS EXPLAINED</h3>';
    body += '<table style="width: 100%; border-collapse: collapse;">';
    body += '<tr style="background: #FFEBEE;"><td style="padding: 8px; border: 1px solid #E0E0E0;"><strong>🔴 Critical</strong></td><td style="padding: 8px; border: 1px solid #E0E0E0;">System down, data loss, security breach → <strong>Hotfix &lt;24h</strong></td></tr>';
    body += '<tr style="background: #FFF3E0;"><td style="padding: 8px; border: 1px solid #E0E0E0;"><strong>🟠 High</strong></td><td style="padding: 8px; border: 1px solid #E0E0E0;">Major feature broken, high business impact → <strong>Fix &lt;48h</strong></td></tr>';
    body += '<tr style="background: #FFF9C4;"><td style="padding: 8px; border: 1px solid #E0E0E0;"><strong>🟡 Medium</strong></td><td style="padding: 8px; border: 1px solid #E0E0E0;">Minor feature issue, workaround available → <strong>Fix in sprint</strong></td></tr>';
    body += '<tr style="background: #F1F8E9;"><td style="padding: 8px; border: 1px solid #E0E0E0;"><strong>🟢 Low</strong></td><td style="padding: 8px; border: 1px solid #E0E0E0;">Cosmetic, low impact → <strong>Fix when convenient</strong></td></tr>';
    body += '</table>';
    body += '</div>';

    // Links section
    body += '<div style="background: #E8EAF6; border-left: 4px solid #3F51B5; padding: 15px; margin: 20px 0;">';
    body += '<h3 style="color: #3F51B5; margin-top: 0;">🔗 QUICK LINKS</h3>';
    body += '<p style="margin: 10px 0;"><a href="' + dashboardUrl + '" style="background: #1976D2; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">📊 View QA Dashboard</a></p>';
    body += '<p style="color: #757575; font-size: 12px; margin: 10px 0;">View individual module bug reports in the "Modules Breakdown" section above</p>';
    body += '</div>';

    body += '<hr style="border: none; border-top: 1px solid #E0E0E0; margin: 20px 0;">';
    body += '<p style="color: #757575; font-size: 11px; text-align: center;">📊 QA Blocker Alert - Automated Notification<br>';
    body += blockerData.modules.length + ' module(s) affected • For questions, contact QA Team</p>';
    body += '</body></html>';

    // Send email
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: body
    });

    Logger.log('✅ Email notification sent to: ' + recipients);
    return true;
  } catch (e) {
    Logger.log('❌ Error sending email notification: ' + e.message);
    return false;
  }
}

/**
 * Send WhatsApp notification via Fonnte API - Plain text format (WhatsApp-friendly)
 *
 * @param {string} groupId - WhatsApp Group ID (format: 120363xxxxx@g.us)
 * @param {object} blockerData - Blocker data with modules array
 * @param {string} fontteToken - Fonnte API token
 * @returns {boolean} - Success status
 */
function sendWhatsAppNotification_(groupId, blockerData, fontteToken) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardUrl = ss.getUrl();
    const dashboardBugsUrl = dashboardUrl + '#gid=' + getDashboardBugsGid_(ss);
    const dashboardOverviewUrl = dashboardUrl + '#gid=' + getDashboardOverviewGid_(ss);

    // Get detailed bug data from QATM BugReport sheets
    const bugDetails = getBugDetailsFromQATM_(blockerData.modules);

    // Build WhatsApp message (plain text, WhatsApp markdown)
    let message = '';

    // ══════════════════════════════════════════════════════════════════════
    // HEADER - Daily Bug Report Summary
    // ══════════════════════════════════════════════════════════════════════
    message += '📊 *DAILY BUG REPORT*\n';
    message += '📅 ' + blockerData.timestamp + '\n';
    message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

    // Calculate total severity counts (same as email)
    const totalCritical = blockerData.modules.reduce((sum, m) => sum + (m.critical || 0), 0);
    const totalHigh = blockerData.modules.reduce((sum, m) => sum + (m.high || 0), 0);
    const totalMedium = blockerData.modules.reduce((sum, m) => sum + (m.medium || 0), 0);

    // Summary
    message += '📈 *SUMMARY*\n';
    message += '🐛 Total Bugs: *' + blockerData.totalBlockers + '*';
    if (blockerData.totalProdBugs > 0) {
      message += '  |  🚨 PROD: *' + blockerData.totalProdBugs + '*';
    }
    message += '\n';

    // Severity breakdown
    message += '*Severity:* ';
    if (totalCritical > 0) message += 'Critical🔴 ' + totalCritical + '  ';
    if (totalHigh > 0) message += 'High🟠 ' + totalHigh + '  ';
    if (totalMedium > 0) message += 'Medium🟡 ' + totalMedium;
    message += '\n\n';

    // ══════════════════════════════════════════════════════════════════════
    // PRODUCTION BUGS SECTION (if any) - TOP PRIORITY
    // ══════════════════════════════════════════════════════════════════════
    const hasProdBugs = blockerData.modules.some(m => m.prodBugs > 0);

    if (hasProdBugs) {
      message += '🚨🚨🚨 *PRODUCTION BUGS* 🚨🚨🚨\n';
      message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

      blockerData.modules.forEach(module => {
        if (module.prodBugs > 0) {
          message += '*' + module.project + ' - ' + module.module + '*\n';
          message += '📌 ' + module.submodule + ': *' + module.prodBugs + '* PROD bug(s)\n';

          if (module.qatmUrl) {
            const gid = getBugReportGidFromQATM_(module.qatmUrl);
            message += '📋 ' + module.qatmUrl + '#gid=' + gid + '\n';
            Logger.log('WhatsApp link (PROD): ' + module.qatmUrl + '#gid=' + gid);
          }
          message += '\n';
        }
      });

      message += '⚠️ *URGENT - IMMEDIATE ACTION REQUIRED!*\n\n';
    }

    // ══════════════════════════════════════════════════════════════════════
    // ALL MODULES - Blocker Breakdown
    // ══════════════════════════════════════════════════════════════════════
    if (blockerData.totalBlockers > 0) {
      message += '═══════════════════════════════════\n';
      message += '⚠️ *BLOCKER BUGS*\n';
      message += '═══════════════════════════════════\n\n';

      blockerData.modules.forEach(module => {
        if (module.blocker > 0 || module.prodBugs > 0) {
          message += '*' + module.project + ' - ' + module.module + '*\n';
          message += '📌 ' + module.submodule + '\n';
          message += '   🐛 Blockers: *' + module.blocker + '*';

          // Severity breakdown
          if ((module.critical || 0) > 0 || (module.high || 0) > 0 || (module.medium || 0) > 0) {
            message += '\n   ';
            if ((module.critical || 0) > 0) message += 'Critical🔴 ' + module.critical + '  ';
            if ((module.high || 0) > 0) message += 'High🟠 ' + module.high + '  ';
            if ((module.medium || 0) > 0) message += 'Medium🟡 ' + module.medium;
          }

          message += '\n';

          if (module.qatmUrl) {
            const gid = getBugReportGidFromQATM_(module.qatmUrl);
            message += '   📋 ' + module.qatmUrl + '#gid=' + gid + '\n';
            Logger.log('WhatsApp link (BLOCKER): ' + module.qatmUrl + '#gid=' + gid);
          }

          message += '\n';
        }
      });
    }

    // ══════════════════════════════════════════════════════════════════════
    // FOOTER: Dashboard Links
    // ══════════════════════════════════════════════════════════════════════
    message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
    message += '🔗 *Dashboard Links:*\n';
    message += '📊 Overview: ' + dashboardOverviewUrl + '\n';
    message += '🐛 Bugs: ' + dashboardBugsUrl + '\n\n';
    message += '_Automated Daily Report - QA Dashboard_';

    // Send via Fonnte API
    const url = 'https://api.fonnte.com/send';

    // For groups: don't include countryCode (causes "invalid group id" error)
    const payload = {
      target: groupId,
      message: message
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': fontteToken
      },
      payload: payload,
      muteHttpExceptions: true
    };

    Logger.log('Sending WhatsApp notification to: ' + groupId);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('WhatsApp Response Code: ' + responseCode);
    Logger.log('WhatsApp Response: ' + responseText);

    if (responseCode === 200) {
      const data = JSON.parse(responseText);

      if (data.status) {
        Logger.log('✅ WhatsApp notification sent successfully');
        return true;
      } else {
        Logger.log('❌ WhatsApp send failed: ' + (data.reason || 'Unknown error'));
        return false;
      }
    } else {
      Logger.log('❌ WhatsApp API error: ' + responseCode + ' - ' + responseText);
      return false;
    }

  } catch (e) {
    Logger.log('❌ Error sending WhatsApp notification: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    return false;
  }
}

/**
 * Get GID for Dashboard Overview tab
 */
function getDashboardOverviewGid_(ss) {
  const overviewSheet = ss.getSheetByName('Overview');
  if (overviewSheet) {
    return overviewSheet.getSheetId();
  }
  return '0'; // Default to first sheet
}

/**
 * Get GID for Dashboard Bugs tab
 */
function getDashboardBugsGid_(ss) {
  const bugsSheet = ss.getSheetByName('Bugs');
  if (bugsSheet) {
    return bugsSheet.getSheetId();
  }
  return '0'; // Default to first sheet
}

/**
 * Get actual BugReport GID from QATM sheet
 * @param {string} qatmUrl - QATM spreadsheet URL
 * @returns {string} - BugReport sheet GID or '2' as fallback
 */
function getBugReportGidFromQATM_(qatmUrl) {
  if (!qatmUrl) return '2';

  try {
    const match = qatmUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) return '2';

    const qatmId = match[1];
    const qatmSs = SpreadsheetApp.openById(qatmId);
    const bugSheet = qatmSs.getSheetByName('BugReport');

    if (bugSheet) {
      return bugSheet.getSheetId().toString();
    }
  } catch (e) {
    Logger.log('Error getting BugReport GID from QATM: ' + e.message);
  }

  return '2'; // Default fallback
}

