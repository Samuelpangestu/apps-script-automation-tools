/**
 * notifications.js
 * Send blocker & PROD BUGS notifications via Google Chat & Email
 * Updated: Changed bullet points from ● to ▬ for consistent formatting
 */

// Latest Web App URL - Update this when deploying new web app version
const LATEST_WEBAPP_URL = 'https://script.google.com/a/macros/inadigital.co.id/s/AKfycbxym3cABwoaZG20jJeyJ1O1UPiz5gDWpEvHiqv67OqJSgevEsFDrnMxqwpgOyk8VyDU4g/exec';

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
    Logger.log('Note: VAPT data is checked per-project and will be included if configured');
    SpreadsheetApp.getUi().alert(
      '✅ All Clear!',
      'Tidak ada Open Blocker atau PROD BUGS.\n\n' +
      'Notification tidak dikirim (tidak ada yang perlu di-alert).\n\n' +
      'Note: VAPT data akan dicek per-project jika dikonfigurasi.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Log summary
  Logger.log('Total modules with bugs: ' + blockerData.modules.length);

  // ── GET SHARED FONNTE TOKEN (row 4, col T) ──────────────────────────────
  // Fonnte Token is SHARED config at row 4, column T (index 20)
  // Used by all modules for sending WhatsApp messages
  let fontteToken = null;
  let globalFallbackGroupId = null;

  if (cfg) {
    fontteToken = String(cfg.getRange(4, 20).getValue()).trim();        // T4 = Fonnte Token (SHARED)
    globalFallbackGroupId = String(cfg.getRange(4, 19).getValue()).trim(); // S4 = Global fallback Group ID
  }

  // ── LOAD VAPT CONFIG FROM ALL MODULES (not just modules with bugs) ────────
  // VAPT config is per-project, but we need to check ALL modules to find valid config
  // Some modules may have VAPT config but no bugs (so not in blockerData.modules)
  // ────────────────────────────────────────────────────────────────────────
  const allVaptConfigs = {};  // { projectName: { spreadsheetId, enabled } }

  if (cfg) {
    const cfgData = cfg.getDataRange().getValues();
    for (let i = 3; i < cfgData.length; i++) {  // Start from row 4 (index 3)
      const projectName = String(cfgData[i][2]).trim();  // Col C = Project
      if (!projectName) break;  // Stop on empty project

      const vaptSpreadsheetId = String(cfgData[i][21]).trim();  // Col V = VAPT Spreadsheet ID
      const vaptEnabled = cfgData[i][22] === true;              // Col W = Enable VAPT

      // Store VAPT config for this project (last valid config wins)
      if (vaptEnabled && vaptSpreadsheetId) {
        allVaptConfigs[projectName] = {
          spreadsheetId: vaptSpreadsheetId,
          enabled: vaptEnabled
        };
        Logger.log('📋 VAPT config loaded from Config: ' + projectName + ' | Spreadsheet: ' + vaptSpreadsheetId);
      }
    }
  }

  // ── GROUP BY PROJECT ────────────────────────────────────────────────────
  // Aggregate modules by PROJECT NAME → send 1 notification per project
  // Notification config (WhatsApp, Email, Chat) diambil dari module pertama yang punya config
  // ────────────────────────────────────────────────────────────────────────

  const projectGroups = {};  // { projectName: { modules: [...], config: {...} } }

  blockerData.modules.forEach(module => {
    const projectName = module.project || 'Unknown';

    if (!projectGroups[projectName]) {
      projectGroups[projectName] = {
        modules: [],
        chatConfig: null,      // First module with chat config
        emailConfig: null,     // First module with email config
        whatsappConfig: null,  // First module with WhatsApp config
        vaptConfig: null       // First module with VAPT config
      };
    }

    projectGroups[projectName].modules.push(module);

    // Store first module with notification config (per type)
    if (module.chatEnabled && module.chatWebhook && !projectGroups[projectName].chatConfig) {
      projectGroups[projectName].chatConfig = {
        webhook: module.chatWebhook,
        enabled: module.chatEnabled
      };
    }

    if (module.emailEnabled && module.emailRecipients && !projectGroups[projectName].emailConfig) {
      projectGroups[projectName].emailConfig = {
        recipients: module.emailRecipients,
        enabled: module.emailEnabled
      };
    }

    if (module.whatsappEnabled && module.whatsappGroupId && !projectGroups[projectName].whatsappConfig) {
      projectGroups[projectName].whatsappConfig = {
        groupId: module.whatsappGroupId,
        enabled: module.whatsappEnabled
      };
    } else if (module.whatsappEnabled && !module.whatsappGroupId && !projectGroups[projectName].whatsappConfig && globalFallbackGroupId && globalFallbackGroupId.includes('@g.us')) {
      // Fallback: use global Group ID if module enabled but no Group ID
      projectGroups[projectName].whatsappConfig = {
        groupId: globalFallbackGroupId,
        enabled: true
      };
    }

    // VAPT config will be set from allVaptConfigs below (not from modules)
    // Removed old logic that only checked modules with bugs
  });

  // ── APPLY VAPT CONFIG FROM ALL MODULES (including those without bugs) ────────
  // Now that projectGroups is created, apply VAPT config from allVaptConfigs
  Object.keys(projectGroups).forEach(projectName => {
    if (allVaptConfigs[projectName]) {
      projectGroups[projectName].vaptConfig = allVaptConfigs[projectName];
      Logger.log('✅ VAPT config applied to project: ' + projectName);
    }
  });

  // Log project summary
  Logger.log('Project Groups: ' + Object.keys(projectGroups).length + ' | Fonnte Token: ' + (fontteToken ? 'SET' : 'NOT SET'));

  // ── SEND PROJECT-BASED NOTIFICATIONS ────────────────────────────────────
  // Send 1 notification per project (all modules in project aggregated)
  // ────────────────────────────────────────────────────────────────────────
  let chatSentCount = 0;
  let emailSentCount = 0;
  let whatsappSentCount = 0;
  let projectsProcessed = 0;

  Object.keys(projectGroups).forEach(projectName => {
    const pg = projectGroups[projectName];
    const modules = pg.modules;

    Logger.log('Processing project: ' + projectName + ' (' + modules.length + ' modules)');

    // Fetch per-project VAPT data (if configured)
    let projectVaptData = {
      vaptBlocker: 0,
      vaptAppsWithBlockers: 0,
      vaptApps: [],
      vaptBreakdown: { critical: 0, high: 0, medium: 0 }
    };

    if (pg.vaptConfig && pg.vaptConfig.enabled && pg.vaptConfig.spreadsheetId) {
      Logger.log('📊 Fetching VAPT data for project: ' + projectName + ' | Spreadsheet: ' + pg.vaptConfig.spreadsheetId);
      projectVaptData = fetchVAPTDataForProject_(pg.vaptConfig.spreadsheetId);
      Logger.log('📊 VAPT data fetched: blocker=' + projectVaptData.vaptBlocker + ', apps=' + projectVaptData.vaptAppsWithBlockers);
    } else {
      Logger.log('⚠️ VAPT config not found or invalid for project: ' + projectName +
                 ' | vaptConfig=' + (pg.vaptConfig ? 'EXISTS' : 'NULL') +
                 ' | enabled=' + (pg.vaptConfig ? pg.vaptConfig.enabled : 'N/A') +
                 ' | spreadsheetId=' + (pg.vaptConfig ? (pg.vaptConfig.spreadsheetId ? 'EXISTS' : 'EMPTY') : 'N/A'));
    }

    // Aggregate project-level data
    const aggregatedData = {
      projectName: projectName,                                     // Project name for subject/header
      modules: modules,
      totalBlockers: modules.reduce((sum, m) => sum + m.blocker, 0),
      totalProdBugs: modules.reduce((sum, m) => sum + m.prodBugs, 0),
      vaptBlocker: projectVaptData.vaptBlocker,                     // Per-project VAPT
      vaptAppsWithBlockers: projectVaptData.vaptAppsWithBlockers,   // Per-project VAPT
      vaptApps: projectVaptData.vaptApps,                           // Per-project VAPT detail
      vaptBreakdown: projectVaptData.vaptBreakdown,                 // Per-project VAPT severity
      timestamp: blockerData.timestamp
    };

    let projectSent = false;

    // Send Google Chat (if configured for this project)
    if (pg.chatConfig && pg.chatConfig.enabled && pg.chatConfig.webhook) {
      if (sendGoogleChatNotification_(pg.chatConfig.webhook, aggregatedData)) {
        chatSentCount++;
        projectSent = true;
      }
    }

    // Send Email (if configured for this project)
    if (pg.emailConfig && pg.emailConfig.enabled && pg.emailConfig.recipients) {
      if (sendEmailNotification_(pg.emailConfig.recipients, aggregatedData)) {
        emailSentCount++;
        projectSent = true;
      }
    }

    // Send WhatsApp (if configured for this project)
    if (pg.whatsappConfig && pg.whatsappConfig.enabled && pg.whatsappConfig.groupId) {
      if (fontteToken && fontteToken.length > 10) {
        if (sendWhatsAppNotification_(pg.whatsappConfig.groupId, aggregatedData, fontteToken)) {
          whatsappSentCount++;
          projectSent = true;
        }
      } else {
        Logger.log('⚠️ WhatsApp enabled but Fonnte Token missing for ' + projectName);
      }
    }

    if (projectSent) {
      projectsProcessed++;
    }
  });

  const skippedCount = Object.keys(projectGroups).length - projectsProcessed;

  // Show result
  let msg = '📤 Notifications Sent (Project-Based)!\n\n';
  if (chatSentCount > 0) msg += '✅ Google Chat: ' + chatSentCount + ' project(s)\n';
  if (emailSentCount > 0) msg += '✅ Email: ' + emailSentCount + ' project(s)\n';
  if (whatsappSentCount > 0) msg += '✅ WhatsApp: ' + whatsappSentCount + ' project(s)\n';
  if (skippedCount > 0) msg += 'ℹ️ Skipped: ' + skippedCount + ' project(s) (all channels disabled)\n';
  msg += '\n📊 Summary:\n';
  msg += '• Total Projects: ' + Object.keys(projectGroups).length + '\n';
  msg += '• Projects notified: ' + projectsProcessed + '\n';
  msg += '• Total Modules with issues: ' + blockerData.modules.length + '\n';
  msg += '• Total Open Blockers: ' + blockerData.totalBlockers + '\n';
  msg += '• Total PROD BUGS: ' + blockerData.totalProdBugs + '\n';
  msg += '• VAPT: Dicek per-project (jika dikonfigurasi)\n\n';
  msg += '💡 Notification config diambil dari module pertama per project';

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
    '📱 GOOGLE CHAT (kolom L-N) - Per Project\n' +
    '  L = Webhook URL (dari Google Chat Space)\n' +
    '  M = Schedule (9,14,18 atau 4h)\n' +
    '  N = ☑ Enable Chat\n\n' +
    '📧 EMAIL (kolom O-P) - Per Project\n' +
    '  O = Email recipients (comma separated)\n' +
    '  P = ☑ Enable Email\n\n' +
    '📲 WHATSAPP (kolom S-U) - Per Project\n' +
    '  S = Group ID (120363xxx@g.us)\n' +
    '  T = Fonnte Token (row 4 ONLY - shared)\n' +
    '  U = ☑ Enable WhatsApp\n\n' +
    '💡 TIP: Cukup isi config di 1 module per project\n' +
    'Semua module dalam 1 project kirim ke channel sama\n\n' +
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
 * Generic helper: Send WhatsApp message via Fonnte
 * @param {string} fontteToken - Fonnte API token
 * @param {string} groupId - WhatsApp group ID (format: 120363xxx@g.us)
 * @param {string} message - Plain text message
 * @returns {boolean} Success status
 */
function sendWhatsApp_(fontteToken, groupId, message) {
  if (!fontteToken || !groupId || !message) {
    Logger.log('⚠️ Missing required parameters for WhatsApp');
    return false;
  }

  try {
    const url = 'https://api.fonnte.com/send';
    const payload = {
      target: groupId,
      message: message,
      countryCode: '62'
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': fontteToken
      },
      payload: payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      Logger.log('✅ WhatsApp sent to: ' + groupId);
      return true;
    } else {
      Logger.log('❌ WhatsApp failed (' + responseCode + '): ' + responseText);
      return false;
    }
  } catch (e) {
    Logger.log('❌ WhatsApp error: ' + e.toString());
    return false;
  }
}

/**
 * Generic helper: Send email with HTML body
 * @param {string} recipients - Email recipients (comma separated)
 * @param {string} subject - Email subject
 * @param {string} htmlBody - HTML email body
 * @returns {boolean} Success status
 */
function sendEmail_(recipients, subject, htmlBody) {
  if (!recipients || !subject || !htmlBody) {
    Logger.log('⚠️ Missing required parameters for Email');
    return false;
  }

  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log('✅ Email sent to: ' + recipients);
    return true;
  } catch (e) {
    Logger.log('❌ Email error: ' + e.toString());
    return false;
  }
}

/**
 * Generic helper: Send Google Chat message
 * @param {string} webhookUrl - Google Chat webhook URL
 * @param {Object} payload - Chat card payload (cardsV2 format)
 * @returns {boolean} Success status
 */
function sendGoogleChat_(webhookUrl, payload) {
  if (!webhookUrl || !payload) {
    Logger.log('⚠️ Missing required parameters for Google Chat');
    return false;
  }

  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      Logger.log('✅ Google Chat sent');
      return true;
    } else {
      Logger.log('❌ Google Chat failed (' + responseCode + '): ' + responseText);
      return false;
    }
  } catch (e) {
    Logger.log('❌ Google Chat error: ' + e.toString());
    return false;
  }
}

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
    Logger.log('⚠️ Bugs tab not found - using Overview as fallback');
    data = overview.getDataRange().getValues();
  } else {
    data = bugsTab.getDataRange().getValues();
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
      // Col S (18) = WhatsApp Group ID (per-module) ✨ NEW
      // Col U (20) = Enable WhatsApp (per-module) ✨ NEW
      // Col T (19) = Fonnte Token (row 4 only, shared for all modules)
      // Col V (21) = VAPT Spreadsheet ID (per-project) ✨ NEW
      // Col W (22) = Enable VAPT (per-project) ✨ NEW
      const chatWebhook = String(cfgData[i][11]).trim();
      const chatEnabled = cfgData[i][13] === true;
      const emailRecipients = String(cfgData[i][14]).trim();
      const emailEnabled = cfgData[i][15] === true;
      const whatsappGroupId = String(cfgData[i][18]).trim(); // Col S (index 18)
      const whatsappEnabled = cfgData[i][20] === true;       // Col U (index 20)
      const vaptSpreadsheetId = String(cfgData[i][21]).trim(); // Col V (index 21)
      const vaptEnabled = cfgData[i][22] === true;            // Col W (index 22)

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
          emailEnabled: emailEnabled,
          whatsappGroupId: whatsappGroupId && whatsappGroupId.includes('@g.us') ? whatsappGroupId : null,
          whatsappEnabled: whatsappEnabled,
          vaptSpreadsheetId: vaptSpreadsheetId && vaptSpreadsheetId.length > 10 ? vaptSpreadsheetId : null,
          vaptEnabled: vaptEnabled
        };
      }
    }
  }

  Logger.log('Config loaded: ' + Object.keys(moduleMap).length + ' modules');

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
      // Col A (index 0) = Project, Col B (index 1) = Modul, Col C (index 2) = Submodul
      if (!row[1] && !row[2]) continue; // Empty row

      project = String(row[0]).trim();        // Col A = Project (e.g., "SIPGN", "INADigital")
      modul = String(row[1]).trim();          // Col B = Modul (e.g., "1", "3", "4")
      moduleName = String(row[2]).trim();     // Col C = Submodul (e.g., "Portal + SSO")

      // Skip rows with NO submodule name
      if (!moduleName || moduleName === '') {
        continue;
      }

      critical = parseInt(row[4]) || 0;       // Col E = Critical
      high = parseInt(row[5]) || 0;           // Col F = High
      medium = parseInt(row[6]) || 0;         // Col G = Medium
      blocker = parseInt(row[9]) || 0;        // Col J = Blocker
      prodBugs = parseInt(row[14]) || 0;      // Col O = Prod (index 14, NOT 11!)
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
        moduleInfo = {};
      }

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
        emailEnabled: moduleInfo.emailEnabled || false,
        whatsappGroupId: moduleInfo.whatsappGroupId || null,   // ✨ NEW - per-module WhatsApp Group ID
        whatsappEnabled: moduleInfo.whatsappEnabled || false,  // ✨ NEW - per-module WhatsApp enable
        vaptSpreadsheetId: moduleInfo.vaptSpreadsheetId || null, // ✨ NEW - per-project VAPT Spreadsheet ID
        vaptEnabled: moduleInfo.vaptEnabled || false           // ✨ NEW - per-project VAPT enable
      });
      totalBlockers += blocker;
      totalProdBugs += prodBugs;
    }
  }

  // ═══════════════════════════════════════════════════════════════════════
  // RETURN MODULE DATA
  // Note: VAPT data is now fetched per-project in sendBlockerNotification()
  // ═══════════════════════════════════════════════════════════════════════

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
    message += '📊 *DAILY BUG REPORT*';
    if (blockerData.projectName) {
      message += ' - ' + blockerData.projectName;
    }
    message += '\n📅 ' + blockerData.timestamp + '\n';
    message += '━━━━━━━━━━━━━\n\n';

    // Calculate totals from blockerData (same data source as Email/WhatsApp for consistency)
    const totalCritical = blockerData.modules.reduce((sum, m) => sum + (m.critical || 0), 0);
    const totalHigh = blockerData.modules.reduce((sum, m) => sum + (m.high || 0), 0);
    const totalMedium = blockerData.modules.reduce((sum, m) => sum + (m.medium || 0), 0);

    // Header summary
    message += '📈 *SUMMARY*\n';
    message += '🐛 QA Bugs: ' + blockerData.totalBlockers + '  |  🚨 PROD: ' + blockerData.totalProdBugs + '\n';

    // QA Severity (show only non-zero)
    const qaSevParts = [];
    if (totalCritical > 0) qaSevParts.push('Critical🟣 ' + totalCritical);
    if (totalHigh > 0) qaSevParts.push('High🔴 ' + totalHigh);
    if (totalMedium > 0) qaSevParts.push('Medium🟠 ' + totalMedium);
    if (qaSevParts.length > 0) {
      message += '  Severity: ' + qaSevParts.join('  ') + '\n';
    }

    // VAPT Blocker summary
    if (blockerData.vaptBlocker > 0) {
      message += '🔒 VAPT Blocker: *' + blockerData.vaptBlocker + '* (' + blockerData.vaptAppsWithBlockers + ' apps)\n';

      // VAPT Severity (show only non-zero)
      const vaptSevParts = [];
      if (blockerData.vaptBreakdown.critical > 0) vaptSevParts.push('Critical🟣 ' + blockerData.vaptBreakdown.critical);
      if (blockerData.vaptBreakdown.high > 0) vaptSevParts.push('High🔴 ' + blockerData.vaptBreakdown.high);
      if (blockerData.vaptBreakdown.medium > 0) vaptSevParts.push('Medium🟠 ' + blockerData.vaptBreakdown.medium);
      if (vaptSevParts.length > 0) {
        message += '  ' + vaptSevParts.join('  ') + '\n';
      }
    }
    message += '\n';

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
    message += '━━━━━━━━━━━━━\n\n';

    // ══════════════════════════════════════════════════════════════════════
    // VAPT BLOCKER DETAIL (if any) - Per aplikasi
    // ══════════════════════════════════════════════════════════════════════
    if (blockerData.vaptBlocker > 0) {
      message += '🔒 *VAPT BLOCKER DETAIL* 🔒\n';
      message += '━━━━━━━━━━━━━\n\n';

      // List apps with blocker > 0
      blockerData.vaptApps.forEach(app => {
        message += app.aplikasi + ': ' + app.blocker + '\n';

        // Show only non-zero severities
        const appSevParts = [];
        if (app.critical > 0) appSevParts.push('Critical🟣 ' + app.critical);
        if (app.high > 0) appSevParts.push('High🔴 ' + app.high);
        if (app.medium > 0) appSevParts.push('Medium🟠 ' + app.medium);
        if (appSevParts.length > 0) {
          message += '   ' + appSevParts.join('  ') + '\n';
        }
      });

      message += '\n';
    }

    // ══════════════════════════════════════════════════════════════════════
    // PRODUCTION BUGS SECTION (ALL MODULES) - PRIORITAS PERTAMA
    // ══════════════════════════════════════════════════════════════════════
    let hasProdBugs = false;
    Object.keys(moduleGroups).forEach(key => {
      if (Object.keys(moduleGroups[key].prodSubmodules).length > 0) hasProdBugs = true;
    });

    if (hasProdBugs) {
      message += '🚨 *PRODUCTION BUGS* 🚨\n';
      message += '<users/all>\n';
      message += '━━━━━━━━━━━━━\n\n';

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

          // Show only non-zero severities
          const prodSevParts = [];
          if (criticalBugs.length > 0) prodSevParts.push('Critical🟣 ' + criticalBugs.length);
          if (highBugs.length > 0) prodSevParts.push('High🔴 ' + highBugs.length);
          if (mediumBugs.length > 0) prodSevParts.push('Medium🟠 ' + mediumBugs.length);

          message += '📌 ' + submodulName + ': ' + bugs.length + ' bug' + (bugs.length > 1 ? 's' : '');
          if (prodSevParts.length > 0) {
            message += ' (' + prodSevParts.join('  ') + ')';
          }
          message += '\n';
        });

        message += '\n';
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
    // Get Web App URL from Script Properties (fallback to latest)
    const scriptProps = PropertiesService.getScriptProperties();
    const dashboardWebAppUrl = scriptProps.getProperty('WEB_APP_URL') || LATEST_WEBAPP_URL;

    message += '━━━━━━━━━━━━━\n';
    message += '🔗 <' + dashboardWebAppUrl + '|📊 Web Dashboard (QA & VAPT)>  •  <' + dashboardOverviewUrl + '|📋 Sheet Overview>  •  <' + dashboardBugsUrl + '|🐛 Sheet Bugs>';
    message += '\n_Automated Daily Report_';

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

    // Build subject based on aggregated data (include project name)
    const projectPrefix = blockerData.projectName ? blockerData.projectName + ' - ' : '';
    let subject = '🚨 ' + projectPrefix + 'QA Blocker Alert: ' + blockerData.totalBlockers + ' blockers';
    if (blockerData.totalProdBugs > 0) {
      subject = '🚨🚨 ' + projectPrefix + 'URGENT PROD BUGS: ' + blockerData.totalProdBugs + ' bugs affecting production';
    }

    let body = '<html><body style="font-family: Arial, sans-serif; max-width: 700px;">';

    // Header
    if (blockerData.totalProdBugs > 0) {
      body += '<div style="background: #D32F2F; color: white; padding: 20px; text-align: center;">';
      body += '<h2 style="margin: 0;">🚨 PRODUCTION BUGS ALERT 🚨</h2>';
      if (blockerData.projectName) {
        body += '<h3 style="margin: 5px 0 0 0; font-weight: normal;">Project: ' + blockerData.projectName + '</h3>';
      }
      body += '</div>';
    } else {
      body += '<div style="background: #FF9800; color: white; padding: 15px; text-align: center;">';
      body += '<h2 style="margin: 0;">⚠️ QA BLOCKER ALERT</h2>';
      if (blockerData.projectName) {
        body += '<h3 style="margin: 5px 0 0 0; font-weight: normal;">Project: ' + blockerData.projectName + '</h3>';
      }
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

    // Severity breakdown (show only non-zero)
    const severityHtml = [];
    if (totalCritical > 0) {
      severityHtml.push('<span style="background: #9C27B0; color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold; margin-right: 4px;">🟣 ' + totalCritical + '</span>');
    }
    if (totalHigh > 0) {
      severityHtml.push('<span style="background: #D32F2F; color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold; margin-right: 4px;">🔴 ' + totalHigh + '</span>');
    }
    if (totalMedium > 0) {
      severityHtml.push('<span style="background: #FF9800; color: white; padding: 2px 8px; border-radius: 3px; font-weight: bold;">🟠 ' + totalMedium + '</span>');
    }
    if (severityHtml.length > 0) {
      body += '<tr><td><strong>Severity:</strong></td><td>' + severityHtml.join(' ') + '</td></tr>';
    }

    if (blockerData.totalProdBugs > 0) {
      body += '<tr><td><strong style="color: #D32F2F;">PROD Bugs:</strong></td><td><span style="background: #D32F2F; color: white; padding: 4px 12px; border-radius: 3px; font-weight: bold; font-size: 18px;">' + blockerData.totalProdBugs + '</span></td></tr>';
    }

    // VAPT Blocker
    if (blockerData.vaptBlocker > 0) {
      body += '<tr><td><strong style="color: #EF6C00;">🔒 VAPT Blocker:</strong></td><td><span style="background: #EF6C00; color: white; padding: 4px 12px; border-radius: 3px; font-weight: bold; font-size: 18px;">' + blockerData.vaptBlocker + '</span></td></tr>';

      // VAPT Severity (show only non-zero)
      const vaptSeverityHtml = [];
      if (blockerData.vaptBreakdown.critical > 0) {
        vaptSeverityHtml.push('<span style="background: #9C27B0; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin-right: 4px;">🟣 ' + blockerData.vaptBreakdown.critical + '</span>');
      }
      if (blockerData.vaptBreakdown.high > 0) {
        vaptSeverityHtml.push('<span style="background: #D32F2F; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin-right: 4px;">🔴 ' + blockerData.vaptBreakdown.high + '</span>');
      }
      if (blockerData.vaptBreakdown.medium > 0) {
        vaptSeverityHtml.push('<span style="background: #FF9800; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px;">🟠 ' + blockerData.vaptBreakdown.medium + '</span>');
      }
      if (vaptSeverityHtml.length > 0) {
        body += '<tr><td style="padding-left: 20px;">VAPT Severity:</td><td>' + vaptSeverityHtml.join(' ') + '</td></tr>';
      }

      body += '<tr><td style="padding-left: 20px;">Apps with blocker:</td><td>' + blockerData.vaptAppsWithBlockers + '</td></tr>';
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

    // VAPT Apps Detail (if any)
    if (blockerData.vaptBlocker > 0 && blockerData.vaptApps.length > 0) {
      body += '<div style="background: #FFF3E0; border-left: 4px solid #EF6C00; padding: 15px; margin: 20px 0;">';
      body += '<h3 style="color: #EF6C00; margin-top: 0;">🔒 VAPT BLOCKER DETAIL</h3>';
      body += '<p style="margin: 5px 0 15px 0; color: #666;">🎯 <strong>Target: 0 blocker di semua aplikasi!</strong></p>';
      body += '<table style="width: 100%; border-collapse: collapse;">';

      blockerData.vaptApps.forEach((app, index) => {
        const bgColor = index % 2 === 0 ? '#FFFAF0' : '#FFFFFF';
        body += '<tr style="background: ' + bgColor + ';">';
        body += '<td style="padding: 8px; border: 1px solid #E0E0E0; font-weight: bold;">' + app.aplikasi + '</td>';
        body += '<td style="padding: 8px; border: 1px solid #E0E0E0;">' + app.blocker + ' findings</td>';
        body += '<td style="padding: 8px; border: 1px solid #E0E0E0;">';
        // Show only non-zero severities
        if (app.critical > 0) {
          body += '<span style="background: #9C27B0; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin-right: 4px;">🟣 ' + app.critical + '</span> ';
        }
        if (app.high > 0) {
          body += '<span style="background: #D32F2F; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin-right: 4px;">🔴 ' + app.high + '</span> ';
        }
        if (app.medium > 0) {
          body += '<span style="background: #FF9800; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px;">🟠 ' + app.medium + '</span>';
        }
        body += '</td>';
        body += '</tr>';
      });

      body += '</table>';
      body += '</div>';
    }

    // QA Modules breakdown
    body += '<div style="margin: 20px 0;">';
    body += '<h3 style="color: #424242; border-bottom: 2px solid #1976D2; padding-bottom: 8px;">📦 QA Modules Breakdown</h3>';

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

      // Severity breakdown (show only non-zero)
      const moduleSeverityHtml = [];
      if ((module.critical || 0) > 0) {
        moduleSeverityHtml.push('<span style="background: #9C27B0; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin-right: 4px;">🟣 ' + module.critical + '</span>');
      }
      if ((module.high || 0) > 0) {
        moduleSeverityHtml.push('<span style="background: #D32F2F; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px; margin-right: 4px;">🔴 ' + module.high + '</span>');
      }
      if ((module.medium || 0) > 0) {
        moduleSeverityHtml.push('<span style="background: #FF9800; color: white; padding: 2px 6px; border-radius: 3px; font-size: 11px;">🟠 ' + module.medium + '</span>');
      }
      if (moduleSeverityHtml.length > 0) {
        body += '<tr><td><strong>Severity:</strong></td><td>' + moduleSeverityHtml.join(' ') + '</td></tr>';
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
    // Get Web App URL from Script Properties (fallback to latest)
    const scriptProps = PropertiesService.getScriptProperties();
    const dashboardWebAppUrl = scriptProps.getProperty('WEB_APP_URL') || LATEST_WEBAPP_URL;

    body += '<div style="background: #E8EAF6; border-left: 4px solid #3F51B5; padding: 15px; margin: 20px 0;">';
    body += '<h3 style="color: #3F51B5; margin-top: 0;">🔗 QUICK LINKS</h3>';
    body += '<p style="margin: 10px 0;"><a href="' + dashboardWebAppUrl + '" style="background: #1976D2; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">📊 Web Dashboard (QA & VAPT)</a></p>';
    body += '<p style="margin: 10px 0;"><a href="' + dashboardUrl + '" style="color: #1976D2; font-weight: bold;">📋 View Google Sheet Dashboard</a></p>';
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
    const dashboardVAPTUrl = dashboardUrl + '#gid=' + getDashboardVAPTGid_(ss);

    // Get Web App URL from Script Properties (fallback to latest)
    const scriptProps = PropertiesService.getScriptProperties();
    const dashboardWebAppUrl = scriptProps.getProperty('WEB_APP_URL') || LATEST_WEBAPP_URL;

    // Get detailed bug data from QATM BugReport sheets
    const bugDetails = getBugDetailsFromQATM_(blockerData.modules);

    // Build WhatsApp message (plain text, WhatsApp markdown)
    let message = '';

    // ══════════════════════════════════════════════════════════════════════
    // HEADER - Daily Bug Report Summary
    // ══════════════════════════════════════════════════════════════════════
    message += '📊 *DAILY BUG REPORT*';
    if (blockerData.projectName) {
      message += ' - ' + blockerData.projectName;
    }
    message += '\n📅 ' + blockerData.timestamp + '\n';
    message += '━━━━━━━━━━━━━\n\n';

    // Calculate total severity counts (same as email)
    const totalCritical = blockerData.modules.reduce((sum, m) => sum + (m.critical || 0), 0);
    const totalHigh = blockerData.modules.reduce((sum, m) => sum + (m.high || 0), 0);
    const totalMedium = blockerData.modules.reduce((sum, m) => sum + (m.medium || 0), 0);

    // Summary with bullet points
    message += '*SUMMARY*\n';

    // Count QA apps with blockers
    const qaAppsWithBlockers = blockerData.modules.filter(m => m.blocker > 0).length;
    message += '▬ QA Bugs: ' + blockerData.totalBlockers + ' (' + qaAppsWithBlockers + ' apps)\n';

    // QA Severity breakdown (hide if 0)
    const qaSeverityParts = [];
    if (totalCritical > 0) qaSeverityParts.push('Critical🟣 ' + totalCritical);
    if (totalHigh > 0) qaSeverityParts.push('High🔴 ' + totalHigh);
    if (totalMedium > 0) qaSeverityParts.push('Medium🟠 ' + totalMedium);
    if (qaSeverityParts.length > 0) {
      message += '  • Severity: ' + qaSeverityParts.join('  ') + '\n';
    }

    // VAPT Blocker summary
    if (blockerData.vaptBlocker > 0) {
      message += '▬ VAPT Blocker: ' + blockerData.vaptBlocker + ' (' + blockerData.vaptAppsWithBlockers + ' apps)\n';

      // VAPT Severity breakdown (hide if 0)
      const vaptSeverityParts = [];
      if (blockerData.vaptBreakdown.critical > 0) vaptSeverityParts.push('Critical🟣 ' + blockerData.vaptBreakdown.critical);
      if (blockerData.vaptBreakdown.high > 0) vaptSeverityParts.push('High🔴 ' + blockerData.vaptBreakdown.high);
      if (blockerData.vaptBreakdown.medium > 0) vaptSeverityParts.push('Medium🟠 ' + blockerData.vaptBreakdown.medium);
      if (vaptSeverityParts.length > 0) {
        message += '  • Severity: ' + vaptSeverityParts.join('  ') + '\n';
      }
    }
    message += '\n';

    // ══════════════════════════════════════════════════════════════════════
    // VAPT BLOCKER DETAIL (if any) - Per aplikasi
    // ══════════════════════════════════════════════════════════════════════
    if (blockerData.vaptBlocker > 0) {
      message += '*VAPT BLOCKER DETAIL*\n';
      message += '━━━━━━━━━━━━━\n\n';

      // List apps with blocker > 0
      blockerData.vaptApps.forEach(app => {
        message += '▬ ' + app.aplikasi + ': ' + app.blocker + '\n';

        // Show only non-zero severities
        const appSeverityParts = [];
        if (app.critical > 0) appSeverityParts.push('Critical🟣 ' + app.critical);
        if (app.high > 0) appSeverityParts.push('High🔴 ' + app.high);
        if (app.medium > 0) appSeverityParts.push('Medium🟠 ' + app.medium);
        if (appSeverityParts.length > 0) {
          message += '  • ' + appSeverityParts.join('  ') + '\n';
        }
      });

      message += '\n';
    }

    // ══════════════════════════════════════════════════════════════════════
    // PRODUCTION BUGS SECTION (if any) - TOP PRIORITY
    // ══════════════════════════════════════════════════════════════════════
    const hasProdBugs = blockerData.modules.some(m => m.prodBugs > 0);

    if (hasProdBugs) {
      message += '🚨 *PRODUCTION BUGS* 🚨\n';
      message += '━━━━━━━━━━━━━\n\n';

      blockerData.modules.forEach(module => {
        if (module.prodBugs > 0) {
          message += '*' + module.project + ' - ' + module.module + '*\n';
          message += '📌 ' + module.submodule + ': *' + module.prodBugs + '* PROD bug(s)\n\n';
        }
      });

      message += '⚠️ *URGENT - IMMEDIATE ACTION REQUIRED!*\n\n';
    }

    // ══════════════════════════════════════════════════════════════════════
    // QA BLOCKER BUGS - Breakdown per module (same format as VAPT)
    // ══════════════════════════════════════════════════════════════════════
    if (blockerData.totalBlockers > 0) {
      message += '*QA BLOCKER BUGS*\n';
      message += '━━━━━━━━━━━━━\n\n';

      blockerData.modules.forEach(module => {
        if (module.blocker > 0 || module.prodBugs > 0) {
          const moduleName = module.project + ' - ' + module.module + ' (' + module.submodule + ')';
          message += '▬ ' + moduleName + ': ' + module.blocker + '\n';

          // Severity breakdown (show only non-zero)
          const moduleSeverityParts = [];
          if ((module.critical || 0) > 0) moduleSeverityParts.push('Critical🟣 ' + module.critical);
          if ((module.high || 0) > 0) moduleSeverityParts.push('High🔴 ' + module.high);
          if ((module.medium || 0) > 0) moduleSeverityParts.push('Medium🟠 ' + module.medium);
          if (moduleSeverityParts.length > 0) {
            message += '  • ' + moduleSeverityParts.join('  ') + '\n';
          }
        }
      });
      message += '\n';
    }

    // ══════════════════════════════════════════════════════════════════════
    // FOOTER: Dashboard Links
    // ══════════════════════════════════════════════════════════════════════
    message += '━━━━━━━━━━━━━\n';
    message += '🔗 *Dashboard Links:*\n';
    message += '📊 Web Dashboard: ' + dashboardWebAppUrl + ' (QA & VAPT)\n';
    message += '📋 Sheet Overview: ' + dashboardOverviewUrl + '\n';
    message += '🐛 Sheet Bugs: ' + dashboardBugsUrl + '\n';
    message += '\n_Automated Daily Report - QA Dashboard_';

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
 * Fetch VAPT data for a specific project from external VAPT spreadsheet
 * @param {string} vaptSpreadsheetId - The spreadsheet ID containing VAPT data
 * @returns {Object} VAPT data { vaptBlocker, vaptAppsWithBlockers, vaptApps, vaptBreakdown }
 */
function fetchVAPTDataForProject_(vaptSpreadsheetId) {
  const result = {
    vaptBlocker: 0,
    vaptAppsWithBlockers: 0,
    vaptApps: [],
    vaptBreakdown: { critical: 0, high: 0, medium: 0 }
  };

  try {
    // Open external VAPT spreadsheet
    const vaptSS = SpreadsheetApp.openById(vaptSpreadsheetId);
    const vaptTab = vaptSS.getSheetByName('VAPT BGN - Helper');

    if (!vaptTab) {
      Logger.log('⚠️ VAPT BGN - Helper tab not found in spreadsheet: ' + vaptSpreadsheetId);
      return result;
    }

    // Fetch VAPT data from VAPT BGN - Helper tab (B4:I36)
    // Simple format: Aplikasi (B) | Critical (G) | High (H) | Medium (I)
    // Row 1-3: Headers
    // Row 4-36: Data
    const data = vaptTab.getRange('B4:I36').getValues();  // B to I = 8 columns

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      // B4:I36 indices: B=0, C=1, D=2, E=3, F=4, G=5, H=6, I=7
      const aplikasi = String(row[0] || '').trim();     // B (index 0 from B)
      const critical = Number(row[5]) || 0;              // G (index 5 from B)
      const high = Number(row[6]) || 0;                  // H (index 6 from B)
      const medium = Number(row[7]) || 0;                // I (index 7 from B)

      // Skip rows with no findings (all 0)
      if (critical === 0 && high === 0 && medium === 0) continue;

      // Skip rows with no aplikasi name
      if (!aplikasi) continue;

      // Total blocker per app (Medium + High + Critical)
      const blocker = medium + high + critical;

      // Accumulate total severity breakdown
      result.vaptBreakdown.medium += medium;
      result.vaptBreakdown.high += high;
      result.vaptBreakdown.critical += critical;

      // Only include apps with blocker > 0
      if (blocker > 0) {
        result.vaptAppsWithBlockers++;

        result.vaptApps.push({
          aplikasi: aplikasi,
          blocker: blocker,
          critical: critical,
          high: high,
          medium: medium
        });
      }
    }

    // Total blocker = sum of all severity blockers
    result.vaptBlocker = result.vaptBreakdown.critical + result.vaptBreakdown.high + result.vaptBreakdown.medium;

    Logger.log('VAPT data fetched: totalBlockers=' + result.vaptBlocker + ', apps=' + result.vaptAppsWithBlockers +
               ', crit=' + result.vaptBreakdown.critical + ', high=' + result.vaptBreakdown.high + ', med=' + result.vaptBreakdown.medium);

  } catch (error) {
    Logger.log('❌ Error fetching VAPT data from spreadsheet ' + vaptSpreadsheetId + ': ' + error.toString());
  }

  return result;
}

/**
 * Get GID for Dashboard VAPT tab
 */
function getDashboardVAPTGid_(ss) {
  const vaptSheet = ss.getSheetByName('VAPT');
  if (vaptSheet) {
    return vaptSheet.getSheetId();
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

/**
 * Set Web App URL in Script Properties
 * Run this once after deploying the web app to set the URL
 *
 * Usage:
 * 1. Deploy web app and get URL
 * 2. Run: setWebAppUrl('https://script.google.com/.../exec')
 */
function setWebAppUrl(url) {
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty('WEB_APP_URL', url);
  Logger.log('✅ Web App URL set to: ' + url);
  Logger.log('This URL will be used in WhatsApp, GChat, and Email notifications');
}

/**
 * Get current Web App URL from Script Properties
 */
function getWebAppUrl() {
  const scriptProps = PropertiesService.getScriptProperties();
  const url = scriptProps.getProperty('WEB_APP_URL');
  if (url) {
    Logger.log('Current Web App URL: ' + url);
  } else {
    Logger.log('⚠️ Web App URL not set. Run setWebAppUrl() first.');
  }
  return url;
}

/**
 * AUTO SET Web App URL - Set latest deployment URL to Script Properties
 * Run this function after deploying new web app version
 * Optional: Notifications will auto-fallback to LATEST_WEBAPP_URL if not set
 */
function autoSetWebAppUrl() {
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.setProperty('WEB_APP_URL', LATEST_WEBAPP_URL);
  Logger.log('Web App URL set to: ' + LATEST_WEBAPP_URL);
  Logger.log('This URL will be used in WhatsApp, GChat, and Email notifications');
  return LATEST_WEBAPP_URL;
}

/**
 * MENU: WhatsApp - Get Groups
 * Get list of WhatsApp groups from Fonnte API
 * Shows group IDs in a dialog for easy copy-paste to Config sheet
 */
function menuTestGetGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');

  if (!cfg) {
    SpreadsheetApp.getUi().alert('❌ Config tab not found');
    return;
  }

  // Get Fonnte Token from Config row 4, col T (column 20)
  const fontteToken = String(cfg.getRange(4, 20).getValue()).trim();

  if (!fontteToken) {
    SpreadsheetApp.getUi().alert(
      '❌ Fonnte Token Missing\n\n' +
      'Please set Fonnte Token in Config sheet:\n' +
      'Row 4, Column T (Fonnte Token)'
    );
    return;
  }

  try {
    // Call Fonnte API to get device info (includes groups)
    const url = 'https://api.fonnte.com/get-devices';

    const options = {
      method: 'post',
      headers: {
        'Authorization': fontteToken
      },
      muteHttpExceptions: true
    };

    Logger.log('Fetching WhatsApp groups from Fonnte API...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('Response Code: ' + responseCode);
    Logger.log('Response: ' + responseText);

    if (responseCode !== 200) {
      SpreadsheetApp.getUi().alert(
        '❌ API Error\n\n' +
        'Response Code: ' + responseCode + '\n' +
        'Response: ' + responseText
      );
      return;
    }

    const data = JSON.parse(responseText);

    // Format response for display
    let message = '📱 WHATSAPP DEVICES & GROUPS\n\n';

    if (data.device && data.device.length > 0) {
      data.device.forEach((device, index) => {
        message += '════════════════════════════\n';
        message += 'DEVICE ' + (index + 1) + ':\n';
        message += '• Name: ' + (device.name || 'N/A') + '\n';
        message += '• Number: ' + (device.device || 'N/A') + '\n';
        message += '• Status: ' + (device.status || 'N/A') + '\n\n';

        // Get groups for this device
        if (device.groups && device.groups.length > 0) {
          message += 'GROUPS (' + device.groups.length + '):\n';
          device.groups.forEach(group => {
            message += '  • ' + group.name + '\n';
            message += '    ID: ' + group.id + '\n';
          });
        } else {
          message += 'GROUPS: None found\n';
        }
        message += '\n';
      });

      message += '════════════════════════════\n\n';
      message += '💡 TIP:\n';
      message += 'Copy Group ID (format: 120363xxx@g.us)\n';
      message += 'Paste to Config sheet Row 4, Col S';

    } else {
      message = '❌ No devices found\n\n' +
                'Raw Response:\n' + responseText;
    }

    SpreadsheetApp.getUi().alert(message);

  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    SpreadsheetApp.getUi().alert(
      '❌ Error getting WhatsApp groups\n\n' +
      'Error: ' + e.message + '\n\n' +
      'Check Execution log for details'
    );
  }
}

/**
 * MENU: WhatsApp - Send Test
 * Send test message to configured WhatsApp group
 */
function menuTestSendToGroup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');

  if (!cfg) {
    SpreadsheetApp.getUi().alert('❌ Config tab not found');
    return;
  }

  // Get WhatsApp config from Config row 4
  const whatsappGroupId = String(cfg.getRange(4, 19).getValue()).trim(); // S4 = Group ID
  const fontteToken = String(cfg.getRange(4, 20).getValue()).trim();      // T4 = Token
  const whatsappEnabled = cfg.getRange(4, 21).getValue() === true;        // U4 = Enable

  // Validate config
  if (!fontteToken) {
    SpreadsheetApp.getUi().alert(
      '❌ Fonnte Token Missing\n\n' +
      'Please set Fonnte Token in Config sheet:\n' +
      'Row 4, Column T (Fonnte Token)'
    );
    return;
  }

  if (!whatsappGroupId || !whatsappGroupId.includes('@g.us')) {
    SpreadsheetApp.getUi().alert(
      '❌ WhatsApp Group ID Missing or Invalid\n\n' +
      'Please set Group ID in Config sheet:\n' +
      'Row 4, Column S (WhatsApp Group ID)\n\n' +
      'Format: 120363xxx@g.us\n\n' +
      'Use "WhatsApp: Get Groups" menu to find your group ID'
    );
    return;
  }

  if (!whatsappEnabled) {
    const response = SpreadsheetApp.getUi().alert(
      '⚠️ WhatsApp Disabled\n\n' +
      'WhatsApp notifications are currently DISABLED.\n' +
      'Row 4, Column U = unchecked\n\n' +
      'Do you want to send test message anyway?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (response !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }

  try {
    // Build test message
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const message =
      '🧪 *TEST MESSAGE*\n' +
      '📅 ' + timestamp + '\n' +
      '━━━━━━━━━━━━━\n\n' +
      '✅ WhatsApp notification is working!\n\n' +
      'This is a test message from QA Portfolio Dashboard.\n\n' +
      '📊 Dashboard: ' + ss.getUrl() + '\n\n' +
      '_Sent via Fonnte API_';

    // Send via Fonnte API
    const url = 'https://api.fonnte.com/send';

    const payload = {
      target: whatsappGroupId,
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

    Logger.log('Sending test message to: ' + whatsappGroupId);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('Response Code: ' + responseCode);
    Logger.log('Response: ' + responseText);

    if (responseCode === 200) {
      const data = JSON.parse(responseText);

      if (data.status) {
        SpreadsheetApp.getUi().alert(
          '✅ Test Message Sent!\n\n' +
          'WhatsApp notification sent successfully to:\n' +
          whatsappGroupId + '\n\n' +
          'Check your WhatsApp group for the test message.'
        );
      } else {
        SpreadsheetApp.getUi().alert(
          '❌ Send Failed\n\n' +
          'Reason: ' + (data.reason || 'Unknown error') + '\n\n' +
          'Response: ' + responseText
        );
      }
    } else {
      SpreadsheetApp.getUi().alert(
        '❌ API Error\n\n' +
        'Response Code: ' + responseCode + '\n' +
        'Response: ' + responseText
      );
    }

  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    SpreadsheetApp.getUi().alert(
      '❌ Error sending test message\n\n' +
      'Error: ' + e.message + '\n\n' +
      'Check Execution log for details'
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════
// TEST EXECUTION NOTIFICATION FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Send test execution notification (Pass Rate summary per module)
 * Grouped by project, sent to configured channels (WhatsApp, Email, Google Chat)
 * 
 * Thresholds (from Dashboard conditional formatting):
 * - Pass Rate (WEB/API): ≥80% 🟢 | 50-79% 🟡 | <50% 🔴
 * - Smoke Pass Rate: ≥80% 🟢 | 50-79% 🟡 | <50% 🔴
 * - Smoke Exec Rate: ≥70% 🟢 | 40-69% 🟡 | <40% 🔴
 */
function sendTestExecutionNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');
  
  if (!cfg) {
    Logger.log('Config tab not found');
    return;
  }
  
  const overview = ss.getSheetByName('Overview');
  if (!overview) {
    Logger.log('Overview tab not found');
    return;
  }
  
  // Get test execution data from Overview tab
  const testData = getTestExecutionData_(overview, cfg);
  
  if (testData.length === 0) {
    Logger.log('No test execution data found');
    SpreadsheetApp.getUi().alert(
      '✅ No Test Data',
      'Tidak ada data test execution untuk dikirim.\n\n' +
      'Pastikan sudah ada module aktif dengan data test.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  Logger.log('Found ' + testData.length + ' projects with test data');
  
  // Get shared Fonnte token (row 4, col T)
  const fontteToken = cfg ? String(cfg.getRange(4, 20).getValue()).trim() : '';
  
  // Group data by project and send notifications
  const projectGroups = {};
  testData.forEach(data => {
    if (!projectGroups[data.projectName]) {
      projectGroups[data.projectName] = [];
    }
    projectGroups[data.projectName].push(data);
  });
  
  let totalSent = 0;
  
  Object.keys(projectGroups).forEach(projectName => {
    const projectModules = projectGroups[projectName];
    
    // Aggregate project data
    const projectData = {
      projectName: projectName,
      timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'),
      modules: projectModules
    };
    
    // Send to each module's configured channels
    projectModules.forEach(moduleData => {
      let sent = 0;
      
      // WhatsApp
      if (moduleData.whatsappEnabled && moduleData.whatsappGroupId && fontteToken) {
        const message = buildWhatsAppTestExecution_(projectData);
        const success = sendWhatsApp_(fontteToken, moduleData.whatsappGroupId, message);
        if (success) sent++;
      }
      
      // Email
      if (moduleData.emailEnabled && moduleData.emailRecipients) {
        const subject = '[' + projectName + '] Daily Test Execution Summary';
        const htmlBody = buildEmailTestExecution_(projectData, ss);
        const success = sendEmail_(moduleData.emailRecipients, subject, htmlBody);
        if (success) sent++;
      }
      
      // Google Chat
      if (moduleData.chatEnabled && moduleData.chatWebhook) {
        const payload = buildGoogleChatTestExecution_(projectData, ss);
        const success = sendGoogleChat_(moduleData.chatWebhook, payload);
        if (success) sent++;
      }
      
      if (sent > 0) {
        Logger.log('✅ Sent test execution notification for: ' + moduleData.modul + ' (' + sent + ' channels)');
        totalSent++;
      }
    });
  });
  
  SpreadsheetApp.getUi().alert(
    '✅ Test Execution Notifications Sent!',
    'Successfully sent test execution summary to ' + totalSent + ' module(s).\n\n' +
    'Projects: ' + Object.keys(projectGroups).join(', '),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Get test execution data from Overview tab
 * Returns array of modules with pass rate data
 */
function getTestExecutionData_(overview, cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const overviewData = overview.getDataRange().getValues();
  
  // Build module config map from Config tab
  const moduleMap = {};
  if (cfg) {
    const cfgData = cfg.getDataRange().getValues();
    for (let i = 3; i < cfgData.length; i++) {
      const active = cfgData[i][0] === true;
      const project = String(cfgData[i][2]).trim();
      const modul = String(cfgData[i][3]).trim();
      const chatWebhook = String(cfgData[i][11]).trim();
      const chatEnabled = cfgData[i][13] === true;
      const emailRecipients = String(cfgData[i][14]).trim();
      const emailEnabled = cfgData[i][15] === true;
      const whatsappGroupId = String(cfgData[i][18]).trim();
      const whatsappEnabled = cfgData[i][20] === true;
      
      if (active && modul) {
        const key = project + '|' + modul;
        moduleMap[key] = {
          project,
          modul,
          chatWebhook: chatWebhook && chatWebhook.includes('chat.googleapis.com') ? chatWebhook : null,
          chatEnabled,
          emailRecipients: emailRecipients || null,
          emailEnabled,
          whatsappGroupId: whatsappGroupId || null,
          whatsappEnabled
        };
      }
    }
  }
  
  // Read test execution data from Overview tab
  // Row structure: 1=webapp link, 2=timestamp, 3=title, 4=group headers, 5=column headers, 6+=data
  const testData = [];
  
  for (let i = 5; i < overviewData.length; i++) {
    const row = overviewData[i];
    const project = String(row[0]).trim();
    const modul = String(row[1]).trim();
    
    // Skip empty rows, TOTAL row, or invalid data
    if (!modul || modul.toUpperCase().includes('TOTAL') || modul.toUpperCase().includes('AVERAGE')) {
      continue;
    }
    
    const key = project + '|' + modul;
    const config = moduleMap[key];
    
    // Skip if module not in config or notification not enabled
    if (!config || (!config.whatsappEnabled && !config.emailEnabled && !config.chatEnabled)) {
      continue;
    }
    
    // Extract test execution data
    // Columns: WEB (I-M: 8-12), API (Q-U: 16-20), Smoke WEB (N-P: 13-15), Smoke API (V-X: 21-23)
    const webTotal = Number(row[8]) || 0;
    const webPass = Number(row[9]) || 0;
    const webPassRate = Number(row[12]) || 0;
    
    const apiTotal = Number(row[16]) || 0;
    const apiPass = Number(row[17]) || 0;
    const apiPassRate = Number(row[20]) || 0;
    
    const smokeWebTotal = Number(row[13]) || 0;
    const smokeWebPassRate = Number(row[14]) || 0;
    
    const smokeApiTotal = Number(row[22]) || 0;
    const smokeApiPassRate = Number(row[23]) || 0;
    
    // Only include modules with test data
    if (webTotal === 0 && apiTotal === 0) {
      continue;
    }
    
    testData.push({
      projectName: project,
      modul: modul,
      webTotal,
      webPass,
      webPassRate,
      apiTotal,
      apiPass,
      apiPassRate,
      smokeWebTotal,
      smokeWebPassRate,
      smokeApiTotal,
      smokeApiPassRate,
      chatWebhook: config.chatWebhook,
      chatEnabled: config.chatEnabled,
      emailRecipients: config.emailRecipients,
      emailEnabled: config.emailEnabled,
      whatsappGroupId: config.whatsappGroupId,
      whatsappEnabled: config.whatsappEnabled
    });
  }
  
  return testData;
}

/**
 * Build WhatsApp message for test execution notification
 */
function buildWhatsAppTestExecution_(projectData) {
  let message = '';

  // Header
  message += '📊 *TEST EXECUTION REPORT*';
  if (projectData.projectName) {
    message += ' - ' + projectData.projectName;
  }
  message += '\n📅 ' + projectData.timestamp + '\n';
  message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

  // Calculate summary stats
  const totalWeb = projectData.modules.reduce((sum, m) => sum + m.webTotal, 0);
  const totalWebPass = projectData.modules.reduce((sum, m) => sum + m.webPass, 0);
  const totalApi = projectData.modules.reduce((sum, m) => sum + m.apiTotal, 0);
  const totalApiPass = projectData.modules.reduce((sum, m) => sum + m.apiPass, 0);

  const modulesBelow80 = projectData.modules.filter(m => {
    const avgRate = m.webTotal > 0 && m.apiTotal > 0
      ? (m.webPassRate + m.apiPassRate) / 2
      : m.webTotal > 0 ? m.webPassRate : m.apiPassRate;
    return avgRate < 0.8;
  });

  // SUMMARY (like blocker notif - summary first)
  message += '📈 *SUMMARY*\n';
  message += '✅ Total Modules: ' + projectData.modules.length + '\n';

  if (modulesBelow80.length > 0) {
    message += '⚠️ Modules Below 80%: ' + modulesBelow80.length;
    message += ' (' + modulesBelow80.map(m => m.modul).join(', ') + ')\n';
  }

  if (totalWeb > 0) {
    const avgWebRate = totalWebPass / totalWeb;
    message += '🎯 Avg WEB Pass Rate: ' + (avgWebRate * 100).toFixed(1) + '%\n';
  }
  if (totalApi > 0) {
    const avgApiRate = totalApiPass / totalApi;
    message += '🎯 Avg API Pass Rate: ' + (avgApiRate * 100).toFixed(1) + '%\n';
  }

  // MODULE DETAILS (per modul, no tree structure)
  message += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  message += '📋 *MODULE DETAILS*\n\n';

  projectData.modules.forEach((mod, idx) => {
    // Module header with status indicator
    const avgPassRate = mod.webTotal > 0 && mod.apiTotal > 0
      ? (mod.webPassRate + mod.apiPassRate) / 2
      : mod.webTotal > 0 ? mod.webPassRate : mod.apiPassRate;

    const statusIcon = avgPassRate >= 0.8 ? '🟢' : avgPassRate >= 0.5 ? '⚠️' : '🔴';
    message += statusIcon + ' *' + mod.modul + '*\n';

    // WEB Test (simple format, no tree)
    if (mod.webTotal > 0) {
      const webIcon = mod.webPassRate >= 0.8 ? '✅' : mod.webPassRate >= 0.5 ? '⚠️' : '❌';
      message += 'WEB: ' + (mod.webPassRate * 100).toFixed(1) + '% ';
      message += '(' + mod.webPass + '/' + mod.webTotal + ' pass)';
      if (mod.webPassRate < 0.8) {
        message += ' ' + webIcon + ' BELOW TARGET';
      }
      message += '\n';

      // Smoke WEB (no tree structure)
      if (mod.smokeWebTotal > 0) {
        message += 'Smoke WEB: ' + (mod.smokeWebPassRate * 100).toFixed(1) + '%\n';
      }
    }

    // API Test (simple format, no tree)
    if (mod.apiTotal > 0) {
      const apiIcon = mod.apiPassRate >= 0.8 ? '✅' : mod.apiPassRate >= 0.5 ? '⚠️' : '❌';
      message += 'API: ' + (mod.apiPassRate * 100).toFixed(1) + '% ';
      message += '(' + mod.apiPass + '/' + mod.apiTotal + ' pass)';
      if (mod.apiPassRate < 0.8) {
        message += ' ' + apiIcon + ' BELOW TARGET';
      }
      message += '\n';

      // Smoke API (no tree structure)
      if (mod.smokeApiTotal > 0) {
        message += 'Smoke API: ' + (mod.smokeApiPassRate * 100).toFixed(1) + '%\n';
      }
    }

    message += '\n';
  });

  // Dashboard link
  message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  const overviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Overview');
  const overviewGid = overviewSheet ? overviewSheet.getSheetId() : 0;
  const dashboardUrl = 'https://docs.google.com/spreadsheets/d/' +
    SpreadsheetApp.getActiveSpreadsheet().getId() + '/edit#gid=' + overviewGid;

  message += '🔗 View Dashboard:\n' + dashboardUrl;

  return message;
}

/**
 * Build HTML email body for test execution notification
 */
function buildEmailTestExecution_(projectData, ss) {
  const overviewSheet = ss.getSheetByName('Overview');
  const overviewGid = overviewSheet ? overviewSheet.getSheetId() : 0;
  const dashboardUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + overviewGid;
  
  // Calculate project stats
  const totalWeb = projectData.modules.reduce((sum, m) => sum + m.webTotal, 0);
  const totalWebPass = projectData.modules.reduce((sum, m) => sum + m.webPass, 0);
  const totalApi = projectData.modules.reduce((sum, m) => sum + m.apiTotal, 0);
  const totalApiPass = projectData.modules.reduce((sum, m) => sum + m.apiPass, 0);
  
  const avgWebRate = totalWeb > 0 ? totalWebPass / totalWeb : 0;
  const avgApiRate = totalApi > 0 ? totalApiPass / totalApi : 0;
  
  const modulesBelow80 = projectData.modules.filter(m => {
    const avgRate = m.webTotal > 0 && m.apiTotal > 0 
      ? (m.webPassRate + m.apiPassRate) / 2 
      : m.webTotal > 0 ? m.webPassRate : m.apiPassRate;
    return avgRate < 0.8;
  });
  
  const statusColor = modulesBelow80.length === 0 ? '#4CAF50' : '#FF9800';
  const statusIcon = modulesBelow80.length === 0 ? '✅' : '⚠️';
  const statusMessage = modulesBelow80.length === 0 
    ? 'All modules meeting target! Great job team!' 
    : modulesBelow80.length + ' module(s) below 80% pass rate — Need attention.';
  
  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
    .container { max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; overflow: hidden; }
    .header { background: ${statusColor}; color: white; padding: 20px; text-align: center; }
    .header h1 { margin: 0; font-size: 24px; }
    .header p { margin: 10px 0 0; opacity: 0.9; }
    .summary { background: #f9f9f9; padding: 15px 20px; border-left: 4px solid ${statusColor}; margin: 20px; }
    .summary p { margin: 5px 0; font-size: 14px; }
    .module { margin: 20px; padding: 15px; border: 1px solid #ddd; border-radius: 4px; }
    .module-header { font-weight: bold; font-size: 16px; margin-bottom: 10px; color: #333; }
    .test-line { padding: 4px 0; }
    .rate-good { color: #2E7D32; font-weight: bold; }
    .rate-warning { color: #F57C00; font-weight: bold; }
    .rate-critical { color: #C62828; font-weight: bold; }
    .footer { text-align: center; padding: 20px; background: #f9f9f9; }
    .button { display: inline-block; padding: 12px 24px; background: #1976D2; color: white; text-decoration: none; border-radius: 4px; margin-top: 10px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>${statusIcon} Test Execution Report</h1>
      <p>${projectData.projectName} | ${projectData.timestamp}</p>
    </div>
    
    <div class="summary">
      <p><strong>📊 Project Status:</strong> ${statusMessage}</p>
      <p><strong>✅ Total Modules:</strong> ${projectData.modules.length}</p>
      ${modulesBelow80.length > 0 ? '<p><strong>⚠️ Modules Below 80%:</strong> ' + modulesBelow80.length + ' (' + modulesBelow80.map(m => m.modul).join(', ') + ')</p>' : ''}
      ${totalWeb > 0 ? '<p><strong>🎯 Avg WEB Pass Rate:</strong> ' + (avgWebRate * 100).toFixed(1) + '%</p>' : ''}
      ${totalApi > 0 ? '<p><strong>🎯 Avg API Pass Rate:</strong> ' + (avgApiRate * 100).toFixed(1) + '%</p>' : ''}
    </div>
    
    <h2 style="margin: 20px; color: #333;">📈 Module Details</h2>
`;
  
  projectData.modules.forEach(mod => {
    const avgPassRate = mod.webTotal > 0 && mod.apiTotal > 0 
      ? (mod.webPassRate + mod.apiPassRate) / 2 
      : mod.webTotal > 0 ? mod.webPassRate : mod.apiPassRate;
    
    const statusIcon = avgPassRate >= 0.8 ? '🟢' : avgPassRate >= 0.5 ? '⚠️' : '🔴';
    
    html += `
    <div class="module">
      <div class="module-header">${statusIcon} ${mod.modul}</div>
`;
    
    // WEB Test
    if (mod.webTotal > 0) {
      const webClass = mod.webPassRate >= 0.8 ? 'rate-good' : mod.webPassRate >= 0.5 ? 'rate-warning' : 'rate-critical';
      html += `
      <div class="test-line">
        <strong>WEB:</strong> <span class="${webClass}">${(mod.webPassRate * 100).toFixed(1)}%</span> 
        (${mod.webPass}/${mod.webTotal} pass)
        ${mod.webPassRate < 0.8 ? ' <span style="color: #F57C00;">⚠️ BELOW TARGET</span>' : ''}
      </div>
`;
      if (mod.smokeWebTotal > 0) {
        html += `      <div class="test-line" style="margin-left: 20px;">└─ Smoke: ${(mod.smokeWebPassRate * 100).toFixed(1)}%</div>\n`;
      }
    }
    
    // API Test
    if (mod.apiTotal > 0) {
      const apiClass = mod.apiPassRate >= 0.8 ? 'rate-good' : mod.apiPassRate >= 0.5 ? 'rate-warning' : 'rate-critical';
      html += `
      <div class="test-line">
        <strong>API:</strong> <span class="${apiClass}">${(mod.apiPassRate * 100).toFixed(1)}%</span> 
        (${mod.apiPass}/${mod.apiTotal} pass)
        ${mod.apiPassRate < 0.8 ? ' <span style="color: #F57C00;">⚠️ BELOW TARGET</span>' : ''}
      </div>
`;
      if (mod.smokeApiTotal > 0) {
        html += `      <div class="test-line" style="margin-left: 20px;">└─ Smoke: ${(mod.smokeApiPassRate * 100).toFixed(1)}%</div>\n`;
      }
    }
    
    html += `
    </div>
`;
  });
  
  html += `
    <div class="footer">
      <p>View full dashboard for detailed test results and trends</p>
      <a href="${dashboardUrl}" class="button">📊 Open Dashboard</a>
      <p style="margin-top: 15px; font-size: 12px; color: #757575;">
        🤖 Automated Test Execution Report<br>
        Target: WEB/API ≥80%, Smoke ≥80%
      </p>
    </div>
  </div>
</body>
</html>
`;
  
  return html;
}

/**
 * Build Google Chat card for test execution notification
 */
function buildGoogleChatTestExecution_(projectData, ss) {
  const overviewSheet = ss.getSheetByName('Overview');
  const overviewGid = overviewSheet ? overviewSheet.getSheetId() : 0;
  const dashboardUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/edit#gid=' + overviewGid;
  
  const widgets = [];
  
  // Calculate project stats
  const totalWeb = projectData.modules.reduce((sum, m) => sum + m.webTotal, 0);
  const totalWebPass = projectData.modules.reduce((sum, m) => sum + m.webPass, 0);
  const totalApi = projectData.modules.reduce((sum, m) => sum + m.apiTotal, 0);
  const totalApiPass = projectData.modules.reduce((sum, m) => sum + m.apiPass, 0);
  
  const avgWebRate = totalWeb > 0 ? totalWebPass / totalWeb : 0;
  const avgApiRate = totalApi > 0 ? totalApiPass / totalApi : 0;
  
  const modulesBelow80 = projectData.modules.filter(m => {
    const avgRate = m.webTotal > 0 && m.apiTotal > 0 
      ? (m.webPassRate + m.apiPassRate) / 2 
      : m.webTotal > 0 ? m.webPassRate : m.apiPassRate;
    return avgRate < 0.8;
  });
  
  const statusIcon = modulesBelow80.length === 0 ? '✅' : '⚠️';
  const statusMessage = modulesBelow80.length === 0 
    ? '<b>All modules meeting target!</b> 🎉 Great job team!' 
    : '<b>' + modulesBelow80.length + ' module(s) below 80% pass rate</b> — Need attention.';
  
  // Header summary
  widgets.push({
    decoratedText: {
      topLabel: statusIcon + ' TEST EXECUTION REPORT',
      text: statusMessage,
      bottomLabel: projectData.timestamp
    }
  });
  
  widgets.push({ divider: {} });
  
  // Project summary
  let summaryText = '<b>📊 Project Status</b><br>';
  summaryText += '✅ Total Modules: ' + projectData.modules.length + '<br>';
  if (modulesBelow80.length > 0) {
    summaryText += '⚠️ Modules Below 80%: ' + modulesBelow80.length + '<br>';
    summaryText += '   (' + modulesBelow80.map(m => m.modul).join(', ') + ')<br>';
  }
  if (totalWeb > 0) {
    summaryText += '🎯 Avg WEB Pass Rate: ' + (avgWebRate * 100).toFixed(1) + '%<br>';
  }
  if (totalApi > 0) {
    summaryText += '🎯 Avg API Pass Rate: ' + (avgApiRate * 100).toFixed(1) + '%';
  }
  
  widgets.push({
    textParagraph: {
      text: summaryText
    }
  });
  
  widgets.push({ divider: {} });
  
  // Module details
  widgets.push({
    textParagraph: {
      text: '<b>📈 Module Details</b>'
    }
  });
  
  projectData.modules.forEach(mod => {
    const avgPassRate = mod.webTotal > 0 && mod.apiTotal > 0 
      ? (mod.webPassRate + mod.apiPassRate) / 2 
      : mod.webTotal > 0 ? mod.webPassRate : mod.apiPassRate;
    
    const statusIcon = avgPassRate >= 0.8 ? '🟢' : avgPassRate >= 0.5 ? '⚠️' : '🔴';
    
    let moduleText = '<b>' + statusIcon + ' ' + mod.modul + '</b><br>';
    
    if (mod.webTotal > 0) {
      moduleText += 'WEB: ' + (mod.webPassRate * 100).toFixed(1) + '% (' + mod.webPass + '/' + mod.webTotal + ')';
      if (mod.webPassRate < 0.8) {
        moduleText += ' ⚠️ BELOW TARGET';
      }
      moduleText += '<br>';
      if (mod.smokeWebTotal > 0) {
        moduleText += '  └─ Smoke: ' + (mod.smokeWebPassRate * 100).toFixed(1) + '%<br>';
      }
    }
    
    if (mod.apiTotal > 0) {
      moduleText += 'API: ' + (mod.apiPassRate * 100).toFixed(1) + '% (' + mod.apiPass + '/' + mod.apiTotal + ')';
      if (mod.apiPassRate < 0.8) {
        moduleText += ' ⚠️ BELOW TARGET';
      }
      moduleText += '<br>';
      if (mod.smokeApiTotal > 0) {
        moduleText += '  └─ Smoke: ' + (mod.smokeApiPassRate * 100).toFixed(1) + '%';
      }
    }
    
    widgets.push({
      textParagraph: {
        text: moduleText
      }
    });
  });
  
  widgets.push({ divider: {} });
  
  // Dashboard button
  widgets.push({
    buttonList: {
      buttons: [{
        text: '📊 Open Dashboard',
        onClick: { openLink: { url: dashboardUrl } }
      }]
    }
  });
  
  return {
    cardsV2: [{
      cardId: 'test-execution-notification',
      card: {
        header: {
          title: projectData.projectName + ' - Test Execution',
          subtitle: 'Daily Test Execution Summary',
          imageUrl: 'https://www.gstatic.com/images/branding/product/1x/keep_48dp.png',
          imageType: 'CIRCLE'
        },
        sections: [{
          widgets: widgets
        }]
      }
    }]
  };
}

// ═══════════════════════════════════════════════════════════════════════
// TEST EXECUTION NOTIFICATION TRIGGER MANAGEMENT
// ═══════════════════════════════════════════════════════════════════════

/**
 * Setup daily test execution notification trigger
 * Uses same schedule as blocker notification (from Config tab col M)
 */
function setupTestExecutionNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');
  const ui = SpreadsheetApp.getUi();
  
  if (!cfg) {
    ui.alert('❌ Config tab not found!');
    return;
  }
  
  const instructionsMsg =
    '🔔 SETUP TEST EXECUTION NOTIFICATION\n\n' +
    'Notification akan menggunakan konfigurasi yang sama dengan Blocker Notification:\n\n' +
    '📱 GOOGLE CHAT (kolom L-N) - Per Project\n' +
    '📧 EMAIL (kolom O-P) - Per Project\n' +
    '📲 WHATSAPP (kolom S-U) - Per Project\n' +
    '⏰ SCHEDULE (kolom M) - Same as blocker notification\n\n' +
    '💡 Test Execution Notification akan dikirim pada jam yang sama dengan Blocker Notification.\n\n' +
    'Lanjutkan setup trigger?';
  
  const response = ui.alert('Setup Test Execution Notification', instructionsMsg, ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.NO) {
    return;
  }
  
  // Check if ANY module has notifications enabled
  const cfgData = cfg.getDataRange().getValues();
  let hasAnyEnabled = false;
  let scheduleStr = '17'; // Default schedule (5 PM)
  
  for (let i = 3; i < cfgData.length; i++) {
    const chatEnabled = cfgData[i][13] === true; // Col N
    const emailEnabled = cfgData[i][15] === true; // Col P
    const whatsappEnabled = cfgData[i][20] === true; // Col U
    const schedule = String(cfgData[i][12]).trim(); // Col M
    
    if (chatEnabled || emailEnabled || whatsappEnabled) {
      hasAnyEnabled = true;
      scheduleStr = schedule || '17';
      break;
    }
  }
  
  if (!hasAnyEnabled) {
    const openConfig = ui.alert(
      '⚠️ No Notifications Enabled',
      'Tidak ada notification yang aktif.\n\n' +
      'Aktifkan minimal 1 channel di Config tab:\n' +
      '• Kolom N = ☑ Enable Google Chat\n' +
      '• Kolom P = ☑ Enable Email\n' +
      '• Kolom U = ☑ Enable WhatsApp\n\n' +
      'Buka Config tab sekarang?',
      ui.ButtonSet.YES_NO
    );
    
    if (openConfig === ui.Button.YES) {
      ss.setActiveSheet(cfg);
      ss.setActiveRange(cfg.getRange('N4'));
    }
    return;
  }
  
  // Parse schedule
  const parsedSchedule = parseSchedule_(scheduleStr);
  
  if (!parsedSchedule.success) {
    ui.alert(
      '❌ Invalid Schedule Format',
      'Schedule format tidak valid: "' + scheduleStr + '"\n\n' +
      '✅ FORMAT VALID:\n' +
      '• Single: 17 atau "17"\n' +
      '• Multiple: 7,12,18\n' +
      '• Interval: 4h (support: 1h, 2h, 4h, 6h, 8h, 12h)\n\n' +
      'Error: ' + parsedSchedule.error,
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendTestExecutionNotification') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  // Create new triggers based on schedule type
  let triggerCount = 0;
  
  if (parsedSchedule.type === 'interval') {
    // Create hourly trigger
    ScriptApp.newTrigger('sendTestExecutionNotification')
      .timeBased()
      .everyHours(parsedSchedule.intervalHours)
      .create();
    triggerCount = 1;
    
    Logger.log('✅ Created test execution notification trigger: Every ' + parsedSchedule.intervalHours + 'h');
  } else {
    // Create daily trigger for each hour
    parsedSchedule.hours.forEach(hour => {
      ScriptApp.newTrigger('sendTestExecutionNotification')
        .timeBased()
        .atHour(hour)
        .everyDays(1)
        .create();
      triggerCount++;
      
      Logger.log('✅ Created test execution notification trigger: Daily at ' + hour + ':00');
    });
  }
  
  ui.alert(
    '✅ Test Execution Notification Setup Complete!',
    'Daily test execution notification trigger berhasil dibuat.\n\n' +
    'Triggers created: ' + triggerCount + '\n' +
    'Schedule: ' + scheduleStr + '\n\n' +
    'Notification akan dikirim sesuai jadwal ke channels yang dikonfigurasi.',
    ui.ButtonSet.OK
  );
}

/**
 * Remove test execution notification trigger
 */
function removeTestExecutionNotification() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendTestExecutionNotification') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  
  if (removed > 0) {
    SpreadsheetApp.getUi().alert(
      '✅ Trigger Removed',
      'Test execution notification trigger telah dihapus.\n\n' +
      'Triggers removed: ' + removed,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log('Removed ' + removed + ' test execution notification trigger(s)');
  } else {
    SpreadsheetApp.getUi().alert(
      'ℹ️ No Trigger Found',
      'Tidak ada test execution notification trigger yang aktif.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
