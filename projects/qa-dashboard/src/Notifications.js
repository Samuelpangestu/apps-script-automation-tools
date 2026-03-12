/**
 * notifications.js
 * Send blocker & PROD BUGS notifications via Google Chat & Email
 */

/**
 * Send blocker notification to Google Chat and Email
 * Run manually for testing or via daily trigger
 *
 * NEW: Supports PER-MODULE webhooks and enable flags from Config
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

  // Send notifications PER MODULE using per-module config
  let chatSentCount = 0;
  let emailSentCount = 0;
  let skippedCount = 0;

  blockerData.modules.forEach(module => {
    // Skip if both notifications disabled for this module
    if (!module.chatEnabled && !module.emailEnabled) {
      Logger.log('Module ' + module.project + '-' + module.module + ' has notifications disabled, skipping');
      skippedCount++;
      return;
    }

    // Create per-module blocker data
    const perModuleData = {
      modules: [module],
      totalBlockers: module.blocker,
      totalProdBugs: module.prodBugs,
      timestamp: blockerData.timestamp
    };

    // Send to Google Chat if enabled for this module
    if (module.chatEnabled && module.chatWebhook) {
      if (sendGoogleChatNotification_(module.chatWebhook, perModuleData)) {
        chatSentCount++;
      }
    }

    // Send Email if enabled for this module
    if (module.emailEnabled && module.emailRecipients) {
      if (sendEmailNotification_(module.emailRecipients, perModuleData)) {
        emailSentCount++;
      }
    }
  });

  // Show result
  let msg = '📤 Notifications Sent!\n\n';
  if (chatSentCount > 0) msg += '✅ Google Chat: ' + chatSentCount + ' message(s) sent\n';
  if (emailSentCount > 0) msg += '✅ Email: ' + emailSentCount + ' message(s) sent\n';
  if (skippedCount > 0) msg += 'ℹ️ Skipped: ' + skippedCount + ' module(s) (notifications disabled)\n';
  msg += '\n📊 Summary:\n';
  msg += '• Total Open Blockers: ' + blockerData.totalBlockers + '\n';
  msg += '• Total PROD BUGS: ' + blockerData.totalProdBugs + '\n';
  msg += '• Modules with issues: ' + blockerData.modules.length + '\n';
  msg += '• Notifications sent: ' + (chatSentCount + emailSentCount) + '\n';
  msg += '• Modules skipped: ' + skippedCount;

  SpreadsheetApp.getUi().alert('Notification Sent', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log('✅ Notifications sent successfully - Chat: ' + chatSentCount + ', Email: ' + emailSentCount + ', Skipped: ' + skippedCount);
}

/**
 * Setup daily blocker notification trigger
 *
 * NEW: Uses per-module notification config when sending
 */
function setupDailyBlockerNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');

  if (!cfg) {
    SpreadsheetApp.getUi().alert('❌ Config tab not found!');
    return;
  }

  // Check if ANY module has notifications enabled
  const cfgData = cfg.getDataRange().getValues();
  let hasAnyEnabled = false;
  let notifHour = 7;

  for (let i = 3; i < cfgData.length; i++) {
    const chatEnabled = cfgData[i][13] === true; // Col N
    const emailEnabled = cfgData[i][15] === true; // Col P
    const hour = parseInt(cfgData[i][12]) || 7; // Col M

    if (chatEnabled || emailEnabled) {
      hasAnyEnabled = true;
      notifHour = hour; // Use the hour from first enabled module
      break;
    }
  }

  if (!hasAnyEnabled) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Notifications Enabled',
      'Tidak ada module yang mengaktifkan Google Chat atau Email notification.\n\n' +
      'Aktifkan di Config tab (kolom N = Enable Notifikasi, kolom P = Enable Email).',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Delete existing trigger
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendBlockerNotification') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new daily trigger at specified hour
  ScriptApp.newTrigger('sendBlockerNotification')
    .timeBased()
    .atHour(notifHour)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Daily Notification Setup!',
    'Blocker notification akan dikirim setiap hari jam ' + notifHour + ':00.\n\n' +
    '📤 Per-module notifications enabled.\n' +
    'Setiap module dengan blockers/PROD bugs akan kirim notifikasi terpisah\n' +
    'sesuai webhook dan enable flags di Config (kolom L-P).\n\n' +
    '💡 Test now: Menu > Notifications > Test Notification Now',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log('✅ Daily blocker notification trigger created at hour ' + notifHour);
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
 * Get blocker data from Overview tab and Config for QATM URLs + per-module notification config
 *
 * NEW: Includes per-module webhook URLs and enable flags from Config
 */
function getBlockerData_(overview, cfg) {
  const data = overview.getDataRange().getValues();

  // Build module map from Config for QATM URLs + notification config
  const moduleMap = {};
  if (cfg) {
    const cfgData = cfg.getDataRange().getValues();
    for (let i = 3; i < cfgData.length; i++) {
      const project = String(cfgData[i][2]).trim();  // Col C
      const modul = String(cfgData[i][3]).trim();     // Col D
      const qatmId = String(cfgData[i][6]).trim();    // Col G

      // Per-module notification config (NEW)
      // Col L (11) = Google Chat Webhook URL
      // Col N (13) = Enable Notifikasi (checkbox)
      // Col O (14) = Email Recipients
      // Col P (15) = Enable Email (checkbox)
      const chatWebhook = String(cfgData[i][11]).trim();
      const chatEnabled = cfgData[i][13] === true;
      const emailRecipients = String(cfgData[i][14]).trim();
      const emailEnabled = cfgData[i][15] === true;

      if (qatmId && qatmId.length > 10) {
        const key = project + '|' + modul;
        moduleMap[key] = {
          qatmUrl: 'https://docs.google.com/spreadsheets/d/' + qatmId + '/edit',
          bugReportGid: '2', // BugReport is typically GID 2
          chatWebhook: chatWebhook && chatWebhook.includes('chat.googleapis.com') ? chatWebhook : null,
          chatEnabled: chatEnabled,
          emailRecipients: emailRecipients && emailRecipients.includes('@') ? emailRecipients : null,
          emailEnabled: emailEnabled
        };
      }
    }
  }

  const modules = [];
  let totalBlockers = 0;
  let totalProdBugs = 0;

  // Start from row 5 (row index 4), skip headers
  for (let i = 4; i < data.length; i++) {
    const row = data[i];

    // Stop at TOTAL row or empty rows
    if (String(row[0]).includes('TOTAL') || String(row[0]).includes('AVERAGE')) break;
    if (!row[1] && !row[2]) continue; // Empty row

    const project = row[0] || '';
    const modul = row[1] || '';  // Col B = Modul
    const moduleName = row[2] || modul || 'Unknown'; // Col C = Submodule
    const blocker = parseInt(row[5]) || 0;  // Col F = Blocker
    const prodBugs = parseInt(row[7]) || 0; // Col H = PROD BUGS

    if (blocker > 0 || prodBugs > 0) {
      const key = project + '|' + modul;
      const moduleInfo = moduleMap[key] || {};

      modules.push({
        project: project,
        module: moduleName,
        blocker: blocker,
        prodBugs: prodBugs,
        qatmUrl: moduleInfo.qatmUrl || null,
        bugReportGid: moduleInfo.bugReportGid || '2',
        // Per-module notification config (NEW)
        chatWebhook: moduleInfo.chatWebhook || null,
        chatEnabled: moduleInfo.chatEnabled || false,
        emailRecipients: moduleInfo.emailRecipients || null,
        emailEnabled: moduleInfo.emailEnabled || false
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
 * Send Google Chat notification - PER MODULE format (concise)
 */
function sendGoogleChatNotification_(webhookUrl, blockerData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardUrl = ss.getUrl();

    // Get the single module (since we're sending per-module now)
    const module = blockerData.modules[0];

    let message = '';

    // Header - different based on whether PROD bugs exist
    if (module.prodBugs > 0) {
      message += '🚨🚨 *PRODUCTION BUGS ALERT* 🚨🚨\n\n';
      message += '🔴 *' + module.project + ' - ' + module.module + '*\n';
      message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
      message += '⚠️ *PROD Bugs: ' + module.prodBugs + '*\n';
      if (module.blocker > 0) {
        message += '⚠️ *Open Blockers: ' + module.blocker + '*\n';
      }
      message += '🕐 ' + blockerData.timestamp + '\n\n';

      message += '🚨 *EMERGENCY - IMMEDIATE ACTION REQUIRED!*\n';
      message += '_Production bugs affect LIVE USERS_\n\n';

      message += '⚡ *PROTOCOL:*\n';
      message += '• Drop other work - TOP PRIORITY\n';
      message += '• Assign senior dev immediately\n';
      message += '• Hourly status updates required\n';
      message += '• Hotfix deployment prioritized\n\n';
    } else {
      message += '⚠️ *QA BLOCKER ALERT*\n\n';
      message += '🟠 *' + module.project + ' - ' + module.module + '*\n';
      message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
      message += '⚠️ *Open Blockers: ' + module.blocker + '*\n';
      message += '🕐 ' + blockerData.timestamp + '\n\n';

      message += '📋 *ACTION REQUIRED:*\n';
      message += '_Critical/High/Medium bugs blocking test execution_\n\n';

      message += '• Impact: Testing cannot proceed\n';
      message += '• Target: Resolve within 48 hours\n';
      message += '• Daily standup: Report status\n';
      message += '• Escalate if stuck >2 days\n\n';
    }

    message += '🔗 *LINKS*\n';
    if (module.qatmUrl) {
      message += '• <' + module.qatmUrl + '#gid=' + module.bugReportGid + '|📋 View Bug Report>\n';
    }
    message += '• <' + dashboardUrl + '|📊 QA Dashboard>\n\n';

    message += '_💡 Automated notification - Reply here for questions_';

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

    // Get the single module (since we're sending per-module now)
    const module = blockerData.modules[0];

    let subject = '🚨 QA Blocker: ' + module.project + ' - ' + module.module;
    if (module.prodBugs > 0) {
      subject = '🚨🚨 URGENT PROD BUGS: ' + module.project + ' - ' + module.module;
    }

    let body = '<html><body style="font-family: Arial, sans-serif; max-width: 700px;">';

    // Header with module info
    if (module.prodBugs > 0) {
      body += '<div style="background: #D32F2F; color: white; padding: 20px; text-align: center;">';
      body += '<h2 style="margin: 0;">🚨🚨 PRODUCTION BUGS ALERT 🚨🚨</h2>';
      body += '</div>';
      body += '<div style="background: #FFCDD2; border-left: 4px solid #D32F2F; padding: 20px; margin: 0;">';
      body += '<h3 style="margin: 0 0 10px 0; color: #D32F2F;">' + module.project + ' - ' + module.module + '</h3>';
      body += '<table style="width: 100%; font-size: 14px;">';
      body += '<tr><td><strong style="color: #D32F2F;">PROD Bugs:</strong></td><td><span style="background: #D32F2F; color: white; padding: 4px 12px; border-radius: 3px; font-weight: bold; font-size: 18px;">' + module.prodBugs + '</span></td></tr>';
      if (module.blocker > 0) {
        body += '<tr><td><strong>Open Blockers:</strong></td><td><strong style="font-size: 16px; color: #E65100;">' + module.blocker + '</strong></td></tr>';
      }
      body += '<tr><td><strong>Timestamp:</strong></td><td>' + blockerData.timestamp + '</td></tr>';
      body += '</table>';
      body += '</div>';

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
      body += '<div style="background: #FF9800; color: white; padding: 15px; text-align: center;">';
      body += '<h2 style="margin: 0;">⚠️ QA BLOCKER ALERT</h2>';
      body += '</div>';
      body += '<div style="background: #FFF3E0; border-left: 4px solid #FF9800; padding: 20px; margin: 0;">';
      body += '<h3 style="margin: 0 0 10px 0; color: #E65100;">' + module.project + ' - ' + module.module + '</h3>';
      body += '<table style="width: 100%; font-size: 14px;">';
      body += '<tr><td><strong>Open Blockers:</strong></td><td><strong style="font-size: 18px; color: #E65100;">' + module.blocker + '</strong></td></tr>';
      body += '<tr><td><strong>Timestamp:</strong></td><td>' + blockerData.timestamp + '</td></tr>';
      body += '</table>';
      body += '</div>';

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
    if (module.qatmUrl) {
      body += '<p style="margin: 10px 0;"><a href="' + module.qatmUrl + '#gid=' + module.bugReportGid + '" style="background: #1976D2; color: white; padding: 8px 16px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">📋 View Bug Report</a></p>';
    }
    body += '<p style="margin: 10px 0;"><a href="' + dashboardUrl + '" style="color: #1976D2; font-weight: bold; font-size: 14px;">📊 QA Portfolio Dashboard</a></p>';
    body += '</div>';

    body += '<hr style="border: none; border-top: 1px solid #E0E0E0; margin: 20px 0;">';
    body += '<p style="color: #757575; font-size: 11px; text-align: center;">📊 ' + module.project + ' - ' + module.module + ' - Automated Notification<br>';
    body += 'This is a per-module alert sent automatically. For questions, contact QA Team.</p>';
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
