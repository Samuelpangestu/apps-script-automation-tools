/**
 * Notifications.js
 * WhatsApp notifications for bi-daily standup
 * - Reminder before standup (30 min before)
 * - Summary after standup (detailed with task names)
 */

/**
 * Send WhatsApp reminder before standup
 * Triggered automatically based on schedule
 * @param {string} projectName - 'SIPGN' or 'INADigital/Internal'
 * @param {boolean} isTest - If true, skip day validation (for testing)
 */
function sendStandupReminder(projectName, isTest = false) {
  try {
    const config = getProjectConfig(projectName);

    if (!config.enableReminder && !isTest) {
      Logger.log(`⚠️ Reminder disabled for ${projectName}`);
      return;
    }

    // Validate WhatsApp config
    if (!config.whatsappGroupId || !config.fontteToken) {
      Logger.log(`⚠️ WhatsApp config incomplete for ${projectName}`);
      return;
    }

    // Get current date and day
    const now = new Date();
    const dayName = Utilities.formatDate(now, Session.getScriptTimeZone(), 'EEEE');
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'EEEE, dd MMMM yyyy');

    // Get standup time for today (already formatted as HH:MM from config)
    let standupTime = getStandupTimeForDay(config, dayName);

    // For testing, use Monday time if no standup scheduled today
    if (!standupTime && isTest) {
      standupTime = config.mondayTime;
      Logger.log(`ℹ️ Test mode: Using Monday time ${standupTime}`);
    }

    if (!standupTime) {
      Logger.log(`⚠️ No standup scheduled for ${dayName}`);
      return;
    }

    // Generate standup rows for today (if not exists)
    // Skip generation in test mode if not a standup day
    if (!isTest || ['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
      try {
        generateStandupRows(projectName, now, true); // skipIfExists = true
      } catch (e) {
        Logger.log(`⚠️ Error generating rows: ${e.message}`);
      }
    } else {
      Logger.log(`ℹ️ Test mode on ${dayName}: Skipping row generation (not a standup day)`);
    }

    // Build reminder message
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUrl = ss.getUrl() + '#gid=' + ss.getSheetByName(projectName + ' Standup').getSheetId();

    const message =
      `🔔 *BI-DAILY STANDUP REMINDER*\n` +
      `${dateStr}\n` +
      `━━━━━━━━━━━━━━\n\n` +
      `Waktu: *${standupTime} WIB*\n` +
      `Project: *${config.projectName}*\n` +
      `Google Meet: ${config.googleMeetLink}\n\n` +
      `@all\n\n` +
      `━━━━━━━━━━━━━━\n\n` +
      `*Please update:*\n` +
      `• Task progress\n` +
      `• Test execution status\n` +
      `• Jira ticket updates\n` +
      `• Create new Jira tasks if needed\n\n` +
      `Standup Sheet: ${sheetUrl}\n\n` +
      `_Auto Reminder - Bi-Daily Standup_`;

    // Send via Fonnte
    const success = sendWhatsAppMessage(config.whatsappGroupId, config.fontteToken, message);

    if (success) {
      Logger.log(`✅ Standup reminder sent for ${projectName}`);
    } else {
      Logger.log(`❌ Failed to send standup reminder for ${projectName}`);
    }

  } catch (e) {
    Logger.log(`❌ Error sending standup reminder for ${projectName}: ${e.message}`);
    Logger.log(e.stack);
  }
}

/**
 * Send WhatsApp summary with detailed task breakdown
 * Triggered at configured summary time
 * @param {string} projectName - 'SIPGN' or 'INADigital/Internal'
 */
function sendStandupSummary(projectName) {
  try {
    const config = getProjectConfig(projectName);

    if (!config.enableSummary) {
      Logger.log(`⚠️ Summary disabled for ${projectName}`);
      return;
    }

    // Validate WhatsApp config
    if (!config.whatsappGroupId || !config.fontteToken) {
      Logger.log(`⚠️ WhatsApp config incomplete for ${projectName}`);
      return;
    }

    // Get today's standup data
    const today = new Date();
    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, dd MMMM yyyy');

    const standupData = getStandupDataForDate(projectName, today);

    if (!standupData || standupData.length === 0) {
      Logger.log(`⚠️ No standup data found for ${projectName} today`);
      return;
    }

    // Group by person
    const summaryByPerson = getStandupSummaryByPerson(standupData);

    if (summaryByPerson.length === 0) {
      Logger.log(`⚠️ No updates found in standup data for ${projectName}`);
      return;
    }

    // Build summary message
    let message = '';

    message += `📊 *BI-DAILY STANDUP SUMMARY*\n`;
    message += `${dateStr}\n`;
    message += `Project: *${config.projectName}*\n`;
    message += `@all\n`;
    message += `━━━━━━━━━━━━━━\n\n`;

    summaryByPerson.forEach(personSummary => {
      message += `*${personSummary.person}:*\n`;

      // Done Since Last
      if (personSummary.done.length > 0) {
        message += `  ✅ *Done:*\n`;
        personSummary.done.forEach(item => {
          message += `    • ${item}\n`;
        });
      }

      // Plan Next 2 Days
      if (personSummary.plan.length > 0) {
        message += `  📋 *Plan:*\n`;
        personSummary.plan.forEach(item => {
          message += `    • ${item}\n`;
        });
      }

      // Blockers
      if (personSummary.blockers.length > 0) {
        message += `  🚫 *Blockers:*\n`;
        personSummary.blockers.forEach(item => {
          message += `    • ${item}\n`;
        });
      }

      // If no updates at all
      if (personSummary.done.length === 0 &&
          personSummary.plan.length === 0 &&
          personSummary.blockers.length === 0) {
        message += `  • No updates\n`;
      }

      message += `\n`;
    });

    // Total summary
    const totalDone = summaryByPerson.reduce((sum, p) => sum + p.done.length, 0);
    const totalPlan = summaryByPerson.reduce((sum, p) => sum + p.plan.length, 0);
    const totalBlockers = summaryByPerson.reduce((sum, p) => sum + p.blockers.length, 0);

    message += `━━━━━━━━━━━━━━\n`;
    message += `*TOTAL:*\n`;
    message += `  • ${totalDone} completed items\n`;
    message += `  • ${totalPlan} planned items\n`;
    message += `  • ${totalBlockers} blockers\n\n`;

    // Sheet link
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUrl = ss.getUrl() + '#gid=' + ss.getSheetByName(projectName + ' Standup').getSheetId();

    message += `Standup Sheet: ${sheetUrl}\n\n`;
    message += `_Auto Summary - Bi-Daily Standup_`;

    // Send via Fonnte
    const success = sendWhatsAppMessage(config.whatsappGroupId, config.fontteToken, message);

    if (success) {
      Logger.log(`✅ Standup summary sent for ${projectName}`);
    } else {
      Logger.log(`❌ Failed to send standup summary for ${projectName}`);
    }

  } catch (e) {
    Logger.log(`❌ Error sending standup summary for ${projectName}: ${e.message}`);
    Logger.log(e.stack);
  }
}

/**
 * Send WhatsApp message via Fonnte API
 * @param {string} groupId - WhatsApp group ID (120363xxx@g.us)
 * @param {string} fontteToken - Fonnte API token
 * @param {string} message - Message text
 * @returns {boolean} Success status
 */
function sendWhatsAppMessage(groupId, fontteToken, message) {
  try {
    const url = 'https://api.fonnte.com/send';

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

    Logger.log(`Sending WhatsApp to: ${groupId}`);
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`Response Code: ${responseCode}`);
    Logger.log(`Response: ${responseText}`);

    if (responseCode === 200) {
      const data = JSON.parse(responseText);

      if (data.status) {
        Logger.log('✅ WhatsApp message sent successfully');
        return true;
      } else {
        Logger.log(`❌ WhatsApp send failed: ${data.reason || 'Unknown error'}`);
        return false;
      }
    } else {
      Logger.log(`❌ WhatsApp API error: ${responseCode} - ${responseText}`);
      return false;
    }

  } catch (e) {
    Logger.log(`❌ Error sending WhatsApp: ${e.message}`);
    Logger.log(e.stack);
    return false;
  }
}
