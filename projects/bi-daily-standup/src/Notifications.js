/**
 * Notifications.js
 * WhatsApp notifications for bi-daily standup
 * - Reminder before standup (30 min before)
 * - Summary after standup (detailed with task names)
 */

/**
 * Send WhatsApp reminder before standup
 * Triggered automatically based on schedule
 * @param {string} projectName - 'Project A' or 'Project B'
 */
function sendStandupReminder(projectName) {
  try {
    const config = getProjectConfig(projectName);

    if (!config.enableReminder) {
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

    // Get standup time for today
    const standupTime = getStandupTimeForDay(config, dayName);

    if (!standupTime) {
      Logger.log(`⚠️ No standup scheduled for ${dayName}`);
      return;
    }

    // Generate standup rows for today (if not exists)
    try {
      generateStandupRows(projectName, now, true); // skipIfExists = true
    } catch (e) {
      Logger.log(`⚠️ Error generating rows: ${e.message}`);
    }

    // Build reminder message
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUrl = ss.getUrl() + '#gid=' + ss.getSheetByName(projectName + ' Standup').getSheetId();

    const message =
      `🔔 *BI-DAILY STANDUP REMINDER*\n` +
      `📅 ${dateStr}\n` +
      `━━━━━━━━━━━━━━━━━━━━\n\n` +
      `⏰ Standup Time: *${standupTime}*\n` +
      `📋 Project: *${config.projectName}*\n\n` +
      `👥 Team: ${config.teamMembers.join(', ')}\n\n` +
      `━━━━━━━━━━━━━━━━━━━━\n\n` +
      `📝 *PLEASE UPDATE:*\n` +
      `✅ Done - Tasks completed since last standup\n` +
      `📋 In Progress - Tasks until next standup\n` +
      `🚨 Blockers - Any blockers or help needed\n\n` +
      `━━━━━━━━━━━━━━━━━━━━\n` +
      `📊 Update sheet: ${sheetUrl}\n\n` +
      `_Automated Reminder - Bi-Daily Standup_`;

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
 * @param {string} projectName - 'Project A' or 'Project B'
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

    // Build summary message (Opsi 2: Detailed with task names)
    let message = '';

    message += `📊 *BI-DAILY STANDUP SUMMARY*\n`;
    message += `📅 ${dateStr}\n`;
    message += `📋 Project: *${config.projectName}*\n`;
    message += `━━━━━━━━━━━━━━━━━━━━\n\n`;

    summaryByPerson.forEach(personSummary => {
      message += `👤 *${personSummary.person}:*\n`;

      // Done tasks
      if (personSummary.done.length > 0) {
        message += `  ✅ *Done:*\n`;
        personSummary.done.forEach(task => {
          message += `    • ${task}\n`;
        });
      } else {
        message += `  ✅ *Done:* None\n`;
      }

      // In Progress tasks
      if (personSummary.inProgress.length > 0) {
        message += `  📋 *In Progress:*\n`;
        personSummary.inProgress.forEach(task => {
          message += `    • ${task}\n`;
        });
      } else {
        message += `  📋 *In Progress:* None\n`;
      }

      // Blockers
      if (personSummary.blockers.length > 0) {
        message += `  🚨 *Blockers:*\n`;
        personSummary.blockers.forEach(blocker => {
          message += `    • ${blocker}\n`;
        });
      } else {
        message += `  🚨 *Blockers:* None\n`;
      }

      message += `\n`;
    });

    // Total summary
    const totalDone = summaryByPerson.reduce((sum, p) => sum + p.done.length, 0);
    const totalInProgress = summaryByPerson.reduce((sum, p) => sum + p.inProgress.length, 0);
    const totalBlockers = summaryByPerson.reduce((sum, p) => sum + p.blockers.length, 0);

    message += `━━━━━━━━━━━━━━━━━━━━\n`;
    message += `📊 *TOTAL:*\n`;
    message += `  ✅ ${totalDone} tasks completed\n`;
    message += `  📋 ${totalInProgress} tasks in progress\n`;
    message += `  🚨 ${totalBlockers} blocker(s)\n\n`;

    // Sheet link
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUrl = ss.getUrl() + '#gid=' + ss.getSheetByName(projectName + ' Standup').getSheetId();

    message += `━━━━━━━━━━━━━━━━━━━━\n`;
    message += `📋 Full details: ${sheetUrl}\n\n`;
    message += `_Automated Summary - Bi-Daily Standup_`;

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
