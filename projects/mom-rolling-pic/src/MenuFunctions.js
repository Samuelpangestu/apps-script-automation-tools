/**
 * MenuFunctions.js
 * Custom menu and user-facing functions for bi-daily standup
 * - onOpen() menu
 * - Initialize functions
 * - Test functions
 * - Manual trigger functions
 */

/**
 * On spreadsheet open - create custom menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📋 Bi-Daily Standup')
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('1️⃣ Initialize Config', 'initializeConfigTab')
      .addItem('2️⃣ Initialize Standup Sheets', 'initializeAllStandupSheets')
      .addSeparator()
      .addItem('🔄 Setup Auto Triggers', 'menuSetupAllTriggers')
      .addItem('🗑️ Remove All Triggers', 'menuRemoveAllTriggers'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Generate Standup')
      .addItem('Generate Today (Both Projects)', 'menuGenerateToday')
      .addItem('Generate for Date...', 'menuGenerateForDate')
      .addSeparator()
      .addItem('🔮 Bulk: Next 3 Months', 'menuGenerateBulk3Months')
      .addItem('🔮 Bulk: Custom Period...', 'menuGenerateBulkCustom')
      .addSeparator()
      .addItem('Generate SIPGN Today', 'menuGenerateProjectAToday')
      .addItem('Generate INADigital/Internal Today', 'menuGenerateProjectBToday'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📱 Test Notifications')
      .addItem('Test Reminder (SIPGN)', 'menuTestReminderA')
      .addItem('Test Reminder (INADigital/Internal)', 'menuTestReminderB')
      .addSeparator()
      .addItem('Test Summary (SIPGN)', 'menuTestSummaryA')
      .addItem('Test Summary (INADigital/Internal)', 'menuTestSummaryB')
      .addSeparator()
      .addItem('Get WhatsApp Groups', 'menuGetWhatsAppGroups'))
    .addSeparator()
    .addItem('📖 Help & Documentation', 'menuShowHelp')
    .addToUi();
}

/**
 * Initialize all standup sheets (both projects)
 */
function initializeAllStandupSheets() {
  try {
    initializeStandupSheet('SIPGN');
    initializeStandupSheet('INADigital/Internal');

    SpreadsheetApp.getUi().alert(
      '✅ Standup Sheets Initialized!\n\n' +
      'Both SIPGN and INADigital/Internal standup sheets have been created.\n\n' +
      'Next steps:\n' +
      '1. Update Config settings (team members, WhatsApp, times)\n' +
      '2. Setup Auto Triggers from menu\n' +
      '3. Test notifications'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Generate standup rows for today (both projects)
 */
function menuGenerateToday() {
  try {
    const ui = SpreadsheetApp.getUi();
    const today = new Date();
    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd MMMM yyyy');
    const dayName = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE');

    // Check if today is a standup day
    if (!['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
      ui.alert(
        '⚠️ Not a Standup Day',
        `Today is ${dayName}, ${dateStr}\n\n` +
        `Bi-daily standup is scheduled for Monday, Wednesday, Friday only.\n\n` +
        `Use "📅 Generate for Specific Date" to generate for a standup day.`,
        ui.ButtonSet.OK
      );
      return;
    }

    let countA = 0;
    let countB = 0;
    let errorA = null;
    let errorB = null;

    try {
      countA = generateStandupRows('SIPGN', today, true);
    } catch (e) {
      errorA = e.message;
      Logger.log('Error generating SIPGN: ' + e.message);
    }

    try {
      countB = generateStandupRows('INADigital/Internal', today, true);
    } catch (e) {
      errorB = e.message;
      Logger.log('Error generating INADigital/Internal: ' + e.message);
    }

    // Get project names from config
    let projectNameA = 'SIPGN';
    let projectNameB = 'INADigital/Internal';

    try {
      const configA = getProjectConfig('SIPGN');
      projectNameA = configA.projectName || 'SIPGN';
    } catch (e) {
      // Use default if config not found
    }

    try {
      const configB = getProjectConfig('INADigital/Internal');
      projectNameB = configB.projectName || 'INADigital/Internal';
    } catch (e) {
      // Use default if config not found
    }

    let message = `📅 Generate Standup for ${dateStr}\n\n`;

    // Project A result
    if (errorA) {
      message += `❌ ${projectNameA}: ${errorA}\n`;
    } else if (countA > 0) {
      message += `✅ ${projectNameA}: ${countA} rows generated\n`;
    } else {
      message += `ℹ️ ${projectNameA}: Already exists (skipped)\n`;
    }

    // Project B result
    if (errorB) {
      message += `❌ ${projectNameB}: ${errorB}\n`;
    } else if (countB > 0) {
      message += `✅ ${projectNameB}: ${countB} rows generated\n`;
    } else {
      message += `ℹ️ ${projectNameB}: Already exists (skipped)\n`;
    }

    // Suggestions if both failed
    if (errorA && errorB) {
      message += `\n💡 Suggestions:\n`;
      message += `• Initialize Standup Sheets first\n`;
      message += `• Check Config (team members set?)\n`;
      message += `• Check Execution log for details`;
    }

    ui.alert(message);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Generate standup for specific date
 */
function menuGenerateForDate() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Generate Standup for Date',
    'Enter date (YYYY-MM-DD format):\nExample: 2026-04-07',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const dateStr = response.getResponseText().trim();

  if (!dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
    ui.alert('❌ Invalid date format. Use YYYY-MM-DD (e.g., 2026-04-07)');
    return;
  }

  try {
    const date = new Date(dateStr);
    const dayName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEEE');

    // Block generation for non-standup days
    if (!['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
      ui.alert(
        '⚠️ Cannot Generate',
        `${dateStr} is ${dayName}\n\n` +
        `Bi-daily standup is scheduled for Monday, Wednesday, Friday only.\n\n` +
        `Please select a standup day.`,
        ui.ButtonSet.OK
      );
      return;
    }

    let countA = 0;
    let countB = 0;
    let errorA = null;
    let errorB = null;

    try {
      countA = generateStandupRows('SIPGN', date, false); // Don't skip if exists
    } catch (e) {
      errorA = e.message;
      Logger.log('Error generating SIPGN: ' + e.message);
    }

    try {
      countB = generateStandupRows('INADigital/Internal', date, false);
    } catch (e) {
      errorB = e.message;
      Logger.log('Error generating INADigital/Internal: ' + e.message);
    }

    // Get project names from config
    let projectNameA = 'SIPGN';
    let projectNameB = 'INADigital/Internal';

    try {
      const configA = getProjectConfig('SIPGN');
      projectNameA = configA.projectName || 'SIPGN';
    } catch (e) {
      // Use default if config not found
    }

    try {
      const configB = getProjectConfig('INADigital/Internal');
      projectNameB = configB.projectName || 'INADigital/Internal';
    } catch (e) {
      // Use default if config not found
    }

    let message = `📅 Generate Standup for ${dateStr} (${dayName})\n\n`;

    // Project A result
    if (errorA) {
      message += `❌ ${projectNameA}: ${errorA}\n`;
    } else {
      message += `✅ ${projectNameA}: ${countA} rows generated\n`;
    }

    // Project B result
    if (errorB) {
      message += `❌ ${projectNameB}: ${errorB}`;
    } else {
      message += `✅ ${projectNameB}: ${countB} rows generated`;
    }

    ui.alert(message);

  } catch (e) {
    ui.alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Bulk generate for next 3 months
 */
function menuGenerateBulk3Months() {
  try {
    const today = new Date();
    const endDate = new Date(today);
    endDate.setMonth(endDate.getMonth() + 3); // 3 months from today

    menuGenerateBulkPeriod(today, endDate);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Bulk generate for custom period
 */
function menuGenerateBulkCustom() {
  const ui = SpreadsheetApp.getUi();

  // Ask for start date
  const startResponse = ui.prompt(
    'Bulk Generate - Start Date',
    'Enter START date (YYYY-MM-DD):\nExample: 2026-04-07',
    ui.ButtonSet.OK_CANCEL
  );

  if (startResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const startDateStr = startResponse.getResponseText().trim();

  if (!startDateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
    ui.alert('❌ Invalid date format. Use YYYY-MM-DD (e.g., 2026-04-07)');
    return;
  }

  // Ask for end date
  const endResponse = ui.prompt(
    'Bulk Generate - End Date',
    'Enter END date (YYYY-MM-DD):\nExample: 2026-07-07',
    ui.ButtonSet.OK_CANCEL
  );

  if (endResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const endDateStr = endResponse.getResponseText().trim();

  if (!endDateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
    ui.alert('❌ Invalid date format. Use YYYY-MM-DD (e.g., 2026-07-07)');
    return;
  }

  try {
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);

    if (endDate <= startDate) {
      ui.alert('❌ Error: End date must be after start date');
      return;
    }

    menuGenerateBulkPeriod(startDate, endDate);

  } catch (e) {
    ui.alert('❌ Error: ' + e.message);
  }
}

/**
 * Generate standup rows for a period (only Mon/Wed/Fri)
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 */
function menuGenerateBulkPeriod(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();

  // Get project names from config
  let projectNameA = 'SIPGN';
  let projectNameB = 'INADigital/Internal';

  try {
    const configA = getProjectConfig('SIPGN');
    projectNameA = configA.projectName || 'SIPGN';
  } catch (e) {
    // Use default if config not found
  }

  try {
    const configB = getProjectConfig('INADigital/Internal');
    projectNameB = configB.projectName || 'INADigital/Internal';
  } catch (e) {
    // Use default if config not found
  }

  const startStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'dd MMM yyyy');
  const endStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'dd MMM yyyy');

  const confirm = ui.alert(
    'Bulk Generate Confirmation',
    `Generate standup rows for:\n\n` +
    `📅 Period: ${startStr} - ${endStr}\n` +
    `📋 Projects: ${projectNameA} & ${projectNameB}\n` +
    `🗓️ Days: Monday, Wednesday, Friday only\n\n` +
    `This may take a few seconds...\n\n` +
    `Continue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }

  try {
    let totalCountA = 0;
    let totalCountB = 0;
    let totalDays = 0;

    const currentDate = new Date(startDate);

    while (currentDate <= endDate) {
      const dayName = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'EEEE');

      // Only generate for Mon/Wed/Fri
      if (['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
        let countA = 0;
        let countB = 0;

        try {
          countA = generateStandupRows('SIPGN', new Date(currentDate), true); // skipIfExists = true
        } catch (e) {
          Logger.log(`Error generating SIPGN for ${currentDate}: ${e.message}`);
        }

        try {
          countB = generateStandupRows('INADigital/Internal', new Date(currentDate), true); // skipIfExists = true
        } catch (e) {
          Logger.log(`Error generating INADigital/Internal for ${currentDate}: ${e.message}`);
        }

        if (countA > 0 || countB > 0) {
          totalCountA += countA;
          totalCountB += countB;
          totalDays++;

          const dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          Logger.log(`✅ Generated for ${dateStr}: A=${countA}, B=${countB}`);
        }
      }

      // Move to next day
      currentDate.setDate(currentDate.getDate() + 1);
    }

    let message = `✅ Bulk Generation Complete!\n\n`;
    message += `📅 Period: ${startStr} - ${endStr}\n\n`;
    message += `📊 Results:\n`;
    message += `• Total standup days: ${totalDays}\n`;
    message += `• ${projectNameA} rows: ${totalCountA}\n`;
    message += `• ${projectNameB} rows: ${totalCountB}\n\n`;
    message += `ℹ️ Existing dates were skipped (data protected)`;

    ui.alert(message);

  } catch (e) {
    ui.alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Generate SIPGN today only
 */
function menuGenerateProjectAToday() {
  try {
    const today = new Date();
    const count = generateStandupRows('SIPGN', today, true);

    if (count > 0) {
      SpreadsheetApp.getUi().alert(`✅ SIPGN: ${count} rows generated`);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ SIPGN: Rows already exist for today');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Generate INADigital/Internal today only
 */
function menuGenerateProjectBToday() {
  try {
    const today = new Date();
    const count = generateStandupRows('INADigital/Internal', today, true);

    if (count > 0) {
      SpreadsheetApp.getUi().alert(`✅ INADigital/Internal: ${count} rows generated`);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ INADigital/Internal: Rows already exist for today');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Test reminder notification for SIPGN
 */
function menuTestReminderA() {
  try {
    sendStandupReminder('SIPGN', true); // isTest = true (skip day validation)
    SpreadsheetApp.getUi().alert('✅ Test reminder sent for SIPGN\n\nCheck your WhatsApp group');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Test reminder notification for INADigital/Internal
 */
function menuTestReminderB() {
  try {
    sendStandupReminder('INADigital/Internal', true); // isTest = true (skip day validation)
    SpreadsheetApp.getUi().alert('✅ Test reminder sent for INADigital/Internal\n\nCheck your WhatsApp group');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Test summary notification for SIPGN
 */
function menuTestSummaryA() {
  try {
    sendStandupSummary('SIPGN');
    SpreadsheetApp.getUi().alert('✅ Test summary sent for SIPGN\n\nCheck your WhatsApp group');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Test summary notification for INADigital/Internal
 */
function menuTestSummaryB() {
  try {
    sendStandupSummary('INADigital/Internal');
    SpreadsheetApp.getUi().alert('✅ Test summary sent for INADigital/Internal\n\nCheck your WhatsApp group');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Get WhatsApp groups from Fonnte
 */
function menuGetWhatsAppGroups() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Get WhatsApp Groups',
    'Enter your Fonnte Token:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const fontteToken = response.getResponseText().trim();

  if (!fontteToken) {
    ui.alert('❌ Fonnte Token cannot be empty');
    return;
  }

  try {
    const url = 'https://api.fonnte.com/get-devices';

    const options = {
      method: 'post',
      headers: {
        'Authorization': fontteToken
      },
      muteHttpExceptions: true
    };

    const apiResponse = UrlFetchApp.fetch(url, options);
    const responseCode = apiResponse.getResponseCode();
    const responseText = apiResponse.getContentText();

    if (responseCode !== 200) {
      ui.alert('❌ API Error\n\nResponse Code: ' + responseCode + '\nResponse: ' + responseText);
      return;
    }

    const data = JSON.parse(responseText);

    let message = '📱 WHATSAPP DEVICES & GROUPS\n\n';

    if (data.device && data.device.length > 0) {
      data.device.forEach((device, index) => {
        message += '═══════════════════\n';
        message += `DEVICE ${index + 1}:\n`;
        message += `• Name: ${device.name || 'N/A'}\n`;
        message += `• Number: ${device.device || 'N/A'}\n`;
        message += `• Status: ${device.status || 'N/A'}\n\n`;

        if (device.groups && device.groups.length > 0) {
          message += `GROUPS (${device.groups.length}):\n`;
          device.groups.forEach(group => {
            message += `  • ${group.name}\n`;
            message += `    ID: ${group.id}\n`;
          });
        } else {
          message += 'GROUPS: None found\n';
        }
        message += '\n';
      });

      message += '═══════════════════\n\n';
      message += '💡 TIP: Copy Group ID (120363xxx@g.us) to Config sheet';

    } else {
      message = '❌ No devices found\n\nRaw Response:\n' + responseText;
    }

    ui.alert(message);

  } catch (e) {
    ui.alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Setup all auto triggers
 */
function menuSetupAllTriggers() {
  try {
    setupAllTriggers();
    SpreadsheetApp.getUi().alert(
      '✅ Triggers Setup Complete!\n\n' +
      'Auto triggers have been created for:\n' +
      '• SIPGN reminder & summary\n' +
      '• INADigital/Internal reminder & summary\n\n' +
      'Check Config sheet for schedule times.'
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Remove all triggers
 */
function menuRemoveAllTriggers() {
  try {
    removeAllTriggers();
    SpreadsheetApp.getUi().alert('✅ All triggers removed');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Show help documentation
 */
function menuShowHelp() {
  const help =
    '📋 BI-DAILY STANDUP - HELP & DOCUMENTATION\n\n' +
    '═══════════════════════════════════\n\n' +
    '🎯 QUICK START:\n\n' +
    '1. Setup → Initialize Config\n' +
    '   Set team members, WhatsApp config, times\n\n' +
    '2. Setup → Initialize Standup Sheets\n' +
    '   Create SIPGN & B sheets\n\n' +
    '3. Setup → Setup Auto Triggers\n' +
    '   Enable automatic reminders & summaries\n\n' +
    '4. Test Notifications\n' +
    '   Send test messages to verify WhatsApp\n\n' +
    '═══════════════════════════════════\n\n' +
    '📅 SCHEDULE:\n' +
    '• Bi-Daily: Monday, Wednesday, Friday\n' +
    '• Reminder: 30 min before standup\n' +
    '• Summary: At configured time\n\n' +
    '═══════════════════════════════════\n\n' +
    '📱 WHATSAPP SETUP:\n' +
    '1. Get Fonnte Token from fonnte.com\n' +
    '2. Use "Get WhatsApp Groups" menu\n' +
    '3. Copy Group ID (120363xxx@g.us)\n' +
    '4. Paste to Config sheet\n\n' +
    '═══════════════════════════════════\n\n' +
    '📊 STANDUP FORMAT:\n' +
    '• 3 rows per person per session\n' +
    '• Done: Tasks completed\n' +
    '• In Progress: Current tasks\n' +
    '• Blocker: Any blockers\n\n' +
    '═══════════════════════════════════';

  SpreadsheetApp.getUi().alert(help);
}
