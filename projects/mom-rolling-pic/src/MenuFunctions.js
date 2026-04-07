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
      .addItem('Generate Project A Today', 'menuGenerateProjectAToday')
      .addItem('Generate Project B Today', 'menuGenerateProjectBToday'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📱 Test Notifications')
      .addItem('Test Reminder (Project A)', 'menuTestReminderA')
      .addItem('Test Reminder (Project B)', 'menuTestReminderB')
      .addSeparator()
      .addItem('Test Summary (Project A)', 'menuTestSummaryA')
      .addItem('Test Summary (Project B)', 'menuTestSummaryB')
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
    initializeStandupSheet('Project A');
    initializeStandupSheet('Project B');

    SpreadsheetApp.getUi().alert(
      '✅ Standup Sheets Initialized!\n\n' +
      'Both Project A and Project B standup sheets have been created.\n\n' +
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
    const today = new Date();
    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd MMMM yyyy');

    let countA = 0;
    let countB = 0;

    try {
      countA = generateStandupRows('Project A', today, true);
    } catch (e) {
      Logger.log('Error generating Project A: ' + e.message);
    }

    try {
      countB = generateStandupRows('Project B', today, true);
    } catch (e) {
      Logger.log('Error generating Project B: ' + e.message);
    }

    let message = `📅 Generate Standup for ${dateStr}\n\n`;

    if (countA > 0) {
      message += `✅ Project A: ${countA} rows generated\n`;
    } else {
      message += `⚠️ Project A: Already exists or error\n`;
    }

    if (countB > 0) {
      message += `✅ Project B: ${countB} rows generated\n`;
    } else {
      message += `⚠️ Project B: Already exists or error\n`;
    }

    SpreadsheetApp.getUi().alert(message);

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

    if (!['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
      ui.alert(`⚠️ Warning: ${dateStr} is ${dayName}\n\nBi-daily standup is scheduled for Monday, Wednesday, Friday only.\n\nContinue anyway?`);
    }

    let countA = 0;
    let countB = 0;

    try {
      countA = generateStandupRows('Project A', date, false); // Don't skip if exists
    } catch (e) {
      Logger.log('Error generating Project A: ' + e.message);
    }

    try {
      countB = generateStandupRows('Project B', date, false);
    } catch (e) {
      Logger.log('Error generating Project B: ' + e.message);
    }

    let message = `📅 Generate Standup for ${dateStr} (${dayName})\n\n`;
    message += `✅ Project A: ${countA} rows generated\n`;
    message += `✅ Project B: ${countB} rows generated`;

    ui.alert(message);

  } catch (e) {
    ui.alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Generate Project A today only
 */
function menuGenerateProjectAToday() {
  try {
    const today = new Date();
    const count = generateStandupRows('Project A', today, true);

    if (count > 0) {
      SpreadsheetApp.getUi().alert(`✅ Project A: ${count} rows generated`);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ Project A: Rows already exist for today');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Generate Project B today only
 */
function menuGenerateProjectBToday() {
  try {
    const today = new Date();
    const count = generateStandupRows('Project B', today, true);

    if (count > 0) {
      SpreadsheetApp.getUi().alert(`✅ Project B: ${count} rows generated`);
    } else {
      SpreadsheetApp.getUi().alert('⚠️ Project B: Rows already exist for today');
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Test reminder notification for Project A
 */
function menuTestReminderA() {
  try {
    sendStandupReminder('Project A');
    SpreadsheetApp.getUi().alert('✅ Test reminder sent for Project A\n\nCheck your WhatsApp group');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Test reminder notification for Project B
 */
function menuTestReminderB() {
  try {
    sendStandupReminder('Project B');
    SpreadsheetApp.getUi().alert('✅ Test reminder sent for Project B\n\nCheck your WhatsApp group');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Test summary notification for Project A
 */
function menuTestSummaryA() {
  try {
    sendStandupSummary('Project A');
    SpreadsheetApp.getUi().alert('✅ Test summary sent for Project A\n\nCheck your WhatsApp group');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/**
 * MENU: Test summary notification for Project B
 */
function menuTestSummaryB() {
  try {
    sendStandupSummary('Project B');
    SpreadsheetApp.getUi().alert('✅ Test summary sent for Project B\n\nCheck your WhatsApp group');
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
      '• Project A reminder & summary\n' +
      '• Project B reminder & summary\n\n' +
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
    '   Create Project A & B sheets\n\n' +
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
