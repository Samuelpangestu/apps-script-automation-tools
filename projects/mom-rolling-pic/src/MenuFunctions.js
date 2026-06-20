/**
 * MenuFunctions.js
 * Spreadsheet menu for QA Bi-Daily.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 QA Bi-Daily')
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('1️⃣ Initialize Config', 'initializeConfigTab')
      .addItem('➕ Add Project', 'addProjectConfig')
      .addItem('2️⃣ Initialize Project Sheets', 'initializeAllStandupSheets')
      .addSeparator()
      .addItem('🔄 Setup Auto Triggers', 'menuSetupAllTriggers')
      .addItem('🗑️ Remove All Triggers', 'menuRemoveAllTriggers'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Generate Standup')
      .addItem('Generate Today (All Projects)', 'menuGenerateToday')
      .addItem('Generate for Date...', 'menuGenerateForDate')
      .addSeparator()
      .addItem('🔮 Bulk: Next 3 Months', 'menuGenerateBulk3Months')
      .addItem('🔮 Bulk: Custom Period...', 'menuGenerateBulkCustom')
      .addSeparator()
      .addItem('Generate One Project Today...', 'menuGenerateOneProjectToday'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📱 Test Notifications')
      .addItem('Test Reminder...', 'menuTestReminder')
      .addItem('Test Summary...', 'menuTestSummary')
      .addSeparator()
      .addItem('Get WhatsApp Groups', 'menuGetWhatsAppGroups'))
    .addSeparator()
    .addItem('📖 Help & Documentation', 'menuShowHelp')
    .addToUi();
}

function initializeAllStandupSheets() {
  const ui = SpreadsheetApp.getUi();
  try {
    const projects = getProjectNames();
    if (!projects.length) throw new Error('No projects found in Config.');
    projects.forEach(initializeStandupSheet);
    ui.alert(`✅ ${projects.length} QA Bi-Daily project sheet(s) initialized:\n\n${projects.join('\n')}`);
  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

function menuGenerateToday() {
  generateForAllProjects_(new Date(), true);
}

function menuGenerateForDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Generate QA Bi-Daily for Date',
    'Enter date in YYYY-MM-DD format:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const value = response.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    ui.alert('❌ Invalid date format. Use YYYY-MM-DD.');
    return;
  }
  generateForAllProjects_(new Date(value + 'T12:00:00'), false);
}

function generateForAllProjects_(date, skipIfExists) {
  const ui = SpreadsheetApp.getUi();
  const dateLabel = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const results = [];

  getProjectNames().forEach(projectName => {
    try {
      const config = getProjectConfig(projectName);
      const dayName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEEE');
      if (!isConfiguredStandupDay_(config, dayName)) {
        results.push(`ℹ️ ${projectName}: not scheduled on ${dayName}`);
        return;
      }
      const count = generateStandupRows(projectName, date, skipIfExists);
      results.push(count > 0
        ? `✅ ${projectName}: ${count} rows generated`
        : `ℹ️ ${projectName}: existing rows skipped`);
    } catch (error) {
      results.push(`❌ ${projectName}: ${error.message}`);
    }
  });

  ui.alert(`📅 QA Bi-Daily generation for ${dateLabel}\n\n${results.join('\n')}`);
}

function menuGenerateBulk3Months() {
  const startDate = new Date();
  const endDate = new Date(startDate);
  endDate.setMonth(endDate.getMonth() + 3);
  menuGenerateBulkPeriod(startDate, endDate);
}

function menuGenerateBulkCustom() {
  const ui = SpreadsheetApp.getUi();
  const startResponse = ui.prompt('Bulk Generate', 'Start date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (startResponse.getSelectedButton() !== ui.Button.OK) return;
  const endResponse = ui.prompt('Bulk Generate', 'End date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (endResponse.getSelectedButton() !== ui.Button.OK) return;

  const startValue = startResponse.getResponseText().trim();
  const endValue = endResponse.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(startValue) || !/^\d{4}-\d{2}-\d{2}$/.test(endValue)) {
    ui.alert('❌ Invalid date format. Use YYYY-MM-DD.');
    return;
  }

  const startDate = new Date(startValue + 'T12:00:00');
  const endDate = new Date(endValue + 'T12:00:00');
  if (endDate < startDate) {
    ui.alert('❌ End date must be on or after start date.');
    return;
  }
  menuGenerateBulkPeriod(startDate, endDate);
}

function menuGenerateBulkPeriod(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  const projects = getProjectNames();
  const totals = {};
  projects.forEach(projectName => { totals[projectName] = 0; });

  const currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    const dayName = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'EEEE');
    projects.forEach(projectName => {
      try {
        const config = getProjectConfig(projectName);
        if (isConfiguredStandupDay_(config, dayName)) {
          totals[projectName] += generateStandupRows(projectName, new Date(currentDate), true);
        }
      } catch (error) {
        Logger.log(`❌ Bulk generation error for ${projectName}: ${error.message}`);
      }
    });
    currentDate.setDate(currentDate.getDate() + 1);
  }

  const resultLines = projects.map(projectName => `• ${projectName}: ${totals[projectName]} rows`);
  ui.alert('✅ QA Bi-Daily bulk generation complete.\n\n' + resultLines.join('\n'));
}

function menuGenerateOneProjectToday() {
  const projectName = promptForProject_('Generate One Project');
  if (!projectName) return;
  try {
    const count = generateStandupRows(projectName, new Date(), true);
    SpreadsheetApp.getUi().alert(`✅ ${projectName}: ${count} rows generated`);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

function menuTestReminder() {
  const projectName = promptForProject_('Test Reminder');
  if (!projectName) return;
  sendStandupReminder(projectName, true);
  SpreadsheetApp.getUi().alert(`✅ Test reminder processed for ${projectName}. Check execution logs and WhatsApp.`);
}

function menuTestSummary() {
  const projectName = promptForProject_('Test Summary');
  if (!projectName) return;
  sendStandupSummary(projectName, true);
  SpreadsheetApp.getUi().alert(`✅ Test summary processed for ${projectName}. Check execution logs and WhatsApp.`);
}

function promptForProject_(title) {
  const ui = SpreadsheetApp.getUi();
  const projects = getProjectNames();
  if (!projects.length) {
    ui.alert('❌ No projects found in Config.');
    return null;
  }

  const response = ui.prompt(
    title,
    'Enter one project name exactly as configured:\n\n' + projects.join('\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const input = response.getResponseText().trim();
  const projectName = projects.find(name => name.toLowerCase() === input.toLowerCase());
  if (!projectName) {
    ui.alert(`❌ Project "${input}" not found.`);
    return null;
  }
  return projectName;
}

function menuGetWhatsAppGroups() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Get WhatsApp Groups', 'Enter your Fonnte Token:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const token = response.getResponseText().trim();
  if (!token) return ui.alert('❌ Fonnte Token cannot be empty.');

  try {
    const apiResponse = UrlFetchApp.fetch('https://api.fonnte.com/get-devices', {
      method: 'post',
      headers: { Authorization: token },
      muteHttpExceptions: true
    });
    if (apiResponse.getResponseCode() !== 200) {
      return ui.alert(`❌ API Error ${apiResponse.getResponseCode()}\n\n${apiResponse.getContentText()}`);
    }

    const data = JSON.parse(apiResponse.getContentText());
    let message = '📱 WHATSAPP DEVICES & GROUPS\n\n';
    (data.device || []).forEach(device => {
      message += `${device.name || device.device || 'Device'}\n`;
      (device.groups || []).forEach(group => {
        message += `• ${group.name}: ${group.id}\n`;
      });
      message += '\n';
    });
    ui.alert(message);
  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
  }
}

function menuSetupAllTriggers() {
  try {
    setupAllTriggers();
    SpreadsheetApp.getUi().alert(
      `✅ QA Bi-Daily scheduler configured for ${getProjectNames().length} project(s).\n\n` +
      'The scheduler checks all project schedules every 5 minutes.'
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

function menuRemoveAllTriggers() {
  removeAllTriggers();
  SpreadsheetApp.getUi().alert('✅ All QA Bi-Daily triggers removed.');
}

function menuShowHelp() {
  SpreadsheetApp.getUi().alert(
    '📋 QA BI-DAILY\n\n' +
    '1. Initialize Config once, or use Add Project for additional projects.\n' +
    '2. Complete schedule, team, and WhatsApp configuration.\n' +
    '3. Initialize Project Sheets.\n' +
    '4. Run Setup Auto Triggers after project or schedule changes.\n\n' +
    'The scheduler supports any number of configured projects.'
  );
}
