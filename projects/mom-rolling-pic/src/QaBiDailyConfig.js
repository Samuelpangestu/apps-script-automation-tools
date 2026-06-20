/**
 * QaBiDailyConfig.js
 * Dynamic configuration management for QA Bi-Daily.
 */

const QA_BI_DAILY_CONFIG_SHEET = 'Config';
const QA_BI_DAILY_SETTINGS = [
  'Project Name',
  'Standup Days',
  'Monday Time',
  'Wednesday Time',
  'Friday Time',
  'Reminder Offset (minutes)',
  'Summary Time',
  'Google Meet Link',
  'WhatsApp Group ID',
  'Fonnte Token',
  'Enable Reminder',
  'Enable Summary'
];

/**
 * Initialize Config with the two existing projects as starter data.
 * Warning: this intentionally resets the Config sheet.
 */
function initializeConfigTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(QA_BI_DAILY_CONFIG_SHEET);

  if (!configSheet) {
    configSheet = ss.insertSheet(QA_BI_DAILY_CONFIG_SHEET, 0);
    configSheet.setTabColor('#4285F4');
  }

  configSheet.clear();
  configSheet.getRange('A1:D1').merge();
  configSheet.getRange('A1').setValue('⚙️ QA BI-DAILY CONFIGURATION')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');

  appendProjectConfigBlock_(configSheet, {
    projectName: 'SIPGN',
    mondayTime: '09:00',
    wednesdayTime: '09:00',
    fridayTime: '09:00',
    summaryTime: '17:00',
    googleMeetLink: 'https://meet.google.com/xxx-xxxx-xxx'
  }, [
    ['Alice', '628123456789'],
    ['Bob', '628129876543'],
    ['Carol', '628131234567']
  ]);

  appendProjectConfigBlock_(configSheet, {
    projectName: 'INADigital/Internal',
    mondayTime: '14:00',
    wednesdayTime: '14:00',
    fridayTime: '14:00',
    summaryTime: '18:00',
    googleMeetLink: 'https://meet.google.com/yyy-yyyy-yyy'
  }, [
    ['David', '628987654321'],
    ['Eve', '628912345678'],
    ['Frank', '628998877665']
  ]);

  appendConfigNotes_(configSheet);
  formatConfigSheet_(configSheet);

  SpreadsheetApp.getUi().alert(
    '✅ QA Bi-Daily Config initialized.\n\n' +
    'Two starter projects were created. Use:\n' +
    'QA Bi-Daily → Setup → Add Project\n\n' +
    'Update team members, WhatsApp Group ID, Fonnte Token, and schedule before setting up triggers.'
  );
}

/**
 * Add another project block without resetting existing configuration.
 */
function addProjectConfig() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Add QA Bi-Daily Project',
    'Enter a unique project name:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const projectName = response.getResponseText().trim();
  if (!projectName) {
    ui.alert('❌ Project name cannot be empty.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(QA_BI_DAILY_CONFIG_SHEET);
  if (!configSheet) {
    ui.alert('❌ Config sheet not found. Initialize Config first.');
    return;
  }

  if (getProjectNames().some(name => name.toLowerCase() === projectName.toLowerCase())) {
    ui.alert(`❌ Project "${projectName}" already exists.`);
    return;
  }

  removeConfigNotes_(configSheet);
  appendProjectConfigBlock_(configSheet, { projectName: projectName }, []);
  appendConfigNotes_(configSheet);
  formatConfigSheet_(configSheet);

  ui.alert(
    `✅ Project "${projectName}" added.\n\n` +
    'Complete its configuration, initialize standup sheets, then run Setup Auto Triggers again.'
  );
}

/**
 * Return every configured project name.
 */
function getProjectNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(QA_BI_DAILY_CONFIG_SHEET);
  if (!configSheet || configSheet.getLastRow() < 1) return [];

  const values = configSheet.getRange(1, 1, configSheet.getLastRow(), 2).getValues();
  return values
    .filter(row => String(row[0]).trim() === 'Project Name' && String(row[1]).trim())
    .map(row => String(row[1]).trim());
}

/**
 * Read one project block by project name.
 */
function getProjectConfig(projectName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(QA_BI_DAILY_CONFIG_SHEET);
  if (!configSheet) {
    throw new Error('Config sheet not found. Run initializeConfigTab() first.');
  }

  const lastRow = configSheet.getLastRow();
  const values = configSheet.getRange(1, 1, lastRow, 2).getValues();
  let settingsStartIndex = -1;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === 'Project Name' &&
        String(values[i][1]).trim().toLowerCase() === String(projectName).trim().toLowerCase()) {
      settingsStartIndex = i;
      break;
    }
  }

  if (settingsStartIndex < 0) {
    throw new Error(`Project "${projectName}" not found in Config.`);
  }

  const settings = {};
  let teamHeaderIndex = -1;
  for (let i = settingsStartIndex; i < values.length; i++) {
    const label = String(values[i][0]).trim();
    if (i > settingsStartIndex && label === 'Project Name') break;
    if (label === 'Name' && String(values[i][1]).trim().indexOf('Phone Number') === 0) {
      teamHeaderIndex = i;
      break;
    }
    if (QA_BI_DAILY_SETTINGS.indexOf(label) >= 0) {
      settings[label] = values[i][1];
    }
  }

  const teamMembers = [];
  const phoneNumbers = [];
  if (teamHeaderIndex >= 0) {
    for (let i = teamHeaderIndex + 1; i < values.length; i++) {
      const name = String(values[i][0] || '').trim();
      const phone = String(values[i][1] || '').trim();
      if (!name && !phone) {
        if (i + 1 >= values.length || !values[i + 1][0]) break;
        continue;
      }
      if (name === 'Project Name' || name === '📝 NOTES') break;
      if (name) {
        teamMembers.push(name);
        phoneNumbers.push(phone);
      }
    }
  }

  return {
    projectName: String(settings['Project Name'] || projectName).trim(),
    teamMembers: teamMembers,
    phoneNumbers: phoneNumbers,
    standupDays: String(settings['Standup Days'] || 'Monday, Wednesday, Friday').trim(),
    mondayTime: extractConfigTime_(settings['Monday Time']),
    wednesdayTime: extractConfigTime_(settings['Wednesday Time']),
    fridayTime: extractConfigTime_(settings['Friday Time']),
    reminderOffset: parseInt(settings['Reminder Offset (minutes)'], 10) || 30,
    summaryTime: extractConfigTime_(settings['Summary Time']),
    googleMeetLink: String(settings['Google Meet Link'] || '').trim(),
    whatsappGroupId: String(settings['WhatsApp Group ID'] || '').trim(),
    fontteToken: String(settings['Fonnte Token'] || '').trim(),
    enableReminder: String(settings['Enable Reminder']).toUpperCase() === 'TRUE',
    enableSummary: String(settings['Enable Summary']).toUpperCase() === 'TRUE'
  };
}

function validateProjectConfig(config) {
  const errors = [];
  if (!config.teamMembers.length) errors.push('Team Members cannot be empty');

  getConfiguredStandupDays_(config).forEach(dayName => {
    const time = getStandupTimeForDay(config, dayName);
    if (!time || !time.match(/^\d{1,2}:\d{2}$/)) {
      errors.push(`${dayName} Time must be in HH:MM format`);
    }
  });

  if (config.enableSummary && (!config.summaryTime || !config.summaryTime.match(/^\d{1,2}:\d{2}$/))) {
    errors.push('Summary Time must be in HH:MM format');
  }

  if (config.enableReminder || config.enableSummary) {
    if (!config.whatsappGroupId || !config.whatsappGroupId.includes('@g.us')) {
      errors.push('WhatsApp Group ID must be in format 120363xxx@g.us');
    }
    if (!config.fontteToken) errors.push('Fonnte Token cannot be empty');
  }

  return { valid: errors.length === 0, errors: errors };
}

function getConfiguredStandupDays_(config) {
  return String(config.standupDays || '')
    .split(',')
    .map(day => day.trim())
    .filter(day => ['Monday', 'Wednesday', 'Friday'].includes(day));
}

function isConfiguredStandupDay_(config, dayName) {
  return getConfiguredStandupDays_(config).includes(dayName);
}

function getStandupTimeForDay(config, dayName) {
  const timeByDay = {
    Monday: config.mondayTime,
    Wednesday: config.wednesdayTime,
    Friday: config.fridayTime
  };
  return timeByDay[dayName] || null;
}

function appendProjectConfigBlock_(sheet, config, teamMembers) {
  const startRow = Math.max(sheet.getLastRow() + 2, 3);
  const settings = [
    ['Project Name', config.projectName],
    ['Standup Days', config.standupDays || 'Monday, Wednesday, Friday'],
    ['Monday Time', config.mondayTime || '09:00'],
    ['Wednesday Time', config.wednesdayTime || '09:00'],
    ['Friday Time', config.fridayTime || '09:00'],
    ['Reminder Offset (minutes)', String(config.reminderOffset || 30)],
    ['Summary Time', config.summaryTime || '17:00'],
    ['Google Meet Link', config.googleMeetLink || ''],
    ['WhatsApp Group ID', config.whatsappGroupId || ''],
    ['Fonnte Token', config.fontteToken || ''],
    ['Enable Reminder', config.enableReminder === false ? 'FALSE' : 'TRUE'],
    ['Enable Summary', config.enableSummary === false ? 'FALSE' : 'TRUE']
  ];

  sheet.getRange(startRow, 1, 1, 4).merge();
  sheet.getRange(startRow, 1).setValue(`📋 ${config.projectName.toUpperCase()} SETTINGS`)
    .setFontSize(12).setFontWeight('bold').setBackground('#E8F0FE').setFontColor('#1A73E8');
  sheet.getRange(startRow + 1, 1, settings.length, 2).setValues(settings);
  sheet.getRange(startRow + 1, 1, settings.length, 1).setFontWeight('bold').setBackground('#F8F9FA');
  sheet.getRange(startRow + 3, 2, 3, 1).setNumberFormat('@STRING@');

  const trueFalseRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true).setAllowInvalid(false).build();
  sheet.getRange(startRow + 11, 2, 2, 1).setDataValidation(trueFalseRule);

  const teamTitleRow = startRow + settings.length + 3;
  sheet.getRange(teamTitleRow, 1, 1, 2).merge();
  sheet.getRange(teamTitleRow, 1).setValue('👥 TEAM MEMBERS')
    .setFontSize(11).setFontWeight('bold').setBackground('#D9EAD3').setFontColor('#38761D');
  sheet.getRange(teamTitleRow + 1, 1, 1, 2).setValues([['Name', 'Phone Number (62xxx)']])
    .setFontWeight('bold').setBackground('#F3F3F3')
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  const rows = teamMembers.slice();
  while (rows.length < 20) rows.push(['', '']);
  sheet.getRange(teamTitleRow + 2, 1, rows.length, 2).setValues(rows)
    .setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
}

function appendConfigNotes_(sheet) {
  const row = sheet.getLastRow() + 2;
  sheet.getRange(row, 1, 1, 4).merge();
  sheet.getRange(row, 1).setValue('📝 NOTES')
    .setFontSize(12).setFontWeight('bold').setBackground('#FFF3E0').setFontColor('#E65100');

  const notes = [
    '• Add projects from QA Bi-Daily → Setup → Add Project.',
    '• Project names must be unique and are used as standup sheet names.',
    '• Empty team rows are skipped; insert more rows before the next project when needed.',
    '• Supported standup days: Monday, Wednesday, Friday.',
    '• After changing projects or schedules, run Setup Auto Triggers again.',
    '• Phone format: 62xxx. WhatsApp Group ID format: 120363xxx@g.us.'
  ];
  notes.forEach((note, index) => {
    sheet.getRange(row + 1 + index, 1, 1, 4).merge();
    sheet.getRange(row + 1 + index, 1).setValue(note).setFontSize(9).setFontColor('#666666');
  });
}

function removeConfigNotes_(sheet) {
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const notesIndex = values.findIndex(row => String(row[0]).trim() === '📝 NOTES');
  if (notesIndex >= 0) {
    sheet.deleteRows(notesIndex + 1, sheet.getLastRow() - notesIndex);
  }
}

function formatConfigSheet_(sheet) {
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 150);
  sheet.setFrozenRows(1);
}

function extractConfigTime_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  const text = String(value || '').trim();
  if (/^\d{4}$/.test(text)) return `${text.slice(0, 2)}:${text.slice(2)}`;
  return text;
}
