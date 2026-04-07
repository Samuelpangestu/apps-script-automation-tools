/**
 * Triggers.js
 * Manage time-based triggers for bi-daily standup
 * - Auto reminder before standup
 * - Auto summary after standup
 * - Auto generate rows (triggered with reminder)
 */

/**
 * Setup all auto triggers for both projects
 * Creates time-based triggers for:
 * - Project A: Reminder + Summary
 * - Project B: Reminder + Summary
 */
function setupAllTriggers() {
  // Remove existing triggers first
  removeAllTriggers();

  const configA = getProjectConfig('Project A');
  const configB = getProjectConfig('Project B');

  // Validate configs
  const validationA = validateProjectConfig(configA);
  const validationB = validateProjectConfig(configB);

  if (!validationA.valid) {
    throw new Error(`Project A config invalid: ${validationA.errors.join(', ')}`);
  }

  if (!validationB.valid) {
    throw new Error(`Project B config invalid: ${validationB.errors.join(', ')}`);
  }

  // Setup triggers for each project
  if (configA.enableReminder) {
    setupReminderTriggers('Project A', configA);
  }

  if (configA.enableSummary) {
    setupSummaryTrigger('Project A', configA);
  }

  if (configB.enableReminder) {
    setupReminderTriggers('Project B', configB);
  }

  if (configB.enableSummary) {
    setupSummaryTrigger('Project B', configB);
  }

  Logger.log('✅ All triggers setup successfully');
}

/**
 * Setup reminder triggers for a project (Mon, Wed, Fri)
 * @param {string} projectName - 'Project A' or 'Project B'
 * @param {Object} config - Project config
 */
function setupReminderTriggers(projectName, config) {
  const projectKey = projectName.replace(' ', ''); // 'ProjectA' or 'ProjectB'

  // Monday reminder
  const mondayTimes = calculateReminderTime(config.mondayTime, config.reminderOffset);
  ScriptApp.newTrigger(`triggerMondayReminder${projectKey}`)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(mondayTimes.hour)
    .nearMinute(mondayTimes.minute)
    .create();

  Logger.log(`✅ Monday reminder trigger created for ${projectName} at ${mondayTimes.hour}:${mondayTimes.minute}`);

  // Wednesday reminder
  const wednesdayTimes = calculateReminderTime(config.wednesdayTime, config.reminderOffset);
  ScriptApp.newTrigger(`triggerWednesdayReminder${projectKey}`)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(wednesdayTimes.hour)
    .nearMinute(wednesdayTimes.minute)
    .create();

  Logger.log(`✅ Wednesday reminder trigger created for ${projectName} at ${wednesdayTimes.hour}:${wednesdayTimes.minute}`);

  // Friday reminder
  const fridayTimes = calculateReminderTime(config.fridayTime, config.reminderOffset);
  ScriptApp.newTrigger(`triggerFridayReminder${projectKey}`)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(fridayTimes.hour)
    .nearMinute(fridayTimes.minute)
    .create();

  Logger.log(`✅ Friday reminder trigger created for ${projectName} at ${fridayTimes.hour}:${fridayTimes.minute}`);
}

/**
 * Setup summary trigger for a project (daily at configured time)
 * @param {string} projectName - 'Project A' or 'Project B'
 * @param {Object} config - Project config
 */
function setupSummaryTrigger(projectName, config) {
  const projectKey = projectName.replace(' ', ''); // 'ProjectA' or 'ProjectB'

  const times = parseTime(config.summaryTime);

  ScriptApp.newTrigger(`triggerSummary${projectKey}`)
    .timeBased()
    .everyDays(1)
    .atHour(times.hour)
    .nearMinute(times.minute)
    .create();

  Logger.log(`✅ Summary trigger created for ${projectName} at ${times.hour}:${times.minute}`);
}

/**
 * Calculate reminder time (standup time - offset minutes)
 * @param {string} standupTime - Time in HH:MM format (e.g., "09:00")
 * @param {number} offsetMinutes - Minutes before standup (e.g., 30)
 * @returns {Object} {hour, minute}
 */
function calculateReminderTime(standupTime, offsetMinutes) {
  const times = parseTime(standupTime);

  // Calculate reminder time
  let reminderMinutes = times.hour * 60 + times.minute - offsetMinutes;

  if (reminderMinutes < 0) {
    reminderMinutes += 24 * 60; // Handle wrap to previous day
  }

  const reminderHour = Math.floor(reminderMinutes / 60);
  const reminderMinute = reminderMinutes % 60;

  return {
    hour: reminderHour,
    minute: reminderMinute
  };
}

/**
 * Parse time string to hour and minute
 * @param {string} timeStr - Time in HH:MM format
 * @returns {Object} {hour, minute}
 */
function parseTime(timeStr) {
  const parts = timeStr.split(':');
  return {
    hour: parseInt(parts[0]),
    minute: parseInt(parts[1])
  };
}

/**
 * Remove all triggers for this project
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const handlerFunction = trigger.getHandlerFunction();

    // Only remove triggers related to bi-daily standup
    if (handlerFunction.startsWith('triggerMonday') ||
        handlerFunction.startsWith('triggerWednesday') ||
        handlerFunction.startsWith('triggerFriday') ||
        handlerFunction.startsWith('triggerSummary')) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`🗑️ Removed trigger: ${handlerFunction}`);
    }
  });

  Logger.log('✅ All standup triggers removed');
}

// ═══════════════════════════════════════════════════════════════
// TRIGGER HANDLER FUNCTIONS (Called by time-based triggers)
// ═══════════════════════════════════════════════════════════════

/**
 * Monday reminder for Project A
 */
function triggerMondayReminderProjectA() {
  Logger.log('📅 Monday Reminder - Project A');
  sendStandupReminder('Project A');
}

/**
 * Monday reminder for Project B
 */
function triggerMondayReminderProjectB() {
  Logger.log('📅 Monday Reminder - Project B');
  sendStandupReminder('Project B');
}

/**
 * Wednesday reminder for Project A
 */
function triggerWednesdayReminderProjectA() {
  Logger.log('📅 Wednesday Reminder - Project A');
  sendStandupReminder('Project A');
}

/**
 * Wednesday reminder for Project B
 */
function triggerWednesdayReminderProjectB() {
  Logger.log('📅 Wednesday Reminder - Project B');
  sendStandupReminder('Project B');
}

/**
 * Friday reminder for Project A
 */
function triggerFridayReminderProjectA() {
  Logger.log('📅 Friday Reminder - Project A');
  sendStandupReminder('Project A');
}

/**
 * Friday reminder for Project B
 */
function triggerFridayReminderProjectB() {
  Logger.log('📅 Friday Reminder - Project B');
  sendStandupReminder('Project B');
}

/**
 * Daily summary for Project A
 */
function triggerSummaryProjectA() {
  Logger.log('📊 Daily Summary - Project A');

  const today = new Date();
  const dayName = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE');

  // Only send summary on standup days
  if (['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
    sendStandupSummary('Project A');
  } else {
    Logger.log(`⚠️ Today is ${dayName}, not a standup day. Skipping summary.`);
  }
}

/**
 * Daily summary for Project B
 */
function triggerSummaryProjectB() {
  Logger.log('📊 Daily Summary - Project B');

  const today = new Date();
  const dayName = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE');

  // Only send summary on standup days
  if (['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
    sendStandupSummary('Project B');
  } else {
    Logger.log(`⚠️ Today is ${dayName}, not a standup day. Skipping summary.`);
  }
}
