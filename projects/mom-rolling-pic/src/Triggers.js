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
 * - SIPGN: Reminder + Summary
 * - INADigital/Internal: Reminder + Summary
 */
function setupAllTriggers() {
  // Remove existing triggers first
  removeAllTriggers();

  const configA = getProjectConfig('SIPGN');
  const configB = getProjectConfig('INADigital/Internal');

  // Validate configs
  const validationA = validateProjectConfig(configA);
  const validationB = validateProjectConfig(configB);

  if (!validationA.valid) {
    throw new Error(`SIPGN config invalid: ${validationA.errors.join(', ')}`);
  }

  if (!validationB.valid) {
    throw new Error(`INADigital/Internal config invalid: ${validationB.errors.join(', ')}`);
  }

  // Setup triggers for each project
  if (configA.enableReminder) {
    setupReminderTriggers('SIPGN', configA);
  }

  if (configA.enableSummary) {
    setupSummaryTrigger('SIPGN', configA);
  }

  if (configB.enableReminder) {
    setupReminderTriggers('INADigital/Internal', configB);
  }

  if (configB.enableSummary) {
    setupSummaryTrigger('INADigital/Internal', configB);
  }

  Logger.log('✅ All triggers setup successfully');
}

/**
 * Setup reminder triggers for a project (Mon, Wed, Fri)
 * @param {string} projectName - 'SIPGN' or 'INADigital/Internal'
 * @param {Object} config - Project config
 */
function setupReminderTriggers(projectName, config) {
  // Map project names to valid function suffixes
  const projectKey = projectName === 'SIPGN' ? 'ProjectA' : 'ProjectB';

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
 * @param {string} projectName - 'SIPGN' or 'INADigital/Internal'
 * @param {Object} config - Project config
 */
function setupSummaryTrigger(projectName, config) {
  // Map project names to valid function suffixes
  const projectKey = projectName === 'SIPGN' ? 'ProjectA' : 'ProjectB';

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
 * Monday reminder for SIPGN
 */
function triggerMondayReminderProjectA() {
  Logger.log('📅 Monday Reminder - SIPGN');
  sendStandupReminder('SIPGN');
}

/**
 * Monday reminder for INADigital/Internal
 */
function triggerMondayReminderProjectB() {
  Logger.log('📅 Monday Reminder - INADigital/Internal');
  sendStandupReminder('INADigital/Internal');
}

/**
 * Wednesday reminder for SIPGN
 */
function triggerWednesdayReminderProjectA() {
  Logger.log('📅 Wednesday Reminder - SIPGN');
  sendStandupReminder('SIPGN');
}

/**
 * Wednesday reminder for INADigital/Internal
 */
function triggerWednesdayReminderProjectB() {
  Logger.log('📅 Wednesday Reminder - INADigital/Internal');
  sendStandupReminder('INADigital/Internal');
}

/**
 * Friday reminder for SIPGN
 */
function triggerFridayReminderProjectA() {
  Logger.log('📅 Friday Reminder - SIPGN');
  sendStandupReminder('SIPGN');
}

/**
 * Friday reminder for INADigital/Internal
 */
function triggerFridayReminderProjectB() {
  Logger.log('📅 Friday Reminder - INADigital/Internal');
  sendStandupReminder('INADigital/Internal');
}

/**
 * Daily summary for SIPGN
 */
function triggerSummaryProjectA() {
  Logger.log('📊 Daily Summary - SIPGN');

  const today = new Date();
  const dayName = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE');

  // Only send summary on standup days
  if (['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
    sendStandupSummary('SIPGN');
  } else {
    Logger.log(`⚠️ Today is ${dayName}, not a standup day. Skipping summary.`);
  }
}

/**
 * Daily summary for INADigital/Internal
 */
function triggerSummaryProjectB() {
  Logger.log('📊 Daily Summary - INADigital/Internal');

  const today = new Date();
  const dayName = Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE');

  // Only send summary on standup days
  if (['Monday', 'Wednesday', 'Friday'].includes(dayName)) {
    sendStandupSummary('INADigital/Internal');
  } else {
    Logger.log(`⚠️ Today is ${dayName}, not a standup day. Skipping summary.`);
  }
}
