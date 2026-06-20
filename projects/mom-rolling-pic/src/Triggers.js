/**
 * Triggers.js
 * Scalable scheduler for every configured QA Bi-Daily project.
 */

const QA_BI_DAILY_SCHEDULER_HANDLER = 'runQaBiDailyScheduler';
const QA_BI_DAILY_SCHEDULER_INTERVAL_MINUTES = 5;

function setupAllTriggers() {
  removeAllTriggers();

  const projects = getProjectNames();
  if (!projects.length) throw new Error('No projects found in Config.');

  const validationErrors = [];
  projects.forEach(projectName => {
    const validation = validateProjectConfig(getProjectConfig(projectName));
    if (!validation.valid) {
      validationErrors.push(`${projectName}: ${validation.errors.join(', ')}`);
    }
  });

  if (validationErrors.length) {
    throw new Error('Invalid project configuration:\n' + validationErrors.join('\n'));
  }

  ScriptApp.newTrigger(QA_BI_DAILY_SCHEDULER_HANDLER)
    .timeBased()
    .everyMinutes(QA_BI_DAILY_SCHEDULER_INTERVAL_MINUTES)
    .create();

  Logger.log(`✅ QA Bi-Daily scheduler created for ${projects.length} project(s)`);
}

/**
 * Poll configured projects and execute actions whose scheduled time is due.
 * Script properties make each project/action idempotent per date.
 */
function runQaBiDailyScheduler() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    const now = new Date();
    const timezone = Session.getScriptTimeZone();
    const dayName = Utilities.formatDate(now, timezone, 'EEEE');
    const dateKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
    const currentHour = parseInt(Utilities.formatDate(now, timezone, 'H'), 10);
    const currentMinute = parseInt(Utilities.formatDate(now, timezone, 'm'), 10);
    const currentMinutes = currentHour * 60 + currentMinute;

    getProjectNames().forEach(projectName => {
      try {
        const config = getProjectConfig(projectName);
        if (!isConfiguredStandupDay_(config, dayName)) return;

        if (config.enableReminder) {
          const reminderTime = calculateReminderTime(
            getStandupTimeForDay(config, dayName),
            config.reminderOffset
          );
          runScheduledActionIfDue_(
            projectName,
            'reminder',
            dateKey,
            reminderTime.hour * 60 + reminderTime.minute,
            currentMinutes,
            () => sendStandupReminder(projectName)
          );
        }

        if (config.enableSummary) {
          const summaryTime = parseTime(config.summaryTime);
          runScheduledActionIfDue_(
            projectName,
            'summary',
            dateKey,
            summaryTime.hour * 60 + summaryTime.minute,
            currentMinutes,
            () => sendStandupSummary(projectName)
          );
        }
      } catch (error) {
        Logger.log(`❌ Scheduler error for ${projectName}: ${error.message}`);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function runScheduledActionIfDue_(projectName, action, dateKey, scheduledMinutes, currentMinutes, callback) {
  const elapsed = currentMinutes - scheduledMinutes;
  if (elapsed < 0 || elapsed >= QA_BI_DAILY_SCHEDULER_INTERVAL_MINUTES) return;

  const propertyKey = `qaBiDaily:${dateKey}:${projectName}:${action}`;
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(propertyKey)) return;

  callback();
  properties.setProperty(propertyKey, new Date().toISOString());
}

function calculateReminderTime(standupTime, offsetMinutes) {
  const times = parseTime(standupTime);
  let reminderMinutes = times.hour * 60 + times.minute - offsetMinutes;
  if (reminderMinutes < 0) reminderMinutes += 24 * 60;
  return {
    hour: Math.floor(reminderMinutes / 60),
    minute: reminderMinutes % 60
  };
}

function parseTime(timeStr) {
  const parts = String(timeStr).split(':');
  return { hour: parseInt(parts[0], 10), minute: parseInt(parts[1], 10) };
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === QA_BI_DAILY_SCHEDULER_HANDLER ||
        handler.startsWith('triggerMonday') ||
        handler.startsWith('triggerWednesday') ||
        handler.startsWith('triggerFriday') ||
        handler.startsWith('triggerSummary')) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`🗑️ Removed trigger: ${handler}`);
    }
  });
}
