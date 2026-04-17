/**
 * Triggers.js — Auto-refresh Dashboard Scheduler
 * ═══════════════════════════════════════════════════════════════════════
 * Configurable scheduled dashboard refresh
 * ═══════════════════════════════════════════════════════════════════════
 */

const TRIGGER_FUNCTION_NAME = 'autoRefreshDashboard';
const TRIGGER_PROPERTY_KEY = 'AUTO_REFRESH_ENABLED';

/**
 * Auto-refresh dashboard (called by trigger)
 */
function autoRefreshDashboard() {
  try {
    Logger.log('🔄 Auto-refreshing dashboard...');
    createDashboard();
    Logger.log('✅ Dashboard auto-refreshed successfully');
  } catch (error) {
    Logger.log('❌ Error auto-refreshing dashboard: ' + error.message);
  }
}

/**
 * Setup auto-refresh trigger
 * @param {number} hours - Hours between refreshes (1, 3, 6, 12, 24)
 */
function setupAutoRefresh(hours) {
  // Remove existing triggers first
  removeAutoRefresh();

  // Create new trigger
  ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
    .timeBased()
    .everyHours(hours)
    .create();

  // Save setting
  PropertiesService.getDocumentProperties().setProperty(TRIGGER_PROPERTY_KEY, hours.toString());

  Logger.log('✅ Auto-refresh enabled: every ' + hours + ' hours');
  return true;
}

/**
 * Remove auto-refresh trigger
 */
function removeAutoRefresh() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('🗑️ Removed existing trigger');
    }
  });

  // Clear setting
  PropertiesService.getDocumentProperties().deleteProperty(TRIGGER_PROPERTY_KEY);

  Logger.log('✅ Auto-refresh disabled');
  return true;
}

/**
 * Get current auto-refresh setting
 */
function getAutoRefreshSetting() {
  const hours = PropertiesService.getDocumentProperties().getProperty(TRIGGER_PROPERTY_KEY);

  if (!hours) {
    return { enabled: false, hours: 0 };
  }

  return {
    enabled: true,
    hours: parseInt(hours)
  };
}

/**
 * Check if auto-refresh is enabled
 */
function isAutoRefreshEnabled() {
  const setting = getAutoRefreshSetting();
  return setting.enabled;
}

/**
 * Menu: Configure auto-refresh
 */
function menuConfigureAutoRefresh() {
  const ui = SpreadsheetApp.getUi();
  const currentSetting = getAutoRefreshSetting();

  let message = 'DASHBOARD AUTO-REFRESH CONFIGURATION\n\n';

  if (currentSetting.enabled) {
    message += 'Current Status: ENABLED ✅\n';
    message += 'Refresh Interval: Every ' + currentSetting.hours + ' hours\n\n';
    message += 'Dashboard will automatically refresh every ' + currentSetting.hours + ' hours.\n\n';
  } else {
    message += 'Current Status: DISABLED ❌\n\n';
    message += 'Dashboard will NOT auto-refresh.\n\n';
  }

  message += 'What would you like to do?';

  const response = ui.alert(
    'Auto-Refresh Dashboard',
    message,
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.YES) {
    // Enable or change interval
    showIntervalDialog();
  } else if (response === ui.Button.NO) {
    // Disable
    if (currentSetting.enabled) {
      removeAutoRefresh();
      ui.alert('Success', 'Auto-refresh has been disabled.', ui.ButtonSet.OK);
    } else {
      ui.alert('Info', 'Auto-refresh is already disabled.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Show interval selection dialog
 */
function showIntervalDialog() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Select Refresh Interval',
    'How often should the dashboard auto-refresh?\n\n' +
    '• Every 1 hour - Very frequent (recommended for active projects)\n' +
    '• Every 3 hours - Moderate frequency\n' +
    '• Every 6 hours - Less frequent\n' +
    '• Every 12 hours - Twice daily\n' +
    '• Every 24 hours - Once daily\n\n' +
    'Choose an interval:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) return;

  // Show input prompt
  const intervalResponse = ui.prompt(
    'Enter Interval',
    'Enter hours (1, 3, 6, 12, or 24):',
    ui.ButtonSet.OK_CANCEL
  );

  if (intervalResponse.getSelectedButton() !== ui.Button.OK) return;

  const hours = parseInt(intervalResponse.getResponseText().trim());

  if (![1, 3, 6, 12, 24].includes(hours)) {
    ui.alert('Invalid Input', 'Please enter: 1, 3, 6, 12, or 24', ui.ButtonSet.OK);
    return;
  }

  try {
    setupAutoRefresh(hours);
    ui.alert(
      'Success! ✅',
      'Dashboard auto-refresh has been enabled.\n\n' +
      'Interval: Every ' + hours + ' hours\n\n' +
      'The dashboard will automatically refresh in the background.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Error', 'Failed to setup auto-refresh: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Get next refresh time
 */
function getNextRefreshTime() {
  const triggers = ScriptApp.getProjectTriggers();

  for (let i = 0; i < triggers.length; i++) {
    const trigger = triggers[i];
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      return trigger.getTriggerSource();
    }
  }

  return null;
}
