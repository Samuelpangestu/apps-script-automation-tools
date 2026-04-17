/**
 * Triggers.js — Auto Sync & Refresh Dashboard Scheduler
 * ═══════════════════════════════════════════════════════════════════════
 * Configurable scheduled sync from production + dashboard refresh
 * ═══════════════════════════════════════════════════════════════════════
 */

const TRIGGER_FUNCTION_NAME = 'autoSyncAndRefresh';
const TRIGGER_PROPERTY_KEY = 'AUTO_REFRESH_ENABLED';

/**
 * DEPRECATED: Use autoSyncAndRefresh() instead
 * Kept for backward compatibility with existing triggers
 */
function autoRefreshDashboard() {
  autoSyncAndRefresh();
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
  const syncSettings = getSyncSettings();

  let message = 'AUTO SYNC & REFRESH CONFIGURATION\n\n';

  // Show sync status
  if (syncSettings.enabled && syncSettings.spreadsheetId) {
    message += 'Production Sync: ENABLED ✅\n';
    message += 'Source: ' + syncSettings.spreadsheetName + '\n\n';
  } else {
    message += 'Production Sync: DISABLED (Local mode)\n\n';
  }

  // Show auto-refresh status
  if (currentSetting.enabled) {
    message += 'Auto Sync & Refresh: ENABLED ✅\n';
    message += 'Interval: Every ' + currentSetting.hours + ' hours\n\n';
    message += 'System will automatically sync from production and refresh dashboard every ' + currentSetting.hours + ' hours.\n\n';
  } else {
    message += 'Auto Sync & Refresh: DISABLED ❌\n\n';
    message += 'Dashboard will NOT auto-refresh.\n\n';
  }

  message += 'What would you like to do?';

  const response = ui.alert(
    'Auto Sync & Refresh Dashboard',
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
      ui.alert('Success', 'Auto sync & refresh has been disabled.', ui.ButtonSet.OK);
    } else {
      ui.alert('Info', 'Auto sync & refresh is already disabled.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Show interval selection dialog
 */
function showIntervalDialog() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Select Sync & Refresh Interval',
    'How often should the system auto-sync and refresh?\n\n' +
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
    const syncSettings = getSyncSettings();

    let successMessage = 'Auto sync & refresh has been enabled.\n\n' +
      'Interval: Every ' + hours + ' hours\n\n';

    if (syncSettings.enabled && syncSettings.spreadsheetId) {
      successMessage += 'System will automatically:\n' +
        '1. Sync data from production (' + syncSettings.spreadsheetName + ')\n' +
        '2. Refresh dashboard\n\n';
    } else {
      successMessage += 'System will automatically refresh dashboard.\n' +
        '(Production sync disabled - using local data)\n\n';
    }

    successMessage += 'Next sync in ' + hours + ' hour(s).';

    ui.alert('Success! ✅', successMessage, ui.ButtonSet.OK);
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
