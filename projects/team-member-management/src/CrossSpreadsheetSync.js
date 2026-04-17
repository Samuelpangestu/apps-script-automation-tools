/**
 * CrossSpreadsheetSync.js — Sync Config from Production Spreadsheet
 * ═══════════════════════════════════════════════════════════════════════
 * Pull Config data from production spreadsheet and refresh dashboard
 * ═══════════════════════════════════════════════════════════════════════
 */

const SYNC_SETTINGS_KEY = 'PRODUCTION_SYNC_SETTINGS';

/**
 * Get production sync settings
 */
function getSyncSettings() {
  const props = PropertiesService.getDocumentProperties();
  const settings = props.getProperty(SYNC_SETTINGS_KEY);

  if (!settings) {
    return {
      enabled: false,
      spreadsheetId: '',
      spreadsheetName: ''
    };
  }

  return JSON.parse(settings);
}

/**
 * Save production sync settings
 */
function saveSyncSettings(spreadsheetId, spreadsheetName) {
  const settings = {
    enabled: true,
    spreadsheetId: spreadsheetId,
    spreadsheetName: spreadsheetName
  };

  PropertiesService.getDocumentProperties()
    .setProperty(SYNC_SETTINGS_KEY, JSON.stringify(settings));

  Logger.log('✅ Sync settings saved: ' + spreadsheetName);
  return true;
}

/**
 * Disable production sync
 */
function disableSyncSettings() {
  const settings = {
    enabled: false,
    spreadsheetId: '',
    spreadsheetName: ''
  };

  PropertiesService.getDocumentProperties()
    .setProperty(SYNC_SETTINGS_KEY, JSON.stringify(settings));

  Logger.log('✅ Sync settings disabled');
  return true;
}

/**
 * Sync data from production spreadsheet
 * Returns: { success: boolean, message: string, synced: { projects, moduls, submoduls } }
 */
function syncFromProduction() {
  const settings = getSyncSettings();

  if (!settings.enabled || !settings.spreadsheetId) {
    Logger.log('⚠️ Production sync not configured');
    return {
      success: false,
      message: 'Production sync not configured. Please configure in Config tab.',
      synced: { projects: 0, moduls: 0, submoduls: 0 }
    };
  }

  try {
    Logger.log('🔄 Syncing from production: ' + settings.spreadsheetName);

    // Open production spreadsheet
    const prodSS = SpreadsheetApp.openById(settings.spreadsheetId);
    const prodConfig = prodSS.getSheetByName(CONFIG_TAB_NAME);

    if (!prodConfig) {
      throw new Error('Config tab not found in production spreadsheet');
    }

    // Open local spreadsheet
    const localSS = SpreadsheetApp.getActiveSpreadsheet();
    const localConfig = localSS.getSheetByName(CONFIG_TAB_NAME);

    if (!localConfig) {
      throw new Error('Config tab not found in local spreadsheet');
    }

    // Sync Projects (fixed range: A10:D20 = 2 sample + 5 blank + buffer)
    const projectRange = 'A' + PROJECT_DATA_START_ROW + ':D' + (PROJECT_DATA_START_ROW + 9);
    const projectData = prodConfig.getRange(projectRange).getValues();
    localConfig.getRange(projectRange).setValues(projectData);

    // Count synced projects
    const projectCount = projectData.filter(row => row[1] && row[1].toString().trim()).length;

    // Sync Moduls (fixed range: A15:E35 = 2 sample + 15 blank + buffer)
    const modulRange = 'A' + MODUL_DATA_START_ROW + ':E' + (MODUL_DATA_START_ROW + 19);
    const modulData = prodConfig.getRange(modulRange).getValues();
    localConfig.getRange(modulRange).setValues(modulData);

    // Count synced moduls
    const modulCount = modulData.filter(row => row[1] && row[1].toString().trim()).length;

    // Sync Submoduls (dynamic range - get all data)
    const submodulLastRow = prodConfig.getLastRow();
    const submodulRange = 'A' + SUBMODUL_DATA_START_ROW + ':H' + Math.max(SUBMODUL_DATA_START_ROW + 40, submodulLastRow);
    const submodulData = prodConfig.getRange(submodulRange).getValues();

    // Clear local submodul area first
    localConfig.getRange(submodulRange).clearContent();

    // Write synced data
    localConfig.getRange(submodulRange).setValues(submodulData);

    // Count synced submoduls
    const submodulCount = submodulData.filter(row => row[1] && row[1].toString().trim()).length;

    // Copy formatting from production (optional - preserve colors)
    copyFormattingFromProduction(prodConfig, localConfig);

    Logger.log('✅ Sync complete: ' + projectCount + ' projects, ' + modulCount + ' moduls, ' + submodulCount + ' submoduls');

    return {
      success: true,
      message: 'Synced from ' + settings.spreadsheetName,
      synced: {
        projects: projectCount,
        moduls: modulCount,
        submoduls: submodulCount
      }
    };

  } catch (error) {
    Logger.log('❌ Sync error: ' + error.message);
    return {
      success: false,
      message: 'Sync failed: ' + error.message,
      synced: { projects: 0, moduls: 0, submoduls: 0 }
    };
  }
}

/**
 * Copy formatting from production (preserve zebra stripes)
 */
function copyFormattingFromProduction(prodConfig, localConfig) {
  try {
    // Copy project formatting
    const projectFmt = prodConfig.getRange('A' + PROJECT_DATA_START_ROW + ':D' + (PROJECT_DATA_START_ROW + 9));
    projectFmt.copyFormatToRange(localConfig, 1, 4, PROJECT_DATA_START_ROW, PROJECT_DATA_START_ROW + 9);

    // Copy modul formatting
    const modulFmt = prodConfig.getRange('A' + MODUL_DATA_START_ROW + ':E' + (MODUL_DATA_START_ROW + 19));
    modulFmt.copyFormatToRange(localConfig, 1, 5, MODUL_DATA_START_ROW, MODUL_DATA_START_ROW + 19);

    // Copy submodul formatting
    const submodulLastRow = prodConfig.getLastRow();
    const submodulFmt = prodConfig.getRange('A' + SUBMODUL_DATA_START_ROW + ':H' + Math.max(SUBMODUL_DATA_START_ROW + 40, submodulLastRow));
    submodulFmt.copyFormatToRange(localConfig, 1, 8, SUBMODUL_DATA_START_ROW, Math.max(SUBMODUL_DATA_START_ROW + 40, submodulLastRow));

    Logger.log('✅ Formatting copied from production');
  } catch (error) {
    Logger.log('⚠️ Formatting copy skipped: ' + error.message);
  }
}

/**
 * Auto sync + refresh (called by trigger)
 * This combines sync from production and dashboard refresh
 */
function autoSyncAndRefresh() {
  try {
    Logger.log('🔄 Auto sync & refresh started...');

    const settings = getSyncSettings();

    // Sync from production if enabled
    if (settings.enabled && settings.spreadsheetId) {
      const syncResult = syncFromProduction();

      if (syncResult.success) {
        Logger.log('✅ Synced: ' + syncResult.synced.projects + ' projects, ' +
                   syncResult.synced.moduls + ' moduls, ' +
                   syncResult.synced.submoduls + ' submoduls');
      } else {
        Logger.log('⚠️ Sync skipped: ' + syncResult.message);
      }
    } else {
      Logger.log('ℹ️ Production sync disabled, skipping sync step');
    }

    // Refresh dashboard
    createDashboard();
    Logger.log('✅ Dashboard refreshed');

    Logger.log('✅ Auto sync & refresh complete');

  } catch (error) {
    Logger.log('❌ Auto sync & refresh error: ' + error.message);
  }
}

/**
 * Test sync connection (verify production spreadsheet access)
 */
function testSyncConnection() {
  const settings = getSyncSettings();

  if (!settings.enabled || !settings.spreadsheetId) {
    return {
      success: false,
      message: 'Production sync not configured'
    };
  }

  try {
    const prodSS = SpreadsheetApp.openById(settings.spreadsheetId);
    const prodConfig = prodSS.getSheetByName(CONFIG_TAB_NAME);

    if (!prodConfig) {
      return {
        success: false,
        message: 'Config tab not found in production spreadsheet'
      };
    }

    const prodName = prodSS.getName();

    return {
      success: true,
      message: 'Connected to: ' + prodName,
      spreadsheetName: prodName
    };

  } catch (error) {
    return {
      success: false,
      message: 'Connection failed: ' + error.message
    };
  }
}
