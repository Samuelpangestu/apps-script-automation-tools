/**
 * CrossSpreadsheetSync.js — Sync Config from Production Spreadsheet
 * ═══════════════════════════════════════════════════════════════════════
 * Pull Config data from production spreadsheet and refresh dashboard
 * ═══════════════════════════════════════════════════════════════════════
 */

const SYNC_SETTINGS_KEY = 'PRODUCTION_SYNC_SETTINGS';
const PROD_ID_ROW = 2; // Row 2 in Config tab
const PROD_ID_COL = 2; // Column B in Config tab

/**
 * Read Production Spreadsheet ID from Config tab
 * Returns the ID from cell B2 in Config tab
 */
function getProductionIdFromConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!configSheet) {
    return '';
  }

  let prodId = configSheet.getRange(PROD_ID_ROW, PROD_ID_COL).getValue();

  if (!prodId || typeof prodId !== 'string') {
    return '';
  }

  prodId = prodId.toString().trim();

  // Skip placeholder text
  if (prodId.includes('Not configured') || prodId.includes('Paste production')) {
    return '';
  }

  // Extract ID from URL if user pasted full URL
  if (prodId.includes('docs.google.com/spreadsheets')) {
    const match = prodId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      return match[1];
    }
  }

  return prodId;
}

/**
 * Get production sync settings
 * Auto-reads from Config tab if not in properties
 */
function getSyncSettings() {
  // First try to read from Config tab
  const prodIdFromConfig = getProductionIdFromConfig();

  if (prodIdFromConfig) {
    // Auto-save to properties if found in Config tab
    try {
      const testSS = SpreadsheetApp.openById(prodIdFromConfig);
      const spreadsheetName = testSS.getName();

      // Save to properties for caching
      const settings = {
        enabled: true,
        spreadsheetId: prodIdFromConfig,
        spreadsheetName: spreadsheetName
      };

      PropertiesService.getDocumentProperties()
        .setProperty(SYNC_SETTINGS_KEY, JSON.stringify(settings));

      return settings;
    } catch (error) {
      Logger.log('⚠️ Could not access production spreadsheet from Config tab: ' + error.message);
    }
  }

  // Fallback to properties
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

    // Sync flat table data (much simpler now!)
    const prodLastRow = prodConfig.getLastRow();
    const dataRowCount = Math.max(60, prodLastRow - CONFIG_DATA_START_ROW + 1); // Min 60 rows

    const dataRange = 'A' + CONFIG_DATA_START_ROW + ':J' + (CONFIG_DATA_START_ROW + dataRowCount - 1);  // Now includes PIC QA column
    const configData = prodConfig.getRange(dataRange).getValues();

    // Clear local data area first
    localConfig.getRange(dataRange).clearContent();

    // Write synced data
    localConfig.getRange(dataRange).setValues(configData);

    // Count synced items
    const activeRows = configData.filter(row => row[0] === true && row[1] && row[2] && row[3]);
    const projectSet = new Set();
    const modulSet = new Set();
    let submodulCount = 0;

    activeRows.forEach(row => {
      const project = row[1].toString().trim();
      const modul = row[2].toString().trim();
      const submodul = row[3].toString().trim();

      if (project) projectSet.add(project);
      if (modul) modulSet.add(project + '|' + modul);
      if (submodul) submodulCount++;
    });

    // Copy formatting from production
    copyFormattingFromProduction(prodConfig, localConfig, dataRange);

    Logger.log('✅ Sync complete: ' + projectSet.size + ' projects, ' + modulSet.size + ' moduls, ' + submodulCount + ' submoduls');

    return {
      success: true,
      message: 'Synced from ' + settings.spreadsheetName,
      synced: {
        projects: projectSet.size,
        moduls: modulSet.size,
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
function copyFormattingFromProduction(prodConfig, localConfig, dataRange) {
  try {
    const fmt = prodConfig.getRange(dataRange);
    fmt.copyFormatToRange(localConfig, 1, 10, CONFIG_DATA_START_ROW, CONFIG_DATA_START_ROW + 59);  // Now 10 columns (A-J)

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
