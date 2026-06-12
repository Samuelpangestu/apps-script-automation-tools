/**
 * CrossSpreadsheetSync.js — Sync Config from QA Dashboard (SSOT)
 * ═══════════════════════════════════════════════════════════════════════
 * Pull Config data from QA Portfolio Dashboard and refresh local tabs
 * SSOT: QA Dashboard Config Sheet
 * ═══════════════════════════════════════════════════════════════════════
 */

const SYNC_SETTINGS_KEY = 'QA_DASHBOARD_SYNC_SETTINGS';
const DASHBOARD_ID_ROW = 2;  // Row 2 in Project tab (Dashboard Spreadsheet ID field at Q2)
const DASHBOARD_ID_COL = 17; // Column Q in Project tab (merged Q2:S2)

/**
 * Read QA Dashboard Spreadsheet ID from Project tab
 * Returns the ID from cell Q2 in Project tab (merged Q2:S2)
 */
function getDashboardIdFromConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);

  if (!configSheet) {
    return '';
  }

  let dashboardId = configSheet.getRange(DASHBOARD_ID_ROW, DASHBOARD_ID_COL).getValue();

  if (!dashboardId || typeof dashboardId !== 'string') {
    return '';
  }

  dashboardId = dashboardId.toString().trim();

  // Skip placeholder text
  if (dashboardId.includes('Not configured') || dashboardId.includes('Paste Dashboard')) {
    return '';
  }

  // Extract ID from URL if user pasted full URL
  if (dashboardId.includes('docs.google.com/spreadsheets')) {
    const match = dashboardId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      return match[1];
    }
  }

  return dashboardId;
}

/**
 * Get QA Dashboard sync settings (read-only, no auto-connection)
 * Reads from Project tab first, then falls back to properties
 */
function getSyncSettings() {
  // First try to read from Project tab
  const dashboardIdFromConfig = getDashboardIdFromConfig();

  if (dashboardIdFromConfig) {
    // Check if we have cached name in properties
    const props = PropertiesService.getDocumentProperties();
    const cached = props.getProperty(SYNC_SETTINGS_KEY);

    if (cached) {
      const cachedSettings = JSON.parse(cached);
      // If same ID, use cached name
      if (cachedSettings.spreadsheetId === dashboardIdFromConfig) {
        return {
          enabled: true,
          spreadsheetId: dashboardIdFromConfig,
          spreadsheetName: cachedSettings.spreadsheetName
        };
      }
    }

    // New ID, return without name (will be fetched on actual sync/test)
    return {
      enabled: true,
      spreadsheetId: dashboardIdFromConfig,
      spreadsheetName: 'QA Dashboard (not yet connected)'
    };
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
 * Save QA Dashboard sync settings
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
 * Disable QA Dashboard sync
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
 * Sync data from QA Dashboard Config Sheet
 * Returns: { success: boolean, message: string, synced: { projects, moduls, submoduls } }
 */
function syncFromDashboard() {
  const settings = getSyncSettings();

  if (!settings.enabled || !settings.spreadsheetId) {
    Logger.log('⚠️ QA Dashboard sync not configured');
    return {
      success: false,
      message: 'QA Dashboard sync not configured. Please configure in Project tab (Q2).',
      synced: { projects: 0, moduls: 0, submoduls: 0 }
    };
  }

  try {
    Logger.log('🔄 Syncing from QA Dashboard: ' + settings.spreadsheetName);

    // Open QA Dashboard spreadsheet
    const dashboardSS = SpreadsheetApp.openById(settings.spreadsheetId);
    const dashboardConfig = dashboardSS.getSheetByName('Config');

    if (!dashboardConfig) {
      throw new Error('Config tab not found in QA Dashboard spreadsheet');
    }

    // Open local spreadsheet
    const localSS = SpreadsheetApp.getActiveSpreadsheet();
    const localConfig = localSS.getSheetByName(CONFIG_TAB_NAME);

    if (!localConfig) {
      throw new Error('Project tab not found in local spreadsheet');
    }

    // QA Dashboard Config structure (row 4 starts data, row 3 is header)
    // Col A: Active, B: Jira Sync, C: Project, D: Modul, E: Submodul, F: PIC QA, G: Spreadsheet ID, H: Link, I: Jira Instance, J: Jira Project

    const dashboardLastRow = dashboardConfig.getLastRow();
    const dataRowCount = Math.max(60, dashboardLastRow - 3); // Dashboard starts from row 4

    // Sync from Dashboard row 4 onwards (skip header rows 1-3)
    const dashboardDataRange = 'A4:F' + (4 + dataRowCount - 1);
    const dashboardData = dashboardConfig.getRange(dashboardDataRange).getValues();

    // Read existing local data to preserve ratings
    const localLastRow = localConfig.getLastRow();
    const existingData = localLastRow >= CONFIG_DATA_START_ROW ?
      localConfig.getRange(CONFIG_DATA_START_ROW, 1, localLastRow - CONFIG_DATA_START_ROW + 1, CONFIG_TOTAL_COLUMNS).getValues() : [];

    // Build map of existing ratings by project|modul|submodul key
    const existingRatingsMap = new Map();
    existingData.forEach(row => {
      const project = row[0] ? row[0].toString().trim() : '';
      const modul = row[1] ? row[1].toString().trim() : '';
      const submodul = row[2] ? row[2].toString().trim() : '';
      if (project && modul && submodul) {
        const key = [project, modul, submodul].join('|');
        existingRatingsMap.set(key, {
          difficulty: row[4] || 0,    // Column E
          risk: row[5] || 0,           // Column F
          complexity: row[6] || 0,     // Column G
          automation: row[7] || 0,     // Column H
          testComplexity: row[8] || 0  // Column I
        });
      }
    });

    // Map Dashboard Config to local Project Config structure
    // Dashboard: A=Active, C=Project, D=Modul, E=Submodul, F=PIC QA
    // Local Project: A=Project, B=Modul, C=Submodul, D=PIC QA, E=Diff, F=Risk, G=Comp, H=Auto, I=Test, J=Status

    const mappedData = [];
    const projectSet = new Set();
    const modulSet = new Set();
    let submodulCount = 0;

    dashboardData.forEach(row => {
      const active = row[0] === true;  // Dashboard col A
      const project = row[2] ? row[2].toString().trim() : '';    // Dashboard col C
      const modul = row[3] ? row[3].toString().trim() : '';      // Dashboard col D
      const submodul = row[4] ? row[4].toString().trim() : '';   // Dashboard col E
      const picQA = row[5] ? row[5].toString().trim() : '';      // Dashboard col F

      if (active && project && modul && submodul) {
        const key = [project, modul, submodul].join('|');
        const existingRatings = existingRatingsMap.get(key);

        // Map to local structure: Project, Modul, Submodul, PIC QA, 5 params (preserve if exist), Status
        mappedData.push([
          project,   // A: Project
          modul,     // B: Modul
          submodul,  // C: Submodul
          picQA,     // D: PIC QA
          existingRatings ? existingRatings.difficulty : 0,    // E: Difficulty (preserve or 0)
          existingRatings ? existingRatings.risk : 0,          // F: Risk (preserve or 0)
          existingRatings ? existingRatings.complexity : 0,    // G: Complexity (preserve or 0)
          existingRatings ? existingRatings.automation : 0,    // H: Automation (preserve or 0)
          existingRatings ? existingRatings.testComplexity : 0,// I: Test Complexity (preserve or 0)
          'Active'   // J: Status
        ]);

        projectSet.add(project);
        modulSet.add(project + '|' + modul);
        submodulCount++;
      }
    });

    if (mappedData.length === 0) {
      Logger.log('⚠️ No active data found in Dashboard Config');
      return {
        success: false,
        message: 'No active data found in Dashboard Config',
        synced: { projects: 0, moduls: 0, submoduls: 0 }
      };
    }

    // Smart sorting: rated items first (by total score desc), unrated items last
    mappedData.sort((a, b) => {
      const aTotalScore = a[4] + a[5] + a[6] + a[7] + a[8]; // Sum of ratings
      const bTotalScore = b[4] + b[5] + b[6] + b[7] + b[8];

      const aHasRatings = aTotalScore > 0;
      const bHasRatings = bTotalScore > 0;

      // If one has ratings and the other doesn't, prioritize the one with ratings
      if (aHasRatings && !bHasRatings) return -1;
      if (!aHasRatings && bHasRatings) return 1;

      // If both have ratings, sort by total score (highest first)
      if (aHasRatings && bHasRatings) return bTotalScore - aTotalScore;

      // If neither has ratings, maintain original order
      return 0;
    });

    // Clear local data area first (from row 4 onwards, keep header)
    const localDataRange = 'A' + CONFIG_DATA_START_ROW + ':J' + (CONFIG_DATA_START_ROW + Math.max(60, mappedData.length) - 1);
    localConfig.getRange(localDataRange).clearContent();

    // Write synced data (now 10 columns: Project, Modul, Submodul, PIC QA, 5 ratings, Status)
    const writeRange = 'A' + CONFIG_DATA_START_ROW + ':J' + (CONFIG_DATA_START_ROW + mappedData.length - 1);
    localConfig.getRange(writeRange).setValues(mappedData);

    // Apply formatting to synced rows
    for (let i = 0; i < mappedData.length; i++) {
      const rowNum = CONFIG_DATA_START_ROW + i;
      const bg = i % 2 === 0 ? '#ffffff' : '#f8f9fa';
      localConfig.getRange(rowNum, 1, 1, CONFIG_TOTAL_COLUMNS).setBackground(bg);
    }

    // Update helper columns
    updateProjectHelperColumns();

    // Cache spreadsheet name for future use
    const dashboardName = dashboardSS.getName();
    saveSyncSettings(settings.spreadsheetId, dashboardName);

    Logger.log('✅ Sync complete: ' + projectSet.size + ' projects, ' + modulSet.size + ' moduls, ' + submodulCount + ' submoduls');

    return {
      success: true,
      message: 'Synced from ' + dashboardName,
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
 * Auto sync + refresh (called by trigger)
 * This combines sync from QA Dashboard and local dashboard refresh
 */
function autoSyncAndRefresh() {
  try {
    Logger.log('🔄 Auto sync & refresh started...');

    const settings = getSyncSettings();

    // Sync from QA Dashboard if enabled
    if (settings.enabled && settings.spreadsheetId) {
      const syncResult = syncFromDashboard();

      if (syncResult.success) {
        Logger.log('✅ Synced: ' + syncResult.synced.projects + ' projects, ' +
                   syncResult.synced.moduls + ' moduls, ' +
                   syncResult.synced.submoduls + ' submoduls');

        // Update helper columns after sync
        updateProjectHelperColumns();

        // Refresh Team Members dropdown validation
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const teamSheet = ss.getSheetByName(TEAM_TAB_NAME);
        if (teamSheet) {
          addTeamDataValidation(teamSheet);
        }

      } else {
        Logger.log('⚠️ Sync skipped: ' + syncResult.message);
      }
    } else {
      Logger.log('ℹ️ QA Dashboard sync disabled, skipping sync step');
    }

    // Refresh local dashboard
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dashSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
    if (dashSheet) ss.deleteSheet(dashSheet);
    SpreadsheetApp.flush();

    createDashboard();
    Logger.log('✅ Dashboard refreshed');

    Logger.log('✅ Auto sync & refresh complete');

  } catch (error) {
    Logger.log('❌ Auto sync & refresh error: ' + error.message);
  }
}

/**
 * Test sync connection (verify QA Dashboard spreadsheet access)
 */
function testSyncConnection() {
  const settings = getSyncSettings();

  if (!settings.enabled || !settings.spreadsheetId) {
    return {
      success: false,
      message: 'QA Dashboard sync not configured'
    };
  }

  try {
    const dashboardSS = SpreadsheetApp.openById(settings.spreadsheetId);
    const dashboardConfig = dashboardSS.getSheetByName('Config');

    if (!dashboardConfig) {
      return {
        success: false,
        message: 'Config tab not found in QA Dashboard spreadsheet'
      };
    }

    const dashboardName = dashboardSS.getName();

    // Cache spreadsheet name for future use
    saveSyncSettings(settings.spreadsheetId, dashboardName);

    // Count active modules
    const lastRow = dashboardConfig.getLastRow();
    if (lastRow < 4) {
      return {
        success: true,
        message: 'Connected to: ' + dashboardName + ' (no data found)',
        spreadsheetName: dashboardName,
        activeModules: 0
      };
    }

    const data = dashboardConfig.getRange('A4:F' + lastRow).getValues();
    const activeCount = data.filter(row => row[0] === true && row[2] && row[3] && row[4]).length;

    return {
      success: true,
      message: 'Connected to: ' + dashboardName + ' (' + activeCount + ' active modules)',
      spreadsheetName: dashboardName,
      activeModules: activeCount
    };

  } catch (error) {
    return {
      success: false,
      message: 'Connection failed: ' + error.message
    };
  }
}
