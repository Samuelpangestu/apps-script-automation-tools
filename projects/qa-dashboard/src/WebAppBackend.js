/**
 * ═══════════════════════════════════════════════════════════════════════
 * WEB APP BACKEND
 * Backend functions for QA Dashboard Web App
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Web App entry point - serves HTML page
 * This function is required for Web App deployment
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('QA Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * CONFIGURATION: Dashboard Spreadsheet IDs
 * These are the actual Google Sheets spreadsheet IDs (not Apps Script IDs)
 *
 * Testing Dashboard:  https://docs.google.com/spreadsheets/d/1-1ztAqfZz-lOCErOA7SQ5rVGsuBoIIVtn4epMsd0rzg/edit
 * Production Dashboard: https://docs.google.com/spreadsheets/d/1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY/edit
 */
// Use different IDs for testing vs production Apps Script projects
const TESTING_SPREADSHEET_ID = '1-1ztAqfZz-lOCErOA7SQ5rVGsuBoIIVtn4epMsd0rzg';
const PRODUCTION_SPREADSHEET_ID = '1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY';

// Auto-detect which environment based on script ID
const CURRENT_SCRIPT_ID = ScriptApp.getScriptId();
const IS_PRODUCTION = CURRENT_SCRIPT_ID === '1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO';
const DEFAULT_SPREADSHEET_ID = IS_PRODUCTION ? PRODUCTION_SPREADSHEET_ID : TESTING_SPREADSHEET_ID;

/**
 * Get the Dashboard spreadsheet
 * HARDCODED to always use Production spreadsheet
 *
 * @returns {Spreadsheet} Dashboard spreadsheet
 */
function getDashboardSpreadsheet_() {
  // FORCE PRODUCTION SPREADSHEET ONLY
  // This ensures web app always shows production data
  return SpreadsheetApp.openById(PRODUCTION_SPREADSHEET_ID);
}


/**
 * Setup function: Store Web App URL for notifications
 * Run this AFTER deploying web app to store the deployment URL
 *
 * Web App URL akan digunakan di WhatsApp notifications (line 1369 Notifications.js)
 */
function setupWebAppUrl() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Setup Web App URL',
    'Paste Web App URL dari deployment:\n\n' +
    '(Format: https://script.google.com/a/macros/.../exec)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const webAppUrl = response.getResponseText().trim();

    if (webAppUrl && webAppUrl.includes('script.google.com')) {
      PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', webAppUrl);

      ui.alert(
        '✅ Web App URL Saved',
        'URL has been stored for notifications.\n\n' +
        'URL: ' + webAppUrl + '\n\n' +
        'This URL will be included in WhatsApp notifications.',
        ui.ButtonSet.OK
      );

      Logger.log('Web App URL saved: ' + webAppUrl);
    } else {
      ui.alert(
        '❌ Invalid URL',
        'Please enter a valid Apps Script Web App URL.\n\n' +
        'Format: https://script.google.com/a/macros/.../exec',
        ui.ButtonSet.OK
      );
    }
  }
}


/**
 * Main function to fetch all dashboard data
 * Called by WebApp.html via google.script.run
 *
 * @returns {Object} Dashboard data including summary, history, and module list
 */
function getDashboardData() {
  try {
    Logger.log('Starting getDashboardData...');
    const ss = getDashboardSpreadsheet_();
    Logger.log('Got spreadsheet');

    const summaryData = getSummaryData_(ss);
    Logger.log('Got summary data');

    const historyData = getHistoryData_(ss);
    Logger.log('Got history data: ' + historyData.length + ' rows');

    const modules = getUniqueModules_(historyData);
    Logger.log('Got modules: ' + modules.length);

    let bugsTableData = [];
    try {
      bugsTableData = getBugsTableData_(ss);
      Logger.log('Got bugs table data: ' + bugsTableData.length + ' rows');
    } catch (bugError) {
      Logger.log('ERROR getting bugs data (continuing with empty): ' + bugError.toString());
      bugsTableData = [];
    }

    let vaptHistoryData = [];
    try {
      vaptHistoryData = getVAPTHistoryData_(ss);
      Logger.log('Got VAPT history data: ' + vaptHistoryData.length + ' rows');
    } catch (vaptError) {
      Logger.log('ERROR getting VAPT data (continuing with empty): ' + vaptError.toString());
      vaptHistoryData = [];
    }

    // Serialize history data (convert Date objects to strings)
    const serializedHistory = historyData.map(row => {
      return row.map(cell => {
        if (cell instanceof Date) {
          return cell.toISOString();
        }
        return cell;
      });
    });

    // Serialize bugs table data (convert Date objects and ensure JSON-safe)
    const serializedBugsTable = bugsTableData.map(row => {
      return row.map(cell => {
        if (cell instanceof Date) {
          return cell.toISOString();
        }
        // Ensure numbers are valid (not NaN or Infinity)
        if (typeof cell === 'number' && !isFinite(cell)) {
          return 0;
        }
        return cell;
      });
    });

    // Ensure summary data is JSON-safe
    const cleanSummary = {};
    for (const key in summaryData) {
      const val = summaryData[key];
      if (val instanceof Date) {
        cleanSummary[key] = val.toISOString();
      } else if (typeof val === 'number' && !isFinite(val)) {
        cleanSummary[key] = 0;
      } else {
        cleanSummary[key] = val;
      }
    }

    // Serialize VAPT history data
    const serializedVaptHistory = vaptHistoryData.map(row => {
      return row.map(cell => {
        if (cell instanceof Date) {
          return cell.toISOString();
        }
        if (typeof cell === 'number' && !isFinite(cell)) {
          return 0;
        }
        return cell;
      });
    });

    const result = {
      summary: cleanSummary,
      history: serializedHistory,
      modules: modules,
      bugsTable: serializedBugsTable,
      vaptHistory: serializedVaptHistory,
      timestamp: new Date().toISOString()
    };

    Logger.log('✅ getDashboardData SUCCESS - returning data');
    Logger.log('Result keys: ' + Object.keys(result).join(', '));
    return result;

  } catch (error) {
    Logger.log('❌ ERROR in getDashboardData: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);

    // Return error with default data instead of throwing
    return {
      error: true,
      errorMessage: error.message,
      errorStack: error.stack,
      summary: getDefaultSummary_(),
      history: [],
      modules: [],
      bugsTable: [],
      vaptHistory: [],
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Get summary data from Overview and Bugs tabs (latest aggregated data)
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Object} Summary metrics
 */
function getSummaryData_(ss) {
  const overview = ss.getSheetByName('Overview');
  const bugs = ss.getSheetByName('Bugs');

  if (!overview || !bugs) {
    return getDefaultSummary_();
  }

  // Get all data rows from Overview (skip headers)
  const overviewData = overview.getDataRange().getValues();

  // Find "TOTAL" or "AVERAGE" row in Overview (usually last row)
  let totalRow = null;
  for (let i = overviewData.length - 1; i >= 0; i--) {
    const moduleName = String(overviewData[i][0]).toUpperCase();
    if (moduleName.includes('TOTAL') || moduleName.includes('AVERAGE')) {
      totalRow = overviewData[i];
      Logger.log('Found TOTAL row at index ' + i + ': ' + moduleName);
      break;
    }
  }

  // If no TOTAL row found, aggregate all rows
  if (!totalRow) {
    Logger.log('No TOTAL row found, aggregating manually...');
    totalRow = aggregateOverviewData_(overviewData);
  }

  // Overview columns (0-indexed) - VERIFIED from testOverviewStructure:
  // [0-3]: Module info (Project, Modul, Submodul, PIC)
  // [4] E: Total Bugs = 16 ✅
  // [5] F: Blocker = 14 ✅ (Medium-Critical bugs, status NOT Closed/Won't Fix/Verified)
  // [6] G: Critical = 3 ✅
  // [7] H: Prod = 0 ✅
  // [8] I: WEB Total = 411
  // [9] J: WEB Pass = 284
  // [10] K: WEB Fail = 0
  // [11] L: WEB Block = 0
  // [12] M: WEB Pass% = 0.69 ✅
  // [13] N: SMOKE WEB Total = 258
  // [14] O: SMOKE WEB Pass% = 0
  // [15] P: SMOKE WEB Exec% = 0
  // [16] Q: API Total = 0
  // [17] R: API Pass = 0
  // [18] S: API Fail = 0
  // [19] T: API Block = 0
  // [20] U: API Pass% = 0
  // [21] V: (unknown/exec?)
  // [22] W: SMOKE API Total = 0
  // [23] X: SMOKE API Pass% = 0
  // [24] Y: PERF = empty
  // [25] Z: Notes = empty

  // Calculate Web Exec% manually from totals to avoid wrong column
  let webExecRate = 0;
  const webTotal = totalRow[8] || 0;  // [8] I: WEB Total
  if (webTotal > 0) {
    const webExecuted = (totalRow[9] || 0) + (totalRow[10] || 0) + (totalRow[11] || 0);  // Pass + Fail + Block
    webExecRate = webExecuted / webTotal;
  }

  // API Exec% - set to 0 if no API tests
  let apiExecRate = 0;
  const apiTotal = totalRow[16] || 0;  // [16] Q: API Total
  if (apiTotal > 0) {
    const apiExecuted = (totalRow[17] || 0) + (totalRow[18] || 0) + (totalRow[19] || 0);  // Pass + Fail + Block
    apiExecRate = apiExecuted / apiTotal;
  }

  // Get VAPT summary data from VAPT tab
  // IMPORTANT: Read from summary cell B5 (Total Blocker), NOT by aggregating all apps
  // Row 5, Col 2 (B5): Total Blocker (Combined = Ad Hoc + Regular)
  // This ensures we get TODAY'S blocker count, not cumulative from all history
  const vaptTab = ss.getSheetByName('VAPT');
  let vaptBlocker = 0;
  let vaptCritical = 0;
  let vaptHigh = 0;
  let vaptMedium = 0;

  if (vaptTab) {
    try {
      // Read Total Blocker from summary cell B5 (row 5, col 2)
      const b5Value = vaptTab.getRange(5, 2).getValue();
      vaptBlocker = Number(b5Value) || 0;

      // Calculate breakdown from app rows (data starts from row 10)
      // This gives us the severity breakdown for today's blocker
      const vaptData = vaptTab.getDataRange().getValues();
      // VAPT columns (NEW): Project(0) | Aplikasi(1) | Blocker(2) | Critical(3) | High(4) | Medium(5) | Low(6) | Info(7)
      for (let i = 9; i < vaptData.length; i++) { // FIXED: Start from index 9 (row 10)
        const row = vaptData[i];

        // Skip empty rows
        if (!row[0] && !row[1]) continue;

        // Skip section headers (merged cells with "═══ PROJECTNAME ═══")
        if (String(row[0]).includes('═══')) continue;

        // Skip column header rows (row with "Aplikasi" in col 1)
        if (String(row[1]).trim() === 'Aplikasi') continue;

        const critical = Number(row[3]) || 0;
        const high = Number(row[4]) || 0;
        const medium = Number(row[5]) || 0;

        vaptCritical += critical;
        vaptHigh += high;
        vaptMedium += medium;
      }
    } catch (vaptError) {
      Logger.log('Error reading VAPT data: ' + vaptError.toString());
    }
  }

  const summary = {
    webPassRate: totalRow[12] || 0,        // [12] M: Web Pass% ✅
    webExecRate: webExecRate,              // Calculated from totals
    apiPassRate: totalRow[20] || 0,        // [20] U: API Pass%
    apiExecRate: apiExecRate,              // Calculated from totals
    smokeWebPassRate: totalRow[14] || 0,   // [14] O: Smoke Web Pass%
    smokeWebExecRate: totalRow[15] || 0,   // [15] P: Smoke Web Exec%
    smokeApiPassRate: totalRow[23] || 0,   // [23] X: Smoke API Pass%
    smokeApiExecRate: totalRow[23] || 0,   // [23] X: Smoke API Exec% (may be same column)
    perfResult: totalRow[24] || '-',       // [24] Y: Performance
    totalBugs: totalRow[4] || 0,           // [4] E: Total Bugs ✅
    openBugs: totalRow[4] - (totalRow[7] || 0) || 0,  // Total - Prod = Open (approximation)
    blockerBugs: totalRow[5] || 0,         // [5] F: Blocker ✅
    criticalBugs: totalRow[6] || 0,        // [6] G: Critical ✅
    vaptBlocker: vaptBlocker,              // VAPT Blocker (Critical + High + Medium)
    vaptCritical: vaptCritical,            // VAPT Critical
    vaptHigh: vaptHigh,                    // VAPT High
    vaptMedium: vaptMedium,                // VAPT Medium
    lastUpdated: getLastRefreshTime_(overview)
  };

  return summary;
}

/**
 * Get historical data from History tab (DAILY snapshots only)
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Array} History data rows (1 entry per day per module)
 */
function getHistoryData_(ss) {
  const history = ss.getSheetByName('History');

  if (!history) {
    return [];
  }

  // Get all data (skip header row 1-2)
  const data = history.getDataRange().getValues();

  // Filter to get DAILY data only (1 entry per day per MODULE)
  const dailyData = {};

  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Skip empty rows

    const timestamp = new Date(row[0]);
    const project = row[1];      // Col B: Project (SIPGN)
    const modul = row[2];         // Col C: Modul (0.E2E, 1, 2, 3, etc)

    // Create unique key per day per module
    const dateKey = timestamp.toISOString().split('T')[0] + '_' + project + '_' + modul; // YYYY-MM-DD_Project_Modul

    // Keep only the LATEST entry for each day per module
    if (!dailyData[dateKey] || new Date(dailyData[dateKey][0]) < timestamp) {
      dailyData[dateKey] = row;
    }
  }

  // Convert back to array and sort by date (newest first)
  return Object.values(dailyData).sort((a, b) => new Date(b[0]) - new Date(a[0]));
}

/**
 * Extract unique modules from history data
 *
 * @param {Array} historyData - History rows
 * @returns {Array} Unique module identifiers (Project - Modul)
 */
function getUniqueModules_(historyData) {
  const moduleSet = new Set();

  historyData.forEach(row => {
    const project = row[1]; // Col B: Project
    const modul = row[2];   // Col C: Modul

    if (project && modul && String(project).trim() !== '' && String(modul).trim() !== '') {
      // Create module identifier: "SIPGN - 0.E2E", "SIPGN - 1", etc
      const moduleId = String(project).trim() + ' - ' + String(modul).trim();
      moduleSet.add(moduleId);
    }
  });

  return Array.from(moduleSet).sort();
}

/**
 * Get last refresh timestamp from sheet
 *
 * @param {Sheet} sheet - Sheet to check
 * @returns {string} Last refresh time
 */
function getLastRefreshTime_(sheet) {
  try {
    // Check if there's a "Last refreshed" note or cell
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // Check top-left area for timestamp (usually row 1)
    for (let r = 1; r <= Math.min(3, lastRow); r++) {
      for (let c = 1; c <= Math.min(5, lastCol); c++) {
        const val = String(sheet.getRange(r, c).getValue());
        if (val.includes('Last refreshed') || val.includes('refreshed:')) {
          return val;
        }
      }
    }

    return 'No refresh data available';
  } catch (error) {
    return 'Unknown';
  }
}

/**
 * Get default summary when sheets are not available
 *
 * @returns {Object} Default summary
 */
function getDefaultSummary_() {
  return {
    webPassRate: 0,
    webExecRate: 0,
    apiPassRate: 0,
    apiExecRate: 0,
    smokeWebPassRate: 0,
    smokeWebExecRate: 0,
    smokeApiPassRate: 0,
    smokeApiExecRate: 0,
    perfResult: '-',
    totalBugs: 0,
    openBugs: 0,
    blockerBugs: 0,
    criticalBugs: 0,
    vaptBlocker: 0,
    vaptCritical: 0,
    vaptHigh: 0,
    vaptMedium: 0,
    lastUpdated: 'No data'
  };
}

/**
 * Aggregate overview data when TOTAL row not found
 *
 * @param {Array} data - Overview data
 * @returns {Array} Aggregated row
 */
function aggregateOverviewData_(data) {
  // Skip header rows (first 4 rows)
  const dataRows = data.slice(4);

  // Initialize totals
  const totals = new Array(26).fill(0);
  totals[0] = 'TOTAL PORTFOLIO'; // Module name

  let count = 0;
  dataRows.forEach(row => {
    const moduleName = String(row[0]).toUpperCase();
    // Skip empty rows and footer rows
    if (!moduleName || moduleName.includes('TOTAL') || moduleName === '') {
      return;
    }

    // Sum numeric columns
    for (let i = 4; i < 26; i++) {
      if (typeof row[i] === 'number') {
        totals[i] += row[i];
      }
    }
    count++;
  });

  // Calculate averages for percentage columns
  // Columns: 12=WebPass%, 14=SmokeWebPass%, 15=SmokeWebExec%, 20=APIPass%, 21=APIExec%, 22-23=SmokeAPI
  if (count > 0) {
    [12, 14, 15, 20, 21, 22, 23].forEach(col => {
      totals[col] = totals[col] / count;
    });
  }

  Logger.log('Aggregated overview data for ' + count + ' modules');
  return totals;
}


/**
 * Get bugs table data from Bugs tab
 * Returns all bug entries grouped by module
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Array} Bugs table rows
 */
function getBugsTableData_(ss) {
  try {
    const bugs = ss.getSheetByName('Bugs');

    if (!bugs) {
      Logger.log('Bugs sheet not found');
      return [];
    }

    const data = bugs.getDataRange().getValues();

    if (data.length <= 1) {
      Logger.log('No bugs table data found');
      return [];
    }

    // Skip header row, get all bug data rows
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[0]) continue;

      const moduleName = String(row[0]).toUpperCase();

      // Skip TOTAL rows
      if (moduleName.includes('TOTAL')) {
        continue;
      }

      rows.push(row);
    }

    Logger.log('Bugs table rows: ' + rows.length);
    return rows;

  } catch (error) {
    Logger.log('ERROR in getBugsTableData_: ' + error.toString());
    return [];
  }
}

/**
 * Get VAPT History data from VAPT History tab
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @returns {Array} VAPT History data rows
 */
function getVAPTHistoryData_(ss) {
  const vaptHistory = ss.getSheetByName('VAPT History');

  if (!vaptHistory) {
    Logger.log('VAPT History tab not found');
    return [];
  }

  try {
    // Get all data (skip header row 1-2, data starts from row 3)
    const data = vaptHistory.getDataRange().getValues();

    // Filter valid rows (skip empty and header rows)
    const historyRows = [];
    for (let i = 2; i < data.length; i++) {  // FIXED: Start from index 2 (row 3)
      const row = data[i];

      // Skip empty timestamp rows
      if (!row[0]) continue;

      // Skip header rows (timestamp column should be a date string, not "Timestamp" text)
      if (String(row[0]).toLowerCase().includes('timestamp')) continue;
      if (String(row[0]).toLowerCase().includes('date')) continue;

      historyRows.push(row);
    }

    // Sort by timestamp (newest first)
    historyRows.sort((a, b) => new Date(b[0]) - new Date(a[0]));

    Logger.log('📊 VAPT History rows loaded: ' + historyRows.length);

    // Log latest entries for debugging
    if (historyRows.length > 0) {
      Logger.log('  Latest entry: ' + historyRows[0][0] + ' | Project=' + historyRows[0][1] + ' | Blocker=' + historyRows[0][2]);
      if (historyRows.length > 1) {
        Logger.log('  2nd entry: ' + historyRows[1][0] + ' | Project=' + historyRows[1][1] + ' | Blocker=' + historyRows[1][2]);
      }
    }

    return historyRows;

  } catch (error) {
    Logger.log('Error reading VAPT History: ' + error.toString());
    return [];
  }
}

/**
 * Get VAPT Dashboard data (for VAPT-specific page)
 * @returns {Object} VAPT dashboard data
 */
function getVAPTDashboardData() {
  try {
    const ss = getDashboardSpreadsheet_();
    const vaptTab = ss.getSheetByName('VAPT');

    if (!vaptTab) {
      return {
        error: true,
        errorMessage: 'VAPT tab not found',
        summary: { vaptBlocker: 0, vaptCritical: 0, vaptHigh: 0, vaptMedium: 0, lastUpdated: 'No data' },
        vaptApps: [],
        vaptHistory: []
      };
    }

    // Get VAPT summary and apps data
    const vaptData = vaptTab.getDataRange().getValues();
    let vaptBlocker = 0;
    let vaptCritical = 0;
    let vaptHigh = 0;
    let vaptMedium = 0;
    const vaptApps = [];

    // Skip header and summary rows (rows 1-9), data starts from row 10
    // VAPT tab columns: Project(0) | Aplikasi(1) | Blocker(2) | Critical(3) | High(4) | Medium(5) | Low(6) | Info(7)
    Logger.log('📊 Reading VAPT tab data for dashboard...');
    for (let i = 9; i < vaptData.length; i++) {  // FIXED: Start from index 9 (row 10)
      const row = vaptData[i];

      // Skip empty rows
      if (!row[0] && !row[1]) continue;

      // Skip section headers (merged cells with "═══ PROJECTNAME ═══")
      if (String(row[0]).includes('═══')) continue;

      // Skip column header rows (row with "Aplikasi" in col 1)
      if (String(row[1]).trim() === 'Aplikasi') continue;

      const blocker = Number(row[2]) || 0;   // Blocker already calculated (Critical + High + Medium)
      const critical = Number(row[3]) || 0;
      const high = Number(row[4]) || 0;
      const medium = Number(row[5]) || 0;
      const low = Number(row[6]) || 0;

      Logger.log('  Row ' + (i+1) + ': ' + row[1] + ' | B=' + blocker + ' C=' + critical + ' H=' + high + ' M=' + medium);

      vaptCritical += critical;
      vaptHigh += high;
      vaptMedium += medium;
      vaptBlocker += blocker;  // Use pre-calculated blocker from col 2

      vaptApps.push({
        aplikasi: row[1],  // FIXED: Aplikasi is in col 1, not col 0
        blocker: blocker,
        critical: critical,
        high: high,
        medium: medium,
        low: low
      });
    }

    Logger.log('📊 VAPT Summary: Blocker=' + vaptBlocker + ' | C=' + vaptCritical + ' | H=' + vaptHigh + ' | M=' + vaptMedium);

    // Get VAPT history
    const vaptHistory = getVAPTHistoryData_(ss);

    // Serialize history
    const serializedHistory = vaptHistory.map(row => {
      return row.map(cell => {
        if (cell instanceof Date) return cell.toISOString();
        if (typeof cell === 'number' && !isFinite(cell)) return 0;
        return cell;
      });
    });

    return {
      summary: {
        vaptBlocker: vaptBlocker,
        vaptCritical: vaptCritical,
        vaptHigh: vaptHigh,
        vaptMedium: vaptMedium,
        lastUpdated: new Date().toISOString()
      },
      vaptApps: vaptApps,
      vaptHistory: serializedHistory
    };

  } catch (error) {
    Logger.log('Error in getVAPTDashboardData: ' + error.toString());
    return {
      error: true,
      errorMessage: error.message,
      summary: { vaptBlocker: 0, vaptCritical: 0, vaptHigh: 0, vaptMedium: 0, lastUpdated: 'Error' },
      vaptApps: [],
      vaptHistory: []
    };
  }
}

/**
 * Test function to verify data structure
 * Run this from Apps Script Editor to debug
 */
function testGetDashboardData() {
  const data = getDashboardData();

  if (data.error) {
    Logger.log('ERROR: ' + data.errorMessage);
    return;
  }

  Logger.log('✅ Data loaded successfully');
  Logger.log('Summary: ' + JSON.stringify(data.summary));
  Logger.log('History rows: ' + data.history.length);
  Logger.log('Modules: ' + data.modules.join(', '));
  Logger.log('Bugs table rows: ' + (data.bugsTable ? data.bugsTable.length : 0));
}
