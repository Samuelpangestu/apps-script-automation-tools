/**
 * ═══════════════════════════════════════════════════════════════════════
 * WEB APP BACKEND
 * Backend functions for QA Dashboard Web App
 * ═══════════════════════════════════════════════════════════════════════
 */

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
 * Web App Entry Point - Serves both HTML and JSON API
 *
 * DUAL MODE SUPPORT:
 * 1. HTML Mode (default): Returns WebApp.html for direct browser access
 * 2. JSON API Mode: Returns JSON data for GitHub Pages frontend
 *
 * Usage:
 * - HTML: https://script.google.com/.../exec
 * - JSON: https://script.google.com/.../exec?format=json
 *
 * @param {Object} e - Event object with parameters
 * @returns {HtmlOutput|TextOutput} HTML page or JSON response
 */
function doGet(e) {
  try {
    // Get user email for authentication
    const userEmail = Session.getActiveUser().getEmail();

    // Check if user is from organization (Workspace only)
    // Comment out this check if you want public access
    if (userEmail && !userEmail.endsWith('@inadigital.co.id')) {
      const errorResponse = {
        error: 'Unauthorized',
        message: 'Access restricted to INA Digital employees only',
        email: userEmail
      };

      return ContentService
        .createTextOutput(JSON.stringify(errorResponse))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Check if request wants JSON (API mode for GitHub Pages)
    if (e.parameter && e.parameter.format === 'json') {
      Logger.log('API Mode: Returning JSON data');

      const data = getDashboardData();

      return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Default: Return HTML (existing Web App mode)
    Logger.log('HTML Mode: Serving WebApp.html');

    return HtmlService.createHtmlOutputFromFile('WebApp')
      .setTitle('QA Dashboard - INA Digital')
      .setFaviconUrl('https://www.inadigital.co.id/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log('ERROR in doGet: ' + error.toString());

    // Return error as JSON if API mode, or simple HTML if web mode
    const errorMessage = 'Failed to load dashboard: ' + error.message;

    if (e.parameter && e.parameter.format === 'json') {
      return ContentService
        .createTextOutput(JSON.stringify({error: errorMessage}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return HtmlService.createHtmlOutput(
      '<h1>Error</h1><p>' + errorMessage + '</p>'
    );
  }
}

/**
 * Setup function: Store current spreadsheet ID for web app access
 * Run this ONCE from spreadsheet context before deploying web app
 */
function setupWebAppSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();

  PropertiesService.getScriptProperties().setProperty('DASHBOARD_SPREADSHEET_ID', ssId);

  SpreadsheetApp.getUi().alert(
    '✅ Web App Setup Complete',
    'Spreadsheet ID has been stored.\n\n' +
    'ID: ' + ssId + '\n\n' +
    'You can now deploy as Web App.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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

    const result = {
      summary: cleanSummary,
      history: serializedHistory,
      modules: modules,
      bugsTable: serializedBugsTable,
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
