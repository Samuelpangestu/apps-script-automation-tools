/**
 * ═══════════════════════════════════════════════════════════════════════
 * WEB APP BACKEND
 * Backend functions for QA Dashboard Web App
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Web App entry point - serves HTML page OR handles API requests
 * This function is required for Web App deployment
 *
 * API Endpoints for Next.js:
 * - ?action=getBugsData        -> Returns bugs from Bugs tab
 * - ?action=getVAPTData        -> Returns VAPT findings
 * - ?action=getKPIData         -> Returns KPI metrics
 * - No action parameter        -> Returns HTML dashboard
 */
function doGet(e) {
  // Check if this is an API request (has action parameter)
  const action = e.parameter.action;

  if (action) {
    // Handle API requests - return JSON
    return handleApiRequest_(action, e);
  }

  // Default: serve HTML dashboard
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('QA Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Web App write entry point for Jenkins automation result ingestion.
 *
 * Expected request:
 * POST { action: "ingestAutomationResult", token: "...", payload: {...} }
 *
 * Token source: Script Properties key AUTOMATION_INGEST_TOKEN.
 */
function doPost(e) {
  try {
    const body = parsePostBody_(e);
    const action = body.action || (e.parameter && e.parameter.action);

    switch (action) {
      case 'ingestAutomationResult':
        return jsonResponse_(ingestAutomationResult_(body.payload || body));

      case 'refreshAutomationHistory':
        return jsonResponse_(refreshAutomationHistoryApi_(body.payload || body));

      case 'sendClosureEmail':
        return jsonResponse_(sendClosureEmail_(body));

      case 'createClosureEmailDraft':
        return jsonResponse_(createClosureEmailDraft_(body));

      default:
        throw new Error('Unknown action: ' + action);
    }
  } catch (error) {
    Logger.log('POST API Error: ' + error.toString());
    return jsonResponse_({
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    });
  }
}

function parsePostBody_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  const type = String(e.postData.type || '').toLowerCase();
  if (type.indexOf('application/json') >= 0 || String(e.postData.contents).trim().charAt(0) === '{') {
    return JSON.parse(e.postData.contents);
  }
  const params = {};
  String(e.postData.contents).split('&').forEach(pair => {
    const parts = pair.split('=');
    if (parts[0]) params[decodeURIComponent(parts[0])] = decodeURIComponent(parts.slice(1).join('=') || '');
  });
  return params;
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function ingestAutomationResult_(payload) {
  validateAutomationIngestToken_(payload);

  const ss = getDashboardSpreadsheet_();
  const ws = ensureAutomationRunsSheet_(ss);
  const timestamp = payload.timestamp ? new Date(payload.timestamp) : new Date();
  const channel = normalizeAutomationChannel_(payload.channel || payload.testType);
  if (channel !== 'web' && channel !== 'api') {
    throw new Error('Automation channel must be web or api');
  }

  const total = Number(payload.total) || 0;
  const passed = Number(payload.passed) || 0;
  const failed = Number(payload.failed) || 0;
  const skipped = Number(payload.skipped) || 0;
  const broken = Number(payload.broken) || 0;
  const flaky = Number(payload.flaky) || 0;
  const passRate = total > 0 ? passed / total : 0;
  const status = String(payload.status || payload.buildStatus || deriveAutomationStatus_(total, failed, broken)).trim();

  const row = [
    timestamp,
    String(payload.project || ''),
    String(payload.module || payload.modul || ''),
    String(payload.submodule || payload.submodul || ''),
    channel,
    String(payload.suite || ''),
    String(payload.environment || payload.env || ''),
    String(payload.contractKey || payload.automationContract || payload.tag || payload.jobName || ''),
    String(payload.tag || ''),
    String(payload.jobName || ''),
    String(payload.buildNumber || ''),
    String(payload.buildUrl || ''),
    String(payload.reportUrl || ''),
    status,
    total,
    passed,
    failed,
    skipped,
    broken,
    flaky,
    passRate,
    String(payload.source || 'jenkins'),
    JSON.stringify(payload)
  ];

  ws.appendRow(row);
  const lastRow = ws.getLastRow();
  ws.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  ws.getRange(lastRow, 21).setNumberFormat('0%');
  SpreadsheetApp.flush();

  let historyRefresh = {success:true};
  try {
    refreshAutomationHistory();
  } catch (historyError) {
    historyRefresh = {
      success:false,
      error:historyError.message
    };
    Logger.log('Automation result saved, but History refresh failed: ' + historyError.toString());
  }

  return {
    success: true,
    data: {
      row: lastRow,
      channel,
      project: row[1],
      module: row[2],
      submodule: row[3],
      status,
      passRate,
      historyRefresh
    },
    timestamp: new Date().toISOString()
  };
}

function refreshAutomationHistoryApi_(payload) {
  validateAutomationIngestToken_(payload);
  const result = refreshAutomationHistory({silent:true});
  return {
    success:true,
    data:result || {},
    timestamp:new Date().toISOString()
  };
}

function validateAutomationIngestToken_(payload) {
  const expectedToken = PropertiesService.getScriptProperties().getProperty('AUTOMATION_INGEST_TOKEN');
  if (!expectedToken) {
    throw new Error('AUTOMATION_INGEST_TOKEN is not configured in Script Properties');
  }
  const providedToken = String((payload && (payload.token || payload.authToken)) || '').trim();
  if (providedToken !== expectedToken) {
    throw new Error('Invalid automation ingestion token');
  }
}

function normalizeAutomationChannel_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (normalized.indexOf('api') >= 0) return 'api';
  if (normalized.indexOf('web') >= 0 || normalized.indexOf('speedweb') >= 0) return 'web';
  return normalized;
}

function deriveAutomationStatus_(total, failed, broken) {
  if (total <= 0) return 'No Tests';
  return (failed > 0 || broken > 0) ? 'Failed' : 'Passed';
}

/**
 * Handle API requests from Next.js frontend
 * Returns JSON response
 */
function handleApiRequest_(action, e) {
  try {
    let data = null;

    switch (action) {
      case 'getBugsData':
        data = getApiBugsData_();
        break;

      case 'getVAPTData':
        data = getApiVAPTData_();
        break;

      case 'getKPIData':
        data = getApiKPIData_();
        break;

      case 'getDashboardSummary':
        data = getApiDashboardSummary_();
        break;

      case 'getExternalTestReports':
        data = getApiExternalTestReports_(e.parameter);
        break;

      default:
        throw new Error('Unknown action: ' + action);
    }

    // Return success response
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: data,
      total: Array.isArray(data) ? data.length : null,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Return error response
    Logger.log('API Error [' + action + ']: ' + error.toString());

    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * API: Get bugs data formatted for Next.js
 * Returns array of bug objects
 */
function getApiBugsData_() {
  const ss = getDashboardSpreadsheet_();
  const bugs = ss.getSheetByName('Bugs');

  if (!bugs) {
    return [];
  }

  const data = bugs.getDataRange().getValues();
  const bugsList = [];

  // Column mapping (adjust based on actual Bugs tab structure):
  // A: Project | B: Module | C: Bug ID | D: Title | E: Status | F: Priority | G: Severity | H: Assignee | I: Reporter | J: Created | K: Resolved | L: Description

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row[0] && !row[1]) continue;

    // Skip TOTAL rows
    if (String(row[0]).toUpperCase().includes('TOTAL')) continue;

    bugsList.push({
      id: String(row[2] || ''),
      project: String(row[0] || ''),
      module: String(row[1] || ''),
      title: String(row[3] || ''),
      status: String(row[4] || ''),
      priority: String(row[5] || ''),
      severity: String(row[6] || ''),
      assignee: String(row[7] || ''),
      reporter: String(row[8] || ''),
      createdDate: row[9] ? new Date(row[9]).toISOString().split('T')[0] : '',
      resolvedDate: row[10] ? new Date(row[10]).toISOString().split('T')[0] : null,
      description: String(row[11] || '')
    });
  }

  return bugsList;
}

/**
 * API: Get VAPT findings formatted for Next.js
 * Returns array of VAPT finding objects
 */
function getApiVAPTData_() {
  const ss = getDashboardSpreadsheet_();
  const vaptTab = ss.getSheetByName('VAPT');

  if (!vaptTab) {
    return [];
  }

  const data = vaptTab.getDataRange().getValues();
  const findings = [];

  // VAPT tab columns: Project(0) | Aplikasi(1) | Blocker(2) | Critical(3) | High(4) | Medium(5) | Low(6) | Info(7)

  for (let i = 9; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row[0] && !row[1]) continue;

    // Skip section headers
    if (String(row[0]).includes('═══')) continue;

    // Skip column headers
    if (String(row[1]).trim() === 'Aplikasi') continue;

    findings.push({
      id: 'VAPT-' + i,
      project: String(row[0] || ''),
      module: String(row[1] || ''),
      app: String(row[1] || ''),
      statusFix: 'Open', // Not in current tab structure
      statusReVAPT: 'Open',
      risk: determineHighestRisk_(row),
      adjustedRisk: determineHighestRisk_(row),
      findingName: String(row[1] || ''),
      reportDate: new Date().toISOString().split('T')[0],
      alreadyInProd: 'Yes',
      blocker: Number(row[2]) || 0,
      critical: Number(row[3]) || 0,
      high: Number(row[4]) || 0,
      medium: Number(row[5]) || 0,
      low: Number(row[6]) || 0
    });
  }

  return findings;
}

/**
 * API: Get KPI data formatted for Next.js
 * Returns array of KPI objects per module
 */
function getApiKPIData_() {
  const ss = getDashboardSpreadsheet_();
  const overview = ss.getSheetByName('Overview');

  if (!overview) {
    return [];
  }

  const data = overview.getDataRange().getValues();
  const kpiList = [];

  // Skip header rows (first 5 rows)
  for (let i = 5; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row[0]) continue;

    // Skip TOTAL rows
    if (String(row[0]).toUpperCase().includes('TOTAL')) continue;

    const webTotal = Number(row[8]) || 0;
    const webPass = Number(row[9]) || 0;
    const webFail = Number(row[10]) || 0;
    const webBlock = Number(row[11]) || 0;

    kpiList.push({
      project: String(row[0] || ''),
      module: String(row[1] || ''),
      testCases: webTotal,
      automated: 0, // Not available in current structure
      passed: webPass,
      failed: webFail,
      blocked: webBlock
    });
  }

  return kpiList;
}

/**
 * API: Get dashboard summary (for overview page)
 */
function getApiDashboardSummary_() {
  const ss = getDashboardSpreadsheet_();
  return getSummaryData_(ss);
}

/**
 * Helper: Determine highest risk level from VAPT row
 */
function determineHighestRisk_(row) {
  const critical = Number(row[3]) || 0;
  const high = Number(row[4]) || 0;
  const medium = Number(row[5]) || 0;
  const low = Number(row[6]) || 0;

  if (critical > 0) return 'Critical';
  if (high > 0) return 'High';
  if (medium > 0) return 'Medium';
  if (low > 0) return 'Low';
  return 'Informational';
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
  return SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
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
    '(Format: https://qa-platform.inadigital.co.id/)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const webAppUrl = response.getResponseText().trim();

    if (webAppUrl && webAppUrl.indexOf('https://') === 0) {
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
        'Please enter a valid HTTPS dashboard URL.\n\n' +
        'Format: https://qa-platform.inadigital.co.id/',
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
  // Skip header rows (first 5 rows: web app link, timestamp, title, group headers, column headers)
  const dataRows = data.slice(5);

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

    // Skip Bugs tab header rows 1-4, get data rows only.
    const rows = [];

    for (let i = 4; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[0]) continue;

      const moduleName = String(row[0]).toUpperCase();

      // Skip TOTAL/header rows
      if (moduleName.includes('TOTAL') || moduleName === 'PROJECT') {
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

    // Log latest entries for debugging (NEW STRUCTURE: Timestamp, Project, Module, Submodule, Total, Critical, High, Medium, Low, Informational, Todo, On Progress, Done, Open, Closed)
    if (historyRows.length > 0) {
      Logger.log('  Latest entry: ' + historyRows[0][0] + ' | Project=' + historyRows[0][1] + ' | Module=' + historyRows[0][2] + ' | Submodule=' + historyRows[0][3] + ' | Total=' + historyRows[0][4]);
      if (historyRows.length > 1) {
        Logger.log('  2nd entry: ' + historyRows[1][0] + ' | Project=' + historyRows[1][1] + ' | Module=' + historyRows[1][2] + ' | Submodule=' + historyRows[1][3] + ' | Total=' + historyRows[1][4]);
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

// ═══════════════════════════════════════════════════════════════════════
// EXTERNAL TEST REPORT & CLOSURE EMAIL APIs
// ═══════════════════════════════════════════════════════════════════════

/**
 * API: Get External Test Reports with optional filter
 * Query params:
 *   - project: Filter by project name
 *   - module: Filter by module name
 *   - status: Filter by overall status
 *
 * Returns array of External Test Report objects
 */
function getApiExternalTestReports_(params) {
  const ss = getDashboardSpreadsheet_();
  const config = ss.getSheetByName('Config');

  if (!config) {
    throw new Error('Config sheet not found');
  }

  const modules = getModuleList_(ss);
  const reports = [];

  modules.forEach(mod => {
    try {
      if (!mod.id || !mod.active) return;

      // Apply filters if provided
      if (params && params.project && mod.project !== params.project) return;
      if (params && params.module && mod.module !== params.module) return;

      const qatmSs = SpreadsheetApp.openById(mod.id);
      const extReportSheet = qatmSs.getSheetByName('External Test Report');

      if (!extReportSheet) {
        Logger.log('No External Test Report tab for: ' + mod.name);
        return;
      }

      const report = readExternalTestReport_(extReportSheet, qatmSs);

      // Apply status filter if provided
      if (params && params.status && report.overallStatus !== params.status) return;

      // Add module metadata
      report.moduleId = mod.id;
      report.project = mod.project;
      report.module = mod.module;
      report.submodule = mod.submodule;
      report.picQA = mod.team;
      report.isExternalQA = mod.externalQA ? mod.externalQA.isExternal : false;

      reports.push(report);
    } catch (error) {
      Logger.log('Error reading External Test Report for ' + mod.name + ': ' + error.message);
    }
  });

  return reports;
}

/**
 * POST API: Send closure email with PDF attachment
 *
 * Request body:
 * {
 *   moduleId: "spreadsheet-id",
 *   emailTo: "recipient@example.com",
 *   emailCc: "cc@example.com",  // optional
 *   emailSubject: "Test Closure Report - Project X",
 *   emailBody: "Please find attached...",
 *   attachPDF: true  // default true
 * }
 */
function sendClosureEmail_(body) {
  try {
    // Validate required fields
    if (!body.moduleId) throw new Error('moduleId is required');
    if (!body.emailTo) throw new Error('emailTo is required');
    if (!body.emailSubject) throw new Error('emailSubject is required');

    const qatmSs = SpreadsheetApp.openById(body.moduleId);
    const extReportSheet = qatmSs.getSheetByName('External Test Report');

    if (!extReportSheet) {
      throw new Error('External Test Report tab not found in QATM');
    }

    // Get report data
    const report = readExternalTestReport_(extReportSheet, qatmSs);

    // Get email sender config (departemen.qa)
    const senderEmail = getEmailSenderConfig_();

    // Generate email body if not provided
    let emailBody = body.emailBody || '';
    if (!emailBody) {
      emailBody = generateClosureEmailTemplate_(report, qatmSs);
    }

    // Prepare email options
    const emailOptions = {
      from: senderEmail,
      name: 'QA Team - INA Digital',
      cc: body.emailCc || '',
      htmlBody: emailBody
    };

    // Attach PDF if requested
    if (body.attachPDF !== false) {
      const pdfBlob = generateClosurePDFBlob_(qatmSs, report);
      if (pdfBlob) {
        emailOptions.attachments = [pdfBlob];
      }
    }

    // Send email
    GmailApp.sendEmail(
      body.emailTo,
      body.emailSubject,
      stripHtmlTags_(emailBody),  // Plain text fallback
      emailOptions
    );

    return {
      success: true,
      message: 'Email sent successfully',
      sentTo: body.emailTo,
      sentCc: body.emailCc || '',
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    Logger.log('sendClosureEmail error: ' + error.stack);
    throw new Error('Failed to send email: ' + error.message);
  }
}

/**
 * POST API: Create Gmail draft (not send)
 * Same params as sendClosureEmail_ but creates draft instead
 */
function createClosureEmailDraft_(body) {
  try {
    // Validate required fields
    if (!body.moduleId) throw new Error('moduleId is required');
    if (!body.emailTo) throw new Error('emailTo is required');
    if (!body.emailSubject) throw new Error('emailSubject is required');

    const qatmSs = SpreadsheetApp.openById(body.moduleId);
    const extReportSheet = qatmSs.getSheetByName('External Test Report');

    if (!extReportSheet) {
      throw new Error('External Test Report tab not found in QATM');
    }

    // Get report data
    const report = readExternalTestReport_(extReportSheet, qatmSs);

    // Generate email body if not provided
    let emailBody = body.emailBody || '';
    if (!emailBody) {
      emailBody = generateClosureEmailTemplate_(report, qatmSs);
    }

    // Prepare email options
    const emailOptions = {
      cc: body.emailCc || '',
      htmlBody: emailBody
    };

    // Attach PDF if requested
    if (body.attachPDF !== false) {
      const pdfBlob = generateClosurePDFBlob_(qatmSs, report);
      if (pdfBlob) {
        emailOptions.attachments = [pdfBlob];
      }
    }

    // Create draft
    const draft = GmailApp.createDraft(
      body.emailTo,
      body.emailSubject,
      stripHtmlTags_(emailBody),  // Plain text fallback
      emailOptions
    );

    return {
      success: true,
      message: 'Email draft created successfully',
      draftId: draft.getId(),
      recipient: body.emailTo,
      subject: body.emailSubject,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    Logger.log('createClosureEmailDraft error: ' + error.stack);
    throw new Error('Failed to create draft: ' + error.message);
  }
}

/**
 * Generate HTML email template for test closure
 */
function generateClosureEmailTemplate_(report, qatmSs) {
  const summary = qatmSs.getSheetByName('Summary');
  let projectName = '';
  let moduleName = '';
  let testPeriod = '';

  if (summary) {
    try {
      projectName = String(summary.getRange('B2').getValue() || '');
      moduleName = String(summary.getRange('B3').getValue() || '');
      // Try to get test period from Summary if available
      testPeriod = String(summary.getRange('B10').getValue() || '');
    } catch(e) {
      Logger.log('Error reading summary: ' + e.message);
    }
  }

  const html = `
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; font-size: 14px; color: #333; }
    .header { background-color: #1976D2; color: white; padding: 20px; text-align: center; }
    .content { padding: 20px; }
    .section { margin-bottom: 20px; }
    .section-title { font-weight: bold; font-size: 16px; color: #1976D2; margin-bottom: 10px; border-bottom: 2px solid #1976D2; padding-bottom: 5px; }
    .field { margin-bottom: 8px; }
    .field-label { font-weight: bold; display: inline-block; width: 200px; }
    .field-value { display: inline-block; }
    .status-approved { color: #2E7D32; font-weight: bold; }
    .status-review { color: #1565C0; font-weight: bold; }
    .status-rejected { color: #C62828; font-weight: bold; }
    .status-warning { color: #F57C00; font-weight: bold; }
    .vapt-blocker { background-color: #FFEBEE; border-left: 4px solid #C62828; padding: 12px; margin: 10px 0; }
    .vapt-approved { background-color: #E8F5E9; border-left: 4px solid #2E7D32; padding: 12px; margin: 10px 0; }
    .vapt-notes { background-color: #FFF3E0; border-left: 4px solid #F57C00; padding: 12px; margin: 10px 0; }
    .footer { background-color: #F5F5F5; padding: 15px; text-align: center; font-size: 12px; color: #666; margin-top: 30px; }
    table { border-collapse: collapse; width: 100%; margin-top: 10px; }
    td { padding: 8px; border: 1px solid #ddd; }
    th { padding: 8px; border: 1px solid #ddd; background-color: #f5f5f5; font-weight: bold; text-align: left; }
  </style>
</head>
<body>
  <div class="header">
    <h2>Test Closure Report</h2>
    <p>${projectName}${moduleName ? ' - ' + moduleName : ''}</p>
  </div>

  <div class="content">
    <div class="section">
      <div class="section-title">External Test Information</div>
      <div class="field">
        <span class="field-label">External Team / Vendor:</span>
        <span class="field-value">${report.externalTeam || '-'}</span>
      </div>
      <div class="field">
        <span class="field-label">Overall Status:</span>
        <span class="field-value ${getStatusClass_(report.overallStatus)}">${report.overallStatus || 'Not Started'}</span>
      </div>
      <div class="field">
        <span class="field-label">Reviewer:</span>
        <span class="field-value">${report.reviewer || '-'}</span>
      </div>
      <div class="field">
        <span class="field-label">Review Date:</span>
        <span class="field-value">${report.reviewDate || '-'}</span>
      </div>
    </div>

    <div class="section">
      <div class="section-title">Testing Evidence & Status</div>
      <table>
        <tr>
          <td><strong>Test Type</strong></td>
          <td><strong>Evidence URL</strong></td>
          <td><strong>Review Status</strong></td>
        </tr>
        <tr>
          <td>Functional Testing</td>
          <td><a href="${report.functionalEvidenceUrl || '#'}">${report.functionalEvidenceUrl ? 'View Evidence' : 'N/A'}</a></td>
          <td class="${getStatusClass_(report.functionalReviewStatus)}">${report.functionalReviewStatus || 'Not Started'}</td>
        </tr>
        <tr>
          <td>Performance Testing</td>
          <td><a href="${report.performanceEvidenceUrl || '#'}">${report.performanceEvidenceUrl ? 'View Evidence' : 'N/A'}</a></td>
          <td class="${getStatusClass_(report.performanceReviewStatus)}">${report.performanceReviewStatus || 'Not Started'}</td>
        </tr>
        <tr>
          <td>VAPT (Security)</td>
          <td><a href="${report.vaptEvidenceUrl || '#'}">${report.vaptEvidenceUrl ? 'View Evidence' : 'N/A'}</a></td>
          <td class="${getStatusClass_(report.vaptReviewStatus)}">${report.vaptReviewStatus || 'Not Started'}</td>
        </tr>
      </table>
    </div>

    <div class="section">
      <div class="section-title">🔒 VAPT Security Findings Summary</div>
      ${report.vaptTotal > 0 ? `
        ${report.vaptBlockerCount > 0 ? `
          <div class="vapt-blocker">
            <strong>⚠️ BLOCKER FINDINGS DETECTED</strong>
            <p style="margin: 8px 0 0 0; font-size: 13px;">
              ${report.vaptBlockerCount} security finding(s) require resolution before approval:
              Critical: ${report.vaptCritical} | High: ${report.vaptHigh} | Medium: ${report.vaptMedium}
            </p>
          </div>
        ` : report.vaptNonBlockerCount > 0 ? `
          <div class="vapt-notes">
            <strong>✅ APPROVED WITH NOTES</strong>
            <p style="margin: 8px 0 0 0; font-size: 13px;">
              No blocker findings detected. ${report.vaptNonBlockerCount} informational finding(s) noted for future improvement:
              Low: ${report.vaptLow} | Informational: ${report.vaptInformational}
            </p>
          </div>
        ` : `
          <div class="vapt-approved">
            <strong>✅ FULLY APPROVED</strong>
            <p style="margin: 8px 0 0 0; font-size: 13px;">
              All ${report.vaptTotal} finding(s) have been resolved. No security blockers remaining.
            </p>
          </div>
        `}

        <table>
          <tr>
            <th colspan="2" style="background-color: #FFEBEE; color: #C62828;">BLOCKER FINDINGS (Require Resolution)</th>
          </tr>
          <tr>
            <td width="50%"><strong>Critical Risk:</strong></td>
            <td width="50%">${report.vaptCritical}</td>
          </tr>
          <tr>
            <td><strong>High Risk:</strong></td>
            <td>${report.vaptHigh}</td>
          </tr>
          <tr>
            <td><strong>Medium Risk:</strong></td>
            <td>${report.vaptMedium}</td>
          </tr>
          <tr style="background-color: #FFEBEE;">
            <td><strong>Total Blocker:</strong></td>
            <td><strong>${report.vaptBlockerCount}</strong></td>
          </tr>
          <tr>
            <th colspan="2" style="background-color: #E8F5E9; color: #2E7D32;">NON-BLOCKER FINDINGS (For Reference)</th>
          </tr>
          <tr>
            <td><strong>Low Risk:</strong></td>
            <td>${report.vaptLow}</td>
          </tr>
          <tr>
            <td><strong>Informational:</strong></td>
            <td>${report.vaptInformational}</td>
          </tr>
          <tr style="background-color: #E8F5E9;">
            <td><strong>Total Non-Blocker:</strong></td>
            <td><strong>${report.vaptNonBlockerCount}</strong></td>
          </tr>
          <tr>
            <th colspan="2" style="background-color: #E3F2FD; color: #1976D2;">STATUS FIX PROGRESS</th>
          </tr>
          <tr>
            <td><strong>Todo:</strong></td>
            <td>${report.vaptTodo}</td>
          </tr>
          <tr>
            <td><strong>On Progress:</strong></td>
            <td>${report.vaptOnProgress}</td>
          </tr>
          <tr>
            <td><strong>Done:</strong></td>
            <td>${report.vaptDone}</td>
          </tr>
          <tr>
            <th colspan="2" style="background-color: #FFF3E0; color: #F57C00;">RE-VAPT STATUS</th>
          </tr>
          <tr>
            <td><strong>Open:</strong></td>
            <td>${report.vaptOpen}</td>
          </tr>
          <tr>
            <td><strong>Closed:</strong></td>
            <td>${report.vaptClosed}</td>
          </tr>
        </table>

        <div style="margin-top: 15px; padding: 10px; background-color: #F5F5F5; border-radius: 4px;">
          <strong>📌 Important Notes:</strong>
          <ul style="margin: 8px 0 0 0; font-size: 13px;">
            <li><strong>Blocker findings (Critical/High/Medium)</strong> must be resolved before final approval</li>
            <li><strong>Non-blocker findings (Low/Informational)</strong> can be approved with notes for future improvement</li>
            <li>Please refer to the attached PDF for detailed findings and remediation recommendations</li>
          </ul>
        </div>
      ` : `
        <p style="color: #666; font-style: italic;">No VAPT data available for this module.</p>
      `}
    </div>

    <div class="section">
      <div class="section-title">Notes</div>
      <p>${report.notes || 'No additional notes.'}</p>
    </div>

    <div class="section">
      <p>Please find the detailed test closure report attached as PDF.</p>
      <p>For any questions or clarifications, please contact the QA team.</p>
    </div>
  </div>

  <div class="footer">
    <p>This is an automated email from QA Dashboard - INA Digital</p>
    <p>Generated on ${new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' })} WIB</p>
  </div>
</body>
</html>
`;

  return html;
}

/**
 * Get status CSS class for email styling
 */
function getStatusClass_(status) {
  const s = String(status || '').toLowerCase();
  if (s.includes('approved') || s.includes('ready')) return 'status-approved';
  if (s.includes('review')) return 'status-review';
  if (s.includes('rejected')) return 'status-rejected';
  return '';
}

/**
 * Generate PDF blob from QATM spreadsheet (simplified version)
 * In production, you might want more sophisticated PDF generation
 */
function generateClosurePDFBlob_(qatmSs, report) {
  try {
    // Convert spreadsheet to PDF
    // Using Summary sheet as the main PDF content
    const summarySheet = qatmSs.getSheetByName('Summary');
    if (!summarySheet) {
      Logger.log('Summary sheet not found, skipping PDF attachment');
      return null;
    }

    const url = 'https://docs.google.com/spreadsheets/d/' + qatmSs.getId() + '/export?format=pdf&gid=' + summarySheet.getSheetId();

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    const blob = response.getBlob();
    const fileName = 'Test_Closure_Report_' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd') + '.pdf';
    blob.setName(fileName);

    return blob;
  } catch (error) {
    Logger.log('PDF generation error: ' + error.message);
    return null;
  }
}

/**
 * Get email sender config from Config sheet or Script Properties
 */
function getEmailSenderConfig_() {
  try {
    const ss = getDashboardSpreadsheet_();
    const config = ss.getSheetByName('Config');

    if (config) {
      // Try to read from Config sheet (you might need to add this column)
      // For now, using Script Properties or default
    }

    // Check Script Properties
    const props = PropertiesService.getScriptProperties();
    const senderEmail = props.getProperty('EMAIL_SENDER_ADDRESS');

    if (senderEmail) {
      return senderEmail;
    }

    // Default to departemen.qa@inadigital.co.id
    return 'departemen.qa@inadigital.co.id';
  } catch (error) {
    Logger.log('getEmailSenderConfig error: ' + error.message);
    return 'departemen.qa@inadigital.co.id';
  }
}

/**
 * Strip HTML tags for plain text email fallback
 */
function stripHtmlTags_(html) {
  return html.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
}
