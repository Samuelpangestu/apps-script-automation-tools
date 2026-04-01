/**
 * VAPTDataFetch.js - Fetch and Process VAPT Data
 *
 * Fetch VAPT findings from source spreadsheet (Ad Hoc + Regular VAPT tabs)
 * Process and combine data for dashboard display
 */

// VAPT Spreadsheet ID (from Config tab)
const VAPT_SPREADSHEET_ID = '17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk';

// ═══════════════════════════════════════════════════════════════════════
// MAIN REFRESH FUNCTION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Refresh VAPT data from source spreadsheet
 * Called by refreshDashboard()
 */
function refreshVAPTData() {
  try {
    Logger.log('Starting VAPT data refresh...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');

    if (!configSheet) {
      Logger.log('⚠️ Config tab not found - skipping VAPT refresh');
      return;
    }

    // Get VAPT spreadsheet ID from Config (will be added later)
    // For now, use hardcoded ID
    const vaptSpreadsheetId = VAPT_SPREADSHEET_ID;

    if (!vaptSpreadsheetId || vaptSpreadsheetId === 'PASTE_VAPT_SPREADSHEET_ID_HERE') {
      Logger.log('⚠️ VAPT Spreadsheet ID not configured - skipping VAPT refresh');
      return;
    }

    // Fetch data from VAPT spreadsheet
    Logger.log('Fetching VAPT data from spreadsheet: ' + vaptSpreadsheetId);
    const vaptData = fetchAndProcessVAPTData_(vaptSpreadsheetId);

    // Write to VAPT tab
    Logger.log('Writing VAPT data to dashboard...');
    writeVAPT(ss, vaptData);

    // Append to history
    Logger.log('Appending to VAPT History...');
    appendVAPTHistory(ss, vaptData);

    Logger.log('✅ VAPT data refresh complete');
  } catch (error) {
    Logger.log('❌ ERROR refreshing VAPT data: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// FETCH AND PROCESS VAPT DATA
// ═══════════════════════════════════════════════════════════════════════

/**
 * Fetch and process VAPT data from both tabs
 */
function fetchAndProcessVAPTData_(vaptSpreadsheetId) {
  const vaptSs = SpreadsheetApp.openById(vaptSpreadsheetId);

  // Fetch Ad Hoc VAPT data
  Logger.log('Fetching Ad Hoc VAPT data...');
  const adHocData = fetchAdHocVAPTData_(vaptSs);
  Logger.log('Ad Hoc VAPT entries: ' + adHocData.length);

  // Fetch Regular VAPT data
  Logger.log('Fetching Regular VAPT data...');
  const regularData = fetchRegularVAPTData_(vaptSs);
  Logger.log('Regular VAPT entries: ' + regularData.length);

  // Process and combine
  Logger.log('Processing and combining VAPT data...');
  const processedData = processVAPTData_(adHocData, regularData);

  return processedData;
}

/**
 * Fetch Ad Hoc VAPT data (C1:U30)
 */
function fetchAdHocVAPTData_(vaptSs) {
  try {
    const sheet = vaptSs.getSheetByName('Ad Hoc VAPT');
    if (!sheet) {
      Logger.log('⚠️ Ad Hoc VAPT tab not found');
      return [];
    }

    const data = sheet.getRange('C1:U100').getValues();  // Extended range to capture all rows
    const entries = [];

    // Skip header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows (check if Aplikasi is empty)
      if (!row[1] || String(row[1]).trim() === '') continue;

      // Parse row data
      // Columns: No, Aplikasi, PIC VAPT, Scope, VAPT Status, Report,
      //          RTR(Crit,High,Med,Low,Info), Open(Crit,High,Med,Low,Info), Closed(Crit,High,Med,Low,Info),
      //          Prod, Formula Updated
      const entry = {
        type: 'Ad Hoc',
        aplikasi: String(row[1]).trim(),
        picVapt: String(row[2]).trim(),
        scope: String(row[3]).trim(),
        status: String(row[4]).trim(),
        report: String(row[5]).trim(),
        readyToRetest: {
          critical: Number(row[6]) || 0,
          high: Number(row[7]) || 0,
          medium: Number(row[8]) || 0,
          low: Number(row[9]) || 0,
          info: Number(row[10]) || 0
        },
        open: {
          critical: Number(row[11]) || 0,
          high: Number(row[12]) || 0,
          medium: Number(row[13]) || 0,
          low: Number(row[14]) || 0,
          info: Number(row[15]) || 0
        },
        closed: {
          critical: Number(row[16]) || 0,
          high: Number(row[17]) || 0,
          medium: Number(row[18]) || 0,
          low: Number(row[19]) || 0,
          info: Number(row[20]) || 0
        },
        prod: row[21] === true || String(row[21]).toUpperCase() === 'TRUE',
        formulaUpdated: row[22] === true || String(row[22]).toUpperCase() === 'TRUE'
      };

      entries.push(entry);
    }

    return entries;
  } catch (error) {
    Logger.log('❌ ERROR fetching Ad Hoc VAPT: ' + error.toString());
    return [];
  }
}

/**
 * Fetch Regular VAPT data (C1:AB30)
 */
function fetchRegularVAPTData_(vaptSs) {
  try {
    const sheet = vaptSs.getSheetByName('Regular VAPT');
    if (!sheet) {
      Logger.log('⚠️ Regular VAPT tab not found');
      return [];
    }

    const data = sheet.getRange('C1:AB100').getValues();  // Extended range
    const entries = [];

    // Skip header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows (check if Aplikasi is empty)
      if (!row[0] || String(row[0]).trim() === '') continue;

      // Parse row data
      // Columns: Aplikasi, PIC VAPT, PIC QA, Product Owner, VAPT MSSP Status, MSSP Report,
      //          MSSP Checklist Status, MSSP Checklist Report, Internal VAPT Status, Report,
      //          Report to PMO, MSSP Reported, Internal Reported,
      //          RTR(Crit,High,Med,Low,Info), Open(Crit,High,Med,Low,Info), Closed(Crit,High,Med,Low,Info),
      //          Prod, Formula Updated
      const entry = {
        type: 'Regular',
        aplikasi: String(row[0]).trim(),
        picVapt: String(row[1]).trim(),
        picQa: String(row[2]).trim(),
        productOwner: String(row[3]).trim(),
        status: String(row[8]).trim(),  // Internal VAPT Status (col 8)
        report: String(row[9]).trim(),   // Report (col 9)
        readyToRetest: {
          critical: Number(row[13]) || 0,
          high: Number(row[14]) || 0,
          medium: Number(row[15]) || 0,
          low: Number(row[16]) || 0,
          info: Number(row[17]) || 0
        },
        open: {
          critical: Number(row[18]) || 0,
          high: Number(row[19]) || 0,
          medium: Number(row[20]) || 0,
          low: Number(row[21]) || 0,
          info: Number(row[22]) || 0
        },
        closed: {
          critical: Number(row[23]) || 0,
          high: Number(row[24]) || 0,
          medium: Number(row[25]) || 0,
          low: Number(row[26]) || 0,
          info: Number(row[27]) || 0
        },
        prod: row[28] === true || String(row[28]).toUpperCase() === 'TRUE',
        formulaUpdated: row[29] === true || String(row[29]).toUpperCase() === 'TRUE'
      };

      entries.push(entry);
    }

    return entries;
  } catch (error) {
    Logger.log('❌ ERROR fetching Regular VAPT: ' + error.toString());
    return [];
  }
}

/**
 * Process and combine Ad Hoc + Regular VAPT data
 * Calculate summary metrics
 */
function processVAPTData_(adHocData, regularData) {
  // Combine all entries
  const allEntries = [...adHocData, ...regularData];

  // Calculate summaries
  const adHocSummary = calculateVAPTSummary_(adHocData);
  const regularSummary = calculateVAPTSummary_(regularData);
  const combinedSummary = calculateVAPTSummary_(allEntries);

  return {
    table: allEntries,
    summary: combinedSummary,
    adHocSummary: adHocSummary,
    regularSummary: regularSummary
  };
}

/**
 * Calculate summary metrics from VAPT entries
 */
function calculateVAPTSummary_(entries) {
  const summary = {
    totalApps: entries.length,
    totalFindings: 0,
    blocker: 0,  // NEW: Medium + High + Critical Open
    blockerBreakdown: {  // NEW: Blocker by severity
      critical: 0,
      high: 0,
      medium: 0
    },
    otherOpen: {  // NEW: Non-blocker open findings
      low: 0,
      info: 0
    },
    totalClosed: 0,  // NEW: Total closed findings
    bySeverity: {
      critical: 0,
      high: 0,
      medium: 0,
      low: 0,
      info: 0
    },
    byStatus: {
      readyToRetest: 0,
      open: 0,
      closed: 0
    },
    byVaptStatus: {
      done: 0,
      inProgress: 0,
      notStarted: 0,
      todo: 0
    },
    prodCount: 0
  };

  entries.forEach(entry => {
    // BLOCKER CALCULATION (Medium + High + Critical Open)
    const blockerForEntry = (entry.open.medium || 0) + (entry.open.high || 0) + (entry.open.critical || 0);
    summary.blocker += blockerForEntry;

    // Blocker Breakdown
    summary.blockerBreakdown.critical += (entry.open.critical || 0);
    summary.blockerBreakdown.high += (entry.open.high || 0);
    summary.blockerBreakdown.medium += (entry.open.medium || 0);

    // Other Open Findings (Low + Info)
    summary.otherOpen.low += (entry.open.low || 0);
    summary.otherOpen.info += (entry.open.info || 0);

    // Total Closed
    const closedTotal = (entry.closed.critical || 0) + (entry.closed.high || 0) + (entry.closed.medium || 0) + (entry.closed.low || 0) + (entry.closed.info || 0);
    summary.totalClosed += closedTotal;

    // By Severity (sum across all statuses)
    summary.bySeverity.critical += (entry.readyToRetest.critical + entry.open.critical + entry.closed.critical);
    summary.bySeverity.high += (entry.readyToRetest.high + entry.open.high + entry.closed.high);
    summary.bySeverity.medium += (entry.readyToRetest.medium + entry.open.medium + entry.closed.medium);
    summary.bySeverity.low += (entry.readyToRetest.low + entry.open.low + entry.closed.low);
    summary.bySeverity.info += (entry.readyToRetest.info + entry.open.info + entry.closed.info);

    // By Status (sum across all severities)
    const rtrTotal = entry.readyToRetest.critical + entry.readyToRetest.high + entry.readyToRetest.medium + entry.readyToRetest.low + entry.readyToRetest.info;
    const openTotal = entry.open.critical + entry.open.high + entry.open.medium + entry.open.low + entry.open.info;

    summary.byStatus.readyToRetest += rtrTotal;
    summary.byStatus.open += openTotal;
    summary.byStatus.closed += closedTotal;

    // VAPT Status
    const status = String(entry.status).toLowerCase();
    if (status === 'done') summary.byVaptStatus.done++;
    else if (status === 'in progress') summary.byVaptStatus.inProgress++;
    else if (status === 'not started') summary.byVaptStatus.notStarted++;
    else if (status === 'todo') summary.byVaptStatus.todo++;

    // Production count
    if (entry.prod) summary.prodCount++;
  });

  // Total findings
  summary.totalFindings = summary.bySeverity.critical + summary.bySeverity.high + summary.bySeverity.medium + summary.bySeverity.low + summary.bySeverity.info;

  return summary;
}
