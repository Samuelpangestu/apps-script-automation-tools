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
 * Fetch Ad Hoc VAPT data (C1:Y100)
 */
function fetchAdHocVAPTData_(vaptSs) {
  try {
    const sheet = vaptSs.getSheetByName('Ad Hoc VAPT');
    if (!sheet) {
      Logger.log('⚠️ Ad Hoc VAPT tab not found');
      return [];
    }

    const data = sheet.getRange('C1:Y100').getValues();  // C to Y = 23 columns (No, Aplikasi, ..., Prod, Formula Updated)
    const entries = [];

    // Skip header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows (check if Aplikasi is empty)
      if (!row[1] || String(row[1]).trim() === '') continue;

      // Parse row data
      // Columns: C=No, D=Aplikasi, E=PIC VAPT, F=Scope, G=VAPT Status, H=Report,
      //          I=RTR(total), J-N=Open(Crit,High,Med,Low,Info), O-S=Closed(Crit,High,Med,Low,Info),
      //          T=Prod, U=Formula Updated
      const entry = {
        type: 'Ad Hoc',
        aplikasi: String(row[1]).trim(),      // D (index 1)
        picVapt: String(row[2]).trim(),       // E (index 2)
        scope: String(row[3]).trim(),         // F (index 3)
        status: String(row[4]).trim(),        // G (index 4)
        report: String(row[5]).trim(),        // H (index 5)
        readyToRetest: {
          critical: 0,  // I is total only, no breakdown
          high: 0,
          medium: 0,
          low: 0,
          info: 0
        },
        open: {
          critical: Number(row[7]) || 0,      // J (index 7)
          high: Number(row[8]) || 0,          // K (index 8)
          medium: Number(row[9]) || 0,        // L (index 9)
          low: Number(row[10]) || 0,          // M (index 10)
          info: Number(row[11]) || 0          // N (index 11)
        },
        closed: {
          critical: Number(row[12]) || 0,     // O (index 12)
          high: Number(row[13]) || 0,         // P (index 13)
          medium: Number(row[14]) || 0,       // Q (index 14)
          low: Number(row[15]) || 0,          // R (index 15)
          info: Number(row[16]) || 0          // S (index 16)
        },
        prod: row[17] === true || String(row[17]).toUpperCase() === 'TRUE',         // T (index 17)
        formulaUpdated: row[18] === true || String(row[18]).toUpperCase() === 'TRUE' // U (index 18)
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
 * Fetch Regular VAPT data (C1:AF100)
 */
function fetchRegularVAPTData_(vaptSs) {
  try {
    const sheet = vaptSs.getSheetByName('Regular VAPT');
    if (!sheet) {
      Logger.log('⚠️ Regular VAPT tab not found');
      return [];
    }

    const data = sheet.getRange('C1:AF100').getValues();  // C to AF = 30 columns (Aplikasi, ..., Prod, Formula Updated)
    const entries = [];

    // Skip header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows (check if Aplikasi is empty)
      if (!row[0] || String(row[0]).trim() === '') continue;

      // Parse row data
      // Columns: C=Aplikasi, D=Product Owner, E=VAPT MSSP, F=MSSP Report,
      //          G=MSSP Checklist Status, H=MSSP Checklist Report, I=Internal VAPT Status, J=Report,
      //          K=Report to PMO, L=MSSP Reported, M=Internal Reported,
      //          N=???, O=???, P=RTR(total),
      //          Q-U=Open(Crit,High,Med,Low,Info), V-Z=Closed(Crit,High,Med,Low,Info),
      //          AA=Prod, AB=Formula Updated
      const entry = {
        type: 'Regular',
        aplikasi: String(row[0]).trim(),       // C (index 0)
        picVapt: String(row[1]).trim(),        // D (index 1) - Product Owner actually, but use as picVapt
        picQa: '',                             // Not available in Regular VAPT
        productOwner: String(row[1]).trim(),   // D (index 1)
        status: String(row[6]).trim(),         // I - Internal VAPT Status (index 6)
        report: String(row[7]).trim(),         // J - Report (index 7)
        readyToRetest: {
          critical: 0,  // P is total only, no breakdown
          high: 0,
          medium: 0,
          low: 0,
          info: 0
        },
        open: {
          critical: Number(row[14]) || 0,      // Q (index 14)
          high: Number(row[15]) || 0,          // R (index 15)
          medium: Number(row[16]) || 0,        // S (index 16)
          low: Number(row[17]) || 0,           // T (index 17)
          info: Number(row[18]) || 0           // U (index 18)
        },
        closed: {
          critical: Number(row[19]) || 0,      // V (index 19)
          high: Number(row[20]) || 0,          // W (index 20)
          medium: Number(row[21]) || 0,        // X (index 21)
          low: Number(row[22]) || 0,           // Y (index 22)
          info: Number(row[23]) || 0           // Z (index 23)
        },
        prod: row[24] === true || String(row[24]).toUpperCase() === 'TRUE',         // AA (index 24)
        formulaUpdated: row[25] === true || String(row[25]).toUpperCase() === 'TRUE' // AB (index 25)
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
  // AGGREGATE: Combine multiple rows for same aplikasi
  const aggregatedAdHoc = aggregateByAplikasi_(adHocData);
  const aggregatedRegular = aggregateByAplikasi_(regularData);

  // Combine all entries
  const allEntries = [...aggregatedAdHoc, ...aggregatedRegular];

  // Calculate summaries
  const adHocSummary = calculateVAPTSummary_(aggregatedAdHoc);
  const regularSummary = calculateVAPTSummary_(aggregatedRegular);
  const combinedSummary = calculateVAPTSummary_(allEntries);

  return {
    table: allEntries,
    summary: combinedSummary,
    adHocSummary: adHocSummary,
    regularSummary: regularSummary
  };
}

/**
 * Aggregate multiple rows with same aplikasi name
 * Sum all findings (RTR, Open, Closed) for each severity
 */
function aggregateByAplikasi_(entries) {
  if (entries.length === 0) return [];

  const grouped = {};

  entries.forEach(entry => {
    const key = entry.aplikasi.toLowerCase().trim();

    if (!grouped[key]) {
      // First entry for this aplikasi - clone it
      grouped[key] = {
        type: entry.type,
        aplikasi: entry.aplikasi,
        picVapt: entry.picVapt,
        scope: entry.scope || '',
        status: entry.status,
        report: entry.report,
        readyToRetest: {
          critical: entry.readyToRetest.critical,
          high: entry.readyToRetest.high,
          medium: entry.readyToRetest.medium,
          low: entry.readyToRetest.low,
          info: entry.readyToRetest.info
        },
        open: {
          critical: entry.open.critical,
          high: entry.open.high,
          medium: entry.open.medium,
          low: entry.open.low,
          info: entry.open.info
        },
        closed: {
          critical: entry.closed.critical,
          high: entry.closed.high,
          medium: entry.closed.medium,
          low: entry.closed.low,
          info: entry.closed.info
        },
        prod: entry.prod,
        formulaUpdated: entry.formulaUpdated,
        _count: 1
      };
    } else {
      // Subsequent entry - sum findings
      grouped[key].readyToRetest.critical += entry.readyToRetest.critical;
      grouped[key].readyToRetest.high += entry.readyToRetest.high;
      grouped[key].readyToRetest.medium += entry.readyToRetest.medium;
      grouped[key].readyToRetest.low += entry.readyToRetest.low;
      grouped[key].readyToRetest.info += entry.readyToRetest.info;

      grouped[key].open.critical += entry.open.critical;
      grouped[key].open.high += entry.open.high;
      grouped[key].open.medium += entry.open.medium;
      grouped[key].open.low += entry.open.low;
      grouped[key].open.info += entry.open.info;

      grouped[key].closed.critical += entry.closed.critical;
      grouped[key].closed.high += entry.closed.high;
      grouped[key].closed.medium += entry.closed.medium;
      grouped[key].closed.low += entry.closed.low;
      grouped[key].closed.info += entry.closed.info;

      // Combine PIC VAPT if different
      if (entry.picVapt && !grouped[key].picVapt.includes(entry.picVapt)) {
        grouped[key].picVapt += ', ' + entry.picVapt;
      }

      // Update status to latest (prefer "In Progress" over "Done")
      if (entry.status === 'In Progress') {
        grouped[key].status = 'In Progress';
      }

      // Prod: TRUE if any entry is TRUE
      if (entry.prod) {
        grouped[key].prod = true;
      }

      // Formula Updated: TRUE if any entry is TRUE
      if (entry.formulaUpdated) {
        grouped[key].formulaUpdated = true;
      }

      grouped[key]._count++;
    }
  });

  // Convert back to array
  const result = Object.values(grouped);

  Logger.log('Aggregation: ' + entries.length + ' rows → ' + result.length + ' unique aplikasi');

  return result;
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
