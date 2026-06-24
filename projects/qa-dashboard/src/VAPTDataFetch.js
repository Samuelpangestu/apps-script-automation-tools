/**
 * VAPTDataFetch.js - Fetch and Process VAPT Data (Per-Project)
 *
 * Fetch VAPT findings from multiple per-project spreadsheets
 * Process and combine data for dashboard display
 */

// ═══════════════════════════════════════════════════════════════════════
// MAIN REFRESH FUNCTION
// ═══════════════════════════════════════════════════════════════════════

/**
 * Refresh VAPT data from per-project spreadsheets
 * Called by refreshDashboard()
 */
function refreshVAPTData() {
  // DISABLED: External VAPT fetch now replaced by per-QATM Summary data
  // All VAPT data comes from individual QATM Summary tabs
  // VAPT tab and VAPT History are populated by per-QATM data in refreshDashboard()
  Logger.log('ℹ️ External VAPT refresh disabled - using per-QATM Summary data only');
  return;

  /* ORIGINAL CODE - DISABLED
  try {
    Logger.log('Starting per-project VAPT data refresh...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');

    if (!configSheet) {
      Logger.log('⚠️ Config tab not found - skipping VAPT refresh');
      return;
    }

    // Read Config to get projects with VAPT enabled
    const cfgData = configSheet.getDataRange().getValues();
    const allProjectData = [];

    // Loop through Config rows (skip header rows 1-3)
    // Row 4 = index 3 = first data row
    // Stop immediately when project name is empty (no more data)
    for (let i = 3; i < cfgData.length; i++) {
      const projectName = String(cfgData[i][2]).trim(); // Col C (index 2) = Project Name

      // Early exit: Stop if project name is empty (no more data)
      if (!projectName) {
        break;
      }

      const vaptSpreadsheetId = String(cfgData[i][21]).trim(); // Col V (index 21) = VAPT Spreadsheet ID
      const vaptEnabled = cfgData[i][22] === true; // Col W (index 22) = Enable VAPT

      // Skip if VAPT not enabled or no spreadsheet ID
      if (!vaptEnabled || !vaptSpreadsheetId) {
        continue;
      }

      Logger.log('✅ Fetching VAPT for project: ' + projectName);

      try {
        // Fetch data from this project's VAPT spreadsheet
        const projectVaptData = fetchAndProcessVAPTData_(vaptSpreadsheetId, projectName);
        allProjectData.push(projectVaptData);
      } catch (error) {
        Logger.log('⚠️ Error fetching VAPT for ' + projectName + ': ' + error.toString());
        // Continue with other projects
      }
    }

    if (allProjectData.length === 0) {
      Logger.log('⚠️ No projects with VAPT enabled - skipping VAPT refresh');
      return;
    }

    // Combine all project data
    Logger.log('Combining VAPT data from ' + allProjectData.length + ' project(s)...');
    const combinedData = combineProjectVAPTData_(allProjectData);

    // Write to VAPT tab
    Logger.log('Writing VAPT data to dashboard...');
    writeVAPT(ss, combinedData);

    // NOTE: VAPT History is now handled by per-submodule data from QATM Summary
    // See refreshDashboard() -> appendVAPTHistory(ss, allData)
    // Old per-project history call removed to avoid conflict

    Logger.log('✅ VAPT data refresh complete');
  } catch (error) {
    Logger.log('❌ ERROR refreshing VAPT data: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
  }
  */
}

/**
 * Combine VAPT data from multiple projects
 */
function combineProjectVAPTData_(allProjectData) {
  const combinedTable = [];
  let combinedSummary = initSummary_();
  const projectSummaries = [];

  allProjectData.forEach(projectData => {
    // Add all table entries (with project name already included)
    combinedTable.push(...projectData.table);

    // Accumulate summary
    combinedSummary = mergeSummaries_(combinedSummary, projectData.summary);

    // Store per-project summary for history
    projectSummaries.push({
      project: projectData.projectName,
      summary: projectData.summary
    });
  });

  return {
    table: combinedTable,
    summary: combinedSummary,
    projectSummaries: projectSummaries
  };
}

/**
 * Initialize empty summary object
 */
function initSummary_() {
  return {
    blocker: 0,
    totalFindings: 0,
    bySeverity: { critical: 0, high: 0, medium: 0, low: 0, info: 0 },
    byStatus: { readyToRetest: 0, open: 0, closed: 0 },
    totalApps: 0,
    byVaptStatus: { done: 0, inProgress: 0 },
    prodCount: 0
  };
}

/**
 * Merge two summary objects
 */
function mergeSummaries_(s1, s2) {
  return {
    blocker: (s1.blocker || 0) + (s2.blocker || 0),
    totalFindings: (s1.totalFindings || 0) + (s2.totalFindings || 0),
    bySeverity: {
      critical: (s1.bySeverity.critical || 0) + (s2.bySeverity.critical || 0),
      high: (s1.bySeverity.high || 0) + (s2.bySeverity.high || 0),
      medium: (s1.bySeverity.medium || 0) + (s2.bySeverity.medium || 0),
      low: (s1.bySeverity.low || 0) + (s2.bySeverity.low || 0),
      info: (s1.bySeverity.info || 0) + (s2.bySeverity.info || 0)
    },
    byStatus: {
      readyToRetest: (s1.byStatus.readyToRetest || 0) + (s2.byStatus.readyToRetest || 0),
      open: (s1.byStatus.open || 0) + (s2.byStatus.open || 0),
      closed: (s1.byStatus.closed || 0) + (s2.byStatus.closed || 0)
    },
    totalApps: (s1.totalApps || 0) + (s2.totalApps || 0),
    byVaptStatus: {
      done: (s1.byVaptStatus.done || 0) + (s2.byVaptStatus.done || 0),
      inProgress: (s1.byVaptStatus.inProgress || 0) + (s2.byVaptStatus.inProgress || 0)
    },
    prodCount: (s1.prodCount || 0) + (s2.prodCount || 0)
  };
}

// ═══════════════════════════════════════════════════════════════════════
// FETCH AND PROCESS VAPT DATA
// ═══════════════════════════════════════════════════════════════════════

/**
 * Fetch and process VAPT data from VAPT BGN - Helper tab
 * @param {string} vaptSpreadsheetId - VAPT spreadsheet ID
 * @param {string} projectName - Project name to tag all entries
 */
function fetchAndProcessVAPTData_(vaptSpreadsheetId, projectName) {
  const vaptSs = SpreadsheetApp.openById(vaptSpreadsheetId);

  // Fetch VAPT data from VAPT BGN - Helper tab (E4:I36)
  Logger.log('  Fetching VAPT data from VAPT BGN - Helper tab...');
  const vaptData = fetchVAPTBGNHelperData_(vaptSs);
  Logger.log('  VAPT entries: ' + vaptData.length);

  // Process data
  Logger.log('  Processing VAPT data...');
  const processedData = processSimpleVAPTData_(vaptData);

  // Add project name to all table entries
  processedData.table.forEach(entry => {
    entry.project = projectName;
  });

  // Add project name to result
  processedData.projectName = projectName;

  return processedData;
}

/**
 * Fetch VAPT data from VAPT BGN - Helper tab (B4:I36)
 * Simple format: Aplikasi (B) | Critical (G) | High (H) | Medium (I)
 */
function fetchVAPTBGNHelperData_(vaptSs) {
  try {
    const sheet = vaptSs.getSheetByName('VAPT BGN - Helper');
    if (!sheet) {
      Logger.log('⚠️ VAPT BGN - Helper tab not found');
      return [];
    }

    // Read B4:I36 to get Aplikasi (B) + Critical/High/Medium (G/H/I)
    const data = sheet.getRange('B4:I36').getValues();  // B to I = 8 columns
    const entries = [];

    Logger.log('  📋 Reading VAPT data from B4:I36...');

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowNumber = i + 4;  // Actual row number in sheet (4-based)

      // B4:I36 indices: B=0, C=1, D=2, E=3, F=4, G=5, H=6, I=7
      const aplikasi = String(row[0] || '').trim();
      const critical = Number(row[5]) || 0;  // G (index 5 from B)
      const high = Number(row[6]) || 0;      // H (index 6 from B)
      const medium = Number(row[7]) || 0;    // I (index 7 from B)

      // Skip rows with no findings (all 0)
      if (critical === 0 && high === 0 && medium === 0) continue;

      // Skip rows with no aplikasi name
      if (!aplikasi) continue;

      Logger.log('  Row ' + rowNumber + ': ' + aplikasi + ' → C=' + critical + ' H=' + high + ' M=' + medium + ' (Blocker=' + (critical + high + medium) + ')');

      const entry = {
        aplikasi: aplikasi,
        open: {
          critical: critical,
          high: high,
          medium: medium,
          low: 0,      // Not in G4:I36 range
          info: 0      // Not in G4:I36 range
        }
      };

      entries.push(entry);
    }

    Logger.log('  ✅ Fetched ' + entries.length + ' aplikasi with findings');

    return entries;
  } catch (error) {
    Logger.log('❌ ERROR fetching VAPT BGN Helper: ' + error.toString());
    return [];
  }
}

/**
 * Process simple VAPT data from VAPT BGN - Helper
 * Calculate summary metrics
 */
function processSimpleVAPTData_(vaptData) {
  const summary = calculateSimpleVAPTSummary_(vaptData);

  return {
    table: vaptData,
    summary: summary
  };
}

/**
 * Calculate summary metrics from simple VAPT entries
 * Simple format: only has aplikasi and open.{critical, high, medium}
 */
function calculateSimpleVAPTSummary_(entries) {
  const summary = {
    totalApps: entries.length,
    totalFindings: 0,
    blocker: 0,  // Critical + High + Medium Open
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

  let appsWithBlocker = 0;

  entries.forEach(entry => {
    // BLOCKER CALCULATION (Critical + High + Medium Open)
    const critical = entry.open.critical || 0;
    const high = entry.open.high || 0;
    const medium = entry.open.medium || 0;
    const low = entry.open.low || 0;
    const info = entry.open.info || 0;

    const blockerForEntry = critical + high + medium;
    summary.blocker += blockerForEntry;

    if (blockerForEntry > 0) {
      appsWithBlocker++;
    }

    // By Severity (only open status in simple format)
    summary.bySeverity.critical += critical;
    summary.bySeverity.high += high;
    summary.bySeverity.medium += medium;
    summary.bySeverity.low += low;
    summary.bySeverity.info += info;

    // By Status (all are open in simple format)
    const openTotal = critical + high + medium + low + info;
    summary.byStatus.open += openTotal;
  });

  // Total findings
  summary.totalFindings = summary.bySeverity.critical + summary.bySeverity.high + summary.bySeverity.medium + summary.bySeverity.low + summary.bySeverity.info;

  // Add apps with blocker count to summary (used in history)
  summary.appsWithBlocker = appsWithBlocker;

  return summary;
}
