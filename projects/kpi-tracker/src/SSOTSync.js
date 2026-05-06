/**
 * SSOTSync.js — Team Member Synchronization
 * ═══════════════════════════════════════════════════════════════════════
 * Sync team members from QA Team Management spreadsheet
 * Source: 1PKZTLAhjcBAoBNcfhhOGdy7JUixEqvLbOwK5YAueSa4SdT5QzH1OHh_U
 * Tab: Team Members (19 columns)
 * ═══════════════════════════════════════════════════════════════════════
 */

// Team Management Spreadsheet Configuration
const SSOT_SPREADSHEET_ID = '11BrTvbh6AWY9cmVIfDx30vfUD6f3vOi0_4D8vHC7Nyw';
const SSOT_TAB_NAME = 'Team Members';
const SSOT_DATA_START_ROW = 2; // Data starts at row 2 (row 1 is header)

// Column mapping from Team Members tab (0-indexed, 19 columns)
// Structure: No, NP, Name, Email, Email 2, HP, Join Date, Title, Role, Lead/PIC, Project, Modul, Submodul, Status, Status Hiring, Automation, Github Personal, VPN ABC, VPN Huwawei
const SSOT_COLUMNS = {
  NO: 0,             // Column A: No
  NP: 1,             // Column B: NP
  NAME: 2,           // Column C: Name
  EMAIL: 3,          // Column D: Email
  EMAIL_2: 4,        // Column E: Email 2
  HP: 5,             // Column F: HP
  JOIN: 6,           // Column G: Join Date
  TITLE: 7,          // Column H: Title
  ROLE: 8,           // Column I: Role
  LEAD_PIC: 9,       // Column J: Lead/PIC
  PROJECT: 10,       // Column K: Project
  MODUL: 11,         // Column L: Modul
  SUBMODUL: 12,      // Column M: Submodul
  STATUS: 13,        // Column N: Status
  STATUS_HIRING: 14, // Column O: Status Hiring
  AUTOMATION: 15,    // Column P: Automation
  GITHUB: 16,        // Column Q: Github Personal
  VPN_ABC: 17,       // Column R: VPN ABC
  VPN_HUWAWEI: 18    // Column S: VPN Huwawei
};

// Role mapping from SSOT to KPI Tracker roles
const ROLE_MAPPING = {
  'Senior Quality Engineer': 'QA Team Lead (CoE)',
  'QA Team Lead': 'QA Team Lead (CoE)',
  'Quality Engineer': 'Quality Engineer (QE)',
  'Intern Quality Engineer': 'Quality Engineer (QE)',
  'PIC QE': 'PIC Project (QE + Koordinator)',
  'Lead Project': 'QA Lead (Project Dedicated)'
};

// Status mapping from Team Members to KPI Tracker
const STATUS_MAPPING = {
  'Active': 'Aktif',
  'Inactive': 'Non-Aktif',
  'On Leave': 'Non-Aktif',
  'Onboard': 'Aktif',
  'Tidak Ada Kabar': 'Non-Aktif',
  'Digispark': 'Non-Aktif'
};

/**
 * Fetch team members from SSOT spreadsheet
 * @returns {Array} Array of team member objects
 */
function fetchFromSSOT() {
  try {
    const ssotSS = SpreadsheetApp.openById(SSOT_SPREADSHEET_ID);
    const ssotSheet = ssotSS.getSheetByName(SSOT_TAB_NAME);

    if (!ssotSheet) {
      throw new Error('SSOT tab "' + SSOT_TAB_NAME + '" not found in spreadsheet');
    }

    const lastRow = ssotSheet.getLastRow();
    if (lastRow < SSOT_DATA_START_ROW) {
      Logger.log('⚠️ No data found in SSOT');
      return [];
    }

    const dataRange = ssotSheet.getRange(
      SSOT_DATA_START_ROW,
      1,
      lastRow - SSOT_DATA_START_ROW + 1,
      19 // Read all 19 columns (A-S)
    );

    const data = dataRange.getValues();
    const members = [];

    data.forEach((row, index) => {
      const name = row[SSOT_COLUMNS.NAME] ? row[SSOT_COLUMNS.NAME].toString().trim() : '';
      const ssotRole = row[SSOT_COLUMNS.ROLE] ? row[SSOT_COLUMNS.ROLE].toString().trim() : '';
      const status = row[SSOT_COLUMNS.STATUS] ? row[SSOT_COLUMNS.STATUS].toString().trim() : '';
      const email = row[SSOT_COLUMNS.EMAIL] ? row[SSOT_COLUMNS.EMAIL].toString().trim() : '';

      // Skip empty rows
      if (!name || name === '') {
        return;
      }

      // Filter only Active members
      if (status !== 'Active') {
        return;
      }

      // Filter only QE/QA roles
      const mappedRole = ROLE_MAPPING[ssotRole];
      if (!mappedRole) {
        // Skip non-QA roles (Security Engineer, UX Research, etc.)
        return;
      }

      // Map status - use Status column (Active/Inactive) with fallback to Status Hiring
      let kpiStatus = STATUS_MAPPING[status];
      if (!kpiStatus) {
        const statusHiring = row[SSOT_COLUMNS.STATUS_HIRING] ? row[SSOT_COLUMNS.STATUS_HIRING].toString().trim() : '';
        kpiStatus = STATUS_MAPPING[statusHiring] || 'Non-Aktif';
      }

      // Extract join date
      const joinDate = row[SSOT_COLUMNS.JOIN];
      let startDate = '';
      if (joinDate instanceof Date) {
        startDate = joinDate;
      } else if (joinDate && joinDate.toString().trim()) {
        // Try to parse date string
        try {
          const parsed = new Date(joinDate.toString());
          if (!isNaN(parsed.getTime())) {
            startDate = parsed;
          }
        } catch (e) {
          // Ignore parse errors
        }
      }

      members.push({
        name: name,
        role: mappedRole,
        email: email,
        status: kpiStatus,
        startDate: startDate,
        ssotRow: index + SSOT_DATA_START_ROW,
        originalRole: ssotRole,
        project: row[SSOT_COLUMNS.PROJECT] ? row[SSOT_COLUMNS.PROJECT].toString().trim() : '',
        hp: row[SSOT_COLUMNS.HP] ? row[SSOT_COLUMNS.HP].toString().trim() : '',
        np: row[SSOT_COLUMNS.NP] ? row[SSOT_COLUMNS.NP].toString().trim() : '',
        modul: row[SSOT_COLUMNS.MODUL] ? row[SSOT_COLUMNS.MODUL].toString().trim() : '',
        submodul: row[SSOT_COLUMNS.SUBMODUL] ? row[SSOT_COLUMNS.SUBMODUL].toString().trim() : ''
      });
    });

    Logger.log('✅ Fetched ' + members.length + ' QA team members from SSOT');
    return members;

  } catch (e) {
    Logger.log('❌ Error fetching from SSOT: ' + e.message);
    throw new Error('Failed to fetch from SSOT: ' + e.message);
  }
}

/**
 * Sync team members from SSOT to Config tab
 * This will overwrite Config tab with SSOT data
 */
function syncFromSSOT() {
  const ui = SpreadsheetApp.getUi();

  // Confirm sync
  const response = ui.alert(
    'Sync from QA Team Management',
    'This will sync team members from QA Team Management spreadsheet.\n\n' +
    'Source: Team Members tab (Active only)\n' +
    'Target: Config tab\n\n' +
    'Current Config data will be overwritten.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // Fetch from Team Management
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Fetching data from QA Team Management...', 'Syncing', 3);

    const members = fetchFromSSOT();

    if (members.length === 0) {
      ui.alert('No Active QA team members found in Team Members tab.\n\nPlease check:\n1. Team Members tab exists\n2. Members have Status = Active\n3. Members have QA/QE roles');
      return;
    }

    // Setup Config tab if not exists
    let config = ss.getSheetByName(CONFIG_TAB_NAME);

    if (!config) {
      ui.alert('Config tab not found. Creating it...');
      setupConfigTab();  // From TeamConfig.js
      config = ss.getSheetByName(CONFIG_TAB_NAME);
    }

    // Clear existing data (keep header)
    const lastRow = config.getLastRow();
    if (lastRow >= TEAM_START_ROW) {
      config.getRange(TEAM_START_ROW, 1, lastRow - TEAM_START_ROW + 1, 7).clearContent();
    }

    // Write data
    ui.alert('Writing ' + members.length + ' members to Config tab...');

    members.forEach((member, idx) => {
      const row = TEAM_START_ROW + idx;

      config.getRange(row, 1).setValue(idx + 1); // No
      config.getRange(row, 2).setValue(member.name); // Name
      config.getRange(row, 3).setValue(member.role); // Role
      config.getRange(row, 4).setValue(member.email); // Email
      config.getRange(row, 5).setValue(member.status); // Status

      if (member.startDate) {
        config.getRange(row, 6).setValue(member.startDate); // Start Date
      }
      config.getRange(row, 7).setValue(''); // End Date

      // Format
      const bg = idx % 2 === 0 ? '#ffffff' : '#f8f9fa';
      config.getRange(row, 1, 1, 7)
        .setBackground(bg)
        .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

      // Center No column
      config.getRange(row, 1).setHorizontalAlignment('center');
    });

    // Apply date format
    config.getRange(TEAM_START_ROW, 6, members.length, 2).setNumberFormat('yyyy-mm-dd');

    // Log sync info
    const syncTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // Add sync info at bottom
    const infoRow = TEAM_START_ROW + members.length + 2;
    config.getRange(infoRow, 1, 1, 7).merge()
      .setValue('Last synced from SSOT: ' + syncTime + ' | ' + members.length + ' members')
      .setFontStyle('italic')
      .setFontColor('#666666')
      .setBackground('#f1f3f4');

    Logger.log('✅ Sync complete: ' + members.length + ' members');

    ui.alert(
      'Sync Complete! ✅',
      'Successfully synced ' + members.length + ' Active QA team members from Team Management.\n\n' +
      'Role distribution:\n' +
      getRoleDistribution(members) + '\n\n' +
      'Time: ' + syncTime,
      ui.ButtonSet.OK
    );

    // Refresh dashboard if exists
    const dashboard = ss.getSheetByName(DASHBOARD_TAB_NAME);
    if (dashboard) {
      createDashboard();
      Logger.log('✅ Dashboard refreshed');
    }

  } catch (e) {
    Logger.log('❌ Sync failed: ' + e.message);
    ui.alert('Sync Failed', 'Error: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Get role distribution summary
 * @param {Array} members - Array of team members
 * @returns {string} Distribution summary
 */
function getRoleDistribution(members) {
  const dist = {};
  members.forEach(m => {
    dist[m.role] = (dist[m.role] || 0) + 1;
  });

  const lines = [];
  Object.keys(dist).sort().forEach(role => {
    lines.push('• ' + role + ': ' + dist[role]);
  });

  return lines.join('\n');
}

/**
 * Show Team Management info
 */
function showSSOTInfo() {
  const ui = SpreadsheetApp.getUi();

  try {
    const members = fetchFromSSOT();

    ui.alert(
      'QA Team Management Info',
      'Spreadsheet: QA Team Management\n' +
      'Spreadsheet ID: ' + SSOT_SPREADSHEET_ID + '\n' +
      'Tab: ' + SSOT_TAB_NAME + '\n\n' +
      'Active QA Team Members: ' + members.length + '\n\n' +
      'Role Distribution:\n' +
      getRoleDistribution(members) + '\n\n' +
      'Status:\n' +
      getStatusDistribution(members),
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('Error', 'Failed to fetch Team Management info:\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Get status distribution
 * @param {Array} members - Array of team members
 * @returns {string} Distribution summary
 */
function getStatusDistribution(members) {
  const dist = {};
  members.forEach(m => {
    dist[m.status] = (dist[m.status] || 0) + 1;
  });

  const lines = [];
  Object.keys(dist).sort().forEach(status => {
    lines.push('• ' + status + ': ' + dist[status]);
  });

  return lines.join('\n');
}

/**
 * Test Team Management connection
 */
function testSSOTConnection() {
  const ui = SpreadsheetApp.getUi();

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Testing connection...', 'Test', 2);

    const ssotSS = SpreadsheetApp.openById(SSOT_SPREADSHEET_ID);
    const ssotSheet = ssotSS.getSheetByName(SSOT_TAB_NAME);

    if (!ssotSheet) {
      throw new Error('Tab "' + SSOT_TAB_NAME + '" not found');
    }

    const lastRow = ssotSheet.getLastRow();

    ui.alert(
      'Connection Test Successful! ✅',
      'QA Team Management Spreadsheet: ' + ssotSS.getName() + '\n' +
      'Tab: ' + SSOT_TAB_NAME + '\n' +
      'Total Rows: ' + lastRow + '\n\n' +
      'Connection is working correctly.',
      ui.ButtonSet.OK
    );

    Logger.log('✅ Team Management connection test passed');

  } catch (e) {
    Logger.log('❌ Team Management connection test failed: ' + e.message);
    ui.alert(
      'Connection Test Failed ❌',
      'Error: ' + e.message + '\n\n' +
      'Please check:\n' +
      '1. Spreadsheet ID is correct\n' +
      '2. You have access to the QA Team Management spreadsheet\n' +
      '3. Tab name "' + SSOT_TAB_NAME + '" exists',
      ui.ButtonSet.OK
    );
  }
}
