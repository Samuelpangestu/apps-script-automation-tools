/**
 * DataInjection.js — Inject Team Member Data
 * ═══════════════════════════════════════════════════════════════════════
 * Script to inject predefined team member data into Team Members tab
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Inject team member data
 */
function injectTeamMemberData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TEAM_TAB_NAME);

  if (!sheet) {
    throw new Error('Team Members tab not found. Please create it first.');
  }

  Logger.log('📥 Injecting team member data...');

  // Raw data structure: [Name, Join, Title, Lead/PIC, Project, NP, Email, Email 2, Status Hiring, Automation, Github Personal, HP, Role, VPN ABC, VPN Huwawei]
  const rawData = [
    ['Team Member 1', '2023-01-15', 'Senior Quality Engineer', 'QA Team Lead', 'Project A', 'EMP-001', 'member1@example.com', '', 'Onboard', '', 'member-1', '+62 800-0000-0001', 'Senior Quality Engineer', TRUE, TRUE],
    ['Team Member 2', '2023-02-01', 'Quality Engineer', 'Team Member 1', 'Project B', 'EMP-002', 'member2@example.com', '', 'Onboard', '', 'member-2', '+62 800-0000-0002', 'Quality Engineer', TRUE, TRUE],
    ['Team Member 3', '2023-03-10', 'Quality Engineer', 'Team Member 1', 'Project A, Project C', 'EMP-003', 'member3@example.com', '', 'Onboard', '', 'member-3', '+62 800-0000-0003', 'Quality Engineer', FALSE, FALSE],
  ];

  // Transform to Team Members format (19 columns)
  // Team format: [No, NP, Name, Email, Email 2, HP, Join Date, Title, Role, Lead/PIC, Project, Modul, Submodul, Status, Status Hiring, Automation, Github Personal, VPN ABC, VPN Huwawei]
  const teamData = [];
  let rowNum = 1;

  rawData.forEach(row => {
    const name = row[0] ? row[0].trim() : '';

    // Skip empty rows
    if (!name) return;

    const joinDate = row[1] ? row[1].trim() : '';
    const title = row[2] ? row[2].trim() : '';
    const leadPic = row[3] ? row[3].trim() : '';
    const project = row[4] ? row[4].trim() : '';
    const np = row[5] ? row[5].trim() : '';
    const email = row[6] ? row[6].trim() : '';
    const email2 = row[7] ? row[7].trim() : '';
    const statusHiring = row[8] ? row[8].trim() : '';
    const automation = row[9] ? row[9].trim() : '';
    const github = row[10] ? row[10].trim() : '';
    const hp = row[11] ? row[11].trim() : '';
    const role = row[12] ? row[12].trim() : '';
    const vpnAbc = row[13] ? row[13].trim() : '';
    const vpnHuwawei = row[14] ? row[14].trim() : '';

    // Determine status
    let status = '';
    if (statusHiring === 'Onboard') {
      status = 'Active';
    } else if (statusHiring === 'Tidak Ada Kabar' || statusHiring === 'Digispark') {
      status = 'Inactive';
    }

    teamData.push([
      rowNum,           // No
      np,               // NP
      name,             // Name
      email,            // Email
      email2,           // Email 2
      hp,               // HP
      joinDate,         // Join Date
      title,            // Title
      role,             // Role
      leadPic,          // Lead/PIC
      project,          // Project
      '',               // Modul (empty - to be filled manually)
      '',               // Submodul (empty - to be filled manually)
      status,           // Status
      statusHiring,     // Status Hiring
      automation,       // Automation
      github,           // Github Personal
      vpnAbc,           // VPN ABC
      vpnHuwawei        // VPN Huwawei
    ]);

    rowNum++;
  });

  // Clear existing data (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > TEAM_HEADER_ROW) {
    sheet.getRange(TEAM_DATA_START_ROW, 1, lastRow - TEAM_HEADER_ROW, TEAM_TOTAL_COLUMNS).clearContent();
  }

  // Write data
  if (teamData.length > 0) {
    sheet.getRange(TEAM_DATA_START_ROW, 1, teamData.length, TEAM_TOTAL_COLUMNS).setValues(teamData);

    // Apply formatting
    applyTeamMemberFormatting();

    Logger.log('✅ Injected ' + teamData.length + ' team members');

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Successfully injected ' + teamData.length + ' team members',
      'Data Injection Complete',
      5
    );

    return {
      success: true,
      count: teamData.length,
      message: 'Injected ' + teamData.length + ' team members'
    };
  }

  return { success: false, message: 'No data to inject' };
}

/**
 * Menu function to inject data
 */
function menuInjectData() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Inject Team Member Data',
    'This will replace all existing data in the Team Members tab with predefined data.\n\n' +
    'Current data will be cleared. Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    injectTeamMemberData();
  } catch (error) {
    ui.alert('Error', 'Failed to inject data: ' + error.message, ui.ButtonSet.OK);
    Logger.log('Error injecting data: ' + error.message);
  }
}
