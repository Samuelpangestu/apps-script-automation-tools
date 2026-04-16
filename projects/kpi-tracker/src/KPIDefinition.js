/**
 * KPIDefinition.js — KPI Definition Management
 * ═══════════════════════════════════════════════════════════════════════
 * Maintainable KPI definitions per role
 * Update KPI target, formula, or add new KPIs here
 * ═══════════════════════════════════════════════════════════════════════
 */

const KPI_DEF_TAB_NAME = 'KPI Definition';

/**
 * Setup KPI Definition tab with all KPIs per role
 */
function setupKPIDefinitionTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let kpiDef = ss.getSheetByName(KPI_DEF_TAB_NAME);

  if (!kpiDef) {
    kpiDef = ss.insertSheet(KPI_DEF_TAB_NAME, 1);
  }

  kpiDef.clear();

  // Set column widths
  kpiDef.setColumnWidth(1, 50);   // A: No
  kpiDef.setColumnWidth(2, 250);  // B: Role
  kpiDef.setColumnWidth(3, 300);  // C: KPI Name
  kpiDef.setColumnWidth(4, 100);  // D: Target
  kpiDef.setColumnWidth(5, 400);  // E: Formula
  kpiDef.setColumnWidth(6, 250);  // F: Data Source
  kpiDef.setColumnWidth(7, 120);  // G: Frequency
  kpiDef.setColumnWidth(8, 150);  // H: Scope

  // Header
  kpiDef.getRange('A1:H1').merge()
    .setValue('📊 KPI DEFINITION — QA DEPARTMENT')
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  kpiDef.setRowHeight(1, 40);

  kpiDef.getRange('A2:H2').merge()
    .setValue('Maintainable KPI definitions per role. Update targets, formulas, or add new KPIs here.')
    .setFontSize(10)
    .setFontStyle('italic')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');

  // Table headers
  const headers = [
    ['No', 'Role', 'KPI Name', 'Target', 'Formula / Calculation', 'Data Source', 'Frequency', 'Scope']
  ];

  kpiDef.getRange(4, 1, 1, 8).setValues(headers)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // KPI Data based on user's table
  const kpiData = [
    // QA Team Lead (CoE)
    [1, 'QA Team Lead (CoE)', 'On-time Delivery Rate', '≥ 80%', 'Milestone on-time ÷ Total milestone × 100', 'Jira (due date vs closed date)', 'Bulanan', '-'],
    [2, 'QA Team Lead (CoE)', 'Defect Detection Efficiency (DDE)', '≥ 75%', 'Bug pre-prod ÷ (Bug pre-prod + Bug prod) × 100', 'Jira (environment label)', 'Per sprint', '-'],
    [3, 'QA Team Lead (CoE)', 'CoE Tool & Template Adoption Rate', '≥ 80%', 'Project pakai CoE ÷ Total project aktif × 100', 'Audit checklist spreadsheet', 'Bulanan', '-'],
    [4, 'QA Team Lead (CoE)', 'Team Training Completion Rate', '≥ 85%', 'Anggota selesai training ÷ Total target × 100', 'Log training / absensi', 'Per kuartal', '-'],

    // QA Lead (Project Dedicated)
    [5, 'QA Lead (Project Dedicated)', 'Test Automation Coverage', '≥ 70%', 'TC smoke diotomasi ÷ Total TC smoke eligible × 100', 'Repo GitHub/GitLab + spreadsheet TC', 'Per sprint', 'Agregat semua project'],
    [6, 'QA Lead (Project Dedicated)', 'Defect Escape Rate', '≤ 10%', 'Bug UAT+prod ÷ Total bug × 100', 'Jira (environment tag)', 'Per sprint', 'Agregat semua project'],
    [7, 'QA Lead (Project Dedicated)', '360 Review Score', '≥ 3.5/5', 'Weighted: TL 35% + PM 25% + QE Peer 15% + Self 10% + Opsional 15%', 'Google Form → tab 360 Review', 'Per semester / kontrak', '-'],
    [8, 'QA Lead (Project Dedicated)', 'Test Pass Rate', '≥ 85%', 'TC passed ÷ Total TC executed × 100', 'Spreadsheet TC / test management tool', 'Per sprint', 'Agregat semua project'],

    // PIC Project (QE + Koordinator)
    [9, 'PIC Project (QE + Koordinator)', 'Test Automation Coverage', '≥ 65%', 'TC smoke diotomasi ÷ Total TC smoke eligible × 100', 'Repo GitHub/GitLab + spreadsheet TC', 'Per sprint', 'Per project'],
    [10, 'PIC Project (QE + Koordinator)', 'Defect Escape Rate', '≤ 10%', 'Bug UAT+prod ÷ Total bug × 100', 'Jira (environment tag)', 'Per sprint', 'Per project'],
    [11, 'PIC Project (QE + Koordinator)', '360 Review Score', '≥ 3.5/5', 'Weighted: TL 35% + PM 25% + QE Peer 15% + Self 10% + Opsional 15%', 'Google Form → tab 360 Review', 'Per semester / kontrak', '-'],
    [12, 'PIC Project (QE + Koordinator)', 'Test Pass Rate', '≥ 80%', 'TC passed ÷ Total TC executed × 100', 'Spreadsheet TC / test management tool', 'Per sprint', 'Per project'],

    // Quality Engineer (QE)
    [13, 'Quality Engineer (QE)', 'Test Automation Coverage', '≥ 60%', 'TC smoke diotomasi ÷ Total TC smoke eligible × 100', 'Repo GitHub/GitLab + spreadsheet TC', 'Per sprint', 'Per project'],
    [14, 'Quality Engineer (QE)', 'Defect Escape Rate', '≤ 10%', 'Bug UAT+prod ÷ Total bug × 100', 'Jira (environment tag)', 'Per sprint', 'Per project'],
    [15, 'Quality Engineer (QE)', '360 Review Score', '≥ 3.5/5', 'Weighted: TL 35% + PM 25% + QE Peer 15% + Self 10% + Opsional 15%', 'Google Form → tab 360 Review', 'Per semester / kontrak', '-'],
    [16, 'Quality Engineer (QE)', 'Test Pass Rate', '≥ 80%', 'TC passed ÷ Total TC executed × 100', 'Spreadsheet TC / test management tool', 'Per sprint', 'Per project']
  ];

  const startRow = 5;
  kpiDef.getRange(startRow, 1, kpiData.length, 8).setValues(kpiData);

  // Format data rows
  for (let i = 0; i < kpiData.length; i++) {
    const row = startRow + i;
    const role = kpiData[i][1];

    // Background color per role
    let bg = '#ffffff';
    if (role.includes('Team Lead (CoE)')) bg = '#e3f2fd';
    else if (role.includes('QA Lead (Project')) bg = '#e8f5e9';
    else if (role.includes('PIC Project')) bg = '#fff3e0';
    else if (role.includes('Quality Engineer')) bg = '#fce4ec';

    kpiDef.getRange(row, 1, 1, 8)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('top')
      .setWrap(true);

    // Number column
    kpiDef.getRange(row, 1).setHorizontalAlignment('center');

    // Target column - bold
    kpiDef.getRange(row, 4).setFontWeight('bold');
  }

  // Set row heights for better readability
  kpiDef.setRowHeights(startRow, kpiData.length, 60);

  // Freeze header
  kpiDef.setFrozenRows(4);
  kpiDef.setFrozenColumns(1);

  // Add note
  kpiDef.getRange(startRow + kpiData.length + 2, 1, 1, 8).merge()
    .setValue('💡 TIP: Update targets, formulas, or add new KPIs by editing this table. Changes will automatically reflect in KPI Tracker.')
    .setBackground('#fff9c4')
    .setFontStyle('italic')
    .setWrap(true);

  Logger.log('✅ KPI Definition tab setup complete with ' + kpiData.length + ' KPIs');
  return kpiDef;
}

/**
 * Get KPI definitions for a specific role
 * @param {string} role - Role name
 * @returns {Array} Array of KPI objects
 */
function getKPIsForRole(role) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiDef = ss.getSheetByName(KPI_DEF_TAB_NAME);

  if (!kpiDef) {
    throw new Error('KPI Definition tab not found. Run setupKPIDefinitionTab() first.');
  }

  const lastRow = kpiDef.getLastRow();
  const data = kpiDef.getRange(5, 1, lastRow - 4, 8).getValues();

  const kpis = [];
  for (let i = 0; i < data.length; i++) {
    const rowRole = data[i][1].toString().trim();
    if (rowRole === role) {
      kpis.push({
        no: data[i][0],
        role: rowRole,
        name: data[i][2],
        target: data[i][3],
        formula: data[i][4],
        dataSource: data[i][5],
        frequency: data[i][6],
        scope: data[i][7]
      });
    }
  }

  return kpis;
}

/**
 * Get all unique roles from KPI Definition
 * @returns {Array} Array of role names
 */
function getAllRoles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiDef = ss.getSheetByName(KPI_DEF_TAB_NAME);

  if (!kpiDef) {
    return [];
  }

  const lastRow = kpiDef.getLastRow();
  const data = kpiDef.getRange(5, 2, lastRow - 4, 1).getValues();

  const roles = [];
  data.forEach(row => {
    const role = row[0].toString().trim();
    if (role && roles.indexOf(role) === -1) {
      roles.push(role);
    }
  });

  return roles;
}
