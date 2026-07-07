/**
 * Users Tab - User Role Management
 *
 * Tab structure:
 * - Email | Role | Active
 * - Used by Next.js platform for RBAC
 */

/**
 * Build Users tab from scratch
 */
function buildUsers(ss) {
  const ws = ss.insertSheet('Users');
  ws.setTabColor('#4A148C');

  // Title row
  ws.getRange(1, 1).setValue('USER ROLES - Manage user access for QA Platform')
    .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold');

  // Info row
  ws.getRange(2, 1, 1, 3).merge()
    .setValue('Kelola role user untuk akses ke QA Platform. Role: admin (full access), qa-engineer (edit + view), viewer (view only)')
    .setBackground('#F3E5F5').setFontColor('#4A148C').setFontStyle('italic')
    .setFontSize(9).setWrap(true);

  // Headers
  ws.getRange(3, 1, 1, 3).setValues([['Email', 'Role', 'Active']])
    .setBackground('#6A1B9A').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Column widths
  ws.setColumnWidth(1, 250); // Email
  ws.setColumnWidth(2, 120); // Role
  ws.setColumnWidth(3, 80);  // Active

  // Default users - sync with Next.js roles.ts
  ws.getRange(4, 1, 9, 3).setValues([
    ['samuel.gonggom@inadigital.co.id', 'admin', 'TRUE'],
    ['muhamad.ramdani@inadigital.co.id', 'admin', 'TRUE'],
    ['departemen.qa@inadigital.co.id', 'admin', 'TRUE'],
    ['tantri.saraswati@digitalperuri.id', 'admin', 'TRUE'],
    ['agung.suwandaru@inadigital.co.id', 'admin', 'TRUE'],
    ['anggitverdan@digitalperuri.id', 'admin', 'TRUE'],
    ['evan.ramadhan@inadigital.co.id', 'admin', 'TRUE'],
    ['gempitaf@digitalperuri.id', 'admin', 'TRUE'],
    ['hamdan.abdul@inadigital.co.id', 'qa-security', 'TRUE']
  ]);

  // Data validation for Role column (B4:B1000)
  const roleValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['admin', 'qa-engineer', 'qa-research', 'qa-security', 'viewer'], true)
    .setAllowInvalid(false)
    .build();
  ws.getRange(4, 2, 997, 1).setDataValidation(roleValidation);

  // Data validation for Active column (C4:C1000)
  const boolValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  ws.getRange(4, 3, 997, 1).setDataValidation(boolValidation);

  // Format Active column as centered
  ws.getRange(4, 3, 997, 1).setHorizontalAlignment('center').setFontWeight('bold');

  // Freeze header rows
  ws.setFrozenRows(3);

  // OPTIONAL: Hide tab (admin bisa unhide via right-click)
  // ws.hideSheet();

  // OPTIONAL: Protect sheet (admin only can edit)
  // const protection = ws.protect().setDescription('User Roles - Admin Only');
  // protection.setWarningOnly(true); // Warning mode (easier for admin)
}

/**
 * Rebuild Users tab (delete old + create new)
 */
function rebuildUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet = ss.getSheetByName('Users');

  if (oldSheet) {
    ss.deleteSheet(oldSheet);
  }

  buildUsers(ss);

  SpreadsheetApp.getUi().alert('✅ Users tab rebuilt!\n\nJangan lupa update menu di MasterDashboard.js jika diperlukan.');
}
