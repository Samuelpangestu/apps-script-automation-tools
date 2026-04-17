/**
 * Dashboard.js — Project Distribution Dashboard
 * ═══════════════════════════════════════════════════════════════════════
 * Show team distribution across projects
 * ═══════════════════════════════════════════════════════════════════════
 */

const DASHBOARD_TAB_NAME = 'Dashboard';

/**
 * Create Dashboard tab
 */
function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check dependencies exist
  if (!ss.getSheetByName(CONFIG_TAB_NAME)) {
    throw new Error('Config tab must be created first');
  }

  if (!ss.getSheetByName(TEAM_TAB_NAME)) {
    throw new Error('Team Members tab must be created first');
  }

  // Create new sheet (deletion handled by caller)
  const sheet = ss.insertSheet(DASHBOARD_TAB_NAME);

  // Build dashboard
  let currentRow = 1;

  // Title
  sheet.getRange(currentRow, 1, 1, 10).merge()
    .setValue('📊 QA TEAM PROJECT & SUBMODUL DISTRIBUTION')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 50);
  currentRow += 2;

  // Get data
  const projects = getActiveProjects();
  const moduls = getActiveModul();
  const submoduls = getActiveSubmodul();
  const teamMembers = getAllActiveTeamMembers();

  // Summary section
  sheet.getRange(currentRow, 1).setValue('📈 SUMMARY').setFontWeight('bold').setFontSize(12);
  currentRow++;

  sheet.getRange(currentRow, 1).setValue('Total Active Projects:');
  sheet.getRange(currentRow, 2).setValue(projects.length);
  currentRow++;

  sheet.getRange(currentRow, 1).setValue('Total Active Modul:');
  sheet.getRange(currentRow, 2).setValue(moduls.length);
  currentRow++;

  sheet.getRange(currentRow, 1).setValue('Total Active Submodul:');
  sheet.getRange(currentRow, 2).setValue(submoduls.length);
  currentRow++;

  sheet.getRange(currentRow, 1).setValue('Total Active Team Members:');
  sheet.getRange(currentRow, 2).setValue(teamMembers.length);
  currentRow += 2;

  // Team distribution by role
  sheet.getRange(currentRow, 1).setValue('👥 TEAM DISTRIBUTION BY ROLE').setFontWeight('bold').setFontSize(12);
  currentRow++;

  const roleCount = {
    'QA Team Lead': 0,
    'QA Lead': 0,
    'PIC Project': 0,
    'Quality Engineer': 0
  };

  teamMembers.forEach(member => {
    if (roleCount[member.role] !== undefined) {
      roleCount[member.role]++;
    }
  });

  Object.keys(roleCount).forEach(role => {
    sheet.getRange(currentRow, 1).setValue(role + ':');
    sheet.getRange(currentRow, 2).setValue(roleCount[role]);
    currentRow++;
  });

  currentRow += 2;

  // Submodul distribution section
  sheet.getRange(currentRow, 1, 1, 10).merge()
    .setValue('🎯 SUBMODUL TEAM ASSIGNMENTS (with Ratings)')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Table header
  const tableHeaders = ['Project', 'Modul', 'Submodul', 'Diff.', 'Risk', 'Comp.', 'QA Team Lead', 'QA Lead', 'PIC Project', 'QE'];
  sheet.getRange(currentRow, 1, 1, 10)
    .setValues([tableHeaders])
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  currentRow++;

  const tableStartRow = currentRow;

  // Build submodul distribution table
  submoduls.forEach((submodul, index) => {
    // Find project from modul
    const modul = moduls.find(m => m.name === submodul.modul);
    const projectName = modul ? modul.project : '-';

    // Get team members for this submodul
    const qaTeamLeads = [];
    const qaLeads = [];
    const pics = [];
    const qes = [];

    teamMembers.forEach(member => {
      if (member.submodul.includes(submodul.name)) {
        if (member.role === 'QA Team Lead') {
          qaTeamLeads.push(member.name);
        } else if (member.role === 'QA Lead') {
          qaLeads.push(member.name);
        } else if (member.role === 'PIC Project') {
          pics.push(member.name);
        } else if (member.role === 'Quality Engineer') {
          qes.push(member.name);
        }
      }
    });

    const rowData = [
      projectName,
      submodul.modul,
      submodul.name,
      submodul.difficulty,
      submodul.risk,
      submodul.complexity,
      qaTeamLeads.join(', ') || '-',
      qaLeads.join(', ') || '-',
      pics.join(', ') || '-',
      qes.join(', ') || '-'
    ];

    sheet.getRange(currentRow, 1, 1, 10).setValues([rowData]);

    // Row styling
    const bg = index % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(currentRow, 1, 1, 10)
      .setBackground(bg)
      .setWrap(true)
      .setVerticalAlignment('top');

    // Rating color coding (difficulty/risk/complexity)
    const avgRating = (submodul.difficulty + submodul.risk + submodul.complexity) / 3;
    let ratingBg = '#ffffff';
    if (avgRating <= 3) {
      ratingBg = '#d4edda'; // Low (green)
    } else if (avgRating <= 6) {
      ratingBg = '#fff3cd'; // Medium (yellow)
    } else {
      ratingBg = '#f8d7da'; // High (red)
    }
    sheet.getRange(currentRow, 4).setBackground(ratingBg);
    sheet.getRange(currentRow, 5).setBackground(ratingBg);
    sheet.getRange(currentRow, 6).setBackground(ratingBg);

    currentRow++;
  });

  // Set column widths
  sheet.setColumnWidth(1, 150); // Project
  sheet.setColumnWidth(2, 150); // Modul
  sheet.setColumnWidth(3, 200); // Submodul
  sheet.setColumnWidth(4, 70);  // Difficulty
  sheet.setColumnWidth(5, 70);  // Risk
  sheet.setColumnWidth(6, 70);  // Complexity
  sheet.setColumnWidth(7, 160); // QA Team Lead
  sheet.setColumnWidth(8, 160); // QA Lead
  sheet.setColumnWidth(9, 160); // PIC Project
  sheet.setColumnWidth(10, 160); // Quality Engineer

  // Auto-resize rows for wrapped text
  for (let i = tableStartRow; i < currentRow; i++) {
    sheet.setRowHeight(i, 60);
  }

  currentRow += 2;

  // Legend
  sheet.getRange(currentRow, 1).setValue('📌 LEGEND:').setFontWeight('bold');
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('Difficulty Levels:');
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('🟢 Easy').setBackground('#d4edda');
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('🟡 Medium').setBackground('#fff3cd');
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('🔴 Hard').setBackground('#f8d7da');
  currentRow += 2;

  // Footer
  sheet.getRange(currentRow, 1, 1, 10).merge()
    .setValue('Last updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'))
    .setFontStyle('italic')
    .setFontSize(9)
    .setHorizontalAlignment('right');

  // Flush changes to ensure completion
  SpreadsheetApp.flush();

  Logger.log('✅ Dashboard created');
}

/**
 * Get project statistics
 */
function getProjectStats() {
  const projects = getActiveProjects();
  const teamMembers = getAllActiveTeamMembers();

  const stats = {
    totalProjects: projects.length,
    totalTeamMembers: teamMembers.length,
    byDifficulty: {
      'Easy': 0,
      'Medium': 0,
      'Hard': 0
    },
    byRole: {
      'QA Team Lead': 0,
      'QA Lead': 0,
      'PIC Project': 0,
      'Quality Engineer': 0
    },
    projectDetails: []
  };

  // Count by difficulty
  projects.forEach(project => {
    if (stats.byDifficulty[project.difficulty] !== undefined) {
      stats.byDifficulty[project.difficulty]++;
    }
  });

  // Count by role
  teamMembers.forEach(member => {
    if (stats.byRole[member.role] !== undefined) {
      stats.byRole[member.role]++;
    }
  });

  // Project details
  projects.forEach(project => {
    const teamCount = teamMembers.filter(m => m.projects.includes(project.name)).length;
    stats.projectDetails.push({
      name: project.name,
      difficulty: project.difficulty,
      teamCount: teamCount
    });
  });

  return stats;
}
