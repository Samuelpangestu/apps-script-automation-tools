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

  // Check dependencies - just create without alerts to avoid conflicts
  if (!ss.getSheetByName(CONFIG_TAB_NAME)) {
    createConfigTab();
  }

  if (!ss.getSheetByName(TEAM_TAB_NAME)) {
    createTeamMemberTab();
  }

  // Delete existing if any
  let existingSheet = ss.getSheetByName(DASHBOARD_TAB_NAME);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }

  // Create new sheet (without position to avoid conflicts)
  const sheet = ss.insertSheet(DASHBOARD_TAB_NAME);

  // Build dashboard
  let currentRow = 1;

  // Title
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('📊 QA TEAM PROJECT DISTRIBUTION DASHBOARD')
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
  const teamMembers = getAllActiveTeamMembers();

  // Summary section
  sheet.getRange(currentRow, 1).setValue('📈 SUMMARY').setFontWeight('bold').setFontSize(12);
  currentRow++;

  sheet.getRange(currentRow, 1).setValue('Total Active Projects:');
  sheet.getRange(currentRow, 2).setValue(projects.length);
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

  // Project distribution section
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('🎯 PROJECT TEAM ASSIGNMENTS')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Table header
  const tableHeaders = ['Project', 'Difficulty', 'QA Team Lead', 'QA Lead', 'PIC Project', 'Quality Engineer'];
  sheet.getRange(currentRow, 1, 1, 6)
    .setValues([tableHeaders])
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  currentRow++;

  const tableStartRow = currentRow;

  // Build project distribution table
  projects.forEach((project, index) => {
    // Get team members for this project
    const qaTeamLeads = [];
    const qaLeads = [];
    const pics = [];
    const qes = [];

    teamMembers.forEach(member => {
      if (member.projects.includes(project.name)) {
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
      project.name,
      project.difficulty,
      qaTeamLeads.join(', ') || '-',
      qaLeads.join(', ') || '-',
      pics.join(', ') || '-',
      qes.join(', ') || '-'
    ];

    sheet.getRange(currentRow, 1, 1, 6).setValues([rowData]);

    // Row styling
    const bg = index % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(currentRow, 1, 1, 6)
      .setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID)
      .setWrap(true)
      .setVerticalAlignment('top');

    // Difficulty color coding
    let difficultyBg = '#ffffff';
    if (project.difficulty === 'Easy') {
      difficultyBg = '#d4edda';
    } else if (project.difficulty === 'Medium') {
      difficultyBg = '#fff3cd';
    } else if (project.difficulty === 'Hard') {
      difficultyBg = '#f8d7da';
    }
    sheet.getRange(currentRow, 2).setBackground(difficultyBg);

    currentRow++;
  });

  // Set column widths
  sheet.setColumnWidth(1, 200); // Project
  sheet.setColumnWidth(2, 100); // Difficulty
  sheet.setColumnWidth(3, 180); // QA Team Lead
  sheet.setColumnWidth(4, 180); // QA Lead
  sheet.setColumnWidth(5, 180); // PIC Project
  sheet.setColumnWidth(6, 180); // Quality Engineer

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
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('Last updated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'))
    .setFontStyle('italic')
    .setFontSize(9)
    .setHorizontalAlignment('right');

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
