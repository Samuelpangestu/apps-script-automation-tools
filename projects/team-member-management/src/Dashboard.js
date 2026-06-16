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
  sheet.getRange(currentRow, 1, 1, 9).merge()
    .setValue('📊 QA TEAM PROJECT & SUBMODUL DISTRIBUTION')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 50);
  currentRow++;

  // Filter note
  sheet.getRange(currentRow, 1, 1, 9).merge()
    .setValue('ℹ️ Only showing projects/moduls/submoduls with Status = Active | Inactive & On Hold are excluded')
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setFontSize(10)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 25);
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

  // Count unique Submodul Leads
  const uniqueSubmodulLeads = new Set();
  submoduls.forEach(sub => {
    if (sub.qaLead) {
      const leads = sub.qaLead.split(',').map(q => q.trim()).filter(q => q);
      leads.forEach(lead => uniqueSubmodulLeads.add(lead));
    }
  });

  const roleCount = {
    'QA Team Lead': 0,
    'Submodul Lead': uniqueSubmodulLeads.size,
    'PIC Project': 0,
    'Quality Engineer': 0
  };

  teamMembers.forEach(member => {
    if (member.role === 'QA Team Lead') roleCount['QA Team Lead']++;
    if (member.role === 'PIC Project') roleCount['PIC Project']++;
    if (member.role === 'Quality Engineer') roleCount['Quality Engineer']++;
  });

  Object.keys(roleCount).forEach(role => {
    sheet.getRange(currentRow, 1).setValue(role + ':');
    sheet.getRange(currentRow, 2).setValue(roleCount[role]);
    currentRow++;
  });

  currentRow += 2;

  // Submodul distribution section
  sheet.getRange(currentRow, 1, 1, 9).merge()
    .setValue('🎯 SUBMODUL TEAM ASSIGNMENTS')
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Table header
  const tableHeaders = ['Project', 'Modul', 'Submodul', 'PIC QA', 'Category', 'QA Team Lead', 'Submodul Lead', 'PIC Project', 'QE'];
  sheet.getRange(currentRow, 1, 1, 9)
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
    const pics = [];
    const qes = [];

    teamMembers.forEach(member => {
      if (member.submodul.includes(submodul.name)) {
        if (member.role === 'QA Team Lead') {
          qaTeamLeads.push(member.name);
        } else if (member.role === 'PIC Project') {
          pics.push(member.name);
        } else if (member.role === 'Quality Engineer') {
          qes.push(member.name);
        }
      }
    });

    // Calculate category from 5 parameters
    const totalScore = calculateTotalScore(
      submodul.difficulty,
      submodul.risk,
      submodul.complexity,
      submodul.automation,
      submodul.testComplexity
    );
    const category = getDifficultyCategory(totalScore);
    const categoryEmoji = getCategoryEmoji(category);
    const categoryDisplay = categoryEmoji + ' ' + category;

    const rowData = [
      projectName,
      submodul.modul,
      submodul.name,
      submodul.picQA || '-', // PIC QA from Config
      categoryDisplay,
      qaTeamLeads.join(', ') || '-',
      submodul.qaLead || '-', // QA Lead from Config
      pics.join(', ') || '-',
      qes.join(', ') || '-'
    ];

    sheet.getRange(currentRow, 1, 1, 9).setValues([rowData]);

    // Row styling
    const bg = index % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(currentRow, 1, 1, 9)
      .setBackground(bg)
      .setWrap(true)
      .setVerticalAlignment('top');

    // Category column color coding (now at column 5 instead of 4)
    const ratingBg = getCategoryColor(category);
    sheet.getRange(currentRow, 5).setBackground(ratingBg).setFontWeight('bold');

    currentRow++;
  });

  // Set column widths (compact)
  sheet.setColumnWidth(1, 120); // Project
  sheet.setColumnWidth(2, 120); // Modul
  sheet.setColumnWidth(3, 160); // Submodul
  sheet.setColumnWidth(4, 110); // PIC QA
  sheet.setColumnWidth(5, 120); // Category
  sheet.setColumnWidth(6, 130); // QA Team Lead
  sheet.setColumnWidth(7, 130); // Submodul Lead
  sheet.setColumnWidth(8, 130); // PIC Project
  sheet.setColumnWidth(9, 130); // Quality Engineer

  // Auto-resize rows for wrapped text
  for (let i = tableStartRow; i < currentRow; i++) {
    sheet.setRowHeight(i, 60);
  }

  currentRow += 2;

  // Person Matrix Section
  sheet.getRange(currentRow, 1, 1, 9).merge()
    .setValue('👥 PERSON ASSIGNMENT MATRIX')
    .setBackground('#ff6d00')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Matrix header
  const matrixHeaders = ['Name', 'Total Projects', 'Total Submoduls', 'Avg Score', 'Category'];
  sheet.getRange(currentRow, 1, 1, 5)
    .setValues([matrixHeaders])
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  currentRow++;

  const matrixStartRow = currentRow;

  // Build person matrix
  const personMap = new Map(); // Map<name, {projects: Set, submoduls: Array<submodul>}>

  teamMembers.forEach(member => {
    if (!personMap.has(member.name)) {
      personMap.set(member.name, {
        projects: new Set(),
        submoduls: []
      });
    }

    const personData = personMap.get(member.name);
    member.projects.forEach(p => personData.projects.add(p));

    // Find submoduls assigned to this person
    member.submodul.forEach(submodulName => {
      const found = submoduls.find(s => s.name === submodulName);
      if (found) {
        personData.submoduls.push(found);
      }
    });
  });

  // Convert to array for sorting
  const personArray = [];
  personMap.forEach((data, name) => {
    const totalProjects = data.projects.size;
    const totalSubmoduls = data.submoduls.length;

    // Calculate average score
    let totalScore = 0;
    data.submoduls.forEach(sub => {
      totalScore += calculateTotalScore(
        sub.difficulty,
        sub.risk,
        sub.complexity,
        sub.automation,
        sub.testComplexity
      );
    });

    const avgScore = totalSubmoduls > 0 ? Math.round(totalScore / totalSubmoduls) : 0;
    const avgCategory = getDifficultyCategory(avgScore);

    personArray.push({
      name: name,
      totalProjects: totalProjects,
      totalSubmoduls: totalSubmoduls,
      avgScore: avgScore,
      avgCategory: avgCategory
    });
  });

  // Sort by avgScore descending (Very Hard first)
  personArray.sort((a, b) => b.avgScore - a.avgScore);

  // Render sorted matrix
  personArray.forEach((person, matrixIndex) => {
    const avgCategoryEmoji = getCategoryEmoji(person.avgCategory);
    const avgCategoryDisplay = avgCategoryEmoji + ' ' + person.avgCategory;

    const matrixRow = [person.name, person.totalProjects, person.totalSubmoduls, person.avgScore, avgCategoryDisplay];
    sheet.getRange(currentRow, 1, 1, 5).setValues([matrixRow]);

    // Row styling
    const bg = matrixIndex % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(currentRow, 1, 1, 5).setBackground(bg);

    // Category color
    const catBg = getCategoryColor(person.avgCategory);
    sheet.getRange(currentRow, 5).setBackground(catBg).setFontWeight('bold');

    currentRow++;
  });

  // Set matrix column widths (compact)
  sheet.setColumnWidth(1, 140); // Name
  sheet.setColumnWidth(2, 90); // Total Projects
  sheet.setColumnWidth(3, 110); // Total Submoduls
  sheet.setColumnWidth(4, 80); // Avg Score
  sheet.setColumnWidth(5, 120); // Category

  currentRow += 2;

  // Submodul Lead Matrix Section
  sheet.getRange(currentRow, 1, 1, 9).merge()
    .setValue('👨‍💼 SUBMODUL LEAD ASSIGNMENT MATRIX')
    .setBackground('#9c27b0')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 35);
  currentRow++;

  // Submodul Lead Matrix header
  const qaLeadHeaders = ['Submodul Lead', 'Total Projects', 'Total Submoduls', 'Submoduls', 'Avg Score', 'Category', 'Team Members'];
  sheet.getRange(currentRow, 1, 1, 7)
    .setValues([qaLeadHeaders])
    .setBackground('#666666')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  currentRow++;

  const qaLeadMatrixStartRow = currentRow;

  // Build QA Lead matrix from active submoduls
  const qaLeadMap = new Map(); // Map<qaLead, {projects: Set, submoduls: Array<submodul>, members: Set}>

  submoduls.forEach(sub => {
    // QA Lead already available from getActiveSubmodul() (from Config column K)
    if (!sub.qaLead) return;

    // Parse comma-separated QA Leads
    const qaLeads = sub.qaLead.split(',').map(q => q.trim()).filter(q => q);

    qaLeads.forEach(qaLead => {
      if (!qaLeadMap.has(qaLead)) {
        qaLeadMap.set(qaLead, {
          projects: new Set(),
          submoduls: [],
          members: new Set()
        });
      }

      const leadData = qaLeadMap.get(qaLead);
      leadData.projects.add(sub.project);
      leadData.submoduls.push(sub);

      // Find team members under this QA Lead
      teamMembers.forEach(member => {
        if (member.qaLead && member.qaLead.includes(qaLead)) {
          leadData.members.add(member.name);
        }
      });
    });
  });

  // Convert to array for sorting
  const qaLeadArray = [];
  qaLeadMap.forEach((data, leadName) => {
    const totalProjects = data.projects.size;
    const totalSubmoduls = data.submoduls.length;

    // Calculate average score
    let totalScore = 0;
    data.submoduls.forEach(sub => {
      totalScore += calculateTotalScore(
        sub.difficulty,
        sub.risk,
        sub.complexity,
        sub.automation,
        sub.testComplexity
      );
    });

    const avgScore = totalSubmoduls > 0 ? Math.round(totalScore / totalSubmoduls) : 0;
    const avgCategory = getDifficultyCategory(avgScore);

    // Extract submodul names with numbering
    const submodulNames = data.submoduls.map((sub, idx) => `${idx + 1}. ${sub.name}`).join('\n');

    qaLeadArray.push({
      name: leadName,
      totalProjects: totalProjects,
      totalSubmoduls: totalSubmoduls,
      submoduls: submodulNames,
      avgScore: avgScore,
      avgCategory: avgCategory,
      members: Array.from(data.members).join(', ')
    });
  });

  // Sort by avgScore descending
  qaLeadArray.sort((a, b) => b.avgScore - a.avgScore);

  // Render Submodul Lead matrix
  qaLeadArray.forEach((lead, leadIndex) => {
    const avgCategoryEmoji = getCategoryEmoji(lead.avgCategory);
    const avgCategoryDisplay = avgCategoryEmoji + ' ' + lead.avgCategory;

    const leadRow = [lead.name, lead.totalProjects, lead.totalSubmoduls, lead.submoduls, lead.avgScore, avgCategoryDisplay, lead.members];
    sheet.getRange(currentRow, 1, 1, 7).setValues([leadRow]);

    // Row styling with wrap text and top alignment
    const bg = leadIndex % 2 === 0 ? '#ffffff' : '#f8f9fa';
    sheet.getRange(currentRow, 1, 1, 7)
      .setBackground(bg)
      .setWrap(true)
      .setVerticalAlignment('top');

    // Category color (now column 6 instead of 5)
    const catBg = getCategoryColor(lead.avgCategory);
    sheet.getRange(currentRow, 6).setBackground(catBg).setFontWeight('bold');

    currentRow++;
  });

  // Set Submodul Lead matrix column widths (compact)
  sheet.setColumnWidth(1, 140); // Submodul Lead
  sheet.setColumnWidth(2, 90); // Total Projects
  sheet.setColumnWidth(3, 110); // Total Submoduls
  sheet.setColumnWidth(4, 240); // Submoduls
  sheet.setColumnWidth(5, 80); // Avg Score
  sheet.setColumnWidth(6, 120); // Category
  sheet.setColumnWidth(7, 200); // Team Members

  currentRow += 2;

  // Legend - Simplified
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('🟢 Very Easy (5-14)  |  🔵 Easy (15-24)  |  🟡 Medium (25-34)  |  🟠 Hard (35-44)  |  🔴 Very Hard (45-50)')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setBackground('#f3f3f3');
  currentRow += 2;

  // Footer
  sheet.getRange(currentRow, 1, 1, 9).merge()
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
