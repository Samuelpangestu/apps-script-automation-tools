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
    ['Samuel Pangestu G Tobing', '', '', 'QA Team Lead', 'SIPGN, INAgov, Emeterai, Integrasi Data Omnyx', 'N053', 'samuel.gonggom@inadigital.co.id', '', 'Onboard', '', 'Samuelpangestu', '6281389088362', 'Senior Quality Engineer', 'TRUE', 'TRUE'],
    ['Muhammad Lutfi Ramdani', '', '', 'QA Team Lead', 'Shared Resource', 'N068', 'muhamad.ramdani@inadigital.co.id', '', 'Onboard', '', 'm-lutfi-ramdani', '', 'Senior Quality Engineer', 'TRUE', 'TRUE'],
    ['Irvan Muhandis', '', '', 'PIC QE', 'INAgov, Wahana', '7815', 'irvan.muhandis@inadigital.co.id', '', 'Onboard', 'INAgov', 'Irvan Muhandis', '6285701514915', 'Quality Engineer', 'TRUE', 'FALSE'],
    ['Muhammad Rizky Ferdiansyah', '', '', 'QE', 'INAgov, Wahana', 'K814', 'muhammad.ferdiansyah@inadigital.co.id', '', 'Onboard', '', 'rizkyfrdiansyah', '6289608603137', 'Quality Engineer', 'TRUE', 'FALSE'],
    ['Wafiq Afifah', '', '', 'PIC QE', 'Digidoc 2.0, Peruri Shield, Penjaminan Online, INAgov', '7818', 'wafiq.afifah@inadigital.co.id', '', 'Onboard', '', 'wifa28', '6283161989370', 'Quality Engineer', 'TRUE', 'FALSE'],
    ['Fresma', '', '', 'PIC QE', 'INAgov, Emeterai, Peruri Shield', '7816', 'muhammad.fresma@inadigital.co.id', '', 'Onboard', '', 'fresmaa', '6282143555445', 'Quality Engineer', 'TRUE', 'FALSE'],
    ['Zahwa', '', '', 'QE', 'Peruri ID, COTS', '', 'zahwa.fairana@inadigital.co.id', '', 'Onboard', '', '', '', 'Intern Quality Engineer', 'TRUE', 'FALSE'],
    ['Putri Handayani', '', '', 'Intern', 'Peruri ID, Digidoc 2.0, COTS', '', 'putri.feby@intern-inadigital.id', '', 'Onboard', '', '', '', 'Intern Quality Engineer', 'TRUE', 'FALSE'],
    ['Mohammad Imam Fauzul', '', '', 'PIC QE', 'Peruri ID, SIPGN', 'K813', 'imam.fauzul@inadigital.co.id', '', 'Onboard', 'Peruri ID', 'imamfauzul', '6289507337201', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Muhammad Farhan Alkautsar', '', '', 'PIC QE', 'Peruri ID, SIPGN', 'P227', 'farhan.intern@digitalperuri.id', '', 'Onboard', 'Peruri ID', 'muhammadfarhanalkautsar@gmail.com', '6281275771360', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Daffa SH', '', '', 'PIC QE', 'Digidoc 2.0, Peruri Shield, CA Life Cycle, CMP, Emudra, COTS', '7814', 'daffa.haramaini@inadigital.co.id', '', 'Onboard', 'COTS', 'shidqiiii', '6281294146841', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Pavita Sherintama G', '', '', 'PIC QE', 'Digidoc 2.0, Peruri Shield, CA Life Cycle, CMP, Penjaminan Online', '7817', 'pavita.giantoro@inadigital.co.id', '', 'Onboard', '', 'pavitasherin', '6282245492793', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Dini Indriyani Putri', '', '', 'Lead Project', 'SIPGN', 'AC000009', 'dini.putri@inadigital.co.id', 'dinindriyanip@gmail.com', 'Onboard', '', '', '', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Gita Gloria', '', '', 'QE', 'SIPGN', 'AC000010', 'gita.gloria@inadigital.co.id', 'gitchaglo@gmail.com', 'Onboard', '', '', '', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Ira Elysa Gurning', '', '', 'QE', 'SIPGN', 'AC000017', 'ira.gurning@inadigital.co.id', 'gurningira97@gmail.com', 'Onboard', '', '', '', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Adinda Fitria Utami', '', '', 'QE', 'SIPGN', 'AC000044', 'adinda.utami@inadigital.co.id', 'adinda.utami@peruri.co.id', 'Onboard', '', '', '6281322515815', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Reiza Nurul Huda', '', '', 'QE', 'SIPGN', 'AC000107', 'reiza.huda@inadigital.co.id', '', 'Onboard', '', '', '', 'Quality Engineer', 'FALSE', 'TRUE'],
    ['Syarif Waliyuddin', '', '', 'QE', 'SIPGN', 'AC000060', 'syarif.waliyuddin@inadigital.co.id', '', 'Onboard', '', '', '', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Denta Devina Tiara Maharani', '', '', 'QE', 'SIPGN', 'AC000098', 'denta.maharani@inadigital.co.id', '', 'Onboard', '', '', '', 'Quality Engineer', 'TRUE', 'TRUE'],
    ['Arief Rahman', '', '', 'QE', 'SIPGN', 'AC000185', 'arief.rahman@inadigital.co.id', '', 'Onboard', '', '', '', 'Quality Engineer', 'FALSE', 'FALSE'],
    ['Muhammad Thariq Fathurahman', '', '', 'QE', 'SIPGN', 'AC000184', 'mfathurahman@inadigital.co.id', '', 'Onboard', '', '', '', 'Quality Engineer', 'FALSE', 'FALSE'],
    ['Radiansyah Amir', '04-04-2026', 'Senior QE', '', 'SIPGN', 'AC000189', 'radiansyah.amir@inadigital.co.id', '', 'Onboard', '', '', '', 'Quality Engineer', 'FALSE', 'FALSE'],
    ['Adika Pambudi', '04-16-2026', 'QE', '', 'SIPGN', 'AC000197', 'adika.pambudi@inadigital.co.id', '', 'Onboard', '', '', '', 'Quality Engineer', 'FALSE', 'TRUE'],
    ['Sela Inriani Hutagalung', '', 'Intern', '', '', 'U419', 'sela.inriani@intern-inadigital.id', 'selainriani17@gmail.com', 'Onboard', '', '', '85183376392', '', 'TRUE', 'FALSE'],
    ['Muhammad Indra Fata', '', 'Intern', '', '', 'U432', 'm.indra@intern-inadigital.id', 'indrafataa19@gmail.com', 'Onboard', '', '', '81228747771', '', 'TRUE', 'FALSE'],
    ['Melinda Triandari Nurlita', '', 'Intern', '', '', 'U426', 'melinda.triandari@intern-inadigital.id', 'melindatriandari01@gmail.com', 'Onboard', '', '', '82119394379', '', 'TRUE', 'FALSE'],
    ['Adam Yudhistira Muhtar', '', 'Intern', '', '', 'U417', 'adam.yudhistira@intern-inadigital.id', 'adamyudhistiramuhtar@gmail.com', 'Onboard', '', '', '82112366831', '', 'TRUE', 'FALSE'],
    ['Aryadi Nugroho', '', '', '', '', '', '', '', 'Onboard', '', '', '6281218461860', 'Security Engineer', 'FALSE', 'FALSE'],
    ['Anggit Verdaningrum', '', '', '', '', '', 'anggitverdan@digitalperuri.id', '', 'Onboard', '', '', '', 'UX Research', '', ''],
    ['Evan Gilang Ramadhan', '', '', '', '', '', 'evan.ramadhan@inadigital.co.id', '', 'Onboard', '', '', '', 'UX Research', '', ''],
    ['Gempitaning Fatihah Fajri', '', '', '', '', '', 'gempitaf@digitalperuri.id', '', 'Onboard', '', '', '', 'UX Research', '', ''],
    ['Hamdan Abdul Aziz', '', '', '', '', 'N061', 'hamdan.abdul@inadigital.co.id', 'hamdan.aziz@peruri.co.id', 'Onboard', '', '', '6282116665772', 'Security Engineer', 'TRUE', 'TRUE'],
    ['Agung Suwandaru', '', '', '', '', 'N056', 'agung.suwandaru@inadigital.co.id', 'agung.suwandaru@peruri.co.id', 'Onboard', '', '', '6281316450824', 'Security Engineer', 'TRUE', 'TRUE'],
    ['Gervasius Laba Kuma', '10-13-2025', '', '', '', 'AC000012', 'gervasius.kuma@inadigital.co.id', 'gervasius.kuma@peruri.co.id', 'Onboard', '', '', '628118121210', 'Security Engineer', 'FALSE', 'TRUE'],
    ['Afin Nur Ihksan', '', '', '', '', 'AC000062', 'afin.ikhsan@inadigital.co.id', 'afin.ikhsan@peruri.co.id', 'Onboard', '', '', '6283867839354', 'Security Engineer', 'FALSE', 'TRUE'],
    ['Kholid Abdulah', '11-24-2025', '', '', '', 'AC000092', 'kholid.abdulah@inadigitl.co.id', '', 'Onboard', '', '', '6285353294055', 'Security Engineer', 'FALSE', 'TRUE']
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
