// ===================================================================
//  QA TEST MANAGEMENT -- Standard Template  v39
//  Tabs: TC_Master | TC_Execution | API_Master | API_Execution | Summary | BugReport | PerfTest | Appendix
//  Run: createQASheet()
//
//  v39 Changes:
//  - Open Blocker formula: ganti 9x COUNTIFS verbose → SUMPRODUCT bersih
//    Objective: hitung bug Open/InProgress/Reopen x Priority Critical/High/Medium
//  - Smoke TOTAL: pisah variabel agar lebih jelas
//  - createQASheet alert: update info
// ===================================================================


// -- Conditional formatting untuk RUN STATUS row ------------------
function addRunStatusCF(ws, statusRow, startCol, endCol) {
  try {
    if (!ws || typeof ws.getRange !== 'function') return;
    if (!statusRow || !startCol || !endCol || endCol < startCol) return;
    const range = ws.getRange(statusRow, startCol, 1, endCol - startCol + 1);
    const rules = ws.getConditionalFormatRules();
    [{val:'FAILED',bg:'#FFCDD2',fg:'#B71C1C'},{val:'IN PROGRESS',bg:'#E3F2FD',fg:'#1565C0'},
      {val:'PASSED',bg:'#C8E6C9',fg:'#1B5E20'},{val:'BLOCKED',bg:'#FFE0B2',fg:'#E65100'},
      {val:'--',bg:'#F5F5F5',fg:'#9E9E9E'}].forEach(s=>rules.push(
        SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(s.val).setBackground(s.bg).setFontColor(s.fg).setBold(true)
            .setRanges([range]).build()
    ));
    ws.setConditionalFormatRules(rules);
  } catch(e) { Logger.log('addRunStatusCF skipped: ' + e.message); }
}


// -- QA PERURI BRANDING ------------------------------------------
const PERURI = {
  primary:   '#0D47A1',   // deep navy
  mid:       '#1565C0',   // main blue
  accent:    '#1976D2',   // medium blue
  light:     '#E3F2FD',   // pale blue bg
  text:      '#FFFFFF',
  copyright: '(c) QA INA Digital  |  Template ini merupakan properti QA Team INA Digital  |  Dilarang digunakan/disebarluaskan tanpa izin  |  departemen.qa@inadigital.co.id',
};

// Safe sheet creator: delete existing then re-insert
function safeSheet(ss, name) {
  try {
    const existing = ss.getSheetByName(name);
    if (existing) { ss.deleteSheet(existing); SpreadsheetApp.flush(); Utilities.sleep(300); }
  } catch(e) {}
  // Retry loop - Apps Script insertSheet can fail if called too quickly
  for (let i = 0; i < 3; i++) {
    try {
      const sh = ss.insertSheet(name);
      if (sh && sh.getRange) { return sh; }
    } catch(e) { Logger.log('safeSheet attempt ' + (i+1) + ' failed: ' + e.message); }
    SpreadsheetApp.flush();
    Utilities.sleep(500);
  }
  throw new Error('safeSheet: could not create sheet ' + name + ' after 3 attempts');
}

function addPeruriFooter(ws, lastRow, totalCols) {
  const r = lastRow + 2;
  ws.getRange(r,1,1,totalCols).merge();
  ws.getRange(r,1)
      .setValue(PERURI.copyright)
      .setBackground(PERURI.primary)
      .setFontColor(PERURI.text)
      .setFontFamily('Arial').setFontSize(8)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(false);
  ws.setRowHeight(r, 20);

  // Add thin top border in accent blue
  ws.getRange(r,1,1,totalCols)
      .setBorder(true, false, false, false, false, false, PERURI.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
function inputBorder(range) {
  range.setBorder(true,true,true,true,false,false,'#1976D2', SpreadsheetApp.BorderStyle.SOLID);
  return range;
}

function createQASheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('QA Test Management -- Template');
  ['TC_Master','TC_Execution','API_Master','API_Execution','Summary','Dashboard','_Dashboard','PerfTest','BugReport','Appendix']
      .forEach(name => { const s=ss.getSheetByName(name); if(s) try{ss.deleteSheet(s);}catch(e){} });
  SpreadsheetApp.flush(); Utilities.sleep(500);
  // Create sheets - flush + sleep between each to avoid insertSheet returning undefined
  Logger.log('Creating TC_Master...');
  createTCMaster(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating API_Master...');
  createAPIMaster(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating TC_Execution...');
  createTCExecution(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating API_Execution...');
  createAPIExecution(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating BugReport...');
  createBugReport(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating PerfTest...');
  createPerfTest(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating Summary...');
  createSummary(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating Appendix...');
  createAppendix(ss); SpreadsheetApp.flush();
  const s1=ss.getSheetByName('Sheet1');
  if(s1&&ss.getSheets().length>1) try{ss.deleteSheet(s1);}catch(e){}
  // Reorder tabs: move each to correct position
  SpreadsheetApp.flush();
  try {
    const order=['Summary','Appendix','BugReport','TC_Execution','TC_Master','API_Execution','API_Master','PerfTest'];
    order.slice().reverse().forEach(name=>{
      const sh=ss.getSheetByName(name);
      if(sh){ ss.setActiveSheet(sh); ss.moveActiveSheet(1); }
    });
  } catch(e) { Logger.log('Tab reorder skipped: '+e.message); }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
      '[OK]  QA Test Management Template v39 berhasil dibuat.\n\n'+
      'TC_Master     -- input test case Web / Mobile\n'+
      'TC_Execution  -- hasil eksekusi per tanggal run\n'+
      'API_Master    -- input test case API\n'+
      'API_Execution -- hasil eksekusi API\n'+
      'Summary       -- info sesi, coverage & run history\n'+
      'BugReport     -- log bug Web, Mobile, dan API\n'+
      'PerfTest      -- rekam & evaluasi hasil performance test\n'+
      'Appendix      -- definisi & panduan\n\n'+
      'IN PROGRESS = ada PASSED tapi masih ada TODO yang belum dieksekusi.'
  );
}

// -- COLORS ----------------------------------------------------------
const C = {
  dark:'#1E2A3A',    blue:'#1A73E8',   teal:'#00897B',
  orange:'#E65100',  purple:'#6A1B9A', red:'#C62828',
  green:'#2E7D32',   grey:'#546E7A',   indigo:'#283593',
  white:'#FFFFFF',   lgrey:'#F5F7FA',  mgrey:'#ECEFF1',
  passed:'#C8E6C9',  failed:'#FFCDD2', blocked:'#FFE0B2',
  section:'#E3F2FD', yellow:'#FFF9C4',
};

// -- HELPERS ---------------------------------------------------------
function bd(r) {
  return r.setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
}
function hdr(r, bg, fg, sz) {
  fg = fg || '#FFFFFF'; sz = sz || 9;
  return bd(r).setBackground(bg).setFontColor(fg).setFontWeight('bold')
      .setFontSize(sz).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
}
function cell_style(c, bg) {
  return c.setBackground(bg||C.white).setFontFamily('Arial').setFontSize(9)
      .setVerticalAlignment('middle');
}
function dv(list) {
  return SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true).setAllowInvalid(false).build();
}
function prioColor(c, v) {
  const m = {Critical:'#B71C1C',High:'#E53935',Medium:'#FB8C00',Low:'#43A047',Lowest:'#90A4AE'};
  if (m[v]) c.setFontColor(m[v]).setFontWeight('bold');
}
function autoColor(c, v) {
  const m = {
    'Automated':           ['#E8F5E9','#1B5E20'],
    'Manual':              ['#E3F2FD','#0D47A1'],
    'To Do':               ['#FFF8E1','#F57F17'],
    'Cannot be Automated': ['#FAFAFA','#546E7A'],
  };
  if (m[v]) c.setBackground(m[v][0]).setFontColor(m[v][1]).setFontWeight('bold').setHorizontalAlignment('center');
}
function statusColor(c, v) {
  const m = {PASSED:C.passed, FAILED:C.failed, BLOCKED:C.blocked, SKIP:'#EEEEEE'};
  if (m[v]) c.setBackground(m[v]);
}

// -- TITLE BAR --------------------------------------------------------
function colLetter(n) {
  let s=''; while(n>0){ n--; s=String.fromCharCode(65+(n%26))+s; n=Math.floor(n/26); } return s;
}
function tlFormula(prioRef) {
  return '=IF('+prioRef+'="","",IF(OR('+prioRef+'="Critical",'+prioRef+'="High",'+prioRef+'="Medium"),"Smoke","Regression"))';
}
function statusCF_(ws, range) {
  const rules = ws.getConditionalFormatRules();
  [{val:'PASSED',bg:'#C8E6C9',fg:'#1B5E20'},{val:'IN PROGRESS',bg:'#E3F2FD',fg:'#1565C0'},
    {val:'FAILED',bg:'#FFCDD2',fg:'#B71C1C'},{val:'BLOCKED',bg:'#FFE0B2',fg:'#E65100'},
    {val:'TODO',bg:'#F5F5F5',fg:'#616161'}].forEach(s=>rules.push(
      SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(s.val).setBackground(s.bg).setFontColor(s.fg).setBold(true)
          .setRanges([range]).build()
  ));
  ws.setConditionalFormatRules(rules);
}
function titleBar(ws, cols, text, bg, sz) {
  ws.getRange(1,1,1,cols).merge();
  hdr(ws.getRange(1,1,1,cols), bg||C.dark, '#FFFFFF', sz||12).setValue(text);
  ws.setRowHeight(1, 32);
}

// -- HEADER ROW -------------------------------------------------------
function headerRow(ws, row, headers, bg) {
  headers.forEach((h, i) => {
    hdr(ws.getRange(row, i+1), bg||C.dark).setValue(h).setWrap(true);
  });
  ws.setRowHeight(row, 42);
}

// ===================================================================
//  TAB 1 -- TC_Master (Web / Mobile)
//  Input utama test case. Tidak ada kolom eksekusi.
// ===================================================================
function createTCMaster(ss) {
  const ws = ss.getSheetByName('TC_Master') || safeSheet(ss,'TC_Master');
  ws.clear(); ws.setTabColor('#1565C0');
  const COLS=14, DS=3, MR=1000;
  ws.getRange(1,1,1,COLS).merge();
  hdr(ws.getRange(1,1,1,COLS),'#0D47A1','#FFFFFF',10)
      .setValue('TC_MASTER  .  Web / Mobile  .  QA PERURI');
  ws.setRowHeight(1,28);
  ['No','SubModul','TC_ID','Feature','Priority','Platform','Test Type','Automated','Version',
    'Role (RBAC)','Scenario','Steps / Gherkin','Expected Result','[AUTO] Test Level']
      .forEach((h,i)=>hdr(ws.getRange(2,i+1),'#0D47A1').setValue(h).setWrap(true));
  ws.setRowHeight(2,38);
  [36,110,100,130,80,80,90,140,65,100,200,260,200,90].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  // Column header notes
  ws.getRange(2,1).setNote('Row number. Auto-filled when data is entered.');
  ws.getRange(2,2).setNote(
      'SubModule — 3rd level in QA hierarchy:\n'+
      '  Project > Module > SubModule > Feature\n'+
      '\n'+
      'SubModule = smallest standalone unit (one app or one domain).\n'+
      '\n'+
      'Layered project (e.g. SIPGN):\n'+
      '  1.1 = Module 1, SubModule 1 (Nutritionist App)\n'+
      '  1.2 = Module 1, SubModule 2 (Courier App)\n'+
      '\n'+
      'Flat project (e.g. INAGOV):\n'+
      '  Direct name: Talenta, e-Office, SIMPEG\n'+
      '\n'+
      'Must be IDENTICAL in TC_Master and API_Master\n'+
      'for Dashboard coverage to merge correctly.'
  );
  ws.getRange(2,3).setNote(
      'TC_ID — Format: [SubModule].[3-digit]  e.g. 1.1.001 / PO.001\n'+
      'Must be UNIQUE. Do not change once Execution results exist.\n'+
      '\n'+
      'Use initials if SubModule name is long:\n'+
      '  Portal       → PO → PO.001\n'+
      '  BackOffice   → BO → BO.001\n'+
      '\n'+
      'Initials: 2-3 uppercase letters, unique per project,\n'+
      'consistent across TC_Master, API_Master, and Execution.'
  );
  ws.getRange(2,4).setNote(
      'Feature — Specific feature or page name.\n'+
      'Examples: Login Page, Checkout Flow, User Management\n'+
      'Used for grouping Coverage per Feature in Summary.'
  );
  ws.getRange(2,5).setNote(
      'Priority & Impact:\n'+
      '\n'+
      'CRITICAL  → Must PASS before release. FAIL = release BLOCKED.\n'+
      'HIGH      → Must PASS in same sprint. FAIL = needs PM approval.\n'+
      'MEDIUM    → Potential blocker. Fix before UAT.\n'+
      'LOW       → Non-blocker. Fix in next sprint.\n'+
      'LOWEST    → Nice to have. Optional.\n'+
      '\n'+
      'Auto Test Level:\n'+
      'Critical / High / Medium → Smoke Test\n'+
      'Low / Lowest             → Regression Test'
  );
  ws.getRange(2,6).setNote('Platform: Web / Mobile / Web & Mobile');
  ws.getRange(2,7).setNote(
      'Test Type:\n'+
      '  Positive   = happy path (valid data, expected flow)\n'+
      '  Negative   = error case (invalid input, rejected action)\n'+
      '  Edge Case  = boundary condition'
  );
  ws.getRange(2,8).setNote(
      'Automation Status:\n'+
      '  Automated           = script exists and runs\n'+
      '  To Do               = planned, not yet done\n'+
      '  Manual              = decided to stay manual\n'+
      '  Cannot be Automated = technically not possible'
  );
  ws.getRange(2,9).setNote('Version: App version when TC was created. e.g. v1.0, v2.3');
  ws.getRange(2,10).setNote(
      'Role (RBAC) — The user role executing this scenario.\n'+
      '\n'+
      'Examples: Admin, Super Admin, User, Viewer, Operator, Supervisor, Guest\n'+
      '\n'+
      'Used to:\n'+
      '  • Verify test coverage per role\n'+
      '  • Confirm access control (RBAC) is correct\n'+
      '  • Ensure 403 Forbidden for unauthorized roles'
  );
  ws.getRange(2,11).setNote(
      'SCENARIO NAMING STANDARD\n'+
      '\n'+
      'Happy Path : [Role] Successfully [Verb] [Object]\n'+
      'Negative   : [Role] Failed to [Verb] [Object] with [Condition]\n'+
      '\n'+
      'Rules:\n'+
      '  • Role, Object → Title Case  (Nutritionist, Meal Plan)\n'+
      '  • Verb → active  (Create, Pick Up, Confirm, Submit)\n'+
      '  • Do not use: success/succeed, do, perform, process\n'+
      '\n'+
      'Standard examples:\n'+
      '  Nutritionist Successfully Creates Meal Plan\n'+
      '  Admin Failed to Delete User with Invalid ID\n'+
      '\n'+
      '───────────────────────────────────────\n'+
      'SCENARIO OUTLINE (write here in this column)\n'+
      'Use when same steps + same outcome type, different data.\n'+
      'Rule: all Examples = Positive only OR Negative only.\n'+
      '\n'+
      'Negative example:\n'+
      'User Failed to Log In with <invalid_credential>\n'+
      'Examples:\n'+
      '- Invalid password\n'+
      '- Empty password\n'+
      '- Expired session\n'+
      '\n'+
      'Positive example:\n'+
      'Admin Successfully Creates User with Role <role>\n'+
      'Examples:\n'+
      '- Viewer\n'+
      '- Operator\n'+
      '- Supervisor'
  );
  ws.getRange(2,12).setNote(
      '[REQUIRED] Steps in Gherkin format:\n'+
      '\n'+
      '  Given : Pre-condition / initial state\n'+
      '          e.g. Given user is on the Login page\n'+
      '  When  : Action performed by the actor\n'+
      '          e.g. When user submits the form\n'+
      '  And   : Additional action if needed\n'+
      '\n'+
      'DO NOT write Then here — Then goes in Expected Result.\n'+
      '\n'+
      'For Scenario Outline, use <angle_brackets>:\n'+
      '  When user submits login with "<invalid_credential>"'
  );
  ws.getRange(2,13).setNote(
      '[REQUIRED] Expected Result in Gherkin Then format:\n'+
      '\n'+
      '  Then : Outcome / state change after action completes\n'+
      '\n'+
      'Be specific — name the UI element, message, or status.\n'+
      '\n'+
      'Positive example:\n'+
      '  Then dashboard is displayed, username shown in header\n'+
      'Negative example:\n'+
      '  Then login is rejected with message "Incorrect password"\n'+
      '\n'+
      'For Scenario Outline: write the common outcome.\n'+
      'Row detail goes in the Examples column.'
  );
  ws.getRange(2,14).setNote('[AUTO — DO NOT EDIT] Test Level auto-calculated from Priority:\nCritical/High/Medium = Smoke  |  Low/Lowest = Regression');

  const data=[
    [1,'1.1','1.1.001','Informasi Program','High','Web','Positive','Automated','v1.0',
      'Viewer','User dapat melihat halaman informasi program',
      'Given user berada di halaman utama\nWhen klik menu Informasi Program',
      'Then halaman tampil, semua konten tersedia','',''],
    [2,'1.1','1.1.002','Informasi Program','Medium','Web','Negative','Manual','v1.0',
      'Viewer','Halaman error saat konten tidak tersedia',
      'Given konten belum tersedia\nWhen user buka halaman',
      'Then pesan error informatif tampil, halaman tidak crash',''],
    [3,'2.1','2.1.001','Perencanaan Pengiriman','Critical','Web','Positive','Automated','v1.1',
      'Admin','Admin membuat rencana pengiriman baru',
      'Given admin di halaman Perencanaan\nWhen isi form dan klik Simpan',
      'Then data tersimpan dan muncul di daftar',''],
    [4,'2.2','2.2.001','Pelacakan Pengiriman','High','Mobile','Positive','Automated','v1.1',
      'Kurir','Kurir melihat rute pengiriman aktif',
      'Given kurir login di mobile\nWhen buka menu Pengiriman',
      'Then rute tampil di peta dengan marker yang tepat',''],
    [5,'7.3','7.3.001','Manajemen Pengguna','Critical','Web','Positive','Automated','v1.0',
      'Admin','Admin menambahkan user baru',
      'Given admin di halaman Manajemen Pengguna\nWhen isi form dan klik Tambah',
      'Then user baru muncul di daftar',''],
    [6,'7.3','7.3.002','Manajemen Pengguna','Low','Web','Negative','Manual','v1.0',
      'Admin','Filter dengan input tidak valid tidak crash\nExamples:\n- Karakter tidak valid <>#$\n- Input kosong',
      'Given admin di halaman filter\nWhen user submits filter with "<invalid_input>"',
      'Then validation message appears, page does not crash',
      ''],
  ];
  data.forEach((row,i)=>{
    const r=DS+i, bg=i%2===0?'#F8F9FA':'#FFFFFF';
    row.forEach((val,c)=>{
      const cell=ws.getRange(r,c+1);
      bd(cell_style(cell,bg)).setValue(val||'');
      cell.setHorizontalAlignment(c===0?'center':'left');
      if(c===4) prioColor(cell,val);
      if(c===7) autoColor(cell,val);
      if(c===9) cell.setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#1565C0');
      if(c===11||c===12) cell.setWrap(true);
      if(c===13) cell.setHorizontalAlignment('center').setFontWeight('bold');
    });
    bd(cell_style(ws.getRange(r,14),bg)).setFormula(tlFormula('E'+r))
        .setFontWeight('bold').setHorizontalAlignment('center');
    ws.setRowHeight(r,58);
  });
  for(let r=DS+data.length;r<=DS+MR;r++) ws.getRange(r,14).setFormula(tlFormula('E'+r));

  const dvEnd=DS+MR;
  // Master sheets: no input borders (clean look)
  ws.getRange('E'+DS+':E'+dvEnd).setDataValidation(dv(['Critical','High','Medium','Low','Lowest']));
  ws.getRange('F'+DS+':F'+dvEnd).setDataValidation(dv(['Web','Mobile','Web & Mobile']));
  ws.getRange('G'+DS+':G'+dvEnd).setDataValidation(dv(['Positive','Negative','Edge Case']));
  ws.getRange('H'+DS+':H'+dvEnd).setDataValidation(dv(['Automated','Manual','To Do','Cannot be Automated']));
  // Automated column CF
  (function(){
    const autoRange = ws.getRange('H'+DS+':H'+(DS+MR));
    const rules = ws.getConditionalFormatRules();
    [{v:'Automated',      bg:'#C8E6C9', fg:'#1B5E20'},
      {v:'Manual',         bg:'#E3F2FD', fg:'#1565C0'},
      {v:'To Do',          bg:'#FFF9C4', fg:'#F57F17'},
      {v:'Cannot be Automated', bg:'#F5F5F5', fg:'#9E9E9E'},
    ].forEach(s => rules.push(
        SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(s.v).setBackground(s.bg).setFontColor(s.fg).setBold(true)
            .setRanges([autoRange]).build()
    ));
    ws.setConditionalFormatRules(rules);
  })();
  ws.getRange('J'+DS+':J'+dvEnd).setDataValidation(dv(['Admin','User','Viewer','Operator','Supervisor','Guest','Super Admin']));
  // TestLevel col protection
  // TestLevel col N: base styling only, CF handles Smoke/Regression colors
  ws.getRange('N'+DS+':N'+dvEnd).setHorizontalAlignment('center').setFontWeight('bold');

  const cf=ws.getConditionalFormatRules();
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Smoke')
      .setBackground('#FFF8F0').setFontColor('#BF360C').setBold(true)
      .setRanges([ws.getRange('N'+DS+':N'+dvEnd)]).build());
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Regression')
      .setBackground('#F1F8E9').setFontColor('#33691E').setBold(true)
      .setRanges([ws.getRange('N'+DS+':N'+dvEnd)]).build());
  [{val:'Critical',fg:'#C62828'},{val:'High',fg:'#D84315'},{val:'Medium',fg:'#E65100'},
    {val:'Low',fg:'#2E7D32'},{val:'Lowest',fg:'#78909C'}].forEach(p=>
      cf.push(SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(p.val).setFontColor(p.fg).setBold(true)
          .setRanges([ws.getRange('E'+DS+':E'+dvEnd)]).build())
  );
  ws.setConditionalFormatRules(cf);
  addPeruriFooter(ws, DS+data.length+3, COLS);
  ws.getRange(DS+data.length+2,1,1,COLS).merge()
      .setValue('Test Level auto-calculated: Critical/High/Medium → Smoke  .  Low/Lowest → Regression. Do not edit column N.')
      .setFontColor('#78909C').setFontStyle('italic').setFontSize(8).setFontFamily('Arial')
      .setBackground('#E3F2FD').setWrap(true);
}

// ===================================================================
//  TAB 2 -- TC_Execution (Web / Mobile)
//  Kolom identitas sync otomatis dari TC_Master.
//  Kolom eksekusi ditambah manual ke kanan.
//  LATEST STATUS otomatis deteksi kolom terisi paling kanan.
// ===================================================================
function createTCExecution(ss) {
  const ws = safeSheet(ss,'TC_Execution');
  ws.clear(); ws.setTabColor('#1565C0');

  // DS=9: Row 1=group label, 2=headers, 3=RUN STATUS, 4-7=counts, 8=sep, 9+=data
  // Col layout: A=TC_ID B=SubModul C=Feature D=Priority E=Platform F=Scenario G=TestLevel
  //             H+=staging runs   Z=LATEST STATUS  AA=Screenshot
  const DS=9, MR=500, STAG=8, STAG_N=3, STATUS_Z=26, SHOT_COL=27, MAX_RUNS=10;

  // -- Row 1: minimal header -- no merges beyond what's needed --
  hdr(ws.getRange(1,1),'#F0F4F8','#90A4AE',8).setValue('');
  hdr(ws.getRange(1,2),'#F0F4F8','#90A4AE',8).setValue('');
  hdr(ws.getRange(1,3),'#F0F4F8','#90A4AE',8).setValue('');
  hdr(ws.getRange(1,4),'#F0F4F8','#90A4AE',8).setValue('');
  hdr(ws.getRange(1,5),'#F0F4F8','#90A4AE',8).setValue('');
  hdr(ws.getRange(1,6),'#F0F4F8','#90A4AE',8).setValue('');
  hdr(ws.getRange(1,7),'#F0F4F8','#90A4AE',8).setValue('TEST LEVEL').setWrap(true);
  // Staging cols instruction -- only label first 3, no merge so user can add freely
  hdr(ws.getRange(1,STAG),'#0D47A1').setValue('<- Tambah kolom tanggal baru ke kanan');
  ws.setRowHeight(1,22);

  // -- Row 2: column headers --
  ['TC_ID','SubModul','Feature','Priority','Platform','Scenario','[AUTO]\nTest Level']
      .forEach((h,i)=>hdr(ws.getRange(2,i+1),'#0D47A1').setValue(h).setWrap(true));
  ['2025-01-20','2025-02-10','2025-02-21']
      .forEach((h,i)=>hdr(ws.getRange(2,STAG+i),'#0D47A1').setValue(h).setWrap(true));
  hdr(ws.getRange(2,STATUS_Z),'#455A64','#FFFFFF',8).setValue('[AUTO]\nLatest Status').setWrap(true);
  // Dynamic screenshot cols: one per run date, labeled with the date
  for(let ri=0; ri<MAX_RUNS; ri++){
    const sc = SHOT_COL+ri;
    // Header: formula links to the corresponding date in row 2
    hdr(ws.getRange(2,sc),'#1565C0','#FFFFFF',8)
        .setFormula('=IFERROR(IF(INDIRECT(ADDRESS(2,'+(STAG+ri)+'))="","Screenshot "+'+(ri+1)+',TEXT(INDIRECT(ADDRESS(2,'+(STAG+ri)+')),"yyyy-mm-dd")&"\nShot"),"Screenshot '+(ri+1)+'")')
        .setWrap(true);
    ws.setColumnWidth(sc, 140);
  }
  ws.setRowHeight(2,38);
  // Column notes
  ws.getRange(2,1).setNote('[AUTO] TC_ID sync dari TC_Master. Jangan edit.');
  ws.getRange(2,2).setNote('[AUTO] SubModul sync dari TC_Master.');
  ws.getRange(2,3).setNote('[AUTO] Feature sync dari TC_Master.');
  ws.getRange(2,4).setNote('[AUTO] Priority sync dari TC_Master.');
  ws.getRange(2,5).setNote('[AUTO] Platform sync dari TC_Master.');
  ws.getRange(2,6).setNote('[AUTO] Skenario sync dari TC_Master.');
  ws.getRange(2,7).setNote('[AUTO] Test Level sync dari TC_Master. Smoke = wajib ditest di setiap run. Regression = full cycle.');
  ws.getRange(2,STATUS_Z).setNote('[AUTO] Status terkini dihitung otomatis dari semua run.\nFAILED > BLOCKED > IN PROGRESS > PASSED > TODO\nIN PROGRESS = ada PASSED di run sebelumnya tapi masih ada TODO.');
  // Screenshot col notes
  for(let ri=0;ri<MAX_RUNS;ri++){
    ws.getRange(2,SHOT_COL+ri).setNote('[INPUT] Screenshot/evidence untuk run ke-'+(ri+1)+'.\nPaste link Google Drive, Jira attachment, atau URL gambar.');
  }
  for(let i=0;i<STAG_N;i++){
    ws.getRange(2,STAG+i).setNote('[INPUT] Isi TODO/PASSED/FAILED/BLOCKED setelah eksekusi pada tanggal ini.\nGunakan dropdown per cell.');
  }

  // -- Rows 3-7: summary labels -- each col independent (NO merge), label in col G --
  [{row:3,label:'> RUN STATUS', bg:'#37474F',fg:'#FFFFFF'},
    {row:4,label:'  PASSED',     bg:'#F1F8E9',fg:'#2E7D32'},
    {row:5,label:'  FAILED',     bg:'#FBE9E7',fg:'#BF360C'},
    {row:6,label:'  BLOCKED',    bg:'#FFF3E0',fg:'#E65100'},
    {row:7,label:'  TODO',       bg:'#F5F5F5',fg:'#546E7A'},
  ].forEach(s=>{
    // Label in cols A?F merged, col G = TestLevel header (no label needed in summary rows)
    ws.getRange(s.row,1,1,6).merge();
    bd(ws.getRange(s.row,1)).setValue(s.label)
        .setBackground(s.bg).setFontColor(s.fg).setFontWeight('bold')
        .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('right').setVerticalAlignment('middle');
    bd(ws.getRange(s.row,7)).setBackground(s.bg); // col G neutral fill
    ws.setRowHeight(s.row,20);
  });

  // Formulas per staging col -- set for initial 3 cols, user copies right for new runs
  for(let i=0;i<STAG_N;i++){
    const col=STAG+i, L=colLetter(col), rng=L+'$'+DS+':'+L+'$'+(DS+MR);
    bd(ws.getRange(3,col)).setFormula(
        '=IF(COUNTA('+rng+')=0,"--",IF(COUNTIF('+rng+',"FAILED")>0,"FAILED",'+
        'IF(COUNTIF('+rng+',"BLOCKED")>0,"BLOCKED",'+
        'IF(COUNTIF('+rng+',"TODO")=COUNTA('+rng+'),"TODO",'+
        'IF(COUNTIF('+rng+',"TODO")>0,"IN PROGRESS","PASSED")))))'
    ).setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    [['PASSED',4],['FAILED',5],['BLOCKED',6],['TODO',7]].forEach(([st,row])=>
        bd(ws.getRange(row,col)).setFormula('=IF(COUNTA('+rng+')=0,"",COUNTIF('+rng+',"'+st+'"))')
            .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center')
    );
  }
  ws.getRange(4,STAG,1,STAG_N).setBackground('#F1F8E9');
  ws.getRange(5,STAG,1,STAG_N).setBackground('#FBE9E7');
  ws.getRange(6,STAG,1,STAG_N).setBackground('#FFF3E0');
  ws.getRange(7,STAG,1,STAG_N).setBackground('#F5F5F5');
  addRunStatusCF(ws,3,STAG,STAG+STAG_N+10);

  // -- Row 8: separator -- single cell, no merge --
  bd(ws.getRange(8,1)).setValue('v Kolom A-G = [AUTO sync TC_Master] . Kolom H dst = [INPUT status per run] . Kolom Z = [AUTO status terkini] . Kolom AA = [INPUT evidence link]')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic').setFontSize(8).setFontFamily('Arial');
  ws.setRowHeight(8,14);

  // -- Row 9+: sync from TC_Master --
  // TC_Master: C=TC_ID B=SubModul D=Feature E=Priority F=Platform J=Scenario M=TestLevel
  ws.getRange(DS,1).setFormula('=ARRAYFORMULA(IF(TC_Master!C3:C'+(MR+2)+'<>"",TC_Master!C3:C'+(MR+2)+',""))');
  ws.getRange(DS,2).setFormula('=ARRAYFORMULA(IF(TC_Master!B3:B'+(MR+2)+'<>"",TC_Master!B3:B'+(MR+2)+',""))');
  ws.getRange(DS,3).setFormula('=ARRAYFORMULA(IF(TC_Master!D3:D'+(MR+2)+'<>"",TC_Master!D3:D'+(MR+2)+',""))');
  ws.getRange(DS,4).setFormula('=ARRAYFORMULA(IF(TC_Master!E3:E'+(MR+2)+'<>"",TC_Master!E3:E'+(MR+2)+',""))');
  ws.getRange(DS,5).setFormula('=ARRAYFORMULA(IF(TC_Master!F3:F'+(MR+2)+'<>"",TC_Master!F3:F'+(MR+2)+',""))');
  ws.getRange(DS,6).setFormula('=ARRAYFORMULA(IF(TC_Master!K3:K'+(MR+2)+'<>"",TC_Master!K3:K'+(MR+2)+',""))');
  ws.getRange(DS,7).setFormula('=ARRAYFORMULA(IF(TC_Master!N3:N'+(MR+2)+'<>"",TC_Master!N3:N'+(MR+2)+',""))');
  ws.getRange(DS,1,MR,7).setBackground('#E8F0FE').setFontColor('#37474F')
      .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle').setWrap(true);

  // CF Priority
  const cf=ws.getConditionalFormatRules();
  [{val:'Critical',bg:'#FFEBEE',fg:'#C62828'},{val:'High',bg:'#FFF3E0',fg:'#D84315'},
    {val:'Medium',bg:'#FFF8E1',fg:'#E65100'},{val:'Low',bg:'#F1F8E9',fg:'#2E7D32'},
    {val:'Lowest',bg:'#ECEFF1',fg:'#78909C'}].forEach(p=>
      cf.push(SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(p.val).setBackground(p.bg).setFontColor(p.fg).setBold(true)
          .setRanges([ws.getRange(DS,4,MR,1)]).build())
  );
  // CF TestLevel
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Smoke')
      .setBackground('#FFF8F0').setFontColor('#BF360C').setBold(true).setRanges([ws.getRange(DS,7,MR,1)]).build());
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Regression')
      .setBackground('#F1F8E9').setFontColor('#33691E').setBold(true).setRanges([ws.getRange(DS,7,MR,1)]).build());
  ws.setConditionalFormatRules(cf);

  // Sample staging data
  [['PASSED','PASSED','PASSED'],['PASSED','PASSED','FAILED'],
    ['FAILED','PASSED','BLOCKED'],['TODO','TODO','TODO'],
    ['PASSED','PASSED','PASSED'],['TODO','TODO','TODO']].forEach((row,i)=>{
    const r=DS+i;
    row.forEach((val,c)=>{
      bd(cell_style(ws.getRange(r,STAG+c),i%2===0?'#F8F9FA':'#FFFFFF'))
          .setValue(val).setHorizontalAlignment('center');
    });
    ws.setRowHeight(r,34);
  });

  // Staging validation & CF -- cols H(8) to Y(25) = 18 cols
  const STAG_COLS = STATUS_Z - STAG; // 18
  statusCF_(ws, ws.getRange(DS, STAG, MR, STAG_COLS));
  ws.getRange(DS, STAG, MR, STAG_COLS).setDataValidation(dv(['TODO','PASSED','FAILED','BLOCKED']));
  inputBorder(ws.getRange(DS, STAG, MR, STAG_COLS));
  for(let ri=0;ri<MAX_RUNS;ri++) inputBorder(ws.getRange(DS, SHOT_COL+ri, MR, 1));
  // Blue border on staging cols = INPUT area
  inputBorder(ws.getRange(DS, STAG, MR, STAG_COLS));
  // Blue border on screenshot cols
  for(let ri=0;ri<MAX_RUNS;ri++) inputBorder(ws.getRange(DS, SHOT_COL+ri, MR, 1));

  // -- LATEST STATUS col Z --
  // Uses ARRAYFORMULA for efficiency -- single formula in Z9
  // Logic: FAILED > BLOCKED > IN PROGRESS (PASSED+TODO) > PASSED > TODO
  for(let r=DS;r<DS+MR;r++){
    ws.getRange(r,STATUS_Z).setFormula(
        '=IF(A'+r+'="","",IF(COUNTIF(H'+r+':Y'+r+',"FAILED")>0,"FAILED",'+
        'IF(COUNTIF(H'+r+':Y'+r+',"BLOCKED")>0,"BLOCKED",'+
        'IF(COUNTA(H'+r+':Y'+r+')=0,"TODO",'+
        'IF(AND(COUNTIF(H'+r+':Y'+r+',"PASSED")>0,COUNTIF(H'+r+':Y'+r+',"TODO")>0),"IN PROGRESS",'+
        'IF(COUNTIF(H'+r+':Y'+r+',"PASSED")>0,"PASSED","TODO"))))))')
    ;
  }
  ws.getRange(DS,STATUS_Z,MR,1)
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  statusCF_(ws, ws.getRange(DS, STATUS_Z, MR, 1));

  // Screenshot col AA -- just set style, no formula
  // Screenshot cols styling
  for(let ri=0;ri<MAX_RUNS;ri++){
    ws.getRange(DS,SHOT_COL+ri,MR,1)
        .setBackground('#F0F4FF').setFontColor('#1A73E8')
        .setFontFamily('Arial').setFontSize(8).setVerticalAlignment('middle')
        .setWrap(false).setHorizontalAlignment('left');
  }

  // Column widths -- tight, leave Z and AA only
  [100,110,130,80,80,180,75].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  for(let i=0;i<STAG_N;i++) ws.setColumnWidth(STAG+i,100);
  ws.setColumnWidth(STATUS_Z,100);
  // screenshot col widths set in header loop above

  // Freeze cols A-G so identitas stay visible while scrolling right
  ws.setFrozenColumns(7);
  ws.setFrozenRows(2);

  ws.getRange(1, STAG+STAG_N).setValue('<- Insert kolom di sini untuk run baru. Isi tanggal di baris 2, copy formula baris 3?7 dari kolom sebelah.')
      .setFontColor('#90A4AE').setFontStyle('italic').setFontSize(8).setFontFamily('Arial');
}

// ===================================================================
//  TAB 3 -- API_Master
// ===================================================================
function createAPIMaster(ss) {
  const ws=safeSheet(ss,'API_Master');
  ws.clear(); ws.setTabColor('#1565C0');
  const COLS=15, DS=3, MR=1000;
  ws.getRange(1,1,1,COLS).merge();
  hdr(ws.getRange(1,1,1,COLS),'#0D47A1','#FFFFFF',10).setValue('API_MASTER  .  QA PERURI  .  Test Level auto-calculated from Priority.');
  ws.setRowHeight(1,28);
  ['No','SubModul','TC_ID','Feature','Method','Endpoint URL','Priority','Auth','Test Type',
    'Automated','Version','Role (RBAC)','Scenario','Expected Result','[AUTO] Test Level']
      .forEach((h,i)=>hdr(ws.getRange(2,i+1),'#0D47A1').setValue(h).setWrap(true));
  ws.setRowHeight(2,38);
  [36,110,110,130,70,220,80,110,90,140,65,200,90,220,110].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  // Column notes for API_Master
  ws.getRange(2,1).setNote('Row number. Auto-filled when data is entered.');
  ws.getRange(2,2).setNote(
      'SubModule — must be IDENTICAL to TC_Master.\n'+
      '  Project > Module > SubModule > Feature\n'+
      '\n'+
      'Layered: 1.1 / 1.2 / 2.1\n'+
      'Flat   : Talenta / e-Office / SIMPEG\n'+
      '\n'+
      'Consistent naming = correct coverage merge in Dashboard.'
  );
  ws.getRange(2,3).setNote(
      'TC_ID API — Numbering format:\n'+
      '\n'+
      '  API.[SVC].[FEAT].[000]\n'+
      '\n'+
      '  [SVC]  = service/domain code, max 4 uppercase letters\n'+
      '           Examples: AUTH, USER, ORD, PAY, INV\n'+
      '  [FEAT] = endpoint/resource code, max 4 uppercase letters\n'+
      '           Examples: LOG, LIST, CRT, UPD, DEL\n'+
      '  [000]  = 3-digit sequence starting from 001\n'+
      '\n'+
      'Full examples:\n'+
      '  API.AUTH.LOG.001 = Auth, Login, 1st TC\n'+
      '  API.USER.CRT.001 = User, Create, 1st TC\n'+
      '  API.USER.CRT.002 = User, Create, 2nd TC (negative)\n'+
      '\n'+
      'Order: Positive (2xx) first → Negative (4xx) → Edge Case'
  );
  ws.getRange(2,4).setNote(
      'Feature — API feature or domain name.\n'+
      'Examples: Authentication, User Management, Meal Plan\n'+
      'Used for grouping Coverage per Feature in Summary.'
  );
  ws.getRange(2,5).setNote('HTTP Method: GET / POST / PUT / PATCH / DELETE');
  ws.getRange(2,6).setNote(
      'Endpoint URL — path only, no base URL.\n'+
      'Example: /api/v1/users/{id}\n'+
      'Dynamic params: use {curly_brackets}.'
  );
  ws.getRange(2,7).setNote(
      'Priority & Impact:\n'+
      '\n'+
      'CRITICAL  → Must PASS before release. FAIL = release BLOCKED.\n'+
      'HIGH      → Must PASS in same sprint. FAIL = needs PM approval.\n'+
      'MEDIUM    → Potential blocker. Fix before UAT.\n'+
      'LOW       → Non-blocker. Fix in next sprint.\n'+
      'LOWEST    → Nice to have. Optional.\n'+
      '\n'+
      'Auto Test Level:\n'+
      'Critical / High / Medium → Smoke Test\n'+
      'Low / Lowest             → Regression Test'
  );
  ws.getRange(2,8).setNote(
      'Auth — Authentication type required.\n'+
      'Examples: Bearer Token, Basic Auth, API Key, (none)'
  );
  ws.getRange(2,9).setNote(
      'Test Type:\n'+
      '  Positive   = happy path (valid payload, expected flow)\n'+
      '  Negative   = error case (invalid input, rejected request)\n'+
      '  Edge Case  = boundary condition'
  );
  ws.getRange(2,10).setNote(
      'Automation Status:\n'+
      '  Automated           = script exists and runs\n'+
      '  To Do               = planned, not yet done\n'+
      '  Manual              = decided to stay manual\n'+
      '  Cannot be Automated = technically not possible'
  );
  ws.getRange(2,11).setNote('Version: API version when TC was created. e.g. v1.0, v2.3');
  ws.getRange(2,12).setNote(
      'Role (RBAC) — The user role accessing this endpoint.\n'+
      '\n'+
      'Examples: Admin, Super Admin, User, Viewer\n'+
      '\n'+
      'Used to:\n'+
      '  • Verify correct role can access endpoint (2xx)\n'+
      '  • Verify wrong role gets 403 Forbidden\n'+
      '  • Verify no token gets 401 Unauthorized'
  );
  ws.getRange(2,13).setNote(
      'SCENARIO NAMING STANDARD\n'+
      '\n'+
      'Happy Path : [Role] Successfully [Verb] [Object] -- [status]\n'+
      'Negative   : [Role] Failed to [Verb] [Object] with [Condition] -- [status]\n'+
      '\n'+
      'Examples:\n'+
      '  User Successfully Creates Meal Plan -- 201\n'+
      '  User Failed to Authenticate with Invalid Token -- 401\n'+
      '\n'+
      'SCENARIO OUTLINE — same steps + same outcome type only:\n'+
      '\n'+
      'User Failed to Authenticate with <auth_condition> -- 401\n'+
      'Examples:\n'+
      '- Invalid token\n'+
      '- Empty token\n'+
      '- Expired token\n'+
      '\n'+
      'Rule: all Examples = same type and same status.\n'+
      'Separate Outlines for 401 vs 403 — different scenarios.'
  );
  ws.getRange(2,14).setNote(
      '[REQUIRED] Expected Result in Gherkin Then format:\n'+
      '\n'+
      '  Then : Outcome after request is sent\n'+
      '\n'+
      'Include: HTTP status code + key response body fields.\n'+
      '\n'+
      'Positive example:\n'+
      '  Then 201 Created\n'+
      '  body: { id, status: "active", created_at }\n'+
      'Negative example:\n'+
      '  Then 400 Bad Request\n'+
      '  body: { error: "field X is required" }\n'+
      'Auth example:\n'+
      '  Then 401 Unauthorized\n'+
      '  body: { message: "Token invalid or expired" }\n'+
      '\n'+
      'For Scenario Outline: write the common outcome.\n'+
      'Row detail goes in the Examples column.'
  );
  ws.getRange(2,15).setNote('[AUTO — DO NOT EDIT] Test Level auto-calculated from Priority:\nCritical/High/Medium = Smoke  |  Low/Lowest = Regression');

  const METHOD_COLORS={GET:['#E8F0FE','#1A237E'],POST:['#E8F5E9','#1B5E20'],
    PUT:['#FFF8E1','#E65100'],DELETE:['#FCE4EC','#880E4F'],PATCH:['#F3E5F5','#4A148C']};

  const data=[
    [1,'1.1','API.1.1.001','Informasi Program','GET','/api/v1/program/info','High','Bearer Token','Positive','Automated','v1.0','Viewer','Get informasi program -- 200','Then 200 OK\nbody: { data: [...], total: n }',''],
    [2,'1.1','API.1.1.002','Informasi Program','GET','/api/v1/program/info','Medium','(none)','Negative','Automated','v1.0','Viewer','Tanpa token ditolak -- 401','Then 401 Unauthorized\nbody: { message: "Token required" }',''],
    [3,'2.1','API.2.1.001','Perencanaan Pengiriman','POST','/api/v1/delivery/plan','Critical','Bearer Token','Positive','Automated','v1.1','Admin','Create rencana berhasil -- 201','Then 201 Created\nbody: { id, status: "pending", created_at }',''],
    [4,'2.1','API.2.1.002','Perencanaan Pengiriman','GET','/api/v1/delivery/plan','High','Bearer Token','Positive','Automated','v1.1','Admin','Ambil list rencana -- 200','Then 200 OK\nbody: { data: [...], total: n, page: 1 }',''],
    [5,'7.3','API.7.3.001','Manajemen Pengguna','POST','/api/v1/users','Critical','Bearer Token + Admin','Positive','Automated','v1.0','Super Admin','User Successfully Creates Account -- 201',
      'Then 201 Created\nbody: { id, email, role, status: "active", created_at }',
      ''],
    [6,'7.3','API.7.3.002','Manajemen Pengguna','POST','/api/v1/users','High','Bearer Token + Admin','Negative','Automated','v1.0','Super Admin','User Failed to Create Account with <invalid_payload> -- 400\nExamples:\n- Missing email\n- Missing password\n- Duplicate email',
      'Then 400 Bad Request\nbody: { error: <error_message> }',
      ''],
    [7,'7.3','API.7.3.003','Manajemen Pengguna','DELETE','/api/v1/users/{id}','Low','Bearer Token + Admin','Positive','Manual','v1.0','Super Admin','Delete user berhasil -- 200','Then 200 OK\nbody: { message: "User deleted successfully" }',''],
  ];
  data.forEach((row,i)=>{
    const r=DS+i, bg=i%2===0?'#F8F9FA':'#FFFFFF';
    row.forEach((val,c)=>{
      const cell=ws.getRange(r,c+1);
      bd(cell_style(cell,bg)).setValue(val||'');
      cell.setHorizontalAlignment(c===0?'center':'left');
      if(c===6) prioColor(cell,val);
      if(c===9) autoColor(cell,val);
      if(c===4){ const mc=METHOD_COLORS[val]; if(mc) cell.setBackground(mc[0]).setFontColor(mc[1]).setFontWeight('bold').setHorizontalAlignment('center'); }
      if(c===11) cell.setHorizontalAlignment('center').setFontWeight('bold').setFontColor('#1565C0');
      if(c===13) cell.setWrap(true);
      if(c===14) cell.setHorizontalAlignment('center').setFontWeight('bold');
    });
    bd(cell_style(ws.getRange(r,15),bg)).setFormula(tlFormula('G'+r)).setFontWeight('bold').setHorizontalAlignment('center');
    ws.setRowHeight(r,44);
  });
  for(let r=DS+data.length;r<=DS+MR;r++) ws.getRange(r,15).setFormula(tlFormula('G'+r));

  const dvEnd=DS+MR;
  // Expected Result col N: wrap + font only, no background override
  ws.getRange('N'+DS+':N'+dvEnd).setFontFamily('Arial').setFontSize(9).setWrap(true).setVerticalAlignment('middle');
  // TestLevel col O: alignment/bold only, CF handles Smoke/Regression colors
  ws.getRange('O'+DS+':O'+dvEnd).setHorizontalAlignment('center').setFontWeight('bold');
  // API_Master: no input borders
  ws.getRange('E'+DS+':E'+dvEnd).setDataValidation(dv(['GET','POST','PUT','DELETE','PATCH']));
  ws.getRange('G'+DS+':G'+dvEnd).setDataValidation(dv(['Critical','High','Medium','Low','Lowest']));
  ws.getRange('I'+DS+':I'+dvEnd).setDataValidation(dv(['Positive','Negative','Edge Case']));
  ws.getRange('J'+DS+':J'+dvEnd).setDataValidation(dv(['Automated','Manual','To Do','Cannot be Automated']));
  // Automated column CF
  (function(){
    const autoRange = ws.getRange('J'+DS+':J'+(DS+MR));
    const rules = ws.getConditionalFormatRules();
    [{v:'Automated',      bg:'#C8E6C9', fg:'#1B5E20'},
      {v:'Manual',         bg:'#E3F2FD', fg:'#1565C0'},
      {v:'To Do',          bg:'#FFF9C4', fg:'#F57F17'},
      {v:'Cannot be Automated', bg:'#F5F5F5', fg:'#9E9E9E'},
    ].forEach(s => rules.push(
        SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(s.v).setBackground(s.bg).setFontColor(s.fg).setBold(true)
            .setRanges([autoRange]).build()
    ));
    ws.setConditionalFormatRules(rules);
  })();
  ws.getRange('L'+DS+':L'+dvEnd).setDataValidation(dv(['Admin','User','Viewer','Operator','Supervisor','Guest','Super Admin']));
  const cf=ws.getConditionalFormatRules();
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Smoke')
      .setBackground('#FFF8F0').setFontColor('#BF360C').setBold(true).setRanges([ws.getRange('O'+DS+':O'+dvEnd)]).build());
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Regression')
      .setBackground('#F1F8E9').setFontColor('#33691E').setBold(true).setRanges([ws.getRange('O'+DS+':O'+dvEnd)]).build());
  Object.entries(METHOD_COLORS).forEach(([m,colors])=>
      cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(m)
          .setBackground(colors[0]).setFontColor(colors[1]).setBold(true)
          .setRanges([ws.getRange('E'+DS+':E'+dvEnd)]).build())
  );
  [{val:'Critical',fg:'#C62828'},{val:'High',fg:'#D84315'},{val:'Medium',fg:'#E65100'},
    {val:'Low',fg:'#2E7D32'},{val:'Lowest',fg:'#78909C'}].forEach(p=>
      cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(p.val)
          .setFontColor(p.fg).setBold(true).setRanges([ws.getRange('G'+DS+':G'+dvEnd)]).build())
  );
  ws.setConditionalFormatRules(cf);
  addPeruriFooter(ws, DS+data.length+3, COLS);
  ws.getRange(DS+data.length+2,1,1,COLS).merge()
      .setValue('Method (E) dan Endpoint URL (F) terpisah. Test Level otomatis dari Priority (G). Role (N) = tim penanggung jawab.')
      .setFontColor('#78909C').setFontStyle('italic').setFontSize(8).setFontFamily('Arial')
      .setBackground('#E3F2FD').setWrap(true);
}

// ===================================================================
//  TAB 4 -- API_Execution
// ===================================================================
function createAPIExecution(ss) {
  const ws = safeSheet(ss,'API_Execution');
  ws.clear(); ws.setTabColor('#283593');

  const DS=9, MR=500, STAG=7, STAG_N=3, STATUS_Z=26, SHOT_COL=27, MAX_RUNS=10;

  // -- Row 1 --
  hdr(ws.getRange(1,1),'#EEF0FB','#90A4AE',8).setValue('');
  for(let c=2;c<=5;c++) hdr(ws.getRange(1,c),'#EEF0FB','#90A4AE',8).setValue('');
  hdr(ws.getRange(1,6),'#EEF0FB','#90A4AE',8).setValue('TEST LEVEL').setWrap(true);
  hdr(ws.getRange(1,STAG),'#283593').setValue('<- Tambah kolom tanggal baru ke kanan');
  ws.setRowHeight(1,22);

  // -- Row 2: headers --
  ['TC_ID','SubModul','Feature','Priority','Endpoint (Method + URL)','[AUTO]\nTest Level']
      .forEach((h,i)=>hdr(ws.getRange(2,i+1),'#0D47A1').setValue(h).setWrap(true));
  ['2025-01-20','2025-02-10','2025-02-21']
      .forEach((h,i)=>hdr(ws.getRange(2,STAG+i),'#283593').setValue(h).setWrap(true));
  hdr(ws.getRange(2,STATUS_Z),'#455A64','#FFFFFF',8).setValue('[AUTO]\nLatest Status').setWrap(true);
  // Dynamic screenshot cols: one per run date
  for(let ri=0; ri<MAX_RUNS; ri++){
    const sc = SHOT_COL+ri;
    hdr(ws.getRange(2,sc),'#1565C0','#FFFFFF',8)
        .setFormula('=IFERROR(IF(INDIRECT(ADDRESS(2,'+(STAG+ri)+'))="","Screenshot "+'+(ri+1)+',TEXT(INDIRECT(ADDRESS(2,'+(STAG+ri)+')),"yyyy-mm-dd")&"\nShot"),"Screenshot '+(ri+1)+'")')
        .setWrap(true);
    ws.setColumnWidth(sc, 140);
  }
  ws.setRowHeight(2,38);


  // Column notes for API_Execution
  ws.getRange(2,1).setNote('[AUTO] TC_ID sync dari API_Master. Jangan edit.');
  ws.getRange(2,2).setNote('[AUTO] SubModul sync dari API_Master.');
  ws.getRange(2,3).setNote('[AUTO] Feature sync dari API_Master.');
  ws.getRange(2,4).setNote('[AUTO] Priority sync dari API_Master.');
  ws.getRange(2,5).setNote('[AUTO] Method + Endpoint sync dari API_Master.');
  ws.getRange(2,6).setNote('[AUTO] Test Level sync dari API_Master.');
  ws.getRange(2,STATUS_Z).setNote('[AUTO] Status terkini dihitung otomatis dari semua run.\nIN PROGRESS = ada PASSED tapi masih ada TODO belum dieksekusi.');
  ws.getRange(2,SHOT_COL).setNote('[INPUT] Link screenshot/response body/Postman result sebagai evidence.');  // -- Rows 3-7 --
  [{row:3,label:'> RUN STATUS',bg:'#37474F',fg:'#FFFFFF'},
    {row:4,label:'  PASSED',    bg:'#F1F8E9',fg:'#2E7D32'},
    {row:5,label:'  FAILED',    bg:'#FBE9E7',fg:'#BF360C'},
    {row:6,label:'  BLOCKED',   bg:'#FFF3E0',fg:'#E65100'},
    {row:7,label:'  TODO',      bg:'#F5F5F5',fg:'#546E7A'},
  ].forEach(s=>{
    ws.getRange(s.row,1,1,5).merge();
    bd(ws.getRange(s.row,1)).setValue(s.label)
        .setBackground(s.bg).setFontColor(s.fg).setFontWeight('bold')
        .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('right').setVerticalAlignment('middle');
    bd(ws.getRange(s.row,6)).setBackground(s.bg);
    ws.setRowHeight(s.row,20);
  });

  for(let i=0;i<STAG_N;i++){
    const col=STAG+i, L=colLetter(col), rng=L+'$'+DS+':'+L+'$'+(DS+MR);
    bd(ws.getRange(3,col)).setFormula(
        '=IF(COUNTA('+rng+')=0,"--",IF(COUNTIF('+rng+',"FAILED")>0,"FAILED",'+
        'IF(COUNTIF('+rng+',"BLOCKED")>0,"BLOCKED",'+
        'IF(COUNTIF('+rng+',"TODO")=COUNTA('+rng+'),"TODO",'+
        'IF(COUNTIF('+rng+',"TODO")>0,"IN PROGRESS","PASSED")))))'
    ).setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    [['PASSED',4],['FAILED',5],['BLOCKED',6],['TODO',7]].forEach(([st,row])=>
        bd(ws.getRange(row,col)).setFormula('=IF(COUNTA('+rng+')=0,"",COUNTIF('+rng+',"'+st+'"))')
            .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center')
    );
  }
  ws.getRange(4,STAG,1,STAG_N).setBackground('#F1F8E9');
  ws.getRange(5,STAG,1,STAG_N).setBackground('#FBE9E7');
  ws.getRange(6,STAG,1,STAG_N).setBackground('#FFF3E0');
  ws.getRange(7,STAG,1,STAG_N).setBackground('#F5F5F5');
  addRunStatusCF(ws,3,STAG,STAG+STAG_N+10);

  bd(ws.getRange(8,1)).setValue('v Kolom A-F = [AUTO sync API_Master] . Kolom G dst = [INPUT status per run] . Kolom Z = [AUTO status terkini] . Kolom AA = [INPUT evidence link]')
      .setBackground('#E8EAF6').setFontColor('#3949AB').setFontStyle('italic').setFontSize(8).setFontFamily('Arial');
  ws.setRowHeight(8,14);

  // Sync API_Master: C=TC_ID B=SubModul D=Feature G=Priority E=Method F=URL N=ExpResult O=TestLevel
  ws.getRange(DS,1).setFormula('=ARRAYFORMULA(IF(API_Master!C3:C'+(MR+2)+'<>"",API_Master!C3:C'+(MR+2)+',""))');
  ws.getRange(DS,2).setFormula('=ARRAYFORMULA(IF(API_Master!B3:B'+(MR+2)+'<>"",API_Master!B3:B'+(MR+2)+',""))');
  ws.getRange(DS,3).setFormula('=ARRAYFORMULA(IF(API_Master!D3:D'+(MR+2)+'<>"",API_Master!D3:D'+(MR+2)+',""))');
  ws.getRange(DS,4).setFormula('=ARRAYFORMULA(IF(API_Master!G3:G'+(MR+2)+'<>"",API_Master!G3:G'+(MR+2)+',""))');
  ws.getRange(DS,5).setFormula('=ARRAYFORMULA(IF(API_Master!C3:C'+(MR+2)+'<>"",API_Master!E3:E'+(MR+2)+'&"  "&API_Master!F3:F'+(MR+2)+',""))');
  ws.getRange(DS,6).setFormula('=ARRAYFORMULA(IF(API_Master!O3:O'+(MR+2)+'<>"",API_Master!O3:O'+(MR+2)+',""))');
  ws.getRange(DS,1,MR,6).setBackground('#E8EAF6').setFontColor('#37474F')
      .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle').setWrap(true);

  const cf=ws.getConditionalFormatRules();
  [{val:'Critical',bg:'#FFEBEE',fg:'#C62828'},{val:'High',bg:'#FFF3E0',fg:'#D84315'},
    {val:'Medium',bg:'#FFF8E1',fg:'#E65100'},{val:'Low',bg:'#F1F8E9',fg:'#2E7D32'},
    {val:'Lowest',bg:'#ECEFF1',fg:'#78909C'}].forEach(p=>
      cf.push(SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(p.val).setBackground(p.bg).setFontColor(p.fg).setBold(true)
          .setRanges([ws.getRange(DS,4,MR,1)]).build())
  );
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Smoke')
      .setBackground('#FFF8F0').setFontColor('#BF360C').setBold(true).setRanges([ws.getRange(DS,6,MR,1)]).build());
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Regression')
      .setBackground('#F1F8E9').setFontColor('#33691E').setBold(true).setRanges([ws.getRange(DS,6,MR,1)]).build());
  ws.setConditionalFormatRules(cf);

  [['PASSED','PASSED','PASSED'],['PASSED','PASSED','FAILED'],
    ['FAILED','PASSED','BLOCKED'],['TODO','TODO','TODO'],
    ['PASSED','PASSED','PASSED'],['TODO','TODO','TODO'],['TODO','TODO','TODO']].forEach((row,i)=>{
    const r=DS+i;
    row.forEach((val,c)=>{ bd(cell_style(ws.getRange(r,STAG+c),i%2===0?'#F8F9FA':'#FFFFFF'))
        .setValue(val).setHorizontalAlignment('center'); });
    ws.setRowHeight(r,34);
  });

  const STAG_COLS = STATUS_Z - STAG; // cols G(7) to Y(25) = 19
  statusCF_(ws, ws.getRange(DS, STAG, MR, STAG_COLS));
  ws.getRange(DS, STAG, MR, STAG_COLS).setDataValidation(dv(['TODO','PASSED','FAILED','BLOCKED']));

  // LATEST STATUS col Z -- ARRAYFORMULA
  for(let r=DS;r<DS+MR;r++){
    ws.getRange(r,STATUS_Z).setFormula(
        '=IF(A'+r+'="","",IF(COUNTIF(G'+r+':Y'+r+',"FAILED")>0,"FAILED",'+
        'IF(COUNTIF(G'+r+':Y'+r+',"BLOCKED")>0,"BLOCKED",'+
        'IF(COUNTA(G'+r+':Y'+r+')=0,"TODO",'+
        'IF(AND(COUNTIF(G'+r+':Y'+r+',"PASSED")>0,COUNTIF(G'+r+':Y'+r+',"TODO")>0),"IN PROGRESS",'+
        'IF(COUNTIF(G'+r+':Y'+r+',"PASSED")>0,"PASSED","TODO"))))))')
    ;
  }
  ws.getRange(DS,STATUS_Z,MR,1)
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  statusCF_(ws, ws.getRange(DS, STATUS_Z, MR, 1));

  for(let ri=0;ri<MAX_RUNS;ri++){
    ws.getRange(DS,SHOT_COL+ri,MR,1)
        .setBackground('#F0F4FF').setFontColor('#1A73E8')
        .setFontFamily('Arial').setFontSize(8).setVerticalAlignment('middle')
        .setWrap(false).setHorizontalAlignment('left');
  }

  [100,110,130,80,220,75].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  for(let i=0;i<STAG_N;i++) ws.setColumnWidth(STAG+i,100);
  ws.setColumnWidth(STATUS_Z,100);
  // screenshot col widths set in header loop above

  ws.setFrozenColumns(6);
  ws.setFrozenRows(2);

  ws.getRange(1, STAG+STAG_N).setValue('<- Insert kolom di sini untuk run baru. Isi tanggal di baris 2, copy formula baris 3?7.')
      .setFontColor('#90A4AE').setFontStyle('italic').setFontSize(8).setFontFamily('Arial');
}

// ===================================================================
//  TAB 5 -- _Dashboard
// ===================================================================
function createSummary(ss) {
  const ws = safeSheet(ss,'Summary');
  ws.clear(); ws.setTabColor('#0D47A1');

  function h_(range, bg, fg, sz) {
    return bd(range).setBackground(bg||'#0D47A1').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  function m_(row,col,nr,nc){ return ws.getRange(row,col,nr,nc).merge(); }
  function lbl(row,col,text,bg){
    bd(ws.getRange(row,col)).setValue(text)
        .setBackground(bg||'#E3F2FD').setFontColor('#0D47A1').setFontWeight('bold')
        .setFontFamily('Arial').setFontSize(9)
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
    ws.setRowHeight(row,24);
  }
  function inp(row,col,ncols,dv_list,isStatus){
    m_(row,col,1,ncols);
    const c=bd(ws.getRange(row,col)).setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(9)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
    if(dv_list) ws.getRange(row,col).setDataValidation(dv(dv_list));
    if(isStatus){ stCF(ws.getRange(row,col)); inputBorder(ws.getRange(row,col,1,ncols)); }
    return c;
  }
  function passRateCF(range){
    const rules=ws.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.8)
        .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.5,0.799)
        .setBackground('#FFF8E1').setFontColor('#E65100').setBold(true).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.5)
        .setBackground('#FFEBEE').setFontColor('#C62828').setBold(true).setRanges([range]).build());
    ws.setConditionalFormatRules(rules);
  }
  function stCF(range){
    const rules=ws.getConditionalFormatRules();
    [{v:'PASSED',bg:'#C8E6C9',fg:'#1B5E20'},{v:'IN PROGRESS',bg:'#E3F2FD',fg:'#1565C0'},
      {v:'FAILED',bg:'#FFCDD2',fg:'#B71C1C'},{v:'BLOCKED',bg:'#FFE0B2',fg:'#E65100'},
      {v:'--',bg:'#F5F5F5',fg:'#9E9E9E'},{v:'PASS',bg:'#C8E6C9',fg:'#1B5E20'},
      {v:'FAIL',bg:'#FFCDD2',fg:'#B71C1C'},{v:'Completed',bg:'#C8E6C9',fg:'#1B5E20'},
      {v:'In Progress',bg:'#E3F2FD',fg:'#1565C0'},{v:'Not Started',bg:'#F5F5F5',fg:'#616161'},
      {v:'On Hold',bg:'#FFE0B2',fg:'#E65100'}].forEach(s=>rules.push(
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s.v)
            .setBackground(s.bg).setFontColor(s.fg).setBold(true).setRanges([range]).build()));
    ws.setConditionalFormatRules(rules);
  }

  // Column layout: L cols 1-10, GAP col 11, R cols 12-21
  const L=1, G=11, R_=12, LW=10, RW=10;
  [95,55,55,55,55,55,60,55,60,55, 10, 95,55,55,55,55,55,60,55,60,55]
      .forEach((w,i)=>ws.setColumnWidth(i+1,w));
  ws.getRange(1,G,400,1).setBackground('#CFD8DC');

  const wZ='TC_Execution!Z9:Z1000', aZ='API_Execution!Z9:Z1000';
  const wTC='TC_Master!C3:C1000',   aTC='API_Master!C3:C1000';
  const wTOT='COUNTA('+wTC+')', aTOT='COUNTA('+aTC+')';

  let R=1;

  // =====================================================================
  // TOP: DESKRIPSI TEST (Test Plan Info) - left=Web/Mobile, right=API
  // =====================================================================
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#0D47A1','#FFFFFF',11).setValue('TEST DESCRIPTION  -  Web / Mobile');
  m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#283593','#FFFFFF',11).setValue('TEST DESCRIPTION  -  API');
  ws.setRowHeight(R,30); R++;

  // Left plan fields
  const leftFields=[
    ['Project:',null],
    ['Modul:',null],
    ['Submodul:',null,false,false,true], // tambah note
    ['QA Lead:',null],
    ['PIC QA:',null],
    ['Environment:',['Dev','Staging / UAT','Production']],
    ['Issue Tracker (URL):',null],
    ['Test Status:',['Not Started','In Progress','Completed','On Hold'],true],
    ['Scope / Notes:',null],
  ];
  const rightFields=[
    ['Base URL:',null],['API Version:',null],['Authentication:',null],
    ['Environment:',['Dev','Staging / UAT','Production']],
    ['Collection / Docs URL:',null],
    ['Test Status:',['Not Started','In Progress','Completed','On Hold'],true],
    ['Perf Test Result:',null,false,true],
    ['Notes:',null],
  ];
  leftFields.forEach(([labelText, dvList, isStatus, isAutoPerf, addNote],i)=>{
    const row=R+i;
    lbl(row,L,labelText,'#E3F2FD');
    if(isAutoPerf){
      m_(row,L+1,1,LW-1);
      bd(ws.getRange(row,L+1)).setFormula('=IFERROR(IF(COUNTA(PerfTest!E16:E45)=0,"--",IF(COUNTIF(PerfTest!L16:L45,"FAIL")>0,"FAIL","PASS")),"--")')
          .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');
      stCF(ws.getRange(row,L+1));
    } else {
      inp(row,L+1,LW-1,dvList,isStatus);
      if(addNote){
        ws.getRange(row,L+1,1,LW-1).setNote('Pisahkan dengan koma jika lebih dari 1 submodul.\nContoh: 1.1,1.2,1.3');
      }
    }
    ws.setRowHeight(row,24);
  });
  rightFields.forEach(([labelText, dvList, isStatus, isAutoPerf],i)=>{
    const row=R+i;
    lbl(row,R_,labelText,'#E8EAF6');
    if(isAutoPerf){
      m_(row,R_+1,1,RW-1);
      bd(ws.getRange(row,R_+1)).setFormula('=IFERROR(IF(COUNTA(PerfTest!E16:E45)=0,"--",IF(COUNTIF(PerfTest!L16:L45,"FAIL")>0,"FAIL","PASS")),"--")')
          .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');
      stCF(ws.getRange(row,R_+1));
    } else {
      inp(row,R_+1,RW-1,dvList,isStatus);
    }
    ws.setRowHeight(row,24);
  });
  R+=Math.max(leftFields.length,rightFields.length);
  ws.setRowHeight(R,6); R++;

  // =====================================================================
  // A. STATUS OVERVIEW
  // =====================================================================
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#1565C0').setValue('A.  STATUS OVERVIEW  -  Web / Mobile');
  m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#283593').setValue('A.  STATUS OVERVIEW  -  API');
  ws.setRowHeight(R,20); R++;

  const kpiLabels=['TOTAL','PASSED','FAILED','BLOCKED','IN PROG','TODO','PASS RATE','AUTO RATE','EXEC RATE','--'];
  const kpiBgs=['#37474F','#2E7D32','#B71C1C','#E65100','#1565C0','#546E7A','#0D47A1','#1976D2','#4A148C',''];
  for(let i=0;i<9;i++){
    h_(ws.getRange(R,L+i),kpiBgs[i]).setValue(kpiLabels[i]).setFontSize(i<6?8:7.5).setWrap(true);
    h_(ws.getRange(R,R_+i),kpiBgs[i]).setValue(kpiLabels[i]).setFontSize(i<6?8:7.5).setWrap(true);
  }
  ws.setRowHeight(R,22); R++;

  const wForms=[
    '=COUNTA('+wTC+')',
    '=COUNTIF('+wZ+',"PASSED")',
    '=COUNTIF('+wZ+',"FAILED")',
    '=COUNTIF('+wZ+',"BLOCKED")',
    '=COUNTIF('+wZ+',"IN PROGRESS")',
    '=COUNTIF('+wZ+',"TODO")',
    '=IFERROR(COUNTIF('+wZ+',"PASSED")/MAX(1,'+wTOT+'),0)',
    '=IFERROR(COUNTIF(TC_Master!H3:H1000,"Automated")/MAX(1,'+wTOT+'),0)',
    '=IFERROR((COUNTIF('+wZ+',"PASSED")+COUNTIF('+wZ+',"FAILED")+COUNTIF('+wZ+',"BLOCKED")+COUNTIF('+wZ+',"IN PROGRESS"))/MAX(1,'+wTOT+'),0)',
  ];
  const aForms=[
    '=COUNTA('+aTC+')',
    '=COUNTIF('+aZ+',"PASSED")',
    '=COUNTIF('+aZ+',"FAILED")',
    '=COUNTIF('+aZ+',"BLOCKED")',
    '=COUNTIF('+aZ+',"IN PROGRESS")',
    '=COUNTIF('+aZ+',"TODO")',
    '=IFERROR(COUNTIF('+aZ+',"PASSED")/MAX(1,'+aTOT+'),0)',
    '=IFERROR(COUNTIF(API_Master!J3:J1000,"Automated")/MAX(1,'+aTOT+'),0)',
    '=IFERROR((COUNTIF('+aZ+',"PASSED")+COUNTIF('+aZ+',"FAILED")+COUNTIF('+aZ+',"BLOCKED")+COUNTIF('+aZ+',"IN PROGRESS"))/MAX(1,'+aTOT+'),0)',
  ];
  wForms.forEach((f,i)=>{
    const c=bd(ws.getRange(R,L+i)).setFormula(f).setBackground('#FFFFFF')
        .setFontWeight('bold').setFontSize(i<6?16:13).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    if(i>=6){ c.setNumberFormat('0%'); passRateCF(ws.getRange(R,L+i)); }
  });
  aForms.forEach((f,i)=>{
    const c=bd(ws.getRange(R,R_+i)).setFormula(f).setBackground('#FFFFFF')
        .setFontWeight('bold').setFontSize(i<6?16:13).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    if(i>=6){ c.setNumberFormat('0%'); passRateCF(ws.getRange(R,R_+i)); }
  });
  ws.setRowHeight(R,36); R++;

  m_(R,L,1,LW);
  ws.getRange(R,L).setValue('Hijau >=80%  |  Kuning 50-79%  |  Merah <50%  |  Exec Rate = % TC sudah dapat hasil (bukan TODO)')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic').setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left');
  m_(R,R_,1,RW);
  ws.getRange(R,R_).setValue('Hijau >=80%  |  Kuning 50-79%  |  Merah <50%')
      .setBackground('#E8EAF6').setFontColor('#283593').setFontStyle('italic').setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(R,14); R++;

  // ── SMOKE TEST SUB-ROW ────────────────────────────────────────────
  // Smoke = Priority Critical / High / Medium (test level auto)
  // Web/Mobile side (left), API side (right)
  const wPrio='TC_Master!E3:E1000', aPrio='API_Master!G3:G1000';
  const smokePrios=['Critical','High','Medium'];
  function smokeCount(zRange, prioRange, stat) {
    return smokePrios.map(p=>'COUNTIFS('+prioRange+',"'+p+'",'+zRange+',"'+stat+'")').join('+');
  }
  function smokeTot(tcRange, prioRange) {
    return smokePrios.map(p=>'COUNTIF('+prioRange+',"'+p+'")').join('+');
  }

  // Sub-header row
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#BF360C').setValue('A1.  SMOKE TEST  -  Web / Mobile  (Critical + High + Medium)');
  m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#4A148C').setValue('A1.  SMOKE TEST  -  API  (Critical + High + Medium)');
  ws.setRowHeight(R,18); R++;

  // KPI label row
  for(let i=0;i<9;i++){
    h_(ws.getRange(R,L+i),kpiBgs[i]).setValue(kpiLabels[i]).setFontSize(i<6?8:7.5).setWrap(true);
    h_(ws.getRange(R,R_+i),kpiBgs[i]).setValue(kpiLabels[i]).setFontSize(i<6?8:7.5).setWrap(true);
  }
  ws.setRowHeight(R,22); R++;

  // Smoke KPI value row
  // FIX: wZ=Z9:Z1000 (992 rows), wPrio=E3:E1000 (998 rows) → mismatch → #VALUE!
  // Solution: use wPrioExec matched to Z9:Z1000 row count (E3:E994 = 992 rows)
  // smokeTot (from TC_Master only) still uses wPrio=E3:E1000 — same-sheet, same size → OK
  const wPrioExec='TC_Master!E3:E994';   // 992 rows → matches wZ=Z9:Z1000
  const aPrioExec='API_Master!G3:G994';  // 992 rows → matches aZ=Z9:Z1000
  const wSmokeTot='('+smokeTot(wTC,wPrio)+')';
  const aSmokeTot='('+smokeTot(aTC,aPrio)+')';
  const wSmokeForms=[
    '='+wSmokeTot,
    '='+smokeCount(wZ,wPrioExec,'PASSED'),
    '='+smokeCount(wZ,wPrioExec,'FAILED'),
    '='+smokeCount(wZ,wPrioExec,'BLOCKED'),
    '='+smokeCount(wZ,wPrioExec,'IN PROGRESS'),
    '='+smokeCount(wZ,wPrioExec,'TODO'),
    // PASS RATE: passed smoke / total smoke TC
    '=IFERROR(('+smokeCount(wZ,wPrioExec,'PASSED')+')/MAX(1,'+wSmokeTot+'),0)',
    // AUTO RATE: fix — wrap all 3 COUNTIFS in () before dividing
    '=IFERROR((COUNTIFS(TC_Master!H3:H994,"Automated",'+wPrioExec+',"Critical")+COUNTIFS(TC_Master!H3:H994,"Automated",'+wPrioExec+',"High")+COUNTIFS(TC_Master!H3:H994,"Automated",'+wPrioExec+',"Medium"))/MAX(1,'+wSmokeTot+'),0)',
    // EXEC RATE: (passed+failed+blocked+inprog) / total smoke
    '=IFERROR(('+smokeCount(wZ,wPrioExec,'PASSED')+'+'+smokeCount(wZ,wPrioExec,'FAILED')+'+'+smokeCount(wZ,wPrioExec,'BLOCKED')+'+'+smokeCount(wZ,wPrioExec,'IN PROGRESS')+')/MAX(1,'+wSmokeTot+'),0)',
  ];
  const aSmokeForms=[
    '='+aSmokeTot,
    '='+smokeCount(aZ,aPrioExec,'PASSED'),
    '='+smokeCount(aZ,aPrioExec,'FAILED'),
    '='+smokeCount(aZ,aPrioExec,'BLOCKED'),
    '='+smokeCount(aZ,aPrioExec,'IN PROGRESS'),
    '='+smokeCount(aZ,aPrioExec,'TODO'),
    '=IFERROR(('+smokeCount(aZ,aPrioExec,'PASSED')+')/MAX(1,'+aSmokeTot+'),0)',
    '=IFERROR((COUNTIFS(API_Master!J3:J994,"Automated",'+aPrioExec+',"Critical")+COUNTIFS(API_Master!J3:J994,"Automated",'+aPrioExec+',"High")+COUNTIFS(API_Master!J3:J994,"Automated",'+aPrioExec+',"Medium"))/MAX(1,'+aSmokeTot+'),0)',
    '=IFERROR(('+smokeCount(aZ,aPrioExec,'PASSED')+'+'+smokeCount(aZ,aPrioExec,'FAILED')+'+'+smokeCount(aZ,aPrioExec,'BLOCKED')+'+'+smokeCount(aZ,aPrioExec,'IN PROGRESS')+')/MAX(1,'+aSmokeTot+'),0)',
  ];
  wSmokeForms.forEach((f,i)=>{
    const c=bd(ws.getRange(R,L+i)).setFormula(f).setBackground('#FFF3E0')
        .setFontWeight('bold').setFontSize(i<6?16:13).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setFontColor('#BF360C');
    if(i>=6){ c.setNumberFormat('0%'); passRateCF(ws.getRange(R,L+i)); }
  });
  aSmokeForms.forEach((f,i)=>{
    const c=bd(ws.getRange(R,R_+i)).setFormula(f).setBackground('#F3E5F5')
        .setFontWeight('bold').setFontSize(i<6?16:13).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setFontColor('#4A148C');
    if(i>=6){ c.setNumberFormat('0%'); passRateCF(ws.getRange(R,R_+i)); }
  });
  ws.setRowHeight(R,36); R++;

  // Smoke legend
  m_(R,L,1,LW);
  ws.getRange(R,L).setValue('Smoke Test = TC dengan Priority Critical / High / Medium  |  Target Pass Rate >= 80% sebelum release')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic').setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left');
  m_(R,R_,1,RW);
  ws.getRange(R,R_).setValue('Smoke Test = API dengan Priority Critical / High / Medium')
      .setBackground('#EDE7F6').setFontColor('#4A148C').setFontStyle('italic').setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(R,14); R++;
  // ── END SMOKE TEST ────────────────────────────────────────────────

  // =====================================================================
  // B. KOMPOSISI STATUS - Header, then chart, then small data table
  // =====================================================================
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#1565C0').setValue('B.  KOMPOSISI STATUS  -  Web + Mobile');
  m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#283593').setValue('B.  KOMPOSISI STATUS  -  API Only');
  ws.setRowHeight(R,20);
  const PIE_HDR_ROW=R; R++;

  // Small data table (cols 1-3 left, 12-14 right) ? chart will sit beside/below
  const PIE_DATA_ROW=R;
  // Headers
  ['Status','N','%'].forEach((h,i)=>{
    h_(ws.getRange(R,L+i),'#0D47A1').setValue(h).setFontSize(8);
    h_(ws.getRange(R,R_+i),'#283593').setValue(h).setFontSize(8);
  });
  ws.setRowHeight(R,16); R++;
  [['PASSED',  '=COUNTIF('+wZ+',"PASSED")',   '=COUNTIF('+aZ+',"PASSED")'],
    ['FAILED',  '=COUNTIF('+wZ+',"FAILED")',   '=COUNTIF('+aZ+',"FAILED")'],
    ['BLOCKED', '=COUNTIF('+wZ+',"BLOCKED")',  '=COUNTIF('+aZ+',"BLOCKED")'],
    ['IN PROG', '=COUNTIF('+wZ+',"IN PROGRESS")','=COUNTIF('+aZ+',"IN PROGRESS")'],
    ['TODO',    '=COUNTIF('+wZ+',"TODO")',      '=COUNTIF('+aZ+',"TODO")'],
  ].forEach((row,i)=>{
    const rr=R+i, bg=i%2===0?'#F8F9FA':'#FFFFFF';
    const totW='SUM('+colLetter(L+1)+(PIE_DATA_ROW+1)+':'+colLetter(L+1)+(PIE_DATA_ROW+5)+')';
    const totA='SUM('+colLetter(R_+1)+(PIE_DATA_ROW+1)+':'+colLetter(R_+1)+(PIE_DATA_ROW+5)+')';
    bd(ws.getRange(rr,L)).setValue(row[0]).setBackground(bg).setFontFamily('Arial').setFontSize(8).setHorizontalAlignment('left');
    bd(ws.getRange(rr,L+1)).setFormula(row[1]).setBackground(bg).setFontFamily('Arial').setFontSize(8).setHorizontalAlignment('center').setFontWeight('bold');
    bd(ws.getRange(rr,L+2)).setFormula('=IFERROR('+colLetter(L+1)+rr+'/MAX(1,'+totW+'),0)').setBackground(bg).setFontFamily('Arial').setFontSize(8).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rr,R_)).setValue(row[0]).setBackground(bg).setFontFamily('Arial').setFontSize(8).setHorizontalAlignment('left');
    bd(ws.getRange(rr,R_+1)).setFormula(row[2]).setBackground(bg).setFontFamily('Arial').setFontSize(8).setHorizontalAlignment('center').setFontWeight('bold');
    bd(ws.getRange(rr,R_+2)).setFormula('=IFERROR('+colLetter(R_+1)+rr+'/MAX(1,'+totA+'),0)').setBackground(bg).setFontFamily('Arial').setFontSize(8).setHorizontalAlignment('center').setNumberFormat('0%');
    ws.setRowHeight(rr,15);
  });
  R+=5;

  // Pie charts anchored BELOW header row (PIE_HDR_ROW+1), offset right of data table
  try {
    ws.insertChart(ws.newChart().setChartType(Charts.ChartType.PIE)
        .addRange(ws.getRange(PIE_DATA_ROW,L,6,2))
        .setPosition(PIE_HDR_ROW+1,L+3,0,0)
        .setOption('title','Web + Mobile').setOption('pieHole',0.4)
        .setOption('colors',['#4CAF50','#F44336','#FF9800','#2196F3','#9E9E9E'])
        .setOption('pieSliceText','percentage')
        .setOption('legend',{position:'right'}).setOption('width',300).setOption('height',170).build());
  }catch(e){}
  try {
    ws.insertChart(ws.newChart().setChartType(Charts.ChartType.PIE)
        .addRange(ws.getRange(PIE_DATA_ROW,R_,6,2))
        .setPosition(PIE_HDR_ROW+1,R_+3,0,0)
        .setOption('title','API Only').setOption('pieHole',0.4)
        .setOption('colors',['#4CAF50','#F44336','#FF9800','#2196F3','#9E9E9E'])
        .setOption('pieSliceText','percentage')
        .setOption('legend',{position:'right'}).setOption('width',300).setOption('height',170).build());
  }catch(e){}
  // space for charts (about 8 rows @ ~22px = 176px)
  for(let i=0;i<8;i++){ ws.setRowHeight(R+i,22); }
  R+=8;

  // =====================================================================
  // C. TREND EKSEKUSI - Header, chart below, date table below chart
  // =====================================================================
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#1565C0').setValue('C.  TREND EKSEKUSI  -  Web / Mobile  (per Tanggal Run)');
  m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#283593').setValue('C.  TREND EKSEKUSI  -  API  (per Tanggal Run)');
  ws.setRowHeight(R,20);
  const TREND_HDR=R; R++;

  // Space for trend charts (7 rows)
  for(let i=0;i<7;i++){ ws.setRowHeight(R+i,22); }
  R+=7;

  // Date table BELOW charts
  ['Tanggal','Passed','Failed','Blocked','Pass%','Exec%','Status'].forEach((h,i)=>{
    h_(ws.getRange(R,L+i),'#0D47A1').setValue(h).setFontSize(8).setWrap(true);
    h_(ws.getRange(R,R_+i),'#283593').setValue(h).setFontSize(8).setWrap(true);
  });
  ws.setRowHeight(R,20);
  const RH_DATE_HDR=R; R++;
  const RH_L=R, RH_A=R;

  // Dynamic trend rows ? dates auto-pulled from TC_Execution row 2 (col H onwards)
  const MAX_RUNS = 15;
  for (let idx=0; idx<MAX_RUNS; idx++) {
    const rl=R+idx, bg=idx%2===0?'#F8F9FA':'#FFFFFF';
    const tcDateCol = 'INDIRECT("TC_Execution!"&ADDRESS(2,'+idx+'+8))'; // col H=8, I=9...
    // Date: pull from TC_Execution row 2, col H+idx
    const dateForm = '=IFERROR(INDEX(TC_Execution!$2:$2,'+(idx+8)+'),"")';
    const aDateForm = '=IFERROR(INDEX(API_Execution!$2:$2,'+(idx+7)+'),"")';
    // Web row - date auto from TC_Execution row 2
    bd(ws.getRange(rl,L)).setFormula(dateForm).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left').setFontWeight('bold').setNumberFormat('yyyy-mm-dd');
    const mfl='MATCH('+colLetter(L)+rl+',TC_Execution!$2:$2,0)';
    bd(ws.getRange(rl,L+1)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",INDEX(TC_Execution!$4:$4,'+mfl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,L+2)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",INDEX(TC_Execution!$5:$5,'+mfl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,L+3)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",INDEX(TC_Execution!$6:$6,'+mfl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,L+4)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",'+colLetter(L+1)+rl+'/MAX(1,'+colLetter(L+1)+rl+'+'+colLetter(L+2)+rl+'+'+colLetter(L+3)+rl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rl,L+5)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'=""," ",('+colLetter(L+1)+rl+'+'+colLetter(L+2)+rl+'+'+colLetter(L+3)+rl+')/MAX(1,'+wTOT+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rl,L+6)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",INDEX(TC_Execution!$3:$3,'+mfl+')),"--")').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
    // API row
    bd(ws.getRange(rl,R_)).setFormula(aDateForm).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left').setFontWeight('bold').setNumberFormat('yyyy-mm-dd');
    const mfa='MATCH('+colLetter(R_)+rl+',API_Execution!$2:$2,0)';
    bd(ws.getRange(rl,R_+1)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",INDEX(API_Execution!$4:$4,'+mfa+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,R_+2)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",INDEX(API_Execution!$5:$5,'+mfa+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,R_+3)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",INDEX(API_Execution!$6:$6,'+mfa+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,R_+4)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",'+colLetter(R_+1)+rl+'/MAX(1,'+colLetter(R_+1)+rl+'+'+colLetter(R_+2)+rl+'+'+colLetter(R_+3)+rl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rl,R_+5)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'=""," ",('+colLetter(R_+1)+rl+'+'+colLetter(R_+2)+rl+'+'+colLetter(R_+3)+rl+')/MAX(1,'+aTOT+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rl,R_+6)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",INDEX(API_Execution!$3:$3,'+mfa+')),"--")').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
    ws.setRowHeight(rl,18);
  }
  passRateCF(ws.getRange(RH_L,L+4,3,1)); passRateCF(ws.getRange(RH_L,L+5,3,1));
  stCF(ws.getRange(RH_L,L+6,3,1));
  passRateCF(ws.getRange(RH_A,R_+4,3,1)); passRateCF(ws.getRange(RH_A,R_+5,3,1));
  stCF(ws.getRange(RH_A,R_+6,3,1));
  R+=3;

  // Trend charts anchored at TREND_HDR+1 (just below section header)
  try {
    ws.insertChart(ws.newChart().setChartType(Charts.ChartType.LINE)
        .addRange(ws.getRange(RH_DATE_HDR,L,4,1))
        .addRange(ws.getRange(RH_DATE_HDR,L+4,4,1))
        .setPosition(TREND_HDR+1,L,0,0)
        .setOption('title','Trend Pass Rate - Web/Mobile')
        .setOption('curveType','function').setOption('legend',{position:'none'})
        .setOption('colors',['#1565C0']).setOption('pointSize',6)
        .setOption('vAxis',{format:'0%',minValue:0,maxValue:1})
        .setOption('width',320).setOption('height',150).build());
  }catch(e){}
  try {
    ws.insertChart(ws.newChart().setChartType(Charts.ChartType.LINE)
        .addRange(ws.getRange(RH_DATE_HDR,R_,4,1))
        .addRange(ws.getRange(RH_DATE_HDR,R_+4,4,1))
        .setPosition(TREND_HDR+1,R_,0,0)
        .setOption('title','Trend Pass Rate - API')
        .setOption('curveType','function').setOption('legend',{position:'none'})
        .setOption('colors',['#283593']).setOption('pointSize',6)
        .setOption('vAxis',{format:'0%',minValue:0,maxValue:1})
        .setOption('width',320).setOption('height',150).build());
  }catch(e){}

  m_(R,L,1,LW);
  ws.getRange(R,L).setValue('+ Tambah baris run baru: isi tanggal di kolom pertama, formula lain pickup otomatis dari Execution sheet.')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic').setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  m_(R,R_,1,RW);
  ws.getRange(R,R_).setValue('+ Tambah baris run baru: isi tanggal di kolom pertama.')
      .setBackground('#E8EAF6').setFontColor('#283593').setFontStyle('italic').setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(R,14); R++;

  // =====================================================================
  // D. BUG SUMMARY (above Coverage)
  const BUG_COL = 'BugReport!D5:D5000'; // Status col (col 4), DS=5
  const BUG_PRO = 'BugReport!C5:C5000'; // Priority col (col 3)
  const BUG_TYP = 'BugReport!B5:B5000'; // Type col (col 2)

  // Left header
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#0D47A1').setValue('D.  BUG SUMMARY  -  Web + Mobile');
  ws.setRowHeight(R,20);
  // Right header
  m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#1565C0').setValue('D.  BUG SUMMARY  -  API');
  R++;

  // Bug KPI cards (left = Web+Mobile, right = API)
  const bugMetrics = [
    ['Total Bugs',  null, null],
    ['Open',        'Open', null],
    ['In Progress', 'In Progress', null],
    ['Fixed',       'Fixed', null],
    ['Verified',    'Verified', null],
    ['Critical',    null, 'Critical'],
    ['High',        null, 'High'],
    ['Medium', null, 'Medium'],
  ];

  const bugColors = {
    'Total Bugs':'#E3F2FD','Open':'#FFCDD2','In Progress':'#E3F2FD',
    'Fixed':'#FFF9C4','Verified':'#C8E6C9','Critical':'#FFCDD2',
    'High':'#FFE0B2','Medium':'#E3F2FD'
  };
  const bugFgColors = {
    'Total Bugs':'#0D47A1','Open':'#B71C1C','In Progress':'#1565C0',
    'Fixed':'#E65100','Verified':'#2E7D32','Critical':'#B71C1C',
    'High':'#E65100','Medium':'#1565C0'
  };

  // Helpers for bug formula
  function bugCount(typFilter, statFilter, prioFilter) {
    let inner;
    if (typFilter === 'API') {
      if (statFilter)      inner = 'COUNTIFS('+BUG_TYP+',"API",'+BUG_COL+',"'+statFilter+'")';
      else if (prioFilter) inner = 'COUNTIFS('+BUG_TYP+',"API",'+BUG_PRO+',"'+prioFilter+'")';
      else                 inner = 'COUNTIF('+BUG_TYP+',"API")';
    } else {
      if (statFilter) {
        inner = 'COUNTIFS('+BUG_TYP+',"Web",'+BUG_COL+',"'+statFilter+'")+'+
            'COUNTIFS('+BUG_TYP+',"Mobile",'+BUG_COL+',"'+statFilter+'")';
      } else if (prioFilter) {
        inner = 'COUNTIFS('+BUG_TYP+',"Web",'+BUG_PRO+',"'+prioFilter+'")+'+
            'COUNTIFS('+BUG_TYP+',"Mobile",'+BUG_PRO+',"'+prioFilter+'")';
      } else {
        inner = 'COUNTIF('+BUG_TYP+',"Web")+COUNTIF('+BUG_TYP+',"Mobile")';
      }
    }
    return '=IFERROR('+inner+',0)';
  }

  bugMetrics.forEach(([label, statF, sevF], i) => {
    const rr = R+i, bg = bugColors[label]||'#FFFFFF', fg = bugFgColors[label]||'#333333';
    // Left (Web+Mobile)
    bd(ws.getRange(rr,L)).setValue(label+':').setBackground('#FFEBEE').setFontFamily('Arial')
        .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right')
        .setFontColor('#C62828').setVerticalAlignment('middle');
    m_(rr,L+1,1,LW-1);
    bd(ws.getRange(rr,L+1)).setFormula(bugCount(null,statF,sevF))
        .setBackground(bg).setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
        .setFontColor(fg).setHorizontalAlignment('center').setVerticalAlignment('middle');
    // Right (API)
    bd(ws.getRange(rr,R_)).setValue(label+':').setBackground('#FFEBEE').setFontFamily('Arial')
        .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right')
        .setFontColor('#B71C1C').setVerticalAlignment('middle');
    m_(rr,R_+1,1,RW-1);
    bd(ws.getRange(rr,R_+1)).setFormula(bugCount('API',statF,sevF))
        .setBackground(bg).setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
        .setFontColor(fg).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setRowHeight(rr,24);
  });
  R += bugMetrics.length;

  // ── SMOKE BLOCKER ROW ─────────────────────────────────────────────
  // Objective: berapa bug yang masih Open/In Progress/Reopen
  //            dengan Priority Critical/High/Medium (= blocker release)
  // Formula: SUMPRODUCT — lebih bersih dari 9x COUNTIFS terpisah
  // BugReport data mulai row 5, col B=Type, C=Priority, D=Status
  const BLOCKER_FORMULA =
      'SUMPRODUCT(' +
      '(ISNUMBER(MATCH(BugReport!D5:D2000,{\"Open\",\"In Progress\",\"Reopen\",\"In Progress VAPT\",\"Done VAPT\"},0)))*' +
      '(ISNUMBER(MATCH(BugReport!C5:C2000,{\"Critical\",\"High\",\"Medium\"},0)))' +
      ')';
  // Sama untuk kedua sisi (Web dan API mengacu ke BugReport yang sama)
  const smokeOpenW = BLOCKER_FORMULA;
  const smokeOpenA = BLOCKER_FORMULA;

  // Separator label row
  m_(R,L,1,LW);
  ws.getRange(R,L).setValue('Open Blocker (Smoke) ↓')
      .setBackground('#BF360C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#E57373',SpreadsheetApp.BorderStyle.SOLID);
  m_(R,R_,1,RW);
  ws.getRange(R,R_).setValue('Open Blocker (Smoke) ↓')
      .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#CE93D8',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(R,14); R++;

  // Smoke Open count label + value
  bd(ws.getRange(R,L)).setValue('Open Blocker:').setBackground('#FFEBEE').setFontFamily('Arial')
      .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right')
      .setFontColor('#C62828').setVerticalAlignment('middle');
  m_(R,L+1,1,LW-1);
  bd(ws.getRange(R,L+1)).setFormula('=IFERROR('+smokeOpenW+',0)')
      .setBackground('#FFCDD2').setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
      .setFontColor('#B71C1C').setHorizontalAlignment('center').setVerticalAlignment('middle');
  bd(ws.getRange(R,R_)).setValue('Open Blocker:').setBackground('#FFEBEE').setFontFamily('Arial')
      .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right')
      .setFontColor('#B71C1C').setVerticalAlignment('middle');
  m_(R,R_+1,1,RW-1);
  bd(ws.getRange(R,R_+1)).setFormula('=IFERROR('+smokeOpenA+',0)')
      .setBackground('#EDE7F6').setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
      .setFontColor('#4A148C').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R,28);

  // Legend
  const smokeNote = ws.getRange(R, L+2);
  smokeNote.setValue('Target: 0 Open Blocker sebelum release')
      .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
  ws.getRange(R,L+2).setNote(
      'Open Blocker (Smoke) - UPDATED WITH VAPT\n\n' +
      'Bug yang dihitung sebagai blocker:\n' +
      '  • Status: Open / In Progress / Reopen / In Progress VAPT / Done VAPT\n' +
      '  • Priority: Critical / High / Medium\n\n' +
      'NOT Blocker:\n' +
      '  • Verified (sudah OK QA, waiting VAPT atau skip VAPT)\n' +
      '  • Closed (final)\n' +
      '  • Won\'t Fix (rejected)\n\n' +
      'Why Done VAPT is still blocker?\n' +
      'Bug perlu re-test QA sebelum Closed. Hanya Closed yang tidak blocker.\n\n' +
      'Target: 0 Open Blocker sebelum release ke production.'
  );
  R += 2;
  // ── END SMOKE BLOCKER ─────────────────────────────────────────────

  // ── PROD BUGS ROW ─────────────────────────────────────────────────
  // Objective: berapa bug yang ada di Production environment
  // Status: Open/In Progress/Reopen/Fixed/Verified (belum Closed)
  // BugReport col I = Environment (row 5 onwards)
  const PROD_BUGS_FORMULA =
      'SUMPRODUCT(' +
      '(ISNUMBER(MATCH(BugReport!D5:D2000,{\"Open\",\"In Progress\",\"Reopen\",\"Fixed\",\"Verified\"},0)))*' +
      '(BugReport!I5:I2000=\"Production\")' +
      ')';

  // Separator label row
  m_(R,L,1,LW);
  ws.getRange(R,L).setValue('Bugs in Production ↓')
      .setBackground('#D32F2F').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#EF5350',SpreadsheetApp.BorderStyle.SOLID);
  m_(R,R_,1,RW);
  ws.getRange(R,R_).setValue('Bugs in Production ↓')
      .setBackground('#D32F2F').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#EF5350',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(R,14); R++;

  // Prod Bugs count label + value
  bd(ws.getRange(R,L)).setValue('PROD BUGS:').setBackground('#FFEBEE').setFontFamily('Arial')
      .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right')
      .setFontColor('#C62828').setVerticalAlignment('middle');
  m_(R,L+1,1,LW-1);
  bd(ws.getRange(R,L+1)).setFormula('=IFERROR('+PROD_BUGS_FORMULA+',0)')
      .setBackground('#FFCDD2').setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
      .setFontColor('#B71C1C').setHorizontalAlignment('center').setVerticalAlignment('middle');
  bd(ws.getRange(R,R_)).setValue('PROD BUGS:').setBackground('#FFEBEE').setFontFamily('Arial')
      .setFontSize(9).setFontWeight('bold').setHorizontalAlignment('right')
      .setFontColor('#B71C1C').setVerticalAlignment('middle');
  m_(R,R_+1,1,RW-1);
  bd(ws.getRange(R,R_+1)).setFormula('=IFERROR('+PROD_BUGS_FORMULA+',0)')
      .setBackground('#FFCDD2').setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
      .setFontColor('#B71C1C').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R,28);

  // Legend
  const prodNote = ws.getRange(R, L+2);
  prodNote.setValue('Target: 0 bugs di Production')
      .setBackground('#FFF3E0').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
  ws.getRange(R,L+2).setNote(
      'PROD BUGS = Bug yang Environment-nya Production dan belum Closed.\n' +
      'Status: Open, In Progress, Reopen, Fixed, Verified.\n' +
      'Target: 0 bugs di Production environment.'
  );
  R += 2;
  // ── END PROD BUGS ─────────────────────────────────────────────────

  // E. COVERAGE PER SUBMODUL (below Bug Summary)
  // =====================================================================
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#1565C0').setValue('E.  COVERAGE PER SUBMODUL  -  Web / Mobile');
  m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#283593').setValue('E.  COVERAGE PER SUBMODUL  -  API');
  ws.setRowHeight(R,20); R++;

  function buildCov(startRow,sc,master,subCol,prioCol,autoCol,execSh,hbg){
    ['SubModul','Total','Smoke','Regression','Auto%','Pass%'].forEach((h,i)=>
        h_(ws.getRange(startRow,sc+i),hbg).setValue(h).setFontSize(8).setWrap(true));
    ws.setRowHeight(startRow,22);
    const DS=startRow+1, MAX=34;
    for(let idx=0;idx<MAX;idx++){
      const row=DS+idx, bg=idx%2===0?(hbg==='#0D47A1'?'#F8F9FA':'#F0F4FF'):'#FFFFFF';
      const ref=colLetter(sc)+row;
      bd(ws.getRange(row,sc)).setFormula('=IFERROR(INDEX(UNIQUE(FILTER('+master+'!'+subCol+'3:'+subCol+'1000,'+master+'!'+subCol+'3:'+subCol+'1000<>"")),'+  (idx+1)+',1),"")').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left').setFontWeight('bold');
      bd(ws.getRange(row,sc+1)).setFormula('=IF('+ref+'="","",COUNTIF('+master+'!'+subCol+'3:'+subCol+'1000,'+ref+'))').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
      bd(ws.getRange(row,sc+2)).setFormula('=IF('+ref+'="","",COUNTIFS('+master+'!'+subCol+'3:'+subCol+'1000,'+ref+','+master+'!'+prioCol+'3:'+prioCol+'1000,"Critical")+COUNTIFS('+master+'!'+subCol+'3:'+subCol+'1000,'+ref+','+master+'!'+prioCol+'3:'+prioCol+'1000,"High")+COUNTIFS('+master+'!'+subCol+'3:'+subCol+'1000,'+ref+','+master+'!'+prioCol+'3:'+prioCol+'1000,"Medium"))').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
      bd(ws.getRange(row,sc+3)).setFormula('=IF('+ref+'="","",COUNTIFS('+master+'!'+subCol+'3:'+subCol+'1000,'+ref+','+master+'!'+prioCol+'3:'+prioCol+'1000,"Low")+COUNTIFS('+master+'!'+subCol+'3:'+subCol+'1000,'+ref+','+master+'!'+prioCol+'3:'+prioCol+'1000,"Lowest"))').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
      bd(ws.getRange(row,sc+4)).setFormula('=IFERROR(COUNTIFS('+master+'!'+subCol+'3:'+subCol+'1000,'+ref+','+master+'!'+autoCol+'3:'+autoCol+'1000,"Automated")/'+colLetter(sc+1)+row+',0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
      bd(ws.getRange(row,sc+5)).setFormula('=IFERROR(COUNTIFS('+execSh+'!B9:B1000,'+ref+','+execSh+'!Z9:Z1000,"PASSED")/MAX(1,COUNTIF('+execSh+'!B9:B1000,'+ref+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
      ws.setRowHeight(row,16);
    }
    passRateCF(ws.getRange(DS,sc+4,MAX,1));
    passRateCF(ws.getRange(DS,sc+5,MAX,1));
    const totBg=hbg==='#0D47A1'?'#E3F2FD':'#E8EAF6', totRow=DS+MAX;
    ['TOTAL','=SUM('+colLetter(sc+1)+DS+':'+colLetter(sc+1)+(totRow-1)+')',
      '=SUM('+colLetter(sc+2)+DS+':'+colLetter(sc+2)+(totRow-1)+')',
      '=SUM('+colLetter(sc+3)+DS+':'+colLetter(sc+3)+(totRow-1)+')',
      '=IFERROR(COUNTIF('+master+'!'+autoCol+'3:'+autoCol+'1000,"Automated")/MAX(1,COUNTA('+master+'!'+subCol+'3:'+subCol+'1000)),0)',
      '=IFERROR(COUNTIF('+execSh+'!Z9:Z1000,"PASSED")/MAX(1,COUNTA('+master+'!'+subCol+'3:'+subCol+'1000)),0)',
    ].forEach((v,i)=>{
      const c=bd(ws.getRange(totRow,sc+i)).setBackground(totBg).setFontWeight('bold')
          .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment(i===0?'left':'center');
      if(typeof v==='string'&&v.startsWith('=')) c.setFormula(v); else c.setValue(v);
      if(i>=4){ c.setNumberFormat('0%'); passRateCF(ws.getRange(totRow,sc+i)); }
    });
    ws.setRowHeight(totRow,18);
  }

  buildCov(R,L,'TC_Master','B','E','H','TC_Execution','#0D47A1');
  buildCov(R,R_,'API_Master','B','G','J','API_Execution','#283593');
  R+=36;


  addPeruriFooter(ws,R+1,21);
  ws.setFrozenRows(0);
}


function createPerfTest(ss) {
  const ws = safeSheet(ss,'PerfTest');
  ws.clear(); ws.setTabColor('#1A237E');

  function h_(range, bg, fg, sz) {
    return bd(range).setBackground(bg||'#37474F').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }
  function lbl(row, col, text) {
    bd(ws.getRange(row,col)).setValue(text).setBackground('#E3F2FD')
        .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
        .setHorizontalAlignment('right').setFontColor('#546E7A').setVerticalAlignment('middle');
  }
  function inp(row, col, ncols, val) {
    const r = ncols>1 ? ws.getRange(row,col,1,ncols).merge() : ws.getRange(row,col);
    bd(r).setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(9)
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    if(val!==undefined) r.setValue(val);
    return r;
  }
  function passfailCF(range){
    const rules=ws.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('PASS').setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('FAIL').setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true).setRanges([range]).build());
    ws.setConditionalFormatRules(rules);
  }

  // -- TITLE ---------------------------------------------------
  ws.getRange(1,1,1,15).merge();
  h_(ws.getRange(1,1,1,15),'#4A148C','#FFFFFF',11).setValue('PERFORMANCE TEST  .  Rekam & Evaluasi Hasil');
  ws.setRowHeight(1,30);
  ws.getRange(2,1,1,15).merge();
  ws.getRange(2,1)
      .setValue('Isi threshold (baris 11) sesuai SLA yang disepakati. Isi kolom Actual setelah menjalankan test. STATUS dihitung otomatis.')
      .setFontColor('#78909C').setFontStyle('italic').setFontSize(9).setFontFamily('Arial')
      .setBackground('#F8F9FA').setHorizontalAlignment('center');
  ws.setRowHeight(2,18);

  // -- SESSION INFO ---------------------------------------------
  ws.getRange(4,1,1,15).merge();
  h_(ws.getRange(4,1,1,15),'#6A1B9A').setValue('SESSION INFO');
  ws.setRowHeight(4,22);

  // Row 5-7: session fields in 2 columns
  const sessionL=[['Project / Sprint:',''],['Tester:',''],['Tanggal Test:','']];
  const sessionR=[['Tool:','K6'],['Environment:','Staging'],['Durasi Run (s):',300]];
  [5,6,7].forEach((r,i)=>{
    lbl(r,1,sessionL[i][0]); inp(r,2,4,sessionL[i][1]);
    lbl(r,8,sessionR[i][0]); inp(r,9,4,sessionR[i][1]);
    ws.setRowHeight(r,24);
  });

  // -- THRESHOLD ------------------------------------------------
  ws.getRange(9,1,1,15).merge();
  h_(ws.getRange(9,1,1,15),'#6A1B9A').setValue('THRESHOLD  .  Batas nilai yang harus dipenuhi (edit baris 11)');
  ws.setRowHeight(9,22);

  // Row 10: metric labels  |  Row 11: threshold values
  // Cols: A=RPS B=ErrRate C=P90 D=P95 E=P99 F=VU G=CPU H=Memory
  const metrics=[
    {col:1, label:'RPS\n(min req/s)', note:'Minimum requests per second yang harus dicapai', default:10,   color:'#512DA8'},
    {col:2, label:'Error Rate\n(max %)',note:'Persentase error maksimum yang diperbolehkan',   default:1,    color:'#512DA8'},
    {col:3, label:'P90\n(max ms)',    note:'Response time P90 tidak boleh melebihi nilai ini', default:500,  color:'#6A1B9A'},
    {col:4, label:'P95\n(max ms)',    note:'Response time P95',                                default:800,  color:'#6A1B9A'},
    {col:5, label:'P99\n(max ms)',    note:'Response time P99',                                default:1500, color:'#6A1B9A'},
    {col:6, label:'VU\n(target)',     note:'Jumlah Virtual User / concurrent user target',     default:100,  color:'#4A148C'},
    {col:7, label:'CPU\n(max %)',     note:'Maksimum CPU usage di server selama test',          default:70,   color:'#7B1FA2'},
    {col:8, label:'Memory\n(max %)', note:'Maksimum memory usage di server selama test',       default:80,   color:'#7B1FA2'},
  ];
  metrics.forEach(m=>{
    h_(ws.getRange(10,m.col),m.color).setValue(m.label).setNote(m.note).setWrap(true).setFontSize(8);
    bd(ws.getRange(11,m.col)).setValue(m.default)
        .setBackground('#F3E5F5').setFontFamily('Arial').setFontSize(12)
        .setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#4A148C').setVerticalAlignment('middle');
  });
  ws.setRowHeight(10,40); ws.setRowHeight(11,32);

  ws.getRange(12,1,1,15).merge();
  ws.getRange(12,1).setValue('? Edit nilai di baris 11 sesuai SLA yang disepakati.')
      .setFontColor('#78909C').setFontStyle('italic').setFontSize(8).setFontFamily('Arial')
      .setBackground('#E3F2FD');
  ws.setRowHeight(12,16);

  // -- RESULTS TABLE ---------------------------------------------
  ws.getRange(13,1,1,15).merge();
  h_(ws.getRange(13,1,1,15),'#6A1B9A').setValue('HASIL EKSEKUSI  .  Isi kolom Actual setelah menjalankan test');
  ws.setRowHeight(13,22);

  // Row 14: column group headers
  const groups=[
    {label:'SCENARIO / ENDPOINT',sc:1,ncols:2},
    {label:'CONFIG',sc:3,ncols:2},
    {label:'THROUGHPUT',sc:5,ncols:2},
    {label:'RESPONSE TIME',sc:7,ncols:3},
    {label:'RESOURCE',sc:10,ncols:2},
    {label:'RESULT',sc:12,ncols:2},
  ];
  groups.forEach(g=>{
    ws.getRange(14,g.sc,1,g.ncols).merge();
    h_(ws.getRange(14,g.sc,1,g.ncols),'#512DA8').setValue(g.label).setFontSize(8);
  });
  ws.setRowHeight(14,22);

  // Row 15: detail col headers
  [{col:1,h:'Scenario'},{col:2,h:'Endpoint'},{col:3,h:'VU'},{col:4,h:'Duration\n(s)'},
    {col:5,h:'RPS\nActual'},{col:6,h:'Error %\nActual'},
    {col:7,h:'P90 ms'},{col:8,h:'P95 ms'},{col:9,h:'P99 ms'},
    {col:10,h:'CPU %'},{col:11,h:'Memory %'},
    {col:12,h:'STATUS'},{col:13,h:'Notes / Observasi'},
    {col:14,h:'Screenshot 1\n(URL)'},{col:15,h:'Screenshot 2\n(URL)'},
  ].forEach(c=>h_(ws.getRange(15,c.col),'#7B1FA2').setValue(c.h).setFontSize(8).setWrap(true));
  ws.setRowHeight(15,38);

  // Notes on result columns
  ws.getRange(15,5).setNote('[INPUT] RPS Actual: requests per second yang berhasil.\nHarus >= threshold baris 11.');
  ws.getRange(15,6).setNote('[INPUT] Error Rate Actual: % request yang error.\nHarus <= threshold baris 11.');
  ws.getRange(15,7).setNote('[INPUT] P90: 90% request selesai dalam waktu ini (ms).\nHarus <= threshold baris 11.');
  ws.getRange(15,8).setNote('[INPUT] P95: 95% request selesai dalam waktu ini (ms).\nHarus <= threshold baris 11.');
  ws.getRange(15,9).setNote('[INPUT] P99: 99% request selesai dalam waktu ini (ms).\nHarus <= threshold baris 11.');
  ws.getRange(15,10).setNote('[INPUT] CPU Usage server selama test (%).\nHarus <= threshold baris 11.');
  ws.getRange(15,11).setNote('[INPUT] Memory Usage server selama test (%).\nHarus <= threshold baris 11.');
  ws.getRange(15,12).setNote('[AUTO] STATUS dihitung otomatis.\nPASS = semua metrik memenuhi threshold.\nFAIL = minimal 1 metrik tidak memenuhi threshold.\nKosong jika belum ada data.');
  ws.getRange(15,14).setNote('[INPUT] Link screenshot grafik K6/Gatling/JMeter.\nContoh: link Google Drive atau Confluence.');
  ws.getRange(15,15).setNote('[INPUT] Link screenshot tambahan (resource monitor, APM, dll).');
  ws.getRange(11,1).setNote('RPS threshold: MINIMUM. Jika actual < nilai ini -> FAIL.');
  ws.getRange(11,2).setNote('Error Rate threshold: MAXIMUM (%). Jika actual > nilai ini -> FAIL.');
  ws.getRange(11,3).setNote('P90 threshold: MAXIMUM (ms). Response time P90 tidak boleh melebihi ini.');
  ws.getRange(11,4).setNote('P95 threshold: MAXIMUM (ms).');
  ws.getRange(11,5).setNote('P99 threshold: MAXIMUM (ms).');
  ws.getRange(11,6).setNote('VU: Jumlah Virtual User yang disimulasikan. Info saja, tidak mempengaruhi PASS/FAIL.');
  ws.getRange(11,7).setNote('CPU threshold: MAXIMUM (%). Pantau resource server selama test.');
  ws.getRange(11,8).setNote('Memory threshold: MAXIMUM (%). Pantau resource server selama test.');
  // Row 16+: data rows
  // STATUS formula (col 12):
  //   PASS if ALL actual values meet threshold:
  //     RPS(col5) >= threshold A11  [higher is better]
  //     Error%(col6) <= threshold B11
  //     P90(col7) <= threshold C11
  //     P95(col8) <= threshold D11
  //     P99(col9) <= threshold E11
  //     CPU(col10) <= threshold G11
  //     Memory(col11) <= threshold H11
  //   Empty if no data in row

  const DS=16, MR=30;
  for(let r=DS;r<DS+MR;r++){
    const bg=((r-DS)%2===0)?'#F8F9FA':'#FFFFFF';
    // Data cols 1-11
    for(let c=1;c<=11;c++){
      bd(ws.getRange(r,c)).setBackground(bg).setFontFamily('Arial').setFontSize(9)
          .setHorizontalAlignment(c<=2?'left':'center').setVerticalAlignment('middle');
    }
    // STATUS col 12
    const f='=IF(AND(E'+r+'="",F'+r+'="",G'+r+'=""),"",'+
        'IF(AND('+
        'IF(E'+r+'<>"",E'+r+'>=$A$11,TRUE),'+      // RPS >= min
        'IF(F'+r+'<>"",F'+r+'<=$B$11,TRUE),'+      // Error% <= max
        'IF(G'+r+'<>"",G'+r+'<=$C$11,TRUE),'+      // P90 <= max
        'IF(H'+r+'<>"",H'+r+'<=$D$11,TRUE),'+      // P95 <= max
        'IF(I'+r+'<>"",I'+r+'<=$E$11,TRUE),'+      // P99 <= max
        'IF(J'+r+'<>"",J'+r+'<=$G$11,TRUE),'+      // CPU <= max
        'IF(K'+r+'<>"",K'+r+'<=$H$11,TRUE)'+       // Memory <= max
        '),"PASS","FAIL"))';
    bd(ws.getRange(r,12)).setFormula(f).setBackground(bg)
        .setFontFamily('Arial').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('center');
    passfailCF(ws.getRange(r,12));
    // Notes col 13 + Screenshot cols 14-15
    bd(ws.getRange(r,13)).setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
    bd(ws.getRange(r,14)).setBackground(bg).setFontColor('#1A73E8').setFontFamily('Arial').setFontSize(9)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
    bd(ws.getRange(r,15)).setBackground(bg).setFontColor('#1A73E8').setFontFamily('Arial').setFontSize(9)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
    ws.setRowHeight(r,22);
  }

  // Sample data row
  [
    ['Login API','/api/v1/auth/login',100,300, 45, 0.2, 320,480,750, 42,61,'','Baseline test, stable environment'],
  ].forEach((row,i)=>{
    row.forEach((val,c)=>{
      if(val!=='') ws.getRange(DS+i,c+1).setValue(val);
    });
  });

  // Summary row after data: overall PASS/FAIL
  const totRow=DS+MR;
  ws.getRange(totRow,1,1,11).merge(); // spans scenario+config cols
  bd(ws.getRange(totRow,1)).setValue('OVERALL RESULT')
      .setBackground('#E3F2FD').setFontWeight('bold').setFontFamily('Arial').setFontSize(10)
      .setHorizontalAlignment('right').setFontColor('#37474F');
  bd(ws.getRange(totRow,12))
      .setFormula('=IF(COUNTIF(L'+DS+':L'+(totRow-1)+',"FAIL")>0,"FAIL","PASS")')
      .setBackground('#E3F2FD').setFontWeight('bold').setFontFamily('Arial').setFontSize(12)
      .setHorizontalAlignment('center').setFontColor('#37474F');
  passfailCF(ws.getRange(totRow,12));
  bd(ws.getRange(totRow,13)).setBackground('#E3F2FD').setFontFamily('Arial').setFontSize(9)
      .setValue('PASS = semua skenario memenuhi threshold  .  Link ke Summary: lihat sel Perf Test Status di tab Summary');
  ws.setRowHeight(totRow,28);

  // Column widths
  [120,160,55,70, 65,75, 65,65,65, 65,70, 70,180, 160,160].forEach((w,i)=>ws.setColumnWidth(i+1,w));

  // Freeze row 15 (header)
  ws.setFrozenRows(15);

  // Guide note
  ws.getRange(totRow+2,1,1,15).merge();
  ws.getRange(totRow+2,1)
      .setValue('Panduan: Kolom STATUS otomatis PASS/FAIL berdasarkan threshold baris 11. '+
          'Edit threshold sesuai SLA. Tambah baris untuk skenario baru. '+
          'OVERALL RESULT di baris terakhir terhubung ke tab Summary (Perf Test Status).')
      .setFontColor('#78909C').setFontStyle('italic').setFontSize(8).setFontFamily('Arial')
      .setBackground('#E3F2FD').setWrap(true);
  ws.setRowHeight(totRow+2,30);
  addPeruriFooter(ws, totRow+4, 15);
}



// ===================================================================
//  TAB 6 -- Appendix
// ===================================================================
function createBugReport(ss) {
  const ws = safeSheet(ss, 'BugReport');
  ws.clear(); ws.setTabColor('#1565C0');

  function h_(r,c,nr,nc,txt,bg,fg,sz){
    const rng=nr>1||nc>1?ws.getRange(r,c,nr,nc).merge():ws.getRange(r,c);
    return rng.setValue(txt||'').setBackground(bg||'#0D47A1').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
  }
  function dv_(list){ return SpreadsheetApp.newDataValidation().requireValueInList(list,true).build(); }

  // Columns (19 total ? Severity removed):
  //  1=BugID  2=Type  3=Priority  4=Status  5=Feature  6=SubModul
  //  7=Title  8=Environment  9=Steps  10=Expected  11=Actual
  //  12=Related TC_ID  13=Reported By  14=Assigned To
  //  15=Date Found  16=Date Fixed  17=Sprint  18=Jira/Link  19=Notes
  [70,75,80,90,100,90, 220,200,90,160,140,160, 110,100,110, 90,90,80,130,150,160]
      .forEach((w,i)=>ws.setColumnWidth(i+1,w));

  // Row 1: title
  h_(1,1,1,21,'BUG REPORT  |  Web  ·  Mobile  ·  API','#0D47A1','#FFFFFF',12);
  ws.setRowHeight(1,30);

  // Row 2: note
  ws.getRange(2,1,1,22).merge()
      .setValue('Priority: Critical = showstopper · High = blocker · Medium = degraded (blocker) · Low = minor  |  Status: Open · In Progress · Fixed · Verified · In Progress VAPT · Done VAPT · Closed')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic')
      .setFontFamily('Arial').setFontSize(8).setHorizontalAlignment('left');
  ws.setRowHeight(2,16);

  // Row 3: group headers
  h_(3,1,1,4,'IDENTIFICATION','#0D47A1');
  h_(3,5,1,2,'CLASSIFICATION','#1565C0');
  h_(3,7,1,6,'DETAIL','#1976D2');
  h_(3,13,1,3,'OWNERSHIP','#1565C0');
  h_(3,16,1,2,'TIMELINE','#0D47A1');
  h_(3,18,1,4,'REFERENCE','#0D47A1');
  ws.setRowHeight(3,18);

  // Row 4: column headers
  ['Bug ID','Type','Priority','Status',
    'Feature','SubModul',
    'Title / Summary','Description',
    'Environment','Steps to Reproduce','Expected Result','Actual Result',
    'Related TC_ID','Reported By','Assigned To',
    'Date Found','Date Fixed',
    'Sprint','Jira / Link','Notes','Screenshot / Evidence'
  ].forEach((h,i) => h_(4,i+1,1,1,h,'#0D47A1'));
  ws.setRowHeight(4,22);
  ws.getRange(4,8).setNote(
      'Description\n\n'+
      'Deskripsi detail bug — sama seperti kolom Description di Jira.\n\n'+
      'Isi dengan konteks tambahan yang tidak muat di Title, misalnya:\n'+
      '  • Kondisi awal saat bug terjadi\n'+
      '  • User role atau permission yang terdampak\n'+
      '  • Data spesifik yang digunakan\n'+
      '  • Frekuensi kemunculan (selalu / kadang / sekali)\n\n'+
      'Boleh dikosongkan jika Title sudah cukup jelas.'
  );

  // Notes
  ws.getRange(4,1).setNote(
      'Bug ID Format: BUG-[TYPE]-[000]\n\n'+
      'BUG-WEB-001  = Web/UI bug ke-1\n'+
      'BUG-MOB-001  = Mobile bug ke-1\n'+
      'BUG-API-001  = API bug ke-1\n\n'+
      'Nomor urut 3 digit, mulai 001.\nJangan reuse ID yang sudah ada.'
  );
  ws.getRange(4,3).setNote(
      'Priority Level:\n\n'+
      'Critical  = Showstopper. Fitur/app tidak bisa digunakan sama sekali. Release DITAHAN.\n'+
      'High      = Blocker. Fitur utama terdampak signifikan.\n'+
      'Medium    = Blocker (degraded). Ada workaround tapi experience buruk.\n'+
      '            --> Medium, High, Critical semua dihitung sebagai blocker di dashboard.\n'+
      'Low       = Minor. Kosmetik atau edge case kecil, tidak blokir release.'
  );
  ws.getRange(4,4).setNote(
      'Status Flow (with VAPT Integration):\n\n'+
      'QA Phase:\n'+
      '  Open → In Progress → Fixed → Verified\n\n'+
      'VAPT Phase (Security Testing):\n'+
      '  Verified → In Progress VAPT → Done VAPT → Closed\n\n'+
      'Exception Flows:\n'+
      '  • Any status → Reopen (bug reappears after fix)\n'+
      '  • Any status → Won\'t Fix (rejected with reason)\n'+
      '  • Verified can skip directly to Closed (no VAPT needed)\n'+
      '  • Done VAPT can return to In Progress VAPT if issues found\n\n'+
      '🚨 BLOCKER STATUS:\n'+
      'Open, In Progress, Reopen, In Progress VAPT, Done VAPT\n'+
      '(with Priority Critical/High/Medium)\n\n'+
      'NOT Blocker: Verified, Closed, Won\'t Fix'
  );

  // Data rows
  const DS = 5, MR = 200;

  ws.getRange(DS,2,MR,1).setDataValidation(dv_(['Web','Mobile','API']));
  ws.getRange(DS,3,MR,1).setDataValidation(dv_(['Critical','High','Medium','Low']));
  ws.getRange(DS,4,MR,1).setDataValidation(dv_(['Open','In Progress','Fixed','Verified','In Progress VAPT','Done VAPT','Closed',"Won't Fix",'Reopen']));
  ws.getRange(DS,9,MR,1).setDataValidation(dv_(['Dev','Staging / UAT','Production','All']));

  // Alternating row bg
  for (let r=DS; r<DS+MR; r++) {
    const bg = (r-DS)%2===0 ? '#F8FBFF' : '#FFFFFF';
    ws.getRange(r,1,1,21).setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(r,22);
  }
  ws.getRange(DS,7,MR,1).setWrap(true);  // Title
  ws.getRange(DS,8,MR,1).setWrap(true);  // Description
  ws.getRange(DS,10,MR,1).setWrap(true); // Steps
  ws.getRange(DS,11,MR,1).setWrap(true); // Expected
  ws.getRange(DS,12,MR,1).setWrap(true); // Actual
  ws.getRange(DS,16,MR,1).setNumberFormat('yyyy-mm-dd');
  ws.getRange(DS,17,MR,1).setNumberFormat('yyyy-mm-dd');
  ws.getRange(DS,21,MR,1).setWrap(false).setFontColor('#1A73E8'); // Screenshot link

  // Conditional Formatting
  const rules = ws.getConditionalFormatRules();
  const prioRange = ws.getRange(DS,3,MR,1);
  const statRange = ws.getRange(DS,4,MR,1);
  const typeRange = ws.getRange(DS,2,MR,1);

  // Priority CF (red-blue theme: critical=deep red, high=orange-red, medium=amber, low=light blue)
  [{v:'Critical',bg:'#FFCDD2',fg:'#B71C1C'},
    {v:'High',    bg:'#FFE0B2',fg:'#E65100'},
    {v:'Medium',  bg:'#E3F2FD',fg:'#1565C0'},
    {v:'Low',     bg:'#F1F8E9',fg:'#388E3C'}
  ].forEach(s=>rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(s.v).setBackground(s.bg).setFontColor(s.fg).setBold(true)
      .setRanges([prioRange]).build()));

  // Status CF
  [{v:'Open',             bg:'#FFCDD2',fg:'#C62828'},
    {v:'In Progress',      bg:'#E3F2FD',fg:'#1565C0'},
    {v:'Fixed',            bg:'#FFF9C4',fg:'#F57F17'},
    {v:'Verified',         bg:'#C8E6C9',fg:'#2E7D32'},
    {v:'In Progress VAPT', bg:'#E1F5FE',fg:'#01579B'},
    {v:'Done VAPT',        bg:'#B2DFDB',fg:'#004D40'},
    {v:'Closed',           bg:'#E8F5E9',fg:'#388E3C'},
    {v:"Won't Fix",        bg:'#F5F5F5',fg:'#9E9E9E'},
    {v:'Reopen',           bg:'#EDE7F6',fg:'#6A1B9A'}
  ].forEach(s=>rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(s.v).setBackground(s.bg).setFontColor(s.fg).setBold(true)
      .setRanges([statRange]).build()));

  // Type CF
  [{v:'Web',    bg:'#E3F2FD',fg:'#1565C0'},
    {v:'Mobile', bg:'#E8F5E9',fg:'#2E7D32'},
    {v:'API',    bg:'#EDE7F6',fg:'#6A1B9A'}
  ].forEach(s=>rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(s.v).setBackground(s.bg).setFontColor(s.fg).setBold(true)
      .setRanges([typeRange]).build()));

  // Sample row
  const sampleRow = [
    'BUG-WEB-001','Web','High','Open','Login','1.1',
    'Tombol Login tidak aktif setelah input password yang benar',
    'Tombol Login disabled meski email & password valid. Terjadi sejak build v1.2.3. Hanya di browser berbasis Chromium.',
    'Staging / UAT',
    '1. Buka halaman Login\n2. Input email & password yang valid\n3. Klik tombol Login',
    'User berhasil login dan diarahkan ke Dashboard',
    'Tombol Login tetap disabled / tidak bisa diklik setelah input benar',
    'WEB.LOG.001','Tester A','Dev B',
    new Date(),'',
    'Sprint 12','','Reproduced di Chrome 121 & Firefox 122. Tidak terjadi di Safari.',
    ''
  ];
  sampleRow.forEach((v,ci) => {
    const c = ws.getRange(DS,ci+1);
    if (ci===15||ci===16) { if(v) c.setValue(v).setNumberFormat('yyyy-mm-dd'); }
    else c.setValue(v||'');
    c.setBackground('#F8FBFF').setFontFamily('Arial').setFontSize(9)
        .setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
  });
  ws.setRowHeight(DS, 60); // taller for wrapped text

  ws.getRange(4,21).setNote('Paste link screenshot/evidence.\nContoh: URL Google Drive, Jira attachment, atau direct image URL.\nPastikan link bisa diakses oleh reviewer.');
  ws.setConditionalFormatRules(rules);
  ws.setFrozenRows(4);
  Logger.log('BugReport created OK');
}


function createAppendix(ss) {
  const ws = safeSheet(ss,'Appendix');
  ws.clear(); ws.setTabColor('#1A237E');

  ws.getRange(1,1,1,4).merge();
  hdr(ws.getRange(1,1,1,4),'#0D47A1','#FFFFFF',11)
      .setValue('APPENDIX  .  Definisi, Konvensi & Panduan  (v39)');
  ws.setRowHeight(1,30);

  let r=2;
  function sec(title, bg){
    ws.getRange(r,1,1,4).merge();
    hdr(ws.getRange(r,1,1,4),bg||'#37474F','#FFFFFF',9).setValue(title);
    ws.setRowHeight(r,24); r++;
  }
  function row2(label,desc,bg){
    bd(ws.getRange(r,1)).setValue(label)
        .setBackground(bg||'#ECEFF1').setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(desc).setBackground('#FFFFFF')
        .setFontFamily('Arial').setFontSize(9).setWrap(true)
        .setHorizontalAlignment('left').setVerticalAlignment('top');
    ws.setRowHeight(r,48); r++;
  }

  sec('0. HIERARKI QA -- PROJECT / MODULE / SUBMODULE','#0D47A1');
  row2('Definisi',
      'Project   = Inisiatif / client / program kerja. Contoh: SIPGN, INAGOV\n'+
      'Module    = Pengelompokan domain fungsional dalam project.\n'+
      '            Kosongkan ("-") jika project tidak punya layer domain (project flat).\n'+
      'SubModule = Unit terkecil yang berdiri sendiri -- 1 aplikasi atau 1 domain.\n'+
      '            Ini adalah ANCHOR utama untuk TC_ID, Coverage, dan Dashboard.\n'+
      'Feature   = Fitur besar dalam SubModule. Dibedakan di kolom Feature, bukan TC_ID.'
  );
  row2('Pola A -- Project Berlayer (SIPGN)',
      'Project  : SIPGN\n'+
      '  Module 1  : Manajemen Gizi\n'+
      '    SubModule 1.1 : Aplikasi Nutritionist\n'+
      '      Feature: Meal Plan, Menu Management\n'+
      '    SubModule 1.2 : Aplikasi Courier\n'+
      '      Feature: Pick Up, Delivery, Return\n'+
      '    SubModule 1.3 : Aplikasi Beneficiary\n'+
      '  Module 2  : Manajemen Distribusi\n'+
      '    SubModule 2.1 : ...'
  );
  row2('Pola B -- Project Flat (INAGOV)',
      'Project   : INAGOV\n'+
      '  Module  : - (kosong)\n'+
      '    SubModule : Talenta\n'+
      '      Feature: Rekrutmen, Penggajian\n'+
      '    SubModule : e-Office\n'+
      '    SubModule : SIMPEG\n'+
      '\n'+
      'Pada pola flat, SubModule setara dengan Module di pola berlayer.\n'+
      'Kolom Module di Summary dan Config dikosongkan.'
  );
  row2('TC_ID per SubModule',
      'Format  : [SubModule].[3-digit]\n'+
      'Pola A  : 1.1.001 (SubModule 1.1, TC ke-1)\n'+
      '          1.2.001 (SubModule 1.2, TC ke-1)\n'+
      'Pola B  : Talenta.001 atau tetap numerik 1.001\n'+
      '\n'+
      'API prefix wajib: API.1.1.001 / API.1.2.001'
  );

  sec('1. STRUKTUR TAB','#0D47A1');
  row2('TC_Master',     'Master list test case Web / Mobile.\nKolom [INPUT]: SubModul, TC_ID, Feature, Priority, Platform, Test Type, Automation, Version, Role (RBAC), Scenario, Steps, Expected Result.\nKolom [AUTO]: Test Level (kolom N) -- jangan diedit.\nFormat TC_ID: [SubModul].[3-digit]  contoh: 1.1.001  2.3.015\nRole = peran RBAC yang menjalankan skenario, contoh: Admin, User, Viewer.');
  row2('TC_Execution',  'Kolom identitas sync otomatis dari TC_Master. Isi kolom staging dengan PASSED / FAILED / BLOCKED / TODO.\nTambah kolom ke kanan untuk setiap run. LATEST STATUS di kolom Z otomatis.\nKolom AA = link screenshot / evidence.\nIN PROGRESS otomatis jika ada PASSED dan TODO di skenario yang sama.');
  row2('API_Master',    'Master list test case API. Method (E) dan Endpoint URL (F) terpisah.\nKolom [INPUT]: SubModul, TC_ID, Feature, Method, Endpoint, Priority, Auth, Test Type, Automation, Version, Role (RBAC), Scenario.\nKolom [AUTO]: Test Level (kolom N) -- jangan diedit.\n\nFormat TC_ID: API.[SVC].[FEAT].[000]\n  [SVC]  = Kode service/domain, maks 3-4 huruf kapital. Contoh: AUTH, USER, ORD, PAY, INV\n  [FEAT] = Kode endpoint/fitur, maks 3-4 huruf kapital. Contoh: LOG, LIST, CRT, UPD, DEL\n  [000]  = Nomor urut 3 digit, mulai 001\n\nContoh: API.AUTH.LOG.001 (Auth, Login, TC-1)  API.USER.CRT.002 (User, Create, TC-2)  API.PAY.CHK.005 (Payment, Checkout, TC-5)\n\nAturan:\n- Harus UNIK -- jangan pernah reuse TC_ID yang sudah ada\n- Jangan ubah TC_ID jika sudah ada hasil di Execution\n- Urutan: Positive dulu (001), baru Negative (002), Edge Case (003)\n\nRole = peran RBAC yang diuji aksesnya, contoh: Admin, Super Admin, User, Viewer.');
  row2('API_Execution', 'Sama seperti TC_Execution namun untuk API. Sync otomatis dari API_Master.\nKolom AA = link screenshot / evidence.');
  row2('Summary',       'Isi bagian Test Plan sebelum memulai eksekusi (Project, Modul, PIC, Jira, status).\nRingkasan otomatis: coverage SubModul & Feature, run history dengan IN PROGRESS.\nPerf Test Status otomatis dari tab PerfTest.');
  row2('PerfTest',      'Rekam hasil performance test. Isi threshold di baris 11 sesuai SLA.\nKolom STATUS otomatis PASS/FAIL per skenario. OVERALL RESULT terhubung ke tab Summary.\nMetrik yang dicek: RPS, Error Rate, P90, P95, P99, VU, CPU, Memory.');
  row2('Appendix',      'Dokumen ini.');
  r++;

  sec('2. STATUS EKSEKUSI','#0D47A1');
  [['PASSED',      'Skenario berhasil -- hasil sesuai Expected Result.','#C8E6C9'],
    ['IN PROGRESS', 'Ada skenario yang sudah PASSED di run sebelumnya namun masih ada TODO -- eksekusi belum selesai.','#E3F2FD'],
    ['FAILED',      'Skenario gagal -- hasil tidak sesuai Expected Result. Wajib buat bug report.','#FFCDD2'],
    ['BLOCKED',     'Tidak bisa dieksekusi -- environment down, dependensi belum siap, atau data belum ada.','#FFE0B2'],
    ['TODO',        'Belum dieksekusi pada run ini.','#F5F5F5'],
  ].forEach(([s,d,bg])=>{
    bd(ws.getRange(r,1)).setValue(s).setBackground(bg).setFontWeight('bold')
        .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(d).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
    ws.setRowHeight(r,34); r++;
  });
  r++;

  sec('3. TEST LEVEL  .  OTOMATIS DARI PRIORITY','#546E7A');
  row2('Smoke',      'Priority Critical / High / Medium ? Smoke.\nTest subset cepat untuk memvalidasi fungsi utama sebelum release atau setelah deployment.','#FFF8F0');
  row2('Regression', 'Priority Low / Lowest ? Regression.\nFull test cycle dijalankan sebelum release ke Production.','#F1F8E9');
  r++;

  sec('4. PRIORITY','#0D47A1');
  [['Critical  [BLOCKER]','Fungsi utama tidak bisa digunakan. Release DITAHAN jika FAIL/BLOCKED.','#FFEBEE'],
    ['High      [BLOCKER]','Fungsi penting terganggu. Jika FAIL: perlu approval PM untuk release.','#FFF3E0'],
    ['Medium    [POTENTIAL BLOCKER]','Fungsi terganggu sebagian. Jika FAIL sebelum UAT: flagged ke tech lead.','#FFF8E1'],
    ['Low',     'Masalah minor: UI, typo, UX kurang optimal.','#F1F8E9'],
    ['Lowest',  'Nice to have. Fix opsional.','#ECEFF1'],
  ].forEach(([p,d,bg])=>{
    bd(ws.getRange(r,1)).setValue(p).setBackground(bg).setFontWeight('bold')
        .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(d).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');
    ws.setRowHeight(r,34); r++;
  });
  r++;

  r++;
  sec('5. ROLE (RBAC) -- KONTROL AKSES','#0D47A1');
  row2('Tujuan','Kolom Role di TC_Master & API_Master bukan role developer, melainkan PERAN PENGGUNA (RBAC) yang menjalankan skenario tersebut.\n\nDigunakan untuk memverifikasi bahwa access control berjalan benar.');
  row2('Contoh Role','Admin: bisa create/edit/delete semua data\nUser/Operator: bisa create & edit data sendiri\nViewer: hanya bisa read/view\nGuest: akses terbatas (belum login)\nSupervisor: approval flow\nSuper Admin: full access termasuk system settings');
  row2('Skenario RBAC','Untuk setiap endpoint/fitur sensitif, buat TC untuk:\n1. Role yang BERHAK -- harus dapat akses (expected: 200/201)\n2. Role yang TIDAK berhak -- harus ditolak (expected: 403 Forbidden)\n3. Tanpa token/login -- harus ditolak (expected: 401 Unauthorized)');
  r++;
  sec('6. AUTOMATION STATUS','#0D47A1');
  [['Automated',          'Script sudah dibuat dan dapat dijalankan via CI/CD atau manual run.'],
    ['Manual',             'Tidak akan diautomasi -- memerlukan penilaian manusia (exploratory, visual, UX).'],
    ['To Do',              'Direncanakan untuk diautomasi, belum dikerjakan.'],
    ['Cannot be Automated','Secara teknis tidak bisa diautomasi (scan QR fisik, biometrik, hardware-dependent).'],
  ].forEach(([s,d])=>row2(s,d));
  r++;

  sec('7. BUG REPORT  —  STATUS & ALUR KERJA (WITH VAPT)','#B71C1C');
  // Status flow table with VAPT
  [['Open',              '#FFCDD2','#B71C1C', 'QA',       'Bug baru ditemukan. Belum ada yang mengerjakan. Masuk antrian Dev.'],
    ['In Progress',      '#E3F2FD','#1565C0', 'Dev',      'Dev sedang mengerjakan fix. Bug belum bisa di-retest.'],
    ['Fixed',            '#FFF9C4','#E65100', 'Dev',      'Dev klaim sudah diperbaiki. Menunggu QA untuk verifikasi ulang.'],
    ['Verified',         '#C8E6C9','#2E7D32', 'QA',       'QA sudah re-test dan bug confirmed fixed. Ready untuk VAPT (Security testing) atau langsung Closed jika skip VAPT.'],
    ['In Progress VAPT', '#E1F5FE','#01579B', 'Security', '🔒 Security team sedang melakukan VAPT testing.'],
    ['Done VAPT',        '#B2DFDB','#004D40', 'Security', '🔒 VAPT testing selesai. Perlu re-test QA sebelum Closed.'],
    ['Closed',           '#E8F5E9','#388E3C', 'QA/Lead',  '✅ Final. Bug sudah selesai dan di-release ke production.'],
    ["Won\'t Fix",       '#F5F5F5','#9E9E9E', 'Lead',     '❌ Tidak diperbaiki (alasan bisnis/teknis). Harus ada komentar.'],
    ['Reopen',           '#EDE7F6','#6A1B9A', 'QA',       '🔄 Bug masih muncul setelah Fixed/Done VAPT. Kembali ke Dev.'],
  ].forEach(function(s) {
    var bg=s[1], fg=s[2], who=s[3], desc=s[4];
    ws.getRange(r,1,1,1).setValue(s[0]).setBackground(bg).setFontColor(fg)
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r,2,1,1).setValue(who).setBackground('#F5F5F5').setFontColor('#424242')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.getRange(r,3,1,2).merge().setValue(desc).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(r,30); r++;
  });
  // Flow legend with VAPT
  ws.getRange(r,1,1,4).merge()
      .setValue('Flow:  Open → In Progress → Fixed → Verified → In Progress VAPT → Done VAPT → Closed  |  Reopen dari any status')
      .setBackground('#FFEBEE').setFontColor('#C62828').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#EF9A9A',SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(r,16); r++;
  // Siapa update apa - with VAPT roles
  row2('Dev update ke',      'In Progress (saat mulai mengerjakan)\nFixed (saat selesai — Dev TIDAK boleh langsung ke Verified/Closed)');
  row2('QA update ke',       'Verified (jika re-test lulus dan siap VAPT)\nReopen (jika bug masih ada)\nClosed (final setelah Done VAPT dan re-test QA)');
  row2('Security update ke', 'In Progress VAPT (saat mulai security testing)\nDone VAPT (testing selesai, siap re-test QA)\nReopen (jika ditemukan issue baru saat VAPT)');
  row2('Lead update ke',     "Won\'t Fix (dengan komentar alasan yang jelas)\nClosed (keputusan akhir)");
  row2('🚨 Open Blocker Calculation (UPDATED WITH VAPT)',
      'Formula Open Blocker di Summary & Dashboard menghitung bug dengan:\n' +
      '  • Status = Open / In Progress / Reopen / In Progress VAPT / Done VAPT\n' +
      '  • Priority = Critical / High / Medium\n\n' +
      'NOT Blocker: Verified, Closed, Won\'t Fix\n\n' +
      '💡 Kenapa Done VAPT masih blocker?\n' +
      'Karena bug perlu di-test ulang oleh QA sebelum bisa Closed.\n' +
      'Hanya setelah Closed baru tidak dihitung blocker.\n\n' +
      'Target: 0 Open Blocker sebelum release ke production.');
  r++;

  sec('9. PERFORMANCE TEST -- METRIK','#4A148C');
  [['RPS (req/s)',    'Requests Per Second -- jumlah request per detik yang berhasil diproses.\nThreshold: MINIMUM. Jika actual < threshold ? FAIL.'],
    ['Error Rate (%)', 'Persentase request yang mengembalikan error (status ? 400 atau timeout).\nThreshold: MAKSIMUM. Jika actual > threshold ? FAIL.'],
    ['P90 / P95 / P99','Percentile response time -- 90% / 95% / 99% dari semua request selesai dalam waktu ini.\nThreshold: MAKSIMUM (ms). Semakin kecil semakin baik.'],
    ['VU (Virtual User)','Jumlah concurrent user yang disimulasikan dalam satu skenario test.\nDiisi sebagai config, bukan sebagai threshold PASS/FAIL.'],
    ['CPU Usage (%)',   'Persentase penggunaan CPU server selama test berlangsung.\nThreshold: MAKSIMUM. Jika actual > threshold ? FAIL.'],
    ['Memory Usage (%)','Persentase penggunaan memory server selama test berlangsung.\nThreshold: MAKSIMUM. Jika actual > threshold ? FAIL.'],
  ].forEach(([m,d])=>row2(m,d));
  r++;

  sec('10. HTTP METHOD','#0D47A1');
  [['GET','Mengambil data. Tidak mengubah state.','#E8F0FE'],
    ['POST','Membuat resource baru.','#E8F5E9'],
    ['PUT','Update resource secara penuh (replace).','#FFF8E1'],
    ['PATCH','Update resource sebagian (partial).','#F3E5F5'],
    ['DELETE','Menghapus resource.','#FCE4EC'],
  ].forEach(([m,d,bg])=>{
    bd(ws.getRange(r,1)).setValue(m).setBackground(bg).setFontWeight('bold')
        .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(r,2,1,3).merge();
    bd(ws.getRange(r,2)).setValue(d).setBackground('#FFFFFF').setFontFamily('Arial')
        .setFontSize(9).setHorizontalAlignment('left').setVerticalAlignment('middle');
    ws.setRowHeight(r,28); r++;
  });

  [100,140,80,200].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  addPeruriFooter(ws, r+2, 4);
}

// ═══════════════════════════════════════════════════════════════════════
// MANUAL JIRA SYNC
// ═══════════════════════════════════════════════════════════════════════

/**
 * Create custom menu on open
 * Adds "Jira Sync" menu item for manual sync
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // Check if Config tab exists (which means Jira Sync is configured)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('Config');

  if (config) {
    ui.createMenu('🔄 Jira Sync')
      .addItem('✅ Sync Now', 'manualSyncJiraFromQATM')
      .addSeparator()
      .addItem('📋 View Config', 'openJiraConfig')
      .addToUi();
  }
}

/**
 * Open Config tab
 */
function openJiraConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('Config');

  if (config) {
    config.activate();
    SpreadsheetApp.getUi().alert('Config tab opened.\n\nYou can view and edit Jira configuration here.');
  } else {
    SpreadsheetApp.getUi().alert('❌ Config tab not found!\n\nPlease run broadcast from Dashboard to create Config tab.');
  }
}

/**
 * Manual Sync Function - Syncs bugs from Jira to BugReport sheet
 *
 * Expected Config tab structure:
 * Row 3: Labels
 * Row 4: Values
 * B4 = Jira Instance (digitalperuri / bgn-peruri)
 * C4 = Jira Project Key (SQA / BGN / etc)
 * D4 = Jira Email
 * E4 = Jira API Token
 * F4 = Module Name/Number
 */
function manualSyncJiraFromQATM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // Get QATM Config
    const config = ss.getSheetByName('Config');
    if (!config) {
      ui.alert('❌ Config tab not found!\n\nPlease create Config tab first via broadcast from Dashboard.');
      return;
    }

    // Read Jira config from QATM Config tab
    // Row 3 = Labels, Row 4 = Values
    const jiraInstance = String(config.getRange('B4').getValue()).trim().toLowerCase();
    const jiraProjectKey = String(config.getRange('C4').getValue()).trim().toUpperCase();
    const jiraEmail = String(config.getRange('D4').getValue()).trim();
    const jiraApiToken = String(config.getRange('E4').getValue()).trim();
    const moduleName = String(config.getRange('F4').getValue()).trim();

    // Validate config
    if (!jiraInstance || !jiraProjectKey || !jiraEmail || !jiraApiToken || !moduleName) {
      let msg = '❌ Jira configuration incomplete!\n\nPlease fill in Config tab (row 4):\n\n';
      if (!jiraInstance) msg += '• B4: Jira Instance (digitalperuri / bgn-peruri)\n';
      if (!jiraProjectKey) msg += '• C4: Jira Project Key (SQA / BGN / etc)\n';
      if (!jiraEmail) msg += '• D4: Jira Email (your.email@company.com)\n';
      if (!jiraApiToken) msg += '• E4: Jira API Token (ATATT3xFf...)\n';
      if (!moduleName) msg += '• F4: Module Name/Number (1 / Portal+SSO / etc)\n';

      ui.alert(msg);
      return;
    }

    ui.alert(
      '🔄 Manual Jira Sync',
      'Starting Jira sync for this QATM...\n\n' +
      'Instance: ' + jiraInstance + '\n' +
      'Project: ' + jiraProjectKey + '\n' +
      'Module: ' + moduleName + '\n\n' +
      'This may take a few minutes.',
      ui.ButtonSet.OK
    );

    // Call the actual sync function
    const result = syncJiraForCurrentQATM_(jiraInstance, jiraProjectKey, jiraApiToken, jiraEmail, moduleName);

    ui.alert(
      '✅ Sync Complete!',
      result,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert('❌ Error', 'Sync failed:\n\n' + e.message + '\n\nCheck Execution log for details.', ui.ButtonSet.OK);
    Logger.log('❌ Manual sync error: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
  }
}

/**
 * Core sync logic - Fetches bugs from Jira and updates BugReport sheet
 */
function syncJiraForCurrentQATM_(jiraInstance, jiraProjectKey, jiraApiToken, email, moduleName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bugSheet = ss.getSheetByName('BugReport');

  if (!bugSheet) {
    Logger.log('❌ BugReport sheet not found');
    return '❌ BugReport sheet not found in this spreadsheet.';
  }

  Logger.log('══════════════════════════════════════════');
  Logger.log('🔄 MANUAL JIRA SYNC');
  Logger.log('══════════════════════════════════════════');
  Logger.log('Instance: ' + jiraInstance);
  Logger.log('Project:  ' + jiraProjectKey);
  Logger.log('Modul:    ' + (moduleName || '(not specified)'));
  Logger.log('');

  try {
    // Fetch issues from Jira
    const issues = fetchJiraIssues_(jiraInstance, jiraProjectKey, moduleName, email, jiraApiToken);

    if (!issues) {
      Logger.log('❌ Failed to fetch issues from Jira');
      return '❌ Failed to fetch issues from Jira.\n\nCheck:\n• Jira Instance URL\n• Project Key\n• API Token\n• Module name';
    }

    Logger.log('✅ Fetched ' + issues.length + ' issue(s) from Jira');

    // Build index of existing bugs in sheet
    const bugIndex = buildBugIndex_(bugSheet);
    Logger.log('📊 Found ' + Object.keys(bugIndex).length + ' existing bug(s) in sheet');

    // Sync bugs
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const instUrl = getJiraInstanceUrl_(jiraInstance);
    let inserted = 0;
    let updated = 0;

    const jiraKeys = new Set();

    issues.forEach(function(issue) {
      const statusName = issue.fields.status && issue.fields.status.name;
      const isClosedStatus = statusName === 'Closed' || statusName === "Won't Fix";

      jiraKeys.add(issue.key);

      // Skip closed bugs - they'll be cleaned up below
      if (isClosedStatus) {
        return;
      }

      // Update or insert active bugs
      if (bugIndex[issue.key] !== undefined) {
        updateBugRow_(bugSheet, bugIndex[issue.key], issue, now, instUrl, jiraInstance);
        updated++;
      } else {
        insertBugRow_(bugSheet, issue, now, instUrl, jiraInstance, moduleName);
        inserted++;
      }
    });

    // Cleanup closed bugs
    const deleted = cleanupClosedBugs_(bugSheet, issues);

    Logger.log('✅ Sync complete:');
    Logger.log('  • Inserted: ' + inserted);
    Logger.log('  • Updated: ' + updated);
    Logger.log('  • Deleted (Closed/Won\'t Fix): ' + deleted);
    Logger.log('══════════════════════════════════════════');

    let result = '📊 Sync Results:\n\n';
    result += '• Inserted: ' + inserted + ' new bug(s)\n';
    result += '• Updated: ' + updated + ' bug(s)\n';
    if (deleted > 0) {
      result += '• Deleted: ' + deleted + ' (Closed/Won\'t Fix)\n';
    }
    result += '\n✅ Total: ' + (inserted + updated) + ' active bug(s)';

    return result;

  } catch (e) {
    Logger.log('❌ Error during sync: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    return '❌ Error during sync:\n\n' + e.message;
  }
}

// ═══════════════════════════════════════════════════════════════════════
// JIRA API & HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

function fetchJiraIssues_(instance, projectKey, moduleName, email, apiToken) {
  const baseUrl = getJiraInstanceUrl_(instance);
  const modulField = getJiraModulField_(instance);

  const isNumeric = /^\d+$/.test(moduleName);
  const modulValue = isNumeric ? moduleName : '"' + moduleName + '"';
  const jql = 'project = "' + projectKey + '" AND issuetype = Bug AND "' + modulField + '" = ' + modulValue + ' ORDER BY priority ASC, updated DESC';

  Logger.log('JQL: ' + jql);
  Logger.log('');

  const auth = Utilities.base64Encode(email + ':' + apiToken);
  const headers = {'Authorization': 'Basic ' + auth, 'Content-Type': 'application/json'};
  const fields = 'summary,description,priority,status,assignee,reporter,resolutiondate,key,created,updated,labels,components,environment';
  const customFields = getJiraCustomFields_(instance);
  const allFields = fields + customFields;
  const allIssues = [];
  let nextPageToken = null;

  do {
    let url = baseUrl + '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) + '&fields=' + encodeURIComponent(allFields) + '&maxResults=100';
    if (nextPageToken) url += '&nextPageToken=' + encodeURIComponent(nextPageToken);

    let response;
    try {
      response = UrlFetchApp.fetch(url, {headers: headers, muteHttpExceptions: true});
    } catch (e) {
      Logger.log('❌ Fetch error: ' + e.message);
      return null;
    }

    if (response.getResponseCode() !== 200) {
      Logger.log('❌ Jira API error ' + response.getResponseCode() + ': ' + response.getContentText().substring(0, 200));
      return null;
    }

    let data;
    try {
      data = JSON.parse(response.getContentText());
    } catch (e) {
      Logger.log('❌ JSON parse error: ' + e.message);
      return null;
    }

    (data.issues || []).forEach(function(issue) { allIssues.push(issue); });
    nextPageToken = data.nextPageToken || null;
    if (nextPageToken) Utilities.sleep(300);
  } while (nextPageToken);

  return allIssues;
}

function buildBugIndex_(bugSheet) {
  const data = bugSheet.getDataRange().getValues();
  const index = {};
  for (let i = 5; i < data.length; i++) {
    const jiraKey = String(data[i][19]).trim();
    if (jiraKey && jiraKey !== '') index[jiraKey] = i + 1;
  }
  return index;
}

function updateBugRow_(bugSheet, rowNum, issue, timestamp, instUrl, instance) {
  const row = [];
  row[1] = 'Functional';
  row[2] = normalizePriority_(issue.fields.priority && issue.fields.priority.name);
  row[3] = normalizeStatus_(issue.fields.status && issue.fields.status.name);
  const fieldMap = getJiraFieldMap_(instance);
  row[4] = getCustomFieldValue_(issue, fieldMap.feature);
  row[5] = getCustomFieldValue_(issue, fieldMap.submodul);
  row[6] = issue.fields.summary || '';
  row[7] = convertADF_(issue.fields.description);
  row[8] = getCustomFieldValue_(issue, fieldMap.environment) || issue.fields.environment || '';
  row[9] = getCustomFieldValue_(issue, fieldMap.steps);
  row[10] = getCustomFieldValue_(issue, fieldMap.expected);
  row[11] = getCustomFieldValue_(issue, fieldMap.actual);
  row[13] = issue.fields.reporter && issue.fields.reporter.displayName || '';
  row[14] = issue.fields.assignee && issue.fields.assignee.displayName || '';
  row[15] = issue.fields.created ? issue.fields.created.substring(0, 10) : '';
  row[18] = instUrl + '/browse/' + issue.key;
  row[19] = issue.key;
  row[20] = timestamp;

  for (let col = 1; col < row.length; col++) {
    if (row[col] !== undefined) bugSheet.getRange(rowNum, col + 1).setValue(row[col]);
  }
}

function insertBugRow_(bugSheet, issue, timestamp, instUrl, instance, moduleName) {
  const lastRow = bugSheet.getLastRow();
  const newRow = lastRow + 1;
  const row = [];
  row[0] = 'BUG-' + String(newRow - 5).padStart(4, '0');
  row[1] = 'Functional';
  row[2] = normalizePriority_(issue.fields.priority && issue.fields.priority.name);
  row[3] = normalizeStatus_(issue.fields.status && issue.fields.status.name);
  const fieldMap = getJiraFieldMap_(instance);
  row[4] = getCustomFieldValue_(issue, fieldMap.feature);
  row[5] = getCustomFieldValue_(issue, fieldMap.submodul) || moduleName || '';
  row[6] = issue.fields.summary || '';
  row[7] = convertADF_(issue.fields.description);
  row[8] = getCustomFieldValue_(issue, fieldMap.environment) || issue.fields.environment || '';
  row[9] = getCustomFieldValue_(issue, fieldMap.steps);
  row[10] = getCustomFieldValue_(issue, fieldMap.expected);
  row[11] = getCustomFieldValue_(issue, fieldMap.actual);
  row[12] = '';
  row[13] = issue.fields.reporter && issue.fields.reporter.displayName || '';
  row[14] = issue.fields.assignee && issue.fields.assignee.displayName || '';
  row[15] = issue.fields.created ? issue.fields.created.substring(0, 10) : '';
  row[16] = '';
  row[17] = '';
  row[18] = instUrl + '/browse/' + issue.key;
  row[19] = issue.key;
  row[20] = timestamp;
  row[21] = '';
  bugSheet.getRange(newRow, 1, 1, row.length).setValues([row]);
}

function cleanupClosedBugs_(bugSheet, jiraIssues) {
  const closedKeys = new Set();
  jiraIssues.forEach(function(issue) {
    const statusName = issue.fields.status && issue.fields.status.name;
    if (statusName === 'Closed' || statusName === "Won't Fix") closedKeys.add(issue.key);
  });
  if (closedKeys.size === 0) return 0;

  const data = bugSheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 5; i--) {
    const jiraKey = String(data[i][19]).trim();
    if (closedKeys.has(jiraKey)) rowsToDelete.push(i + 1);
  }
  rowsToDelete.forEach(function(rowNum) { bugSheet.deleteRow(rowNum); });
  return rowsToDelete.length;
}

function getJiraInstanceUrl_(instance) {
  const instances = {'digitalperuri': 'https://digitalperuri.atlassian.net', 'bgn-peruri': 'https://bgn-peruri.atlassian.net'};
  return instances[instance] || instances['digitalperuri'];
}

function getJiraModulField_(instance) {
  const fields = {'digitalperuri': 'cf[10097]', 'bgn-peruri': 'cf[10289]'};
  return fields[instance] || 'cf[10097]';
}

function getJiraCustomFields_(instance) {
  const fields = {
    'digitalperuri': ',customfield_11090,customfield_10095,customfield_10560,customfield_10561,customfield_10562,customfield_11354',
    'bgn-peruri': ',customfield_10298,customfield_10291,customfield_10292,customfield_10293,customfield_10294,customfield_10300'
  };
  return fields[instance] || fields['digitalperuri'];
}

function getJiraFieldMap_(instance) {
  const maps = {
    'digitalperuri': {
      feature: 'customfield_11090', environment: 'customfield_10095', steps: 'customfield_10560',
      expected: 'customfield_10561', actual: 'customfield_10562', submodul: 'customfield_11354'
    },
    'bgn-peruri': {
      feature: 'customfield_10298', environment: 'customfield_10291', steps: 'customfield_10292',
      expected: 'customfield_10293', actual: 'customfield_10294', submodul: 'customfield_10300'
    }
  };
  return maps[instance] || maps['digitalperuri'];
}

function getCustomFieldValue_(issue, fieldId) {
  if (!fieldId || !issue.fields) return '';
  const value = issue.fields[fieldId];
  if (!value) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'object' && value.value) return value.value;
  if (typeof value === 'object' && value.content) return convertADF_(value);
  return '';
}

function normalizePriority_(priority) {
  if (!priority) return '';
  const map = {'Highest': 'Critical', 'Critical': 'Critical', 'High': 'High', 'Medium': 'Medium', 'Low': 'Low', 'Lowest': 'Low', 'Minor': 'Low', 'Trivial': 'Low'};
  return map[priority] || priority;
}

function normalizeStatus_(status) {
  if (!status) return '';
  const sl = status.toLowerCase();
  if (['open', 'to do', 'backlog', 'new'].includes(sl)) return 'Open';
  if (['in progress', 'in review', 'review', 'testing'].includes(sl)) return 'In Progress';
  if (['fixed', 'ready for qa', 'ready for review'].includes(sl)) return 'Fixed';
  if (['verified', 'qa verified'].includes(sl)) return 'Verified';
  if (['closed', 'done'].includes(sl)) return 'Closed';
  if (["won't fix", 'wontfix', 'not a bug', 'invalid'].includes(sl)) return "Won't Fix";
  if (['reopened', 'reopen'].includes(sl)) return 'Reopen';
  return 'Open';
}

function convertADF_(adf) {
  if (!adf) return '';
  if (typeof adf === 'string') return adf;
  function extract(node) {
    if (!node) return '';
    if (node.type === 'text') return node.text || '';
    if (node.type === 'hardBreak') return '\n';
    if (node.type === 'paragraph') return (node.content || []).map(extract).join('') + '\n';
    if (node.type === 'bulletList' || node.type === 'orderedList') {
      return (node.content || []).map(function(item, i) {
        return (node.type === 'orderedList' ? (i + 1) + '. ' : ' • ') + (item.content || []).map(extract).join('');
      }).join('\n') + '\n';
    }
    if (node.type === 'mediaSingle' || node.type === 'mediaInline' || node.type === 'media') return '[Image]';
    if (node.content) return node.content.map(extract).join('');
    return '';
  }
  try {
    return extract(adf).trim();
  } catch (e) {
    return '';
  }
}
