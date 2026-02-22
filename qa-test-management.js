// ===================================================================
//  QA TEST MANAGEMENT -- Standard Template
//  Tabs: TC_Master | TC_Execution | API_Master | API_Execution | _Dashboard | Appendix
//  Run: createQASheet()
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
  ['TC_Master','TC_Execution','API_Master','API_Execution','Summary','Dashboard','_Dashboard','PerfTest','Appendix']
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
  Logger.log('Creating Summary...');
  createSummary(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating PerfTest...');
  createPerfTest(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  Logger.log('Creating Appendix...');
  createAppendix(ss); SpreadsheetApp.flush();
  const s1=ss.getSheetByName('Sheet1');
  if(s1&&ss.getSheets().length>1) try{ss.deleteSheet(s1);}catch(e){}
  // Reorder tabs: move each to correct position
  SpreadsheetApp.flush();
  try {
    const order=['Summary','Appendix','TC_Execution','TC_Master','API_Execution','API_Master','PerfTest'];
    order.slice().reverse().forEach(name=>{
      const sh=ss.getSheetByName(name);
      if(sh){ ss.setActiveSheet(sh); ss.moveActiveSheet(1); }
    });
  } catch(e) { Logger.log('Tab reorder skipped: '+e.message); }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '[OK]  QA Test Management Template berhasil dibuat.\n\n'+
    'TC_Master     -- input test case Web / Mobile\n'+
    'TC_Execution  -- hasil eksekusi per tanggal run\n'+
    'API_Master    -- input test case API\n'+
    'API_Execution -- hasil eksekusi API\n'+
    'Summary       -- info sesi, coverage & run history\n'+
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
  // Column header notes ? guides QA what to fill vs what is auto
  ws.getRange(2,1).setNote('No. Urut. Auto-fill saat input data.');
  ws.getRange(2,2).setNote('SubModul: Kode modul, misal 1.1 / 2.1 / 7.3. Gunakan konsisten agar coverage Dashboard akurat.');
  ws.getRange(2,3).setNote(
    'TC_ID -- Pola Penomoran:\n'+
    '\n'+
    'Format  : [APP].[FEAT].[000]\n'+
    '\n'+
    '[APP]   = Kode aplikasi/modul, maks 3-4 huruf kapital\n'+
    '          Contoh: WEB, MOB, ADM, USR, SHP\n'+
    '[FEAT]  = Kode fitur/halaman, maks 3-4 huruf kapital\n'+
    '          Contoh: LOG, DASH, PRF, CHK, RPT\n'+
    '[000]   = Nomor urut 3 digit, mulai dari 001\n'+
    '\n'+
    'Contoh lengkap:\n'+
    '  WEB.LOG.001  = Web, Login, TC pertama\n'+
    '  WEB.LOG.002  = Web, Login, TC kedua\n'+
    '  MOB.DASH.001 = Mobile, Dashboard, TC pertama\n'+
    '  ADM.USR.015  = Admin, User Mgmt, TC ke-15\n'+
    '\n'+
    'Aturan:\n'+
    '- Harus UNIK -- jangan pernah reuse TC_ID yang sudah ada\n'+
    '- Jangan ubah TC_ID jika sudah ada hasil di Execution\n'+
    '- Urutan: Positive dulu (001), baru Negative (002), Edge Case (003)'
  );
  ws.getRange(2,4).setNote('Feature: Nama fitur/halaman spesifik, misal "Login Page", "Checkout Flow".\nDigunakan untuk grouping Coverage per Feature di Summary.');
  ws.getRange(2,5).setNote(
    'Priority & Dampak:\n'+
    '\n'+
    'CRITICAL  -> Blocker utama. WAJIB PASS sebelum release.\n'+
    '            Jika FAIL/BLOCKED: release DITAHAN.\n'+
    'HIGH      -> Blocker. Harus PASS di sprint yang sama.\n'+
    '            Jika FAIL: perlu approval PM untuk release.\n'+
    'MEDIUM    -> Potential blocker. Fix sebelum UAT.\n'+
    '            Jika FAIL: flagged ke tech lead.\n'+
    'LOW       -> Non-blocker. Fix di sprint berikutnya.\n'+
    'LOWEST    -> Nice to have. Opsional.\n'+
    '\n'+
    'Test Level otomatis:\n'+
    'Critical/High/Medium -> Smoke Test\n'+
    'Low/Lowest           -> Regression Test'
  );
  ws.getRange(2,6).setNote('Platform: Web / Mobile / Web & Mobile');
  ws.getRange(2,7).setNote('Test Type: Positive = happy path | Negative = error case | Edge Case = boundary condition');
  ws.getRange(2,8).setNote('Automation Status:\nAutomated = script sudah ada\nManual = tidak akan diautomasi\nTo Do = belum dikerjakan\nCannot be Automated = teknis tidak memungkinkan');
  ws.getRange(2,9).setNote('Version: Versi aplikasi saat TC ini dibuat, misal v1.0, v2.3');
  ws.getRange(2,10).setNote('[INPUT WAJIB] Skenario singkat dalam 1 kalimat.\nContoh: "User berhasil login dengan kredensial valid"');
  ws.getRange(2,11).setNote('[INPUT WAJIB] Steps dalam format Gherkin:\nGiven [kondisi awal]\nWhen [aksi]\nThen [hasil yang diharapkan]');
  ws.getRange(2,12).setNote('[INPUT WAJIB] Hasil yang diharapkan secara spesifik.\nContoh: "Halaman dashboard tampil, user name muncul di header"');
  ws.getRange(2,14).setNote('[AUTO - JANGAN EDIT] Test Level dihitung otomatis dari Priority:\nCritical/High/Medium = Smoke\nLow/Lowest = Regression');
  ws.getRange(2,10).setNote(
    'Role (RBAC): Peran pengguna yang menjalankan skenario ini.\n'+
    '\n'+
    'Contoh peran: Admin, User, Viewer, Operator, Supervisor, Guest, Super Admin\n'+
    '\n'+
    'Digunakan untuk:\n'+
    '- Memastikan test coverage per role\n'+
    '- Verifikasi akses kontrol (RBAC) berjalan benar\n'+
    '- Contoh: Admin bisa create user, Viewer hanya bisa read'
  );

  const data=[
    [1,'1.1','1.1.001','Informasi Program','High','Web','Positive','Automated','v1.0',
     'Viewer','User dapat melihat halaman informasi program',
     'Given user berada di halaman utama\nWhen klik menu Informasi Program\nThen halaman tampil',
     'Halaman tampil, semua konten tersedia',''],
    [2,'1.1','1.1.002','Informasi Program','Medium','Web','Negative','Manual','v1.0',
     'Viewer','Halaman error saat konten tidak tersedia',
     'Given konten belum tersedia\nWhen user buka halaman\nThen pesan error tampil',
     'Pesan error informatif, halaman tidak crash',''],
    [3,'2.1','2.1.001','Perencanaan Pengiriman','Critical','Web','Positive','Automated','v1.1',
     'Admin','Admin membuat rencana pengiriman baru',
     'Given admin di halaman Perencanaan\nWhen isi form dan klik Simpan\nThen data tersimpan',
     'Data tersimpan dan muncul di daftar',''],
    [4,'2.2','2.2.001','Pelacakan Pengiriman','High','Mobile','Positive','Automated','v1.1',
     'Kurir','Kurir melihat rute pengiriman aktif',
     'Given kurir login di mobile\nWhen buka menu Pengiriman\nThen rute tampil di peta',
     'Rute tampil dengan marker yang tepat',''],
    [5,'7.3','7.3.001','Manajemen Pengguna','Critical','Web','Positive','Automated','v1.0',
     'Admin','Admin menambahkan user baru',
     'Given admin di halaman Manajemen Pengguna\nWhen isi form dan klik Tambah\nThen user ditambahkan',
     'User baru muncul di daftar',''],
    [6,'7.3','7.3.002','Manajemen Pengguna','Low','Web','Negative','Manual','v1.0',
     'Admin','Filter dengan input tidak valid tidak crash',
     'Given admin di halaman filter\nWhen input karakter tidak valid\nThen halaman tetap stabil',
     'Pesan validasi tampil, tidak crash',''],
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
      if(c===13) cell.setHorizontalAlignment('center').setFontWeight('bold').setBackground('#E3F2FD');
    });
    bd(cell_style(ws.getRange(r,14),bg)).setFormula(tlFormula('E'+r))
      .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#E3F2FD');
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
  // Also protect TestLevel col visually
  ws.getRange('N'+DS+':N'+dvEnd).setBackground('#E3F2FD').setFontColor('#1565C0');

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
    .setValue('Test Level otomatis: Critical/High/Medium ? Smoke  .  Low/Lowest ? Regression. Jangan edit kolom M.')
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
  const DS=9, MR=500, STAG=8, STAG_N=3, STATUS_Z=26, SHOT_COL=27;

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
  hdr(ws.getRange(2,SHOT_COL),'#0D47A1','#FFFFFF',8).setValue('[INPUT]\nEvidence Link').setWrap(true);
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
  ws.getRange(2,SHOT_COL).setNote('[INPUT] Paste link screenshot/evidence setelah eksekusi.\nContoh: link Google Drive, Jira attachment, atau URL gambar.');
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
  inputBorder(ws.getRange(DS, SHOT_COL, MR, 1));
  // Blue border on staging cols = INPUT area
  inputBorder(ws.getRange(DS, STAG, MR, STAG_COLS));
  // Blue border on evidence col
  inputBorder(ws.getRange(DS, SHOT_COL, MR, 1));

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
  ws.getRange(DS, SHOT_COL, MR, 1)
    .setBackground('#FAFAFA').setFontColor('#1A73E8')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle').setWrap(true);

  // Column widths -- tight, leave Z and AA only
  [100,110,130,80,80,180,75].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  for(let i=0;i<STAG_N;i++) ws.setColumnWidth(STAG+i,100);
  ws.setColumnWidth(STATUS_Z,100);
  ws.setColumnWidth(SHOT_COL,160);

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
  ws.clear(); ws.setTabColor('#283593');
  const COLS=14, DS=3, MR=1000;
  ws.getRange(1,1,1,COLS).merge();
  hdr(ws.getRange(1,1,1,COLS),'#283593','#FFFFFF',10).setValue('API_MASTER  .  QA PERURI  .  Test Level otomatis dari Priority.');
  ws.setRowHeight(1,28);
  ['No','SubModul','TC_ID','Feature','Method','Endpoint URL','Priority','Auth','Test Type',
   'Automated','Version','Role (RBAC)','Scenario','[AUTO] Test Level']
    .forEach((h,i)=>hdr(ws.getRange(2,i+1),'#0D47A1').setValue(h).setWrap(true));
  ws.setRowHeight(2,38);
  [36,110,110,130,70,220,80,110,90,140,65,200,90,110].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  // Column notes for API_Master
  ws.getRange(2,1).setNote('No. Urut.');
  ws.getRange(2,2).setNote('SubModul: Kode modul. Gunakan kode yang sama dengan TC_Master untuk konsistensi.');
  ws.getRange(2,3).setNote(
    'TC_ID API -- Pola Penomoran:\n'+
    '\n'+
    'Format  : API.[SVC].[FEAT].[000]\n'+
    '\n'+
    'API.    = Prefix wajib untuk semua API TC\n'+
    '[SVC]   = Kode service/domain, maks 3-4 huruf kapital\n'+
    '          Contoh: AUTH, USER, ORD, PAY, INV\n'+
    '[FEAT]  = Kode endpoint/resource, maks 3-4 huruf kapital\n'+
    '          Contoh: LOG, LIST, CRT, UPD, DEL\n'+
    '[000]   = Nomor urut 3 digit, mulai dari 001\n'+
    '\n'+
    'Contoh lengkap:\n'+
    '  API.AUTH.LOG.001  = Auth service, Login endpoint, TC pertama\n'+
    '  API.USER.CRT.001  = User service, Create, TC pertama\n'+
    '  API.USER.CRT.002  = User service, Create, TC kedua (negative)\n'+
    '  API.PAY.CHK.005   = Payment, Checkout, TC ke-5\n'+
    '\n'+
    'Urutan: Positive (2xx) dulu -> Negative (4xx) -> Edge Case'
  );
  ws.getRange(2,4).setNote('Feature: Nama fitur/domain API. Digunakan untuk grouping di Summary.');
  ws.getRange(2,5).setNote('HTTP Method: GET/POST/PUT/DELETE/PATCH');
  ws.getRange(2,6).setNote('Endpoint URL: Path saja tanpa base URL.\nContoh: /api/v1/users/{id}\nParameter dinamis gunakan {curly bracket}.');
  ws.getRange(2,7).setNote(
    'Priority API & Dampak:\n'+
    '\n'+
    'CRITICAL  -> Blocker. API utama tidak bisa digunakan.\n'+
    '            FAIL = release DITAHAN.\n'+
    'HIGH      -> Blocker. Fungsi penting terganggu.\n'+
    '            FAIL = perlu approval PM.\n'+
    'MEDIUM    -> Potential blocker. Flagged ke tech lead.\n'+
    'LOW       -> Non-blocker. Fix sprint berikutnya.\n'+
    'LOWEST    -> Opsional.\n'+
    '\n'+
    'Test Level: Critical/High/Medium -> Smoke | Low/Lowest -> Regression'
  );
  ws.getRange(2,8).setNote('Auth: Jenis autentikasi yang diperlukan.\nContoh: Bearer Token, Basic Auth, API Key, (none)');
  ws.getRange(2,9).setNote('Test Type: Positive = happy path | Negative = error/invalid input | Edge Case = boundary');
  ws.getRange(2,10).setNote('Automation Status: Automated/Manual/To Do/Cannot be Automated');
  ws.getRange(2,11).setNote('Version: Versi API saat TC dibuat.');
  ws.getRange(2,12).setNote('[INPUT WAJIB] Skenario singkat: apa yang ditest dan expected HTTP status.\nContoh: "User login berhasil -- 200"\nContoh: "Token expired ditolak -- 401"');
  ws.getRange(2,13).setNote('[AUTO - JANGAN EDIT] Test Level otomatis dari Priority.');
  ws.getRange(2,12).setNote(
    'Role (RBAC): Peran yang memiliki akses ke endpoint ini.\n'+
    '\n'+
    'Contoh: Admin, Super Admin, User, Viewer\n'+
    '\n'+
    'Gunakan untuk verifikasi authorization:\n'+
    '- Apakah role yang sesuai bisa akses endpoint?\n'+
    '- Apakah role yang tidak sesuai mendapat 403 Forbidden?'
  );

  const METHOD_COLORS={GET:['#E8F0FE','#1A237E'],POST:['#E8F5E9','#1B5E20'],
    PUT:['#FFF8E1','#E65100'],DELETE:['#FCE4EC','#880E4F'],PATCH:['#F3E5F5','#4A148C']};

  const data=[
    [1,'1.1','API.1.1.001','Informasi Program','GET','/api/v1/program/info','High','Bearer Token','Positive','Automated','v1.0','Viewer','Get informasi program -- 200',''],
    [2,'1.1','API.1.1.002','Informasi Program','GET','/api/v1/program/info','Medium','(none)','Negative','Automated','v1.0','Viewer','Tanpa token ditolak -- 401',''],
    [3,'2.1','API.2.1.001','Perencanaan Pengiriman','POST','/api/v1/delivery/plan','Critical','Bearer Token','Positive','Automated','v1.1','Admin','Create rencana berhasil -- 201',''],
    [4,'2.1','API.2.1.002','Perencanaan Pengiriman','GET','/api/v1/delivery/plan','High','Bearer Token','Positive','Automated','v1.1','Admin','Ambil list rencana -- 200',''],
    [5,'7.3','API.7.3.001','Manajemen Pengguna','POST','/api/v1/users','Critical','Bearer Token + Admin','Positive','Automated','v1.0','Super Admin','Create user berhasil -- 201',''],
    [6,'7.3','API.7.3.002','Manajemen Pengguna','POST','/api/v1/users','High','Bearer Token + Admin','Negative','Automated','v1.0','Super Admin','Email duplikat ditolak -- 422',''],
    [7,'7.3','API.7.3.003','Manajemen Pengguna','DELETE','/api/v1/users/{id}','Low','Bearer Token + Admin','Positive','Manual','v1.0','Super Admin','Delete user berhasil -- 200',''],
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
      if(c===13) cell.setHorizontalAlignment('center').setFontWeight('bold').setBackground('#E3F2FD');
    });
    bd(cell_style(ws.getRange(r,14),bg)).setFormula(tlFormula('G'+r)).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#E3F2FD');
    ws.setRowHeight(r,44);
  });
  for(let r=DS+data.length;r<=DS+MR;r++) ws.getRange(r,14).setFormula(tlFormula('G'+r));

  const dvEnd=DS+MR;
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
  ws.getRange('N'+DS+':N'+dvEnd).setBackground('#E3F2FD').setFontColor('#1565C0');

  const cf=ws.getConditionalFormatRules();
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Smoke')
    .setBackground('#FFF8F0').setFontColor('#BF360C').setBold(true).setRanges([ws.getRange('M'+DS+':M'+dvEnd)]).build());
  cf.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Regression')
    .setBackground('#F1F8E9').setFontColor('#33691E').setBold(true).setRanges([ws.getRange('M'+DS+':M'+dvEnd)]).build());
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

  const DS=9, MR=500, STAG=7, STAG_N=3, STATUS_Z=26, SHOT_COL=27;

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
  hdr(ws.getRange(2,SHOT_COL),'#0D47A1','#FFFFFF',8).setValue('[INPUT]\nEvidence Link').setWrap(true);
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

  // Sync API_Master: C=TC_ID B=SubModul D=Feature G=Priority E=Method F=URL M=TestLevel
  ws.getRange(DS,1).setFormula('=ARRAYFORMULA(IF(API_Master!C3:C'+(MR+2)+'<>"",API_Master!C3:C'+(MR+2)+',""))');
  ws.getRange(DS,2).setFormula('=ARRAYFORMULA(IF(API_Master!B3:B'+(MR+2)+'<>"",API_Master!B3:B'+(MR+2)+',""))');
  ws.getRange(DS,3).setFormula('=ARRAYFORMULA(IF(API_Master!D3:D'+(MR+2)+'<>"",API_Master!D3:D'+(MR+2)+',""))');
  ws.getRange(DS,4).setFormula('=ARRAYFORMULA(IF(API_Master!G3:G'+(MR+2)+'<>"",API_Master!G3:G'+(MR+2)+',""))');
  ws.getRange(DS,5).setFormula('=ARRAYFORMULA(IF(API_Master!C3:C'+(MR+2)+'<>"",API_Master!E3:E'+(MR+2)+'&"  "&API_Master!F3:F'+(MR+2)+',""))');
  ws.getRange(DS,6).setFormula('=ARRAYFORMULA(IF(API_Master!N3:N'+(MR+2)+'<>"",API_Master!N3:N'+(MR+2)+',""))');
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

  ws.getRange(DS, SHOT_COL, MR, 1)
    .setBackground('#FAFAFA').setFontColor('#1A73E8')
    .setFontFamily('Arial').setFontSize(9).setVerticalAlignment('middle').setWrap(true);

  [100,110,130,80,220,75].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  for(let i=0;i<STAG_N;i++) ws.setColumnWidth(STAG+i,100);
  ws.setColumnWidth(STATUS_Z,100);
  ws.setColumnWidth(SHOT_COL,160);

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
    ['Project / Sprint:',null],['Period:',null],['QA Lead:',null],['PIC QA:',null],
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
  leftFields.forEach(([labelText, dvList, isStatus, isAutoPerf],i)=>{
    const row=R+i;
    lbl(row,L,labelText,'#E3F2FD');
    if(isAutoPerf){
      m_(row,L+1,1,LW-1);
      bd(ws.getRange(row,L+1)).setFormula('=IFERROR(IF(COUNTA(PerfTest!E16:E45)=0,"--",IF(COUNTIF(PerfTest!L16:L45,"FAIL")>0,"FAIL","PASS")),"--")')
        .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');
      stCF(ws.getRange(row,L+1));
    } else {
      inp(row,L+1,LW-1,dvList,isStatus);
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
  const MAX_RUNS = 10;
  for (let idx=0; idx<MAX_RUNS; idx++) {
    const rl=R+idx, bg=idx%2===0?'#F8F9FA':'#FFFFFF';
    const tcDateCol = 'INDIRECT("TC_Execution!"&ADDRESS(2,'+idx+'+8))'; // col H=8, I=9...
    // Date: pull from TC_Execution row 2, col H+idx
    const dateForm = '=IFERROR(INDEX(TC_Execution!2:2,'+(idx+8)+'),"")';
    const aDateForm = '=IFERROR(INDEX(API_Execution!2:2,'+(idx+7)+'),"")';
    // Web row - date auto from TC_Execution row 2
    bd(ws.getRange(rl,L)).setFormula(dateForm).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left').setFontWeight('bold').setNumberFormat('yyyy-mm-dd');
    const mfl='MATCH('+colLetter(L)+rl+',TC_Execution!2:2,0)';
    bd(ws.getRange(rl,L+1)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",INDEX(TC_Execution!4:4,'+mfl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,L+2)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",INDEX(TC_Execution!5:5,'+mfl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,L+3)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",INDEX(TC_Execution!6:6,'+mfl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,L+4)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",'+colLetter(L+1)+rl+'/MAX(1,'+colLetter(L+1)+rl+'+'+colLetter(L+2)+rl+'+'+colLetter(L+3)+rl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rl,L+5)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'=""," ",('+colLetter(L+1)+rl+'+'+colLetter(L+2)+rl+'+'+colLetter(L+3)+rl+')/MAX(1,'+wTOT+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rl,L+6)).setFormula('=IFERROR(IF('+colLetter(L)+rl+'="","",INDEX(TC_Execution!3:3,'+mfl+')),"--")').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
    // API row
    bd(ws.getRange(rl,R_)).setFormula(aDateForm).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left').setFontWeight('bold').setNumberFormat('yyyy-mm-dd');
    const mfa='MATCH('+colLetter(R_)+rl+',API_Execution!2:2,0)';
    bd(ws.getRange(rl,R_+1)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",INDEX(API_Execution!4:4,'+mfa+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,R_+2)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",INDEX(API_Execution!5:5,'+mfa+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,R_+3)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",INDEX(API_Execution!6:6,'+mfa+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center');
    bd(ws.getRange(rl,R_+4)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",'+colLetter(R_+1)+rl+'/MAX(1,'+colLetter(R_+1)+rl+'+'+colLetter(R_+2)+rl+'+'+colLetter(R_+3)+rl+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rl,R_+5)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'=""," ",('+colLetter(R_+1)+rl+'+'+colLetter(R_+2)+rl+'+'+colLetter(R_+3)+rl+')/MAX(1,'+aTOT+')),0)').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setNumberFormat('0%');
    bd(ws.getRange(rl,R_+6)).setFormula('=IFERROR(IF('+colLetter(R_)+rl+'="","",INDEX(API_Execution!3:3,'+mfa+')),"--")').setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
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
  // D. COVERAGE PER SUBMODUL
  // =====================================================================
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#1565C0').setValue('D.  COVERAGE PER SUBMODUL  -  Web / Mobile');
  m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#283593').setValue('D.  COVERAGE PER SUBMODUL  -  API');
  ws.setRowHeight(R,20); R++;

  function buildCov(startRow,sc,master,subCol,prioCol,autoCol,execSh,hbg){
    ['SubModul','Total','Smoke','Regression','Auto%','Pass%'].forEach((h,i)=>
      h_(ws.getRange(startRow,sc+i),hbg).setValue(h).setFontSize(8).setWrap(true));
    ws.setRowHeight(startRow,22);
    const DS=startRow+1, MAX=12;
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
  R+=14;

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
function createAppendix(ss) {
  const ws = safeSheet(ss,'Appendix');
  ws.clear(); ws.setTabColor('#1A237E');

  ws.getRange(1,1,1,4).merge();
  hdr(ws.getRange(1,1,1,4),'#0D47A1','#FFFFFF',11)
    .setValue('APPENDIX  .  Definisi, Konvensi & Panduan');
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

  sec('1. STRUKTUR TAB','#0D47A1');
  row2('TC_Master',     'Master list test case Web / Mobile.\nKolom [INPUT]: SubModul, TC_ID, Feature, Priority, Platform, Test Type, Automation, Version, Role (RBAC), Scenario, Steps, Expected Result.\nKolom [AUTO]: Test Level (kolom N) -- jangan diedit.\nFormat TC_ID: [SubModul].[3-digit]  contoh: 1.1.001  2.3.015\nRole = peran RBAC yang menjalankan skenario, contoh: Admin, User, Viewer.');
  row2('TC_Execution',  'Kolom identitas sync otomatis dari TC_Master. Isi kolom staging dengan PASSED / FAILED / BLOCKED / TODO.\nTambah kolom ke kanan untuk setiap run. LATEST STATUS di kolom Z otomatis.\nKolom AA = link screenshot / evidence.\nIN PROGRESS otomatis jika ada PASSED dan TODO di skenario yang sama.');
  row2('API_Master',    'Master list test case API. Method (E) dan Endpoint URL (F) terpisah.\nKolom [INPUT]: SubModul, TC_ID, Feature, Method, Endpoint, Priority, Auth, Test Type, Automation, Version, Role (RBAC), Scenario.\nKolom [AUTO]: Test Level (kolom N) -- jangan diedit.\nFormat TC_ID: API.[SubModul].[3-digit]  contoh: API.1.1.001\nRole = peran RBAC yang diuji aksesnya, contoh: Admin, Super Admin, Viewer.');
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
