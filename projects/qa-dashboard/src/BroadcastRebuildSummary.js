/**
 * BroadcastRebuildSummary.js
 *
 * Rebuild Summary tabs in all QATM with latest template (includes VAPT Blocker Breakdown)
 * Preserves user input fields (Project, Module, etc.)
 */

function broadcastRebuildSummary() {
  const dashboardSs = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const modules = getModuleList_(dashboardSs);
  const targets = [];
  const seen = {};

  modules.forEach(mod => {
    if (!mod || !mod.id || seen[mod.id]) return;
    seen[mod.id] = true;
    targets.push(mod);
  });

  if (targets.length === 0) {
    ui.alert('No active QATM targets found in Config.');
    return;
  }

  const response = ui.alert(
    'Broadcast: Rebuild Summary Tabs',
    'Will REBUILD Summary tab in ' + targets.length + ' QATM sheets with latest template.\n\n' +
    '⚠️ WARNING:\n' +
    '• Summary tab will be DELETED and RECREATED\n' +
    '• Manual edits in Summary will be LOST\n' +
    '• Input fields (Project/Module) will be RESTORED\n' +
    '• Formulas rebuilt from template\n\n' +
    '✅ What will be added:\n' +
    '• F. VAPT FINDINGS SUMMARY section\n' +
    '• VAPT Blocker Breakdown (Critical/High/Medium)\n' +
    '• Updated formulas to reference "VAPT - Detail Finding"\n\n' +
    '✅ Safe:\n' +
    '• Other tabs (TC, BugReport, VAPT) unchanged\n' +
    '• User inputs preserved\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let created = 0;
  let skipped = 0;
  let failed = 0;
  const errors = [];

  targets.forEach(mod => {
    try {
      const qatmSs = SpreadsheetApp.openById(mod.id);
      const existingSummary = qatmSs.getSheetByName('Summary');

      if (!existingSummary) {
        skipped++;
        return;
      }

      // Save input values
      let savedInputs = {};
      try {
        const data = existingSummary.getDataRange().getValues();
        // Find input rows (usually B2:B10 in Summary)
        savedInputs = {
          project: getValue_(data, 'Project:', 1),
          module: getValue_(data, 'Modul:', 1),
          submodule: getValue_(data, 'Submodul:', 1),
          qaLead: getValue_(data, 'QA Lead:', 1),
          picQA: getValue_(data, 'PIC QA:', 1),
          environment: getValue_(data, 'Environment:', 1),
          issueTracker: getValue_(data, 'Issue Tracker', 1),
          testStatus: getValue_(data, 'Test Status:', 1),
          scopeNotes: getValue_(data, 'Scope / Notes:', 1),
        };
      } catch (e) {
        Logger.log('Could not save inputs from ' + mod.id + ': ' + e.message);
      }

      // Call rebuild function from QATM
      rebuildSummaryInQATM_(qatmSs, savedInputs);
      created++;

    } catch (error) {
      failed++;
      errors.push((mod.project || '') + ' / ' + (mod.module || '') + ': ' + error.message);
      Logger.log('Rebuild Summary broadcast failed [' + (mod.id || '-') + ']: ' + error.stack);
    }
  });

  ui.alert(
    'Summary Rebuild Broadcast Complete',
    'Rebuilt: ' + created + '\n' +
    'Skipped (no Summary tab): ' + skipped + '\n' +
    'Failed: ' + failed +
    (errors.length ? '\n\nErrors:\n' + errors.slice(0, 8).join('\n') : '') +
    '\n\n✅ Summary tabs now include VAPT Blocker Breakdown section.',
    ui.ButtonSet.OK
  );
}

function getValue_(data, label, colOffset) {
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      if (String(data[i][j]).includes(label)) {
        return data[i][j + colOffset] || '';
      }
    }
  }
  return '';
}

function rebuildSummaryInQATM_(ss, savedInputs) {
  // Delete existing Summary
  const existingSummary = ss.getSheetByName('Summary');
  if (existingSummary) {
    ss.deleteSheet(existingSummary);
    SpreadsheetApp.flush();
    Utilities.sleep(300);
  }

  // Recreate Summary using inline version of createSummary
  createSummaryFromTemplate_(ss);

  // Restore inputs
  const newSummary = ss.getSheetByName('Summary');
  if (newSummary && savedInputs && savedInputs.project) {
    try {
      newSummary.getRange('B2').setValue(savedInputs.project);
      newSummary.getRange('B3').setValue(savedInputs.module);
      newSummary.getRange('B4').setValue(savedInputs.submodule);
      newSummary.getRange('B5').setValue(savedInputs.qaLead);
      newSummary.getRange('B6').setValue(savedInputs.picQA);
      newSummary.getRange('B7').setValue(savedInputs.environment);
      newSummary.getRange('B8').setValue(savedInputs.issueTracker);
      newSummary.getRange('B9').setValue(savedInputs.testStatus);
      newSummary.getRange('B10').setValue(savedInputs.scopeNotes);
    } catch (e) {
      Logger.log('Could not restore inputs: ' + e.message);
    }
  }
}

// ===== Helper functions from MasterQATCM.js =====

// Safe sheet creator
function safeSheet(ss, name) {
  try {
    const existing = ss.getSheetByName(name);
    if (existing) { ss.deleteSheet(existing); SpreadsheetApp.flush(); Utilities.sleep(300); }
  } catch(e) {}
  for (let i = 0; i < 3; i++) {
    try {
      const sh = ss.insertSheet(name, 0);
      if (sh && sh.getRange) { return sh; }
    } catch(e) { Logger.log('safeSheet attempt ' + (i+1) + ' failed: ' + e.message); }
    SpreadsheetApp.flush();
    Utilities.sleep(500);
  }
  throw new Error('safeSheet: could not create sheet ' + name + ' after 3 attempts');
}

// Peruri footer
function addPeruriFooter(ws, lastRow, totalCols) {
  const r = lastRow + 2;
  ws.getRange(r,1,1,totalCols).merge();
  ws.getRange(r,1)
    .setValue('(c) QA INA Digital  |  Template ini merupakan properti QA Team INA Digital  |  Dilarang digunakan/disebarluaskan tanpa izin  |  departemen.qa@inadigital.co.id')
    .setBackground('#ECEFF1').setFontColor('#546E7A')
    .setFontSize(7).setFontStyle('italic').setFontFamily('Arial')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange(r,1,1,totalCols)
    .setBorder(true, false, false, false, false, false, '#1976D2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

// Border helper
function bd(r) {
  return r.setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
}

// Dropdown validation
function dv(list) {
  return SpreadsheetApp.newDataValidation().requireValueInList(list,true).setAllowInvalid(false).build();
}

// Input border
function inputBorder(range) {
  range.setBorder(true,true,true,true,false,false,'#1976D2', SpreadsheetApp.BorderStyle.SOLID);
}

function createSummaryFromTemplate_(ss) {
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
      '(ISNUMBER(MATCH(BugReport!D5:D2000,{\"Open\",\"In Progress\",\"Reopen\",\"Fixed\",\"Verified\",\"In Progress VAPT\",\"Done VAPT\"},0)))*' +
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
      '  • Status: Open / In Progress / Reopen / Fixed / Verified / In Progress VAPT / Done VAPT\n' +
      '  • Priority: Critical / High / Medium\n\n' +
      'NOT Blocker:\n' +
      '  • Closed (final)\n' +
      '  • Won\'t Fix (rejected)\n\n' +
      'Why Fixed/Verified/Done VAPT still blocker?\n' +
      'Semua status sebelum Closed masih blocker karena belum release.\n' +
      'Hanya Closed yang tidak blocker.\n\n' +
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

  // =====================================================================
  // F. VAPT FINDINGS SUMMARY
  // =====================================================================
  ws.setRowHeight(R,6); R++;
  m_(R,L,1,LW); h_(ws.getRange(R,L),'#BF360C').setValue('F.  VAPT FINDINGS SUMMARY');
  ws.setRowHeight(R,20); R++;

  // Overview row
  lbl(R,L,'Total VAPT Findings','#FFEBEE');
  m_(R,L+1,1,LW-1);
  bd(ws.getRange(R,L+1)).setFormula('=IFERROR(COUNTA(FILTER(\'VAPT - Detail Finding\'!A:A,\'VAPT - Detail Finding\'!A:A<>\"\",\'VAPT - Detail Finding\'!A:A<>"Finding ID")),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R,24); R++;

  // By Risk Level header
  ws.setRowHeight(R,6); R++;
  m_(R,L,1,5);
  bd(ws.getRange(R,L)).setValue('By Risk Level (Adjusted Risk)')
      .setBackground('#FFCCBC').setFontColor('#BF360C').setFontWeight('bold')
      .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
  ws.setRowHeight(R,18); R++;

  // Risk level table headers
  ['Risk Level','Count','Risk Level','Count'].forEach((h,i)=>{
    const col = i<2 ? L+i : L+i+1;
    h_(ws.getRange(R,col),'#D84315').setValue(h).setFontSize(8);
  });
  ws.setRowHeight(R,18); R++;

  // Risk level data rows
  const riskLevels = [
    ['Critical', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"Critical"),0)', 'Low', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"Low"),0)'],
    ['High', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"High"),0)', 'Informational', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"Informational"),0)'],
    ['Medium', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!E:E,"Medium"),0)', '', ''],
  ];
  riskLevels.forEach((row,i)=>{
    const rr=R+i, bg=i%2===0?'#FFF3E0':'#FFFFFF';
    bd(ws.getRange(rr,L)).setValue(row[0]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
    bd(ws.getRange(rr,L+1)).setFormula(row[1]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
    bd(ws.getRange(rr,L+3)).setValue(row[2]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
    if(row[3]) bd(ws.getRange(rr,L+4)).setFormula(row[3]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
    ws.setRowHeight(rr,16);
  });
  R+=3;

  // By Status Fix header
  ws.setRowHeight(R,6); R++;
  m_(R,L,1,5);
  bd(ws.getRange(R,L)).setValue('By Status Fix (Dev Team)')
      .setBackground('#E3F2FD').setFontColor('#0D47A1').setFontWeight('bold')
      .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
  ws.setRowHeight(R,18); R++;

  // Status Fix table headers
  ['Status','Count','Status','Count'].forEach((h,i)=>{
    const col = i<2 ? L+i : L+i+1;
    h_(ws.getRange(R,col),'#1565C0').setValue(h).setFontSize(8);
  });
  ws.setRowHeight(R,18); R++;

  // Status Fix data rows
  const statusFix = [
    ['Todo', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"Todo"),0)', 'Done', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"Done"),0)'],
    ['On Progress Remediation', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"On Progress Remediation"),0)', 'Accepted', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"Accepted"),0)'],
    ['Ready to Retest', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"Ready to Retest"),0)', 'False Positive', '=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!H:H,"False Positive"),0)'],
  ];
  statusFix.forEach((row,i)=>{
    const rr=R+i, bg=i%2===0?'#F8F9FA':'#FFFFFF';
    bd(ws.getRange(rr,L)).setValue(row[0]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
    bd(ws.getRange(rr,L+1)).setFormula(row[1]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
    bd(ws.getRange(rr,L+3)).setValue(row[2]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
    bd(ws.getRange(rr,L+4)).setFormula(row[3]).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
    ws.setRowHeight(rr,16);
  });
  R+=3;

  // By Status Re-VAPT header
  ws.setRowHeight(R,6); R++;
  m_(R,L,1,5);
  bd(ws.getRange(R,L)).setValue('By Status Re-VAPT (Pentester)')
      .setBackground('#F3E5F5').setFontColor('#6A1B9A').setFontWeight('bold')
      .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
  ws.setRowHeight(R,18); R++;

  // Status Re-VAPT simple row
  const reVaptRow = R;
  bd(ws.getRange(reVaptRow,L)).setValue('Open').setBackground('#F8F9FA').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
  bd(ws.getRange(reVaptRow,L+1)).setFormula('=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!I:I,"Open"),0)').setBackground('#F8F9FA').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
  bd(ws.getRange(reVaptRow,L+3)).setValue('Closed').setBackground('#F8F9FA').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
  bd(ws.getRange(reVaptRow,L+4)).setFormula('=IFERROR(COUNTIF(\'VAPT - Detail Finding\'!I:I,"Closed"),0)').setBackground('#F8F9FA').setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setFontWeight('bold');
  ws.setRowHeight(reVaptRow,16); R++;

  // BLOCKER BREAKDOWN section
  ws.setRowHeight(R,6); R++;
  m_(R,L,1,5);
  bd(ws.getRange(R,L)).setValue('VAPT Blockers (Status: Todo / On Progress / Ready to Retest / Accepted)')
      .setBackground('#FFCDD2').setFontColor('#C62828').setFontWeight('bold')
      .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('left');
  ws.setRowHeight(R,18); R++;

  // Blocker by severity
  lbl(R,L,'VAPT Blocker Count','#FFEBEE');
  m_(R,L+1,1,LW-1);
  bd(ws.getRange(R,L+1)).setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*((\'VAPT - Detail Finding\'!E:E="Critical")+(\'VAPT - Detail Finding\'!E:E="High")+(\'VAPT - Detail Finding\'!E:E="Medium"))),0)')
      .setBackground('#FFCDD2').setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
      .setFontColor('#B71C1C').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R,28); R++;

  lbl(R,L,'VAPT Blocker Critical','#FFF3E0');
  m_(R,L+1,1,LW-1);
  bd(ws.getRange(R,L+1)).setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*(\'VAPT - Detail Finding\'!E:E="Critical")),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R,22); R++;

  lbl(R,L,'VAPT Blocker High','#FFF3E0');
  m_(R,L+1,1,LW-1);
  bd(ws.getRange(R,L+1)).setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*(\'VAPT - Detail Finding\'!E:E="High")),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R,22); R++;

  lbl(R,L,'VAPT Blocker Medium','#FFF3E0');
  m_(R,L+1,1,LW-1);
  bd(ws.getRange(R,L+1)).setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*(\'VAPT - Detail Finding\'!E:E="Medium")),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R,22); R++;

  lbl(R,L,'VAPT Non Blocker Count','#E8F5E9');
  m_(R,L+1,1,LW-1);
  bd(ws.getRange(R,L+1)).setFormula('=IFERROR(SUMPRODUCT((\'VAPT - Detail Finding\'!H:H<>"Done")*(\'VAPT - Detail Finding\'!H:H<>"False Positive")*((\'VAPT - Detail Finding\'!E:E="Low")+(\'VAPT - Detail Finding\'!E:E="Informational"))),0)')
      .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(R,22); R++;

  // Note
  m_(R,L,1,LW);
  ws.getRange(R,L).setValue('Target: VAPT Blocker Count = 0 sebelum closure sign-off. Blocker = severity Critical/High/Medium dengan status selain Done/False Positive.')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic').setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(R,14); R++;

  addPeruriFooter(ws,R+1,21);
  ws.setFrozenRows(0);
}


