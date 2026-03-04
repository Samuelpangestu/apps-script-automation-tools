/**
 * QA_Portfolio_Dashboard.js  —  v40
 * ═══════════════════════════════════════════════════════════════════════
 * Paste SELURUH FILE ini ke Apps Script QA PORTFOLIO DASHBOARD.
 *
 * SETUP PERTAMA KALI:
 *   createDashboard()    → buat semua tab (Config, Overview, Smoke, ...)
 *
 * REFRESH DATA:
 *   refreshDashboard()   → pull data dari semua modul aktif di Config
 *
 * AUTO TRIGGER:
 *   setupTrigger()       → auto refresh setiap 1 jam
 *
 * TABS YANG DIBUAT:
 *   Config   — daftar modul (Active, Project, Module, SubModule, PIC, QA Lead, SpreadsheetID)
 *   Overview — KPI portfolio: WEB | SMOKE WEB | API | SMOKE API | BUGS (25 col)
 *   Smoke    — dedicated Smoke Test view + 5 charts per modul
 *   Blockers — TC Critical/High yang FAILED/BLOCKED
 *   Coverage — coverage per SubModul
 *   History  — trend data tiap refresh (termasuk Smoke trend)
 *   _Raw     — cache internal (jangan edit manual)
 *
 * LAYOUT OVERVIEW (25 kolom):
 *   Col  1-4 : MODULE INFO  (SubModule, Project, Module, PIC/Team)
 *   Col  5-9 : WEB/MOBILE  (Total, Pass, Fail, Block, Pass%)
 *   Col 10-12: SMOKE WEB   (Total, Pass%, Exec%)
 *   Col 13-17: API          (Total, Pass, Fail, Block, Pass%)
 *   Col 18-20: SMOKE API    (Total, Pass%, Exec%)
 *   Col 21   : PERF
 *   Col 22-24: BUGS         (Total, Blocker, Critical)
 *   Col 25   : NOTES
 *
 * QA TEAM LEAD:
 *   - Dibaca dari Config col F (manual input)
 *   - ATAU auto-filled dari Summary baris B4 saat refresh
 *   - Tampil di tab Smoke col "QA Lead"
 * ═══════════════════════════════════════════════════════════════════════
 */


// ═══════════════════════════════════════════════════════════════════════
// SETUP & REFRESH
// ═══════════════════════════════════════════════════════════════════════

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Overview','Smoke','Blockers','Coverage','History','_Raw','Config'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });
  buildConfig(ss);
  buildOverview(ss);
  buildSmoke(ss);
  buildBlockers(ss);
  buildCoverage(ss);
  buildHistory(ss);
  buildRaw(ss);
  ss.setActiveSheet(ss.getSheetByName('Config'));
  safeAlert_('Dashboard berhasil dibuat!\n\nLangkah selanjutnya:\n1. Isi tab Config dengan Spreadsheet ID modul\n2. Isi kolom "QA Team Lead" manual (atau otomatis dari Summary B4 saat refresh)\n3. Jalankan refreshDashboard()');
}

function refreshDashboard() {
  Logger.log('refreshDashboard START: ' + new Date());
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const modules = getModuleList_(ss);
  if (modules.length === 0) {
    safeAlert_('Belum ada modul aktif di Config.\nIsi tab Config dulu lalu refresh.');
    return;
  }

  const allData = [];
  modules.forEach(mod => {
    Logger.log('Pulling: ' + mod.name + ' [' + mod.id + ']');
    try {
      allData.push(pullModuleData_(mod));
      Logger.log('OK: ' + mod.name);
    } catch(e) {
      Logger.log('ERROR ' + mod.name + ': ' + e.message);
      allData.push(emptyModuleData_(mod, 'ERROR: ' + e.message));
    }
    Utilities.sleep(150);
  });

  writeOverview(ss, allData);
  writeSmoke(ss, allData);
  writeBlockers(ss, allData);
  writeCoverage(ss, allData);
  appendHistory(ss, allData);
  updateRaw(ss, allData);
  updateConfig(ss, allData);  // write back PIC + QA Lead from Summary

  const ts = 'Last refreshed: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss');
  ['Overview','Smoke'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) sh.getRange(1,1).setValue(ts);
  });

  Logger.log('refreshDashboard DONE');
  safeAlert_('Refresh selesai! ' + allData.length + ' modul di-update.\n\nQA Lead otomatis diisi dari Summary B4 (jika tersedia).');
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshDashboard').timeBased().everyHours(1).create();
  safeAlert_('Trigger set! Auto-refresh setiap 1 jam.');
}

function safeAlert_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}


// ═══════════════════════════════════════════════════════════════════════
// MODULE LIST
// ═══════════════════════════════════════════════════════════════════════

function getModuleList_(ss) {
  const cfg = ss.getSheetByName('Config');
  if (!cfg) return [];
  const data = cfg.getDataRange().getValues();
  const modules = [];
  for (let i = 3; i < data.length; i++) {
    const active    = String(data[i][0]).trim().toUpperCase();
    const project   = String(data[i][1]).trim();
    const module    = String(data[i][2]).trim();
    const submodule = String(data[i][3]).trim();
    const team      = String(data[i][4]).trim();
    const lead      = String(data[i][5]).trim();   // col F = QA Team Lead (Config manual)
    const id        = String(data[i][6]).trim();
    if (active !== 'Y' || !id || id.length < 10 || id === 'PASTE_SPREADSHEET_ID_HERE') continue;
    modules.push({ name: submodule||project, id, project, module, submodule, team, lead, active: true });
  }
  return modules;
}


// ═══════════════════════════════════════════════════════════════════════
// DATA PULL — per module
// ═══════════════════════════════════════════════════════════════════════

function pullModuleData_(mod) {
  const src  = SpreadsheetApp.openById(mod.id);
  const tcm  = src.getSheetByName('TC_Master');
  const tce  = src.getSheetByName('TC_Execution');
  const apim = src.getSheetByName('API_Master');
  const apie = src.getSheetByName('API_Execution');
  const perf = src.getSheetByName('PerfTest');
  const summ = src.getSheetByName('Summary');
  const bugr = src.getSheetByName('BugReport');

  const SUMM_KPI_ROW = 13;  // Summary row 13: main Web+API KPI values

  let picQA = mod.team||'', qaLead = mod.lead||'', projectSprint = '';
  let wTotal=0,wPassed=0,wFailed=0,wBlocked=0,wInProg=0,wTodo=0,wPassRate=0,wAutoRate=0,wExecRate=0;
  let aTotal=0,aPassed=0,aFailed=0,aBlocked=0,aInProg=0,aTodo=0,aPassRate=0,aAutoRate=0,aExecRate=0;
  let wSmokeTotal=0,wSmokePassed=0,wSmokeFailed=0,wSmokeBlocked=0,wSmokeInProg=0,wSmokeTodo=0;
  let wSmokePassRate=0,wSmokeAutoRate=0,wSmokeExecRate=0;
  let aSmokeTotal=0,aSmokePassed=0,aSmokePassRate=0,aSmokeAutoRate=0,aSmokeExecRate=0;
  let perfResult = '--';

  try {
    if (summ) {
      const ps  = summ.getRange(2,2).getValue();  // B2 = Project/Sprint
      const ql  = summ.getRange(4,2).getValue();  // B4 = QA Lead (from Summary)
      const pic = summ.getRange(5,2).getValue();  // B5 = PIC QA
      if (ps  && String(ps).trim())  projectSprint = String(ps).trim();
      if (ql  && String(ql).trim())  qaLead        = String(ql).trim(); // Summary overrides Config
      if (pic && String(pic).trim()) picQA         = String(pic).trim();

      const perfVal = summ.getRange(8,13).getValue(); // M8 = Perf result
      if (perfVal && String(perfVal).trim()) perfResult = String(perfVal).trim();

      // Web KPI — row 13, cols A-I (1-9)
      const wKpi = summ.getRange(SUMM_KPI_ROW,1,1,9).getValues()[0];
      wTotal=+wKpi[0]||0; wPassed=+wKpi[1]||0; wFailed=+wKpi[2]||0;
      wBlocked=+wKpi[3]||0; wInProg=+wKpi[4]||0; wTodo=+wKpi[5]||0;
      wPassRate=+wKpi[6]||0; wAutoRate=+wKpi[7]||0; wExecRate=+wKpi[8]||0;

      // API KPI — row 13, cols L-T (12-20)
      const aKpi = summ.getRange(SUMM_KPI_ROW,12,1,9).getValues()[0];
      aTotal=+aKpi[0]||0; aPassed=+aKpi[1]||0; aFailed=+aKpi[2]||0;
      aBlocked=+aKpi[3]||0; aInProg=+aKpi[4]||0; aTodo=+aKpi[5]||0;
      aPassRate=+aKpi[6]||0; aAutoRate=+aKpi[7]||0; aExecRate=+aKpi[8]||0;

      Logger.log(mod.name + ' | wTotal=' + wTotal + ' wPass=' + pct_(wPassRate) + ' | aTotal=' + aTotal + ' aPass=' + pct_(aPassRate));

      // Smoke KPI — cari row secara DINAMIS (v2 — paling robust)
      //
      // STRATEGI 2 LANGKAH:
      //   1. textFinder('SMOKE TEST') → dapat posisi smoke header secara pasti
      //      (tidak bisa nyangkut di STATUS OVERVIEW — section itu tidak punya teks ini)
      //   2. Dari header, scan ke bawah (+1 s/d +5): cari row pertama
      //      di mana col A bernilai NUMBER (bukan text label "TOTAL")
      //      → itu adalah baris values yang sesungguhnya
      //
      // Ini menggantikan pendekatan lama (scan label row dari row 14) yang
      // bisa salah tangkap STATUS OVERVIEW label row jika sheet punya extra rows.
      //
      // Layout Summary smoke section (bisa bervariasi per versi broadcast):
      //   Row +0 : header  "A1. SMOKE TEST — Web / Mobile ..."  ← textFinder
      //   Row +1 : labels  TOTAL PASSED ... PASS RATE EXEC RATE  (text, di-skip)
      //   Row +2 : values  5  2  2  0  0  1  40%  80%  80%       ← target
      //  -- atau versi lama tanpa label row:
      //   Row +0 : header
      //   Row +1 : values  ← langsung ketemu di scan pertama
      //
      // Web (L=col 1): A=Total  G=Pass%  I=Exec%
      // API (R=col12): L=Total  R=Pass%  T=Exec%
      try {
        // Cari smokeValRow secara dinamis:
        //
        // KENAPA TIDAK "cari numeric di col A":
        //   Broadcast Fix D merges seluruh row (col 1-10) dengan formula COUNTIFS
        //   → col A di baris Open Blocker = NUMBER (bug count), bukan smoke total.
        //   Scan numeric akan nyangkut di situ.
        //
        // STRATEGI BENAR: cari LABEL ROW (col A = "TOTAL", col G = "PASS RATE"),
        //   lalu smokeValRow = labelRow + 1.
        //
        // Layout Summary setelah broadcast Fix D:
        //   Row 15: SMOKE TEST header
        //   Row 16: "Open Blocker (Smoke) ↓"   → col A = text, skip
        //   Row 17: =COUNTIFS(...)              → col A = NUMBER (bug count), skip ← jebakan
        //   Row 18: TOTAL PASSED... PASS RATE   → col A = "TOTAL" ← cari ini
        //   Row 19: 52 22 0 0 14 0 40% 80% 80% → smokeValRow = 18+1 = 19 ✓
        //
        // Layout fresh sheet (tanpa broadcast Fix D):
        //   Row 15: SMOKE TEST header
        //   Row 16: TOTAL PASSED... PASS RATE   → col A = "TOTAL" ← cari ini
        //   Row 17: 1 ...                       → smokeValRow = 16+1 = 17 ✓
        //
        // Scan dimulai dari row 14 (setelah Status Overview labels di row 12).

        let smokeLabelRow = -1;
        const SCAN_START = 14, SCAN_LEN = 22;  // scan rows 14–35
        const scanGrid = summ.getRange(SCAN_START, 1, SCAN_LEN, 9).getValues();
        for (let i = 0; i < scanGrid.length; i++) {
          const c1 = String(scanGrid[i][0]).trim().toUpperCase();  // col A
          const c7 = String(scanGrid[i][6]).trim().toUpperCase();  // col G
          if (c1 === 'TOTAL' && c7 === 'PASS RATE') {
            smokeLabelRow = SCAN_START + i;
            break;
          }
        }
        if (smokeLabelRow === -1) throw new Error('Smoke label row not found (TOTAL+PASS RATE) in rows 14-35');
        const smokeValRow = smokeLabelRow + 1;
        Logger.log(mod.name + ' | smokeLabelRow=' + smokeLabelRow + ' smokeValRow=' + smokeValRow);

        // Web / Mobile smoke (col 1–9)
        const wSm = summ.getRange(smokeValRow, 1, 1, 9).getValues()[0];
        wSmokeTotal    = +wSm[0]||0;   // col A = TOTAL
        wSmokePassed   = +wSm[1]||0;   // col B = PASSED
        wSmokeFailed   = +wSm[2]||0;   // col C = FAILED
        wSmokeBlocked  = +wSm[3]||0;   // col D = BLOCKED
        wSmokeInProg   = +wSm[4]||0;   // col E = IN PROG
        wSmokeTodo     = +wSm[5]||0;   // col F = TODO
        wSmokePassRate = +wSm[6]||0;   // col G = PASS RATE
        wSmokeAutoRate = +wSm[7]||0;   // col H = AUTO RATE
        wSmokeExecRate = +wSm[8]||0;   // col I = EXEC RATE

        // API smoke (col 12–20, same layout offset R_=12)
        const aSm = summ.getRange(smokeValRow, 12, 1, 9).getValues()[0];
        aSmokeTotal    = +aSm[0]||0;   // col L = TOTAL
        aSmokePassed   = +aSm[1]||0;   // col M = PASSED
        aSmokePassRate = +aSm[6]||0;   // col R = PASS RATE
        aSmokeAutoRate = +aSm[7]||0;   // col S = AUTO RATE
        aSmokeExecRate = +aSm[8]||0;   // col T = EXEC RATE

        Logger.log(mod.name
            + ' | smokeValRow=' + smokeValRow
            + ' | wSmokeTotal=' + wSmokeTotal
            + ' wSmokePass='  + pct_(wSmokePassRate)
            + ' wSmokeExec='  + pct_(wSmokeExecRate)
            + ' | aTotal='    + aSmokeTotal
            + ' aPass='       + pct_(aSmokePassRate));
      } catch(se) {
        Logger.log('Smoke KPI skip [' + mod.name + ']: ' + se.message);
      }

    } else {
      // Fallback: hitung dari raw sheets
      const wS = getSheetStats_(tcm,tce,'TC'), aS = getSheetStats_(apim,apie,'API');
      ({total:wTotal,passed:wPassed,failed:wFailed,blocked:wBlocked,inprog:wInProg,todo:wTodo,passRate:wPassRate,autoRate:wAutoRate,execRate:wExecRate} = wS);
      ({total:aTotal,passed:aPassed,failed:aFailed,blocked:aBlocked,inprog:aInProg,todo:aTodo,passRate:aPassRate,autoRate:aAutoRate,execRate:aExecRate} = aS);
      perfResult = getPerfResult_(perf);
    }
  } catch(e) {
    Logger.log('pullModuleData_ error [' + mod.name + ']: ' + e.message);
    try {
      const wS = getSheetStats_(tcm,tce,'TC'), aS = getSheetStats_(apim,apie,'API');
      wTotal=wS.total; wPassed=wS.passed; wFailed=wS.failed; wBlocked=wS.blocked; wPassRate=wS.passRate;
      aTotal=aS.total; aPassed=aS.passed; aFailed=aS.failed; aBlocked=aS.blocked; aPassRate=aS.passRate;
    } catch(e2) {}
  }

  return {
    name:mod.name, team:picQA, lead:qaLead, id:mod.id,
    sprint:projectSprint, project:mod.project, module:mod.module, submodule:mod.submodule,
    refreshed:new Date(),
    wTotal,wPassed,wFailed,wBlocked,wInProg,wTodo,wPassRate,wAutoRate,wExecRate,
    aTotal,aPassed,aFailed,aBlocked,aInProg,aTodo,aPassRate,aAutoRate,aExecRate,
    wSmokeTotal,wSmokePassed,wSmokeFailed,wSmokeBlocked,wSmokeInProg,wSmokeTodo,
    wSmokePassRate,wSmokeAutoRate,wSmokeExecRate,
    aSmokeTotal,aSmokePassed,aSmokePassRate,aSmokeAutoRate,aSmokeExecRate,
    perfResult,
    blockers: getBlockers_(tcm,tce,apim,apie,mod.name),
    coverage: getCoverage_(tcm,tce,apim,apie),
    bugStats: getBugStats_(bugr),
    error: '',
  };
}

function emptyModuleData_(mod, errorMsg) {
  return {
    name:mod.name,team:mod.team||'',lead:mod.lead||'',id:mod.id,
    sprint:'',project:mod.project||'',module:mod.module||'',submodule:mod.submodule||'',
    refreshed:new Date(),error:errorMsg,
    wTotal:0,wPassed:0,wFailed:0,wBlocked:0,wInProg:0,wTodo:0,wPassRate:0,wAutoRate:0,wExecRate:0,
    aTotal:0,aPassed:0,aFailed:0,aBlocked:0,aInProg:0,aTodo:0,aPassRate:0,aAutoRate:0,aExecRate:0,
    wSmokeTotal:0,wSmokePassed:0,wSmokeFailed:0,wSmokeBlocked:0,wSmokeInProg:0,wSmokeTodo:0,
    wSmokePassRate:0,wSmokeAutoRate:0,wSmokeExecRate:0,
    aSmokeTotal:0,aSmokePassed:0,aSmokePassRate:0,aSmokeAutoRate:0,aSmokeExecRate:0,
    perfResult:'--',blockers:[],coverage:[],
    bugStats:{total:0,open:0,inprog:0,fixed:0,verified:0,critical:0,high:0,medium:0,low:0,blocker:0},
  };
}


// ═══════════════════════════════════════════════════════════════════════
// STAT HELPERS
// ═══════════════════════════════════════════════════════════════════════

function getSheetStats_(masterSheet, execSheet, type) {
  const empty = {total:0,passed:0,failed:0,blocked:0,inprog:0,todo:0,passRate:0,autoRate:0,execRate:0};
  if (!masterSheet || !execSheet) return empty;
  try {
    const mData = masterSheet.getDataRange().getValues().slice(2);
    const eData = execSheet.getDataRange().getValues().slice(8);
    const statusMap = {};
    eData.forEach(r => { if(r[0]&&r[25]) statusMap[r[0]] = String(r[25]).trim(); });
    const autoCol = type==='TC' ? 7 : 9;
    let total=0,passed=0,failed=0,blocked=0,inprog=0,todo=0,auto=0;
    mData.forEach(r => {
      if (!r[2]) return; total++;
      if (r[autoCol]==='Automated') auto++;
      const st = statusMap[r[2]]||'TODO';
      if (st==='PASSED') passed++;
      else if (st==='FAILED') failed++;
      else if (st==='BLOCKED') blocked++;
      else if (st==='IN PROGRESS') inprog++;
      else todo++;
    });
    return {total,passed,failed,blocked,inprog,todo,
      passRate:total?passed/total:0, autoRate:total?auto/total:0,
      execRate:total?(passed+failed+blocked+inprog)/total:0};
  } catch(e) { return empty; }
}

function getBlockers_(tcm, tce, apim, apie, moduleName) {
  const bl = [];
  try {
    if (tcm&&tce) {
      const sm={}; tce.getDataRange().getValues().slice(8).forEach(r=>{if(r[0]&&r[25])sm[r[0]]=r[25];});
      tcm.getDataRange().getValues().slice(2).forEach(r=>{
        if(!r[2]||(r[4]!=='Critical'&&r[4]!=='High'))return;
        const st=sm[r[2]]||'TODO';
        if(st==='FAILED'||st==='BLOCKED') bl.push({module:moduleName,type:'Web',tcId:r[2],prio:r[4],feature:r[3],scenario:String(r[10]).substring(0,80),status:st});
      });
    }
  } catch(e) {}
  try {
    if (apim&&apie) {
      const sm={}; apie.getDataRange().getValues().slice(8).forEach(r=>{if(r[0]&&r[25])sm[r[0]]=r[25];});
      apim.getDataRange().getValues().slice(2).forEach(r=>{
        if(!r[2]||(r[6]!=='Critical'&&r[6]!=='High'))return;
        const st=sm[r[2]]||'TODO';
        if(st==='FAILED'||st==='BLOCKED') bl.push({module:moduleName,type:'API',tcId:r[2],prio:r[6],feature:r[3],scenario:String(r[12]).substring(0,80),status:st});
      });
    }
  } catch(e) {}
  return bl;
}

function getCoverage_(tcm, tce, apim, apie) {
  const cov=[];
  try {
    if (tcm) {
      const sm={}; (tce?tce.getDataRange().getValues().slice(8):[]).forEach(r=>{if(r[0]&&r[25])sm[r[0]]=r[25];});
      const map={};
      tcm.getDataRange().getValues().slice(2).forEach(r=>{
        if(!r[2]||!r[1])return;
        if(!map[r[1]])map[r[1]]={sub:r[1],total:0,passed:0,failed:0,auto:0,type:'Web'};
        map[r[1]].total++;
        if(r[7]==='Automated')map[r[1]].auto++;
        const st=sm[r[2]]||'';
        if(st==='PASSED')map[r[1]].passed++;
        if(st==='FAILED')map[r[1]].failed++;
      });
      Object.values(map).forEach(v=>cov.push(v));
    }
  } catch(e) {}
  try {
    if (apim) {
      const sm={}; (apie?apie.getDataRange().getValues().slice(8):[]).forEach(r=>{if(r[0]&&r[25])sm[r[0]]=r[25];});
      const map={};
      apim.getDataRange().getValues().slice(2).forEach(r=>{
        if(!r[2]||!r[1])return;
        if(!map[r[1]])map[r[1]]={sub:r[1],total:0,passed:0,failed:0,auto:0,type:'API'};
        map[r[1]].total++;
        if(r[9]==='Automated')map[r[1]].auto++;
        const st=sm[r[2]]||'';
        if(st==='PASSED')map[r[1]].passed++;
        if(st==='FAILED')map[r[1]].failed++;
      });
      Object.values(map).forEach(v=>cov.push(v));
    }
  } catch(e) {}
  return cov;
}

function getPerfResult_(perfSheet) {
  if (!perfSheet) return '--';
  try {
    const data = perfSheet.getDataRange().getValues().slice(15,45);
    if (!data.some(r=>r[4]&&r[4]!=='')) return '--';
    return data.some(r=>r[11]==='FAIL') ? 'FAIL' : 'PASS';
  } catch(e) { return '--'; }
}

function getBugStats_(bugSheet) {
  const empty={total:0,open:0,inprog:0,fixed:0,verified:0,critical:0,high:0,medium:0,low:0,blocker:0};
  if (!bugSheet) return empty;
  try {
    const rows=bugSheet.getDataRange().getValues().slice(4).filter(r=>r[0]&&r[0]!=='');
    const cnt=(fn)=>rows.filter(fn).length;
    return {
      total:rows.length,
      open:    cnt(r=>r[3]==='Open'),
      inprog:  cnt(r=>r[3]==='In Progress'),
      fixed:   cnt(r=>r[3]==='Fixed'),
      verified:cnt(r=>r[3]==='Verified'),
      critical:cnt(r=>r[2]==='Critical'),
      high:    cnt(r=>r[2]==='High'),
      medium:  cnt(r=>r[2]==='Medium'),
      low:     cnt(r=>r[2]==='Low'),
      blocker: cnt(r=>['Open','In Progress','Reopen'].includes(r[3])&&['Critical','High','Medium'].includes(r[2])),
    };
  } catch(e) { return empty; }
}

function pct_(v) { return Math.round((v||0)*100)+'%'; }


// ═══════════════════════════════════════════════════════════════════════
// CONFIG TAB
// ═══════════════════════════════════════════════════════════════════════

function buildConfig(ss) {
  const ws = ss.insertSheet('Config');
  ws.setTabColor('#37474F');
  ws.clear();
  function hdr(r,c,txt,w,note){
    ws.getRange(r,c).setValue(txt).setBackground('#0D47A1').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setColumnWidth(c,w);
    if(note) ws.getRange(r,c).setNote(note);
  }
  ws.getRange(1,1,1,9).merge().setValue('QA PORTFOLIO DASHBOARD  —  Module Config')
      .setBackground('#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,32);
  ws.getRange(2,1,1,9).merge()
      .setValue('Spreadsheet ID ada di URL modul: docs.google.com/spreadsheets/d/[ID]/edit  |  QA Team Lead: isi manual ATAU otomatis dari Summary B4 saat refresh')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic').setFontSize(8).setFontFamily('Arial');
  ws.setRowHeight(2,16);

  hdr(3,1,'Active (Y/N)',80,'Y = aktif di-pull saat refresh\nN = skip');
  hdr(3,2,'Project',120,'Project / inisiatif / client. Contoh: SIPGN, INAGOV');
  hdr(3,3,'Module',120,'Domain dalam project. Kosongkan jika project flat.');
  hdr(3,4,'SubModule',160,'Unit terkecil → 1 spreadsheet QA.');
  hdr(3,5,'PIC / Team',150,'Auto-filled dari Summary B5 saat refresh.');
  hdr(3,6,'QA Team Lead',140,'Auto-filled dari Summary B4 saat refresh.\nBisa diisi manual jika Summary belum ada.');
  hdr(3,7,'Spreadsheet ID',380,'URL: https://docs.google.com/spreadsheets/d/[ID]/edit');
  hdr(3,8,'Link',60);
  hdr(3,9,'Notes',180);
  ws.setRowHeight(3,22);

  [['Y','SIPGN','1 - Manajemen Gizi','1.1 Aplikasi Nutritionist','QA Team','Andi','1evhTCv0gyfsTxkh5SusXvK_GD68HRJkH9QFZB3-jDmg','','Sample - ganti dengan ID aktual'],
    ['Y','SIPGN','1 - Manajemen Gizi','1.2 Aplikasi Courier','QA Team','Budi','PASTE_SPREADSHEET_ID_HERE','',''],
    ['N','INAGOV','','Talenta','QA Team','Citra','PASTE_SPREADSHEET_ID_HERE','','Flat project'],
  ].forEach((row,i)=>{ws.getRange(4+i,1,1,row.length).setValues([row]);ws.setRowHeight(4+i,22);});
}


// ═══════════════════════════════════════════════════════════════════════
// OVERVIEW TAB — build + write + charts
// ═══════════════════════════════════════════════════════════════════════

function buildOverview(ss) {
  const ws = ss.insertSheet('Overview');
  ws.setTabColor('#0D47A1');
  ws.clear();
  initOverviewHeaders_(ws);
  ws.getRange(5,1,1,25).merge()
      .setValue('▶ Run refreshDashboard() untuk mengisi data')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(4);
}

function initOverviewHeaders_(ws) {
  const lastCol = Math.max(ws.getLastColumn()||1, 25);
  try { ws.getRange(1,1,4,lastCol).breakApart(); } catch(e) {}
  ws.getRange(1,1,4,lastCol).clearContent().clearFormat();

  function h_(r,c,nr,nc,txt,bg,fg,sz){
    const rng=(nr>1||nc>1)?ws.getRange(r,c,nr,nc).merge():ws.getRange(r,c);
    rng.setValue(txt||'').setBackground(bg||'#0D47A1').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  }

  // Col widths — 25 cols
  [145,100,100,125, 52,56,52,52,68, 62,68,58, 52,56,52,52,68, 62,68,58, 68, 52,62,56, 165]
      .forEach((w,i)=>ws.setColumnWidth(i+1,w));

  // Row 1 — last refresh
  ws.getRange(1,1,1,25).merge().setValue('Last refreshed: —')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  // Row 2 — title
  h_(2,1,1,25,'QA DASHBOARD  |  PORTFOLIO OVERVIEW','#0D47A1','#FFFFFF',13);
  ws.setRowHeight(2,30);

  // Row 3 — group headers
  h_(3,1, 1,4, 'MODULE INFO',    '#263238');
  h_(3,5, 1,5, 'WEB / MOBILE',  '#1565C0');
  h_(3,10,1,3, '🔥 SMOKE WEB',  '#BF360C');
  h_(3,13,1,5, 'API',            '#283593');
  h_(3,18,1,3, '🔥 SMOKE API',  '#4A148C');
  h_(3,21,1,1, 'PERF',           '#004D40');
  h_(3,22,1,3, 'BUGS',           '#B71C1C');
  h_(3,25,1,1, 'NOTES',          '#37474F');
  ws.setRowHeight(3,22);

  // Row 4 — column headers
  ['SubModule','Project','Module','PIC / Team',
    'Total','Pass','Fail','Block','Pass%',
    'Total','Pass%','Exec%',
    'Total','Pass','Fail','Block','Pass%',
    'Total','Pass%','Exec%',
    'Perf','Bugs','Blocker','Critical','Notes'
  ].forEach((lbl,i)=>h_(4,i+1,1,1,lbl,'#1565C0'));
  ws.getRange(4,10).setNote('Smoke Web: TC Priority Critical+High+Medium');
  ws.getRange(4,11).setNote('Smoke Web Pass Rate (target ≥80%)');
  ws.getRange(4,12).setNote('Smoke Web Exec Rate (% TC sudah ada hasil)');
  ws.getRange(4,18).setNote('Smoke API: TC Priority Critical+High+Medium');
  ws.getRange(4,19).setNote('Smoke API Pass Rate (target ≥80%)');
  ws.getRange(4,20).setNote('Smoke API Exec Rate');
  ws.setRowHeight(4,26);
  ws.setFrozenRows(4);
}

function writeOverview(ss, allData) {
  let ws = ss.getSheetByName('Overview');
  if (!ws) { buildOverview(ss); ws = ss.getSheetByName('Overview'); }

  initOverviewHeaders_(ws);  // safe rebuild — breakApart dulu

  const lastRow = Math.max(ws.getLastRow(),5);
  if (lastRow>=5) ws.getRange(5,1,lastRow-4,25).clearContent().clearFormat();

  const rules = [];

  allData.forEach((d,i)=>{
    const r  = 5+i;
    const bg = i%2===0 ? '#F9FAFB' : '#FFFFFF';
    const bs = d.bugStats||{};
    const hasSmoke = d.wSmokeTotal>0||d.aSmokeTotal>0;

    function cell(col,val,fmt){
      const c=ws.getRange(r,col).setValue(val==null?'':val).setBackground(bg)
          .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
      if(fmt)c.setNumberFormat(fmt);
      return c;
    }

    ws.getRange(r,1).setValue(d.name).setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    cell(2,d.project||d.sprint||'');
    cell(3,d.module||'');
    cell(4,d.team||'');

    // Web
    cell(5,d.wTotal); cell(6,d.wPassed); cell(7,d.wFailed); cell(8,d.wBlocked);
    cell(9,d.error?'ERR':d.wPassRate,'0%');

    // Smoke Web
    cell(10,hasSmoke?d.wSmokeTotal:'--');
    cell(11,hasSmoke?d.wSmokePassRate:'--',hasSmoke?'0%':null);
    cell(12,hasSmoke?d.wSmokeExecRate:'--',hasSmoke?'0%':null);

    // API
    cell(13,d.aTotal); cell(14,d.aPassed); cell(15,d.aFailed); cell(16,d.aBlocked);
    cell(17,d.error?'ERR':d.aPassRate,'0%');

    // Smoke API
    cell(18,hasSmoke?d.aSmokeTotal:'--');
    cell(19,hasSmoke?d.aSmokePassRate:'--',hasSmoke?'0%':null);
    cell(20,hasSmoke?d.aSmokeExecRate:'--',hasSmoke?'0%':null);

    cell(21,d.perfResult);
    cell(22,bs.total||0); cell(23,bs.blocker||0); cell(24,bs.critical||0);

    ws.getRange(r,25).setValue(d.error||'').setBackground(bg).setFontFamily('Arial').setFontSize(8)
        .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
        .setBorder(true,true,true,true,false,false,'#E0E0E0',SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(r,22);

    // RAG Pass%
    [9,17].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.8,0.5)));
    // RAG Smoke Pass%
    [11,19].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.8,0.5)));
    // RAG Smoke Exec%
    [12,20].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.7,0.4)));
    // Failed > 0
    [7,15].forEach(col=>rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFCDD2').setFontColor('#C62828').setBold(true)
        .setRanges([ws.getRange(r,col)]).build()));
    // Blocked > 0
    [8,16].forEach(col=>rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
        .setRanges([ws.getRange(r,col)]).build()));
    // Blocker/Critical > 0
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
        .setRanges([ws.getRange(r,23),ws.getRange(r,24)]).build());
    // Perf
    [['PASS','#C8E6C9','#1B5E20'],['FAIL','#FFCDD2','#C62828'],['--','#F5F5F5','#9E9E9E']]
        .forEach(([v,bg2,fg])=>rules.push(SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(v).setBackground(bg2).setFontColor(fg).setBold(true)
            .setRanges([ws.getRange(r,21)]).build()));
  });

  // TOTAL row
  if (allData.length > 0) {
    const tr = 5+allData.length;
    ws.getRange(tr,1,1,4).merge().setValue('TOTAL / AVERAGE')
        .setBackground('#E3F2FD').setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    [[5,'wTotal'],[6,'wPassed'],[7,'wFailed'],[8,'wBlocked'],
      [13,'aTotal'],[14,'aPassed'],[15,'aFailed'],[16,'aBlocked']].forEach(([col,key])=>{
      ws.getRange(tr,col).setValue(allData.reduce((a,d)=>a+(d[key]||0),0))
          .setBackground('#DDEEFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    });
    const avg=(key)=>allData.reduce((a,d)=>a+(d[key]||0),0)/allData.length;
    [[9,'wPassRate'],[17,'aPassRate'],[11,'wSmokePassRate'],[19,'aSmokePassRate']].forEach(([col,key])=>
        ws.getRange(tr,col).setValue(avg(key)).setNumberFormat('0%')
            .setBackground(col>=10&&col<=12||col>=18&&col<=20?'#FFF3E0':'#DDEEFF')
            .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center'));
    [[22,'total'],[23,'blocker'],[24,'critical']].forEach(([col,key])=>
        ws.getRange(tr,col).setValue(allData.reduce((a,d)=>a+((d.bugStats||{})[key]||0),0))
            .setBackground('#DDEEFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center'));
    ws.setRowHeight(tr,22);
  }

  ws.setConditionalFormatRules(rules);
  buildOverviewCharts_(ws, allData);
}

function buildOverviewCharts_(ws, allData) {
  if (!allData||allData.length===0) return;
  ws.getCharts().forEach(c=>ws.removeChart(c));
  const n=allData.length, dRow=5, cRow=dRow+n+4;

  tryChart_(()=>ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(4,1,n+1,1))
      .addRange(ws.getRange(4,9,n+1,1))    // Web Pass%
      .addRange(ws.getRange(4,17,n+1,1))   // API Pass%
      .setPosition(cRow,1,0,0)
      .setOption('title','Pass Rate — Web vs API')
      .setOption('hAxis',{title:'Pass Rate',format:'#%',minValue:0,maxValue:1})
      .setOption('colors',['#1565C0','#283593'])
      .setOption('legend',{position:'top'})
      .setOption('width',460).setOption('height',270).build()));

  tryChart_(()=>ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(4,1,n+1,1))
      .addRange(ws.getRange(4,11,n+1,1))   // Smoke Web Pass%
      .addRange(ws.getRange(4,19,n+1,1))   // Smoke API Pass%
      .setPosition(cRow,9,0,0)
      .setOption('title','🔥 Smoke Pass Rate — Web vs API')
      .setOption('hAxis',{title:'Pass Rate',format:'#%',minValue:0,maxValue:1})
      .setOption('colors',['#BF360C','#4A148C'])
      .setOption('legend',{position:'top'})
      .setOption('width',460).setOption('height',270).build()));

  tryChart_(()=>ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(ws.getRange(4,1,n+1,1))
      .addRange(ws.getRange(4,6,n+1,1))    // Web Passed
      .addRange(ws.getRange(4,7,n+1,1))    // Web Failed
      .addRange(ws.getRange(4,8,n+1,1))    // Web Blocked
      .setPosition(cRow+20,1,0,0)
      .setOption('title','Web/Mobile TC Status per Module')
      .setOption('isStacked',true)
      .setOption('colors',['#4CAF50','#F44336','#FF9800'])
      .setOption('legend',{position:'top'})
      .setOption('width',460).setOption('height',250).build()));

  if (allData.some(d=>(d.bugStats||{}).blocker>0))
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(ws.getRange(4,1,n+1,1))
        .addRange(ws.getRange(4,23,n+1,1)) // Blocker
        .addRange(ws.getRange(4,24,n+1,1)) // Critical
        .setPosition(cRow+20,9,0,0)
        .setOption('title','🚨 Open Blocker & Critical Bugs')
        .setOption('colors',['#FF9800','#F44336'])
        .setOption('legend',{position:'top'})
        .setOption('width',460).setOption('height',250).build()));
}


// ═══════════════════════════════════════════════════════════════════════
// SMOKE TAB — dedicated smoke view + 5 charts
// ═══════════════════════════════════════════════════════════════════════

function buildSmoke(ss) {
  const ws = ss.insertSheet('Smoke');
  ws.setTabColor('#BF360C');
  ws.clear();
  initSmokeHeaders_(ws);
  ws.getRange(5,1,1,13).merge()
      .setValue('▶ Run refreshDashboard() untuk mengisi data')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(4);
}

function initSmokeHeaders_(ws) {
  function h_(r,c,nr,nc,txt,bg,fg,sz){
    const rng=(nr>1||nc>1)?ws.getRange(r,c,nr,nc).merge():ws.getRange(r,c);
    rng.setValue(txt||'').setBackground(bg||'#BF360C').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#FFCCBC',SpreadsheetApp.BorderStyle.SOLID);
  }
  const lastCol=Math.max(ws.getLastColumn()||1,13);
  try{ws.getRange(1,1,4,lastCol).breakApart();}catch(e){}
  ws.getRange(1,1,4,lastCol).clearContent().clearFormat();

  [148,105,145,140, 62,70,60, 62,70,60, 62,62, 175]
      .forEach((w,i)=>ws.setColumnWidth(i+1,w));

  ws.getRange(1,1,1,13).merge().setValue('Last refreshed: —')
      .setBackground('#FBE9E7').setFontColor('#BF360C').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  h_(2,1,1,13,'🔥  SMOKE TEST DASHBOARD  —  Priority: Critical + High + Medium','#BF360C','#FFFFFF',13);
  ws.setRowHeight(2,30);

  h_(3,1, 1,4,'MODULE INFO',           '#263238');
  h_(3,5, 1,3,'SMOKE WEB / MOBILE',    '#BF360C');
  h_(3,8, 1,3,'SMOKE API',             '#4A148C');
  h_(3,11,1,2,'OPEN BLOCKER (BUG)',    '#B71C1C');
  h_(3,13,1,1,'STATUS',                '#37474F');
  ws.setRowHeight(3,20);

  ['SubModule','Project','PIC / Team','QA Lead',
    'Total','Pass%','Exec%',
    'Total','Pass%','Exec%',
    'Web Bug','API Bug','Smoke Status'
  ].forEach((lbl,i)=>h_(4,i+1,1,1,lbl,'#1565C0'));
  ws.getRange(4,4).setNote('QA Team Lead\nAuto-filled dari Config / Summary B4');
  ws.getRange(4,5).setNote('Total TC Smoke Web (Critical+High+Medium)');
  ws.getRange(4,6).setNote('Pass Rate Smoke Web (target ≥80%)');
  ws.getRange(4,7).setNote('Exec Rate Smoke Web (% TC sudah ada hasil)');
  ws.getRange(4,8).setNote('Total TC Smoke API (Critical+High+Medium)');
  ws.getRange(4,9).setNote('Pass Rate Smoke API (target ≥80%)');
  ws.getRange(4,11).setNote('Bug Open/InProg/Reopen priority Med-Critical (Web type)');
  ws.getRange(4,12).setNote('Bug Open/InProg/Reopen priority Critical');
  ws.setRowHeight(4,26);
  ws.setFrozenRows(4);
}

function writeSmoke(ss, allData) {
  let ws = ss.getSheetByName('Smoke');
  if (!ws) { buildSmoke(ss); ws = ss.getSheetByName('Smoke'); }

  initSmokeHeaders_(ws);

  const lastRow=Math.max(ws.getLastRow(),5);
  if (lastRow>=5) ws.getRange(5,1,lastRow-4,13).clearContent().clearFormat();

  const rules=[];

  allData.forEach((d,i)=>{
    const r=5+i, bg=i%2===0?'#FFF8F6':'#FFFFFF';
    const bs=d.bugStats||{};
    const hasW=d.wSmokeTotal>0, hasA=d.aSmokeTotal>0;

    function cell(col,val,fmt){
      const c=ws.getRange(r,col).setValue(val==null?'':val).setBackground(bg)
          .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
          .setBorder(true,true,true,true,false,false,'#FFCCBC',SpreadsheetApp.BorderStyle.SOLID);
      if(fmt)c.setNumberFormat(fmt);
      return c;
    }

    ws.getRange(r,1).setValue(d.name).setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#FFCCBC',SpreadsheetApp.BorderStyle.SOLID);
    cell(2,d.project||d.sprint||'');
    cell(3,d.team||'');
    cell(4,d.lead||'');        // QA Team Lead

    cell(5, hasW?d.wSmokeTotal:'--');
    cell(6, hasW?d.wSmokePassRate:'--', hasW?'0%':null);
    cell(7, hasW?d.wSmokeExecRate:'--', hasW?'0%':null);

    cell(8, hasA?d.aSmokeTotal:'--');
    cell(9, hasA?d.aSmokePassRate:'--', hasA?'0%':null);
    cell(10,hasA?d.aSmokeExecRate:'--', hasA?'0%':null);

    cell(11,bs.blocker||0);
    cell(12,bs.critical||0);

    // Smoke Status badge
    const smokePct = hasW ? d.wSmokePassRate : (hasA ? d.aSmokePassRate : null);
    const status = d.error           ? '❌ Error'
        : smokePct===null    ? '⬜ Belum ada'
            : smokePct>=0.8      ? '✅ Ready'
                : smokePct>=0.5      ? '⚠️ Perlu perhatian'
                    : '🚨 Blocker';
    const stBg   = d.error?'#FFCDD2':smokePct===null?'#F5F5F5':smokePct>=0.8?'#C8E6C9':smokePct>=0.5?'#FFF9C4':'#FFCDD2';
    const stFg   = d.error?'#C62828':smokePct===null?'#757575':smokePct>=0.8?'#1B5E20':smokePct>=0.5?'#E65100':'#B71C1C';
    ws.getRange(r,13).setValue(status).setBackground(stBg).setFontColor(stFg)
        .setFontFamily('Arial').setFontSize(8).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#FFCCBC',SpreadsheetApp.BorderStyle.SOLID);

    ws.setRowHeight(r,22);

    [6,9].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.8,0.5)));
    [7,10].forEach(col=>rules.push(...ragRules_(ws.getRange(r,col),0.7,0.4)));
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0).setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true)
        .setRanges([ws.getRange(r,11),ws.getRange(r,12)]).build());
  });

  // Summary row
  if (allData.length>0) {
    const tr=5+allData.length;
    ws.getRange(tr,1,1,4).merge().setValue('RATA-RATA / TOTAL')
        .setBackground('#FCE4EC').setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    const avg=(key)=>allData.reduce((a,d)=>a+(d[key]||0),0)/allData.length;
    const sum=(key)=>allData.reduce((a,d)=>a+(d[key]||0),0);
    [[5,sum('wSmokeTotal'),null],[6,avg('wSmokePassRate'),'0%'],
      [8,sum('aSmokeTotal'),null],[9,avg('aSmokePassRate'),'0%'],
      [11,sum(d=>(d.bugStats||{}).blocker||0),null],[12,sum(d=>(d.bugStats||{}).critical||0),null]
    ].forEach(([col,val,fmt])=>{
      const c=ws.getRange(tr,col).setValue(typeof val==='function'?allData.reduce((a,d)=>a+val(d),0):val)
          .setBackground('#FCE4EC').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
      if(fmt)c.setNumberFormat(fmt);
    });
    ws.setRowHeight(tr,22);
  }

  ws.setConditionalFormatRules(rules);
  buildSmokeCharts_(ws, allData);
}

function buildSmokeCharts_(ws, allData) {
  if (!allData||allData.length===0) return;
  ws.getCharts().forEach(c=>ws.removeChart(c));
  const n=allData.length, dRow=5, cRow=dRow+n+4, tmp=15;

  // Chart 1 — Smoke Pass Rate per Module (Web + API)
  tryChart_(()=>ws.insertChart(ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(dRow,1,n,1))
      .addRange(ws.getRange(dRow,6,n,1))    // Smoke Web Pass%
      .addRange(ws.getRange(dRow,9,n,1))    // Smoke API Pass%
      .setPosition(cRow,1,0,0)
      .setOption('title','🔥 Smoke Pass Rate per Module (Web & API)')
      .setOption('hAxis',{title:'Pass Rate',format:'#%',minValue:0,maxValue:1})
      .setOption('series',{0:{color:'#BF360C',labelInLegend:'Web/Mobile'},1:{color:'#4A148C',labelInLegend:'API'}})
      .setOption('legend',{position:'top'})
      .setOption('chartArea',{left:150,top:40,right:20,bottom:30})
      .setOption('width',510).setOption('height',Math.max(230,n*32+90)).build()));

  // Chart 2 — Smoke Total TC distribution (donut)
  const d2=[['Module','Smoke Total']];
  allData.forEach(d=>{if(d.wSmokeTotal>0)d2.push([d.name,d.wSmokeTotal]);});
  if (d2.length>1) {
    d2.forEach((row,ri)=>ws.getRange(cRow+ri,tmp,1,2).setValues([row]));
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(ws.getRange(cRow,tmp,d2.length,2))
        .setPosition(cRow,9,0,0)
        .setOption('title','Smoke TC Distribution (Web)')
        .setOption('pieHole',0.45).setOption('pieSliceText','percentage')
        .setOption('legend',{position:'right'})
        .setOption('width',340).setOption('height',220).build()));
  }

  // Chart 3 — Smoke Status Breakdown (stacked column)
  const d3=[['Module','Passed','Failed','Blocked','In Prog','Todo']];
  allData.forEach(d=>{if(d.wSmokeTotal>0)d3.push([d.name,d.wSmokePassed||0,d.wSmokeFailed||0,d.wSmokeBlocked||0,d.wSmokeInProg||0,d.wSmokeTodo||0]);});
  if (d3.length>1) {
    const sRow=cRow+Math.max(allData.length,3)+3;
    d3.forEach((row,ri)=>ws.getRange(sRow+ri,tmp,1,6).setValues([row]));
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(ws.getRange(sRow,tmp,d3.length,6))
        .setPosition(cRow+18,1,0,0)
        .setOption('title','Smoke TC Status Breakdown (Web/Mobile)')
        .setOption('isStacked',true)
        .setOption('series',{0:{color:'#4CAF50'},1:{color:'#F44336'},2:{color:'#FF9800'},3:{color:'#2196F3'},4:{color:'#9E9E9E'}})
        .setOption('legend',{position:'top'})
        .setOption('width',510).setOption('height',260).build()));
  }

  // Chart 4 — Open Blocker Bugs
  if (allData.some(d=>(d.bugStats||{}).blocker>0))
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(ws.getRange(dRow,1,n,1))
        .addRange(ws.getRange(dRow,11,n,1)) // Web blocker
        .setPosition(cRow+18,9,0,0)
        .setOption('title','🚨 Open Blocker Bugs per Module')
        .setOption('hAxis',{title:'Jumlah Bug',minValue:0})
        .setOption('colors',['#B71C1C'])
        .setOption('legend',{position:'top'})
        .setOption('width',340).setOption('height',260).build()));

  // Chart 5 — Exec Rate vs target 100%
  const d5=[['Module','Exec%','Target']];
  allData.forEach(d=>{if(d.wSmokeTotal>0)d5.push([d.name,d.wSmokeExecRate,1]);});
  if (d5.length>1) {
    const eRow=cRow+Math.max(allData.length,3)*2+6;
    d5.forEach((row,ri)=>ws.getRange(eRow+ri,tmp,1,3).setValues([row]));
    tryChart_(()=>ws.insertChart(ws.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(ws.getRange(eRow,tmp,d5.length,3))
        .setPosition(cRow+36,1,0,0)
        .setOption('title','Smoke Exec Rate vs Target 100%')
        .setOption('vAxis',{format:'#%',minValue:0,maxValue:1,title:'Exec Rate'})
        .setOption('series',{0:{color:'#0288D1',labelInLegend:'Exec Rate'},1:{color:'#E53935',type:'line',labelInLegend:'Target 100%',lineWidth:2}})
        .setOption('legend',{position:'top'})
        .setOption('width',510).setOption('height',240).build()));
  }
}


// ═══════════════════════════════════════════════════════════════════════
// BLOCKERS TAB
// ═══════════════════════════════════════════════════════════════════════

function buildBlockers(ss) {
  const ws=ss.insertSheet('Blockers'); ws.setTabColor('#B71C1C'); ws.clear();
  function h_(c,txt){ws.getRange(2,c).setValue(txt).setBackground('#B71C1C').setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#E57373',SpreadsheetApp.BorderStyle.SOLID);}
  [120,80,85,75,75,100,250,85].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  ws.getRange(1,1,1,8).merge().setValue('BLOCKER ALERT  —  Critical & High  |  Status: FAILED / BLOCKED')
      .setBackground('#B71C1C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(12).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,28);
  ['Modul','Type','TC_ID','Priority','Status','Feature','Scenario','Refresh'].forEach((t,i)=>h_(i+1,t));
  ws.setRowHeight(2,20);
  ws.getRange(3,1,1,8).merge().setValue('▶ Run refreshDashboard()')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(2);
}

function writeBlockers(ss, allData) {
  const ws=ss.getSheetByName('Blockers'); if(!ws)return;
  const lastRow=Math.max(ws.getLastRow(),3);
  if(lastRow>=3)ws.getRange(3,1,lastRow-2,8).clearContent().clearFormat();
  const all=[];
  allData.forEach(d=>d.blockers.forEach(b=>all.push({...b,refreshed:d.refreshed})));
  if(all.length===0){
    ws.getRange(3,1,1,8).merge().setValue('✅ Tidak ada blocker! Semua Critical & High TC passed.')
        .setBackground('#C8E6C9').setFontColor('#1B5E20').setFontWeight('bold')
        .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.setRowHeight(3,28);
  } else {
    all.sort((a,b)=>{if(a.status!==b.status)return a.status==='FAILED'?-1:1;if(a.prio!==b.prio)return a.prio==='Critical'?-1:1;return a.module.localeCompare(b.module);});
    const rules=[];
    all.forEach((b,i)=>{
      const r=3+i,bg=i%2===0?'#FFF8F8':'#FFFFFF';
      [b.module,b.type,b.tcId,b.prio,b.status,b.feature,b.scenario,
        Utilities.formatDate(b.refreshed,Session.getScriptTimeZone(),'dd/MM HH:mm')
      ].forEach((v,ci)=>ws.getRange(r,ci+1).setValue(v).setBackground(bg).setFontFamily('Arial').setFontSize(9)
          .setHorizontalAlignment(ci>1&&ci<6?'center':'left').setVerticalAlignment('middle').setWrap(ci===6)
          .setBorder(true,true,true,true,false,false,'#E57373',SpreadsheetApp.BorderStyle.SOLID));
      ws.setRowHeight(r,20);
      const sR=ws.getRange(r,5),pR=ws.getRange(r,4);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FAILED').setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true).setRanges([sR]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('BLOCKED').setBackground('#FFE0B2').setFontColor('#E65100').setBold(true).setRanges([sR]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Critical').setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true).setRanges([pR]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('High').setBackground('#FFE0B2').setFontColor('#E65100').setBold(true).setRanges([pR]).build());
    });
    ws.setConditionalFormatRules(rules);
  }
  ws.getRange(1,1,1,8).clear();
  ws.getRange(1,1,1,8).merge().setValue(
      'BLOCKER ALERT  |  Total:'+all.length+'  FAILED:'+all.filter(b=>b.status==='FAILED').length+
      '  BLOCKED:'+all.filter(b=>b.status==='BLOCKED').length+
      '  Critical:'+all.filter(b=>b.prio==='Critical').length+
      '  High:'+all.filter(b=>b.prio==='High').length
  ).setBackground('#B71C1C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
}


// ═══════════════════════════════════════════════════════════════════════
// COVERAGE TAB
// ═══════════════════════════════════════════════════════════════════════

function buildCoverage(ss) {
  const ws=ss.insertSheet('Coverage'); ws.setTabColor('#1B5E20'); ws.clear();
  [120,90,60,60,60,60,65].forEach((w,i)=>ws.setColumnWidth(i+1,w));
  ws.getRange(1,1,1,7).merge().setValue('COVERAGE PER SUBMODUL  —  All Modules')
      .setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(12).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,28);
  ['Modul','SubModul','Type','Total','Passed','Failed','Auto%'].forEach((h,i)=>
      ws.getRange(2,i+1).setValue(h).setBackground('#2E7D32').setFontColor('#FFFFFF')
          .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center')
          .setBorder(true,true,true,true,false,false,'#81C784',SpreadsheetApp.BorderStyle.SOLID));
  ws.setRowHeight(2,20);
  ws.getRange(3,1,1,7).merge().setValue('▶ Run refreshDashboard()')
      .setBackground('#F1F8E9').setFontColor('#33691E').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(2);
}

function writeCoverage(ss, allData) {
  const ws=ss.getSheetByName('Coverage'); if(!ws)return;
  const lastRow=Math.max(ws.getLastRow(),3);
  if(lastRow>=3)ws.getRange(3,1,lastRow-2,7).clearContent().clearFormat();
  let r=3; const rules=[];
  allData.forEach(d=>{
    if(!d.coverage||d.coverage.length===0)return;
    d.coverage.forEach((cov,i)=>{
      const bg=i%2===0?'#F1F8E9':'#FFFFFF';
      const autoRate=cov.total>0?cov.auto/cov.total:0;
      [d.name,cov.sub,cov.type,cov.total,cov.passed,cov.failed,autoRate].forEach((v,ci)=>{
        const c=ws.getRange(r,ci+1).setValue(v).setBackground(bg).setFontFamily('Arial').setFontSize(9)
            .setHorizontalAlignment(ci<3?'left':'center').setVerticalAlignment('middle')
            .setBorder(true,true,true,true,false,false,'#81C784',SpreadsheetApp.BorderStyle.SOLID);
        if(ci===6)c.setNumberFormat('0%');
      });
      rules.push(...ragRules_(ws.getRange(r,7),0.8,0.5));
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
          .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true).setRanges([ws.getRange(r,6)]).build());
      ws.setRowHeight(r,20); r++;
    });
  });
  ws.setConditionalFormatRules(rules);
}


// ═══════════════════════════════════════════════════════════════════════
// HISTORY TAB — trend data + chart
// ═══════════════════════════════════════════════════════════════════════

function buildHistory(ss) {
  const ws=ss.insertSheet('History'); ws.setTabColor('#4A148C'); ws.clear();
  const hdrs=['Timestamp','SubModule','Project','Module','PIC','QA Lead',
    'wPass%','wExec%','aPass%','aExec%',
    'wSmokePass%','wSmokeExec%','aSmokePass%','aSmokeExec%',
    'Perf','Bugs','Open','Blocker','Critical'];
  ws.getRange(1,1,1,hdrs.length).merge().setValue('HISTORY  —  Trend Data (auto-appended setiap refresh)')
      .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold')
      .setBackground('#6A1B9A').setFontColor('#FFFFFF');
  ws.setFrozenRows(2);
  ws.setColumnWidth(1,130);
  [2,3,4,5,6].forEach(c=>ws.setColumnWidth(c,110));
  for(let c=7;c<=hdrs.length;c++)ws.setColumnWidth(c,72);
}

function appendHistory(ss, allData) {
  const ws=ss.getSheetByName('History'); if(!ws)return;
  const ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm');
  allData.forEach(d=>{
    const bs=d.bugStats||{};
    ws.appendRow([ts,d.submodule||d.name,d.project||'',d.module||'',d.team||'',d.lead||'',
      d.wPassRate,d.wExecRate,d.aPassRate,d.aExecRate,
      d.wSmokePassRate,d.wSmokeExecRate,d.aSmokePassRate,d.aSmokeExecRate,
      d.perfResult,bs.total||0,bs.open||0,bs.blocker||0,bs.critical||0]);
  });
  const lastRow=ws.getLastRow();
  if(lastRow>=3){
    for(const col of [7,8,9,10,11,12,13,14])ws.getRange(3,col,lastRow-2,1).setNumberFormat('0%');
    tryChart_(()=>{
      ws.getCharts().forEach(c=>ws.removeChart(c));
      ws.insertChart(ws.newChart()
          .setChartType(Charts.ChartType.LINE)
          .addRange(ws.getRange(2,1,lastRow-1,1))
          .addRange(ws.getRange(2,7,lastRow-1,1))    // wPass%
          .addRange(ws.getRange(2,9,lastRow-1,1))    // aPass%
          .addRange(ws.getRange(2,11,lastRow-1,1))   // wSmokePass%
          .addRange(ws.getRange(2,13,lastRow-1,1))   // aSmokePass%
          .setPosition(3,21,0,0)
          .setOption('title','Pass Rate Trend Over Time')
          .setOption('curveType','function')
          .setOption('series',{
            0:{color:'#1565C0',labelInLegend:'Web Pass%'},
            1:{color:'#283593',labelInLegend:'API Pass%'},
            2:{color:'#BF360C',labelInLegend:'Smoke Web%',lineDashStyle:[6,3]},
            3:{color:'#4A148C',labelInLegend:'Smoke API%',lineDashStyle:[6,3]}
          })
          .setOption('vAxis',{title:'Pass Rate',format:'#%',minValue:0,maxValue:1})
          .setOption('hAxis',{title:'Refresh Time'})
          .setOption('legend',{position:'top'})
          .setOption('width',720).setOption('height',380).build());
    });
  }
}


// ═══════════════════════════════════════════════════════════════════════
// RAW TAB
// ═══════════════════════════════════════════════════════════════════════

function buildRaw(ss) {
  const ws=ss.insertSheet('_Raw'); ws.setTabColor('#546E7A'); ws.clear();
  ws.getRange(1,1).setValue('Internal cache — jangan edit manual.')
      .setBackground('#546E7A').setFontColor('#FFFFFF').setFontSize(9).setFontFamily('Arial');
  ws.setRowHeight(1,16);
}

function updateRaw(ss, allData) {
  const ws=ss.getSheetByName('_Raw'); if(!ws)return;
  ws.clearContents();
  ws.getRange(1,1).setValue('Refreshed: '+new Date());
  const hdrs=['Modul','PIC','QA Lead','Project','wTotal','wPass','wFail','wBlock','wPass%',
    'aTotal','aPass','aFail','aBlock','aPass%',
    'wSmokeTotal','wSmokePass%','wSmokeExec%',
    'aSmokeTotal','aSmokePass%','aSmokeExec%',
    'Perf','Bugs','Blocker','Error'];
  ws.getRange(2,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold').setBackground('#607D8B').setFontColor('#FFFFFF');
  allData.forEach((d,i)=>{
    const bs=d.bugStats||{};
    ws.getRange(3+i,1,1,hdrs.length).setValues([[
      d.name,d.team,d.lead,d.sprint,
      d.wTotal,d.wPassed,d.wFailed,d.wBlocked,d.wPassRate,
      d.aTotal,d.aPassed,d.aFailed,d.aBlocked,d.aPassRate,
      d.wSmokeTotal,d.wSmokePassRate,d.wSmokeExecRate,
      d.aSmokeTotal,d.aSmokePassRate,d.aSmokeExecRate,
      d.perfResult,bs.total||0,bs.blocker||0,d.error
    ]]);
  });
}


// ═══════════════════════════════════════════════════════════════════════
// UPDATE CONFIG — write back PIC + QA Lead dari Summary
// ═══════════════════════════════════════════════════════════════════════

function updateConfig(ss, allData) {
  const cfg=ss.getSheetByName('Config'); if(!cfg)return;
  const cfgData=cfg.getDataRange().getValues();
  allData.forEach(d=>{
    for(let i=3;i<cfgData.length;i++){
      if(String(cfgData[i][6]).trim()===d.id){
        if(d.team)cfg.getRange(i+1,5).setValue(d.team);  // col E = PIC
        if(d.lead)cfg.getRange(i+1,6).setValue(d.lead);  // col F = QA Team Lead
        break;
      }
    }
  });
}


// ═══════════════════════════════════════════════════════════════════════
// SHARED HELPERS
// ═══════════════════════════════════════════════════════════════════════

function ragRules_(rng, greenMin, yellowMin) {
  return [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(greenMin)
        .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(yellowMin, greenMin-0.001)
        .setBackground('#FFF9C4').setFontColor('#E65100').setBold(true).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(yellowMin)
        .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true).setRanges([rng]).build(),
  ];
}

function tryChart_(fn) {
  try { fn(); } catch(e) { Logger.log('Chart skip: ' + e.message); }
}

function getActiveModules_() {
  return getModuleList_(SpreadsheetApp.getActiveSpreadsheet());
}
