// =============================================================================
//  QA PORTFOLIO DASHBOARD  |  Apps Script Aggregator
//  File ini dijalankan di Google Sheets TERPISAH dari modul sheets
//
//  SETUP:
//  1. Buat Google Sheets baru -> Extensions -> Apps Script
//  2. Paste semua kode ini -> Save -> Run createDashboard()
//  3. Isi tab Config: Modul Name + Spreadsheet ID tiap modul
//  4. Run refreshDashboard() untuk update data
//  5. Opsional: Setup trigger otomatis via setupTrigger()
// =============================================================================


// ?? MAIN ENTRY POINTS ??????????????????????????????????????????????????????
function createDashboard() {
  Logger.log('createDashboard START');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('QA Dashboard');

  ['Config','Overview','Blockers','Coverage','History','_Raw'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) try { ss.deleteSheet(s); } catch(e) {}
  });
  SpreadsheetApp.flush(); Utilities.sleep(800);

  buildConfig(ss);  SpreadsheetApp.flush(); Utilities.sleep(500);
  buildOverview(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  buildBlockers(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  buildCoverage(ss); SpreadsheetApp.flush(); Utilities.sleep(500);
  buildHistory(ss);  SpreadsheetApp.flush(); Utilities.sleep(500);
  buildRaw(ss);      SpreadsheetApp.flush();

  // Delete default Sheet1
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) try { ss.deleteSheet(s1); } catch(e) {}

  SpreadsheetApp.getUi().alert(
      'Dashboard berhasil dibuat!\n\n' +
      '1. Buka tab Config ? isi Modul Name & Spreadsheet ID\n' +
      '2. Run refreshDashboard() untuk load data\n' +
      '3. Run setupTrigger() untuk auto-refresh setiap hari\n\n' +
      'Cara dapat Spreadsheet ID:\n' +
      'Buka modul sheet ? copy dari URL:\n' +
      'docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit'
  );
  Logger.log('createDashboard DONE');
}

function refreshDashboard() {
  Logger.log('refreshDashboard START: ' + new Date());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modules = getModuleList(ss);
  if (modules.length === 0) {
    SpreadsheetApp.getUi().alert('Belum ada modul di Config. Isi dulu tab Config.');
    return;
  }
  Logger.log('Modules: ' + modules.length);

  const allData = [];
  modules.forEach((mod, i) => {
    if (!mod.active) { Logger.log('Skip (inactive): ' + mod.name); return; }
    Logger.log('Pulling: ' + mod.name + ' [' + mod.id + ']');
    try {
      const data = pullModuleData(mod);
      allData.push(data);
      Logger.log('OK: ' + mod.name + ' | Web Pass=' + data.wPassRate + ' API Pass=' + data.aPassRate);
    } catch(e) {
      Logger.log('ERROR ' + mod.name + ': ' + e.message);
      allData.push(emptyModuleData(mod, 'ERROR: ' + e.message));
    }
    Utilities.sleep(200);
  });

  writeOverview(ss, allData);
  writeBlockers(ss, allData);
  writeCoverage(ss, allData);
  appendHistory(ss, allData);
  updateRaw(ss, allData);
  updateConfig(ss, allData);

  // Update last refresh timestamp in Overview
  const ov = ss.getSheetByName('Overview');
  if (ov) {
    ov.getRange(2, 1).setValue('Last refreshed: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss'));
  }
  Logger.log('refreshDashboard DONE');
  try { SpreadsheetApp.getUi().alert('Refresh selesai! ' + allData.length + ' modul di-update.'); } catch(e) {}
}

function setupTrigger() {
  // Delete existing triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshDashboard') ScriptApp.deleteTrigger(t);
  });
  // Create daily trigger at 07:00
  ScriptApp.newTrigger('refreshDashboard')
      .timeBased().everyHours(1).create();
  try { SpreadsheetApp.getUi().alert('Trigger set! Dashboard auto-refresh setiap 1 jam sekali.'); } catch(e) {}
}


// ?? DATA PULLER ????????????????????????????????????????????????????????????
function pullModuleData(mod) {
  const src = SpreadsheetApp.openById(mod.id);

  const tcm  = src.getSheetByName('TC_Master');
  const tce  = src.getSheetByName('TC_Execution');
  const apim = src.getSheetByName('API_Master');
  const apie = src.getSheetByName('API_Execution');
  const perf = src.getSheetByName('PerfTest');
  const summ = src.getSheetByName('Summary');
  const bugr = src.getSheetByName('BugReport');

  // ?? Summary cell map (QA Test Management v38+) ???????????????????????
  // TEST DESCRIPTION: row 1=title, rows 2-9=fields, value col=B(2) left / M(13) right
  //   B2=Project/Sprint  B3=Period    B4=QA Lead   B5=PIC QA
  //   B6=Environment     B7=Issue Tracker           B8=Test Status  B9=Scope
  //   M2=Base URL        M6=Collection URL          M7=Test Status(API)
  //   M8=Perf Test Result
  // STATUS OVERVIEW KPI: row 13 (L=1, R_=12)
  //   Web: A13=Total B13=Passed C13=Failed D13=Blocked E13=InProg F13=TODO
  //        G13=PassRate H13=AutoRate I13=ExecRate
  //   API: L13=Total M13=Passed N13=Failed O13=Blocked P13=InProg Q13=TODO
  //        R13=PassRate S13=AutoRate T13=ExecRate
  // SMOKE TEST STATUS OVERVIEW (if exists): row determined by finding "A2. STATUS OVERVIEW - Smoke Test"
  //   Web: cols A-I (Total, Passed, Failed, Blocked, InProg, Todo, PassRate, AutoRate, ExecRate)
  //   API: cols L-T (same structure)
  const SUMM_KPI_ROW = 13;

  let picQA = mod.team || '';
  let projectSprint = mod.sprint || '';
  let wTotal=0, wPassed=0, wFailed=0, wBlocked=0, wInProg=0, wTodo=0;
  let wPassRate=0, wAutoRate=0, wExecRate=0;
  let aTotal=0, aPassed=0, aFailed=0, aBlocked=0, aInProg=0, aTodo=0;
  let aPassRate=0, aAutoRate=0, aExecRate=0;
  let perfResult = '--';

  // Smoke Test KPIs
  let smokeWTotal=0, smokeWPassed=0, smokeWFailed=0, smokeWBlocked=0;
  let smokeWPassRate=0, smokeWExecRate=0;
  let smokeATotal=0, smokeAPassed=0, smokeAFailed=0, smokeABlocked=0;
  let smokeAPassRate=0, smokeAExecRate=0;

  try {
    if (summ) {
      // Description fields
      const ps  = summ.getRange(2, 2).getValue();
      const pic = summ.getRange(5, 2).getValue();
      if (ps  && String(ps).trim())  projectSprint = String(ps).trim();
      if (pic && String(pic).trim()) picQA         = String(pic).trim();

      // Perf result from Summary M8 (rightFields index 6 = Perf Test Result)
      const perfVal = summ.getRange(8, 13).getValue();
      if (perfVal && String(perfVal).trim()) perfResult = String(perfVal).trim();

      // Web/Mobile KPI from Summary row 13, cols A-I (1-9)
      const wKpi = summ.getRange(SUMM_KPI_ROW, 1, 1, 9).getValues()[0];
      wTotal    = Number(wKpi[0]) || 0;
      wPassed   = Number(wKpi[1]) || 0;
      wFailed   = Number(wKpi[2]) || 0;
      wBlocked  = Number(wKpi[3]) || 0;
      wInProg   = Number(wKpi[4]) || 0;
      wTodo     = Number(wKpi[5]) || 0;
      wPassRate = Number(wKpi[6]) || 0;
      wAutoRate = Number(wKpi[7]) || 0;
      wExecRate = Number(wKpi[8]) || 0;

      // API KPI from Summary row 13, cols L-T (12-20)
      const aKpi = summ.getRange(SUMM_KPI_ROW, 12, 1, 9).getValues()[0];
      aTotal    = Number(aKpi[0]) || 0;
      aPassed   = Number(aKpi[1]) || 0;
      aFailed   = Number(aKpi[2]) || 0;
      aBlocked  = Number(aKpi[3]) || 0;
      aInProg   = Number(aKpi[4]) || 0;
      aTodo     = Number(aKpi[5]) || 0;
      aPassRate = Number(aKpi[6]) || 0;
      aAutoRate = Number(aKpi[7]) || 0;
      aExecRate = Number(aKpi[8]) || 0;

      // Smoke Test KPI (if section exists)
      try {
        const allData = summ.getDataRange().getValues();
        let smokeKpiRow = -1;

        // Find "A2. STATUS OVERVIEW - Smoke Test"
        for (let i = 0; i < allData.length; i++) {
          const cellValue = allData[i][0] ? allData[i][0].toString() : '';
          if (cellValue.includes('A2.') && cellValue.includes('STATUS OVERVIEW') && cellValue.includes('Smoke Test')) {
            // Values are 2 rows below header (header row + labels row + values row)
            smokeKpiRow = i + 3; // Convert to 1-indexed (+1) then skip header (+1) and labels (+1) = +3
            break;
          }
        }

        if (smokeKpiRow > 0) {
          // Web Smoke Test KPI from cols A-I
          const smokeWKpi = summ.getRange(smokeKpiRow, 1, 1, 9).getValues()[0];
          smokeWTotal    = Number(smokeWKpi[0]) || 0;
          smokeWPassed   = Number(smokeWKpi[1]) || 0;
          smokeWFailed   = Number(smokeWKpi[2]) || 0;
          smokeWBlocked  = Number(smokeWKpi[3]) || 0;
          smokeWPassRate = Number(smokeWKpi[6]) || 0;
          smokeWExecRate = Number(smokeWKpi[8]) || 0;

          // API Smoke Test KPI from cols L-T (12-20)
          const smokeAKpi = summ.getRange(smokeKpiRow, 12, 1, 9).getValues()[0];
          smokeATotal    = Number(smokeAKpi[0]) || 0;
          smokeAPassed   = Number(smokeAKpi[1]) || 0;
          smokeAFailed   = Number(smokeAKpi[2]) || 0;
          smokeABlocked  = Number(smokeAKpi[3]) || 0;
          smokeAPassRate = Number(smokeAKpi[6]) || 0;
          smokeAExecRate = Number(smokeAKpi[8]) || 0;

          Logger.log(mod.name + ' | Smoke Test KPI found at row ' + smokeKpiRow + ' | smokeWTotal=' + smokeWTotal + ' smokeWPass%=' + Math.round(smokeWPassRate*100) + '%');
        } else {
          Logger.log(mod.name + ' | Smoke Test STATUS OVERVIEW not found (optional section)');
        }
      } catch(smokeErr) {
        Logger.log(mod.name + ' | Error reading Smoke Test KPI: ' + smokeErr.message);
      }

      Logger.log(mod.name + ' | Summary KPI read OK | wTotal=' + wTotal + ' wPass%=' + Math.round(wPassRate*100) + '% | aTotal=' + aTotal + ' aPass%=' + Math.round(aPassRate*100) + '%');
    } else {
      // Fallback: Summary missing ? calculate from raw sheets
      Logger.log(mod.name + ' | Summary not found, falling back to raw sheet calculation');
      const wStats = getSheetStats(tcm, tce, 'TC');
      const aStats = getSheetStats(apim, apie, 'API');
      wTotal=wStats.total; wPassed=wStats.passed; wFailed=wStats.failed;
      wBlocked=wStats.blocked; wInProg=wStats.inprog; wTodo=wStats.todo;
      wPassRate=wStats.passRate; wAutoRate=wStats.autoRate; wExecRate=wStats.execRate;
      aTotal=aStats.total; aPassed=aStats.passed; aFailed=aStats.failed;
      aBlocked=aStats.blocked; aInProg=aStats.inprog; aTodo=aStats.todo;
      aPassRate=aStats.passRate; aAutoRate=aStats.autoRate; aExecRate=aStats.execRate;
      perfResult = getPerfResult(perf);
    }
  } catch(e) {
    Logger.log('pullModuleData error [' + mod.name + ']: ' + e.message);
    // Fallback on error too
    try {
      const wStats = getSheetStats(tcm, tce, 'TC');
      const aStats = getSheetStats(apim, apie, 'API');
      wTotal=wStats.total; wPassed=wStats.passed; wFailed=wStats.failed;
      wBlocked=wStats.blocked; wInProg=wStats.inprog; wTodo=wStats.todo;
      wPassRate=wStats.passRate; wAutoRate=wStats.autoRate; wExecRate=wStats.execRate;
      aTotal=aStats.total; aPassed=aStats.passed; aFailed=aStats.failed;
      aBlocked=aStats.blocked; aInProg=aStats.inprog; aTodo=aStats.todo;
      aPassRate=aStats.passRate; aAutoRate=aStats.autoRate; aExecRate=aStats.execRate;
    } catch(e2) { Logger.log('Fallback also failed: ' + e2.message); }
  }

  const blockers  = getBlockers(tcm, tce, apim, apie, mod.name);
  const coverage  = getCoverage(tcm, tce, apim, apie);

  return {
    name:        mod.name,
    team:        picQA,
    id:          mod.id,
    sprint:      projectSprint,
    refreshed:   new Date(),
    wTotal, wPassed, wFailed, wBlocked, wInProg, wTodo,
    wPassRate, wAutoRate, wExecRate,
    aTotal, aPassed, aFailed, aBlocked, aInProg, aTodo,
    aPassRate, aAutoRate, aExecRate,
    perfResult,
    blockers,
    coverage,
    bugStats:    getBugStats(bugr),
    // Smoke Test metrics
    smokeWTotal, smokeWPassed, smokeWFailed, smokeWBlocked, smokeWPassRate, smokeWExecRate,
    smokeATotal, smokeAPassed, smokeAFailed, smokeABlocked, smokeAPassRate, smokeAExecRate,
    error:       '',
  };
}

function getSheetStats(masterSheet, execSheet, type) {
  const empty = {total:0,passed:0,failed:0,blocked:0,inprog:0,todo:0,passRate:0,autoRate:0,execRate:0};
  if (!masterSheet || !execSheet) return empty;

  try {
    // ?? Master: count total TC and auto rate ??????????????????????????????
    // TC_Master:  A=No B=SubModul C=TC_ID(2) ... H=Automated(7)
    // API_Master: A=No B=SubModul C=TC_ID(2) ... J=Automated(9)
    // Data starts row 3 = slice(2)
    const mData = masterSheet.getDataRange().getValues();
    const mRows = mData.slice(2).filter(r => r[2] && r[2] !== ''); // col C = TC_ID
    const total  = mRows.length;
    const autoIdx = type === 'TC' ? 7 : 9;
    const automated = mRows.filter(r => r[autoIdx] === 'Automated').length;
    const autoRate  = total > 0 ? automated / total : 0;

    // ?? Execution: read STATUS from col Z (col 26) EXPLICITLY ????????????
    // Root cause of 0%: getDataRange() stops at last col with direct input (often col J).
    // Cols K-Y are empty, col Z has ARRAYFORMULA ? getDataRange() may not include Z.
    // Fix: read col Z directly regardless of getDataRange range.
    // TC_Execution DS=9, API_Execution DS=9
    const DS = 9;
    const lastRow = execSheet.getLastRow();
    const STATUS_COL = 26; // col Z (1-indexed)
    const ID_COL     = 1;  // col A = TC_ID

    if (lastRow < DS) return { total, passed:0, failed:0, blocked:0, inprog:0, todo:0,
      passRate:0, autoRate, execRate:0 };

    const numRows  = lastRow - DS + 1;
    const idVals   = execSheet.getRange(DS, ID_COL,     numRows, 1).getValues();
    const statVals = execSheet.getRange(DS, STATUS_COL, numRows, 1).getValues();

    let passed=0, failed=0, blocked=0, inprog=0, todo=0;
    idVals.forEach((row, i) => {
      if (!row[0] || row[0] === '') return; // skip empty rows
      const st = String(statVals[i][0] || '').trim();
      if (st === 'PASSED')      passed++;
      else if (st === 'FAILED') failed++;
      else if (st === 'BLOCKED') blocked++;
      else if (st === 'IN PROGRESS') inprog++;
      else if (st === 'TODO')   todo++;
    });
    const executed = passed + failed + blocked + inprog;

    return {
      total, passed, failed, blocked, inprog, todo,
      passRate:  total > 0 ? passed   / total : 0,
      autoRate,
      execRate:  total > 0 ? executed / total : 0,
    };
  } catch(e) {
    Logger.log('getSheetStats error: ' + e.message);
    return empty;
  }
}

function getBlockers(tcm, tce, apim, apie, moduleName) {
  const blockers = [];
  if (!tcm || !tce) return blockers;

  try {
    // TC Blockers
    const mData = tcm.getDataRange().getValues();
    const eData = tce.getDataRange().getValues();

    // Build TC_ID ? Latest Status map from execution
    const statusMap = {};
    eData.slice(8).forEach(r => {
      if (r[0] && r[25]) statusMap[r[0]] = r[25]; // col A=TC_ID, col Z=status
    });

    // TC_Master: A=No B=SubModul C=TC_ID D=Feature E=Priority ... K=Scenario
    mData.slice(2).forEach(r => {
      const tcId    = r[2];
      const prio    = r[4]; // col E
      const scenario = r[10]; // col K
      const feature  = r[3];  // col D
      if (!tcId) return;
      if (prio !== 'Critical' && prio !== 'High') return;
      const status = statusMap[tcId] || 'TODO';
      if (status === 'FAILED' || status === 'BLOCKED') {
        blockers.push({ module: moduleName, type: 'Web/Mobile', tcId, prio, feature, scenario: String(scenario).substring(0,80), status });
      }
    });
  } catch(e) { Logger.log('getBlockers TC error: ' + e.message); }

  try {
    // API Blockers
    const amData = apim.getDataRange().getValues();
    const aeData = apie.getDataRange().getValues();
    const aStatusMap = {};
    aeData.slice(8).forEach(r => { if (r[0] && r[25]) aStatusMap[r[0]] = r[25]; });

    // API_Master: A=No B=SubModul C=TC_ID D=Feature E=Method F=URL G=Priority
    amData.slice(2).forEach(r => {
      const tcId    = r[2];
      const prio    = r[6]; // col G
      const feature  = r[3];
      const scenario = r[12]; // col M = Scenario
      if (!tcId) return;
      if (prio !== 'Critical' && prio !== 'High') return;
      const status = aStatusMap[tcId] || 'TODO';
      if (status === 'FAILED' || status === 'BLOCKED') {
        blockers.push({ module: moduleName, type: 'API', tcId, prio, feature, scenario: String(scenario).substring(0,80), status });
      }
    });
  } catch(e) { Logger.log('getBlockers API error: ' + e.message); }

  return blockers;
}

function getCoverage(tcm, tce, apim, apie) {
  const coverage = [];
  if (!tcm) return coverage;

  try {
    const mData = tcm.getDataRange().getValues();
    const eData = tce ? tce.getDataRange().getValues() : [];
    const statusMap = {};
    eData.slice(8).forEach(r => { if (r[0] && r[25]) statusMap[r[0]] = r[25]; });

    // Group by SubModul (col B = index 1)
    const submodMap = {};
    mData.slice(2).forEach(r => {
      const sub = r[1]; const tcId = r[2]; const prio = r[4]; const autoVal = r[7];
      if (!tcId || !sub) return;
      if (!submodMap[sub]) submodMap[sub] = {sub, total:0, passed:0, failed:0, auto:0, type:'Web'};
      submodMap[sub].total++;
      if (autoVal === 'Automated') submodMap[sub].auto++;
      const st = statusMap[tcId] || '';
      if (st === 'PASSED') submodMap[sub].passed++;
      if (st === 'FAILED') submodMap[sub].failed++;
    });
    Object.values(submodMap).forEach(v => coverage.push(v));
  } catch(e) { Logger.log('getCoverage TC error: ' + e.message); }

  try {
    if (!apim) return coverage;
    const amData = apim.getDataRange().getValues();
    const aeData = apie ? apie.getDataRange().getValues() : [];
    const aStatusMap = {};
    aeData.slice(8).forEach(r => { if (r[0] && r[25]) aStatusMap[r[0]] = r[25]; });

    const asubMap = {};
    amData.slice(2).forEach(r => {
      const sub = r[1]; const tcId = r[2]; const prio = r[6]; const autoVal = r[9];
      if (!tcId || !sub) return;
      if (!asubMap[sub]) asubMap[sub] = {sub, total:0, passed:0, failed:0, auto:0, type:'API'};
      asubMap[sub].total++;
      if (autoVal === 'Automated') asubMap[sub].auto++;
      const st = aStatusMap[tcId] || '';
      if (st === 'PASSED') asubMap[sub].passed++;
      if (st === 'FAILED') asubMap[sub].failed++;
    });
    Object.values(asubMap).forEach(v => coverage.push(v));
  } catch(e) { Logger.log('getCoverage API error: ' + e.message); }

  return coverage;
}

function getPerfResult(perfSheet) {
  if (!perfSheet) return '--';
  try {
    const data = perfSheet.getDataRange().getValues();
    // DS=16, MR=30, totRow=46, col L=12 (index 11)
    // Check if any data in col E (index 4) from row 16 (index 15)
    const dataRows = data.slice(15, 45);
    const hasData = dataRows.some(r => r[4] && r[4] !== '');
    if (!hasData) return '--';
    const hasFail = dataRows.some(r => r[11] === 'FAIL');
    return hasFail ? 'FAIL' : 'PASS';
  } catch(e) { return '--'; }
}

function getBugStats(bugSheet) {
  const empty = {total:0,open:0,inprog:0,fixed:0,verified:0,critical:0,high:0,medium:0,low:0,blocker:0};
  if (!bugSheet) return empty;
  try {
    const data = bugSheet.getDataRange().getValues();
    // Row 1=title, 2=note, 3=groups, 4=headers, DS=5
    const rows = data.slice(4).filter(r => r[0] && r[0] !== ''); // col A = Bug ID
    const total    = rows.length;
    const open     = rows.filter(r => r[3] === 'Open').length;          // col D = Status
    const inprog   = rows.filter(r => r[3] === 'In Progress').length;
    const fixed    = rows.filter(r => r[3] === 'Fixed').length;
    const verified = rows.filter(r => r[3] === 'Verified').length;
    const critical = rows.filter(r => r[2] === 'Critical').length;      // col C = Priority
    const high     = rows.filter(r => r[2] === 'High').length;
    const medium   = rows.filter(r => r[2] === 'Medium').length;
    const low      = rows.filter(r => r[2] === 'Low').length;
    // Blocker = Open/In Progress/Reopen bugs with priority Medium, High, OR Critical
    const blocker  = rows.filter(r =>
        ['Open','In Progress','Reopen'].includes(r[3]) &&
        ['Critical','High','Medium'].includes(r[2])
    ).length;
    return {total, open, inprog, fixed, verified, critical, high, medium, low, blocker};
  } catch(e) {
    Logger.log('getBugStats error: ' + e.message);
    return {total:0,open:0,inprog:0,fixed:0,verified:0,critical:0,high:0,medium:0,low:0,blocker:0};
  }
}

function emptyModuleData(mod, errorMsg) {
  return { name: mod.name, team: mod.team||'', id: mod.id, sprint: mod.sprint||'',
    refreshed: new Date(), error: errorMsg,
    wTotal:0,wPassed:0,wFailed:0,wBlocked:0,wInProg:0,wTodo:0,wPassRate:0,wAutoRate:0,wExecRate:0,
    aTotal:0,aPassed:0,aFailed:0,aBlocked:0,aInProg:0,aTodo:0,aPassRate:0,aAutoRate:0,aExecRate:0,
    smokeWTotal:0,smokeWPassed:0,smokeWFailed:0,smokeWBlocked:0,smokeWPassRate:0,smokeWExecRate:0,
    smokeATotal:0,smokeAPassed:0,smokeAFailed:0,smokeABlocked:0,smokeAPassRate:0,smokeAExecRate:0,
    perfResult:'--', blockers:[], coverage:[], bugStats:{total:0,open:0,critical:0,high:0,medium:0,blocker:0} };
}


// ?? CONFIG SHEET ???????????????????????????????????????????????????????????
function buildConfig(ss) {
  const ws = ss.insertSheet('Config');
  ws.setTabColor('#37474F');
  ws.clear();

  function hdr(r,c,txt,bg){ ws.getRange(r,c).setValue(txt).setBackground(bg||'#0D47A1')
      .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle'); }

  // Title
  ws.getRange(1,1,1,7).merge();
  ws.getRange(1,1).setValue('QA PORTFOLIO DASHBOARD  ?  Module Config')
      .setBackground('#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,32);

  // Instruction
  ws.getRange(2,1,1,7).merge();
  ws.getRange(2,1).setValue('Isi baris di bawah ini. Spreadsheet ID ada di URL modul sheet: docs.google.com/spreadsheets/d/[ID]/edit')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic').setFontSize(8).setFontFamily('Arial');
  ws.setRowHeight(2,16);

  // Headers
  ['Active (Y/N)','Modul Name','PIC / Team / Squad','Project / Sprint','Spreadsheet ID','Link','Notes'].forEach((h,i) => {
    hdr(3, i+1, h); ws.setColumnWidth(i+1, [80,140,140,110,380,60,160][i]);
  });
  ws.setRowHeight(3,22);

  // Sample rows
  [
    ['Y','QA-TEMPLATE','QA Team','Sprint 1','1evhTCv0gyfsTxkh5SusXvK_GD68HRJkH9QFZB3-jDmg','','Sample - template modul QA'],
    ['Y','MOD-AUTH','Team Platform','Sprint 12','PASTE_SPREADSHEET_ID_HERE','','Modul Authentication'],
    ['N','MOD-PAYMENT','Team Payment','Sprint 11','PASTE_SPREADSHEET_ID_HERE','','Inactive - next sprint'],
  ].forEach((row,i) => {
    const r = 4+i, bg = i%2===0 ? '#F8F9FA' : '#FFFFFF';
    row.forEach((v,ci) => {
      ws.getRange(r,ci+1).setValue(v).setBackground(bg).setFontFamily('Arial').setFontSize(9)
          .setHorizontalAlignment(ci===0?'center':'left').setVerticalAlignment('middle')
          .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    });
    ws.setRowHeight(r,20);
    // Link formula
    ws.getRange(r,6).setFormula('=IFERROR(HYPERLINK("https://docs.google.com/spreadsheets/d/"&E'+r+',"Open"),"")');
    // Active CF
    const rules = ws.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Y')
        .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true).setRanges([ws.getRange(r,1)]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('N')
        .setBackground('#FFEBEE').setFontColor('#C62828').setRanges([ws.getRange(r,1)]).build());
    ws.setConditionalFormatRules(rules);
  });

  // Add 20 empty rows for user input
  for (let i=7; i<=26; i++) {
    ws.getRange(i,1).setValue('Y').setHorizontalAlignment('center').setBackground('#FAFAFA').setFontFamily('Arial').setFontSize(9);
    ws.getRange(i,2,1,6).setBackground('#FAFAFA');
    ws.setRowHeight(i,20);
  }

  // DV for Active column
  ws.getRange(4,1,23,1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Y','N'],true).build());

  ws.setFrozenRows(3);
}

function getModuleList(ss) {
  const cfg = ss.getSheetByName('Config');
  if (!cfg) return [];
  const data = cfg.getDataRange().getValues();
  const modules = [];
  data.slice(3).forEach(row => {
    const active = String(row[0]).trim().toUpperCase() === 'Y';
    const name   = String(row[1]).trim();
    const team   = String(row[2]).trim();
    const sprint = String(row[3]).trim();
    const id     = String(row[4]).trim();
    if (!name || !id || id === 'PASTE_SPREADSHEET_ID_HERE' || id === '') return;
    modules.push({ active, name, team, sprint, id });
  });
  return modules;
}


// ?? OVERVIEW SHEET ?????????????????????????????????????????????????????????
function buildOverview(ss) {
  const ws = ss.insertSheet('Overview');
  ws.setTabColor('#0D47A1');
  ws.clear();

  function h_(r,c,nr,nc,txt,bg,fg,sz){
    const rng = nr>1||nc>1 ? ws.getRange(r,c,nr,nc).merge() : ws.getRange(r,c);
    return rng.setValue(txt||'').setBackground(bg||'#0D47A1').setFontColor(fg||'#FFFFFF')
        .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
  }

  // Column widths (added Smoke Test columns)
  [130,80,80, 55,60,60,60,65, 55,60,60,60,65, 55,65, 55,65, 70,55,60,60,200].forEach((w,i)=>ws.setColumnWidth(i+1,w));

  // Row 1: last refresh placeholder
  ws.getRange(1,1,1,22).merge();
  ws.getRange(1,1).setValue('Last refreshed: ?')
      .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic')
      .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
  ws.setRowHeight(1,16);

  // Row 2: main title
  h_(2,1,1,22,'QA DASHBOARD  |  PORTFOLIO OVERVIEW','#0D47A1','#FFFFFF',13);
  ws.setRowHeight(2,30);

  // Row 3: group headers
  h_(3,1,1,3,'MODULE INFO','#263238');
  h_(3,4,1,5,'WEB / MOBILE','#1565C0');
  h_(3,9,1,5,'API','#283593');
  h_(3,14,1,2,'SMOKE - WEB','#E65100');
  h_(3,16,1,2,'SMOKE - API','#BF360C');
  h_(3,18,1,1,'PERF','#004D40');
  h_(3,19,1,3,'BUGS','#B71C1C');
  h_(3,22,1,1,'NOTES','#37474F');
  ws.setRowHeight(3,20);

  // Row 4: column headers
  ['Modul','PIC / Team / Squad','Project / Sprint',
    'Total','Passed','Failed','Block','Pass%',
    'Total','Passed','Failed','Block','Pass%',
    'Total','Pass%', 'Total','Pass%',
    'Perf','Bugs','Blocker','Critical','Error / Info'].forEach((h,i) => h_(4,i+1,1,1,h,'#0D47A1'));
  ws.setRowHeight(4,20);

  // Placeholder message row 5
  ws.getRange(5,1,1,15).merge();
  ws.getRange(5,1).setValue('? Run refreshDashboard() untuk mengisi data')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(5,28);

  ws.setFrozenRows(4);
}

function writeOverview(ss, allData) {
  const ws = ss.getSheetByName('Overview');
  if (!ws) return;

  // Clear data rows
  const lastRow = Math.max(ws.getLastRow(), 5);
  if (lastRow >= 5) ws.getRange(5, 1, lastRow-4, 15).clearContent().clearFormat();

  const rules = [];

  allData.forEach((d, i) => {
    const r = 5 + i;
    const bg = i%2===0 ? '#F8F9FA' : '#FFFFFF';

    function cell(col, val, fmt) {
      const c = ws.getRange(r,col).setValue(val).setBackground(bg)
          .setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
      if (fmt) c.setNumberFormat(fmt);
      return c;
    }

    // Module info
    ws.getRange(r,1).setValue(d.name).setBackground(bg).setFontFamily('Arial').setFontSize(9)
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    cell(2, d.team);
    cell(3, d.sprint);

    // Web stats
    cell(4,  d.wTotal);
    cell(5,  d.wPassed);
    cell(6,  d.wFailed);
    cell(7,  d.wBlocked);
    cell(8,  d.error ? 'ERR' : d.wPassRate, '0%');

    // API stats
    cell(9,  d.aTotal);
    cell(10, d.aPassed);
    cell(11, d.aFailed);
    cell(12, d.aBlocked);
    cell(13, d.error ? 'ERR' : d.aPassRate, '0%');

    // Smoke Test - Web
    cell(14, d.smokeWTotal || 0);
    cell(15, d.error ? 'ERR' : (d.smokeWPassRate || 0), '0%');

    // Smoke Test - API
    cell(16, d.smokeATotal || 0);
    cell(17, d.error ? 'ERR' : (d.smokeAPassRate || 0), '0%');

    // Perf
    cell(18, d.perfResult);

    // Bug stats
    const bs = d.bugStats || {};
    cell(19, bs.total  || 0);
    cell(20, bs.blocker|| 0);
    cell(21, bs.critical||0);
    // Notes / error
    ws.getRange(r,22).setValue(d.error || '')
        .setBackground(bg).setFontFamily('Arial').setFontSize(8).setHorizontalAlignment('left')
        .setVerticalAlignment('middle').setWrap(true)
        .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    // Bug CF
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true)
        .setRanges([ws.getRange(r,20),ws.getRange(r,21)]).build());

    ws.setRowHeight(r, 22);

    // RAG CF for Pass% columns (8, 13, 15-Smoke Web, 17-Smoke API)
    [8,13,15,17].forEach(col => {
      const rng = ws.getRange(r,col);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.8)
          .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true).setRanges([rng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.5,0.799)
          .setBackground('#FFF9C4').setFontColor('#E65100').setBold(true).setRanges([rng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.5)
          .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true).setRanges([rng]).build());
    });
    // Perf CF (column now 18 instead of 14)
    ['PASS','FAIL','--'].forEach(v => {
      const perfBg = v==='PASS'?'#C8E6C9':v==='FAIL'?'#FFCDD2':'#F5F5F5';
      const perfFg = v==='PASS'?'#1B5E20':v==='FAIL'?'#C62828':'#9E9E9E';
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
          .setBackground(perfBg).setFontColor(perfFg).setBold(true)
          .setRanges([ws.getRange(r,18)]).build());
    });
    // Failed > 0 highlight
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true)
        .setRanges([ws.getRange(r,6), ws.getRange(r,11)]).build());
    // Blocked > 0
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true)
        .setRanges([ws.getRange(r,7), ws.getRange(r,12)]).build());
  });

  // TOTAL row
  if (allData.length > 0) {
    const tr = 5 + allData.length;
    ws.getRange(tr,1,1,3).merge();
    ws.getRange(tr,1).setValue('TOTAL / AVERAGE').setBackground('#E3F2FD')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    [[4,'wTotal'],[5,'wPassed'],[6,'wFailed'],[7,'wBlocked'],
      [9,'aTotal'],[10,'aPassed'],[11,'aFailed'],[12,'aBlocked']].forEach(([col,key]) => {
      const sum = allData.reduce((acc,d) => acc+(d[key]||0), 0);
      ws.getRange(tr,col).setValue(sum).setBackground('#E3F2FD')
          .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    });
    // Bug totals
    [[15,'total'],[16,'blocker'],[17,'critical']].forEach(([col,key]) => {
      const sum = allData.reduce((acc,d) => acc+((d.bugStats||{})[key]||0), 0);
      ws.getRange(tr,col).setValue(sum).setBackground('#E3F2FD').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    });
    const avgWPass = allData.length>0 ? allData.reduce((a,d)=>a+d.wPassRate,0)/allData.length : 0;
    const avgAPass = allData.length>0 ? allData.reduce((a,d)=>a+d.aPassRate,0)/allData.length : 0;
    ws.getRange(tr,8).setValue(avgWPass).setNumberFormat('0%').setBackground('#E3F2FD').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.getRange(tr,13).setValue(avgAPass).setNumberFormat('0%').setBackground('#E3F2FD').setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.setRowHeight(tr, 22);
  }

  ws.setConditionalFormatRules(rules);

  // ?? Build Charts for Stakeholders ??????????????????????????????????????
  buildOverviewCharts(ws, allData);
}

function buildOverviewCharts(ws, allData) {
  if (!allData || allData.length === 0) return;

  // Remove existing charts
  ws.getCharts().forEach(c => ws.removeChart(c));

  const dataRow = 5; // first data row
  const n = allData.length;
  const totalRow = dataRow + n; // summary row
  const chartStartRow = totalRow + 3; // charts placed below data table

  // ?? Chart 1: Pass Rate Comparison (Clustered Bar) ???????????????????
  // Uses Web Pass% (col 8) and API Pass% (col 13) for each module
  const barChart = ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(ws.getRange(4, 1, n+1, 1))    // Module names
      .addRange(ws.getRange(4, 8, n+1, 1))    // Web Pass%
      .addRange(ws.getRange(4, 13, n+1, 1))   // API Pass%
      .setPosition(chartStartRow, 1, 0, 0)
      .setOption('title', 'Pass Rate per Module (Web vs API)')
      .setOption('hAxis', {title: 'Pass Rate', format: '#%', minValue: 0, maxValue: 1})
      .setOption('colors', ['#1565C0','#283593'])
      .setOption('legend', {position: 'top'})
      .setOption('width', 480).setOption('height', 280)
      .setOption('series', {0:{targetAxisIndex:0}, 1:{targetAxisIndex:0}})
      .build();
  ws.insertChart(barChart);

  // ?? Chart 2: Bug Distribution Pie ???????????????????????????????????
  // Total bugs per module
  const bugPieData = [['Module','Total Bugs']];
  allData.forEach(d => { if ((d.bugStats||{}).total > 0) bugPieData.push([d.name, (d.bugStats||{}).total||0]); });

  if (bugPieData.length > 1) {
    // Write temp data for pie (hidden area, far right)
    const tmpCol = 21;
    bugPieData.forEach((r,i) => ws.getRange(chartStartRow+i, tmpCol, 1, 2).setValues([r]));
    const pieChart = ws.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(ws.getRange(chartStartRow, tmpCol, bugPieData.length, 2))
        .setPosition(chartStartRow, 10, 0, 0)
        .setOption('title', 'Bug Distribution per Module')
        .setOption('pieHole', 0.4)
        .setOption('pieSliceText', 'percentage')
        .setOption('legend', {position: 'right'})
        .setOption('width', 340).setOption('height', 280)
        .build();
    ws.insertChart(pieChart);
  }

  // ?? Chart 3: Exec Rate Bar ???????????????????????????????????????????
  // Web Exec% (col 5 = Passed / Total approximation) vs total
  // Use Passed count (col 5) and Failed (col 6) stacked for execution status
  const stackChart = ws.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(ws.getRange(4, 1, n+1, 1))   // Module names
      .addRange(ws.getRange(4, 5, n+1, 1))   // Web Passed
      .addRange(ws.getRange(4, 6, n+1, 1))   // Web Failed
      .addRange(ws.getRange(4, 7, n+1, 1))   // Web Blocked
      .setPosition(chartStartRow + 20, 1, 0, 0)
      .setOption('title', 'Web/Mobile TC Status per Module')
      .setOption('isStacked', true)
      .setOption('colors', ['#4CAF50','#F44336','#FF9800'])
      .setOption('legend', {position: 'top'})
      .setOption('width', 480).setOption('height', 260)
      .build();
  ws.insertChart(stackChart);

  // ?? Chart 4: Blocker Trend (from History if available) ??????????????
  // Simple: show blocker bug count per module as horizontal bar
  const hasBlockers = allData.some(d => (d.bugStats||{}).blocker > 0);
  if (hasBlockers) {
    const blockerChart = ws.newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(ws.getRange(4, 1, n+1, 1))   // Module names
        .addRange(ws.getRange(4, 16, n+1, 1))  // Blocker count
        .addRange(ws.getRange(4, 17, n+1, 1))  // Critical count
        .setPosition(chartStartRow + 20, 10, 0, 0)
        .setOption('title', 'Open Blockers & Critical Bugs per Module')
        .setOption('colors', ['#FF9800','#F44336'])
        .setOption('legend', {position: 'top'})
        .setOption('width', 340).setOption('height', 260)
        .build();
    ws.insertChart(blockerChart);
  }
}


// ?? BLOCKERS SHEET ??????????????????????????????????????????????????????????
function buildBlockers(ss) {
  const ws = ss.insertSheet('Blockers');
  ws.setTabColor('#B71C1C');
  ws.clear();

  function h_(r,c,txt,bg){ ws.getRange(r,c).setValue(txt).setBackground(bg||'#B71C1C')
      .setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true,true,true,true,false,false,'#E57373',SpreadsheetApp.BorderStyle.SOLID); }

  [120,85,85,75,75,100,250,85].forEach((w,i)=>ws.setColumnWidth(i+1,w));

  ws.getRange(1,1,1,8).merge();
  ws.getRange(1,1).setValue('BLOCKER ALERT  ?  CRITICAL & HIGH  |  Status: FAILED / BLOCKED')
      .setBackground('#B71C1C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(12).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,28);

  ['Modul','Type','TC_ID','Priority','Status','Feature','Scenario','Refresh'].forEach((h,i)=>h_(2,i+1,h));
  ws.setRowHeight(2,20);

  ws.getRange(3,1,1,8).merge();
  ws.getRange(3,1).setValue('? Run refreshDashboard() untuk mengisi data')
      .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');

  ws.setFrozenRows(2);
}

function writeBlockers(ss, allData) {
  const ws = ss.getSheetByName('Blockers');
  if (!ws) return;
  const lastRow = Math.max(ws.getLastRow(), 3);
  if (lastRow >= 3) ws.getRange(3, 1, lastRow-2, 8).clearContent().clearFormat();

  const allBlockers = [];
  allData.forEach(d => d.blockers.forEach(b => allBlockers.push({...b, refreshed: d.refreshed})));

  // Also add open/in-progress bugs Medium?Critical from all modules
  allData.forEach(d => {
    if (!d.bugStats) return;
    // We show count summary ? full list would require pulling BugReport data again
    // This is annotated in Overview col 16 (Blocker bugs count)
  });

  if (allBlockers.length === 0) {
    ws.getRange(3,1,1,8).merge();
    ws.getRange(3,1).setValue('Tidak ada blocker! Semua Critical & High TC passed.')
        .setBackground('#C8E6C9').setFontColor('#1B5E20').setFontWeight('bold')
        .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.setRowHeight(3,28);
    return;
  }

  // Sort: FAILED first, then BLOCKED; Critical before High
  allBlockers.sort((a,b)=>{
    if (a.status!==b.status) return a.status==='FAILED'?-1:1;
    if (a.prio!==b.prio) return a.prio==='Critical'?-1:1;
    return a.module.localeCompare(b.module);
  });

  const rules = [];
  allBlockers.forEach((b,i) => {
    const r = 3+i, bg = i%2===0?'#FFF8F8':'#FFFFFF';
    [b.module,b.type,b.tcId,b.prio,b.status,b.feature,b.scenario,
      Utilities.formatDate(b.refreshed,Session.getScriptTimeZone(),'dd/MM HH:mm')
    ].forEach((v,ci) => {
      ws.getRange(r,ci+1).setValue(v).setBackground(bg).setFontFamily('Arial').setFontSize(9)
          .setHorizontalAlignment(ci>1&&ci<6?'center':'left').setVerticalAlignment('middle').setWrap(ci===6)
          .setBorder(true,true,true,true,false,false,'#E57373',SpreadsheetApp.BorderStyle.SOLID);
    });
    ws.setRowHeight(r,20);
    // Status CF
    const sRng = ws.getRange(r,5);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('FAILED')
        .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true).setRanges([sRng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('BLOCKED')
        .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true).setRanges([sRng]).build());
    // Priority CF
    const pRng = ws.getRange(r,4);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Critical')
        .setBackground('#FFCDD2').setFontColor('#B71C1C').setBold(true).setRanges([pRng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('High')
        .setBackground('#FFE0B2').setFontColor('#E65100').setBold(true).setRanges([pRng]).build());
  });
  ws.setConditionalFormatRules(rules);

  // Summary count at top
  ws.getRange(1,1,1,8).clear();
  ws.getRange(1,1,1,8).merge();
  ws.getRange(1,1).setValue(
      'BLOCKER ALERT  ?  Total: ' + allBlockers.length +
      '  |  FAILED: ' + allBlockers.filter(b=>b.status==='FAILED').length +
      '  |  BLOCKED: ' + allBlockers.filter(b=>b.status==='BLOCKED').length +
      '  |  Critical: ' + allBlockers.filter(b=>b.prio==='Critical').length +
      '  |  High: ' + allBlockers.filter(b=>b.prio==='High').length
  ).setBackground('#B71C1C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
}


// ?? COVERAGE SHEET ??????????????????????????????????????????????????????????
function buildCoverage(ss) {
  const ws = ss.insertSheet('Coverage');
  ws.setTabColor('#1B5E20');
  ws.clear();

  [120,90,60,60,60,65,65].forEach((w,i)=>ws.setColumnWidth(i+1,w));

  ws.getRange(1,1,1,7).merge();
  ws.getRange(1,1).setValue('COVERAGE PER SUBMODUL  ?  All Modules')
      .setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(12).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,28);

  ['Modul','SubModul','Type','Total','Passed','Failed','Auto%'].forEach((h,i)=>{
    ws.getRange(2,i+1).setValue(h).setBackground('#2E7D32').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center')
        .setBorder(true,true,true,true,false,false,'#81C784',SpreadsheetApp.BorderStyle.SOLID);
  });
  ws.setRowHeight(2,20);

  ws.getRange(3,1,1,7).merge();
  ws.getRange(3,1).setValue('? Run refreshDashboard() untuk mengisi data')
      .setBackground('#F1F8E9').setFontColor('#33691E').setFontStyle('italic')
      .setFontSize(10).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setFrozenRows(2);
}

function writeCoverage(ss, allData) {
  const ws = ss.getSheetByName('Coverage');
  if (!ws) return;
  const lastRow = Math.max(ws.getLastRow(), 3);
  if (lastRow >= 3) ws.getRange(3, 1, lastRow-2, 7).clearContent().clearFormat();

  let r = 3;
  const rules = [];
  allData.forEach(d => {
    if (!d.coverage || d.coverage.length === 0) return;
    d.coverage.forEach((cov, i) => {
      const bg = i%2===0 ? '#F1F8E9':'#FFFFFF';
      const autoRate = cov.total > 0 ? cov.auto/cov.total : 0;
      [d.name, cov.sub, cov.type, cov.total, cov.passed, cov.failed, autoRate].forEach((v,ci) => {
        const c = ws.getRange(r,ci+1).setValue(v).setBackground(bg).setFontFamily('Arial').setFontSize(9)
            .setHorizontalAlignment(ci<3?'left':'center').setVerticalAlignment('middle')
            .setBorder(true,true,true,true,false,false,'#81C784',SpreadsheetApp.BorderStyle.SOLID);
        if (ci===6) c.setNumberFormat('0%');
      });
      // Auto% CF
      const aRng = ws.getRange(r,7);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.8)
          .setBackground('#C8E6C9').setFontColor('#1B5E20').setBold(true).setRanges([aRng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.5,0.799)
          .setBackground('#FFF9C4').setFontColor('#E65100').setBold(true).setRanges([aRng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.5)
          .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true).setRanges([aRng]).build());
      // Failed > 0
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
          .setBackground('#FFCDD2').setFontColor('#C62828').setBold(true)
          .setRanges([ws.getRange(r,6)]).build());
      ws.setRowHeight(r, 20);
      r++;
    });
  });
  ws.setConditionalFormatRules(rules);
}


// ?? HISTORY SHEET (Trend data) ?????????????????????????????????????????????
function buildHistory(ss) {
  const ws = ss.insertSheet('History');
  ws.setTabColor('#4A148C');
  ws.clear();

  ws.getRange(1,1,1,12).merge();
  ws.getRange(1,1).setValue('TREND HISTORY  ?  Pass Rate per Refresh  (auto-appended setiap refresh)')
      .setBackground('#4A148C').setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setFontFamily('Arial').setHorizontalAlignment('center');
  ws.setRowHeight(1,24);

  ['Timestamp','Modul','PIC / Team / Squad','Web Pass%','Web Exec%','API Pass%','API Exec%','Perf','Total Bugs','Open Bugs','Blocker Bugs','Critical Bugs'].forEach((h,i)=>{
    ws.getRange(2,i+1).setValue(h).setBackground('#6A1B9A').setFontColor('#FFFFFF')
        .setFontWeight('bold').setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center');
    ws.setColumnWidth(i+1,[130,120,140,75,75,75,75,60,70,70,80,80][i]||70);
  });
  ws.setRowHeight(2,20);
  ws.setFrozenRows(2);
}

function appendHistory(ss, allData) {
  const ws = ss.getSheetByName('History');
  if (!ws) return;
  const ts = new Date();
  const tsStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  allData.forEach(d => {
    const bs = d.bugStats || {};
    ws.appendRow([tsStr, d.name, d.team,
      d.wPassRate, d.wExecRate, d.aPassRate, d.aExecRate, d.perfResult,
      bs.total||0, bs.open||0, bs.blocker||0, bs.critical||0]);
  });
  // Format % columns
  const lastRow = ws.getLastRow();
  if (lastRow >= 3) {
    [4,5,6,7].forEach(col => {
      ws.getRange(3, col, lastRow-2, 1).setNumberFormat('0%');
    });
    // Trend chart ? Web & API Pass Rate over time
    try {
      ws.getCharts().forEach(c => ws.removeChart(c));
      const trendChart = ws.newChart()
          .setChartType(Charts.ChartType.LINE)
          .addRange(ws.getRange(2, 1, lastRow-1, 1))  // Timestamp
          .addRange(ws.getRange(2, 4, lastRow-1, 1))  // Web Pass%
          .addRange(ws.getRange(2, 6, lastRow-1, 1))  // API Pass%
          .setPosition(2, 14, 0, 0)
          .setOption('title', 'Pass Rate Trend Over Time')
          .setOption('curveType', 'function')
          .setOption('colors', ['#1565C0','#283593'])
          .setOption('vAxis', {title:'Pass Rate', format:'#%', minValue:0, maxValue:1})
          .setOption('hAxis', {title:'Refresh Time'})
          .setOption('legend', {position:'top'})
          .setOption('width', 600).setOption('height', 320)
          .build();
      ws.insertChart(trendChart);
    } catch(e) { Logger.log('History chart skipped: ' + e.message); }
  }
}


// ?? RAW DATA SHEET ?????????????????????????????????????????????????????????
function buildRaw(ss) {
  const ws = ss.insertSheet('_Raw');
  ws.setTabColor('#546E7A');
  ws.clear();
  ws.getRange(1,1).setValue('Internal cache ? jangan diedit manual. Data di-overwrite setiap refresh.')
      .setBackground('#546E7A').setFontColor('#FFFFFF').setFontSize(9).setFontFamily('Arial');
  ws.setRowHeight(1,16);
}

function updateRaw(ss, allData) {
  const ws = ss.getSheetByName('_Raw');
  if (!ws) return;
  ws.clearContents();
  ws.getRange(1,1).setValue('Refreshed: ' + new Date());
  const headers = ['Modul','PIC / Team / Squad','Project / Sprint','wTotal','wPass','wFail','wBlock','wPassRate',
    'aTotal','aPass','aFail','aBlock','aPassRate','Perf','Blockers','Error'];
  ws.getRange(2,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#607D8B').setFontColor('#FFFFFF');
  allData.forEach((d,i) => {
    ws.getRange(3+i,1,1,16).setValues([[
      d.name,d.team,d.sprint,d.wTotal,d.wPassed,d.wFailed,d.wBlocked,d.wPassRate,
      d.aTotal,d.aPassed,d.aFailed,d.aBlocked,d.aPassRate,d.perfResult,d.blockers.length,d.error
    ]]);
  });
}

function updateConfig(ss, allData) {
  // Auto-write PIC/Team/Squad and Project/Sprint back to Config from pulled Summary data
  const cfg = ss.getSheetByName('Config');
  if (!cfg) return;
  const cfgData = cfg.getDataRange().getValues();
  allData.forEach(d => {
    // Find matching row by Spreadsheet ID (col 5 = index 4)
    for (let i = 3; i < cfgData.length; i++) {
      if (String(cfgData[i][4]).trim() === d.id) {
        if (d.team)   cfg.getRange(i+1, 3).setValue(d.team);   // col 3 = PIC / Team / Squad
        if (d.sprint) cfg.getRange(i+1, 4).setValue(d.sprint); // col 4 = Project / Sprint
        break;
      }
    }
  });
}
