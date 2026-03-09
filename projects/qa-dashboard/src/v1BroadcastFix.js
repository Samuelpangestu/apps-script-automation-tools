/**
 * broadcast_master.js  —  v40
 * ═══════════════════════════════════════════════════════════════════════
 * Paste ke Apps Script QA PORTFOLIO DASHBOARD (bukan ke modul).
 *
 * CARA PAKAI:
 *   masterBroadcastAll()    → jalankan SEMUA fix ke semua modul aktif
 *   broadcastXxx()          → jalankan 1 fix saja ke semua modul
 *   testSingleSheet()       → test 1 fix ke 1 Spreadsheet ID tertentu
 *
 * DAFTAR FIX (urutan eksekusi di masterBroadcastAll):
 *   A. Dashboard Headers   — rebuild Overview + buat tab Smoke jika belum ada
 *   B. Smoke Section       — insert A1. SMOKE TEST di Summary
 *   C. Smoke Formulas      — fix #VALUE! & range mismatch di formula Smoke
 *   D. Smoke Blocker Row   — tambah baris "BUG BLOCKER" di Bug Summary
 *   E. Description Column  — insert kolom Description di BugReport col 8
 *   F. Medium Wording      — ganti "Medium (Blocker)" → "Medium" di Summary
 *   G. Coverage Rows       — expand COVERAGE PER SUBMODUL ke 34 baris
 *   H. API Expected Result — insert kolom Expected Result di API_Master col 14
 *   I. Recreate Appendix   — tulis ulang tab Appendix lengkap
 *
 * CATATAN RANGE FORMULA SMOKE (Fix C):
 *   TC_Master starts row 3, TC_Execution starts row 9 → offset 6.
 *   TC_Execution!Z9:Z1000  = 992 rows
 *   TC_Master!E3:E994      = 992 rows  ← HARUS SAMA PERSIS
 *   Default sheet = 1000 rows. Jangan pakai >1000 → clips differently → #VALUE!
 * ═══════════════════════════════════════════════════════════════════════
 */


// ═══════════════════════════════════════════════════════════════════════
// MASTER RUNNER
// ═══════════════════════════════════════════════════════════════════════

function masterBroadcastAll() {
    Logger.log('masterBroadcastAll mulai...');

    const STEPS = [
        ['Dashboard Headers',   () => fixDashboardHeaders_()],
        ['Summary Fields',      () => runBroadcast_(fixSummaryFields_,     'Summary Fields')],
        ['QA Lead Label',       () => runBroadcast_(fixQALeadLabel_,       'QA Lead Label')],
        ['Submodul Note',       () => runBroadcast_(fixSubmodulNote_,      'Submodul Note')],
        ['PIC QA Dropdown',     () => runBroadcast_(fixPICQADropdown_,     'PIC QA Dropdown')],
        ['Smoke Section',       () => runBroadcast_(addSmokeSection_,      'Smoke Section')],
        ['Smoke Formulas',      () => runBroadcast_(fixSmokeForms_,        'Smoke Formulas')],
        ['Smoke Blocker Row',   () => runBroadcast_(addSmokeBlockerRow_,   'Smoke Blocker')],
        ['Description Column',  () => runBroadcast_(addDescriptionCol_,    'Description')],
        ['Medium Wording',      () => runBroadcast_(fixMediumWording_,     'Medium Wording')],
        ['Coverage Rows',       () => runBroadcast_(addCoverageRows_,      'Coverage')],
        ['API Expected Result', () => runBroadcast_(fixApiExpectedResult_, 'API Expected')],
        ['Recreate Appendix',   () => runBroadcastAppendix_()],
        ['Fix Open Blocker Formula', () => runBroadcast_(fixSimplifyFormulas_, 'Open Blocker Formula')],
    ];

    const lines = [];
    STEPS.forEach(([name, fn]) => {
        try {
            const result = fn();
            const ok = !result.includes('ERR') || result.includes('Err:0');
            Logger.log((ok?'OK  ':'ERR ') + name + ': ' + result);
            lines.push((ok?'✅':'❌') + ' ' + name + ': ' + result);
        } catch(e) {
            Logger.log('ERR ' + name + ': ' + e.message);
            lines.push('❌ ' + name + ': ' + e.message);
        }
    });

    const msg = '📋 masterBroadcastAll — Selesai\n' + lines.join('\n');
    Logger.log(msg);
    safeAlert_(msg);
}

/**
 * 🚀 ALL-IN-ONE MIGRATION for QATM modules
 *
 * Jalankan fungsi ini SEKALI untuk migrate semua QATM modules ke struktur baru.
 * Script ini AMAN dan IDEMPOTENT (bisa dijalankan berkali-kali tanpa error).
 *
 * Yang dilakukan (hanya fixes yang relevan untuk migration hari ini):
 * 1. fixSummaryFields - Restructure Summary fields (Project, Modul, Submodul, QA Lead, PIC QA)
 * 2. fixQALeadLabel - Update "QA Team Lead:" → "QA Lead:"
 * 3. fixSubmodulNote - Update Submodul note example → "1.1,1.2,1.3"
 * 4. fixPICQADropdown - Remove dropdown validation dari PIC QA (jadi free text)
 *
 * Skip fixes yang tidak relevan hari ini (Smoke, Blocker, Coverage, dll).
 */
function migrateQATMModulesToNewStructure() {
    Logger.log('═══════════════════════════════════════════════════════════');
    Logger.log('🚀 QATM MODULES MIGRATION START: ' + new Date());
    Logger.log('═══════════════════════════════════════════════════════════');

    const MIGRATION_STEPS = [
        ['Summary Fields',      () => runBroadcast_(fixSummaryFields_,     'Summary Fields')],
        ['QA Lead Label',       () => runBroadcast_(fixQALeadLabel_,       'QA Lead Label')],
        ['Submodul Note',       () => runBroadcast_(fixSubmodulNote_,      'Submodul Note')],
        ['PIC QA Dropdown',     () => runBroadcast_(fixPICQADropdown_,     'PIC QA Dropdown')],
    ];

    const lines = [];
    MIGRATION_STEPS.forEach(([name, fn]) => {
        try {
            Logger.log('\n📝 Running: ' + name);
            const result = fn();
            const ok = !result.includes('ERR') || result.includes('Err:0');
            Logger.log((ok?'✅ ':'❌ ') + name + ': ' + result);
            lines.push((ok?'✅':'❌') + ' ' + name + ': ' + result);
        } catch(e) {
            Logger.log('❌ ' + name + ': ' + e.message);
            lines.push('❌ ' + name + ': ' + e.message);
        }
    });

    Logger.log('\n═══════════════════════════════════════════════════════════');
    Logger.log('🎉 QATM MODULES MIGRATION COMPLETE: ' + new Date());
    Logger.log('═══════════════════════════════════════════════════════════');

    const msg = '🚀 QATM MODULES MIGRATION COMPLETE\n\n' + lines.join('\n') +
                '\n\n📝 Next Steps:\n1. Verify Summary tab di semua QATM\n2. Jalankan migrateDashboardToNewStructure() di Dashboard\n3. Run refreshDashboard() untuk sync data';

    Logger.log('\n' + msg);
    safeAlert_(msg);
}

// Individual callers
function broadcastDashboardHeaders()   { safeAlert_(fixDashboardHeaders_()); }
function broadcastSummaryFields()      { safeAlert_('Summary Fields: '     + runBroadcast_(fixSummaryFields_,     'Summary Fields')); }
function broadcastQALeadLabel()        { safeAlert_('QA Lead Label: '      + runBroadcast_(fixQALeadLabel_,       'QA Lead Label')); }
function broadcastSubmodulNote()       { safeAlert_('Submodul Note: '      + runBroadcast_(fixSubmodulNote_,      'Submodul Note')); }
function broadcastPICQADropdown()      { safeAlert_('PIC QA Dropdown: '    + runBroadcast_(fixPICQADropdown_,     'PIC QA Dropdown')); }
function broadcastSmokeSection()       { safeAlert_('Smoke Section: '      + runBroadcast_(addSmokeSection_,      'Smoke Section')); }
function broadcastSmokeFormulas()      { safeAlert_('Smoke Formulas: '     + runBroadcast_(fixSmokeForms_,        'Smoke Formulas')); }
function broadcastSmokeBlockerRow()    { safeAlert_('Smoke Blocker Row: '  + runBroadcast_(addSmokeBlockerRow_,   'Smoke Blocker')); }
function broadcastDescriptionCol()     { safeAlert_('Description Col: '    + runBroadcast_(addDescriptionCol_,    'Description')); }
function broadcastMediumWording()      { safeAlert_('Medium Wording: '     + runBroadcast_(fixMediumWording_,     'Medium Wording')); }
function broadcastCoverageRows()       { safeAlert_('Coverage Rows: '      + runBroadcast_(addCoverageRows_,      'Coverage')); }
function broadcastApiExpectedResult()  { safeAlert_('API Expected: '       + runBroadcast_(fixApiExpectedResult_, 'API Expected')); }
function broadcastRecreateAppendix()   { safeAlert_('Appendix: '           + runBroadcastAppendix_()); }
function broadcastFixOpenBlocker()     { safeAlert_('Open Blocker: '       + runBroadcast_(fixSimplifyFormulas_, 'Open Blocker Formula')); }


// ═══════════════════════════════════════════════════════════════════════
// HELPERS — iterator + alert
// ═══════════════════════════════════════════════════════════════════════

function runBroadcast_(perSheetFn, label) {
    const ids = getActiveIds_();
    let ok=0, skip=0, err=0, errList=[];
    ids.forEach(id => {
        try {
            const result = perSheetFn(id);
            if (result && result.includes('skipped')) {
                skip++;
                Logger.log('SKIP [' + label + '] ' + id + ' — ' + result);
            } else {
                ok++;
                Logger.log('OK   [' + label + '] ' + id + ' — ' + (result||'done'));
            }
        } catch(e) {
            err++;
            errList.push(id + ': ' + e.message);
            Logger.log('ERR  [' + label + '] ' + id + ' — ' + e.message);
        }
    });
    let summary = 'OK:' + ok + ' Skip:' + skip + ' Err:' + err;
    if (errList.length > 0) summary += '\n  • ' + errList.join('\n  • ');
    return summary;
}

function getActiveIds_() {
    const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!cfg) throw new Error('Config tab tidak ditemukan di Dashboard');
    const data = cfg.getDataRange().getValues();
    const ids  = [];
    for (let i = 3; i < data.length; i++) {
        const active = data[i][0];  // col A = Active (TRUE/FALSE or Y/N for backward compat)
        const id     = String(data[i][6]).trim();  // col G = Spreadsheet ID

        // Support both TRUE/FALSE (new) and Y/N (old)
        const isActive = (active === true || String(active).trim().toUpperCase() === 'Y');

        if (isActive && id && id.length > 20 && id !== 'PASTE_SPREADSHEET_ID_HERE') {
            ids.push(id);
            Logger.log('[getActiveIds_] Found active module: ' + id);
        }
    }
    Logger.log('[getActiveIds_] Total active modules: ' + ids.length);
    return ids;
}

function safeAlert_(msg) {
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// Interactive single-sheet tester
function testSingleSheet() {
    const ui = SpreadsheetApp.getUi();
    const r1 = ui.prompt('Test Single Sheet',
        '1=Smoke Section\n2=Smoke Formulas\n3=Smoke Blocker\n4=Description Col\n5=Medium Wording\n6=Coverage Rows\n7=API Expected\n8=Appendix',
        ui.ButtonSet.OK_CANCEL);
    if (r1.getSelectedButton() !== ui.Button.OK) return;
    const r2 = ui.prompt('Spreadsheet ID', 'Paste Spreadsheet ID modul:', ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK) return;
    const id = r2.getResponseText().trim();
    const fns = {
        '1':addSmokeSection_, '2':fixSmokeForms_, '3':addSmokeBlockerRow_,
        '4':addDescriptionCol_, '5':fixMediumWording_, '6':addCoverageRows_,
        '7':fixApiExpectedResult_,
        '8': (i) => { buildAppendix_(SpreadsheetApp.openById(i)); return 'Appendix rebuilt'; }
    };
    const fn = fns[r1.getResponseText().trim()];
    if (!fn) { ui.alert('Pilihan tidak valid'); return; }
    try { ui.alert('Hasil: ' + fn(id)); } catch(e) { ui.alert('ERROR: ' + e.message); }
}


// ═══════════════════════════════════════════════════════════════════════
// FIX A — DASHBOARD HEADERS
// Rebuild Overview header rows 1-4 (tanpa menyentuh data rows 5+).
// Buat tab Smoke di Dashboard jika belum ada.
// Strategy: breakApart DULU seluruh rows 1-4, baru tulis ulang — aman.
// ═══════════════════════════════════════════════════════════════════════

function fixDashboardHeaders_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = [];

    // ── Overview ──────────────────────────────────────────────────────────
    const ov = ss.getSheetByName('Overview');
    if (ov) {
        const lastCol = Math.max(ov.getLastColumn()||1, 25);
        try { ov.getRange(1,1,4,lastCol).breakApart(); } catch(e) {}
        ov.getRange(1,1,4,lastCol).clearContent().clearFormat();
        rebuildOverviewHeaders_(ov);
        log.push('Overview headers rebuilt (25 cols, Smoke bersebelahan WEB/API)');
    } else {
        log.push('Overview: tab tidak ditemukan (jalankan createDashboard() dulu)');
    }

    // ── Smoke tab ──────────────────────────────────────────────────────────
    let smokeSheet = ss.getSheetByName('Smoke');
    if (!smokeSheet) {
        smokeSheet = ss.insertSheet('Smoke');
        smokeSheet.setTabColor('#BF360C');
        smokeSheet.clear();
        smokeSheet.getRange(1,1,1,13).merge().setValue('Last refreshed: —')
            .setBackground('#FBE9E7').setFontColor('#BF360C').setFontStyle('italic')
            .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
        smokeSheet.getRange(2,1,1,13).merge()
            .setValue('🔥 SMOKE TEST DASHBOARD — Jalankan refreshDashboard() untuk isi data')
            .setBackground('#BF360C').setFontColor('#FFFFFF').setFontWeight('bold')
            .setFontSize(12).setFontFamily('Arial').setHorizontalAlignment('center');
        log.push('Smoke tab created');
    } else {
        log.push('Smoke tab sudah ada');
    }

    const result = log.join(' | ') + '. Jalankan refreshDashboard() untuk isi data.';
    Logger.log('FIX A: ' + result);
    return result;
}

function rebuildOverviewHeaders_(ov) {
    function h_(r,c,nr,nc,txt,bg,fg,sz) {
        const rng=(nr>1||nc>1)?ov.getRange(r,c,nr,nc).merge():ov.getRange(r,c);
        rng.setValue(txt||'').setBackground(bg||'#0D47A1').setFontColor(fg||'#FFFFFF')
            .setFontWeight('bold').setFontSize(sz||9).setFontFamily('Arial')
            .setHorizontalAlignment('center').setVerticalAlignment('middle')
            .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    }

    // 25 cols: 4 MODULE | 5 WEB | 3 SMOKE WEB | 5 API | 3 SMOKE API | 1 PERF | 3 BUGS | 1 NOTES
    [145,100,100,125, 52,56,52,52,68, 62,68,58, 52,56,52,52,68, 62,68,58, 68, 52,62,56, 165]
        .forEach((w,i) => ov.setColumnWidth(i+1,w));

    // Row 1 — refresh timestamp
    ov.getRange(1,1,1,25).merge().setValue('Last refreshed: —')
        .setBackground('#E3F2FD').setFontColor('#1565C0').setFontStyle('italic')
        .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left');
    ov.setRowHeight(1,16);

    // Row 2 — title
    h_(2,1,1,25,'QA DASHBOARD  |  PORTFOLIO OVERVIEW','#0D47A1','#FFFFFF',13);
    ov.setRowHeight(2,30);

    // Row 3 — group headers
    h_(3,1, 1,4,'MODULE INFO',   '#263238');
    h_(3,5, 1,5,'WEB / MOBILE',  '#1565C0');
    h_(3,10,1,3,'🔥 SMOKE WEB',  '#BF360C');
    h_(3,13,1,5,'API',            '#283593');
    h_(3,18,1,3,'🔥 SMOKE API',  '#4A148C');
    h_(3,21,1,1,'PERF',           '#004D40');
    h_(3,22,1,3,'BUGS',           '#B71C1C');
    h_(3,25,1,1,'NOTES',          '#37474F');
    ov.setRowHeight(3,22);

    // Row 4 — column headers
    ['SubModule','Project','Module','PIC / Team',
        'Total','Pass','Fail','Block','Pass%',
        'Total','Pass%','Exec%',
        'Total','Pass','Fail','Block','Pass%',
        'Total','Pass%','Exec%',
        'Perf','Bugs','Blocker','Critical','Notes'
    ].forEach((lbl,i) => h_(4,i+1,1,1,lbl,'#1565C0'));

    ov.getRange(4,10).setNote('Smoke Web: TC Priority Critical+High+Medium');
    ov.getRange(4,11).setNote('Smoke Web Pass Rate (target ≥80%)');
    ov.getRange(4,12).setNote('Smoke Web Exec Rate');
    ov.getRange(4,18).setNote('Smoke API: TC Priority Critical+High+Medium');
    ov.getRange(4,19).setNote('Smoke API Pass Rate (target ≥80%)');
    ov.getRange(4,20).setNote('Smoke API Exec Rate');
    ov.setRowHeight(4,26);
    ov.setFrozenRows(4);
}


// ═══════════════════════════════════════════════════════════════════════
// FIX B — ADD SMOKE TEST SECTION TO SUMMARY
// Insert 4 baris A1. SMOKE TEST sebelum section B (KOMPOSISI STATUS).
// ═══════════════════════════════════════════════════════════════════════

function addSmokeSection_(spreadsheetId) {
    const src = SpreadsheetApp.openById(spreadsheetId);
    const ws  = src.getSheetByName('Summary');
    if (!ws) return 'skipped (no Summary)';
    if (ws.createTextFinder('SMOKE TEST').matchEntireCell(false).findNext())
        return 'skipped (SMOKE TEST already exists)';

    const secB = ws.createTextFinder('KOMPOSISI STATUS').matchEntireCell(false).findNext();
    if (!secB) return 'skipped (KOMPOSISI STATUS not found)';

    const secBRow = secB.getRow();
    ws.insertRowsBefore(secBRow, 4);
    let R = secBRow;
    const L=1, LW=10, R_=12, RW=10;

    function h_(rng,bg) {
        return rng.setBackground(bg||'#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold')
            .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center').setVerticalAlignment('middle')
            .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
    }
    function m_(r,c,nr,nc) { if(nr>1||nc>1) ws.getRange(r,c,nr,nc).merge(); }
    function bd_(rng) {
        return rng.setFontFamily('Arial').setFontWeight('bold')
            .setHorizontalAlignment('center').setVerticalAlignment('middle')
            .setBorder(true,true,true,true,false,false,'#90CAF9',SpreadsheetApp.BorderStyle.SOLID);
    }

    // Row 1: headers
    m_(R,L,1,LW); h_(ws.getRange(R,L),'#BF360C').setValue('A1.  SMOKE TEST  —  Web / Mobile  (Critical + High + Medium)');
    m_(R,R_,1,RW); h_(ws.getRange(R,R_),'#4A148C').setValue('A1.  SMOKE TEST  —  API  (Critical + High + Medium)');
    ws.setRowHeight(R,18); R++;

    // Row 2: KPI labels
    const kpiLabels = ['TOTAL','PASSED','FAILED','BLOCKED','IN PROG','TODO','PASS RATE','AUTO RATE','EXEC RATE'];
    const kpiBgs    = ['#0D47A1','#1B5E20','#B71C1C','#E65100','#1565C0','#455A64','#004D40','#4A148C','#1A237E'];
    for (let i=0; i<9; i++) {
        h_(ws.getRange(R,L+i), kpiBgs[i]).setValue(kpiLabels[i]).setFontSize(i<6?8:7.5).setWrap(true);
        h_(ws.getRange(R,R_+i),kpiBgs[i]).setValue(kpiLabels[i]).setFontSize(i<6?8:7.5).setWrap(true);
    }
    ws.setRowHeight(R,22); R++;

    // Row 3: formulas
    // KEY: Z9:Z1000 = 992 rows. E3:E994 = 992 rows. Must match exactly.
    const wZ='TC_Execution!Z9:Z1000',   aZ='API_Execution!Z9:Z1000';
    const wP='TC_Master!E3:E994',       aP='API_Master!G3:G994';
    const wPF='TC_Master!E3:E1000',     aPF='API_Master!G3:G1000';
    const wH='TC_Master!H3:H994',       aJ='API_Master!J3:J994';

    function sc_(z,p,s){ return ['Critical','High','Medium'].map(v=>`COUNTIFS(${p},"${v}",${z},"${s}")`).join('+'); }
    function st_(pf){ return ['Critical','High','Medium'].map(v=>`COUNTIF(${pf},"${v}")`).join('+'); }

    const wT=`(${st_(wPF)})`, aT=`(${st_(aPF)})`;
    const wF=[
        `=${wT}`,
        `=${sc_(wZ,wP,'PASSED')}`,`=${sc_(wZ,wP,'FAILED')}`,`=${sc_(wZ,wP,'BLOCKED')}`,`=${sc_(wZ,wP,'IN PROGRESS')}`,`=${sc_(wZ,wP,'TODO')}`,
        `=IFERROR((${sc_(wZ,wP,'PASSED')})/MAX(1,${wT}),0)`,
        `=IFERROR((COUNTIFS(${wH},"Automated",${wP},"Critical")+COUNTIFS(${wH},"Automated",${wP},"High")+COUNTIFS(${wH},"Automated",${wP},"Medium"))/MAX(1,${wT}),0)`,
        `=IFERROR((${sc_(wZ,wP,'PASSED')}+${sc_(wZ,wP,'FAILED')}+${sc_(wZ,wP,'BLOCKED')}+${sc_(wZ,wP,'IN PROGRESS')})/MAX(1,${wT}),0)`,
    ];
    const aF=[
        `=${aT}`,
        `=${sc_(aZ,aP,'PASSED')}`,`=${sc_(aZ,aP,'FAILED')}`,`=${sc_(aZ,aP,'BLOCKED')}`,`=${sc_(aZ,aP,'IN PROGRESS')}`,`=${sc_(aZ,aP,'TODO')}`,
        `=IFERROR((${sc_(aZ,aP,'PASSED')})/MAX(1,${aT}),0)`,
        `=IFERROR((COUNTIFS(${aJ},"Automated",${aP},"Critical")+COUNTIFS(${aJ},"Automated",${aP},"High")+COUNTIFS(${aJ},"Automated",${aP},"Medium"))/MAX(1,${aT}),0)`,
        `=IFERROR((${sc_(aZ,aP,'PASSED')}+${sc_(aZ,aP,'FAILED')}+${sc_(aZ,aP,'BLOCKED')}+${sc_(aZ,aP,'IN PROGRESS')})/MAX(1,${aT}),0)`,
    ];

    wF.forEach((f,i) => {
        const c=bd_(ws.getRange(R,L+i)).setFormula(f).setBackground('#FFF3E0').setFontSize(i<6?16:13).setFontColor('#BF360C');
        if(i>=6) c.setNumberFormat('0%');
    });
    aF.forEach((f,i) => {
        const c=bd_(ws.getRange(R,R_+i)).setFormula(f).setBackground('#F3E5F5').setFontSize(i<6?16:13).setFontColor('#4A148C');
        if(i>=6) c.setNumberFormat('0%');
    });
    ws.setRowHeight(R,36); R++;

    // Row 4: legend
    m_(R,L,1,LW); ws.getRange(R,L).setValue('Smoke Test = TC Priority Critical/High/Medium  |  Target Pass Rate ≥80% sebelum release')
        .setBackground('#FFF8E1').setFontColor('#E65100').setFontStyle('italic').setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left');
    m_(R,R_,1,RW); ws.getRange(R,R_).setValue('Smoke Test = API Priority Critical/High/Medium')
        .setBackground('#EDE7F6').setFontColor('#4A148C').setFontStyle('italic').setFontSize(7).setFontFamily('Arial').setHorizontalAlignment('left');
    ws.setRowHeight(R,14);

    return 'Smoke Test section added (4 rows before row ' + secBRow + ')';
}


// ═══════════════════════════════════════════════════════════════════════
// FIX C — SMOKE FORMULAS
// Overwrite formulas Smoke row dengan range yang tepat (992 rows).
// Fix #VALUE! (array size mismatch) & AUTO RATE 400%.
// ═══════════════════════════════════════════════════════════════════════

function fixSmokeForms_(spreadsheetId) {
    const src = SpreadsheetApp.openById(spreadsheetId);
    const ws  = src.getSheetByName('Summary');
    if (!ws) return 'skipped (no Summary)';
    const smokeCell = ws.createTextFinder('SMOKE TEST').matchEntireCell(false).findNext();
    if (!smokeCell) return 'skipped (no SMOKE TEST section)';
    const valRow = smokeCell.getRow() + 2;
    const testF  = ws.getRange(valRow,1).getFormula();
    if (!testF || !testF.includes('COUNTIF')) return 'skipped (value row not found or no formula)';

    const L=1, R_=12;
    const wZ='TC_Execution!Z9:Z1000',   aZ='API_Execution!Z9:Z1000';
    const wP='TC_Master!E3:E994',       aP='API_Master!G3:G994';
    const wPF='TC_Master!E3:E1000',     aPF='API_Master!G3:G1000';
    const wH='TC_Master!H3:H994',       aJ='API_Master!J3:J994';

    function sc_(z,p,s){ return ['Critical','High','Medium'].map(v=>`COUNTIFS(${p},"${v}",${z},"${s}")`).join('+'); }
    function st_(pf){ return ['Critical','High','Medium'].map(v=>`COUNTIF(${pf},"${v}")`).join('+'); }

    const wT=`(${st_(wPF)})`, aT=`(${st_(aPF)})`;
    const wF=[
        `=${wT}`,`=${sc_(wZ,wP,'PASSED')}`,`=${sc_(wZ,wP,'FAILED')}`,`=${sc_(wZ,wP,'BLOCKED')}`,`=${sc_(wZ,wP,'IN PROGRESS')}`,`=${sc_(wZ,wP,'TODO')}`,
        `=IFERROR((${sc_(wZ,wP,'PASSED')})/MAX(1,${wT}),0)`,
        `=IFERROR((COUNTIFS(${wH},"Automated",${wP},"Critical")+COUNTIFS(${wH},"Automated",${wP},"High")+COUNTIFS(${wH},"Automated",${wP},"Medium"))/MAX(1,${wT}),0)`,
        `=IFERROR((${sc_(wZ,wP,'PASSED')}+${sc_(wZ,wP,'FAILED')}+${sc_(wZ,wP,'BLOCKED')}+${sc_(wZ,wP,'IN PROGRESS')})/MAX(1,${wT}),0)`,
    ];
    const aF=[
        `=${aT}`,`=${sc_(aZ,aP,'PASSED')}`,`=${sc_(aZ,aP,'FAILED')}`,`=${sc_(aZ,aP,'BLOCKED')}`,`=${sc_(aZ,aP,'IN PROGRESS')}`,`=${sc_(aZ,aP,'TODO')}`,
        `=IFERROR((${sc_(aZ,aP,'PASSED')})/MAX(1,${aT}),0)`,
        `=IFERROR((COUNTIFS(${aJ},"Automated",${aP},"Critical")+COUNTIFS(${aJ},"Automated",${aP},"High")+COUNTIFS(${aJ},"Automated",${aP},"Medium"))/MAX(1,${aT}),0)`,
        `=IFERROR((${sc_(aZ,aP,'PASSED')}+${sc_(aZ,aP,'FAILED')}+${sc_(aZ,aP,'BLOCKED')}+${sc_(aZ,aP,'IN PROGRESS')})/MAX(1,${aT}),0)`,
    ];

    let n=0;
    wF.forEach((f,i) => { const c=ws.getRange(valRow,L+i); if(c.getFormula()!==f){c.setFormula(f);if(i>=6)c.setNumberFormat('0%');n++;} });
    aF.forEach((f,i) => { const c=ws.getRange(valRow,R_+i); if(c.getFormula()!==f){c.setFormula(f);if(i>=6)c.setNumberFormat('0%');n++;} });

    return n===0 ? 'skipped (already correct)' : n + ' formulas fixed at row ' + valRow;
}


// ═══════════════════════════════════════════════════════════════════════
// HELPER — Remove duplicate BUG BLOCKER rows
// ═══════════════════════════════════════════════════════════════════════

function removeDuplicateBugBlocker_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const summ = src.getSheetByName('Summary');
    if (!summ) return 'skipped (no Summary)';

    const finder = summ.createTextFinder('BUG BLOCKER').matchEntireCell(false).findAll();
    if (finder.length <= 1) return 'skipped (no duplicates)';

    // Keep the last one (row 18 - "Open Blocker (Smoke)" location), delete the first one (row 16 - newly created)
    // Find which row is closer to "A1. SMOKE TEST" - that's the duplicate to remove
    const smokeRow = summ.createTextFinder('A1. SMOKE TEST').matchEntireCell(false).findNext();
    if (!smokeRow) return 'skipped (A1. SMOKE TEST not found)';

    const smokeRowNum = smokeRow.getRow();
    let rowToDelete = null;

    finder.forEach(cell => {
        const rowNum = cell.getRow();
        // Delete the one that's closer to SMOKE TEST header (the newly inserted one)
        if (Math.abs(rowNum - smokeRowNum) < 5) {
            rowToDelete = rowNum;
        }
    });

    if (rowToDelete) {
        summ.deleteRows(rowToDelete, 2);  // Delete 2 rows (label + data)
        return 'deleted duplicate at row ' + rowToDelete;
    }

    return 'skipped (could not identify duplicate)';
}

// ═══════════════════════════════════════════════════════════════════════
// FIX D — SMOKE BLOCKER ROW di Bug Summary
// Rename "Open Blocker (Smoke)" → "BUG BLOCKER" atau buat baru jika belum ada.
// ═══════════════════════════════════════════════════════════════════════

function addSmokeBlockerRow_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const summ = src.getSheetByName('Summary');
    if (!summ) return 'skipped (no Summary)';

    const L=1, LW=10, R_=12, RW=10;

    // Check if "Open Blocker (Smoke)" exists - rename it
    const oldBlockerCell = summ.createTextFinder('Open Blocker').matchEntireCell(false).findNext();
    if (oldBlockerCell) {
        const row = oldBlockerCell.getRow();
        summ.getRange(row,L,1,LW).setValue('BUG BLOCKER');
        summ.getRange(row,R_,1,RW).setValue('BUG BLOCKER');
        return 'renamed "Open Blocker (Smoke)" → "BUG BLOCKER"';
    }

    // Check if "BUG BLOCKER" already exists
    if (summ.createTextFinder('BUG BLOCKER').matchEntireCell(false).findNext())
        return 'skipped (already exists)';

    // If neither exists, create new row
    const medCell = summ.createTextFinder('Medium').matchEntireCell(false).findNext();
    if (!medCell) return 'skipped (Medium row not found)';
    const insertRow = medCell.getRow() + 1;
    summ.insertRowsAfter(insertRow-1, 2);

    // Separator row
    summ.getRange(insertRow,L,1,LW).merge().setValue('BUG BLOCKER')
        .setBackground('#FFEBEE').setFontColor('#C62828').setFontWeight('bold').setFontSize(8).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    summ.getRange(insertRow,R_,1,RW).merge().setValue('BUG BLOCKER')
        .setBackground('#EDE7F6').setFontColor('#4A148C').setFontWeight('bold').setFontSize(8).setFontFamily('Arial')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    summ.setRowHeight(insertRow,16);

    // Data row — COUNTIFS BugReport Open/InProg/Reopen x Medium/High/Critical
    const dataRow = insertRow + 1;
    const wBug = `=COUNTIFS(BugReport!D:D,"Open",BugReport!C:C,"Critical")`+
        `+COUNTIFS(BugReport!D:D,"Open",BugReport!C:C,"High")`+
        `+COUNTIFS(BugReport!D:D,"Open",BugReport!C:C,"Medium")`+
        `+COUNTIFS(BugReport!D:D,"In Progress",BugReport!C:C,"Critical")`+
        `+COUNTIFS(BugReport!D:D,"In Progress",BugReport!C:C,"High")`+
        `+COUNTIFS(BugReport!D:D,"In Progress",BugReport!C:C,"Medium")`+
        `+COUNTIFS(BugReport!D:D,"Reopen",BugReport!C:C,"Critical")`+
        `+COUNTIFS(BugReport!D:D,"Reopen",BugReport!C:C,"High")`+
        `+COUNTIFS(BugReport!D:D,"Reopen",BugReport!C:C,"Medium")`;

    summ.getRange(dataRow,L,1,LW).merge().setFormula(wBug)
        .setBackground('#FFCDD2').setFontColor('#B71C1C').setFontWeight('bold').setFontSize(16).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    summ.getRange(dataRow,R_,1,RW).merge().setFormula(wBug)
        .setBackground('#EDE7F6').setFontColor('#4A148C').setFontWeight('bold').setFontSize(16).setFontFamily('Arial')
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    summ.setRowHeight(dataRow,36);

    return 'Open Blocker rows added at rows ' + insertRow + '-' + dataRow;
}


// ═══════════════════════════════════════════════════════════════════════
// FIX E — DESCRIPTION COLUMN di BugReport
// Insert kolom "Description" di col 8 (setelah Title col 7).
// ═══════════════════════════════════════════════════════════════════════

function addDescriptionCol_(spreadsheetId) {
    const src = SpreadsheetApp.openById(spreadsheetId);
    const ws  = src.getSheetByName('BugReport');
    if (!ws) return 'skipped (no BugReport)';
    const hdrRow  = 4;
    const lastCol = ws.getLastColumn();
    const hdrs    = ws.getRange(hdrRow,1,1,lastCol).getValues()[0];
    if (hdrs.some(h => String(h).toLowerCase().includes('description')))
        return 'skipped (Description already exists)';

    const insertCol = 8;
    ws.insertColumnBefore(insertCol);
    ws.setColumnWidth(insertCol, 220);
    ws.getRange(hdrRow,insertCol).setValue('Description')
        .setBackground('#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setNote('Konteks yang tidak muat di Title:\n• Kondisi/step reproduksi singkat\n• Role/user yang terdampak\n• Frekuensi kejadian\n• Data yang digunakan');

    // Extend DETAIL group merge (row 3) jika ada
    try {
        const grpRow = hdrRow - 1;
        ws.getRange(grpRow,1,1,ws.getLastColumn()).getMergedRanges().forEach(m => {
            if (m.getColumn() <= insertCol-1 && m.getLastColumn() >= insertCol-1) {
                m.breakApart();
                ws.getRange(grpRow, m.getColumn(), 1, m.getNumColumns()+1).merge()
                    .setValue(ws.getRange(grpRow, m.getColumn()).getValue());
            }
        });
    } catch(e) {}

    return 'Description col inserted at col ' + insertCol;
}


// ═══════════════════════════════════════════════════════════════════════
// FIX F — MEDIUM WORDING
// Ganti "Medium (Blocker)" → "Medium" di Summary.
// Pakai replaceAllWith → handle merged cells tanpa touch merge API.
// Formula text: per-cell try/catch (merged cell non-top-left → skip).
// ═══════════════════════════════════════════════════════════════════════

function fixMediumWording_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const summ = src.getSheetByName('Summary');
    if (!summ) return 'skipped (no Summary)';

    // replaceAllWith handles merged cell display values tanpa merge API
    let n1=0, n2=0;
    try { n1 = summ.createTextFinder('Medium (Blocker):').replaceAllWith('Medium:'); } catch(e) {}
    try { n2 = summ.createTextFinder('Medium (Blocker)').replaceAllWith('Medium');  } catch(e) {}

    // Fix formula COUNTIF criteria — skip merged non-top-left cells (setFormula would fail)
    let nF=0;
    summ.createTextFinder('Medium (Blocker)').matchFormulaText(true).findAll().forEach(c => {
        const f = c.getFormula();
        if (!f || !f.includes('Medium (Blocker)')) return;
        try {
            c.setFormula(f.replace(/Medium \(Blocker\)/g, 'Medium'));
            nF++;
        } catch(e) {
            Logger.log('  Medium formula skip (merged): row=' + c.getRow() + ' — ' + e.message);
        }
    });

    const total = n1+n2+nF;
    return total===0 ? 'skipped (already correct or not found)' : total + ' replaced';
}


// ═══════════════════════════════════════════════════════════════════════
// FIX G — COVERAGE ROWS (expand ke 34 baris)
// ═══════════════════════════════════════════════════════════════════════

function addCoverageRows_(spreadsheetId) {
    const src = SpreadsheetApp.openById(spreadsheetId);
    const ws  = src.getSheetByName('Summary');
    if (!ws) return 'skipped (no Summary)';

    const covCell = ws.createTextFinder('COVERAGE PER SUBMODUL').matchEntireCell(false).findNext();
    if (!covCell) return 'skipped (COVERAGE section not found)';

    const hdrRow = covCell.getRow();
    const nextSec = ws.createTextFinder('AUTOMATION STATUS').matchEntireCell(false).findNext();
    const endRow  = nextSec ? nextSec.getRow()-1 : hdrRow+35;
    const currentRows = endRow - hdrRow;

    if (currentRows >= 34) return 'skipped (already ' + currentRows + ' rows)';
    const toAdd = 34 - currentRows;
    ws.insertRowsAfter(endRow, toAdd);
    // Copy format from last data row
    ws.getRange(endRow,1,1,ws.getLastColumn()).copyTo(
        ws.getRange(endRow+1,1,toAdd,ws.getLastColumn()), {contentsOnly:false});
    ws.getRange(endRow+1,1,toAdd,ws.getLastColumn()).clearContent();

    return toAdd + ' rows added (now 34)';
}


// ═══════════════════════════════════════════════════════════════════════
// FIX H — API EXPECTED RESULT COLUMN
// Insert kolom "Expected Result" di API_Master col 14.
// Fix formula API_Execution TestLevel (O → P setelah insert).
// ═══════════════════════════════════════════════════════════════════════

function fixApiExpectedResult_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const apim = src.getSheetByName('API_Master');
    if (!apim) return 'skipped (no API_Master)';

    const hdrRow  = 2;
    const lastCol = apim.getLastColumn();
    const hdrs    = apim.getRange(hdrRow,1,1,lastCol).getValues()[0];
    if (hdrs.some(h => String(h).toLowerCase().includes('expected')))
        return 'skipped (already exists)';

    const insertCol = 14;
    apim.insertColumnBefore(insertCol);
    apim.setColumnWidth(insertCol, 220);
    apim.getRange(hdrRow,insertCol).setValue('Expected Result')
        .setBackground('#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setNote('[WAJIB] Then: HTTP status + response body\nContoh: 200 OK + {"status":"success"}\nContoh: 400 Bad Request + {"message":"email invalid"}');

    const dataRows = Math.max(apim.getLastRow()-hdrRow, 1);
    apim.getRange(hdrRow+1,insertCol,dataRows,1).setWrap(true);

    // Fix API_Execution formula O→P
    const apie = src.getSheetByName('API_Execution');
    if (apie && apie.getLastRow() >= 9) {
        try {
            const cell = apie.getRange(9,6);
            const f    = cell.getFormula();
            if (f && f.includes('API_Master!O')) {
                cell.setFormula(f.replace(/API_Master!O(\d+):O(\d+)/g,'API_Master!P$1:P$2')
                    .replace(/API_Master!O(\d+)/g,'API_Master!P$1'));
            }
        } catch(e) {}
    }

    return 'API Expected Result col inserted at col ' + insertCol;
}


// ═══════════════════════════════════════════════════════════════════════
// FIX I — RECREATE APPENDIX
// Tulis ulang tab Appendix lengkap (hapus lama, buat baru).
// ═══════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════
// FIX J — FIX BUG BLOCKER FORMULA
// Objective: menghitung jumlah bug Open/In Progress/Reopen/Fixed/Verified
//            dengan priority Critical/High/Medium di BugReport.
//
//  Lama (9× COUNTIFS verbose):
//    =COUNTIFS(D:D,"Open",C:C,"Critical")+COUNTIFS(D:D,"Open",C:C,"High")+...
//
//  Baru (1× SUMPRODUCT bersih):
//    =SUMPRODUCT(
//      (ISNUMBER(MATCH(BugReport!D5:D2000,{"Open","In Progress","Reopen","Fixed","Verified"},0)))*
//      (ISNUMBER(MATCH(BugReport!C5:C2000,{"Critical","High","Medium"},0)))
//    )
//
//  FORCE OVERWRITE — selalu dijalankan ke semua modul (tidak ada skip).
// ═══════════════════════════════════════════════════════════════════════

function fixSimplifyFormulas_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const summ = src.getSheetByName('Summary');
    if (!summ) return 'skipped (no Summary)';

    // Formula bersih — BugReport data mulai row 5, limit 2000 untuk performa
    const FORMULA =
        '=SUMPRODUCT(' +
        '(ISNUMBER(MATCH(BugReport!D5:D2000,{"Open","In Progress","Reopen","Fixed","Verified"},0)))*' +
        '(ISNUMBER(MATCH(BugReport!C5:C2000,{"Critical","High","Medium"},0)))' +
        ')';

    // Cari baris "BUG BLOCKER" — label separator
    const blockerCell = summ.createTextFinder('BUG BLOCKER').matchEntireCell(false).findNext();
    if (!blockerCell) return 'skipped (BUG BLOCKER row not found — jalankan Smoke Blocker dulu)';

    const dataRow = blockerCell.getRow() + 1;
    let fixed = 0;

    // Kiri (col 1) — web side, biasanya merged
    try {
        const L = summ.getRange(dataRow, 1);
        const cur = L.getFormula();
        if (cur !== FORMULA) {
            L.setFormula(FORMULA);
            fixed++;
        }
    } catch(e) {
        Logger.log('Open Blocker left skip [' + spreadsheetId + ']: ' + e.message);
    }

    // Kanan (col 12) — api side, biasanya merged
    try {
        const R = summ.getRange(dataRow, 12);
        const cur = R.getFormula();
        if (cur !== FORMULA) {
            R.setFormula(FORMULA);
            fixed++;
        }
    } catch(e) {
        Logger.log('Open Blocker right skip [' + spreadsheetId + ']: ' + e.message);
    }

    return fixed === 0
        ? 'skipped (already correct)'
        : fixed + ' formula replaced at row ' + dataRow;
}

// ═══════════════════════════════════════════════════════════════════════
// FIX K — BUG BLOCKER COMPLETE FIX (Gabungan Remove Duplicate + Rename + Update Formula)
// 1. Hapus duplikat BUG BLOCKER jika ada
// 2. Rename "Open Blocker (Smoke)" → "BUG BLOCKER"
// 3. Update formula untuk include Fixed & Verified
// ═══════════════════════════════════════════════════════════════════════

function fixBugBlockerComplete_(spreadsheetId) {
    const results = [];

    // Step 1: Remove duplicates
    const step1 = removeDuplicateBugBlocker_(spreadsheetId);
    results.push('1. Remove duplicate: ' + step1);

    // Step 2: Rename or create BUG BLOCKER row
    const step2 = addSmokeBlockerRow_(spreadsheetId);
    results.push('2. Rename/Create: ' + step2);

    // Step 3: Update formula
    const step3 = fixSimplifyFormulas_(spreadsheetId);
    results.push('3. Update formula: ' + step3);

    return results.join(' | ');
}

// Broadcast function untuk menjalankan ke semua modul (PUBLIC - tanpa underscore)
function broadcastBugBlockerFix() {
    const ids = getActiveIds_();
    let ok = 0;
    const errList = [];

    ids.forEach(id => {
        try {
            const result = fixBugBlockerComplete_(id);
            Logger.log('✅ ' + id + ': ' + result);
            ok++;
        } catch (e) {
            const msg = id + ': ' + e.message;
            Logger.log('❌ ' + msg);
            errList.push(msg);
        }
    });

    const summary = ok + '/' + ids.length + ' modules fixed';
    if (errList.length > 0) {
        Browser.msgBox('BUG BLOCKER Fix Complete', summary + '\n\nErrors:\n' + errList.join('\n'), Browser.Buttons.OK);
    } else {
        Browser.msgBox('BUG BLOCKER Fix Complete', summary + '\n\nAll modules updated successfully!', Browser.Buttons.OK);
    }
}

// Alias dengan underscore untuk backward compatibility
function runBroadcastBugBlockerFix_() {
    broadcastBugBlockerFix();
}


function runBroadcastAppendix_() {
    const ids = getActiveIds_();
    let ok=0; const errList=[];
    ids.forEach(id => {
        try {
            buildAppendix_(SpreadsheetApp.openById(id));
            Logger.log('Appendix OK: ' + id);
            ok++;
        } catch(e) {
            errList.push(id + ': ' + e.message);
            Logger.log('Appendix ERR: ' + id + ' — ' + e.message);
        }
    });
    let s = 'OK:' + ok + ' Err:' + errList.length;
    if (errList.length > 0) s += '\n  • ' + errList.join('\n  • ');
    return s;
}

function buildAppendix_(ss) {
    const existing = ss.getSheetByName('Appendix');
    if (existing) ss.deleteSheet(existing);
    const ws = ss.insertSheet('Appendix');
    ws.setTabColor('#37474F');

    let r = 1;

    // Col A = white margin, content from col B
    function apxHdr_(title, bg) {
        ws.setRowHeight(r,22);
        ws.getRange(r,1).setValue('').setBackground('#FFFFFF');
        ws.getRange(r,2,1,3).merge().setValue('  ' + title)
            .setBackground(bg||'#37474F').setFontColor('#FFFFFF').setFontWeight('bold')
            .setFontSize(9).setFontFamily('Arial').setHorizontalAlignment('left').setVerticalAlignment('middle');
        r++;
    }
    function apxRow_(label, desc, labelBg, rowH) {
        ws.setRowHeight(r, rowH||48);
        ws.getRange(r,1).setValue('').setBackground('#FFFFFF');
        ws.getRange(r,2).setValue(label).setBackground(labelBg||'#FAFAFA')
            .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
            .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
        ws.getRange(r,3).setValue(desc).setBackground('#FFFFFF')
            .setFontSize(9).setFontFamily('Arial')
            .setHorizontalAlignment('left').setVerticalAlignment('top').setWrap(true);
        ws.getRange(r,4).setValue('').setBackground('#FFFFFF');
        r++;
    }
    function apxStatusRow_(label, desc, bg, rowH) {
        ws.setRowHeight(r, rowH||28);
        ws.getRange(r,1).setValue('').setBackground('#FFFFFF');
        ws.getRange(r,2).setValue(label).setBackground(bg||'#EEEEEE')
            .setFontWeight('bold').setFontSize(9).setFontFamily('Arial')
            .setHorizontalAlignment('center').setVerticalAlignment('middle');
        ws.getRange(r,3,1,2).merge().setValue(desc).setBackground('#FFFFFF')
            .setFontSize(9).setFontFamily('Arial')
            .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
        r++;
    }
    function gap_(h) {
        ws.setRowHeight(r,h||8);
        ws.getRange(r,1,1,4).setBackground('#FFFFFF');
        r++;
    }

    ws.setColumnWidth(1,16); ws.setColumnWidth(2,160); ws.setColumnWidth(3,340); ws.setColumnWidth(4,80);

    // Title
    ws.setRowHeight(r,32);
    ws.getRange(r,1).setValue('').setBackground('#FFFFFF');
    ws.getRange(r,2,1,3).merge().setValue('📋  QA Test Management  —  Panduan Penggunaan  (Template v38+)')
        .setBackground('#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold')
        .setFontSize(13).setFontFamily('Arial').setHorizontalAlignment('left').setVerticalAlignment('middle');
    r++;
    ws.setRowHeight(r,16);
    ws.getRange(r,1).setValue('').setBackground('#FFFFFF');
    ws.getRange(r,2,1,3).merge().setValue('  Dokumen ini berisi panduan pengisian semua tab. Auto-updated via Dashboard broadcast.')
        .setBackground('#E3F2FD').setFontColor('#1565C0').setFontSize(9).setFontFamily('Arial')
        .setFontStyle('italic').setHorizontalAlignment('left').setVerticalAlignment('middle');
    r++; gap_();

    // 0. Hierarki
    apxHdr_('0.  HIERARKI ORGANISASI','#263238');
    apxRow_('Project','Inisiatif / client / program kerja. Contoh: SIPGN, INAGOV, COTS.');
    apxRow_('Module','Domain dalam project. Contoh: 1 - Manajemen Gizi. Kosongkan jika project flat.');
    apxRow_('SubModule','Unit terkecil → 1 aplikasi / 1 domain → 1 Spreadsheet QA.');
    apxRow_('Feature','Fungsionalitas dalam SubModule. Contoh: Login, Dashboard, Export PDF.');
    gap_();

    // 1. Struktur Tab
    apxHdr_('1.  STRUKTUR TAB','#1565C0');
    apxRow_('Summary','KPI ringkasan: Status Overview, Smoke Test (row 17), Komposisi Status, Coverage, Automation, Bug Summary.');
    apxRow_('TC_Master','Daftar semua Test Case (Web/Mobile). 1 baris = 1 TC.');
    apxRow_('TC_Execution','Hasil eksekusi TC. 1 kolom per siklus/sprint. Kolom Z = status terakhir (formula auto).');
    apxRow_('API_Master','Daftar semua API Test Case. Wajib isi: Method, URL, Scenario, Expected Result.');
    apxRow_('API_Execution','Hasil eksekusi API. Struktur sama TC_Execution.');
    apxRow_('PerfTest','Hasil K6 performance test. DS=row16, MR=row30, Total=row46, col L = PASS/FAIL.');
    apxRow_('BugReport','Daftar semua bug. Status diisi QA & Dev sesuai alur (lihat Section 8).');
    apxRow_('Appendix','Panduan ini. Di-overwrite otomatis dari Dashboard broadcast.');
    gap_();

    // 2. Status Eksekusi
    apxHdr_('2.  STATUS EKSEKUSI TC','#2E7D32');
    [['PASSED','#C8E6C9','TC berhasil — semua Expected Result terpenuhi.'],
        ['FAILED','#FFCDD2','TC gagal — ada Expected Result yang tidak terpenuhi.'],
        ['BLOCKED','#FFE0B2','TC tidak bisa dijalankan — dependensi/blocker/env down.'],
        ['IN PROGRESS','#BBDEFB','TC sedang dalam proses eksekusi.'],
        ['TODO','#F5F5F5','TC belum dieksekusi (default awal).'],
        ['N/A','#F3E5F5','TC tidak relevan untuk siklus ini (fitur disable).'],
    ].forEach(s => apxStatusRow_(s[0],s[2],s[1]));
    gap_();

    // 3. Scenario Naming
    apxHdr_('3.  SCENARIO NAMING (Gherkin)','#4A148C');
    apxRow_('Format Dasar','Given [kondisi awal]\nWhen [aksi]\nThen [expected result]\nAnd [kondisi tambahan]');
    apxRow_('Contoh Positive','Scenario: User successfully logs in with valid credentials\nGiven user is on login page\nWhen user enters valid email and password\nThen user is redirected to dashboard\nAnd welcome message is displayed');
    apxRow_('Contoh Negative','Scenario: User fails to log in with invalid password\nGiven user is on login page\nWhen user enters valid email and wrong password\nThen error message "Wrong password" is displayed\nAnd user remains on login page');
    apxRow_('Aturan Then (Expected)','Then WAJIB berisi kondisi yang bisa diverifikasi:\n✅ Then HTTP 200 OK + {"status":"success"}\n✅ Then user melihat pesan "Data berhasil disimpan"\n❌ Then berhasil (terlalu abstrak)');
    gap_();

    // 4. Test Level
    apxHdr_('4.  TEST LEVEL','#E65100');
    apxRow_('Smoke Test','TC dengan Priority Critical, High, atau Medium.\nTarget: Pass Rate ≥80% sebelum release ke staging/production.\nDijalankan di awal sprint / setelah deployment.\nDi-track di Summary row 17 + Dashboard tab Smoke.');
    apxRow_('Regression Test','Semua TC (termasuk Low). Dijalankan lengkap di akhir sprint sebelum release final.');
    gap_();

    // 5. Priority
    apxHdr_('5.  PRIORITY','#B71C1C');
    [['Critical','#FFCDD2','Fitur utama tidak bisa digunakan. Tidak bisa release.'],
        ['High',    '#FFE0B2','Fitur penting tidak bisa digunakan, ada workaround. Harus fix before release.'],
        ['Medium',  '#FFF9C4','Fitur minor tidak bisa digunakan. Bisa release dengan catatan.'],
        ['Low',     '#E8F5E9','Kosmetik / UI. Tidak mempengaruhi fungsi. Fix di sprint berikutnya.'],
        ['Trivial', '#F3E5F5','Sangat minor, tidak terdeteksi user biasa.'],
    ].forEach(s => apxStatusRow_(s[0],s[2],s[1]));
    gap_();

    // 6. Role RBAC
    apxHdr_('6.  ROLE RBAC','#006064');
    apxRow_('Konvensi','Format: [Role] successfully/fails to [verb] [object]\nContoh: Admin successfully creates new user\nContoh: Staff fails to access admin panel');
    apxRow_('Cara Tulis','Selalu sebutkan role di awal scenario jika TC berhubungan dengan permission atau akses.','#F1F8E9');
    gap_();

    // 7. Automation Status
    apxHdr_('7.  AUTOMATION STATUS','#1A237E');
    [['Automated',   '#E8EAF6','TC sudah ada script yang berjalan di CI/CD.'],
        ['In Progress', '#E3F2FD','Script automasi sedang dibuat.'],
        ['Planned',     '#FFF9C4','Direncanakan untuk diotomasi tapi belum dikerjakan.'],
        ['Manual',      '#F5F5F5','Hanya dijalankan manual, tidak akan diotomasi.'],
    ].forEach(s => apxStatusRow_(s[0],s[2],s[1]));
    gap_();

    // 8. Bug Report
    apxHdr_('8.  BUG REPORT  —  STATUS & ALUR KERJA','#B71C1C');
    ws.setRowHeight(r,16);
    ws.getRange(r,1).setValue('').setBackground('#FFFFFF');
    ws.getRange(r,2,1,3).merge().setValue('STATUS & ALUR KERJA')
        .setBackground('#B71C1C').setFontColor('#FFCDD2').setFontWeight('bold')
        .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left').setVerticalAlignment('middle');
    r++;
    [['Open',       '#FFCDD2','Bug baru ditemukan QA. Dev belum mulai fix.'],
        ['In Progress','#FFE0B2','Dev sedang mengerjakan fix.'],
        ['Fixed',      '#BBDEFB','Dev selesai fix, deploy ke test env. Menunggu verifikasi QA.'],
        ['Verified',   '#C8E6C9','QA konfirmasi bug sudah fix. Siap close.'],
        ['Closed',     '#EEEEEE','Bug selesai & terverifikasi.'],
        ["Won't Fix",  '#F3E5F5','Tidak perlu diperbaiki (out of scope / by design).'],
        ['Reopen',     '#FFCDD2','Bug masih ada setelah diklaim Fixed. Kembali ke In Progress.'],
    ].forEach(s => apxStatusRow_(s[0],s[2],s[1]));
    ws.setRowHeight(r,16);
    ws.getRange(r,1).setValue('').setBackground('#FFFFFF');
    ws.getRange(r,2,1,3).merge().setValue('Flow:  Open → In Progress → Fixed → Verified → Closed   |   Reopen → kembali ke In Progress')
        .setBackground('#FFEBEE').setFontColor('#C62828').setFontStyle('italic')
        .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('left').setVerticalAlignment('middle');
    r++; gap_();
    apxRow_('Aturan Update Status',
        'DEV: In Progress (mulai) → Fixed (selesai, deploy)\nQA: Verified (OK) → Reopen (masih bug) → Closed (final)\nJangan skip step.');
    apxRow_('BUG BLOCKER',
        'Bug Open/In Progress/Reopen/Fixed/Verified dengan priority Critical/High/Medium.\nTarget: BUG BLOCKER = 0 sebelum release ke production.');
    gap_();

    // 9. Performance Test
    apxHdr_('9.  PERFORMANCE TEST (K6)','#004D40');
    apxRow_('Data Start Row','Row 16 (DS=16). Isi dari baris ini.','#E0F2F1');
    apxRow_('Max Row','Row 30 (MR=30). Batas bawah data per skenario.','#E0F2F1');
    apxRow_('Total Row','Row 46. Agregasi total semua skenario.','#E0F2F1');
    apxRow_('Kolom L (PASS/FAIL)','Formula otomatis. PASS jika semua threshold terpenuhi.','#E0F2F1');
    apxRow_('Threshold Standar','p95 < 2000ms  |  Error Rate < 1%  |  RPS sesuai target skenario','#E0F2F1');
    gap_();

    // 10. HTTP Methods
    apxHdr_('10.  HTTP METHODS (API)','#1565C0');
    [['GET','Mengambil data. Tidak mengubah state server. Idempotent.'],
        ['POST','Membuat resource baru. Tidak idempotent.'],
        ['PUT','Update resource existing (replace seluruh resource). Idempotent.'],
        ['PATCH','Update sebagian field resource.'],
        ['DELETE','Menghapus resource. Idempotent.'],
    ].forEach(s => apxRow_(s[0],s[1],'#E3F2FD',28));
    gap_();

    // Footer
    ws.setRowHeight(r,22);
    ws.getRange(r,1).setValue('').setBackground('#FFFFFF');
    ws.getRange(r,2,1,3).merge().setValue('QA Team  ·  Template v38+  ·  2026  ·  Auto-updated via Dashboard Broadcast')
        .setBackground('#0D47A1').setFontColor('#BBDEFB')
        .setFontSize(8).setFontFamily('Arial').setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setFrozenRows(0);
}


// ═══════════════════════════════════════════════════════════════════════
// FIX K — UPDATE SUMMARY FIELDS (Project, Modul, Submodul restructure)
// Ubah "Project / Sprint" → "Project", tambah "Modul" dan "Submodul"
// Shift QA Lead & PIC QA down 1 row
// ═══════════════════════════════════════════════════════════════════════

function fixSummaryFields_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const summ = src.getSheetByName('Summary');
    if (!summ) return 'skipped (no Summary)';

    try {
        // Cek apakah sudah di-update (cek label di A3 - seharusnya "Modul:" kalau sudah update)
        const a3Label = String(summ.getRange(3, 1).getValue()).trim();
        Logger.log('[fixSummaryFields] A3 label: "' + a3Label + '"');

        if (a3Label === 'Modul:') {
            Logger.log('[fixSummaryFields] Already updated, skipping');
            return 'skipped (already updated)';
        }

        // Backup values yang akan di-shift
        const oldA2 = String(summ.getRange(2, 1).getValue()).trim();  // Label: Project/Sprint atau Period
        const oldB2 = summ.getRange(2, 2).getValue();  // Project/Sprint value
        const oldA3 = String(summ.getRange(3, 1).getValue()).trim();  // Label: bisa Period atau QA Lead
        const oldB3 = summ.getRange(3, 2).getValue();  // Period value atau env
        const oldA4 = String(summ.getRange(4, 1).getValue()).trim();  // Label
        const oldB4 = summ.getRange(4, 2).getValue();  // QA Lead value
        const oldA5 = String(summ.getRange(5, 1).getValue()).trim();  // Label
        const oldB5 = summ.getRange(5, 2).getValue();  // PIC QA value

        Logger.log('[fixSummaryFields] Old structure: A2="'+oldA2+'" A3="'+oldA3+'" A4="'+oldA4+'" A5="'+oldA5+'"');

        // Update labels (col A)
        summ.getRange(2, 1).setValue('Project:');
        summ.getRange(3, 1).setValue('Modul:');
        summ.getRange(4, 1).setValue('Submodul:');
        summ.getRange(5, 1).setValue('QA Lead:');
        summ.getRange(6, 1).setValue('PIC QA:');

        // Update values (col B)
        // B2 = Project (keep old value jika ada)
        if (oldB2) summ.getRange(2, 2).setValue(oldB2);

        // B3 = Modul (kosongkan, akan diisi manual atau dari config)
        summ.getRange(3, 2).setValue('');

        // B4 = Submodul (ambil dari old B3 jika labelnya "Period")
        const submodVal = (oldA3.toLowerCase().includes('period') || oldA3.toLowerCase().includes('sprint')) ? oldB3 : '';
        summ.getRange(4, 2).setValue(submodVal || '');

        // Add note untuk Submodul
        summ.getRange(4, 2).setNote('Pisahkan dengan koma jika lebih dari 1 submodul.\nContoh: 1.1,1.2,1.3');

        // B5 = QA Lead (shift dari old B4 jika labelnya QA)
        const qaLeadVal = (oldA4.toLowerCase().includes('qa') || oldA4.toLowerCase().includes('lead')) ? oldB4 : '';
        summ.getRange(5, 2).setValue(qaLeadVal || '');

        // B6 = PIC QA (shift dari old B5 jika labelnya PIC)
        const picVal = (oldA5.toLowerCase().includes('pic') || oldA5.toLowerCase().includes('qa')) ? oldB5 : '';
        summ.getRange(6, 2).setValue(picVal || '');

        Logger.log('[fixSummaryFields] Updated successfully');
        return 'updated (fields restructured)';

    } catch(e) {
        Logger.log('[fixSummaryFields] Error: ' + e.message);
        return 'error: ' + e.message;
    }
}


// ═══════════════════════════════════════════════════════════════════════
// FIX L — UPDATE QA TEAM LEAD LABEL
// Ubah "QA Team Lead:" → "QA Lead:" di Summary
// ═══════════════════════════════════════════════════════════════════════

function fixQALeadLabel_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const summ = src.getSheetByName('Summary');
    if (!summ) return 'skipped (no Summary)';

    try {
        // Cari label "QA Team Lead:" di kolom A (biasanya di row 4 atau 5)
        const searchRange = summ.getRange('A1:A20');
        const finder = searchRange.createTextFinder('QA Team Lead:').matchEntireCell(true);
        const found = finder.findNext();

        if (!found) {
            // Cek apakah sudah "QA Lead:" (sudah di-update)
            const finderNew = searchRange.createTextFinder('QA Lead:').matchEntireCell(true);
            if (finderNew.findNext()) {
                return 'skipped (already QA Lead:)';
            }
            return 'skipped (label not found)';
        }

        // Update label
        const row = found.getRow();
        summ.getRange(row, 1).setValue('QA Lead:');
        Logger.log('[fixQALeadLabel] Updated row ' + row + ' from "QA Team Lead:" to "QA Lead:"');

        return 'updated (QA Team Lead → QA Lead)';

    } catch(e) {
        Logger.log('[fixQALeadLabel] Error: ' + e.message);
        return 'error: ' + e.message;
    }
}


// ═══════════════════════════════════════════════════════════════════════
// FIX M — UPDATE SUBMODUL NOTE
// Update note di Submodul: "Login, Dashboard, Profile" → "1.1,1.2,1.3"
// ═══════════════════════════════════════════════════════════════════════

function fixSubmodulNote_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const summ = src.getSheetByName('Summary');
    if (!summ) return 'skipped (no Summary)';

    try {
        // Cari label "Submodul:" di kolom A
        const searchRange = summ.getRange('A1:A20');
        const finder = searchRange.createTextFinder('Submodul:').matchEntireCell(true);
        const found = finder.findNext();

        if (!found) return 'skipped (Submodul label not found)';

        const row = found.getRow();
        const valueCell = summ.getRange(row, 2);  // Col B

        // Cek note yang ada
        const currentNote = valueCell.getNote();

        // Jika note sudah benar, skip
        if (currentNote && currentNote.includes('1.1,1.2,1.3')) {
            return 'skipped (note already updated)';
        }

        // Update note
        valueCell.setNote('Pisahkan dengan koma jika lebih dari 1 submodul.\nContoh: 1.1,1.2,1.3');
        Logger.log('[fixSubmodulNote] Updated note at row ' + row);

        return 'updated (note → 1.1,1.2,1.3)';

    } catch(e) {
        Logger.log('[fixSubmodulNote] Error: ' + e.message);
        return 'error: ' + e.message;
    }
}


// ═══════════════════════════════════════════════════════════════════════
// FIX N — REMOVE PIC QA DROPDOWN
// Hapus data validation (dropdown) di PIC QA, jadikan free text
// ═══════════════════════════════════════════════════════════════════════

function fixPICQADropdown_(spreadsheetId) {
    const src  = SpreadsheetApp.openById(spreadsheetId);
    const summ = src.getSheetByName('Summary');
    if (!summ) return 'skipped (no Summary)';

    try {
        // Cari label "PIC QA:" atau "PIC / Team:" di kolom A
        const searchRange = summ.getRange('A1:A20');
        let finder = searchRange.createTextFinder('PIC QA:').matchEntireCell(true);
        let found = finder.findNext();

        if (!found) {
            // Try alternatif label
            finder = searchRange.createTextFinder('PIC / Team:').matchEntireCell(true);
            found = finder.findNext();
        }

        if (!found) return 'skipped (PIC QA label not found)';

        const row = found.getRow();
        const valueCell = summ.getRange(row, 2);  // Col B

        // Cek apakah ada data validation
        const validation = valueCell.getDataValidation();
        if (!validation) {
            return 'skipped (already free text)';
        }

        // Hapus data validation
        valueCell.setDataValidation(null);
        Logger.log('[fixPICQADropdown] Removed dropdown at row ' + row);

        return 'updated (dropdown removed)';

    } catch(e) {
        Logger.log('[fixPICQADropdown] Error: ' + e.message);
        return 'error: ' + e.message;
    }
}


