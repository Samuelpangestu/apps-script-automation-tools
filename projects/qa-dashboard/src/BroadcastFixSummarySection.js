/**
 * broadcast_fix_summary_sections.js
 * ─────────────────────────────────────────────────────────────────────────
 * Paste ke Apps Script editor QA SHEET (bukan Dashboard).
 * Jalankan dari masing-masing modul, ATAU gunakan broadcastFromDashboard
 * jika di-paste ke Dashboard dan Config berisi Spreadsheet IDs.
 *
 * PERUBAHAN:
 *   Summary tab:
 *     D. COVERAGE PER SUBMODUL  →  E. COVERAGE PER SUBMODUL  (turun)
 *     E. BUG SUMMARY            →  D. BUG SUMMARY             (naik)
 *     Coverage rows             :  12 → 34 rows + 1 TOTAL = 35 rows
 *
 * CARA PAKAI:
 *   Dari modul QA Sheet       : jalankan fixSummarySections()
 *   Dari Dashboard (broadcast): jalankan broadcastFixSummarySections()
 * ─────────────────────────────────────────────────────────────────────────
 */


// ══════════════════════════════════════════════════════════════════════════
// BROADCAST — jalankan dari Dashboard
// ══════════════════════════════════════════════════════════════════════════
function broadcastFixSummarySections() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');
  if (!cfg) { Logger.log('ERROR: Config tab tidak ditemukan.'); return; }

  var allData  = cfg.getDataRange().getValues();
  var hdrRow   = allData[2] || [];
  var idColIdx = detectIdColSec_(hdrRow);

  var ok = 0, skip = 0, err = 0, errList = [];

  for (var i = 3; i < allData.length; i++) {
    var row = allData[i];
    if (String(row[0]).trim().toUpperCase() !== 'Y') continue;
    var id = String(row[idColIdx]).trim();
    if (!id || id.length < 10 || id === 'PASTE_SPREADSHEET_ID_HERE') continue;

    try {
      var src = SpreadsheetApp.openById(id);
      var res = fixSummarySectionsInSheet_(src);
      if (res === 'skipped') { skip++; Logger.log('SKIP: ' + id.substring(0,25)); }
      else                   { ok++;   Logger.log('OK  : ' + id.substring(0,25) + ' — ' + res); }
    } catch(e) {
      err++;
      errList.push('  • ' + id.substring(0,25) + ': ' + e.message);
      Logger.log('ERR : ' + id.substring(0,25) + ' — ' + e.message);
    }
  }

  var summary =
    'broadcastFixSummarySections selesai\n\n' +
    'Berhasil : ' + ok   + ' modul\n' +
    'Skip     : ' + skip + ' modul\n' +
    'Gagal    : ' + err  + ' modul';
  if (errList.length) summary += '\n' + errList.join('\n');
  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert(summary); } catch(e) {}
}


// ══════════════════════════════════════════════════════════════════════════
// SINGLE SHEET — jalankan dari modul QA Sheet langsung
// ══════════════════════════════════════════════════════════════════════════
function fixSummarySections() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var res = fixSummarySectionsInSheet_(ss);
  try {
    SpreadsheetApp.getUi().alert(
      res === 'skipped' ? '⏭  Skip — Summary tidak perlu diubah.' : '✅  ' + res
    );
  } catch(e) { Logger.log(res); }
}


// ══════════════════════════════════════════════════════════════════════════
// CORE FIX
// ══════════════════════════════════════════════════════════════════════════
function fixSummarySectionsInSheet_(ss) {
  var ws = ss.getSheetByName('Summary');
  if (!ws) return 'skipped (no Summary tab)';

  // ── Layout constants (must match buildSummary) ──
  var L = 1, R_ = 12, LW = 10, RW = 10;
  var MAX_NEW = 34; // 34 data rows + 1 TOTAL = 35 rows total

  // ── Find section header rows via TextFinder ──
  var covRow = findSectionRow_(ws, 'COVERAGE PER SUBMODUL');
  var bugRow = findSectionRow_(ws, 'BUG SUMMARY');

  if (covRow === -1 && bugRow === -1) return 'skipped (sections not found)';

  Logger.log('  Found Coverage at row ' + covRow + ', Bug Summary at row ' + bugRow);

  // Determine current order
  var alreadyCorrect = (bugRow !== -1 && covRow !== -1 && bugRow < covRow);

  // ── Step 1: Rename section labels ──
  renameSectionLabel_(ws, 'D.  COVERAGE PER SUBMODUL', 'E.  COVERAGE PER SUBMODUL');
  renameSectionLabel_(ws, 'E.  COVERAGE PER SUBMODUL', 'E.  COVERAGE PER SUBMODUL'); // already correct
  renameSectionLabel_(ws, 'E.  BUG SUMMARY', 'D.  BUG SUMMARY');
  renameSectionLabel_(ws, 'D.  BUG SUMMARY', 'D.  BUG SUMMARY'); // already correct

  var log = [];

  // ── Step 2: Expand Coverage section from 12 rows to 34 rows ──
  // Find the current Coverage header row (after potential rename)
  var covRowCurrent = findSectionRow_(ws, 'COVERAGE PER SUBMODUL');
  if (covRowCurrent !== -1) {
    var expanded = expandCoverageRows_(ws, covRowCurrent, MAX_NEW, L, R_, LW, RW, ss);
    log.push(expanded);
  }

  // ── Step 3: Swap sections if Bug is below Coverage ──
  // Re-detect after possible expansion
  var covRowFinal = findSectionRow_(ws, 'COVERAGE PER SUBMODUL');
  var bugRowFinal = findSectionRow_(ws, 'BUG SUMMARY');

  if (covRowFinal !== -1 && bugRowFinal !== -1 && covRowFinal < bugRowFinal) {
    var swapped = swapSections_(ws, covRowFinal, bugRowFinal);
    log.push('Sections swapped: Bug now above Coverage');
  } else if (bugRowFinal < covRowFinal) {
    log.push('Order already correct (Bug above Coverage)');
  }

  return log.join('; ') || 'done';
}


// ── Rename all matching section header cells ──────────────────────────────
function renameSectionLabel_(ws, oldText, newText) {
  if (oldText === newText) return;
  var finder = ws.createTextFinder(oldText).matchEntireCell(false);
  var cells  = finder.findAll();
  cells.forEach(function(cell) {
    var v = cell.getValue();
    cell.setValue(v.replace(oldText, newText));
  });
}


// ── Find row number of a section header ──────────────────────────────────
function findSectionRow_(ws, keyword) {
  var finder = ws.createTextFinder(keyword).matchEntireCell(false);
  var cell   = finder.findNext();
  return cell ? cell.getRow() : -1;
}


// ── Expand Coverage section from current size to MAX_NEW rows ─────────────
function expandCoverageRows_(ws, headerRow, maxNew, L, R_, LW, RW, ss) {
  // Sub-header row = headerRow + 1 (SubModul, Total, Smoke, Regression, Auto%, Pass%)
  // Data starts at headerRow + 2
  var subHdrRow = headerRow + 1;
  var dataStart = headerRow + 2;

  // Detect current MAX by counting consecutive non-TOTAL rows
  var currentMax = 0;
  var lastRow = ws.getLastRow();
  for (var r = dataStart; r <= lastRow; r++) {
    var v = String(ws.getRange(r, L).getValue()).trim();
    if (v === 'TOTAL') break;
    currentMax++;
  }

  if (currentMax >= maxNew) {
    return 'Coverage already has ' + currentMax + ' rows — skip expand';
  }

  var rowsToAdd = maxNew - currentMax;
  var insertAfterRow = dataStart + currentMax - 1; // last data row before TOTAL

  // Insert blank rows before TOTAL
  ws.insertRowsAfter(insertAfterRow, rowsToAdd);
  Logger.log('  Inserted ' + rowsToAdd + ' rows after row ' + insertAfterRow);

  // Re-write formulas for new rows (Web/Mobile left side)
  writeCovRows_(ws, dataStart, maxNew, L, 'TC_Master', 'B', 'E', 'H', 'TC_Execution', '#0D47A1');
  // Re-write formulas for new rows (API right side)
  writeCovRows_(ws, dataStart, maxNew, R_, 'API_Master', 'B', 'G', 'J', 'API_Execution', '#283593');

  // Re-write TOTAL row (now at dataStart + maxNew)
  var totRow = dataStart + maxNew;
  writeCovTotal_(ws, totRow, L,  maxNew, dataStart, 'TC_Master',  'B', 'E', 'H', 'TC_Execution',  '#0D47A1');
  writeCovTotal_(ws, totRow, R_, maxNew, dataStart, 'API_Master', 'B', 'G', 'J', 'API_Execution', '#283593');

  return 'Coverage expanded: ' + currentMax + ' → ' + maxNew + ' rows (+' + rowsToAdd + ')';
}


// ── Write Coverage data row formulas ─────────────────────────────────────
function writeCovRows_(ws, dataStart, maxNew, sc, master, subCol, prioCol, autoCol, execSh, hbg) {
  for (var idx = 0; idx < maxNew; idx++) {
    var row = dataStart + idx;
    var bg  = idx % 2 === 0 ? (hbg === '#0D47A1' ? '#F8F9FA' : '#F0F4FF') : '#FFFFFF';
    var ref = colLetterSec_(sc) + row;

    var cell0 = ws.getRange(row, sc);
    cell0.setFormula(
      '=IFERROR(INDEX(UNIQUE(FILTER(' + master + '!' + subCol + '3:' + subCol + '1000,' +
      master + '!' + subCol + '3:' + subCol + '1000<>"")),'+  (idx+1) + ',1),"")'
    ).setBackground(bg).setFontFamily('Arial').setFontSize(9)
     .setHorizontalAlignment('left').setFontWeight('bold')
     .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);

    ws.getRange(row, sc+1).setFormula(
      '=IF(' + ref + '="","",COUNTIF(' + master + '!' + subCol + '3:' + subCol + '1000,' + ref + '))'
    ).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center')
     .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);

    ws.getRange(row, sc+2).setFormula(
      '=IF(' + ref + '="","",COUNTIFS(' + master + '!' + subCol + '3:' + subCol + '1000,' + ref +
      ',' + master + '!' + prioCol + '3:' + prioCol + '1000,"Critical")+COUNTIFS(' +
      master + '!' + subCol + '3:' + subCol + '1000,' + ref +
      ',' + master + '!' + prioCol + '3:' + prioCol + '1000,"High")+COUNTIFS(' +
      master + '!' + subCol + '3:' + subCol + '1000,' + ref +
      ',' + master + '!' + prioCol + '3:' + prioCol + '1000,"Medium"))'
    ).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center')
     .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);

    ws.getRange(row, sc+3).setFormula(
      '=IF(' + ref + '="","",COUNTIFS(' + master + '!' + subCol + '3:' + subCol + '1000,' + ref +
      ',' + master + '!' + prioCol + '3:' + prioCol + '1000,"Low")+COUNTIFS(' +
      master + '!' + subCol + '3:' + subCol + '1000,' + ref +
      ',' + master + '!' + prioCol + '3:' + prioCol + '1000,"Lowest"))'
    ).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center')
     .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);

    ws.getRange(row, sc+4).setFormula(
      '=IFERROR(COUNTIFS(' + master + '!' + subCol + '3:' + subCol + '1000,' + ref +
      ',' + master + '!' + autoCol + '3:' + autoCol + '1000,"Automated")/' +
      colLetterSec_(sc+1) + row + ',0)'
    ).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center')
     .setNumberFormat('0%')
     .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);

    ws.getRange(row, sc+5).setFormula(
      '=IFERROR(COUNTIFS(' + execSh + '!B9:B1000,' + ref +
      ',' + execSh + '!Z9:Z1000,"PASSED")/MAX(1,COUNTIF(' + execSh + '!B9:B1000,' + ref + ')),0)'
    ).setBackground(bg).setFontFamily('Arial').setFontSize(9).setHorizontalAlignment('center')
     .setNumberFormat('0%')
     .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);

    ws.setRowHeight(row, 16);
  }
}


// ── Write Coverage TOTAL row ──────────────────────────────────────────────
function writeCovTotal_(ws, totRow, sc, maxNew, dataStart, master, subCol, prioCol, autoCol, execSh, hbg) {
  var totBg = hbg === '#0D47A1' ? '#E3F2FD' : '#E8EAF6';
  var DS    = dataStart;
  var DE    = dataStart + maxNew - 1; // last data row

  var values = [
    'TOTAL',
    '=SUM(' + colLetterSec_(sc+1) + DS + ':' + colLetterSec_(sc+1) + DE + ')',
    '=SUM(' + colLetterSec_(sc+2) + DS + ':' + colLetterSec_(sc+2) + DE + ')',
    '=SUM(' + colLetterSec_(sc+3) + DS + ':' + colLetterSec_(sc+3) + DE + ')',
    '=IFERROR(COUNTIF(' + master + '!' + autoCol + '3:' + autoCol + '1000,"Automated")/MAX(1,COUNTA(' + master + '!' + subCol + '3:' + subCol + '1000)),0)',
    '=IFERROR(COUNTIF(' + execSh + '!Z9:Z1000,"PASSED")/MAX(1,COUNTA(' + master + '!' + subCol + '3:' + subCol + '1000)),0)',
  ];

  values.forEach(function(v, i) {
    var cell = ws.getRange(totRow, sc + i)
      .setBackground(totBg).setFontWeight('bold')
      .setFontFamily('Arial').setFontSize(9)
      .setHorizontalAlignment(i === 0 ? 'left' : 'center')
      .setBorder(true,true,true,true,false,false,'#CFD8DC',SpreadsheetApp.BorderStyle.SOLID);
    if (typeof v === 'string' && v.startsWith('=')) cell.setFormula(v);
    else cell.setValue(v);
    if (i >= 4) cell.setNumberFormat('0%');
  });
  ws.setRowHeight(totRow, 18);
}


// ── Swap Coverage and Bug Summary sections ────────────────────────────────
function swapSections_(ws, covRow, bugRow) {
  // Coverage is above Bug Summary (covRow < bugRow)
  // We need Bug Summary to be above Coverage
  // Strategy: cut Bug Summary rows, insert above Coverage

  var lastRow      = ws.getLastRow();
  var bugEndRow    = covRow - 1;   // Bug Summary ends just before Coverage
  var bugRowCount  = bugEndRow - bugRow + 1;
  var covEndRow    = findSectionEnd_(ws, covRow, lastRow);
  var covRowCount  = covEndRow - covRow + 1;

  Logger.log('  Coverage: rows ' + covRow + '–' + covEndRow + ' (' + covRowCount + ' rows)');
  Logger.log('  Bug Sum:  rows ' + bugRow + '–' + bugEndRow + ' (' + bugRowCount + ' rows)');

  // Cut Bug Summary (copy values + formulas to temp, delete, insert before Coverage)
  // Approach: use moveTo / copy then delete
  // Simplest robust approach: insert blank rows before Coverage, copy Bug into them, delete old Bug

  // 1. Insert bugRowCount blank rows right before Coverage
  ws.insertRowsBefore(covRow, bugRowCount);

  // 2. Bug Summary is now at bugRow + bugRowCount (shifted down by bugRowCount)
  var bugRowShifted = bugRow + bugRowCount;

  // 3. Copy Bug Summary rows to the newly inserted blank rows
  var bugRange  = ws.getRange(bugRowShifted, 1, bugRowCount, ws.getLastColumn());
  var destRange = ws.getRange(covRow,        1, bugRowCount, ws.getLastColumn());
  bugRange.copyTo(destRange);

  // 4. Delete the original Bug Summary rows (now shifted further down)
  ws.deleteRows(bugRowShifted, bugRowCount);

  return 'swapped';
}


// ── Find where a section ends (next section header or end of sheet) ────────
function findSectionEnd_(ws, sectionStartRow, lastRow) {
  // A section ends when we find the next section header (cell starting with a letter + period)
  // or reach the last data row
  var sectionLetters = ['A.','B.','C.','D.','E.','F.','G.'];
  for (var r = sectionStartRow + 1; r <= lastRow; r++) {
    var v = String(ws.getRange(r, 1).getValue()).trim();
    for (var s = 0; s < sectionLetters.length; s++) {
      if (v.indexOf(sectionLetters[s]) === 0) return r - 1;
    }
    // Also check col 12 (right side headers)
    var v2 = String(ws.getRange(r, 12).getValue()).trim();
    for (var s2 = 0; s2 < sectionLetters.length; s2++) {
      if (v2.indexOf(sectionLetters[s2]) === 0) return r - 1;
    }
  }
  return lastRow;
}


// ── Column letter helper ──────────────────────────────────────────────────
function colLetterSec_(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter  = String.fromCharCode(65 + rem) + letter;
    col     = Math.floor((col - 1) / 26);
  }
  return letter;
}


// ── Config ID column detector ─────────────────────────────────────────────
function detectIdColSec_(headerRow) {
  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i]).trim().toUpperCase();
    if (h === 'SPREADSHEET ID' || h === 'SPREADSHEET_ID') return i;
  }
  return 6;
}