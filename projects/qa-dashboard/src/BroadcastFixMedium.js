/**
 * broadcast_fix_medium_wording.js
 * ─────────────────────────────────────────────────────────────────────────
 * Paste ke Apps Script editor QA PORTFOLIO DASHBOARD.
 *
 * FUNGSI:
 *   broadcastFixMediumWording()  → fix semua modul aktif
 *   fixMediumSingleSheet()       → fix satu modul (minta input ID)
 *
 * PERUBAHAN:
 *   Summary tab — Bug Priority label:
 *     "Medium (Blocker):" → "Medium:"
 *     "Medium (Blocker)"  → "Medium"   (fallback tanpa colon)
 *
 * Pakai TextFinder — tidak terpengaruh merged cells, formula, atau spacing.
 * ─────────────────────────────────────────────────────────────────────────
 */

function broadcastFixMediumWording() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName('Config');
  if (!cfg) { Logger.log('ERROR: Config tab tidak ditemukan.'); return; }

  var allData  = cfg.getDataRange().getValues();
  var hdrRow   = allData[2] || [];
  var idColIdx = detectIdColMed_(hdrRow);

  var ok = 0, skip = 0, err = 0, errList = [];

  for (var i = 3; i < allData.length; i++) {
    var row = allData[i];
    if (String(row[0]).trim().toUpperCase() !== 'Y') continue;
    var id = String(row[idColIdx]).trim();
    if (!id || id.length < 10 || id === 'PASTE_SPREADSHEET_ID_HERE') continue;

    try {
      var res = fixMediumInSheet_(id);
      if (res === 'skipped') {
        skip++;
        Logger.log('SKIP: ' + id.substring(0, 25));
      } else {
        ok++;
        Logger.log('OK  : ' + id.substring(0, 25) + ' — ' + res);
      }
    } catch(e) {
      err++;
      errList.push('  • ' + id.substring(0, 25) + ': ' + e.message);
      Logger.log('ERR : ' + id.substring(0, 25) + ' — ' + e.message);
    }
  }

  var summary =
    'broadcastFixMediumWording selesai\n\n' +
    'Berhasil : ' + ok   + ' modul\n' +
    'Skip     : ' + skip + ' modul (tidak ada "Medium (Blocker)")\n' +
    'Gagal    : ' + err  + ' modul';
  if (errList.length) summary += '\n' + errList.join('\n');
  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert(summary); } catch(e) {}
}


// ── Core fix for one spreadsheet ──────────────────────────────────────────
function fixMediumInSheet_(spreadsheetId) {
  var src  = SpreadsheetApp.openById(spreadsheetId);
  var summ = src.getSheetByName('Summary');
  if (!summ) return 'skipped (no Summary tab)';

  var totalFixed = 0;

  // TextFinder handles merged cells, all cell types, no issue with getValues
  // Fix "Medium (Blocker):" (with colon — how the label is written)
  var finder1 = summ.createTextFinder('Medium (Blocker):')
    .matchEntireCell(true)
    .matchCase(true);
  var found1 = finder1.findAll();
  found1.forEach(function(cell) {
    cell.setValue('Medium:');
    totalFixed++;
    Logger.log('  Fixed cell ' + cell.getA1Notation() + ': "Medium (Blocker):" → "Medium:"');
  });

  // Also catch without colon (fallback)
  var finder2 = summ.createTextFinder('Medium (Blocker)')
    .matchEntireCell(true)
    .matchCase(true);
  var found2 = finder2.findAll();
  found2.forEach(function(cell) {
    cell.setValue('Medium');
    totalFixed++;
    Logger.log('  Fixed cell ' + cell.getA1Notation() + ': "Medium (Blocker)" → "Medium"');
  });

  // Also fix inside formulas (COUNTIF referencing "Medium (Blocker)")
  var finder3 = summ.createTextFinder('Medium \\(Blocker\\)')
    .matchEntireCell(false)
    .useRegularExpression(true);
  var found3 = finder3.findAll();
  var formulaFixed = 0;
  found3.forEach(function(cell) {
    var f = cell.getFormula();
    if (f && f.indexOf('Medium (Blocker)') > -1) {
      cell.setFormula(f.replace(/Medium \(Blocker\)/g, 'Medium'));
      formulaFixed++;
      Logger.log('  Fixed formula at ' + cell.getA1Notation());
    }
  });

  if (totalFixed === 0 && formulaFixed === 0) return 'skipped';

  return totalFixed + ' label(s) fixed' +
    (formulaFixed ? ', ' + formulaFixed + ' formula(s) fixed' : '');
}


// ── Single sheet test ─────────────────────────────────────────────────────
function fixMediumSingleSheet() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Fix Single Sheet',
    'Masukkan Spreadsheet ID modul:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var id = resp.getResponseText().trim();
  if (!id) { ui.alert('ID tidak boleh kosong.'); return; }

  try {
    var result = fixMediumInSheet_(id);
    ui.alert(result === 'skipped'
      ? '⏭  Skip — "Medium (Blocker)" tidak ditemukan di Summary.'
      : '✅  ' + result);
  } catch(e) {
    ui.alert('❌  Error: ' + e.message);
  }
}


// ── Helper ────────────────────────────────────────────────────────────────
function detectIdColMed_(headerRow) {
  for (var i = 0; i < headerRow.length; i++) {
    var h = String(headerRow[i]).trim().toUpperCase();
    if (h === 'SPREADSHEET ID' || h === 'SPREADSHEET_ID') return i;
  }
  return 6;
}