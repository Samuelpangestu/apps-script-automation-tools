/**
 * InitVAPTTabs.js - Initialize 4 VAPT tabs in QATM template
 *
 * Tabs created:
 * 1. Detail Finding - Regular VAPT
 * 2. Evidence - Regular VAPT
 * 3. Detail Finding - Ad Hoc VAPT
 * 4. Evidence - Ad Hoc VAPT
 *
 * Usage:
 * - For new templates: Called from createQASheet()
 * - For existing: Use broadcastVAPTTabs() in BroadcastVAPTTabs.js
 */

// ═══════════════════════════════════════════════════════════════════════
// VAPT TAB COLORS
// ═══════════════════════════════════════════════════════════════════════

const VAPT_COLORS = {
  regular: '#FF6F00',     // Deep Orange for Regular VAPT
  adhoc: '#7B1FA2',       // Purple for Ad Hoc VAPT
  header: '#263238',      // Dark Blue Grey for headers
  white: '#FFFFFF',
  evidence: '#004D40',    // Teal for Evidence tabs
};

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Create data validation dropdown
 */
function vaptDropdown_(values) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
}

/**
 * Apply border to range
 */
function vaptBorder_(range) {
  return range.setBorder(
    true, true, true, true, false, false,
    '#CFD8DC', SpreadsheetApp.BorderStyle.SOLID
  );
}

/**
 * Create header cell
 */
function vaptHeader_(range, bg, fg, size) {
  fg = fg || VAPT_COLORS.white;
  size = size || 9;
  return vaptBorder_(range)
    .setBackground(bg)
    .setFontColor(fg)
    .setFontWeight('bold')
    .setFontSize(size)
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

// ═══════════════════════════════════════════════════════════════════════
// TAB 1: Detail Finding - Regular VAPT
// ═══════════════════════════════════════════════════════════════════════

function createDetailFindingRegularVAPT(ss) {
  const ws = ss.getSheetByName('Detail Finding - Regular VAPT') ||
             ss.insertSheet('Detail Finding - Regular VAPT');

  ws.clear();
  ws.setTabColor(VAPT_COLORS.regular);

  // Title row
  const title = 'VAPT DETAIL FINDING  .  Regular VAPT  .  QA PERURI';
  ws.getRange(1, 1, 1, 33).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.header, VAPT_COLORS.white, 11)
    .setValue(title);
  ws.setRowHeight(1, 32);

  // Header row
  const headers = [
    'Finding ID',                    // A
    'Application',                   // B
    'Category',                      // C
    'Scope',                         // D
    'Status Fix (Dev)',              // E
    'Status Re-VAPT (Pentester)',    // F
    'Original Risk',                 // G
    'Adjusted Risk',                 // H
    'Finding Name',                  // I
    'Description',                   // J
    'Impact',                        // K
    'Affected URL/Endpoint',         // L
    'Recommendation',                // M
    'Remediation Steps',             // N
    'Report Date',                   // O
    'Pentester',                     // P
    'Target Remediation Date',       // Q
    'Time to Remediate (Days)',      // R
    'Actual Fix Date',               // S
    'Verified By (QA)',              // T
    'Verified Date',                 // U
    'CVSS Score',                    // V
    'CWE ID',                        // W
    'OWASP Top 10',                  // X
    'PoC Available',                 // Y
    'Re-Test Required',              // Z
    'Re-Test Date',                  // AA
    'Re-Test Result',                // AB
    'Notes / Comments',              // AC
    'Jira Ticket',                   // AD
    'Acceptance Proof / MAoR',       // AE
    'Created By',                    // AF
    'Last Updated',                  // AG
  ];

  headers.forEach((header, i) => {
    vaptHeader_(ws.getRange(2, i + 1), VAPT_COLORS.regular, VAPT_COLORS.white)
      .setValue(header);
  });
  ws.setRowHeight(2, 50);

  // Set column widths
  const widths = [
    120, 120, 100, 80, 120, 120, 90, 90, 200,   // A-I
    250, 180, 180, 220, 220, 100,                // J-O
    100, 120, 100, 100, 100, 100,                // P-U
    80, 80, 100, 80, 80, 100, 100,               // V-AB
    200, 120, 200, 100, 120                      // AC-AG
  ];
  widths.forEach((width, i) => ws.setColumnWidth(i + 1, width));

  // Data validation dropdowns
  const lastRow = 1000;

  // E: Status Fix (Dev)
  const statusFixValues = [
    'Todo',
    'On Progress Remediation',
    'Ready to Retest',
    'On Progress Retest',
    'Done',
    'Accepted',
    'False Positive',
    'Duplicated',
    'Out of Scope'
  ];
  ws.getRange(`E3:E${lastRow}`).setDataValidation(vaptDropdown_(statusFixValues));

  // F: Status Re-VAPT (Pentester)
  const statusReVAPTValues = ['Open', 'Closed'];
  ws.getRange(`F3:F${lastRow}`).setDataValidation(vaptDropdown_(statusReVAPTValues));

  // G & H: Risk levels
  const riskValues = ['Informational', 'Low', 'Medium', 'High', 'Critical', 'False Positive'];
  ws.getRange(`G3:G${lastRow}`).setDataValidation(vaptDropdown_(riskValues));
  ws.getRange(`H3:H${lastRow}`).setDataValidation(vaptDropdown_(riskValues));

  // Y: PoC Available
  ws.getRange(`Y3:Y${lastRow}`).setDataValidation(vaptDropdown_(['Yes', 'No']));

  // Z: Re-Test Required
  ws.getRange(`Z3:Z${lastRow}`).setDataValidation(vaptDropdown_(['Yes', 'No']));

  // AB: Re-Test Result
  ws.getRange(`AB3:AB${lastRow}`).setDataValidation(vaptDropdown_(['Pass', 'Fail', 'Pending']));

  // Conditional formatting for Status Fix
  addVAPTStatusConditionalFormatting_(ws, `E3:E${lastRow}`);

  // Conditional formatting for Status Re-VAPT
  addVAPTReVAPTConditionalFormatting_(ws, `F3:F${lastRow}`);

  // Conditional formatting for Risk levels
  addRiskConditionalFormatting_(ws, `G3:G${lastRow}`);
  addRiskConditionalFormatting_(ws, `H3:H${lastRow}`);

  // Freeze header rows (with error handling for merged cells)
  try {
    ws.setFrozenRows(2);
    ws.setFrozenColumns(1);
  } catch (e) {
    // Silent skip - merged cells conflict with freeze
  }

  // Add column notes
  addDetailFindingNotes_(ws);

  Logger.log('✅ Detail Finding - Regular VAPT created');
}

// ═══════════════════════════════════════════════════════════════════════
// TAB 2: Evidence - Regular VAPT
// ═══════════════════════════════════════════════════════════════════════

function createEvidenceRegularVAPT(ss) {
  const ws = ss.getSheetByName('Evidence - Regular VAPT') ||
             ss.insertSheet('Evidence - Regular VAPT');

  ws.clear();
  ws.setTabColor(VAPT_COLORS.evidence);

  // Title row
  const title = 'VAPT EVIDENCE  .  Regular VAPT  .  Proof of Concept & Re-Test Evidence';
  ws.getRange(1, 1, 1, 26).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.evidence, VAPT_COLORS.white, 11)
    .setValue(title);
  ws.setRowHeight(1, 32);

  // Header row
  const headers = [
    'No',                               // A
    'App',                              // B
    'Scp',                              // C
    'Finding Name',                     // D
    'Proof of Concept (PoC) Description', // E
    'PoC Evidence 1',                   // F
    'PoC Evidence 2',                   // G
    'PoC Evidence 3',                   // H
    'PoC Evidence 4',                   // I
    'PoC Evidence 5',                   // J
    'PoC Evidence 6',                   // K
    'PoC Evidence 7',                   // L
    'PoC Evidence 8',                   // M
    'PoC Evidence 9',                   // N
    'PoC Evidence 10',                  // O
    'Re-VAPT Description',              // P
    'Re-VAPT Evidence 1',               // Q
    'Re-VAPT Evidence 2',               // R
    'Re-VAPT Evidence 3',               // S
    'Re-VAPT Evidence 4',               // T
    'Re-VAPT Evidence 5',               // U
    'Re-VAPT Evidence 6',               // V
    'Re-VAPT Evidence 7',               // W
    'Re-VAPT Evidence 8',               // X
    'Re-VAPT Evidence 9',               // Y
    'Re-VAPT Evidence 10',              // Z
  ];

  headers.forEach((header, i) => {
    vaptHeader_(ws.getRange(2, i + 1), VAPT_COLORS.evidence, VAPT_COLORS.white)
      .setValue(header);
  });
  ws.setRowHeight(2, 50);

  // Set column widths
  const widths = [
    120, 120, 80, 200, 300,  // A-E (No, App, Scp, Finding Name, PoC Desc)
    250, 250, 250, 250, 250, 250, 250, 250, 250, 250,  // F-O (PoC Evidence 1-10)
    300,  // P (Re-VAPT Description)
    250, 250, 250, 250, 250, 250, 250, 250, 250, 250   // Q-Z (Re-VAPT Evidence 1-10)
  ];
  widths.forEach((width, i) => ws.setColumnWidth(i + 1, width));

  // Freeze header rows (with error handling for merged cells)
  try {
    ws.setFrozenRows(2);
    ws.setFrozenColumns(4);  // Freeze first 4 columns (No, App, Scp, Finding Name)
  } catch (e) {
    // Silent skip - merged cells conflict with freeze
  }

  // Add column notes
  addEvidenceNotes_(ws);

  Logger.log('✅ Evidence - Regular VAPT created');
}

// ═══════════════════════════════════════════════════════════════════════
// TAB 3: Detail Finding - Ad Hoc VAPT
// ═══════════════════════════════════════════════════════════════════════

function createDetailFindingAdHocVAPT(ss) {
  const ws = ss.getSheetByName('Detail Finding - Ad Hoc VAPT') ||
             ss.insertSheet('Detail Finding - Ad Hoc VAPT');

  ws.clear();
  ws.setTabColor(VAPT_COLORS.adhoc);

  // Title row
  const title = 'VAPT DETAIL FINDING  .  Ad Hoc VAPT  .  QA PERURI';
  ws.getRange(1, 1, 1, 33).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.adhoc, VAPT_COLORS.white, 11)
    .setValue(title);
  ws.setRowHeight(1, 32);

  // Header row (same as Regular VAPT)
  const headers = [
    'Finding ID', 'Application', 'Category', 'Scope',
    'Status Fix (Dev)', 'Status Re-VAPT (Pentester)', 'Original Risk', 'Adjusted Risk',
    'Finding Name', 'Description', 'Impact', 'Affected URL/Endpoint',
    'Recommendation', 'Remediation Steps', 'Report Date', 'Pentester',
    'Target Remediation Date', 'Time to Remediate (Days)', 'Actual Fix Date',
    'Verified By (QA)', 'Verified Date', 'CVSS Score', 'CWE ID',
    'OWASP Top 10', 'PoC Available', 'Re-Test Required', 'Re-Test Date',
    'Re-Test Result', 'Notes / Comments', 'Jira Ticket', 'Acceptance Proof / MAoR',
    'Created By', 'Last Updated'
  ];

  headers.forEach((header, i) => {
    vaptHeader_(ws.getRange(2, i + 1), VAPT_COLORS.adhoc, VAPT_COLORS.white)
      .setValue(header);
  });
  ws.setRowHeight(2, 50);

  // Set column widths (same as Regular)
  const widths = [
    120, 120, 100, 80, 120, 120, 90, 90, 200,
    250, 180, 180, 220, 220, 100,
    100, 120, 100, 100, 100, 100,
    80, 80, 100, 80, 80, 100, 100,
    200, 120, 200, 100, 120
  ];
  widths.forEach((width, i) => ws.setColumnWidth(i + 1, width));

  // Data validation (same as Regular)
  const lastRow = 1000;
  const statusFixValues = ['Todo', 'On Progress Remediation', 'Ready to Retest', 'On Progress Retest', 'Done', 'Accepted', 'False Positive', 'Duplicated', 'Out of Scope'];
  const statusReVAPTValues = ['Open', 'Closed'];
  const riskValues = ['Informational', 'Low', 'Medium', 'High', 'Critical', 'False Positive'];

  ws.getRange(`E3:E${lastRow}`).setDataValidation(vaptDropdown_(statusFixValues));
  ws.getRange(`F3:F${lastRow}`).setDataValidation(vaptDropdown_(statusReVAPTValues));
  ws.getRange(`G3:G${lastRow}`).setDataValidation(vaptDropdown_(riskValues));
  ws.getRange(`H3:H${lastRow}`).setDataValidation(vaptDropdown_(riskValues));
  ws.getRange(`Y3:Y${lastRow}`).setDataValidation(vaptDropdown_(['Yes', 'No']));
  ws.getRange(`Z3:Z${lastRow}`).setDataValidation(vaptDropdown_(['Yes', 'No']));
  ws.getRange(`AB3:AB${lastRow}`).setDataValidation(vaptDropdown_(['Pass', 'Fail', 'Pending']));

  // Conditional formatting
  addVAPTStatusConditionalFormatting_(ws, `E3:E${lastRow}`);
  addVAPTReVAPTConditionalFormatting_(ws, `F3:F${lastRow}`);
  addRiskConditionalFormatting_(ws, `G3:G${lastRow}`);
  addRiskConditionalFormatting_(ws, `H3:H${lastRow}`);

  // Freeze (with error handling for merged cells)
  try {
    ws.setFrozenRows(2);
    ws.setFrozenColumns(1);
  } catch (e) {
    // Silent skip - merged cells conflict with freeze
  }

  // Add notes
  addDetailFindingNotes_(ws);

  Logger.log('✅ Detail Finding - Ad Hoc VAPT created');
}

// ═══════════════════════════════════════════════════════════════════════
// TAB 4: Evidence - Ad Hoc VAPT
// ═══════════════════════════════════════════════════════════════════════

function createEvidenceAdHocVAPT(ss) {
  const ws = ss.getSheetByName('Evidence - Ad Hoc VAPT') ||
             ss.insertSheet('Evidence - Ad Hoc VAPT');

  ws.clear();
  ws.setTabColor(VAPT_COLORS.adhoc);

  // Title row
  const title = 'VAPT EVIDENCE  .  Ad Hoc VAPT  .  Proof of Concept & Re-Test Evidence';
  ws.getRange(1, 1, 1, 26).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.adhoc, VAPT_COLORS.white, 11)
    .setValue(title);
  ws.setRowHeight(1, 32);

  // Header row (same as Regular VAPT Evidence)
  const headers = [
    'No', 'App', 'Scp', 'Finding Name', 'Proof of Concept (PoC) Description',
    'PoC Evidence 1', 'PoC Evidence 2', 'PoC Evidence 3', 'PoC Evidence 4', 'PoC Evidence 5',
    'PoC Evidence 6', 'PoC Evidence 7', 'PoC Evidence 8', 'PoC Evidence 9', 'PoC Evidence 10',
    'Re-VAPT Description',
    'Re-VAPT Evidence 1', 'Re-VAPT Evidence 2', 'Re-VAPT Evidence 3', 'Re-VAPT Evidence 4', 'Re-VAPT Evidence 5',
    'Re-VAPT Evidence 6', 'Re-VAPT Evidence 7', 'Re-VAPT Evidence 8', 'Re-VAPT Evidence 9', 'Re-VAPT Evidence 10'
  ];

  headers.forEach((header, i) => {
    vaptHeader_(ws.getRange(2, i + 1), VAPT_COLORS.adhoc, VAPT_COLORS.white)
      .setValue(header);
  });
  ws.setRowHeight(2, 50);

  // Set column widths (same as Regular Evidence)
  const widths = [
    120, 120, 80, 200, 300,
    250, 250, 250, 250, 250, 250, 250, 250, 250, 250,
    300,
    250, 250, 250, 250, 250, 250, 250, 250, 250, 250
  ];
  widths.forEach((width, i) => ws.setColumnWidth(i + 1, width));

  // Freeze (with error handling for merged cells)
  try {
    ws.setFrozenRows(2);
    ws.setFrozenColumns(4);
  } catch (e) {
    // Silent skip - merged cells conflict with freeze
  }

  // Add notes
  addEvidenceNotes_(ws);

  Logger.log('✅ Evidence - Ad Hoc VAPT created');
}

// ═══════════════════════════════════════════════════════════════════════
// CONDITIONAL FORMATTING HELPERS
// ═══════════════════════════════════════════════════════════════════════

function addVAPTStatusConditionalFormatting_(ws, rangeA1) {
  const range = ws.getRange(rangeA1);
  const rules = ws.getConditionalFormatRules();

  const statusFormats = [
    {value: 'Todo', bg: '#FFF3E0', fg: '#E65100'},
    {value: 'On Progress Remediation', bg: '#E3F2FD', fg: '#1565C0'},
    {value: 'Ready to Retest', bg: '#FFF9C4', fg: '#F57F17'},
    {value: 'On Progress Retest', bg: '#B3E5FC', fg: '#0277BD'},
    {value: 'Done', bg: '#C8E6C9', fg: '#2E7D32'},
    {value: 'Accepted', bg: '#A5D6A7', fg: '#1B5E20'},
    {value: 'False Positive', bg: '#E0E0E0', fg: '#424242'},
    {value: 'Duplicated', bg: '#F5F5F5', fg: '#757575'},
    {value: 'Out of Scope', bg: '#FFCCBC', fg: '#BF360C'}
  ];

  statusFormats.forEach(fmt => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(fmt.value)
        .setBackground(fmt.bg)
        .setFontColor(fmt.fg)
        .setBold(true)
        .setRanges([range])
        .build()
    );
  });

  ws.setConditionalFormatRules(rules);
}

function addVAPTReVAPTConditionalFormatting_(ws, rangeA1) {
  const range = ws.getRange(rangeA1);
  const rules = ws.getConditionalFormatRules();

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Open')
      .setBackground('#FFEBEE')
      .setFontColor('#C62828')
      .setBold(true)
      .setRanges([range])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Closed')
      .setBackground('#E8F5E9')
      .setFontColor('#2E7D32')
      .setBold(true)
      .setRanges([range])
      .build()
  );

  ws.setConditionalFormatRules(rules);
}

function addRiskConditionalFormatting_(ws, rangeA1) {
  const range = ws.getRange(rangeA1);
  const rules = ws.getConditionalFormatRules();

  const riskFormats = [
    {value: 'Critical', bg: '#FFEBEE', fg: '#B71C1C'},
    {value: 'High', bg: '#FFCDD2', fg: '#C62828'},
    {value: 'Medium', bg: '#FFF9C4', fg: '#F57F17'},
    {value: 'Low', bg: '#FFF8E1', fg: '#F9A825'},
    {value: 'Informational', bg: '#E3F2FD', fg: '#1565C0'},
    {value: 'False Positive', bg: '#F5F5F5', fg: '#757575'}
  ];

  riskFormats.forEach(fmt => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(fmt.value)
        .setBackground(fmt.bg)
        .setFontColor(fmt.fg)
        .setBold(true)
        .setRanges([range])
        .build()
    );
  });

  ws.setConditionalFormatRules(rules);
}

// ═══════════════════════════════════════════════════════════════════════
// COLUMN NOTES
// ═══════════════════════════════════════════════════════════════════════

function addDetailFindingNotes_(ws) {
  ws.getRange(2, 1).setNote(
    'Finding ID\n\n' +
    'Unique identifier untuk setiap finding VAPT.\n' +
    'Format: [PROJECT]-[TYPE]-[NUMBER]\n\n' +
    'Example: BGN-REG-001, SIPGN-ADHOC-045\n\n' +
    'TYPE: REG (Regular VAPT), ADHOC (Ad Hoc VAPT)'
  );

  ws.getRange(2, 5).setNote(
    'Status Fix (Dev/Pentester)\n\n' +
    'Status perbaikan oleh Development Team:\n\n' +
    '• Todo: Belum dikerjakan\n' +
    '• On Progress Remediation: Sedang diperbaiki\n' +
    '• Ready to Retest: Siap untuk di-test ulang\n' +
    '• On Progress Retest: Sedang di-test ulang\n' +
    '• Done: Selesai diperbaiki dan sudah di-verify\n' +
    '• Accepted: Diterima as-is (risk accepted)\n' +
    '• False Positive: Bukan vulnerability sebenarnya\n' +
    '• Duplicated: Duplikat dari finding lain\n' +
    '• Out of Scope: Di luar scope VAPT'
  );

  ws.getRange(2, 6).setNote(
    'Status Re-VAPT (Pentester)\n\n' +
    'Status dari Pentester setelah re-test:\n\n' +
    '• Open: Vulnerability masih ada / belum diperbaiki\n' +
    '• Closed: Vulnerability sudah berhasil diperbaiki dan verified'
  );

  ws.getRange(2, 7).setNote(
    'Original Risk\n\n' +
    'Risk level awal dari Pentester saat pertama kali ditemukan.\n\n' +
    'Levels: Critical, High, Medium, Low, Informational'
  );

  ws.getRange(2, 8).setNote(
    'Adjusted Risk\n\n' +
    'Risk level yang sudah disesuaikan setelah analisis lebih lanjut.\n\n' +
    'Bisa berbeda dari Original Risk karena:\n' +
    '• Mitigasi temporary sudah diterapkan\n' +
    '• Konteks bisnis / environment berbeda\n' +
    '• Impact analysis yang lebih detail\n\n' +
    'Levels: Critical, High, Medium, Low, Informational, False Positive'
  );

  ws.getRange(2, 18).setNote(
    'Time to Remediate (Days)\n\n' +
    'Jumlah hari yang dibutuhkan untuk memperbaiki vulnerability.\n\n' +
    'Auto-calculated dari:\n' +
    'Target Remediation Date - Report Date\n\n' +
    'atau\n\n' +
    'Actual Fix Date - Report Date (jika sudah selesai)'
  );
}

function addEvidenceNotes_(ws) {
  ws.getRange(2, 1).setNote(
    'No (Finding ID)\n\n' +
    'Finding ID yang sama dengan di tab Detail Finding.\n\n' +
    'Format: BGN-REG-001, SIPGN-ADHOC-045'
  );

  ws.getRange(2, 5).setNote(
    'Proof of Concept (PoC) Description\n\n' +
    'Penjelasan langkah-langkah exploit:\n\n' +
    '1. Pre-conditions / setup required\n' +
    '2. Step-by-step exploitation\n' +
    '3. Expected result / impact\n\n' +
    'Harus cukup detail agar Dev Team bisa reproduce.'
  );

  ws.getRange(2, 6).setNote(
    'PoC Evidence (Screenshots/Videos)\n\n' +
    'Upload screenshot atau video ke Google Drive.\n' +
    'Paste link di kolom ini.\n\n' +
    'Tips:\n' +
    '• Gunakan folder per-project di Drive\n' +
    '• Set permission: Anyone with link can view\n' +
    '• Beri nama file yang jelas (contoh: BGN-REG-001-PoC-1.png)'
  );

  ws.getRange(2, 16).setNote(
    'Re-VAPT Description\n\n' +
    'Deskripsi hasil re-test setelah vulnerability diperbaiki:\n\n' +
    '• Apa yang di-test ulang?\n' +
    '• Apakah masih bisa di-exploit?\n' +
    '• Verifikasi fix sudah benar?\n\n' +
    'Jika Closed: jelaskan kenapa vulnerability sudah tidak ada.\n' +
    'Jika Open: jelaskan kenapa masih vulnerable.'
  );

  ws.getRange(2, 17).setNote(
    'Re-VAPT Evidence (Screenshots/Videos)\n\n' +
    'Screenshot/video hasil re-test.\n\n' +
    'Upload ke Drive dan paste link di sini.'
  );
}

// ═══════════════════════════════════════════════════════════════════════
// MASTER FUNCTION - CREATE ALL 4 VAPT TABS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Create all 4 VAPT tabs in current spreadsheet
 * Call this from MasterQATCM.js createQASheet() or run standalone
 */
function createAllVAPTTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    Logger.log('🔒 Creating VAPT tabs...');

    createDetailFindingRegularVAPT(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    createEvidenceRegularVAPT(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    createDetailFindingAdHocVAPT(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    createEvidenceAdHocVAPT(ss);
    SpreadsheetApp.flush();

    ui.alert(
      '✅ VAPT Tabs Created!',
      'Successfully created 4 VAPT tabs:\n\n' +
      '1. Detail Finding - Regular VAPT\n' +
      '2. Evidence - Regular VAPT\n' +
      '3. Detail Finding - Ad Hoc VAPT\n' +
      '4. Evidence - Ad Hoc VAPT\n\n' +
      'All tabs include:\n' +
      '• Data validation dropdowns\n' +
      '• Conditional formatting\n' +
      '• Column notes for guidance',
      ui.ButtonSet.OK
    );

    Logger.log('✅ All VAPT tabs created successfully');

  } catch (e) {
    Logger.log('❌ Error creating VAPT tabs: ' + e.message);
    ui.alert(
      '❌ Error',
      'Failed to create VAPT tabs:\n\n' + e.message + '\n\n' +
      'Check Execution log for details.',
      ui.ButtonSet.OK
    );
  }
}

// ═══════════════════════════════════════════════════════════════════════
// SIMPLIFIED VAPT TABS (2 TABS ONLY)
// ═══════════════════════════════════════════════════════════════════════

/**
 * Create Detail Finding - VAPT tab (simplified - no Regular/Ad Hoc split)
 * Single unified VAPT tab for all findings
 */
function createDetailFindingVAPT(ss) {
  const ws = ss.getSheetByName('Detail Finding - VAPT') ||
             ss.insertSheet('Detail Finding - VAPT');

  ws.clear();
  ws.setTabColor('#FF6F00');  // Deep Orange

  // Title row
  const title = 'VAPT DETAIL FINDING  .  QA PERURI';
  ws.getRange(1, 1, 1, 33).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.header, VAPT_COLORS.white, 11)
    .setValue(title);
  ws.setRowHeight(1, 32);

  // Header row
  const headers = [
    'Finding ID',                    // A
    'Application',                   // B
    'Category',                      // C
    'Scope',                         // D
    'Status Fix (Dev)',              // E
    'Status Re-VAPT (Pentester)',    // F
    'Original Risk',                 // G
    'Adjusted Risk',                 // H
    'Finding Name',                  // I
    'Description',                   // J
    'Impact',                        // K
    'Affected URL/Endpoint',         // L
    'Recommendation',                // M
    'Remediation Steps',             // N
    'Report Date',                   // O
    'Pentester',                     // P
    'Target Remediation Date',       // Q
    'Time to Remediate (Days)',      // R
    'Actual Fix Date',               // S
    'Verified By (QA)',              // T
    'Verified Date',                 // U
    'CVSS Score',                    // V
    'CWE ID',                        // W
    'OWASP Top 10',                  // X
    'PoC Available',                 // Y
    'Re-Test Required',              // Z
    'Re-Test Date',                  // AA
    'Re-Test Result',                // AB
    'Notes / Comments',              // AC
    'Jira Ticket',                   // AD
    'Acceptance Proof / MAoR',       // AE
    'Created By',                    // AF
    'Last Updated',                  // AG
  ];

  headers.forEach((header, i) => {
    vaptHeader_(ws.getRange(2, i + 1), '#FF6F00', VAPT_COLORS.white)
      .setValue(header);
  });
  ws.setRowHeight(2, 50);

  // Set column widths
  const widths = [
    120, 120, 100, 80, 120, 120, 90, 90, 200,   // A-I
    250, 180, 180, 220, 220, 100,                // J-O
    100, 120, 100, 100, 100, 100,                // P-U
    80, 80, 100, 80, 80, 100, 100,               // V-AB
    200, 120, 200, 100, 120                      // AC-AG
  ];
  widths.forEach((width, i) => ws.setColumnWidth(i + 1, width));

  // Data validation dropdowns
  const lastRow = 1000;

  // E: Status Fix (Dev)
  const statusFixValues = [
    'Todo',
    'On Progress Remediation',
    'Ready to Retest',
    'On Progress Retest',
    'Done',
    'Accepted',
    'False Positive',
    'Duplicated',
    'Out of Scope'
  ];
  ws.getRange(`E3:E${lastRow}`).setDataValidation(vaptDropdown_(statusFixValues));

  // F: Status Re-VAPT (Pentester)
  const statusReVAPTValues = ['Open', 'Closed'];
  ws.getRange(`F3:F${lastRow}`).setDataValidation(vaptDropdown_(statusReVAPTValues));

  // G & H: Risk levels
  const riskValues = ['Informational', 'Low', 'Medium', 'High', 'Critical', 'False Positive'];
  ws.getRange(`G3:G${lastRow}`).setDataValidation(vaptDropdown_(riskValues));
  ws.getRange(`H3:H${lastRow}`).setDataValidation(vaptDropdown_(riskValues));

  // Y: PoC Available
  ws.getRange(`Y3:Y${lastRow}`).setDataValidation(vaptDropdown_(['Yes', 'No']));

  // Z: Re-Test Required
  ws.getRange(`Z3:Z${lastRow}`).setDataValidation(vaptDropdown_(['Yes', 'No']));

  // AB: Re-Test Result
  ws.getRange(`AB3:AB${lastRow}`).setDataValidation(vaptDropdown_(['Pass', 'Fail', 'Pending']));

  // Conditional formatting for Status Fix
  addVAPTStatusConditionalFormatting_(ws, `E3:E${lastRow}`);

  // Conditional formatting for Status Re-VAPT
  addVAPTReVAPTConditionalFormatting_(ws, `F3:F${lastRow}`);

  // Conditional formatting for Risk levels
  addRiskConditionalFormatting_(ws, `G3:G${lastRow}`);
  addRiskConditionalFormatting_(ws, `H3:H${lastRow}`);

  // Freeze header rows (with error handling for merged cells)
  try {
    ws.setFrozenRows(2);
    ws.setFrozenColumns(1);
  } catch (e) {
    // Silent skip - merged cells conflict with freeze
  }

  // Add column notes
  addDetailFindingNotes_(ws);

  Logger.log('✅ Detail Finding - VAPT created');
}

/**
 * Create Evidence - VAPT tab (simplified - no Regular/Ad Hoc split)
 * Single unified Evidence tab for all VAPT findings
 */
function createEvidenceVAPT(ss) {
  const ws = ss.getSheetByName('Evidence - VAPT') ||
             ss.insertSheet('Evidence - VAPT');

  ws.clear();
  ws.setTabColor('#004D40');  // Teal

  // Title row
  const title = 'VAPT EVIDENCE  .  QA PERURI';
  ws.getRange(1, 1, 1, 26).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.header, VAPT_COLORS.white, 11)
    .setValue(title);
  ws.setRowHeight(1, 32);

  // Header row
  const headers = [
    'Finding ID',          // A
    'Application',         // B
    'Finding Name',        // C
    'Adjusted Risk',       // D
    'PoC Evidence 1',      // E
    'PoC Evidence 2',      // F
    'PoC Evidence 3',      // G
    'PoC Evidence 4',      // H
    'PoC Evidence 5',      // I
    'PoC Evidence 6',      // J
    'PoC Evidence 7',      // K
    'PoC Evidence 8',      // L
    'PoC Evidence 9',      // M
    'PoC Evidence 10',     // N
    'Re-VAPT Evidence 1',  // O
    'Re-VAPT Evidence 2',  // P
    'Re-VAPT Evidence 3',  // Q
    'Re-VAPT Evidence 4',  // R
    'Re-VAPT Evidence 5',  // S
    'Re-VAPT Evidence 6',  // T
    'Re-VAPT Evidence 7',  // U
    'Re-VAPT Evidence 8',  // V
    'Re-VAPT Evidence 9',  // W
    'Re-VAPT Evidence 10', // X
    'Created By',          // Y
    'Last Updated'         // Z
  ];

  headers.forEach((header, i) => {
    vaptHeader_(ws.getRange(2, i + 1), '#004D40', VAPT_COLORS.white)
      .setValue(header);
  });
  ws.setRowHeight(2, 50);

  // Set column widths
  const widths = [
    120, 120, 200, 90,                          // A-D
    150, 150, 150, 150, 150, 150, 150, 150, 150, 150,  // E-N (PoC)
    150, 150, 150, 150, 150, 150, 150, 150, 150, 150,  // O-X (Re-VAPT)
    100, 120                                    // Y-Z
  ];
  widths.forEach((width, i) => ws.setColumnWidth(i + 1, width));

  // Data validation for Risk level
  const lastRow = 1000;
  const riskValues = ['Informational', 'Low', 'Medium', 'High', 'Critical', 'False Positive'];
  ws.getRange(`D3:D${lastRow}`).setDataValidation(vaptDropdown_(riskValues));

  // Conditional formatting for Risk level
  addRiskConditionalFormatting_(ws, `D3:D${lastRow}`);

  // Freeze header rows (with error handling for merged cells)
  try {
    ws.setFrozenRows(2);
    ws.setFrozenColumns(3);
  } catch (e) {
    // Silent skip - merged cells conflict with freeze
  }

  // Add column notes
  addEvidenceNotes_(ws);

  Logger.log('✅ Evidence - VAPT created');
}
