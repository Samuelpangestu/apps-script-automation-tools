/**
 * InitVAPTTabs.js - Initialize 3 VAPT tabs in QATM template
 *
 * NEW STRUCTURE (Simplified):
 * 1. VAPT - Helper (Dashboard/Tracking)
 * 2. VAPT - Detail Finding (Combined Regular + Ad Hoc)
 * 3. VAPT - Evidence (Combined Regular + Ad Hoc)
 *
 * Usage:
 * - For new templates: Called from createQASheet()
 * - For existing: Use broadcastVAPTTabs() in BroadcastVAPTTabs.js
 */

// ═══════════════════════════════════════════════════════════════════════
// VAPT TAB COLORS
// ═══════════════════════════════════════════════════════════════════════

const VAPT_COLORS = {
  helper: '#1A237E',      // Dark Blue for Helper
  detail: '#263238',      // Dark Grey for Detail Finding
  evidence: '#004D40',    // Dark Teal for Evidence
  white: '#FFFFFF',
  orange: '#FF6F00',      // Orange accent
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
// TAB 1: VAPT - Helper (Dashboard/Tracking)
// ═══════════════════════════════════════════════════════════════════════

function createVAPTHelper(ss) {
  const ws = ss.getSheetByName('VAPT - Helper') || ss.insertSheet('VAPT - Helper');

  ws.clear();
  ws.setTabColor(VAPT_COLORS.helper);

  // Title row
  const title = 'VAPT HELPER  .  QA PERURI';
  ws.getRange(1, 1, 1, 32).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.helper, VAPT_COLORS.white, 12)
    .setValue(title);
  ws.setRowHeight(1, 35);

  // Section 1: Vulnerabilities in Production (Row 3)
  ws.getRange(3, 1, 1, 32).merge();
  ws.getRange(3, 1)
    .setValue('Vulnerabilities in Production')
    .setBackground('#E3F2FD')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  // Headers for metrics (Row 4)
  const metricsHeaders = [
    '', '', 'Ready to Retest', '', '', '', '', 'Open', '', '', '', '', 'Closed', '', '', '', '',
    'Accepted - Open', '', '', '', '', 'Accepted - Closed', '', '', '', '', '', '', '', '', ''
  ];

  ws.getRange(4, 1, 1, 32).setValues([metricsHeaders])
    .setBackground(VAPT_COLORS.detail)
    .setFontColor(VAPT_COLORS.white)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Risk level headers (Row 5)
  const riskHeaders = ['Critical', 'High', 'Medium', 'Low', 'Informational'];
  for (let i = 0; i < 6; i++) {
    const startCol = 3 + (i * 5);
    riskHeaders.forEach((risk, idx) => {
      ws.getRange(5, startCol + idx).setValue(risk);
    });
  }
  ws.getRange(5, 1, 1, 32)
    .setBackground('#B0BEC5')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Data row with formulas (Row 6) - Count vulnerabilities in production
  const sheetName = "'VAPT - Detail Finding'";
  const dataRange = sheetName + "!A3:AF";

  // Risk levels for columns
  const risks = ['Critical', 'High', 'Medium', 'Low', 'Informational'];

  // Helper function to create COUNTIFS formula
  const countFormula = (statusCol, statusVal, riskLevel, extraCondition = '') => {
    let formula = `=COUNTIFS(${sheetName}!AA:AA,"Yes",${sheetName}!${statusCol}:${statusCol},"${statusVal}",${sheetName}!H:H,"${riskLevel}"`;
    if (extraCondition) formula += extraCondition;
    formula += ')';
    return formula;
  };

  const formulas = [
    '', // Col A
    '', // Col B
    // Ready to Retest (Cols C-G)
    ...risks.map(r => countFormula('E', 'Ready to Retest', r)),
    // Open (Cols H-L)
    ...risks.map(r => countFormula('F', 'Open', r)),
    // Closed (Cols M-Q)
    ...risks.map(r => countFormula('F', 'Closed', r)),
    // Accepted - Open (Cols R-V)
    ...risks.map(r => countFormula('E', 'Accepted', r, `,${sheetName}!F:F,"Open"`)),
    // Accepted - Closed (Cols W-AA)
    ...risks.map(r => countFormula('E', 'Accepted', r, `,${sheetName}!F:F,"Closed"`)),
    // Reserved columns (Cols AB-AF)
    '', '', '', '', ''
  ];

  ws.getRange(6, 1, 1, 32).setFormulas([formulas])
    .setNumberFormat('0')
    .setHorizontalAlignment('center');

  ws.setRowHeight(4, 25);
  ws.setRowHeight(5, 25);

  // Section 2: Tracking Vulnerability in Production (Row 10)
  ws.getRange(10, 1, 1, 11).merge();
  ws.getRange(10, 1)
    .setValue('Tracking Vulnerability in Production')
    .setBackground('#FFF3E0')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  const trackProdHeaders = [
    'VAPT Type', 'Application', 'Scp', 'Status Fix\n(Filled by Dev Team)', 'Adjusted Risk',
    'Finding ID', 'Finding Name', 'Report Date', 'Target Remediation Date',
    'Time to Remediate (Days)', 'Acceptance Proof / MAoR'
  ];

  ws.getRange(11, 1, 1, 11).setValues([trackProdHeaders]);
  vaptHeader_(ws.getRange(11, 1, 1, 11), VAPT_COLORS.orange, VAPT_COLORS.white);
  ws.setRowHeight(11, 40);

  // Add FILTER formula for Tracking in Production (Row 12, Cols A-K)
  const filterProdFormula =
    `=ARRAYFORMULA(IF(ISBLANK(FILTER(${sheetName}!AA:AA,${sheetName}!AA:AA="Yes")),"",` +
    `{` +
    `"Regular VAPT",` + // VAPT Type (placeholder, could be enhanced)
    `FILTER(${sheetName}!B:B,${sheetName}!AA:AA="Yes"),` + // App
    `FILTER(${sheetName}!D:D,${sheetName}!AA:AA="Yes"),` + // Scp
    `FILTER(${sheetName}!E:E,${sheetName}!AA:AA="Yes"),` + // Status Fix
    `FILTER(${sheetName}!H:H,${sheetName}!AA:AA="Yes"),` + // Adjusted Risk
    `FILTER(${sheetName}!A:A,${sheetName}!AA:AA="Yes"),` + // Finding ID
    `FILTER(${sheetName}!I:I,${sheetName}!AA:AA="Yes"),` + // Finding Name
    `FILTER(${sheetName}!O:O,${sheetName}!AA:AA="Yes"),` + // Report Date
    `FILTER(${sheetName}!Q:Q,${sheetName}!AA:AA="Yes"),` + // Target Remediation Date
    `FILTER(${sheetName}!R:R,${sheetName}!AA:AA="Yes"),` + // Time to Remediate
    `FILTER(${sheetName}!AE:AE,${sheetName}!AA:AA="Yes")` + // Acceptance Proof
    `}))`;

  ws.getRange(12, 1).setFormula(filterProdFormula);

  // Section 3: Tracking Vulnerability Overall (Row 10, starting Col M)
  ws.getRange(10, 13, 1, 12).merge();
  ws.getRange(10, 13)
    .setValue('Tracking Vulnerability Overall')
    .setBackground('#E8F5E9')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  const trackOverallHeaders = [
    'VAPT Type', 'Application', 'Scp', 'Status Fix\n(Filled by Dev Team)', 'Risk', 'Adjusted Risk',
    'Finding ID', 'Finding Name', 'Report Date', 'Target Remediation Date',
    'Time to Remediate (Days)', 'Acceptance Proof / MAoR'
  ];

  ws.getRange(11, 13, 1, 12).setValues([trackOverallHeaders]);
  vaptHeader_(ws.getRange(11, 13, 1, 12), '#2E7D32', VAPT_COLORS.white);

  // Add FILTER formula for Tracking Overall (Row 12, Cols M-X)
  const filterOverallFormula =
    `=ARRAYFORMULA(IF(ISBLANK(${sheetName}!A3:A),"",` +
    `{` +
    `"Regular VAPT",` + // VAPT Type (placeholder)
    `${sheetName}!B3:B,` + // App
    `${sheetName}!D3:D,` + // Scp
    `${sheetName}!E3:E,` + // Status Fix
    `${sheetName}!G3:G,` + // Risk
    `${sheetName}!H3:H,` + // Adjusted Risk
    `${sheetName}!A3:A,` + // Finding ID
    `${sheetName}!I3:I,` + // Finding Name
    `${sheetName}!O3:O,` + // Report Date
    `${sheetName}!Q3:Q,` + // Target Remediation Date
    `${sheetName}!R3:R,` + // Time to Remediate
    `${sheetName}!AE3:AE` + // Acceptance Proof
    `}))`;

  ws.getRange(12, 13).setFormula(filterOverallFormula);

  // Set column widths
  const widths = [100, 120, 80, 150, 100, 120, 200, 100, 120, 100, 180];
  widths.forEach((width, i) => {
    ws.setColumnWidth(i + 1, width);
    if (i < 11) ws.setColumnWidth(i + 13, width); // For second table
  });

  // Freeze rows
  try {
    ws.setFrozenRows(11);
  } catch (e) {
    // Silent skip
  }

  Logger.log('✅ VAPT - Helper created');
}

// ═══════════════════════════════════════════════════════════════════════
// TAB 2: VAPT - Detail Finding
// ═══════════════════════════════════════════════════════════════════════

function createDetailFindingVAPT(ss) {
  const ws = ss.getSheetByName('VAPT - Detail Finding') ||
             ss.insertSheet('VAPT - Detail Finding');

  ws.clear();
  ws.setTabColor(VAPT_COLORS.detail);

  // Title row
  const title = 'VAPT DETAIL FINDING  .  QA PERURI';
  ws.getRange(1, 1, 1, 32).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.detail, VAPT_COLORS.white, 11)
    .setValue(title);
  ws.setRowHeight(1, 32);

  // Header row (32 columns)
  const headers = [
    'No',                                    // A
    'App',                                   // B
    'Ver',                                   // C
    'Scp',                                   // D
    'Status Fix (Dev / Pentester)',          // E
    'Status Re-VAPT (Pentester)',            // F
    'Risk',                                  // G
    'Adjusted Risk',                         // H
    'Finding Name',                          // I
    'Report Document',                       // J
    'Description',                           // K
    'Impact',                                // L
    'Recommendation',                        // M
    'Affected Target',                       // N
    'Report Date',                           // O
    'Finding Closed Date',                   // P
    'Target Remediation Date (Automatic)',   // Q
    'Time to Remediate (Automatic)',         // R
    'Security Tester',                       // S
    'Perban BSSN',                           // T
    'OWASP\n(All Platform)',                 // U
    'OWASP\n(Specific Platform)',            // V
    'Tested By',                             // W
    'Reporter',                              // X
    'Envi',                                  // Y
    'System PIC',                            // Z
    'Already in Prod',                       // AA
    'Go to Prod Date',                       // AB
    'CVSS 4.0',                              // AC
    'Ticket ID / Request ID',                // AD
    'Acceptance Proof / MAoR',               // AE
    'Notes'                                  // AF
  ];

  headers.forEach((header, i) => {
    vaptHeader_(ws.getRange(2, i + 1), VAPT_COLORS.detail, VAPT_COLORS.white)
      .setValue(header);
  });
  ws.setRowHeight(2, 50);

  // Set column widths
  const widths = [
    120, 120, 80, 80, 150, 150, 90, 90, 200,      // A-I
    180, 250, 180, 220, 180, 100,                  // J-O
    100, 120, 100, 120, 180,                       // P-T
    180, 200, 100, 100, 100, 120, 100, 100,        // U-AB
    100, 150, 200, 200                             // AC-AF
  ];
  widths.forEach((width, i) => ws.setColumnWidth(i + 1, width));

  // Data validation dropdowns
  const lastRow = 1000;

  // E: Status Fix (Dev / Pentester)
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
  ws.getRange('E3:E' + lastRow).setDataValidation(vaptDropdown_(statusFixValues));

  // F: Status Re-VAPT (Pentester)
  const statusReVAPTValues = ['Open', 'Closed'];
  ws.getRange('F3:F' + lastRow).setDataValidation(vaptDropdown_(statusReVAPTValues));

  // G & H: Risk levels
  const riskValues = ['Informational', 'Low', 'Medium', 'High', 'Critical', 'False Positive'];
  ws.getRange('G3:G' + lastRow).setDataValidation(vaptDropdown_(riskValues));
  ws.getRange('H3:H' + lastRow).setDataValidation(vaptDropdown_(riskValues));

  // S: Security Tester
  const securityTesterValues = ['Peruri', 'BSSN', 'MII', 'Telkomsigma'];
  ws.getRange('S3:S' + lastRow).setDataValidation(vaptDropdown_(securityTesterValues));

  // T: Perban BSSN
  const perbanBSSNValues = [
    'Web - Autentikasi',
    'Web - Manajemen Sesi',
    'Web - Persyaratan Kontrol Akses',
    'Web - Validasi Input',
    'Web - Kriptografi pada Verifikasi Statis',
    'Web - Penanganan Eror dan Pencatatan log',
    'Web - Proteksi Data',
    'Web - Keamanan Komunikasi',
    'Web - Pengendalian Kode Berbahaya',
    'Web - Logika Bisnis',
    'Web - File',
    'Web - Keamanan API dan Web Service',
    'Web - Keamanan Konfigurasi',
    'Mobile - Penyimpanan Data dan persyaratan privasi',
    'Mobile - Kriptografi',
    'Mobile - Autentikasi dan Manajemen Sesi',
    'Mobile - Komunikasi Jaringan',
    'Mobile - Interaksi Platform',
    'Mobile - Kualitas Kode dan Pengaturan Build',
    'Mobile - Ketahanan',
    'N/A'
  ];
  ws.getRange('T3:T' + lastRow).setDataValidation(vaptDropdown_(perbanBSSNValues));

  // U: OWASP (All Platform)
  const owaspAllValues = [
    'A01:2025 - Broken Access Control',
    'A02:2025 - Security Misconfiguration',
    'A03:2025 - Software Supply Chain Failures',
    'A04:2025 - Cryptographic Failures',
    'A05:2025 - Injection',
    'A06:2025 - Insecure Design',
    'A07:2025 - Authentication Failures',
    'A08:2025 - Software or Data Integrity Failures',
    'A09:2025 - Security Logging & Alerting Failures',
    'A10:2025 - Mishandling of Exceptional Conditions'
  ];
  ws.getRange('U3:U' + lastRow).setDataValidation(vaptDropdown_(owaspAllValues));

  // V: OWASP (Specific Platform) - Very long list
  const owaspSpecificValues = [
    'N/A',
    'A01:2025 - Broken Access Control',
    'A02:2025 - Security Misconfiguration',
    'A03:2025 - Software Supply Chain Failures',
    'A04:2025 - Cryptographic Failures',
    'A05:2025 - Injection',
    'A06:2025 - Insecure Design',
    'A07:2025 - Authentication Failures',
    'A08:2025 - Software or Data Integrity Failures',
    'A09:2025 - Security Logging and Alerting Failures',
    'A10:2025 - Mishandling of Exceptional Conditions',
    'M1: 2024 - Improper Credential Usage',
    'M2: 2024 - Inadequate Supply Chain Security',
    'M3: 2024 - Insecure Authentication/Authorization',
    'M4: 2024 - Insufficient Input/Output Validation',
    'M5: 2024 - Insecure Communication',
    'M6: 2024 - Inadequate Privacy Controls',
    'M7: 2024 - Insufficient Binary Protections',
    'M8: 2024 - Security Misconfiguration',
    'M9: 2024 - Insecure Data Storage',
    'M10: 2024 - Insufficient Cryptography',
    'API1: 2023 - Broken Object Level Authorization',
    'API2: 2023 - Broken Authentication',
    'API3: 2023 - Broken Object Property Level Authorization',
    'API4: 2023 - Unrestricted Resource Consumption',
    'API5: 2023 - Broken Function Level Authorization',
    'API6: 2023 - Unrestricted Access to Sensitive Business Flows',
    'API7: 2023 - Server Side Request Forgery',
    'API8: 2023 - Security Misconfiguration',
    'API9: 2023 - Improper Inventory Management',
    'API10: 2023 - Unsafe Consumption of APIs',
    'LLM01: 2025 - Prompt Injection',
    'LLM02: 2025 - Insecure Output Handling',
    'LLM03: 2025 - Training Data Poisoning',
    'LLM04: 2025 - Model Denial of Service',
    'LLM05: 2025 - Supply Chain Vulnerabilities',
    'LLM06: 2025 - Sensitive Information Disclosure',
    'LLM07: 2025 - Insecure Plugin Design',
    'LLM08: 2025 - Excessive Agency',
    'LLM09: 2025 - Overreliance',
    'LLM10: 0225 - Model Theft',
    'ISR01:2024 – Outdated Software',
    'ISR02:2024 – Insufficient Threat Detection',
    'ISR03:2024 – Insecure Configurations',
    'ISR04:2024 – Insecure Resource and User Management',
    'ISR05:2024 – Insecure Use of Cryptography',
    'ISR06:2024 – Insecure Network Access Management',
    'ISR07:2024 – Insecure Authentication Methods and Default Credentials',
    'ISR08:2024 – Information Leakage',
    'ISR09:2024 – Insecure Access to Resources and Management Components',
    'ISR10:2024 – Insufficient Asset Management and Documentation'
  ];
  ws.getRange('V3:V' + lastRow).setDataValidation(vaptDropdown_(owaspSpecificValues));

  // Y: Envi
  const enviValues = ['Local', 'Development', 'Staging', 'Production'];
  ws.getRange('Y3:Y' + lastRow).setDataValidation(vaptDropdown_(enviValues));

  // AA: Already in Prod
  ws.getRange('AA3:AA' + lastRow).setDataValidation(vaptDropdown_(['Yes', 'No']));

  // Conditional formatting for Status Fix
  addVAPTStatusConditionalFormatting_(ws, 'E3:E' + lastRow);

  // Conditional formatting for Status Re-VAPT
  addVAPTReVAPTConditionalFormatting_(ws, 'F3:F' + lastRow);

  // Conditional formatting for Risk levels
  addRiskConditionalFormatting_(ws, 'G3:G' + lastRow);
  addRiskConditionalFormatting_(ws, 'H3:H' + lastRow);

  // Freeze header rows
  try {
    ws.setFrozenRows(2);
    ws.setFrozenColumns(1);
  } catch (e) {
    // Silent skip
  }

  // Add column notes
  addDetailFindingNotes_(ws);

  Logger.log('✅ VAPT - Detail Finding created');
}

// ═══════════════════════════════════════════════════════════════════════
// TAB 3: VAPT - Evidence
// ═══════════════════════════════════════════════════════════════════════

function createEvidenceVAPT(ss) {
  const ws = ss.getSheetByName('VAPT - Evidence') ||
             ss.insertSheet('VAPT - Evidence');

  ws.clear();
  ws.setTabColor(VAPT_COLORS.evidence);

  // Title row
  const title = 'VAPT EVIDENCE  .  QA PERURI';
  ws.getRange(1, 1, 1, 26).merge();
  vaptHeader_(ws.getRange(1, 1), VAPT_COLORS.evidence, VAPT_COLORS.white, 11)
    .setValue(title);
  ws.setRowHeight(1, 32);

  // Header row
  const headers = [
    'No',                                    // A
    'App',                                   // B
    'Scp',                                   // C
    'Finding Name',                          // D
    'Proof of Concept (PoC) Description',    // E
    'PoC Evidence 1',                        // F
    'PoC Evidence 2',                        // G
    'PoC Evidence 3',                        // H
    'PoC Evidence 4',                        // I
    'PoC Evidence 5',                        // J
    'PoC Evidence 6',                        // K
    'PoC Evidence 7',                        // L
    'PoC Evidence 8',                        // M
    'PoC Evidence 9',                        // N
    'PoC Evidence 10',                       // O
    'Re-VAPT Description',                   // P
    'Re-VAPT Evidence 1',                    // Q
    'Re-VAPT Evidence 2',                    // R
    'Re-VAPT Evidence 3',                    // S
    'Re-VAPT Evidence 4',                    // T
    'Re-VAPT Evidence 5',                    // U
    'Re-VAPT Evidence 6',                    // V
    'Re-VAPT Evidence 7',                    // W
    'Re-VAPT Evidence 8',                    // X
    'Re-VAPT Evidence 9',                    // Y
    'Re-VAPT Evidence 10',                   // Z
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

  // Freeze header rows
  try {
    ws.setFrozenRows(2);
    ws.setFrozenColumns(4);  // Freeze first 4 columns (No, App, Scp, Finding Name)
  } catch (e) {
    // Silent skip
  }

  // Add column notes
  addEvidenceNotes_(ws);

  Logger.log('✅ VAPT - Evidence created');
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
    'No (Finding ID)\n\n' +
    'Unique identifier untuk setiap finding VAPT.\n' +
    'Format: [PROJECT]-[NUMBER]\n\n' +
    'Example: BGN-001, SIPGN-045'
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
    'Risk\n\n' +
    'Risk level dari Pentester.\n\n' +
    'Levels: Critical, High, Medium, Low, Informational, False Positive'
  );

  ws.getRange(2, 8).setNote(
    'Adjusted Risk\n\n' +
    'Risk level yang sudah disesuaikan setelah analisis lebih lanjut.\n\n' +
    'Bisa berbeda dari Risk karena:\n' +
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
    'Finding Closed Date - Report Date (jika sudah selesai)'
  );

  ws.getRange(2, 25).setNote(
    'Envi (Environment)\n\n' +
    'Environment dimana vulnerability ditemukan:\n\n' +
    '• Local: Development local\n' +
    '• Development: Dev server\n' +
    '• Staging: Staging server\n' +
    '• Production: Production server'
  );
}

function addEvidenceNotes_(ws) {
  ws.getRange(2, 1).setNote(
    'No (Finding ID)\n\n' +
    'Finding ID yang sama dengan di tab VAPT - Detail Finding.\n\n' +
    'Format: BGN-001, SIPGN-045'
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
    '• Beri nama file yang jelas (contoh: BGN-001-PoC-1.png)'
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
// MASTER FUNCTION - CREATE ALL 3 VAPT TABS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Create all 3 VAPT tabs in current spreadsheet
 * Call this from MasterQATCM.js createQASheet() or run standalone
 */
function createAllVAPTTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    Logger.log('🔒 Creating VAPT tabs...');

    createVAPTHelper(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    createDetailFindingVAPT(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    createEvidenceVAPT(ss);
    SpreadsheetApp.flush();

    ui.alert(
      '✅ VAPT Tabs Created!',
      'Successfully created 3 VAPT tabs:\n\n' +
      '1. VAPT - Helper (Dashboard/Tracking)\n' +
      '2. VAPT - Detail Finding\n' +
      '3. VAPT - Evidence\n\n' +
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
