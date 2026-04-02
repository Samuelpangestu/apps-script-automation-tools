/**
 * FixTCNotes.js
 *
 * Fix TC_Master column notes yang tergeser
 * Broadcast ke semua QATM spreadsheets yang terdaftar di Config
 */

/**
 * Broadcast TC_Master notes fix ke semua QATM spreadsheets
 * Reads from Config tab untuk dapat list QATM spreadsheet IDs
 */
function broadcastFixTCNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '🔧 Fix TC_Master Column Notes',
    'Fix ini akan memperbaiki column notes di TC_Master yang tergeser:\n\n' +
    '• Col 1 (No) → TEST CASE ID notes\n' +
    '• Col 2 (SubModul) → MODUL notes\n' +
    '• Col 3 (TC_ID) → TEST SCENARIO notes\n' +
    '• All 14 columns repositioned correctly\n\n' +
    'Headers TIDAK diubah (hanya notes).\n\n' +
    'Akan di-broadcast ke semua QATM spreadsheets di Config.\n\n' +
    'Lanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Fix dibatalkan.');
    return;
  }

  try {
    Logger.log('🔧 Starting TC_Master notes broadcast fix...');

    // Get QATM spreadsheet IDs from Config
    const spreadsheetIds = getQATMSpreadsheetIds_();

    if (spreadsheetIds.length === 0) {
      ui.alert(
        'No QATM Spreadsheets Found',
        'Tidak ada QATM spreadsheet IDs yang ditemukan di Config tab.\n\n' +
        'Pastikan Config tab sudah disetup dengan benar.',
        ui.ButtonSet.OK
      );
      return;
    }

    Logger.log(`Found ${spreadsheetIds.length} QATM spreadsheet(s) to fix`);

    let successCount = 0;
    let failCount = 0;
    const errors = [];

    spreadsheetIds.forEach((id, index) => {
      try {
        Logger.log(`Processing spreadsheet ${index + 1}/${spreadsheetIds.length}: ${id}`);

        const qatmSS = SpreadsheetApp.openById(id);
        const ws = qatmSS.getSheetByName('TC_Master');

        if (!ws) {
          errors.push(`${qatmSS.getName()}: TC_Master tab not found`);
          failCount++;
          return;
        }

        // Apply fixes
        applyTCNotesFix_(ws);

        Logger.log(`✅ Fixed: ${qatmSS.getName()}`);
        successCount++;

      } catch (e) {
        Logger.log(`❌ Error on ${id}: ${e.message}`);
        errors.push(`${id}: ${e.message}`);
        failCount++;
      }
    });

    // Show results
    let message = `Broadcast Complete!\n\n`;
    message += `✅ Success: ${successCount} spreadsheet(s)\n`;
    message += `❌ Failed: ${failCount} spreadsheet(s)`;

    if (errors.length > 0) {
      message += `\n\nErrors:\n${errors.join('\n')}`;
    }

    ui.alert('✅ Broadcast Results', message, ui.ButtonSet.OK);
    Logger.log(message);

  } catch (e) {
    Logger.log('❌ Error in broadcast fix: ' + e.message);
    ui.alert(
      '❌ Error',
      'Gagal apply broadcast fix:\n' + e.message + '\n\n' +
      'Check Executions log untuk detail.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Get QATM Spreadsheet IDs from Config tab
 * @returns {Array<string>} Array of QATM spreadsheet IDs
 */
function getQATMSpreadsheetIds_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');

  if (!configSheet) {
    Logger.log('Config tab not found');
    return [];
  }

  const data = configSheet.getDataRange().getValues();
  const spreadsheetIds = [];

  // Look for QATM spreadsheet IDs in Config
  // Format expected: Column B contains spreadsheet IDs
  // Skip header rows (first 2 rows)
  for (let i = 2; i < data.length; i++) {
    const row = data[i];

    // Column A = Setting name, Column B = Value
    const settingName = String(row[0] || '').trim();
    const value = String(row[1] || '').trim();

    // Skip if not a QATM spreadsheet ID row
    if (!settingName || !value) continue;
    if (settingName.toLowerCase().includes('vapt')) continue; // Skip VAPT IDs
    if (value === 'PASTE_QATM_SPREADSHEET_ID_HERE') continue; // Skip placeholder

    // Check if value looks like a spreadsheet ID (alphanumeric, length > 30)
    if (value.length > 30 && /^[a-zA-Z0-9_-]+$/.test(value)) {
      spreadsheetIds.push(value);
      Logger.log(`Found QATM ID: ${value} (${settingName})`);
    }
  }

  return spreadsheetIds;
}

/**
 * Helper: Apply TC notes fix to a worksheet
 * @param {Sheet} ws - TC_Master worksheet
 */
function applyTCNotesFix_(ws) {
  // Headers unchanged - only fixing notes!
  const notes = [
    // Col 1: TEST CASE ID
    'TEST CASE ID\nUnique identifier untuk test case\n\nFORMAT RECOMMENDED:\nTC001, TC002, TC003...\natau TC-LOGIN-001, TC-DASHBOARD-001\n\nPENTING:\n• Harus unique dalam 1 modul\n• Jangan diubah setelah dibuat (untuk traceability)\n• Gunakan format yang konsisten',

    // Col 2: MODUL
    'MODUL\nModul yang ditest - harus match Summary > Modul\n\nContoh: 1, 2, 3.1, Login, Dashboard\n\nGUNAKAN:\nKonsisten dengan naming di Summary untuk grouping yang benar\n\nSubModule = smallest standalone unit (one app or one domain).\nMust be IDENTICAL in TC_Master and API_Master for Dashboard coverage to merge correctly.',

    // Col 3: TEST SCENARIO
    'TEST SCENARIO\nDeskripsi skenario yang akan ditest\n\nContoh:\n• "Verify user can login with valid credentials"\n• "Verify error message shown for invalid email format"\n• "Verify dashboard loads within 3 seconds"\n\nTIPS:\nMulai dengan "Verify" atau "Test that"\nJelaskan WHAT akan ditest, bukan HOW',

    // Col 4: Feature
    'Feature — Specific feature or page name.\nExamples: Login Page, Checkout Flow, User Management\nUsed for grouping Coverage per Feature in Summary.',

    // Col 5: Priority
    'Priority & Impact:\n\nCRITICAL  → Must PASS before release. FAIL = release BLOCKED.\nHIGH      → Must PASS in same sprint. FAIL = needs PM approval.\nMEDIUM    → Potential blocker. Fix before UAT.\nLOW       → Non-blocker. Fix in next sprint.\nLOWEST    → Nice to have. Optional.\n\nAuto Test Level:\nCritical / High / Medium → Smoke Test\nLow / Lowest             → Regression Test',

    // Col 6: Platform
    'Platform: Web / Mobile / Web & Mobile',

    // Col 7: Test Type
    'Test Type:\n  Positive   = happy path (valid data, expected flow)\n  Negative   = error case (invalid input, rejected action)\n  Edge Case  = boundary condition',

    // Col 8: Automated
    'Automation Status:\n  Automated           = script exists and runs\n  To Do               = planned, not yet done\n  Manual              = decided to stay manual\n  Cannot be Automated = technically not possible',

    // Col 9: Version
    'Version: App version when TC was created. e.g. v1.0, v2.3',

    // Col 10: Role (RBAC)
    'Role (RBAC) — The user role executing this scenario.\n\nExamples: Admin, Super Admin, User, Viewer, Operator, Supervisor, Guest\n\nUsed to:\n  • Verify test coverage per role\n  • Confirm access control (RBAC) is correct\n  • Ensure 403 Forbidden for unauthorized roles',

    // Col 11: Scenario
    'SCENARIO NAMING STANDARD\n\nHappy Path : [Role] Successfully [Verb] [Object]\nNegative   : [Role] Failed to [Verb] [Object] with [Condition]\n\nRules:\n  • Role, Object → Title Case  (Nutritionist, Meal Plan)\n  • Verb → active  (Create, Pick Up, Confirm, Submit)\n  • Do not use: success/succeed, do, perform, process\n\nStandard examples:\n  Nutritionist Successfully Creates Meal Plan\n  Admin Failed to Delete User with Invalid ID\n\n───────────────────────────────────────\nSCENARIO OUTLINE (write here in this column)\nUse when same steps + same outcome type, different data.\nRule: all Examples = Positive only OR Negative only.\n\nNegative example:\nUser Failed to Log In with <invalid_credential>\nExamples:\n- Invalid password\n- Empty password\n- Expired session\n\nPositive example:\nAdmin Successfully Creates User with Role <role>\nExamples:\n- Viewer\n- Operator\n- Supervisor',

    // Col 12: Steps / Gherkin
    '[REQUIRED] Steps in Gherkin format:\n\n  Given : Pre-condition / initial state\n          e.g. Given user is on the Login page\n  When  : Action performed by the actor\n          e.g. When user submits the form\n  And   : Additional action if needed\n\nDO NOT write Then here — Then goes in Expected Result.\n\nFor Scenario Outline, use <angle_brackets>:\n  When user submits login with "<invalid_credential>"',

    // Col 13: Expected Result
    '[REQUIRED] Expected Result in Gherkin Then format:\n\n  Then : Outcome / state change after action completes\n\nBe specific — name the UI element, message, or status.\n\nPositive example:\n  Then dashboard is displayed, username shown in header\nNegative example:\n  Then login is rejected with message "Incorrect password"\n\nFor Scenario Outline: write the common outcome.\nRow detail goes in the Examples column.',

    // Col 14: [AUTO] Test Level
    '[AUTO — DO NOT EDIT] Test Level auto-calculated from Priority:\nCritical/High/Medium = Smoke  |  Low/Lowest = Regression'
  ];

  notes.forEach((note, i) => {
    ws.getRange(2, i + 1).setNote(note);
  });

  Logger.log(`✅ Applied notes fix to: ${ws.getParent().getName()}`);
}
