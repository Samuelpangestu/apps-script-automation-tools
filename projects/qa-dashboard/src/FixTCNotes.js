/**
 * FixTCNotes.js
 *
 * Fix TC_Master column notes yang tergeser
 * Broadcast ke semua QATM spreadsheets yang terdaftar di Config
 */

/**
 * Broadcast TC_Master notes fix ke semua QATM spreadsheets (Non-interactive version)
 * Can be run from triggers or scripts without user interaction
 * Reads from Config tab untuk dapat list QATM spreadsheet IDs
 */
function broadcastFixTCNotesAuto() {
  Logger.log('🔧 Starting TC_Master notes broadcast fix (auto mode)...');
  return executeBroadcastFix_();
}

/**
 * Broadcast TC_Master notes fix ke semua QATM spreadsheets (Interactive version)
 * Requires user confirmation via UI
 * Reads from Config tab untuk dapat list QATM spreadsheet IDs
 */
function broadcastFixTCNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Fix TC_Master Column Notes',
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

  const result = executeBroadcastFix_();

  // Show results in UI
  const ui2 = SpreadsheetApp.getUi();
  ui2.alert('Broadcast Results', result.message, ui2.ButtonSet.OK);
}

/**
 * Execute broadcast fix (shared logic for both interactive and non-interactive)
 * @returns {Object} Result object with success/fail counts and message
 */
function executeBroadcastFix_() {

  try {
    // Get QATM spreadsheet IDs from Config
    const spreadsheetIds = getQATMSpreadsheetIds_();

    if (spreadsheetIds.length === 0) {
      const errorMsg = 'No QATM Spreadsheets Found. Tidak ada QATM spreadsheet IDs yang ditemukan di Config tab.';
      Logger.log('❌ ' + errorMsg);
      return {
        success: false,
        message: errorMsg,
        successCount: 0,
        failCount: 0,
        errors: [errorMsg]
      };
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
          const error = `${qatmSS.getName()}: TC_Master tab not found`;
          errors.push(error);
          Logger.log('❌ ' + error);
          failCount++;
          return;
        }

        // Apply fixes
        applyTCNotesFix_(ws);

        Logger.log(`✅ Fixed: ${qatmSS.getName()}`);
        successCount++;

      } catch (e) {
        const error = `${id}: ${e.message}`;
        Logger.log(`❌ Error on ${id}: ${e.message}`);
        errors.push(error);
        failCount++;
      }
    });

    // Build result message
    let message = `Broadcast Complete!\n\n`;
    message += `✅ Success: ${successCount} spreadsheet(s)\n`;
    message += `❌ Failed: ${failCount} spreadsheet(s)`;

    if (errors.length > 0) {
      message += `\n\nErrors:\n${errors.join('\n')}`;
    }

    Logger.log(message);

    return {
      success: successCount > 0,
      message: message,
      successCount: successCount,
      failCount: failCount,
      errors: errors
    };

  } catch (e) {
    const errorMsg = 'Error in broadcast fix: ' + e.message;
    Logger.log('❌ ' + errorMsg);
    return {
      success: false,
      message: errorMsg,
      successCount: 0,
      failCount: 1,
      errors: [errorMsg]
    };
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

  Logger.log(`Config tab has ${data.length} rows`);

  // Look for QATM spreadsheet IDs in Config
  // Config structure in Dashboard: Project, Modul, Submodul, PIC QA, QATM URL, Bug Report GID
  // We need to extract spreadsheet IDs from QATM URLs (column E, index 4)

  // Skip header rows (first 4 rows based on Config structure)
  for (let i = 4; i < data.length; i++) {
    const row = data[i];

    // Skip empty rows
    if (!row[0]) continue;

    const project = String(row[0] || '').trim();
    const qatmUrl = String(row[4] || '').trim(); // Column E = QATM URL

    // Extract spreadsheet ID from URL
    // Format: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit
    if (qatmUrl && qatmUrl.includes('spreadsheets/d/')) {
      const match = qatmUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        const id = match[1];
        spreadsheetIds.push(id);
        Logger.log(`Found QATM ID: ${id} (Project: ${project})`);
      }
    }
  }

  Logger.log(`Total QATM spreadsheets found: ${spreadsheetIds.length}`);
  return spreadsheetIds;
}

/**
 * Debug function to check Config tab structure
 */
function debugConfigStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');

  if (!configSheet) {
    Logger.log('Config tab not found');
    return;
  }

  const data = configSheet.getDataRange().getValues();

  Logger.log('=== CONFIG TAB STRUCTURE ===');
  Logger.log(`Total rows: ${data.length}`);
  Logger.log('');
  Logger.log('First 5 rows:');

  for (let i = 0; i < Math.min(5, data.length); i++) {
    Logger.log(`Row ${i}: ${JSON.stringify(data[i])}`);
  }

  Logger.log('');
  Logger.log('Sample data rows (row 4-8):');
  for (let i = 4; i < Math.min(9, data.length); i++) {
    const row = data[i];
    Logger.log(`Row ${i}: Project="${row[0]}", Modul="${row[1]}", QATM URL="${row[4]}"`);
  }
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
