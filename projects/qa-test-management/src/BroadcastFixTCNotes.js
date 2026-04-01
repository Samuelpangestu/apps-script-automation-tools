/**
 * BroadcastFixTCNotes.js
 *
 * Fix TC_Master column notes yang tergeser
 * Broadcast ke semua QATM spreadsheets
 */

/**
 * Fix TC_Master column notes only (headers unchanged)
 * Run this on individual QATM spreadsheet
 */
function fixTCMasterNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('TC_Master');

  if (!ws) {
    SpreadsheetApp.getUi().alert('TC_Master tab not found!');
    return;
  }

  Logger.log('Fixing TC_Master column notes...');

  // Headers unchanged - only fixing notes!

  // Column 1: TEST CASE ID
  ws.getRange(2, 1).setNote(
    'TEST CASE ID\n' +
    'Unique identifier untuk test case\n' +
    '\n' +
    'FORMAT RECOMMENDED:\n' +
    'TC001, TC002, TC003...\n' +
    'atau TC-LOGIN-001, TC-DASHBOARD-001\n' +
    '\n' +
    'PENTING:\n' +
    '• Harus unique dalam 1 modul\n' +
    '• Jangan diubah setelah dibuat (untuk traceability)\n' +
    '• Gunakan format yang konsisten'
  );

  // Column 2: MODUL
  ws.getRange(2, 2).setNote(
    'MODUL\n' +
    'Modul yang ditest - harus match Summary > Modul\n' +
    '\n' +
    'Contoh: 1, 2, 3.1, Login, Dashboard\n' +
    '\n' +
    'GUNAKAN:\n' +
    'Konsisten dengan naming di Summary untuk grouping yang benar\n' +
    '\n' +
    'SubModule = smallest standalone unit (one app or one domain).\n' +
    'Must be IDENTICAL in TC_Master and API_Master for Dashboard coverage to merge correctly.'
  );

  // Column 3: TEST SCENARIO
  ws.getRange(2, 3).setNote(
    'TEST SCENARIO\n' +
    'Deskripsi skenario yang akan ditest\n' +
    '\n' +
    'Contoh:\n' +
    '• "Verify user can login with valid credentials"\n' +
    '• "Verify error message shown for invalid email format"\n' +
    '• "Verify dashboard loads within 3 seconds"\n' +
    '\n' +
    'TIPS:\n' +
    'Mulai dengan "Verify" atau "Test that"\n' +
    'Jelaskan WHAT akan ditest, bukan HOW'
  );

  // Column 4: Feature (keep existing note, just ensure it's correct)
  ws.getRange(2, 4).setNote(
    'Feature — Specific feature or page name.\n' +
    'Examples: Login Page, Checkout Flow, User Management\n' +
    'Used for grouping Coverage per Feature in Summary.'
  );

  // Column 5: Priority (keep existing note)
  ws.getRange(2, 5).setNote(
    'Priority & Impact:\n' +
    '\n' +
    'CRITICAL  → Must PASS before release. FAIL = release BLOCKED.\n' +
    'HIGH      → Must PASS in same sprint. FAIL = needs PM approval.\n' +
    'MEDIUM    → Potential blocker. Fix before UAT.\n' +
    'LOW       → Non-blocker. Fix in next sprint.\n' +
    'LOWEST    → Nice to have. Optional.\n' +
    '\n' +
    'Auto Test Level:\n' +
    'Critical / High / Medium → Smoke Test\n' +
    'Low / Lowest             → Regression Test'
  );

  // Column 6: Platform
  ws.getRange(2, 6).setNote('Platform: Web / Mobile / Web & Mobile');

  // Column 7: Test Type
  ws.getRange(2, 7).setNote(
    'Test Type:\n' +
    '  Positive   = happy path (valid data, expected flow)\n' +
    '  Negative   = error case (invalid input, rejected action)\n' +
    '  Edge Case  = boundary condition'
  );

  // Column 8: Automated
  ws.getRange(2, 8).setNote(
    'Automation Status:\n' +
    '  Automated           = script exists and runs\n' +
    '  To Do               = planned, not yet done\n' +
    '  Manual              = decided to stay manual\n' +
    '  Cannot be Automated = technically not possible'
  );

  // Column 9: Version
  ws.getRange(2, 9).setNote('Version: App version when TC was created. e.g. v1.0, v2.3');

  // Column 10: Role (RBAC)
  ws.getRange(2, 10).setNote(
    'Role (RBAC) — The user role executing this scenario.\n' +
    '\n' +
    'Examples: Admin, Super Admin, User, Viewer, Operator, Supervisor, Guest\n' +
    '\n' +
    'Used to:\n' +
    '  • Verify test coverage per role\n' +
    '  • Confirm access control (RBAC) is correct\n' +
    '  • Ensure 403 Forbidden for unauthorized roles'
  );

  // Column 11: Scenario
  ws.getRange(2, 11).setNote(
    'SCENARIO NAMING STANDARD\n' +
    '\n' +
    'Happy Path : [Role] Successfully [Verb] [Object]\n' +
    'Negative   : [Role] Failed to [Verb] [Object] with [Condition]\n' +
    '\n' +
    'Rules:\n' +
    '  • Role, Object → Title Case  (Nutritionist, Meal Plan)\n' +
    '  • Verb → active  (Create, Pick Up, Confirm, Submit)\n' +
    '  • Do not use: success/succeed, do, perform, process\n' +
    '\n' +
    'Standard examples:\n' +
    '  Nutritionist Successfully Creates Meal Plan\n' +
    '  Admin Failed to Delete User with Invalid ID\n' +
    '\n' +
    '───────────────────────────────────────\n' +
    'SCENARIO OUTLINE (write here in this column)\n' +
    'Use when same steps + same outcome type, different data.\n' +
    'Rule: all Examples = Positive only OR Negative only.\n' +
    '\n' +
    'Negative example:\n' +
    'User Failed to Log In with <invalid_credential>\n' +
    'Examples:\n' +
    '- Invalid password\n' +
    '- Empty password\n' +
    '- Expired session\n' +
    '\n' +
    'Positive example:\n' +
    'Admin Successfully Creates User with Role <role>\n' +
    'Examples:\n' +
    '- Viewer\n' +
    '- Operator\n' +
    '- Supervisor'
  );

  // Column 12: Steps / Gherkin
  ws.getRange(2, 12).setNote(
    '[REQUIRED] Steps in Gherkin format:\n' +
    '\n' +
    '  Given : Pre-condition / initial state\n' +
    '          e.g. Given user is on the Login page\n' +
    '  When  : Action performed by the actor\n' +
    '          e.g. When user submits the form\n' +
    '  And   : Additional action if needed\n' +
    '\n' +
    'DO NOT write Then here — Then goes in Expected Result.\n' +
    '\n' +
    'For Scenario Outline, use <angle_brackets>:\n' +
    '  When user submits login with "<invalid_credential>"'
  );

  // Column 13: Expected Result
  ws.getRange(2, 13).setNote(
    '[REQUIRED] Expected Result in Gherkin Then format:\n' +
    '\n' +
    '  Then : Outcome / state change after action completes\n' +
    '\n' +
    'Be specific — name the UI element, message, or status.\n' +
    '\n' +
    'Positive example:\n' +
    '  Then dashboard is displayed, username shown in header\n' +
    'Negative example:\n' +
    '  Then login is rejected with message "Incorrect password"\n' +
    '\n' +
    'For Scenario Outline: write the common outcome.\n' +
    'Row detail goes in the Examples column.'
  );

  // Column 14: [AUTO] Test Level
  ws.getRange(2, 14).setNote(
    '[AUTO — DO NOT EDIT] Test Level auto-calculated from Priority:\n' +
    'Critical/High/Medium = Smoke  |  Low/Lowest = Regression'
  );

  Logger.log('✅ TC_Master notes fixed');

  SpreadsheetApp.getUi().alert(
    '✅ TC_Master Notes Fixed!',
    'Column notes have been corrected for all 14 columns:\n\n' +
    '• Col 1: No → TEST CASE ID notes\n' +
    '• Col 2: SubModul → MODUL notes\n' +
    '• Col 3: TC_ID → TEST SCENARIO notes\n' +
    '• All other columns: notes repositioned correctly\n\n' +
    'Headers unchanged (tetap sama)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Broadcast fix to multiple QATM spreadsheets
 * Edit the spreadsheetIds array below with your QATM spreadsheet IDs
 */
function broadcastFixTCNotes() {
  // ⚠️ EDIT THIS: Add your QATM Spreadsheet IDs here
  const spreadsheetIds = [
    // Example:
    // '1ABC123...',
    // '1XYZ789...',
  ];

  if (spreadsheetIds.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Spreadsheets Configured',
      'Please edit broadcastFixTCNotes() function and add spreadsheet IDs to the array.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  let successCount = 0;
  let failCount = 0;
  const errors = [];

  spreadsheetIds.forEach((id, index) => {
    try {
      Logger.log(`Processing spreadsheet ${index + 1}/${spreadsheetIds.length}: ${id}`);

      const ss = SpreadsheetApp.openById(id);
      const ws = ss.getSheetByName('TC_Master');

      if (!ws) {
        errors.push(`${ss.getName()}: TC_Master tab not found`);
        failCount++;
        return;
      }

      // Apply fixes (same logic as fixTCMasterNotes but without UI alerts)
      applyTCNotesFix_(ws);

      Logger.log(`✅ Fixed: ${ss.getName()}`);
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

  SpreadsheetApp.getUi().alert('Broadcast Results', message, SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log(message);
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
}
