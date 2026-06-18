/**
 * VerifyEpicMenu.js - Verification script untuk Epic Management menu
 *
 * HOW TO USE:
 * 1. Push script ini ke Apps Script: clasp push
 * 2. Run function: verifyEpicMenuExists()
 * 3. Check execution log untuk hasil
 */

/**
 * Verify Epic Management functions exist and are callable
 */
function verifyEpicMenuExists() {
  Logger.log('═══════════════════════════════════════════');
  Logger.log('🔍 VERIFYING EPIC MANAGEMENT DEPLOYMENT');
  Logger.log('═══════════════════════════════════════════');

  const checks = {
    passed: [],
    failed: []
  };

  // Check 1: Functions exist
  Logger.log('\n1️⃣ Checking if Epic Management functions exist...');

  const functionsToCheck = [
    'createOrUpdateJiraEpics',
    'showEpicStatus',
    '_getEpicModules_',
    '_createJiraEpic_',
    '_updateJiraEpic_',
    '_getJiraEpic_'
  ];

  functionsToCheck.forEach(fnName => {
    try {
      const fn = this[fnName] || eval(fnName);
      if (typeof fn === 'function') {
        checks.passed.push('✅ Function exists: ' + fnName);
        Logger.log('  ✅ ' + fnName + ' exists');
      } else {
        checks.failed.push('❌ Not a function: ' + fnName);
        Logger.log('  ❌ ' + fnName + ' is not a function');
      }
    } catch (e) {
      checks.failed.push('❌ Function not found: ' + fnName + ' - ' + e.message);
      Logger.log('  ❌ ' + fnName + ' not found: ' + e.message);
    }
  });

  // Check 2: Constants exist
  Logger.log('\n2️⃣ Checking if constants exist...');

  try {
    if (typeof JIRA_INSTANCES_ !== 'undefined') {
      checks.passed.push('✅ JIRA_INSTANCES_ constant exists');
      Logger.log('  ✅ JIRA_INSTANCES_ exists');
      Logger.log('     Instances: ' + Object.keys(JIRA_INSTANCES_).join(', '));
    } else {
      checks.failed.push('❌ JIRA_INSTANCES_ not found');
      Logger.log('  ❌ JIRA_INSTANCES_ not found');
    }
  } catch (e) {
    checks.failed.push('❌ JIRA_INSTANCES_ error: ' + e.message);
    Logger.log('  ❌ JIRA_INSTANCES_ error: ' + e.message);
  }

  try {
    if (typeof EPIC_COL_INDEX_ !== 'undefined') {
      checks.passed.push('✅ EPIC_COL_INDEX_ constant exists');
      Logger.log('  ✅ EPIC_COL_INDEX_ exists (value: ' + EPIC_COL_INDEX_ + ')');
    } else {
      checks.failed.push('❌ EPIC_COL_INDEX_ not found');
      Logger.log('  ❌ EPIC_COL_INDEX_ not found');
    }
  } catch (e) {
    checks.failed.push('❌ EPIC_COL_INDEX_ error: ' + e.message);
    Logger.log('  ❌ EPIC_COL_INDEX_ error: ' + e.message);
  }

  // Check 3: Config sheet has Column X
  Logger.log('\n3️⃣ Checking Config sheet Column X...');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = ss.getSheetByName('Config');

    if (!cfg) {
      checks.failed.push('❌ Config sheet not found');
      Logger.log('  ❌ Config sheet not found');
    } else {
      const headerX = cfg.getRange(3, 24).getValue();
      if (String(headerX).includes('Epic')) {
        checks.passed.push('✅ Column X (Epic Key) exists in Config');
        Logger.log('  ✅ Column X header: "' + headerX + '"');
      } else {
        checks.failed.push('❌ Column X header incorrect: ' + headerX);
        Logger.log('  ❌ Column X header incorrect: "' + headerX + '"');
        Logger.log('     Expected: "Jira Epic Key"');
      }
    }
  } catch (e) {
    checks.failed.push('❌ Config check error: ' + e.message);
    Logger.log('  ❌ Config check error: ' + e.message);
  }

  // Check 4: Test onOpen menu structure
  Logger.log('\n4️⃣ Testing menu creation...');

  try {
    // Force reload menu
    onOpen();
    checks.passed.push('✅ onOpen() executed successfully');
    Logger.log('  ✅ onOpen() executed successfully');
    Logger.log('  ℹ️  Menu should now be visible in spreadsheet');
  } catch (e) {
    checks.failed.push('❌ onOpen() error: ' + e.message);
    Logger.log('  ❌ onOpen() error: ' + e.message);
  }

  // Summary
  Logger.log('\n═══════════════════════════════════════════');
  Logger.log('📊 VERIFICATION SUMMARY');
  Logger.log('═══════════════════════════════════════════');
  Logger.log('Passed: ' + checks.passed.length);
  Logger.log('Failed: ' + checks.failed.length);

  if (checks.failed.length === 0) {
    Logger.log('\n✅ ALL CHECKS PASSED!');
    Logger.log('Epic Management is deployed correctly.');
    Logger.log('\nNext steps:');
    Logger.log('1. Refresh your spreadsheet');
    Logger.log('2. Menu should appear: 🎯 QA Dashboard → 📋 Epic Management');
  } else {
    Logger.log('\n❌ SOME CHECKS FAILED:');
    checks.failed.forEach(f => Logger.log('  ' + f));
  }

  // Show alert to user
  const msg = '🔍 EPIC MANAGEMENT VERIFICATION\n\n' +
    'Passed: ' + checks.passed.length + '\n' +
    'Failed: ' + checks.failed.length + '\n\n' +
    (checks.failed.length === 0
      ? '✅ All checks passed!\n\nRefresh spreadsheet to see menu:\n🎯 QA Dashboard → 📋 Epic Management'
      : '❌ Some checks failed:\n\n' + checks.failed.join('\n'));

  SpreadsheetApp.getUi().alert(msg);

  return checks;
}

/**
 * Quick test: Just reload menu without verification
 */
function quickReloadEpicMenu() {
  Logger.log('🔄 Quick reload menu...');
  try {
    onOpen();
    Logger.log('✅ Menu reloaded');
    SpreadsheetApp.getUi().alert('✅ Menu reloaded!\n\nRefresh spreadsheet to see changes.');
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Error reloading menu:\n\n' + e.message);
  }
}
