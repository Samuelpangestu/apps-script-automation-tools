/**
 * AutomationConfigGenerator.js
 * ═══════════════════════════════════════════════════════════════════════
 * Auto-generate automation ingest config for rows with empty Tab Sheet Name
 * Columns Y-AK (25-37): Tab Sheet Name, Webhook Env Key, Webhook URL, etc.
 * ═══════════════════════════════════════════════════════════════════════
 */

/**
 * Generate automation config for empty rows in Config tab
 * Only fills rows where Tab Sheet Name (col Y) is empty
 * Preserves existing data if Tab Sheet Name already exists
 */
function generateAutomationConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Config');

  if (!ws) {
    throw new Error('Config tab not found. Please run Create Dashboard first.');
  }

  const lastRow = ws.getLastRow();
  if (lastRow < 4) {
    SpreadsheetApp.getUi().alert('No data rows found in Config tab.');
    return;
  }

  // Read all data from Config
  const data = ws.getRange(4, 1, lastRow - 3, 37).getValues(); // A4:AK (cols 1-37)

  let generatedCount = 0;
  const updates = [];

  data.forEach((row, index) => {
    const rowNum = index + 4; // Actual row number

    // Check if this row has module data (Active checkbox + Project + Modul + Submodul)
    const active = row[0]; // Col A
    const project = row[2] ? row[2].toString().trim() : ''; // Col C
    const modul = row[3] ? row[3].toString().trim() : ''; // Col D
    const submodul = row[4] ? row[4].toString().trim() : ''; // Col E
    const tabSheetName = row[24] ? row[24].toString().trim() : ''; // Col Y (index 24)

    // Only generate if:
    // 1. Row has Project/Modul/Submodul data
    // 2. Tab Sheet Name is empty (not yet generated)
    if (active && project && modul && submodul && !tabSheetName) {
      const config = generateConfigForModule_(project, modul, submodul);
      updates.push({
        row: rowNum,
        config: config
      });
      generatedCount++;
    }
  });

  if (updates.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Empty Rows Found',
      'All rows with module data already have automation config.\n\n' +
      'Tab Sheet Name (column Y) exists for all active modules.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Write generated configs
  updates.forEach(update => {
    ws.getRange(update.row, 25, 1, 13).setValues([update.config]); // Y-AK (cols 25-37)
  });

  SpreadsheetApp.getUi().alert(
    'Automation Config Generated! ✅',
    'Generated automation config for ' + generatedCount + ' modules.\n\n' +
    'Existing configs were preserved (not overwritten).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log('✅ Generated automation config for ' + generatedCount + ' modules');
}

/**
 * Generate automation config for a single module
 * @param {string} project - Project name
 * @param {string} modul - Modul name
 * @param {string} submodul - Submodul name
 * @returns {Array} Config array [Tab Sheet Name, Webhook Env Key, ..., Report Channel]
 */
function generateConfigForModule_(project, modul, submodul) {
  // Helper: Convert to camelCase (e.g., "E2E DAPUR" -> "e2eDapur")
  const toCamelCase = (str) => {
    return str.toString()
      .trim()
      .split(/[\s-_.]+/)
      .map((word, index) => {
        if (index === 0) return word.toLowerCase();
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
      })
      .join('');
  };

  // Helper: Convert to uppercase no spaces (e.g., "E2E DAPUR" -> "E2EDAPUR")
  const toUpperNoSpace = (str) => {
    return str.toString()
      .trim()
      .replace(/[\s-_.]+/g, '')
      .toUpperCase();
  };

  // Helper: Convert to lowercase with hyphens (e.g., "E2E DAPUR" -> "e2e-dapur")
  const toLowerHyphen = (str) => {
    return str.toString()
      .trim()
      .toLowerCase()
      .replace(/[\s_]+/g, '-')
      .replace(/\.+/g, '.');
  };

  // Helper: Convert to lowercase with hyphens, handling special chars
  const toLowerHyphenClean = (str) => {
    return str.toString()
      .trim()
      .toLowerCase()
      .replace(/[\s_]+/g, '-')
      .replace(/[()&]/g, '-')
      .replace(/\.+/g, '.')
      .replace(/-+/g, '-')
      .replace(/^-|-$/g, '');
  };

  const projectCamel = toCamelCase(project);
  const modulCamel = toCamelCase(modul);
  const submodulCamel = toCamelCase(submodul);

  const modulUpper = toUpperNoSpace(modul);
  const submodulUpper = toUpperNoSpace(submodul);

  const projectLower = toLowerHyphen(project);
  const modulLower = toLowerHyphenClean(modul);
  const submodulLower = toLowerHyphenClean(submodul);

  return [
    // Y: Tab Sheet Name (MODUL-SUBMODUL uppercase no space)
    modulUpper + '-' + submodulUpper,

    // Z: Webhook Env Key (WEBHOOK_URL_MODUL_SUBMODUL)
    'WEBHOOK_URL_' + modulUpper + '_' + submodulUpper,

    // AA: Webhook URL (empty by default)
    '',

    // AB: Test Env (default: Staging)
    'Staging',

    // AC: Web Job Pattern
    'qa-speedweb-' + projectLower + '-' + modulCamel + '-' + submodulCamel + '-[suite]-[env]',

    // AD: API Job Pattern
    'qa-api-' + projectLower + '-' + modulCamel + '-' + submodulCamel + '-[suite]-[env]',

    // AE: Web Tag Pattern
    modulCamel + '-' + submodulCamel + '-[suite]-[env]',

    // AF: API Tag Pattern
    modulCamel + '-' + submodulCamel + '-[suite]-[env]',

    // AG: Web Enabled (default: FALSE)
    false,

    // AH: API Enabled (default: FALSE)
    false,

    // AI: Automation Owner (empty by default)
    '',

    // AJ: Automation Notes (empty by default)
    '',

    // AK: Report Channel
    'qa-' + modulLower + '-' + submodulLower
  ];
}

/**
 * Menu function wrapper for generateAutomationConfig
 */
function menuGenerateAutomationConfig() {
  try {
    generateAutomationConfig();
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to generate automation config:\n\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    Logger.log('Generate Automation Config error: ' + error.stack);
  }
}
