/**
 * EpicManagement.js  —  Jira Epic Creation & Management for QA Dashboard
 * ═══════════════════════════════════════════════════════════════════════
 *
 * FEATURES:
 *   - Create Jira Epic per active module (from Config tab)
 *   - Update Epic summary if Module/Submodule changed
 *   - Auto-skip if Epic already exists (based on Epic Key in Config col X)
 *   - Show Epic status summary
 *
 * USAGE:
 *   Menu: 🎯 QA Dashboard > 📋 Epic Management
 *     - ⭐ Create/Update Epic for Active Modules
 *     - 🔍 Show Epic Status
 *
 * EPIC FORMAT:
 *   Summary: [Module] Submodule
 *   Description: (empty - user fills manually in Jira)
 *
 * CONFIG:
 *   - Column X (24): Jira Epic Key (auto-filled after creation)
 *   - Empty Epic Key = will create new Epic
 *   - Existing Epic Key = will check and update if Module/Submodule changed
 *
 * DEPENDENCIES:
 *   - JIRA_INSTANCES_ from JiraSync.js (reused constants)
 *   - _getCreds_() from JiraSync.js (reused credentials reader)
 * ═══════════════════════════════════════════════════════════════════════
 */

// Constants from JiraSync.js (should be in global scope)
// If not available, define them here:
// const JIRA_INSTANCES_ = {
//   'digitalperuri': 'https://digitalperuri.atlassian.net',
//   'bgn-peruri':    'https://bgn-peruri.atlassian.net',
// };

const EPIC_COL_INDEX_ = 23; // Column X (0-indexed = 23, 1-indexed = 24)

// ═══════════════════════════════════════════════════════════════════════
// MENU FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * Create/Update Jira Epic for all active modules
 *
 * Logic:
 * 1. Get all modules where: Active=TRUE + Jira Instance valid + Project Key valid
 * 2. For each module:
 *    - If Epic Key empty → CREATE new Epic → Store key in Config col X
 *    - If Epic Key exists → CHECK if Module/Submodule changed → UPDATE Epic
 * 3. Show summary: Created, Updated, Skipped, Failed
 */
function createOrUpdateJiraEpics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('═══════════════════════════════════════════');
  Logger.log('📋 CREATE/UPDATE JIRA EPIC - Started');
  Logger.log('═══════════════════════════════════════════');

  try {
    // Get credentials
    const creds = _getCreds_(ss);
    if (!creds || Object.keys(creds).length === 0) {
      safeAlert_('❌ Jira credentials tidak ditemukan!\n\nSetup credentials di tab Credentials terlebih dahulu.');
      return;
    }

    // Get modules that need Epic
    const modules = _getEpicModules_(ss);

    if (modules.length === 0) {
      safeAlert_('⚠️ Tidak ada modul yang memenuhi syarat untuk Epic creation.\n\n' +
        'Pastikan:\n' +
        '1. Active = TRUE (col A)\n' +
        '2. Jira Instance valid (col I)\n' +
        '3. Jira Project valid (col J)');
      return;
    }

    Logger.log('Found ' + modules.length + ' module(s) to process');

    // Process each module
    const results = {
      created: [],
      updated: [],
      skipped: [],
      failed: []
    };

    const cfg = ss.getSheetByName('Config');

    modules.forEach((mod, idx) => {
      Logger.log('');
      Logger.log('─────────────────────────────────────────');
      Logger.log('[' + (idx + 1) + '/' + modules.length + '] Processing: ' + mod.name);
      Logger.log('Instance: ' + mod.instance);
      Logger.log('Project: ' + mod.projectKey);
      Logger.log('Current Epic Key: ' + (mod.epicKey || '(empty)'));

      try {
        const cred = creds[mod.instance];
        if (!cred) {
          throw new Error('Credentials tidak ada untuk instance: ' + mod.instance);
        }

        // Epic summary format: [Module] Submodule
        const epicSummary = '[' + (mod.module || 'Unknown') + '] ' + (mod.submodule || mod.module || 'Unknown');

        if (!mod.epicKey) {
          // CREATE new Epic
          Logger.log('Creating new Epic...');
          const epicKey = _createJiraEpic_(mod.instance, mod.projectKey, epicSummary, '', cred);

          if (epicKey) {
            // Update Config with Epic Key
            cfg.getRange(mod.rowIndex, EPIC_COL_INDEX_ + 1).setValue(epicKey);
            results.created.push(mod.name + ' → ' + epicKey);
            Logger.log('✅ Epic created: ' + epicKey);
          } else {
            throw new Error('Failed to create Epic (API error)');
          }

        } else {
          // CHECK if Epic exists and needs update
          Logger.log('Checking existing Epic...');
          const epicData = _getJiraEpic_(mod.instance, mod.epicKey, cred);

          if (!epicData) {
            // Epic doesn't exist (deleted or moved), clear Epic Key and create new
            Logger.log('⚠️ Epic not found in Jira, creating new...');
            cfg.getRange(mod.rowIndex, EPIC_COL_INDEX_ + 1).setValue('');

            const epicKey = _createJiraEpic_(mod.instance, mod.projectKey, epicSummary, '', cred);
            if (epicKey) {
              cfg.getRange(mod.rowIndex, EPIC_COL_INDEX_ + 1).setValue(epicKey);
              results.created.push(mod.name + ' → ' + epicKey + ' (replaced invalid key)');
              Logger.log('✅ Epic created: ' + epicKey);
            } else {
              throw new Error('Failed to create Epic (API error)');
            }

          } else {
            // Epic exists, check if summary changed
            if (epicData.summary !== epicSummary) {
              Logger.log('Summary changed: "' + epicData.summary + '" → "' + epicSummary + '"');
              const updated = _updateJiraEpic_(mod.instance, mod.epicKey, epicSummary, cred);

              if (updated) {
                results.updated.push(mod.name + ' → ' + mod.epicKey);
                Logger.log('✅ Epic updated');
              } else {
                throw new Error('Failed to update Epic (API error)');
              }

            } else {
              results.skipped.push(mod.name + ' → ' + mod.epicKey + ' (unchanged)');
              Logger.log('ℹ️ Epic unchanged, skipped');
            }
          }
        }

      } catch (e) {
        results.failed.push(mod.name + ' → ' + e.message);
        Logger.log('❌ Error: ' + e.message);
      }

      // Rate limiting
      Utilities.sleep(300);
    });

    // Show results
    Logger.log('');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('📋 CREATE/UPDATE JIRA EPIC - Completed');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('Created: ' + results.created.length);
    Logger.log('Updated: ' + results.updated.length);
    Logger.log('Skipped: ' + results.skipped.length);
    Logger.log('Failed:  ' + results.failed.length);

    let msg = '✅ Jira Epic Management - Selesai!\n\n';

    if (results.created.length > 0) {
      msg += '🆕 CREATED (' + results.created.length + '):\n';
      results.created.forEach(r => msg += '  • ' + r + '\n');
      msg += '\n';
    }

    if (results.updated.length > 0) {
      msg += '🔄 UPDATED (' + results.updated.length + '):\n';
      results.updated.forEach(r => msg += '  • ' + r + '\n');
      msg += '\n';
    }

    if (results.skipped.length > 0) {
      msg += '⏭️ SKIPPED (' + results.skipped.length + '):\n';
      results.skipped.forEach(r => msg += '  • ' + r + '\n');
      msg += '\n';
    }

    if (results.failed.length > 0) {
      msg += '❌ FAILED (' + results.failed.length + '):\n';
      results.failed.forEach(r => msg += '  • ' + r + '\n');
    }

    safeAlert_(msg);

  } catch (e) {
    Logger.log('❌ Fatal error: ' + e.message);
    Logger.log('Stack trace: ' + e.stack);
    safeAlert_('❌ Error:\n\n' + e.message);
  }
}

/**
 * Show Epic Status - Display summary of all Epics in Config
 */
function showEpicStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const modules = _getEpicModules_(ss);

  if (modules.length === 0) {
    safeAlert_('⚠️ Tidak ada modul aktif dengan Jira configuration.\n\n' +
      'Pastikan Active=TRUE dan Jira Instance/Project valid.');
    return;
  }

  let msg = '📋 JIRA EPIC STATUS\n';
  msg += '═════════════════════════════════════════\n\n';

  const withEpic = modules.filter(m => m.epicKey);
  const withoutEpic = modules.filter(m => !m.epicKey);

  msg += 'Total Modules: ' + modules.length + '\n';
  msg += 'With Epic: ' + withEpic.length + '\n';
  msg += 'Without Epic: ' + withoutEpic.length + '\n\n';

  if (withEpic.length > 0) {
    msg += '─────────────────────────────────────────\n';
    msg += '✅ MODULES WITH EPIC:\n';
    msg += '─────────────────────────────────────────\n\n';

    withEpic.forEach(m => {
      const base = JIRA_INSTANCES_[m.instance];
      const url = base + '/browse/' + m.epicKey;
      msg += '• ' + m.name + '\n';
      msg += '  Epic: ' + m.epicKey + '\n';
      msg += '  Instance: ' + m.instance + '\n';
      msg += '  Project: ' + m.projectKey + '\n';
      msg += '  URL: ' + url + '\n\n';
    });
  }

  if (withoutEpic.length > 0) {
    msg += '─────────────────────────────────────────\n';
    msg += '⚠️ MODULES WITHOUT EPIC:\n';
    msg += '─────────────────────────────────────────\n\n';

    withoutEpic.forEach(m => {
      msg += '• ' + m.name + '\n';
      msg += '  Instance: ' + m.instance + '\n';
      msg += '  Project: ' + m.projectKey + '\n\n';
    });

    msg += '💡 Tip: Jalankan "Create/Update Epic" untuk membuat Epic.\n';
  }

  safeAlert_(msg);
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS - Module Data
// ═══════════════════════════════════════════════════════════════════════

/**
 * Get modules that need Epic creation/management
 *
 * Filter:
 * - Col A (Active) = TRUE
 * - Col I (Jira Instance) valid and not '-'
 * - Col J (Jira Project) valid and not '-'
 *
 * Returns array with: name, instance, projectKey, module, submodule, epicKey, rowIndex
 */
function _getEpicModules_(ss) {
  const cfg = ss.getSheetByName('Config');
  if (!cfg) return [];

  const data = cfg.getDataRange().getValues();
  const modules = [];

  for (let i = 3; i < data.length; i++) { // Start from row 4 (index 3)
    const active      = data[i][0] === true;           // Col A = Active
    const project     = String(data[i][2]).trim();     // Col C = Project
    const module      = String(data[i][3]).trim();     // Col D = Modul
    const submodule   = String(data[i][4]).trim();     // Col E = Submodul
    const instance    = String(data[i][8]).trim().toLowerCase();  // Col I = Jira Instance
    const projectKey  = String(data[i][9]).trim().toUpperCase();  // Col J = Jira Project
    const epicKey     = String(data[i][EPIC_COL_INDEX_]).trim();  // Col X = Jira Epic Key

    // Validate module
    if (!active) continue;
    if (!instance || instance === '-') continue;
    if (!projectKey || projectKey === '-') continue;
    if (!JIRA_INSTANCES_[instance]) {
      Logger.log('⚠️ Unknown Jira instance: ' + instance + ' (skipped)');
      continue;
    }

    modules.push({
      name: submodule || module || 'Unknown',
      instance: instance,
      projectKey: projectKey,
      project: project,
      module: module,
      submodule: submodule,
      epicKey: epicKey || '',
      rowIndex: i + 1  // 1-indexed for getRange()
    });
  }

  return modules;
}

/**
 * Web App API: create/update Epic and the default QA planning tickets for a QATM row.
 *
 * POST payload:
 * {
 *   action: "bootstrapQatmJiraTickets",
 *   payload: { project,module,submodule,jiraInstance,jiraProject,spreadsheetId,configRow }
 * }
 *
 * Optional Script Properties:
 * - QATM_JIRA_BOOTSTRAP_TOKEN
 * - QATM_TEST_EXECUTION_ISSUE_TYPE (default: Test Execution)
 * - QATM_TEST_AUTOMATION_ISSUE_TYPE (default: Test Automation)
 */
function bootstrapQatmJiraTickets_(payload) {
  const expectedToken = PropertiesService.getScriptProperties().getProperty('QATM_JIRA_BOOTSTRAP_TOKEN');
  if (expectedToken && String(payload.token || '') !== expectedToken) {
    throw new Error('Invalid QATM Jira bootstrap token');
  }

  const ss = typeof getDashboardSpreadsheet_ === 'function'
    ? getDashboardSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName('Config');
  if (!cfg) throw new Error('Config sheet not found');

  const instance = String(payload.jiraInstance || '').trim().toLowerCase();
  const projectKey = String(payload.jiraProject || '').trim().toUpperCase();
  if (!instance || instance === '-' || !JIRA_INSTANCES_[instance]) {
    return { success: false, skipped: true, error: 'Jira Instance is empty or unsupported' };
  }
  if (!projectKey || projectKey === '-') {
    return { success: false, skipped: true, error: 'Jira Project Key is empty' };
  }

  const creds = _getCreds_(ss);
  const cred = creds[instance];
  if (!cred) throw new Error('Jira credentials not found for instance: ' + instance);

  const configRow = _findQatmConfigRow_(cfg, payload);
  if (!configRow) throw new Error('Dashboard Config row not found for QATM');

  const moduleName = String(payload.module || configRow.values[3] || '').trim();
  const submodule = String(payload.submodule || configRow.values[4] || moduleName || '').trim();
  const qatmId = String(payload.spreadsheetId || configRow.values[6] || '').trim();
  const epicSummary = '[' + (moduleName || 'Unknown') + '] ' + (submodule || moduleName || 'Unknown');

  let epicKey = String(configRow.values[EPIC_COL_INDEX_] || '').trim();
  if (!epicKey) {
    epicKey = _createJiraEpic_(instance, projectKey, epicSummary, '', cred);
    if (!epicKey) throw new Error('Failed to create Jira Epic');
    cfg.getRange(configRow.rowIndex, EPIC_COL_INDEX_ + 1).setValue(epicKey);
  }

  const existingEpicTickets = _getJiraEpicChildren_(instance, projectKey, epicKey, cred);
  if (existingEpicTickets.length > 0) {
    return {
      success: true,
      skipped: true,
      reason: 'Epic already has child tickets. Jira bootstrap skipped to avoid duplicate tickets.',
      epicKey: epicKey,
      tickets: existingEpicTickets.map(function(issue) {
        return {
          kind: 'existing',
          issueType: String(issue.fields && issue.fields.issuetype && issue.fields.issuetype.name || ''),
          key: issue.key,
          created: false,
          url: JIRA_INSTANCES_[instance] + '/browse/' + issue.key
        };
      }),
      configRow: configRow.rowIndex
    };
  }

  const props = PropertiesService.getScriptProperties();
  const issueTypes = {
    testExecution: props.getProperty('QATM_TEST_EXECUTION_ISSUE_TYPE') || 'Test Execution',
    testAutomation: props.getProperty('QATM_TEST_AUTOMATION_ISSUE_TYPE') || 'Test Automation'
  };
  const label = _jiraLabel_('qatm-' + (qatmId || moduleName + '-' + submodule).slice(0, 18));
  const baseDescription = [
    'Generated from QA Dashboard AI QATM Generator.',
    'Project: ' + String(payload.project || configRow.values[2] || '-'),
    'Module: ' + moduleName,
    'Submodule: ' + submodule,
    qatmId ? 'QATM: https://docs.google.com/spreadsheets/d/' + qatmId + '/edit' : ''
  ].filter(Boolean).join('\n');

  const tickets = [
    {
      kind: 'testExecution',
      issueType: issueTypes.testExecution,
      summary: '[UI/API] ' + epicSummary + ' - Test Execution',
      description: baseDescription + '\n\nScope: Manual test execution planning and tracking.'
    },
    {
      kind: 'uiAutomation',
      issueType: issueTypes.testAutomation,
      summary: '[UI] ' + epicSummary + ' - Test Automation',
      description: baseDescription + '\n\nScope: UI automation implementation and pipeline integration.'
    },
    {
      kind: 'apiAutomation',
      issueType: issueTypes.testAutomation,
      summary: '[API] ' + epicSummary + ' - Test Automation',
      description: baseDescription + '\n\nScope: API automation implementation and pipeline integration.'
    }
  ].map(function(ticket) {
    return _ensureJiraIssue_(instance, projectKey, epicKey, ticket, cred, [label, 'qa-dashboard', 'qatm-bootstrap']);
  });

  return {
    success: tickets.every(function(ticket) { return ticket.key; }),
    epicKey: epicKey,
    tickets: tickets,
    configRow: configRow.rowIndex
  };
}

function _findQatmConfigRow_(cfg, payload) {
  const data = cfg.getDataRange().getValues();
  const explicitRow = Number(payload.configRow) || 0;
  if (explicitRow >= 4 && explicitRow <= data.length) {
    return { rowIndex: explicitRow, values: data[explicitRow - 1] };
  }

  const project = String(payload.project || '').trim();
  const moduleName = String(payload.module || '').trim();
  const submodule = String(payload.submodule || '').trim();
  const qatmId = String(payload.spreadsheetId || '').trim();
  for (let i = 3; i < data.length; i++) {
    const sameModule = String(data[i][2] || '').trim() === project
      && String(data[i][3] || '').trim() === moduleName
      && String(data[i][4] || '').trim() === submodule;
    const sameQatm = qatmId && String(data[i][6] || '').trim() === qatmId;
    if (sameModule || sameQatm) return { rowIndex: i + 1, values: data[i] };
  }
  return null;
}

function _ensureJiraIssue_(instanceKey, projectKey, epicKey, ticket, cred, labels) {
  const existing = _findJiraIssueBySummary_(instanceKey, projectKey, ticket.summary, labels[0], cred);
  if (existing) {
    return {
      kind: ticket.kind,
      issueType: ticket.issueType,
      key: existing.key,
      created: false,
      url: JIRA_INSTANCES_[instanceKey] + '/browse/' + existing.key
    };
  }

  const key = _createJiraIssue_(instanceKey, projectKey, epicKey, ticket, cred, labels);
  return {
    kind: ticket.kind,
    issueType: ticket.issueType,
    key: key || '',
    created: Boolean(key),
    url: key ? JIRA_INSTANCES_[instanceKey] + '/browse/' + key : ''
  };
}

function _findJiraIssueBySummary_(instanceKey, projectKey, summary, label, cred) {
  const jql = 'project = "' + projectKey + '" AND labels = "' + label + '" ORDER BY created DESC';
  const data = _jiraRequest_(instanceKey, '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) + '&fields=summary,parent&maxResults=20', 'get', null, cred);
  const issues = data && data.issues ? data.issues : [];
  for (let i = 0; i < issues.length; i++) {
    if (String(issues[i].fields && issues[i].fields.summary || '') === summary) return issues[i];
  }
  return null;
}

function _getJiraEpicChildren_(instanceKey, projectKey, epicKey, cred) {
  const jql = 'project = "' + projectKey + '" AND parent = "' + epicKey + '" ORDER BY created ASC';
  const data = _jiraRequest_(instanceKey, '/rest/api/3/search/jql?jql=' + encodeURIComponent(jql) + '&fields=summary,issuetype,parent&maxResults=50', 'get', null, cred);
  return data && data.issues ? data.issues : [];
}

function _createJiraIssue_(instanceKey, projectKey, epicKey, ticket, cred, labels) {
  const payload = {
    fields: {
      project: { key: projectKey },
      summary: ticket.summary,
      description: _jiraDoc_(ticket.description),
      issuetype: { name: ticket.issueType },
      parent: epicKey ? { key: epicKey } : undefined,
      labels: labels
    }
  };
  const data = _jiraRequest_(instanceKey, '/rest/api/3/issue', 'post', payload, cred);
  return data && data.key ? data.key : null;
}

function _jiraRequest_(instanceKey, path, method, payload, cred) {
  const auth = Utilities.base64Encode(cred.email + ':' + cred.token);
  const response = UrlFetchApp.fetch(JIRA_INSTANCES_[instanceKey] + path, {
    method: method,
    headers: {
      Authorization: 'Basic ' + auth,
      'Content-Type': 'application/json'
    },
    payload: payload ? JSON.stringify(payload) : undefined,
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Jira API failed (' + code + '): ' + text.substring(0, 500));
  }
  return text ? JSON.parse(text) : {};
}

function _jiraDoc_(text) {
  return {
    type: 'doc',
    version: 1,
    content: String(text || '').split('\n').map(function(line) {
      return {
        type: 'paragraph',
        content: line ? [{ type: 'text', text: line }] : []
      };
    })
  };
}

function _jiraLabel_(value) {
  return String(value || 'qatm-bootstrap')
    .toLowerCase()
    .replace(/[^a-z0-9-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 40) || 'qatm-bootstrap';
}

// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS - Jira API
// ═══════════════════════════════════════════════════════════════════════

/**
 * Create Epic in Jira
 *
 * @param {string} instanceKey - Jira instance key (digitalperuri, bgn-peruri)
 * @param {string} projectKey - Jira project key (TEST, SQA, etc.)
 * @param {string} summary - Epic summary (format: [Module] Submodule)
 * @param {string} description - Epic description (can be empty)
 * @param {object} cred - Credentials object {email, token}
 * @return {string|null} Epic key (PROJ-123) or null if failed
 */
function _createJiraEpic_(instanceKey, projectKey, summary, description, cred) {
  const base = JIRA_INSTANCES_[instanceKey];
  const url = base + '/rest/api/3/issue';

  const auth = Utilities.base64Encode(cred.email + ':' + cred.token);
  const headers = {
    'Authorization': 'Basic ' + auth,
    'Content-Type': 'application/json'
  };

  const payload = {
    fields: {
      project: {
        key: projectKey
      },
      summary: summary,
      description: description ? {
        type: 'doc',
        version: 1,
        content: [{
          type: 'paragraph',
          content: [{
            type: 'text',
            text: description
          }]
        }]
      } : undefined,
      issuetype: {
        name: 'Epic'
      }
    }
  };

  try {
    Logger.log('API Request: POST ' + url);
    Logger.log('Payload: ' + JSON.stringify(payload, null, 2));

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('Response Code: ' + responseCode);
    Logger.log('Response: ' + responseText.substring(0, 500));

    if (responseCode === 201) {
      const data = JSON.parse(responseText);
      return data.key || null;
    } else {
      Logger.log('❌ Failed to create Epic: HTTP ' + responseCode);
      Logger.log('Response: ' + responseText);
      return null;
    }

  } catch (e) {
    Logger.log('❌ Exception during Epic creation: ' + e.message);
    return null;
  }
}

/**
 * Update Epic summary in Jira
 *
 * @param {string} instanceKey - Jira instance key
 * @param {string} epicKey - Epic key to update (PROJ-123)
 * @param {string} summary - New Epic summary
 * @param {object} cred - Credentials object {email, token}
 * @return {boolean} True if updated, false if failed
 */
function _updateJiraEpic_(instanceKey, epicKey, summary, cred) {
  const base = JIRA_INSTANCES_[instanceKey];
  const url = base + '/rest/api/3/issue/' + epicKey;

  const auth = Utilities.base64Encode(cred.email + ':' + cred.token);
  const headers = {
    'Authorization': 'Basic ' + auth,
    'Content-Type': 'application/json'
  };

  const payload = {
    fields: {
      summary: summary
    }
  };

  try {
    Logger.log('API Request: PUT ' + url);
    Logger.log('Payload: ' + JSON.stringify(payload, null, 2));

    const response = UrlFetchApp.fetch(url, {
      method: 'put',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('Response Code: ' + responseCode);
    if (responseText) {
      Logger.log('Response: ' + responseText.substring(0, 500));
    }

    if (responseCode === 204 || responseCode === 200) {
      return true;
    } else {
      Logger.log('❌ Failed to update Epic: HTTP ' + responseCode);
      Logger.log('Response: ' + responseText);
      return false;
    }

  } catch (e) {
    Logger.log('❌ Exception during Epic update: ' + e.message);
    return false;
  }
}

/**
 * Get Epic data from Jira (for verification)
 *
 * @param {string} instanceKey - Jira instance key
 * @param {string} epicKey - Epic key to fetch (PROJ-123)
 * @param {object} cred - Credentials object {email, token}
 * @return {object|null} Epic data {key, summary, ...} or null if not found
 */
function _getJiraEpic_(instanceKey, epicKey, cred) {
  const base = JIRA_INSTANCES_[instanceKey];
  const url = base + '/rest/api/3/issue/' + epicKey + '?fields=summary,status';

  const auth = Utilities.base64Encode(cred.email + ':' + cred.token);
  const headers = {
    'Authorization': 'Basic ' + auth,
    'Content-Type': 'application/json'
  };

  try {
    Logger.log('API Request: GET ' + url);

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: headers,
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('Response Code: ' + responseCode);

    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      return {
        key: data.key,
        summary: data.fields.summary || '',
        status: data.fields.status ? data.fields.status.name : ''
      };
    } else if (responseCode === 404) {
      Logger.log('Epic not found: ' + epicKey);
      return null;
    } else {
      Logger.log('❌ Failed to get Epic: HTTP ' + responseCode);
      Logger.log('Response: ' + responseText.substring(0, 500));
      return null;
    }

  } catch (e) {
    Logger.log('❌ Exception during Epic fetch: ' + e.message);
    return null;
  }
}
