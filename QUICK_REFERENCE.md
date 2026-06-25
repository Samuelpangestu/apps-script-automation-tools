# Quick Reference - Dashboard Architecture

## File Locations

| Purpose | File | Location |
|---------|------|----------|
| Dashboard Menu & Creation | MasterDashboard.js | `/projects/qa-dashboard/src/` |
| Jira Sync & Bug Import | JiraSync.js | `/projects/qa-dashboard/src/` |
| Notifications | Notifications.js | `/projects/qa-dashboard/src/` |
| Automation Config | AutomationConfigGenerator.js | `/projects/qa-dashboard/src/` |

## Config Sheet Column Map

```
A: Active               (BOOLEAN: TRUE/FALSE)
B: Jira Sync            (BOOLEAN: TRUE/FALSE)
C: Project              (TEXT: from Summary B2)
D: Modul                (TEXT: from Summary B3)
E: Submodul             (TEXT: from Summary B4)
F: PIC QA               (TEXT: from Summary B6)
G: Spreadsheet ID       (TEXT: module ID - REQUIRED)
H: Link                 (INTERNAL)
I: Jira Instance        (TEXT: digitalperuri/bgn-peruri)
J: Jira Project Key     (TEXT: e.g., TEST, SQA)
K: QA Lead              (TEXT: from Summary B5)
L: Chat Webhook         (TEXT: Google Chat URL)
M: Chat Schedule        (TEXT: e.g., 7,12,18)
N: Chat Enable          (BOOLEAN)
O: Email Recipients     (TEXT: comma-separated)
P: Email Enable         (BOOLEAN)
Q: Refresh Interval     (NUMBER: minutes)
R: Refresh Enable       (BOOLEAN)
S: WhatsApp Group ID    (TEXT: Global fallback)
T: Fonnte Token         (TEXT: SHARED, Row 4 only)
U: VAPT Enable          (BOOLEAN)
V: VAPT Spreadsheet ID  (TEXT)
W: Reserved             (FUTURE)
X-AK: Automation Config (13 columns auto-generated)
```

## Key Functions & When to Use

### Dashboard Setup
- `createDashboard()` - First-time setup, creates all 12 sheets
- `refreshDashboard()` - Refresh all data (no Jira sync)
- `refreshDashboardWithJiraSync()` - Jira sync THEN refresh (sequential)

### Jira Integration
- `jiraSetup()` - Initialize Jira integration (run once)
- `jiraActivate()` - Activate Jira with triggers (run after jiraSetup)
- `syncAllJira()` - Manual sync all modules from Jira
- `showJiraJQL()` - Debug: show JQL queries being used
- `removeJiraTriggers()` - Disable Jira auto-sync

### Notifications
- `sendBlockerNotification()` - Send blocker alerts (manual test)
- `setupDailyBlockerNotification()` - Enable daily alerts
- `setupAutoRefreshTrigger()` - Enable auto-refresh

### Automation Config
- `generateAutomationConfig()` - Auto-fill columns Y-AK when empty

## Module Determination

A module is **ACTIVE** (pulled during refresh) when:
```
Config Column A = TRUE
AND
Config Column G (Spreadsheet ID):
  - Not empty
  - Length >= 10 characters
  - NOT "PASTE_SPREADSHEET_ID_HERE"
```

## Dashboard Sheets

Created by `createDashboard()`:

1. **Config** - Module registry (manual + auto-filled)
2. **Credentials** - Jira & webhook secrets
3. **Overview** - Aggregated KPI metrics
4. **Bugs** - Consolidated bug list (from Jira)
5. **VAPT** - Security findings
6. **VAPT History** - VAPT trends
7. **Smoke** - Smoke test summary
8. **Failure Scenario** - Failed/blocked test cases
9. **Coverage** - Test coverage by feature
10. **History** - Trend snapshots (append-only)
11. **Automation Runs** - Pipeline results
12. **_Raw** - Internal cache

## Jira Sync Process

### Step 1: Every Sync
```
For each module with Active=TRUE and Jira Sync=TRUE:
  1. Clear ALL data rows (keep headers)
  2. Fetch bugs from Jira (via JQL + modul field)
  3. Filter out: "Closed" and "Won't Fix" status
  4. Insert all active bugs as new rows
  5. Add timestamps & hyperlinks
```

### Step 2: Status Update (Daily 23:00)
```
For each module with Active=TRUE:
  1. Re-fetch from Jira
  2. Update Status column only (preserve other edits)
```

### Step 3: Notifications (Daily 07:00)
```
Scan all BugReport sheets:
  1. Find Open/In Progress/Reopen bugs
  2. Group by project
  3. Send Google Chat + Email alerts
```

## Duplicate Prevention

**Current Strategy: FULL DATA REPLACEMENT**

```
Every sync:
  1. Clear all data rows from BugReport
  2. Fetch fresh bugs from Jira
  3. Insert all (closed bugs automatically excluded)
  
Result: No duplicates possible (full refresh)
Trade-off: Manual edits to bugs are lost
```

## API Integration Patterns

### Jira Cloud
```javascript
// Basic Auth
const auth = Utilities.base64Encode(email + ':' + token);
const resp = UrlFetchApp.fetch(url, {
  headers: { 'Authorization': 'Basic ' + auth }
});
```

### Google Chat
```javascript
UrlFetchApp.fetch(webhookUrl, {
  method: 'post',
  contentType: 'application/json',
  payload: JSON.stringify(messageObject)
});
```

### Email (MailApp)
```javascript
MailApp.sendEmail({ to, subject, htmlBody });
```

## Error Handling Best Practices

- Use `muteHttpExceptions: true` to prevent crashes
- Check HTTP status code before parsing
- Return `null` on error (don't throw)
- Always log full URL and query for debugging
- Handle credential not found gracefully

## Performance Tips

- `getModuleList_()` is called frequently - cache result if looping
- Jira API has rate limiting - add `Utilities.sleep(400)` between modules
- `clearContent()` instead of `deleteRows()` to preserve frozen rows
- Use `getLastRow()` only when necessary (expensive operation)

## Testing Checklist

Before deploying:
- [ ] Config sheet has valid Spreadsheet IDs (≥10 chars)
- [ ] Jira Instance selected (digitalperuri or bgn-peruri)
- [ ] Jira Project Key is correct (e.g., TEST, SQA)
- [ ] Credentials filled in (email + API token)
- [ ] At least one module has Active=TRUE
- [ ] Run `jiraActivate()` (tests connection)
- [ ] Check Execution log for errors
- [ ] Manual sync works: `syncAllJira()`
- [ ] Trigger created: check Triggers in Apps Script editor
- [ ] Email/Chat webhook URLs valid (if enabling notifications)

## Common Issues & Solutions

| Issue | Cause | Solution |
|-------|-------|----------|
| No data in Overview | No modules with Active=TRUE | Check Config column A |
| Jira sync fails | Wrong credentials | Re-run jiraSetup() + jiraActivate() |
| Jira API 401 | Invalid API token | Get new token from id.atlassian.com |
| Jira API 404 | Wrong project key | Verify in Jira URL (..projects/TEST) |
| Duplicate bugs | Manual rows added | Use separate sheet (full refresh expected) |
| Frozen rows issue | Used deleteRows() | Use clearContent() instead |
| Timeout errors | Too many modules | Split into multiple dashboards |

## Column References (by purpose)

**Identity:**
- A (Active), G (Spreadsheet ID) → Determines if module loads

**Metadata (auto-filled):**
- C (Project), D (Module), E (Submodule), F (PIC QA), K (QA Lead)

**Integration:**
- I (Jira Instance), J (Jira Project Key)
- L-N (Google Chat), O-P (Email), S-T (WhatsApp)
- U-W (VAPT), X-AK (Automation)

**Data Storage:**
- Columns Y-AK populated by `generateAutomationConfig()`

