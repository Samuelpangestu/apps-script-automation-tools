# Codebase Exploration Index

This index provides a roadmap to the comprehensive codebase exploration that has been completed for the QA Test Management Template.

## Documentation Files Created

### 1. **CODEBASE_ANALYSIS.md** (22KB)
**Comprehensive technical deep-dive documentation**

Complete reference for architecture, implementation patterns, and data structures.

**Sections:**
1. Dashboard Sheet Structure
2. Config Tab Structure (Detailed Reference)
3. Existing Jira Integration
4. External API Integration Patterns
5. Data Structures & Key Objects
6. Duplicate Prevention & Sync Tracking
7. Refresh Flow & Timing
8. Critical Implementation Details
9. File Reference Guide
10. Summary: Duplicate Prevention Strategy
11. Next Steps for Implementation

**Best for:** Understanding the complete system, detailed implementation decisions, code patterns, data flow

---

### 2. **QUICK_REFERENCE.md** (6.6KB)
**Quick-lookup guide and checklists**

Condensed reference for common tasks, troubleshooting, and quick lookups.

**Contents:**
- File locations
- Config sheet column map
- Key functions and when to use them
- Module determination logic
- Dashboard sheet list
- Jira sync process (3 steps)
- Duplicate prevention strategy
- API integration patterns
- Error handling best practices
- Performance tips
- Testing checklist
- Common issues & solutions
- Column references by purpose

**Best for:** Quick lookups, testing, troubleshooting, development reference

---

## Quick Links to Key Information

### Architecture Overview
- **Dashboard Structure:** See CODEBASE_ANALYSIS.md section 1
- **Config Organization:** See CODEBASE_ANALYSIS.md section 2 or QUICK_REFERENCE.md "Config Sheet Column Map"
- **File Reference:** See CODEBASE_ANALYSIS.md section 9 or QUICK_REFERENCE.md "File Locations"

### Jira Integration
- **Complete Details:** See CODEBASE_ANALYSIS.md section 3 (5,000+ words)
- **Quick Setup:** See QUICK_REFERENCE.md "Jira Integration" section
- **Troubleshooting:** See QUICK_REFERENCE.md "Common Issues & Solutions"

### API Patterns
- **All Integrations:** See CODEBASE_ANALYSIS.md section 4
- **Code Examples:** Each API has implementation examples
- **Error Handling:** See QUICK_REFERENCE.md "Error Handling Best Practices"

### Data Structures
- **Module Object:** See CODEBASE_ANALYSIS.md section 5
- **Bug Report Schema:** See CODEBASE_ANALYSIS.md section 5 (BC_ column definitions)
- **Blocker Data:** See CODEBASE_ANALYSIS.md section 5

### Duplicate Prevention
- **Strategy:** See CODEBASE_ANALYSIS.md section 6
- **Implementation:** See QUICK_REFERENCE.md "Duplicate Prevention"
- **Recommendations:** See CODEBASE_ANALYSIS.md section 10

### Refresh & Sync Flows
- **Complete Flow:** See CODEBASE_ANALYSIS.md section 7
- **Quick Summary:** See QUICK_REFERENCE.md "Jira Sync Process"
- **Triggers:** See QUICK_REFERENCE.md "Key Functions & When to Use"

---

## Key Findings Summary

### 1. Dashboard Structure
- 12 automatically-created sheets
- Config sheet is the system of record (SSOT)
- 27+ configuration columns for integrations
- Active modules determined by Column A (Active) + Column G (valid Spreadsheet ID)

### 2. Config Sheet Organization
**File:** `/projects/qa-dashboard/src/MasterDashboard.js` (line 1328+, buildConfig function)

| Column | Range | Purpose |
|--------|-------|---------|
| A-K | Core | Active flag, Jira sync, project metadata |
| L-N | Google Chat | Webhook URL, Schedule, Enable |
| O-P | Email | Recipients, Enable |
| Q-R | Auto-Refresh | Interval, Enable |
| S-T | WhatsApp | Group ID (global), Fonnte Token |
| U-W | VAPT | Enable, Spreadsheet ID |
| X-AK | Automation | Tab Name, Webhook, Job Patterns (auto-generated) |

### 3. Jira Integration
**File:** `/projects/qa-dashboard/src/JiraSync.js` (2,147 lines)

**Architecture:**
- Two Jira instances: digitalperuri, bgn-peruri
- Basic Auth (email + API token)
- Custom field mappings per instance
- JQL-based filtering by module
- Full data replacement strategy (no incremental sync)

**Key Functions:**
- `jiraSetup()` - Initialize Jira integration
- `jiraActivate()` - Activate with triggers
- `syncJiraToAllSheets()` - Manual sync (1-hour trigger)
- `_syncMod_()` - Sync single module
- `_fetch_()` - Fetch from Jira API

**Triggers:**
- Every 1 hour: `syncJiraToAllSheets()`
- Daily 07:00: `sendBugNotification()`
- Daily 23:00: `syncJiraStatusAll()`

### 4. API Integrations

| API | Method | Authentication | Key Patterns |
|-----|--------|----------------|--------------|
| Jira Cloud | UrlFetchApp | Basic Auth | JQL pagination with nextPageToken |
| Google Chat | UrlFetchApp | Webhook URL | Card format JSON |
| Email | MailApp | Built-in | HTML tables, color-coded |
| WhatsApp | UrlFetchApp | Fonnte token | Reference only (not fully implemented) |
| Automation | Webhook | Custom | Tab Sheet Name identifier |

### 5. Duplicate Prevention
**Strategy:** FULL DATA REPLACEMENT

```
Every sync:
1. Clear all data rows from BugReport
2. Fetch fresh bugs from Jira
3. Insert all (Closed/Won't Fix filtered)
Result: No duplicates possible
Trade-off: Manual edits are lost
```

### 6. Data Structures

**Module Object (returned by getModuleList_):**
```
{
  name, id, project, module, submodule, team,
  active, jiraSync, jiraInst, jiraProj,
  externalQA { isExternal, retestByPeruri, notes },
  automationContracts { all, web, api }
}
```

**Bug Report Columns (23 total):**
```
A: BUG_ID (hyperlinked)
B-U: Bug details (priority, status, environment, etc.)
V: SCREENSHOT URLs
W: JIRA_KEY (raw identifier)
X: SYNCED (timestamp)
```

---

## How to Use These Documents

### For Understanding Architecture
1. Start with QUICK_REFERENCE.md "Dashboard Sheets" section
2. Read CODEBASE_ANALYSIS.md section 1 for structure
3. Read CODEBASE_ANALYSIS.md section 2 for Config details

### For Implementing New Features
1. Check CODEBASE_ANALYSIS.md section 11 "Next Steps"
2. Review relevant API patterns (section 4)
3. Follow existing patterns from QUICK_REFERENCE.md

### For Troubleshooting
1. Check QUICK_REFERENCE.md "Common Issues & Solutions"
2. Run tests from QUICK_REFERENCE.md "Testing Checklist"
3. Use debugging tips from CODEBASE_ANALYSIS.md section 8

### For Integration Work
1. Review API patterns (CODEBASE_ANALYSIS.md section 4)
2. Check data structures (CODEBASE_ANALYSIS.md section 5)
3. Follow error handling patterns (QUICK_REFERENCE.md)

### For Performance Optimization
1. See QUICK_REFERENCE.md "Performance Tips"
2. Review trigger setup (QUICK_REFERENCE.md "Jira Integration")
3. Check implementation details (CODEBASE_ANALYSIS.md section 8)

---

## File Locations Reference

### Core Dashboard Files
**Location:** `/projects/qa-dashboard/src/`

| File | Lines | Key Functions |
|------|-------|---------------|
| MasterDashboard.js | 3,105 | createDashboard, refreshDashboard, getModuleList_ |
| JiraSync.js | 2,147 | jiraSetup, jiraActivate, syncJiraToAllSheets, _syncMod_, _fetch_ |
| Notifications.js | ~1,500 | sendBlockerNotification, fetchVAPTDataForProject_ |
| AutomationConfigGenerator.js | 211 | generateAutomationConfig, generateConfigForModule_ |
| WebAppBackend.js | ~800 | Web UI backend |
| VAPTDataFetch.js | ~500 | VAPT aggregation |
| Other | Various | VAPTTab.js, BugsTab.js, InitVAPTTabs.js, etc. |

### Configuration Files
```
appsscript.json     - Apps Script manifest
clasp.json          - Clasp deployment config
.env                - Environment variables (local dev)
```

---

## Important Implementation Notes

### Active Module Determination
```javascript
// From getModuleList_() line 822
if (!active || !id || id.length < 10 || id === 'PASTE_SPREADSHEET_ID_HERE') continue;
```

A module is ACTIVE when:
1. Column A (Active) = TRUE
2. Column G (Spreadsheet ID) not empty
3. Spreadsheet ID length >= 10 chars
4. Spreadsheet ID != "PASTE_SPREADSHEET_ID_HERE"

### Jira Sync Process
```
Setup:    jiraSetup()  → Adds columns + Credentials sheet
Activate: jiraActivate() → Tests connection + creates triggers
Execute:  _syncMod_()  → Clears rows, fetches, inserts bugs
```

### Error Handling Pattern
```javascript
// Used throughout codebase
try {
  const resp = UrlFetchApp.fetch(url, {
    headers: { /* auth */ },
    muteHttpExceptions: true  // CRITICAL: prevent crashes
  });
  if (resp.getResponseCode() !== 200) {
    Logger.log('Error: ' + resp.getResponseCode());
    return null;  // Don't throw
  }
  return JSON.parse(resp.getContentText());
} catch (e) {
  Logger.log('Exception: ' + e.message);
  return null;  // Graceful fallback
}
```

---

## Next Steps

1. **Review CODEBASE_ANALYSIS.md** for comprehensive understanding (22KB, all 11 sections)
2. **Keep QUICK_REFERENCE.md handy** for daily development (print or bookmark)
3. **Use specific sections** as needed for implementing features
4. **Reference "Next Steps for Implementation"** in CODEBASE_ANALYSIS.md section 11

---

**Documents Created:** June 17, 2026
**Exploration Scope:** Comprehensive (Dashboard structure, Config, Jira, APIs, Data structures, Sync flows, Duplicates)
**Total Documentation:** 28.6KB across 2 files
