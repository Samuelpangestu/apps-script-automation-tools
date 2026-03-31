# 🏗️ System Architecture - QA Test Management & Dashboard

**Last Updated:** March 31, 2026
**Version:** 2.0 (with VAPT Integration)

> Complete architectural documentation for QA Test Management System and Portfolio Dashboard. This document serves as the single source of truth for understanding system design, data flows, and integration patterns.

---

## 📋 Table of Contents

1. [System Overview](#system-overview)
2. [Architecture Diagram](#architecture-diagram)
3. [Component Breakdown](#component-breakdown)
4. [Data Flow](#data-flow)
5. [Blocker Calculation Logic](#blocker-calculation-logic)
6. [VAPT Workflow Integration](#vapt-workflow-integration)
7. [Key Files Reference](#key-files-reference)
8. [Database Schema](#database-schema)
9. [Integration Points](#integration-points)
10. [Common Patterns](#common-patterns)
11. [How to Extend](#how-to-extend)
12. [Troubleshooting](#troubleshooting)

---

## 🎯 System Overview

### Architecture Pattern
**Hub-and-Spoke Model**
- **Hub:** QA Portfolio Dashboard (centralized aggregator)
- **Spokes:** Individual QATM (QA Test Management) modules per project/module
- **External:** Jira (bug tracking system)

### Core Components

```
┌─────────────────────────────────────────────────────────────────┐
│                    QA PORTFOLIO DASHBOARD                        │
│  (1lHO8yKyqKs1_n5GV1m-SJMACLS95Jc7yy6dM_ItyT-l_-GdmkGQk3OIO)   │
│                                                                  │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐         │
│  │   Overview   │  │     Bugs     │  │    Smoke     │         │
│  │  (KPI Cards) │  │  (Tracking)  │  │  (Critical)  │         │
│  └──────────────┘  └──────────────┘  └──────────────┘         │
│                                                                  │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐         │
│  │   History    │  │   Coverage   │  │   Failure    │         │
│  │  (Timeline)  │  │  (Progress)  │  │  (Analysis)  │         │
│  └──────────────┘  └──────────────┘  └──────────────┘         │
│                                                                  │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │                    Config Sheet                           │  │
│  │  • Module registry (Active/Inactive)                     │  │
│  │  • Jira Sync configuration                               │  │
│  │  • Spreadsheet IDs for QATM modules                      │  │
│  └──────────────────────────────────────────────────────────┘  │
└────────────────────┬─────────────────────────────────────┬─────┘
                     │                                     │
                     │ refreshDashboard()                  │
                     │ pullModuleData()                    │
                     │                                     │
          ┌──────────▼──────────┐              ┌──────────▼──────────┐
          │  QATM Module 1      │              │  QATM Module N      │
          │  (Project A)        │              │  (Project Z)        │
          │                     │              │                     │
          │  • TC_Master        │              │  • TC_Master        │
          │  • TC_Execution     │              │  • TC_Execution     │
          │  • API_Master       │              │  • API_Master       │
          │  • API_Execution    │              │  • API_Execution    │
          │  • BugReport        │              │  • BugReport        │
          │  • Summary          │              │  • Summary          │
          │  • Appendix         │              │  • Appendix         │
          └──────────┬──────────┘              └──────────┬──────────┘
                     │                                     │
                     │ syncJiraBugs()                     │
                     │ updateStatus()                      │
                     │                                     │
          ┌──────────▼─────────────────────────────────────▼─────────┐
          │                      JIRA                                  │
          │  • Bug tracking                                           │
          │  • Status sync (Open → Fixed → Verified → VAPT → Closed) │
          │  • Priority management                                     │
          └───────────────────────────────────────────────────────────┘
```

---

## 📊 Architecture Diagram

### System Layers

```
┌─────────────────────────────────────────────────────────────────┐
│                      PRESENTATION LAYER                          │
├─────────────────────────────────────────────────────────────────┤
│  • Google Sheets UI (tabs, charts, conditional formatting)      │
│  • Web App UI (HTML/CSS/JS served via doGet/doPost)            │
│  • Custom menus (onOpen triggers)                               │
│  • Data validation dropdowns                                     │
└───────────────────────────────┬─────────────────────────────────┘
                                │
┌───────────────────────────────▼─────────────────────────────────┐
│                      APPLICATION LAYER                           │
├─────────────────────────────────────────────────────────────────┤
│  Dashboard Scripts:                                              │
│  • MasterDashboard.js  - Main orchestrator                      │
│  • BugsTab.js          - Bug aggregation & delta tracking       │
│  • JiraSync.js         - Jira integration                       │
│  • Notifications.js    - WhatsApp/Email alerts                  │
│  • VAPTBroadcast.js    - Bulk update utility                    │
│                                                                  │
│  QATM Scripts:                                                   │
│  • MasterQATCM.js      - Template generator                     │
│  • AddHeaderNotes.js   - Documentation helper                   │
└───────────────────────────────┬─────────────────────────────────┘
                                │
┌───────────────────────────────▼─────────────────────────────────┐
│                        DATA LAYER                                │
├─────────────────────────────────────────────────────────────────┤
│  • Google Sheets (structured data storage)                      │
│  • Script Properties (configuration, credentials)               │
│  • Cache Service (temporary data, 6 hour TTL)                   │
│  • External APIs (Jira REST API)                                │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🧩 Component Breakdown

### 1. QA Portfolio Dashboard

**Purpose:** Centralized monitoring and reporting for all QA activities across projects

**Core Files:**
- `MasterDashboard.js` - Main orchestrator, dashboard builder
- `BugsTab.js` - Bug metrics aggregation with historical tracking
- `JiraSync.js` - Jira integration and bug synchronization
- `Notifications.js` - Alert system (WhatsApp, Email, Google Chat)
- `VAPTBroadcast.js` - Bulk operations for VAPT workflow updates

**Key Functions:**

```javascript
// Main entry point - pull data from all QATM modules
refreshDashboard()
  ├─ getModuleList_(ss)              // Read Config sheet
  ├─ pullModuleData_(mod)            // Extract data from each QATM
  ├─ writeOverview(ss, allData)      // KPI cards
  ├─ writeBugs(ss, allData)          // Bug tracking with delta
  ├─ writeSmoke(ss, allData)         // Critical test results
  ├─ writeCoverage(ss, allData)      // Test coverage metrics
  └─ appendHistory(ss, allData)      // Historical timeline

// Jira integration
syncAndRefresh()
  ├─ syncAllJiraBugs()               // Fetch from Jira API
  ├─ updateBugReportSheets()         // Write to QATM BugReport
  └─ refreshDashboard()              // Update aggregated views
```

**Tabs Structure:**

| Tab | Purpose | Update Frequency |
|-----|---------|------------------|
| **Overview** | Portfolio-level KPIs | On refresh (manual) |
| **Bugs** | Bug metrics with delta tracking | On refresh |
| **Smoke** | Critical smoke test results | On refresh |
| **Coverage** | Test coverage progress | On refresh |
| **Failure Scenario** | Failed test analysis | On refresh |
| **History** | Timeline of metrics | Append-only |
| **Config** | Module registry | Manual edit |
| **Credentials** | Jira auth config | Manual edit |
| **_Raw** | Debug data dump | On refresh |

---

### 2. QATM (QA Test Management) Module

**Purpose:** Project-specific test case repository and execution tracking

**Core Files:**
- `MasterQATCM.js` - Template generator (creates new QATM)
- `AddHeaderNotes.js` - Adds documentation to headers

**Tabs Structure:**

| Tab | Purpose | Data Source |
|-----|---------|-------------|
| **Summary** | KPIs and metrics dashboard | Calculated from execution sheets |
| **TC_Master** | Web/Mobile test case repository | Manual entry |
| **TC_Execution** | Web/Mobile test execution log | Manual entry |
| **API_Master** | API test case repository | Manual entry |
| **API_Execution** | API test execution log | Manual entry |
| **BugReport** | Bug tracking with VAPT workflow | Jira sync + manual |
| **PerfTest** | Performance test results (K6/JMeter) | Manual entry |
| **Appendix** | Documentation and workflow guide | Auto-generated |

**Summary Metrics:**

```
Web Testing Section:
├─ Total Test Cases
├─ Passed / Failed / Blocked / In Progress / TODO
├─ Pass Rate % (target: ≥95%)
├─ Automation Rate %
├─ Execution Rate %
└─ Open Blocker (Critical metric!)

API Testing Section:
├─ [Same structure as Web]

Smoke Testing:
├─ Web Smoke (subset of critical tests)
├─ API Smoke
└─ Open Blocker Smoke (release blocker)

Bugs Section:
├─ Open Blocker (Medium-Critical, not Closed)
├─ Prod Bugs (bugs in Production environment)
```

---

## 🔄 Data Flow

### Flow 1: Dashboard Refresh (Pull Model)

```
USER TRIGGER
   │
   ▼
refreshDashboard()
   │
   ├─► getModuleList_(ss) ─────► Read Config sheet
   │                              ├─ Active = TRUE
   │                              └─ Spreadsheet ID exists
   ▼
FOR EACH MODULE:
   │
   ├─► SpreadsheetApp.openById(moduleId)
   │
   ├─► pullModuleData_(mod)
   │    ├─ Read Summary sheet (KPIs)
   │    ├─ Read BugReport sheet (bug counts)
   │    ├─ Read TC_Execution (test results)
   │    ├─ Read API_Execution (API results)
   │    └─ Calculate metrics
   │
   └─► Return moduleData object

AGGREGATE:
   │
   ├─► writeOverview(ss, allData)
   │    └─ Show: Total modules, Total tests, Pass rate, Blockers
   │
   ├─► writeBugs(ss, allData)
   │    ├─ Aggregate bugs by module
   │    ├─ Calculate delta vs previous refresh
   │    └─ Color coding: Green=down, Red=up
   │
   ├─► writeSmoke(ss, allData)
   │    └─ Critical smoke tests only
   │
   ├─► writeCoverage(ss, allData)
   │    └─ Test coverage by module/feature
   │
   └─► appendHistory(ss, allData)
        └─ Append row with timestamp + metrics
```

### Flow 2: Jira Bug Sync (Push Model)

```
TRIGGER: Time-based or Manual
   │
   ▼
syncAllJiraBugs()
   │
   ├─► Read Dashboard Config sheet
   │    └─ Get modules with Jira Sync = TRUE
   │
   ▼
FOR EACH MODULE:
   │
   ├─► Read Credentials sheet
   │    ├─ Jira Instance (e.g., digitalperuri.atlassian.net)
   │    ├─ Jira Project Key (e.g., TEST)
   │    └─ API Token
   │
   ├─► Fetch bugs from Jira API
   │    └─ GET /rest/api/3/search?jql=project=TEST
   │
   ├─► Parse Jira response
   │    ├─ Bug ID (e.g., TEST-123)
   │    ├─ Summary (title)
   │    ├─ Priority (Critical/High/Medium/Low)
   │    ├─ Status (Open/In Progress/Fixed/Verified/VAPT/Closed)
   │    ├─ Environment (Dev/UAT/Production)
   │    └─ Link (Jira ticket URL)
   │
   ├─► Open QATM BugReport sheet
   │
   └─► Write or Update bugs
        ├─ Match by Bug ID (column F)
        ├─ Update: Status, Priority
        └─ Add new bugs if not exists
```

### Flow 3: VAPT Workflow Broadcast

```
TRIGGER: Manual run broadcastVAPTStatusUpdate()
   │
   ▼
Read Dashboard Config
   │
   ├─► Get all ACTIVE modules
   │
   ▼
FOR EACH MODULE:
   │
   ├─► Open QATM spreadsheet
   │
   ├─► Update BugReport sheet
   │    ├─ Status dropdown: Add "In Progress VAPT" & "Done VAPT"
   │    ├─ Status note (D4): Add VAPT workflow documentation
   │    ├─ Header row: Update with VAPT statuses
   │    └─ Conditional formatting: Add VAPT colors
   │
   ├─► Update Summary sheet
   │    ├─ Open Blocker formula: Include VAPT statuses
   │    └─ Note: Update blocker documentation
   │
   └─► Recreate Appendix
        └─ Section 7: Add VAPT workflow and roles
```

---

## 🚨 Blocker Calculation Logic

### Critical Rule (Post-VAPT Integration)

**BLOCKER STATUS:**
```javascript
// All statuses BEFORE Closed are blockers (belum release ke production)
const BLOCKER_STATUSES = [
  'Open',
  'In Progress',
  'Reopen',
  'Fixed',           // ✅ BLOCKER (Dev claimed fixed, waiting QA)
  'Verified',        // ✅ BLOCKER (QA verified, waiting VAPT or release)
  'In Progress VAPT',// ✅ BLOCKER (Security testing in progress)
  'Done VAPT'        // ✅ BLOCKER (VAPT done, waiting final QA sign-off)
];

// Only these are NOT blockers
const NOT_BLOCKER = [
  'Closed',          // ❌ NOT BLOCKER (released to production)
  "Won't Fix"        // ❌ NOT BLOCKER (rejected, won't be fixed)
];
```

**Priority Filter:**
```javascript
const BLOCKER_PRIORITIES = [
  'Critical',   // Showstopper
  'High',       // Blocker
  'Medium'      // Degraded (still blocker)
];

// Low and Lowest are NOT blockers (minor issues)
```

**Formula (Google Sheets):**

```javascript
// QATM Summary - Open Blocker calculation
=SUMPRODUCT(
  (ISNUMBER(MATCH(
    BugReport!D5:D2000,
    {"Open","In Progress","Reopen","Fixed","Verified","In Progress VAPT","Done VAPT"},
    0
  )))*
  (ISNUMBER(MATCH(
    BugReport!C5:C2000,
    {"Critical","High","Medium"},
    0
  )))
)
```

**JavaScript (JiraSync.js):**

```javascript
// Line 1937 - JiraSync blocker detection
const BLOCKER_STATUS = [
  'open',
  'in progress',
  'reopen',
  'fixed',           // ← Fixed included
  'verified',        // ← Verified included
  'in progress vapt',
  'done vapt'
];

// Usage
if (BLOCKER_STATUS.includes(status.toLowerCase()) &&
    ['Critical', 'High', 'Medium'].includes(priority)) {
  // This is a blocker!
}
```

**Dashboard BugsTab.js:**

```javascript
// Line 362-366 - Status-agnostic blocker detection
if (status !== 'Closed' && status !== "Won't Fix" &&
    ['Critical','Highest','High','Medium'].includes(priority)) {
  stats.blocker++;
}
```

### Why This Logic?

**Rationale:**
- Bug is only NOT a blocker after it's **Closed** (released to production)
- All intermediate statuses (Fixed, Verified, Done VAPT) are still **blockers** because:
  - **Fixed:** Waiting QA re-test (could fail)
  - **Verified:** Waiting VAPT or final sign-off (not released yet)
  - **Done VAPT:** Waiting final QA confirmation before Closed

**Target:** `0 Open Blocker` before production release

---

## 🔒 VAPT Workflow Integration

### Bug Status Lifecycle (with VAPT)

```
┌─────────────────────────────────────────────────────────────────┐
│                      BUG LIFECYCLE                               │
└─────────────────────────────────────────────────────────────────┘

QA PHASE:
   Open ────────► In Progress ────────► Fixed ────────► Verified
    │                  │                   │                │
    │                  │                   │                │
    └──────────────────┴───────────────────┴────────────────┘
                             │
                             │ (Reopen if bug reappears)
                             ▼
                          Reopen ─────────┐
                                          │
                                          │
VAPT PHASE (Security Testing):           │
                                          │
   Verified ────► In Progress VAPT ────► Done VAPT ────► Closed
       │               │                      │              │
       │               │                      │              │
       │               └──────────────────────┘              │
       │               (Reopen if VAPT finds issues)         │
       │                                                     │
       └─────────────────────────────────────────────────────┘
       (Skip VAPT - direct to Closed for non-security bugs)

EXCEPTION FLOWS:
   Any Status ────► Won't Fix (rejected with reason + Lead approval)
```

### Role Responsibilities

| Role | Can Update To |
|------|--------------|
| **Dev** | In Progress (start work)<br>Fixed (claim fixed - Dev CANNOT go directly to Verified/Closed) |
| **QA** | Verified (re-test passed, ready for VAPT or release)<br>Reopen (bug still exists)<br>Closed (final sign-off after Done VAPT) |
| **Security** | In Progress VAPT (start security testing)<br>Done VAPT (VAPT complete, ready for QA final check)<br>Reopen (if new security issues found) |
| **Lead** | Won't Fix (business/technical decision with comment)<br>Closed (final authority) |

### Status Colors (Conditional Formatting)

```javascript
// BugReport sheet conditional formatting
const STATUS_COLORS = {
  'Open':             { bg: '#FFCDD2', fg: '#C62828' }, // Red
  'In Progress':      { bg: '#E3F2FD', fg: '#1565C0' }, // Blue
  'Fixed':            { bg: '#FFF9C4', fg: '#F57F17' }, // Yellow
  'Verified':         { bg: '#C8E6C9', fg: '#2E7D32' }, // Light green
  'In Progress VAPT': { bg: '#E1F5FE', fg: '#01579B' }, // Light blue
  'Done VAPT':        { bg: '#B2DFDB', fg: '#004D40' }, // Teal
  'Closed':           { bg: '#E8F5E9', fg: '#388E3C' }, // Green
  "Won't Fix":        { bg: '#F5F5F5', fg: '#9E9E9E' }, // Gray
  'Reopen':           { bg: '#EDE7F6', fg: '#6A1B9A' }  // Purple
};
```

---

## 📁 Key Files Reference

### Dashboard (`projects/qa-dashboard/src/`)

| File | Lines | Purpose | Key Functions |
|------|-------|---------|---------------|
| **MasterDashboard.js** | 2000+ | Main orchestrator | `refreshDashboard()`, `pullModuleData_()`, `getModuleList_()` |
| **BugsTab.js** | 500 | Bug aggregation with delta | `writeBugs()`, `collectBugsFromModules_()` |
| **JiraSync.js** | 2000+ | Jira integration | `syncAllJiraBugs()`, `_getBlockerData_()` |
| **Notifications.js** | 1500 | Alert system | `sendWhatsAppNotification()`, `sendEmailReport()` |
| **VAPTBroadcast.js** | 900 | Bulk update utility | `broadcastVAPTStatusUpdate()`, `broadcastFixNoteAndAppendix()` |
| **WebApp.html** | 31KB | Web dashboard UI | HTML/CSS/JS frontend |
| **WebAppBackend.js** | 14KB | Web API backend | `doGet()`, `getModuleData()` |
| **BroadcastFixes.js** | 500 | Legacy broadcast utils | (Deprecated) |

### Template (`projects/qa-test-management/src/`)

| File | Lines | Purpose | Key Functions |
|------|-------|---------|---------------|
| **MasterQATCM.js** | 2500+ | QATM template generator | `createNewQATCM()`, `buildSummary()`, `buildBugReport()` |
| **AddHeaderNotes.js** | 300 | Documentation helper | `addNotesToQATCM()` |

---

## 🗄️ Database Schema

### Dashboard Config Sheet

| Column | Name | Type | Description |
|--------|------|------|-------------|
| A | Active | BOOLEAN | TRUE = pull data on refresh |
| B | Jira Sync | BOOLEAN | TRUE = sync bugs from Jira |
| C | Project | STRING | Project name (from QATM Summary B2) |
| D | Modul | STRING | Module name (from QATM Summary B3) |
| E | Submodul | STRING | Submodule name (from QATM Summary B4) |
| F | PIC QA | STRING | QA Engineer (from QATM Summary B6) |
| G | Spreadsheet ID | STRING | QATM Google Sheets ID |
| H | Link | FORMULA | `=HYPERLINK("https://docs.google.com/spreadsheets/d/"&G2, "Open")` |
| I | Jira Instance | STRING | e.g., "digitalperuri" |
| J | Jira Project | STRING | e.g., "TEST" |

### QATM BugReport Sheet

| Column | Name | Type | Description |
|--------|------|------|-------------|
| A | No | NUMBER | Auto-increment |
| B | Type | STRING | Web / Mobile / API |
| C | Priority | STRING | Critical / High / Medium / Low / Lowest |
| D | Status | STRING | Open / In Progress / Fixed / Verified / In Progress VAPT / Done VAPT / Closed / Won't Fix / Reopen |
| E | Feature | STRING | Feature name where bug found |
| F | Bug ID | STRING | Jira ticket ID (e.g., TEST-123) |
| G | Title | STRING | Bug title/summary |
| H | Link | STRING | URL to Jira ticket or test evidence |
| I | Environment | STRING | Dev / Staging / UAT / Production |
| J | Reporter | STRING | Person who found the bug |
| K | Comment | STRING | Additional notes |

### Dashboard Bugs Tab

| Column | Name | Type | Description |
|--------|------|------|-------------|
| A | Project | STRING | Project name |
| B | Modul | STRING | Module name |
| C | Submodul | STRING | Submodule name |
| D | Total | NUMBER | All bugs (not Closed) |
| E | Critical | NUMBER | Critical priority count |
| F | High | NUMBER | High priority count |
| G | Medium | NUMBER | Medium priority count |
| H | Low | NUMBER | Low priority count |
| I | Lowest | NUMBER | Lowest priority count |
| J | Blocker | NUMBER | Medium-Critical bugs, not Closed/Won't Fix |
| K | Δ Total | NUMBER | Delta vs previous refresh |
| L | Δ Blocker | NUMBER | Delta blocker vs previous |
| M | Dev | NUMBER | Bugs in Dev environment |
| N | UAT | NUMBER | Bugs in UAT environment |
| O | Prod | NUMBER | Bugs in Production (CRITICAL!) |
| P | Prev Total | NUMBER | Previous total (for delta calc) |
| Q | Last Updated | TIMESTAMP | Last refresh time |

---

## 🔌 Integration Points

### 1. Jira REST API

**Endpoint:** `https://{instance}.atlassian.net/rest/api/3/search`

**Authentication:** Basic Auth (email + API token)

**Query:**
```javascript
const jql = `project=${projectKey} AND type=Bug`;
const url = `https://${instance}.atlassian.net/rest/api/3/search?jql=${jql}&maxResults=100&fields=summary,status,priority,customfield_xxxxx`;
```

**Response Mapping:**
```javascript
{
  "issues": [
    {
      "key": "TEST-123",                    // → Bug ID (column F)
      "fields": {
        "summary": "Login button broken", // → Title (column G)
        "status": {
          "name": "In Progress"           // → Status (column D, lowercase)
        },
        "priority": {
          "name": "High"                  // → Priority (column C)
        },
        "customfield_10050": {            // Custom field for Environment
          "value": "Production"           // → Environment (column I)
        }
      }
    }
  ]
}
```

**Sync Frequency:**
- Time-based trigger: Every 1 hour
- Manual: Run `syncAllJiraBugs()` from Dashboard

---

### 2. WhatsApp Notifications (via Fonnte API)

**Endpoint:** `https://api.fonnte.com/send`

**Authentication:** API Key in header

**Payload:**
```javascript
{
  "target": "6281234567890",  // Phone number
  "message": "🚨 *BLOCKER ALERT*\n\n...",
  "countryCode": "62"
}
```

**Trigger Conditions:**
- `totalBlockers > 0` (any blocker exists)
- Manual: Run notification function

---

## 🎨 Common Patterns

### Pattern 1: Safe Sheet Access

```javascript
// Always check if sheet exists before reading
function safeGetSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`⚠️ Sheet "${sheetName}" not found`);
    return null;
  }
  return sheet;
}
```

### Pattern 2: Batch Operations

```javascript
// Read/write in batches for performance
const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();  // Batch read

// Process data...

sheet.getRange(1, 1, data.length, data[0].length).setValues(data);  // Batch write
```

### Pattern 3: Config-Driven Modules

```javascript
// Use Config sheet to drive which modules to process
const modules = getModuleList_(ss);
modules.forEach(mod => {
  if (!mod.active) return;  // Skip inactive
  processModule(mod);
});
```

### Pattern 4: Idempotent Operations

```javascript
// Make operations repeatable without side effects
function updateOrCreate(sheet, identifier, data) {
  const existingRow = findRow(sheet, identifier);
  if (existingRow) {
    // Update
    sheet.getRange(existingRow, 1, 1, data.length).setValues([data]);
  } else {
    // Create
    sheet.appendRow(data);
  }
}
```

---

## 🚀 How to Extend

### Add a New Bug Status

**Example:** Adding "In Review" status

1. **Update QATM Template** (`MasterQATCM.js`):
```javascript
// Line ~2008: Status dropdown
ws.getRange(DS,4,MR,1).setDataValidation(
  dv_(['Open','In Progress','Fixed','In Review','Verified',...])
);

// Line ~2036: Conditional formatting
{v:'In Review', bg:'#F3E5F5', fg:'#7B1FA2'}
```

2. **Update Blocker Formula** (if status is a blocker):
```javascript
// Line ~1516: BLOCKER_FORMULA
'(ISNUMBER(MATCH(BugReport!D5:D2000,{"Open","In Progress","Reopen","Fixed","In Review","Verified",...},0)))*'
```

3. **Update JiraSync** (`JiraSync.js`):
```javascript
// Line 1937: BLOCKER_STATUS
const BLOCKER_STATUS = ['open', 'in progress', 'reopen', 'fixed', 'in review', 'verified', ...];

// Line 909-910: Status colors
const SBG = {..., 'In Review':'#F3E5F5'};
const SFG = {..., 'In Review':'#7B1FA2'};
```

4. **Broadcast Update** (`VAPTBroadcast.js`):
```javascript
// Line ~95: Status dropdown in broadcast
.requireValueInList(['Open','In Progress','Fixed','In Review','Verified',...], true)

// Line ~177: Formula update
'{"Open","In Progress","Reopen","Fixed","In Review","Verified",...}'

// Line ~104-119: Status note
'🚨 BLOCKER STATUS:\nOpen, In Progress, Reopen, Fixed, In Review, Verified, ...'
```

5. **Deploy:**
   - Push Dashboard to Testing/Production
   - Run `broadcastFixNoteAndAppendix()` to update existing QATMs

---

### Add a New Dashboard Tab

**Example:** Adding "Performance" tab

1. **Create Builder Function** (`projects/qa-dashboard/src/NewTab.js`):
```javascript
function buildPerformance(ss) {
  const ws = ss.insertSheet('Performance', 2);
  ws.setTabColor('#4A148C');

  // Headers
  ws.getRange(1,1).setValue('Performance Test Results');
  ws.getRange(2,1,1,5).setValues([['Project','Module','RPS','Response Time','Result']]);

  // Styling...
}

function writePerformance(ss, allData) {
  let ws = ss.getSheetByName('Performance');
  if (!ws) { buildPerformance(ss); ws = ss.getSheetByName('Performance'); }

  // Clear old data
  ws.getRange(3, 1, ws.getMaxRows()-2, 5).clearContent();

  // Write new data
  const perfData = allData.map(mod => [
    mod.project,
    mod.module,
    mod.perfRPS || '—',
    mod.perfResponseTime || '—',
    mod.perfResult || 'N/A'
  ]);

  ws.getRange(3, 1, perfData.length, 5).setValues(perfData);
}
```

2. **Integrate into Refresh** (`MasterDashboard.js`):
```javascript
// Line ~320: Add to refreshDashboard()
writePerformance(ss, allData);
```

3. **Update Step Creator** (`MasterDashboard.js`):
```javascript
// Line ~250: Add to step3_createDataTabs()
Logger.log('Creating Performance...');
buildPerformance(ss);
Logger.log('✅ Performance created');
```

---

## 🔧 Troubleshooting

### Issue 1: "Requested entity was not found"

**Symptom:** Deployment fails with script ID error

**Cause:** Invalid or deleted Apps Script project ID

**Fix:**
```bash
# Check .clasp.json
cat .clasp.json

# Verify script ID in Google Apps Script console
# If wrong, update .clasp.json with correct ID
```

---

### Issue 2: "Config kosong! Isi tab Config dulu"

**Symptom:** Broadcast or refresh fails with "no modules found"

**Cause:**
- Config sheet missing
- No rows with Active = TRUE
- Spreadsheet ID column empty or too short

**Fix:**
```javascript
// Check Config structure
// Row 1: Title
// Row 2: Subtitle
// Row 3: Headers
// Row 4+: Data

// Verify:
// Column A (Active) = TRUE (boolean, not string)
// Column G (Spreadsheet ID) = valid ID (44+ chars)
```

---

### Issue 3: Blocker Count Mismatch

**Symptom:** Dashboard blocker count ≠ QATM Summary blocker count

**Cause:** Inconsistent blocker formula or status list

**Debug:**
```javascript
// Check QATM Summary formula (Web/API Open Blocker cells)
// Should include: Open, In Progress, Reopen, Fixed, Verified, In Progress VAPT, Done VAPT

// Check JiraSync.js BLOCKER_STATUS (line 1937)
// Should match formula statuses (lowercase)

// Check Dashboard BugsTab.js (line 362-366)
// Should exclude only Closed and Won't Fix
```

---

### Issue 4: Jira Sync Not Working

**Symptom:** Bugs not syncing from Jira

**Debug Checklist:**
1. ✅ Config sheet: Jira Sync = TRUE for module
2. ✅ Credentials sheet: API token valid (not expired)
3. ✅ Jira Instance: Correct (e.g., "digitalperuri" not full URL)
4. ✅ Jira Project: Correct project key (e.g., "TEST")
5. ✅ Network: API quota not exceeded
6. ✅ Permissions: Service account has Jira access

**Test:**
```javascript
// Run from Apps Script editor
function testJiraConnection() {
  const instance = 'digitalperuri';
  const project = 'TEST';
  const email = 'your-email@domain.com';
  const token = 'your-api-token';

  const url = `https://${instance}.atlassian.net/rest/api/3/search?jql=project=${project}&maxResults=1`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(email + ':' + token),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}
```

---

## 📚 Additional Resources

- **Clasp Workflow:** See `CLASP_WORKFLOW.md`
- **Environment Setup:** See `ENVIRONMENTS.md`
- **Project Commands:** See `.claude/commands/qa-help.md`
- **Google Apps Script Docs:** https://developers.google.com/apps-script
- **Jira REST API:** https://developer.atlassian.com/cloud/jira/platform/rest/v3/

---

## 📝 Version History

| Version | Date | Changes |
|---------|------|---------|
| **2.0** | 2026-03-31 | Added VAPT workflow integration |
| **1.5** | 2026-03-17 | Added dual environment support (Testing + Production) |
| **1.0** | 2026-02-27 | Initial architecture documentation |

---

**Maintained by:** QA Team @ INA Digital
**Last Review:** March 31, 2026
