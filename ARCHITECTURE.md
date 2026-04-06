# 🏗️ System Architecture - QA Test Management & Dashboard

**Last Updated:** April 6, 2026
**Version:** 2.1 (VAPT External Source Integration)

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
┌──────────────────────────────────────────────────────────────────┐
│                     QA PORTFOLIO DASHBOARD                        │
│   (1b2RBemEgo5B0YfUJHqAw8D0dH9Pg2Avgcngb7iz1PxY)                │
│                                                                   │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐          │
│  │   Overview   │  │     Bugs     │  │     VAPT     │          │
│  │  (KPI Cards) │  │  (Tracking)  │  │  (Security)  │          │
│  └──────────────┘  └──────────────┘  └──────────────┘          │
│                                                                   │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐          │
│  │    Smoke     │  │   History    │  │   Coverage   │          │
│  │  (Critical)  │  │  (Timeline)  │  │  (Progress)  │          │
│  └──────────────┘  └──────────────┘  └──────────────┘          │
│                                                                   │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │                     Config Sheet                          │   │
│  │  • Module registry (Active/Inactive)                      │   │
│  │  • Jira Sync configuration                                │   │
│  │  • Spreadsheet IDs for QATM modules                       │   │
│  │  • Notification config (WhatsApp, Email, Google Chat)     │   │
│  └──────────────────────────────────────────────────────────┘   │
└────────┬──────────────────────────────────┬───────────────┬──────┘
         │                                  │               │
         │ refreshDashboard()               │               │
         │ pullModuleData()                 │               │
         │                                  │               │ refreshVAPTData()
         │                                  │               │ fetchAndProcessVAPTData_()
         │                                  │               │
┌────────▼─────────┐            ┌──────────▼──────────┐   │
│  QATM Module 1   │            │  QATM Module N      │   │
│  (Project A)     │            │  (Project Z)        │   │
│                  │            │                     │   │
│ • TC_Master      │            │ • TC_Master         │   │
│ • TC_Execution   │            │ • TC_Execution      │   │
│ • API_Master     │            │ • API_Master        │   │
│ • API_Execution  │            │ • API_Execution     │   │
│ • BugReport      │            │ • BugReport         │   │
│ • Summary        │            │ • Summary           │   │
│ • Appendix       │            │ • Appendix          │   │
└────────┬─────────┘            └──────────┬──────────┘   │
         │                                 │               │
         │ syncJiraBugs()                  │               │
         │ updateStatus()                  │               │
         │                                 │               │
┌────────▼─────────────────────────────────▼─────────┐    │
│                      JIRA                           │    │
│  • Bug tracking                                     │    │
│  • Status sync (Open → Fixed → Verified → Closed)  │    │
│  • Priority management                              │    │
└─────────────────────────────────────────────────────┘    │
                                                            │
                        ┌───────────────────────────────────▼────────┐
                        │      VAPT SOURCE SPREADSHEET              │
                        │  (17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBM...) │
                        │                                            │
                        │  • Ad Hoc VAPT (security findings)        │
                        │  • Regular VAPT (scheduled assessments)   │
                        │  • Severity: Critical, High, Medium, Low  │
                        │  • Status: Open, Ready to Retest, Closed  │
                        └────────────────────────────────────────────┘
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
│  • VAPTTab.js          - VAPT findings dashboard builder        │
│  • VAPTDataFetch.js    - Fetch/process VAPT data               │
│  • JiraSync.js         - Jira integration                       │
│  • Notifications.js    - WhatsApp/Email/GChat alerts            │
│  • VAPTBroadcast.js    - Bulk update utility                    │
│  • WebAppBackend.js    - Web dashboard API                      │
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
- `VAPTTab.js` - VAPT findings dashboard builder
- `VAPTDataFetch.js` - Fetch and process VAPT data from external source
- `JiraSync.js` - Jira integration and bug synchronization
- `Notifications.js` - Alert system (WhatsApp, Email, Google Chat)
- `VAPTBroadcast.js` - Bulk operations for VAPT workflow updates
- `WebAppBackend.js` - Web dashboard API backend

**Key Functions:**

```javascript
// Main entry point - pull data from all QATM modules
refreshDashboard()
  ├─ getModuleList_(ss)              // Read Config sheet
  ├─ pullModuleData_(mod)            // Extract data from each QATM
  ├─ refreshVAPTData()               // NEW: Fetch VAPT data from external source
  ├─ writeOverview(ss, allData)      // KPI cards
  ├─ writeBugs(ss, allData)          // Bug tracking with delta
  ├─ writeVAPT(ss, vaptData)         // NEW: VAPT findings with blocker tracking
  ├─ writeSmoke(ss, allData)         // Critical test results
  ├─ writeCoverage(ss, allData)      // Test coverage metrics
  ├─ appendHistory(ss, allData)      // Historical timeline
  └─ appendVAPTHistory(ss, vaptData) // NEW: VAPT historical timeline

// VAPT data refresh (NEW)
refreshVAPTData()
  ├─ fetchAndProcessVAPTData_()      // Fetch from VAPT spreadsheet
  │   ├─ fetchAdHocVAPTData_()       // Ad Hoc VAPT findings
  │   ├─ fetchRegularVAPTData_()     // Regular VAPT findings
  │   └─ processVAPTData_()          // Combine and calculate metrics
  ├─ writeVAPT(ss, vaptData)         // Write to VAPT tab
  └─ appendVAPTHistory(ss, vaptData) // Append to history

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
| **VAPT** | Security findings blocker tracking | On refresh |
| **Smoke** | Critical smoke test results | On refresh |
| **Coverage** | Test coverage progress | On refresh |
| **Failure Scenario** | Failed test analysis | On refresh |
| **History** | Timeline of metrics | Append-only |
| **VAPT History** | Timeline of VAPT blocker trends | Append-only |
| **Config** | Module registry & notification config | Manual edit |
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

### Flow 3: VAPT Data Refresh (Pull from External Source)

```
TRIGGER: refreshDashboard() → refreshVAPTData()
   │
   ▼
Read Config
   │
   ├─► Get VAPT Spreadsheet ID (external source)
   │
   ▼
Open VAPT Spreadsheet
   │
   ├─► fetchAdHocVAPTData_(vaptSs)
   │    ├─ Read "Ad Hoc VAPT" tab (C1:Y100)
   │    ├─ Skip header rows (1-2)
   │    ├─ Extract data:
   │    │   ├─ Aplikasi (application name)
   │    │   ├─ Severity breakdown (Critical, High, Medium, Low, Info)
   │    │   ├─ Status (Open, Ready to Retest, Closed)
   │    │   └─ Production status
   │    └─ Return array of findings
   │
   ├─► fetchRegularVAPTData_(vaptSs)
   │    ├─ Read "Regular VAPT" tab (same structure)
   │    └─ Return array of findings
   │
   ▼
Process and Combine Data
   │
   ├─► processVAPTData_(adHocData, regularData)
   │    ├─ Group by Aplikasi
   │    ├─ Aggregate severity counts
   │    ├─ Calculate blocker count (Medium-Critical Open)
   │    ├─ Combine Ad Hoc + Regular findings
   │    └─ Return processed dataset
   │
   ▼
Write to Dashboard
   │
   ├─► writeVAPT(ss, vaptData)
   │    ├─ Clear VAPT tab data rows
   │    ├─ Write summary metrics (Total blocker, Apps with blocker)
   │    ├─ Write detail table per aplikasi
   │    └─ Apply conditional formatting (red for blocker > 0)
   │
   └─► appendVAPTHistory(ss, vaptData)
        ├─ Read current date
        ├─ Append row: [Date, Total Blocker, Ad Hoc Count, Regular Count]
        └─ Update VAPT chart data range
```

### Flow 4: VAPT Workflow Broadcast (Legacy)

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

## 🔒 VAPT Integration

### Overview

The VAPT (Vulnerability Assessment & Penetration Testing) integration provides centralized security findings tracking in the QA Portfolio Dashboard. This system pulls data from an external VAPT spreadsheet containing both Ad Hoc and Regular VAPT assessment results.

**Key Features:**
- Automated data fetch from external VAPT source spreadsheet
- Combined view of Ad Hoc + Regular VAPT findings
- Blocker tracking (Medium-Critical Open findings)
- Per-application severity breakdown (Critical, High, Medium, Low, Info)
- Historical trending with VAPT History timeline
- Integration with notifications (WhatsApp, Email, Google Chat)

### VAPT Data Source

**External Spreadsheet:** `17qeErP3VHxN7qcNQqhT6zGLukxZU4OKLmBMbsgsl1Rk`

**Tabs:**
- **Ad Hoc VAPT:** Security findings from ad-hoc assessments (C1:Y100)
- **Regular VAPT:** Scheduled/routine VAPT results (same structure)

**Data Structure:**
| Column | Field | Description |
|--------|-------|-------------|
| C | No | Entry number |
| D | Aplikasi | Application name |
| E-I | Severity | Critical, High, Medium, Low, Info counts |
| J-N | Status | Open, Ready to Retest, Closed counts |
| O | Production | Production status flag |

### VAPT Blocker Definition

**VAPT Blocker = Medium-Critical Open findings**

```javascript
// Blocker calculation
const blocker = (criticalOpen + highOpen + mediumOpen) || 0;

// Only count "Open" status, not "Ready to Retest" or "Closed"
```

**Target:** 0 VAPT blocker across all applications before production release

### VAPT Dashboard Tab Structure

**Summary Section (Rows 1-9):**
- Last refresh timestamp
- Total Blocker count (all apps combined)
- Apps with Blocker count
- Breakdown: Ad Hoc vs Regular VAPT

**Detail Table (Rows 10+):**
- Per-application severity breakdown
- Columns: Aplikasi | Blocker | Critical | High | Medium | Low | Info
- Conditional formatting: Red background if blocker > 0
- Target reminder: "🎯 Target: 0 blocker di semua aplikasi!"

**VAPT History Tab:**
- Timeline of blocker trends (append-only)
- Columns: Date | Total Blocker | Ad Hoc Count | Regular Count
- Used for trend charts in Web App dashboard

### Integration with Notifications

VAPT blocker data is included in all notification channels:

**WhatsApp / Email / Google Chat Format:**
```
📊 DAILY BUG REPORT
━━━━━━━━━━━━━

SUMMARY
▬ QA Bugs: 67 (10 apps)
  • Severity: Critical🟣 5  High🔴 8  Medium🟠 54
▬ VAPT Blocker: 20 (7 apps)
  • Severity: High🔴 3  Medium🟠 17

VAPT BLOCKER DETAIL
━━━━━━━━━━━━━

▬ Fleet Management Mobile: 5
  • High🔴 1  Medium🟠 4
▬ Dialur: 5
  • High🔴 2  Medium🟠 3
...
```

**Notification Logic:**
- Include if `vaptBlocker > 0`
- Show total blocker count + apps affected
- Show severity breakdown (hide zero counts)
- List detail per application with blocker > 0

---

## 🔒 VAPT Workflow Integration (Legacy)

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
| **VAPTTab.js** | 500 | VAPT dashboard builder | `buildVAPT()`, `writeVAPT()`, `appendVAPTHistory()` |
| **VAPTDataFetch.js** | 400 | VAPT data fetch & process | `refreshVAPTData()`, `fetchAdHocVAPTData_()`, `fetchRegularVAPTData_()`, `processVAPTData_()` |
| **JiraSync.js** | 2000+ | Jira integration | `syncAllJiraBugs()`, `_getBlockerData_()` |
| **Notifications.js** | 1750 | Alert system (WhatsApp, Email, GChat) | `sendWhatsAppNotification_()`, `sendEmailNotification_()`, `sendGoogleChatNotification_()` |
| **VAPTBroadcast.js** | 900 | Bulk update utility (Legacy) | `broadcastVAPTStatusUpdate()`, `broadcastFixNoteAndAppendix()` |
| **WebApp.html** | 31KB | Web dashboard UI | HTML/CSS/JS frontend with VAPT charts |
| **WebAppBackend.js** | 14KB | Web API backend | `doGet()`, `getModuleData()`, `getVAPTData()` |
| **InitVAPT.js** | 200 | VAPT initialization | `initVAPTConfig()` |
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

## 📢 Notification System Architecture

### Overview

The notification system provides automated alerts for QA bugs and VAPT security findings through multiple channels: WhatsApp (Fonnte), Email (Gmail), and Google Chat. Notifications are triggered on schedule or manually, aggregating data from Dashboard tabs.

**Key Components:**
- **Data Source:** Dashboard Bugs tab + VAPT tab (aggregated from all QATM modules)
- **Trigger:** Time-based (configurable: hourly, daily, custom schedule)
- **Channels:** WhatsApp (Fonnte API), Email (MailApp), Google Chat (Webhook)
- **Config:** Per-module notification settings in Config sheet (columns L-U)

### End-to-End Notification Flow

```
┌─────────────────────────────────────────────────────────────────────┐
│                    STAGE 1: DATA COLLECTION                          │
└─────────────────────────────────────────────────────────────────────┘

[JIRA]
  │
  │ 1. Jira Sync Trigger (Every 1 hour)
  ▼
syncAllJiraBugs()
  │
  ├─► Read Config sheet
  │    └─ Get modules with Jira Sync = TRUE
  │
  ├─► FOR EACH MODULE:
  │    ├─ Fetch bugs from Jira API
  │    │   └─ GET /rest/api/3/search?jql=project=XXX
  │    │
  │    ├─ Parse response:
  │    │   ├─ Bug ID (TEST-123)
  │    │   ├─ Summary, Priority, Status
  │    │   ├─ Environment (Production/UAT/Dev)
  │    │   └─ Custom fields
  │    │
  │    └─► Write to QATM BugReport sheet
  │         ├─ Match by Bug ID (update if exists)
  │         └─ Add new bugs
  │
  └─► Log sync results

┌─────────────────────────────────────────────────────────────────────┐
│                STAGE 2: DASHBOARD AGGREGATION                        │
└─────────────────────────────────────────────────────────────────────┘

refreshDashboard()
  │
  ├─► Read Config sheet
  │    └─ Get all ACTIVE modules
  │
  ├─► FOR EACH MODULE:
  │    ├─ Open QATM spreadsheet
  │    ├─ Read BugReport sheet
  │    │   ├─ Count bugs by priority (Critical, High, Medium, Low)
  │    │   ├─ Count bugs by environment (Prod, UAT, Dev)
  │    │   ├─ Calculate blocker count (Medium-Critical, not Closed)
  │    │   └─ Calculate PROD bugs (Environment = Production)
  │    │
  │    └─► Aggregate to Dashboard Bugs tab
  │         └─ Columns: Project, Modul, Submodul, Total, Critical, High,
  │            Medium, Low, Blocker, Dev, UAT, Prod
  │
  ├─► Fetch VAPT data (refreshVAPTData)
  │    ├─ Read external VAPT spreadsheet
  │    ├─ Combine Ad Hoc + Regular VAPT
  │    ├─ Calculate blocker count per app (Medium-Critical Open)
  │    └─► Write to Dashboard VAPT tab
  │
  └─► Update Overview summary metrics

┌─────────────────────────────────────────────────────────────────────┐
│                 STAGE 3: NOTIFICATION TRIGGER                        │
└─────────────────────────────────────────────────────────────────────┘

[TIME-BASED TRIGGER] or [MANUAL TRIGGER]
  │
  │ setupDailyBlockerNotification()
  │ - Schedule options:
  │   • Single: 9 (daily at 9:00)
  │   • Multiple: 9,14,18 (3x per day)
  │   • Interval: 4h (every 4 hours)
  │
  ▼
sendBlockerNotification()
  │
  ├─► Read Config sheet
  │    ├─ Get modules with notifications enabled
  │    └─ Get notification config per module:
  │         ├─ Google Chat Webhook (col L)
  │         ├─ Email recipients (col O)
  │         └─ Enable flags (col N, P)
  │
  ├─► Read Dashboard Overview + Bugs + VAPT tabs
  │    └─ getBlockerData_(overview, cfg)
  │         ├─ Read Bugs tab data (all modules)
  │         ├─ Read VAPT tab data (all apps)
  │         ├─ Calculate totals:
  │         │   ├─ Total QA Blocker count
  │         │   ├─ Total PROD bugs count
  │         │   ├─ Total VAPT Blocker count
  │         │   └─ Per-module breakdown
  │         │
  │         └─ Return blockerData object:
  │              {
  │                modules: [...],         // Per-module bug data
  │                totalBlockers: N,
  │                totalProdBugs: N,
  │                vaptBlocker: N,
  │                vaptApps: [...],        // VAPT detail per app
  │                vaptBreakdown: {...},   // Severity counts
  │                timestamp: "..."
  │              }
  │
  ├─► Check if notification needed
  │    └─ Skip if: totalBlockers = 0 AND totalProdBugs = 0 AND vaptBlocker = 0
  │
  └─► If blockers exist, send notifications

┌─────────────────────────────────────────────────────────────────────┐
│                  STAGE 4: MULTI-CHANNEL DELIVERY                     │
└─────────────────────────────────────────────────────────────────────┘

[GROUP BY WEBHOOK/EMAIL]
  │
  ├─► Group modules by Google Chat Webhook
  │    └─ {webhookUrl: [module1, module2, ...]}
  │
  ├─► Group modules by Email recipients
  │    └─ {emailList: [module1, module2, ...]}
  │
  └─► Get GLOBAL WhatsApp config (row 4, col S-U)

[SEND NOTIFICATIONS IN PARALLEL]
  │
  ├─► Google Chat (Per Webhook)
  │    │
  │    FOR EACH webhookUrl:
  │      │
  │      ├─ Build plain text message (Google Chat markdown)
  │      │   ├─ Header: "📊 DAILY BUG REPORT"
  │      │   ├─ Summary section:
  │      │   │   ├─ QA Bugs: N (X apps) + severity breakdown
  │      │   │   └─ VAPT Blocker: N (X apps) + severity breakdown
  │      │   │
  │      │   ├─ VAPT Blocker Detail (if any)
  │      │   │   └─ List apps with blocker > 0
  │      │   │
  │      │   ├─ Production Bugs section (if any)
  │      │   │   └─ <users/all> mention + per-module breakdown
  │      │   │
  │      │   ├─ QA Blocker Bugs section
  │      │   │   └─ Per-module breakdown with severity
  │      │   │
  │      │   └─ Footer: Dashboard links (Web App, Sheet)
  │      │
  │      ├─ Send to Google Chat webhook
  │      │   └─ POST https://chat.googleapis.com/v1/spaces/.../messages
  │      │
  │      └─ Log result
  │
  ├─► Email (Per Recipient Group)
  │    │
  │    FOR EACH emailList:
  │      │
  │      ├─ Build HTML email
  │      │   ├─ Header (red if PROD bugs, orange if blockers)
  │      │   ├─ Summary section (table with severity badges)
  │      │   ├─ VAPT Apps Detail (if any)
  │      │   ├─ Alert message (emergency protocol if PROD bugs)
  │      │   ├─ QA Modules breakdown (styled tables)
  │      │   ├─ Priority levels explanation
  │      │   └─ Quick links section
  │      │
  │      ├─ Send via MailApp.sendEmail()
  │      │   └─ Subject: "🚨🚨 URGENT PROD BUGS" or "🚨 QA BLOCKER ALERT"
  │      │
  │      └─ Log result
  │
  └─► WhatsApp (GLOBAL - Single Group)
       │
       IF whatsappEnabled AND groupId valid:
         │
         ├─ Build plain text message (WhatsApp markdown)
         │   ├─ Header: "📊 *DAILY BUG REPORT*"
         │   ├─ Summary section (▬ bullets)
         │   ├─ VAPT Blocker Detail (if any)
         │   ├─ Production Bugs section (if any)
         │   ├─ QA Blocker Bugs section
         │   └─ Footer: Dashboard links
         │
         ├─ Send to Fonnte API
         │   └─ POST https://api.fonnte.com/send
         │        {
         │          target: "120363xxx@g.us",
         │          message: "...",
         │          Authorization: fontteToken
         │        }
         │
         └─ Log result

┌─────────────────────────────────────────────────────────────────────┐
│                     STAGE 5: CONFIRMATION                            │
└─────────────────────────────────────────────────────────────────────┘

Show Alert Dialog:
  │
  └─ "📤 Notifications Sent!
      ✅ Google Chat: N message(s) sent
      ✅ Email: N message(s) sent
      ✅ WhatsApp: N message(s) sent

      📊 Summary:
      • Total Open Blockers: N
      • Total PROD BUGS: N
      • VAPT Blocker: N (X apps)
      • Modules with issues: N"
```

### Notification Configuration

**Config Sheet Structure (Row 4+):**

| Column | Field | Type | Description |
|--------|-------|------|-------------|
| L | Google Chat Webhook | STRING | Webhook URL from Google Chat Space |
| M | Schedule | STRING | "9", "9,14,18", or "4h" (interval) |
| N | Enable Chat | BOOLEAN | ☑ TRUE to send Google Chat notifications |
| O | Email Recipients | STRING | Comma-separated emails |
| P | Enable Email | BOOLEAN | ☑ TRUE to send Email notifications |
| S (row 4) | WhatsApp Group ID | STRING | Format: 120363xxx@g.us (GLOBAL) |
| T (row 4) | Fonnte Token | STRING | API token from Fonnte (GLOBAL) |
| U (row 4) | Enable WhatsApp | BOOLEAN | ☑ TRUE to send WhatsApp notifications (GLOBAL) |

**Schedule Format Examples:**
- `9` → Daily at 9:00
- `9,14,18` → 3x per day (9:00, 14:00, 18:00)
- `4h` → Every 4 hours (supports: 1h, 2h, 4h, 6h, 8h, 12h)

**Notification Strategy:**
- **Per-Module (Google Chat, Email):** Each module can have different webhooks/emails
- **GLOBAL (WhatsApp):** Single WhatsApp group receives all notifications
- **Aggregation:** Multiple modules with same webhook/email receive 1 combined message

### Message Format Comparison

**Google Chat (Markdown with hyperlinks):**
```
📊 *DAILY BUG REPORT*
━━━━━━━━━━━━━

SUMMARY
▬ QA Bugs: 67 (10 apps)
  • Severity: Critical🟣 5  High🔴 8  Medium🟠 54
▬ VAPT Blocker: 20 (7 apps)
  • Severity: High🔴 3  Medium🟠 17

🚨🚨🚨 *PRODUCTION BUGS* 🚨🚨🚨
<users/all>
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

*SIPGN - 4*
📌 AI Surveillance: *2* PROD bug(s)

⚠️ *URGENT - IMMEDIATE ACTION REQUIRED!*

...

🔗 <https://script.google.com/.../exec|📊 Web Dashboard (QA & VAPT)>
```

**Email (HTML with styled tables):**
- Red/Orange header based on severity
- Styled summary table with severity badges
- Alert box with emergency protocol (if PROD bugs)
- Per-module breakdown with color coding
- Quick links section with buttons

**WhatsApp (Plain text, WhatsApp-friendly):**
```
📊 *DAILY BUG REPORT*
📅 2026-04-06 14:59:53
━━━━━━━━━━━━━

SUMMARY
▬ QA Bugs: 67 (10 apps)
  • Severity: Critical🟣 5  High🔴 8  Medium🟠 54
▬ VAPT Blocker: 20 (7 apps)
  • Severity: High🔴 3  Medium🟠 17

...

━━━━━━━━━━━━━
🔗 *Dashboard Links:*
📊 Web Dashboard: https://script.google.com/.../exec (QA & VAPT)
📋 Sheet Overview: https://docs.google.com/...
🐛 Sheet Bugs: https://docs.google.com/...

_Automated Daily Report - QA Dashboard_
```

### Error Handling & Retry Logic

**Notification Failures:**
```javascript
// Each channel has independent error handling
try {
  sendGoogleChatNotification_(webhookUrl, blockerData);
} catch (e) {
  Logger.log('❌ Google Chat notification failed: ' + e.message);
  // Continue with other channels
}
```

**Common Failure Scenarios:**
1. **Google Chat:** Invalid webhook URL, space deleted
2. **Email:** Invalid email address, quota exceeded (MailApp: 100/day)
3. **WhatsApp:** Invalid group ID, expired token, rate limit

**No Retry:** System does NOT retry failed notifications (to avoid spam)

**Logging:** All results logged to Apps Script Execution log

### Web App URL Management

**Problem:** Notifications need latest Web App deployment URL

**Solution:** Script Properties (persistent storage)

**Setup:**
```javascript
// Run once after deploying new Web App version
function autoSetWebAppUrl() {
  const LATEST_WEBAPP_URL = 'https://script.google.com/.../exec';
  PropertiesService.getScriptProperties()
    .setProperty('WEB_APP_URL', LATEST_WEBAPP_URL);
}
```

**Usage in Notifications:**
```javascript
// Notifications.js reads from Script Properties
const scriptProps = PropertiesService.getScriptProperties();
const dashboardWebAppUrl = scriptProps.getProperty('WEB_APP_URL') || LATEST_WEBAPP_URL;
```

**Fallback:** If property not set, fallback to `LATEST_WEBAPP_URL` constant (line 8)

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
| **2.1** | 2026-04-06 | VAPT external source integration: VAPTTab.js, VAPTDataFetch.js, automated blocker tracking from external VAPT spreadsheet, notification integration |
| **2.0** | 2026-03-31 | Added VAPT workflow integration (Bug status lifecycle) |
| **1.5** | 2026-03-17 | Added dual environment support (Testing + Production) |
| **1.0** | 2026-02-27 | Initial architecture documentation |

---

**Maintained by:** QA Team @ INA Digital
**Last Review:** April 6, 2026
