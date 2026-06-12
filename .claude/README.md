# Apps Script Automation Tools - AI Context

**Centralized AI context documentation for all automation projects**

This document provides comprehensive context for Claude Code to assist with development across all Apps Script automation tools.

---

## 🎯 Projects Overview

### 1. QA Test Management Template
**Purpose:** Comprehensive test case management and execution tracking
**Location:** `projects/qa-test-management/`
**Key Features:** Instant setup, test case repository, execution tracking, smart dashboards, performance testing

### 2. QA Portfolio Dashboard
**Purpose:** Centralized monitoring dashboard aggregating test results from multiple projects
**Location:** `projects/qa-dashboard/`
**Key Features:** Multi-project aggregation, consolidated KPIs, auto-refresh, broadcast updates, automation config generation

### 3. QA Team Management (SSOT Architecture)
**Purpose:** End-to-end team management with auto-sync from central dashboard
**Location:** `projects/team-member-management/`
**Key Features:** Auto-sync from QA Dashboard, smart team generation, data preservation, difficulty-based sorting

**Architecture:**
```
QA Dashboard (Config Sheet) ← SSOT: Add submodule + PIC QA
         ↓ Auto-Sync (manual or scheduled)
QA Team Management
  ├─ Project Config (synced from Dashboard)
  ├─ Team Members (auto-generated, manual data preserved)
  └─ Dashboard (Person Matrix sorted by difficulty)
```

**Core Principles:**
- **SSOT (Single Source of Truth):** QA Dashboard Config Sheet is master data source
- **Data Preservation:** Manual fills (Email, Phone, NP, etc.) never lost on regeneration
- **Smart Sorting:** Rated items first (by total difficulty), unrated last; Person Matrix by avgScore descending
- **Column Segregation:** Auto-generated left (A-F: No, Name, Projects, Modules, Submodules, Status), Manual right (G-S: NP, Email, Phone, etc.)
- **End-to-End Workflow:** Run All = Sync → Generate → Refresh (same for scheduled triggers)

**Key Files:**
- `CrossSpreadsheetSync.js` - Sync data from Dashboard (Dashboard ID at cell Q2)
- `TeamMember.js` - Team member management with data preservation
- `Dashboard.js` - Person Assignment Matrix (sorted by difficulty)
- `ProjectConfig.js` - Project/Module/Submodule configuration
- `Menu.js` - User interface and workflow orchestration

**Data Preservation Pattern:**
```javascript
// Read existing manual data into Map
const existingManualDataMap = new Map();
existingData.forEach(row => {
  const name = row[1]; // Column B (Name)
  if (name) {
    existingManualDataMap.set(name, row.slice(6, 19)); // Preserve G-S
  }
});

// Apply preserved data when regenerating
if (existingManual) {
  row.push(...existingManual); // Preserve existing
} else {
  row.push('', '', '', 'Quality Engineer', ...); // Defaults
}
```

**Validation Strategy:**
- Auto-generated columns (A-F): `.setAllowInvalid(true)` - Allow synced values from Dashboard
- Manual columns: Strict validation where appropriate
- Dropdowns from helper columns (K, L, M) with suggestions but allow override

### 4. MOM Rolling & PIC Reminder
**Purpose:** Automated rotation system for meeting note-takers
**Location:** `projects/mom-rolling-pic/`
**Key Features:** Smart rotation, schedule management, auto doc creation, email notifications

### 5. KPI Tracker
**Purpose:** Team performance tracking and KPI monitoring
**Location:** `projects/kpi-tracker/`
**Key Features:** KPI dashboard, individual tracking, trend analysis, goal setting

---

## 📋 Available Commands

Ketik command ini di Claude Code untuk menjalankan workflow otomatis:

### 🚀 Deployment

- **`/deploy-dashboard`** - Deploy QA Dashboard ke Apps Script
  - Check status & uncommitted changes
  - Push menggunakan clasp
  - Verify deployment berhasil

- **`/deploy-template`** - Deploy QA Test Management template ke Apps Script
  - Check status & uncommitted changes
  - Push menggunakan clasp
  - Verify deployment berhasil

### 🔄 Sync & Update

- **`/sync-clasp`** - Pull latest changes dari Apps Script projects
  - Sync Dashboard project
  - Sync Template project
  - Show diff dan ask untuk commit

### 🔧 Development

- **`/fix-dashboard`** - Debug dan fix Dashboard issues
  - Analyze error atau issue
  - Propose fixes dengan explanation
  - Implement dan test changes
  - Commit dan deploy

- **`/update-template`** - Update template dengan features baru
  - Add new functionality
  - Maintain backward compatibility
  - Update documentation
  - Deploy changes

### 📚 Help

- **`/qa-help`** - Show comprehensive help
  - List semua commands
  - Project structure
  - Quick workflows
  - Documentation links
  - Tips & tricks

## 🎯 Usage Examples

```bash
# Deploy dashboard after making changes
/deploy-dashboard

# Pull latest from Apps Script before starting work
/sync-clasp

# Get help when stuck
/qa-help

# Fix a specific issue
/fix-dashboard
# Then describe the issue in natural language
```

## 💡 Tips

1. **Start with `/qa-help`** untuk overview lengkap
2. Gunakan **`/sync-clasp`** sebelum mulai development
3. Gunakan **`/deploy-*`** untuk quick deployment tanpa manual clasp commands
4. Gunakan **`/fix-*`** untuk debugging dengan AI assistance

## 🔗 Related Files

- `CLASP_WORKFLOW.md` - Detailed clasp workflow guide
- `setup-clasp.sh` - Quick setup script
- `.github/workflows/deploy-apps-script.yml` - CI/CD pipeline

## 📝 Notes

Commands ini adalah wrapper untuk common tasks yang:
- Reduce manual typing
- Ensure best practices
- Automate repetitive workflows
- Provide contextual help

Setiap command sudah include project-specific context (Script IDs, folder structure, etc.).

---

## 🏗️ Common Architecture Patterns

### 1. SSOT (Single Source of Truth)
**Used in:** QA Team Management, QA Dashboard
**Pattern:** One master data source, all other sheets consume from it
**Implementation:** Cross-spreadsheet sync with SpreadsheetApp.openById()

### 2. Data Preservation with Map
**Used in:** QA Team Management (TeamMember.js), QA Dashboard (Sync)
**Pattern:** Store existing data in Map before regenerating, reapply after
```javascript
const existingDataMap = new Map();
existingData.forEach(row => {
  const key = row[keyColumn];
  existingDataMap.set(key, preservedColumns);
});
```

### 3. Smart Sorting
**Used in:** QA Team Management (Dashboard, Sync)
**Pattern:** Multi-tier sorting - rated first (by score desc), unrated last
```javascript
data.sort((a, b) => {
  const aHasData = a.score > 0;
  const bHasData = b.score > 0;
  if (aHasData && !bHasData) return -1;
  if (!aHasData && bHasData) return 1;
  if (aHasData && bHasData) return b.score - a.score;
  return 0;
});
```

### 4. Column Segregation
**Used in:** QA Team Management, QA Dashboard
**Pattern:** Auto-generated columns left, manual fill columns right
**Benefit:** Clear separation of concerns, easier data preservation

### 5. Helper Columns for Validation
**Used in:** QA Team Management (ProjectConfig.js)
**Pattern:** Create helper columns (K, L, M) with UNIQUE() formulas for dropdown validation
**Implementation:** Data validation from range with `.setAllowInvalid(true)` for flexibility

### 6. End-to-End Workflows
**Used in:** QA Team Management (Menu.js, CrossSpreadsheetSync.js)
**Pattern:** Chain operations in consistent order: Sync → Generate → Refresh
**Implementation:** Both manual (Run All) and scheduled (trigger) execute same workflow

---

## 🔧 Troubleshooting Guide

### Common Issues & Solutions

#### 1. Data Validation Errors
**Error:** "The data you entered violates data validation rules"
**Cause:** Strict validation (`.setAllowInvalid(false)`) blocking synced data
**Solution:** Use `.setAllowInvalid(true)` for auto-generated columns
**Location:** `projects/team-member-management/src/TeamMember.js:162,178,194`

#### 2. Dashboard ID Issues
**Error:** "Illegal spreadsheet id or key"
**Cause:** Dashboard ID not properly configured or wrong cell location
**Solution:** Check cell Q2 in Project Config tab, paste full URL or just ID
**Location:** `projects/team-member-management/src/CrossSpreadsheetSync.js:17`

#### 3. False Automation Matching
**Issue:** External QA showing automation results despite no internal testing
**Cause:** Generic fallback matching by module name
**Solution:** Prioritize contract-based matching, skip External QA modules
**Location:** `projects/qa-dashboard/src/MasterDashboard.js:getDashboardAutomationAliases_()`

#### 4. Missing Data After Refresh
**Issue:** Manual data (Email, Phone) lost after regenerating Team Members
**Cause:** Data preservation logic not implemented
**Solution:** Use Map-based preservation pattern (see Architecture Patterns #2)
**Location:** `projects/team-member-management/src/TeamMember.js:generateTeamMembersFromConfig()`

#### 5. Auto-Sync Incomplete
**Issue:** Scheduled trigger only syncs + refreshes, missing Generate step
**Cause:** autoSyncAndRefresh() missing Generate call
**Solution:** Add Generate step between Sync and Refresh
**Location:** `projects/team-member-management/src/CrossSpreadsheetSync.js:autoSyncAndRefresh()`

---

## 🎯 Development Principles

### Code Quality
- ✅ Modular architecture with separation of concerns
- ✅ Comprehensive error handling with try-catch
- ✅ Logging for debugging and audit trails
- ✅ Idempotent operations (safe to re-run)
- ✅ Input validation and sanitization
- ✅ Performance optimized (batch operations, minimal API calls)

### Data Integrity
- ✅ SSOT architecture - one master source
- ✅ Data preservation - never lose manual fills
- ✅ Smart sorting - preserve user intent
- ✅ Validation with flexibility - allow programmatic updates
- ✅ Clear column segregation - auto vs manual

### User Experience
- ✅ One-click workflows (Run All)
- ✅ Scheduled automation (set and forget)
- ✅ Clear menu organization
- ✅ Informative success/error messages
- ✅ Dashboard visualizations

---

## 📚 Reference Documentation

- **[Main README](../README.md)** - Comprehensive project overview
- **[Architecture](../ARCHITECTURE.md)** - System design & data flows
- **[Clasp Workflow](../CLASP_WORKFLOW.md)** - Development workflow guide
- **[Environments](../ENVIRONMENTS.md)** - Testing vs Production setup

---

## 💡 Quick Reference

### Cell Locations (QA Team Management)
- Dashboard Spreadsheet ID: **Q2** (merged Q2:S2) in Project Config tab
- Helper Columns: **K** (Projects), **L** (Modules), **M** (Submodules)
- Auto-Generated Columns: **A-F** (No, Name, Projects, Modules, Submodules, Status)
- Manual Fill Columns: **G-S** (NP, Email, Email2, HP, JoinDate, Title, Role, etc.)

### Key Constants
- `CONFIG_TAB_NAME = 'Project'`
- `TEAM_TAB_NAME = 'Team Members'`
- `DASHBOARD_TAB_NAME = 'Dashboard'`
- `DASHBOARD_ID_COL = 17` (Column Q)
- `TEAM_AUTO_GENERATED_COLS = 6` (Columns A-F)

### Run All Workflow
```
1. Sync from QA Dashboard (CrossSpreadsheetSync.syncFromDashboard)
2. Generate Team Members (TeamMember.generateTeamMembersFromConfig)
3. Refresh Dashboard (Dashboard.createDashboard)
```

---

**Last Updated:** 2026-06-13
**Version:** 1.0
